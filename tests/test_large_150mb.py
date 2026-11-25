from __future__ import annotations

import io
import json
import os
import socket
import sys
import threading
import time
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
os.environ.setdefault("CLICKSHIFT_HOME", str(Path(__file__).parent / "_tmp_home"))

from clickshift import cli, debuglog, engine, learning, monitor, updater  # noqa: E402
from clickshift.formats import ConversionResult  # noqa: E402


def test_zip_600mb_limits_bomb_guard_and_150mb_ooxml_sanitizer(tmp_path):
    assert engine._MAX_ZIP_ENTRY_BYTES >= 600 * 1024 * 1024
    assert engine._MAX_ZIP_TOTAL_BYTES >= 2400 * 1024 * 1024
    assert engine._MAX_ZIP_ENTRIES >= 50_000

    # 1. Zip-bomb ratio > 200x on >32 MiB entry is rejected
    bomb_info = zipfile.ZipInfo("word/document.xml")
    bomb_info.file_size = 100 * 1024 * 1024
    bomb_info.compress_size = 100 * 1024  # 1024x ratio
    with pytest.raises(ValueError, match="compression ratio"):
        engine._check_zip_member_safety(bomb_info, engine._MAX_ZIP_ENTRY_BYTES)

    # 2. Clean 150 MiB .docx with safe internal rels returns src directly (zero temp copy!)
    clean_docx = tmp_path / "clean_150mb.docx"
    safe_rels = (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>'
        b"</Relationships>"
    )
    chunk_1mb = os.urandom(1024 * 1024)
    with zipfile.ZipFile(clean_docx, "w", allowZip64=True) as zf:
        zf.writestr("word/_rels/document.xml.rels", safe_rels, compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("word/document.xml", b"<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>")
        info = zipfile.ZipInfo("word/media/clip.bin")
        info.compress_type = zipfile.ZIP_STORED
        with zf.open(info, "w", force_zip64=True) as mfh:
            for _ in range(150):
                mfh.write(chunk_1mb)

    assert clean_docx.stat().st_size >= 150 * 1024 * 1024
    called_tmp = []
    out_clean = engine._sanitize_ooxml_for_com(clean_docx, lambda: (called_tmp.append(1) or tmp_path))
    assert out_clean == clean_docx
    assert not called_tmp

    # 3. Unsafe external template rel in a >100 MiB .docx is stripped via streaming copy
    dirty_docx = tmp_path / "dirty_large.docx"
    bad_rels = (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="https://evil.example/t.dotm" TargetMode="External"/>'
        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://safe.example" TargetMode="External"/>'
        b"</Relationships>"
    )
    with zipfile.ZipFile(dirty_docx, "w", allowZip64=True) as zf:
        zf.writestr("word/_rels/document.xml.rels", bad_rels, compress_type=zipfile.ZIP_DEFLATED)
        info = zipfile.ZipInfo("word/media/huge.bin")
        info.compress_type = zipfile.ZIP_STORED
        with zf.open(info, "w", force_zip64=True) as mfh:
            for _ in range(110):  # 110 MiB > old 100 MiB cap
                mfh.write(chunk_1mb)

    stage_dir = tmp_path / "san_out"
    stage_dir.mkdir()
    sanitized = engine._sanitize_ooxml_for_com(dirty_docx, stage_dir)
    assert sanitized != dirty_docx and sanitized.exists()
    with zipfile.ZipFile(sanitized, "r") as zcheck:
        rels_after = zcheck.read("word/_rels/document.xml.rels")
        assert b"evil.example" not in rels_after
        assert b"safe.example" in rels_after
        assert zcheck.getinfo("word/media/huge.bin").file_size == 110 * 1024 * 1024

    # 4. RTF scanner detects unsafe DDEAUTO past the old 16 MiB cutoff
    late_rtf = tmp_path / "late_dde.rtf"
    with open(late_rtf, "wb") as rfh:
        rfh.write(b"{\\rtf1\\ansi\n")
        block = b"safe paragraph text " * 52_000  # ~1 MB
        for _ in range(20):
            rfh.write(block)
        rfh.write(b"{\\field{\\*\\fldinst DDEAUTO cmd.exe}}}")
    assert late_rtf.stat().st_size > 19 * 1024 * 1024
    with pytest.raises(RuntimeError, match="unsafe external template/DDE"):
        engine._check_rtf_for_com(late_rtf)


def test_real_150mb_streaming_text_to_html_and_docx(tmp_path):
    src_md = tmp_path / "report_150mb.md"
    target_bytes = 150 * 1024 * 1024  # exact 150 MiB (157,286,400 bytes)
    line_block = (
        "# Section Heading\n"
        "- Bullet item with **bold** and `inline_code` & <tag> safety\n"
        + ("Plain body paragraph line for 150 MB streaming verification 0123456789.\n" * 140)
    )
    block_bytes = line_block.encode("utf-8")
    repeats = (target_bytes // len(block_bytes)) + 1
    with open(src_md, "wb") as fh:
        written = 0
        for _ in range(repeats):
            take = min(len(block_bytes), target_bytes - written)
            if take <= 0:
                break
            fh.write(block_bytes[:take])
            written += take
    assert src_md.stat().st_size == target_bytes

    eng = engine.ConversionEngine(enable_ml=False)
    try:
        # 150 MiB .md -> .html via streaming chunk writer
        res_html = eng.convert(src_md, "html")
        assert res_html.ok, res_html.error
        out_html = Path(res_html.target)
        assert out_html.exists() and out_html.stat().st_size > target_bytes
        with open(out_html, "rb") as hfh:
            head = hfh.read(512)
            hfh.seek(max(0, out_html.stat().st_size - 64))
            tail = hfh.read(64)
        assert b"<!DOCTYPE html>" in head and b"&lt;tag&gt;" in head
        assert b"</pre></body></html>" in tail
        out_html.unlink()

        # 150 MiB .md -> .docx via streaming OOXML writer
        steps: list[float] = []
        with engine.subprogress_scope(lambda d, t, _m: steps.append(d / float(t))):
            res_docx = eng.convert(src_md, "docx")
        assert res_docx.ok, res_docx.error
        assert steps and steps[-1] > 0.5
        out_docx = Path(res_docx.target)
        assert out_docx.exists() and out_docx.stat().st_size > 100_000
        with zipfile.ZipFile(out_docx, "r") as zf:
            names = set(zf.namelist())
            assert {"[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/styles.xml"} <= names
            with zf.open("word/document.xml") as dfh:
                xml_head = dfh.read(4096)
            assert b"Section Heading" in xml_head and b"Heading1" in xml_head
        out_docx.unlink()
    finally:
        eng.shutdown()
        src_md.unlink(missing_ok=True)


def test_epub_preserves_opf_spine_order(tmp_path):
    epub_path = tmp_path / "book.epub"
    container_xml = (
        '<?xml version="1.0"?>'
        '<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">'
        '<rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>'
        "</container>"
    )
    # Spine intentionally orders chap2 -> chap10 -> chap1 (alphabetical sort would produce chap1, chap10, chap2!)
    opf_xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<package xmlns="http://www.idpf.org/2007/opf" version="2.0">'
        "<manifest>"
        '<item id="c1" href="chap1.xhtml" media-type="application/xhtml+xml"/>'
        '<item id="c2" href="chap2.xhtml" media-type="application/xhtml+xml"/>'
        '<item id="c10" href="chap10.xhtml" media-type="application/xhtml+xml"/>'
        "</manifest>"
        "<spine>"
        '<itemref idref="c2"/>'
        '<itemref idref="c10"/>'
        '<itemref idref="c1"/>'
        "</spine>"
        "</package>"
    )
    with zipfile.ZipFile(epub_path, "w") as zf:
        zf.writestr("META-INF/container.xml", container_xml)
        zf.writestr("OEBPS/content.opf", opf_xml)
        zf.writestr("OEBPS/chap1.xhtml", "<html><head><style>body{color:red}</style></head><body><p>THIRD_IN_SPINE</p></body></html>")
        zf.writestr("OEBPS/chap2.xhtml", "<html><body><script>var x=1;</script><p>FIRST_IN_SPINE</p></body></html>")
        zf.writestr("OEBPS/chap10.xhtml", "<html><body><p>SECOND_IN_SPINE</p></body></html>")

    out_txt = tmp_path / "book.txt"
    engine.conv_epub_extract(epub_path, out_txt)
    txt = out_txt.read_text(encoding="utf-8-sig")
    assert "var x=1" not in txt and "color:red" not in txt
    assert txt.index("FIRST_IN_SPINE") < txt.index("SECOND_IN_SPINE") < txt.index("THIRD_IN_SPINE")


def test_pillow_200mp_webp_16383_clamp_and_smart_embed(tmp_path):
    from PIL import Image

    # 1. >180 MP image opens and converts without DecompressionBombError
    big_png = tmp_path / "scan_200mp.png"
    with engine._allow_large_images(Image):
        im_200mp = Image.new("L", (20_000, 10_000), color=240)
        im_200mp.save(big_png, format="PNG", compress_level=1)
        im_200mp.close()

    out_jpg = tmp_path / "scan_200mp.jpg"
    engine.conv_pillow(big_png, out_jpg)
    with engine._allow_large_images(Image), Image.open(out_jpg) as chk:
        assert chk.size == (20_000, 10_000)

    # 2. >16,383 px wide image clamps proportionally for WebP instead of failing
    wide_png = tmp_path / "wide.png"
    im_wide = Image.new("RGB", (18_000, 360), (20, 80, 160))
    im_wide.save(wide_png, format="PNG", compress_level=1)
    im_wide.close()

    out_webp = tmp_path / "wide.webp"
    engine.conv_pillow(wide_png, out_webp)
    with Image.open(out_webp) as wchk:
        assert wchk.width == 16_383
        assert 300 <= wchk.height <= 360

    # 3. Smart raster normalization: alpha stays PNG, CMYK/opaque photo becomes JPEG
    rgba_buf = io.BytesIO()
    Image.new("RGBA", (120, 120), (255, 0, 0, 128)).save(rgba_buf, format="PNG")
    res_rgba = engine._normalize_image_bytes(rgba_buf.getvalue())
    assert res_rgba is not None and res_rgba[1] == "image/png"

    cmyk_buf = io.BytesIO()
    Image.new("CMYK", (400, 400), (10, 20, 30, 5)).save(cmyk_buf, format="TIFF")
    res_cmyk = engine._normalize_image_bytes(cmyk_buf.getvalue())
    assert res_cmyk is not None and res_cmyk[1] == "image/jpeg"


def test_pptx_to_pdf_deduplicates_images_and_docx_table_pagination(tmp_path):
    import docx
    import pptx
    from pptx.util import Inches
    from PIL import Image

    fitz = engine._import_fitz()

    # 1. 12 slides sharing the same photo -> PDF embeds 1 shared image xref
    img_path = tmp_path / "bg.jpg"
    Image.new("RGB", (600, 400), (30, 90, 200)).save(img_path, format="JPEG", quality=90)
    deck_path = tmp_path / "shared_bg.pptx"
    prs = pptx.Presentation()
    blank = prs.slide_layouts[6]
    for i in range(12):
        sl = prs.slides.add_slide(blank)
        sl.shapes.add_picture(str(img_path), Inches(1), Inches(1), width=Inches(5))
    prs.save(str(deck_path))

    out_pdf = tmp_path / "shared_bg.pdf"
    engine.conv_pptx_to_pdf(deck_path, out_pdf)
    with fitz.open(str(out_pdf)) as pdf_doc:
        assert len(pdf_doc) == 12
        all_xrefs = {img[0] for page in pdf_doc for img in page.get_images(full=True)}
        assert len(all_xrefs) == 1

    # Also verify streaming ODP deduplicates the 12 identical slide images into 1 Pictures/* entry
    out_odp = tmp_path / "shared_bg.odp"
    engine.conv_pptx_to_odp(deck_path, out_odp)
    with zipfile.ZipFile(out_odp, "r") as zodp:
        names = zodp.namelist()
        assert names[0] == "mimetype"
        pic_entries = [n for n in names if n.startswith("Pictures/")]
        assert len(pic_entries) == 1

    # 2. 55-row DOCX table paginates across >= 3 PPTX slides and extracts cleanly to text
    big_tbl_docx = tmp_path / "table55.docx"
    d = docx.Document()
    d.add_heading("Sales Report", level=1)
    tbl = d.add_table(rows=55, cols=3)
    tbl.cell(0, 0).text = "ID"
    tbl.cell(0, 1).text = "Region"
    tbl.cell(0, 2).text = "Revenue"
    for r in range(1, 55):
        tbl.cell(r, 0).text = f"R{r}"
        tbl.cell(r, 1).text = "EMEA"
        tbl.cell(r, 2).text = str(r * 100)
    d.save(str(big_tbl_docx))

    out_pptx = tmp_path / "table55.pptx"
    engine.conv_docx_to_pptx(big_tbl_docx, out_pptx)
    prs_out = pptx.Presentation(str(out_pptx))
    assert len(prs_out.slides) >= 3

    out_txt = tmp_path / "table55.txt"
    engine.conv_docx_to_text(big_tbl_docx, out_txt)
    txt_content = out_txt.read_text(encoding="utf-8-sig")
    assert "R54\tEMEA\t5400" in txt_content


def test_memory_gate_reentrant_and_csv_pdf_and_page_scale_routing(tmp_path):
    gate = engine._MemoryBudgetGate(budget_mb=512)
    # Nested/reentrant 150 MB acquisitions on the same thread must NEVER deadlock
    with gate.acquire(150 * 1024 * 1024):
        with gate.acquire(150 * 1024 * 1024):
            assert gate._in_use_mb > 0
    assert gate._in_use_mb == 0

    # Concurrent heavy acquisitions serialize so active usage never exceeds budget
    active = 0
    max_concurrent = 0
    lk = threading.Lock()

    def _heavy_worker():
        nonlocal active, max_concurrent
        with gate.acquire(120 * 1024 * 1024):  # cost = 512 MB (full budget)
            with lk:
                active += 1
                max_concurrent = max(max_concurrent, active)
            time.sleep(0.04)
            with lk:
                active -= 1

    threads = [threading.Thread(target=_heavy_worker) for _ in range(3)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=5.0)
    assert max_concurrent == 1

    # Page-aware pdf -> docx scale routing: <= 250 pages keeps pdf2docx first; > 250 pages demotes it
    fitz = engine._import_fitz()
    fake_routes = [
        engine.Route("pdf2docx", lambda s, d: None, 95),
        engine.Route("office-com", lambda s, d: None, 85),
        engine.Route("pymupdf-docx", lambda s, d: None, 75),
    ]
    short_pdf = tmp_path / "short.pdf"
    d_short = fitz.open()
    for _ in range(5):
        d_short.new_page()
    d_short.save(str(short_pdf))
    d_short.close()
    kept = engine._adjust_routes_for_scale(short_pdf, 80 * 1024 * 1024, ".docx", fake_routes)
    assert kept[0].name == "pdf2docx"

    long_pdf = tmp_path / "long300.pdf"
    d_long = fitz.open()
    for _ in range(260):
        d_long.new_page()
    d_long.save(str(long_pdf))
    d_long.close()
    demoted = engine._adjust_routes_for_scale(long_pdf, 80 * 1024 * 1024, ".docx", fake_routes)
    assert demoted[0].name == "office-com"
    assert demoted[-1].name == "pdf2docx"


def test_ipc_heartbeats_and_mid_job_reattach_no_duplicate_run(tmp_path, monkeypatch):
    monkeypatch.setattr(monitor, "DATA_DIR", tmp_path)
    monkeypatch.delenv("CLICKSHIFT_NO_IPC", raising=False)

    run_calls = 0

    class SlowFakeEngine:
        def convert_many(self, jobs, progress=None, step_progress=None):
            nonlocal run_calls
            run_calls += 1
            time.sleep(0.28)
            out = []
            for s, t in jobs:
                r = ConversionResult(s, s + "." + t, "slow-fake", 280.0, True, bytes_in=150_000_000, bytes_out=50_000)
                out.append(r)
                if progress:
                    progress(r)
            return out

    monkeypatch.setattr(engine, "get_engine", lambda **_kw: SlowFakeEngine())
    monkeypatch.setattr(monitor.DaemonIPCServer, "HEARTBEAT_INTERVAL_S", 0.06)

    srv = monitor.DaemonIPCServer({"ml_enabled": False})
    srv.start_and_publish()
    # Allow same-process test client by spoofing meta pid to parent/other alive pid
    meta = json.loads(monitor.ipc_file().read_text(encoding="utf-8"))
    monkeypatch.setattr(monitor, "_pid_alive", lambda _pid: True)
    monitor.ipc_file().write_text(
        json.dumps({"port": meta["port"], "token": meta["token"], "pid": os.getpid() + 99999}),
        encoding="utf-8",
    )

    try:
        # 1. End-to-end streaming over multiple heartbeat ticks runs job ONCE
        res = monitor.try_daemon_convert([("doc_150mb.pdf", "docx")], enable_ml=False)
        assert res is not None and len(res) == 1 and res[0].ok
        assert run_calls == 1

        # 2. Manual mid-job disconnect + attach resumes the SAME job without re-running!
        job_id = "reattach_test_job_01"
        s1 = socket.create_connection(("127.0.0.1", srv.port), timeout=2.0)
        r1 = s1.makefile("r", encoding="utf-8")
        w1 = s1.makefile("w", encoding="utf-8")
        w1.write(json.dumps({"token": srv.token, "job_id": job_id, "jobs": [("a.pdf", "docx")]}) + "\n")
        w1.flush()
        ack_line = json.loads(r1.readline())
        assert ack_line["type"] == "ack" and ack_line["job_id"] == job_id
        # Abruptly drop client socket while daemon worker is still converting!
        r1.close()
        w1.close()
        s1.close()

        # Re-attach on a fresh socket and verify we get the result with only 1 engine invocation
        s2 = socket.create_connection(("127.0.0.1", srv.port), timeout=2.0)
        with s2:
            r2 = s2.makefile("r", encoding="utf-8")
            w2 = s2.makefile("w", encoding="utf-8")
            w2.write(json.dumps({"token": srv.token, "action": "attach", "job_id": job_id}) + "\n")
            w2.flush()
            frames = [json.loads(ln) for ln in r2 if ln.strip()]
        result_frames = [f for f in frames if "ok" in f]
        assert len(result_frames) == 1 and result_frames[0]["ok"] is True
        assert run_calls == 2  # NOT 3!
    finally:
        srv.stop()


def test_watch_folder_filters_partial_downloads_and_zero_byte_files(tmp_path, monkeypatch):
    drop = tmp_path / "watch"
    drop.mkdir()
    (drop / "big.pdf.crdownload").write_bytes(b"partial chrome download")
    (drop / "movie.part").write_bytes(b"partial firefox")
    (drop / ".~clickshift-1234.docx").write_bytes(b"staging")
    (drop / "zero.txt").write_bytes(b"")
    (drop / "unknown.xyz123").write_bytes(b"unsupported")
    valid = drop / "ready.txt"
    valid.write_text("ready content", encoding="utf-8")

    converted: list[str] = []

    class FakeWatchEngine:
        def convert(self, src_p, dst_ext):
            converted.append(Path(src_p).name)
            out_p = Path(src_p).with_suffix("." + dst_ext.lstrip("."))
            out_p.write_text("<p>done</p>", encoding="utf-8")
            return ConversionResult(str(src_p), str(out_p), "text-passthru", 2.0, True)

    monkeypatch.setattr(engine, "get_engine", lambda **_kw: FakeWatchEngine())
    monkeypatch.setattr(monitor.WatchFolder, "POLL_SECONDS", 0.03)
    wf = monitor.WatchFolder(drop, {"auto_target": "html"})
    wf.start()
    try:
        deadline = time.monotonic() + 2.0
        while not converted and time.monotonic() < deadline:
            time.sleep(0.03)
        time.sleep(0.18)
    finally:
        wf.stop_event.set()
        wf.join(timeout=1.5)

    assert converted == ["ready.txt"]


def test_office_sta_worker_timeout_protects_user_open_office_pids(monkeypatch):
    killed_pids: list[tuple[int, str]] = []
    monkeypatch.setattr(engine, "_terminate_owned_pid", lambda pid, exe: killed_pids.append((pid, exe)))

    # Case A: Pre-existing PowerPoint PID (4242) must NEVER be marked owned or killed on timeout
    w_shared = engine._OfficeStaWorker("powerpoint", "PowerPoint.Application")
    w_shared._owned_pid = None  # attached to user's open PowerPoint
    w_shared._thread = threading.Thread(target=lambda: time.sleep(0.3), daemon=True)
    w_shared._thread.start()
    with pytest.raises(TimeoutError):
        w_shared.submit(lambda _app: None, timeout_s=0.05)
    assert killed_pids == []

    # Case B: Exclusively spawned Word PID (9876) IS terminated cleanly on timeout
    w_owned = engine._OfficeStaWorker("word", "Word.Application")
    w_owned._owned_pid = 9876
    w_owned._thread = threading.Thread(target=lambda: time.sleep(0.3), daemon=True)
    w_owned._thread.start()
    with pytest.raises(TimeoutError):
        w_owned.submit(lambda _app: None, timeout_s=0.05)
    assert killed_pids == [(9876, "winword.exe")]


def test_learning_150mb_tail_probe_normalized_reward_and_eta_overrun(tmp_path):
    # 1. EOF tail probe detects PDF trailer markers & OOXML central directory past 32 KB
    tail_pdf = tmp_path / "trailer.pdf"
    with open(tail_pdf, "wb") as fh:
        fh.write(b"%PDF-1.7\n" + (b"0" * 80_000))
        fh.write(b"\n/Type /Page /Subtype /Image /Subtype /Image /Encrypt 12 0 R\n%%EOF")
    feats = learning._features(tail_pdf, ".docx")
    assert len(feats) == learning.DIM
    assert feats[7] > 0.5   # media_ratio detected from EOF tail!
    assert feats[8] >= 0.3  # /Encrypt complexity detected from EOF tail!

    # 2. Bandit reward at 150 MB differentiates 2-min (120_000 ms) vs 10-min (600_000 ms) backends
    sel = learning.BanditSelector(history_file=tmp_path / "hist.jsonl")
    small_pdf = tmp_path / "tiny.pdf"
    small_pdf.write_bytes(b"%PDF-1.7 small")
    sel.record(small_pdf, ".docx", "fast", 250.0, True, src_size=100_000)

    size_150mb = 150 * 1024 * 1024
    for _ in range(10):
        sel.record(tail_pdf, ".docx", "fast_150", 90_000.0, True, src_size=size_150mb)
        sel.record(tail_pdf, ".docx", "slow_150", 600_000.0, True, src_size=size_150mb)

    routes = [
        engine.Route("slow_150", lambda s, d: None, 95),
        engine.Route("fast_150", lambda s, d: None, 70),
    ]
    ranked = sel.rank(tail_pdf, ".docx", routes, src_size=size_150mb)
    assert ranked[0].name == "fast_150"
    # Baseline median_ms remains normalized (< 15_000 ms), NOT polluted to 350_000 ms!
    assert sel.median_ms["pdf>docx"] < 15_000.0

    # 3. estimate_job_duration_ms on a 150 MB PDF never returns a tiny ~5s estimate
    fake_150_stat_pdf = tmp_path / "ref.pdf"
    fake_150_stat_pdf.write_bytes(b"%PDF")
    orig_stat = Path.stat

    def _mock_stat(self, *a, **kw):
        st = orig_stat(self, *a, **kw)
        if self == fake_150_stat_pdf:
            class _S:
                st_size = size_150mb
            return _S()  # type: ignore[return-value]
        return st

    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(Path, "stat", _mock_stat)
        est_ms = updater.estimate_job_duration_ms(fake_150_stat_pdf, "docx", median_map={"pdf>docx": 180.0})
    assert est_ms >= 90_000.0  # heuristic floor prevents 5s estimate!

    # 4. ConversionProgressPopup never freezes at 95% and remaining_s never counts upward on overrun
    pop = updater.ConversionProgressPopup([("huge.pdf", "docx")], estimated_total_ms=10_000.0, enabled=False).start()
    t0 = pop._t0
    snaps = [pop.compute_snapshot(now=t0 + dt) for dt in (2.0, 6.0, 9.5, 12.0, 20.0, 45.0, 90.0)]
    pcts = [s["pct"] for s in snaps]
    rems = [s["remaining_s"] for s in snaps]
    assert pcts == sorted(pcts)
    assert pcts[-1] > 95.0
    for earlier, later in zip(rems, rems[1:]):
        assert later <= earlier + 1e-6


def test_filelist_utf8_bom_and_debuglog_releases_live_exception_frames(tmp_path, monkeypatch):
    # 1. UTF-8 BOM in --filelist does not corrupt first file path
    first_file = tmp_path / "alpha.txt"
    first_file.write_text("ok", encoding="utf-8")
    fl = tmp_path / "batch_99_1.filelist"
    fl.write_bytes(b"\xef\xbb\xbf" + f"{first_file}\n".encode("utf-8"))
    loaded = cli._load_filelist(fl)
    assert loaded == [first_file]
    assert not fl.exists()  # batch_*.filelist auto-cleaned!

    # 2. Failed conversion attempt records formatted traceback string without pinning live exc object
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "onerror")
    monkeypatch.setattr(monitor, "DATA_DIR", tmp_path)
    captured_sessions: list[dict] = []
    orig_finalize = debuglog.finalize

    def _spy_finalize(sess, res):
        captured_sessions.append(sess)
        return orig_finalize(sess, res)

    monkeypatch.setattr(debuglog, "finalize", _spy_finalize)
    monkeypatch.setattr(engine.BACKENDS, "office_app", {})
    monkeypatch.setattr(engine.BACKENDS, "libreoffice", "")
    bad_src = tmp_path / "broken.pdf"
    bad_src.write_bytes(b"not a valid pdf")
    eng = engine.ConversionEngine(enable_ml=False)
    try:
        res = eng.convert(bad_src, "docx")
    finally:
        eng.shutdown()
    assert not res.ok
    assert captured_sessions
    for att in captured_sessions[0]["attempts"]:
        assert "exc" not in att
        assert isinstance(att.get("traceback"), str) and len(att["traceback"]) > 0


def test_multipage_pdf_html_text_mode_md_streaming_and_subprocess_job_timeout(tmp_path):
    fitz = engine._import_fitz()

    # 1. 75-page PDF (>60 pages) switches conv_pdf_to_html_fidelity to compact <text> SVG (not Bezier glyph bloat)
    many_pdf = tmp_path / "book75.pdf"
    doc = fitz.open()
    for pno in range(1, 76):
        pg = doc.new_page()
        pg.insert_text((72, 90), f"Heading Page {pno}", fontsize=18)
        pg.insert_text((72, 130), f"Body paragraph content on page {pno}.", fontsize=11)
    doc.save(str(many_pdf))
    doc.close()

    out_html = tmp_path / "book75.html"
    engine.conv_pdf_to_html_fidelity(many_pdf, out_html)
    html_txt = out_html.read_text(encoding="utf-8")
    assert 'id="page-75"' in html_txt
    assert "<text" in html_txt  # text_as_path=False kept selectable <text> nodes

    # 2. Structured PDF -> Markdown handles cross-page code merge in O(N)
    code_pdf = tmp_path / "code_pages.pdf"
    cdoc = fitz.open()
    for pno in range(1, 8):
        pg = cdoc.new_page()
        pg.insert_text((72, 80), f"int step_{pno} = {pno};", fontname="cour", fontsize=10)
    cdoc.save(str(code_pdf))
    cdoc.close()
    out_md = tmp_path / "code_pages.md"
    engine.conv_pdf_to_md_structured(code_pdf, out_md)
    md_txt = out_md.read_text(encoding="utf-8")
    assert "step_1" in md_txt and "step_7" in md_txt

    # 3. _run_subprocess_job_tree cleanly terminates timed-out child trees
    with pytest.raises(engine.subprocess.TimeoutExpired):
        engine._run_subprocess_job_tree([sys.executable, "-c", "import time; time.sleep(10)"], timeout=0.2)


def test_audit_regressions_ipc_zero_step_heartbeats_oserror_single_try_csv_pdf_and_utf16_progress(tmp_path, monkeypatch):
    # 1. IPC emits real heartbeat frames even when step_seq stays 0 (zero step_progress callbacks)
    monkeypatch.setattr(monitor, "DATA_DIR", tmp_path)
    monkeypatch.delenv("CLICKSHIFT_NO_IPC", raising=False)
    monkeypatch.setattr(monitor.DaemonIPCServer, "HEARTBEAT_INTERVAL_S", 0.05)

    class SilentBlockingEngine:
        def combine_images(self, files, target, step_progress=None):
            # Intentionally never calls step_progress so step_seq remains 0
            time.sleep(0.18)
            return ConversionResult(files[0], files[0] + ".pdf", "silent", 180.0, True)

    monkeypatch.setattr(engine, "get_engine", lambda **_kw: SilentBlockingEngine())
    srv = monitor.DaemonIPCServer({"ml_enabled": False})
    srv.start_and_publish()
    try:
        s = socket.create_connection(("127.0.0.1", srv.port), timeout=2.0)
        with s:
            r = s.makefile("r", encoding="utf-8")
            w = s.makefile("w", encoding="utf-8")
            w.write(
                json.dumps({"token": srv.token, "action": "combine_images", "files": ["img1.png"], "target": "pdf"})
                + "\n"
            )
            w.flush()
            frames = [json.loads(ln) for ln in r if ln.strip()]
        assert frames[0]["type"] == "ack"
        hb_frames = [f for f in frames if f.get("type") == "heartbeat"]
        assert len(hb_frames) >= 2
        assert any(f.get("ok") is True for f in frames)
    finally:
        srv.stop()

    # 2. A backend raising OSError (e.g., UnidentifiedImageError / FileNotFoundError) runs ONCE, never twice!
    oserror_calls = 0

    def _failing_oserror_route(_s, _d):
        nonlocal oserror_calls
        oserror_calls += 1
        raise OSError("simulated corrupt stream")

    def _fallback_ok_route(_s, d):
        Path(d).write_bytes(b"ok")

    src_dummy = tmp_path / "corrupt.png"
    src_dummy.write_bytes(b"bad")
    monkeypatch.setattr(
        engine,
        "_routes_for",
        lambda _se, _de: [
            engine.Route("pillow", _failing_oserror_route, 100),
            engine.Route("fallback", _fallback_ok_route, 50),
        ],
    )
    eng = engine.ConversionEngine(enable_ml=False)
    try:
        res = eng.convert(src_dummy, "jpg")
    finally:
        eng.shutdown()
    assert res.ok and res.backend == "fallback"
    assert oserror_calls == 1

    monkeypatch.undo()

    # 3. csv -> pdf and tsv -> pdf route cleanly on Word-only vs Excel-installed machines
    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(engine, "IS_WINDOWS", True)
        mp.setattr(engine.BACKENDS, "office_app", {"word": "Word.Application"})
        mp.setattr(engine.BACKENDS, "libreoffice", "soffice")
        csv_routes_word_only = [r.name for r in engine._routes_for("csv", "pdf")]
        tsv_routes_word_only = [r.name for r in engine._routes_for("tsv", "pdf")]
        assert csv_routes_word_only[0] == "libreoffice"
        assert tsv_routes_word_only[0] == "libreoffice"

        mp.setattr(engine.BACKENDS, "office_app", {"word": "Word.Application", "excel": "Excel.Application"})
        csv_routes_excel = [r.name for r in engine._routes_for("csv", "pdf")]
        tsv_routes_excel = [r.name for r in engine._routes_for("tsv", "pdf")]
        assert csv_routes_excel[0] == "office-com"
        assert tsv_routes_excel[0] == "office-com"

    # 4. >8 MiB UTF-16 streaming text conversion reports accurate byte progress up to 100%
    utf16_src = tmp_path / "wide_log.txt"
    utf16_payload = ("Line 0123456789 UTF-16 progress check\r\n" * 130_000).encode("utf-16")
    assert len(utf16_payload) > 9 * 1024 * 1024
    utf16_src.write_bytes(utf16_payload)
    utf16_out = tmp_path / "wide_log.html"
    prog_fracs: list[float] = []
    with engine.subprogress_scope(lambda d, t, _m: prog_fracs.append(d / float(t))):
        engine.conv_text_generic(utf16_src, utf16_out)
    assert prog_fracs and prog_fracs[-1] >= 0.98


