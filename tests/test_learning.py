import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
os.environ.setdefault("CLICKSHIFT_HOME", str(Path(__file__).parent / "_tmp_home"))

import numpy as np  # noqa: E402
from clickshift.learning import BanditSelector, _features, DIM  # noqa: E402
from clickshift.engine import Route  # noqa: E402


def make_routes():
    return [Route("fast", lambda s, d: None, 90), Route("slow", lambda s, d: None, 60)]


def test_features_dim():
    p = Path(__file__)
    f = _features(p, ".docx")
    assert len(f) == DIM


def test_bandit_learns_faster_backend(tmp_path):
    sel = BanditSelector(history_file=tmp_path / "h.jsonl")
    src = tmp_path / "doc.pdf"
    src.write_bytes(b"%PDF-1.7 fake")
    routes = make_routes()

    # simulate: 'fast' always ~100ms, 'slow' always ~900ms
    for i in range(24):
        ranked = sel.rank(src, ".docx", routes)
        chosen = ranked[0]
        ms = 100 if chosen.name == "fast" else 900
        sel.record(src, ".docx", chosen.name, ms, True)

    final = sel.rank(src, ".docx", routes)
    assert final[0].name == "fast", "bandit should converge to the faster backend"


def test_state_persists(tmp_path):
    sel = BanditSelector(history_file=tmp_path / "h.jsonl")
    src = tmp_path / "d.pdf"
    src.write_bytes(b"x")
    sel.record(src, ".docx", "fast", 120, True)
    sel2 = BanditSelector(history_file=tmp_path / "h.jsonl")
    assert sum(sel2.bandit.n.values()) >= 1


def test_content_aware_features_distinguish_scanned_vs_text_pdf(tmp_path):
    text_pdf = tmp_path / "text.pdf"
    scan_pdf = tmp_path / "scan.pdf"
    text_pdf.write_bytes(b"%PDF-1.7\n/Type /Page\n/Font /F1\n/Font /F2\n" + b"a" * 500)
    scan_pdf.write_bytes(b"%PDF-1.7\n/Type /Page\n/Subtype /Image\n/Subtype /Image\n" + b"a" * 500)
    f_text = _features(text_pdf, ".docx")
    f_scan = _features(scan_pdf, ".docx")
    assert len(f_text) == DIM and len(f_scan) == DIM
    # media_ratio (index 7) is higher for the image-heavy PDF than the font-only PDF
    assert f_scan[7] > f_text[7]


def test_sherman_morrison_matches_exact_inverse(tmp_path):
    sel = BanditSelector(history_file=tmp_path / "h.jsonl")
    rng = np.random.default_rng(42)
    arm = "pdf>docx|pdf2docx"
    for _ in range(15):
        x = rng.normal(size=DIM)
        x[0] = 1.0
        sel.bandit.update(arm, x, float(rng.uniform(-1.0, 1.0)))
    exact_inv = np.linalg.inv(sel.bandit.A[arm])
    assert np.allclose(sel.bandit.A_inv[arm], exact_inv, atol=1e-6)


def test_batch_persistence_and_history_tail_compaction(tmp_path, monkeypatch):
    from clickshift import learning as learning_mod

    src = tmp_path / "sample.txt"
    src.write_text("hello world", encoding="utf-8")
    sel = learning_mod.BanditSelector(history_file=tmp_path / "h.jsonl")

    saves = 0
    orig_save = sel._save_state

    def counted_save():
        nonlocal saves
        saves += 1
        orig_save()

    monkeypatch.setattr(sel, "_save_state", counted_save)
    with sel.batch_persistence():
        for i in range(5):
            sel.record(src, ".html", "text-passthru", 10.0 + i, True)
        assert saves == 0
    assert saves == 1
    assert len(sel.history_file.read_text(encoding="utf-8").splitlines()) == 5

    # Synthetic >1 MB history triggers binary tail compaction keeping <= 3000 valid lines
    big_line = '{"pair":"txt>html","arm":"txt>html|text-passthru","ms":5.0,"ok":true,"reward":1.0,"file":"abc.txt","t":1}\n'
    repeat = (1_050_000 // len(big_line)) + 10
    sel.history_file.write_text(big_line * repeat, encoding="utf-8")
    sel.record(src, ".html", "text-passthru", 12.0, True)
    kept_lines = sel.history_file.read_text(encoding="utf-8").splitlines()
    assert 1 <= len(kept_lines) <= 3000
    assert kept_lines[-1].startswith("{") and kept_lines[-1].endswith("}")

