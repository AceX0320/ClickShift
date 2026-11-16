"""Cross-platform unit tests for the ConFas engine (run anywhere)."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
os.environ["CONFAS_HOME"] = str(Path(__file__).parent / "_tmp_home")

import pytest  # noqa: E402
from confas.engine import ConversionEngine, newest_target_for, _routes_for  # noqa: E402


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
    from confas.engine import _flatten_path_opacity
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
    from confas.engine import _merge_shaded_line_tables
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
    from confas.monitor import menu_targets

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
    from confas.engine import conv_pptx_to_pdf

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
    from confas.monitor import menu_targets

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




