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


