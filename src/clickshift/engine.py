"""ClickShift conversion engine.

Design goals: SPEED + PERFECTION (lossless-fidelity conversions).

Strategy (fastest path first, newest format versions by default):

  * DOCX / XLSX / PPTX -> PDF : Microsoft Word/Excel/PowerPoint COM automation
        (the installed Office writes native OOXML; fastest & pixel-perfect on
        Windows machines that have Office).  Fallback: LibreOffice headless.
  * PDF -> DOCX              : pdf2docx (preserves layout, tables, images),
        then Microsoft Word PDF-reflow (Office COM), then LibreOffice writer
        filter ``MS Word 2007 XML`` (= .docx, OOXML 2007+ - NOT legacy .doc).
  * Legacy Office (.doc/.xls/.ppt) -> modern (.docx/.xlsx/.pptx) :
        Office COM with FileFormat codes 12 / 51 / 24, or LibreOffice filters
        ``MS Word 2007 XML`` / ``MS Excel 2007 XML`` / ``MS PowerPoint 2007 XML``.
  * Modern -> legacy is NEVER produced unless explicitly requested.
  * Images                     : Pillow (C-accelerated) re-encoded to the
        newest codec generation (JPEG XL > WebP > AVIF/HEIF > JPEG/PNG...).
  * Text/code/markup           : chardet auto-encoding -> UTF-8 (+ BOM for
        Windows-native UTF-16 consumers), HTML <-> DOCX via python-docx,
        Markdown -> DOCX/HTML/PDF.
  * EPUB                       : epub -> text/HTML via stdlib zipfile.

Everything runs in a persistent worker process pool, so consecutive
conversions are near-instant (no interpreter/COM spin-up per file).

The module is import-safe on Linux/macOS as well (Windows-only backends are
simply unavailable there) which makes development/testing possible anywhere.
"""
# pyright: reportMissingImports=false, reportMissingModuleSource=false
from __future__ import annotations

import contextlib
import html as _html_mod
import importlib
import io
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
import zipfile
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, Optional, Sequence
from xml.etree import ElementTree as ET

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    __package__ = "clickshift"

from .formats import (
    IMAGE_EXTS,
    NEWEST_GENERATION,
    OFFICE_DOC_IN,
    OFFICE_SHEET_IN,
    OFFICE_SLIDE_IN,
    PAGED_IN,
    PAGE_IMAGE_EXTS,
    PDF_CAPABLE_IN,
    SUPPORTED_INPUT_EXTS,
    TEXT_EXTS,
    ConversionResult,
    _norm_ext,
    all_targets_for,
    is_pages_to_images,
    newest_target_for,
)

IS_WINDOWS = sys.platform == "win32"


def _opt_import(name: str) -> Any:
    try:
        return importlib.import_module(name)
    except Exception:  # pragma: no cover
        return None


def _import_fitz() -> Any:
    """Import PyMuPDF via modern ``pymupdf`` first to avoid deprecation warnings."""
    try:
        return importlib.import_module("pymupdf")
    except Exception:
        return importlib.import_module("fitz")


winreg: Any = _opt_import("winreg") if IS_WINDOWS else None


def _ensure_stdio() -> None:
    """In PyInstaller --noconsole mode sys.stdout/stderr are None, breaking pdf2docx/logging."""
    try:
        if sys.stdout is None:
            sys.stdout = open(os.devnull, "w", encoding="utf-8", errors="replace")
        if sys.stderr is None:
            sys.stderr = open(os.devnull, "w", encoding="utf-8", errors="replace")
    except Exception:
        pass


_ensure_stdio()


# --------------------------------------------------------------------------- #
#  Backend discovery (lazy per-attribute probing; zero cold import overhead)
# --------------------------------------------------------------------------- #

_BACKEND_ATTRS = (
    "libreoffice",
    "office_app",
    "pillow",
    "pdf2docx",
    "pymupdf",
    "pypdf",
    "python_docx",
    "python_pptx",
    "chardet",
)
_MOD_PROBES = {
    "pillow": ("PIL",),
    "pymupdf": ("pymupdf", "fitz"),
    "pypdf": ("pypdf",),
    "python_docx": ("docx",),
    "python_pptx": ("pptx",),
    "chardet": ("chardet",),
}
_PDF2DOCX_PATCH_LOCK = threading.Lock()
_PDF2DOCX_PATCHED = False


class Backends:
    def __init__(self) -> None:
        self._probe_lock = threading.Lock()

    def _probe_libreoffice(self) -> Optional[str]:
        if IS_WINDOWS:
            candidates = (
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
            )
            for c in candidates:
                if Path(c).exists():
                    return c
            return shutil.which("soffice")
        return shutil.which("soffice") or shutil.which("libreoffice")

    def _probe_office_app(self) -> dict[str, str]:
        apps: dict[str, str] = {}
        if IS_WINDOWS and winreg is not None:
            try:
                for app, progid in (
                    ("word", "Word.Application"),
                    ("excel", "Excel.Application"),
                    ("powerpoint", "PowerPoint.Application"),
                ):
                    try:
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                            apps[app] = progid
                    except OSError:
                        continue
            except Exception:
                pass
        return apps

    def _probe_attr(self, name: str) -> Any:
        if name == "libreoffice":
            return self._probe_libreoffice()
        if name == "office_app":
            return self._probe_office_app()
        if name in _MOD_PROBES:
            for mod_name in _MOD_PROBES[name]:
                try:
                    importlib.import_module(mod_name)
                    return True
                except Exception:
                    continue
            return False
        if name == "pdf2docx":
            try:
                _patch_pdf2docx_compat()
                importlib.import_module("pdf2docx")
                return True
            except Exception:
                return False
        raise AttributeError(name)

    def __getattr__(self, name: str) -> Any:
        if name not in _BACKEND_ATTRS:
            raise AttributeError(f"{type(self).__name__!s} has no attribute {name!r}")
        with self._probe_lock:
            if name in self.__dict__:
                return self.__dict__[name]
            val = self._probe_attr(name)
            self.__dict__[name] = val
            return val

    def detect(self) -> None:
        with self._probe_lock:
            for name in _BACKEND_ATTRS:
                self.__dict__[name] = self._probe_attr(name)

    def summary(self) -> dict:
        return {
            "libreoffice": self.libreoffice,
            "office_apps": dict(self.office_app),
            "pillow": self.pillow,
            "pdf2docx": self.pdf2docx,
            "pymupdf": self.pymupdf,
            "pypdf": self.pypdf,
            "python_docx": self.python_docx,
            "python_pptx": self.python_pptx,
            "chardet": self.chardet,
        }


def _patch_pdf2docx_compat() -> None:
    """Fix upstream PyMuPDF >= 1.24.3 and NumPy >= 2.0 breaking changes in pdf2docx 0.5.8."""
    global _PDF2DOCX_PATCHED
    if _PDF2DOCX_PATCHED:
        return
    with _PDF2DOCX_PATCH_LOCK:
        if _PDF2DOCX_PATCHED:
            return
        try:
            np: Any = importlib.import_module("numpy")
            for alias, target in (
                ("float_", np.float64),
                ("int_", np.int64),
                ("int0", np.intp),
                ("bool8", np.bool_),
                ("Inf", np.inf),
                ("NaN", np.nan),
            ):
                if not hasattr(np, alias):
                    setattr(np, alias, target)
        except Exception:
            pass
        try:
            fitz: Any = _import_fitz()
            # PyMuPDF 1.24.3+ removed Rect.get_area() / IRect.get_area() which pdf2docx calls
            for cls_name in ("Rect", "IRect"):
                cls = getattr(fitz, cls_name, None)
                if cls is not None and not hasattr(cls, "get_area"):
                    setattr(cls, "get_area", lambda self: abs(self))
        except Exception:
            pass
        _patch_pdf2docx_opacity()
        _PDF2DOCX_PATCHED = True


def _blend_on_white(color: Any, opacity: Any) -> Any:
    """Composite an RGB(0..1) colour with ``opacity`` over a white page."""
    if not color or opacity is None:
        return color
    try:
        op = float(opacity)
    except (TypeError, ValueError):
        return color
    if op >= 0.999:
        return color
    op = max(0.0, op)
    return tuple(1.0 - op * (1.0 - float(c)) for c in color)


def _flatten_path_opacity(raw_paths: Any) -> Any:
    """pdf2docx ignores fill/stroke opacity, so a 10%-opaque dark box (a common
    way to draw light-grey panels) becomes solid charcoal cell shading in
    Word. Pre-blend translucent colours over white and drop invisible paths."""
    out = []
    for p in raw_paths or []:
        if not isinstance(p, dict):
            out.append(p)
            continue
        has_fill = p.get("fill") is not None
        has_color = p.get("color") is not None
        fop = p.get("fill_opacity")
        sop = p.get("stroke_opacity")
        fill_visible = has_fill and (fop is None or fop > 0.01)
        stroke_visible = has_color and (sop is None or sop > 0.01)
        if (has_fill or has_color) and not (fill_visible or stroke_visible):
            continue  # fully transparent path: nothing to render
        needs_blend = (has_fill and fop is not None and fop < 0.999) or (has_color and sop is not None and sop < 0.999)
        if not needs_blend:
            out.append(p)
            continue
        p = dict(p)
        if has_fill:
            p["fill"] = _blend_on_white(p["fill"], fop) if fill_visible else None
            p["fill_opacity"] = 1.0
        if has_color:
            p["color"] = _blend_on_white(p["color"], sop) if stroke_visible else None
            p["stroke_opacity"] = 1.0
        out.append(p)
    return out


def _patch_pdf2docx_opacity() -> None:
    try:
        mod: Any = importlib.import_module("pdf2docx.page.RawPageFitz")
        paths_mod: Any = importlib.import_module("pdf2docx.shape.Paths")
    except Exception:
        return
    cls = getattr(mod, "RawPageFitz", None)
    if cls is None or getattr(cls, "_clickshift_opacity_patched", False):
        return

    def _init_paths(self: Any, **settings: Any) -> Any:
        raw_paths = _flatten_path_opacity(self.page_engine.get_cdrawings())
        return paths_mod.Paths(parent=self).restore(raw_paths)

    cls._init_paths = _init_paths
    cls._clickshift_opacity_patched = True


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MAX_MERGE_GAP_TWIPS = 300  # spacer paragraphs up to 15pt are part of one box


def _merge_shaded_line_tables_with_check(xml: bytes) -> tuple[bytes, int, bool]:
    """Merge runs of per-line shaded tables and check document non-emptiness in one parse."""
    etree: Any = importlib.import_module("lxml.etree")
    w = "{%s}" % _W_NS
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
    root = etree.fromstring(xml, parser=parser)

    def ival(el: Any, attr: str = "w") -> int:
        if el is None:
            return 0
        try:
            return int(float(el.get(w + attr) or 0))
        except ValueError:
            return 0

    def single_cell_fill(tbl: Any) -> Optional[str]:
        rows = tbl.findall(w + "tr")
        if len(rows) != 1:
            return None
        cells = rows[0].findall(w + "tc")
        if len(cells) != 1:
            return None
        shd = cells[0].find(f"{w}tcPr/{w}shd")
        fill = (shd.get(w + "fill") if shd is not None else None) or ""
        if not fill or fill.lower() == "auto":
            return None
        return fill.lower()

    def small_spacer(p: Any) -> bool:
        if p.tag != w + "p":
            return False
        if any((t.text or "").strip() for t in p.iter(w + "t")):
            return False
        if p.find(f".//{w}drawing") is not None or p.find(f".//{w}pict") is not None:
            return False
        sp = p.find(f"{w}pPr/{w}spacing")
        if sp is None or sp.get(w + "lineRule") != "exact":
            return False
        total = ival(sp, "line") + ival(sp, "before") + ival(sp, "after")
        return total <= _MAX_MERGE_GAP_TWIPS

    def geometry(tbl: Any) -> tuple[int, int]:
        ind = ival(tbl.find(f"{w}tblPr/{w}tblInd"))
        tcw = ival(tbl.find(f"{w}tr/{w}tc/{w}tcPr/{w}tcW"))
        return ind, tcw

    merged = 0
    for body in list(root.iter(w + "body", w + "tc", w + "txbxContent")):
        children = list(body)
        i = 0
        while i < len(children):
            el = children[i]
            fill = single_cell_fill(el) if el.tag == w + "tbl" else None
            if not fill:
                i += 1
                continue
            group = [el]
            spacers: list[list[Any]] = []
            j = i + 1
            while j < len(children):
                pending = []
                k = j
                while k < len(children) and small_spacer(children[k]):
                    pending.append(children[k])
                    k += 1
                if k < len(children) and children[k].tag == w + "tbl" and single_cell_fill(children[k]) == fill:
                    spacers.append(pending)
                    group.append(children[k])
                    j = k + 1
                else:
                    break
            if len(group) < 2:
                i += 1
                continue
            geos = [geometry(t) for t in group]
            left = min(g[0] for g in geos)
            right = max(g[0] + g[1] for g in geos)
            first = group[0]
            cell = first.find(f"{w}tr/{w}tc")
            # first table: anchor at the leftmost edge, widen, free the row height
            for idx, (tbl, (ind, _)) in enumerate(zip(group, geos)):
                shift = ind - left
                if shift:
                    for pind in tbl.iter(w + "ind"):
                        pind.set(w + "left", str(ival(pind, "left") + shift))
                if idx == 0:
                    continue
                for sp in spacers[idx - 1]:
                    cell.append(sp)
                src_cell = tbl.find(f"{w}tr/{w}tc")
                for child in list(src_cell):
                    if child.tag != w + "tcPr":
                        cell.append(child)
                tbl.getparent().remove(tbl)
            tbl_ind = first.find(f"{w}tblPr/{w}tblInd")
            if tbl_ind is not None:
                tbl_ind.set(w + "w", str(left))
            tcw = cell.find(f"{w}tcPr/{w}tcW")
            if tcw is not None:
                tcw.set(w + "w", str(right - left))
            for gc in first.iter(w + "gridCol"):
                gc.set(w + "w", str(max(ival(gc), right - left)))
            for trh in first.iter(w + "trHeight"):
                trh.getparent().remove(trh)
            merged += len(group) - 1
            i = j
    has_content = (
        root.find(f".//{w}tbl") is not None
        or root.find(f".//{w}drawing") is not None
        or root.find(f".//{w}pict") is not None
        or any((t.text or "").strip() for t in root.iter(w + "t"))
    )
    if not merged:
        return xml, 0, has_content
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), merged, has_content


def _merge_shaded_line_tables(xml: bytes) -> tuple[bytes, int]:
    """Merge runs of per-line shaded tables; return (xml, merged_count)."""
    out_xml, merged, _ = _merge_shaded_line_tables_with_check(xml)
    return out_xml, merged


_MAX_ZIP_ENTRY_BYTES = 50 * 1024 * 1024
_MAX_ZIP_TOTAL_BYTES = 500 * 1024 * 1024
_MAX_ZIP_ENTRIES = 5000


def _zip_read_bounded(z: zipfile.ZipFile, name: str, max_bytes: int = _MAX_ZIP_ENTRY_BYTES) -> bytes:
    """Safely read a ZIP member up to ``max_bytes`` (guards against zip bombs)."""
    info = z.getinfo(name)
    if info.file_size > max_bytes:
        raise ValueError(f"ZIP entry '{name}' exceeds size limit ({info.file_size} > {max_bytes})")
    chunks: list[bytes] = []
    total = 0
    with z.open(info, "r") as stream:
        while True:
            chunk = stream.read(min(1024 * 1024, (max_bytes + 1) - total))
            if not chunk:
                break
            total += len(chunk)
            if total > max_bytes:
                raise ValueError(f"ZIP entry '{name}' exceeded decompressed size cap ({max_bytes} bytes)")
            chunks.append(chunk)
    return b"".join(chunks)


def _safe_xml_fromstring(data: bytes) -> ET.Element:
    """Parse XML bytes rejecting DTD/ENTITY declarations."""
    if b"<!ENTITY" in data[:16384].upper() or b"<!DOCTYPE" in data[:16384].upper():
        raise ValueError("Refusing XML with DTD/ENTITY declarations")
    return ET.fromstring(data)


def _postprocess_pdf2docx(dst: Path) -> Optional[bool]:
    """Rewrite word/document.xml with layout clean-ups; return has_content if checked."""
    tmp_name: Optional[str] = None
    try:
        with zipfile.ZipFile(dst) as zin:
            xml = _zip_read_bounded(zin, "word/document.xml", _MAX_ZIP_ENTRY_BYTES)
            new_xml, merged, has_content = _merge_shaded_line_tables_with_check(xml)
            if not merged:
                return has_content
            fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=str(dst.parent))
            os.close(fd)
            with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = (
                        new_xml
                        if item.filename == "word/document.xml"
                        else _zip_read_bounded(zin, item.filename, _MAX_ZIP_ENTRY_BYTES * 2)
                    )
                    zout.writestr(item, data)
        os.replace(tmp_name, str(dst))
        tmp_name = None
        return has_content
    except Exception:  # never fail a conversion because of cosmetics
        return None
    finally:
        if tmp_name:
            try:
                os.unlink(tmp_name)
            except OSError:
                pass


BACKENDS = Backends()


# --------------------------------------------------------------------------- #
#  Backend implementations
# --------------------------------------------------------------------------- #

BackendFn = Callable[[Path, Path], None]


def _unique(path: Path) -> Path:
    """Return ``path`` if free, else name_1.ext, name_2.ext ..."""
    if not path.exists():
        return path
    stem, suf = path.stem, path.suffix
    for i in range(1, 500):
        cand = path.with_name(f"{stem}_{i}{suf}")
        if not cand.exists():
            return cand
    return path.with_name(f"{stem}_{uuid.uuid4().hex[:6]}{suf}")


# ---- LibreOffice (universal fallback, batch-friendly) ---------------------- #

_LO_FILTERS = {
    "pdf": "writer_pdf_Export",
    "docx": "MS Word 2007 XML",          # OOXML 2007+ (NOT 'MS Word 97')
    "xlsx": "MS Excel 2007 XML",
    "pptx": "MS PowerPoint 2007 XML",
    "doc": "MS Word 97",
    "xls": "MS Excel 97",
    "ppt": "MS PowerPoint 97",
    "odt": "writer8",
    "ods": "calc8",
    "odp": "impress8",
    "rtf": "Rich Text Format",
    "html": "HTML (StarWriter)",
    "txt": "Text",
    "csv": "Text - txt - csv (StarCalc)",
    "png": "draw_png_Export",
    "jpg": "draw_jpeg_Export",
    "svg": "draw_svg_Export",
}


_LO_XCU_HARDENED = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
    'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
    '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
    '<prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse">'
    '<value>true</value></prop></item>\n'
    '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
    '<prop oor:name="DisableMacrosExecution" oor:op="fuse">'
    '<value>true</value></prop></item>\n'
    '</oor:items>\n'
)


_LO_SEM = threading.Semaphore(max(2, min(4, (os.cpu_count() or 4) // 2 or 2)))


def _lo_infilter_for(src_ext: str, dst_ext: str) -> Optional[str]:
    # Crucial for PDF -> DOCX/ODT/RTF/TXT/HTML: force LibreOffice Writer import
    # instead of Draw (which rejects MS Word 2007 XML export with error 207).
    if src_ext == "pdf" and dst_ext in ("docx", "doc", "odt", "rtf", "txt", "html", "htm"):
        return "--infilter=writer_pdf_import"
    if src_ext == "pdf" and dst_ext in ("pptx", "ppt", "odp"):
        return "--infilter=impress_pdf_import"
    return None


def _lo_build_base_args(prof_uri: str, dst_ext: str, infilter: Optional[str], extra: Optional[list[str]]) -> list[str]:
    filt = _LO_FILTERS.get(dst_ext)
    args = [
        BACKENDS.libreoffice, "--headless", "--norestore", "--nologo", "--nofirststartwizard",
        f"-env:UserInstallation={prof_uri}",
    ]
    if infilter:
        args.append(infilter)
    conv_spec = f"{dst_ext}:{filt}" if filt else dst_ext
    if extra:
        conv_spec += ": " + ", ".join(extra)
    args += ["--convert-to", conv_spec]
    return args


def lo_convert(src: Path, dst: Path, extra: Optional[list[str]] = None) -> None:
    if not BACKENDS.libreoffice:
        raise RuntimeError("LibreOffice not found")
    outdir = dst.parent
    src_resolved = Path(src).resolve()
    dst_ext = dst.suffix.lstrip(".").lower()
    src_ext = src_resolved.suffix.lstrip(".").lower()
    infilter = _lo_infilter_for(src_ext, dst_ext)
    with _LO_SEM, tempfile.TemporaryDirectory(prefix="clickshift_lo_") as prof_tmp:
        prof_dir = Path(prof_tmp)
        user_dir = prof_dir / "user"
        user_dir.mkdir(parents=True, exist_ok=True)
        (user_dir / "registrymodifications.xcu").write_text(_LO_XCU_HARDENED, encoding="utf-8")
        prof_uri = prof_dir.resolve().as_uri()
        args = _lo_build_base_args(prof_uri, dst_ext, infilter, extra)
        args += ["--outdir", str(outdir), str(src_resolved)]
        proc = subprocess.run(args, capture_output=True, timeout=300)
    produced = outdir / (src_resolved.stem + dst.suffix)
    if produced.exists() and produced != dst:
        produced.replace(dst)
    if not dst.exists():
        raise RuntimeError(f"LibreOffice produced no output ({proc.returncode}): "
                           f"{proc.stderr.decode(errors='replace')[:400]}")


def _lo_convert_batch(
    pairs: list[tuple[Path, Path]],
    dst_ext: str,
    infilter: Optional[str] = None,
) -> tuple[dict[Path, bool], str]:
    """Convert multiple distinct-stem files in one soffice invocation into a staging dir.

    Returns ``(succeeded_map, err_msg)`` keyed by source ``Path``.
    """
    if not BACKENDS.libreoffice or not pairs:
        return {s: False for s, _ in pairs}, "LibreOffice not found"
    ext_dot = "." + dst_ext.lstrip(".")
    ok_map: dict[Path, bool] = {s: False for s, _ in pairs}
    err_msg = ""
    with _LO_SEM, tempfile.TemporaryDirectory(prefix="clickshift_lobatch_") as prof_tmp:
        root_tmp = Path(prof_tmp)
        user_dir = root_tmp / "user"
        stage_out = root_tmp / "out"
        user_dir.mkdir(parents=True, exist_ok=True)
        stage_out.mkdir(parents=True, exist_ok=True)
        (user_dir / "registrymodifications.xcu").write_text(_LO_XCU_HARDENED, encoding="utf-8")
        prof_uri = (root_tmp / "prof").resolve().as_uri()
        args = _lo_build_base_args(prof_uri, dst_ext, infilter, None)
        args += ["--outdir", str(stage_out)] + [str(s.resolve()) for s, _ in pairs]
        try:
            proc = subprocess.run(args, capture_output=True, timeout=max(300, 60 * len(pairs)))
            if proc.returncode != 0:
                err_msg = proc.stderr.decode(errors="replace")[:400]
        except Exception as exc:
            return ok_map, str(exc)
        for src_p, dst_p in pairs:
            produced = stage_out / (src_p.stem + ext_dot)
            if produced.exists() and produced.stat().st_size > 0:
                dst_p.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(produced), str(dst_p))
                ok_map[src_p] = dst_p.exists()
    return ok_map, err_msg


def conv_libreoffice(src: Path, dst: Path) -> None:
    lo_convert(src, dst)


# ---- Microsoft Office COM (fastest & most faithful when Office is present) - #

_FILEFORMAT = {  # https://learn.microsoft.com/office/vba/api/wdwdformat etc.
    "docx": 12,   # wdFormatXMLDocument  (OOXML, Word 2007+)
    "pdf": 17,    # wdFormatPDF
    "xlsx": 51,   # xlOpenXMLWorkbook
    "pptx": 24,   # ppSaveAsOpenXMLPresentation
    "mht": 10,
    "html": 8,    # wdFormatHTML  (Word only)
    "rtf": 6,     # wdFormatRTF   (Word only)
    "txt": 2,     # wdFormatText  (Word only)
}
_WORD_ONLY_FMTS = {"html", "rtf", "txt"}


class _OfficeStaWorker:
    """Dedicated STA thread keeping one Office COM application warm across files."""

    IDLE_TIMEOUT_S = 45.0

    def __init__(self, app_name: str, progid: str) -> None:
        self.app_name = app_name
        self.progid = progid
        self._q: "queue.Queue" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.Lock()

    def _ensure_thread(self) -> None:
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(
                    target=self._loop, daemon=True, name=f"clickshift-com-{self.app_name}"
                )
                self._thread.start()

    def submit(self, fn: Callable[[Any], None]) -> None:
        self._ensure_thread()
        res_q: "queue.Queue[tuple[bool, Any]]" = queue.Queue(maxsize=1)
        self._q.put((fn, res_q))
        ok, payload = res_q.get()
        if not ok:
            raise payload

    def shutdown(self) -> None:
        with self._lock:
            if self._thread and self._thread.is_alive():
                self._q.put(None)

    def _spawn_app(self) -> Any:
        win32_client: Any = importlib.import_module("win32com.client")
        app = win32_client.DispatchEx(self.progid)
        for attr, val in (("Visible", False), ("DisplayAlerts", 0), ("ScreenUpdating", False)):
            try:
                setattr(app, attr, val)
            except Exception:
                pass
        if _office_disable_macros_enabled():
            try:
                app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
        return app

    def _quit_app(self, app: Any) -> None:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    @staticmethod
    def _is_rpc_disconnect(exc: BaseException) -> bool:
        hr = getattr(exc, "hresult", None)
        if hr in (-2147023174, -2147417848, -2147418111, -2147220995, -2147023170):
            return True
        msg = str(exc).upper()
        return "RPC" in msg or "SERVER_UNAVAILABLE" in msg or "DISCONNECTED" in msg

    def _loop(self) -> None:
        pythoncom: Any = importlib.import_module("pythoncom")
        pythoncom.CoInitialize()
        app = None
        try:
            while True:
                try:
                    item = self._q.get(timeout=self.IDLE_TIMEOUT_S)
                except queue.Empty:
                    # Idle timeout expired -> release COM server until next batch
                    self._quit_app(app)
                    app = None
                    continue
                if item is None:
                    break
                fn, res_q = item
                try:
                    if app is None:
                        app = self._spawn_app()
                    try:
                        fn(app)
                    except Exception as first_exc:
                        if not self._is_rpc_disconnect(first_exc):
                            raise
                        # Retry once with a fresh COM process only if RPC server died
                        self._quit_app(app)
                        app = self._spawn_app()
                        fn(app)
                    res_q.put((True, None))
                except Exception as exc:
                    if self._is_rpc_disconnect(exc):
                        self._quit_app(app)
                        app = None
                    res_q.put((False, exc))
        finally:
            self._quit_app(app)
            pythoncom.CoUninitialize()


class _OfficeComPool:
    def __init__(self) -> None:
        self._workers: dict[str, _OfficeStaWorker] = {}
        self._lock = threading.Lock()

    def run(self, app_name: str, progid: str, fn: Callable[[Any], None]) -> None:
        with self._lock:
            w = self._workers.get(app_name)
            if w is None:
                w = _OfficeStaWorker(app_name, progid)
                self._workers[app_name] = w
        w.submit(fn)

    def shutdown(self) -> None:
        with self._lock:
            workers = list(self._workers.values())
            self._workers.clear()
        for w in workers:
            w.shutdown()


_OFFICE_POOL = _OfficeComPool()


_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_RTF_UNSAFE_RE = re.compile(
    br"(?is)(?:\\\*\\template\b|\\template\b|\bDDEAUTO\b|\bINCLUDETEXT\b)"
)
_OFFICE_ALLOWED_DST = set(_FILEFORMAT) | {"doc", "xls", "ppt", "odt", "ods", "odp", "csv"}


def _office_disable_macros_enabled() -> bool:
    try:
        from . import monitor
        return bool(monitor.load_config().get("office_disable_macros", True))
    except Exception:
        return True


def _is_unsafe_external_rel(rel_type: str, target: str) -> bool:
    t = (target or "").strip()
    tl = t.lower()
    if t.startswith(("\\\\", "//")) or tl.startswith("file:"):
        return True
    return not rel_type.lower().endswith("/hyperlink")


def _sanitize_ooxml_for_com(src: Path, tmp_dir: Any) -> Path:
    """Strip external templates/links from OOXML `.rels` before opening in Office COM."""
    if src.suffix.lower() not in (".docx", ".xlsx", ".pptx"):
        return src
    rewrites: dict[str, bytes] = {}
    try:
        zin = zipfile.ZipFile(src, "r")
    except (zipfile.BadZipFile, OSError):
        return src
    with zin:
        infos = zin.infolist()
        if len(infos) > _MAX_ZIP_ENTRIES:
            raise ValueError("OOXML archive exceeds maximum entry count")
        for info in infos:
            if not info.filename.lower().endswith(".rels"):
                continue
            raw = _zip_read_bounded(zin, info.filename, 10 * 1024 * 1024)
            root = _safe_xml_fromstring(raw)
            changed = False
            for rel in list(root):
                if (rel.get("TargetMode") or "").strip().lower() != "external":
                    continue
                if _is_unsafe_external_rel(rel.get("Type") or "", rel.get("Target") or ""):
                    root.remove(rel)
                    changed = True
            if changed:
                ET.register_namespace("", _RELS_NS)
                rewrites[info.filename] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        if not rewrites:
            return src
        out_dir = tmp_dir() if callable(tmp_dir) else Path(tmp_dir)
        sanitized = out_dir / src.name
        with zipfile.ZipFile(sanitized, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in infos:
                data = rewrites[info.filename] if info.filename in rewrites else _zip_read_bounded(zin, info.filename, _MAX_ZIP_ENTRY_BYTES * 2)
                zout.writestr(info, data)
    return sanitized


def _check_rtf_for_com(src: Path) -> None:
    if src.suffix.lower() != ".rtf":
        return
    try:
        max_bytes = 16 * 1024 * 1024
        chunk_bytes = 256 * 1024
        overlap = b""
        seen = 0
        with open(src, "rb") as fh:
            while seen < max_bytes:
                chunk = fh.read(min(chunk_bytes, max_bytes - seen))
                if not chunk:
                    break
                seen += len(chunk)
                buf = overlap + chunk
                if _RTF_UNSAFE_RE.search(buf):
                    raise RuntimeError("Refusing unsafe external template/DDE reference in RTF via Office COM")
                overlap = buf[-64:]
    except RuntimeError:
        raise
    except OSError:
        return


def conv_office_com(src: Path, dst: Path) -> None:
    if not IS_WINDOWS:
        raise RuntimeError("Office COM requires Windows")
    ext = dst.suffix.lstrip(".").lower()
    if ext not in _OFFICE_ALLOWED_DST:
        raise RuntimeError(f"Unsupported Office COM target extension: .{ext}")
    if src.suffix.lstrip(".").lower() in OFFICE_SHEET_IN and ext in ("xlsx", "pdf", "csv", "xls"):
        app_name = "excel"
    elif src.suffix.lstrip(".").lower() in OFFICE_SLIDE_IN and ext in ("pptx", "pdf", "ppt"):
        app_name = "powerpoint"
    else:
        app_name = "word"
    progid = BACKENDS.office_app.get(app_name)
    if not progid:
        raise RuntimeError(f"Microsoft {app_name.title()} not installed")
    _check_rtf_for_com(src)
    target = Path(dst)
    if target.exists():
        target = _unique(target)
    dst_abs = str(target.resolve())
    disable_macros = _office_disable_macros_enabled()

    tmp_obj: Optional[tempfile.TemporaryDirectory] = None

    def _lazy_tmp_dir() -> Path:
        nonlocal tmp_obj
        if tmp_obj is None:
            tmp_obj = tempfile.TemporaryDirectory(prefix="clickshift_com_")
        return Path(tmp_obj.name)

    try:
        safe_src = _sanitize_ooxml_for_com(src.resolve(), _lazy_tmp_dir)
        src_abs = str(safe_src.resolve())

        def _do_convert(app: Any) -> None:
            if disable_macros:
                try:
                    app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                except Exception:
                    pass
            if app_name == "word":
                doc = app.Documents.Open(
                    src_abs,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                    NoEncodingDialog=True,
                )
                try:
                    fmt = _FILEFORMAT.get(ext)
                    if ext == "pdf":
                        doc.ExportAsFixedFormat(dst_abs, 17)  # wdExportFormatPDF
                    elif fmt is not None:
                        doc.SaveAs2(dst_abs, FileFormat=fmt)
                    else:
                        doc.SaveAs2(dst_abs)
                finally:
                    doc.Close(SaveChanges=False)
            elif app_name == "excel":
                wb = app.Workbooks.Open(src_abs, UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
                try:
                    if ext == "pdf":
                        wb.ExportAsFixedFormat(-1, dst_abs)  # xlTypePDF
                    else:
                        wb.SaveAs(dst_abs, FileFormat=_FILEFORMAT.get(ext, 51))
                finally:
                    wb.Close(SaveChanges=False)
            else:  # powerpoint
                pres = app.Presentations.Open(src_abs, WithWindow=False)
                try:
                    if ext == "pdf":
                        pres.SaveAs(dst_abs, 32)  # ppSaveAsPDF
                    else:
                        pres.SaveAs(dst_abs, _FILEFORMAT.get(ext, 24))
                finally:
                    pres.Close()

        _OFFICE_POOL.run(app_name, progid, _do_convert)
    finally:
        if tmp_obj is not None:
            tmp_obj.cleanup()
    out_file = Path(dst_abs)
    if not out_file.exists():
        raise RuntimeError("Office COM produced no output")
    wanted = Path(dst)
    if out_file != wanted.resolve():
        out_file.replace(wanted)


# ---- pdf2docx + PyMuPDF/pypdf fallbacks (PDF -> DOCX) ---------------------- #

def conv_pdf2docx(src: Path, dst: Path) -> None:
    _ensure_stdio()
    _patch_pdf2docx_compat()
    pdf2docx_mod: Any = importlib.import_module("pdf2docx")
    cv = pdf2docx_mod.Converter(str(src))
    try:
        cv.convert(str(dst), start=0, end=None)
    finally:
        cv.close()
    if not dst.exists() or dst.stat().st_size == 0:
        raise RuntimeError("pdf2docx produced no output file")
    has_content = _postprocess_pdf2docx(dst)
    if has_content is False:
        dst.unlink(missing_ok=True)
        raise RuntimeError("pdf2docx produced an empty DOCX (likely scanned PDF)")
    # Fallback emptiness check only if XML postprocess could not inspect word/document.xml
    if has_content is None and BACKENDS.python_docx:
        try:
            docx: Any = importlib.import_module("docx")
            d = docx.Document(str(dst))
            if not (any(p.text.strip() for p in d.paragraphs) or d.tables or d.inline_shapes):
                dst.unlink(missing_ok=True)
                raise RuntimeError("pdf2docx produced an empty DOCX (likely scanned PDF)")
        except RuntimeError:
            raise
        except Exception:
            pass


_PYPDF_PARA_RE = re.compile(r"\n\s*\n")
_WR_HEAD_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_WR_BULLET_RE = re.compile(r"^\s*[-*+]\s+")
_WR_NUM_RE = re.compile(r"^\s*\d+[.)]\s+")
_WR_QUOTE_RE = re.compile(r"^>\s?")
_RICH_MD_RE = re.compile(r"(\*\*.+?\*\*|`.+?`)")
_EPUB_SCRIPT_RE = re.compile(r"(?is)<(script|style).+?</\1>")
_EPUB_BREAK_RE = re.compile(r"(?i)</(p|div|h[1-6]|li|br)>")
_EPUB_TAG_RE = re.compile(r"(?s)<[^>]+>")
_EPUB_WS_RE = re.compile(r"[ \t]+")
_EPUB_NL_RE = re.compile(r"\n\s*\n+")


def conv_pdf_to_docx_pymupdf(src: Path, dst: Path) -> None:
    """High-reliability PDF -> DOCX via PyMuPDF (fitz) + python-docx.

    Preserves headings, bold/italic text runs, and embedded images; if a page
    is scanned (image-only with no text), renders the page at 200 DPI into the
    Word document so scanned PDFs convert cleanly.
    """
    fitz: Any = _import_fitz()
    docx: Any = importlib.import_module("docx")
    docx_shared: Any = importlib.import_module("docx.shared")
    Inches = docx_shared.Inches
    Pt = docx_shared.Pt

    pdf = fitz.open(str(src))
    if len(pdf) == 0:
        pdf.close()
        raise RuntimeError("PDF has 0 pages")
    doc = docx.Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    text_flags = getattr(fitz, "TEXTFLAGS_DICT", 0) & ~getattr(fitz, "TEXT_PRESERVE_IMAGES", 0)

    try:
        for page_idx, page in enumerate(pdf):
            if page_idx > 0:
                doc.add_page_break()
            pdict = page.get_text("dict", flags=text_flags) if text_flags else page.get_text("dict")
            blocks = pdict.get("blocks", []) if isinstance(pdict, dict) else []
            wrote_any_text = False
            for b in blocks:
                if b.get("type") == 0:  # text block
                    for line in b.get("lines", []):
                        spans = line.get("spans", [])
                        if not spans:
                            continue
                        line_text = "".join(s.get("text", "") for s in spans).strip()
                        if not line_text:
                            continue
                        wrote_any_text = True
                        max_size = max((float(s.get("size", 11.0)) for s in spans), default=11.0)
                        if max_size >= 18.0:
                            doc.add_heading(line_text, level=1)
                        elif max_size >= 14.5:
                            doc.add_heading(line_text, level=2)
                        else:
                            p = doc.add_paragraph()
                            runs_buf: list[tuple[bool, bool, list[str]]] = []
                            for s in spans:
                                txt = s.get("text", "")
                                if not txt:
                                    continue
                                flags = int(s.get("flags", 0))
                                font_name = (s.get("font") or "").lower()
                                is_b = bool((flags & 16) or "bold" in font_name)
                                is_i = bool((flags & 2) or "italic" in font_name or "oblique" in font_name)
                                if runs_buf and runs_buf[-1][0] == is_b and runs_buf[-1][1] == is_i:
                                    runs_buf[-1][2].append(txt)
                                else:
                                    runs_buf.append((is_b, is_i, [txt]))
                            for is_b, is_i, parts in runs_buf:
                                r = p.add_run("".join(parts))
                                if is_b:
                                    r.bold = True
                                if is_i:
                                    r.italic = True
            # If the page had no extractable text (e.g. scanned page), embed rendered page image
            if not wrote_any_text:
                pix = page.get_pixmap(dpi=200, alpha=False)
                img_bytes = pix.tobytes("png")
                with io.BytesIO(img_bytes) as bio:
                    doc.add_picture(bio, width=Inches(6.2))
    finally:
        pdf.close()
    doc.save(str(dst))


def conv_pdf_to_docx_pypdf(src: Path, dst: Path) -> None:
    """Pure-Python fallback PDF -> DOCX via pypdf + python-docx."""
    pypdf_mod: Any = importlib.import_module("pypdf")
    docx: Any = importlib.import_module("docx")
    docx_shared: Any = importlib.import_module("docx.shared")
    PdfReader = pypdf_mod.PdfReader
    Pt = docx_shared.Pt

    reader = PdfReader(str(src), strict=True)
    if len(reader.pages) == 0:
        raise RuntimeError("PDF has 0 pages")
    doc = docx.Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    wrote_any = False
    for idx, page in enumerate(reader.pages):
        if idx > 0:
            doc.add_page_break()
        txt = (page.extract_text() or "").strip()
        if txt:
            wrote_any = True
            for para in _PYPDF_PARA_RE.split(txt):
                clean = para.strip()
                if clean:
                    doc.add_paragraph(clean)
    if not wrote_any:
        raise RuntimeError("pypdf extracted no text from PDF")
    doc.save(str(dst))


# ---- Pillow image pipeline (C-speed) --------------------------------------- #

_IMG_QUALITY = {"jpg": 95, "jpeg": 95, "webp": 95, "avif": 60, "heic": 60,
                "jxl": 90, "tiff": 90, "png": None}


def _exif_orientation(im: Any) -> int:
    try:
        exif = im.getexif()
        return int(exif.get(274, 1)) if exif else 1
    except Exception:
        return 1


def conv_pillow(src: Path, dst: Path) -> None:
    Image: Any = importlib.import_module("PIL.Image")
    ext = dst.suffix.lstrip(".").lower()
    if ext == "jpg":
        ext = "jpeg"
    with Image.open(src) as im_orig:
        if _exif_orientation(im_orig) != 1:
            ImageOps: Any = importlib.import_module("PIL.ImageOps")
            im = ImageOps.exif_transpose(im_orig) or im_orig
        else:
            im = im_orig
        if ext in ("jpeg", "bmp", "ico") and im.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            rgba = im.convert("RGBA") if im.mode != "RGBA" else im
            bg.paste(rgba, mask=rgba.getchannel("A"))
            im = bg
        elif ext in ("jpeg", "bmp") and im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        save_kw: dict[str, Any] = {}
        q = _IMG_QUALITY.get(ext)
        if q is not None and ext in ("jpeg", "webp", "tiff"):
            save_kw["quality"] = q
        if ext == "webp":
            save_kw["method"] = 4          # fast but strong compression
        if ext in ("png", "tiff"):
            save_kw["optimize"] = True
        if ext == "jpeg":
            save_kw["subsampling"] = 0     # 4:4:4 = maximum fidelity
            exif_bytes = im.info.get("exif", b"")
            if exif_bytes:
                save_kw["exif"] = exif_bytes
        im.save(dst, **save_kw)


# ---- Text-family conversions ----------------------------------------------- #

def _read_text_auto(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    if not raw:
        return "", "utf-8"
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig", errors="replace"), "utf-8-sig"
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16", errors="replace"), "utf-16"
    try:
        return raw.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    enc = "utf-8"
    sample = raw[:65_536]
    if BACKENDS.chardet:
        chardet: Any = importlib.import_module("chardet")
        guess = chardet.detect(sample)
        if guess.get("encoding"):
            enc = guess["encoding"]
    try:
        return raw.decode(enc, errors="replace"), enc
    except LookupError:
        return raw.decode("utf-8", errors="replace"), "utf-8"


def conv_text_generic(src: Path, dst: Path) -> None:
    text, _enc = _read_text_auto(src)
    ext = dst.suffix.lstrip(".").lower()
    if ext in (TEXT_EXTS - {"html", "htm"}):
        dst.write_text(text, encoding="utf-8-sig")
    elif ext in ("html", "htm"):
        body = _html_mod.escape(text)
        dst.write_text(
            "<!DOCTYPE html><html><head><meta charset='utf-8'><title>"
            f"{_html_mod.escape(src.stem)}</title></head><body><pre>{body}</pre></body></html>",
            encoding="utf-8")
    else:
        raise RuntimeError(f"unsupported text target .{ext}")


_MD_HEURISTIC = re.compile(r"^(#{1,6}\s|[*-]\s|\d+.\s|>\s|`|\*\*)", re.M)


def _write_text_to_docx(text: str, dst: Path, md: bool) -> None:
    docx: Any = importlib.import_module("docx")
    docx_shared: Any = importlib.import_module("docx.shared")
    Pt = docx_shared.Pt
    lines = text.splitlines()
    d = docx.Document()
    style = d.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    in_code = False
    code_buf: list[str] = []
    for line in lines:
        if md and line.strip().startswith("```"):
            if in_code:
                p = d.add_paragraph("\n".join(code_buf))
                for r in p.runs:
                    r.font.name = "Consolas"
                    r.font.size = Pt(9)
                code_buf = []
            in_code = not in_code
            continue
        if in_code:
            code_buf.append(line)
            continue
        m = _WR_HEAD_RE.match(line) if md else None
        if m:
            d.add_heading(m.group(2).strip(), level=min(len(m.group(1)), 4))
            continue
        if md:
            mb = _WR_BULLET_RE.match(line)
            if mb:
                d.add_paragraph(line[mb.end():], style="List Bullet")
                continue
            mn = _WR_NUM_RE.match(line)
            if mn:
                d.add_paragraph(line[mn.end():], style="List Number")
                continue
            mq = _WR_QUOTE_RE.match(line)
            if mq:
                p = d.add_paragraph(line[mq.end():])
                p.paragraph_format.left_indent = Pt(18)
                continue
        _add_rich_md_paragraph(d, line, md)
    if code_buf:
        d.add_paragraph("\n".join(code_buf))
    d.save(dst)


def conv_text_to_docx(src: Path, dst: Path) -> None:
    text, _ = _read_text_auto(src)
    md = src.suffix.lstrip(".").lower() in ("md", "markdown") or bool(_MD_HEURISTIC.search(text[:4000]))
    _write_text_to_docx(text, dst, md=md)


def _add_rich_md_paragraph(doc, line: str, md: bool) -> None:
    p = doc.add_paragraph()
    if not md or ("**" not in line and "`" not in line):
        p.add_run(line)
        return
    parts = _RICH_MD_RE.split(line)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            p.add_run(part[2:-2]).bold = True
        elif part.startswith("`") and part.endswith("`"):
            r = p.add_run(part[1:-1])
            r.font.name = "Consolas"
        else:
            p.add_run(part)


def conv_html_to_docx(src: Path, dst: Path) -> None:
    # route through LibreOffice: it maps real HTML/CSS structure into OOXML
    lo_convert(src, dst)


def conv_epub_extract(src: Path, dst: Path) -> None:
    texts: list[str] = []
    with zipfile.ZipFile(src) as z:
        all_infos = z.infolist()
        if len(all_infos) > _MAX_ZIP_ENTRIES:
            raise ValueError(f"EPUB archive exceeds maximum entries ({len(all_infos)} > {_MAX_ZIP_ENTRIES})")
        name_set = {i.filename for i in all_infos}
        names = [n for n in name_set if n.lower().endswith((".xhtml", ".html", ".htm"))]
        total_read = 0
        # respect spine order when possible
        try:
            container_bytes = _zip_read_bounded(z, "META-INF/container.xml", _MAX_ZIP_ENTRY_BYTES)
            total_read += len(container_bytes)
            root = _safe_xml_fromstring(container_bytes)
            rootfiles = root.findall(".//{urn:oasis:names:tc:opendocument:xmlns:container}rootfile")
            opf_path = rootfiles[0].get("full-path") if rootfiles else None
            if opf_path:
                opf_dir = str(Path(opf_path).parent)
                opf_bytes = _zip_read_bounded(z, opf_path, _MAX_ZIP_ENTRY_BYTES)
                total_read += len(opf_bytes)
                opf = _safe_xml_fromstring(opf_bytes)
                ns = {"p": "http://www.idpf.org/2007/opf", "c": "http://www.idpf.org/2007/opf"}
                manifest: dict[str, str] = {}
                for it in opf.findall(".//p:manifest/p:item", ns):
                    item_id = it.get("id")
                    item_href = it.get("href")
                    if item_id and item_href:
                        manifest[item_id] = item_href
                ordered: list[str] = []
                for it in opf.findall(".//p:spine/p:itemref", ns):
                    idref = it.get("idref")
                    href = manifest.get(idref) if idref else None
                    if href:
                        ordered.append((opf_dir + "/" + href).replace("\\", "/").lstrip("./"))
                names = [n for n in ordered if n in name_set] or names
        except ValueError:
            raise
        except Exception:
            pass
        for n in sorted(names):
            remaining = _MAX_ZIP_TOTAL_BYTES - total_read
            if remaining <= 0:
                raise ValueError("EPUB exceeded maximum cumulative decompressed size")
            try:
                raw_entry = _zip_read_bounded(z, n, min(_MAX_ZIP_ENTRY_BYTES, remaining))
                total_read += len(raw_entry)
                html = raw_entry.decode("utf-8", errors="replace")
            except KeyError:
                continue
            html = _EPUB_SCRIPT_RE.sub(" ", html)
            # crude but robust text extraction w/ paragraph breaks
            html = _EPUB_BREAK_RE.sub("\n", html)
            text = _EPUB_TAG_RE.sub(" ", html)
            text = _EPUB_WS_RE.sub(" ", text)
            text = _EPUB_NL_RE.sub("\n\n", text).strip()
            if text:
                texts.append(text)
    joined = "\n\n\n".join(texts) + "\n"
    ext = dst.suffix.lstrip(".").lower()
    if ext in ("docx",):
        _write_text_to_docx(joined, dst, md=bool(_MD_HEURISTIC.search(joined[:4000])))
    elif ext in ("html", "htm"):
        body = _html_mod.escape(joined)
        dst.write_text(f"<!DOCTYPE html><html><head><meta charset='utf-8'></head>"
                       f"<body><pre>{body}</pre></body></html>", encoding="utf-8")
    else:
        dst.write_text(joined, encoding="utf-8-sig")


def conv_docx_to_text(src: Path, dst: Path) -> None:
    docx: Any = importlib.import_module("docx")
    d = docx.Document(str(src))
    lines = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            lines.append("\t".join(c.text for c in row.cells))
    dst.write_text("\n".join(lines), encoding="utf-8-sig")


def _load_transposed_rgb(src: Path) -> Any:
    """Open an image with EXIF orientation applied and alpha composited onto white."""
    Image: Any = importlib.import_module("PIL.Image")
    with Image.open(src) as raw:
        if _exif_orientation(raw) != 1:
            ImageOps: Any = importlib.import_module("PIL.ImageOps")
            im = ImageOps.exif_transpose(raw)
        else:
            im = raw
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            rgba = im.convert("RGBA") if im.mode != "RGBA" else im
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.getchannel("A"))
            return bg
        if im.mode == "RGB":
            if im is raw:
                raw.load()
                return raw.copy()
            return im
        return im.convert("RGB")


def _image_to_embed_stream(src: Path) -> tuple[io.BytesIO, int, int]:
    """Return (JPEG/PNG BytesIO, width_px, height_px) for embedding into DOCX/PPTX."""
    Image: Any = importlib.import_module("PIL.Image")
    with Image.open(src) as raw:
        fmt = (raw.format or "").upper()
        orient = _exif_orientation(raw)
        if fmt == "JPEG" and orient == 1 and raw.mode in ("RGB", "L"):
            w, h = raw.size
            return io.BytesIO(src.read_bytes()), w, h
        if orient != 1:
            ImageOps: Any = importlib.import_module("PIL.ImageOps")
            im = ImageOps.exif_transpose(raw)
        else:
            im = raw
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            rgba = im.convert("RGBA") if im.mode != "RGBA" else im
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.getchannel("A"))
            im = bg
        elif im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        w, h = im.size
        buf = io.BytesIO()
        if fmt == "JPEG":
            im.save(buf, format="JPEG", quality=95, subsampling=0)
        else:
            im.save(buf, format="PNG")
        buf.seek(0)
        return buf, w, h


def conv_images_to_pdf(srcs: list[Path], dst: Path) -> None:
    """Combine one or more images (in exact order) into a multi-page PDF using O(1) RAM."""
    if not srcs:
        raise ValueError("No images provided")
    if not BACKENDS.pillow:
        if len(srcs) == 1:
            lo_convert(srcs[0], dst)
            return
        raise RuntimeError("Pillow is required to combine images into PDF")
    if BACKENDS.pymupdf:
        fitz: Any = _import_fitz()
        Image: Any = importlib.import_module("PIL.Image")
        doc = fitz.open()
        try:
            for p in srcs:
                raw_inserted = False
                with Image.open(p) as hdr:
                    fmt = (hdr.format or "").upper()
                    orient = _exif_orientation(hdr)
                    w_px, h_px = hdr.size
                    mode = hdr.mode
                if orient == 1 and (
                    (fmt == "JPEG" and mode in ("RGB", "L"))
                    or (fmt == "PNG" and mode in ("RGB", "RGBA", "L", "LA"))
                ):
                    w_pt = max(1.0, float(w_px) * 72.0 / 200.0)
                    h_pt = max(1.0, float(h_px) * 72.0 / 200.0)
                    page = doc.new_page(width=w_pt, height=h_pt)
                    page.insert_image(page.rect, stream=p.read_bytes())
                    raw_inserted = True
                if not raw_inserted:
                    im = _load_transposed_rgb(p)
                    buf = io.BytesIO()
                    try:
                        w_px, h_px = im.size
                        w_pt = max(1.0, float(w_px) * 72.0 / 200.0)
                        h_pt = max(1.0, float(h_px) * 72.0 / 200.0)
                        if p.suffix.lower() in (".jpg", ".jpeg"):
                            im.save(buf, format="JPEG", quality=95, subsampling=0)
                        else:
                            im.save(buf, format="PNG")
                        page = doc.new_page(width=w_pt, height=h_pt)
                        page.insert_image(page.rect, stream=buf.getvalue())
                    finally:
                        buf.close()
                        im.close()
            doc.save(str(dst), deflate=True, garbage=1)
        finally:
            doc.close()
        return

    pages = [_load_transposed_rgb(p) for p in srcs]
    try:
        pages[0].save(
            dst,
            "PDF",
            resolution=200.0,
            save_all=True,
            append_images=pages[1:],
        )
    finally:
        for pg in pages:
            pg.close()


def conv_image_to_pdf(src: Path, dst: Path) -> None:
    conv_images_to_pdf([src], dst)


def conv_images_to_docx(srcs: list[Path], dst: Path) -> None:
    """Combine one or more images (in exact order) into a Word document (.docx)."""
    if not srcs:
        raise ValueError("No images provided")
    docx_mod: Any = importlib.import_module("docx")
    docx_shared: Any = importlib.import_module("docx.shared")
    docx_text_enum: Any = importlib.import_module("docx.enum.text")
    Inches = docx_shared.Inches
    Pt = docx_shared.Pt
    WD_ALIGN_PARAGRAPH = docx_text_enum.WD_ALIGN_PARAGRAPH

    doc = docx_mod.Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.6)
    sec.bottom_margin = Inches(0.6)
    sec.left_margin = Inches(0.6)
    sec.right_margin = Inches(0.6)
    # Leave a small height cushion so Word's line-descent / rounding never pushes
    # an exact-fit image or its break onto a blank separator page.
    max_w_in = float(sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches) - 0.05
    max_h_in = float(sec.page_height.inches - sec.top_margin.inches - sec.bottom_margin.inches) - 0.15

    for idx, src in enumerate(srcs):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        if idx > 0:
            p.paragraph_format.page_break_before = True
        stream, w_px, h_px = _image_to_embed_stream(src)
        scale = min(max_w_in / max(1, w_px), max_h_in / max(1, h_px))
        try:
            p.add_run().add_picture(
                stream,
                width=Inches(max(0.5, w_px * scale)),
                height=Inches(max(0.5, h_px * scale)),
            )
        finally:
            stream.close()

    doc.save(str(dst))


def conv_images_to_pptx(srcs: list[Path], dst: Path) -> None:
    """Combine one or more images (in exact order) into a widescreen PowerPoint (.pptx)."""
    if not srcs:
        raise ValueError("No images provided")
    pptx_mod: Any = importlib.import_module("pptx")
    pptx_util: Any = importlib.import_module("pptx.util")
    Inches = pptx_util.Inches

    prs = pptx_mod.Presentation()
    SLIDE_W, SLIDE_H = 13.333, 7.5
    MARGIN = 0.25
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank_layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]
    avail_w = SLIDE_W - 2 * MARGIN
    avail_h = SLIDE_H - 2 * MARGIN

    for src in srcs:
        slide = prs.slides.add_slide(blank_layout)
        stream, w_px, h_px = _image_to_embed_stream(src)
        scale = min(avail_w / max(1, w_px), avail_h / max(1, h_px))
        w_in = w_px * scale
        h_in = h_px * scale
        left = (SLIDE_W - w_in) / 2.0
        top = (SLIDE_H - h_in) / 2.0
        try:
            slide.shapes.add_picture(
                stream, Inches(left), Inches(top), width=Inches(w_in), height=Inches(h_in)
            )
        finally:
            stream.close()

    prs.save(str(dst))


_SVG_HDR_RE = re.compile(r"^\s*(?:<\?xml[^>]*\?>\s*)?(?:<!DOCTYPE[^>]*>\s*)?", re.I)
_SVG_ID_REF_RE = re.compile(r'(\bid=")([^"]+)"|(url\(#)([^)]+)\)|((?:xlink:)?href="#)([^"]+)"')


def conv_pdf_to_html_fidelity(src: Path, dst: Path) -> None:
    """Pixel-faithful PDF -> HTML: every page becomes an inline SVG produced by
    MuPDF with exact embedded font glyph paths, vector shapes, gradients, masks,
    and raster images, laid out like a PDF viewer."""
    fitz: Any = _import_fitz()
    doc = fitz.open(str(src))
    try:
        if doc.needs_pass:
            raise RuntimeError("PDF is password protected")
        title = _html_mod.escape((doc.metadata or {}).get("title") or src.stem)
        with open(dst, "w", encoding="utf-8") as out_fh:
            out_fh.write(
                "<!DOCTYPE html>\n<html lang=\"en\"><head><meta charset=\"utf-8\">"
                "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">"
                f"<title>{title}</title><style>"
                "html,body{margin:0;background:#525659}"
                "body{padding:24px 12px;display:flex;flex-direction:column;align-items:center;gap:24px}"
                ".page{width:100%;background:#fff;box-shadow:0 2px 10px rgba(0,0,0,.45);overflow:hidden}"
                ".page>svg{display:block;width:100%;height:100%}"
                "@media print{html,body{background:#fff;padding:0;gap:0}"
                ".page{box-shadow:none;max-width:none!important;page-break-after:always}}"
                "</style></head><body>\n"
            )
            for i, page in enumerate(doc, start=1):
                svg = page.get_svg_image(matrix=fitz.Identity, text_as_path=True)
                head = _SVG_HDR_RE.sub("", svg[:512], count=1)
                svg = head + svg[512:] if len(svg) > 512 else head
                prefix = f"p{i}_"

                def _prefix_svg_id(m: re.Match) -> str:
                    if m.group(1) is not None:
                        return f'{m.group(1)}{prefix}{m.group(2)}"'
                    if m.group(3) is not None:
                        return f"{m.group(3)}{prefix}{m.group(4)})"
                    return f'{m.group(5)}{prefix}{m.group(6)}"'

                svg = _SVG_ID_REF_RE.sub(_prefix_svg_id, svg)
                w, h = page.rect.width, page.rect.height
                out_fh.write(
                    f'<section class="page" id="page-{i}" style="aspect-ratio:{w:.2f}/{h:.2f};'
                    f'max-width:{w * 4 / 3:.0f}px">{svg}</section>\n'
                )
            out_fh.write("</body></html>\n")
    finally:
        doc.close()


def conv_pdf_to_text(src: Path, dst: Path) -> None:
    ext = dst.suffix.lstrip(".").lower()
    if BACKENDS.pymupdf:
        fitz: Any = _import_fitz()
        doc = fitz.open(str(src))
        try:
            if ext in ("html", "htm"):
                with open(dst, "w", encoding="utf-8") as out_fh:
                    out_fh.write("<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>\n")
                    for idx, page in enumerate(doc):
                        if idx > 0:
                            out_fh.write("\n<hr/>\n")
                        out_fh.write(page.get_text("html"))
                    out_fh.write("\n</body></html>")
                return
            with open(dst, "w", encoding="utf-8-sig") as out_fh:
                for idx, page in enumerate(doc):
                    if idx > 0:
                        out_fh.write("\n\n")
                    out_fh.write((page.get_text("text") or "").strip())
            return
        finally:
            doc.close()
    try:
        pypdf_mod: Any = importlib.import_module("pypdf")
        PdfReader = pypdf_mod.PdfReader
    except ImportError:
        lo_convert(src, dst)
        return
    reader = PdfReader(str(src))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)
    if ext in ("html", "htm"):
        body = _html_mod.escape(text)
        dst.write_text(
            f"<!DOCTYPE html><html><head><meta charset='utf-8'></head><body><pre>{body}</pre></body></html>",
            encoding="utf-8",
        )
    else:
        dst.write_text(text, encoding="utf-8-sig")


# --------------------------------------------------------------------------- #
#  Structured PDF -> Markdown (PyMuPDF layout & typography analysis)
# --------------------------------------------------------------------------- #

_PDF_PUA_MAP = str.maketrans({
    "\ue084": "<", "\ue085": ">", "\ue086": "≤", "\ue087": "≥",
    "\ue09e": "(", "\ue09f": ")", "\ue0a0": "[", "\ue0a1": "]",
    "\ue0c1": "λ", "\uf0b7": "•", "\uf0a7": "▪", "\uf076": "✓",
})
_MD_MONO_RE = re.compile(r"(courier|consolas|mono|menlo|monaco|inconsolata|fira|sourcecode)", re.I)
_MD_SYM_RE = re.compile(r"(opensymbol|wingdings|webdings|zapfdingbats|^symbol$)", re.I)
_MD_BULLET_RE = re.compile(r"^([•◦▪▫●○■□–—-]|\uf0b7|\*)\s+(.*)$")
_MD_SOLO_BULLET_RE = re.compile(r"^([•◦▪▫●○■□]|\*)$")
_MD_NUM_RE = re.compile(r"^(\d{1,3})[.)]\s+(.*)$")
_MD_SOLO_NUM_RE = re.compile(r"^(\d{1,3})\)$")
_MD_REF_RE = re.compile(r"^(\[\d{1,4}\])\s+(.*)$")
_MD_SOLO_REF_RE = re.compile(r"^(\[\d{1,4}\])$")
_MD_SPACES_RE = re.compile(r" {2,}")
_MD_WORD_RE = re.compile(r"[A-Za-z]{4,}")
_MD_MATH_OR_OP_RE = re.compile(r"[={}<>≤≥∑⋅]|^\w\s*=")
_MD_TRAILING_COMMA_RE = re.compile(r"[,;]$")
_MD_TRAILING_PUNCT_RE = re.compile(r"[,;:!?]$")
_MD_NUM_PREFIX_RE = re.compile(r"^\d+(\.\d+)*\.")
_MD_WS_COLLAPSE_RE = re.compile(r"\s+")
_MD_TRAILING_WORD_RE = re.compile(r"[A-Za-z]+$")
_MD_LEADING_WORD_RE = re.compile(r"^[A-Za-z]+")
_MD_STAR_SPACE_RE = re.compile(r"\* \*")
_MD_DBL_STAR_SPACE_RE = re.compile(r"\*\* \*\*")
_MD_PAGE_NUM_RE = re.compile(r"(?:page\s*)?\d{1,4}|[-–—]\s*\d{1,4}\s*[-–—]", re.I)


def _md_clean_text(s: str) -> str:
    s = s.translate(_PDF_PUA_MAP).replace("\u00a0", " ").replace("\ufeff", "")
    return _MD_SPACES_RE.sub(" ", s)


def _md_is_mono(font: str) -> bool:
    return bool(_MD_MONO_RE.search(font or ""))


def _md_is_sym(font: str) -> bool:
    return bool(_MD_SYM_RE.search(font or ""))


def _md_is_bold(font: str, flags: int) -> bool:
    fl = (font or "").lower()
    return bool(flags & 16) or "bold" in fl or "black" in fl


def _md_is_italic(font: str, flags: int) -> bool:
    fl = (font or "").lower()
    return bool(flags & 2) or "italic" in fl or "oblique" in fl


def _md_analyze_typography(doc: Any, text_flags: int = 0) -> tuple[float, list[float], set[str], dict[int, list]]:
    """Return (body_size, heading_tiers_desc, doc_words_lower, cached_page_blocks)."""
    size_counts: Counter[float] = Counter()
    words: set[str] = set()
    cached_blocks: dict[int, list] = {}
    for i in range(min(len(doc), 60)):
        page = doc[i]
        pdict = page.get_text("dict", flags=text_flags) if text_flags else page.get_text("dict")
        text_blocks = [b for b in pdict.get("blocks", ()) if b.get("type") == 0]
        cached_blocks[i] = text_blocks
        for b in text_blocks:
            for ln in b["lines"]:
                line_txt: list[str] = []
                for sp in ln["spans"]:
                    txt = sp["text"]
                    stripped = txt.strip()
                    if not stripped:
                        continue
                    line_txt.append(txt)
                    if not _md_is_mono(sp["font"]) and not _md_is_sym(sp["font"]):
                        size_counts[round(sp["size"], 1)] += len(stripped)
                for w in _MD_WORD_RE.findall("".join(line_txt)):
                    words.add(w.lower())

    body_size = size_counts.most_common(1)[0][0] if size_counts else 11.0
    large = sorted((sz for sz, cnt in size_counts.items() if sz >= body_size * 1.11 and cnt >= 3), reverse=True)
    tiers: list[float] = []
    for sz in large:
        if not tiers or abs(tiers[-1] - sz) > 1.2:
            tiers.append(sz)
    return body_size, tiers[:4], words, cached_blocks


def _md_heading_level(size: float, body_size: float, tiers: list[float],
                      is_bold: bool, has_sym: bool, text: str, line_count: int) -> int:
    if not text or has_sym or len(text) <= 1 or len(text) > 160 or line_count > 3:
        return 0
    if _MD_MATH_OR_OP_RE.search(text):
        return 0
    for idx, tier_sz in enumerate(tiers):
        if abs(size - tier_sz) <= 1.2 or size >= tier_sz - 0.5:
            return min(idx + 1, 4)
    if size >= body_size * 1.11 and not _MD_TRAILING_COMMA_RE.search(text):
        return 2 if size < body_size * 1.35 else 1
    if is_bold and line_count == 1 and len(text) <= 72:
        if not _MD_TRAILING_PUNCT_RE.search(text) and (not text.endswith(".") or _MD_NUM_PREFIX_RE.match(text)):
            if not _MD_REF_RE.match(text) and not text.startswith("©") and len(text.split()) <= 10:
                return min(len(tiers) + 1, 4) if tiers else 3
    return 0


def _md_diagram_boxes(page: Any) -> list[tuple[float, float, float, float]]:
    """Cluster dense non-frame vector drawings into diagram regions via sweep-line."""
    pw, ph = page.rect.width, page.rect.height
    rects = []
    for d in page.get_drawings():
        r = d.get("rect")
        if not r or (r.width > pw * 0.88 and r.height > ph * 0.85) or (r.width < 1 and r.height < 1):
            continue
        rects.append((r.x0, r.y0, r.x1, r.y1))
    if len(rects) < 4:
        return []
    pad = 14.0
    active: list[list[Any]] = []
    retired: list[list[Any]] = []
    for x0, y0, x1, y1 in sorted(rects, key=lambda r: (r[1], r[0])):
        still_active: list[list[Any]] = []
        hit = None
        for c in active:
            cb = c[0]
            if cb[3] + pad < y0:
                retired.append(c)
                continue
            still_active.append(c)
            if not (x1 + pad < cb[0] or x0 - pad > cb[2] or y1 + pad < cb[1]):
                if hit is None:
                    hit = c
                else:
                    # Merge overlapping active cluster directly
                    hb = hit[0]
                    hb[0], hb[1] = min(hb[0], cb[0]), min(hb[1], cb[1])
                    hb[2], hb[3] = max(hb[2], cb[2]), max(hb[3], cb[3])
                    hit[1] += c[1]
                    still_active.pop()
        active = still_active
        if hit is not None:
            cb = hit[0]
            cb[0], cb[1], cb[2], cb[3] = min(cb[0], x0), min(cb[1], y0), max(cb[2], x1), max(cb[3], y1)
            hit[1] += 1
        else:
            active.append([[x0, y0, x1, y1], 1])
    clusters = retired + active
    changed = True
    while changed and len(clusters) > 1:
        changed = False
        nxt: list[list[Any]] = []
        for b, cnt in sorted(clusters, key=lambda c: (c[0][1], c[0][0])):
            hit = next(
                (o for o in nxt if not (b[2] + pad < o[0][0] or b[0] - pad > o[0][2] or b[3] + pad < o[0][1] or b[1] - pad > o[0][3])),
                None,
            )
            if hit is not None:
                ob = hit[0]
                ob[0], ob[1], ob[2], ob[3] = min(ob[0], b[0]), min(ob[1], b[1]), max(ob[2], b[2]), max(ob[3], b[3])
                hit[1] += cnt
                changed = True
            else:
                nxt.append([b, cnt])
        clusters = nxt
    return [(b[0], b[1], b[2], b[3]) for b, cnt in clusters if cnt >= 4 and (b[2] - b[0]) >= 35 and (b[3] - b[1]) >= 20]


def _md_overlap(inner: tuple[float, float, float, float],
                outer: tuple[float, float, float, float], pad: float = 0.0) -> float:
    ix0, iy0 = max(inner[0], outer[0] - pad), max(inner[1], outer[1] - pad)
    ix1, iy1 = min(inner[2], outer[2] + pad), min(inner[3], outer[3] + pad)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return ((ix1 - ix0) * (iy1 - iy0)) / max((inner[2] - inner[0]) * (inner[3] - inner[1]), 1.0)


def _md_real_tables(page: Any) -> list[tuple[tuple[float, float, float, float], str]]:
    """Extract genuine grid tables while ignoring sparse flowchart diagram boxes."""
    out: list[tuple[tuple[float, float, float, float], str]] = []
    try:
        finder = page.find_tables()
    except Exception:
        return out
    for t in getattr(finder, "tables", []):
        if t.row_count < 2 or t.col_count < 2:
            continue
        rows = t.extract()
        if not rows:
            continue
        non_empty = [sum(1 for c in row if c and c.strip()) for row in rows]
        if sum(1 for k in non_empty if k >= 2) < 2 or sum(non_empty) < (t.row_count * t.col_count) * 0.5:
            continue
        if max((len((c or "").splitlines()) for row in rows for c in row), default=1) > 4:
            continue
        clean = [[_MD_WS_COLLAPSE_RE.sub(" ", (c or "")).strip().replace("|", "\\|") for c in row] for row in rows]
        hdr = "| " + " | ".join(clean[0]) + " |"
        sep = "| " + " | ".join("---" for _ in clean[0]) + " |"
        body = ["| " + " | ".join(r) + " |" for r in clean[1:]]
        tb = t.bbox
        out.append(((float(tb[0]), float(tb[1]), float(tb[2]), float(tb[3])), "\n".join([hdr, sep, *body])))
    return out


def _md_format_spans(spans: list[dict]) -> str:
    runs: list[tuple[str, str]] = []
    for sp in spans:
        txt = _md_clean_text(sp["text"])
        if not txt:
            continue
        if _md_is_sym(sp["font"]):
            st = ""
        elif _md_is_mono(sp["font"]):
            st = "c"
        else:
            b, it = _md_is_bold(sp["font"], sp["flags"]), _md_is_italic(sp["font"], sp["flags"])
            st = "bi" if (b and it) else ("b" if b else ("i" if it else ""))
        if runs and runs[-1][0] == st:
            runs[-1] = (st, runs[-1][1] + txt)
        else:
            runs.append((st, txt))
    wrap = {"b": "**", "i": "*", "bi": "***", "c": "`"}
    parts = []
    for st, txt in runs:
        m = wrap.get(st)
        if not m or not txt.strip():
            parts.append(txt)
            continue
        lead = txt[: len(txt) - len(txt.lstrip())]
        trail = txt[len(txt.rstrip()) :]
        parts.append(f"{lead}{m}{txt.strip()}{m}{trail}")
    return "".join(parts)


def _md_join_lines(lines: list[str], doc_words: set[str]) -> str:
    out = ""
    for raw in lines:
        seg = raw.strip()
        if not seg:
            continue
        if not out:
            out = seg
            continue
        if out.endswith("-") and len(out) >= 2 and out[-2].isalpha() and seg[0].islower():
            lw = _MD_TRAILING_WORD_RE.findall(out[:-1])
            rw = _MD_LEADING_WORD_RE.findall(seg)
            out = (out[:-1] + seg) if (lw and rw and (lw[0] + rw[0]).lower() in doc_words) else (out + seg)
        else:
            out = out + " " + seg
    return _MD_STAR_SPACE_RE.sub(" ", _MD_DBL_STAR_SPACE_RE.sub(" ", out))


def _md_merge_split_markers(raw_lines: list[str], styled_lines: list[str]) -> tuple[list[str], list[str]]:
    p_out, s_out, i = [], [], 0
    while i < len(raw_lines):
        rp = raw_lines[i]
        sp = styled_lines[i] if i < len(styled_lines) else rp
        if (_MD_SOLO_NUM_RE.match(rp) or _MD_SOLO_BULLET_RE.match(rp) or _MD_SOLO_REF_RE.match(rp)) and i + 1 < len(raw_lines):
            sp_n = styled_lines[i + 1] if i + 1 < len(styled_lines) else raw_lines[i + 1]
            p_out.append(f"{rp} {raw_lines[i + 1]}")
            s_out.append(f"{sp} {sp_n}")
            i += 2
        else:
            p_out.append(rp)
            s_out.append(sp)
            i += 1
    return p_out, s_out


def conv_pdf_to_md_structured(src: Path, dst: Path) -> None:
    """Convert PDF to clean structured GitHub-Flavored Markdown via PyMuPDF."""
    fitz: Any = _import_fitz()
    doc = fitz.open(str(src))
    try:
        if doc.needs_pass:
            raise RuntimeError("PDF is password protected")
        text_flags = getattr(fitz, "TEXTFLAGS_DICT", 0) & ~getattr(fitz, "TEXT_PRESERVE_IMAGES", 0)
        body_size, tiers, doc_words, cached_blocks = _md_analyze_typography(doc, text_flags=text_flags)
        items_all: list[tuple[str, str]] = []

        for page_idx, page in enumerate(doc):
            ph = page.rect.height
            diag_boxes = _md_diagram_boxes(page)
            tables = _md_real_tables(page)
            tbl_boxes = [tb for tb, _ in tables]
            blocks = cached_blocks.pop(page_idx, None)
            if blocks is None:
                pdict = page.get_text("dict", flags=text_flags) if text_flags else page.get_text("dict")
                blocks = [b for b in pdict.get("blocks", ()) if b.get("type") == 0]
            blocks.sort(key=lambda b: (round(b["bbox"][1] / 5.0), b["bbox"][0]))

            page_items: list[tuple[float, str, str]] = [(tb[1], "table", md) for tb, md in tables]
            for b in blocks:
                rb = b["bbox"]
                bbox: tuple[float, float, float, float] = (float(rb[0]), float(rb[1]), float(rb[2]), float(rb[3]))
                if any(_md_overlap(bbox, tb) >= 0.5 for tb in tbl_boxes):
                    continue
                spans = [sp for ln in b["lines"] for sp in ln["spans"] if sp["text"].strip()]
                if not spans:
                    continue
                raw_lines = [_md_clean_text("".join(sp["text"] for sp in ln["spans"])).strip() for ln in b["lines"]]
                raw_lines = [x for x in raw_lines if x]
                plain = " ".join(raw_lines).strip()
                if not plain:
                    continue
                if (bbox[1] > ph * 0.87 or bbox[3] < ph * 0.08) and _MD_PAGE_NUM_RE.fullmatch(plain):
                    continue
                size_hist: dict[float, int] = {}
                for sp in spans:
                    sz = round(sp["size"], 1)
                    size_hist[sz] = size_hist.get(sz, 0) + len(sp["text"].strip())
                dom_size = max(size_hist.items(), key=lambda kv: kv[1])[0]
                dom_bold = all(_md_is_bold(sp["font"], sp["flags"]) for sp in spans)
                has_sym = any(_md_is_sym(sp["font"]) for sp in spans)
                all_mono = all(_md_is_mono(sp["font"]) for sp in spans)

                if dom_size < body_size * 0.85 and any(_md_overlap(bbox, db, pad=28.0) >= 0.5 for db in diag_boxes):
                    continue

                if all_mono:
                    min_x = min(ln["bbox"][0] for ln in b["lines"] if any(sp["text"].strip() for sp in ln["spans"]))
                    char_w = max(dom_size * 0.58, 3.0)
                    code_lines = []
                    for ln in b["lines"]:
                        lt = "".join(sp["text"] for sp in ln["spans"]).replace("\u00a0", " ").rstrip()
                        if not lt.strip():
                            continue
                        if lt == lt.lstrip():
                            lt = (" " * max(0, int(round((ln["bbox"][0] - min_x) / char_w)))) + lt
                        code_lines.append(lt)
                    body = "\n".join(code_lines)
                    lang = "c" if ("#include " in body or ("double " in body and "{" in body)) else ""
                    page_items.append((bbox[1], "code", f"```{lang}\n{body}\n```"))
                    continue

                h_lvl = _md_heading_level(dom_size, body_size, tiers, dom_bold, has_sym, plain, len(raw_lines))
                if h_lvl:
                    page_items.append((bbox[1], "h", f"{'#' * h_lvl} {_md_clean_text(_md_join_lines(raw_lines, doc_words)).strip()}"))
                    continue

                styled = [x for x in (_md_format_spans(ln["spans"]).strip() for ln in b["lines"]) if x]
                raw_m, styled_m = _md_merge_split_markers(raw_lines, styled)
                if any(_MD_BULLET_RE.match(p) or _MD_NUM_RE.match(p) or _MD_REF_RE.match(p) for p in raw_m):
                    items_md, cur, pfx = [], [], ""
                    for pl, sl in zip(raw_m, styled_m):
                        mb, mn, mr = _MD_BULLET_RE.match(pl), _MD_NUM_RE.match(pl), _MD_REF_RE.match(pl)
                        if mb or mn or mr:
                            if cur:
                                items_md.append(f"{pfx}{_md_join_lines(cur, doc_words)}")
                            if mb is not None:
                                pfx = "- "
                                cur = [_MD_BULLET_RE.sub(r"\2", sl, 1)]
                            elif mr is not None:
                                pfx = "- "
                                cur = [sl]
                            elif mn is not None:
                                pfx = f"{mn.group(1)}. "
                                cur = [_MD_NUM_RE.sub(r"\2", sl, 1)]
                        else:
                            cur.append(sl)
                    if cur:
                        items_md.append(f"{pfx}{_md_join_lines(cur, doc_words)}")
                    page_items.append((bbox[1], "list", "\n".join(items_md)))
                    continue

                para_md = (
                    "  \n".join(styled)
                    if len(raw_lines) <= 4 and all(len(x) < 42 and not x.endswith(".") for x in raw_lines)
                    else _md_join_lines(styled, doc_words)
                )
                page_items.append((bbox[1], "para", para_md))

            page_items.sort(key=lambda it: it[0])
            for _y0, kind, md in page_items:
                if (
                    kind == "code" and items_all and items_all[-1][0] == "code"
                    and items_all[-1][1].startswith("```\n") and md.startswith("```\n")
                ):
                    items_all[-1] = ("code", f"```\n{items_all[-1][1][4:-4]}\n\n{md[4:-4]}\n```")
                else:
                    items_all.append((kind, md))

        dst.write_text("\n\n".join(md for _k, md in items_all).strip() + "\n", encoding="utf-8")
    finally:
        doc.close()


# --------------------------------------------------------------------------- #
#  Presentations (PDF / DOCX <-> PPTX)
# --------------------------------------------------------------------------- #

_PPTX_EMU_PER_PT = 12700.0
_PPTX_EMU_PER_IN = 914400.0
_PPTX_MAX_PT = 4032.0  # PowerPoint OOXML max slide side = 56 in
_PPTX_MIN_PT = 72.0    # 1 in minimum


def _clamp_slide_pts(w_pt: float, h_pt: float) -> tuple[float, float]:
    """Clamp slide dimensions (in points) within PowerPoint's supported bounds."""
    if w_pt <= 0 or h_pt <= 0:
        return (960.0, 540.0)
    top = max(w_pt, h_pt)
    if top > _PPTX_MAX_PT:
        scale = _PPTX_MAX_PT / top
        w_pt *= scale
        h_pt *= scale
    return (max(_PPTX_MIN_PT, w_pt), max(_PPTX_MIN_PT, h_pt))


def _rgb_int_to_tuple(color_int: int) -> tuple[int, int, int]:
    """Unpack PyMuPDF packed 0xRRGGBB integer into (r, g, b) bytes."""
    return ((color_int >> 16) & 0xFF, (color_int >> 8) & 0xFF, color_int & 0xFF)


def _normalize_image_stream(blob: bytes) -> Optional[io.BytesIO]:
    """Ensure image bytes are in a universally accepted raster format (PNG/JPEG)."""
    if not blob:
        return None
    if BACKENDS.pillow:
        try:
            Image: Any = importlib.import_module("PIL.Image")
            with Image.open(io.BytesIO(blob)) as im:
                fmt = (im.format or "").upper()
                if fmt in ("PNG", "JPEG") and im.mode in ("RGB", "RGBA", "L"):
                    return io.BytesIO(blob)
                out = io.BytesIO()
                if im.mode not in ("RGB", "RGBA", "L"):
                    im = im.convert("RGBA" if "A" in im.getbands() else "RGB")
                im.save(out, format="PNG")
                out.seek(0)
                return out
        except Exception:
            return None
    return io.BytesIO(blob)


def _flatten_pptx_shapes(shapes: Any, _sort: bool = True) -> list[tuple[int, int, Any]]:
    """Collect leaf shapes recursively, sorted top-to-bottom then left-to-right."""
    items: list[tuple[int, int, Any]] = []
    for sh in shapes:
        sub_shapes = getattr(sh, "shapes", None)
        if sub_shapes is not None:
            items.extend(_flatten_pptx_shapes(sub_shapes, _sort=False))
        else:
            top = int(getattr(sh, "top", 0) or 0)
            left = int(getattr(sh, "left", 0) or 0)
            items.append((top, left, sh))
    if _sort:
        items.sort(key=lambda t: (t[0], t[1]))
    return items


def conv_pdf_to_pptx(src: Path, dst: Path) -> None:
    """Convert PDF pages to PowerPoint slides (.pptx) via PyMuPDF + python-pptx.

    Text pages emit positioned editable textboxes and extracted images; scanned
    or textless pages fall back to a 200 DPI page image so nothing is dropped.
    """
    fitz: Any = _import_fitz()
    pptx_mod: Any = importlib.import_module("pptx")
    pptx_util: Any = importlib.import_module("pptx.util")
    pptx_color: Any = importlib.import_module("pptx.dml.color")
    Pt = pptx_util.Pt
    RGBColor = pptx_color.RGBColor
    text_flags = getattr(fitz, "TEXTFLAGS_DICT", 0) & ~getattr(fitz, "TEXT_PRESERVE_IMAGES", 0)

    pdf = fitz.open(str(src))
    try:
        if getattr(pdf, "needs_pass", False):
            raise RuntimeError("PDF is password protected")
        if len(pdf) == 0:
            raise RuntimeError("PDF has 0 pages")

        sw, sh = _clamp_slide_pts(float(pdf[0].rect.width), float(pdf[0].rect.height))
        prs = pptx_mod.Presentation()
        prs.slide_width = Pt(sw)
        prs.slide_height = Pt(sh)
        blank_layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]
        xref_cache: dict[int, Optional[bytes]] = {}

        for page in pdf:
            slide = prs.slides.add_slide(blank_layout)
            pw = max(1.0, float(page.rect.width))
            ph = max(1.0, float(page.rect.height))
            scale = min(sw / pw, sh / ph)
            ox = (sw - pw * scale) / 2.0
            oy = (sh - ph * scale) / 2.0

            pdict = page.get_text("dict", flags=text_flags) if text_flags else page.get_text("dict")
            blocks = pdict.get("blocks", []) if isinstance(pdict, dict) else []
            text_blocks = [b for b in blocks if b.get("type") == 0]
            has_text = any(
                "".join(s.get("text", "") for s in ln.get("spans", [])).strip()
                for b in text_blocks
                for ln in b.get("lines", [])
            )

            if not has_text:
                pix = page.get_pixmap(dpi=200, alpha=False)
                with io.BytesIO(pix.tobytes("png")) as stream:
                    slide.shapes.add_picture(
                        stream, Pt(ox), Pt(oy), Pt(pw * scale), Pt(ph * scale)
                    )
                continue

            # Place extracted raster images first so text sits cleanly on top
            seen_img_rects: set[tuple[int, int, int, int]] = set()
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                if xref in xref_cache:
                    raw_norm = xref_cache[xref]
                else:
                    raw_norm = None
                    try:
                        extracted = pdf.extract_image(xref)
                    except Exception:
                        extracted = None
                    if extracted and extracted.get("image"):
                        s_norm = _normalize_image_stream(extracted["image"])
                        if s_norm is not None:
                            raw_norm = s_norm.getvalue()
                            s_norm.close()
                    xref_cache[xref] = raw_norm
                if raw_norm is None:
                    continue
                try:
                    rects = page.get_image_rects(xref)
                except Exception:
                    rects = []
                for rect in rects:
                    rw = float(rect.width)
                    rh = float(rect.height)
                    if rw < 4.0 or rh < 4.0:
                        continue
                    key = (round(rect.x0), round(rect.y0), round(rw), round(rh))
                    if key in seen_img_rects:
                        continue
                    seen_img_rects.add(key)
                    with io.BytesIO(raw_norm) as img_stream:
                        try:
                            slide.shapes.add_picture(
                                img_stream,
                                Pt(ox + rect.x0 * scale),
                                Pt(oy + rect.y0 * scale),
                                Pt(rw * scale),
                                Pt(rh * scale),
                            )
                        except Exception:
                            pass

            for b in text_blocks:
                lines = b.get("lines", [])
                non_empty_lines = [
                    ln for ln in lines
                    if "".join(s.get("text", "") for s in ln.get("spans", [])).strip()
                ]
                if not non_empty_lines:
                    continue
                bx0, by0, bx1, by1 = b.get("bbox", (0.0, 0.0, pw, ph))
                bw = max(24.0, (bx1 - bx0) * scale + 6.0)
                bh = max(14.0, (by1 - by0) * scale + 4.0)
                tb = slide.shapes.add_textbox(
                    Pt(max(0.0, ox + bx0 * scale)),
                    Pt(max(0.0, oy + by0 * scale)),
                    Pt(min(bw, sw)),
                    Pt(min(bh, sh)),
                )
                tf = tb.text_frame
                tf.word_wrap = True
                tf.margin_left = Pt(1)
                tf.margin_right = Pt(1)
                tf.margin_top = Pt(1)
                tf.margin_bottom = Pt(1)

                for l_idx, ln in enumerate(non_empty_lines):
                    p = tf.paragraphs[0] if l_idx == 0 else tf.add_paragraph()
                    p.space_before = Pt(0)
                    p.space_after = Pt(1)
                    for sp in ln.get("spans", []):
                        txt = sp.get("text", "")
                        if not txt:
                            continue
                        run = p.add_run()
                        run.text = txt
                        fsize = max(6.0, min(96.0, float(sp.get("size", 11.0)) * scale * 0.96))
                        run.font.size = Pt(fsize)
                        flags = int(sp.get("flags", 0) or 0)
                        if flags & 16:
                            run.font.bold = True
                        if flags & 2:
                            run.font.italic = True
                        color_val = sp.get("color")
                        if isinstance(color_val, int):
                            r, g, bl = _rgb_int_to_tuple(color_val)
                            run.font.color.rgb = RGBColor(r, g, bl)

        prs.save(str(dst))
    finally:
        pdf.close()


def _copy_docx_runs_to_pptx_para(src_p: Any, dst_p: Any, Pt: Any, RGBColor: Any, base_pt: float) -> None:
    """Copy run text and basic inline formatting from a python-docx paragraph to pptx."""
    if not getattr(src_p, "runs", None):
        r = dst_p.add_run()
        r.text = src_p.text
        r.font.size = Pt(base_pt)
        return
    for srun in src_p.runs:
        if not srun.text:
            continue
        trun = dst_p.add_run()
        trun.text = srun.text
        if srun.bold is not None:
            trun.font.bold = bool(srun.bold)
        if srun.italic is not None:
            trun.font.italic = bool(srun.italic)
        if srun.underline is not None:
            trun.font.underline = bool(srun.underline)
        if srun.font.size is not None:
            trun.font.size = Pt(max(8.0, min(48.0, float(srun.font.size.pt))))
        else:
            trun.font.size = Pt(base_pt)
        rgb = getattr(getattr(srun.font, "color", None), "rgb", None)
        if rgb is not None and len(str(rgb)) == 6:
            try:
                val = int(str(rgb), 16)
                trun.font.color.rgb = RGBColor(*_rgb_int_to_tuple(val))
            except Exception:
                pass
    if not dst_p.runs and src_p.text.strip():
        r = dst_p.add_run()
        r.text = src_p.text
        r.font.size = Pt(base_pt)


_DRAWING_BLIP_TAG = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
_VML_IMAGE_TAG = "{urn:schemas-microsoft-com:vml}imagedata"
_REL_EMBED_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
_REL_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_W_BR_TAG = f"{{{_W_NS}}}br"
_W_TYPE_ATTR = f"{{{_W_NS}}}type"


def _scan_docx_paragraph(child: Any, check_break: bool = True) -> tuple[bool, list[str]]:
    has_break = False
    rids: list[str] = []
    for node in child.iter():
        t = node.tag
        if check_break and not has_break and t == _W_BR_TAG and node.get(_W_TYPE_ATTR) == "page":
            has_break = True
        elif t == _DRAWING_BLIP_TAG:
            rid = node.get(_REL_EMBED_ATTR)
            if rid:
                rids.append(rid)
        elif t == _VML_IMAGE_TAG:
            rid = node.get(_REL_ID_ATTR)
            if rid:
                rids.append(rid)
    return has_break, rids


def conv_docx_to_pptx(src: Path, dst: Path) -> None:
    """Convert Word (.docx) to widescreen PowerPoint (.pptx), splitting on headings/page breaks."""
    docx_mod: Any = importlib.import_module("docx")
    docx_para_mod: Any = importlib.import_module("docx.text.paragraph")
    docx_tbl_mod: Any = importlib.import_module("docx.table")
    pptx_mod: Any = importlib.import_module("pptx")
    pptx_util: Any = importlib.import_module("pptx.util")
    pptx_color: Any = importlib.import_module("pptx.dml.color")
    pptx_align: Any = importlib.import_module("pptx.enum.text")
    Inches = pptx_util.Inches
    Pt = pptx_util.Pt
    RGBColor = pptx_color.RGBColor
    PP_ALIGN = pptx_align.PP_ALIGN

    doc = docx_mod.Document(str(src))
    prs = pptx_mod.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]

    MARGIN_X = 0.7
    BODY_W = 11.93
    BODY_TOP_TITLED = 1.45
    BODY_TOP_PLAIN = 0.65
    BODY_BOTTOM = 6.95

    current_slide: Any = None
    current_title = ""
    current_tf: Any = None
    y_cursor = BODY_TOP_PLAIN
    slide_has_content = False

    def _new_slide(title_text: str = "") -> Any:
        nonlocal current_slide, current_title, current_tf, y_cursor, slide_has_content
        slide = prs.slides.add_slide(blank_layout)
        current_slide = slide
        current_title = title_text.strip()
        current_tf = None
        slide_has_content = False
        if current_title:
            tb = slide.shapes.add_textbox(Inches(MARGIN_X), Inches(0.45), Inches(BODY_W), Inches(0.85))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = current_title
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
            y_cursor = BODY_TOP_TITLED
        else:
            y_cursor = BODY_TOP_PLAIN
        return slide

    def _ensure_body_tf() -> Any:
        nonlocal current_tf
        slide = current_slide if current_slide is not None else _new_slide()
        if current_tf is None:
            tb = slide.shapes.add_textbox(
                Inches(MARGIN_X),
                Inches(y_cursor),
                Inches(BODY_W),
                Inches(max(0.8, BODY_BOTTOM - y_cursor)),
            )
            current_tf = tb.text_frame
            current_tf.word_wrap = True
        return current_tf

    related_parts = getattr(doc.part, "related_parts", {}) or {}
    children = list(doc.element.body)
    idx = 0
    while idx < len(children):
        child = children[idx]
        idx += 1
        tag = child.tag.rsplit("}", 1)[-1]

        if tag == "p":
            para = docx_para_mod.Paragraph(child, doc)
            text = para.text.strip()
            style_name = (getattr(para.style, "name", "") or "").lower()
            has_break = bool(getattr(para.paragraph_format, "page_break_before", False))
            found_break, rids = _scan_docx_paragraph(child, check_break=not has_break)
            if found_break:
                has_break = True

            if has_break and current_slide is not None and (slide_has_content or current_title):
                _new_slide()

            img_blobs: list[bytes] = []
            if rids:
                rel_parts = getattr(getattr(para, "part", None), "related_parts", related_parts) or {}
                for rid in rids:
                    part = rel_parts.get(rid)
                    blob = getattr(part, "blob", None) if part is not None else None
                    if blob:
                        img_blobs.append(blob)

            if style_name == "title" and text:
                title_slide = prs.slides.add_slide(blank_layout)
                tb = title_slide.shapes.add_textbox(Inches(1.0), Inches(2.4), Inches(11.333), Inches(1.4))
                tf = tb.text_frame
                tf.word_wrap = True
                p0 = tf.paragraphs[0]
                p0.alignment = PP_ALIGN.CENTER
                r0 = p0.add_run()
                r0.text = text
                r0.font.size = Pt(38)
                r0.font.bold = True
                r0.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
                if idx < len(children) and children[idx].tag.rsplit("}", 1)[-1] == "p":
                    nxt = docx_para_mod.Paragraph(children[idx], doc)
                    if (getattr(nxt.style, "name", "") or "").lower() == "subtitle" and nxt.text.strip():
                        p1 = tf.add_paragraph()
                        p1.alignment = PP_ALIGN.CENTER
                        p1.space_before = Pt(10)
                        r1 = p1.add_run()
                        r1.text = nxt.text.strip()
                        r1.font.size = Pt(20)
                        r1.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
                        idx += 1
                current_slide = None
                current_tf = None
                slide_has_content = False
                continue

            if style_name in ("heading 1", "heading 2") and text:
                if current_slide is None or slide_has_content or current_title:
                    _new_slide(text)
                else:
                    _new_slide(text)
                continue

            if not text and not img_blobs:
                continue

            if text:
                is_subhead = style_name.startswith("heading ")
                is_list = style_name.startswith("list") or bool(getattr(getattr(para._p, "pPr", None), "numPr", None) is not None)
                line_est = max(1, (len(text) // 85) + 1 + text.count("\n"))
                step_in = line_est * 0.34 + (0.16 if is_subhead else 0.10)
                if current_slide is None:
                    _new_slide()
                elif y_cursor + step_in > BODY_BOTTOM and slide_has_content:
                    cont = f"{current_title} (cont.)" if current_title and not current_title.endswith("(cont.)") else current_title
                    _new_slide(cont)

                tf = _ensure_body_tf()
                p_out = tf.paragraphs[0] if not slide_has_content else tf.add_paragraph()
                p_out.space_before = Pt(6 if is_subhead else 3)
                p_out.space_after = Pt(4)
                if is_list:
                    _copy_bullet_prefix = True
                else:
                    _copy_bullet_prefix = False
                if _copy_bullet_prefix and not text.startswith(("•", "-", "*")):
                    br = p_out.add_run()
                    br.text = "• "
                    br.font.size = Pt(16)
                _copy_docx_runs_to_pptx_para(para, p_out, Pt, RGBColor, 18.0 if is_subhead else 16.0)
                if is_subhead:
                    for rr in p_out.runs:
                        rr.font.bold = True
                y_cursor += step_in
                slide_has_content = True

            for blob in img_blobs:
                stream = _normalize_image_stream(blob)
                if stream is None:
                    continue
                if current_slide is None or (y_cursor > BODY_BOTTOM - 2.0 and slide_has_content):
                    cont = f"{current_title} (cont.)" if current_title and not current_title.endswith("(cont.)") else current_title
                    slide_target = _new_slide(cont)
                else:
                    slide_target = current_slide
                avail_h = max(1.6, BODY_BOTTOM - y_cursor)
                pic_h = min(3.4, avail_h)
                try:
                    pic = slide_target.shapes.add_picture(stream, Inches(MARGIN_X), Inches(y_cursor), height=Inches(pic_h))
                    if pic.width > Inches(BODY_W):
                        ratio = float(Inches(BODY_W)) / float(pic.width)
                        pic.width = Inches(BODY_W)
                        pic.height = int(float(pic.height) * ratio)
                    y_cursor += (float(pic.height) / _PPTX_EMU_PER_IN) + 0.2
                    current_tf = None
                    slide_has_content = True
                except Exception:
                    pass

        elif tag == "tbl":
            tbl = docx_tbl_mod.Table(child, doc)
            nrows = len(tbl.rows)
            ncols = len(tbl.columns) if nrows else 0
            if nrows == 0 or ncols == 0:
                continue
            est_h = min(5.4, max(0.8, nrows * 0.44))
            if current_slide is None or (y_cursor + est_h > BODY_BOTTOM and slide_has_content):
                cont = f"{current_title} (cont.)" if current_title and not current_title.endswith("(cont.)") else current_title
                slide_target = _new_slide(cont)
            else:
                slide_target = current_slide
            tbl_shape = slide_target.shapes.add_table(
                nrows, ncols, Inches(MARGIN_X), Inches(y_cursor), Inches(BODY_W), Inches(est_h)
            )
            ptbl = tbl_shape.table
            for r_idx, row in enumerate(tbl.rows):
                for c_idx in range(min(ncols, len(row.cells))):
                    cell = ptbl.cell(r_idx, c_idx)
                    cell.text = row.cells[c_idx].text.strip()
                    for p in cell.text_frame.paragraphs:
                        for rr in p.runs:
                            rr.font.size = Pt(12)
                            if r_idx == 0:
                                rr.font.bold = True
            y_cursor += est_h + 0.22
            current_tf = None
            slide_has_content = True

    if len(prs.slides) == 0:
        _new_slide(src.stem)
    prs.save(str(dst))


def conv_pptx_to_docx(src: Path, dst: Path) -> None:
    """Convert PowerPoint (.pptx) slides to a structured Word document (.docx)."""
    pptx_mod: Any = importlib.import_module("pptx")
    pptx_shapes_enum: Any = importlib.import_module("pptx.enum.shapes")
    pptx_color_enum: Any = importlib.import_module("pptx.enum.dml")
    docx_mod: Any = importlib.import_module("docx")
    docx_shared: Any = importlib.import_module("docx.shared")
    Inches = docx_shared.Inches
    Pt = docx_shared.Pt
    DocxRGB = docx_shared.RGBColor
    MSO_SHAPE_TYPE = pptx_shapes_enum.MSO_SHAPE_TYPE
    MSO_COLOR_TYPE = pptx_color_enum.MSO_COLOR_TYPE

    prs = pptx_mod.Presentation(str(src))
    doc = docx_mod.Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    if len(prs.slides) == 0:
        doc.add_heading(src.stem, level=1)
        doc.save(str(dst))
        return

    for s_idx, slide in enumerate(prs.slides):
        if s_idx > 0:
            doc.add_page_break()
        title_shape = None
        try:
            title_shape = slide.shapes.title
        except Exception:
            title_shape = None

        title_text = ""
        if title_shape is not None and getattr(title_shape, "has_text_frame", False):
            title_text = title_shape.text_frame.text.strip()
        if title_text:
            doc.add_heading(title_text, level=1)
        else:
            doc.add_heading(f"Slide {s_idx + 1}", level=2)

        for _top, _left, sh in _flatten_pptx_shapes(slide.shapes):
            if title_shape is not None and sh == title_shape:
                continue

            if getattr(sh, "has_table", False):
                stbl = sh.table
                nrows = len(stbl.rows)
                ncols = len(stbl.columns) if nrows else 0
                if nrows > 0 and ncols > 0:
                    dtbl = doc.add_table(rows=nrows, cols=ncols)
                    try:
                        dtbl.style = "Table Grid"
                    except Exception:
                        pass
                    for r_idx, row in enumerate(stbl.rows):
                        for c_idx in range(min(ncols, len(row.cells))):
                            txt = row.cells[c_idx].text.strip()
                            dcell = dtbl.cell(r_idx, c_idx)
                            dcell.text = ""
                            dp = dcell.paragraphs[0]
                            dr = dp.add_run(txt)
                            if r_idx == 0:
                                dr.bold = True
                    doc.add_paragraph("")
                continue

            if getattr(sh, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE or hasattr(sh, "image"):
                try:
                    blob = sh.image.blob
                except Exception:
                    blob = None
                if blob:
                    stream = _normalize_image_stream(blob)
                    if stream is not None:
                        w_in = (float(sh.width or 0) / _PPTX_EMU_PER_IN) if getattr(sh, "width", None) else 4.5
                        w_in = max(0.8, min(6.2, w_in))
                        try:
                            doc.add_picture(stream, width=Inches(w_in))
                        except Exception:
                            pass
                continue

            if getattr(sh, "has_text_frame", False):
                for p in sh.text_frame.paragraphs:
                    ptext = p.text.strip()
                    if not ptext:
                        continue
                    lvl = int(getattr(p, "level", 0) or 0)
                    style = "List Bullet" if lvl > 0 else None
                    try:
                        dp = doc.add_paragraph(style=style) if style else doc.add_paragraph()
                    except Exception:
                        dp = doc.add_paragraph()
                    if lvl > 1:
                        dp.paragraph_format.left_indent = Inches(0.25 * lvl)
                    if p.runs:
                        for pr in p.runs:
                            if not pr.text:
                                continue
                            dr = dp.add_run(pr.text)
                            if pr.font.bold is not None:
                                dr.bold = bool(pr.font.bold)
                            if pr.font.italic is not None:
                                dr.italic = bool(pr.font.italic)
                            if pr.font.underline is not None:
                                dr.underline = bool(pr.font.underline)
                            if pr.font.size is not None:
                                dr.font.size = Pt(max(8.0, min(28.0, float(pr.font.size.pt))))
                            try:
                                clr = pr.font.color
                                if clr is not None and clr.type == MSO_COLOR_TYPE.RGB and clr.rgb is not None:
                                    val = int(str(clr.rgb), 16)
                                    dr.font.color.rgb = DocxRGB(*_rgb_int_to_tuple(val))
                            except Exception:
                                pass
                    else:
                        dp.add_run(ptext)

        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    np = doc.add_paragraph()
                    nr = np.add_run(f"Notes: {notes}")
                    nr.italic = True
                    nr.font.size = Pt(9.5)
                    nr.font.color.rgb = DocxRGB(0x64, 0x74, 0x8B)
        except Exception:
            pass

    doc.save(str(dst))


def _pptx_fill_rgb01(fill: Any, pptx_dml: Any = None) -> Optional[tuple[float, float, float]]:
    """Extract normalized (0..1) RGB fill color from a python-pptx FillFormat if solid."""
    try:
        if pptx_dml is None:
            pptx_dml = importlib.import_module("pptx.enum.dml")
        if fill.type != pptx_dml.MSO_FILL.SOLID:
            return None
        fc = fill.fore_color
        if fc.type == pptx_dml.MSO_COLOR_TYPE.RGB and fc.rgb is not None:
            r, g, b = _rgb_int_to_tuple(int(str(fc.rgb), 16))
            return (r / 255.0, g / 255.0, b / 255.0)
    except Exception:
        pass
    return None


def conv_pptx_to_pdf(src: Path, dst: Path) -> None:
    """Pure-Python PPTX -> PDF renderer via python-pptx + PyMuPDF (no Office/LO required)."""
    pptx_mod: Any = importlib.import_module("pptx")
    pptx_shapes_enum: Any = importlib.import_module("pptx.enum.shapes")
    pptx_dml: Any = importlib.import_module("pptx.enum.dml")
    pptx_align: Any = importlib.import_module("pptx.enum.text")
    fitz: Any = _import_fitz()

    prs = pptx_mod.Presentation(str(src))
    sw = float(prs.slide_width or 914400 * 10) / _PPTX_EMU_PER_PT
    sh = float(prs.slide_height or int(914400 * 5.625)) / _PPTX_EMU_PER_PT
    pdf = fitz.open()
    probe_doc: Any = None
    probe_page: Any = None
    norm_cache: dict[int, Optional[bytes]] = {}

    align_map = {
        pptx_align.PP_ALIGN.CENTER: fitz.TEXT_ALIGN_CENTER,
        pptx_align.PP_ALIGN.RIGHT: fitz.TEXT_ALIGN_RIGHT,
        pptx_align.PP_ALIGN.JUSTIFY: fitz.TEXT_ALIGN_JUSTIFY,
    }

    try:
        if len(prs.slides) == 0:
            pdf.new_page(width=sw, height=sh)
        for slide in prs.slides:
            page = pdf.new_page(width=sw, height=sh)
            bg_rgb = _pptx_fill_rgb01(slide.background.fill, pptx_dml)
            if bg_rgb is not None:
                page.draw_rect(page.rect, color=None, fill=bg_rgb)

            # Collect in z-order (slide.shapes iteration order preserves draw stack)
            stack: list[Any] = []
            def _collect_z(container: Any) -> None:
                for item in container:
                    sub = getattr(item, "shapes", None)
                    if sub is not None:
                        _collect_z(sub)
                    else:
                        stack.append(item)
            _collect_z(slide.shapes)

            for shape in stack:
                x0 = float(getattr(shape, "left", 0) or 0) / _PPTX_EMU_PER_PT
                y0 = float(getattr(shape, "top", 0) or 0) / _PPTX_EMU_PER_PT
                w = float(getattr(shape, "width", 0) or 0) / _PPTX_EMU_PER_PT
                h = float(getattr(shape, "height", 0) or 0) / _PPTX_EMU_PER_PT
                if w <= 0.5 or h <= 0.5:
                    continue
                rect = fitz.Rect(x0, y0, x0 + w, y0 + h)

                if getattr(shape, "shape_type", None) == pptx_shapes_enum.MSO_SHAPE_TYPE.PICTURE or hasattr(shape, "image"):
                    try:
                        raw_blob = shape.image.blob
                        bkey = hash(raw_blob)
                        if bkey not in norm_cache:
                            stream = _normalize_image_stream(raw_blob)
                            norm_cache[bkey] = stream.getvalue() if stream is not None else None
                            if stream is not None:
                                stream.close()
                        cached = norm_cache[bkey]
                        if cached is not None:
                            page.insert_image(rect, stream=cached)
                    except Exception:
                        pass
                    continue

                if getattr(shape, "has_table", False):
                    tbl = shape.table
                    col_ws = [float(c.width) / _PPTX_EMU_PER_PT for c in tbl.columns]
                    row_hs = [float(r.height) / _PPTX_EMU_PER_PT for r in tbl.rows]
                    cy = y0
                    for r_idx, row in enumerate(tbl.rows):
                        rh = row_hs[r_idx] if r_idx < len(row_hs) and row_hs[r_idx] > 0 else (h / max(1, len(tbl.rows)))
                        cx = x0
                        for c_idx in range(len(tbl.columns)):
                            cw = col_ws[c_idx] if c_idx < len(col_ws) and col_ws[c_idx] > 0 else (w / max(1, len(tbl.columns)))
                            crect = fitz.Rect(cx, cy, cx + cw, cy + rh)
                            cfill = _pptx_fill_rgb01(row.cells[c_idx].fill, pptx_dml) or ((0.94, 0.96, 0.98) if r_idx == 0 else None)
                            page.draw_rect(crect, color=(0.72, 0.76, 0.82), width=0.6, fill=cfill)
                            ctext = row.cells[c_idx].text.strip()
                            if ctext:
                                inner = fitz.Rect(cx + 3, cy + 2, cx + cw - 3, cy + rh - 2)
                                if inner.width > 2 and inner.height > 2:
                                    page.insert_textbox(
                                        inner,
                                        ctext,
                                        fontsize=10,
                                        fontname="hebo" if r_idx == 0 else "helv",
                                        color=(0.08, 0.09, 0.12),
                                    )
                            cx += cw
                        cy += rh
                    continue

                fill_rgb = _pptx_fill_rgb01(getattr(shape, "fill", None), pptx_dml)
                if fill_rgb is not None:
                    page.draw_rect(rect, color=None, fill=fill_rgb)

                if getattr(shape, "has_text_frame", False):
                    lines_out: list[tuple[str, float, str, tuple[float, float, float], int]] = []
                    for p in shape.text_frame.paragraphs:
                        ptxt = p.text
                        if not ptxt.strip():
                            continue
                        p_align = align_map.get(getattr(p, "alignment", None), fitz.TEXT_ALIGN_LEFT)
                        fsize = 16.0
                        bold = False
                        italic = False
                        col_rgb = (0.08, 0.09, 0.12)
                        for rn in p.runs:
                            if rn.font.size is not None:
                                fsize = float(rn.font.size.pt)
                            if rn.font.bold:
                                bold = True
                            if rn.font.italic:
                                italic = True
                            try:
                                if (
                                    rn.font.color is not None
                                    and rn.font.color.type == pptx_dml.MSO_COLOR_TYPE.RGB
                                    and rn.font.color.rgb is not None
                                ):
                                    rr, gg, bb = _rgb_int_to_tuple(int(str(rn.font.color.rgb), 16))
                                    col_rgb = (rr / 255.0, gg / 255.0, bb / 255.0)
                            except Exception:
                                pass
                            break
                        fname = "hebi" if (bold and italic) else ("hebo" if bold else ("heit" if italic else "helv"))
                        lines_out.append((ptxt, max(7.0, min(60.0, fsize)), fname, col_rgb, p_align))

                    if lines_out:
                        cur_y = y0 + 2.0
                        bottom = y0 + h - 2.0
                        for ptxt, fs, fname, col_rgb, p_align in lines_out:
                            if cur_y >= bottom - 4.0:
                                break
                            box = fitz.Rect(x0 + 3.0, cur_y, x0 + w - 3.0, bottom)
                            if box.width <= 4 or box.height <= 4:
                                break
                            # Fit text inside remaining box height; binary search shrink on probe page if needed
                            used_fs = fs
                            rc = page.insert_textbox(
                                box, ptxt, fontsize=used_fs, fontname=fname, color=col_rgb, align=p_align
                            )
                            if rc < 0 and fs >= 9.0:
                                max_k = int((fs - 7.0) // 2.0)
                                if max_k > 0:
                                    if probe_doc is None:
                                        probe_doc = fitz.open()
                                        probe_page = probe_doc.new_page(width=sw, height=sh)
                                    pbox = fitz.Rect(0.0, 0.0, box.width, box.height)
                                    lo, hi = 1, max_k
                                    best_fs: Optional[float] = None
                                    while lo <= hi:
                                        mid = (lo + hi) // 2
                                        cand = fs - 2.0 * mid
                                        if probe_page.insert_textbox(
                                            pbox, ptxt, fontsize=cand, fontname=fname, color=col_rgb, align=p_align
                                        ) >= 0:
                                            best_fs = cand
                                            hi = mid - 1
                                        else:
                                            lo = mid + 1
                                    if best_fs is not None:
                                        used_fs = best_fs
                                        page.insert_textbox(
                                            box, ptxt, fontsize=used_fs, fontname=fname, color=col_rgb, align=p_align
                                        )
                                    else:
                                        used_fs = max(5.0, fs - 2.0 * (max_k + 1))
                                else:
                                    used_fs = max(5.0, fs - 2.0)
                            elif rc < 0:
                                used_fs = max(5.0, fs - 2.0)
                            est_lines = max(1, int((len(ptxt) * used_fs * 0.52) // max(20.0, box.width)) + ptxt.count("\n") + 1)
                            cur_y += est_lines * used_fs * 1.28 + 3.0

        pdf.save(str(dst), deflate=True, garbage=1)
    finally:
        if probe_doc is not None:
            try:
                probe_doc.close()
            except Exception:
                pass
        pdf.close()


def conv_via_docx_to_pptx(src: Path, dst: Path) -> None:
    """Convert non-DOCX documents (txt/md/html/doc/odt/rtf) -> temp .docx -> .pptx."""
    s = src.suffix.lstrip(".").lower()
    with tempfile.TemporaryDirectory(prefix="clickshift_pptx_") as td:
        tmp_docx = Path(td) / f"{src.stem}.docx"
        if s in TEXT_EXTS:
            if s in ("html", "htm") and BACKENDS.libreoffice:
                conv_html_to_docx(src, tmp_docx)
            else:
                conv_text_to_docx(src, tmp_docx)
        elif IS_WINDOWS and "word" in BACKENDS.office_app:
            conv_office_com(src, tmp_docx)
        else:
            lo_convert(src, tmp_docx)
        if not tmp_docx.exists():
            raise RuntimeError(f"Failed intermediate conversion {s} -> docx")
        conv_docx_to_pptx(tmp_docx, dst)


def conv_via_pptx_to_docx(src: Path, dst: Path) -> None:
    """Convert legacy/open slide formats (.ppt/.odp) -> temp .pptx -> .docx."""
    with tempfile.TemporaryDirectory(prefix="clickshift_docx_") as td:
        tmp_pptx = Path(td) / f"{src.stem}.pptx"
        if IS_WINDOWS and "powerpoint" in BACKENDS.office_app:
            conv_office_com(src, tmp_pptx)
        else:
            lo_convert(src, tmp_pptx)
        if not tmp_pptx.exists():
            raise RuntimeError(f"Failed intermediate conversion {src.suffix} -> pptx")
        conv_pptx_to_docx(tmp_pptx, dst)


def conv_docx_to_pdf(src: Path, dst: Path) -> None:
    """Pure-Python DOCX -> PDF fallback via python-docx + PyMuPDF."""
    docx_mod: Any = importlib.import_module("docx")
    docx_para_mod: Any = importlib.import_module("docx.text.paragraph")
    docx_tbl_mod: Any = importlib.import_module("docx.table")
    fitz: Any = importlib.import_module("fitz")

    doc = docx_mod.Document(str(src))
    pdf = fitz.open()
    PAGE_W, PAGE_H = 595.0, 842.0  # A4 in pt
    MARGIN_X, MARGIN_Y = 54.0, 54.0
    BODY_W = PAGE_W - 2 * MARGIN_X
    BOTTOM_Y = PAGE_H - MARGIN_Y

    page = pdf.new_page(width=PAGE_W, height=PAGE_H)
    y = MARGIN_Y

    def _next_page() -> None:
        nonlocal page, y
        page = pdf.new_page(width=PAGE_W, height=PAGE_H)
        y = MARGIN_Y

    try:
        for child in doc.element.body:
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "p":
                para = docx_para_mod.Paragraph(child, doc)
                p_break = bool(getattr(para.paragraph_format, "page_break_before", False))
                xml_break, rids = _scan_docx_paragraph(child, check_break=not p_break)
                if (p_break or xml_break) and y > MARGIN_Y + 4.0:
                    _next_page()

                text = para.text.strip()
                style_name = (getattr(para.style, "name", "") or "").lower()
                if text:
                    if style_name == "title":
                        fs, fname, sp_after = 20.0, "hebo", 10.0
                    elif style_name.startswith("heading 1"):
                        fs, fname, sp_after = 16.0, "hebo", 8.0
                    elif style_name.startswith("heading"):
                        fs, fname, sp_after = 13.0, "hebo", 6.0
                    else:
                        fs, fname, sp_after = 11.0, "helv", 5.0
                    if style_name.startswith("list") and not text.startswith(("•", "-", "*")):
                        text = f"• {text}"
                    lines_est = max(1, int((len(text) * fs * 0.52) // BODY_W) + text.count("\n") + 1)
                    needed = lines_est * fs * 1.32 + sp_after
                    if y + needed > BOTTOM_Y and y > MARGIN_Y + 4.0:
                        _next_page()
                    box = fitz.Rect(MARGIN_X, y, MARGIN_X + BODY_W, BOTTOM_Y)
                    page.insert_textbox(box, text, fontsize=fs, fontname=fname, color=(0.08, 0.09, 0.12))
                    y += needed

                for rid in rids:
                    part = doc.part.related_parts.get(str(rid))
                    blob = getattr(part, "blob", None)
                    if not isinstance(blob, (bytes, bytearray)) or not blob:
                        continue
                    stream = _normalize_image_stream(bytes(blob))
                    if stream is None:
                        continue
                    img_h = 180.0
                    if y + img_h > BOTTOM_Y and y > MARGIN_Y + 4.0:
                        _next_page()
                    rect = fitz.Rect(MARGIN_X, y, MARGIN_X + min(BODY_W, 360.0), min(BOTTOM_Y, y + img_h))
                    try:
                        page.insert_image(rect, stream=stream.getvalue())
                        y += img_h + 8.0
                    except Exception:
                        pass
                    finally:
                        stream.close()

            elif tag == "tbl":
                tbl = docx_tbl_mod.Table(child, doc)
                nrows = len(tbl.rows)
                ncols = len(tbl.columns) if nrows else 0
                if nrows == 0 or ncols == 0:
                    continue
                cw = BODY_W / ncols
                rh = 22.0
                if y + rh * min(nrows, 2) > BOTTOM_Y and y > MARGIN_Y + 4.0:
                    _next_page()
                for r_idx, row in enumerate(tbl.rows):
                    if y + rh > BOTTOM_Y:
                        _next_page()
                    for c_idx in range(min(ncols, len(row.cells))):
                        crect = fitz.Rect(MARGIN_X + c_idx * cw, y, MARGIN_X + (c_idx + 1) * cw, y + rh)
                        page.draw_rect(
                            crect,
                            color=(0.72, 0.76, 0.82),
                            width=0.6,
                            fill=(0.94, 0.96, 0.98) if r_idx == 0 else None,
                        )
                        ctext = row.cells[c_idx].text.strip()
                        if ctext:
                            inner = fitz.Rect(crect.x0 + 3, crect.y0 + 2, crect.x1 - 3, crect.y1 - 2)
                            page.insert_textbox(
                                inner,
                                ctext,
                                fontsize=9.5,
                                fontname="hebo" if r_idx == 0 else "helv",
                                color=(0.08, 0.09, 0.12),
                            )
                    y += rh
                y += 8.0

        pdf.save(str(dst), deflate=True, garbage=1)
    finally:
        pdf.close()


# --------------------------------------------------------------------------- #
#  Multi-page documents -> image folders (every page becomes one image)
# --------------------------------------------------------------------------- #

# Sources that contain pages/slides/sheets and must never be reduced to one image.
PAGED_IN = {"pdf", "doc", "docx", "odt", "rtf", "ppt", "pptx", "odp",
            "xls", "xlsx", "ods", "epub", "xps"}
PAGE_IMAGE_EXTS = {"png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff", "gif", "avif", "jxl"}
PAGE_RENDER_DPI = 200


def is_pages_to_images(src: Path, dst_ext: str) -> bool:
    return (src.suffix.lstrip(".").lower() in PAGED_IN
            and dst_ext.lstrip(".").lower() in PAGE_IMAGE_EXTS)


def _unique_dir(path: Path) -> Path:
    """Return ``path`` if free, else name_1, name_2 ... (folder variant of _unique)."""
    if not path.exists():
        return path
    for i in range(1, 500):
        cand = path.with_name(f"{path.name}_{i}")
        if not cand.exists():
            return cand
    return path.with_name(f"{path.name}_{uuid.uuid4().hex[:6]}")


def _dir_size(path: Path) -> int:
    total = 0
    try:
        for p in path.rglob("*"):
            if p.is_file():
                total += p.stat().st_size
    except OSError:
        pass
    return total


def _render_pdf_pages(pdf_path: Path, out_dir: Path, ext: str) -> int:
    """Render every page of ``pdf_path`` into ``out_dir`` as page_001.<ext> ...; return page count."""
    fitz: Any = _import_fitz()
    ext = ext.lstrip(".").lower()
    pil_fmt = {"jpg": "JPEG", "jpeg": "JPEG", "png": "PNG", "webp": "WEBP", "bmp": "BMP",
               "tif": "TIFF", "tiff": "TIFF", "gif": "GIF", "avif": "AVIF", "jxl": "JXL"}[ext]
    doc = fitz.open(str(pdf_path))
    try:
        if getattr(doc, "needs_pass", False):
            raise RuntimeError("PDF is password protected")
        n = len(doc)
        if n == 0:
            raise RuntimeError("Document has 0 pages")
        width = max(3, len(str(n)))
        out_dir.mkdir(parents=True, exist_ok=True)
        if ext == "png":
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=PAGE_RENDER_DPI, alpha=False)
                pix.save(str(out_dir / f"page_{i:0{width}d}.{ext}"))
            return n
        if not BACKENDS.pillow:
            raise RuntimeError(f"Pillow is required to write .{ext} pages")
        Image: Any = importlib.import_module("PIL.Image")
        kwargs: dict[str, Any] = {}
        if pil_fmt == "JPEG":
            kwargs = {"quality": 92, "optimize": True}
        elif pil_fmt == "WEBP":
            kwargs = {"quality": 90}

        def _encode_one(pw: int, ph: int, samples: bytes, target_str: str) -> None:
            im = Image.frombytes("RGB", (pw, ph), samples)
            try:
                if pil_fmt == "GIF":
                    im = im.convert("P", palette=Image.ADAPTIVE)
                im.save(target_str, format=pil_fmt, **kwargs)
            finally:
                im.close()

        if n == 1:
            pix = doc[0].get_pixmap(dpi=PAGE_RENDER_DPI, alpha=False)
            _encode_one(pix.width, pix.height, pix.samples, str(out_dir / f"page_{1:0{width}d}.{ext}"))
            return 1

        max_workers = max(2, min(4, os.cpu_count() or 2))
        pending: list[Any] = []
        with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="clickshift-imgenc") as pool:
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=PAGE_RENDER_DPI, alpha=False)
                target = str(out_dir / f"page_{i:0{width}d}.{ext}")
                pending.append(pool.submit(_encode_one, pix.width, pix.height, pix.samples, target))
                if len(pending) >= max_workers * 2:
                    pending.pop(0).result()
            for fut in pending:
                fut.result()
        return n
    finally:
        doc.close()


def conv_pages_to_image_folder(src: Path, out_dir: Path, ext: str) -> None:
    """Convert a multi-page document into a folder (``out_dir``) holding one image per page."""
    s = src.suffix.lstrip(".").lower()
    try:
        if s in ("pdf", "xps", "epub"):
            _render_pdf_pages(src, out_dir, ext)
            return
        with tempfile.TemporaryDirectory(prefix="clickshift_pages_") as td:
            tmp_pdf = Path(td) / f"{src.stem}.pdf"
            last_err = ""
            for r in _routes_for(s, "pdf"):
                try:
                    r.fn(src, tmp_pdf)
                    if tmp_pdf.exists() and tmp_pdf.stat().st_size > 0:
                        break
                except Exception as exc:
                    last_err = f"{r.name}: {type(exc).__name__}: {exc}"
                    tmp_pdf.unlink(missing_ok=True)
            if not tmp_pdf.exists():
                raise RuntimeError(f"Could not render pages ({last_err or 'no PDF route available'})")
            _render_pdf_pages(tmp_pdf, out_dir, ext)
    except Exception:
        shutil.rmtree(out_dir, ignore_errors=True)  # never leave a half-filled folder
        raise


# --------------------------------------------------------------------------- #
#  Routing table
# --------------------------------------------------------------------------- #

@dataclass
class Route:
    name: str
    fn: BackendFn
    weight: int = 0   # static preference: higher = try first


_ROUTES_LOCK = threading.Lock()
_ROUTES_CACHE: dict[
    tuple[str, str],
    tuple[int, bool, tuple[tuple[str, Any], ...], tuple[Route, ...]],
] = {}


def _snapshot_backend_val(val: Any) -> Any:
    return tuple(sorted(val.items())) if isinstance(val, dict) else val


def _routes_for(src_ext: str, dst_ext: str) -> list[Route]:
    s, d = src_ext.lower().lstrip("."), dst_ext.lower().lstrip(".")
    if s == d:
        return []
    b_id = id(BACKENDS)
    cached = _ROUTES_CACHE.get((s, d))
    if cached is not None:
        c_bid, c_win, c_deps, c_routes = cached
        if (
            c_bid == b_id
            and c_win == IS_WINDOWS
            and all(_snapshot_backend_val(getattr(BACKENDS, k)) == v for k, v in c_deps)
        ):
            return list(c_routes)

    accessed: dict[str, Any] = {}

    def _b(attr: str) -> Any:
        if attr not in accessed:
            accessed[attr] = getattr(BACKENDS, attr)
        return accessed[attr]

    def _has_office() -> bool:
        return IS_WINDOWS and bool(_b("office_app"))

    def _has_word() -> bool:
        return IS_WINDOWS and "word" in _b("office_app")

    def _has_ppt() -> bool:
        return IS_WINDOWS and "powerpoint" in _b("office_app")

    routes: list[Route] = []

    def add(name: str, fn: BackendFn, w: int = 0):
        routes.append(Route(name, fn, w))

    # same family shortcuts ---------------------------------------------------
    if s in IMAGE_EXTS and d in IMAGE_EXTS:
        if _b("pillow"):
            add("pillow", conv_pillow, 100)
        add("libreoffice", conv_libreoffice, 20)
    elif s == "pdf" and d == "docx":
        if _b("pdf2docx"):
            add("pdf2docx", conv_pdf2docx, 95)
        if _has_word():
            # Word 2013+ PDF reflow: excellent for text-centric PDFs
            add("office-com", conv_office_com, 85)
        if _b("pymupdf") and _b("python_docx"):
            add("pymupdf-docx", conv_pdf_to_docx_pymupdf, 75)
        add("libreoffice", conv_libreoffice, 65)
        if _b("pypdf") and _b("python_docx"):
            add("pypdf-docx", conv_pdf_to_docx_pypdf, 50)
    elif s == "pdf" and d == "pptx":
        if _b("pymupdf") and _b("python_pptx"):
            add("pymupdf-python-pptx", conv_pdf_to_pptx, 95)
        add("libreoffice", conv_libreoffice, 40)
    elif s == "pdf" and d in ("md", "markdown"):
        if _b("pymupdf"):
            add("pymupdf-md", conv_pdf_to_md_structured, 99)
        else:
            add("pypdf", conv_pdf_to_text, 90)
        add("libreoffice", conv_libreoffice, 40)
    elif s == "pdf" and d in ("txt", "csv", "html", "htm"):
        if d in ("html", "htm") and _b("pymupdf"):
            add("pymupdf-html", conv_pdf_to_html_fidelity, 99)
        else:
            add("pypdf", conv_pdf_to_text, 90 if d not in ("html", "htm") else 45)
        if d in _WORD_ONLY_FMTS and _has_word():
            add("office-com", conv_office_com, 60)
        add("libreoffice", conv_libreoffice, 40)
    elif s == "pdf" and d in ("odt", "rtf"):
        if d == "rtf" and _has_word():
            add("office-com", conv_office_com, 85)
        add("libreoffice", conv_libreoffice, 75)
    elif s in PAGED_IN and d in PAGE_IMAGE_EXTS:
        if _b("pymupdf") and (d == "png" or _b("pillow")):
            _page_img_ext = d
            add("pages-to-images", lambda src_, dst_, _e=_page_img_ext: conv_pages_to_image_folder(src_, dst_, _e), 100)
        if s == "pdf":
            add("libreoffice", conv_libreoffice, 30)
    elif s in IMAGE_EXTS and d == "pdf":
        if _b("pillow"):
            add("pillow-pdf", conv_image_to_pdf, 100)
        add("libreoffice", conv_libreoffice, 30)
    elif s in IMAGE_EXTS and d == "docx":
        if _b("pillow") and _b("python_docx"):
            add("pillow-python-docx", lambda s_, d_: conv_images_to_docx([s_], d_), 100)
    elif s in IMAGE_EXTS and d == "pptx":
        if _b("pillow") and _b("python_pptx"):
            add("pillow-python-pptx", lambda s_, d_: conv_images_to_pptx([s_], d_), 100)
    elif s == "docx" and d == "pptx":
        if _b("python_docx") and _b("python_pptx"):
            add("python-docx-python-pptx", conv_docx_to_pptx, 95)
        add("libreoffice", conv_libreoffice, 30)
    elif s in (OFFICE_DOC_IN | {"md", "markdown", "rst"}) - {"docx"} and d == "pptx":
        if s in TEXT_EXTS and _b("python_docx") and _b("python_pptx"):
            add("text-python-docx-python-pptx", conv_via_docx_to_pptx, 85)
        elif _b("python_docx") and _b("python_pptx") and (_has_word() or _b("libreoffice")):
            add("via-python-docx-python-pptx", conv_via_docx_to_pptx, 85)
        add("libreoffice", conv_libreoffice, 30)
    elif s == "pptx" and d == "docx":
        if _b("python_pptx") and _b("python_docx"):
            add("python-pptx-python-docx", conv_pptx_to_docx, 95)
        add("libreoffice", conv_libreoffice, 30)
    elif s in ("ppt", "odp") and d == "docx":
        if _b("python_pptx") and _b("python_docx") and (_has_ppt() or _b("libreoffice")):
            add("via-python-pptx-python-docx", conv_via_pptx_to_docx, 85)
        add("libreoffice", conv_libreoffice, 30)
    elif s in ("docx",) and d in TEXT_EXTS - {"docx"}:
        if s == "docx" and d in ("txt", "md", "log"):
            add("python-docx", conv_docx_to_text, 90)
        if d in _WORD_ONLY_FMTS and _has_word():
            add("office-com", conv_office_com, 45)
        add("libreoffice", conv_libreoffice, 50)
    elif s in TEXT_EXTS and d == "pdf":
        if _has_word():
            # Word opens txt/md/html and exports a perfect PDF
            add("office-com", conv_office_com, 85)
        if _b("libreoffice"):
            add("libreoffice", conv_libreoffice, 60)
    elif s in TEXT_EXTS and d == "docx":
        if _b("python_docx"):
            if s in ("html", "htm") and _b("libreoffice"):
                add("libreoffice-html", conv_html_to_docx, 80)
            add("python-docx-md", conv_text_to_docx, 70)
        add("libreoffice", conv_libreoffice, 50)
    elif s == "epub":
        if _b("python_docx") and d == "docx":
            add("epub->docx", conv_epub_extract, 80)
        if d in ("txt", "md", "html", "htm"):
            add("epub-extract", conv_epub_extract, 90)
        add("libreoffice", conv_libreoffice, 30)
    elif s in TEXT_EXTS and d in TEXT_EXTS:
        add("text-passthru", conv_text_generic, 100)
        if s in OFFICE_DOC_IN and d in ("html", "txt"):
            if _has_office():
                add("office-com", conv_office_com, 60)
            add("libreoffice", conv_libreoffice, 40)
    elif s in ("doc", "xls", "ppt", "odt", "ods", "odp", "rtf") and d in ("docx", "xlsx", "pptx", "pdf"):
        if _has_office():
            add("office-com", conv_office_com, 100)
        add("libreoffice", conv_libreoffice, 70)
    elif s in OFFICE_SLIDE_IN and d == "pdf":
        if _has_ppt():
            add("office-com", conv_office_com, 100)
        add("libreoffice", conv_libreoffice, 70)
        if s == "pptx" and _b("python_pptx") and _b("pymupdf"):
            add("python-pptx-pymupdf-pdf", conv_pptx_to_pdf, 50)
    elif s in OFFICE_DOC_IN | OFFICE_SHEET_IN and d == "pdf":
        if _has_office():
            add("office-com", conv_office_com, 100)
        add("libreoffice", conv_libreoffice, 70)
        if s == "docx" and _b("python_docx") and _b("pymupdf"):
            add("python-docx-pymupdf-pdf", conv_docx_to_pdf, 45)
    elif s in OFFICE_DOC_IN and d in ("docx", "odt", "rtf", "html"):
        if _has_office():
            add("office-com", conv_office_com, 100)
        add("libreoffice", conv_libreoffice, 70)
        if d == "docx" and _b("python_docx"):
            add("python-docx-md", conv_text_to_docx, 65)
        if s in TEXT_EXTS and d in ("html", "htm"):
            add("text-passthru", conv_text_generic, 90)
    elif s in ("docx", "xlsx", "pptx") and d in ("pdf", "odt", "ods", "odp", "rtf", "html", "txt", "csv"):
        if d == "pdf" and _has_office():
            add("office-com", conv_office_com, 100)
        elif s == "docx" and d in _WORD_ONLY_FMTS and _has_word():
            add("office-com", conv_office_com, 45)
        add("libreoffice", conv_libreoffice, 70)
    else:
        known_targets = (
            IMAGE_EXTS | TEXT_EXTS | OFFICE_DOC_IN | OFFICE_SHEET_IN | OFFICE_SLIDE_IN
            | {"pdf", "docx", "xlsx", "pptx", "doc", "xls", "ppt", "odt", "ods", "odp", "rtf", "epub", "xps"}
        )
        if s in (SUPPORTED_INPUT_EXTS | TEXT_EXTS) and d in known_targets:
            if s in TEXT_EXTS and d in TEXT_EXTS:
                add("text-passthru", conv_text_generic, 50)
            add("libreoffice", conv_libreoffice, 10)
            if d in _OFFICE_ALLOWED_DST and _has_office():
                add("office-com", conv_office_com, 5)

    # prune routes whose deps are missing
    live: list[Route] = []
    for r in routes:
        if r.name.startswith("pillow") and not _b("pillow"):
            continue
        if ("python-docx" in r.name) and not _b("python_docx"):
            continue
        if ("python-pptx" in r.name) and not _b("python_pptx"):
            continue
        if ("pymupdf" in r.name or "pages-to-images" in r.name) and not _b("pymupdf"):
            continue
        if r.name == "office-com" and not (IS_WINDOWS and _b("office_app")):
            continue
        if r.name.startswith("libreoffice") and not _b("libreoffice"):
            continue
        if r.name == "pdf2docx" and not _b("pdf2docx"):
            continue
        live.append(r)
    live.sort(key=lambda r: -r.weight)

    deps = tuple((k, _snapshot_backend_val(v)) for k, v in sorted(accessed.items()))
    with _ROUTES_LOCK:
        if len(_ROUTES_CACHE) > 512:
            _ROUTES_CACHE.clear()
        _ROUTES_CACHE[(s, d)] = (b_id, IS_WINDOWS, deps, tuple(live))
    return live


# --------------------------------------------------------------------------- #
#  Engine
# --------------------------------------------------------------------------- #

class ConversionEngine:
    """Thread-safe orchestrator with optional ML backend selection."""

    def __init__(self, enable_ml: bool = True, history_file: Optional[Path] = None) -> None:
        self.enable_ml = enable_ml and os.environ.get("CLICKSHIFT_DISABLE_ML") != "1"
        self._selector: Any = None
        self.cloud: Any = None
        if self.enable_ml:
            try:
                from .learning import BanditSelector
                from .monitor import DATA_DIR   # honours CLICKSHIFT_HOME env var
                self._selector = BanditSelector(history_file or (DATA_DIR / "history.jsonl"))
            except Exception:
                self._selector = None
            if self._selector is not None:
                try:
                    from .cloud import CloudSync
                    cloud = CloudSync(self._selector)
                    self._selector.attach_cloud(cloud)
                    self.cloud = cloud
                    cloud.start()          # no-op unless cloud_sync is on
                except Exception:
                    self.cloud = None
        self._pool = ThreadPoolExecutor(max_workers=max(2, (os.cpu_count() or 4)))
        self._lock = threading.Lock()

    # -- single ------------------------------------------------------------- #
    def convert(self, src: str | Path, dst_ext: str, out_dir: Optional[str | Path] = None,
                out_path: Optional[str | Path] = None) -> ConversionResult:
        from . import debuglog
        dbg_mode = debuglog.get_mode()
        sess = debuglog.new_session(src, dst_ext) if dbg_mode != debuglog.MODE_OFF else None
        src = Path(src)
        try:
            src_st = src.stat()
        except OSError:
            res = ConversionResult(str(src), "", "-", 0, False, "source file not found")
            if sess is not None:
                sess.update(ok=False, final_error=res.error, attempts=[],
                            planned_out="", out_path="")
                debuglog.finalize(sess, res)
            return res
        src_size = src_st.st_size
        dst_ext = _norm_ext(dst_ext)
        if not dst_ext:
            res = ConversionResult(str(src), "", "-", 0, False, "missing target extension")
            if sess is not None:
                sess.update(ok=False, final_error=res.error, attempts=[],
                            planned_out="", out_path="")
                debuglog.finalize(sess, res)
            return res
        base = Path(out_path) if out_path else (Path(out_dir) if out_dir else src.parent) / (src.stem + dst_ext)
        pages_mode = is_pages_to_images(src, dst_ext)
        if pages_mode:
            # every page -> one image, collected in a folder named after the document
            folder = Path(out_path) if out_path else (Path(out_dir) if out_dir else src.parent) / src.stem
            base = _unique_dir(folder)
        else:
            base = _unique(Path(base))
        if sess is not None:
            sess["planned_out"] = str(base)
            sess["dst_ext"] = dst_ext
            sess["bytes_in"] = src_size
        routes = _routes_for(src.suffix, dst_ext)
        if sess is not None:
            sess["routes_static"] = [r.name for r in routes]
        sel = self._selector
        ml_vec: Any = None
        sel_feat = bool(sel is not None and hasattr(sel, "extract_features"))
        if sel is not None:
            if sel_feat:
                ml_vec = sel.extract_features(src, dst_ext, src_size=src_size)
            if hasattr(sel, "rank_debug"):
                routes, ml_info = (
                    sel.rank_debug(src, dst_ext, routes, features=ml_vec, src_size=src_size)
                    if sel_feat else sel.rank_debug(src, dst_ext, routes)
                )
                if sess is not None:
                    sess["ml"].update(ml_info)
            else:
                routes = (
                    sel.rank(src, dst_ext, routes, features=ml_vec, src_size=src_size)
                    if sel_feat else sel.rank(src, dst_ext, routes)
                )
        if sess is not None:
            sess["routes"] = [{"name": r.name, "weight": r.weight} for r in routes]
        if not routes:
            res = ConversionResult(str(src), dst_ext, "-", 0, False,
                                   f"No conversion path known for {src.suffix} -> {dst_ext}")
            if sess is not None:
                sess.update(ok=False, final_error=res.error, attempts=[])
                debuglog.finalize(sess, res)
            return res
        t0 = time.perf_counter()
        err = ""
        attempts: Optional[list[dict]] = [] if sess is not None else None
        mkdir_done = False
        for r in routes:
            ta = time.perf_counter()
            try:
                if not mkdir_done:
                    base.parent.mkdir(parents=True, exist_ok=True)
                    mkdir_done = True
                r.fn(src, base)
                now = time.perf_counter()
                ms = (now - t0) * 1000
                att_ms = (now - ta) * 1000
                if attempts is not None:
                    attempts.append({"backend": r.name, "ok": True, "ms": att_ms,
                                     "error": "", "traceback": ""})
                if sel is not None:
                    if sel_feat:
                        sel.record(src, dst_ext, r.name, att_ms, True,
                                   features=ml_vec, src_size=src_size)
                    else:
                        sel.record(src, dst_ext, r.name, att_ms, True)
                if pages_mode:
                    out_bytes = _dir_size(base) if base.is_dir() else 0
                else:
                    try:
                        out_bytes = base.stat().st_size
                    except OSError:
                        out_bytes = 0
                res = ConversionResult(str(src), str(base), r.name, ms, True,
                                       bytes_in=src_size, bytes_out=out_bytes)
                if sess is not None:
                    sess.update(ok=True, winning_backend=r.name, total_ms=ms,
                                attempts=attempts, out_path=str(base),
                                bytes_out=out_bytes)
                    debuglog.finalize(sess, res)
                return res
            except Exception as exc:  # try next backend
                att_ms = (time.perf_counter() - ta) * 1000
                if sel is not None:
                    if sel_feat:
                        sel.record(src, dst_ext, r.name, att_ms, False,
                                   features=ml_vec, src_size=src_size)
                    else:
                        sel.record(src, dst_ext, r.name, att_ms, False)
                err = f"{type(exc).__name__}: {exc}"
                if attempts is not None:
                    attempts.append({"backend": r.name, "ok": False, "ms": att_ms,
                                     "error": err, "exc": exc})
                if dbg_mode != debuglog.MODE_OFF:
                    traceback.print_exc()
        ms = (time.perf_counter() - t0) * 1000
        res = ConversionResult(str(src), str(base), routes[0].name, ms, False, err)
        if sess is not None:
            sess.update(ok=False, final_error=err, attempts=attempts,
                        total_ms=ms, out_path=str(base))
            debuglog.finalize(sess, res)
        return res

    # -- combine multiple images -> single pdf/docx/pptx -------------------- #
    def combine_images(self, srcs: Sequence[str | Path], dst_ext: str,
                       out_dir: Optional[str | Path] = None,
                       out_path: Optional[str | Path] = None) -> ConversionResult:
        from . import debuglog
        paths = [Path(s) for s in srcs if Path(s).exists()]
        ext = _norm_ext(dst_ext)
        if not paths:
            return ConversionResult("", "", "-", 0, False, "no source images found")
        if ext not in (".pdf", ".docx", ".pptx"):
            return ConversionResult(str(paths[0]), "", "-", 0, False,
                                    f"unsupported combine target: {dst_ext}")
        base = Path(out_path) if out_path else (Path(out_dir) if out_dir else paths[0].parent) / (paths[0].stem + ext)
        base = _unique(Path(base))
        dbg_mode = debuglog.get_mode()
        sess = debuglog.new_session(paths[0], ext) if dbg_mode != debuglog.MODE_OFF else None
        bytes_in = sum((p.stat().st_size for p in paths), 0)
        if sess is not None:
            sess["planned_out"] = str(base)
            sess["dst_ext"] = ext
            sess["bytes_in"] = bytes_in
        backend_name = {".pdf": "pillow-pdf", ".docx": "pillow-python-docx", ".pptx": "pillow-python-pptx"}[ext]
        fn = {".pdf": conv_images_to_pdf, ".docx": conv_images_to_docx, ".pptx": conv_images_to_pptx}[ext]
        t0 = time.perf_counter()
        try:
            base.parent.mkdir(parents=True, exist_ok=True)
            fn(paths, base)
            ms = (time.perf_counter() - t0) * 1000
            try:
                bytes_out = base.stat().st_size
            except OSError:
                bytes_out = 0
            res = ConversionResult(str(paths[0]), str(base), backend_name, ms, True,
                                   bytes_in=bytes_in, bytes_out=bytes_out)
            if sess is not None:
                sess.update(ok=True, winning_backend=backend_name, total_ms=ms,
                            out_path=str(base), bytes_out=bytes_out)
                debuglog.finalize(sess, res)
            return res
        except Exception as exc:
            ms = (time.perf_counter() - t0) * 1000
            err = f"{type(exc).__name__}: {exc}"
            res = ConversionResult(str(paths[0]), str(base), backend_name, ms, False, err,
                                   bytes_in=bytes_in)
            if sess is not None:
                sess.update(ok=False, final_error=err, total_ms=ms, out_path=str(base),
                            attempts=[{"backend": backend_name, "ok": False, "ms": ms,
                                       "error": err, "exc": exc}])
                debuglog.finalize(sess, res)
            return res

    # -- batch -------------------------------------------------------------- #
    def convert_many(self, jobs: list[tuple[str, str]],
                     progress: Optional[Callable[[ConversionResult], None]] = None) -> list[ConversionResult]:
        if not jobs:
            return []
        if len(jobs) == 1:
            res = self.convert(jobs[0][0], jobs[0][1])
            if progress:
                try:
                    progress(res)
                except Exception:
                    pass
            return [res]

        def _emit(res: ConversionResult) -> None:
            if progress:
                try:
                    progress(res)
                except Exception:
                    pass

        n_jobs = len(jobs)
        results: list[Optional[ConversionResult]] = [None] * n_jobs
        sel = self._selector
        batch_ctx = (
            sel.batch_persistence()
            if sel is not None and hasattr(sel, "batch_persistence")
            else contextlib.nullcontext()
        )
        with batch_ctx:
            remaining: list[tuple[int, str, str]] = []
            lo_groups: dict[tuple[str, Optional[str]], list[tuple[int, str, str, Path, Path, str, int, Any]]] = {}
            reserved_outs: set[str] = set()
            sel_feat = bool(sel is not None and hasattr(sel, "extract_features"))
            has_bandit_data = (
                bool(sel is not None and (not hasattr(sel, "bandit") or sel.bandit.total_samples() > 0))
            )

            for idx, (s, e) in enumerate(jobs):
                dst_ext = _norm_ext(e)
                src_p = Path(s)
                if not dst_ext or is_pages_to_images(src_p, dst_ext):
                    remaining.append((idx, s, e))
                    continue
                routes = _routes_for(src_p.suffix, dst_ext)
                if not routes or not any(r.name == "libreoffice" and r.fn is conv_libreoffice for r in routes):
                    remaining.append((idx, s, e))
                    continue
                if not has_bandit_data and (routes[0].name != "libreoffice" or routes[0].fn is not conv_libreoffice):
                    remaining.append((idx, s, e))
                    continue
                try:
                    src_st = src_p.stat()
                except OSError:
                    remaining.append((idx, s, e))
                    continue
                src_size = src_st.st_size
                ml_vec = None
                if sel is not None:
                    if sel_feat:
                        ml_vec = sel.extract_features(src_p, dst_ext, src_size=src_size)
                        routes = sel.rank(src_p, dst_ext, routes, features=ml_vec, src_size=src_size)
                    else:
                        routes = sel.rank(src_p, dst_ext, routes)
                if routes and routes[0].name == "libreoffice" and routes[0].fn is conv_libreoffice:
                    infilter = _lo_infilter_for(src_p.suffix.lstrip(".").lower(), dst_ext.lstrip("."))
                    base = _unique(src_p.parent / (src_p.stem + dst_ext))
                    norm_out = os.path.normcase(os.path.abspath(str(base)))
                    if norm_out in reserved_outs:
                        remaining.append((idx, s, e))
                        continue
                    reserved_outs.add(norm_out)
                    lo_groups.setdefault((dst_ext, infilter), []).append(
                        (idx, s, e, src_p, base, dst_ext, src_size, ml_vec)
                    )
                else:
                    remaining.append((idx, s, e))

            if lo_groups:
                from . import debuglog
                dbg_mode = debuglog.get_mode()
                for (dst_ext, infilter), items in lo_groups.items():
                    if len(items) < 2:
                        for item in items:
                            remaining.append((item[0], item[1], item[2]))
                        continue
                    # Split into sub-batches with distinct file stems so LibreOffice
                    # --outdir filenames never collide inside the staging directory.
                    sub_batches: list[list[tuple[int, str, str, Path, Path, str, int, Any]]] = []
                    for item in items:
                        stem_key = item[3].stem.lower()
                        placed = False
                        for sb in sub_batches:
                            if stem_key not in {it[3].stem.lower() for it in sb}:
                                sb.append(item)
                                placed = True
                                break
                        if not placed:
                            sub_batches.append([item])
                    for sb in sub_batches:
                        if len(sb) < 2:
                            remaining.append((sb[0][0], sb[0][1], sb[0][2]))
                            continue
                        pairs = [(it[3], it[4]) for it in sb]
                        t0 = time.perf_counter()
                        ok_map, _ = _lo_convert_batch(pairs, dst_ext.lstrip("."), infilter)
                        per_ms = ((time.perf_counter() - t0) * 1000) / len(sb)
                        for idx, s, e, src_p, base, de, bytes_in, ml_vec in sb:
                            if ok_map.get(src_p) and base.exists():
                                if sel is not None:
                                    if sel_feat:
                                        sel.record(src_p, de, "libreoffice", per_ms, True,
                                                   features=ml_vec, src_size=bytes_in)
                                    else:
                                        sel.record(src_p, de, "libreoffice", per_ms, True)
                                try:
                                    out_bytes = base.stat().st_size
                                except OSError:
                                    out_bytes = 0
                                res = ConversionResult(
                                    str(src_p), str(base), "libreoffice", per_ms, True,
                                    bytes_in=bytes_in, bytes_out=out_bytes,
                                )
                                if dbg_mode == debuglog.MODE_ALWAYS:
                                    sess = debuglog.new_session(src_p, de)
                                    sess.update(
                                        ok=True, winning_backend="libreoffice", total_ms=per_ms,
                                        planned_out=str(base), out_path=str(base),
                                        dst_ext=de, bytes_in=bytes_in, bytes_out=out_bytes,
                                        routes_static=["libreoffice"],
                                        routes=[{"name": "libreoffice", "weight": 0}],
                                        attempts=[{"backend": "libreoffice", "ok": True,
                                                   "ms": per_ms, "error": "", "traceback": ""}],
                                    )
                                    debuglog.finalize(sess, res)
                                results[idx] = res
                                _emit(res)
                            else:
                                remaining.append((idx, s, e))

            if len(remaining) == 1:
                idx, s, e = remaining[0]
                res = self.convert(s, e)
                results[idx] = res
                _emit(res)
            elif remaining:
                fut_map = {self._pool.submit(self.convert, s, e): idx for idx, s, e in remaining}
                for fut in as_completed(fut_map):
                    idx = fut_map[fut]
                    res = fut.result()
                    results[idx] = res
                    _emit(res)
        return [r for r in results if r is not None]

    def shutdown(self) -> None:
        self._pool.shutdown(wait=False)
        _OFFICE_POOL.shutdown()
        if self.cloud is not None:
            try:
                self.cloud.stop()
            except Exception:
                pass


ENGINE_SINGLETON: Optional[ConversionEngine] = None
_ENGINE_LOCK = threading.Lock()


def get_engine(**kw) -> ConversionEngine:
    global ENGINE_SINGLETON
    with _ENGINE_LOCK:
        if ENGINE_SINGLETON is None:
            ENGINE_SINGLETON = ConversionEngine(**kw)
    return ENGINE_SINGLETON
