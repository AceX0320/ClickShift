"""Lightweight conversion metadata & result types (zero third-party imports).

Keeping format tables, extension helpers, and ``ConversionResult`` isolated
from ``clickshift.engine`` allows ``cli``, ``monitor``, and IPC clients to run
without importing heavy document/ML/imaging backends.
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path


def _norm_ext(ext: str) -> str:
    ext = (ext or "").strip().lower()
    if ext and not ext.startswith("."):
        ext = "." + ext
    return ext


@dataclass
class ConversionResult:
    source: str
    target: str
    backend: str
    elapsed_ms: float
    ok: bool
    error: str = ""
    bytes_in: int = 0
    bytes_out: int = 0

    def as_dict(self) -> dict:
        return {
            "source": self.source,
            "target": self.target,
            "backend": self.backend,
            "elapsed_ms": round(self.elapsed_ms, 1),
            "ok": self.ok,
            "error": self.error,
            "bytes_in": self.bytes_in,
            "bytes_out": self.bytes_out,
            "t": time.time(),
        }


OFFICE_DOC_IN = {"doc", "docx", "odt", "rtf", "mht", "html", "htm", "txt", "wpd"}
OFFICE_SHEET_IN = {"xls", "xlsx", "ods", "csv", "txt"}
OFFICE_SLIDE_IN = {"ppt", "pptx", "odp"}
PDF_CAPABLE_IN = OFFICE_DOC_IN | OFFICE_SLIDE_IN | OFFICE_SHEET_IN | {"png", "jpg", "jpeg", "tif", "tiff", "epub", "xps"}
IMAGE_EXTS = {
    "png", "jpg", "jpeg", "gif", "bmp", "tiff", "tif", "webp",
    "avif", "heic", "heif", "jxl", "ico", "ppm", "pgm", "pbm", "jp2",
}
TEXT_EXTS = {
    "txt", "md", "markdown", "csv", "tsv", "json", "xml", "html", "htm",
    "log", "ini", "py", "js", "c", "cpp", "h", "java", "yml", "yaml", "srt",
}
SUPPORTED_INPUT_EXTS = (
    OFFICE_DOC_IN
    | OFFICE_SHEET_IN
    | OFFICE_SLIDE_IN
    | PDF_CAPABLE_IN
    | IMAGE_EXTS
    | TEXT_EXTS
    | {"pdf", "epub", "xps", "wmf", "emf"}
)

PAGED_IN = {
    "pdf", "doc", "docx", "odt", "rtf", "ppt", "pptx", "odp",
    "xls", "xlsx", "ods", "epub", "xps",
}
PAGE_IMAGE_EXTS = {"png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff", "gif", "avif", "jxl"}

NEWEST_GENERATION = {
    "pdf": "docx",
    "doc": "docx",
    "xls": "xlsx",
    "ppt": "pptx",
    "odt": "docx",
    "ods": "xlsx",
    "odp": "pptx",
    "rtf": "docx",
    "htm": "docx",
    "html": "docx",
    "wmf": "png",
    "emf": "png",
}


@lru_cache(maxsize=128)
def _cached_targets_tuple(src_ext: str) -> tuple[str, ...]:
    src_ext = src_ext.lower().lstrip(".")
    out: list[str] = []
    g = NEWEST_GENERATION.get(src_ext)
    if g:
        out.append(g)
    if src_ext == "pdf":
        out += ["docx", "pptx", "txt", "md", "html", "csv", "png", "jpg", "webp", "odt", "rtf"]
    if src_ext in OFFICE_DOC_IN:
        out += ["pdf", "docx", "pptx", "txt", "md", "html", "rtf", "odt", "png", "jpg", "webp"]
    if src_ext in OFFICE_SHEET_IN or src_ext == "tsv":
        out += ["xlsx", "pdf", "csv", "ods", "txt", "html", "png", "jpg", "webp"]
    if src_ext in OFFICE_SLIDE_IN:
        out += ["pptx", "pdf", "docx", "odp", "png", "jpg", "webp"]
    if src_ext in IMAGE_EXTS or src_ext in ("wmf", "emf"):
        out += ["webp", "png", "jpg", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp", "ico", "gif"]
    if src_ext in TEXT_EXTS:
        out += ["docx", "pdf", "txt", "md", "html", "rtf", "odt"]
    if src_ext == "epub":
        out += ["txt", "docx", "html", "md", "pdf", "png", "jpg", "webp"]
    if src_ext == "xps":
        out += ["pdf", "docx", "png", "jpg"]
    seen: set[str] = set()
    res: list[str] = []
    for e in out:
        if e != src_ext and e not in seen:
            seen.add(e)
            res.append(e)
    return tuple(res)


def all_targets_for(src_ext: str) -> list[str]:
    """Return ALL supported conversion target extensions for ``src_ext`` (ordered best-first)."""
    return list(_cached_targets_tuple(src_ext))


def newest_target_for(src_ext: str) -> list[str]:
    """Suggested 'latest version' targets for a given source extension."""
    return all_targets_for(src_ext)


def is_pages_to_images(src: Path, dst_ext: str) -> bool:
    return (
        src.suffix.lstrip(".").lower() in PAGED_IN
        and dst_ext.lstrip(".").lower() in PAGE_IMAGE_EXTS
    )
