# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for ClickShift.exe  (build with:  pyinstaller clickshift.spec)
import re
from pathlib import Path

ROOT = Path(SPECPATH)
_init_text = (ROOT / "src" / "clickshift" / "__init__.py").read_text(encoding="utf-8")
_ver_match = re.search(r'__version__\s*=\s*"([^"]+)"', _init_text)
EXE_NAME = f"ClickShift-v{_ver_match.group(1)}" if _ver_match else "ClickShift"

a = Analysis(
    [str(ROOT / "entry_point.py")],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=[(str(ROOT / "assets"), "assets")] if (ROOT / "assets").exists() else [],
    hiddenimports=[
        "clickshift.engine", "clickshift.learning", "clickshift.monitor", "clickshift.cli",
        "clickshift.cloud", "clickshift.debuglog", "clickshift.updater",
        # Dynamically imported backends and Windows-only modules:
        "win32com.client", "pythoncom", "pywintypes",
        "pystray", "pystray._win32",
        "tkinter", "tkinter.ttk",
        "PIL", "PIL.Image", "PIL.ImageOps", "PIL.ImageDraw",
        "numpy", "docx", "docx.shared", "pptx", "pptx.util", "pptx.dml.color",
        "pdf2docx", "pypdf", "fitz",
        "chardet", "charset_normalizer",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "PyQt5", "PySide6"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    strip=False,
    upx=False,
    console=False,          # windowed app -> no black console flash
    icon=str(ROOT / "assets" / "clickshift.ico") if (ROOT / "assets" / "clickshift.ico").exists() else None,
)
