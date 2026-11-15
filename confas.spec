# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for ConFas.exe  (build with:  pyinstaller confas.spec)
import re
from pathlib import Path

ROOT = Path(SPECPATH)
_init_text = (ROOT / "src" / "confas" / "__init__.py").read_text(encoding="utf-8")
_ver_match = re.search(r'__version__\s*=\s*"([^"]+)"', _init_text)
EXE_NAME = f"ConFas-v{_ver_match.group(1)}" if _ver_match else "ConFas"

a = Analysis(
    [str(ROOT / "entry_point.py")],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=[(str(ROOT / "assets"), "assets")] if (ROOT / "assets").exists() else [],
    hiddenimports=[
        "confas.engine", "confas.learning", "confas.monitor", "confas.cli",
        "confas.cloud", "confas.debuglog", "confas.updater",
        # Dynamically imported backends and Windows-only modules:
        "win32com.client", "pythoncom", "pywintypes",
        "pystray", "pystray._win32",
        "tkinter", "tkinter.ttk",
        "PIL", "PIL.Image", "PIL.ImageOps", "PIL.ImageDraw",
        "numpy", "docx", "docx.shared", "pdf2docx", "pypdf", "fitz",
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
    icon=str(ROOT / "assets" / "confas.ico") if (ROOT / "assets" / "confas.ico").exists() else None,
)
