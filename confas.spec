# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for ConFas.exe  (build with:  pyinstaller confas.spec)
from pathlib import Path

ROOT = Path(SPECPATH)

a = Analysis(
    [str(ROOT / "entry_point.py")],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=[],
    hiddenimports=[
        "confas.engine", "confas.learning", "confas.monitor", "confas.cli",
        "confas.cloud", "confas.debuglog",
        # Dynamically imported backends and Windows-only modules:
        "win32com.client", "pythoncom", "pywintypes",
        "pystray", "pystray._win32",
        "PIL", "PIL.Image", "PIL.ImageOps", "PIL.ImageDraw",
        "numpy", "docx", "docx.shared", "pdf2docx", "pypdf", "fitz",
        "chardet", "charset_normalizer",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "PyQt5", "PySide6"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ConFas",
    debug=False,
    strip=False,
    upx=False,
    console=False,          # windowed app -> no black console flash
    icon=str(ROOT / "assets" / "confas.ico") if (ROOT / "assets" / "confas.ico").exists() else None,
)
