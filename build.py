#!/usr/bin/env python3
"""One-shot builder: creates the Windows EXE with PyInstaller.
Run on Windows:      python build.py
Produces:            dist/ConFas.exe   (single file, windowed/no-console)
The same script sanity-checks imports before freezing so you get an early,
clear error instead of a broken EXE.
"""
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
SPEC = ROOT / "confas.spec"


def check_imports() -> None:
    sys.path.insert(0, str(SRC))
    import confas.engine, confas.learning, confas.monitor, confas.cli, confas.cloud, confas.debuglog, confas.updater  # noqa: F401
    print("imports OK")


def _app_version() -> str:
    sys.path.insert(0, str(SRC))
    from confas import __version__
    return __version__


def build() -> int:
    ver = _app_version()
    exe_name = f"ConFas-v{ver}"
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--name", exe_name,
        "--onefile",
        "--noconsole",                      # GUI subsystem: no black console window
        "--specpath", str(ROOT),
        "--paths", str(SRC),
        "--collect-all", "pdf2docx",
        # pdf2docx's native deps were missing inside the frozen EXE (that is
        # why BACKENDS.pdf2docx showed False on users' machines): bundle them.
        "--collect-all", "fitz",             # PyMuPDF (pdf rendering/reflow)
        "--collect-all", "cv2",              # OpenCV (pdf2docx layout analysis)
        "--collect-all", "docx",             # python-docx templates + shared
        "--hidden-import", "numpy",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "PIL.ImageOps",
        "--hidden-import", "PIL.ImageDraw",
        "--hidden-import", "docx",
        "--hidden-import", "docx.shared",
        "--hidden-import", "pdf2docx",
        "--hidden-import", "pypdf",
        "--hidden-import", "fitz",
        "--hidden-import", "chardet",
        "--hidden-import", "charset_normalizer",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--hidden-import", "pystray",
        "--hidden-import", "pystray._win32",
        "--hidden-import", "confas.engine",
        "--hidden-import", "confas.learning",
        "--hidden-import", "confas.monitor",
        "--hidden-import", "confas.cli",
        "--hidden-import", "confas.cloud",
        "--hidden-import", "confas.debuglog",
        "--hidden-import", "confas.updater",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
    ]
    assets_dir = ROOT / "assets"
    if assets_dir.exists():
        args += ["--add-data", f"{assets_dir}{os.pathsep}assets"]
    icon = assets_dir / "confas.ico"
    if icon.exists():
        args += ["--icon", str(icon)]
    args.append(str(ROOT / "entry_point.py"))
    print("$", " ".join(args))
    return subprocess.call(args)


if __name__ == "__main__":
    check_imports()
    raise SystemExit(build())
