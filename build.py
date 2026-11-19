#!/usr/bin/env python3
"""One-shot builder: creates the Windows EXE with PyInstaller.
Run on Windows:      python build.py
Produces:            dist/ClickShift-v<ver>.exe   (single file, windowed/no-console)
The same script sanity-checks imports before freezing so you get an early,
clear error instead of a broken EXE.
"""
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
SPEC = ROOT / "clickshift.spec"


def check_imports() -> None:
    sys.path.insert(0, str(SRC))
    import clickshift.engine, clickshift.learning, clickshift.monitor, clickshift.cli, clickshift.cloud, clickshift.debuglog, clickshift.updater  # noqa: F401
    print("imports OK")


def _app_version() -> str:
    sys.path.insert(0, str(SRC))
    from clickshift import __version__
    return __version__


def _write_version_file(ver: str) -> Path:
    parts = [int(p) if p.isdigit() else 0 for p in (ver.split(".") + ["0", "0", "0"])[:4]]
    vtup = f"({parts[0]}, {parts[1]}, {parts[2]}, {parts[3]})"
    vstr = f"{parts[0]}.{parts[1]}.{parts[2]}.{parts[3]}"
    content = f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={vtup},
    prodvers={vtup},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [
          StringStruct('CompanyName', 'ClickShift'),
          StringStruct('FileDescription', 'ClickShift'),
          StringStruct('FileVersion', '{vstr}'),
          StringStruct('InternalName', 'ClickShift'),
          StringStruct('OriginalFilename', 'ClickShift.exe'),
          StringStruct('ProductName', 'ClickShift'),
          StringStruct('ProductVersion', '{ver}')
        ])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""
    out = ROOT / "build" / "version_info.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(content, encoding="utf-8")
    return out


def build() -> int:
    ver = _app_version()
    exe_name = f"ClickShift-v{ver}"
    ver_file = _write_version_file(ver)
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--name", exe_name,
        "--onefile",
        "--noconsole",                      # GUI subsystem: no black console window
        "--version-file", str(ver_file),
        "--specpath", str(ROOT),
        "--paths", str(SRC),
        "--collect-all", "pdf2docx",
        # pdf2docx's native deps were missing inside the frozen EXE (that is
        # why BACKENDS.pdf2docx showed False on users' machines): bundle them.
        "--collect-all", "fitz",             # PyMuPDF (pdf rendering/reflow)
        "--collect-all", "cv2",              # OpenCV (pdf2docx layout analysis)
        "--collect-all", "docx",             # python-docx templates + shared
        "--collect-all", "pptx",             # python-pptx default.pptx template
        "--hidden-import", "numpy",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "PIL.ImageOps",
        "--hidden-import", "PIL.ImageDraw",
        "--hidden-import", "docx",
        "--hidden-import", "docx.shared",
        "--hidden-import", "pptx",
        "--hidden-import", "pptx.util",
        "--hidden-import", "pptx.dml.color",
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
        "--hidden-import", "clickshift.formats",
        "--hidden-import", "clickshift.engine",
        "--hidden-import", "clickshift.learning",
        "--hidden-import", "clickshift.monitor",
        "--hidden-import", "clickshift.cli",
        "--hidden-import", "clickshift.cloud",
        "--hidden-import", "clickshift.debuglog",
        "--hidden-import", "clickshift.updater",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
    ]
    assets_dir = ROOT / "assets"
    if assets_dir.exists():
        args += ["--add-data", f"{assets_dir}{os.pathsep}assets"]
    icon = assets_dir / "clickshift.ico"
    if icon.exists():
        args += ["--icon", str(icon)]
    args.append(str(ROOT / "entry_point.py"))
    print("$", " ".join(args))
    return subprocess.call(args)


if __name__ == "__main__":
    check_imports()
    raise SystemExit(build())
