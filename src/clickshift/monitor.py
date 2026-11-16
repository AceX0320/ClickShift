r"""ClickShift always-on services:

Tray daemon          - runs forever, shows an icon in the notification area
                       (Task Manager > Startup apps friendly), lets you
                       pause/quit and open your watch folder.
Watch-folder         - drop any document into `%USERPROFILE%/ClickShiftDrop` (or a folder
watcher              you pick) and it is converted automatically to its
                       "latest generation" format using the engine + ML.
Startup registration - `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` entry
                       so ClickShift starts with Windows (appears under
                       Task Manager -> Startup apps).
Context-menu         - per-extension shell keys (HKCU\Software\Classes) that add
installer              "Convert with ClickShift" + one "Convert to <FMT> with ClickShift"
                       row per sensible target to the Explorer right-click menu
                       (classic menu / "Show more options" on Win11), and also
                       registers ClickShift under "Open with" so it is reachable
                       from the Windows 11 first-level menu.  Verbs are written
                       to *every* location modern Windows consults:
                         1. the effective ProgID (FileExts UserChoice),
                         2. SystemFileAssociations\.ext,
                         3. the raw .ext class key (legacy).
No admin rights required, no external .dll handlers.
Everything degrades gracefully on non-Windows systems (the pure-python parts
are still importable there for development & CI testing).
"""
# pyright: reportMissingImports=false, reportMissingModuleSource=false
from __future__ import annotations

import ctypes
import importlib
import json
import os
import socket
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path
from typing import Any, Optional

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    __package__ = "clickshift"

IS_WINDOWS = sys.platform == "win32"


def _opt_import(name: str) -> Any:
    try:
        return importlib.import_module(name)
    except Exception:  # pragma: no cover
        return None


winreg: Any = _opt_import("winreg") if IS_WINDOWS else None
APP_NAME = "ClickShift"

DATA_DIR = Path(os.environ.get("CLICKSHIFT_HOME", str(Path.home() / ".clickshift")))


def data_dir() -> Path:
    """Resolve the data dir at call time (honours CLICKSHIFT_HOME changes/tests)."""
    return Path(os.environ.get("CLICKSHIFT_HOME", str(Path.home() / ".clickshift")))


DROP_DIR = Path(os.environ.get("CLICKSHIFT_DROP", str(Path.home() / "ClickShiftDrop")))
CONFIG_FILE = DATA_DIR / "config.json"
LOG_FILE = DATA_DIR / "daemon.log"
PID_FILE = DATA_DIR / "daemon.pid"

# extensions that get right-click entries
MENU_EXTENSIONS = [
    "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "ods", "odp",
    "rtf", "txt", "md", "markdown", "csv", "tsv", "html", "htm", "epub", "xps",
    "mht", "wpd", "json", "xml", "yml", "yaml", "ini", "log", "srt",
    "png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp", "avif", "heic",
    "heif", "jxl", "ico", "wmf", "emf", "ppm", "pgm", "pbm", "jp2",
]

# Full target choices offered per extension in the right-click menus
_FAMILY_TARGETS = {
    "pdf": ["docx", "pptx", "txt", "md", "html", "csv", "png", "jpg", "webp", "odt", "rtf"],
    "doc": ["pdf", "docx", "pptx", "txt", "md", "html", "rtf", "odt", "png", "jpg", "webp"],
    "docx": ["pdf", "pptx", "txt", "md", "html", "rtf", "odt", "png", "jpg", "webp"],
    "odt": ["docx", "pdf", "pptx", "txt", "md", "html", "rtf", "png", "jpg", "webp"],
    "rtf": ["docx", "pdf", "pptx", "txt", "md", "html", "odt", "png", "jpg", "webp"],
    "mht": ["docx", "pdf", "html", "txt"],
    "wpd": ["docx", "pdf", "txt", "rtf", "odt"],
    "xls": ["xlsx", "pdf", "csv", "ods", "txt", "html", "png", "jpg", "webp"],
    "xlsx": ["pdf", "csv", "ods", "txt", "html", "png", "jpg", "webp"],
    "ods": ["xlsx", "pdf", "csv", "txt", "html", "png", "jpg", "webp"],
    "csv": ["xlsx", "pdf", "ods", "txt", "html", "docx"],
    "tsv": ["xlsx", "pdf", "csv", "ods", "txt", "html"],
    "ppt": ["pptx", "pdf", "docx", "odp", "png", "jpg", "webp"],
    "pptx": ["pdf", "docx", "odp", "png", "jpg", "webp"],
    "odp": ["pptx", "pdf", "docx", "png", "jpg", "webp"],
    "txt": ["docx", "pdf", "md", "html", "rtf", "odt"],
    "md": ["docx", "pdf", "html", "txt", "rtf", "odt"],
    "markdown": ["docx", "pdf", "html", "txt", "rtf", "odt"],
    "html": ["docx", "pdf", "txt", "md", "odt", "rtf"],
    "htm": ["docx", "pdf", "txt", "md", "odt", "rtf"],
    "epub": ["docx", "pdf", "txt", "html", "md", "png", "jpg", "webp"],
    "xps": ["pdf", "docx", "png", "jpg"],
    "png": ["webp", "jpg", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp", "ico", "gif"],
    "jpg": ["webp", "png", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp", "ico", "gif"],
    "jpeg": ["webp", "png", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp", "ico", "gif"],
    "gif": ["webp", "png", "jpg", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp", "ico"],
    "bmp": ["png", "webp", "jpg", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "ico", "gif"],
    "tif": ["png", "webp", "jpg", "pdf", "docx", "pptx", "avif", "jxl", "bmp", "ico", "gif"],
    "tiff": ["png", "webp", "jpg", "pdf", "docx", "pptx", "avif", "jxl", "bmp", "ico", "gif"],
    "webp": ["png", "jpg", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp", "ico", "gif"],
    "avif": ["png", "jpg", "webp", "pdf", "docx", "pptx", "jxl", "tiff", "bmp", "ico", "gif"],
    "heic": ["jpg", "png", "webp", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp"],
    "heif": ["jpg", "png", "webp", "pdf", "docx", "pptx", "avif", "jxl", "tiff", "bmp"],
    "jxl": ["png", "jpg", "webp", "pdf", "docx", "pptx", "avif", "tiff", "bmp", "ico"],
    "ico": ["png", "webp", "jpg", "pdf", "docx", "pptx", "bmp"],
    "wmf": ["png", "webp", "jpg", "pdf", "docx", "pptx"],
    "emf": ["png", "webp", "jpg", "pdf", "docx", "pptx"],
}


def menu_targets(ext: str) -> list[str]:
    ext_clean = ext.lower().lstrip(".")
    if ext_clean in _FAMILY_TARGETS:
        return list(_FAMILY_TARGETS[ext_clean])
    try:
        from .engine import all_targets_for
        targets = all_targets_for(ext_clean)
        if targets:
            return targets
    except Exception:
        pass
    return ["pdf", "docx", "txt"]


# --------------------------------------------------------------------------- #
#  Config
# --------------------------------------------------------------------------- #

DEFAULT_CONFIG = {
    "watch_folder": str(DROP_DIR),
    "watch_enabled": True,
    "auto_target": "pdf",          # what dropped files are converted to by default
    "ml_enabled": True,
    "notify": True,
    "win11_both_menus": True,      # surface full ClickShift menu on Win11 1st & 2nd right-click menus
    # --- opt-in diagnostic logging (off by default: no *.clickshift-debug.txt) #
    "debug_reports": "off",         # off | onerror | always
    "debug_keep": 200,              # central reports kept in ~/.clickshift/debug/
    # --- shared-learnings cloud sync (opt-in, aggregate stats only) ------ #
    "cloud_enabled": False,         # enable with: clickshift config cloud_sync=on
    "cloud_mode": "gist",           # "gist" (zero-backend) or "url" (own server)
    "cloud_gist_id": "REPLACE_WITH_YOUR_GIST_ID",
    "cloud_url": "",                # used when cloud_mode=url
    "cloud_interval_s": 900,        # background push/pull cadence
    "cloud_debounce_s": 60,         # min seconds between uploads
    "cloud_pull_on_start": True,
    # --- GitHub auto-updater --------------------------------------------- #
    "auto_update_check": True,
    "update_check_interval_s": 43200,   # check every 12 hours
    "last_update_check_ts": 0.0,
    "skipped_update_version": "",
}


def config_file() -> Path:
    """Resolve the config file path at call time (honours CLICKSHIFT_HOME changes/tests)."""
    return data_dir() / "config.json"


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    try:
        cfg.update(json.loads(config_file().read_text(encoding="utf-8")))
    except FileNotFoundError:
        pass
    except Exception:
        pass
    return cfg


def save_config(cfg: dict) -> None:
    cfile = config_file()
    cfile.parent.mkdir(parents=True, exist_ok=True)
    tmp = cfile.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    tmp.replace(cfile)


def exe_path() -> str:
    """Path of the running executable (frozen) or python launcher."""
    if getattr(sys, "frozen", False):
        return sys.executable
    return sys.executable


def launch_command(args: list[str]) -> list[str]:
    """Command list that re-launches this program with given args."""
    if getattr(sys, "frozen", False):
        return [sys.executable] + args
    return [sys.executable, "-m", "clickshift"] + args


def _log(msg: str) -> None:
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n"
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(line)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
#  Startup registration  (Task Manager > Startup apps)
# --------------------------------------------------------------------------- #

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def register_startup() -> bool:
    """Add ClickShift to the current user's Run key -> shows up in Task Manager."""
    if not IS_WINDOWS or winreg is None:
        return False
    try:
        cmd = subprocess.list2cmdline(launch_command(["run"]))
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                            winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, cmd)
            try:
                winreg.DeleteValue(k, "ConFas")
            except OSError:
                pass
        _log(f"startup registered: {cmd}")
        return True
    except Exception as exc:
        _log(f"register_startup failed: {exc}")
        return False


def unregister_startup() -> bool:
    if not IS_WINDOWS or winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                            winreg.KEY_SET_VALUE) as k:
            for name in (APP_NAME, "ConFas"):
                try:
                    winreg.DeleteValue(k, name)
                except FileNotFoundError:
                    pass
        return True
    except FileNotFoundError:
        return True
    except Exception:
        return False


def is_startup_registered() -> bool:
    if not IS_WINDOWS or winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, APP_NAME)
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------- #
#  Right-click context menu (Cascading submenu + Full verbs + Win11 both menus)
# --------------------------------------------------------------------------- #

VERB_MENU = "ClickShift.Menu"
VERB_AUTO = "ClickShift.Convert"
VERB_CHOOSE = "ClickShift.Choose"
OPENWITH_APP = "ClickShift.exe"
WIN11_CLASSIC_CLSID = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
UNIVERSAL_TARGETS = ["pdf", "docx", "xlsx", "pptx", "txt", "md", "html", "png", "jpg", "webp", "csv"]


def _verb_cmd(target: Optional[str], choose: bool = False) -> str:
    args = ["convert"]
    if choose:
        args.append("--choose")
    args.append("%1")
    if target:
        args += ["-t", target]
    return subprocess.list2cmdline(launch_command(args))


def _user_choice_progid(ext: str) -> str:
    """ProgID Windows actually uses for .ext (FileExts UserChoice)."""
    if not IS_WINDOWS or winreg is None:
        return ""
    try:
        path = (rf"Software\Microsoft\Windows\CurrentVersion"
                rf"\Explorer\FileExts\.{ext}\UserChoice")
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
            return str(winreg.QueryValueEx(k, "ProgId")[0] or "").strip()
    except Exception:
        return ""


def _legacy_progid(ext: str) -> str:
    """ProgID stored as default value of the .ext class key (HKCU or HKCR)."""
    if not IS_WINDOWS or winreg is None:
        return ""
    try:
        for hive, sub in (
            (winreg.HKEY_CURRENT_USER, rf"Software\Classes\.{ext}"),
            (winreg.HKEY_CLASSES_ROOT, rf".{ext}"),
        ):
            try:
                with winreg.OpenKey(hive, sub) as k:
                    val = str(winreg.QueryValue(k, "") or "").strip()
                    if val:
                        return val
            except OSError:
                continue
    except Exception:
        pass
    return ""


def _openwith_progids(ext: str) -> list[str]:
    """All ProgIDs registered under OpenWithProgids for .ext (e.g. MSEdgePDF, Word.Document.12)."""
    if not IS_WINDOWS or winreg is None:
        return []
    out: list[str] = []
    try:
        for hive, sub in (
            (winreg.HKEY_CURRENT_USER,
             rf"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.{ext}\OpenWithProgids"),
            (winreg.HKEY_CLASSES_ROOT, rf".{ext}\OpenWithProgids"),
        ):
            try:
                with winreg.OpenKey(hive, sub) as k:
                    i = 0
                    while True:
                        try:
                            name, _val, _typ = winreg.EnumValue(k, i)
                            i += 1
                            if name and name not in out:
                                out.append(name)
                        except OSError:
                            break
            except OSError:
                continue
    except Exception:
        pass
    return out


def _verb_locations(ext: str) -> list[str]:
    """All registry roots modern Windows consults for per-extension verbs."""
    targets = [rf"Software\Classes\SystemFileAssociations\.{ext}"]
    progids = [_user_choice_progid(ext), _legacy_progid(ext)] + _openwith_progids(ext)
    for progid in progids:
        candidate = rf"Software\Classes\{progid}" if progid else ""
        if candidate and candidate not in targets:
            targets.append(candidate)
    targets.append(rf"Software\Classes\.{ext}")
    return targets


def _write_verb(base: str, verb: str, label: str, cmd: str) -> None:
    assert winreg is not None
    key = base + rf"\shell\{verb}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key) as k:
        winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, label)
        winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, exe_path() + ",0")
        with winreg.CreateKey(k, "command") as kc:
            winreg.SetValue(kc, "", winreg.REG_SZ, cmd)


def _write_cascading_menu(base: str, targets: list[str], default_target: Optional[str] = None) -> None:
    """Write a cascading 'Convert with ClickShift' flyout menu containing ALL target choices."""
    assert winreg is not None
    menu_key = base + rf"\shell\{VERB_MENU}"
    # Clean any stale subcommands first (including legacy ConFas menu if upgrading)
    _delete_key_tree(winreg.HKEY_CURRENT_USER, base + r"\shell\ConFas.Menu")
    _delete_key_tree(winreg.HKEY_CURRENT_USER, menu_key)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, menu_key) as km:
        winreg.SetValueEx(km, "MUIVerb", 0, winreg.REG_SZ, "Convert with ClickShift")
        winreg.SetValueEx(km, "Icon", 0, winreg.REG_SZ, exe_path() + ",0")
        winreg.SetValueEx(km, "SubCommands", 0, winreg.REG_SZ, "")
        winreg.SetValueEx(km, "Position", 0, winreg.REG_SZ, "Top")

    # 00: Auto (latest generation format)
    auto_lbl = (f"Auto — Convert to {default_target.upper()} (latest format)"
                if default_target else "Auto — Convert to latest format")
    _write_verb(menu_key, "00_Auto", auto_lbl, _verb_cmd(None))

    # 01..N: Every single target choice for this extension
    for idx, t in enumerate(targets, start=1):
        sub_verb = f"{idx:02d}_To_{t}"
        _write_verb(menu_key, sub_verb, f"Convert to {t.upper()}", _verb_cmd(t))

    # 99: Interactive format chooser popup
    _write_verb(menu_key, "99_Choose", "Choose format...", _verb_cmd(None, choose=True))


def _delete_verb(base: str, verb: str) -> bool:
    assert winreg is not None
    key = base + rf"\shell\{verb}"
    if not _key_exists(winreg.HKEY_CURRENT_USER, key):
        return False
    _delete_key_tree(winreg.HKEY_CURRENT_USER, key)
    return True


def _key_exists(root, path: str) -> bool:
    assert winreg is not None
    try:
        with winreg.OpenKey(root, path):
            return True
    except OSError:
        return False


def _verbs_for(ext: str) -> list[tuple[str, str, str]]:
    """(verb key name, menu label, command) tuples for one extension."""
    targets = menu_targets(ext)
    best = targets[0].upper() if targets else "LATEST"
    out = [
        (VERB_AUTO, f"Convert with ClickShift (Auto -> {best})", _verb_cmd(None)),
    ]
    for t in targets:
        out.append((f"ClickShift.To{t.capitalize()}",
                    f"Convert to {t.upper()} with ClickShift",
                    _verb_cmd(t)))
    out.append((VERB_CHOOSE, "Convert with ClickShift (Choose format...)", _verb_cmd(None, choose=True)))
    return out


def _delete_key_tree(root, path: str) -> None:
    assert winreg is not None
    try:
        with winreg.OpenKey(root, path) as k:
            subs = []
            i = 0
            while True:
                try:
                    subs.append(winreg.EnumKey(k, i))
                    i += 1
                except OSError:
                    break
        for s in subs:
            _delete_key_tree(root, path + "\\" + s)
        winreg.DeleteKey(root, path)
    except FileNotFoundError:
        pass
    except OSError:
        pass


def _register_openwith(ext: str) -> None:
    """Make ClickShift appear in Explorer's 'Open with' (Windows 11 first menu) with all choices."""
    assert winreg is not None
    app = OPENWITH_APP
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                          rf"Software\Classes\Applications\{app}") as k:
        winreg.SetValueEx(k, "FriendlyAppName", 0, winreg.REG_SZ,
                          "ClickShift — Convert Document (All Formats)")
        winreg.SetValueEx(k, "ApplicationDescription", 0, winreg.REG_SZ,
                          "Ultra-fast local document converter")
        with winreg.CreateKey(k, "DefaultIcon") as ki:
            winreg.SetValue(ki, "", winreg.REG_SZ, exe_path() + ",0")
        # Launching via 'Open with' on Win11's 1st menu opens the native cursor popup
        # listing every available conversion target for the clicked file!
        with winreg.CreateKey(k, r"shell\open\command") as kc:
            winreg.SetValue(kc, "", winreg.REG_SZ, _verb_cmd(None, choose=True))
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                          rf"Software\Classes\.{ext}\OpenWithList") as kol:
        winreg.SetValueEx(kol, app, 0, winreg.REG_SZ, "")


def _unregister_openwith(ext: str) -> None:
    assert winreg is not None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            rf"Software\Classes\.{ext}\OpenWithList", 0,
                            winreg.KEY_SET_VALUE) as kol:
            for app in (OPENWITH_APP, "ConFas.exe"):
                try:
                    winreg.DeleteValue(kol, app)
                except OSError:
                    pass
    except Exception:
        pass


def _set_win11_both_menus(enable: bool) -> None:
    """Ensure Windows 11 surfaces full shell context menu verbs on primary right-click too."""
    if not IS_WINDOWS or winreg is None:
        return
    try:
        if enable:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                                  WIN11_CLASSIC_CLSID + r"\InprocServer32") as k:
                winreg.SetValue(k, "", winreg.REG_SZ, "")
        else:
            _delete_key_tree(winreg.HKEY_CURRENT_USER, WIN11_CLASSIC_CLSID)
    except Exception as exc:
        _log(f"win11 menu bridge update failed: {exc}")


def _assets_dir() -> Path:
    """Locate bundled assets directory (works both frozen in _MEIPASS and from source)."""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        cand = Path(meipass) / "assets"
        if cand.exists():
            return cand
    return Path(__file__).resolve().parents[2] / "assets"


def _save_exe_path_registry() -> None:
    """Store the current ClickShift executable path in HKCU\\Software\\ClickShift for ClickShiftShell.dll."""
    if not IS_WINDOWS or winreg is None:
        return
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\ClickShift") as k:
            winreg.SetValueEx(k, "ExePath", 0, winreg.REG_SZ, exe_path())
        _delete_key_tree(winreg.HKEY_CURRENT_USER, r"Software\ConFas")
    except Exception as exc:
        _log(f"save ExePath registry failed: {exc}")


def _install_win11_shell_extension() -> bool:
    """Install the packaged IExplorerCommand shell extension (ClickShiftShell.dll) so
    the expandable 'Convert with ClickShift >' submenu appears in the Windows 11 main menu.
    """
    if not IS_WINDOWS:
        return False
    try:
        import shutil
        src_assets = _assets_dir()
        dll_src = src_assets / "ClickShiftShell.dll"
        manifest_src = src_assets / "AppxManifest.xml"
        if not dll_src.exists() or not manifest_src.exists():
            _log(f"Win11 shell extension assets not found in {src_assets}")
            return False

        local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        shell_dir = local_app / "ClickShift" / "win11_shell"
        shell_dir.mkdir(parents=True, exist_ok=True)

        for fname in ("ClickShiftShell.dll", "ClickShiftStub.exe", "AppxManifest.xml",
                      "clickshift.png", "clickshift.ico", "ClickShiftShell.msix", "ClickShiftShell.cer"):
            s = src_assets / fname
            if s.exists():
                try:
                    shutil.copy2(s, shell_dir / fname)
                except Exception:
                    pass

        msix_path = shell_dir / "ClickShiftShell.msix"
        cer_path = shell_dir / "ClickShiftShell.cer"
        manifest_path = shell_dir / "AppxManifest.xml"

        # 1. Remove any stale version (or legacy ConFas extension) and try direct registration
        prep_remove = (
            'Get-AppxPackage -Name "ClickShift.ShellExtension" -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue; '
            'Get-AppxPackage -Name "ConFas.ShellExtension" -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue; '
        )
        direct_cmd = prep_remove + (
            f'Add-AppxPackage -Path "{msix_path}" -ForceApplicationShutdown -ErrorAction Stop'
            if msix_path.exists()
            else f'Add-AppxPackage -Register "{manifest_path}" -ForceApplicationShutdown -ErrorAction Stop'
        )
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", direct_cmd],
            capture_output=True, text=True, timeout=25,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if r.returncode == 0:
            _log("Win11 IExplorerCommand shell extension registered directly")
            return True

        # 2. One-time elevated trust + registration (enables sideloading/dev-unlock + trusts cert + registers)
        elev_parts = [
            'New-Item -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AppModelUnlock" -Force | Out-Null',
            'Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AppModelUnlock" -Name "AllowDevelopmentWithoutDevLicense" -Value 1 -Type DWord -Force',
            'Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AppModelUnlock" -Name "AllowAllTrustedApps" -Value 1 -Type DWord -Force',
        ]
        if cer_path.exists():
            elev_parts.append(
                f'Import-Certificate -FilePath "{cer_path}" -CertStoreLocation "Cert:\\LocalMachine\\TrustedPeople" | Out-Null'
            )
            elev_parts.append(
                f'Import-Certificate -FilePath "{cer_path}" -CertStoreLocation "Cert:\\LocalMachine\\Root" | Out-Null'
            )
        elev_Scripts = "; ".join(elev_parts)
        outer_cmd = (
            f'Start-Process powershell -Verb RunAs -WindowStyle Hidden -Wait '
            f'-ArgumentList \'-NoProfile -WindowStyle Hidden -Command "{elev_Scripts}"\'; '
            f'{direct_cmd}'
        )
        r2 = subprocess.run(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", outer_cmd],
            capture_output=True, text=True, timeout=45,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if r2.returncode == 0:
            _log("Win11 IExplorerCommand shell extension registered after one-time trust setup")
            return True
        _log(f"Win11 shell extension registration stderr: {r2.stderr.strip() or r.stderr.strip()}")
    except Exception as exc:
        _log(f"Win11 shell extension install error: {exc}")
    return False


def _remove_win11_shell_extension() -> None:
    if not IS_WINDOWS:
        return
    try:
        cmd = (
            'Get-AppxPackage -Name "ClickShift.ShellExtension" -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue; '
            'Get-AppxPackage -Name "ConFas.ShellExtension" -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue'
        )
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", cmd],
            capture_output=True, text=True, timeout=20,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        pass


def _purge_duplicate_class_verbs() -> None:
    """Remove any ClickShift* / ConFas* verbs from HKCU\\Software\\Classes\\<subkey>\\shell
    (outside SystemFileAssociations) so the classic menu never shows duplicate entries.
    """
    if not IS_WINDOWS or winreg is None:
        return
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Classes") as root:
            subs: list[str] = []
            i = 0
            while True:
                try:
                    subs.append(winreg.EnumKey(root, i))
                    i += 1
                except OSError:
                    break
        for sub in subs:
            if sub.lower() == "systemfileassociations":
                continue
            sh_path = rf"Software\Classes\{sub}\shell"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sh_path) as sh_key:
                    verbs: list[str] = []
                    j = 0
                    while True:
                        try:
                            v = winreg.EnumKey(sh_key, j)
                            j += 1
                            if v.lower().startswith(("clickshift", "confas")):
                                verbs.append(v)
                        except OSError:
                            break
                for v in verbs:
                    _delete_key_tree(winreg.HKEY_CURRENT_USER, rf"{sh_path}\{v}")
            except OSError:
                continue
    except Exception:
        pass


def install_context_menu(
    extensions: Optional[list[str]] = None,
    progress_cb: Optional[Any] = None,
) -> int:
    """Create the single expandable 'Convert with ClickShift >' menu on both the
    Windows 11 primary context menu (via ClickShiftShell.dll IExplorerCommand) and
    the secondary classic menu (via cascading SubCommands), removing any
    duplicate flat or ProgID verb entries.
    Returns number of extensions handled.
    """
    if not IS_WINDOWS:
        return 0

    def _report(status: str, pct: float) -> None:
        if progress_cb is not None:
            try:
                progress_cb(status, pct)
            except Exception:
                pass

    _report("Saving executable path for Explorer...", 22)
    _save_exe_path_registry()
    _report("Cleaning up duplicate shell entries...", 30)
    _purge_duplicate_class_verbs()
    if winreg is not None:
        _delete_key_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Applications\ConFas.exe")
    # Windows 11: the packaged IExplorerCommand extension shows ONE expandable
    # 'Convert with ClickShift >' in the main menu AND automatically in
    # 'Show more options'. Writing the registry cascade as well would make the
    # classic menu show it twice, so the registry cascade is only a fallback
    # (Windows 10, or when the package could not be registered).
    _set_win11_both_menus(False)
    _report("Registering Windows 11 shell extension...", 42)
    shell_ext_ok = _is_win11() and _install_win11_shell_extension()
    exts = extensions or MENU_EXTENSIONS
    total_exts = max(1, len(exts))
    n = 0
    for idx, ext in enumerate(exts):
        if idx % 6 == 0 or idx == total_exts - 1:
            pct = 55.0 + (idx / total_exts) * 30.0
            _report(f"Configuring right-click actions (.{ext})...", pct)
        locations = _verb_locations(ext)
        primary_loc = rf"Software\Classes\SystemFileAssociations\.{ext}"
        targets = menu_targets(ext)
        default_t = targets[0] if targets else None
        ok = False
        for loc in locations:
            # Remove any legacy flat top-level verbs and any duplicate ProgID cascading menus
            for verb, _label, _cmd in _verbs_for(ext):
                _delete_verb(loc, verb)
            if shell_ext_ok or loc != primary_loc:
                _delete_verb(loc, VERB_MENU)
        if shell_ext_ok:
            ok = True
        else:
            try:
                _write_cascading_menu(primary_loc, targets, default_t)
                ok = True
            except Exception as exc:
                _log(f"cascading menu .{ext} failed at {primary_loc}: {exc}")
        try:
            _register_openwith(ext)
        except Exception as exc:
            _log(f"open-with registration for .{ext} failed: {exc}")
        if ok:
            n += 1
    # Clean up universal '*' verbs so they don't duplicate or override per-extension target lists
    if winreg is not None:
        _delete_verb(r"Software\Classes\*", VERB_MENU)
        _delete_verb(r"Software\Classes\*", VERB_CHOOSE)
        _delete_verb(r"Software\Classes\*", VERB_AUTO)

    _log(f"context menu installed for {n}/{len(exts)} extensions "
         f"({'Win11 shell extension' if shell_ext_ok else 'registry cascade'}; single entry per menu)")
    return n


def _is_win11() -> bool:
    try:
        return IS_WINDOWS and sys.getwindowsversion().build >= 22000  # type: ignore[attr-defined]
    except Exception:
        return False


def remove_context_menu(extensions: Optional[list[str]] = None) -> int:
    if not IS_WINDOWS or winreg is None:
        return 0
    exts = extensions or MENU_EXTENSIONS
    n = 0
    for ext in exts:
        for base in _verb_locations(ext):
            if _delete_verb(base, VERB_MENU):
                n += 1
            for verb, _label, _cmd in _verbs_for(ext):
                if _delete_verb(base, verb):
                    n += 1
        _unregister_openwith(ext)
    _delete_verb(r"Software\Classes\*", VERB_MENU)
    _delete_verb(r"Software\Classes\*", VERB_CHOOSE)
    _delete_verb(r"Software\Classes\*", VERB_AUTO)
    _delete_key_tree(winreg.HKEY_CURRENT_USER,
                     rf"Software\Classes\Applications\{OPENWITH_APP}")
    if winreg is not None:
        _delete_key_tree(winreg.HKEY_CURRENT_USER, r"Software\Classes\Applications\ConFas.exe")
    _delete_key_tree(winreg.HKEY_CURRENT_USER, r"Software\ClickShift")
    _delete_key_tree(winreg.HKEY_CURRENT_USER, r"Software\ConFas")
    _set_win11_both_menus(False)
    _remove_win11_shell_extension()
    _log(f"context menu removed ({n} verb keys deleted)")
    return n


def refresh_explorer() -> None:
    """Best-effort notify Shell about changes (no kill needed, no admin)."""
    if not IS_WINDOWS:
        return
    try:
        windll: Any = getattr(ctypes, "windll", None)
        if windll is None:
            return
        HWND_BROADCAST = 0xFFFF
        WM_SETTINGCHANGE = 0x1A
        SHChangeNotify = windll.shell32.SHChangeNotify
        SHCNE_ASSOCCHANGED = 0x08000000
        SHCNF_IDLIST = 0x0000
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
        windll.user32.SendMessageW(HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment")
    except Exception:
        pass


# --------------------------------------------------------------------------- #
#  Watch folder daemon
# --------------------------------------------------------------------------- #

class WatchFolder(threading.Thread):
    """Poll-based watcher (dependency-free, works from frozen EXE reliably)."""

    POLL_SECONDS = 1.0
    STABLE_CHECKS = 2   # file size must be unchanged N polls before we convert

    def __init__(self, folder: Path, cfg: dict, on_done=None) -> None:
        super().__init__(daemon=True, name="clickshift-watch")
        self.folder = folder
        self.cfg = cfg
        self.on_done = on_done
        self.stop_event = threading.Event()
        self._seen: dict[str, tuple[int, int]] = {}   # path -> (size, stable_count)
        self._produced: set[str] = set()              # outputs created by watcher itself

    def run(self) -> None:
        from .engine import newest_target_for, get_engine
        self.folder.mkdir(parents=True, exist_ok=True)
        _log(f"watching {self.folder} (auto target: {self.cfg.get('auto_target')})")
        while not self.stop_event.is_set():
            try:
                now = {}
                for p in self.folder.iterdir():
                    if (p.is_file() and not p.name.startswith("~$")
                            and not p.name.endswith((".clickshift-debug.txt", ".confas-debug.txt"))):
                        try:
                            now[str(p)] = p.stat().st_size
                        except OSError:
                            continue
                for path_str, size in now.items():
                    if path_str in self._produced:
                        continue
                    prev = self._seen.get(path_str)
                    if prev is None:
                        self._seen[path_str] = (size, 0)
                        continue
                    old_size, stable = prev
                    if size != old_size:
                        self._seen[path_str] = (size, 0)      # still being written
                        continue
                    if stable < self.STABLE_CHECKS:
                        self._seen[path_str] = (size, stable + 1)
                        continue
                    # file is complete -> convert once (skip our own outputs)
                    del self._seen[path_str]
                    src = Path(path_str)
                    out_ext = (self.cfg.get("auto_target") or "").lower()
                    if out_ext and src.stem.lower().endswith("_" + out_ext):
                        continue
                    self._convert(src, newest_target_for, get_engine)
                # forget vanished files
                for gone in set(self._seen) - set(now):
                    del self._seen[gone]
                for gone in set(self._produced) - set(now):
                    self._produced.discard(gone)
            except Exception as exc:
                _log(f"watch loop error: {exc}")
            self.stop_event.wait(self.POLL_SECONDS)

    def _convert(self, src: Path, newest_fn, engine_getter) -> None:
        target = (self.cfg.get("auto_target") or "pdf").lower().lstrip(".")
        src_ext = src.suffix.lstrip(".").lower()
        if src_ext == target:
            # e.g. user dropped a .pdf when auto_target is "pdf" -> convert to its
            # newest target ("docx") instead of silently skipping the file!
            sugg = newest_fn(src.suffix)
            if not sugg:
                return
            target = sugg[0]
        eng = engine_getter(enable_ml=bool(self.cfg.get("ml_enabled", True)))
        res = eng.convert(src, target)
        if res.ok and res.target:
            self._produced.add(str(Path(res.target)))
        _log(f"watch-converted {src.name} -> {res.target} via {res.backend} "
             f"({res.elapsed_ms:.0f} ms, ok={res.ok}{', err=' + res.error if res.error else ''})")
        if self.on_done:
            try:
                self.on_done(res)
            except Exception:
                pass


# --------------------------------------------------------------------------- #
#  Single-instance guard
# --------------------------------------------------------------------------- #

def acquire_single_instance() -> bool:
    """Return True if this process is the only ClickShift daemon."""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        if PID_FILE.exists():
            pid = int(PID_FILE.read_text().strip() or "0")
            if pid and _pid_alive(pid):
                return False
        PID_FILE.write_text(str(os.getpid()))
        return True
    except Exception:
        return True


def _pid_alive(pid: int) -> bool:
    if IS_WINDOWS:
        try:
            windll: Any = getattr(ctypes, "windll", None)
            if windll is None:
                return False
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            k32 = windll.kernel32
            h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if not h:
                return False
            code = ctypes.c_ulong()
            ok = k32.GetExitCodeProcess(h, ctypes.byref(code))
            k32.CloseHandle(h)
            return bool(ok) and code.value == STILL_ACTIVE
        except Exception:
            return False
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


# --------------------------------------------------------------------------- #
#  Local Warm-Daemon IPC (fast path for Explorer right-click & CLI)
# --------------------------------------------------------------------------- #

def ipc_file() -> Path:
    return data_dir() / "daemon_ipc.json"


class DaemonIPCServer(threading.Thread):
    """Loopback IPC server so CLI/right-click conversions run inside the warm daemon."""

    def __init__(self, cfg: dict) -> None:
        super().__init__(daemon=True, name="clickshift-ipc")
        self.cfg = cfg
        self.token = uuid.uuid4().hex
        self.stop_event = threading.Event()
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(16)
        self._sock.settimeout(1.0)
        self.port = int(self._sock.getsockname()[1])

    def start_and_publish(self) -> None:
        f = ipc_file()
        f.parent.mkdir(parents=True, exist_ok=True)
        try:
            fd = os.open(str(f), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                json.dump({"port": self.port, "token": self.token, "pid": os.getpid()}, fh)
        except Exception:
            f.write_text(json.dumps({"port": self.port, "token": self.token, "pid": os.getpid()}),
                         encoding="utf-8")
        self.start()

    def stop(self) -> None:
        self.stop_event.set()
        try:
            self._sock.close()
        except Exception:
            pass
        try:
            ipc_file().unlink(missing_ok=True)
        except Exception:
            pass

    def run(self) -> None:
        from .engine import get_engine
        # Pre-warm the singleton engine on daemon start
        try:
            get_engine(enable_ml=bool(self.cfg.get("ml_enabled", True)))
        except Exception:
            pass
        while not self.stop_event.is_set():
            try:
                conn, _ = self._sock.accept()
            except socket.timeout:
                continue
            except OSError:
                break
            threading.Thread(target=self._handle, args=(conn,), daemon=True).start()

    def _handle(self, conn) -> None:
        from .engine import get_engine
        try:
            conn.settimeout(300.0)
            with conn:
                rfile = conn.makefile("r", encoding="utf-8")
                wfile = conn.makefile("w", encoding="utf-8")
                line = rfile.readline()
                if not line:
                    return
                req = json.loads(line)
                if req.get("token") != self.token:
                    return
                enable_ml = bool(req.get("enable_ml", True))
                eng = get_engine(enable_ml=enable_ml)
                write_lock = threading.Lock()

                def on_done(res):
                    with write_lock:
                        wfile.write(json.dumps(res.as_dict()) + "\n")
                        wfile.flush()

                if req.get("action") == "combine_images":
                    files = [str(p) for p in (req.get("files") or [])]
                    target = str(req.get("target") or "pdf")
                    on_done(eng.combine_images(files, target))
                else:
                    jobs = [(str(s), str(t)) for s, t in (req.get("jobs") or [])]
                    eng.convert_many(jobs, progress=on_done)
        except Exception as exc:
            _log(f"ipc handler error: {exc}")


def _ipc_request(payload_fields: dict, progress=None) -> Optional[list]:
    """Send an authenticated JSON request to the running warm daemon."""
    if os.environ.get("CLICKSHIFT_NO_IPC") == "1":
        return None
    f = ipc_file()
    if not f.exists():
        return None
    try:
        from .engine import ConversionResult
        meta = json.loads(f.read_text(encoding="utf-8"))
        pid = int(meta.get("pid", 0))
        if pid == os.getpid() or (pid and not _pid_alive(pid)):
            return None
        port = int(meta["port"])
        token = str(meta["token"])
        results = []
        with socket.create_connection(("127.0.0.1", port), timeout=1.5) as sock:
            sock.settimeout(300.0)
            rfile = sock.makefile("r", encoding="utf-8")
            wfile = sock.makefile("w", encoding="utf-8")
            body = {"token": token, **payload_fields}
            wfile.write(json.dumps(body) + "\n")
            wfile.flush()
            for line in rfile:
                if not line.strip():
                    continue
                d = json.loads(line)
                res = ConversionResult(
                    source=d["source"],
                    target=d["target"],
                    backend=d["backend"],
                    elapsed_ms=float(d["elapsed_ms"]),
                    ok=bool(d["ok"]),
                    error=d.get("error", ""),
                    bytes_in=int(d.get("bytes_in", 0)),
                    bytes_out=int(d.get("bytes_out", 0)),
                )
                results.append(res)
                if progress:
                    try:
                        progress(res)
                    except Exception:
                        pass
        return results
    except Exception:
        return None


def try_daemon_convert(jobs: list[tuple[str, str]], enable_ml: bool = True,
                       progress=None) -> Optional[list]:
    """Send conversion jobs to the warm daemon if running; return None on any IPC miss."""
    results = _ipc_request({"jobs": jobs, "enable_ml": enable_ml}, progress=progress)
    return results if (results is not None and len(results) == len(jobs)) else None


def try_daemon_combine(files: list[str], target: str, enable_ml: bool = True,
                       progress=None) -> Optional[list]:
    """Combine ordered images into one document inside the warm daemon if running."""
    results = _ipc_request(
        {"action": "combine_images", "files": files, "target": target, "enable_ml": enable_ml},
        progress=progress,
    )
    return results if (results is not None and len(results) == 1) else None


# --------------------------------------------------------------------------- #
#  System tray daemon
# --------------------------------------------------------------------------- #

def _make_tray_icon_image():
    """Load assets/clickshift.ico (or clickshift.png) for the system tray, with fallback."""
    Image: Any = importlib.import_module("PIL.Image")
    ImageDraw: Any = importlib.import_module("PIL.ImageDraw")
    assets = _assets_dir()
    for name in ("clickshift.ico", "clickshift.png"):
        p = assets / name
        if p.exists():
            try:
                img = Image.open(p)
                img.load()
                return img.convert("RGBA").resize((64, 64), Image.LANCZOS)
            except Exception:
                pass
    # Fallback matching the navy / split-doc / gold-bolt icon design
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([3, 3, 61, 61], radius=14, fill=(13, 30, 62, 255))
    d.rounded_rectangle([15, 11, 49, 53], radius=4, fill=(56, 189, 248, 255))
    d.rectangle([15, 11, 32, 53], fill=(248, 250, 252, 255))
    d.polygon([(35, 17), (24, 35), (32, 35), (29, 49), (40, 30), (33, 30)], fill=(250, 204, 21, 255))
    return img


DEFAULT_UPDATE_INTERVAL_S = 43200  # 12 hours


def update_check_due_in(cfg: dict, now: Optional[float] = None) -> tuple[bool, float]:
    """Return (is_due_now, seconds_until_next_check) for the 12-hour background updater."""
    interval = max(300, int(cfg.get("update_check_interval_s", DEFAULT_UPDATE_INTERVAL_S)))
    last_ts = float(cfg.get("last_update_check_ts", 0.0) or 0.0)
    cur = time.time() if now is None else float(now)
    elapsed = cur - last_ts
    if last_ts <= 0.0 or elapsed >= interval:
        return True, 0.0
    return False, max(1.0, float(interval) - elapsed)


def perform_background_update_check(
    current_version: str,
    on_before_restart: Optional[Any] = None,
    updater_mod: Optional[Any] = None,
    now: Optional[float] = None,
) -> str:
    """Record check timestamp and silently download+install if a newer release exists.

    Returns one of: 'updated', 'not_frozen', 'up_to_date', 'disabled', 'error'.
    """
    cfg = load_config()
    if not cfg.get("auto_update_check", True):
        return "disabled"
    if updater_mod is None:
        from . import updater as updater_mod
    cfg["last_update_check_ts"] = float(time.time() if now is None else now)
    save_config(cfg)
    try:
        info = updater_mod.check_for_update(current_version=current_version)
        if info is None:
            return "up_to_date"
        return str(
            updater_mod.run_background_update(
                info,
                current_version=current_version,
                on_before_restart=on_before_restart,
            )
        )
    except Exception as exc:
        _log(f"background update check failed: {exc}")
        return "error"


def run_daemon(headless: bool = False) -> int:
    """Main entry for `ClickShift.exe run` - stays alive forever."""
    cfg = load_config()
    if not acquire_single_instance():
        _log("another instance already running; exiting")
        return 0
    DROP_DIR.mkdir(parents=True, exist_ok=True)
    Path(cfg["watch_folder"]).mkdir(parents=True, exist_ok=True)
    ipc = None
    try:
        ipc = DaemonIPCServer(cfg)
        ipc.start_and_publish()
        globals()["_ACTIVE_IPC"] = ipc
    except Exception as exc:
        _log(f"ipc server unavailable: {exc}")
    watch = None
    if cfg.get("watch_enabled", True):
        watch = WatchFolder(Path(cfg["watch_folder"]), cfg)
        watch.start()
        globals()["_ACTIVE_WATCH"] = watch
    if headless or not IS_WINDOWS:
        _log(f"daemon started headless (pid {os.getpid()})")
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            pass
        finally:
            if ipc:
                ipc.stop()
        return 0

    # ---- pystray tray icon ------------------------------------------------ #
    update_stop = threading.Event()
    update_lock = threading.Lock()
    try:
        pystray: Any = importlib.import_module("pystray")
        from . import __version__ as _ver
        from . import updater as _updater

        paused = {"v": False}
        icon_holder: dict[str, Any] = {"icon": None}

        def _stop_for_update() -> None:
            update_stop.set()
            if watch:
                watch.stop_event.set()
            if ipc:
                ipc.stop()
            ic = icon_holder.get("icon")
            if ic is not None:
                try:
                    ic.stop()
                except Exception:
                    pass

        def _check_for_update(manual: bool = False) -> None:
            if not update_lock.acquire(blocking=False):
                return
            try:
                if not manual:
                    perform_background_update_check(
                        current_version=_ver,
                        on_before_restart=_stop_for_update,
                        updater_mod=_updater,
                    )
                    return
                cfg_now = load_config()
                cfg_now["last_update_check_ts"] = time.time()
                save_config(cfg_now)
                info = _updater.check_for_update(current_version=_ver)
                if info is None:
                    if icon_holder.get("icon") is not None:
                        try:
                            icon_holder["icon"].notify(f"ClickShift v{_ver} is up to date.")
                        except Exception:
                            pass
                    return
                _updater.run_interactive_update(
                    info,
                    current_version=_ver,
                    on_before_restart=_stop_for_update,
                )
            except Exception as exc:
                _log(f"update check failed: {exc}")
            finally:
                update_lock.release()

        def _auto_update_loop() -> None:
            if update_stop.wait(5.0):
                return
            while not update_stop.is_set():
                cur_cfg = load_config()
                if cur_cfg.get("auto_update_check", True):
                    due, wait_s = update_check_due_in(cur_cfg)
                    if due:
                        _check_for_update(manual=False)
                        wait_s = max(300, int(load_config().get("update_check_interval_s", DEFAULT_UPDATE_INTERVAL_S)))
                else:
                    wait_s = max(300, int(cur_cfg.get("update_check_interval_s", DEFAULT_UPDATE_INTERVAL_S)))
                if update_stop.wait(wait_s):
                    return

        def toggle_pause(icon, item):
            paused["v"] = not paused["v"]
            if watch:
                if paused["v"]:
                    watch.stop_event.set()
                else:
                    # restart a fresh watcher thread (stop_event is one-shot)
                    w2 = WatchFolder(Path(cfg["watch_folder"]), cfg)
                    w2.start()
                    globals()["_ACTIVE_WATCH"] = w2
            icon.notify(f"ClickShift {'paused' if paused['v'] else 'active'}")

        def open_drop(icon, item):
            subprocess.Popen(["explorer", str(Path(cfg["watch_folder"]))])

        def check_updates_menu(icon, item):
            threading.Thread(
                target=lambda: _check_for_update(manual=True),
                daemon=True,
            ).start()

        def quit_app(icon, item):
            update_stop.set()
            if watch:
                watch.stop_event.set()
            if ipc:
                ipc.stop()
            icon.stop()

        def _on_tray_ready(ic) -> None:
            ic.visible = True
            notice = _updater.consume_update_notice()
            if notice:
                title, msg = _updater.format_update_notification(notice)
                _log(msg)
                time.sleep(1.0)
                try:
                    ic.notify(msg, title)
                except Exception:
                    pass
            if os.environ.get("CLICKSHIFT_NO_POPUP") != "1":
                threading.Thread(target=_auto_update_loop, daemon=True).start()

        menu = pystray.Menu(
            pystray.MenuItem(lambda i: f"ClickShift v{_ver} - {'paused' if paused['v'] else 'watching ' + cfg['watch_folder']}",
                             None, enabled=False),
            pystray.MenuItem("Pause/Resume watching", toggle_pause),
            pystray.MenuItem("Open drop folder", open_drop),
            pystray.MenuItem("Check for updates...", check_updates_menu),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit ClickShift", quit_app),
        )
        icon = pystray.Icon(APP_NAME, _make_tray_icon_image(), f"ClickShift v{_ver} - document converter", menu)
        icon_holder["icon"] = icon
        _log("tray daemon started")
        icon.run(setup=_on_tray_ready)
        return 0
    except Exception as exc:
        update_stop.set()
        _log(f"tray unavailable ({exc}); falling back to headless loop")
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            return 0
        finally:
            if ipc:
                ipc.stop()
