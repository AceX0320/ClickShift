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

import base64
import ctypes
import hashlib
import hmac
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
_REAL_WINREG: Any = winreg
_REAL_SUBPROCESS_RUN: Any = subprocess.run
APP_NAME = "ClickShift"


def _is_pytest_running() -> bool:
    return (
        ("PYTEST_CURRENT_TEST" in os.environ or "pytest" in sys.modules)
        and os.environ.get("CLICKSHIFT_ALLOW_PYTEST_REGISTRY") != "1"
    )


def _can_write_system_registry() -> bool:
    """Allow registry writes unless pytest is running against the real Windows winreg module."""
    if not IS_WINDOWS or winreg is None:
        return False
    if _is_pytest_running() and _REAL_WINREG is not None and winreg is _REAL_WINREG:
        return False
    return True


def _is_python_interpreter(path_str: str) -> bool:
    if not path_str:
        return True
    name = path_str.replace("\\", "/").rsplit("/", 1)[-1].lower()
    return (
        name in {"python.exe", "pythonw.exe", "python", "python3", "python3.exe"}
        or name.startswith("python3.")
    )


def _find_installed_clickshift_exe() -> Optional[Path]:
    """Find a real frozen ClickShift*.exe binary (never python.exe)."""
    if getattr(sys, "frozen", False) and not _is_python_interpreter(sys.executable):
        p = Path(sys.executable)
        if p.is_file():
            return p
    if IS_WINDOWS and winreg is not None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\ClickShift") as k:
                val, _ = winreg.QueryValueEx(k, "ExePath")
            if isinstance(val, str) and val.strip() and not _is_python_interpreter(val):
                p = Path(val.strip())
                if p.is_file() and p.suffix.lower() == ".exe":
                    return p
        except Exception:
            pass
    local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
    cand_local = local_app / "ClickShift" / "ClickShift.exe"
    if cand_local.is_file():
        return cand_local
    downloads = Path.home() / "Downloads"
    if downloads.is_dir():
        try:
            exes = sorted(
                (
                    p for p in downloads.glob("ClickShift*.exe")
                    if p.is_file() and not p.name.lower().endswith(".old")
                ),
                key=lambda x: x.stat().st_mtime,
                reverse=True,
            )
            if exes:
                return exes[0]
        except Exception:
            pass
    return None

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
        from .formats import all_targets_for
        targets = all_targets_for(ext_clean)
        if targets:
            return targets
    except Exception:
        pass
    return ["pdf", "docx", "txt"]


# --------------------------------------------------------------------------- #
#  Config
# --------------------------------------------------------------------------- #

DEFAULT_CONFIG: dict[str, Any] = {
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
    "cloud_enabled": True,          # shared community ML pool (disable with: cloud_sync=off)
    "cloud_mode": "url",            # "url" (shared Cloudflare pool) or "gist" (GitHub Gist)
    "cloud_gist_id": "REPLACE_WITH_YOUR_GIST_ID",
    "cloud_url": "https://clickshift-cloud.clickshift.workers.dev",
    "cloud_interval_s": 900,        # background push/pull cadence
    "cloud_debounce_s": 60,         # min seconds between uploads
    "cloud_pull_on_start": True,
    "cloud_require_registered": True,  # only registered download client_ids can push to cloud
    # --- GitHub auto-updater --------------------------------------------- #
    "auto_update_check": True,
    "update_check_interval_s": 43200,   # check every 12 hours
    "last_update_check_ts": 0.0,
    "skipped_update_version": "",
    # --- Security hardening ---------------------------------------------- #
    "office_disable_macros": True,
}


def _restrict_to_current_user(path: Path) -> bool:
    """Restrict file/directory permissions to the current user (NTFS DACL on Windows, 0o600/0o700 on POSIX)."""
    try:
        os.chmod(path, 0o700 if path.is_dir() else 0o600)
    except OSError:
        pass
    if not IS_WINDOWS:
        return True
    try:
        from ctypes import wintypes
        advapi32 = ctypes.windll.advapi32
        kernel32 = ctypes.windll.kernel32
        TOKEN_QUERY = 0x0008
        TokenUser = 1
        SE_FILE_OBJECT = 1
        DACL_SECURITY_INFORMATION = 0x00000004
        PROTECTED_DACL_SECURITY_INFORMATION = 0x80000000
        SDDL_REVISION_1 = 1

        class SID_AND_ATTRIBUTES(ctypes.Structure):
            _fields_ = [("Sid", ctypes.c_void_p), ("Attributes", wintypes.DWORD)]

        h_token = wintypes.HANDLE()
        if not advapi32.OpenProcessToken(kernel32.GetCurrentProcess(), TOKEN_QUERY, ctypes.byref(h_token)):
            return False
        try:
            needed = wintypes.DWORD(0)
            advapi32.GetTokenInformation(h_token, TokenUser, None, 0, ctypes.byref(needed))
            if not needed.value:
                return False
            buf = ctypes.create_string_buffer(needed.value)
            if not advapi32.GetTokenInformation(h_token, TokenUser, buf, needed, ctypes.byref(needed)):
                return False
            tu = SID_AND_ATTRIBUTES.from_buffer(buf)
            sid_ptr = ctypes.c_wchar_p()
            if not advapi32.ConvertSidToStringSidW(tu.Sid, ctypes.byref(sid_ptr)):
                return False
            try:
                sid_str = sid_ptr.value or ""
            finally:
                kernel32.LocalFree(sid_ptr)
        finally:
            kernel32.CloseHandle(h_token)

        if not sid_str:
            return False
        # Grant Full Access exclusively to SYSTEM and the current user SID; block inherited ACEs
        sddl = f"D:P(A;;FA;;;SY)(A;;FA;;;{sid_str})"
        sd = ctypes.c_void_p()
        if not advapi32.ConvertStringSecurityDescriptorToSecurityDescriptorW(
            sddl, SDDL_REVISION_1, ctypes.byref(sd), None
        ):
            return False
        try:
            dacl = ctypes.c_void_p()
            dacl_present = wintypes.BOOL()
            dacl_defaulted = wintypes.BOOL()
            if not advapi32.GetSecurityDescriptorDacl(
                sd, ctypes.byref(dacl_present), ctypes.byref(dacl), ctypes.byref(dacl_defaulted)
            ) or not dacl_present.value:
                return False
            rc = advapi32.SetNamedSecurityInfoW(
                str(path),
                SE_FILE_OBJECT,
                DACL_SECURITY_INFORMATION | PROTECTED_DACL_SECURITY_INFORMATION,
                None,
                None,
                dacl,
                None,
            )
            return rc == 0
        finally:
            kernel32.LocalFree(sd)
    except Exception:
        return False


def validate_watch_folder(path: Path) -> Path:
    """Resolve and verify a watch folder isn't a protected system/profile root."""
    resolved = path.resolve()
    if len(resolved.parts) <= 1:
        raise ValueError(f"Refusing root path as watch folder: {resolved}")
    home = Path.home().resolve()
    protected = {
        home,
        home / "Desktop",
        home / "Documents",
        home / "Downloads",
        home / "Pictures",
        home / "Videos",
        home / "Music",
    }
    if IS_WINDOWS:
        win_root = Path(os.environ.get("SystemRoot", r"C:\Windows")).resolve()
        pf = Path(os.environ.get("ProgramFiles", r"C:\Program Files")).resolve()
        pf86 = Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")).resolve()
        for sys_dir in (win_root, pf, pf86):
            if resolved == sys_dir or sys_dir in resolved.parents:
                raise ValueError(f"Refusing system directory as watch folder: {resolved}")
    if resolved in protected:
        raise ValueError(
            f"Refusing user library folder as watch folder (it would auto-convert all existing files): {resolved}"
        )
    return resolved


_INITIAL_DATA_DIR = DATA_DIR
_CONFIG_CACHE_LOCK = threading.Lock()
_CONFIG_CACHE: Optional[tuple[str, int, int, dict]] = None


def _active_data_dir() -> Path:
    if DATA_DIR != _INITIAL_DATA_DIR:
        return DATA_DIR
    return data_dir()


def config_file() -> Path:
    """Resolve the config file path at call time (honours CLICKSHIFT_HOME changes/tests)."""
    return _active_data_dir() / "config.json"


def load_config() -> dict:
    global _CONFIG_CACHE
    cfile = config_file()
    cfg = dict(DEFAULT_CONFIG)
    try:
        st = cfile.stat()
    except FileNotFoundError:
        return cfg
    except OSError:
        return cfg
    key_path = str(cfile)
    with _CONFIG_CACHE_LOCK:
        if (_CONFIG_CACHE is not None
                and _CONFIG_CACHE[0] == key_path
                and _CONFIG_CACHE[1] == st.st_mtime_ns
                and _CONFIG_CACHE[2] == st.st_size):
            cfg.update(_CONFIG_CACHE[3])
            return cfg
    try:
        loaded = json.loads(cfile.read_text(encoding="utf-8"))
        if isinstance(loaded, dict):
            with _CONFIG_CACHE_LOCK:
                _CONFIG_CACHE = (key_path, st.st_mtime_ns, st.st_size, dict(loaded))
            cfg.update(loaded)
    except Exception:
        pass
    return cfg


def save_config(cfg: dict) -> None:
    global _CONFIG_CACHE
    cfile = config_file()
    cfile.parent.mkdir(parents=True, exist_ok=True)
    tmp = cfile.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    tmp.replace(cfile)
    try:
        st = cfile.stat()
        with _CONFIG_CACHE_LOCK:
            _CONFIG_CACHE = (str(cfile), st.st_mtime_ns, st.st_size, dict(cfg))
    except OSError:
        with _CONFIG_CACHE_LOCK:
            _CONFIG_CACHE = None


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
STARTUP_APPROVED_KEY = r"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
_STARTUP_ENABLED_BIN = b"\x02" + b"\x00" * 11


def _extract_exe_from_cmd(raw: str) -> str:
    s = (raw or "").strip()
    if not s:
        return ""
    if s[0] == '"':
        end_q = s.find('"', 1)
        if end_q != -1:
            return s[1:end_q]
    low = s.lower()
    idx = 0
    while True:
        pos = low.find(".exe", idx)
        if pos == -1:
            break
        end_exe = pos + 4
        if end_exe == len(s) or s[end_exe] in (" ", "\t"):
            return s[:end_exe]
        idx = end_exe
    return s.split()[0]


def register_startup() -> bool:
    """Add ClickShift to the current user's Run key and ensure StartupApproved is enabled."""
    if not _can_write_system_registry():
        return False
    try:
        if not getattr(sys, "frozen", False) and winreg is _REAL_WINREG:
            frozen_exe = _find_installed_clickshift_exe()
            if frozen_exe is not None:
                cmd = subprocess.list2cmdline([str(frozen_exe), "run"])
            else:
                cmd = subprocess.list2cmdline(launch_command(["run"]))
        else:
            cmd = subprocess.list2cmdline(launch_command(["run"]))
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                                winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, cmd)
            try:
                winreg.DeleteValue(k, "ConFas")
            except OSError:
                pass
        # Windows 10/11 blocks Run entries whose StartupApproved\Run byte0 is odd (e.g. 0x03).
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, STARTUP_APPROVED_KEY, 0,
                                    winreg.KEY_SET_VALUE) as ak:
                winreg.SetValueEx(ak, APP_NAME, 0, winreg.REG_BINARY, _STARTUP_ENABLED_BIN)
                try:
                    winreg.DeleteValue(ak, "ConFas")
                except OSError:
                    pass
        except OSError:
            pass
        _save_exe_path_registry()
        _log(f"startup registered: {cmd}")
        return True
    except Exception as exc:
        _log(f"register_startup failed: {exc}")
        return False


def unregister_startup() -> bool:
    if not _can_write_system_registry():
        return False
    ok = True
    for key_path in (RUN_KEY, STARTUP_APPROVED_KEY):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                winreg.KEY_SET_VALUE) as k:
                for name in (APP_NAME, "ConFas"):
                    try:
                        winreg.DeleteValue(k, name)
                    except FileNotFoundError:
                        pass
        except FileNotFoundError:
            pass
        except Exception:
            if key_path == RUN_KEY:
                ok = False
    return ok


def is_startup_registered() -> bool:
    if not IS_WINDOWS or winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            val, _ = winreg.QueryValueEx(k, APP_NAME)
        if not isinstance(val, str) or not val.strip():
            return False
        target = _extract_exe_from_cmd(val)
        if getattr(sys, "frozen", False):
            if not target:
                return False
            try:
                if os.path.normcase(os.path.abspath(target)) != os.path.normcase(os.path.abspath(sys.executable)):
                    return False
                if not Path(target).exists():
                    return False
            except Exception:
                return False
        elif winreg is _REAL_WINREG and _is_python_interpreter(target) and _find_installed_clickshift_exe() is not None:
            return False
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_APPROVED_KEY) as ak:
                raw, _ = winreg.QueryValueEx(ak, APP_NAME)
            if isinstance(raw, (bytes, bytearray)) and len(raw) >= 1 and (raw[0] & 1) != 0:
                return False
        except FileNotFoundError:
            pass
        except OSError:
            pass
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


def _verb_names_for(ext: str) -> list[str]:
    """Verb key names for one extension (fast: skips command string construction)."""
    targets = menu_targets(ext)
    out = [VERB_AUTO]
    out.extend(f"ClickShift.To{t.capitalize()}" for t in targets)
    out.append(VERB_CHOOSE)
    return out


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


_OPENWITH_APP_LOCAL = threading.local()


def _register_openwith_app() -> None:
    """Register HKCU\\Software\\Classes\\Applications\\ClickShift.exe once per batch."""
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
        with winreg.CreateKey(k, r"shell\open\command") as kc:
            winreg.SetValue(kc, "", winreg.REG_SZ, _verb_cmd(None, choose=True))


def _register_openwith(ext: str) -> None:
    """Make ClickShift appear in Explorer's 'Open with' (Windows 11 first menu) with all choices."""
    assert winreg is not None
    app = OPENWITH_APP
    if not getattr(_OPENWITH_APP_LOCAL, "registered", False):
        _register_openwith_app()
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
    if not _can_write_system_registry():
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


def _sync_localappdata_copy(src_exe: Path) -> None:
    """Maintain a resilient backup copy at %LOCALAPPDATA%\\ClickShift\\ClickShift.exe."""
    if not IS_WINDOWS or _is_pytest_running():
        return
    try:
        import shutil
        local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        dst_exe = local_app / "ClickShift" / "ClickShift.exe"
        dst_exe.parent.mkdir(parents=True, exist_ok=True)
        if src_exe.resolve() == dst_exe.resolve() or _files_identical(src_exe, dst_exe):
            return
        tmp = dst_exe.with_suffix(".exe.stage")
        shutil.copy2(src_exe, tmp)
        try:
            os.replace(tmp, dst_exe)
        except OSError:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
    except Exception:
        pass


def _save_exe_path_registry() -> None:
    """Store the active frozen ClickShift executable path in HKCU\\Software\\ClickShift for ClickShiftShell.dll."""
    if not _can_write_system_registry():
        return
    try:
        if getattr(sys, "frozen", False) and not _is_python_interpreter(sys.executable):
            target_path = exe_path()
            _sync_localappdata_copy(Path(target_path))
        else:
            found = _find_installed_clickshift_exe()
            if found is None:
                return
            target_path = str(found)
        create_fn = getattr(winreg, "CreateKey", None) or getattr(winreg, "CreateKeyEx", None)
        if create_fn is None:
            return
        with create_fn(winreg.HKEY_CURRENT_USER, r"Software\ClickShift") as k:
            winreg.SetValueEx(k, "ExePath", 0, winreg.REG_SZ, target_path)
        _delete_key_tree(winreg.HKEY_CURRENT_USER, r"Software\ConFas")
    except Exception as exc:
        _log(f"save ExePath registry failed: {exc}")


def _cer_thumbprint(cer_path: Path) -> str:
    """Compute uppercase SHA-1 thumbprint of a DER/PEM .cer file."""
    try:
        raw = cer_path.read_bytes()
        if b"-----BEGIN CERTIFICATE-----" in raw:
            import ssl
            txt = raw.decode("ascii", errors="ignore")
            body = "".join(
                ln.strip() for ln in txt.splitlines()
                if ln.strip() and not ln.strip().startswith("-----")
            )
            raw = base64.b64decode(body)
        return hashlib.sha1(raw).hexdigest().upper()
    except Exception:
        return ""


def _show_admin_required_popup() -> None:
    """Notify the user when UAC elevation was declined or unavailable."""
    msg = "ClickShift needs the admin rights to work!"
    if not IS_WINDOWS or os.environ.get("CLICKSHIFT_NO_POPUP") == "1":
        print(msg)
        return
    try:
        ctypes.windll.user32.MessageBoxW(None, msg, "ClickShift", 0x00000030 | 0x00040000)
    except Exception:
        print(msg)


def _ps_encoded(script_text: str) -> str:
    return base64.b64encode(script_text.encode("utf-16le")).decode("ascii")


def _files_identical(src: Path, dst: Path) -> bool:
    try:
        if not dst.exists():
            return False
        s_st, d_st = src.stat(), dst.stat()
        if s_st.st_size != d_st.st_size:
            return False
        if s_st.st_mtime_ns == d_st.st_mtime_ns:
            return True
        return src.read_bytes() == dst.read_bytes()
    except OSError:
        return False


def _shell_ext_pkg_registered() -> bool:
    if winreg is None:
        return False
    pkg_reg = r"Software\Classes\PackagedCom\Package"
    has_clickshift = False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, pkg_reg) as k:
            i = 0
            while True:
                try:
                    name = winreg.EnumKey(k, i)
                    i += 1
                    if name.startswith("ConFas.ShellExtension_"):
                        return False
                    if name.startswith("ClickShift.ShellExtension_"):
                        has_clickshift = True
                except OSError:
                    break
    except OSError:
        return False
    return has_clickshift


def _shell_assets_digest(shell_dir: Path) -> str:
    h = hashlib.sha256()
    for fname in ("ClickShiftShell.msix", "ClickShiftShell.dll", "ClickShiftStub.exe", "AppxManifest.xml"):
        p = shell_dir / fname
        if p.exists():
            h.update(fname.encode("ascii"))
            h.update(p.read_bytes())
    return h.hexdigest()


def _install_win11_shell_extension() -> bool:
    """Install the packaged IExplorerCommand shell extension (ClickShiftShell.dll) so
    the expandable 'Convert with ClickShift >' submenu appears in the Windows 11 main menu.
    """
    if not IS_WINDOWS:
        return False
    if _is_pytest_running() and subprocess.run is _REAL_SUBPROCESS_RUN:
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

        msix_path = shell_dir / "ClickShiftShell.msix"
        cer_path = shell_dir / "ClickShiftShell.cer"
        manifest_path = shell_dir / "AppxManifest.xml"
        thumb_file = shell_dir / "installed_cert.thumbprint"
        stamp_file = shell_dir / ".installed_sha256"

        # If ClickShift.ShellExtension is already registered in Windows 11 and the native
        # C++ binaries (ClickShiftShell.dll / ClickShiftStub.exe) haven't changed, keep the
        # active registration intact instead of tearing down the package for a new CI cert.
        already_reg = _shell_ext_pkg_registered()
        if already_reg and (shell_dir / "ClickShiftShell.dll").exists():
            dll_dst = shell_dir / "ClickShiftShell.dll"
            stub_dst = shell_dir / "ClickShiftStub.exe"
            stub_src = src_assets / "ClickShiftStub.exe"
            if _files_identical(dll_src, dll_dst) and (not stub_src.exists() or _files_identical(stub_src, stub_dst)):
                return True
            src_cer = src_assets / "ClickShiftShell.cer"
            prev_tp = thumb_file.read_text(encoding="utf-8", errors="ignore").strip().upper() if thumb_file.exists() else ""
            new_tp = _cer_thumbprint(src_cer).upper() if src_cer.exists() else ""
            if prev_tp and new_tp and prev_tp != new_tp:
                return True

        for fname in ("ClickShiftShell.dll", "ClickShiftStub.exe", "AppxManifest.xml",
                      "clickshift.png", "clickshift.ico", "ClickShiftShell.msix", "ClickShiftShell.cer"):
            s = src_assets / fname
            d = shell_dir / fname
            if s.exists() and not _files_identical(s, d):
                try:
                    shutil.copy2(s, d)
                except Exception:
                    pass

        cur_digest = _shell_assets_digest(shell_dir)

        try:
            if (
                stamp_file.exists()
                and stamp_file.read_text(encoding="utf-8", errors="ignore").strip() == cur_digest
                and _shell_ext_pkg_registered()
            ):
                return True
        except Exception:
            pass

        # 1. Try direct in-place Add-AppxPackage registration (without removing ClickShift.ShellExtension first,
        # so an untrusted cert or slow AppXSvc never destroys an already-installed working shell extension).
        msix_lit = str(msix_path).replace("'", "''")
        manifest_lit = str(manifest_path).replace("'", "''")
        add_pkg_expr = (
            f"Add-AppxPackage -Path '{msix_lit}' -ForceApplicationShutdown -ErrorAction Stop"
            if msix_path.exists()
            else f"Add-AppxPackage -Register '{manifest_lit}' -ForceApplicationShutdown -ErrorAction Stop"
        )
        direct_cmd = (
            "Get-AppxPackage -Name 'ConFas.ShellExtension' -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue; "
            f"try {{ {add_pkg_expr} }} catch {{ "
            "Get-AppxPackage -Name 'ClickShift.ShellExtension' -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue; "
            f"{add_pkg_expr} }}"
        )
        r = None
        try:
            r = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", _ps_encoded(direct_cmd)],
                capture_output=True, text=True, timeout=60,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired:
            _log("Win11 direct Add-AppxPackage timed out after 60s; proceeding to trust verification")
            if _shell_ext_pkg_registered():
                return True
        if r is not None and r.returncode == 0:
            try:
                stamp_file.write_text(cur_digest, encoding="utf-8")
            except Exception:
                pass
            _log("Win11 IExplorerCommand shell extension registered directly")
            return True

        # 2. One-time elevated trust + registration (triggers UAC prompt).
        # Only unlock AllowDevelopmentWithoutDevLicense when unpacked manifest fallback is used.
        elev_parts = [
            'New-Item -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AppModelUnlock" -Force | Out-Null',
            'Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AppModelUnlock" -Name "AllowAllTrustedApps" -Value 1 -Type DWord -Force',
        ]
        if not msix_path.exists():
            elev_parts.append(
                'Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AppModelUnlock" -Name "AllowDevelopmentWithoutDevLicense" -Value 1 -Type DWord -Force'
            )
        if cer_path.exists():
            cer_lit = str(cer_path).replace("'", "''")
            elev_parts.append(
                f"Import-Certificate -FilePath '{cer_lit}' -CertStoreLocation 'Cert:\\LocalMachine\\TrustedPeople' | Out-Null"
            )
            elev_parts.append(
                f"Import-Certificate -FilePath '{cer_lit}' -CertStoreLocation 'Cert:\\LocalMachine\\Root' | Out-Null"
            )
            tp = _cer_thumbprint(cer_path)
            if tp:
                try:
                    thumb_file.write_text(tp, encoding="utf-8")
                except Exception:
                    pass

        elev_b64 = _ps_encoded("; ".join(elev_parts))
        uac_wrapper = (
            "try {\n"
            f"  $p = Start-Process powershell -Verb RunAs -WindowStyle Hidden -PassThru -Wait -ArgumentList '-NoProfile','-NonInteractive','-WindowStyle','Hidden','-EncodedCommand','{elev_b64}'\n"
            "  if ($null -eq $p) { exit 5 }\n"
            "  exit $p.ExitCode\n"
            "} catch [System.ComponentModel.Win32Exception] {\n"
            "  if ($_.Exception.NativeErrorCode -eq 1223) { exit 1223 }\n"
            "  exit 5\n"
            "}\n"
        )
        r_uac = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", _ps_encoded(uac_wrapper)],
            capture_output=True, text=True, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if r_uac.returncode != 0:
            _log(f"Win11 admin elevation declined or failed (exit={r_uac.returncode}): {r_uac.stderr.strip()}")
            if r_uac.returncode in (1223, 5):
                _show_admin_required_popup()
            return False

        r2 = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", _ps_encoded(direct_cmd)],
            capture_output=True, text=True, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if r2.returncode == 0:
            try:
                stamp_file.write_text(cur_digest, encoding="utf-8")
            except Exception:
                pass
            _log("Win11 IExplorerCommand shell extension registered after one-time trust setup")
            return True
        r_err = r.stderr.strip() if (r is not None and getattr(r, "stderr", None)) else ""
        _log(f"Win11 shell extension registration stderr: {(r2.stderr or '').strip() or r_err}")
    except Exception as exc:
        _log(f"Win11 shell extension install error: {exc}")
    return False


def _remove_win11_shell_extension() -> None:
    if not IS_WINDOWS:
        return
    if _is_pytest_running() and subprocess.run is _REAL_SUBPROCESS_RUN:
        return
    try:
        local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        shell_dir = local_app / "ClickShift" / "win11_shell"
        thumb_file = shell_dir / "installed_cert.thumbprint"
        tp = ""
        if thumb_file.exists():
            tp = thumb_file.read_text(encoding="utf-8", errors="ignore").strip()
            if not all(c in "0123456789ABCDEFabcdef" for c in tp):
                tp = ""
        if shell_dir.exists() or tp or _shell_ext_pkg_registered():
            cmd = (
                "Get-AppxPackage -Name 'ClickShift.ShellExtension' -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue; "
                "Get-AppxPackage -Name 'ConFas.ShellExtension' -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue"
            )
            if tp:
                cmd += (
                    f"; Remove-Item 'Cert:\\LocalMachine\\TrustedPeople\\{tp}' -ErrorAction SilentlyContinue"
                    f"; Remove-Item 'Cert:\\LocalMachine\\Root\\{tp}' -ErrorAction SilentlyContinue"
                )
            subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", _ps_encoded(cmd)],
                capture_output=True, text=True, timeout=20,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        import shutil
        if shell_dir.exists():
            shutil.rmtree(shell_dir, ignore_errors=True)
    except Exception:
        pass


_SKIP_CLASS_KEYS = frozenset({
    "systemfileassociations", "clsid", "interface", "typelib", "appid",
    "directshow", "media type", "mediafoundation", "packagedcom",
    "installer", "local settings", "wow6432node", "activatableclasses",
    "extensions", "component categories", "mime",
})


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
            sl = sub.lower()
            if sl in _SKIP_CLASS_KEYS or sub.startswith("{"):
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
    if _is_pytest_running() and _REAL_WINREG is not None and winreg is _REAL_WINREG:
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
    from .updater import is_packaged_store_app
    if is_packaged_store_app():
        shell_ext_ok = _is_win11()
    else:
        shell_ext_ok = _is_win11() and _install_win11_shell_extension()
    exts = extensions or MENU_EXTENSIONS
    total_exts = max(1, len(exts))
    n = 0
    if winreg is not None:
        try:
            _register_openwith_app()
        except Exception:
            pass
    _OPENWITH_APP_LOCAL.registered = True
    try:
        for idx, ext in enumerate(exts):
            if idx % 6 == 0 or idx == total_exts - 1:
                pct = 55.0 + (idx / total_exts) * 30.0
                _report(f"Configuring right-click actions (.{ext})...", pct)
            locations = _verb_locations(ext)
            primary_loc = rf"Software\Classes\SystemFileAssociations\.{ext}"
            targets = menu_targets(ext)
            default_t = targets[0] if targets else None
            verb_names = [f"ClickShift.To{t.capitalize()}" for t in targets]
            all_verbs = [VERB_AUTO, *verb_names, VERB_CHOOSE]
            ok = False
            for loc in locations:
                for verb in all_verbs:
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
    finally:
        _OPENWITH_APP_LOCAL.registered = False
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
    if not _can_write_system_registry():
        return 0
    exts = extensions or MENU_EXTENSIONS
    n = 0
    for ext in exts:
        verb_names = _verb_names_for(ext)
        for base in _verb_locations(ext):
            if _delete_verb(base, VERB_MENU):
                n += 1
            for verb in verb_names:
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
        SMTO_ABORTIFHUNG = 0x0002
        SHChangeNotify = windll.shell32.SHChangeNotify
        SHCNE_ASSOCCHANGED = 0x08000000
        SHCNF_IDLIST = 0x0000
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
        windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment",
            SMTO_ABORTIFHUNG, 250, ctypes.byref(ctypes.c_size_t()),
        )
    except Exception:
        pass


# --------------------------------------------------------------------------- #
#  Watch folder daemon
# --------------------------------------------------------------------------- #

class WatchFolder(threading.Thread):
    """Poll-based watcher (dependency-free, works from frozen EXE reliably)."""

    POLL_SECONDS = 0.5
    STABLE_CHECKS = 2   # file size must be unchanged N polls before we convert

    def __init__(self, folder: Path, cfg: dict, on_done=None) -> None:
        super().__init__(daemon=True, name="clickshift-watch")
        self.folder = folder
        self.cfg = cfg
        self.on_done = on_done
        self.stop_event = threading.Event()
        self._seen: dict[str, tuple[int, int]] = {}          # path -> (size, stable_count)
        self._done: dict[str, tuple[int, int]] = {}          # path -> (mtime_ns, size) already processed
        self._produced: set[str] = set()                     # outputs created by watcher itself
        self._inflight: set[str] = set()
        self._state_lock = threading.Lock()

    def run(self) -> None:
        from concurrent.futures import ThreadPoolExecutor
        from .formats import newest_target_for
        from .engine import get_engine

        self.folder.mkdir(parents=True, exist_ok=True)
        _log(f"watching {self.folder} (auto target: {self.cfg.get('auto_target')})")
        with ThreadPoolExecutor(max_workers=2, thread_name_prefix="clickshift-drop") as pool:
            while not self.stop_event.is_set():
                try:
                    now: dict[str, tuple[int, int]] = {}
                    with os.scandir(self.folder) as it:
                        for entry in it:
                            nm = entry.name
                            if nm.startswith("~$") or nm.endswith((".clickshift-debug.txt", ".confas-debug.txt")):
                                continue
                            try:
                                if not entry.is_file(follow_symlinks=False):
                                    continue
                                st = entry.stat(follow_symlinks=False)
                                now[entry.path] = (st.st_mtime_ns, st.st_size)
                            except OSError:
                                continue
                    with self._state_lock:
                        produced = set(self._produced)
                        inflight = set(self._inflight)
                        done_map = dict(self._done)
                    for path_str, (mtime_ns, size) in now.items():
                        if path_str in produced or path_str in inflight:
                            continue
                        if done_map.get(path_str) == (mtime_ns, size):
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
                        # file is complete -> dispatch once (skip our own outputs)
                        del self._seen[path_str]
                        src = Path(path_str)
                        out_ext = (self.cfg.get("auto_target") or "").lower()
                        if out_ext and src.stem.lower().endswith("_" + out_ext):
                            with self._state_lock:
                                self._done[path_str] = (mtime_ns, size)
                            continue
                        with self._state_lock:
                            self._inflight.add(path_str)
                        pool.submit(self._run_job, path_str, (mtime_ns, size), newest_target_for, get_engine)
                    # forget vanished files
                    now_keys = set(now)
                    for gone in set(self._seen) - now_keys:
                        del self._seen[gone]
                    with self._state_lock:
                        for gone in self._produced - now_keys:
                            self._produced.discard(gone)
                        for gone in set(self._done) - now_keys:
                            del self._done[gone]
                except Exception as exc:
                    _log(f"watch loop error: {exc}")
                self.stop_event.wait(self.POLL_SECONDS)

    def _run_job(self, path_str: str, stamp: tuple[int, int], newest_fn, engine_getter) -> None:
        try:
            self._convert(Path(path_str), newest_fn, engine_getter)
        finally:
            with self._state_lock:
                self._inflight.discard(path_str)
                self._done[path_str] = stamp

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
        popup: Any = None
        try:
            from .updater import ConversionProgressPopup
            popup = ConversionProgressPopup([(str(src), target)]).start()
        except Exception:
            popup = None
        try:
            eng = engine_getter(enable_ml=bool(self.cfg.get("ml_enabled", True)))
            res = eng.convert(src, target)
            if popup is not None:
                popup.on_file_done(res)
                if res.ok:
                    popup.complete(delay_s=0.35)
                else:
                    popup.error(res.error or "Conversion failed", delay_s=1.2)
        except Exception as exc:
            if popup is not None:
                popup.error(str(exc), delay_s=1.2)
            raise
        if res.ok and res.target:
            with self._state_lock:
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

_INSTANCE_MUTEX_HANDLE: Any = None


def _release_single_instance() -> None:
    """Release the Win32 singleton mutex and remove stale daemon.pid on clean exit."""
    global _INSTANCE_MUTEX_HANDLE
    h = _INSTANCE_MUTEX_HANDLE
    _INSTANCE_MUTEX_HANDLE = None
    if h and IS_WINDOWS:
        try:
            k32 = ctypes.WinDLL("kernel32")
            k32.CloseHandle.argtypes = [ctypes.c_void_p]
            k32.CloseHandle(h)
        except Exception:
            pass
    try:
        (_active_data_dir() / "daemon.pid").unlink(missing_ok=True)
    except Exception:
        pass


def acquire_single_instance() -> bool:
    """Return True if this process is the only ClickShift daemon."""
    global _INSTANCE_MUTEX_HANDLE
    d = _active_data_dir()
    pid_file = d / "daemon.pid"
    try:
        d.mkdir(parents=True, exist_ok=True)
        if IS_WINDOWS and getattr(ctypes, "WinDLL", None) is not None:
            k32 = ctypes.WinDLL("kernel32", use_last_error=True)
            k32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
            k32.CreateMutexW.restype = ctypes.c_void_p
            k32.CloseHandle.argtypes = [ctypes.c_void_p]
            norm_dir = os.path.normcase(os.path.abspath(str(d)))
            tag = hashlib.sha256(norm_dir.encode("utf-8")).hexdigest()[:16]
            mutex_name = f"Local\\ClickShiftDaemon_{tag}"
            handle = k32.CreateMutexW(None, False, mutex_name)
            last_err = ctypes.get_last_error()
            if handle:
                ERROR_ALREADY_EXISTS = 183
                if last_err == ERROR_ALREADY_EXISTS:
                    k32.CloseHandle(handle)
                    return False
                _INSTANCE_MUTEX_HANDLE = handle
                pid_file.write_text(str(os.getpid()), encoding="utf-8")
                return True
        if pid_file.exists():
            pid = int(pid_file.read_text().strip() or "0")
            if pid and pid != os.getpid() and _pid_alive(pid):
                return False
        pid_file.write_text(str(os.getpid()), encoding="utf-8")
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


def _recv_line_bounded(conn, max_bytes: int = 65536) -> str:
    """Read a single newline-delimited UTF-8 request capped at ``max_bytes``."""
    chunks: list[bytes] = []
    total = 0
    while True:
        chunk = conn.recv(min(4096, max_bytes - total))
        if not chunk:
            break
        nl = chunk.find(b"\n")
        if nl != -1:
            chunks.append(chunk[:nl])
            break
        chunks.append(chunk)
        total += len(chunk)
        if total >= max_bytes:
            raise ValueError("IPC request exceeded maximum line length")
    return b"".join(chunks).decode("utf-8", errors="strict")


class DaemonIPCServer(threading.Thread):
    """Loopback IPC server so CLI/right-click conversions run inside the warm daemon."""

    MAX_REQUEST_BYTES = 65536
    MAX_CONCURRENT_CONNS = 16

    def __init__(self, cfg: dict) -> None:
        super().__init__(daemon=True, name="clickshift-ipc")
        self.cfg = cfg
        self.token = uuid.uuid4().hex
        self.stop_event = threading.Event()
        self._conn_sem = threading.BoundedSemaphore(self.MAX_CONCURRENT_CONNS)
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(16)
        self._sock.settimeout(1.0)
        self.port = int(self._sock.getsockname()[1])

    def start_and_publish(self) -> None:
        f = ipc_file()
        f.parent.mkdir(parents=True, exist_ok=True)
        _restrict_to_current_user(f.parent)
        try:
            fd = os.open(str(f), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                json.dump({"port": self.port, "token": self.token, "pid": os.getpid()}, fh)
        except Exception:
            f.write_text(json.dumps({"port": self.port, "token": self.token, "pid": os.getpid()}),
                         encoding="utf-8")
        _restrict_to_current_user(f)
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
            if not self._conn_sem.acquire(blocking=False):
                try:
                    conn.close()
                except Exception:
                    pass
                continue
            threading.Thread(target=self._handle_guarded, args=(conn,), daemon=True).start()

    def _handle_guarded(self, conn) -> None:
        try:
            self._handle(conn)
        finally:
            self._conn_sem.release()

    def _handle(self, conn) -> None:
        from .engine import get_engine
        try:
            conn.settimeout(2.0)
            with conn:
                line = _recv_line_bounded(conn, max_bytes=self.MAX_REQUEST_BYTES)
                if not line:
                    return
                req = json.loads(line)
                if not isinstance(req, dict):
                    return
                tok = str(req.get("token") or "")
                if not hmac.compare_digest(tok, self.token):
                    return
                conn.settimeout(300.0)
                wfile = conn.makefile("w", encoding="utf-8")
                enable_ml = bool(req.get("enable_ml", True))
                eng = get_engine(enable_ml=enable_ml)
                write_lock = threading.Lock()

                def on_done(res):
                    with write_lock:
                        wfile.write(json.dumps(res.as_dict()) + "\n")
                        wfile.flush()

                def on_step(info: dict):
                    with write_lock:
                        wfile.write(json.dumps({"type": "progress", **info}) + "\n")
                        wfile.flush()

                if req.get("action") == "combine_images":
                    files = [str(p) for p in (req.get("files") or [])]
                    target = str(req.get("target") or "pdf")
                    if hasattr(eng, "combine_images"):
                        try:
                            on_done(eng.combine_images(files, target, step_progress=on_step))
                        except TypeError:
                            on_done(eng.combine_images(files, target))
                else:
                    jobs = [(str(s), str(t)) for s, t in (req.get("jobs") or [])]
                    try:
                        eng.convert_many(jobs, progress=on_done, step_progress=on_step)
                    except TypeError:
                        eng.convert_many(jobs, progress=on_done)
        except Exception as exc:
            _log(f"ipc handler error: {exc}")


def _ipc_request(payload_fields: dict, progress=None, step_progress=None) -> Optional[list]:
    """Send an authenticated JSON request to the running warm daemon."""
    if os.environ.get("CLICKSHIFT_NO_IPC") == "1":
        return None
    f = ipc_file()
    try:
        raw_meta = f.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None
    except Exception:
        return None
    try:
        from .formats import ConversionResult
        meta = json.loads(raw_meta)
        pid = int(meta.get("pid", 0))
        if pid == os.getpid() or (pid and not _pid_alive(pid)):
            return None
        port = int(meta["port"])
        token = str(meta["token"])
        results = []
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(0.4)
        with sock:
            sock.connect(("127.0.0.1", port))
            try:
                sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            except OSError:
                pass
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
                if isinstance(d, dict) and d.get("type") == "progress":
                    if step_progress is not None:
                        try:
                            step_progress(d)
                        except Exception:
                            pass
                    continue
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
                       progress=None, step_progress=None) -> Optional[list]:
    """Send conversion jobs to the warm daemon if running; return None on any IPC miss."""
    results = _ipc_request(
        {"jobs": jobs, "enable_ml": enable_ml},
        progress=progress,
        step_progress=step_progress,
    )
    return results if (results is not None and len(results) == len(jobs)) else None


def try_daemon_combine(files: list[str], target: str, enable_ml: bool = True,
                       progress=None, step_progress=None) -> Optional[list]:
    """Combine ordered images into one document inside the warm daemon if running."""
    results = _ipc_request(
        {"action": "combine_images", "files": files, "target": target, "enable_ml": enable_ml},
        progress=progress,
        step_progress=step_progress,
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
    # Fallback matching the obsidian / crimson-white stacked pages + cursor icon
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([3, 3, 61, 61], radius=14, fill=(20, 10, 11, 255))
    d.rounded_rectangle([26, 8, 51, 41], radius=4, fill=(225, 29, 46, 255))
    d.rounded_rectangle([13, 16, 39, 51], radius=4, fill=(255, 255, 255, 255), outline=(11, 11, 11, 255), width=2)
    d.polygon([(28, 31), (28, 53), (34, 48), (38, 57), (42, 55), (38, 46), (46, 46)],
              fill=(11, 11, 11, 255), outline=(255, 255, 255, 255))
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
    is_store_fn = getattr(updater_mod, "is_packaged_store_app", None)
    if callable(is_store_fn) and is_store_fn():
        return "store_managed"
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
    if IS_WINDOWS and getattr(sys, "frozen", False):
        _save_exe_path_registry()
        if not is_startup_registered():
            register_startup()
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
            _release_single_instance()
        return 0

    # ---- pystray tray icon ------------------------------------------------ #
    update_stop = threading.Event()
    update_lock = threading.Lock()
    stopped_cleanly = {"v": False}
    try:
        pystray: Any = importlib.import_module("pystray")
        from . import __version__ as _ver
        from . import updater as _updater

        paused = {"v": False}
        auto_loop_started = {"v": False}
        icon_holder: dict[str, Any] = {"icon": None}

        def _stop_for_update() -> None:
            stopped_cleanly["v"] = True
            update_stop.set()
            if watch:
                watch.stop_event.set()
            if ipc:
                ipc.stop()
            _release_single_instance()
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
            stopped_cleanly["v"] = True
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
            if not auto_loop_started["v"] and os.environ.get("CLICKSHIFT_NO_POPUP") != "1":
                auto_loop_started["v"] = True
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
        icon_img = _make_tray_icon_image()
        last_exc: Optional[Exception] = None
        for attempt, delay_s in enumerate((2.0, 4.0, 8.0, 0.0)):
            try:
                icon = pystray.Icon(APP_NAME, icon_img, f"ClickShift v{_ver} - document converter", menu)
                icon_holder["icon"] = icon
                _log("tray daemon started" if attempt == 0 else f"tray daemon started (retry {attempt})")
                icon.run(setup=_on_tray_ready)
                last_exc = None
                break
            except Exception as exc:
                last_exc = exc
                if stopped_cleanly["v"] or delay_s <= 0:
                    raise
                _log(f"tray startup retry after delay ({exc})")
                if update_stop.wait(delay_s):
                    break
        if last_exc is not None and not stopped_cleanly["v"]:
            raise last_exc
        _release_single_instance()
        return 0
    except Exception as exc:
        update_stop.set()
        _log(f"tray unavailable ({exc}); falling back to headless loop")
        try:
            while not stopped_cleanly["v"]:
                time.sleep(3600)
        except KeyboardInterrupt:
            pass
        finally:
            if ipc:
                ipc.stop()
            _release_single_instance()
        return 0
