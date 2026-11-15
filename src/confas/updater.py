"""GitHub auto-updater and branded setup/update progress UI for ConFas."""
from __future__ import annotations

import json
import os
import queue
import re
import ssl
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from confas import APP_NAME, __version__


def _log_info(msg: str) -> None:
    try:
        from confas import monitor
        monitor._log(msg)
    except Exception:
        pass


def _log_error(msg: str) -> None:
    _log_info(f"ERROR: {msg}")

GITHUB_OWNER = "AceX0320"
GITHUB_REPO = "ConFas"
LATEST_RELEASE_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
TAGS_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/tags?per_page=10"
RELEASE_BY_TAG_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/tags/{{tag}}"
RELEASES_PAGE_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

MAX_JSON_BYTES = 2 * 1024 * 1024
MAX_EXE_BYTES = 512 * 1024 * 1024
DEFAULT_MIN_EXE_BYTES = 1024 * 1024

_ALLOWED_API_HOSTS = {"api.github.com"}
_ALLOWED_DOWNLOAD_HOSTS = {
    "github.com",
    "objects.githubusercontent.com",
    "release-assets.githubusercontent.com",
}

_VERSION_RE = re.compile(r"^\s*v?(\d+(?:\.\d+)*)")


@dataclass(frozen=True)
class ReleaseInfo:
    """Metadata for an available ConFas GitHub release."""

    tag_name: str
    version: str
    download_url: str
    release_notes: str = ""
    html_url: str = RELEASES_PAGE_URL
    asset_size: int = 0


def parse_version(version_str: str) -> tuple[int, ...]:
    """Parse a SemVer-like tag or version string (e.g. 'v1.0.11') into ints."""
    if not version_str:
        return ()
    m = _VERSION_RE.match(str(version_str))
    if not m:
        return ()
    return tuple(int(part) for part in m.group(1).split("."))


def is_newer_version(remote_version: str, current_version: str) -> bool:
    """Return True if remote_version is strictly newer than current_version."""
    r = parse_version(remote_version)
    c = parse_version(current_version)
    if not r or not c:
        return False
    width = max(len(r), len(c))
    r_pad = r + (0,) * (width - len(r))
    c_pad = c + (0,) * (width - len(c))
    return r_pad > c_pad


def find_icon_path() -> Optional[Path]:
    """Locate bundled or repository assets/confas.ico without hardcoding user paths."""
    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "assets" / "confas.ico")
    candidates.append(Path(__file__).resolve().parents[2] / "assets" / "confas.ico")
    for p in candidates:
        if p.exists():
            return p
    return None


def is_allowed_download_url(url: str) -> bool:
    """Verify URL points only to HTTPS GitHub release/asset hosts."""
    if not url or not isinstance(url, str):
        return False
    parsed = urllib.parse.urlparse(url.strip())
    if parsed.scheme != "https":
        return False
    host = (parsed.hostname or "").lower()
    if host not in _ALLOWED_DOWNLOAD_HOSTS:
        return False
    if host == "github.com":
        return parsed.path.startswith(f"/{GITHUB_OWNER}/{GITHUB_REPO}/")
    return True


def _is_allowed_api_url(url: str) -> bool:
    parsed = urllib.parse.urlparse(url.strip())
    return parsed.scheme == "https" and (parsed.hostname or "").lower() in _ALLOWED_API_HOSTS


class _SafeRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Reject HTTP redirects outside GitHub release / asset hosts."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if not (is_allowed_download_url(newurl) or _is_allowed_api_url(newurl)):
            raise ValueError(f"Blocked untrusted redirect target: {newurl}")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _fetch_json(url: str, timeout: float = 8.0) -> object:
    if not _is_allowed_api_url(url):
        raise ValueError(f"Blocked API URL: {url}")
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"{APP_NAME}/{__version__}",
        },
        method="GET",
    )
    ctx = ssl.create_default_context()
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=ctx),
        _SafeRedirectHandler(),
    )
    with opener.open(req, timeout=timeout) as resp:
        data = resp.read(MAX_JSON_BYTES + 1)
    if len(data) > MAX_JSON_BYTES:
        raise ValueError("GitHub API response exceeded safety size limit")
    return json.loads(data.decode("utf-8"))


def parse_release_payload(data: object) -> Optional[ReleaseInfo]:
    """Extract ReleaseInfo from a GitHub Release API JSON object."""
    if not isinstance(data, dict):
        return None
    if data.get("draft") or data.get("prerelease"):
        return None
    tag = str(data.get("tag_name") or "").strip()
    if not parse_version(tag):
        return None

    download_url = ""
    asset_size = 0
    assets = data.get("assets")
    if isinstance(assets, list):
        for asset in assets:
            if not isinstance(asset, dict):
                continue
            name = str(asset.get("name") or "")
            url = str(asset.get("browser_download_url") or "")
            if name.lower() == "confas.exe" and is_allowed_download_url(url):
                download_url = url
                asset_size = int(asset.get("size") or 0)
                break
        if not download_url:
            for asset in assets:
                if not isinstance(asset, dict):
                    continue
                name = str(asset.get("name") or "")
                url = str(asset.get("browser_download_url") or "")
                if name.lower().endswith(".exe") and is_allowed_download_url(url):
                    download_url = url
                    asset_size = int(asset.get("size") or 0)
                    break

    if not download_url:
        return None

    html_url = str(data.get("html_url") or RELEASES_PAGE_URL)
    if not is_allowed_download_url(html_url):
        html_url = RELEASES_PAGE_URL

    return ReleaseInfo(
        tag_name=tag,
        version=".".join(str(p) for p in parse_version(tag)),
        download_url=download_url,
        release_notes=str(data.get("body") or "").strip(),
        html_url=html_url,
        asset_size=asset_size,
    )


def check_for_update(
    current_version: str = __version__,
    timeout: float = 8.0,
) -> Optional[ReleaseInfo]:
    """Check GitHub Releases (and newer tags with released assets) for an update."""
    candidate: Optional[ReleaseInfo] = None

    try:
        latest_data = _fetch_json(LATEST_RELEASE_API, timeout=timeout)
        candidate = parse_release_payload(latest_data)
    except urllib.error.HTTPError as e:
        if e.code != 404:
            _log_info(f"updater: latest release check HTTP {e.code}")
            return None
    except Exception as e:
        _log_info(f"updater: latest release check failed: {e}")
        return None

    # Also check recent tags in case a newer version tag already published assets
    # without being marked 'latest'.
    try:
        tags_data = _fetch_json(TAGS_API, timeout=timeout)
        if isinstance(tags_data, list):
            best_tag = ""
            base_ver = candidate.version if candidate else current_version
            for item in tags_data:
                if not isinstance(item, dict):
                    continue
                tname = str(item.get("name") or "").strip()
                if is_newer_version(tname, best_tag or base_ver):
                    best_tag = tname
            if best_tag and is_newer_version(best_tag, base_ver):
                tag_url = RELEASE_BY_TAG_API.format(tag=urllib.parse.quote(best_tag, safe=""))
                try:
                    tagged_rel = parse_release_payload(_fetch_json(tag_url, timeout=timeout))
                    if tagged_rel is not None:
                        candidate = tagged_rel
                except Exception:
                    pass
    except Exception:
        pass

    if candidate and is_newer_version(candidate.version, current_version):
        return candidate
    return None


def download_update(
    url: str,
    dest: Path,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    timeout: float = 120.0,
    min_bytes: int = DEFAULT_MIN_EXE_BYTES,
) -> Path:
    """Download a verified Windows PE executable from GitHub Releases to dest."""
    if not is_allowed_download_url(url):
        raise ValueError(f"Refusing download from non-GitHub URL: {url}")

    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    part = dest.with_suffix(dest.suffix + ".part")

    req = urllib.request.Request(
        url,
        headers={"User-Agent": f"{APP_NAME}/{__version__}"},
        method="GET",
    )
    ctx = ssl.create_default_context()
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=ctx),
        _SafeRedirectHandler(),
    )

    downloaded = 0
    try:
        with opener.open(req, timeout=timeout) as resp:
            if not is_allowed_download_url(resp.geturl()):
                raise ValueError(f"Blocked redirected URL: {resp.geturl()}")
            total = int(resp.headers.get("Content-Length") or 0)
            if total > MAX_EXE_BYTES:
                raise ValueError(f"Update asset too large ({total} bytes)")
            with open(part, "wb") as fh:
                while True:
                    chunk = resp.read(65536)
                    if not chunk:
                        break
                    downloaded += len(chunk)
                    if downloaded > MAX_EXE_BYTES:
                        raise ValueError("Download exceeded maximum allowed size")
                    fh.write(chunk)
                    if progress_cb is not None:
                        progress_cb(downloaded, total)

        size = part.stat().st_size
        if size < min_bytes:
            raise ValueError(f"Downloaded binary too small ({size} bytes)")
        with open(part, "rb") as fh:
            magic = fh.read(2)
        if magic != b"MZ":
            raise ValueError("Downloaded file is not a valid Windows executable")

        os.replace(part, dest)
        return dest
    except Exception:
        try:
            if part.exists():
                part.unlink()
        except Exception:
            pass
        raise


def apply_windows_update(
    downloaded_exe: Path,
    target_exe: Path,
    pid_to_wait: Optional[int] = None,
) -> Path:
    """Spawn a detached Windows helper script that replaces target_exe and relaunches it."""
    downloaded_exe = Path(downloaded_exe).resolve()
    target_exe = Path(target_exe).resolve()
    if not downloaded_exe.exists():
        raise FileNotFoundError(f"Missing downloaded update: {downloaded_exe}")

    pid = pid_to_wait or os.getpid()
    tmp_dir = Path(tempfile.gettempdir())
    script = tmp_dir / f"confas_apply_update_{pid}.cmd"

    # Wait up to ~20s for the running ConFas process to exit cleanly, then force-stop
    # any lingering ConFas lock, swap in the new EXE with retries, relaunch setup/tray,
    # and self-delete.
    src_str = str(downloaded_exe)
    dst_str = str(target_exe)
    bat = (
        "@echo off\r\n"
        "setlocal enableextensions\r\n"
        f"set \"SRC={src_str}\"\r\n"
        f"set \"DST={dst_str}\"\r\n"
        f"set \"OLDPID={pid}\"\r\n"
        "for /L %%I in (1,1,20) do (\r\n"
        "  tasklist /FI \"PID eq %OLDPID%\" 2>nul | find \"%OLDPID%\" >nul\r\n"
        "  if errorlevel 1 goto copy_now\r\n"
        "  if %%I EQU 8 taskkill /F /PID %OLDPID% >nul 2>nul\r\n"
        "  ping 127.0.0.1 -n 2 >nul\r\n"
        ")\r\n"
        ":copy_now\r\n"
        "for /L %%R in (1,1,15) do (\r\n"
        "  move /Y \"%SRC%\" \"%DST%\" >nul 2>nul && goto launch_new\r\n"
        "  copy /Y \"%SRC%\" \"%DST%\" >nul 2>nul && (del /F /Q \"%SRC%\" >nul 2>nul & goto launch_new)\r\n"
        "  ping 127.0.0.1 -n 2 >nul\r\n"
        ")\r\n"
        "exit /b 1\r\n"
        ":launch_new\r\n"
        "start \"\" \"%DST%\"\r\n"
        "(goto) 2>nul & del /F /Q \"%~f0\"\r\n"
    )
    script.write_text(bat, encoding="utf-8")

    creationflags = 0
    for flag in ("CREATE_NO_WINDOW", "DETACHED_PROCESS", "CREATE_NEW_PROCESS_GROUP"):
        creationflags |= getattr(subprocess, flag, 0)

    subprocess.Popen(
        ["cmd.exe", "/c", str(script)],
        creationflags=creationflags,
        close_fds=True,
    )
    return script


class InstallProgressPopup:
    """Thread-safe native progress window for setup and update feedback.

    Suppressed automatically when CONFAS_NO_POPUP=1 or when tkinter is
    unavailable in headless environments.
    """

    def __init__(
        self,
        title: str = "ConFas Setup",
        heading: str = "Setting up ConFas...",
        enabled: Optional[bool] = None,
    ) -> None:
        if enabled is None:
            enabled = (os.name == "nt") and (os.environ.get("CONFAS_NO_POPUP") != "1")
        self.enabled = bool(enabled)
        self.title = title
        self.heading = heading
        self._q: queue.Queue[tuple[str, object]] = queue.Queue()
        self._ready = threading.Event()
        self._closed = threading.Event()
        self._thread: Optional[threading.Thread] = None

    def start(self) -> "InstallProgressPopup":
        if not self.enabled:
            self._ready.set()
            self._closed.set()
            return self
        self._thread = threading.Thread(target=self._run_ui, daemon=True)
        self._thread.start()
        self._ready.wait(timeout=2.0)
        return self

    def update(self, status: str, progress: Optional[float] = None, detail: str = "") -> None:
        if self.enabled and not self._closed.is_set():
            self._q.put(("update", (status, progress, detail)))

    def complete(self, message: str = "Setup complete!", delay_s: float = 0.8) -> None:
        if not self.enabled:
            return
        self._q.put(("complete", (message, delay_s)))
        if self._thread is not None:
            self._thread.join(timeout=max(2.0, delay_s + 2.0))

    def error(self, message: str, delay_s: float = 2.0) -> None:
        if not self.enabled:
            return
        self._q.put(("error", (message, delay_s)))
        if self._thread is not None:
            self._thread.join(timeout=max(3.0, delay_s + 2.0))

    def close(self) -> None:
        if not self.enabled:
            return
        self._q.put(("close", None))
        if self._thread is not None:
            self._thread.join(timeout=2.0)

    def _run_ui(self) -> None:
        try:
            import tkinter as tk
            from tkinter import ttk
        except Exception as e:
            _log_info(f"InstallProgressPopup disabled (tkinter unavailable): {e}")
            self._ready.set()
            self._closed.set()
            return

        root = tk.Tk()
        root.title(self.title)
        root.resizable(False, False)
        root.configure(bg="#0F172A")
        try:
            ico = find_icon_path()
            if ico is not None:
                root.iconbitmap(default=str(ico))
        except Exception:
            pass

        width, height = 430, 178
        sx = root.winfo_screenwidth()
        sy = root.winfo_screenheight()
        root.geometry(f"{width}x{height}+{(sx - width) // 2}+{(sy - height) // 2}")
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass

        style = ttk.Style(root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "ConFas.Horizontal.TProgressbar",
            troughcolor="#1E293B",
            background="#38BDF8",
            bordercolor="#1E293B",
            lightcolor="#38BDF8",
            darkcolor="#0284C7",
            thickness=10,
        )

        frame = tk.Frame(root, bg="#0F172A", padx=22, pady=18)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text=self.heading,
            bg="#0F172A",
            fg="#F8FAFC",
            font=("Segoe UI", 12, "bold"),
            anchor="w",
        ).pack(fill="x")

        status_lbl = tk.Label(
            frame,
            text="Starting...",
            bg="#0F172A",
            fg="#CBD5E1",
            font=("Segoe UI", 9),
            anchor="w",
        )
        status_lbl.pack(fill="x", pady=(8, 8))

        bar_mode = {"indeterminate": True}
        bar = ttk.Progressbar(
            frame,
            style="ConFas.Horizontal.TProgressbar",
            orient="horizontal",
            length=386,
            mode="indeterminate",
            maximum=100,
        )
        bar.pack(fill="x")
        bar.start(14)

        detail_lbl = tk.Label(
            frame,
            text=f"ConFas v{__version__}",
            bg="#0F172A",
            fg="#64748B",
            font=("Segoe UI", 8),
            anchor="w",
        )
        detail_lbl.pack(fill="x", pady=(8, 0))

        self._ready.set()

        def poll() -> None:
            while True:
                try:
                    kind, payload = self._q.get_nowait()
                except queue.Empty:
                    break

                if kind == "update":
                    status_text, pct, detail_text = payload  # type: ignore[misc]
                    if status_text:
                        status_lbl.configure(text=str(status_text))
                    if detail_text:
                        detail_lbl.configure(text=str(detail_text))
                    if pct is not None:
                        if bar_mode["indeterminate"]:
                            bar.stop()
                            bar.configure(mode="determinate")
                            bar_mode["indeterminate"] = False
                        bar["value"] = max(0.0, min(100.0, float(pct)))
                elif kind == "complete":
                    msg, delay_s = payload  # type: ignore[misc]
                    if bar_mode["indeterminate"]:
                        bar.stop()
                        bar.configure(mode="determinate")
                        bar_mode["indeterminate"] = False
                    bar["value"] = 100.0
                    status_lbl.configure(text=str(msg))
                    detail_lbl.configure(text="Ready in system tray")
                    root.after(int(max(0.1, float(delay_s)) * 1000), root.destroy)
                    return
                elif kind == "error":
                    msg, delay_s = payload  # type: ignore[misc]
                    if bar_mode["indeterminate"]:
                        bar.stop()
                        bar.configure(mode="determinate")
                        bar_mode["indeterminate"] = False
                    status_lbl.configure(text=str(msg))
                    detail_lbl.configure(text="Setup encountered an error")
                    root.after(int(max(0.2, float(delay_s)) * 1000), root.destroy)
                    return
                elif kind == "close":
                    root.destroy()
                    return

            root.after(45, poll)

        root.after(45, poll)
        try:
            root.mainloop()
        finally:
            import gc
            try:
                bar.stop()
            except Exception:
                pass
            poll = None  # type: ignore[assignment]
            del status_lbl, detail_lbl, bar, style, frame, root
            gc.collect()
            self._closed.set()


def prompt_update_dialog(info: ReleaseInfo, current_version: str = __version__) -> str:
    """Show a native update prompt asking the user to update, defer, or skip.

    Returns 'update', 'later', or 'skip'.
    """
    if os.environ.get("CONFAS_NO_POPUP") == "1":
        return "later"

    try:
        import tkinter as tk
    except Exception:
        if os.name == "nt":
            try:
                import ctypes
                res = ctypes.windll.user32.MessageBoxW(
                    0,
                    f"ConFas {info.tag_name} is available (you have v{current_version}).\n\nUpdate now?",
                    "ConFas Update Available",
                    0x04 | 0x40,
                )
                return "update" if res == 6 else "later"
            except Exception:
                return "later"
        return "later"

    choice = {"value": "later"}
    root = tk.Tk()
    root.title("ConFas Update Available")
    root.resizable(False, False)
    root.configure(bg="#0F172A")
    try:
        ico = find_icon_path()
        if ico is not None:
            root.iconbitmap(default=str(ico))
    except Exception:
        pass

    notes = (info.release_notes or "").strip()
    if len(notes) > 360:
        notes = notes[:357].rstrip() + "..."
    size_mb = f" ({info.asset_size / (1024 * 1024):.1f} MB)" if info.asset_size else ""

    width, height = 450, 265 if notes else 195
    sx = root.winfo_screenwidth()
    sy = root.winfo_screenheight()
    root.geometry(f"{width}x{height}+{(sx - width) // 2}+{(sy - height) // 2}")
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    wrap = tk.Frame(root, bg="#0F172A", padx=22, pady=18)
    wrap.pack(fill="both", expand=True)

    tk.Label(
        wrap,
        text=f"ConFas {info.tag_name} is available",
        bg="#0F172A",
        fg="#F8FAFC",
        font=("Segoe UI", 12, "bold"),
        anchor="w",
    ).pack(fill="x")

    tk.Label(
        wrap,
        text=f"Installed: v{current_version}   ->   Latest: {info.tag_name}{size_mb}",
        bg="#0F172A",
        fg="#38BDF8",
        font=("Segoe UI", 9),
        anchor="w",
    ).pack(fill="x", pady=(6, 10))

    if notes:
        box = tk.Label(
            wrap,
            text=notes,
            bg="#1E293B",
            fg="#CBD5E1",
            font=("Segoe UI", 8),
            anchor="nw",
            justify="left",
            wraplength=390,
            padx=10,
            pady=8,
        )
        box.pack(fill="both", expand=True, pady=(0, 12))

    btn_row = tk.Frame(wrap, bg="#0F172A")
    btn_row.pack(fill="x", side="bottom")

    def choose(val: str) -> None:
        choice["value"] = val
        root.destroy()

    update_btn = tk.Button(
        btn_row,
        text="Update Now",
        command=lambda: choose("update"),
        bg="#0284C7",
        fg="#FFFFFF",
        activebackground="#0369A1",
        activeforeground="#FFFFFF",
        relief="flat",
        padx=14,
        pady=6,
        font=("Segoe UI", 9, "bold"),
        cursor="hand2",
    )
    update_btn.pack(side="right")
    update_btn.focus_set()

    tk.Button(
        btn_row,
        text="Later",
        command=lambda: choose("later"),
        bg="#334155",
        fg="#E2E8F0",
        activebackground="#475569",
        activeforeground="#FFFFFF",
        relief="flat",
        padx=12,
        pady=6,
        font=("Segoe UI", 9),
        cursor="hand2",
    ).pack(side="right", padx=(0, 8))

    tk.Button(
        btn_row,
        text="Skip version",
        command=lambda: choose("skip"),
        bg="#0F172A",
        fg="#94A3B8",
        activebackground="#1E293B",
        activeforeground="#CBD5E1",
        relief="flat",
        padx=8,
        pady=6,
        font=("Segoe UI", 8),
        cursor="hand2",
    ).pack(side="left")

    root.protocol("WM_DELETE_WINDOW", lambda: choose("later"))
    root.mainloop()
    return choice["value"]


def run_interactive_update(
    info: ReleaseInfo,
    current_version: str = __version__,
    on_before_restart: Optional[Callable[[], None]] = None,
    is_frozen: Optional[bool] = None,
    target_exe: Optional[Path] = None,
) -> str:
    """Prompt the user, download the verified update, apply it, and restart ConFas.

    Returns one of: 'updated', 'opened_browser', 'later', 'skip', 'error'.
    """
    decision = prompt_update_dialog(info, current_version=current_version)
    if decision in ("later", "skip"):
        return decision

    frozen = getattr(sys, "frozen", False) if is_frozen is None else is_frozen
    if not frozen or os.name != "nt":
        webbrowser.open(info.html_url or RELEASES_PAGE_URL)
        return "opened_browser"

    dest_exe = Path(target_exe or sys.executable).resolve()
    staging_exe = Path(tempfile.gettempdir()) / f"ConFas-update-{info.tag_name}.exe"

    popup = InstallProgressPopup(
        title="Updating ConFas",
        heading=f"Updating to ConFas {info.tag_name}",
    ).start()

    try:
        popup.update("Downloading new ConFas release...", progress=0.0, detail=info.tag_name)

        def on_progress(done: int, total: int) -> None:
            if total > 0:
                pct = min(95.0, (done / total) * 95.0)
                mb_done = done / (1024 * 1024)
                mb_tot = total / (1024 * 1024)
                popup.update(
                    f"Downloading ConFas {info.tag_name}...",
                    progress=pct,
                    detail=f"{mb_done:.1f} MB / {mb_tot:.1f} MB",
                )
            else:
                popup.update(
                    f"Downloading ConFas {info.tag_name}...",
                    progress=None,
                    detail=f"{done / (1024 * 1024):.1f} MB downloaded",
                )

        download_update(info.download_url, staging_exe, progress_cb=on_progress)
        popup.update("Installing update and restarting...", progress=98.0, detail="Replacing executable")
        apply_windows_update(staging_exe, dest_exe, pid_to_wait=os.getpid())
        popup.complete(f"Restarting into {info.tag_name}...", delay_s=0.4)
        if on_before_restart is not None:
            on_before_restart()
        return "updated"
    except Exception as e:
        _log_error(f"updater: failed to apply {info.tag_name}: {e}")
        popup.error(f"Update failed: {e}", delay_s=2.5)
        return "error"
