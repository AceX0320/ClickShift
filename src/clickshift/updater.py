"""GitHub auto-updater and branded setup/update progress UI for ClickShift."""
from __future__ import annotations

import base64
import ctypes
import hashlib
import hmac
import json
import os
import queue
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable, Optional, Sequence

if TYPE_CHECKING:
    import ssl as _ssl_t

from clickshift import APP_NAME, __version__


def _log_info(msg: str) -> None:
    try:
        from clickshift import monitor
        monitor._log(msg)
    except Exception:
        pass


def _log_error(msg: str) -> None:
    _log_info(f"ERROR: {msg}")

GITHUB_OWNER = "AceX0320"
GITHUB_REPO = "ClickShift"
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

# Dual-slot ECDSA P-256 public key list for zero-downtime key rotation:
# Add new keys here (base64 of 72-byte CNG ECCPUBLICBLOB, 65-byte X9.62, or SPKI DER)
# or override via CLICKSHIFT_UPDATE_PUBKEYS_B64 during testing.
TRUSTED_ECDSA_PUBKEYS_B64: tuple[str, ...] = (
    "RUNTMSAAAAA51GhTNj8aBWoIBxWojEWaMNYmLZBR+9rwzScwqaFE/nMIqrMDoTuPrnN/VFzAM8Biv7QwpRHdLnn6z9Op3OhU",
)


def active_trusted_pubkeys() -> tuple[str, ...]:
    env = os.environ.get("CLICKSHIFT_UPDATE_PUBKEYS_B64", "").strip()
    if env:
        return tuple(k.strip() for k in env.split(",") if k.strip())
    return TRUSTED_ECDSA_PUBKEYS_B64


@dataclass(frozen=True)
class ReleaseInfo:
    """Metadata for an available ClickShift GitHub release."""

    tag_name: str
    version: str
    download_url: str
    release_notes: str = ""
    html_url: str = RELEASES_PAGE_URL
    asset_size: int = 0
    sha256_url: str = ""
    sig_url: str = ""
    expected_sha256: str = ""


_STORE_IDENTITY_CACHE: Optional[bool] = None


def is_packaged_store_app() -> bool:
    """Return True only when running inside ClickShift's own MSIX / Store package identity."""
    global _STORE_IDENTITY_CACHE
    if sys.platform != "win32" or os.environ.get("CLICKSHIFT_ALLOW_SELF_UPDATE") == "1":
        return False
    if _STORE_IDENTITY_CACHE is not None:
        return _STORE_IDENTITY_CACHE
    try:
        length = ctypes.c_uint32(0)
        rc = ctypes.windll.kernel32.GetCurrentPackageFullName(ctypes.byref(length), None)
        if rc != 122 or not length.value:
            _STORE_IDENTITY_CACHE = False
            return False
        buf = ctypes.create_unicode_buffer(length.value)
        rc2 = ctypes.windll.kernel32.GetCurrentPackageFullName(ctypes.byref(length), buf)
        if rc2 != 0:
            _STORE_IDENTITY_CACHE = False
            return False
        pkg_low = buf.value.lower()
        _STORE_IDENTITY_CACHE = (
            pkg_low.startswith(("clickshift", "confas"))
            and not pkg_low.startswith(("clickshift.shellextension_", "confas.shellextension_"))
        )
        return _STORE_IDENTITY_CACHE
    except Exception:
        _STORE_IDENTITY_CACHE = False
        return False


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
    """Locate bundled or repository assets/clickshift.ico without hardcoding user paths."""
    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "assets" / "clickshift.ico")
    candidates.append(Path(__file__).resolve().parents[2] / "assets" / "clickshift.ico")
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


_SSL_CTX_LOCK = threading.Lock()
_SSL_CTX: Optional["_ssl_t.SSLContext"] = None


def _get_ssl_context() -> "_ssl_t.SSLContext":
    global _SSL_CTX
    if _SSL_CTX is None:
        with _SSL_CTX_LOCK:
            if _SSL_CTX is None:
                import ssl

                _SSL_CTX = ssl.create_default_context()
    assert _SSL_CTX is not None
    return _SSL_CTX


def _fetch_json(url: str, timeout: float = 8.0) -> object:
    if not _is_allowed_api_url(url):
        raise ValueError(f"Blocked API URL: {url}")
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "Accept-Encoding": "gzip",
            "User-Agent": f"{APP_NAME}/{__version__}",
        },
        method="GET",
    )
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=_get_ssl_context()),
        _SafeRedirectHandler(),
    )
    with opener.open(req, timeout=timeout) as resp:
        data = resp.read(MAX_JSON_BYTES + 1)
        resp_headers = getattr(resp, "headers", None)
    if len(data) > MAX_JSON_BYTES:
        raise ValueError("GitHub API response exceeded safety size limit")
    if resp_headers and str(resp_headers.get("Content-Encoding", "")).lower() == "gzip":
        import gzip

        data = gzip.decompress(data)
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
    tag_ver = parse_version(tag)
    if not tag_ver:
        return None
    ver_str = ".".join(str(p) for p in tag_ver)

    download_url = ""
    matched_asset_name = ""
    asset_size = 0
    sha256_url = ""
    sig_url = ""
    assets = data.get("assets")
    if isinstance(assets, list):
        # Single pass over assets while preserving strict priority ordering
        preferred = (
            f"clickshift-{tag}.exe".lower(),
            f"clickshift-v{ver_str}.exe".lower(),
            "clickshift.exe",
        )
        pref_rank = {n: i for i, n in enumerate(preferred)}
        best_rank = len(preferred) + 1
        sig_by_lower: dict[str, str] = {}
        for asset in assets:
            if not isinstance(asset, dict):
                continue
            name = str(asset.get("name") or "")
            url = str(asset.get("browser_download_url") or "")
            nl = name.lower()
            if nl == "sha256sums.txt":
                if is_allowed_download_url(url):
                    sha256_url = url
                continue
            if nl.endswith(".sig"):
                if is_allowed_download_url(url):
                    sig_by_lower[nl] = url
                continue
            rank = pref_rank.get(nl)
            if rank is None and nl.startswith(("clickshift", "confas")) and nl.endswith(".exe"):
                rank = len(preferred)
            if rank is not None and rank < best_rank and is_allowed_download_url(url):
                best_rank = rank
                download_url = url
                matched_asset_name = name
                asset_size = int(asset.get("size") or 0)
        if matched_asset_name:
            sig_url = sig_by_lower.get(f"{matched_asset_name.lower()}.sig", "")

    if not download_url:
        return None

    html_url = str(data.get("html_url") or RELEASES_PAGE_URL)
    if not is_allowed_download_url(html_url):
        html_url = RELEASES_PAGE_URL

    return ReleaseInfo(
        tag_name=tag,
        version=ver_str,
        download_url=download_url,
        release_notes=str(data.get("body") or "").strip(),
        html_url=html_url,
        asset_size=asset_size,
        sha256_url=sha256_url,
        sig_url=sig_url,
    )


def _fetch_text_asset(url: str, max_bytes: int = 65536, timeout: float = 15.0) -> bytes:
    if not is_allowed_download_url(url):
        raise ValueError(f"Blocked untrusted asset URL: {url}")
    req = urllib.request.Request(
        url,
        headers={"User-Agent": f"{APP_NAME}/{__version__}"},
        method="GET",
    )
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=_get_ssl_context()),
        _SafeRedirectHandler(),
    )
    with opener.open(req, timeout=timeout) as resp:
        if not is_allowed_download_url(resp.geturl()):
            raise ValueError(f"Blocked redirected asset URL: {resp.geturl()}")
        data = resp.read(max_bytes + 1)
    if len(data) > max_bytes:
        raise ValueError("Metadata asset exceeded maximum allowed size")
    return data


def parse_sha256sums(text: str, asset_filename: str) -> Optional[str]:
    """Extract lowercase SHA-256 hex digest for ``asset_filename`` from SHA256SUMS.txt."""
    want = Path(asset_filename).name.lower()
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r"^([0-9a-fA-F]{64})\s+\*?(.+)$", line)
        if not m:
            continue
        digest, fname = m.group(1).lower(), Path(m.group(2).strip()).name.lower()
        if fname == want:
            return digest
    return None


def _to_bcrypt_p256_blob(raw_or_b64: str) -> bytes:
    raw = base64.b64decode(raw_or_b64.strip())
    if len(raw) == 72 and raw[:4] == b"ECS1":
        return raw
    if len(raw) == 65 and raw[0] == 0x04:
        return struct.pack("<II", 0x31534345, 32) + raw[1:65]
    if len(raw) == 91 and raw[-65] == 0x04:  # X.509 SPKI DER for P-256
        return struct.pack("<II", 0x31534345, 32) + raw[-64:]
    raise ValueError("Unsupported ECDSA P-256 public key format")


def _decode_ecdsa_sig_raw64(sig_bytes: bytes) -> bytes:
    """Accept raw 64-byte (r||s), base64-encoded, or ASN.1 DER ECDSA P-256 signatures."""
    data = sig_bytes.strip()
    if len(data) == 64:
        return data
    if data and data.isascii():
        try:
            decoded = base64.b64decode(data, validate=True)
            if len(decoded) == 64:
                return decoded
            data = decoded
        except Exception:
            pass
    # ASN.1 DER: 0x30 <len> 0x02 <rlen> <r> 0x02 <slen> <s>
    if len(data) >= 8 and data[0] == 0x30:
        idx = 2 if data[1] < 0x80 else 3
        if idx < len(data) and data[idx] == 0x02:
            rlen = data[idx + 1]
            r = data[idx + 2:idx + 2 + rlen].lstrip(b"\x00")
            pos = idx + 2 + rlen
            if pos + 2 <= len(data) and data[pos] == 0x02:
                slen = data[pos + 1]
                s = data[pos + 2:pos + 2 + slen].lstrip(b"\x00")
                if len(r) <= 32 and len(s) <= 32:
                    return r.rjust(32, b"\x00") + s.rjust(32, b"\x00")
    if len(sig_bytes) == 64:
        return sig_bytes
    raise ValueError("Invalid ECDSA P-256 signature encoding")


_BCRYPT_LOCK = threading.Lock()
_BCRYPT_DLL: Any = None


def _get_bcrypt_dll() -> Any:
    global _BCRYPT_DLL
    if _BCRYPT_DLL is None:
        with _BCRYPT_LOCK:
            if _BCRYPT_DLL is None:
                dll: Any = ctypes.WinDLL("bcrypt")
                dll.BCryptOpenAlgorithmProvider.argtypes = [
                    ctypes.POINTER(ctypes.c_void_p), ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint32
                ]
                dll.BCryptOpenAlgorithmProvider.restype = ctypes.c_uint32
                dll.BCryptCloseAlgorithmProvider.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
                dll.BCryptCloseAlgorithmProvider.restype = ctypes.c_uint32
                dll.BCryptImportKeyPair.argtypes = [
                    ctypes.c_void_p, ctypes.c_void_p, ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_void_p),
                    ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32,
                ]
                dll.BCryptImportKeyPair.restype = ctypes.c_uint32
                dll.BCryptDestroyKey.argtypes = [ctypes.c_void_p]
                dll.BCryptDestroyKey.restype = ctypes.c_uint32
                dll.BCryptVerifySignature.argtypes = [
                    ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint32,
                    ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32,
                ]
                dll.BCryptVerifySignature.restype = ctypes.c_uint32
                _BCRYPT_DLL = dll
    return _BCRYPT_DLL


def verify_ecdsa_p256_sha256(
    sha256_digest: bytes,
    sig_bytes: bytes,
    pubkeys_b64: Sequence[str],
) -> bool:
    """Verify a SHA-256 digest against any trusted ECDSA P-256 public key."""
    if len(sha256_digest) != 32 or not pubkeys_b64:
        return False
    try:
        sig64 = _decode_ecdsa_sig_raw64(sig_bytes)
    except ValueError:
        return False

    if sys.platform == "win32":
        try:
            bcrypt = _get_bcrypt_dll()

            h_alg = ctypes.c_void_p()
            if bcrypt.BCryptOpenAlgorithmProvider(ctypes.byref(h_alg), "ECDSA_P256", None, 0) != 0:
                return False
            try:
                digest_buf = (ctypes.c_ubyte * 32).from_buffer_copy(sha256_digest)
                sig_buf = (ctypes.c_ubyte * 64).from_buffer_copy(sig64)
                for key_b64 in pubkeys_b64:
                    try:
                        blob = _to_bcrypt_p256_blob(key_b64)
                    except Exception:
                        continue
                    blob_buf = (ctypes.c_ubyte * len(blob)).from_buffer_copy(blob)
                    h_key = ctypes.c_void_p()
                    if bcrypt.BCryptImportKeyPair(
                        h_alg, None, "ECCPUBLICBLOB", ctypes.byref(h_key), blob_buf, len(blob), 0
                    ) != 0:
                        continue
                    try:
                        if bcrypt.BCryptVerifySignature(
                            h_key, None, digest_buf, 32, sig_buf, 64, 0
                        ) == 0:
                            return True
                    finally:
                        bcrypt.BCryptDestroyKey(h_key)
                return False
            finally:
                bcrypt.BCryptCloseAlgorithmProvider(h_alg, 0)
        except Exception:
            pass

    # Fallback (Linux/macOS tests or when cryptography is installed)
    try:
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.asymmetric import ec
        from cryptography.hazmat.primitives.asymmetric.utils import Prehashed, encode_dss_signature

        r = int.from_bytes(sig64[:32], "big")
        s = int.from_bytes(sig64[32:], "big")
        der_sig = encode_dss_signature(r, s)
        for key_b64 in pubkeys_b64:
            try:
                blob = _to_bcrypt_p256_blob(key_b64)
                x962 = b"\x04" + blob[8:72]
                pub = ec.EllipticCurvePublicKey.from_encoded_point(ec.SECP256R1(), x962)
                pub.verify(der_sig, sha256_digest, ec.ECDSA(Prehashed(hashes.SHA256())))
                return True
            except Exception:
                continue
    except Exception:
        pass
    return False


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
    expected_sha256: str = "",
    sha256_url: str = "",
    sig_url: str = "",
    trusted_pubkeys: Optional[Sequence[str]] = None,
    require_signature: bool = False,
) -> Path:
    """Download and verify a Windows PE executable from GitHub Releases to dest."""
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
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=_get_ssl_context()),
        _SafeRedirectHandler(),
    )

    downloaded = 0
    magic = b""
    last_cb_bytes = -1
    last_cb_ts = 0.0
    sha = hashlib.sha256()
    try:
        with opener.open(req, timeout=timeout) as resp:
            if not is_allowed_download_url(resp.geturl()):
                raise ValueError(f"Blocked redirected URL: {resp.geturl()}")
            total = int(resp.headers.get("Content-Length") or 0)
            if total > MAX_EXE_BYTES:
                raise ValueError(f"Update asset too large ({total} bytes)")
            with open(part, "wb") as fh:
                while True:
                    chunk = resp.read(524_288)
                    if not chunk:
                        break
                    if downloaded == 0:
                        magic = chunk[:2]
                    elif len(magic) < 2:
                        magic += chunk[: 2 - len(magic)]
                    downloaded += len(chunk)
                    if downloaded > MAX_EXE_BYTES:
                        raise ValueError("Download exceeded maximum allowed size")
                    sha.update(chunk)
                    fh.write(chunk)
                    if progress_cb is not None:
                        now = time.monotonic()
                        if downloaded - last_cb_bytes >= 262_144 or (now - last_cb_ts) >= 0.05:
                            last_cb_bytes = downloaded
                            last_cb_ts = now
                            progress_cb(downloaded, total)
        if progress_cb is not None and downloaded != last_cb_bytes:
            progress_cb(downloaded, total)

        if downloaded < min_bytes:
            raise ValueError(f"Downloaded binary too small ({downloaded} bytes)")
        if magic != b"MZ":
            raise ValueError("Downloaded file is not a valid Windows executable")

        actual_hex = sha.hexdigest().lower()
        want_hex = (expected_sha256 or "").strip().lower()
        if not want_hex and sha256_url:
            sums_txt = _fetch_text_asset(sha256_url).decode("utf-8", errors="replace")
            asset_name = Path(urllib.parse.urlparse(url).path).name
            parsed_hex = parse_sha256sums(sums_txt, asset_name)
            if not parsed_hex:
                raise ValueError(f"SHA256SUMS.txt missing entry for {asset_name}")
            want_hex = parsed_hex
        if want_hex and not hmac.compare_digest(actual_hex, want_hex):
            raise ValueError("Downloaded binary SHA-256 mismatch")

        pubkeys = tuple(trusted_pubkeys) if trusted_pubkeys is not None else active_trusted_pubkeys()
        if pubkeys or require_signature:
            if not sig_url:
                raise ValueError("Missing cryptographic signature (.sig) for release")
            sig_bytes = _fetch_text_asset(sig_url, max_bytes=16384)
            if not verify_ecdsa_p256_sha256(sha.digest(), sig_bytes, pubkeys):
                raise ValueError("Downloaded binary failed ECDSA signature verification")

        os.replace(part, dest)
        return dest
    except Exception:
        try:
            if part.exists():
                part.unlink()
        except Exception:
            pass
        raise


def _resolve_staging_dir(data_dir: Optional[Path] = None) -> Path:
    """Create a private per-update staging folder outside shared %TEMP%."""
    base_env = os.environ.get("CLICKSHIFT_STAGING_DIR", "").strip()
    if base_env:
        root = Path(base_env)
    elif data_dir is not None:
        root = Path(data_dir) / "staging"
    elif sys.platform == "win32":
        local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        root = local_app / "ClickShift" / "staging"
    else:
        root = _resolve_data_dir(data_dir) / "staging"
    stage = root / f"upd-{uuid.uuid4().hex[:12]}"
    stage.mkdir(parents=True, exist_ok=True)
    try:
        from clickshift.monitor import _restrict_to_current_user
        _restrict_to_current_user(root)
        _restrict_to_current_user(stage)
    except Exception:
        pass
    return stage


def cleanup_replaced_exe(old_pid: int, old_exe: Path, wait_ms: int = 20000) -> bool:
    """Wait on the exited previous EXE PID and remove the superseded binary safely."""
    try:
        old_path = Path(old_exe).resolve()
        cur_dir = Path(sys.executable).resolve().parent
        # Safety check: never delete outside our own install/staging folders or non-ClickShift names
        allowed_name = bool(re.match(r"^(?:clickshift|confas).*\.exe(?:\.old)?$", old_path.name, re.I))
        in_app_dir = (old_path.parent == cur_dir or "clickshift" in str(old_path.parent).lower())
        if not (allowed_name and in_app_dir):
            return False

        if sys.platform == "win32" and old_pid > 0 and old_pid != os.getpid():
            SYNCHRONIZE = 0x00100000
            h = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, False, int(old_pid))
            if h:
                try:
                    ctypes.windll.kernel32.WaitForSingleObject(h, int(wait_ms))
                finally:
                    ctypes.windll.kernel32.CloseHandle(h)

        for _ in range(15):
            if not old_path.exists():
                return True
            try:
                old_path.unlink()
                return True
            except OSError:
                time.sleep(0.5)

        if sys.platform == "win32" and old_path.exists():
            MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004
            ctypes.windll.kernel32.MoveFileExW(str(old_path), None, MOVEFILE_DELAY_UNTIL_REBOOT)
    except Exception:
        pass
    return False


def _place_executable(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if src.resolve() == dst.resolve():
        return
    try:
        os.replace(src, dst)
    except OSError:
        tmp_dst = dst.with_suffix(".exe.stage")
        shutil.copy2(src, tmp_dst)
        os.replace(tmp_dst, dst)
        try:
            src.unlink(missing_ok=True)
        except OSError:
            pass
    # Clear Mark-of-the-Web ADS if present so Windows doesn't repeatedly warn on our signed/verified update
    if sys.platform == "win32":
        try:
            os.unlink(f"{dst}:Zone.Identifier")
        except OSError:
            pass


def apply_windows_update(
    downloaded_exe: Path,
    target_exe: Path,
    pid_to_wait: Optional[int] = None,
    new_exe_path: Optional[Path] = None,
    relaunch_args: Optional[list[str]] = None,
) -> Path:
    """Install ``downloaded_exe`` natively without PowerShell bypass scripts and relaunch."""
    downloaded_exe = Path(downloaded_exe).resolve()
    target_exe = Path(target_exe).resolve()
    final_exe = Path(new_exe_path).resolve() if new_exe_path is not None else target_exe
    if not downloaded_exe.exists():
        raise FileNotFoundError(f"Missing downloaded update: {downloaded_exe}")

    pid = int(pid_to_wait or os.getpid())
    old_to_cleanup: Optional[Path] = None

    if final_exe == target_exe and target_exe.exists():
        backup = target_exe.with_name(f"{target_exe.name}.{pid}.old")
        os.replace(target_exe, backup)
        _place_executable(downloaded_exe, final_exe)
        old_to_cleanup = backup
    else:
        _place_executable(downloaded_exe, final_exe)
        if target_exe != final_exe and target_exe.exists():
            old_to_cleanup = target_exe

    child_env = {
        k: v for k, v in os.environ.items()
        if k != "_MEIPASS2" and not k.startswith("_PYI_")
    }
    child_env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"

    cmd = [str(final_exe)]
    if relaunch_args:
        cmd.extend(relaunch_args)
    if old_to_cleanup is not None:
        cmd.extend(["--cleanup-pid", str(pid), "--cleanup-old", str(old_to_cleanup)])

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(
        subprocess, "CREATE_NEW_PROCESS_GROUP", 0
    )
    subprocess.Popen(
        cmd,
        cwd=str(final_exe.parent),
        env=child_env,
        creationflags=creationflags,
        close_fds=True,
    )
    return final_exe


class InstallProgressPopup:
    """Thread-safe native progress window for setup and update feedback.

    Suppressed automatically when CLICKSHIFT_NO_POPUP=1 or when tkinter is
    unavailable in headless environments.
    """

    def __init__(
        self,
        title: str = "ClickShift Setup",
        heading: str = "Setting up ClickShift...",
        enabled: Optional[bool] = None,
    ) -> None:
        if enabled is None:
            enabled = (os.name == "nt") and (os.environ.get("CLICKSHIFT_NO_POPUP") != "1")
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

        def _run_mainloop() -> None:
            root = tk.Tk()
            root.title(self.title)
            root.resizable(False, False)
            root.configure(bg="#140A0B")
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
                "ClickShift.Horizontal.TProgressbar",
                troughcolor="#2A1215",
                background="#E11D2E",
                bordercolor="#2A1215",
                lightcolor="#FF3B4A",
                darkcolor="#B80F1F",
                thickness=10,
            )

            frame = tk.Frame(root, bg="#140A0B", padx=22, pady=18)
            frame.pack(fill="both", expand=True)

            tk.Label(
                frame,
                text=self.heading,
                bg="#140A0B",
                fg="#FFFFFF",
                font=("Segoe UI", 12, "bold"),
                anchor="w",
            ).pack(fill="x")

            status_lbl = tk.Label(
                frame,
                text="Starting...",
                bg="#140A0B",
                fg="#E9E9E9",
                font=("Segoe UI", 9),
                anchor="w",
            )
            status_lbl.pack(fill="x", pady=(8, 8))

            bar_mode = {"indeterminate": True}
            bar = ttk.Progressbar(
                frame,
                style="ClickShift.Horizontal.TProgressbar",
                orient="horizontal",
                length=386,
                mode="indeterminate",
                maximum=100,
            )
            bar.pack(fill="x")
            bar.start(14)

            detail_lbl = tk.Label(
                frame,
                text=f"ClickShift v{__version__}",
                bg="#140A0B",
                fg="#A1A1AA",
                font=("Segoe UI", 8),
                anchor="w",
            )
            detail_lbl.pack(fill="x", pady=(8, 0))

            self._ready.set()

            def _destroy_cleanly() -> None:
                try:
                    bar.stop()
                except Exception:
                    pass
                try:
                    root.destroy()
                except Exception:
                    pass

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
                        root.after(int(max(0.1, float(delay_s)) * 1000), _destroy_cleanly)
                        return
                    elif kind == "error":
                        msg, delay_s = payload  # type: ignore[misc]
                        if bar_mode["indeterminate"]:
                            bar.stop()
                            bar.configure(mode="determinate")
                            bar_mode["indeterminate"] = False
                        status_lbl.configure(text=str(msg))
                        detail_lbl.configure(text="Setup encountered an error")
                        root.after(int(max(0.2, float(delay_s)) * 1000), _destroy_cleanly)
                        return
                    elif kind == "close":
                        _destroy_cleanly()
                        return

                root.after(45, poll)

            root.after(45, poll)
            root.mainloop()

        import gc
        try:
            _run_mainloop()
        finally:
            gc.collect()
            self._closed.set()


def _norm_vtag(ver: str) -> str:
    v = str(ver or "").strip()
    return v if v.startswith("v") else f"v{v}"


def _resolve_data_dir(data_dir: Optional[Path] = None) -> Path:
    if data_dir is not None:
        return Path(data_dir)
    try:
        from clickshift import monitor
        return monitor.data_dir()
    except Exception:
        return Path(os.environ.get("CLICKSHIFT_HOME", str(Path.home() / ".clickshift")))


def update_notice_file(data_dir: Optional[Path] = None) -> Path:
    return _resolve_data_dir(data_dir) / "update_notice.json"


def write_update_notice(
    from_version: str,
    to_version: str,
    silent: bool = True,
    data_dir: Optional[Path] = None,
) -> Path:
    """Persist pending post-update notification marker for the restarted process."""
    p = update_notice_file(data_dir)
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "from_version": _norm_vtag(from_version),
        "to_version": _norm_vtag(to_version),
        "silent": bool(silent),
        "ts": time.time(),
    }
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload), encoding="utf-8")
    tmp.replace(p)
    return p


def peek_update_notice(data_dir: Optional[Path] = None) -> Optional[dict]:
    """Read pending update notice without removing it."""
    p = update_notice_file(data_dir)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else None
    except FileNotFoundError:
        return None
    except Exception:
        return None


def consume_update_notice(data_dir: Optional[Path] = None) -> Optional[dict]:
    """Read and clear the pending post-update notification marker."""
    p = update_notice_file(data_dir)
    try:
        raw = p.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None
    except Exception:
        try:
            p.unlink(missing_ok=True)
        except Exception:
            pass
        return None
    try:
        p.unlink(missing_ok=True)
    except Exception:
        pass
    try:
        data = json.loads(raw)
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def format_update_notification(notice: dict) -> tuple[str, str]:
    """Return (title, message) summarizing an applied auto-update."""
    from_v = str(notice.get("from_version") or "").strip()
    to_v = str(notice.get("to_version") or "").strip()
    if from_v and to_v:
        return (
            f"{APP_NAME} Updated",
            f"{APP_NAME} updated automatically from {from_v} to {to_v}.",
        )
    return (
        f"{APP_NAME} Updated",
        f"{APP_NAME} updated automatically to v{__version__}.",
    )


def prompt_update_dialog(info: ReleaseInfo, current_version: str = __version__) -> str:
    """Show a minimal update prompt with current/available version and an Update button.

    Returns 'update' or 'later'.
    """
    if os.environ.get("CLICKSHIFT_NO_POPUP") == "1":
        return "later"

    try:
        import tkinter as tk
    except Exception:
        if os.name == "nt":
            try:
                import ctypes
                res = ctypes.windll.user32.MessageBoxW(
                    0,
                    f"Current version: v{current_version}\nAvailable version: {info.tag_name}",
                    "ClickShift Update",
                    0x04 | 0x40,
                )
                return "update" if res == 6 else "later"
            except Exception:
                return "later"
        return "later"

    choice = {"value": "later"}
    root = tk.Tk()
    root.title("ClickShift Update")
    root.resizable(False, False)
    root.configure(bg="#140A0B")
    try:
        ico = find_icon_path()
        if ico is not None:
            root.iconbitmap(default=str(ico))
    except Exception:
        pass

    wrap = tk.Frame(root, bg="#140A0B", padx=24, pady=20)
    wrap.pack(fill="both", expand=True)

    tk.Label(
        wrap,
        text=f"Installed: v{current_version}",
        bg="#140A0B",
        fg="#E9E9E9",
        font=("Segoe UI", 10),
        anchor="w",
    ).pack(fill="x")

    tk.Label(
        wrap,
        text=f"Available: {info.tag_name}",
        bg="#140A0B",
        fg="#FF3B4A",
        font=("Segoe UI", 11, "bold"),
        anchor="w",
    ).pack(fill="x", pady=(6, 16))

    def choose(val: str) -> None:
        choice["value"] = val
        root.destroy()

    update_btn = tk.Button(
        wrap,
        text="Update",
        command=lambda: choose("update"),
        bg="#E11D2E",
        fg="#FFFFFF",
        activebackground="#B80F1F",
        activeforeground="#FFFFFF",
        relief="flat",
        padx=20,
        pady=7,
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
    )
    update_btn.pack(fill="x")
    update_btn.focus_set()

    root.update_idletasks()
    width = max(300, root.winfo_reqwidth())
    height = root.winfo_reqheight()
    sx = root.winfo_screenwidth()
    sy = root.winfo_screenheight()
    root.geometry(f"{width}x{height}+{(sx - width) // 2}+{(sy - height) // 2}")
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    root.protocol("WM_DELETE_WINDOW", lambda: choose("later"))
    root.mainloop()
    return choice["value"]


def _download_verified_release(
    info: ReleaseInfo,
    dest: Path,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> Path:
    """Call download_update with typed arguments (avoids **dict unpacking and preserves stubbed callbacks)."""
    if info.expected_sha256 or info.sha256_url or info.sig_url:
        return download_update(
            info.download_url,
            dest,
            progress_cb=progress_cb,
            expected_sha256=info.expected_sha256,
            sha256_url=info.sha256_url,
            sig_url=info.sig_url,
        )
    if progress_cb is not None:
        return download_update(info.download_url, dest, progress_cb=progress_cb)
    return download_update(info.download_url, dest)


def run_background_update(
    info: ReleaseInfo,
    current_version: str = __version__,
    on_before_restart: Optional[Callable[[], None]] = None,
    is_frozen: Optional[bool] = None,
    target_exe: Optional[Path] = None,
    data_dir: Optional[Path] = None,
) -> str:
    """Silently download, apply, and restart ClickShift in the background without prompts.

    Returns one of: 'updated', 'not_frozen', 'store_managed', 'error'.
    """
    if is_packaged_store_app():
        _log_info("updater: Microsoft Store package detected; OS Store manages updates")
        return "store_managed"
    frozen = getattr(sys, "frozen", False) if is_frozen is None else is_frozen
    if not frozen or sys.platform != "win32":
        _log_info(f"updater: {info.tag_name} available (skipping silent replace in source/non-Windows run)")
        return "not_frozen"

    dest_exe = Path(target_exe or sys.executable).resolve()
    versioned_exe = dest_exe.parent / f"{APP_NAME}-{info.tag_name}.exe"
    stage_dir = _resolve_staging_dir(data_dir)
    staging_exe = stage_dir / f"ClickShift-update-{info.tag_name}.exe"

    try:
        _log_info(f"updater: silently downloading {info.tag_name} in background")
        _download_verified_release(info, staging_exe)
        write_update_notice(current_version, info.tag_name, silent=True, data_dir=data_dir)
        apply_windows_update(
            staging_exe,
            dest_exe,
            pid_to_wait=os.getpid(),
            new_exe_path=versioned_exe,
            relaunch_args=["--post-update"],
        )
        if on_before_restart is not None:
            on_before_restart()
        return "updated"
    except Exception as e:
        _log_error(f"updater: background update to {info.tag_name} failed: {e}")
        return "error"


def run_interactive_update(
    info: ReleaseInfo,
    current_version: str = __version__,
    on_before_restart: Optional[Callable[[], None]] = None,
    is_frozen: Optional[bool] = None,
    target_exe: Optional[Path] = None,
    data_dir: Optional[Path] = None,
) -> str:
    """Prompt the user, download the verified update, apply it, and restart ClickShift.

    Returns one of: 'updated', 'opened_browser', 'later', 'skip', 'store_managed', 'error'.
    """
    if is_packaged_store_app():
        return "store_managed"
    decision = prompt_update_dialog(info, current_version=current_version)
    if decision in ("later", "skip"):
        return decision

    frozen = getattr(sys, "frozen", False) if is_frozen is None else is_frozen
    if not frozen or sys.platform != "win32":
        import webbrowser

        webbrowser.open(info.html_url or RELEASES_PAGE_URL)
        return "opened_browser"

    dest_exe = Path(target_exe or sys.executable).resolve()
    versioned_exe = dest_exe.parent / f"{APP_NAME}-{info.tag_name}.exe"
    stage_dir = _resolve_staging_dir(data_dir)
    staging_exe = stage_dir / f"ClickShift-update-{info.tag_name}.exe"

    popup = InstallProgressPopup(
        title="Updating ClickShift",
        heading=f"Updating to ClickShift {info.tag_name}",
    ).start()

    try:
        popup.update("Downloading new ClickShift release...", progress=0.0, detail=info.tag_name)

        def on_progress(done: int, total: int) -> None:
            if total > 0:
                pct = min(95.0, (done / total) * 95.0)
                mb_done = done / (1024 * 1024)
                mb_tot = total / (1024 * 1024)
                popup.update(
                    f"Downloading ClickShift {info.tag_name}...",
                    progress=pct,
                    detail=f"{mb_done:.1f} MB / {mb_tot:.1f} MB",
                )
            else:
                popup.update(
                    f"Downloading ClickShift {info.tag_name}...",
                    progress=None,
                    detail=f"{done / (1024 * 1024):.1f} MB downloaded",
                )

        _download_verified_release(info, staging_exe, progress_cb=on_progress)
        popup.update("Installing update and restarting...", progress=98.0, detail="Replacing executable")
        write_update_notice(current_version, info.tag_name, silent=False, data_dir=data_dir)
        apply_windows_update(
            staging_exe,
            dest_exe,
            pid_to_wait=os.getpid(),
            new_exe_path=versioned_exe,
            relaunch_args=["--post-update"],
        )
        popup.complete(f"Restarting into {info.tag_name}...", delay_s=0.4)
        if on_before_restart is not None:
            on_before_restart()
        return "updated"
    except Exception as e:
        _log_error(f"updater: failed to apply {info.tag_name}: {e}")
        popup.error(f"Update failed: {e}", delay_s=2.5)
        return "error"
