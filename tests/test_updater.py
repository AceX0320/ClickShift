import io
import os
import sys
from pathlib import Path
from unittest import mock

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from clickshift import updater  # noqa: E402


def test_parse_and_compare_versions():
    assert updater.parse_version("v1.0.11") == (1, 0, 11)
    assert updater.parse_version("1.0.9") == (1, 0, 9)
    assert updater.parse_version("bad") == ()
    assert updater.is_newer_version("v1.0.11", "1.0.10") is True
    assert updater.is_newer_version("1.1.0", "v1.0.99") is True
    assert updater.is_newer_version("v1.0.10", "1.0.10") is False
    assert updater.is_newer_version("v1.0.9", "1.0.10") is False


def test_allowed_download_url_policy():
    assert updater.is_allowed_download_url(
        "https://github.com/AceX0320/ClickShift/releases/download/v1.0.11/ClickShift.exe"
    )
    assert updater.is_allowed_download_url(
        "https://objects.githubusercontent.com/github-production-release-asset/123/ClickShift.exe"
    )
    assert not updater.is_allowed_download_url(
        "http://github.com/AceX0320/ClickShift/releases/download/v1.0.11/ClickShift.exe"
    )
    assert not updater.is_allowed_download_url("https://github.com/evil/Repo/releases/ClickShift.exe")
    assert not updater.is_allowed_download_url("https://example.com/ClickShift.exe")


def test_parse_release_payload_picks_clickshift_exe():
    payload = {
        "tag_name": "v1.0.12",
        "draft": False,
        "prerelease": False,
        "body": "Bug fixes and new icon",
        "html_url": "https://github.com/AceX0320/ClickShift/releases/tag/v1.0.12",
        "assets": [
            {
                "name": "build-info.txt",
                "size": 91,
                "browser_download_url": "https://github.com/AceX0320/ClickShift/releases/download/v1.0.12/build-info.txt",
            },
            {
                "name": "ClickShift.exe",
                "size": 150_000_000,
                "browser_download_url": "https://github.com/AceX0320/ClickShift/releases/download/v1.0.12/ClickShift.exe",
            },
        ],
    }
    info = updater.parse_release_payload(payload)
    assert info is not None
    assert info.tag_name == "v1.0.12"
    assert info.version == "1.0.12"
    assert info.download_url.endswith("/ClickShift.exe")
    assert info.asset_size == 150_000_000


def test_check_for_update_prefers_newer_tagged_release(monkeypatch):
    def fake_fetch(url, timeout=8.0):
        if url == updater.LATEST_RELEASE_API:
            return {
                "tag_name": "v1.0.10",
                "assets": [{
                    "name": "ClickShift.exe",
                    "size": 1000,
                    "browser_download_url": "https://github.com/AceX0320/ClickShift/releases/download/v1.0.10/ClickShift.exe",
                }],
            }
        if url == updater.TAGS_API:
            return [{"name": "v1.0.12"}, {"name": "v1.0.10"}]
        if "v1.0.12" in url:
            return {
                "tag_name": "v1.0.12",
                "assets": [{
                    "name": "ClickShift.exe",
                    "size": 2000,
                    "browser_download_url": "https://github.com/AceX0320/ClickShift/releases/download/v1.0.12/ClickShift.exe",
                }],
            }
        raise AssertionError(f"Unexpected URL: {url}")

    monkeypatch.setattr(updater, "_fetch_json", fake_fetch)
    rel = updater.check_for_update(current_version="1.0.10")
    assert rel is not None
    assert rel.tag_name == "v1.0.12"

    assert updater.check_for_update(current_version="1.0.12") is None


def test_download_update_verifies_pe_header_and_writes(tmp_path, monkeypatch):
    payload = b"MZ" + (b"\x00" * 2046)
    url = "https://github.com/AceX0320/ClickShift/releases/download/v1.0.11/ClickShift.exe"

    class _FakeResp(io.BytesIO):
        headers = {"Content-Length": str(len(payload))}

        def geturl(self):
            return url

    class _FakeOpener:
        def open(self, req, timeout=60.0):
            return _FakeResp(payload)

    monkeypatch.setattr(updater.urllib.request, "build_opener", lambda *a: _FakeOpener())

    seen_progress = []
    dest = tmp_path / "ClickShift-new.exe"
    out = updater.download_update(
        url, dest, progress_cb=lambda d, t: seen_progress.append((d, t)), min_bytes=512, trusted_pubkeys=()
    )
    assert out == dest
    assert dest.read_bytes() == payload
    assert seen_progress and seen_progress[-1] == (len(payload), len(payload))

    # Reject non-MZ payload
    bad_dest = tmp_path / "bad.exe"
    class _BadOpener:
        def open(self, req, timeout=60.0):
            return _FakeResp(b"<html>not an exe</html>" * 50)

    monkeypatch.setattr(updater.urllib.request, "build_opener", lambda *a: _BadOpener())
    with pytest.raises(ValueError, match="not a valid Windows executable"):
        updater.download_update(url, bad_dest, min_bytes=128, trusted_pubkeys=())
    assert not bad_dest.exists()


def test_apply_windows_update_uses_native_relaunch(tmp_path, monkeypatch):
    popen_calls = []
    monkeypatch.setattr(
        updater.subprocess,
        "Popen",
        lambda cmd, **kw: popen_calls.append((cmd, kw)),
    )
    monkeypatch.setenv("_MEIPASS2", r"C:\Temp\_MEI999")
    monkeypatch.setenv("_PYI_APPLICATION_HOME_DIR", r"C:\Temp\_MEI999")

    src = tmp_path / "staging" / "ClickShift-update-v1.0.99.exe"
    src.parent.mkdir(parents=True, exist_ok=True)
    old_dst = tmp_path / "ClickShift-v1.0.98.exe"
    new_dst = tmp_path / "ClickShift-v1.0.99.exe"
    src.write_bytes(b"MZ" + b"x" * 100)
    old_dst.write_bytes(b"MZ" + b"y" * 100)

    out = updater.apply_windows_update(
        src,
        old_dst,
        pid_to_wait=99999,
        new_exe_path=new_dst,
        relaunch_args=["--post-update"],
    )
    assert out == new_dst.resolve()
    assert new_dst.exists()
    assert len(popen_calls) == 1
    cmd, kw = popen_calls[0]
    assert Path(cmd[0]) == new_dst.resolve()
    assert "--post-update" in cmd
    assert "--cleanup-pid" in cmd and "99999" in cmd
    assert "--cleanup-old" in cmd and str(old_dst.resolve()) in cmd
    assert not any(str(a).lower().endswith("powershell.exe") for a in cmd)
    assert kw["env"].get("PYINSTALLER_RESET_ENVIRONMENT") == "1"
    assert "_MEIPASS2" not in kw["env"]
    assert "_PYI_APPLICATION_HOME_DIR" not in kw["env"]


def test_parse_release_payload_rejects_arbitrary_exe():
    payload = {
        "tag_name": "v1.0.19",
        "assets": [
            {
                "name": "unrelated.exe",
                "size": 5_000_000,
                "browser_download_url": "https://github.com/AceX0320/ClickShift/releases/download/v1.0.19/unrelated.exe",
            }
        ],
    }
    assert updater.parse_release_payload(payload) is None


def test_download_update_verifies_sha256_and_ecdsa_signature(tmp_path, monkeypatch):
    import base64
    import hashlib
    import struct
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.asymmetric import ec
    from cryptography.hazmat.primitives.asymmetric.utils import Prehashed, decode_dss_signature

    priv = ec.generate_private_key(ec.SECP256R1())
    nums = priv.public_key().public_numbers()
    cng_blob = struct.pack("<II", 0x31534345, 32) + nums.x.to_bytes(32, "big") + nums.y.to_bytes(32, "big")
    pub_b64 = base64.b64encode(cng_blob).decode("ascii")

    payload = b"MZ" + (b"\x90" * 2046)
    digest = hashlib.sha256(payload).digest()
    der_sig = priv.sign(digest, ec.ECDSA(Prehashed(hashes.SHA256())))
    r, s = decode_dss_signature(der_sig)
    sig_b64 = base64.b64encode(r.to_bytes(32, "big") + s.to_bytes(32, "big"))

    exe_url = "https://github.com/AceX0320/ClickShift/releases/download/v1.0.19/ClickShift-v1.0.19.exe"
    sums_url = "https://github.com/AceX0320/ClickShift/releases/download/v1.0.19/SHA256SUMS.txt"
    sig_url = "https://github.com/AceX0320/ClickShift/releases/download/v1.0.19/ClickShift-v1.0.19.exe.sig"
    sums_bytes = f"{digest.hex()}  ClickShift-v1.0.19.exe\n".encode("utf-8")

    routes = {exe_url: payload, sums_url: sums_bytes, sig_url: sig_b64}

    class _FakeResp(io.BytesIO):
        def __init__(self, data, target_url):
            super().__init__(data)
            self._url = target_url
            self.headers = {"Content-Length": str(len(data))}

        def geturl(self):
            return self._url

    class _FakeOpener:
        def open(self, req, timeout=60.0):
            u = req.full_url
            return _FakeResp(routes[u], u)

    monkeypatch.setattr(updater.urllib.request, "build_opener", lambda *a: _FakeOpener())

    dest = tmp_path / "ok.exe"
    out = updater.download_update(
        exe_url,
        dest,
        min_bytes=512,
        sha256_url=sums_url,
        sig_url=sig_url,
        trusted_pubkeys=("AAAA_ROTATED_OLD_KEY_INVALID_LEN", pub_b64),
        require_signature=True,
    )
    assert out == dest and dest.read_bytes() == payload

    # Bad SHA-256 fails closed
    routes[sums_url] = f"{'0' * 64}  ClickShift-v1.0.19.exe\n".encode("utf-8")
    with pytest.raises(ValueError, match="SHA-256 mismatch"):
        updater.download_update(exe_url, tmp_path / "bad_sha.exe", min_bytes=512, sha256_url=sums_url)

    # Restore SHA-256, break ECDSA sig -> fails closed
    routes[sums_url] = sums_bytes
    routes[sig_url] = base64.b64encode(b"\x01" * 64)
    with pytest.raises(ValueError, match="ECDSA signature verification"):
        updater.download_update(
            exe_url,
            tmp_path / "bad_sig.exe",
            min_bytes=512,
            sha256_url=sums_url,
            sig_url=sig_url,
            trusted_pubkeys=(pub_b64,),
            require_signature=True,
        )

    # Missing .sig when pubkey configured fails closed
    with pytest.raises(ValueError, match="Missing cryptographic signature"):
        updater.download_update(
            exe_url,
            tmp_path / "missing_sig.exe",
            min_bytes=512,
            trusted_pubkeys=(pub_b64,),
        )


def test_staging_dir_avoids_shared_temp(tmp_path, monkeypatch):
    monkeypatch.delenv("CLICKSHIFT_STAGING_DIR", raising=False)
    monkeypatch.setattr(updater.tempfile, "gettempdir", lambda: r"C:\SharedTemp")
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setattr(updater.sys, "platform", "win32")
    st = updater._resolve_staging_dir(None)
    assert st.exists()
    assert st.parent == tmp_path / "ClickShift" / "staging"
    assert Path(r"C:\SharedTemp") not in st.parents


def test_install_progress_popup_suppressed_by_env(monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_NO_POPUP", "1")
    popup = updater.InstallProgressPopup().start()
    assert popup.enabled is False
    popup.update("Step 1", progress=50, detail="Testing")
    popup.complete("Done")


def test_run_interactive_update_handles_frozen_apply(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_NO_POPUP", "1")
    info = updater.ReleaseInfo(
        tag_name="v1.0.99",
        version="1.0.99",
        download_url="https://github.com/AceX0320/ClickShift/releases/download/v1.0.99/ClickShift.exe",
    )
    monkeypatch.setattr(updater, "prompt_update_dialog", lambda *a, **k: "update")
    monkeypatch.setattr(updater, "download_update", lambda u, d, progress_cb=None: d.write_bytes(b"MZ"))
    applied = []
    monkeypatch.setattr(
        updater,
        "apply_windows_update",
        lambda s, t, pid_to_wait=None, new_exe_path=None, relaunch_args=None: applied.append((s, t, new_exe_path, relaunch_args)),
    )
    stopped = []

    monkeypatch.setattr(updater.sys, "platform", "win32")
    monkeypatch.setattr(updater.tempfile, "gettempdir", lambda: str(tmp_path))
    target = tmp_path / "ClickShift.exe"
    res = updater.run_interactive_update(
        info,
        current_version="1.0.10",
        on_before_restart=lambda: stopped.append(True),
        is_frozen=True,
        target_exe=target,
        data_dir=tmp_path,
    )
    assert res == "updated"
    assert len(applied) == 1
    assert applied[0][1] == target
    assert applied[0][2] == tmp_path / "ClickShift-v1.0.99.exe"
    assert applied[0][3] == ["--post-update"]
    assert stopped == [True]


def test_run_background_update_silent_and_records_notice(tmp_path, monkeypatch):
    """Background updates must run without prompts/popups and save update_notice.json."""
    info = updater.ReleaseInfo(
        tag_name="v1.0.18",
        version="1.0.18",
        download_url="https://github.com/AceX0320/ClickShift/releases/download/v1.0.18/ClickShift-v1.0.18.exe",
    )
    monkeypatch.setattr(
        updater,
        "prompt_update_dialog",
        lambda *a, **k: (_ for _ in ()).throw(AssertionError("prompt must not run in background update")),
    )
    monkeypatch.setattr(updater, "download_update", lambda u, d, progress_cb=None: d.write_bytes(b"MZ"))
    applied = []
    monkeypatch.setattr(
        updater,
        "apply_windows_update",
        lambda s, t, pid_to_wait=None, new_exe_path=None, relaunch_args=None: applied.append((s, t, new_exe_path, relaunch_args)),
    )
    stopped = []
    monkeypatch.setattr(updater.sys, "platform", "win32")
    monkeypatch.setattr(updater.tempfile, "gettempdir", lambda: str(tmp_path))
    target = tmp_path / "ClickShift-v1.0.17.exe"
    res = updater.run_background_update(
        info,
        current_version="1.0.17",
        on_before_restart=lambda: stopped.append(True),
        is_frozen=True,
        target_exe=target,
        data_dir=tmp_path,
    )
    assert res == "updated"
    assert stopped == [True]
    assert applied and applied[0][2] == tmp_path / "ClickShift-v1.0.18.exe"
    assert applied[0][3] == ["--post-update"]

    peeked = updater.peek_update_notice(tmp_path)
    assert peeked is not None and peeked["silent"] is True
    assert peeked["from_version"] == "v1.0.17" and peeked["to_version"] == "v1.0.18"

    consumed = updater.consume_update_notice(tmp_path)
    assert consumed is not None
    assert consumed == peeked
    assert updater.consume_update_notice(tmp_path) is None
    title, body = updater.format_update_notification(consumed)
    assert title == "ClickShift Updated"
    assert "v1.0.17" in body and "v1.0.18" in body


def test_parse_release_payload_matches_versioned_exe():
    payload = {
        "tag_name": "v1.0.18",
        "assets": [
            {
                "name": "ClickShift-v1.0.18.exe",
                "size": 151_000_000,
                "browser_download_url": "https://github.com/AceX0320/ClickShift/releases/download/v1.0.18/ClickShift-v1.0.18.exe",
            }
        ],
    }
    info = updater.parse_release_payload(payload)
    assert info is not None
    assert info.download_url.endswith("/ClickShift-v1.0.18.exe")


def test_sign_release_sums_expands_globs_and_fails_on_empty(tmp_path, monkeypatch):
    import hashlib
    import importlib.util

    script_path = Path(__file__).resolve().parent.parent / "scripts" / "sign_release.py"
    spec = importlib.util.spec_from_file_location("sign_release", script_path)
    assert spec is not None and spec.loader is not None
    sign_release = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(sign_release)

    dist_dir = tmp_path / "dist"
    dist_dir.mkdir()
    exe = dist_dir / "ClickShift-v1.0.19.exe"
    payload = b"MZ" + b"\x00" * 1024
    exe.write_bytes(payload)

    sums_out = dist_dir / "SHA256SUMS.txt"
    pattern = str(dist_dir / "ClickShift-v*.exe")
    monkeypatch.setattr(sys, "argv", ["sign_release.py", "sums", pattern, "--out", str(sums_out)])
    assert sign_release.main() == 0
    assert sums_out.is_file() and sums_out.stat().st_size > 0
    parsed = updater.parse_sha256sums(sums_out.read_text(encoding="utf-8"), exe.name)
    assert parsed == hashlib.sha256(payload).hexdigest()

    missing_out = dist_dir / "EMPTY_SUMS.txt"
    missing_pattern = str(dist_dir / "NoSuch-v*.exe")
    assert sign_release.cmd_sums([missing_pattern], missing_out) == 1
    assert not missing_out.exists()


def test_shell_extension_identity_not_treated_as_store_app(monkeypatch):
    import ctypes

    pkg_holder = {"name": "ClickShift.ShellExtension_1.0.22.0_x64__5g5shzxksyewg"}

    def fake_get_pkg_full_name(len_ptr, buf):
        val = pkg_holder["name"]
        need = len(val) + 1
        cur = len_ptr._obj.value
        len_ptr._obj.value = need
        if buf is None or cur < need:
            return 122  # ERROR_INSUFFICIENT_BUFFER
        buf.value = val
        return 0

    class FakeK32:
        GetCurrentPackageFullName = staticmethod(fake_get_pkg_full_name)

    class FakeWindll:
        kernel32 = FakeK32()

    class FakeCtypes:
        windll = FakeWindll()

        def __getattr__(self, name):
            return getattr(ctypes, name)

    monkeypatch.setattr(updater.sys, "platform", "win32")
    monkeypatch.delenv("CLICKSHIFT_ALLOW_SELF_UPDATE", raising=False)
    monkeypatch.setattr(updater, "ctypes", FakeCtypes())

    monkeypatch.setattr(updater, "_STORE_IDENTITY_CACHE", None)
    assert updater.is_packaged_store_app() is False

    pkg_holder["name"] = "ClickShift.Converter_1.0.22.0_x64__abcdef"
    monkeypatch.setattr(updater, "_STORE_IDENTITY_CACHE", None)
    assert updater.is_packaged_store_app() is True
    monkeypatch.setattr(updater, "_STORE_IDENTITY_CACHE", None)


def _make_synthetic_pyi_exe(chunks: list[tuple[str, bytes]]) -> bytes:
    """Build a minimal PE + PyInstaller CArchive binary for delta tests."""
    import struct
    import zlib

    # DOS header ("MZ") + padding up to payload
    pe_stub = b"MZ" + b"\x90" * 4094
    body = bytearray(pe_stub)

    toc_entries = bytearray()
    for name, raw_data in chunks:
        comp = zlib.compress(raw_data, level=1)
        pos = len(body)
        body.extend(comp)
        name_bytes = name.encode("utf-8") + b"\x00"
        # Entry struct: !IIIIBc + padded name (16-byte aligned)
        unpadded_len = 18 + len(name_bytes)
        entry_size = (unpadded_len + 15) & ~15
        pad = b"\x00" * (entry_size - unpadded_len)
        toc_entries.extend(
            struct.pack("!IIIIBc", entry_size, pos, len(comp), len(raw_data), 1, b"x")
            + name_bytes
            + pad
        )

    toc_pos = len(body)
    body.extend(toc_entries)
    toc_len = len(toc_entries)

    # PyInstaller 88-byte COOKIE: !8sIIII64s (pkg_len counts from start of stub here)
    cookie_magic = b"MEI\014\013\012\013\016"
    cookie_size = 88
    pkg_len = len(body) + cookie_size
    cookie = struct.pack("!8sIIII64s", cookie_magic, pkg_len, toc_pos, toc_len, 312, b"python312.dll")
    body.extend(cookie)
    return bytes(body)


def _install_delta_http_mock(
    monkeypatch,
    routes: dict[str, bytes],
    full_fetches: list[str],
    range_fetches: list[tuple[int, int]],
    force_200_on_range: bool = False,
):
    """Install a urllib.request opener mock supporting HTTP 206 Range requests."""
    import email.message
    import io
    import urllib.request

    class _Resp(io.BytesIO):
        def __init__(self, payload: bytes, status: int, url: str, extra_headers: dict[str, str] | None = None):
            super().__init__(payload)
            self.status = status
            self._url = url
            self.headers = email.message.Message()
            self.headers["Content-Length"] = str(len(payload))
            for k, v in (extra_headers or {}).items():
                self.headers[k] = v

        def getcode(self):
            return self.status

        def geturl(self):
            return self._url

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            self.close()
            return False

    class _Opener:
        def open(self, req, timeout=None):
            url = req.full_url if hasattr(req, "full_url") else str(req)
            if url not in routes:
                raise AssertionError(f"Unexpected URL requested: {url}")
            data = routes[url]
            headers = {k.lower(): v for k, v in getattr(req, "headers", {}).items()}
            range_hdr = headers.get("range", "")
            if range_hdr and not force_200_on_range:
                assert range_hdr.startswith("bytes=")
                start_s, end_s = range_hdr[len("bytes="):].split("-", 1)
                start = int(start_s)
                end = int(end_s)
                assert 0 <= start <= end < len(data)
                slice_bytes = data[start : end + 1]
                range_fetches.append((start, end))
                return _Resp(
                    slice_bytes,
                    206,
                    url,
                    {"Content-Range": f"bytes {start}-{end}/{len(data)}"},
                )
            full_fetches.append(url)
            return _Resp(data, 200, url)

    monkeypatch.setattr(urllib.request, "build_opener", lambda *a, **kw: _Opener())


def test_parse_release_payload_detects_blockmap():
    base = f"{updater.RELEASES_PAGE_URL}/download/v1.0.27"
    payload = {
        "tag_name": "v1.0.27",
        "html_url": f"{updater.RELEASES_PAGE_URL}/tag/v1.0.27",
        "body": "Delta update release",
        "assets": [
            {
                "name": "ClickShift-v1.0.27.exe",
                "browser_download_url": f"{base}/ClickShift-v1.0.27.exe",
                "size": 150_000_000,
            },
            {
                "name": "ClickShift-v1.0.27.exe.blockmap",
                "browser_download_url": f"{base}/ClickShift-v1.0.27.exe.blockmap",
                "size": 85_000,
            },
            {
                "name": "SHA256SUMS.txt",
                "browser_download_url": f"{base}/SHA256SUMS.txt",
            },
            {
                "name": "SHA256SUMS.txt.sig",
                "browser_download_url": f"{base}/SHA256SUMS.txt.sig",
            },
        ],
    }
    info = updater.parse_release_payload(payload)
    assert info is not None
    assert info.download_url == f"{base}/ClickShift-v1.0.27.exe"
    assert info.blockmap_url == f"{base}/ClickShift-v1.0.27.exe.blockmap"

    # Untrusted host on .blockmap must be dropped while keeping trusted .exe
    untrusted_payload = {
        **payload,
        "assets": [
            payload["assets"][0],
            {
                "name": "ClickShift-v1.0.27.exe.blockmap",
                "browser_download_url": "https://evil.example.com/ClickShift-v1.0.27.exe.blockmap",
            },
        ],
    }
    info2 = updater.parse_release_payload(untrusted_payload)
    assert info2 is not None
    assert info2.download_url == f"{base}/ClickShift-v1.0.27.exe"
    assert info2.blockmap_url == ""


def test_delta_download_uses_local_basis_and_range_requests(tmp_path, monkeypatch):
    import hashlib
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import ec

    def _pseudorand(tag: str, size: int) -> bytes:
        out = bytearray()
        counter = 0
        while len(out) < size:
            out.extend(hashlib.sha256(f"{tag}:{counter}".encode()).digest())
            counter += 1
        return bytes(out[:size])

    shared_chunks = [
        (f"mod_{i}.pyc", _pseudorand(f"shared-{i}", 8192))
        for i in range(12)
    ]
    old_exe = _make_synthetic_pyi_exe(
        [("header.dat", _pseudorand("hdr-old", 4096)), *shared_chunks, ("app.pyc", _pseudorand("app-old", 8192))]
    )
    # Shift offsets via header change and modify app.pyc only
    new_exe = _make_synthetic_pyi_exe(
        [("header.dat", _pseudorand("hdr-new", 4608)), *shared_chunks, ("app.pyc", _pseudorand("app-new", 8192))]
    )
    assert len(old_exe) > 4096
    assert len(new_exe) > 4096

    basis_path = tmp_path / "basis.exe"
    basis_path.write_bytes(old_exe)
    new_path_src = tmp_path / "new_build.exe"
    new_path_src.write_bytes(new_exe)

    blockmap_bytes = updater.build_blockmap(new_path_src)
    parsed = updater.parse_blockmap(blockmap_bytes)
    new_sha = hashlib.sha256(new_exe).hexdigest()
    assert parsed["size"] == len(new_exe)
    assert parsed["sha256"] == new_sha
    assert len(parsed["chunks"]) > 4

    import base64

    sums_bytes = f"{new_sha}  ClickShift-v1.0.27.exe\n".encode("utf-8")
    priv_key = ec.generate_private_key(ec.SECP256R1())
    sig_bytes = priv_key.sign(new_exe, ec.ECDSA(hashes.SHA256()))
    pub_b64 = base64.b64encode(
        priv_key.public_key().public_bytes(
            serialization.Encoding.DER,
            serialization.PublicFormat.SubjectPublicKeyInfo,
        )
    ).decode("ascii")

    base = f"{updater.RELEASES_PAGE_URL}/download/v1.0.27"
    exe_url = f"{base}/ClickShift-v1.0.27.exe"
    blockmap_url = f"{base}/ClickShift-v1.0.27.exe.blockmap"
    sha_url = f"{base}/SHA256SUMS.txt"
    sig_url = f"{base}/ClickShift-v1.0.27.exe.sig"

    full_fetches: list[str] = []
    range_fetches: list[tuple[int, int]] = []
    _install_delta_http_mock(
        monkeypatch,
        routes={
            exe_url: new_exe,
            blockmap_url: blockmap_bytes,
            sha_url: sums_bytes,
            sig_url: sig_bytes,
        },
        full_fetches=full_fetches,
        range_fetches=range_fetches,
    )

    dest = tmp_path / "downloaded.exe"
    monkeypatch.delenv("CLICKSHIFT_NO_DELTA", raising=False)
    updater.download_update(
        exe_url,
        dest,
        min_bytes=1024,
        sha256_url=sha_url,
        sig_url=sig_url,
        trusted_pubkeys=[pub_b64],
        require_signature=True,
        blockmap_url=blockmap_url,
        basis_exe=basis_path,
    )

    assert dest.read_bytes() == new_exe
    assert exe_url not in full_fetches
    downloaded_range_bytes = sum((end - start + 1) for start, end in range_fetches)
    assert downloaded_range_bytes > 0
    assert downloaded_range_bytes < len(new_exe) // 2


def test_delta_download_falls_back_cleanly_on_http_200_or_bad_blockmap(tmp_path, monkeypatch):
    import hashlib

    shared_chunks = [
        (f"mod_{i}.pyc", (f"shared-module-{i}-" * 400).encode("utf-8"))
        for i in range(6)
    ]
    old_exe = _make_synthetic_pyi_exe([*shared_chunks, ("app.pyc", b"old" * 1000)])
    new_exe = _make_synthetic_pyi_exe([*shared_chunks, ("app.pyc", b"new" * 1000)])

    basis_path = tmp_path / "basis.exe"
    basis_path.write_bytes(old_exe)
    new_src = tmp_path / "new.exe"
    new_src.write_bytes(new_exe)

    blockmap_bytes = updater.build_blockmap(new_src)
    new_sha = hashlib.sha256(new_exe).hexdigest()
    sums_bytes = f"{new_sha}  ClickShift.exe\n".encode("utf-8")

    base = f"{updater.RELEASES_PAGE_URL}/download/v1.0.27"
    exe_url = f"{base}/ClickShift.exe"
    blockmap_url = f"{base}/ClickShift.exe.blockmap"
    sha_url = f"{base}/SHA256SUMS.txt"

    # Case 1: Server ignores Range header and returns HTTP 200
    full_fetches: list[str] = []
    range_fetches: list[tuple[int, int]] = []
    _install_delta_http_mock(
        monkeypatch,
        routes={exe_url: new_exe, blockmap_url: blockmap_bytes, sha_url: sums_bytes},
        full_fetches=full_fetches,
        range_fetches=range_fetches,
        force_200_on_range=True,
    )
    dest1 = tmp_path / "out1.exe"
    updater.download_update(
        exe_url,
        dest1,
        min_bytes=1024,
        sha256_url=sha_url,
        trusted_pubkeys=(),
        blockmap_url=blockmap_url,
        basis_exe=basis_path,
    )
    assert dest1.read_bytes() == new_exe
    assert exe_url in full_fetches

    # Case 2: Corrupt blockmap triggers clean fallback to full download
    full_fetches.clear()
    range_fetches.clear()
    _install_delta_http_mock(
        monkeypatch,
        routes={exe_url: new_exe, blockmap_url: b"corrupt-blockmap-payload", sha_url: sums_bytes},
        full_fetches=full_fetches,
        range_fetches=range_fetches,
    )
    dest2 = tmp_path / "out2.exe"
    updater.download_update(
        exe_url,
        dest2,
        min_bytes=1024,
        sha256_url=sha_url,
        trusted_pubkeys=(),
        blockmap_url=blockmap_url,
        basis_exe=basis_path,
    )
    assert dest2.read_bytes() == new_exe
    assert exe_url in full_fetches

    # Case 3: Stale blockmap SHA mismatch in SHA256SUMS.txt triggers fallback
    stale_sums = (
        f"{new_sha}  ClickShift.exe\n"
        f"{'0' * 64}  ClickShift.exe.blockmap\n"
    ).encode("utf-8")
    full_fetches.clear()
    range_fetches.clear()
    _install_delta_http_mock(
        monkeypatch,
        routes={exe_url: new_exe, blockmap_url: blockmap_bytes, sha_url: stale_sums},
        full_fetches=full_fetches,
        range_fetches=range_fetches,
    )
    dest3 = tmp_path / "out3.exe"
    updater.download_update(
        exe_url,
        dest3,
        min_bytes=1024,
        sha256_url=sha_url,
        trusted_pubkeys=(),
        blockmap_url=blockmap_url,
        basis_exe=basis_path,
    )
    assert dest3.read_bytes() == new_exe
    assert exe_url in full_fetches
    assert range_fetches == []

    # Case 4: CLICKSHIFT_NO_DELTA=1 opts out of delta transfer
    full_fetches.clear()
    range_fetches.clear()
    _install_delta_http_mock(
        monkeypatch,
        routes={exe_url: new_exe, blockmap_url: blockmap_bytes, sha_url: sums_bytes},
        full_fetches=full_fetches,
        range_fetches=range_fetches,
    )
    monkeypatch.setenv("CLICKSHIFT_NO_DELTA", "1")
    dest4 = tmp_path / "out4.exe"
    updater.download_update(
        exe_url,
        dest4,
        min_bytes=1024,
        sha256_url=sha_url,
        trusted_pubkeys=(),
        blockmap_url=blockmap_url,
        basis_exe=basis_path,
    )
    assert dest4.read_bytes() == new_exe
    assert exe_url in full_fetches
    assert range_fetches == []


def test_sign_release_blockmap_subcommand(tmp_path):
    import hashlib
    import sys
    from pathlib import Path

    scripts_dir = str(Path(__file__).resolve().parent.parent / "scripts")
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import sign_release

    exe_bytes = _make_synthetic_pyi_exe([("main.pyc", b"hello-clickshift" * 500)])
    exe_path = tmp_path / "ClickShift-v1.0.27.exe"
    exe_path.write_bytes(exe_bytes)

    rc = sign_release.cmd_blockmap([str(tmp_path / "ClickShift*.exe")])
    assert rc == 0

    blockmap_path = tmp_path / "ClickShift-v1.0.27.exe.blockmap"
    assert blockmap_path.is_file()
    parsed = updater.parse_blockmap(blockmap_path.read_bytes())
    assert parsed["size"] == len(exe_bytes)
    assert parsed["sha256"] == hashlib.sha256(exe_bytes).hexdigest()
    assert len(parsed["chunks"]) >= 2

    # Unmatched glob fails closed
    assert sign_release.cmd_blockmap([str(tmp_path / "DoesNotExist*.exe")]) == 1

