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
