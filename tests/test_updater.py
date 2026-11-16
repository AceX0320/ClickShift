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
        url, dest, progress_cb=lambda d, t: seen_progress.append((d, t)), min_bytes=512
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
        updater.download_update(url, bad_dest, min_bytes=128)
    assert not bad_dest.exists()


def test_apply_windows_update_writes_self_deleting_script(tmp_path, monkeypatch):
    monkeypatch.setattr(updater.tempfile, "gettempdir", lambda: str(tmp_path))
    popen_calls = []
    monkeypatch.setattr(updater.subprocess, "Popen", lambda cmd, **kw: popen_calls.append(cmd))

    src = tmp_path / "new.exe"
    dst = tmp_path / "ClickShift.exe"
    src.write_bytes(b"MZ" + b"x" * 100)
    dst.write_bytes(b"MZ" + b"y" * 100)

    script = updater.apply_windows_update(src, dst, pid_to_wait=99999, relaunch_args=["--post-update"])
    assert script.exists()
    content = script.read_text(encoding="utf-8-sig")
    assert script.suffix == ".ps1"
    assert str(src.resolve()) in content
    assert str(dst.resolve()) in content
    assert "99999" in content
    assert "--post-update" in content
    assert "PYINSTALLER_RESET_ENVIRONMENT" in content
    assert len(popen_calls) == 1
    assert popen_calls[0][0].lower().endswith("powershell.exe")
    assert popen_calls[0][-2:] == ["-File", str(script)]


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
