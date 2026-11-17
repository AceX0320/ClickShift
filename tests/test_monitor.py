import os
import sys
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
_tmp = Path(__file__).parent / "_tmp_home"
os.environ["CLICKSHIFT_HOME"] = str(_tmp)
os.environ["CLICKSHIFT_DROP"] = str(_tmp / "drop")

from clickshift import monitor  # noqa: E402
from clickshift.monitor import WatchFolder  # noqa: E402


def test_config_roundtrip():
    cfg = monitor.load_config()
    cfg["auto_target"] = "docx"
    monitor.save_config(cfg)
    assert monitor.load_config()["auto_target"] == "docx"


def test_watch_folder_converts(tmp_path):
    drop = tmp_path / "drop"
    drop.mkdir()
    (drop / "hello.txt").write_text("line one\nline two\n", encoding="utf-8")
    done = []
    w = WatchFolder(drop, {"auto_target": "docx", "ml_enabled": False},
                    on_done=done.append)
    w.start()
    deadline = time.time() + 15
    while time.time() < deadline and not done:
        time.sleep(0.3)
    w.stop_event.set()
    assert done, "watcher did not convert the dropped file"
    assert done[0].ok, done[0].error
    assert (drop / "hello.docx").exists()


def test_nonwindows_no_registry_calls():
    if not monitor.IS_WINDOWS:
        assert monitor.register_startup() is False
        assert monitor.install_context_menu() == 0


def test_all_choices_available_in_menus():
    """Every file family must expose all supported conversion targets, not just 2-3."""
    pdf_targets = monitor.menu_targets("pdf")
    for expected in ("docx", "txt", "md", "html", "csv", "png", "jpg", "webp"):
        assert expected in pdf_targets
    docx_targets = monitor.menu_targets("docx")
    for expected in ("pdf", "txt", "md", "html", "rtf", "odt"):
        assert expected in docx_targets
    png_targets = monitor.menu_targets("png")
    for expected in ("webp", "jpg", "pdf", "avif", "jxl", "tiff", "bmp", "ico", "gif"):
        assert expected in png_targets


def test_daemon_ipc_delegation(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    src = tmp_path / "ipc_note.txt"
    src.write_text("Hello via warm daemon IPC\n", encoding="utf-8")
    srv = monitor.DaemonIPCServer({"ml_enabled": False})
    srv.start_and_publish()
    try:
        # Pretend the IPC file belongs to a different live PID so client connects
        meta = __import__("json").loads(monitor.ipc_file().read_text(encoding="utf-8"))
        meta["pid"] = 0
        monitor.ipc_file().write_text(__import__("json").dumps(meta), encoding="utf-8")
        seen = []
        res = monitor.try_daemon_convert([(str(src), "html")], enable_ml=False, progress=seen.append)
        assert res is not None and len(res) == 1
        assert res[0].ok and len(seen) == 1
        assert (tmp_path / "ipc_note.html").exists()
    finally:
        srv.stop()


def test_win11_shell_ext_skips_duplicate_classic_cascade(monkeypatch):
    """When Win11 IExplorerCommand shell ext installs, remove registry ClickShift.Menu
    instead of writing a second cascading entry in the 'Show more options' menu."""
    deleted, written = [], []
    monkeypatch.setattr(monitor, "IS_WINDOWS", True)
    monkeypatch.setattr(monitor, "winreg", None)
    monkeypatch.setattr(monitor, "_save_exe_path_registry", lambda: None)
    monkeypatch.setattr(monitor, "_purge_duplicate_class_verbs", lambda: None)
    monkeypatch.setattr(monitor, "_set_win11_both_menus", lambda _v: None)
    monkeypatch.setattr(monitor, "_register_openwith", lambda _ext: None)
    monkeypatch.setattr(monitor, "_verb_locations", lambda ext: [rf"Software\Classes\SystemFileAssociations\.{ext}"])
    monkeypatch.setattr(monitor, "_delete_verb", lambda loc, verb: deleted.append((loc, verb)))
    monkeypatch.setattr(monitor, "_write_cascading_menu", lambda loc, *_a: written.append(loc))

    monkeypatch.setattr(monitor, "_is_win11", lambda: True)
    monkeypatch.setattr(monitor, "_install_win11_shell_extension", lambda: True)
    assert monitor.install_context_menu(extensions=["pdf"]) == 1
    assert any(verb == monitor.VERB_MENU for _loc, verb in deleted)
    assert written == []

    # Fallback (e.g. Win10) still writes the registry cascade
    deleted.clear()
    monkeypatch.setattr(monitor, "_is_win11", lambda: False)
    assert monitor.install_context_menu(extensions=["pdf"]) == 1
    assert written == [r"Software\Classes\SystemFileAssociations\.pdf"]


def test_tray_icon_loads_new_asset_and_cli_install_quiet(tmp_path, monkeypatch):
    img = monitor._make_tray_icon_image()
    assert img.mode == "RGBA"
    assert img.size == (64, 64)

    monkeypatch.setenv("CLICKSHIFT_NO_POPUP", "1")
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    monkeypatch.setattr(monitor, "DATA_DIR", tmp_path)
    monkeypatch.setitem(monitor.DEFAULT_CONFIG, "watch_folder", str(tmp_path / "drop"))
    monkeypatch.setattr(monitor, "install_context_menu", lambda **_kw: 3)
    monkeypatch.setattr(monitor, "register_startup", lambda: True)
    monkeypatch.setattr(monitor, "refresh_explorer", lambda: None)

    from clickshift import cli
    rc = cli.main(["install", "--quiet"])
    assert rc == 0


def test_auto_update_12h_interval_and_background_check(tmp_path, monkeypatch):
    assert monitor.DEFAULT_CONFIG["update_check_interval_s"] == 43200
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    monkeypatch.setattr(monitor, "DATA_DIR", tmp_path)

    due, wait_s = monitor.update_check_due_in({"update_check_interval_s": 43200, "last_update_check_ts": 0.0}, now=1000.0)
    assert due is True and wait_s == 0.0

    due, wait_s = monitor.update_check_due_in({"update_check_interval_s": 43200, "last_update_check_ts": 1000.0}, now=4600.0)
    assert due is False and wait_s == 43200 - 3600

    due, wait_s = monitor.update_check_due_in({"update_check_interval_s": 43200, "last_update_check_ts": 1000.0}, now=1000.0 + 43200.0)
    assert due is True and wait_s == 0.0

    from clickshift import updater
    fake_info = updater.ReleaseInfo(
        tag_name="v1.0.99",
        version="1.0.99",
        download_url="https://github.com/AceX0320/ClickShift/releases/download/v1.0.99/ClickShift-v1.0.99.exe",
    )
    bg_calls = []

    class _FakeUpdater:
        @staticmethod
        def check_for_update(current_version=""):
            return fake_info

        @staticmethod
        def run_background_update(info, current_version="", on_before_restart=None):
            bg_calls.append((info.tag_name, current_version))
            return "updated"

    res = monitor.perform_background_update_check(
        current_version="1.0.18",
        updater_mod=_FakeUpdater,
        now=50_000.0,
    )
    assert res == "updated"
    assert bg_calls == [("v1.0.99", "1.0.18")]
    assert monitor.load_config()["last_update_check_ts"] == 50_000.0


def test_win11_uac_decline_shows_admin_popup(tmp_path, monkeypatch):
    """If the user selects 'No' at the UAC prompt, show the exact admin rights message."""
    import base64

    assets = tmp_path / "assets"
    assets.mkdir()
    (assets / "ClickShiftShell.dll").write_bytes(b"MZ")
    (assets / "AppxManifest.xml").write_text("<Package/>", encoding="utf-8")
    (assets / "ClickShiftShell.msix").write_bytes(b"PK")
    (assets / "ClickShiftShell.cer").write_bytes(b"FAKECER")

    monkeypatch.setattr(monitor, "IS_WINDOWS", True)
    monkeypatch.setattr(monitor, "_assets_dir", lambda: assets)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))

    runs = []

    class _FakeProc:
        def __init__(self, rc):
            self.returncode = rc
            self.stderr = ""
            self.stdout = ""

    def fake_run(cmd, **_kw):
        runs.append(cmd)
        # First call = direct Add-AppxPackage fails; second call = UAC declined (1223)
        return _FakeProc(1 if len(runs) == 1 else 1223)

    popups = []
    monkeypatch.setattr(monitor.subprocess, "run", fake_run)
    monkeypatch.setattr(monitor, "_show_admin_required_popup", lambda: popups.append("ClickShift needs the admin rights to work!"))

    ok = monitor._install_win11_shell_extension()
    assert ok is False
    assert popups == ["ClickShift needs the admin rights to work!"]
    assert "-EncodedCommand" in runs[1]
    uac_script = base64.b64decode(runs[1][runs[1].index("-EncodedCommand") + 1]).decode("utf-16le")
    assert "Start-Process powershell -Verb RunAs" in uac_script
    # MSIX mode must not unlock AllowDevelopmentWithoutDevLicense
    assert "AllowDevelopmentWithoutDevLicense" not in uac_script


def test_daemon_ipc_rejects_bad_token_and_oversized_line(tmp_path, monkeypatch):
    import json
    import socket

    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    srv = monitor.DaemonIPCServer({"ml_enabled": False})
    srv.start_and_publish()
    try:
        # 1. Bad token returns empty response cleanly
        with socket.create_connection(("127.0.0.1", srv.port), timeout=2.0) as s:
            s.sendall((json.dumps({"token": "wrong", "jobs": []}) + "\n").encode("utf-8"))
            assert s.recv(1024) == b""

        # 2. Oversized header without newline fails cleanly without OOM/hang
        with socket.create_connection(("127.0.0.1", srv.port), timeout=2.0) as s:
            try:
                s.sendall(b"x" * (srv.MAX_REQUEST_BYTES + 512))
                assert s.recv(1024) == b""
            except (ConnectionResetError, BrokenPipeError):
                pass  # Windows Winsock sends RST when server closes socket with unread bytes
    finally:
        srv.stop()


def test_validate_watch_folder_rejects_protected_dirs(tmp_path):
    with pytest.raises(ValueError):
        monitor.validate_watch_folder(Path.home())
    with pytest.raises(ValueError):
        monitor.validate_watch_folder(Path.home() / "Desktop")
    assert monitor.validate_watch_folder(tmp_path / "drop") == (tmp_path / "drop").resolve()


def test_store_msix_skips_self_signed_win11_uac_install(monkeypatch):
    from clickshift import updater
    monkeypatch.setattr(monitor, "IS_WINDOWS", True)
    monkeypatch.setattr(monitor, "winreg", None)
    monkeypatch.setattr(monitor, "_save_exe_path_registry", lambda: None)
    monkeypatch.setattr(monitor, "_purge_duplicate_class_verbs", lambda: None)
    monkeypatch.setattr(monitor, "_set_win11_both_menus", lambda _v: None)
    monkeypatch.setattr(monitor, "_register_openwith", lambda _ext: None)
    monkeypatch.setattr(monitor, "_verb_locations", lambda ext: [rf"Software\Classes\SystemFileAssociations\.{ext}"])
    monkeypatch.setattr(monitor, "_delete_verb", lambda *_a: None)
    monkeypatch.setattr(monitor, "_is_win11", lambda: True)
    monkeypatch.setattr(updater, "is_packaged_store_app", lambda: True)
    monkeypatch.setattr(
        monitor,
        "_install_win11_shell_extension",
        lambda: (_ for _ in ()).throw(AssertionError("Store MSIX must never invoke UAC self-signed cert install")),
    )
    assert monitor.install_context_menu(extensions=["pdf"]) == 1


def test_load_config_stat_cache_and_cli_tail_lines(tmp_path, monkeypatch):
    import json
    from clickshift import cli, monitor

    cfg_path = tmp_path / "config.json"
    monkeypatch.setattr(monitor, "DATA_DIR", tmp_path)
    monkeypatch.setattr(monitor, "CONFIG_FILE", cfg_path)
    monitor._CONFIG_CACHE = None

    cfg_path.write_text(json.dumps({"auto_target": "pdf"}), encoding="utf-8")
    read_calls = 0
    orig_read = json.loads

    def wrapped_loads(s, *a, **kw):
        nonlocal read_calls
        read_calls += 1
        return orig_read(s, *a, **kw)

    monkeypatch.setattr(json, "loads", wrapped_loads)
    assert monitor.load_config()["auto_target"] == "pdf"
    assert monitor.load_config()["auto_target"] == "pdf"
    assert read_calls == 1

    # Verify WatchFolder converts once without re-converting unchanged source/output
    watch_dir = tmp_path / "drop"
    watch_dir.mkdir()
    doc = watch_dir / "note.txt"
    doc.write_text("hello watch folder", encoding="utf-8")

    converts: list[str] = []

    class FakeEngine:
        def convert(self, src_path, dst_ext, out_path=None):
            converts.append(str(src_path))
            out_p = Path(out_path)
            out_p.write_text("converted", encoding="utf-8")
            from clickshift.formats import ConversionResult
            return ConversionResult(str(src_path), str(out_p), "text-passthru", 1.0, True)

    import clickshift.engine as eng_mod
    monkeypatch.setattr(eng_mod, "get_engine", lambda **_kw: FakeEngine())
    monkeypatch.setattr(monitor.WatchFolder, "POLL_SECONDS", 0.03)
    wf = monitor.WatchFolder(watch_dir, {"auto_target": "html"})
    wf.start()
    try:
        import time
        deadline = time.monotonic() + 2.0
        while not converts and time.monotonic() < deadline:
            time.sleep(0.03)
        time.sleep(0.18)
    finally:
        wf.stop_event.set()
        wf.join(timeout=1.5)
    assert len(converts) == 1

    # Verify _tail_lines across small, empty, and multi-chunk (>64 KiB) files
    empty = tmp_path / "empty.jsonl"
    empty.write_text("", encoding="utf-8")
    assert cli._tail_lines(empty, 10) == []

    big = tmp_path / "big.jsonl"
    lines = [f'{{"i":{i},"pad":"' + ("x" * 60) + '"}' for i in range(1500)]
    big.write_bytes(("\r\n".join(lines) + "\r\n").encode("utf-8"))
    assert cli._tail_lines(big, 25) == lines[-25:]




