import os
import sys
import time
from pathlib import Path

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

