"""PyInstaller entry point for ClickShift.exe.

Behaviour when double-clicked with NO arguments (first run):
  * installs the right-click menu entries,
  * registers itself in Startup apps (Task Manager > Startup apps),
  * starts the always-on tray daemon + watch-folder converter.
"""
import multiprocessing
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from clickshift.cli import main                      # noqa: E402
from clickshift import monitor, updater              # noqa: E402


def _strip_cleanup_args(argv: list[str]) -> tuple[list[str], int, str]:
    """Extract internal post-update cleanup flags without breaking CLI parsing."""
    rest: list[str] = []
    old_pid = 0
    old_exe = ""
    i = 0
    while i < len(argv):
        tok = argv[i]
        if tok == "--cleanup-pid" and i + 1 < len(argv):
            try:
                old_pid = int(argv[i + 1])
            except ValueError:
                old_pid = 0
            i += 2
            continue
        if tok == "--cleanup-old" and i + 1 < len(argv):
            old_exe = argv[i + 1]
            i += 2
            continue
        rest.append(tok)
        i += 1
    return rest, old_pid, old_exe


def _already_installed() -> bool:
    if not monitor.IS_WINDOWS:
        return monitor.config_file().exists()
    return monitor.config_file().exists() and monitor.is_startup_registered()


def first_run_wizard(quiet: bool = False) -> int:
    """Double-click / post-update experience: install everything then stay running."""
    notice = updater.peek_update_notice(monitor.data_dir())
    is_silent = quiet or bool(notice and notice.get("silent"))
    rc = 0
    if is_silent or not _already_installed():
        rc = main(["install", "--quiet"] if is_silent else ["install"])
        print("Starting ClickShift in the background... (it will now start automatically"
              " with Windows)")
    return monitor.run_daemon(headless=False) or rc


if __name__ == "__main__":
    multiprocessing.freeze_support()
    args, cleanup_pid, cleanup_old = _strip_cleanup_args(sys.argv[1:])
    if cleanup_old:
        updater.cleanup_replaced_exe(cleanup_pid, Path(cleanup_old))
    if not args:
        sys.exit(first_run_wizard(quiet=False))
    if len(args) == 1 and args[0] == "--post-update":
        sys.exit(first_run_wizard(quiet=True))
    sys.exit(main(args))
