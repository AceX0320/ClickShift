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


def first_run_wizard(quiet: bool = False) -> int:
    """Double-click / post-update experience: install everything then stay running."""
    notice = updater.peek_update_notice(monitor.data_dir())
    is_silent = quiet or bool(notice and notice.get("silent"))
    rc = main(["install", "--quiet"] if is_silent else ["install"])
    print("Starting ClickShift in the background... (it will now start automatically"
          " with Windows)")
    return monitor.run_daemon(headless=False) or rc


if __name__ == "__main__":
    multiprocessing.freeze_support()
    if len(sys.argv) == 1:
        sys.exit(first_run_wizard(quiet=False))
    if len(sys.argv) == 2 and sys.argv[1] == "--post-update":
        sys.exit(first_run_wizard(quiet=True))
    sys.exit(main())
