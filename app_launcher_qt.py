import os
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _setup_crash_capture():
    crash_dir = Path(os.environ.get("APPDATA", Path.home()))
    crash_log = crash_dir / "aliyun_photo_manager_qt_crash.log"

    if sys.stderr is None or getattr(sys.stderr, "fileno", lambda: -1)() < 0:
        try:
            sys.stderr = open(crash_log, "a", encoding="utf-8")
        except Exception:
            sys.stderr = open(os.devnull, "w")

    if sys.stdout is None:
        try:
            sys.stdout = open(os.devnull, "w")
        except Exception:
            pass

    return crash_log


if __name__ == "__main__":
    crash_log = _setup_crash_capture()
    try:
        from aliyun_photo_manager.qt_gui import main

        main()
    except Exception:
        import traceback

        with open(crash_log, "a", encoding="utf-8") as file_obj:
            file_obj.write(f"\n{'=' * 60}\nTop-level crash:\n{traceback.format_exc()}\n")
        raise
