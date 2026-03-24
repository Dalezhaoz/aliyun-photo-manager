import os
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _setup_crash_capture():
    """在最早期重定向 stderr 并启用 faulthandler，捕获原生崩溃。"""
    crash_dir = Path(os.environ.get("APPDATA", Path.home()))
    crash_log = crash_dir / "aliyun_photo_manager_crash.log"

    # --windowed 打包后 sys.stderr 可能是 None，先修复它
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

    # faulthandler 能捕获 SIGSEGV / access violation 等原生崩溃
    try:
        import faulthandler
        fault_file = open(crash_log, "a", encoding="utf-8")
        faulthandler.enable(file=fault_file)
    except Exception:
        pass

    return crash_log


if __name__ == "__main__":
    crash_log = _setup_crash_capture()
    try:
        from aliyun_photo_manager.gui import main
        main()
    except Exception:
        import traceback
        with open(crash_log, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\nTop-level crash:\n{traceback.format_exc()}\n")
        raise
