from pathlib import Path

import PyInstaller.__main__


PROJECT_ROOT = Path(__file__).resolve().parent
ENTRY_FILE = PROJECT_ROOT / "windows_updater.py"
DIST_DIR = PROJECT_ROOT / "dist_updater"
BUILD_DIR = PROJECT_ROOT / "build_updater"
ICON_PATH = PROJECT_ROOT / "assets" / "app_icon.ico"


def main() -> None:
    args = [
        str(ENTRY_FILE),
        "--name=aliyun_photo_manager_updater",
        "--windowed",
        "--noconfirm",
        f"--distpath={DIST_DIR}",
        f"--workpath={BUILD_DIR}",
        f"--specpath={PROJECT_ROOT}",
        "--clean",
    ]
    if ICON_PATH.exists():
        args.append(f"--icon={ICON_PATH}")
    PyInstaller.__main__.run(args)


if __name__ == "__main__":
    main()
