from pathlib import Path

import PyInstaller.__main__


PROJECT_ROOT = Path(__file__).resolve().parent
ENTRY_FILE = PROJECT_ROOT / "app_launcher.py"
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
ICON_PATH = PROJECT_ROOT / "assets" / "app_icon.ico"
HELPER_PUBLISH_DIR = (
    PROJECT_ROOT
    / "tools"
    / "PhoneDecryptHelper"
    / "bin"
    / "Release"
    / "net8.0-windows"
    / "win-x86"
    / "publish"
)


def main() -> None:
    args = [
        str(ENTRY_FILE),
        "--name=aliyun_photo_manager",
        "--windowed",
        "--noconfirm",
        f"--distpath={DIST_DIR}",
        f"--workpath={BUILD_DIR}",
        f"--specpath={PROJECT_ROOT}",
        f"--paths={PROJECT_ROOT / 'src'}",
        "--hidden-import=aliyun_photo_manager",
        "--hidden-import=aliyun_photo_manager.gui",
        "--hidden-import=aliyun_photo_manager.app",
        "--hidden-import=aliyun_photo_manager.config",
        "--hidden-import=aliyun_photo_manager.downloader",
        "--hidden-import=aliyun_photo_manager.excel_classifier",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=openpyxl",
        "--hidden-import=oss2",
        "--clean",
    ]
    if ICON_PATH.exists():
        args.append(f"--icon={ICON_PATH}")
    if HELPER_PUBLISH_DIR.exists():
        for helper_file in HELPER_PUBLISH_DIR.iterdir():
            if helper_file.is_file():
                args.append(f"--add-binary={helper_file};.")
    PyInstaller.__main__.run(args)


if __name__ == "__main__":
    main()
