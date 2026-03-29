import json
import argparse
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
UPDATER_EXE = (
    PROJECT_ROOT
    / "dist_updater"
    / "aliyun_photo_manager_updater"
    / "aliyun_photo_manager_updater.exe"
)


def write_release_config(channel: str) -> Path:
    config_path = PROJECT_ROOT / "release_config.json"
    config_path.write_text(
        json.dumps(
            {
                "channel": channel,
                "show_experimental": channel == "beta",
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return config_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Windows desktop app.")
    parser.add_argument("--channel", choices=["stable", "beta"], default="stable")
    return parser.parse_args()


def main(channel: str = "stable") -> None:
    release_config = write_release_config(channel)
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
        "--hidden-import=aliyun_photo_manager.release_config",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=openpyxl",
        "--hidden-import=oss2",
        "--clean",
        f"--add-data={release_config};.",
    ]
    if ICON_PATH.exists():
        args.append(f"--icon={ICON_PATH}")
    if HELPER_PUBLISH_DIR.exists():
        for helper_file in HELPER_PUBLISH_DIR.iterdir():
            if helper_file.is_file():
                args.append(f"--add-binary={helper_file};.")
    if UPDATER_EXE.exists():
        args.append(f"--add-binary={UPDATER_EXE};.")
    PyInstaller.__main__.run(args)


if __name__ == "__main__":
    args = parse_args()
    main(channel=args.channel)
