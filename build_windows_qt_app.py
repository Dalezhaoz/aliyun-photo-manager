import json
import argparse
from pathlib import Path

import PyInstaller.__main__


PROJECT_ROOT = Path(__file__).resolve().parent
ENTRY_FILE = PROJECT_ROOT / "app_launcher_qt.py"
DIST_DIR = PROJECT_ROOT / "dist_qt"
BUILD_DIR = PROJECT_ROOT / "build_qt"
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
UPDATE_CONFIG_PATH = PROJECT_ROOT / "update_config.json"


def write_release_config(channel: str) -> Path:
    config_path = PROJECT_ROOT / "release_config.json"
    existing: dict = {}
    if config_path.exists():
        try:
            existing = json.loads(config_path.read_text(encoding="utf-8"))
        except Exception:
            existing = {}
    payload = dict(existing)
    explicit_show_experimental = existing.get("show_experimental")
    payload.update(
        {
            "channel": channel,
            "show_experimental": (
                bool(explicit_show_experimental)
                if explicit_show_experimental is not None
                else channel == "beta"
            ),
        }
    )
    config_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return config_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Windows Qt desktop app.")
    parser.add_argument("--channel", choices=["stable", "beta"], default="stable")
    return parser.parse_args()


def main(channel: str = "stable") -> None:
    release_config = write_release_config(channel)
    args = [
        str(ENTRY_FILE),
        "--name=aliyun_photo_manager_qt",
        "--windowed",
        "--noconfirm",
        f"--distpath={DIST_DIR}",
        f"--workpath={BUILD_DIR}",
        f"--specpath={PROJECT_ROOT}",
        f"--paths={PROJECT_ROOT / 'src'}",
        "--hidden-import=aliyun_photo_manager",
        "--hidden-import=aliyun_photo_manager.qt_gui",
        "--hidden-import=aliyun_photo_manager.release_config",
        "--hidden-import=PySide6.QtCore",
        "--hidden-import=PySide6.QtGui",
        "--hidden-import=PySide6.QtWidgets",
        "--hidden-import=PySide6.QtPrintSupport",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=openpyxl",
        "--hidden-import=oss2",
        "--exclude-module=tkinter",
        "--exclude-module=_tkinter",
        "--exclude-module=tkinter.ttk",
        "--exclude-module=tkinter.filedialog",
        "--exclude-module=tkinter.messagebox",
        "--exclude-module=PySide6.QtQml",
        "--exclude-module=PySide6.QtQuick",
        "--exclude-module=PySide6.QtPdf",
        "--exclude-module=PySide6.QtPdfWidgets",
        "--clean",
        f"--add-data={release_config};.",
    ]
    if ICON_PATH.exists():
        args.append(f"--icon={ICON_PATH}")
    if UPDATE_CONFIG_PATH.exists():
        args.append(f"--add-data={UPDATE_CONFIG_PATH};.")
    if HELPER_PUBLISH_DIR.exists():
        for helper_file in HELPER_PUBLISH_DIR.iterdir():
            if helper_file.is_file():
                args.append(f"--add-binary={helper_file};.")
    PyInstaller.__main__.run(args)


if __name__ == "__main__":
    args = parse_args()
    main(channel=args.channel)
