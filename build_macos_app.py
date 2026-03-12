from pathlib import Path
import subprocess

import PyInstaller.__main__


PROJECT_ROOT = Path(__file__).resolve().parent
ENTRY_FILE = PROJECT_ROOT / "app_launcher.py"
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
APP_PATH = DIST_DIR / "阿里云照片下载与分类.app"


def postprocess_app() -> None:
    if not APP_PATH.exists():
        return

    subprocess.run(
        ["xattr", "-dr", "com.apple.quarantine", str(APP_PATH)],
        check=False,
    )
    subprocess.run(
        ["codesign", "--force", "--deep", "--sign", "-", str(APP_PATH)],
        check=False,
    )


def main() -> None:
    PyInstaller.__main__.run(
        [
            str(ENTRY_FILE),
            "--name=阿里云照片下载与分类",
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
    )
    postprocess_app()


if __name__ == "__main__":
    main()
