from pathlib import Path
import shutil
import subprocess
import tempfile

from PIL import Image


PROJECT_ROOT = Path(__file__).resolve().parent
ASSETS_DIR = PROJECT_ROOT / "assets"
SOURCE_IMAGE = ASSETS_DIR / "app_icon.png"
MAC_ICON = ASSETS_DIR / "app_icon.icns"
WIN_ICON = ASSETS_DIR / "app_icon.ico"


def ensure_source_image() -> None:
    if not SOURCE_IMAGE.exists():
        raise FileNotFoundError(
            f"未找到图标源文件：{SOURCE_IMAGE}\n"
            "请先把你的图片保存成 assets/app_icon.png"
        )


def build_windows_icon(image: Image.Image) -> None:
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    WIN_ICON.parent.mkdir(parents=True, exist_ok=True)
    image.save(WIN_ICON, format="ICO", sizes=sizes)


def build_macos_icon(image: Image.Image) -> None:
    iconutil = shutil.which("iconutil")
    if iconutil is None:
        raise RuntimeError("当前系统缺少 iconutil，无法生成 .icns 文件。")

    with tempfile.TemporaryDirectory(prefix="iconset_") as temp_dir:
        iconset_dir = Path(temp_dir) / "app.iconset"
        iconset_dir.mkdir(parents=True, exist_ok=True)
        for size in (16, 32, 64, 128, 256, 512):
            normal = image.resize((size, size), Image.LANCZOS)
            retina = image.resize((size * 2, size * 2), Image.LANCZOS)
            normal.save(iconset_dir / f"icon_{size}x{size}.png")
            retina.save(iconset_dir / f"icon_{size}x{size}@2x.png")
        subprocess.run(
            [iconutil, "-c", "icns", str(iconset_dir), "-o", str(MAC_ICON)],
            check=True,
        )


def main() -> None:
    ensure_source_image()
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    image = Image.open(SOURCE_IMAGE).convert("RGBA")
    build_windows_icon(image)
    try:
        build_macos_icon(image)
    except Exception as exc:
        print(f"macOS 图标生成失败：{exc}")
    print(f"Windows 图标：{WIN_ICON}")
    print(f"macOS 图标：{MAC_ICON}")


if __name__ == "__main__":
    main()
