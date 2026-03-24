from pathlib import Path
import subprocess
import sys


PROJECT_ROOT = Path(__file__).resolve().parent
HELPER_PROJECT = PROJECT_ROOT / "tools" / "PhoneDecryptHelper" / "PhoneDecryptHelper.csproj"
OUTPUT_DIR = PROJECT_ROOT / "tools" / "PhoneDecryptHelper" / "bin" / "Release" / "net8.0-windows" / "win-x86" / "publish"


def main() -> None:
    if sys.platform != "win32":
        raise SystemExit("电话解密 helper 只能在 Windows 上构建。")
    if not (PROJECT_ROOT / "Interop.DeDll.dll").exists() or not (PROJECT_ROOT / "DeDLL.dll").exists():
        raise SystemExit("请先把 Interop.DeDll.dll 和 DeDLL.dll 放到项目根目录。")

    command = [
        "dotnet",
        "publish",
        str(HELPER_PROJECT),
        "-c",
        "Release",
        "-r",
        "win-x86",
        "--self-contained",
        "true",
    ]
    completed = subprocess.run(command, check=False)
    if completed.returncode != 0:
        raise SystemExit(completed.returncode)
    print(f"Helper 已构建：{OUTPUT_DIR / 'PhoneDecryptHelper.exe'}")
