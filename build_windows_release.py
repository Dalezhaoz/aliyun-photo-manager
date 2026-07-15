from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from build_windows_app import main as build_app_main
from build_windows_updater import main as build_updater_main

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from aliyun_photo_manager import __version__
APP_DIR = PROJECT_ROOT / "dist" / "aliyun_photo_manager"
UPDATER_EXE = (
    PROJECT_ROOT
    / "dist_updater"
    / "aliyun_photo_manager_updater"
    / "aliyun_photo_manager_updater.exe"
)
RELEASES_DIR = PROJECT_ROOT / "releases" / "windows"
MANIFEST_NAME = "update_manifest.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Windows release artifacts with incremental patch support.")
    parser.add_argument("--skip-build", action="store_true")
    parser.add_argument("--base-url", default="")
    parser.add_argument("--previous-version", default="")
    parser.add_argument("--channel", choices=["stable", "beta"], default="stable")
    return parser.parse_args()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_obj:
        for chunk in iter(lambda: file_obj.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_manifest(app_dir: Path, version: str) -> dict:
    files = []
    for file_path in sorted(app_dir.rglob("*")):
        if not file_path.is_file():
            continue
        relative_path = file_path.relative_to(app_dir).as_posix()
        files.append(
            {
                "path": relative_path,
                "size": file_path.stat().st_size,
                "sha256": sha256_file(file_path),
            }
        )
    return {
        "app": "aliyun_photo_manager",
        "version": version,
        "package_type": "full",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "files": files,
        "removed": [],
    }


def load_previous_manifest(version: str) -> dict | None:
    manifest_path = RELEASES_DIR / version / MANIFEST_NAME
    if not manifest_path.exists():
        return None
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def diff_manifests(old_manifest: dict, new_manifest: dict) -> dict:
    old_files = {item["path"]: item for item in old_manifest.get("files", [])}
    new_files = {item["path"]: item for item in new_manifest.get("files", [])}
    changed = []
    removed = []
    for path, item in new_files.items():
        previous = old_files.get(path)
        if previous != item:
            changed.append(item)
    for path in old_files:
        if path not in new_files:
            removed.append(path)
    return {
        "app": "aliyun_photo_manager",
        "version": new_manifest["version"],
        "from_version": old_manifest.get("version", ""),
        "package_type": "patch",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "files": changed,
        "removed": removed,
    }


def write_package(zip_path: Path, app_dir: Path, manifest: dict) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(MANIFEST_NAME, json.dumps(manifest, ensure_ascii=False, indent=2))
        for item in manifest.get("files", []):
            relative_path = item["path"]
            archive.write(app_dir / relative_path, f"files/{relative_path}")


def write_latest_metadata(release_dir: Path, manifest: dict, patch_manifest: dict | None, base_url: str) -> None:
    if not base_url:
        return
    version = manifest["version"]
    latest_payload = {
        "app": "aliyun_photo_manager",
        "latest_version": version,
        "published_at": datetime.now().isoformat(timespec="seconds"),
        "notes": "",
        "full_package": {
            "url": f"{base_url.rstrip('/')}/{version}/aliyun_photo_manager_full_{version}.zip",
            "sha256": sha256_file(release_dir / f"aliyun_photo_manager_full_{version}.zip"),
            "notes": "",
        },
        "patches": {},
    }
    if patch_manifest:
        previous = patch_manifest.get("from_version", "")
        latest_payload["patches"][previous] = {
            "url": f"{base_url.rstrip('/')}/{version}/aliyun_photo_manager_patch_{previous}_to_{version}.zip",
            "sha256": sha256_file(
                release_dir / f"aliyun_photo_manager_patch_{previous}_to_{version}.zip"
            ),
            "notes": "",
        }
    (RELEASES_DIR / "latest.json").write_text(
        json.dumps(latest_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def ensure_updater_bundled(app_dir: Path) -> None:
    if not UPDATER_EXE.exists():
        raise RuntimeError("未找到 updater 可执行文件，请先构建 updater。")
    shutil.copy2(UPDATER_EXE, app_dir / UPDATER_EXE.name)


def main() -> None:
    args = parse_args()
    version = __version__
    if not args.skip_build:
        build_app_main(channel=args.channel)
        build_updater_main()

    if not APP_DIR.exists():
        raise RuntimeError(f"未找到 Windows 打包目录：{APP_DIR}")

    ensure_updater_bundled(APP_DIR)

    manifest = build_manifest(APP_DIR, version)
    manifest["channel"] = args.channel
    (APP_DIR / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    release_dir = RELEASES_DIR / version
    release_dir.mkdir(parents=True, exist_ok=True)
    (release_dir / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    full_zip = release_dir / f"aliyun_photo_manager_full_{version}.zip"
    write_package(full_zip, APP_DIR, manifest)

    patch_manifest = None
    previous_version = args.previous_version.strip()
    if previous_version:
        previous_manifest = load_previous_manifest(previous_version)
        if previous_manifest:
            patch_manifest = diff_manifests(previous_manifest, manifest)
            patch_zip = release_dir / f"aliyun_photo_manager_patch_{previous_version}_to_{version}.zip"
            write_package(patch_zip, APP_DIR, patch_manifest)

    write_latest_metadata(release_dir, manifest, patch_manifest, args.base_url.strip())
    print(f"已生成 Windows 发布目录：{release_dir}")


if __name__ == "__main__":
    main()
