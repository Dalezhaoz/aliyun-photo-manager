from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from build_windows_qt_app import main as build_qt_app_main
from build_windows_updater import main as build_updater_main

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from aliyun_photo_manager import __version__
from aliyun_photo_manager.release_config import load_release_config

QT_APP_DIR = PROJECT_ROOT / "dist_qt" / "aliyun_photo_manager_qt"
LAUNCHER_EXE = PROJECT_ROOT / "dist_launcher" / "aliyun_photo_manager_launcher.exe"
PORTABLE_DIR = PROJECT_ROOT / "dist_portable" / "aliyun_photo_manager"
APP_DIR = PORTABLE_DIR / "app"
RELEASES_DIR = PROJECT_ROOT / "releases" / "windows"
MANIFEST_NAME = "update_manifest.json"
UPDATE_CONFIG_NAME = "update_config.json"
DEPLOY_BUNDLE_NAME = "aliyun_photo_manager_qt_update_site_{version}.zip"
PORTABLE_BUNDLE_NAME = "aliyun_photo_manager_portable_{version}.zip"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Windows Qt release artifacts.")
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


def write_latest_metadata(version: str, manifest: dict, patch_manifest: dict | None, base_url: str) -> None:
    if not base_url:
        return
    latest_payload = {
        "app": "aliyun_photo_manager",
        "latest_version": version,
        "published_at": datetime.now().isoformat(timespec="seconds"),
        "notes": "",
        "full_package": {
            "url": f"{base_url.rstrip('/')}/{version}/aliyun_photo_manager_qt_full_{version}.zip",
            "sha256": sha256_file(RELEASES_DIR / version / f"aliyun_photo_manager_qt_full_{version}.zip"),
            "notes": "",
        },
        "patches": {},
    }
    if patch_manifest:
        previous = patch_manifest.get("from_version", "")
        latest_payload["patches"][previous] = {
            "url": f"{base_url.rstrip('/')}/{version}/aliyun_photo_manager_qt_patch_{previous}_to_{version}.zip",
            "sha256": sha256_file(
                RELEASES_DIR / version / f"aliyun_photo_manager_qt_patch_{previous}_to_{version}.zip"
            ),
            "notes": "",
        }
    (RELEASES_DIR / "latest.json").write_text(
        json.dumps(latest_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def write_update_config(app_dir: Path, base_url: str) -> None:
    if not base_url:
        return
    config_path = app_dir / UPDATE_CONFIG_NAME
    feed_url = f"{base_url.rstrip('/')}/latest.json"
    config_path.write_text(
        json.dumps({"feed_url": feed_url}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def detect_previous_version(current_version: str) -> str:
    latest_path = RELEASES_DIR / "latest.json"
    if not latest_path.exists():
        return ""
    try:
        payload = json.loads(latest_path.read_text(encoding="utf-8"))
    except Exception:
        return ""
    latest_version = str(payload.get("latest_version", "")).strip()
    if not latest_version or latest_version == current_version:
        return ""
    return latest_version


def build_deploy_bundle(version: str, release_dir: Path) -> Path | None:
    latest_json_path = RELEASES_DIR / "latest.json"
    if not latest_json_path.exists():
        return None

    deploy_root = release_dir / "deploy"
    if deploy_root.exists():
        shutil.rmtree(deploy_root)
    deploy_root.mkdir(parents=True, exist_ok=True)

    deploy_version_dir = deploy_root / version
    deploy_version_dir.mkdir(parents=True, exist_ok=True)

    shutil.copy2(latest_json_path, deploy_root / "latest.json")
    for item in release_dir.iterdir():
        if item.name == "deploy":
            continue
        if item.is_file():
            shutil.copy2(item, deploy_version_dir / item.name)

    bundle_path = release_dir / DEPLOY_BUNDLE_NAME.format(version=version)
    if bundle_path.exists():
        bundle_path.unlink()
    with zipfile.ZipFile(bundle_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(deploy_root / "latest.json", "latest.json")
        for file_path in sorted(deploy_version_dir.rglob("*")):
            if not file_path.is_file():
                continue
            relative_path = file_path.relative_to(deploy_root).as_posix()
            archive.write(file_path, relative_path)
    return bundle_path


def assemble_portable_package() -> None:
    if not QT_APP_DIR.exists():
        raise RuntimeError(f"Qt app directory not found: {QT_APP_DIR}")
    if not LAUNCHER_EXE.exists():
        raise RuntimeError(f"Launcher not found: {LAUNCHER_EXE}")
    if PORTABLE_DIR.exists():
        shutil.rmtree(PORTABLE_DIR)
    APP_DIR.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(QT_APP_DIR, APP_DIR)
    shutil.copy2(LAUNCHER_EXE, PORTABLE_DIR / LAUNCHER_EXE.name)


def write_portable_bundle(version: str, release_dir: Path) -> Path:
    bundle_path = release_dir / PORTABLE_BUNDLE_NAME.format(version=version)
    if bundle_path.exists():
        bundle_path.unlink()
    with zipfile.ZipFile(bundle_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in sorted(PORTABLE_DIR.rglob("*")):
            if not file_path.is_file():
                continue
            relative_path = file_path.relative_to(PORTABLE_DIR).as_posix()
            archive.write(file_path, relative_path)
    return bundle_path


def main() -> None:
    args = parse_args()
    version = __version__
    release_config = load_release_config()
    base_url = args.base_url.strip() or release_config.update_base_url.strip()

    if not args.skip_build:
        build_updater_main()
        build_qt_app_main(channel=args.channel)

    assemble_portable_package()
    write_update_config(APP_DIR, base_url)

    manifest = build_manifest(APP_DIR, version)
    manifest["channel"] = args.channel
    (APP_DIR / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    release_dir = RELEASES_DIR / version
    release_dir.mkdir(parents=True, exist_ok=True)
    (release_dir / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    full_zip = release_dir / f"aliyun_photo_manager_qt_full_{version}.zip"
    write_package(full_zip, APP_DIR, manifest)

    patch_manifest = None
    previous_version = args.previous_version.strip() or detect_previous_version(version)
    if previous_version:
        previous_manifest = load_previous_manifest(previous_version)
        if previous_manifest:
            patch_manifest = diff_manifests(previous_manifest, manifest)
            patch_zip = release_dir / f"aliyun_photo_manager_qt_patch_{previous_version}_to_{version}.zip"
            write_package(patch_zip, APP_DIR, patch_manifest)

    write_latest_metadata(version, manifest, patch_manifest, base_url)
    deploy_bundle = build_deploy_bundle(version, release_dir)
    portable_bundle = write_portable_bundle(version, release_dir)
    print(f"Qt release generated at: {release_dir}")
    if deploy_bundle is not None:
        print(f"Upload-ready bundle: {deploy_bundle}")
    print(f"Portable package: {portable_bundle}")


if __name__ == "__main__":
    main()
