from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Apply Windows incremental updates for aliyun_photo_manager.")
    parser.add_argument("--app-dir", required=True)
    parser.add_argument("--package", required=True)
    parser.add_argument("--wait-pid", type=int, required=True)
    parser.add_argument("--restart-exe", default="")
    return parser.parse_args()


def wait_for_process_exit(pid: int, timeout_seconds: int = 90) -> None:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        try:
            import os

            os.kill(pid, 0)
        except OSError:
            return
        time.sleep(0.5)
    raise RuntimeError("等待主程序退出超时。")


def load_manifest(extract_dir: Path) -> dict:
    manifest_path = extract_dir / "update_manifest.json"
    if not manifest_path.exists():
        raise RuntimeError("更新包中缺少 update_manifest.json。")
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def apply_update(app_dir: Path, package_path: Path) -> None:
    extract_dir = Path(tempfile.mkdtemp(prefix="aliyun_photo_manager_patch_"))
    backup_dir = Path(tempfile.mkdtemp(prefix="aliyun_photo_manager_backup_"))
    try:
        with zipfile.ZipFile(package_path) as archive:
            archive.extractall(extract_dir)

        manifest = load_manifest(extract_dir)
        removed = [str(item) for item in manifest.get("removed", [])]
        files = manifest.get("files", [])

        for relative_path in removed:
            target_path = app_dir / relative_path
            if target_path.exists():
                backup_path = backup_dir / relative_path
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target_path, backup_path)
                if target_path.is_file():
                    target_path.unlink()

        for item in files:
            relative_path = str(item.get("path", "")).strip()
            if not relative_path:
                continue
            source_path = extract_dir / "files" / relative_path
            target_path = app_dir / relative_path
            if not source_path.exists():
                raise RuntimeError(f"更新包缺少文件：{relative_path}")
            if target_path.exists():
                backup_path = backup_dir / relative_path
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target_path, backup_path)
            target_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, target_path)

        (app_dir / "installed_manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        restore_backup(app_dir, backup_dir)
        raise
    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)
        shutil.rmtree(backup_dir, ignore_errors=True)


def restore_backup(app_dir: Path, backup_dir: Path) -> None:
    if not backup_dir.exists():
        return
    for backup_path in sorted(backup_dir.rglob("*")):
        if backup_path.is_dir():
            continue
        relative_path = backup_path.relative_to(backup_dir)
        target_path = app_dir / relative_path
        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(backup_path, target_path)


def restart_app(app_dir: Path, restart_exe: str) -> None:
    if not restart_exe:
        return
    exe_path = app_dir / restart_exe
    if exe_path.exists():
        subprocess.Popen([str(exe_path)], cwd=str(app_dir), close_fds=True)


def main() -> int:
    args = parse_args()
    app_dir = Path(args.app_dir).resolve()
    package_path = Path(args.package).resolve()

    try:
        wait_for_process_exit(args.wait_pid)
        apply_update(app_dir, package_path)
        restart_app(app_dir, args.restart_exe)
    except Exception as exc:
        print(f"更新失败：{exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
