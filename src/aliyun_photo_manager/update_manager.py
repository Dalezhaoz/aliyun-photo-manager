from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional
from urllib.error import URLError
from urllib.request import urlopen

from . import __version__


UPDATE_FEED_ENV = "ALIYUN_PHOTO_MANAGER_UPDATE_FEED"
UPDATE_CONFIG_NAME = "update_config.json"
UPDATER_EXE_NAME = "aliyun_photo_manager_updater.exe"
LAUNCHER_EXE_NAME = "aliyun_photo_manager_launcher.exe"


class UpdateError(RuntimeError):
    pass


@dataclass(frozen=True)
class UpdatePackage:
    version: str
    package_type: str
    url: str
    notes: str
    sha256: str


@dataclass(frozen=True)
class UpdateCheckResult:
    current_version: str
    latest_version: str
    package: UpdatePackage | None
    notes: str

    @property
    def has_update(self) -> bool:
        return self.package is not None


def _app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def get_update_feed_url() -> str:
    env_value = os.environ.get(UPDATE_FEED_ENV, "").strip()
    if env_value:
        return env_value
    config_path = _app_root() / UPDATE_CONFIG_NAME
    if not config_path.exists():
        raise UpdateError(f"未配置更新地址。请设置环境变量 {UPDATE_FEED_ENV} 或在程序目录放置 {UPDATE_CONFIG_NAME}。")
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise UpdateError(f"读取更新配置失败：{exc}") from exc
    feed_url = str(config.get("feed_url", "")).strip()
    if not feed_url:
        raise UpdateError(f"{UPDATE_CONFIG_NAME} 中缺少 feed_url。")
    return feed_url


def _fetch_json(url: str) -> dict[str, Any]:
    try:
        with urlopen(url, timeout=12) as response:
            data = response.read()
    except URLError as exc:
        raise UpdateError(f"请求更新信息失败：{exc}") from exc
    try:
        payload = json.loads(data.decode("utf-8"))
    except Exception as exc:
        raise UpdateError(f"解析更新信息失败：{exc}") from exc
    if not isinstance(payload, dict):
        raise UpdateError("更新信息格式不正确。")
    return payload


def _version_tuple(value: str) -> tuple[int, ...]:
    cleaned = value.strip().lstrip("vV")
    if not cleaned:
        return (0,)
    try:
        return tuple(int(part) for part in cleaned.split("."))
    except ValueError:
        return (0,)


def check_for_updates() -> UpdateCheckResult:
    feed_url = get_update_feed_url()
    payload = _fetch_json(feed_url)
    latest_version = str(payload.get("latest_version", "")).strip()
    if not latest_version:
        raise UpdateError("更新信息中缺少 latest_version。")
    current_version = __version__
    if _version_tuple(latest_version) <= _version_tuple(current_version):
        return UpdateCheckResult(
            current_version=current_version,
            latest_version=latest_version,
            package=None,
            notes=str(payload.get("notes", "")).strip(),
        )

    patches = payload.get("patches", {})
    package_data = None
    if isinstance(patches, dict):
        package_data = patches.get(current_version)
    package_type = "patch"
    if not isinstance(package_data, dict):
        package_data = payload.get("full_package")
        package_type = "full"
    if not isinstance(package_data, dict):
        raise UpdateError("更新信息中缺少可用的增量包或整包。")
    package_url = str(package_data.get("url", "")).strip()
    if not package_url:
        raise UpdateError("更新包配置中缺少 url。")
    package_sha256 = str(package_data.get("sha256", "")).strip().lower()
    if len(package_sha256) != 64 or any(character not in "0123456789abcdef" for character in package_sha256):
        raise UpdateError("更新包配置中缺少有效的 sha256，已拒绝不可验证的更新。")
    notes = str(package_data.get("notes", "")).strip() or str(payload.get("notes", "")).strip()
    package = UpdatePackage(
        version=latest_version,
        package_type=package_type,
        url=package_url,
        notes=notes,
        sha256=package_sha256,
    )
    return UpdateCheckResult(
        current_version=current_version,
        latest_version=latest_version,
        package=package,
        notes=notes,
    )


def download_update_package(
    package: UpdatePackage,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="aliyun_photo_manager_update_"))
    target_path = temp_dir / f"{package.package_type}_{package.version}.zip"
    try:
        digest = hashlib.sha256()
        with urlopen(package.url, timeout=120) as response:
            total = int(response.headers.get("Content-Length", 0))
            downloaded = 0
            with target_path.open("wb") as out_file:
                while True:
                    chunk = response.read(256 * 1024)
                    if not chunk:
                        break
                    out_file.write(chunk)
                    digest.update(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total > 0:
                        progress_callback(downloaded, total)
        actual_sha256 = digest.hexdigest()
        if actual_sha256 != package.sha256:
            raise UpdateError(
                f"更新包完整性校验失败：期望 {package.sha256}，实际 {actual_sha256}。"
            )
    except Exception as exc:
        if target_path.exists():
            target_path.unlink(missing_ok=True)
        if isinstance(exc, UpdateError):
            raise
        raise UpdateError(f"下载更新包失败：{exc}") from exc
    return target_path


def launch_windows_updater(package_path: Path) -> None:
    if sys.platform != "win32":
        raise UpdateError("在线更新目前只支持 Windows。")
    app_dir = _app_root()
    launcher_path = app_dir.parent / LAUNCHER_EXE_NAME
    if not launcher_path.exists():
        launcher_path = app_dir / UPDATER_EXE_NAME
    if not launcher_path.exists():
        raise UpdateError(f"未找到启动器：{LAUNCHER_EXE_NAME}")
    restart_exe = Path(sys.executable).name if getattr(sys, "frozen", False) else ""
    command = [
        str(launcher_path),
        "--apply-update",
        "--app-dir",
        str(app_dir),
        "--package",
        str(package_path),
        "--wait-pid",
        str(os.getpid()),
    ]
    if restart_exe:
        command.extend(["--restart-exe", restart_exe])
    try:
        subprocess.Popen(command, cwd=str(app_dir), close_fds=True)
    except Exception as exc:
        raise UpdateError(f"启动更新器失败：{exc}") from exc
