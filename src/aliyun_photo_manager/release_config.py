from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ReleaseConfig:
    channel: str = "stable"
    show_experimental: bool = False
    update_base_url: str = ""


def _app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def get_release_config_path() -> Path:
    return _app_root() / "release_config.json"


def load_release_config() -> ReleaseConfig:
    path = get_release_config_path()
    if not path.exists():
        return ReleaseConfig()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return ReleaseConfig()
    channel = str(payload.get("channel", "stable")).strip() or "stable"
    show_experimental = bool(payload.get("show_experimental", channel == "beta"))
    update_base_url = str(payload.get("update_base_url", "")).strip()
    return ReleaseConfig(
        channel=channel,
        show_experimental=show_experimental,
        update_base_url=update_base_url,
    )
