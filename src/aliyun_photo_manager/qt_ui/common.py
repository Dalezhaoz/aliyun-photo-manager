from __future__ import annotations

from pathlib import Path


def format_path(value: str) -> str:
    return str(Path(value).expanduser()) if value else ""
