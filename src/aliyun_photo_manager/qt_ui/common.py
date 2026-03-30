from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QWheelEvent
from PySide6.QtWidgets import QComboBox


def format_path(value: str) -> str:
    return str(Path(value).expanduser()) if value else ""


class AppComboBox(QComboBox):
    def wheelEvent(self, event: QWheelEvent) -> None:
        if not self.view().isVisible():
            event.ignore()
            return
        super().wheelEvent(event)
