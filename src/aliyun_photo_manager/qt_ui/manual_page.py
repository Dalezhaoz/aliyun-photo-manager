from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QLabel, QPlainTextEdit, QVBoxLayout, QWidget

from .base_page import Card


class ManualPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = Card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)
        title = QLabel("使用说明")
        title.setProperty("heroTitle", True)
        intro = QLabel("查看工具的完整使用说明。")
        intro.setProperty("heroText", True)
        intro.setWordWrap(True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        card = Card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(24, 22, 24, 24)
        content = QPlainTextEdit()
        content.setReadOnly(True)
        readme_path = Path(__file__).resolve().parents[3] / "README.md"
        try:
            content.setPlainText(readme_path.read_text(encoding="utf-8"))
        except OSError as exc:
            content.setPlainText(f"读取 README 失败：{exc}")
        card_layout.addWidget(content)
        root.addWidget(card, 1)
