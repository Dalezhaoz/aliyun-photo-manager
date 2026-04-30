from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class Card(QFrame):
    def __init__(self, title: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setProperty("pageCard", True)
        self.body_layout = QVBoxLayout(self)
        self.body_layout.setContentsMargins(20, 18, 20, 20)
        self.body_layout.setSpacing(14)
        self.title: QLabel | None = None
        if title:
            self.title = QLabel(title)
            self.title.setProperty("sectionTitle", True)
            self.body_layout.addWidget(self.title)


class BaseToolPage(QWidget):
    def __init__(
        self,
        *,
        title: str,
        description: str = "",
        show_steps: bool = False,
        show_log: bool = True,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.home_callback = None
        self.favorite_callback = None
        self.root = QVBoxLayout(self)
        self.root.setContentsMargins(0, 0, 0, 0)
        self.root.setSpacing(14)

        self.header = Card()
        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(12)
        title_box = QVBoxLayout()
        title_box.setContentsMargins(0, 0, 0, 0)
        title_box.setSpacing(6)
        title_label = QLabel(title)
        title_label.setProperty("heroTitle", True)
        desc_label = QLabel(description)
        desc_label.setProperty("heroText", True)
        desc_label.setWordWrap(True)
        title_box.addWidget(title_label)
        if description:
            title_box.addWidget(desc_label)
        header_row.addLayout(title_box, 1)
        self.home_button = QPushButton("返回首页")
        self.home_button.setObjectName("HelpButton")
        self.home_button.clicked.connect(self._handle_home)
        header_row.addWidget(self.home_button)
        self.guide_button = QPushButton("使用指南")
        self.guide_button.setObjectName("HelpButton")
        header_row.addWidget(self.guide_button)
        self.history_button = QPushButton("历史记录")
        self.history_button.setObjectName("HelpButton")
        header_row.addWidget(self.history_button)
        self.favorite_button = QPushButton("收藏")
        self.favorite_button.setObjectName("HelpButton")
        self.favorite_button.setCheckable(True)
        self.favorite_button.clicked.connect(self._handle_favorite)
        header_row.addWidget(self.favorite_button)
        self.header.body_layout.addLayout(header_row)
        self.root.addWidget(self.header)

        if show_steps:
            self.steps_card = Card()
            steps = QHBoxLayout()
            steps.setContentsMargins(0, 0, 0, 0)
            steps.setSpacing(10)
            for index, text in enumerate(("选择数据", "设置规则", "执行处理", "查看结果"), start=1):
                label = QLabel(f"{index}  {text}")
                label.setObjectName("StepItem")
                label.setAlignment(Qt.AlignCenter)
                steps.addWidget(label)
            self.steps_card.body_layout.addLayout(steps)
            self.root.addWidget(self.steps_card)

        self.content_area = QHBoxLayout()
        self.content_area.setSpacing(14)
        self.root.addLayout(self.content_area, 1)
        self.left_card = Card()
        self.right_card = Card()
        self.content_area.addWidget(self.left_card, 3)
        self.content_area.addWidget(self.right_card, 2)

        self.log_card: Card | None = None
        self.log_text: QPlainTextEdit | None = None
        if show_log:
            self.log_card = Card("执行日志")
            self.log_text = QPlainTextEdit()
            self.log_text.setReadOnly(True)
            self.log_text.setMaximumHeight(120)
            self.log_text.setPlaceholderText("暂无日志")
            self.log_card.body_layout.addWidget(self.log_text)
            self.root.addWidget(self.log_card)

    def add_log(self, text: str) -> None:
        if self.log_text is None:
            return
        current = self.log_text.toPlainText().rstrip()
        self.log_text.setPlainText(f"{current}\n{text}" if current else text)

    def set_home_callback(self, callback) -> None:
        self.home_callback = callback

    def set_favorite_callback(self, callback) -> None:
        self.favorite_callback = callback

    def set_favorite_state(self, is_favorite: bool) -> None:
        self.favorite_button.setChecked(is_favorite)
        self.favorite_button.setText("已收藏" if is_favorite else "收藏")

    def _handle_home(self) -> None:
        if callable(self.home_callback):
            self.home_callback()

    def _handle_favorite(self) -> None:
        if callable(self.favorite_callback):
            self.favorite_callback()
