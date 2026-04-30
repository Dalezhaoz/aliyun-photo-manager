from __future__ import annotations

from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class FormRow(QWidget):
    def __init__(self, label: str, widget: QWidget, button_text: str | None = None) -> None:
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        label_widget = QLabel(label)
        label_widget.setProperty("formLabel", True)
        label_widget.setFixedWidth(110)
        layout.addWidget(label_widget)
        layout.addWidget(widget, 1)
        self.button: QPushButton | None = None
        if button_text:
            self.button = SecondaryButton(button_text)
            self.button.setFixedWidth(110)
            layout.addWidget(self.button)


class PathInput(QWidget):
    def __init__(self, placeholder: str = "", button_text: str = "选择文件") -> None:
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText(placeholder)
        self.button = SecondaryButton(button_text)
        self.button.setFixedWidth(110)
        layout.addWidget(self.edit, 1)
        layout.addWidget(self.button)


class Section(QWidget):
    def __init__(self, title: str) -> None:
        super().__init__()
        self.body_layout = QVBoxLayout(self)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(10)
        title_label = QLabel(title)
        title_label.setObjectName("SectionTitle")
        self.body_layout.addWidget(title_label)

    def add(self, widget: QWidget) -> None:
        self.body_layout.addWidget(widget)


class PrimaryButton(QPushButton):
    def __init__(self, text: str) -> None:
        super().__init__(text)
        self.setProperty("accent", True)
        self.setMinimumHeight(40)


class SecondaryButton(QPushButton):
    def __init__(self, text: str) -> None:
        super().__init__(text)
        self.setMinimumHeight(40)
