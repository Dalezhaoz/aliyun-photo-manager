from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
)

from ..word_to_html import WordExportResult, export_word_to_html
from .base_page import BaseToolPage
from .widgets import FormRow, PathInput, PrimaryButton, SecondaryButton, Section


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class TemplateWorker(QRunnable):
    def __init__(
        self,
        source_path: Path,
        variant: str,
        log_fn: Callable[[str], None],
        on_success: Callable[[WordExportResult], None],
        on_error: Callable[[str], None],
    ) -> None:
        super().__init__()
        self.source_path = source_path
        self.variant = variant
        self.log_fn = log_fn
        self.signals = WorkerSignals()
        self.signals.success.connect(on_success)
        self.signals.error.connect(on_error)

    def run(self) -> None:
        try:
            result = export_word_to_html(self.source_path, self.variant, logger=self.log_fn)
        except Exception as exc:
            self.signals.error.emit(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}")
        else:
            self.signals.success.emit(result)


class TemplateConvertPage(BaseToolPage):
    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="表样转换",
            description="将 Word / Excel 表样转换为 HTML，支持 Net 版和 Java 版占位符。",
            show_steps=False,
            show_log=False,
        )
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.last_result: WordExportResult | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        if self.left_card.title is not None:
            self.left_card.title.setText("转换参数")

        source_section = Section("文件选择")
        self.source_input = PathInput("请选择表样文件 (.doc/.docx/.xlsx)", "选择文件")
        self.source_edit = self.source_input.edit
        self.source_input.button.clicked.connect(self.choose_source)
        source_section.add(FormRow("表样文件", self.source_input))
        self.left_card.body_layout.addWidget(source_section)

        action_row = QHBoxLayout()
        self.net_button = PrimaryButton("Net版导出")
        self.net_button.clicked.connect(lambda: self.start_export("net"))
        action_row.addWidget(self.net_button)
        self.java_button = SecondaryButton("Java版导出")
        self.java_button.clicked.connect(lambda: self.start_export("java"))
        action_row.addWidget(self.java_button)
        self.copy_button = SecondaryButton("复制 HTML")
        self.copy_button.clicked.connect(self.copy_html)
        self.copy_button.setEnabled(False)
        action_row.addWidget(self.copy_button)
        action_row.addStretch(1)
        self.left_card.body_layout.addLayout(action_row)
        self.left_card.body_layout.addStretch(1)

        if self.right_card.title is not None:
            self.right_card.title.setText("HTML 代码")

        self.code_text = QPlainTextEdit()
        self.right_card.body_layout.addWidget(self.code_text, 1)

        self.status_label = QLabel("选择文件后点击 Net版导出 或 Java版导出")
        self.status_label.setWordWrap(True)
        self.right_card.body_layout.addWidget(self.status_label)

    def choose_source(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "选择表样文件",
            "",
            "表样文件 (*.doc *.docx *.xlsx);;所有文件 (*.*)",
        )
        if selected:
            self.source_edit.setText(selected)

    def start_export(self, variant: str) -> None:
        source_path = Path(self.source_edit.text().strip())
        if not source_path.exists():
            QMessageBox.critical(self, "参数错误", "请选择有效的表样文件。")
            return

        self.net_button.setEnabled(False)
        self.java_button.setEnabled(False)
        self.copy_button.setEnabled(False)
        self.code_text.setPlainText("正在导出 HTML...")
        self.status_label.setText("处理中...")
        self.log_fn(f"启动表样转换：{source_path.name} ({variant})")
        worker = TemplateWorker(source_path, variant, self.log_fn, self.on_success, self.on_error)
        self.thread_pool.start(worker)

    def on_success(self, result: WordExportResult) -> None:
        self.last_result = result
        self.net_button.setEnabled(True)
        self.java_button.setEnabled(True)
        self.copy_button.setEnabled(True)
        self.code_text.setPlainText(result.html_content)
        self.status_label.setText(f"源文件：{result.source_path}\n导出类型：{result.variant}")

    def on_error(self, error_text: str) -> None:
        self.net_button.setEnabled(True)
        self.java_button.setEnabled(True)
        self.copy_button.setEnabled(False)
        self.code_text.setPlainText(error_text)
        self.status_label.setText("导出失败。")
        QMessageBox.critical(self, "导出失败", error_text.splitlines()[0])

    def copy_html(self) -> None:
        if self.last_result is None:
            return
        QApplication.clipboard().setText(self.last_result.html_content)
        self.log_fn("已复制 HTML。")
