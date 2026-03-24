from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, Signal, QRunnable, QThreadPool
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from ..word_to_html import WordExportResult, export_word_to_html


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


class TemplateConvertPage(QWidget):
    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.last_result: WordExportResult | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._create_card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)
        title = QLabel("表样转换")
        title.setProperty("heroTitle", True)
        intro = QLabel("将 Word / Excel 表样转换为 HTML。第一版先提供代码结果和浏览器预览入口。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        controls = self._create_card()
        controls_layout = QVBoxLayout(controls)
        controls_layout.setContentsMargins(24, 22, 24, 24)
        controls_layout.setSpacing(18)
        controls_title = QLabel("转换参数")
        controls_title.setProperty("sectionTitle", True)
        controls_layout.addWidget(controls_title)
        row = QHBoxLayout()
        row.setSpacing(10)
        self.source_edit = QLineEdit()
        row.addWidget(self.source_edit, 1)
        choose_button = QPushButton("选择文件")
        choose_button.setFixedWidth(144)
        choose_button.clicked.connect(self.choose_source)
        row.addWidget(choose_button)
        controls_layout.addLayout(row)

        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        self.net_button = QPushButton("Net版导出")
        self.net_button.setProperty("accent", True)
        self.net_button.setFixedWidth(132)
        self.net_button.clicked.connect(lambda: self.start_export("net"))
        action_row.addWidget(self.net_button)
        self.java_button = QPushButton("Java版导出")
        self.java_button.setFixedWidth(132)
        self.java_button.clicked.connect(lambda: self.start_export("java"))
        action_row.addWidget(self.java_button)
        self.copy_button = QPushButton("复制 HTML")
        self.copy_button.setFixedWidth(132)
        self.copy_button.clicked.connect(self.copy_html)
        self.copy_button.setEnabled(False)
        action_row.addWidget(self.copy_button)
        action_row.addStretch(1)
        controls_layout.addLayout(action_row)
        root.addWidget(controls)

        splitter = QSplitter()
        code_card = self._create_card()
        code_layout = QVBoxLayout(code_card)
        code_layout.setContentsMargins(24, 22, 24, 24)
        code_layout.setSpacing(14)
        code_title = QLabel("HTML 代码")
        code_title.setProperty("sectionTitle", True)
        code_layout.addWidget(code_title)
        self.code_text = QPlainTextEdit()
        code_layout.addWidget(self.code_text)

        preview_card = self._create_card()
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(24, 22, 24, 24)
        preview_layout.setSpacing(14)
        preview_title = QLabel("预览说明")
        preview_title.setProperty("sectionTitle", True)
        preview_layout.addWidget(preview_title)
        self.preview_text = QPlainTextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)

        splitter.addWidget(code_card)
        splitter.addWidget(preview_card)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        root.addWidget(splitter, 1)

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

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
        self.preview_text.setPlainText("处理中...")
        self.log_fn(f"启动表样转换：{source_path.name} ({variant})")
        worker = TemplateWorker(source_path, variant, self.log_fn, self.on_success, self.on_error)
        self.thread_pool.start(worker)

    def on_success(self, result: WordExportResult) -> None:
        self.last_result = result
        self.net_button.setEnabled(True)
        self.java_button.setEnabled(True)
        self.copy_button.setEnabled(True)
        self.code_text.setPlainText(result.html_content)
        self.preview_text.setPlainText(
            f"源文件：{result.source_path}\n"
            f"导出类型：{result.variant}\n\n"
            "第一版 Qt 页面先显示 HTML 和说明，浏览器预览入口后续补。"
        )

    def on_error(self, error_text: str) -> None:
        self.net_button.setEnabled(True)
        self.java_button.setEnabled(True)
        self.copy_button.setEnabled(False)
        self.code_text.setPlainText(error_text)
        self.preview_text.setPlainText("导出失败。")
        QMessageBox.critical(self, "导出失败", error_text.splitlines()[0])

    def copy_html(self) -> None:
        if self.last_result is None:
            return
        QApplication.clipboard().setText(self.last_result.html_content)
        self.log_fn("已复制 HTML。")
