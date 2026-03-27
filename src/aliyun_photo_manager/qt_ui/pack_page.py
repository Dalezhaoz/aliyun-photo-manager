from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..result_packer import PackSummary, pack_encrypted_folder, query_pack_history


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class SimpleWorker(QRunnable):
    def __init__(self, task: Callable[[], object], on_success, on_error) -> None:
        super().__init__()
        self.task = task
        self.signals = WorkerSignals()
        self.signals.success.connect(on_success)
        self.signals.error.connect(on_error)

    def run(self) -> None:
        try:
            result = self.task()
        except Exception as exc:
            self.signals.error.emit(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}")
        else:
            self.signals.success.emit(result)


class PackPage(QWidget):
    LABEL_WIDTH = 132
    ACTION_WIDTH = 144

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.last_summary: PackSummary | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        title = QLabel("结果打包")
        title.setProperty("heroTitle", True)
        intro = QLabel("对任意结果文件或文件夹进行压缩与 AES 加密，并支持历史密码查询。")
        intro.setProperty("heroText", True)
        intro.setWordWrap(True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        body = QGridLayout()
        body.setHorizontalSpacing(16)
        body.setVerticalSpacing(16)
        root.addLayout(body, 1)

        left = self._card()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(24, 22, 24, 24)
        left_layout.setSpacing(18)
        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.source_edit = QLineEdit()
        source_row = QWidget()
        source_layout = QHBoxLayout(source_row)
        source_layout.setContentsMargins(0, 0, 0, 0)
        source_layout.setSpacing(10)
        source_layout.addWidget(self.source_edit, 1)
        file_button = QPushButton("选择文件")
        file_button.setFixedWidth(self.ACTION_WIDTH)
        file_button.clicked.connect(self.choose_source_file)
        source_layout.addWidget(file_button)
        dir_button = QPushButton("选择文件夹")
        dir_button.setFixedWidth(self.ACTION_WIDTH)
        dir_button.clicked.connect(self.choose_source_dir)
        source_layout.addWidget(dir_button)
        self._add_row(form, 0, "待打包对象", source_row)

        self.output_edit = QLineEdit()
        self._add_row(form, 1, "输出目录", self._with_dir_button(self.output_edit))

        self.custom_password_checkbox = QCheckBox("手动设置密码")
        self.custom_password_checkbox.toggled.connect(self.update_password_state)
        form.addWidget(self.custom_password_checkbox, 2, 1)

        self.password_edit = QLineEdit()
        self.password_edit.setReadOnly(True)
        self._add_row(form, 3, "打包密码", self.password_edit)
        left_layout.addLayout(form)

        actions = QHBoxLayout()
        self.run_button = QPushButton("一键打包并加密")
        self.run_button.setProperty("accent", True)
        self.run_button.clicked.connect(self.start_pack)
        actions.addWidget(self.run_button)
        self.copy_button = QPushButton("复制密码")
        self.copy_button.clicked.connect(self.copy_password)
        self.copy_button.setEnabled(False)
        actions.addWidget(self.copy_button)
        actions.addStretch(1)
        left_layout.addLayout(actions)
        body.addWidget(left, 0, 0)

        right = self._card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_title = QLabel("打包结果与历史")
        right_title.setProperty("sectionTitle", True)
        right_layout.addWidget(right_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        right_layout.addWidget(self.result_text, 1)

        query_row = QHBoxLayout()
        self.query_edit = QLineEdit()
        self.query_edit.setPlaceholderText("按文件名、来源名或密码查询")
        query_row.addWidget(self.query_edit, 1)
        query_button = QPushButton("查询历史")
        query_button.setFixedWidth(self.ACTION_WIDTH)
        query_button.clicked.connect(self.run_query)
        query_row.addWidget(query_button)
        right_layout.addLayout(query_row)
        body.addWidget(right, 0, 1)

        body.setColumnStretch(0, 3)
        body.setColumnStretch(1, 2)

    def _card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> None:
        label = QLabel(label_text)
        label.setProperty("formLabel", True)
        label.setFixedWidth(self.LABEL_WIDTH)
        layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

    def _with_dir_button(self, line_edit: QLineEdit) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton("选择目录")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(lambda: self.choose_directory(line_edit))
        row_layout.addWidget(button)
        return row

    def choose_source_file(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择文件")
        if selected:
            self.source_edit.setText(selected)

    def choose_source_dir(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择文件夹", self.source_edit.text().strip() or str(Path.cwd()))
        if selected:
            self.source_edit.setText(selected)

    def choose_directory(self, line_edit: QLineEdit) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择目录", line_edit.text().strip() or str(Path.cwd()))
        if selected:
            line_edit.setText(selected)

    def update_password_state(self, checked: bool) -> None:
        self.password_edit.setReadOnly(not checked)
        if checked:
            self.password_edit.clear()

    def start_pack(self) -> None:
        source_text = self.source_edit.text().strip()
        output_text = self.output_edit.text().strip()
        if not source_text or not output_text:
            QMessageBox.critical(self, "参数错误", "请先选择待打包对象和输出目录。")
            return
        source_path = Path(source_text)
        output_dir = Path(output_text)
        password = self.password_edit.text().strip() if self.custom_password_checkbox.isChecked() else None
        self.run_button.setEnabled(False)
        self.result_text.setPlainText("正在打包，请稍候...")
        worker = SimpleWorker(
            task=lambda: pack_encrypted_folder(source_path, output_dir, password=password, logger=self.log_fn),
            on_success=self.on_pack_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def on_pack_success(self, summary: PackSummary) -> None:
        self.run_button.setEnabled(True)
        self.copy_button.setEnabled(True)
        self.last_summary = summary
        self.password_edit.setText(summary.password)
        self.result_text.setPlainText(
            "\n".join(
                [
                    f"来源：{summary.source_path}",
                    f"压缩包：{summary.output_path}",
                    f"文件数：{summary.file_count}",
                    f"密码：{summary.password}",
                    f"时间：{summary.created_at}",
                ]
            )
        )

    def run_query(self) -> None:
        records = query_pack_history(self.query_edit.text().strip())
        if not records:
            self.result_text.setPlainText("未找到匹配记录。")
            return
        lines = []
        for item in records:
            lines.append(
                f"{item.get('archive_name','')}\n来源：{item.get('source_name','')}\n密码：{item.get('password','')}\n时间：{item.get('created_at','')}"
            )
        self.result_text.setPlainText("\n\n".join(lines))

    def copy_password(self) -> None:
        if self.last_summary is None:
            return
        QApplication.clipboard().setText(self.last_summary.password)
        self.log_fn("已复制打包密码。")

    def on_error(self, error_text: str) -> None:
        self.run_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])
