from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QWidget,
)

from ..phone_decrypt import (
    PhoneDecryptOptions,
    PhoneDecryptSummary,
    load_filter_id_cards,
    run_phone_decrypt,
)
from .base_page import BaseToolPage
from .widgets import FormRow, PrimaryButton, SecondaryButton, Section


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class PhoneWorker(QRunnable):
    def __init__(
        self,
        options: PhoneDecryptOptions,
        log_fn: Callable[[str], None],
        on_success: Callable[[PhoneDecryptSummary], None],
        on_error: Callable[[str], None],
    ) -> None:
        super().__init__()
        self.options = options
        self.log_fn = log_fn
        self.signals = WorkerSignals()
        self.signals.success.connect(on_success)
        self.signals.error.connect(on_error)

    def run(self) -> None:
        try:
            summary = run_phone_decrypt(self.options, logger=self.log_fn)
        except Exception as exc:
            self.signals.error.emit(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}")
        else:
            self.signals.success.emit(summary)


class PhoneDecryptPage(BaseToolPage):
    ACTION_WIDTH = 116

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="电话解密",
            description="通过 helper 解密 web_info.info1，并回写到考生表备用3。",
            show_steps=False,
            show_log=True,
        )
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self._build_ui()

    def _build_ui(self) -> None:
        if self.left_card.title is not None:
            self.left_card.title.setText("解密参数")

        conn_section = Section("数据库连接")
        self.server_edit = QLineEdit()
        conn_section.add(FormRow("服务器", self.server_edit))
        self.port_edit = QLineEdit("1433")
        conn_section.add(FormRow("端口", self.port_edit))
        self.username_edit = QLineEdit()
        conn_section.add(FormRow("用户名", self.username_edit))
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        conn_section.add(FormRow("密码", self.password_edit))
        self.signup_db_edit = QLineEdit()
        conn_section.add(FormRow("报名数据库", self.signup_db_edit))
        self.phone_db_edit = QLineEdit()
        conn_section.add(FormRow("电话数据库", self.phone_db_edit))
        self.left_card.body_layout.addWidget(conn_section)

        table_section = Section("考生表配置")
        self.exam_sort_edit = QLineEdit()
        table_section.add(FormRow("考试代码", self.exam_sort_edit))
        table_row = QWidget()
        table_layout = QHBoxLayout(table_row)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.setSpacing(10)
        self.table_edit = QLineEdit()
        table_layout.addWidget(self.table_edit, 1)
        fill_button = QPushButton("生成表名")
        fill_button.setFixedWidth(self.ACTION_WIDTH)
        fill_button.clicked.connect(self.fill_table_name)
        table_layout.addWidget(fill_button)
        table_section.add(FormRow("考生表名", table_row))
        self.left_card.body_layout.addWidget(table_section)

        mode_section = Section("解密模式")
        mode_row = QWidget()
        mode_layout = QHBoxLayout(mode_row)
        mode_layout.setContentsMargins(0, 0, 0, 0)
        mode_layout.setSpacing(16)
        self.mode_all = QRadioButton("解密全部电话")
        self.mode_partial = QRadioButton("按名单解密")
        self.mode_all.setChecked(True)
        self.mode_all.toggled.connect(self.update_mode_state)
        mode_layout.addWidget(self.mode_all)
        mode_layout.addWidget(self.mode_partial)
        mode_layout.addStretch(1)
        mode_section.add(FormRow("解密模式", mode_row))

        filter_row = QWidget()
        filter_layout = QHBoxLayout(filter_row)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(10)
        self.filter_edit = QLineEdit()
        filter_layout.addWidget(self.filter_edit, 1)
        self.filter_button = QPushButton("选择文件")
        self.filter_button.setFixedWidth(self.ACTION_WIDTH)
        self.filter_button.clicked.connect(self.choose_filter_file)
        filter_layout.addWidget(self.filter_button)
        mode_section.add(FormRow("名单文件", filter_row))
        self.left_card.body_layout.addWidget(mode_section)

        action_row = QHBoxLayout()
        self.run_button = PrimaryButton("开始解密")
        self.run_button.clicked.connect(self.start_run)
        action_row.addWidget(self.run_button)
        action_row.addStretch(1)
        self.left_card.body_layout.addLayout(action_row)
        self.left_card.body_layout.addStretch(1)

        if self.right_card.title is not None:
            self.right_card.title.setText("解密结果")

        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.right_card.body_layout.addWidget(self.result_text, 1)

        self.update_mode_state()

    def fill_table_name(self) -> None:
        exam_sort = self.exam_sort_edit.text().strip()
        if exam_sort:
            self.table_edit.setText(f"考生表{exam_sort}")

    def choose_filter_file(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "选择名单文件",
            "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)",
        )
        if selected:
            self.filter_edit.setText(selected)

    def update_mode_state(self) -> None:
        enabled = self.mode_partial.isChecked()
        self.filter_edit.setEnabled(enabled)
        self.filter_button.setEnabled(enabled)

    def start_run(self) -> None:
        try:
            options = self.build_options()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return
        self.run_button.setEnabled(False)
        self.result_text.setPlainText("正在查询、解密并回写，请稍候...")
        self.log_fn(f"启动电话解密：{options.signup_database}.{options.candidate_table}")
        worker = PhoneWorker(options, self.log_fn, self.on_success, self.on_error)
        self.thread_pool.start(worker)

    def build_options(self) -> PhoneDecryptOptions:
        server = self.server_edit.text().strip()
        username = self.username_edit.text().strip()
        password = self.password_edit.text().strip()
        signup_db = self.signup_db_edit.text().strip()
        phone_db = self.phone_db_edit.text().strip() or signup_db
        table_name = self.table_edit.text().strip()
        if not table_name and self.exam_sort_edit.text().strip():
            table_name = f"考生表{self.exam_sort_edit.text().strip()}"
            self.table_edit.setText(table_name)
        if not table_name:
            raise ValueError("请输入考生表名，或先填写考试代码生成表名。")
        try:
            port = int(self.port_edit.text().strip() or "1433")
        except ValueError as exc:
            raise ValueError("端口必须是数字。") from exc

        id_cards = None
        mode = "partial" if self.mode_partial.isChecked() else "all"
        if mode == "partial":
            filter_path = Path(self.filter_edit.text().strip())
            if not filter_path.exists():
                raise ValueError("按名单解密时，请选择有效名单文件。")
            id_cards = load_filter_id_cards(filter_path)
            if not id_cards:
                raise ValueError("名单文件中没有可用身份证号。")

        return PhoneDecryptOptions(
            server=server,
            port=port,
            username=username,
            password=password,
            signup_database=signup_db,
            phone_database=phone_db,
            candidate_table=table_name,
            candidate_filter_mode=mode,
            candidate_id_cards=id_cards,
        )

    def on_success(self, summary: PhoneDecryptSummary) -> None:
        self.run_button.setEnabled(True)
        lines = [
            f"报名库：{summary.signup_database}",
            f"电话库：{summary.phone_database}",
            f"考生表：{summary.candidate_table}",
            f"解密组件：{summary.backend_name}",
            f"总记录：{summary.total_rows}",
            f"命中密文：{summary.matched_info_rows}",
            f"解密成功：{summary.decrypted_rows}",
            f"回写备用3：{summary.updated_rows}",
            f"跳过：{summary.skipped_rows}",
            f"失败：{summary.failed_rows}",
        ]
        self.result_text.setPlainText("\n".join(lines))

    def on_error(self, error_text: str) -> None:
        self.run_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "电话解密失败", error_text.splitlines()[0])
