from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, Signal, QRunnable, QThreadPool
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from ..phone_decrypt import (
    PhoneDecryptOptions,
    PhoneDecryptSummary,
    load_filter_id_cards,
    run_phone_decrypt,
)


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


class PhoneDecryptPage(QWidget):
    LABEL_WIDTH = 132
    ACTION_WIDTH = 144

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._create_card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)
        title = QLabel("电话解密")
        title.setProperty("heroTitle", True)
        intro = QLabel("通过 helper 解密 `web_info.info1`，并回写到考生表 `备用3`。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        body = QGridLayout()
        body.setHorizontalSpacing(16)
        body.setVerticalSpacing(16)
        root.addLayout(body, 1)

        card = self._create_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(24, 22, 24, 24)
        card_layout.setSpacing(18)
        section_title = QLabel("解密参数")
        section_title.setProperty("sectionTitle", True)
        card_layout.addWidget(section_title)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)
        form.setColumnStretch(0, 0)
        form.setColumnStretch(1, 1)
        row_index = 0

        self.server_edit = QLineEdit()
        self._add_form_row(form, row_index, "服务器", self.server_edit)
        row_index += 1
        self.port_edit = QLineEdit("1433")
        self._add_form_row(form, row_index, "端口", self.port_edit)
        row_index += 1
        self.username_edit = QLineEdit()
        self._add_form_row(form, row_index, "用户名", self.username_edit)
        row_index += 1
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self._add_form_row(form, row_index, "密码", self.password_edit)
        row_index += 1
        self.signup_db_edit = QLineEdit()
        self._add_form_row(form, row_index, "报名数据库", self.signup_db_edit)
        row_index += 1
        self.phone_db_edit = QLineEdit()
        self._add_form_row(form, row_index, "电话数据库", self.phone_db_edit)
        row_index += 1
        self.exam_sort_edit = QLineEdit()
        self._add_form_row(form, row_index, "考试代码", self.exam_sort_edit)
        row_index += 1

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
        self._add_form_row(form, row_index, "考生表名", table_row)
        row_index += 1

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
        self._add_form_row(form, row_index, "解密模式", mode_row)
        row_index += 1

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
        self._add_form_row(form, row_index, "名单文件", filter_row)
        row_index += 1
        card_layout.addLayout(form)
        card_layout.addStretch(1)

        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 8, 0, 0)
        self.run_button = QPushButton("开始解密")
        self.run_button.setProperty("accent", True)
        self.run_button.setFixedWidth(150)
        self.run_button.clicked.connect(self.start_run)
        action_row.addWidget(self.run_button)
        action_row.addStretch(1)
        card_layout.addLayout(action_row)
        body.addWidget(card, 0, 0)

        result_card = self._create_card()
        result_layout = QVBoxLayout(result_card)
        result_layout.setContentsMargins(24, 22, 24, 24)
        result_layout.setSpacing(14)
        result_title = QLabel("解密结果")
        result_title.setProperty("sectionTitle", True)
        result_layout.addWidget(result_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        result_layout.addWidget(self.result_text)
        body.addWidget(result_card, 0, 1)

        body.setColumnStretch(0, 3)
        body.setColumnStretch(1, 2)

        self.update_mode_state()

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_form_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> None:
        label = QLabel(label_text)
        label.setProperty("formLabel", True)
        label.setFixedWidth(self.LABEL_WIDTH)
        layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

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
