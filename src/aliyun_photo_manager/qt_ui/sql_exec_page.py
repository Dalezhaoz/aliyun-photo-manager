from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
)

from ..sql_template_executor import render_sql_template
from .base_page import BaseToolPage
from .widgets import FormRow, PathInput, PrimaryButton, SecondaryButton, Section


class SqlExecPage(BaseToolPage):
    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="SQL 配置执行",
            description="通过现成 SQL 模板填入考试代码、考试年月和时间参数，生成一键执行 SQL。",
            show_steps=False,
            show_log=True,
        )
        self.log_fn = log_fn
        self.last_sql = ""
        self._build_ui()

    def _build_ui(self) -> None:
        if self.left_card.title is not None:
            self.left_card.title.setText("模板参数")

        template_section = Section("模板")
        self.template_input = PathInput("选择 SQL 模板文件", "选择文件")
        self.template_edit = self.template_input.edit
        self.template_input.button.clicked.connect(self.choose_template)
        template_section.add(FormRow("SQL 模板", self.template_input))
        self.left_card.body_layout.addWidget(template_section)

        params_section = Section("查询参数")
        self.exam_code_edit = QLineEdit()
        params_section.add(FormRow("考试代码", self.exam_code_edit))
        self.exam_date_edit = QLineEdit()
        params_section.add(FormRow("考试年月", self.exam_date_edit))
        self.signup_start_edit = QLineEdit()
        params_section.add(FormRow("报名开始时间", self.signup_start_edit))
        self.signup_end_edit = QLineEdit()
        params_section.add(FormRow("报名结束时间", self.signup_end_edit))
        self.audit_start_edit = QLineEdit()
        params_section.add(FormRow("审核开始时间", self.audit_start_edit))
        self.audit_end_edit = QLineEdit()
        params_section.add(FormRow("审核结束时间", self.audit_end_edit))
        self.left_card.body_layout.addWidget(params_section)

        action_row = QHBoxLayout()
        generate_button = PrimaryButton("生成 SQL")
        generate_button.clicked.connect(self.render_sql)
        action_row.addWidget(generate_button)
        copy_button = SecondaryButton("复制 SQL")
        copy_button.clicked.connect(self.copy_sql)
        action_row.addWidget(copy_button)
        action_row.addStretch(1)
        self.left_card.body_layout.addLayout(action_row)
        self.left_card.body_layout.addStretch(1)

        if self.right_card.title is not None:
            self.right_card.title.setText("生成结果")

        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.right_card.body_layout.addWidget(self.result_text, 1)

    def choose_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self, "选择 SQL 模板", "", "SQL 文件 (*.sql);;所有文件 (*.*)"
        )
        if selected:
            self.template_edit.setText(selected)

    def render_sql(self) -> None:
        try:
            result = render_sql_template(
                source_path=Path(self.template_edit.text().strip()),
                exam_code=self.exam_code_edit.text().strip(),
                exam_date=self.exam_date_edit.text().strip(),
                signup_start=self.signup_start_edit.text().strip(),
                signup_end=self.signup_end_edit.text().strip(),
                audit_start=self.audit_start_edit.text().strip(),
                audit_end=self.audit_end_edit.text().strip(),
            )
        except Exception as exc:
            QMessageBox.critical(self, "生成失败", str(exc))
            return
        self.last_sql = result.sql_content
        self.result_text.setPlainText(result.sql_content)
        self.log_fn("已生成 SQL 模板结果。")

    def copy_sql(self) -> None:
        if not self.last_sql:
            return
        QApplication.clipboard().setText(self.last_sql)
        self.log_fn("已复制 SQL。")
