from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import (
    QApplication,
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

from ..sql_template_executor import render_sql_template


class SqlExecPage(QWidget):
    LABEL_WIDTH = 132
    ACTION_WIDTH = 144

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.last_sql = ""
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        title = QLabel("SQL 配置执行")
        title.setProperty("heroTitle", True)
        intro = QLabel("通过现成 SQL 模板填入考试代码、考试年月和时间参数，生成一键执行 SQL。")
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

        self.template_edit = QLineEdit()
        self._add_row(form, 0, "SQL 模板", self._with_file_button(self.template_edit))
        self.exam_code_edit = QLineEdit()
        self._add_row(form, 1, "考试代码", self.exam_code_edit)
        self.exam_date_edit = QLineEdit()
        self._add_row(form, 2, "考试年月", self.exam_date_edit)
        self.signup_start_edit = QLineEdit()
        self._add_row(form, 3, "报名开始时间", self.signup_start_edit)
        self.signup_end_edit = QLineEdit()
        self._add_row(form, 4, "报名结束时间", self.signup_end_edit)
        self.audit_start_edit = QLineEdit()
        self._add_row(form, 5, "审核开始时间", self.audit_start_edit)
        self.audit_end_edit = QLineEdit()
        self._add_row(form, 6, "审核结束时间", self.audit_end_edit)
        left_layout.addLayout(form)
        left_layout.addStretch(1)

        actions = QHBoxLayout()
        generate_button = QPushButton("生成 SQL")
        generate_button.setProperty("accent", True)
        generate_button.clicked.connect(self.render_sql)
        actions.addWidget(generate_button)
        copy_button = QPushButton("复制 SQL")
        copy_button.clicked.connect(self.copy_sql)
        actions.addWidget(copy_button)
        actions.addStretch(1)
        left_layout.addLayout(actions)
        body.addWidget(left, 0, 0)

        right = self._card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_title = QLabel("生成结果")
        right_title.setProperty("sectionTitle", True)
        right_layout.addWidget(right_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        right_layout.addWidget(self.result_text, 1)
        body.addWidget(right, 0, 1)
        body.setColumnStretch(0, 3)
        body.setColumnStretch(1, 4)

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

    def _with_file_button(self, line_edit: QLineEdit) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton("选择文件")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(self.choose_template)
        row_layout.addWidget(button)
        return row

    def choose_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择 SQL 模板", "", "SQL 文件 (*.sql);;所有文件 (*.*)")
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
