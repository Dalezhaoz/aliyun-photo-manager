from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QWidget,
)

from ..update_sql_generator import (
    export_update_sql_template,
    load_update_field_mappings,
    render_update_sql,
)
from .base_page import BaseToolPage
from .common import AppComboBox
from .widgets import FormRow, PathInput, PrimaryButton, SecondaryButton, Section


class UpdateSqlPage(BaseToolPage):
    ACTION_WIDTH = 116

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="更新 SQL 生成",
            description="通过字段映射模板生成标准 UPDATE SQL，可选择忽略空值覆盖正式表。",
            show_steps=False,
            show_log=True,
        )
        self.log_fn = log_fn
        self.last_sql = ""
        self._build_ui()

    def _build_ui(self) -> None:
        if self.left_card.title is not None:
            self.left_card.title.setText("模板与关联设置")

        mapping_section = Section("字段映射")
        self.mapping_input = PathInput("请选择字段映射模板", "选择文件")
        self.mapping_edit = self.mapping_input.edit
        self.mapping_input.button.clicked.connect(self.choose_mapping)
        mapping_section.add(FormRow("映射模板", self.mapping_input))

        btn_row = QWidget()
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(10)
        load_button = QPushButton("加载字段")
        load_button.setFixedWidth(self.ACTION_WIDTH)
        load_button.clicked.connect(self.load_headers)
        export_button = QPushButton("导出模板")
        export_button.setFixedWidth(self.ACTION_WIDTH)
        export_button.clicked.connect(self.export_template)
        btn_layout.addWidget(load_button)
        btn_layout.addWidget(export_button)
        btn_layout.addStretch(1)
        mapping_section.add(FormRow("", btn_row))
        self.left_card.body_layout.addWidget(mapping_section)

        table_section = Section("表名与关联字段")
        self.target_table_edit = QLineEdit()
        table_section.add(FormRow("考生表名称", self.target_table_edit))
        self.source_table_edit = QLineEdit()
        table_section.add(FormRow("临时表名称", self.source_table_edit))
        self.target_key_combo = AppComboBox()
        table_section.add(FormRow("考生表关联字段", self.target_key_combo))
        self.source_key_combo = AppComboBox()
        table_section.add(FormRow("临时表关联字段", self.source_key_combo))
        self.ignore_empty_checkbox = QCheckBox("忽略空值，不覆盖正式表")
        table_section.add(self.ignore_empty_checkbox)
        self.left_card.body_layout.addWidget(table_section)

        action_row = QHBoxLayout()
        self.run_button = PrimaryButton("生成 SQL")
        self.run_button.clicked.connect(self.render_sql)
        action_row.addWidget(self.run_button)
        self.copy_button = SecondaryButton("复制 SQL")
        self.copy_button.setEnabled(False)
        self.copy_button.clicked.connect(self.copy_sql)
        action_row.addWidget(self.copy_button)
        action_row.addStretch(1)
        self.left_card.body_layout.addLayout(action_row)
        self.left_card.body_layout.addStretch(1)

        if self.right_card.title is not None:
            self.right_card.title.setText("生成结果")

        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.right_card.body_layout.addWidget(self.result_text, 1)

    def choose_mapping(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self, "选择映射模板", "", "Excel 文件 (*.xlsx *.xls)"
        )
        if selected:
            self.mapping_edit.setText(selected)

    def export_template(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(
            self, "导出字段映射模板", "更新SQL字段映射模板.xlsx", "Excel 文件 (*.xlsx)"
        )
        if not selected:
            return
        summary = export_update_sql_template(Path(selected))
        self.result_text.setPlainText(f"模板已导出：{summary.output_path}")
        self.log_fn(f"已导出更新 SQL 模板：{summary.output_path}")

    def load_headers(self) -> None:
        mapping_path = Path(self.mapping_edit.text().strip())
        if not mapping_path.exists():
            QMessageBox.critical(self, "缺少模板", "请先选择有效的映射模板。")
            return
        try:
            _, target_headers, source_headers = load_update_field_mappings(mapping_path)
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            return
        self.target_key_combo.clear()
        self.target_key_combo.addItems(target_headers)
        self.source_key_combo.clear()
        self.source_key_combo.addItems(source_headers)
        self.log_fn(f"已加载更新 SQL 字段：目标 {len(target_headers)} 列，来源 {len(source_headers)} 列。")

    def render_sql(self) -> None:
        mapping_path = Path(self.mapping_edit.text().strip())
        try:
            result = render_update_sql(
                mapping_path=mapping_path,
                target_table=self.target_table_edit.text().strip(),
                source_table=self.source_table_edit.text().strip(),
                target_key_column=self.target_key_combo.currentText().strip(),
                source_key_column=self.source_key_combo.currentText().strip(),
                ignore_empty=self.ignore_empty_checkbox.isChecked(),
                logger=self.log_fn,
            )
        except Exception as exc:
            QMessageBox.critical(self, "生成失败", str(exc))
            return
        self.last_sql = result.sql_content
        self.result_text.setPlainText(result.sql_content)
        self.copy_button.setEnabled(True)

    def copy_sql(self) -> None:
        if not self.last_sql:
            return
        QApplication.clipboard().setText(self.last_sql)
        self.log_fn("已复制更新 SQL。")
