from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
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
    QVBoxLayout,
    QWidget,
)

from ..update_sql_generator import (
    export_update_sql_template,
    load_update_field_mappings,
    render_update_sql,
)


class UpdateSqlPage(QWidget):
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

        hero = self._create_card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)
        title = QLabel("更新 SQL 生成")
        title.setProperty("heroTitle", True)
        intro = QLabel("通过字段映射模板生成标准 UPDATE SQL，可选择忽略空值覆盖正式表。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        body = QGridLayout()
        body.setHorizontalSpacing(16)
        body.setVerticalSpacing(16)
        root.addLayout(body, 1)

        left = self._create_card()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(24, 22, 24, 24)
        left_layout.setSpacing(18)
        left_title = QLabel("模板与关联设置")
        left_title.setProperty("sectionTitle", True)
        left_layout.addWidget(left_title)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.mapping_edit = QLineEdit()
        mapping_row = self._with_file_button(self.mapping_edit, self.choose_mapping, "选择文件")
        self._add_row(form, 0, "映射模板", mapping_row)

        template_actions = QWidget()
        template_actions_layout = QHBoxLayout(template_actions)
        template_actions_layout.setContentsMargins(0, 0, 0, 0)
        template_actions_layout.setSpacing(10)
        load_button = QPushButton("加载字段")
        load_button.setFixedWidth(self.ACTION_WIDTH)
        load_button.clicked.connect(self.load_headers)
        export_button = QPushButton("导出模板")
        export_button.setFixedWidth(self.ACTION_WIDTH)
        export_button.clicked.connect(self.export_template)
        template_actions_layout.addWidget(load_button)
        template_actions_layout.addWidget(export_button)
        template_actions_layout.addStretch(1)
        self._add_row(form, 1, "", template_actions)

        self.target_table_edit = QLineEdit()
        self._add_row(form, 2, "考生表名称", self.target_table_edit)
        self.source_table_edit = QLineEdit()
        self._add_row(form, 3, "临时表名称", self.source_table_edit)

        self.target_key_combo = QComboBox()
        self._add_row(form, 4, "考生表关联字段", self.target_key_combo)
        self.source_key_combo = QComboBox()
        self._add_row(form, 5, "临时表关联字段", self.source_key_combo)

        self.ignore_empty_checkbox = QCheckBox("忽略空值，不覆盖正式表")
        form.addWidget(self.ignore_empty_checkbox, 6, 1)
        left_layout.addLayout(form)
        left_layout.addStretch(1)

        action_row = QHBoxLayout()
        self.run_button = QPushButton("生成 SQL")
        self.run_button.setProperty("accent", True)
        self.run_button.setFixedWidth(132)
        self.run_button.clicked.connect(self.render_sql)
        self.copy_button = QPushButton("复制 SQL")
        self.copy_button.setFixedWidth(132)
        self.copy_button.setEnabled(False)
        self.copy_button.clicked.connect(self.copy_sql)
        action_row.addWidget(self.run_button)
        action_row.addWidget(self.copy_button)
        action_row.addStretch(1)
        left_layout.addLayout(action_row)
        body.addWidget(left, 0, 0)

        right = self._create_card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_layout.setSpacing(14)
        right_title = QLabel("生成结果")
        right_title.setProperty("sectionTitle", True)
        right_layout.addWidget(right_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        right_layout.addWidget(self.result_text, 1)
        body.addWidget(right, 0, 1)

        body.setColumnStretch(0, 3)
        body.setColumnStretch(1, 4)

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> None:
        if label_text:
            label = QLabel(label_text)
            label.setProperty("formLabel", True)
            label.setFixedWidth(self.LABEL_WIDTH)
            layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

    def _with_file_button(self, line_edit: QLineEdit, callback, button_text: str) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton(button_text)
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(callback)
        row_layout.addWidget(button)
        return row

    def choose_mapping(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "选择映射模板",
            "",
            "Excel 文件 (*.xlsx *.xls)",
        )
        if selected:
            self.mapping_edit.setText(selected)

    def export_template(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "导出字段映射模板",
            "更新SQL字段映射模板.xlsx",
            "Excel 文件 (*.xlsx)",
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
