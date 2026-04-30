from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
)

from ..ordered_unit_audit import (
    OrderedNameAuditResult,
    export_ordered_name_audit,
    load_excel_headers,
    run_ordered_name_audit,
)
from .base_page import BaseToolPage
from .common import AppComboBox
from .widgets import FormRow, PathInput, PrimaryButton, SecondaryButton, Section


class OrderedUnitAuditPage(BaseToolPage):
    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="单列顺序核对",
            description="严格按原表顺序检查你指定的一列，找出上下相似和分段重复的情况。",
            show_steps=False,
            show_log=True,
        )
        self.log_fn = log_fn
        self.result: OrderedNameAuditResult | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        if self.left_card.title is not None:
            self.left_card.title.setText("核对参数")

        data_section = Section("数据配置")
        self.excel_input = PathInput("请选择数据表文件", "选择文件")
        self.excel_edit = self.excel_input.edit
        self.excel_input.button.clicked.connect(self.choose_file)
        data_section.add(FormRow("数据表", self.excel_input))
        self.column_combo = AppComboBox()
        data_section.add(FormRow("目标列", self.column_combo))
        self.left_card.body_layout.addWidget(data_section)

        action_row = QHBoxLayout()
        load_button = QPushButton("加载列")
        load_button.clicked.connect(self.load_headers)
        run_button = PrimaryButton("开始核对")
        run_button.clicked.connect(self.run_audit)
        export_button = SecondaryButton("导出结果")
        export_button.clicked.connect(self.export_result)
        action_row.addWidget(load_button)
        action_row.addWidget(run_button)
        action_row.addWidget(export_button)
        action_row.addStretch(1)
        self.left_card.body_layout.addLayout(action_row)

        tips = QPlainTextEdit()
        tips.setReadOnly(True)
        tips.setPlainText(
            "核对规则：\n"
            "1. 严格按原表顺序核对，不会先排序。\n"
            "2. 只比较你当前选中的这一列。\n"
            "3. 如果上下两条名称相似，会摘出来人工核对。\n"
            "4. 如果某名称前面出现过，中间夹了别的内容，后面又重新出现，也会标出来。"
        )
        tips.setMinimumHeight(170)
        self.left_card.body_layout.addWidget(tips)

        self.summary_text = QPlainTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setPlaceholderText("核对完成后，这里会显示摘要。")
        self.summary_text.setMinimumHeight(180)
        self.left_card.body_layout.addWidget(self.summary_text, 1)

        if self.right_card.title is not None:
            self.right_card.title.setText("需核对清单")

        self.issue_table = QTableWidget(0, 6)
        self.issue_table.setHorizontalHeaderLabels(
            ["问题类型", "上一处行号", "当前行号", "上一处名称", "当前名称", "说明"]
        )
        self.issue_table.horizontalHeader().setStretchLastSection(True)
        self.issue_table.verticalHeader().setVisible(False)
        self.issue_table.setAlternatingRowColors(True)
        self.right_card.body_layout.addWidget(self.issue_table, 1)

    def choose_file(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)"
        )
        if selected:
            self.excel_edit.setText(selected)

    def load_headers(self) -> None:
        excel_path = Path(self.excel_edit.text().strip())
        if not excel_path.exists():
            QMessageBox.warning(self, "参数错误", "请先选择数据表。")
            return
        try:
            headers = load_excel_headers(excel_path)
        except Exception as exc:
            QMessageBox.critical(self, "加载失败", str(exc))
            return
        self.column_combo.clear()
        self.column_combo.addItems(headers)

    def run_audit(self) -> None:
        excel_path = Path(self.excel_edit.text().strip())
        if not excel_path.exists():
            QMessageBox.warning(self, "参数错误", "请先选择数据表。")
            return
        target_column = self.column_combo.currentText().strip()
        if not target_column:
            QMessageBox.warning(self, "参数错误", "请先加载列并选择目标列。")
            return
        try:
            self.result = run_ordered_name_audit(excel_path, target_column)
        except Exception as exc:
            QMessageBox.critical(self, "核对失败", str(exc))
            return
        self._render_summary(self.result)
        self._render_issues(self.result)
        self.log_fn(
            f"单列顺序核对完成：{excel_path.name}，目标列 {target_column}，"
            f"共检查 {self.result.total_rows} 行，发现 {len(self.result.issues)} 条问题。"
        )

    def _render_summary(self, result: OrderedNameAuditResult) -> None:
        type_counts: dict[str, int] = {}
        for issue in result.issues:
            type_counts[issue.issue_type] = type_counts.get(issue.issue_type, 0) + 1
        lines = [
            f"源文件：{result.source_path}",
            f"目标列：{result.target_column}",
            f"参与核对行数：{result.total_rows}",
            f"问题条数：{len(result.issues)}",
            "",
            "问题分布：",
        ]
        if type_counts:
            lines.extend(
                f"- {issue_type}：{count}" for issue_type, count in sorted(type_counts.items())
            )
        else:
            lines.append("- 当前未发现问题")
        self.summary_text.setPlainText("\n".join(lines))

    def _render_issues(self, result: OrderedNameAuditResult) -> None:
        self.issue_table.setRowCount(0)
        for issue in result.issues:
            row = self.issue_table.rowCount()
            self.issue_table.insertRow(row)
            values = [
                issue.issue_type,
                str(issue.first_row),
                str(issue.second_row),
                issue.first_name,
                issue.second_name,
                issue.description,
            ]
            for column, value in enumerate(values):
                self.issue_table.setItem(row, column, QTableWidgetItem(value))
        self.issue_table.resizeColumnsToContents()

    def export_result(self) -> None:
        if self.result is None:
            QMessageBox.information(self, "提示", "请先完成一次核对。")
            return
        default_name = self.result.source_path.with_name(
            f"{self.result.source_path.stem}_单列顺序核对.xlsx"
        )
        selected, _ = QFileDialog.getSaveFileName(self, "导出结果", str(default_name), "Excel 文件 (*.xlsx)")
        if not selected:
            return
        try:
            output_path = export_ordered_name_audit(self.result, Path(selected))
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))
            return
        QMessageBox.information(self, "导出成功", f"结果已导出到：\n{output_path}")
        self.log_fn(f"单列顺序核对结果已导出：{output_path}")
