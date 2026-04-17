from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..job_code_audit import (
    AuditResult,
    export_job_audit_result,
    load_job_audit_headers,
    run_job_code_audit,
    validate_required_columns,
)


class JobCodeAuditPage(QWidget):
    LABEL_WIDTH = 118
    ACTION_WIDTH = 110

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.result: AuditResult | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        title = QLabel("岗位表核对")
        title.setProperty("heroTitle", True)
        intro = QLabel("导入岗位表后，校验编码格式、上下级绑定，以及地市下主管、主管下单位、单位下岗位的顺延规则。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
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
        self.excel_edit = QLineEdit()
        self._add_row(form, 0, "岗位表", self._with_file_button(self.excel_edit))
        left_layout.addLayout(form)

        tips = QPlainTextEdit()
        tips.setReadOnly(True)
        tips.setPlainText(
            "校验规则：\n"
            "1. 编码结构固定为 2 + 3 + 3 + 2。\n"
            "2. 相同编码不能绑定不同名称，相同名称不能绑定不同编码。\n"
            "3. 地市下主管从 001 开始，主管下单位从 001 开始，单位下岗位从 01 开始。\n"
            "4. 主管部门、报考单位、报考岗位名称允许为空，但空名称也占一个编码位。"
        )
        tips.setMinimumHeight(160)
        left_layout.addWidget(tips)

        action_row = QHBoxLayout()
        run_button = QPushButton("开始核对")
        run_button.setProperty("accent", True)
        run_button.clicked.connect(self.run_audit)
        export_button = QPushButton("导出问题")
        export_button.clicked.connect(self.export_result)
        action_row.addWidget(run_button)
        action_row.addWidget(export_button)
        action_row.addStretch(1)
        left_layout.addLayout(action_row)

        self.summary_text = QPlainTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setPlaceholderText("核对完成后，这里会显示摘要。")
        self.summary_text.setMinimumHeight(180)
        left_layout.addWidget(self.summary_text, 1)
        body.addWidget(left, 0, 0)

        right = self._card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_layout.setSpacing(16)
        title = QLabel("问题明细")
        title.setProperty("sectionTitle", True)
        right_layout.addWidget(title)

        self.issue_table = QTableWidget(0, 12)
        self.issue_table.setHorizontalHeaderLabels(
            [
                "问题类型",
                "层级",
                "Excel行号",
                "地市",
                "地市代码",
                "主管部门",
                "主管部门编码",
                "报考单位",
                "报考单位编码",
                "报考岗位",
                "报考岗位编码",
                "问题说明",
            ]
        )
        self.issue_table.horizontalHeader().setStretchLastSection(True)
        self.issue_table.verticalHeader().setVisible(False)
        self.issue_table.setAlternatingRowColors(True)
        right_layout.addWidget(self.issue_table, 1)
        body.addWidget(right, 0, 1)
        body.setColumnStretch(0, 2)
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
        button.clicked.connect(lambda: self.choose_file(line_edit))
        row_layout.addWidget(button)
        return row

    def choose_file(self, line_edit: QLineEdit) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if selected:
            line_edit.setText(selected)

    def run_audit(self) -> None:
        excel_path = Path(self.excel_edit.text().strip())
        if not excel_path.exists():
            QMessageBox.warning(self, "参数错误", "请先选择岗位表。")
            return

        try:
            headers = load_job_audit_headers(excel_path)
            missing_columns = validate_required_columns(headers)
            if missing_columns:
                raise ValueError(f"缺少必填列：{'、'.join(missing_columns)}")
            self.result = run_job_code_audit(excel_path)
        except Exception as exc:
            QMessageBox.critical(self, "核对失败", str(exc))
            return

        self._render_summary(self.result)
        self._render_issues(self.result)
        self.log_fn(f"岗位表核对完成：{excel_path.name}，共 {self.result.total_rows} 行，发现 {self.result.issue_count} 条问题。")

    def _render_summary(self, result: AuditResult) -> None:
        if result.issue_count == 0:
            self.summary_text.setPlainText(
                f"源文件：{result.source_path}\n"
                f"数据行数：{result.total_rows}\n"
                "问题条数：0\n\n"
                "当前未发现问题。"
            )
            return

        type_counts: dict[str, int] = {}
        for issue in result.issues:
            type_counts[issue.issue_type] = type_counts.get(issue.issue_type, 0) + 1
        lines = [
            f"源文件：{result.source_path}",
            f"数据行数：{result.total_rows}",
            f"问题条数：{result.issue_count}",
            "",
            "问题分布：",
        ]
        lines.extend(f"- {issue_type}：{count}" for issue_type, count in sorted(type_counts.items()))
        self.summary_text.setPlainText("\n".join(lines))

    def _render_issues(self, result: AuditResult) -> None:
        self.issue_table.setRowCount(0)
        for issue in result.issues:
            row = self.issue_table.rowCount()
            self.issue_table.insertRow(row)
            values = [
                issue.issue_type,
                issue.level,
                issue.excel_rows,
                issue.city,
                issue.city_code,
                issue.department,
                issue.department_code,
                issue.unit,
                issue.unit_code,
                issue.job,
                issue.job_code,
                issue.description,
            ]
            for column, value in enumerate(values):
                self.issue_table.setItem(row, column, QTableWidgetItem(value))
        self.issue_table.resizeColumnsToContents()

    def export_result(self) -> None:
        if self.result is None:
            QMessageBox.information(self, "提示", "请先完成一次核对。")
            return
        default_dir = str(self.result.source_path.parent / f"{self.result.source_path.stem}_岗位核对结果")
        selected = QFileDialog.getExistingDirectory(
            self,
            "选择导出目录",
            default_dir,
        )
        if not selected:
            return
        try:
            output_paths = export_job_audit_result(self.result, Path(selected))
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))
            return
        QMessageBox.information(
            self,
            "导出成功",
            f"已按地市拆分导出 {len(output_paths)} 个文件到：\n{selected}",
        )
        self.log_fn(f"岗位表核对结果已按地市拆分导出：{selected}，共 {len(output_paths)} 个文件。")
