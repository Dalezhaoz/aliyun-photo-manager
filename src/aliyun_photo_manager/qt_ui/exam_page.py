from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from ..exam_arranger import (
    ExamArrangeOptions,
    ExamArrangeSummary,
    ExamRuleItem,
    export_exam_templates,
    run_exam_arrangement,
)
from .base_page import BaseToolPage
from .common import AppComboBox
from .widgets import FormRow, PathInput, PrimaryButton, SecondaryButton, Section


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class ExamWorker(QRunnable):
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


class ExamPage(BaseToolPage):
    ACTION_WIDTH = 116

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="考场编排",
            description="按考生明细、岗位归组和编排片段生成考号、考点、考场与座号。",
            show_steps=False,
            show_log=True,
        )
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self._build_ui()

    def _build_ui(self) -> None:
        if self.left_card.title is not None:
            self.left_card.title.setText("编排参数")

        files_section = Section("数据文件")
        self._add_file_row(files_section, "考生明细表", self._make_file_edit("candidate"))
        self._add_file_row(files_section, "岗位归组表", self._make_file_edit("group"))
        self._add_file_row(files_section, "编排片段表", self._make_file_edit("plan"))
        self.output_input = PathInput("选择输出文件路径", "选择文件")
        self.output_edit = self.output_input.edit
        self.output_input.button.clicked.connect(self._choose_output_file)
        files_section.add(FormRow("输出文件", self.output_input))
        self.left_card.body_layout.addWidget(files_section)

        digits_section = Section("位数设置")
        self.point_digits_edit = QLineEdit("2")
        digits_section.add(FormRow("考点位数", self.point_digits_edit))
        self.room_digits_edit = QLineEdit("3")
        digits_section.add(FormRow("考场位数", self.room_digits_edit))
        self.seat_digits_edit = QLineEdit("2")
        digits_section.add(FormRow("座号位数", self.seat_digits_edit))
        self.serial_digits_edit = QLineEdit("3")
        digits_section.add(FormRow("流水号位数", self.serial_digits_edit))
        self.sort_mode_combo = AppComboBox()
        self.sort_mode_combo.addItem("按原顺序", "original")
        self.sort_mode_combo.addItem("随机打乱", "random")
        self.sort_mode_combo.setCurrentIndex(1)
        digits_section.add(FormRow("组内顺序", self.sort_mode_combo))
        self.left_card.body_layout.addWidget(digits_section)

        rule_section = Section("考号规则")
        rule_row = QHBoxLayout()
        self.rule_type_combo = AppComboBox()
        self.rule_type_combo.addItems(["考点", "考场", "座号", "流水号", "岗位编码", "科目号", "自定义"])
        self.rule_type_combo.currentTextChanged.connect(self.update_rule_custom_state)
        self.rule_custom_edit = QLineEdit()
        self.rule_custom_edit.setPlaceholderText("自定义内容")
        self.rule_custom_edit.setEnabled(False)
        add_rule_button = QPushButton("添加规则")
        add_rule_button.clicked.connect(self.add_rule)
        rule_row.addWidget(self.rule_type_combo, 1)
        rule_row.addWidget(self.rule_custom_edit, 1)
        rule_row.addWidget(add_rule_button)
        rule_section.body_layout.addLayout(rule_row)
        self.rule_table = QTableWidget(0, 2)
        self.rule_table.setHorizontalHeaderLabels(["项目", "自定义内容"])
        self.rule_table.horizontalHeader().setStretchLastSection(True)
        self.rule_table.verticalHeader().setVisible(False)
        self.rule_table.setMinimumHeight(140)
        rule_section.add(self.rule_table)
        rule_actions = QHBoxLayout()
        up_button = QPushButton("上移")
        up_button.clicked.connect(self.move_rule_up)
        down_button = QPushButton("下移")
        down_button.clicked.connect(self.move_rule_down)
        del_button = QPushButton("删除所选")
        del_button.clicked.connect(self.remove_rule)
        rule_actions.addWidget(up_button)
        rule_actions.addWidget(down_button)
        rule_actions.addWidget(del_button)
        rule_actions.addStretch(1)
        rule_section.body_layout.addLayout(rule_actions)
        self.left_card.body_layout.addWidget(rule_section)

        action_row = QHBoxLayout()
        export_button = QPushButton("导出标准模板")
        export_button.clicked.connect(self.export_templates)
        self.run_button = PrimaryButton("开始编排")
        self.run_button.clicked.connect(self.start_run)
        action_row.addWidget(export_button)
        action_row.addWidget(self.run_button)
        action_row.addStretch(1)
        self.left_card.body_layout.addLayout(action_row)
        self.left_card.body_layout.addStretch(1)

        if self.right_card.title is not None:
            self.right_card.title.setText("编排结果")

        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.right_card.body_layout.addWidget(self.result_text, 1)

        self.update_rule_custom_state(self.rule_type_combo.currentText())

    def _make_file_edit(self, key: str) -> PathInput:
        input_widget = PathInput(f"选择{key}表", "选择文件")
        setattr(self, f"{key}_edit", input_widget.edit)
        input_widget.button.clicked.connect(lambda checked=False, k=key: self._choose_input_file(k))
        return input_widget

    def _add_file_row(self, section: Section, label: str, input_widget: PathInput) -> None:
        section.add(FormRow(label, input_widget))

    def _choose_input_file(self, key: str) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if selected:
            edit: QLineEdit = getattr(self, f"{key}_edit")
            edit.setText(selected)

    def _choose_output_file(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(self, "选择输出文件", "", "Excel 文件 (*.xlsx)")
        if selected:
            self.output_edit.setText(selected)

    def add_rule(self) -> None:
        rule_type = self.rule_type_combo.currentText().strip()
        custom_text = self.rule_custom_edit.text().strip() if rule_type == "自定义" else ""
        row = self.rule_table.rowCount()
        self.rule_table.insertRow(row)
        self.rule_table.setItem(row, 0, QTableWidgetItem(rule_type))
        self.rule_table.setItem(row, 1, QTableWidgetItem(custom_text))
        if rule_type == "自定义":
            self.rule_custom_edit.clear()

    def update_rule_custom_state(self, value: str) -> None:
        enabled = value.strip() == "自定义"
        self.rule_custom_edit.setEnabled(enabled)
        if not enabled:
            self.rule_custom_edit.clear()

    def remove_rule(self) -> None:
        row = self.rule_table.currentRow()
        if row >= 0:
            self.rule_table.removeRow(row)

    def move_rule_up(self) -> None:
        row = self.rule_table.currentRow()
        if row <= 0:
            return
        self._swap_rows(row, row - 1)
        self.rule_table.selectRow(row - 1)

    def move_rule_down(self) -> None:
        row = self.rule_table.currentRow()
        if row < 0 or row >= self.rule_table.rowCount() - 1:
            return
        self._swap_rows(row, row + 1)
        self.rule_table.selectRow(row + 1)

    def _swap_rows(self, a: int, b: int) -> None:
        for col in range(self.rule_table.columnCount()):
            a_item = self.rule_table.item(a, col)
            b_item = self.rule_table.item(b, col)
            a_text = a_item.text() if a_item else ""
            b_text = b_item.text() if b_item else ""
            self.rule_table.setItem(a, col, QTableWidgetItem(b_text))
            self.rule_table.setItem(b, col, QTableWidgetItem(a_text))

    def export_templates(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择导出目录", str(Path.cwd()))
        if not selected:
            return
        try:
            summary = export_exam_templates(Path(selected))
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))
            return
        self.candidate_edit.setText(str(summary.candidate_template_path))
        self.group_edit.setText(str(summary.group_template_path))
        self.plan_edit.setText(str(summary.plan_template_path))
        self.result_text.setPlainText(
            "\n".join([
                f"考生模板：{summary.candidate_template_path}",
                f"岗位归组模板：{summary.group_template_path}",
                f"编排片段模板：{summary.plan_template_path}",
            ])
        )

    def _collect_rules(self) -> list[ExamRuleItem]:
        rules: list[ExamRuleItem] = []
        for row in range(self.rule_table.rowCount()):
            type_item = self.rule_table.item(row, 0)
            custom_item = self.rule_table.item(row, 1)
            item_type = type_item.text().strip() if type_item else ""
            custom_text = custom_item.text().strip() if custom_item else ""
            if item_type:
                rules.append(ExamRuleItem(item_type=item_type, custom_text=custom_text))
        return rules

    def start_run(self) -> None:
        try:
            candidate_path = Path(self.candidate_edit.text().strip())
            group_path = Path(self.group_edit.text().strip())
            plan_path = Path(self.plan_edit.text().strip())
            output_path = Path(self.output_edit.text().strip()) if self.output_edit.text().strip() else None
            options = ExamArrangeOptions(
                candidate_path=candidate_path,
                group_path=group_path,
                plan_path=plan_path,
                output_path=output_path,
                exam_point_digits=int(self.point_digits_edit.text().strip()),
                room_digits=int(self.room_digits_edit.text().strip()),
                seat_digits=int(self.seat_digits_edit.text().strip()),
                serial_digits=int(self.serial_digits_edit.text().strip()),
                sort_mode=str(self.sort_mode_combo.currentData()),
                rule_items=self._collect_rules(),
            )
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return
        self.run_button.setEnabled(False)
        self.result_text.setPlainText("正在编排，请稍候...")
        worker = ExamWorker(
            task=lambda: run_exam_arrangement(options, logger=self.log_fn),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def on_success(self, summary: ExamArrangeSummary) -> None:
        self.run_button.setEnabled(True)
        self.result_text.setPlainText(
            "\n".join([
                f"结果文件：{summary.output_path}",
                f"总人数：{summary.total_candidates}",
                f"成功编排：{summary.arranged_candidates}",
                f"未找到岗位归组：{summary.missing_groups}",
                f"未找到编排片段：{summary.missing_plan_groups}",
                f"重复岗位归组：{summary.duplicate_group_rows}",
                f"剩余空座：{summary.unused_plan_slots}",
            ])
        )

    def on_error(self, error_text: str) -> None:
        self.run_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "编排失败", error_text.splitlines()[0])
