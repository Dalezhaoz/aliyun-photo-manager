from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
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
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..data_matcher import ColumnMapping, DataMatchOptions, DataMatchSummary, list_headers, run_data_match
from .base_page import BaseToolPage
from .common import AppComboBox
from .widgets import FormRow, PathInput, PrimaryButton, SecondaryButton, Section


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class MatchWorker(QRunnable):
    def __init__(self, options: DataMatchOptions, log_fn: Callable[[str], None], on_success, on_error) -> None:
        super().__init__()
        self.options = options
        self.log_fn = log_fn
        self.signals = WorkerSignals()
        self.signals.success.connect(on_success)
        self.signals.error.connect(on_error)

    def run(self) -> None:
        try:
            summary = run_data_match(self.options, logger=self.log_fn)
        except Exception as exc:
            self.signals.error.emit(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}")
        else:
            self.signals.success.emit(summary)


class MatchPage(BaseToolPage):
    LABEL_WIDTH = 118
    ACTION_WIDTH = 116

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="数据匹配",
            description="按主键和附加匹配列把来源表字段补回目标表，输出带匹配结果清单的新文件。",
            show_steps=True,
            show_log=True,
        )
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.target_headers: list[str] = []
        self.source_headers: list[str] = []
        self._build_ui()

    def _build_ui(self) -> None:
        if self.left_card.title is None:
            self.left_card.title = QLabel("匹配配置")
            self.left_card.title.setProperty("sectionTitle", True)
            self.left_card.body_layout.insertWidget(0, self.left_card.title)
        else:
            self.left_card.title.setText("匹配配置")

        file_section = Section("文件设置")
        self.target_input = PathInput("请选择目标表文件", "选择文件")
        self.target_edit = self.target_input.edit
        self.target_input.button.clicked.connect(lambda: self.choose_file(self.target_edit))
        self.source_input = PathInput("请选择来源表文件", "选择文件")
        self.source_edit = self.source_input.edit
        self.source_input.button.clicked.connect(lambda: self.choose_file(self.source_edit))
        self.output_input = PathInput("请选择输出文件路径", "选择文件")
        self.output_edit = self.output_input.edit
        self.output_input.button.clicked.connect(lambda: self.choose_file(self.output_edit, True))
        file_section.add(FormRow("目标表", self.target_input))
        file_section.add(FormRow("来源表", self.source_input))
        file_section.add(FormRow("输出文件", self.output_input))
        load_button = SecondaryButton("加载表头")
        load_button.clicked.connect(self.load_headers)
        file_section.add(load_button)
        self.left_card.body_layout.addWidget(file_section)

        match_section = Section("主键匹配")
        self.target_key_combo = AppComboBox()
        self.source_key_combo = AppComboBox()
        match_section.add(FormRow("目标表匹配列", self.target_key_combo))
        match_section.add(FormRow("来源表匹配列", self.source_key_combo))
        self.left_card.body_layout.addWidget(match_section)

        extra_section = Section("附加匹配列")
        extra_editor = QHBoxLayout()
        self.extra_target_combo = AppComboBox()
        self.extra_target_combo.setPlaceholderText("目标表列")
        self.extra_source_combo = AppComboBox()
        self.extra_source_combo.setPlaceholderText("来源表列")
        add_extra = SecondaryButton("添加映射")
        add_extra.clicked.connect(self.add_extra_mapping)
        extra_editor.addWidget(self.extra_target_combo, 1)
        extra_editor.addWidget(self.extra_source_combo, 1)
        extra_editor.addWidget(add_extra)
        extra_section.body_layout.addLayout(extra_editor)
        self.extra_table = self._mapping_table("目标表列", "来源表列")
        extra_section.add(self.extra_table)

        extra_actions = QHBoxLayout()
        del_extra = SecondaryButton("删除所选")
        del_extra.clicked.connect(lambda: self._remove_selected_row(self.extra_table))
        extra_actions.addWidget(del_extra)
        extra_actions.addStretch(1)
        extra_section.body_layout.addLayout(extra_actions)
        self.left_card.body_layout.addWidget(extra_section)

        transfer_section = Section("补充列映射")
        transfer_editor = QHBoxLayout()
        self.transfer_name_edit = QLineEdit()
        self.transfer_name_edit.setPlaceholderText("结果列名")
        self.transfer_source_combo = AppComboBox()
        self.transfer_source_combo.setPlaceholderText("来源表列")
        add_transfer = SecondaryButton("添加补充列")
        add_transfer.clicked.connect(self.add_transfer_mapping)
        transfer_editor.addWidget(self.transfer_name_edit, 1)
        transfer_editor.addWidget(self.transfer_source_combo, 1)
        transfer_editor.addWidget(add_transfer)
        transfer_section.body_layout.addLayout(transfer_editor)
        self.transfer_table = self._mapping_table("结果列名", "来源表列")
        transfer_section.add(self.transfer_table)
        transfer_actions = QHBoxLayout()
        del_transfer = SecondaryButton("删除所选")
        del_transfer.clicked.connect(lambda: self._remove_selected_row(self.transfer_table))
        transfer_actions.addWidget(del_transfer)
        transfer_actions.addStretch(1)
        transfer_section.body_layout.addLayout(transfer_actions)
        self.left_card.body_layout.addWidget(transfer_section)

        run_row = QHBoxLayout()
        self.run_button = PrimaryButton("开始匹配")
        self.run_button.clicked.connect(self.start_run)
        run_row.addWidget(self.run_button)
        run_row.addStretch(1)
        self.left_card.body_layout.addLayout(run_row)

        if self.right_card.title is None:
            self.right_card.title = QLabel("匹配结果")
            self.right_card.title.setProperty("sectionTitle", True)
            self.right_card.body_layout.insertWidget(0, self.right_card.title)
        else:
            self.right_card.title.setText("匹配结果")
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.right_card.body_layout.addWidget(self.result_text, 1)

    def _card(self) -> QFrame:
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

    def _with_file_button(self, line_edit: QLineEdit, save_mode: bool = False) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton("选择文件")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(lambda: self.choose_file(line_edit, save_mode))
        row_layout.addWidget(button)
        return row

    def choose_file(self, line_edit: QLineEdit, save_mode: bool = False) -> None:
        if save_mode:
            selected, _ = QFileDialog.getSaveFileName(self, "选择输出文件", "", "Excel 文件 (*.xlsx)")
        else:
            selected, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if selected:
            line_edit.setText(selected)
            if not save_mode and line_edit is self.target_edit and not self.output_edit.text().strip():
                target_path = Path(selected)
                self.output_edit.setText(str(target_path.with_name(f"{target_path.stem}_数据匹配结果.xlsx")))

    def _mapping_table(self, first_header: str, second_header: str) -> QTableWidget:
        table = QTableWidget(0, 2)
        table.setHorizontalHeaderLabels([first_header, second_header])
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.horizontalHeader().setStretchLastSection(True)
        table.verticalHeader().setVisible(False)
        table.setMinimumHeight(140)
        return table

    def _remove_selected_row(self, table: QTableWidget) -> None:
        row = table.currentRow()
        if row >= 0:
            table.removeRow(row)

    def load_headers(self) -> None:
        try:
            target_path = Path(self.target_edit.text().strip())
            source_path = Path(self.source_edit.text().strip())
            self.target_headers = list_headers(target_path)
            self.source_headers = list_headers(source_path)
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            return
        self._reset_combo(self.target_key_combo, self.target_headers)
        self._reset_combo(self.source_key_combo, self.source_headers)
        self._reset_combo(self.extra_target_combo, self.target_headers, placeholder="目标表列")
        self._reset_combo(self.extra_source_combo, self.source_headers, placeholder="来源表列")
        self._reset_combo(self.transfer_source_combo, self.source_headers, placeholder="来源表列")
        self.result_text.setPlainText(
            "目标表列：\n" + "\n".join(self.target_headers) + "\n\n来源表列：\n" + "\n".join(self.source_headers)
        )

    def _reset_combo(self, combo: QComboBox, values: list[str], placeholder: str = "") -> None:
        current = combo.currentText().strip()
        combo.clear()
        if placeholder:
            combo.addItem(placeholder, "")
        for value in values:
            combo.addItem(value, value)
        if current and current in values:
            combo.setCurrentText(current)
        elif values:
            combo.setCurrentIndex(1 if placeholder else 0)

    def _combo_value(self, combo: QComboBox) -> str:
        data = combo.currentData()
        if isinstance(data, str) and data:
            return data
        return combo.currentText().strip()

    def _append_mapping_row(self, table: QTableWidget, first: str, second: str) -> None:
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(first))
        table.setItem(row, 1, QTableWidgetItem(second))

    def add_extra_mapping(self) -> None:
        target = self._combo_value(self.extra_target_combo)
        source = self._combo_value(self.extra_source_combo)
        if not target or not source:
            QMessageBox.warning(self, "参数不完整", "请先选择目标表列和来源表列。")
            return
        self._append_mapping_row(self.extra_table, target, source)

    def add_transfer_mapping(self) -> None:
        result_name = self.transfer_name_edit.text().strip()
        source = self._combo_value(self.transfer_source_combo)
        if not result_name or not source:
            QMessageBox.warning(self, "参数不完整", "请先填写结果列名并选择来源表列。")
            return
        self._append_mapping_row(self.transfer_table, result_name, source)
        self.transfer_name_edit.clear()

    def _collect_mappings(self, table: QTableWidget) -> list[ColumnMapping]:
        mappings: list[ColumnMapping] = []
        for row in range(table.rowCount()):
            first_item = table.item(row, 0)
            second_item = table.item(row, 1)
            first = first_item.text().strip() if first_item else ""
            second = second_item.text().strip() if second_item else ""
            if first and second:
                mappings.append(ColumnMapping(first, second))
        return mappings

    def start_run(self) -> None:
        try:
            options = DataMatchOptions(
                target_path=Path(self.target_edit.text().strip()),
                source_path=Path(self.source_edit.text().strip()),
                target_key_column=self._combo_value(self.target_key_combo),
                source_key_column=self._combo_value(self.source_key_combo),
                extra_match_mappings=self._collect_mappings(self.extra_table),
                transfer_mappings=self._collect_mappings(self.transfer_table),
                output_path=Path(self.output_edit.text().strip()) if self.output_edit.text().strip() else None,
            )
            if not options.transfer_mappings:
                raise ValueError("请至少添加一条补充列映射。")
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return
        self.run_button.setEnabled(False)
        self.result_text.setPlainText("正在匹配，请稍候...")
        worker = MatchWorker(options, self.log_fn, self.on_success, self.on_error)
        self.thread_pool.start(worker)

    def on_success(self, summary: DataMatchSummary) -> None:
        self.run_button.setEnabled(True)
        lines = [
            f"结果文件：{summary.output_path}",
            f"总行数：{summary.total_rows}",
            f"已匹配：{summary.matched_rows}",
            f"未匹配：{summary.unmatched_rows}",
            f"来源重复键：{summary.duplicate_source_keys}",
            f"歧义行：{summary.ambiguous_rows}",
        ]
        self.result_text.setPlainText("\n".join(lines))

    def on_error(self, error_text: str) -> None:
        self.run_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "匹配失败", error_text.splitlines()[0])
