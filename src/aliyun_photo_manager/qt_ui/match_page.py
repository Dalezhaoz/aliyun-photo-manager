from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
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


class MatchPage(QWidget):
    LABEL_WIDTH = 132
    ACTION_WIDTH = 144

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.target_headers: list[str] = []
        self.source_headers: list[str] = []
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        title = QLabel("数据匹配")
        title.setProperty("heroTitle", True)
        intro = QLabel("按主键和附加匹配列把来源表字段补回目标表，输出带匹配结果清单的新文件。")
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

        self.target_edit = QLineEdit()
        self._add_row(form, 0, "目标表", self._with_file_button(self.target_edit))
        self.source_edit = QLineEdit()
        self._add_row(form, 1, "来源表", self._with_file_button(self.source_edit))
        self.output_edit = QLineEdit()
        self._add_row(form, 2, "输出文件", self._with_file_button(self.output_edit, save_mode=True))

        header_action = QWidget()
        header_layout = QHBoxLayout(header_action)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(10)
        load_button = QPushButton("加载表头")
        load_button.setFixedWidth(self.ACTION_WIDTH)
        load_button.clicked.connect(self.load_headers)
        header_layout.addWidget(load_button)
        header_layout.addStretch(1)
        self._add_row(form, 3, "", header_action)

        self.target_key_edit = QLineEdit()
        self._add_row(form, 4, "目标表匹配列", self.target_key_edit)
        self.source_key_edit = QLineEdit()
        self._add_row(form, 5, "来源表匹配列", self.source_key_edit)
        left_layout.addLayout(form)

        extra_title = QLabel("附加匹配列")
        extra_title.setProperty("sectionTitle", True)
        left_layout.addWidget(extra_title)
        self.extra_table = self._mapping_table("目标表列", "来源表列")
        left_layout.addWidget(self.extra_table)

        extra_actions = QHBoxLayout()
        add_extra = QPushButton("添加映射")
        add_extra.clicked.connect(lambda: self.extra_table.insertRow(self.extra_table.rowCount()))
        del_extra = QPushButton("删除所选")
        del_extra.clicked.connect(lambda: self._remove_selected_row(self.extra_table))
        extra_actions.addWidget(add_extra)
        extra_actions.addWidget(del_extra)
        extra_actions.addStretch(1)
        left_layout.addLayout(extra_actions)

        transfer_title = QLabel("补充列映射")
        transfer_title.setProperty("sectionTitle", True)
        left_layout.addWidget(transfer_title)
        self.transfer_table = self._mapping_table("结果列名", "来源表列")
        left_layout.addWidget(self.transfer_table)
        transfer_actions = QHBoxLayout()
        add_transfer = QPushButton("添加补充列")
        add_transfer.clicked.connect(lambda: self.transfer_table.insertRow(self.transfer_table.rowCount()))
        del_transfer = QPushButton("删除所选")
        del_transfer.clicked.connect(lambda: self._remove_selected_row(self.transfer_table))
        transfer_actions.addWidget(add_transfer)
        transfer_actions.addWidget(del_transfer)
        transfer_actions.addStretch(1)
        left_layout.addLayout(transfer_actions)

        run_row = QHBoxLayout()
        self.run_button = QPushButton("开始匹配")
        self.run_button.setProperty("accent", True)
        self.run_button.setFixedWidth(132)
        self.run_button.clicked.connect(self.start_run)
        run_row.addWidget(self.run_button)
        run_row.addStretch(1)
        left_layout.addLayout(run_row)
        body.addWidget(left, 0, 0)

        right = self._card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_title = QLabel("匹配结果")
        right_title.setProperty("sectionTitle", True)
        right_layout.addWidget(right_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        right_layout.addWidget(self.result_text, 1)
        body.addWidget(right, 0, 1)

        body.setColumnStretch(0, 3)
        body.setColumnStretch(1, 2)

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

    def _mapping_table(self, first_header: str, second_header: str) -> QTableWidget:
        table = QTableWidget(0, 2)
        table.setHorizontalHeaderLabels([first_header, second_header])
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
        if self.target_headers and not self.target_key_edit.text().strip():
            self.target_key_edit.setText(self.target_headers[0])
        if self.source_headers and not self.source_key_edit.text().strip():
            self.source_key_edit.setText(self.source_headers[0])
        self.result_text.setPlainText(
            "目标表列：\n" + "\n".join(self.target_headers) + "\n\n来源表列：\n" + "\n".join(self.source_headers)
        )

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
                target_key_column=self.target_key_edit.text().strip(),
                source_key_column=self.source_key_edit.text().strip(),
                extra_match_mappings=self._collect_mappings(self.extra_table),
                transfer_mappings=self._collect_mappings(self.transfer_table),
                output_path=Path(self.output_edit.text().strip()) if self.output_edit.text().strip() else None,
            )
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
