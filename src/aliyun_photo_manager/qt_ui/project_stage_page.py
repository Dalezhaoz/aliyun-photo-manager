from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import traceback
from dataclasses import asdict
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QCheckBox,
    QComboBox,
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

from ..project_stage_report import (
    ProjectStageSummary,
    StageServerConfig,
    dump_status_query_payload,
    export_project_stages,
    query_project_stages,
    summary_from_dict,
)


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class SimpleWorker(QRunnable):
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


class ProjectStagePage(QWidget):
    LABEL_WIDTH = 132

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.server_configs: list[StageServerConfig] = []
        self.selected_index: int | None = None
        self.last_summary: ProjectStageSummary | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        title = QLabel("项目阶段汇总")
        title.setProperty("heroTitle", True)
        intro = QLabel("一次查看多台 SQL Server 上报名项目的阶段状态，支持服务器配置和结果导出。")
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
        self.server_name_edit = QLineEdit()
        self._add_row(form, 0, "服务器名称", self.server_name_edit)
        self.server_host_edit = QLineEdit()
        self._add_row(form, 1, "数据库地址", self.server_host_edit)
        self.server_port_edit = QLineEdit("1433")
        self._add_row(form, 2, "端口", self.server_port_edit)
        self.server_user_edit = QLineEdit()
        self._add_row(form, 3, "用户名", self.server_user_edit)
        self.server_password_edit = QLineEdit()
        self.server_password_edit.setEchoMode(QLineEdit.Password)
        self._add_row(form, 4, "密码", self.server_password_edit)
        self.server_enabled_checkbox = QCheckBox("启用此服务器")
        self.server_enabled_checkbox.setChecked(True)
        form.addWidget(self.server_enabled_checkbox, 5, 1)
        left_layout.addLayout(form)

        server_actions = QHBoxLayout()
        save_button = QPushButton("新增/更新")
        save_button.clicked.connect(self.save_server)
        clear_button = QPushButton("清空表单")
        clear_button.clicked.connect(self.clear_server_form)
        delete_button = QPushButton("删除所选")
        delete_button.clicked.connect(self.delete_server)
        test_button = QPushButton("测试连接")
        test_button.clicked.connect(self.test_server)
        server_actions.addWidget(save_button)
        server_actions.addWidget(clear_button)
        server_actions.addWidget(delete_button)
        server_actions.addWidget(test_button)
        server_actions.addStretch(1)
        left_layout.addLayout(server_actions)

        self.server_table = QTableWidget(0, 3)
        self.server_table.setHorizontalHeaderLabels(["服务器名称", "地址", "启用"])
        self.server_table.horizontalHeader().setStretchLastSection(True)
        self.server_table.verticalHeader().setVisible(False)
        self.server_table.setMinimumHeight(140)
        self.server_table.itemSelectionChanged.connect(self.on_server_selected)
        left_layout.addWidget(self.server_table)

        filters = QGridLayout()
        filters.setHorizontalSpacing(14)
        filters.setVerticalSpacing(14)
        self.status_filter_combo = QComboBox()
        self.status_filter_combo.addItems(["正在进行 + 即将开始", "全部", "只看正在进行", "只看即将开始"])
        self._add_row(filters, 0, "状态", self.status_filter_combo)
        self.stage_keyword_edit = QLineEdit()
        self._add_row(filters, 1, "阶段关键字", self.stage_keyword_edit)
        self.project_keyword_edit = QLineEdit()
        self._add_row(filters, 2, "项目关键字", self.project_keyword_edit)
        left_layout.addLayout(filters)

        query_actions = QHBoxLayout()
        self.run_button = QPushButton("开始查询")
        self.run_button.setProperty("accent", True)
        self.run_button.clicked.connect(self.start_query)
        export_button = QPushButton("导出 Excel")
        export_button.clicked.connect(self.export_result)
        query_actions.addWidget(self.run_button)
        query_actions.addWidget(export_button)
        query_actions.addStretch(1)
        left_layout.addLayout(query_actions)
        body.addWidget(left, 0, 0)

        right = self._card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        title = QLabel("查询结果")
        title.setProperty("sectionTitle", True)
        right_layout.addWidget(title)
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
        label = QLabel(label_text)
        label.setProperty("formLabel", True)
        label.setFixedWidth(self.LABEL_WIDTH)
        layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

    def refresh_server_table(self) -> None:
        self.server_table.setRowCount(0)
        for index, item in enumerate(self.server_configs):
            self.server_table.insertRow(index)
            self.server_table.setItem(index, 0, QTableWidgetItem(item.name))
            self.server_table.setItem(index, 1, QTableWidgetItem(f"{item.host}:{item.port}"))
            self.server_table.setItem(index, 2, QTableWidgetItem("是" if item.enabled else "否"))

    def clear_server_form(self) -> None:
        self.selected_index = None
        self.server_name_edit.clear()
        self.server_host_edit.clear()
        self.server_port_edit.setText("1433")
        self.server_user_edit.clear()
        self.server_password_edit.clear()
        self.server_enabled_checkbox.setChecked(True)

    def on_server_selected(self) -> None:
        row = self.server_table.currentRow()
        if row < 0 or row >= len(self.server_configs):
            return
        self.selected_index = row
        item = self.server_configs[row]
        self.server_name_edit.setText(item.name)
        self.server_host_edit.setText(item.host)
        self.server_port_edit.setText(str(item.port))
        self.server_user_edit.setText(item.username)
        self.server_password_edit.setText(item.password)
        self.server_enabled_checkbox.setChecked(item.enabled)

    def save_server(self) -> None:
        try:
            config = StageServerConfig(
                name=self.server_name_edit.text().strip(),
                host=self.server_host_edit.text().strip(),
                port=int(self.server_port_edit.text().strip() or "1433"),
                username=self.server_user_edit.text().strip(),
                password=self.server_password_edit.text().strip(),
                enabled=self.server_enabled_checkbox.isChecked(),
            )
            if not config.name or not config.host or not config.username or not config.password:
                raise ValueError("请完整填写服务器名称、地址、用户名和密码。")
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return
        if self.selected_index is None:
            self.server_configs.append(config)
        else:
            self.server_configs[self.selected_index] = config
        self.refresh_server_table()
        self.clear_server_form()

    def delete_server(self) -> None:
        row = self.server_table.currentRow()
        if row < 0 or row >= len(self.server_configs):
            return
        del self.server_configs[row]
        self.refresh_server_table()
        self.clear_server_form()

    def test_server(self) -> None:
        try:
            config = StageServerConfig(
                name=self.server_name_edit.text().strip(),
                host=self.server_host_edit.text().strip(),
                port=int(self.server_port_edit.text().strip() or "1433"),
                username=self.server_user_edit.text().strip(),
                password=self.server_password_edit.text().strip(),
                enabled=True,
            )
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        worker = SimpleWorker(
            task=lambda: query_project_stages([config], status_filter="全部"),
            on_success=lambda summary: self.result_text.setPlainText(
                f"连接成功。\n遍历数据库：{summary.visited_databases}\n匹配业务库：{summary.matched_databases}"
            ),
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def start_query(self) -> None:
        if not self.server_configs:
            QMessageBox.critical(self, "参数错误", "请先添加至少一台数据库服务器。")
            return
        self.run_button.setEnabled(False)
        self.result_text.setPlainText("正在查询，请稍候...")

        def task() -> ProjectStageSummary:
            with tempfile.TemporaryDirectory(prefix="project_stage_qt_") as temp_dir:
                temp_path = Path(temp_dir)
                input_path = temp_path / "query.json"
                output_path = temp_path / "result.json"
                dump_status_query_payload(
                    servers=self.server_configs,
                    status_filter=self.status_filter_combo.currentText().strip(),
                    stage_keyword=self.stage_keyword_edit.text().strip(),
                    project_keyword=self.project_keyword_edit.text().strip(),
                    output_path=input_path,
                )
                runner_path = Path(__file__).resolve().parents[1] / "project_stage_runner.py"
                result = subprocess.run(
                    [sys.executable, str(runner_path), str(input_path), str(output_path)],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                if result.returncode != 0:
                    raise RuntimeError((result.stderr or result.stdout or "项目阶段汇总查询失败").strip())
                return summary_from_dict(json.loads(output_path.read_text(encoding="utf-8")))

        worker = SimpleWorker(task=task, on_success=self.on_query_success, on_error=self.on_error)
        self.thread_pool.start(worker)

    def on_query_success(self, summary: ProjectStageSummary) -> None:
        self.run_button.setEnabled(True)
        self.last_summary = summary
        lines = [
            f"启用服务器：{summary.enabled_servers}",
            f"遍历数据库：{summary.visited_databases}",
            f"匹配业务库：{summary.matched_databases}",
            f"正在进行：{summary.ongoing_count}",
            f"即将开始：{summary.upcoming_count}",
            "",
        ]
        for item in summary.records[:50]:
            lines.append(
                f"{item.server_name} / {item.database_name} / {item.project_name} / {item.stage_name} / {item.status}"
            )
        self.result_text.setPlainText("\n".join(lines))

    def export_result(self) -> None:
        if self.last_summary is None or not self.last_summary.records:
            return
        selected, _ = QFileDialog.getSaveFileName(self, "导出项目阶段汇总", "项目阶段汇总.xlsx", "Excel 文件 (*.xlsx)")
        if not selected:
            return
        try:
            output_path = export_project_stages(self.last_summary, Path(selected))
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))
            return
        self.result_text.appendPlainText(f"\n已导出：{output_path}")

    def on_error(self, error_text: str) -> None:
        self.run_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])
