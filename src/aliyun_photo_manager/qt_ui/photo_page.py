from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
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

from ..app import RunOptions, WorkflowSummary, run_photo_classification_only, run_photo_download_and_template
from ..config import OssConfig, validate_oss_config


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class PhotoWorker(QRunnable):
    def __init__(self, task: Callable[[], object], on_success: Callable[[object], None], on_error: Callable[[str], None]) -> None:
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


class PhotoPage(QWidget):
    LABEL_WIDTH = 132
    ACTION_WIDTH = 144

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self._active_action = "download"
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._create_card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)
        title = QLabel("照片下载与分类")
        title.setProperty("heroTitle", True)
        intro = QLabel("支持本地目录生成模板和按模板分类，也支持云存储下载后生成模板。")
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
        left_title = QLabel("执行参数")
        left_title.setProperty("sectionTitle", True)
        left_layout.addWidget(left_title)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.source_mode_combo = QComboBox()
        self.source_mode_combo.addItems(["本地目录", "云存储下载后处理"])
        self.source_mode_combo.currentIndexChanged.connect(self.update_source_mode_state)
        self._add_row(form, 0, "数据来源", self.source_mode_combo)

        self.download_dir_edit = QLineEdit()
        self._add_row(form, 1, "下载目录", self._with_dir_button(self.download_dir_edit))

        self.sorted_dir_edit = QLineEdit()
        self._add_row(form, 2, "分类目录", self._with_dir_button(self.sorted_dir_edit))
        left_layout.addLayout(form)

        self.cloud_section = QWidget()
        cloud_layout = QVBoxLayout(self.cloud_section)
        cloud_layout.setContentsMargins(0, 0, 0, 0)
        cloud_layout.setSpacing(12)
        cloud_title = QLabel("云存储配置")
        cloud_title.setProperty("sectionTitle", True)
        cloud_layout.addWidget(cloud_title)

        cloud_form = QGridLayout()
        cloud_form.setHorizontalSpacing(14)
        cloud_form.setVerticalSpacing(14)

        self.cloud_type_combo = QComboBox()
        self.cloud_type_combo.addItems(["aliyun", "tencent"])
        self._add_row(cloud_form, 0, "云类型", self.cloud_type_combo)

        self.endpoint_edit = QLineEdit()
        self._add_row(cloud_form, 1, "Endpoint / Region", self.endpoint_edit)

        self.access_key_id_edit = QLineEdit()
        self._add_row(cloud_form, 2, "AccessKey ID", self.access_key_id_edit)

        self.access_key_secret_edit = QLineEdit()
        self.access_key_secret_edit.setEchoMode(QLineEdit.Password)
        self._add_row(cloud_form, 3, "AccessKey Secret", self.access_key_secret_edit)

        self.bucket_edit = QLineEdit()
        self._add_row(cloud_form, 4, "Bucket", self.bucket_edit)

        self.prefix_edit = QLineEdit()
        self._add_row(cloud_form, 5, "前缀", self.prefix_edit)

        cloud_layout.addLayout(cloud_form)
        left_layout.addWidget(self.cloud_section)

        self.dry_run_checkbox = QCheckBox("仅预览，不实际执行")
        left_layout.addWidget(self.dry_run_checkbox)
        self.skip_existing_checkbox = QCheckBox("下载时跳过已存在文件")
        self.skip_existing_checkbox.setChecked(True)
        left_layout.addWidget(self.skip_existing_checkbox)
        left_layout.addStretch(1)

        actions = QHBoxLayout()
        self.generate_button = QPushButton("下载并生成模板")
        self.generate_button.setProperty("accent", True)
        self.generate_button.setFixedWidth(156)
        self.generate_button.clicked.connect(self.start_generate)
        actions.addWidget(self.generate_button)
        self.classify_button = QPushButton("按模板分类")
        self.classify_button.setFixedWidth(132)
        self.classify_button.clicked.connect(self.start_classify)
        actions.addWidget(self.classify_button)
        actions.addStretch(1)
        left_layout.addLayout(actions)
        body.addWidget(left, 0, 0)

        right = self._create_card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_layout.setSpacing(14)
        right_title = QLabel("执行结果")
        right_title.setProperty("sectionTitle", True)
        right_layout.addWidget(right_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        right_layout.addWidget(self.result_text, 1)
        body.addWidget(right, 0, 1)

        body.setColumnStretch(0, 3)
        body.setColumnStretch(1, 2)
        self.update_source_mode_state()

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

    def _with_dir_button(self, line_edit: QLineEdit) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton("选择目录")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(lambda: self.choose_directory(line_edit))
        row_layout.addWidget(button)
        return row

    def choose_directory(self, line_edit: QLineEdit) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择目录", line_edit.text().strip() or str(Path.cwd()))
        if selected:
            line_edit.setText(selected)

    def update_source_mode_state(self) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        self.cloud_section.setVisible(cloud_mode)
        for widget in (
            self.cloud_type_combo,
            self.endpoint_edit,
            self.access_key_id_edit,
            self.access_key_secret_edit,
            self.bucket_edit,
            self.prefix_edit,
            self.skip_existing_checkbox,
        ):
            widget.setVisible(cloud_mode)
            widget.setEnabled(cloud_mode)

    def build_options(self) -> RunOptions:
        download_dir = self.download_dir_edit.text().strip()
        sorted_dir = self.sorted_dir_edit.text().strip()
        if not download_dir or not sorted_dir:
            raise ValueError("请先选择下载目录和分类目录。")
        return RunOptions(
            download_dir=Path(download_dir),
            sorted_dir=Path(sorted_dir),
            prefix=self.prefix_edit.text().strip(),
            skip_download=self.source_mode_combo.currentIndex() == 0,
            dry_run=self.dry_run_checkbox.isChecked(),
            skip_existing=self.skip_existing_checkbox.isChecked(),
        )

    def build_cloud_config(self) -> OssConfig:
        return validate_oss_config(
            OssConfig(
                cloud_type=self.cloud_type_combo.currentText().strip(),
                access_key_id=self.access_key_id_edit.text().strip(),
                access_key_secret=self.access_key_secret_edit.text().strip(),
                endpoint=self.endpoint_edit.text().strip(),
                bucket_name=self.bucket_edit.text().strip(),
            )
        )

    def start_generate(self) -> None:
        try:
            options = self.build_options()
            oss_config = None if options.skip_download else self.build_cloud_config()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return
        self._active_action = "download"
        self._set_running(True)
        self.result_text.setPlainText("正在处理，请稍候...")
        worker = PhotoWorker(
            task=lambda: run_photo_download_and_template(options, oss_config=oss_config, logger=self.log_fn),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def start_classify(self) -> None:
        try:
            options = self.build_options()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return
        self._active_action = "classify"
        self._set_running(True)
        self.result_text.setPlainText("正在按模板分类，请稍候...")
        worker = PhotoWorker(
            task=lambda: run_photo_classification_only(options, logger=self.log_fn),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def _set_running(self, running: bool) -> None:
        self.generate_button.setEnabled(not running)
        self.classify_button.setEnabled(not running)

    def on_success(self, summary: WorkflowSummary) -> None:
        self._set_running(False)
        lines = [
            f"下载目录：{summary.download_dir}",
            f"分类目录：{summary.sorted_dir}",
            f"模板：{summary.template_path}",
        ]
        if summary.download_result is not None:
            lines.extend(
                [
                    f"找到文件：{summary.download_result.total_found}",
                    f"已下载：{summary.download_result.downloaded_count}",
                    f"跳过已存在：{summary.download_result.skipped_existing_count}",
                ]
            )
        if self._active_action == "download":
            lines.append(f"模板文件数：{summary.template_file_count}")
        else:
            lines.append(f"分类完成：{summary.classified_count}")
            if summary.report_path is not None:
                lines.append(f"结果清单：{summary.report_path}")
        self.result_text.setPlainText("\n".join(lines))

    def on_error(self, error_text: str) -> None:
        self._set_running(False)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])
