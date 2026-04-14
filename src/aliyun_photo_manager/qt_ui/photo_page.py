from __future__ import annotations

import traceback
import json
from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt, QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..app import (
    RunOptions,
    WorkflowSummary,
    run_photo_classification_only,
    run_photo_download_and_template,
    run_photo_template_only,
)
from ..certificate_filter import list_template_headers
from ..config import OssConfig, validate_oss_config
from ..downloader import BrowserEntry, list_browser_entries, list_buckets
from .common import AppComboBox


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
    LABEL_WIDTH = 118
    ACTION_WIDTH = 116
    SETTINGS_FILE = Path(__file__).resolve().parents[3] / ".gui_settings.json"

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.photo_filter_headers: list[str] = []
        self.browser_entries: list[BrowserEntry] = []
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
        intro = QLabel("支持云存储下载照片、按表过滤下载，以及根据本地目录生成模板后再按模板分类。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        body = self._create_card()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(24, 22, 24, 24)
        body_layout.setSpacing(18)
        root.addWidget(body, 1)

        title_label = QLabel("执行参数")
        title_label.setProperty("sectionTitle", True)
        body_layout.addWidget(title_label)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.source_mode_combo = AppComboBox()
        self.source_mode_combo.addItems(["本地目录", "云存储下载后处理"])
        self.source_mode_combo.currentIndexChanged.connect(self.update_source_mode_state)
        self._add_row(form, 0, "数据来源", self.source_mode_combo)

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

        self.cloud_type_combo = AppComboBox()
        self.cloud_type_combo.addItems(["aliyun", "tencent"])
        self._add_row(cloud_form, 0, "云类型", self.cloud_type_combo)

        self.endpoint_edit = QLineEdit()
        self._add_row(cloud_form, 1, "Endpoint / Region", self.endpoint_edit)

        self.access_key_id_edit = QLineEdit()
        self._add_row(cloud_form, 2, "AccessKey ID", self.access_key_id_edit)

        self.access_key_secret_edit = QLineEdit()
        self.access_key_secret_edit.setEchoMode(QLineEdit.Password)
        self._add_row(cloud_form, 3, "AccessKey Secret", self.access_key_secret_edit)

        bucket_row = QWidget()
        bucket_layout = QHBoxLayout(bucket_row)
        bucket_layout.setContentsMargins(0, 0, 0, 0)
        bucket_layout.setSpacing(10)
        self.bucket_combo = AppComboBox()
        self.bucket_combo.setEditable(False)
        bucket_layout.addWidget(self.bucket_combo, 1)
        self.load_bucket_button = QPushButton("加载 Bucket")
        self.load_bucket_button.setFixedWidth(self.ACTION_WIDTH)
        self.load_bucket_button.clicked.connect(self.load_buckets)
        bucket_layout.addWidget(self.load_bucket_button)
        self._add_row(cloud_form, 4, "Bucket", bucket_row)

        self.prefix_edit = QLineEdit()
        self.prefix_edit.setPlaceholderText("当前前缀，留空表示根目录")
        self._add_row(cloud_form, 5, "当前前缀", self.prefix_edit)

        browser_controls = QWidget()
        browser_controls_layout = QHBoxLayout(browser_controls)
        browser_controls_layout.setContentsMargins(0, 0, 0, 0)
        browser_controls_layout.setSpacing(10)
        self.load_prefix_button = QPushButton("加载当前层级")
        self.load_prefix_button.clicked.connect(self.load_browser_entries)
        browser_controls_layout.addWidget(self.load_prefix_button)
        self.parent_prefix_button = QPushButton("返回上一级")
        self.parent_prefix_button.clicked.connect(self.go_to_parent_prefix)
        browser_controls_layout.addWidget(self.parent_prefix_button)
        browser_controls_layout.addStretch(1)
        self._add_row(cloud_form, 6, "云端目录", browser_controls)

        self.browser_list = QListWidget()
        self.browser_list.setMinimumHeight(150)
        self.browser_list.itemClicked.connect(self.on_browser_item_clicked)
        self.browser_list.itemDoubleClicked.connect(self.on_browser_item_double_clicked)
        self._add_row(cloud_form, 7, "", self.browser_list)

        self.browser_status_label = QLabel("请先选择 Bucket，再加载当前层级。")
        self.browser_status_label.setWordWrap(True)
        self._add_row(cloud_form, 8, "", self.browser_status_label)

        cloud_layout.addLayout(cloud_form)
        form.addWidget(self.cloud_section, 1, 0, 1, 2)

        self.download_dir_edit = QLineEdit()
        self._add_row(form, 2, "下载目录", self._with_dir_button(self.download_dir_edit))

        self.sorted_dir_edit = QLineEdit()
        self._add_row(form, 3, "分类目录", self._with_dir_button(self.sorted_dir_edit))

        self.filter_download_checkbox = QCheckBox("云下载时按表过滤")
        self.filter_download_checkbox.toggled.connect(self.update_filter_state)
        form.addWidget(self.filter_download_checkbox, 4, 1)

        self.filter_template_edit = QLineEdit()
        filter_template_row = self._with_file_button(self.filter_template_edit, self.choose_filter_template)
        self._add_row(form, 5, "过滤表", filter_template_row)

        self.filter_column_combo = AppComboBox()
        self.filter_column_combo.setEditable(False)
        load_filter_headers_btn = QPushButton("加载列")
        load_filter_headers_btn.clicked.connect(self.load_filter_headers)
        filter_column_row = QWidget()
        filter_column_layout = QHBoxLayout(filter_column_row)
        filter_column_layout.setContentsMargins(0, 0, 0, 0)
        filter_column_layout.setSpacing(10)
        filter_column_layout.addWidget(self.filter_column_combo, 1)
        filter_column_layout.addWidget(load_filter_headers_btn)
        self._add_row(form, 6, "前缀列", filter_column_row)

        body_layout.addLayout(form)

        self.dry_run_checkbox = QCheckBox("仅预览，不实际执行")
        body_layout.addWidget(self.dry_run_checkbox)

        self.skip_existing_checkbox = QCheckBox("下载时跳过已存在文件")
        self.skip_existing_checkbox.setChecked(True)
        body_layout.addWidget(self.skip_existing_checkbox)

        actions = QHBoxLayout()
        actions.setContentsMargins(0, 8, 0, 0)
        actions.setSpacing(10)

        self.download_button = QPushButton("下载照片")
        self.download_button.setProperty("accent", True)
        self.download_button.setFixedWidth(132)
        self.download_button.clicked.connect(self.start_download)
        actions.addWidget(self.download_button)

        self.generate_button = QPushButton("生成模板")
        self.generate_button.setFixedWidth(132)
        self.generate_button.clicked.connect(self.start_generate_template)
        actions.addWidget(self.generate_button)

        self.classify_button = QPushButton("按模板分类")
        self.classify_button.setFixedWidth(132)
        self.classify_button.clicked.connect(self.start_classify)
        actions.addWidget(self.classify_button)

        actions.addStretch(1)
        body_layout.addLayout(actions)

        result_title = QLabel("执行结果")
        result_title.setProperty("sectionTitle", True)
        body_layout.addWidget(result_title)

        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(160)
        self.result_text.setPlaceholderText("这里会显示下载、模板生成和分类结果。")
        body_layout.addWidget(self.result_text)
        body_layout.addStretch(1)

        self.update_source_mode_state()
        self.update_filter_state()
        self._bind_cloud_cache_events()
        self._load_cached_cloud_settings()
        self._show_browser_placeholder()

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

    def _with_file_button(self, line_edit: QLineEdit, callback) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton("选择文件")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(callback)
        row_layout.addWidget(button)
        return row

    def choose_directory(self, line_edit: QLineEdit) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择目录", line_edit.text().strip() or str(Path.cwd()))
        if selected:
            line_edit.setText(selected)

    def choose_filter_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择过滤表", "", "Excel 文件 (*.xlsx)")
        if selected:
            self.filter_template_edit.setText(selected)
            self.load_filter_headers()

    def load_filter_headers(self) -> None:
        template_path = self.filter_template_edit.text().strip()
        if not template_path:
            QMessageBox.information(self, "缺少过滤表", "请先选择过滤表文件。")
            return
        try:
            headers = list_template_headers(Path(template_path))
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            return
        self.photo_filter_headers = headers
        self.filter_column_combo.clear()
        self.filter_column_combo.addItems(headers)
        self.log_fn(f"已读取照片过滤表列：{', '.join(headers) if headers else '无'}")

    def update_source_mode_state(self) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        self.cloud_section.setVisible(cloud_mode)
        for widget in (
            self.cloud_type_combo,
            self.endpoint_edit,
            self.access_key_id_edit,
            self.access_key_secret_edit,
            self.bucket_combo,
            self.prefix_edit,
            self.skip_existing_checkbox,
            self.filter_download_checkbox,
        ):
            widget.setVisible(cloud_mode)
            widget.setEnabled(cloud_mode)
        self.download_button.setVisible(cloud_mode)
        self.download_button.setEnabled(True)
        self.update_filter_state()

    def update_filter_state(self) -> None:
        enabled = self.source_mode_combo.currentIndex() == 1 and self.filter_download_checkbox.isChecked()
        self.filter_template_edit.setEnabled(enabled)
        self.filter_column_combo.setEnabled(enabled)

    def _show_browser_placeholder(self, text: str = "请先选择 Bucket，再加载当前层级。") -> None:
        self.browser_entries = []
        self.browser_list.clear()
        item = QListWidgetItem(text)
        item.setFlags(Qt.NoItemFlags)
        self.browser_list.addItem(item)
        self.browser_status_label.setText(text)

    def build_options(self) -> RunOptions:
        download_dir = self.download_dir_edit.text().strip()
        sorted_dir = self.sorted_dir_edit.text().strip()
        if not download_dir:
            raise ValueError("请先选择下载目录。")
        if not sorted_dir:
            raise ValueError("请先选择分类目录。")

        filter_template_path = None
        filter_column = ""
        filter_download = self.source_mode_combo.currentIndex() == 1 and self.filter_download_checkbox.isChecked()
        if filter_download:
            filter_template_text = self.filter_template_edit.text().strip()
            if not filter_template_text:
                raise ValueError("已启用按表过滤下载，请先选择过滤表。")
            filter_template_path = Path(filter_template_text)
            if not filter_template_path.exists():
                raise ValueError("过滤表文件不存在。")
            filter_column = self.filter_column_combo.currentText().strip()
            if not filter_column:
                raise ValueError("已启用按表过滤下载，请先选择文件名前缀列。")

        return RunOptions(
            download_dir=Path(download_dir),
            sorted_dir=Path(sorted_dir),
            prefix=self.prefix_edit.text().strip(),
            skip_download=self.source_mode_combo.currentIndex() == 0,
            dry_run=self.dry_run_checkbox.isChecked(),
            skip_existing=self.skip_existing_checkbox.isChecked(),
            filter_download=filter_download,
            filter_template_path=filter_template_path,
            filter_column=filter_column,
        )

    def build_cloud_config(self) -> OssConfig:
        return validate_oss_config(
            OssConfig(
                cloud_type=self.cloud_type_combo.currentText().strip(),
                access_key_id=self.access_key_id_edit.text().strip(),
                access_key_secret=self.access_key_secret_edit.text().strip(),
                endpoint=self.endpoint_edit.text().strip(),
                bucket_name=self.bucket_combo.currentText().strip(),
            )
        )

    def load_buckets(self) -> None:
        try:
            cloud_type = self.cloud_type_combo.currentText().strip()
            access_key_id = self.access_key_id_edit.text().strip()
            access_key_secret = self.access_key_secret_edit.text().strip()
            endpoint = self.endpoint_edit.text().strip()
            if not access_key_id or not access_key_secret or not endpoint:
                raise ValueError("请先填写云类型、Endpoint / Region、AccessKey。")
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self.load_bucket_button.setEnabled(False)
        self.result_text.setPlainText("正在加载 Bucket 列表...")
        worker = PhotoWorker(
            task=lambda: list_buckets(
                access_key_id=access_key_id,
                access_key_secret=access_key_secret,
                endpoint=endpoint,
                cloud_type=cloud_type,
            ),
            on_success=self.on_buckets_loaded,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def on_buckets_loaded(self, buckets: list[str]) -> None:
        self.load_bucket_button.setEnabled(True)
        current = self.bucket_combo.currentText().strip()
        self.bucket_combo.clear()
        self.bucket_combo.addItems(buckets)
        if current:
            index = self.bucket_combo.findText(current)
            if index >= 0:
                self.bucket_combo.setCurrentIndex(index)
        self._save_cached_cloud_settings()
        self.result_text.setPlainText(f"已加载 {len(buckets)} 个 Bucket。")

        self._show_browser_placeholder("Bucket 已加载，请点“加载当前层级”选择目录。")

    def load_browser_entries(self) -> None:
        try:
            config = self.build_cloud_config()
            prefix = self.prefix_edit.text().strip()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self.load_prefix_button.setEnabled(False)
        self.browser_status_label.setText("正在加载当前层级...")
        worker = PhotoWorker(
            task=lambda: list_browser_entries(config, prefix),
            on_success=self.on_browser_entries_loaded,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def on_browser_entries_loaded(self, entries: list[BrowserEntry]) -> None:
        self.load_prefix_button.setEnabled(True)
        self.browser_entries = entries
        self.browser_list.clear()
        if not entries:
            self._show_browser_placeholder("当前层级没有子文件夹或文件。")
            return
        for entry in entries:
            marker = "[目录]" if entry.entry_type == "folder" else "[文件]"
            item = QListWidgetItem(f"{marker} {entry.display_name}")
            item.setData(Qt.UserRole, entry.key)
            item.setData(Qt.UserRole + 1, entry.entry_type)
            self.browser_list.addItem(item)
        folder_count = len([entry for entry in entries if entry.entry_type == "folder"])
        file_count = len(entries) - folder_count
        current_prefix = self.prefix_edit.text().strip() or "/"
        self.browser_status_label.setText(f"当前：{current_prefix}，找到 {folder_count} 个文件夹，{file_count} 个文件。")

    def on_browser_item_clicked(self, item: QListWidgetItem) -> None:
        entry_key = item.data(Qt.UserRole)
        entry_type = item.data(Qt.UserRole + 1)
        if not entry_key or not entry_type:
            return
        if entry_type == "folder":
            self.prefix_edit.setText(str(entry_key))
            self.browser_status_label.setText(f"已选文件夹：{entry_key}")
            self._save_cached_cloud_settings()
        else:
            self.browser_status_label.setText(f"当前文件：{entry_key}")

    def on_browser_item_double_clicked(self, item: QListWidgetItem) -> None:
        entry_key = item.data(Qt.UserRole)
        entry_type = item.data(Qt.UserRole + 1)
        if entry_type != "folder" or not entry_key:
            return
        self.prefix_edit.setText(str(entry_key))
        self._save_cached_cloud_settings()
        self.load_browser_entries()

    def go_to_parent_prefix(self) -> None:
        current = self.prefix_edit.text().strip().strip("/")
        if not current:
            self.prefix_edit.clear()
        else:
            self.prefix_edit.setText("/".join(current.split("/")[:-1]))
        self._save_cached_cloud_settings()
        self.load_browser_entries()

    def _read_settings(self) -> dict:
        if not self.SETTINGS_FILE.exists():
            return {}
        try:
            return json.loads(self.SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _write_settings(self, settings: dict) -> None:
        self.SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")

    def _bind_cloud_cache_events(self) -> None:
        self.cloud_type_combo.currentTextChanged.connect(self._on_cloud_type_changed)
        for widget in (
            self.endpoint_edit,
            self.access_key_id_edit,
            self.access_key_secret_edit,
            self.bucket_combo,
            self.prefix_edit,
        ):
            if hasattr(widget, "editingFinished"):
                widget.editingFinished.connect(self._save_cached_cloud_settings)
        self.bucket_combo.currentTextChanged.connect(self._save_cached_cloud_settings)

    def _save_cached_cloud_settings(self) -> None:
        settings = self._read_settings()
        cloud_type = self.cloud_type_combo.currentText().strip() or "aliyun"
        profiles = settings.setdefault("cloud_profiles", {})
        profile = profiles.setdefault(cloud_type, {})
        profile["access_key_id"] = self.access_key_id_edit.text().strip()
        profile["access_key_secret"] = self.access_key_secret_edit.text().strip()
        profile["endpoint"] = self.endpoint_edit.text().strip()
        profile["bucket_name"] = self.bucket_combo.currentText().strip()
        profile["prefix"] = self.prefix_edit.text().strip()
        settings["cloud_type"] = cloud_type
        self._write_settings(settings)

    def _load_cached_cloud_settings(self) -> None:
        settings = self._read_settings()
        cloud_type = settings.get("cloud_type", self.cloud_type_combo.currentText().strip() or "aliyun")
        index = self.cloud_type_combo.findText(cloud_type)
        if index >= 0:
            self.cloud_type_combo.setCurrentIndex(index)
        self._apply_cached_cloud_profile(cloud_type)

    def _apply_cached_cloud_profile(self, cloud_type: str) -> None:
        settings = self._read_settings()
        profile = settings.get("cloud_profiles", {}).get(cloud_type, {})
        self.access_key_id_edit.setText(profile.get("access_key_id", ""))
        self.access_key_secret_edit.setText(profile.get("access_key_secret", ""))
        self.endpoint_edit.setText(profile.get("endpoint", ""))
        bucket_name = profile.get("bucket_name", "")
        self.bucket_combo.clear()
        if bucket_name:
            self.bucket_combo.addItem(bucket_name)
        self.prefix_edit.setText(profile.get("prefix", ""))
        self._show_browser_placeholder("请点“加载当前层级”选择目录。")

    def _on_cloud_type_changed(self, cloud_type: str) -> None:
        self._apply_cached_cloud_profile(cloud_type)
        self._save_cached_cloud_settings()

    def start_download(self) -> None:
        try:
            options = self.build_options()
            if options.skip_download:
                QMessageBox.information(self, "本地模式", "本地目录模式不需要下载，直接点击“生成模板”即可。")
                return
            oss_config = self.build_cloud_config()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._active_action = "download"
        self._set_running(True)
        self.result_text.setPlainText("正在下载并生成模板，请稍候...")
        worker = PhotoWorker(
            task=lambda: run_photo_download_and_template(options, oss_config=oss_config, logger=self.log_fn),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def start_generate_template(self) -> None:
        try:
            options = self.build_options()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._active_action = "template"
        self._set_running(True)
        self.result_text.setPlainText("正在生成模板，请稍候...")
        worker = PhotoWorker(
            task=lambda: run_photo_template_only(options, logger=self.log_fn),
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
        self.download_button.setEnabled(not running)
        self.generate_button.setEnabled(not running)
        self.classify_button.setEnabled(not running)

    def on_success(self, summary: WorkflowSummary) -> None:
        self._set_running(False)
        if hasattr(self, "load_prefix_button"):
            self.load_prefix_button.setEnabled(True)
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
            if self.filter_download_checkbox.isChecked():
                lines.append(f"下载过滤列：{self.filter_column_combo.currentText().strip()}")
        if self._active_action == "download":
            lines.append(f"模板文件数：{summary.template_file_count}")
            lines.append("本次已完成下载并生成模板。")
        elif self._active_action == "template":
            lines.append(f"模板文件数：{summary.template_file_count}")
            lines.append("已根据当前下载目录或本地目录生成模板。")
        else:
            lines.append(f"分类完成：{summary.classified_count}")
            if summary.report_path is not None:
                lines.append(f"结果清单：{summary.report_path}")
        self.result_text.setPlainText("\n".join(lines))

    def on_error(self, error_text: str) -> None:
        self._set_running(False)
        if hasattr(self, "load_prefix_button"):
            self.load_prefix_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])
