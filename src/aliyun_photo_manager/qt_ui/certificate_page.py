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

from ..certificate_filter import (
    CertificateFilterOptions,
    CertificateFilterSummary,
    CertificateTemplateSummary,
    generate_certificate_template,
    list_template_headers,
    load_column_values,
    run_certificate_filter,
)
from ..config import OssConfig, validate_oss_config
from ..downloader import BrowserEntry, DownloadResult, download_objects, list_browser_entries, list_buckets
from .common import AppComboBox


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class SimpleWorker(QRunnable):
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


class CertificatePage(QWidget):
    LABEL_WIDTH = 118
    ACTION_WIDTH = 116
    SETTINGS_FILE = Path(__file__).resolve().parents[3] / ".gui_settings.json"

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.browser_entries: list[BrowserEntry] = []
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._create_card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        title = QLabel("证件资料处理")
        title.setProperty("heroTitle", True)
        intro = QLabel("支持云下载按表过滤，以及按模板动态分层导出。")
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

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.source_mode_combo = AppComboBox()
        self.source_mode_combo.addItems(["本地目录", "云存储下载后处理"])
        self.source_mode_combo.currentIndexChanged.connect(self.update_source_mode_state)
        self._add_row(form, 0, "数据来源", self.source_mode_combo)

        self.cloud_section = QWidget()
        cloud_form = QGridLayout(self.cloud_section)
        cloud_form.setContentsMargins(0, 0, 0, 0)
        cloud_form.setHorizontalSpacing(14)
        cloud_form.setVerticalSpacing(14)

        self.cloud_type_combo = AppComboBox()
        self.cloud_type_combo.addItems(["aliyun", "tencent"])
        self.cloud_type_combo.currentTextChanged.connect(self._on_cloud_type_changed)
        self._add_row(cloud_form, 0, "云类型", self.cloud_type_combo)

        self.endpoint_edit = QLineEdit()
        self.endpoint_edit.setPlaceholderText("阿里云填 Endpoint，腾讯云填 Region 或 Endpoint")
        self._add_row(cloud_form, 1, "Endpoint / Region", self.endpoint_edit)

        self.access_key_id_edit = QLineEdit()
        self._add_row(cloud_form, 2, "AccessKey ID", self.access_key_id_edit)

        self.access_key_secret_edit = QLineEdit()
        self.access_key_secret_edit.setEchoMode(QLineEdit.Password)
        self._add_row(cloud_form, 3, "AccessKey Secret", self.access_key_secret_edit)

        self.bucket_combo = AppComboBox()
        self.bucket_combo.setEditable(False)
        self.load_bucket_button = QPushButton("加载 Bucket")
        self.load_bucket_button.setFixedWidth(self.ACTION_WIDTH)
        self.load_bucket_button.clicked.connect(self.load_buckets)
        self._add_row(cloud_form, 4, "Bucket", self._wrap_with_button(self.bucket_combo, self.load_bucket_button))

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

        form.addWidget(self.cloud_section, 1, 0, 1, 2)

        self.source_edit = QLineEdit()
        self._add_row(form, 2, "资料目录", self._with_dir_button(self.source_edit))

        self.template_edit = QLineEdit()
        self._add_row(form, 3, "处理模板", self._with_file_button(self.template_edit, self.choose_template))

        self.match_combo = AppComboBox()
        load_headers_btn = QPushButton("加载列")
        load_headers_btn.clicked.connect(self.load_headers)
        self._add_row(form, 4, "匹配列", self._wrap_with_button(self.match_combo, load_headers_btn))

        self.output_edit = QLineEdit()
        self._add_row(form, 5, "输出目录", self._with_dir_button(self.output_edit))

        self.filter_download_checkbox = QCheckBox("云下载时按表过滤")
        self.filter_download_checkbox.toggled.connect(self.update_filter_state)
        form.addWidget(self.filter_download_checkbox, 6, 1)

        self.filter_template_edit = QLineEdit()
        self._add_row(form, 7, "过滤表", self._with_file_button(self.filter_template_edit, self.choose_filter_template))

        self.filter_column_combo = AppComboBox()
        load_filter_headers_btn = QPushButton("加载列")
        load_filter_headers_btn.clicked.connect(self.load_filter_headers)
        self._add_row(form, 8, "文件夹名列", self._wrap_with_button(self.filter_column_combo, load_filter_headers_btn))

        self.mode_combo = AppComboBox()
        self.mode_combo.addItems(["复制整个人员文件夹", "只复制关键词文件"])
        self.mode_combo.currentIndexChanged.connect(self.update_keyword_state)
        self._add_row(form, 9, "筛选模式", self.mode_combo)

        self.keyword_edit = QLineEdit("学历证书")
        self._add_row(form, 10, "文件关键词", self.keyword_edit)

        self.rename_checkbox = QCheckBox("导出后文件夹重命名")
        self.rename_checkbox.toggled.connect(self.update_rename_state)
        form.addWidget(self.rename_checkbox, 11, 1)

        self.folder_name_combo = AppComboBox()
        self.folder_name_combo.setEnabled(False)
        self._add_row(form, 12, "导出名称列", self.folder_name_combo)

        self.classify_checkbox = QCheckBox("按所选列建立目录")
        self.classify_checkbox.setChecked(True)
        self.classify_checkbox.toggled.connect(self.update_classify_state)
        form.addWidget(self.classify_checkbox, 13, 1)

        self.classify_columns_list = QListWidget()
        self.classify_columns_list.setMinimumHeight(84)
        self._add_row(form, 14, "分层列", self.classify_columns_list)

        self.dry_run_checkbox = QCheckBox("仅预览，不实际执行")
        form.addWidget(self.dry_run_checkbox, 15, 1)

        body_layout.addLayout(form)

        actions = QHBoxLayout()
        self.download_button = QPushButton("下载")
        self.download_button.setProperty("accent", True)
        self.download_button.clicked.connect(self.start_download)
        actions.addWidget(self.download_button)
        self.generate_button = QPushButton("生成模板")
        self.generate_button.clicked.connect(self.start_generate_template)
        actions.addWidget(self.generate_button)
        self.run_button = QPushButton("按模板分类")
        self.run_button.clicked.connect(self.start_run)
        actions.addWidget(self.run_button)
        actions.addStretch(1)
        body_layout.addLayout(actions)

        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(180)
        body_layout.addWidget(self.result_text)

        self.update_source_mode_state()
        self.update_filter_state()
        self.update_keyword_state()
        self.update_classify_state()
        self._bind_cloud_cache_events()
        self._load_cached_cloud_settings()
        self._show_classify_placeholder()
        self._show_browser_placeholder()

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> None:
        label = QLabel(label_text)
        label.setProperty("formLabel", True)
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        label.setFixedWidth(self.LABEL_WIDTH)
        layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

    def _with_dir_button(self, line_edit: QLineEdit) -> QWidget:
        button = QPushButton("选择目录")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(lambda: self.choose_directory(line_edit))
        return self._wrap_with_button(line_edit, button)

    def _with_file_button(self, line_edit: QLineEdit, callback) -> QWidget:
        button = QPushButton("选择文件")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(callback)
        return self._wrap_with_button(line_edit, button)

    def _wrap_with_button(self, field: QWidget, button: QPushButton) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        layout.addWidget(field, 1)
        layout.addWidget(button)
        return row

    def choose_directory(self, line_edit: QLineEdit) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择目录", line_edit.text().strip() or str(Path.cwd()))
        if selected:
            line_edit.setText(selected)

    def choose_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择处理模板", "", "Excel 文件 (*.xlsx)")
        if selected:
            self.template_edit.setText(selected)
            self.load_headers()

    def choose_filter_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择过滤表", "", "Excel 文件 (*.xlsx)")
        if selected:
            self.filter_template_edit.setText(selected)
            self.load_filter_headers()

    def load_headers(self) -> None:
        headers = list_template_headers(Path(self.template_edit.text().strip()))
        self.match_combo.clear()
        self.folder_name_combo.clear()
        self.classify_columns_list.clear()
        self.match_combo.addItems(headers)
        self.folder_name_combo.addItems(headers)
        default_columns = {"分类一", "分类二", "分类三"}
        for header in headers:
            item = QListWidgetItem(header)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if header in default_columns else Qt.Unchecked)
            self.classify_columns_list.addItem(item)
        if not headers:
            self._show_classify_placeholder("模板里没有可用列。")

    def load_filter_headers(self) -> None:
        headers = list_template_headers(Path(self.filter_template_edit.text().strip()))
        self.filter_column_combo.clear()
        self.filter_column_combo.addItems(headers)

    def _selected_classify_columns(self) -> list[str]:
        selected: list[str] = []
        for index in range(self.classify_columns_list.count()):
            item = self.classify_columns_list.item(index)
            if item.checkState() == Qt.Checked:
                selected.append(item.text())
        return selected

    def _show_classify_placeholder(self, text: str = "请先选择处理模板并点击“加载列”。") -> None:
        self.classify_columns_list.clear()
        item = QListWidgetItem(text)
        item.setFlags(Qt.NoItemFlags)
        self.classify_columns_list.addItem(item)

    def update_source_mode_state(self) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        self.cloud_section.setVisible(cloud_mode)
        self.cloud_section.setEnabled(cloud_mode)
        self.filter_download_checkbox.setEnabled(cloud_mode)
        self.filter_download_checkbox.setVisible(cloud_mode)
        self.download_button.setVisible(cloud_mode)

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

    def update_keyword_state(self) -> None:
        self.keyword_edit.setEnabled(self.mode_combo.currentIndex() == 1)

    def update_rename_state(self, checked: bool) -> None:
        self.folder_name_combo.setEnabled(checked)

    def update_classify_state(self) -> None:
        self.classify_columns_list.setEnabled(self.classify_checkbox.isChecked())

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
        worker = SimpleWorker(
            task=lambda: list_buckets(
                access_key_id=access_key_id,
                access_key_secret=access_key_secret,
                endpoint=endpoint,
                cloud_type=cloud_type,
            ),
            on_success=self.on_buckets_loaded,
            on_error=self.on_worker_error,
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
        worker = SimpleWorker(
            task=lambda: list_browser_entries(config, prefix),
            on_success=self.on_browser_entries_loaded,
            on_error=self.on_worker_error,
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
        profile["certificate_bucket_name"] = self.bucket_combo.currentText().strip()
        profile["certificate_prefix"] = self.prefix_edit.text().strip()
        settings["cloud_type"] = cloud_type
        settings["certificate_bucket_name"] = self.bucket_combo.currentText().strip()
        settings["certificate_prefix"] = self.prefix_edit.text().strip()
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
        bucket_name = profile.get("certificate_bucket_name", profile.get("bucket_name", ""))
        self.bucket_combo.clear()
        if bucket_name:
            self.bucket_combo.addItem(bucket_name)
        self.prefix_edit.setText(profile.get("certificate_prefix", profile.get("prefix", "")))
        self._show_browser_placeholder("请点“加载当前层级”选择目录。")

    def _on_cloud_type_changed(self, cloud_type: str) -> None:
        self._apply_cached_cloud_profile(cloud_type)
        self._save_cached_cloud_settings()

    def start_download(self) -> None:
        try:
            config = self.build_cloud_config()
            target_dir = Path(self.source_edit.text().strip())
            prefix = self.prefix_edit.text().strip()
            key_filter = None
            filter_enabled = self.filter_download_checkbox.isChecked()
            filter_column = ""
            if filter_enabled:
                filter_template = Path(self.filter_template_edit.text().strip())
                filter_column = self.filter_column_combo.currentText().strip()
                allowed_people = set(load_column_values(filter_template, filter_column))

                def key_filter(object_key: str) -> bool:
                    relative_path = object_key
                    normalized_prefix = prefix.strip().strip("/")
                    if normalized_prefix:
                        prefix_with_slash = normalized_prefix + "/"
                        if object_key.startswith(prefix_with_slash):
                            relative_path = object_key[len(prefix_with_slash):]
                    parts = Path(relative_path.lstrip("/")).parts
                    return bool(parts) and parts[0] in allowed_people
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._set_running(True)
        worker = SimpleWorker(
            task=lambda: download_objects(
                config=config,
                prefix=prefix,
                download_dir=target_dir,
                dry_run=self.dry_run_checkbox.isChecked(),
                skip_existing=True,
                logger=self.log_fn,
                key_filter=key_filter,
                stage="certificate_download",
            ),
            on_success=lambda result: self.on_download_success(result, filter_enabled, filter_column),
            on_error=self.on_worker_error,
        )
        self.thread_pool.start(worker)

    def on_download_success(self, result: DownloadResult, filter_enabled: bool, filter_column: str) -> None:
        self._set_running(False)
        if hasattr(self, "load_prefix_button"):
            self.load_prefix_button.setEnabled(True)
        lines = [
            f"云端可下载文件：{result.total_found}",
            f"已下载：{result.downloaded_count}",
            f"跳过已存在：{result.skipped_existing_count}",
        ]
        if filter_enabled:
            lines.append(f"下载过滤列：{filter_column}")
        self.result_text.setPlainText("\n".join(lines))

    def start_generate_template(self) -> None:
        try:
            source_dir = Path(self.source_edit.text().strip())
            output_dir = Path(self.output_edit.text().strip() or self.source_edit.text().strip())
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._set_running(True)
        worker = SimpleWorker(
            task=lambda: generate_certificate_template(
                source_dir=source_dir,
                output_dir=output_dir,
                dry_run=self.dry_run_checkbox.isChecked(),
                logger=self.log_fn,
            ),
            on_success=self.on_template_success,
            on_error=self.on_worker_error,
        )
        self.thread_pool.start(worker)

    def start_run(self) -> None:
        try:
            classify_columns = self._selected_classify_columns() if self.classify_checkbox.isChecked() else []
            options = CertificateFilterOptions(
                template_path=Path(self.template_edit.text().strip()),
                source_dir=Path(self.source_edit.text().strip()),
                output_dir=Path(self.output_edit.text().strip()),
                match_column=self.match_combo.currentText().strip(),
                rename_folder=self.rename_checkbox.isChecked(),
                folder_name_column=self.folder_name_combo.currentText().strip() if self.rename_checkbox.isChecked() else "",
                classify_output=self.classify_checkbox.isChecked(),
                classify_columns=classify_columns,
                keyword=self.keyword_edit.text().strip() if self.mode_combo.currentIndex() == 1 else "",
                dry_run=self.dry_run_checkbox.isChecked(),
            )
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._set_running(True)
        worker = SimpleWorker(
            task=lambda: run_certificate_filter(options, logger=self.log_fn),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def _set_running(self, running: bool) -> None:
        self.download_button.setEnabled(not running)
        self.generate_button.setEnabled(not running)
        self.run_button.setEnabled(not running)

    def on_template_success(self, summary: CertificateTemplateSummary) -> None:
        self._set_running(False)
        self.template_edit.setText(str(summary.template_path))
        self.load_headers()
        self.result_text.setPlainText(f"已生成模板：{summary.template_path}")

    def on_success(self, summary: CertificateFilterSummary) -> None:
        self._set_running(False)
        lines = [
            f"模板：{summary.template_path}",
            f"来源目录：{summary.source_dir}",
            f"输出目录：{summary.output_dir}",
            f"命中人员：{summary.matched_people}",
            f"复制文件：{summary.copied_files}",
        ]
        if summary.classify_columns:
            lines.append(f"分层列：{' / '.join(summary.classify_columns)}")
        self.result_text.setPlainText("\n".join(lines))

    def on_error(self, error_text: str) -> None:
        self._set_running(False)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "按模板分类失败", error_text.splitlines()[0])

    def on_worker_error(self, error_text: str) -> None:
        self._set_running(False)
        if hasattr(self, "load_prefix_button"):
            self.load_prefix_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])
