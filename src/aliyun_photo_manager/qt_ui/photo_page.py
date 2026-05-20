from __future__ import annotations

import json
import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Qt, Signal
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
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
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
from ..downloader import BrowserEntry, list_browser_entries, list_bucket_infos, search_folder_entries
from .base_page import BaseToolPage
from .common import AppComboBox


EXCEL_FILE_FILTER = "Excel 文件 (*.xlsx *.xls)"


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)
    progress = Signal(str, int, int, str)


class PhotoWorker(QRunnable):
    def __init__(self, task: Callable[[], object], on_success: Callable[[object], None], on_error: Callable[[str], None]) -> None:
        super().__init__()
        self.setAutoDelete(False)
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


class PhotoPage(BaseToolPage):
    LABEL_WIDTH = 118
    ACTION_WIDTH = 116
    SETTINGS_FILE = Path(__file__).resolve().parents[3] / ".gui_settings.json"

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="照片下载与分类",
            description="支持本地目录直接生成模板并分类，也支持云存储按表过滤下载后再处理。",
            show_steps=True,
            show_log=False,
        )
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.photo_filter_headers: list[str] = []
        self.browser_entries: list[BrowserEntry] = []
        self.browser_cache: dict[str, list[BrowserEntry]] = {}
        self.browser_listing_prefix = ""
        self.selected_prefix = ""
        self._loading_bucket_list = False
        self._browser_request_id = 0
        self._active_action = "download"
        self.progress_signal = WorkerSignals()
        self.progress_signal.progress.connect(self.update_progress)
        self._build_ui()

    def _build_ui(self) -> None:
        body_layout = self.left_card.body_layout
        if self.left_card.title is not None:
            self.left_card.title.setText("执行参数")
        if self.right_card.title is not None:
            self.right_card.title.setText("执行结果")

        title_label = QLabel("执行参数")
        title_label.setProperty("sectionTitle", True)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.source_mode_combo = AppComboBox()
        self.source_mode_combo.addItems(["本地目录", "云存储"])
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
        self.cloud_type_combo.addItem("阿里云 OSS", "aliyun")
        self.cloud_type_combo.addItem("腾讯云 COS", "tencent")
        self.cloud_type_combo.currentTextChanged.connect(self._on_cloud_type_changed)
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

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("可留空，留空时搜索全部目录")
        self._add_row(cloud_form, 5, "前缀搜索", self.search_edit)

        browser_controls = QWidget()
        browser_controls_layout = QHBoxLayout(browser_controls)
        browser_controls_layout.setContentsMargins(0, 0, 0, 0)
        browser_controls_layout.setSpacing(10)
        self.load_prefix_button = QPushButton("搜索")
        self.load_prefix_button.clicked.connect(self.search_browser_entries)
        browser_controls_layout.addWidget(self.load_prefix_button)
        self.parent_prefix_button = QPushButton("返回上一层")
        self.parent_prefix_button.clicked.connect(self.go_to_parent_prefix)
        browser_controls_layout.addWidget(self.parent_prefix_button)
        browser_controls_layout.addStretch(1)
        self._add_row(cloud_form, 6, "目录搜索", browser_controls)

        self.browser_list = QListWidget()
        self.browser_list.setMinimumHeight(150)
        self.browser_list.itemClicked.connect(self.on_browser_item_clicked)
        self._add_row(cloud_form, 7, "", self.browser_list)

        self.browser_status_label = QLabel("请先选择 Bucket，再输入前缀后点击搜索。")
        self.browser_status_label.setWordWrap(True)
        self._add_row(cloud_form, 8, "", self.browser_status_label)

        self.selected_path_label = QLabel("未选择下载路径")
        self.selected_path_label.setWordWrap(True)
        self._add_row(cloud_form, 9, "已选路径", self.selected_path_label)

        cloud_layout.addLayout(cloud_form)
        form.addWidget(self.cloud_section, 1, 0, 1, 2)

        self.download_dir_edit = QLineEdit()
        self.download_dir_label = self._add_row(form, 2, "本地目录", self._with_dir_button(self.download_dir_edit))

        self.sorted_dir_edit = QLineEdit()
        self.sorted_dir_label = self._add_row(form, 3, "分类目录", self._with_dir_button(self.sorted_dir_edit))

        self.filter_download_checkbox = QCheckBox("云下载时按表过滤")
        self.filter_download_checkbox.setChecked(True)
        self.filter_download_checkbox.toggled.connect(self.update_filter_state)
        form.addWidget(self.filter_download_checkbox, 4, 1)

        self.filter_template_edit = QLineEdit()
        filter_template_row = self._with_file_button(self.filter_template_edit, self.choose_filter_template)
        self.filter_template_label = self._add_row(form, 5, "过滤表", filter_template_row)

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
        self.filter_column_label = self._add_row(form, 6, "前缀列", filter_column_row)

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
        body_layout.addStretch(1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.right_card.body_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("未开始")
        self.progress_label.setWordWrap(True)
        self.right_card.body_layout.addWidget(self.progress_label)

        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(160)
        self.result_text.setPlaceholderText("这里会显示下载、模板生成和分类结果。")
        self.right_card.body_layout.addWidget(self.result_text, 1)

        self._bind_cloud_cache_events()
        self._load_cached_cloud_settings()
        self._show_browser_placeholder()
        self.update_source_mode_state()
        self.update_filter_state()

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> QLabel | None:
        label = None
        if label_text:
            label = QLabel(label_text)
            label.setProperty("formLabel", True)
            label.setFixedWidth(self.LABEL_WIDTH)
            layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)
        return label

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
        selected, _ = QFileDialog.getOpenFileName(self, "选择过滤表", "", EXCEL_FILE_FILTER)
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
        self.filter_download_checkbox.setVisible(cloud_mode)
        self.filter_download_checkbox.setEnabled(cloud_mode)
        self.skip_existing_checkbox.setVisible(cloud_mode)
        self.skip_existing_checkbox.setEnabled(cloud_mode)
        self.download_button.setVisible(cloud_mode)
        self.download_button.setEnabled(True)
        if self.download_dir_label is not None:
            self.download_dir_label.setText("下载目录" if cloud_mode else "本地目录")
        if self.sorted_dir_label is not None:
            self.sorted_dir_label.setText("分类目录")
        self.update_filter_state()

    def update_filter_state(self) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        enabled = cloud_mode and self.filter_download_checkbox.isChecked()
        self.filter_template_edit.setEnabled(enabled)
        self.filter_column_combo.setEnabled(enabled)
        if self.filter_template_label is not None:
            self.filter_template_label.setVisible(cloud_mode)
        if self.filter_column_label is not None:
            self.filter_column_label.setVisible(cloud_mode)
        self.filter_template_edit.parentWidget().setVisible(cloud_mode)
        self.filter_column_combo.parentWidget().setVisible(cloud_mode)

    def _show_browser_placeholder(self, text: str = "请先选择 Bucket，再输入前缀后点击搜索。") -> None:
        self.browser_entries = []
        self.browser_listing_prefix = ""
        self.browser_list.clear()
        item = QListWidgetItem(text)
        item.setFlags(Qt.NoItemFlags)
        self.browser_list.addItem(item)
        self.browser_status_label.setText(text)
        self.selected_path_label.setText("未选择下载路径")
        self.selected_prefix = ""

    def build_options(self, *, require_sorted_dir: bool, require_selected_prefix: bool) -> RunOptions:
        download_dir = self.download_dir_edit.text().strip()
        if not download_dir:
            raise ValueError("请先选择目录。")

        sorted_dir_text = self.sorted_dir_edit.text().strip()
        sorted_dir = Path(sorted_dir_text) if sorted_dir_text else None
        if require_sorted_dir and sorted_dir is None:
            raise ValueError("请先选择分类目录。")

        cloud_mode = self.source_mode_combo.currentIndex() == 1
        if require_selected_prefix and cloud_mode and not self.selected_prefix:
            raise ValueError("请先从下方搜索结果中选择下载路径。")

        filter_template_path = None
        filter_column = ""
        filter_download = cloud_mode and self.filter_download_checkbox.isChecked()
        if filter_download:
            filter_template_text = self.filter_template_edit.text().strip()
            if not filter_template_text:
                raise ValueError("已启用按表过滤下载，请先选择过滤表。")
            filter_template_path = Path(filter_template_text)
            if not filter_template_path.exists():
                raise ValueError("过滤表文件不存在。")
            filter_column = self.filter_column_combo.currentText().strip()
            if not filter_column:
                raise ValueError("已启用按表过滤下载，请先选择前缀列。")

        return RunOptions(
            download_dir=Path(download_dir),
            sorted_dir=sorted_dir,
            prefix=self.selected_prefix if cloud_mode else "",
            skip_download=not cloud_mode,
            dry_run=self.dry_run_checkbox.isChecked(),
            skip_existing=self.skip_existing_checkbox.isChecked(),
            filter_download=filter_download,
            filter_template_path=filter_template_path,
            filter_column=filter_column,
        )

    def build_cloud_config(self) -> OssConfig:
        cloud_type = self.cloud_type_combo.currentData() or "aliyun"
        endpoint = self.endpoint_edit.text().strip()
        bucket_location = self.bucket_combo.currentData()
        if cloud_type == "tencent" and isinstance(bucket_location, str) and bucket_location.strip():
            endpoint = bucket_location.strip()
        return validate_oss_config(
            OssConfig(
                cloud_type=cloud_type,
                access_key_id=self.access_key_id_edit.text().strip(),
                access_key_secret=self.access_key_secret_edit.text().strip(),
                endpoint=endpoint,
                bucket_name=self.bucket_combo.currentText().strip(),
            )
        )

    def _append_result_log(self, line: str) -> None:
        current = self.result_text.toPlainText().rstrip()
        self.result_text.setPlainText(f"{current}\n{line}" if current else line)

    def _format_browser_debug(
        self,
        *,
        action: str,
        config: OssConfig,
        prefix: str | None = None,
        keyword: str | None = None,
        count: int | None = None,
        entries: list[BrowserEntry] | None = None,
    ) -> str:
        parts = [
            f"[browser] {action}",
            f"cloud={config.cloud_type}",
            f"endpoint={config.endpoint}",
            f"bucket={config.bucket_name}",
        ]
        if prefix is not None:
            parts.append(f"prefix={prefix or '/'}")
        if keyword is not None:
            parts.append(f"keyword={keyword or '<empty>'}")
        if count is not None:
            parts.append(f"count={count}")
        if entries is not None:
            folder_count = sum(1 for entry in entries if entry.entry_type == "folder")
            file_count = sum(1 for entry in entries if entry.entry_type == "file")
            sample = ", ".join(entry.key for entry in entries[:8])
            parts.append(f"folders={folder_count}")
            parts.append(f"files={file_count}")
            if sample:
                parts.append(f"sample={sample}")
        return " | ".join(parts)

    def load_buckets(self) -> None:
        try:
            cloud_type = self.cloud_type_combo.currentData() or "aliyun"
            access_key_id = self.access_key_id_edit.text().strip()
            access_key_secret = self.access_key_secret_edit.text().strip()
            endpoint = self.endpoint_edit.text().strip()
            if not access_key_id or not access_key_secret or not endpoint:
                raise ValueError("请先填写云类型、Endpoint / Region 和 AccessKey。")
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self.load_bucket_button.setEnabled(False)
        self._loading_bucket_list = True
        self.result_text.setPlainText(
            "\n".join(
                [
                    "[bucket] start",
                    f"cloud={cloud_type}",
                    f"endpoint={endpoint}",
                    f"access_key_id_len={len(access_key_id)}",
                ]
            )
        )
        worker = PhotoWorker(
            task=lambda: list_bucket_infos(
                access_key_id=access_key_id,
                access_key_secret=access_key_secret,
                endpoint=endpoint,
                cloud_type=cloud_type,
            ),
            on_success=self.on_buckets_loaded,
            on_error=self.on_bucket_load_error,
        )
        self.thread_pool.start(worker)

    def on_buckets_loaded(self, buckets: list[str]) -> None:
        self.load_bucket_button.setEnabled(True)
        current = self.bucket_combo.currentText().strip()
        self.bucket_combo.clear()
        bucket_names: list[str] = []
        bucket_debug: list[str] = []
        for bucket in buckets:
            name = getattr(bucket, "name", str(bucket))
            location = getattr(bucket, "location", "") or ""
            self.bucket_combo.addItem(name, location)
            bucket_names.append(name)
            bucket_debug.append(f"{name}:{location or '-'}")
        if current:
            index = self.bucket_combo.findText(current)
            if index >= 0:
                self.bucket_combo.setCurrentIndex(index)
        self._apply_current_bucket_location()
        self._loading_bucket_list = False
        self._save_cached_cloud_settings()
        self.result_text.setPlainText(
            "\n".join(
                [
                    "[bucket] done",
                    f"count={len(bucket_names)}",
                    f"buckets={', '.join(bucket_debug[:20])}",
                ]
            )
        )
        self._show_browser_placeholder("Bucket 已加载，请输入前缀后点击搜索。")
        try:
            config = self.build_cloud_config()
        except Exception:
            self.browser_status_label.setText("Bucket 已加载，可直接点搜索查看根层目录，或输入前缀后搜索。")
            return
        self._append_result_log(
            self._format_browser_debug(action="bucket-ready", config=config, prefix=self.selected_prefix)
        )

    def _load_browser_entries(self, config: OssConfig, prefix: str, status_template: str) -> None:
        request_id, context = self._start_browser_request(config)
        self._append_result_log(
            self._format_browser_debug(action=f"list-start#{request_id}", config=config, prefix=prefix)
        )
        worker = PhotoWorker(
            task=lambda: list_browser_entries(config, prefix),
            on_success=lambda entries: self._accept_browser_entries_loaded(
                entries,
                request_id=request_id,
                context=context,
                prefix=prefix,
                status=status_template.format(prefix=prefix or "/"),
                is_directory_listing=True,
            ),
            on_error=self.on_error,
        )
        self._browser_worker = worker
        self.thread_pool.start(worker)

    def _start_global_browser_search(self, config: OssConfig, keyword: str) -> None:
        self.browser_status_label.setText("正在搜索目录...")
        request_id, context = self._start_browser_request(config)
        self._append_result_log(
            self._format_browser_debug(action=f"search-start#{request_id}", config=config, keyword=keyword)
        )
        worker = PhotoWorker(
            task=lambda: search_folder_entries(config, keyword),
            on_success=lambda entries: self._accept_browser_entries_loaded(
                entries,
                request_id=request_id,
                context=context,
                prefix=None,
                status=f"找到 {len(entries)} 个匹配条目，请在下方选择目录作为下载路径。",
                is_directory_listing=False,
            ),
            on_error=self.on_error,
        )
        self._browser_worker = worker
        self.thread_pool.start(worker)

    def _cloud_context(self, config: OssConfig) -> tuple[str, str, str]:
        return (config.cloud_type, config.endpoint, config.bucket_name)

    def _start_browser_request(self, config: OssConfig) -> tuple[int, tuple[str, str, str]]:
        self._browser_request_id += 1
        return self._browser_request_id, self._cloud_context(config)

    def _is_current_browser_request(self, request_id: int, context: tuple[str, str, str]) -> bool:
        if request_id != self._browser_request_id:
            return False
        try:
            current_config = self.build_cloud_config()
        except Exception:
            return False
        return self._cloud_context(current_config) == context

    def _search_in_current_entries(self, config: OssConfig, keyword: str, entries: list[BrowserEntry]) -> None:
        keyword_lower = keyword.lower()
        folder_candidates = [
            entry
            for entry in entries
            if entry.entry_type == "folder"
            and (
                entry.display_name.lower().startswith(keyword_lower)
                or entry.key.strip("/").lower().startswith(keyword_lower)
                or f"/{keyword_lower}" in f"/{entry.key.strip('/').lower()}"
            )
        ]
        exact_matches = [
            entry
            for entry in folder_candidates
            if entry.display_name.lower() == keyword_lower
            or entry.key.strip("/").lower() == keyword_lower
        ]
        matched_folder = None
        if len(exact_matches) == 1:
            matched_folder = exact_matches[0]
        elif len(folder_candidates) == 1:
            matched_folder = folder_candidates[0]

        if matched_folder is not None:
            target_prefix = matched_folder.key
            self.browser_status_label.setText(f"正在进入目录：{target_prefix}")
            self._load_browser_entries(config, target_prefix, "当前目录：{prefix}")
            return

        self._start_global_browser_search(config, keyword)

    def search_browser_entries(self) -> None:
        try:
            config = self.build_cloud_config()
            keyword = self.search_edit.text().strip()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self.load_prefix_button.setEnabled(False)
        current_prefix = self.selected_prefix

        if not keyword:
            self.browser_status_label.setText("正在加载目录...")
            cached_entries = self.browser_cache.get(current_prefix)
            if cached_entries is not None:
                self.on_browser_entries_loaded(
                    cached_entries,
                    prefix=current_prefix,
                    status=f"当前目录：{current_prefix or '/'}",
                    is_directory_listing=True,
                )
                return
            self._load_browser_entries(config, current_prefix, "当前目录：{prefix}")
            return

        cached_entries = self.browser_cache.get(current_prefix)
        if cached_entries is not None:
            self._search_in_current_entries(config, keyword, cached_entries)
            return

        self._start_global_browser_search(config, keyword)

    def _search_after_loading_current(
        self,
        config: OssConfig,
        keyword: str,
        current_prefix: str,
        entries: list[BrowserEntry],
    ) -> None:
        try:
            if self._cloud_context(config) != self._cloud_context(self.build_cloud_config()):
                return
        except Exception:
            return
        self.browser_cache[current_prefix] = entries
        self.browser_listing_prefix = current_prefix
        self._search_in_current_entries(config, keyword, entries)

    def _accept_browser_entries_loaded(
        self,
        entries: list[BrowserEntry],
        *,
        request_id: int,
        context: tuple[str, str, str],
        prefix: str | None = None,
        status: str | None = None,
        is_directory_listing: bool = False,
    ) -> None:
        if not self._is_current_browser_request(request_id, context):
            self.load_prefix_button.setEnabled(True)
            self._append_result_log(
                f"[browser] stale-response ignored | request_id={request_id} | context={context}"
            )
            return
        self.on_browser_entries_loaded(
            entries,
            prefix=prefix,
            status=status,
            is_directory_listing=is_directory_listing,
        )

    def on_browser_entries_loaded(
        self,
        entries: list[BrowserEntry],
        *,
        prefix: str | None = None,
        status: str | None = None,
        is_directory_listing: bool = False,
    ) -> None:
        self.load_prefix_button.setEnabled(True)
        self.browser_entries = entries
        if is_directory_listing:
            cache_prefix = prefix or ""
            self.browser_cache[cache_prefix] = entries
            self.browser_listing_prefix = cache_prefix
        else:
            self.browser_listing_prefix = ""
        self.browser_list.clear()
        try:
            config = self.build_cloud_config()
            self._append_result_log(
                self._format_browser_debug(
                    action="list-done" if is_directory_listing else "search-done",
                    config=config,
                    prefix=prefix,
                    count=len(entries),
                    entries=entries,
                )
            )
        except Exception:
            self._append_result_log(f"[browser] done | count={len(entries)}")
        if not entries:
            self._show_browser_placeholder("没有找到匹配的条目。")
            return
        if prefix is not None:
            self.selected_prefix = prefix
            self.selected_path_label.setText(prefix or "未选择下载路径")
        for entry in entries:
            label_prefix = "[目录]" if entry.entry_type == "folder" else "[文件]"
            item = QListWidgetItem(f"{label_prefix} {entry.display_name}")
            item.setData(Qt.UserRole, entry)
            self.browser_list.addItem(item)
        self.browser_status_label.setText(status or f"已加载 {len(entries)} 个条目，请在下方选择目录。")

    def on_browser_item_clicked(self, item: QListWidgetItem) -> None:
        entry = item.data(Qt.UserRole)
        if not entry:
            return
        if entry.entry_type != "folder":
            self.browser_status_label.setText("当前选中的是文件，请选择目录作为下载路径。")
            return
        self.selected_prefix = str(entry.key)
        self.selected_path_label.setText(self.selected_prefix)
        self.browser_status_label.setText(f"已选择下载路径：{self.selected_prefix}")
        self._save_cached_cloud_settings()

    def go_to_parent_prefix(self) -> None:
        current = self.selected_prefix.strip().strip("/")
        if not current:
            self.selected_prefix = ""
            self.selected_path_label.setText("未选择下载路径")
            self.browser_status_label.setText("当前已在根目录。")
        else:
            parent = "/".join(current.split("/")[:-1]).strip("/")
            self.selected_prefix = parent + "/" if parent else ""
            self.selected_path_label.setText(self.selected_prefix or "未选择下载路径")
            self.browser_status_label.setText(
                f"已返回上一层：{self.selected_prefix or '/'}"
            )
        self._save_cached_cloud_settings()

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
            self.search_edit,
        ):
            if hasattr(widget, "editingFinished"):
                widget.editingFinished.connect(self._save_cached_cloud_settings)
        self.bucket_combo.currentTextChanged.connect(self._on_bucket_changed)

    def _on_bucket_changed(self, bucket_name: str) -> None:
        if self._loading_bucket_list:
            return
        self._apply_current_bucket_location()
        self._browser_request_id += 1
        self.browser_cache.clear()
        self.browser_entries = []
        self.browser_listing_prefix = ""
        self.selected_prefix = ""
        self.selected_path_label.setText("未选择下载路径")
        self._save_cached_cloud_settings()
        if not bucket_name.strip():
            self._show_browser_placeholder("请先选择 Bucket，再输入前缀后点击搜索。")
            return
        self._show_browser_placeholder("Bucket 已切换，请输入前缀后点击搜索。")
        try:
            config = self.build_cloud_config()
        except Exception as exc:
            self.browser_status_label.setText(str(exc))
            return
        self._append_result_log(self._format_browser_debug(action="bucket-changed", config=config, prefix=""))

    def _apply_current_bucket_location(self) -> None:
        if (self.cloud_type_combo.currentData() or "aliyun") != "tencent":
            return
        location = self.bucket_combo.currentData()
        if isinstance(location, str) and location.strip():
            self.endpoint_edit.setText(location.strip())

    def _save_cached_cloud_settings(self) -> None:
        settings = self._read_settings()
        cloud_type = self.cloud_type_combo.currentData() or "aliyun"
        profiles = settings.setdefault("cloud_profiles", {})
        profile = profiles.setdefault(cloud_type, {})
        profile["access_key_id"] = self.access_key_id_edit.text().strip()
        profile["access_key_secret"] = self.access_key_secret_edit.text().strip()
        profile["endpoint"] = self.endpoint_edit.text().strip()
        profile["bucket_name"] = self.bucket_combo.currentText().strip()
        profile["search_keyword"] = self.search_edit.text().strip()
        profile["prefix"] = self.selected_prefix
        settings["cloud_type"] = cloud_type
        self._write_settings(settings)

    def _load_cached_cloud_settings(self) -> None:
        from ..config import normalize_cloud_type

        settings = self._read_settings()
        raw_cloud_type = settings.get("cloud_type", "aliyun")
        try:
            cloud_type = normalize_cloud_type(raw_cloud_type)
        except ValueError:
            cloud_type = "aliyun"
        index = self.cloud_type_combo.findData(cloud_type)
        if index >= 0:
            self.cloud_type_combo.setCurrentIndex(index)
        self._apply_cached_cloud_profile(cloud_type)

    def _apply_cached_cloud_profile(self, cloud_type: str) -> None:
        from ..config import normalize_cloud_type

        try:
            normalized_type = normalize_cloud_type(cloud_type)
        except ValueError:
            normalized_type = "aliyun"
        index = self.cloud_type_combo.findData(normalized_type)
        if index < 0:
            index = self.cloud_type_combo.findData("aliyun")
        if index >= 0:
            self.cloud_type_combo.setCurrentIndex(index)
        settings = self._read_settings()
        profile = settings.get("cloud_profiles", {}).get(normalized_type, {})
        self.access_key_id_edit.setText(profile.get("access_key_id", ""))
        self.access_key_secret_edit.setText(profile.get("access_key_secret", ""))
        self.endpoint_edit.setText(profile.get("endpoint", ""))
        bucket_name = profile.get("bucket_name", "")
        selected_prefix = profile.get("prefix", "")
        self.bucket_combo.clear()
        if bucket_name:
            self.bucket_combo.addItem(bucket_name)
        self.search_edit.setText(profile.get("search_keyword", ""))
        self._show_browser_placeholder("请输入前缀后点击搜索。")
        if selected_prefix:
            self.selected_prefix = selected_prefix
            self.selected_path_label.setText(selected_prefix)

    def _on_cloud_type_changed(self, display_text: str) -> None:
        self._browser_request_id += 1
        self._loading_bucket_list = True
        self._apply_cached_cloud_profile(str(self.cloud_type_combo.currentData() or "aliyun"))
        self._loading_bucket_list = False
        self._save_cached_cloud_settings()

    def start_download(self) -> None:
        try:
            options = self.build_options(require_sorted_dir=False, require_selected_prefix=True)
            oss_config = self.build_cloud_config()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._active_action = "download"
        self._set_running(True)
        self.reset_progress("准备下载...")
        self.result_text.setPlainText("正在下载并生成模板，请稍候...")
        worker = PhotoWorker(
            task=lambda: run_photo_download_and_template(
                options,
                oss_config=oss_config,
                logger=self.log_fn,
                progress_callback=self.make_progress_callback(),
            ),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def start_generate_template(self) -> None:
        try:
            options = self.build_options(require_sorted_dir=False, require_selected_prefix=False)
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._active_action = "template"
        self._set_running(True)
        self.reset_progress("正在生成模板...")
        self.result_text.setPlainText("正在生成模板，请稍候...")
        worker = PhotoWorker(
            task=lambda: run_photo_template_only(options, logger=self.log_fn),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self.thread_pool.start(worker)

    def start_classify(self) -> None:
        try:
            options = self.build_options(require_sorted_dir=True, require_selected_prefix=False)
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._active_action = "classify"
        self._set_running(True)
        self.reset_progress("正在分类...")
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
        self.load_prefix_button.setEnabled(not running)

    def make_progress_callback(self):
        def progress(stage: str, current: int, total: int, current_file: str) -> None:
            self.progress_signal.progress.emit(stage, current, total, current_file)

        return progress

    def reset_progress(self, text: str) -> None:
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_label.setText(text)

    def update_progress(self, stage: str, current: int, total: int, current_file: str) -> None:
        maximum = max(total, 1)
        self.progress_bar.setRange(0, maximum)
        self.progress_bar.setValue(min(current, maximum))
        filename = Path(current_file).name if current_file else ""
        if total:
            self.progress_label.setText(f"下载进度：{current}/{total} {filename}".strip())
        else:
            self.progress_label.setText("正在统计下载文件...")

    def on_success(self, summary: WorkflowSummary) -> None:
        self._set_running(False)
        if self._active_action == "download" and summary.download_result is not None:
            self.update_progress(
                "download",
                summary.download_result.total_found,
                summary.download_result.total_found,
                "",
            )
        else:
            self.progress_label.setText("任务完成")
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
            if self.selected_prefix:
                lines.append(f"下载路径：{self.selected_prefix}")
        if self._active_action == "download":
            lines.append(f"模板文件数：{summary.template_file_count}")
            lines.append("本次已完成下载并生成模板。")
        elif self._active_action == "template":
            lines.append(f"模板文件数：{summary.template_file_count}")
            lines.append("已根据当前目录生成模板。")
        else:
            lines.append(f"分类完成：{summary.classified_count}")
            if summary.report_path is not None:
                lines.append(f"结果清单：{summary.report_path}")
        self.result_text.setPlainText("\n".join(lines))

    def on_error(self, error_text: str) -> None:
        self._set_running(False)
        self.progress_label.setText("任务失败")
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])

    def on_bucket_load_error(self, error_text: str) -> None:
        self._loading_bucket_list = False
        self.load_bucket_button.setEnabled(True)
        self.on_error(error_text)
