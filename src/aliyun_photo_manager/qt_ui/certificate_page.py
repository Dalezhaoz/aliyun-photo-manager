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
from ..downloader import DownloadResult, download_objects, list_browser_entries, list_bucket_infos, search_folder_entries
from .base_page import BaseToolPage
from .common import AppComboBox


EXCEL_FILE_FILTER = "Excel 文件 (*.xlsx *.xls)"


def ensure_child_directory(base_dir: Path, child_name: str) -> Path:
    normalized = base_dir.expanduser().resolve()
    if normalized.name == child_name:
        return normalized
    return normalized / child_name


def build_prefixed_directory(base_dir: Path, prefix: str, child_name: str) -> Path:
    normalized = base_dir.expanduser().resolve()
    if normalized.name == child_name:
        return normalized
    cleaned_prefix = prefix.strip().strip("/")
    if cleaned_prefix:
        return normalized.joinpath(*cleaned_prefix.split("/"), child_name)
    return ensure_child_directory(normalized, child_name)


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)
    progress = Signal(str, int, int, str)


class SimpleWorker(QRunnable):
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


class CertificatePage(BaseToolPage):
    LABEL_WIDTH = 118
    ACTION_WIDTH = 116
    SETTINGS_FILE = Path(__file__).resolve().parents[3] / ".gui_settings.json"

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__(
            title="证件资料处理",
            description="本地模式直接按模板分类，云模式先按表过滤下载，再继续处理。",
            show_steps=True,
            show_log=False,
        )
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.browser_entries: list[BrowserEntry] = []
        self.browser_cache: dict[str, list[BrowserEntry]] = {}
        self.browser_listing_prefix = ""
        self.selected_prefix = ""
        self._loading_bucket_list = False
        self._browser_request_id = 0
        self.cloud_download_ready = False
        self.progress_signal = WorkerSignals()
        self.progress_signal.progress.connect(self.update_progress)
        self._build_ui()

    def _build_ui(self) -> None:
        body_layout = self.left_card.body_layout
        if self.left_card.title is not None:
            self.left_card.title.setText("执行参数")
        if self.right_card.title is not None:
            self.right_card.title.setText("执行结果")

        base_form = QGridLayout()
        base_form.setHorizontalSpacing(14)
        base_form.setVerticalSpacing(14)

        self.source_mode_combo = AppComboBox()
        self.source_mode_combo.addItems(["本地目录", "云存储"])
        self.source_mode_combo.currentIndexChanged.connect(self.update_source_mode_state)
        self._add_row(base_form, 0, "数据来源", self.source_mode_combo)

        self.source_edit = QLineEdit()
        self.source_label = self._add_row(base_form, 1, "本地目录", self._with_dir_button(self.source_edit))

        self.output_edit = QLineEdit()
        self.output_row = self._with_dir_button(self.output_edit)
        self.output_label = self._add_row(base_form, 2, "输出目录", self.output_row)

        body_layout.addLayout(base_form)

        self.cloud_section = self._create_card()
        cloud_layout = QVBoxLayout(self.cloud_section)
        cloud_layout.setContentsMargins(18, 18, 18, 18)
        cloud_layout.setSpacing(12)
        cloud_title = QLabel("云下载")
        cloud_title.setProperty("sectionTitle", True)
        cloud_layout.addWidget(cloud_title)

        cloud_form = QGridLayout()
        cloud_form.setHorizontalSpacing(14)
        cloud_form.setVerticalSpacing(14)

        self.cloud_type_combo = AppComboBox()
        self.cloud_type_combo.addItems(["aliyun", "tencent"])
        self.cloud_type_combo.currentTextChanged.connect(self._on_cloud_type_changed)
        self._add_row(cloud_form, 0, "云类型", self.cloud_type_combo)

        self.endpoint_edit = QLineEdit()
        self.endpoint_edit.setPlaceholderText("阿里云填写 Endpoint，腾讯云填写 Region 或 Endpoint")
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

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("可留空，留空时搜索全部目录")
        self._add_row(cloud_form, 5, "前缀搜索", self.search_edit)

        self.search_button = QPushButton("搜索")
        self.search_button.clicked.connect(self.search_browser_entries)
        search_controls = QWidget()
        search_controls_layout = QHBoxLayout(search_controls)
        search_controls_layout.setContentsMargins(0, 0, 0, 0)
        search_controls_layout.setSpacing(10)
        search_controls_layout.addWidget(self.search_button)
        self.parent_prefix_button = QPushButton("返回上一层")
        self.parent_prefix_button.clicked.connect(self.go_to_parent_prefix)
        search_controls_layout.addWidget(self.parent_prefix_button)
        search_controls_layout.addStretch(1)
        self._add_row(cloud_form, 6, "目录搜索", search_controls)

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

        self.filter_download_checkbox = QCheckBox("云下载时按表过滤")
        self.filter_download_checkbox.setChecked(True)
        self.filter_download_checkbox.toggled.connect(self.update_filter_state)
        cloud_form.addWidget(self.filter_download_checkbox, 10, 1)

        self.filter_template_edit = QLineEdit()
        self.filter_template_label = self._add_row(
            cloud_form,
            11,
            "过滤表",
            self._with_file_button(self.filter_template_edit, self.choose_filter_template),
        )

        self.filter_column_combo = AppComboBox()
        load_filter_headers_btn = QPushButton("加载列")
        load_filter_headers_btn.clicked.connect(self.load_filter_headers)
        self.filter_column_row = self._wrap_with_button(self.filter_column_combo, load_filter_headers_btn)
        self.filter_column_label = self._add_row(cloud_form, 12, "文件夹名列", self.filter_column_row)

        cloud_layout.addLayout(cloud_form)

        cloud_actions = QHBoxLayout()
        self.download_button = QPushButton("下载")
        self.download_button.setProperty("accent", True)
        self.download_button.clicked.connect(self.start_download)
        cloud_actions.addWidget(self.download_button)
        cloud_actions.addStretch(1)
        cloud_layout.addLayout(cloud_actions)
        body_layout.addWidget(self.cloud_section)

        self.processing_section = self._create_card()
        processing_layout = QVBoxLayout(self.processing_section)
        processing_layout.setContentsMargins(18, 18, 18, 18)
        processing_layout.setSpacing(12)
        processing_title = QLabel("处理")
        processing_title.setProperty("sectionTitle", True)
        processing_layout.addWidget(processing_title)

        processing_form = QGridLayout()
        processing_form.setHorizontalSpacing(14)
        processing_form.setVerticalSpacing(14)

        self.template_edit = QLineEdit()
        self.template_label = self._add_row(
            processing_form,
            0,
            "处理模板",
            self._with_file_button(self.template_edit, self.choose_template),
        )

        self.match_combo = AppComboBox()
        load_headers_btn = QPushButton("加载列")
        load_headers_btn.clicked.connect(self.load_headers)
        self._add_row(processing_form, 1, "匹配列", self._wrap_with_button(self.match_combo, load_headers_btn))

        self.mode_combo = AppComboBox()
        self.mode_combo.addItems(["复制整个人员文件夹", "只复制关键词文件"])
        self.mode_combo.currentIndexChanged.connect(self.update_keyword_state)
        self._add_row(processing_form, 2, "筛选模式", self.mode_combo)

        self.keyword_edit = QLineEdit("学历证书")
        self._add_row(processing_form, 3, "文件关键词", self.keyword_edit)

        self.rename_checkbox = QCheckBox("导出后文件夹重命名")
        self.rename_checkbox.toggled.connect(self.update_rename_state)
        processing_form.addWidget(self.rename_checkbox, 4, 1)

        self.folder_name_combo = AppComboBox()
        self.folder_name_combo.setEnabled(False)
        self._add_row(processing_form, 5, "导出名称列", self.folder_name_combo)

        self.classify_checkbox = QCheckBox("按所选列建立目录")
        self.classify_checkbox.setChecked(True)
        self.classify_checkbox.toggled.connect(self.update_classify_state)
        processing_form.addWidget(self.classify_checkbox, 6, 1)

        self.classify_columns_list = QListWidget()
        self.classify_columns_list.setMinimumHeight(84)
        self._add_row(processing_form, 7, "分层列", self.classify_columns_list)

        self.dry_run_checkbox = QCheckBox("仅预览，不实际执行")
        processing_form.addWidget(self.dry_run_checkbox, 8, 1)

        processing_layout.addLayout(processing_form)

        processing_actions = QHBoxLayout()
        self.generate_button = QPushButton("生成模板")
        self.generate_button.clicked.connect(self.start_generate_template)
        processing_actions.addWidget(self.generate_button)
        self.run_button = QPushButton("按模板分类")
        self.run_button.clicked.connect(self.start_run)
        processing_actions.addWidget(self.run_button)
        processing_actions.addStretch(1)
        processing_layout.addLayout(processing_actions)
        body_layout.addWidget(self.processing_section)
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
        self.result_text.setMinimumHeight(180)
        self.right_card.body_layout.addWidget(self.result_text, 1)

        self._bind_cloud_cache_events()
        self._load_cached_cloud_settings()
        self._show_classify_placeholder()
        self._show_browser_placeholder()
        self.update_source_mode_state()
        self.update_filter_state()
        self.update_keyword_state()
        self.update_classify_state()

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> QLabel | None:
        label = None
        if label_text:
            label = QLabel(label_text)
            label.setProperty("formLabel", True)
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedWidth(self.LABEL_WIDTH)
            layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)
        return label

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
        selected, _ = QFileDialog.getOpenFileName(self, "选择处理模板", "", EXCEL_FILE_FILTER)
        if selected:
            self.template_edit.setText(selected)
            self.load_headers()

    def choose_filter_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择过滤表", "", EXCEL_FILE_FILTER)
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

    def _show_browser_placeholder(self, text: str = "请先选择 Bucket，再输入前缀后点击搜索。") -> None:
        self.browser_entries = []
        self.browser_listing_prefix = ""
        self.browser_list.clear()
        item = QListWidgetItem(text)
        item.setFlags(Qt.NoItemFlags)
        self.browser_list.addItem(item)
        self.browser_status_label.setText(text)
        self.selected_prefix = ""
        self.selected_path_label.setText("未选择下载路径")

    def update_source_mode_state(self) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        if self.source_label is not None:
            self.source_label.setText("下载目录" if cloud_mode else "本地目录")
        self.cloud_section.setVisible(cloud_mode)
        self.download_button.setVisible(cloud_mode)
        self.output_row.setVisible(True)
        if self.output_label is not None:
            self.output_label.setVisible(True)
        self.processing_section.setVisible(True)
        self.generate_button.setEnabled(not cloud_mode or self.cloud_download_ready)
        self.run_button.setEnabled(not cloud_mode or self.cloud_download_ready)
        self.update_filter_state()

    def update_filter_state(self) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        enabled = cloud_mode and self.filter_download_checkbox.isChecked()
        self.filter_template_edit.setEnabled(enabled)
        self.filter_column_combo.setEnabled(enabled)
        self.filter_template_edit.parentWidget().setVisible(cloud_mode)
        self.filter_column_row.setVisible(cloud_mode)
        self.filter_download_checkbox.setVisible(cloud_mode)

    def update_keyword_state(self) -> None:
        self.keyword_edit.setEnabled(self.mode_combo.currentIndex() == 1)

    def update_rename_state(self, checked: bool) -> None:
        self.folder_name_combo.setEnabled(checked)

    def update_classify_state(self) -> None:
        self.classify_columns_list.setEnabled(self.classify_checkbox.isChecked())

    def mark_cloud_download_ready(self, ready: bool) -> None:
        self.cloud_download_ready = ready
        self.update_source_mode_state()

    def build_cloud_config(self) -> OssConfig:
        cloud_type = self.cloud_type_combo.currentText().strip()
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
            cloud_type = self.cloud_type_combo.currentText().strip()
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
        worker = SimpleWorker(
            task=lambda: list_bucket_infos(
                access_key_id=access_key_id,
                access_key_secret=access_key_secret,
                endpoint=endpoint,
                cloud_type=cloud_type,
            ),
            on_success=self.on_buckets_loaded,
            on_error=self.on_bucket_load_error,
        )
        self._last_worker = worker
        self._last_worker = worker
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
        worker = SimpleWorker(
            task=lambda: list_browser_entries(config, prefix),
            on_success=lambda entries: self._accept_browser_entries_loaded(
                entries,
                request_id=request_id,
                context=context,
                prefix=prefix,
                status=status_template.format(prefix=prefix or "/"),
                is_directory_listing=True,
            ),
            on_error=self.on_worker_error,
        )
        self._last_worker = worker
        self.thread_pool.start(worker)

    def _start_global_browser_search(self, config: OssConfig, keyword: str) -> None:
        self.browser_status_label.setText("正在搜索目录...")
        request_id, context = self._start_browser_request(config)
        self._append_result_log(
            self._format_browser_debug(action=f"search-start#{request_id}", config=config, keyword=keyword)
        )
        worker = SimpleWorker(
            task=lambda: search_folder_entries(config, keyword),
            on_success=lambda entries: self._accept_browser_entries_loaded(
                entries,
                request_id=request_id,
                context=context,
                prefix=None,
                status=f"找到 {len(entries)} 个匹配条目，请在下方选择目录作为下载路径。",
                is_directory_listing=False,
            ),
            on_error=self.on_worker_error,
        )
        self._last_worker = worker
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

        self.search_button.setEnabled(False)
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
            self.search_button.setEnabled(True)
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
        self.search_button.setEnabled(True)
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
            self.browser_status_label.setText(f"已返回上一层：{self.selected_prefix or '/'}")
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
        self.mark_cloud_download_ready(False)
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
        if self.cloud_type_combo.currentText().strip() != "tencent":
            return
        location = self.bucket_combo.currentData()
        if isinstance(location, str) and location.strip():
            self.endpoint_edit.setText(location.strip())

    def _save_cached_cloud_settings(self) -> None:
        settings = self._read_settings()
        cloud_type = self.cloud_type_combo.currentText().strip() or "aliyun"
        profiles = settings.setdefault("cloud_profiles", {})
        profile = profiles.setdefault(cloud_type, {})
        profile["access_key_id"] = self.access_key_id_edit.text().strip()
        profile["access_key_secret"] = self.access_key_secret_edit.text().strip()
        profile["endpoint"] = self.endpoint_edit.text().strip()
        profile["certificate_bucket_name"] = self.bucket_combo.currentText().strip()
        profile["certificate_search_keyword"] = self.search_edit.text().strip()
        profile["certificate_prefix"] = self.selected_prefix
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
        bucket_name = profile.get("certificate_bucket_name", profile.get("bucket_name", ""))
        selected_prefix = profile.get("certificate_prefix", profile.get("prefix", ""))
        self.bucket_combo.clear()
        if bucket_name:
            self.bucket_combo.addItem(bucket_name)
        self.search_edit.setText(profile.get("certificate_search_keyword", ""))
        self._show_browser_placeholder("请输入前缀后点击搜索。")
        if selected_prefix:
            self.selected_prefix = selected_prefix
            self.selected_path_label.setText(selected_prefix)

    def _on_cloud_type_changed(self, cloud_type: str) -> None:
        self._browser_request_id += 1
        self._loading_bucket_list = True
        self._apply_cached_cloud_profile(cloud_type)
        self._loading_bucket_list = False
        self._save_cached_cloud_settings()

    def start_download(self) -> None:
        try:
            config = self.build_cloud_config()
            target_parent = Path(self.source_edit.text().strip())
            if not str(target_parent).strip():
                raise ValueError("请先选择下载目录。")
            if not self.selected_prefix:
                raise ValueError("请先从下方搜索结果中选择下载路径。")
            target_dir = build_prefixed_directory(target_parent, self.selected_prefix, "证件资料下载")
            filter_column = ""
            key_filter = None
            if self.filter_download_checkbox.isChecked():
                filter_template = Path(self.filter_template_edit.text().strip())
                if not str(filter_template).strip():
                    raise ValueError("请先选择过滤表。")
                filter_column = self.filter_column_combo.currentText().strip()
                if not filter_column:
                    raise ValueError("请先选择文件夹名列。")
                allowed_people = set(load_column_values(filter_template, filter_column))
                selected_prefix = self.selected_prefix

                def key_filter(object_key: str) -> bool:
                    relative_path = object_key
                    normalized_prefix = selected_prefix.strip().strip("/")
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
        self.reset_progress("准备下载...")
        worker = SimpleWorker(
            task=lambda: download_objects(
                config=config,
                prefix=self.selected_prefix,
                download_dir=target_dir,
                dry_run=False,
                skip_existing=True,
                logger=self.log_fn,
                progress_callback=self.make_progress_callback(),
                key_filter=key_filter,
                stage="certificate_download",
            ),
            on_success=lambda result: self.on_download_success(result, filter_column, target_dir),
            on_error=self.on_worker_error,
        )
        self._last_worker = worker
        self.thread_pool.start(worker)

    def on_download_success(self, result: DownloadResult, filter_column: str, target_dir: Path) -> None:
        self._set_running(False)
        self.update_progress("certificate_download", result.total_found, result.total_found, "")
        self.source_edit.setText(str(target_dir))
        self.mark_cloud_download_ready(True)
        if not self.output_edit.text().strip():
            self.output_edit.setText(str(target_dir.parent / "证件资料导出"))
        lines = [
            f"已选路径：{self.selected_prefix}",
            f"实际下载目录：{target_dir}",
            f"云端可下载文件：{result.total_found}",
            f"已下载：{result.downloaded_count}",
            f"跳过已存在：{result.skipped_existing_count}",
        ]
        if filter_column:
            lines.append(f"下载过滤列：{filter_column}")
        self.result_text.setPlainText("\n".join(lines))

    def start_generate_template(self) -> None:
        try:
            source_dir = Path(self.source_edit.text().strip())
            if not str(source_dir).strip():
                raise ValueError("请先选择本地目录。")
            output_text = self.output_edit.text().strip() or self.source_edit.text().strip()
            output_dir = ensure_child_directory(Path(output_text), "证件资料导出")
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self._set_running(True)
        self.reset_progress("正在生成模板...")
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
        self._last_worker = worker
        self.thread_pool.start(worker)

    def start_run(self) -> None:
        try:
            source_dir = Path(self.source_edit.text().strip())
            output_dir = Path(self.output_edit.text().strip())
            if not str(source_dir).strip():
                raise ValueError("请先选择本地目录。")
            if not str(output_dir).strip():
                raise ValueError("请先选择输出目录。")
            output_dir = ensure_child_directory(output_dir, "证件资料导出")
            classify_columns = self._selected_classify_columns() if self.classify_checkbox.isChecked() else []
            options = CertificateFilterOptions(
                template_path=Path(self.template_edit.text().strip()),
                source_dir=source_dir,
                output_dir=output_dir,
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
        self.reset_progress("正在筛选...")
        worker = SimpleWorker(
            task=lambda: run_certificate_filter(
                options,
                logger=self.log_fn,
                progress_callback=self.make_progress_callback(),
            ),
            on_success=self.on_success,
            on_error=self.on_error,
        )
        self._last_worker = worker
        self.thread_pool.start(worker)

    def _set_running(self, running: bool) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        can_process = not cloud_mode or self.cloud_download_ready
        self.download_button.setEnabled(not running)
        self.generate_button.setEnabled((not running) and can_process)
        self.run_button.setEnabled((not running) and can_process)
        self.search_button.setEnabled(not running)

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
        if stage == "certificate":
            self.progress_label.setText(f"筛选进度：{current}/{total} {filename}".strip() if total else "正在读取模板...")
            return
        if total:
            self.progress_label.setText(f"下载进度：{current}/{total} {filename}".strip())
        else:
            self.progress_label.setText("正在统计下载文件...")

    def on_template_success(self, summary: CertificateTemplateSummary) -> None:
        self._set_running(False)
        self.progress_label.setText("模板生成完成")
        self.template_edit.setText(str(summary.template_path))
        self.load_headers()
        self.result_text.setPlainText(f"已生成模板：{summary.template_path}")

    def on_success(self, summary: CertificateFilterSummary) -> None:
        self._set_running(False)
        self.progress_label.setText("筛选完成")
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
        self.progress_label.setText("任务失败")
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "按模板分类失败", error_text.splitlines()[0])

    def on_worker_error(self, error_text: str) -> None:
        self._set_running(False)
        self.progress_label.setText("任务失败")
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])

    def on_bucket_load_error(self, error_text: str) -> None:
        self._loading_bucket_list = False
        self.load_bucket_button.setEnabled(True)
        self.on_worker_error(error_text)
