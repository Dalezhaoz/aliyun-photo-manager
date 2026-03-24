from __future__ import annotations

import traceback
from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt, QObject, Signal, QRunnable, QThreadPool
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

from ..certificate_filter import (
    CertificateFilterOptions,
    CertificateFilterSummary,
    list_template_headers,
    load_match_values,
    run_certificate_filter,
)
from ..config import OssConfig, validate_oss_config
from ..downloader import DownloadResult, download_objects, list_browser_entries, list_buckets


class WorkerSignals(QObject):
    success = Signal(object)
    error = Signal(str)


class CertificateWorker(QRunnable):
    def __init__(
        self,
        options: CertificateFilterOptions,
        log_fn: Callable[[str], None],
        on_success: Callable[[CertificateFilterSummary], None],
        on_error: Callable[[str], None],
    ) -> None:
        super().__init__()
        self.options = options
        self.log_fn = log_fn
        self.signals = WorkerSignals()
        self.signals.success.connect(on_success)
        self.signals.error.connect(on_error)

    def run(self) -> None:
        try:
            summary = run_certificate_filter(
                self.options,
                logger=self.log_fn,
            )
        except Exception as exc:
            self.signals.error.emit(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}")
        else:
            self.signals.success.emit(summary)


class CertificatePage(QWidget):
    LABEL_WIDTH = 132
    ACTION_WIDTH = 144

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.thread_pool = QThreadPool.globalInstance()
        self.certificate_headers: list[str] = []
        self.current_prefix = ""
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._create_card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)
        title = QLabel("证件资料筛选")
        title.setProperty("heroTitle", True)
        intro = QLabel("支持本地目录筛选，也支持先从云存储按模板匹配值下载证件资料，再执行筛选。")
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
        section_title = QLabel("筛选参数")
        section_title.setProperty("sectionTitle", True)
        left_layout.addWidget(section_title)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)
        form.setColumnStretch(0, 0)
        form.setColumnStretch(1, 1)
        row_index = 0

        self.source_mode_combo = QComboBox()
        self.source_mode_combo.addItems(["本地目录", "云存储下载后处理"])
        self.source_mode_combo.currentIndexChanged.connect(self.update_source_mode_state)
        self._add_form_row(form, row_index, "数据来源", self.source_mode_combo)
        row_index += 1

        self.template_edit = QLineEdit()
        template_row = self._with_file_button(self.template_edit, self.choose_template)
        self._add_form_row(form, row_index, "人员模板", template_row)
        row_index += 1

        self.match_combo = QComboBox()
        self.match_combo.setEditable(False)
        load_headers_btn = QPushButton("加载模板列")
        load_headers_btn.clicked.connect(self.load_headers)
        match_row = QWidget()
        match_layout = QHBoxLayout(match_row)
        match_layout.setContentsMargins(0, 0, 0, 0)
        match_layout.addWidget(self.match_combo, 1)
        match_layout.addWidget(load_headers_btn)
        self._add_form_row(form, row_index, "匹配列", match_row)
        row_index += 1

        self.source_edit = QLineEdit()
        source_row = self._with_dir_button(self.source_edit)
        self._add_form_row(form, row_index, "证件资料目录", source_row)
        row_index += 1

        self.output_edit = QLineEdit()
        output_row = self._with_dir_button(self.output_edit)
        self._add_form_row(form, row_index, "输出目录", output_row)
        row_index += 1

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["复制整个人员文件夹", "只复制关键词文件"])
        self.mode_combo.currentIndexChanged.connect(self.update_keyword_state)
        self._add_form_row(form, row_index, "筛选模式", self.mode_combo)
        row_index += 1

        self.keyword_edit = QLineEdit("学历证书")
        self._add_form_row(form, row_index, "文件关键词", self.keyword_edit)
        row_index += 1

        self.folder_name_combo = QComboBox()
        self.folder_name_combo.setEditable(False)
        self.folder_name_combo.setEnabled(False)
        self._add_form_row(form, row_index, "名称列", self.folder_name_combo)
        row_index += 1

        left_layout.addLayout(form)

        self.cloud_section = QWidget()
        cloud_section_layout = QVBoxLayout(self.cloud_section)
        cloud_section_layout.setContentsMargins(0, 0, 0, 0)
        cloud_section_layout.setSpacing(12)
        self.cloud_section_title = QLabel("云存储下载")
        self.cloud_section_title.setProperty("sectionTitle", True)
        cloud_section_layout.addWidget(self.cloud_section_title)

        cloud_form = QGridLayout()
        cloud_form.setHorizontalSpacing(14)
        cloud_form.setVerticalSpacing(14)

        self.cloud_type_combo = QComboBox()
        self.cloud_type_combo.addItems(["aliyun", "tencent"])
        self._add_form_row(cloud_form, 0, "云类型", self.cloud_type_combo)

        self.endpoint_edit = QLineEdit()
        self.endpoint_edit.setPlaceholderText("阿里云填 Endpoint，腾讯云填 Region 或 Endpoint")
        self._add_form_row(cloud_form, 1, "Endpoint / Region", self.endpoint_edit)

        self.access_key_id_edit = QLineEdit()
        self._add_form_row(cloud_form, 2, "AccessKey ID", self.access_key_id_edit)

        self.access_key_secret_edit = QLineEdit()
        self.access_key_secret_edit.setEchoMode(QLineEdit.Password)
        self._add_form_row(cloud_form, 3, "AccessKey Secret", self.access_key_secret_edit)

        bucket_row = QWidget()
        bucket_layout = QHBoxLayout(bucket_row)
        bucket_layout.setContentsMargins(0, 0, 0, 0)
        bucket_layout.setSpacing(10)
        self.bucket_combo = QComboBox()
        self.bucket_combo.setEditable(False)
        self.bucket_combo.currentTextChanged.connect(self.on_bucket_changed)
        bucket_layout.addWidget(self.bucket_combo, 1)
        self.load_bucket_button = QPushButton("加载 Bucket")
        self.load_bucket_button.setFixedWidth(self.ACTION_WIDTH)
        self.load_bucket_button.clicked.connect(self.load_buckets)
        bucket_layout.addWidget(self.load_bucket_button)
        self._add_form_row(cloud_form, 4, "Bucket", bucket_row)

        prefix_row = QWidget()
        prefix_layout = QHBoxLayout(prefix_row)
        prefix_layout.setContentsMargins(0, 0, 0, 0)
        prefix_layout.setSpacing(10)
        self.prefix_edit = QLineEdit()
        self.prefix_edit.setPlaceholderText("当前前缀，留空表示根目录")
        prefix_layout.addWidget(self.prefix_edit, 1)
        self.load_prefix_button = QPushButton("加载当前层级")
        self.load_prefix_button.setFixedWidth(self.ACTION_WIDTH)
        self.load_prefix_button.clicked.connect(self.load_browser_entries)
        prefix_layout.addWidget(self.load_prefix_button)
        self._add_form_row(cloud_form, 5, "当前前缀", prefix_row)

        browser_row = QWidget()
        browser_layout = QHBoxLayout(browser_row)
        browser_layout.setContentsMargins(0, 0, 0, 0)
        browser_layout.setSpacing(10)
        self.browser_combo = QComboBox()
        self.browser_combo.setEditable(False)
        browser_layout.addWidget(self.browser_combo, 1)
        self.enter_folder_button = QPushButton("进入目录")
        self.enter_folder_button.setFixedWidth(self.ACTION_WIDTH)
        self.enter_folder_button.clicked.connect(self.enter_selected_folder)
        browser_layout.addWidget(self.enter_folder_button)
        self._add_form_row(cloud_form, 6, "子目录", browser_row)

        nav_row = QWidget()
        nav_layout = QHBoxLayout(nav_row)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(10)
        self.back_prefix_button = QPushButton("返回上级")
        self.back_prefix_button.setFixedWidth(self.ACTION_WIDTH)
        self.back_prefix_button.clicked.connect(self.go_parent_prefix)
        nav_layout.addWidget(self.back_prefix_button)
        self.download_button = QPushButton("下载证件资料")
        self.download_button.setFixedWidth(self.ACTION_WIDTH)
        self.download_button.clicked.connect(self.start_cloud_download)
        nav_layout.addWidget(self.download_button)
        nav_layout.addStretch(1)
        self._add_form_row(cloud_form, 7, "", nav_row)

        cloud_section_layout.addLayout(cloud_form)
        left_layout.addWidget(self.cloud_section)

        option_title = QLabel("执行选项")
        option_title.setProperty("sectionTitle", True)
        left_layout.addWidget(option_title)

        option_box = QWidget()
        option_layout = QVBoxLayout(option_box)
        option_layout.setContentsMargins(0, 0, 0, 0)
        option_layout.setSpacing(10)

        self.rename_checkbox = QCheckBox("导出后文件夹重命名")
        self.rename_checkbox.toggled.connect(self.update_rename_state)
        option_layout.addWidget(self.rename_checkbox)

        self.classify_checkbox = QCheckBox("按分类一 / 分类二 / 分类三建立目录")
        self.classify_checkbox.setChecked(True)
        option_layout.addWidget(self.classify_checkbox)

        self.dry_run_checkbox = QCheckBox("仅预览，不实际执行")
        option_layout.addWidget(self.dry_run_checkbox)
        left_layout.addWidget(option_box)
        left_layout.addStretch(1)

        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 8, 0, 0)
        self.run_button = QPushButton("开始筛选")
        self.run_button.setProperty("accent", True)
        self.run_button.setFixedWidth(150)
        self.run_button.clicked.connect(self.start_run)
        action_row.addWidget(self.run_button)
        action_row.addStretch(1)
        left_layout.addLayout(action_row)
        body.addWidget(left, 0, 0)

        right = self._create_card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_layout.setSpacing(14)
        summary_title = QLabel("筛选结果")
        summary_title.setProperty("sectionTitle", True)
        right_layout.addWidget(summary_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        right_layout.addWidget(self.result_text, 1)
        body.addWidget(right, 0, 1)

        body.setColumnStretch(0, 3)
        body.setColumnStretch(1, 2)
        self.update_keyword_state()
        self.update_source_mode_state()

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_form_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> None:
        if label_text:
            label = QLabel(label_text)
            label.setProperty("formLabel", True)
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedWidth(self.LABEL_WIDTH)
            layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

    def _with_file_button(self, line_edit: QLineEdit, callback) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        layout.addWidget(line_edit, 1)
        button = QPushButton("选择文件")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(callback)
        layout.addWidget(button)
        return row

    def _with_dir_button(self, line_edit: QLineEdit) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        layout.addWidget(line_edit, 1)
        button = QPushButton("选择目录")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(lambda: self.choose_directory(line_edit))
        layout.addWidget(button)
        return row

    def choose_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择人员模板", "", "Excel 文件 (*.xlsx)")
        if selected:
            self.template_edit.setText(selected)
            self.load_headers()

    def choose_directory(self, line_edit: QLineEdit) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择目录", line_edit.text().strip() or str(Path.cwd()))
        if selected:
            line_edit.setText(selected)

    def load_headers(self) -> None:
        template_path = self.template_edit.text().strip()
        if not template_path:
            QMessageBox.information(self, "缺少模板", "请先选择人员模板文件。")
            return
        try:
            headers = list_template_headers(Path(template_path))
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            return
        self.certificate_headers = headers
        self.match_combo.clear()
        self.folder_name_combo.clear()
        self.match_combo.addItems(headers)
        self.folder_name_combo.addItems(headers)
        self.log_fn(f"已读取模板列：{', '.join(headers) if headers else '无'}")

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

    def update_source_mode_state(self) -> None:
        cloud_mode = self.source_mode_combo.currentIndex() == 1
        self.cloud_section.setVisible(cloud_mode)
        for widget in (
            self.cloud_type_combo,
            self.endpoint_edit,
            self.access_key_id_edit,
            self.access_key_secret_edit,
            self.bucket_combo,
            self.load_bucket_button,
            self.prefix_edit,
            self.load_prefix_button,
            self.browser_combo,
            self.enter_folder_button,
            self.back_prefix_button,
            self.download_button,
        ):
            widget.setVisible(cloud_mode)
            widget.setEnabled(cloud_mode)

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
        self.result_text.setPlainText("正在加载 bucket 列表...")

        def task() -> list[str]:
            return list_buckets(
                access_key_id=access_key_id,
                access_key_secret=access_key_secret,
                endpoint=endpoint,
                cloud_type=cloud_type,
            )

        worker = SimpleWorker(
            task=task,
            on_success=self.on_buckets_loaded,
            on_error=self.on_worker_error,
        )
        self.thread_pool.start(worker)

    def on_buckets_loaded(self, buckets: list[str]) -> None:
        self.load_bucket_button.setEnabled(True)
        self.bucket_combo.clear()
        self.bucket_combo.addItems(buckets)
        self.result_text.setPlainText(f"已加载 {len(buckets)} 个 bucket。")
        self.log_fn(f"证件资料 bucket 加载完成，共 {len(buckets)} 个。")

    def on_bucket_changed(self, _value: str) -> None:
        self.current_prefix = ""
        self.prefix_edit.setText("")
        self.browser_combo.clear()

    def load_browser_entries(self) -> None:
        try:
            config = self.build_cloud_config()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return
        prefix = self.prefix_edit.text().strip()
        self.load_prefix_button.setEnabled(False)
        self.result_text.setPlainText(f"正在加载当前层级：{prefix or '/'}")

        worker = SimpleWorker(
            task=lambda: list_browser_entries(config, prefix),
            on_success=self.on_browser_entries_loaded,
            on_error=self.on_worker_error,
        )
        self.thread_pool.start(worker)

    def on_browser_entries_loaded(self, entries: list) -> None:
        self.load_prefix_button.setEnabled(True)
        self.browser_combo.clear()
        folder_entries = [entry for entry in entries if entry.entry_type == "folder"]
        for entry in folder_entries:
            self.browser_combo.addItem(entry.display_name, entry.key)
        self.result_text.setPlainText(
            f"当前层级共 {len([entry for entry in entries if entry.entry_type == 'folder'])} 个文件夹，"
            f"{len([entry for entry in entries if entry.entry_type == 'file'])} 个文件。"
        )

    def enter_selected_folder(self) -> None:
        next_prefix = self.browser_combo.currentData()
        if not next_prefix:
            return
        self.current_prefix = str(next_prefix)
        self.prefix_edit.setText(self.current_prefix)
        self.load_browser_entries()

    def go_parent_prefix(self) -> None:
        current = self.prefix_edit.text().strip().strip("/")
        if not current:
            return
        parts = current.split("/")
        parent = "/".join(parts[:-1])
        self.current_prefix = parent + "/" if parent else ""
        self.prefix_edit.setText(self.current_prefix)
        self.load_browser_entries()

    def start_cloud_download(self) -> None:
        try:
            config = self.build_cloud_config()
            template_path = Path(self.template_edit.text().strip())
            match_column = self.match_combo.currentText().strip()
            target_dir_text = self.source_edit.text().strip()
            if not target_dir_text:
                raise ValueError("请先选择本地证件资料目录。")
            target_dir = Path(target_dir_text)
            if not template_path.exists():
                raise ValueError("请先选择有效的人员模板。")
            if not match_column:
                raise ValueError("请先加载模板列并选择匹配列。")
            match_values = load_match_values(template_path, match_column)
            if not match_values:
                raise ValueError("模板匹配列没有可用数据，无法下载证件资料。")
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        allowed_people = set(match_values)
        prefix = self.prefix_edit.text().strip()
        target_dir.mkdir(parents=True, exist_ok=True)
        self.download_button.setEnabled(False)
        self.result_text.setPlainText("正在从云端下载证件资料，请稍候...")
        self.log_fn(f"启动证件资料云下载：{prefix or '/'} -> {target_dir}")

        def key_filter(object_key: str) -> bool:
            relative_path = object_key
            normalized_prefix = prefix.strip().strip("/")
            if normalized_prefix:
                prefix_with_slash = normalized_prefix + "/"
                if object_key.startswith(prefix_with_slash):
                    relative_path = object_key[len(prefix_with_slash):]
            relative_path = relative_path.lstrip("/")
            if not relative_path:
                return False
            parts = Path(relative_path).parts
            return bool(parts) and parts[0] in allowed_people

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
            on_success=self.on_download_success,
            on_error=self.on_worker_error,
        )
        self.thread_pool.start(worker)

    def on_download_success(self, result: DownloadResult) -> None:
        self.download_button.setEnabled(True)
        lines = [
            f"云端可下载文件：{result.total_found}",
            f"已下载：{result.downloaded_count}",
            f"跳过已存在：{result.skipped_existing_count}",
            f"本地证件资料目录：{self.source_edit.text().strip()}",
        ]
        self.result_text.setPlainText("\n".join(lines))
        self.log_fn("证件资料云下载完成。")

    def update_keyword_state(self) -> None:
        self.keyword_edit.setEnabled(self.mode_combo.currentIndex() == 1)

    def update_rename_state(self, checked: bool) -> None:
        self.folder_name_combo.setEnabled(checked)

    def start_run(self) -> None:
        try:
            options = self.build_options()
        except Exception as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        self.run_button.setEnabled(False)
        self.result_text.setPlainText("正在筛选，请稍候...")
        self.log_fn(f"启动证件资料筛选：{options.source_dir} -> {options.output_dir}")
        worker = CertificateWorker(options, self.log_fn, self.on_success, self.on_error)
        self.thread_pool.start(worker)

    def build_options(self) -> CertificateFilterOptions:
        template_path = Path(self.template_edit.text().strip())
        source_dir = Path(self.source_edit.text().strip())
        output_dir = Path(self.output_edit.text().strip())
        match_column = self.match_combo.currentText().strip()
        if not template_path.exists():
            raise ValueError("请选择有效的人员模板。")
        if not source_dir.exists():
            raise ValueError("请选择有效的证件资料目录。")
        if not match_column:
            raise ValueError("请先加载模板列并选择匹配列。")

        keyword = self.keyword_edit.text().strip() if self.mode_combo.currentIndex() == 1 else ""
        if self.mode_combo.currentIndex() == 1 and not keyword:
            raise ValueError("关键词模式下请输入文件关键词。")

        folder_name_column = self.folder_name_combo.currentText().strip() if self.rename_checkbox.isChecked() else ""
        if self.rename_checkbox.isChecked() and not folder_name_column:
            raise ValueError("已勾选重命名，请选择名称列。")

        return CertificateFilterOptions(
            template_path=template_path,
            source_dir=source_dir,
            output_dir=output_dir,
            match_column=match_column,
            rename_folder=self.rename_checkbox.isChecked(),
            folder_name_column=folder_name_column,
            classify_output=self.classify_checkbox.isChecked(),
            keyword=keyword,
            dry_run=self.dry_run_checkbox.isChecked(),
        )

    def on_success(self, summary: CertificateFilterSummary) -> None:
        self.run_button.setEnabled(True)
        lines = [
            f"模板：{summary.template_path}",
            f"来源目录：{summary.source_dir}",
            f"输出目录：{summary.output_dir}",
            f"模板有效行数：{summary.total_rows}",
            f"命中人员：{summary.matched_people}",
            f"缺失人员：{summary.missing_people}",
            f"复制人数：{summary.copied_people}",
            f"复制文件：{summary.copied_files}",
        ]
        if summary.report_path is not None:
            lines.append(f"结果清单：{summary.report_path}")
        self.result_text.setPlainText("\n".join(lines))

    def on_error(self, error_text: str) -> None:
        self.run_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "筛选失败", error_text.splitlines()[0])

    def on_worker_error(self, error_text: str) -> None:
        self.run_button.setEnabled(True)
        self.download_button.setEnabled(True)
        self.load_bucket_button.setEnabled(True)
        self.load_prefix_button.setEnabled(True)
        self.result_text.setPlainText(error_text)
        QMessageBox.critical(self, "执行失败", error_text.splitlines()[0])


class SimpleWorker(QRunnable):
    def __init__(
        self,
        task: Callable[[], object],
        on_success: Callable[[object], None],
        on_error: Callable[[str], None],
    ) -> None:
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
