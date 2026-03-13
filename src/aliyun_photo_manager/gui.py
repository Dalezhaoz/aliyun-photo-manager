import queue
import threading
import tkinter as tk
import json
import os
import socket
import subprocess
import sys
import tempfile
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from .app import (
    RunOptions,
    WorkflowSummary,
    build_prefixed_directory,
    resolve_photo_directories,
    run_photo_classification_only,
    run_photo_download_and_template,
)
from . import __version__
from .certificate_filter import (
    CertificateFilterOptions,
    CertificateFilterSummary,
    list_template_headers,
    load_match_values,
    run_certificate_filter,
)
from .config import OssConfig, validate_oss_config, validate_oss_credentials
from .downloader import (
    BrowserEntry,
    count_photos_in_prefix,
    download_objects,
    is_photo_key,
    list_buckets,
    list_browser_entries,
    list_folder_prefixes,
)
from .excel_classifier import generate_template
from .result_packer import PackSummary, pack_encrypted_folder, query_pack_history
from .data_matcher import (
    ColumnMapping,
    DataMatchOptions,
    DataMatchSummary,
    list_headers as list_match_headers,
    run_data_match,
)
from .exam_arranger import (
    ExamArrangeOptions,
    ExamArrangeSummary,
    ExamRuleItem,
    ExamTemplateExportSummary,
    export_exam_templates,
    list_headers as list_exam_headers,
    run_exam_arrangement,
)
from .word_to_html import WordExportResult, export_word_to_html


class App:
    SETTINGS_FILE = Path(__file__).resolve().parents[2] / ".gui_settings.json"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"报名系统工具箱 v{__version__}")
        self.root.geometry("1080x760")
        self.root.minsize(920, 680)

        self.log_queue: "queue.Queue[object]" = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.cancel_event = threading.Event()

        self.prefix_var = tk.StringVar()
        self.photo_source_mode_var = tk.StringVar(value="oss")
        self.photo_template_var = tk.StringVar()
        self.photo_match_column_var = tk.StringVar()
        self.photo_filter_by_template_var = tk.BooleanVar(value=False)
        self.search_keyword_var = tk.StringVar()
        self.bucket_values = []
        self.folder_values = [""]
        self.download_dir_var = tk.StringVar(value=str(Path.cwd() / "downloads"))
        self.sorted_dir_var = tk.StringVar(value=str(Path.cwd() / "sorted"))
        self.cloud_type_var = tk.StringVar(value="aliyun")
        self.access_key_id_var = tk.StringVar()
        self.access_key_secret_var = tk.StringVar()
        self.endpoint_var = tk.StringVar(value="https://oss-cn-hangzhou.aliyuncs.com")
        self.bucket_name_var = tk.StringVar()

        self.skip_download_var = tk.BooleanVar(value=False)
        self.flat_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=True)
        self.include_duplicates_var = tk.BooleanVar(value=False)
        self.move_sorted_files_var = tk.BooleanVar(value=False)
        self.skip_existing_var = tk.BooleanVar(value=True)
        self.certificate_template_var = tk.StringVar()
        self.certificate_source_mode_var = tk.StringVar(value="local")
        self.certificate_source_dir_var = tk.StringVar()
        self.certificate_output_dir_var = tk.StringVar()
        self.certificate_match_column_var = tk.StringVar()
        self.certificate_rename_folder_var = tk.BooleanVar(value=False)
        self.certificate_folder_name_column_var = tk.StringVar()
        self.certificate_keyword_var = tk.StringVar(value="学历证书")
        self.certificate_classify_var = tk.BooleanVar(value=True)
        self.certificate_dry_run_var = tk.BooleanVar(value=False)
        self.certificate_mode_var = tk.StringVar(value="folder")
        self.certificate_bucket_name_var = tk.StringVar()
        self.certificate_prefix_var = tk.StringVar()
        self.word_source_var = tk.StringVar()
        self.pack_source_dir_var = tk.StringVar()
        self.pack_output_dir_var = tk.StringVar(value=str(Path.cwd() / "packed_results"))
        self.pack_use_custom_password_var = tk.BooleanVar(value=False)
        self.pack_password_var = tk.StringVar()
        self.pack_query_var = tk.StringVar()
        self.match_target_var = tk.StringVar()
        self.match_source_var = tk.StringVar()
        self.match_target_key_var = tk.StringVar()
        self.match_source_key_var = tk.StringVar()
        self.match_output_var = tk.StringVar()
        self.match_extra_target_var = tk.StringVar()
        self.match_extra_source_var = tk.StringVar()
        self.match_transfer_target_var = tk.StringVar()
        self.match_transfer_source_var = tk.StringVar()
        self.exam_candidate_var = tk.StringVar()
        self.exam_group_var = tk.StringVar()
        self.exam_plan_var = tk.StringVar()
        self.exam_output_var = tk.StringVar()
        self.exam_point_digits_var = tk.StringVar(value="2")
        self.exam_room_digits_var = tk.StringVar(value="3")
        self.exam_seat_digits_var = tk.StringVar(value="2")
        self.exam_serial_digits_var = tk.StringVar(value="4")
        self.exam_sort_mode_var = tk.StringVar(value="random")
        self.exam_rule_type_var = tk.StringVar(value="自定义")
        self.exam_rule_custom_var = tk.StringVar()

        self.status_var = tk.StringVar(value="就绪")
        self.progress_text_var = tk.StringVar(value="未开始")
        self.summary_text_var = tk.StringVar(value="任务结果会显示在这里")
        self.certificate_status_var = tk.StringVar(value="未开始证件资料筛选")
        self.certificate_progress_text_var = tk.StringVar(value="未开始")
        self.certificate_summary_text_var = tk.StringVar(value="证件资料筛选结果会显示在这里")
        self.bucket_status_var = tk.StringVar(value="未加载 bucket 列表")
        self.folder_status_var = tk.StringVar(value="未加载 bucket 文件夹")
        self.search_status_var = tk.StringVar(value="未搜索文件夹")
        self.selected_folder_info_var = tk.StringVar(value="当前未选择 bucket 文件夹")
        self.certificate_bucket_status_var = tk.StringVar(value="未加载 bucket 列表")
        self.certificate_folder_status_var = tk.StringVar(value="未加载 bucket 文件夹")
        self.certificate_search_status_var = tk.StringVar(value="未搜索文件夹")
        self.certificate_selected_folder_info_var = tk.StringVar(value="当前未选择 bucket 文件夹")
        self.word_status_var = tk.StringVar(value="未开始导出")
        self.word_result_var = tk.StringVar(value="表样转换结果会显示在这里")
        self.word_preview_status_var = tk.StringVar(value="未生成预览")
        self.pack_status_var = tk.StringVar(value="未开始打包")
        self.pack_result_var = tk.StringVar(value="结果打包信息会显示在这里")
        self.match_status_var = tk.StringVar(value="未开始匹配")
        self.match_result_var = tk.StringVar(value="数据匹配结果会显示在这里")
        self.exam_status_var = tk.StringVar(value="未开始编排")
        self.exam_result_var = tk.StringVar(value="考场编排结果会显示在这里")
        self.folder_tree: Optional[ttk.Treeview] = None
        self.certificate_folder_tree: Optional[ttk.Treeview] = None
        self.folder_nodes: Dict[str, BrowserEntry] = {}
        self.certificate_folder_nodes: Dict[str, BrowserEntry] = {}
        self.current_folder_entries: List[BrowserEntry] = []
        self.current_certificate_folder_entries: List[BrowserEntry] = []
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.certificate_progress_bar: Optional[ttk.Progressbar] = None
        self.last_summary: Optional[WorkflowSummary] = None
        self.last_certificate_summary: Optional[CertificateFilterSummary] = None
        self.last_word_export: Optional[WordExportResult] = None
        self.last_pack_summary: Optional[PackSummary] = None
        self.last_match_summary: Optional[DataMatchSummary] = None
        self.last_exam_summary: Optional[ExamArrangeSummary] = None
        self.last_exam_template_export: Optional[ExamTemplateExportSummary] = None
        self.word_code_text = None
        self.word_preview_widget = None
        self.match_result_text = None
        self.pack_query_result_text = None
        self.exam_result_text = None
        self.certificate_headers: List[str] = []
        self.photo_headers: List[str] = []
        self.match_target_headers: List[str] = []
        self.match_source_headers: List[str] = []
        self.match_extra_mappings: List[ColumnMapping] = []
        self.match_transfer_mappings: List[ColumnMapping] = []
        self.exam_rule_items: List[ExamRuleItem] = []
        self.exam_group_headers: List[str] = []
        self.exam_rule_base_types = ["自定义", "考点", "考场", "座号", "流水号"]
        self.certificate_bucket_values: List[str] = []
        self.certificate_folder_values: List[str] = [""]
        self.default_sash_pending = {"photo": True, "certificate": True}
        self.bucket_load_token = 0
        self.certificate_bucket_load_token = 0
        self.cloud_profiles: Dict[str, Dict[str, str]] = {
            "aliyun": self.default_cloud_profile("aliyun"),
            "tencent": self.default_cloud_profile("tencent"),
        }

        self.load_saved_settings()
        self.build_ui()
        if self.photo_template_var.get().strip() and Path(
            self.photo_template_var.get().strip()
        ).exists():
            self.load_photo_headers()
        if self.certificate_template_var.get().strip() and Path(
            self.certificate_template_var.get().strip()
        ).exists():
            self.load_certificate_headers()
        if (
            self.match_target_var.get().strip()
            and self.match_source_var.get().strip()
            and Path(self.match_target_var.get().strip()).exists()
            and Path(self.match_source_var.get().strip()).exists()
        ):
            self.load_match_headers()
        self.root.after(150, self.flush_logs)

    def build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        title = ttk.Label(
            container,
            text=f"报名系统工具箱 v{__version__}",
            font=("Helvetica", 16, "bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        notebook = ttk.Notebook(container)
        notebook.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.notebook = notebook
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        settings_frame = ttk.Frame(notebook, padding=14)
        settings_frame.columnconfigure(0, weight=1)
        settings_frame.rowconfigure(0, weight=1)
        notebook.add(settings_frame, text="照片下载与分类")

        photo_paned = ttk.Panedwindow(settings_frame, orient="horizontal")
        photo_paned.grid(row=0, column=0, sticky="nsew")
        self.photo_paned = photo_paned

        left_outer = ttk.Frame(photo_paned)
        left_outer.columnconfigure(0, weight=1)
        left_outer.rowconfigure(0, weight=1)

        left_canvas = tk.Canvas(left_outer, highlightthickness=0)
        left_canvas.grid(row=0, column=0, sticky="nsew")

        left_scrollbar = ttk.Scrollbar(left_outer, orient="vertical", command=left_canvas.yview)
        left_scrollbar.grid(row=0, column=1, sticky="ns")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        left_frame = ttk.Frame(left_canvas)
        left_frame.columnconfigure(0, weight=1)

        left_window = left_canvas.create_window((0, 0), window=left_frame, anchor="nw")

        def sync_left_scroll(_event=None) -> None:
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def sync_left_width(event) -> None:
            left_canvas.itemconfigure(left_window, width=event.width)

        left_frame.bind("<Configure>", sync_left_scroll)
        left_canvas.bind("<Configure>", sync_left_width)
        self.bind_mousewheel_to_canvas(left_canvas, left_frame)

        self.photo_right_frame = ttk.Frame(settings_frame)
        self.photo_right_frame.columnconfigure(0, weight=1)
        self.photo_right_frame.rowconfigure(0, weight=1)
        photo_paned.add(left_outer, weight=1)
        photo_paned.add(self.photo_right_frame, weight=1)

        photo_source_frame = ttk.LabelFrame(left_frame, text="数据来源", padding=12)
        photo_source_frame.grid(row=0, column=0, sticky="ew")
        ttk.Radiobutton(
            photo_source_frame,
            text="从云存储下载后处理",
            variable=self.photo_source_mode_var,
            value="oss",
            command=self.update_photo_source_mode_ui,
        ).pack(side="left")
        ttk.Radiobutton(
            photo_source_frame,
            text="直接处理本地目录",
            variable=self.photo_source_mode_var,
            value="local",
            command=self.update_photo_source_mode_ui,
        ).pack(side="left", padx=(10, 0))

        paths_frame = ttk.LabelFrame(left_frame, text="目录", padding=12)
        paths_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        paths_frame.columnconfigure(1, weight=1)

        self.add_path_row(
            paths_frame,
            row=0,
            label="下载目录",
            variable=self.download_dir_var,
        )
        self.add_path_row(
            paths_frame,
            row=1,
            label="分类目录",
            variable=self.sorted_dir_var,
        )

        photo_template_frame = ttk.LabelFrame(left_frame, text="名单下载", padding=12)
        photo_template_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        photo_template_frame.columnconfigure(1, weight=1)
        self.add_file_row(
            photo_template_frame,
            row=0,
            label="人员模板",
            variable=self.photo_template_var,
            command=self.choose_photo_template,
            button_text="选择文件",
        )
        ttk.Label(photo_template_frame, text="匹配列", width=16).grid(row=1, column=0, sticky="w", pady=4)
        photo_match_row = ttk.Frame(photo_template_frame)
        photo_match_row.grid(row=1, column=1, columnspan=2, sticky="ew", pady=4)
        photo_match_row.columnconfigure(0, weight=1)
        self.photo_match_combo = ttk.Combobox(
            photo_match_row,
            textvariable=self.photo_match_column_var,
            values=self.photo_headers,
            state="readonly",
        )
        self.photo_match_combo.grid(row=0, column=0, sticky="ew")
        ttk.Button(photo_match_row, text="加载模板列", command=self.load_photo_headers).grid(
            row=0, column=1, padx=(8, 0)
        )
        self.add_tick_checkbutton(
            photo_template_frame,
            text="只下载模板中的人员",
            variable=self.photo_filter_by_template_var,
        ).grid(row=2, column=1, sticky="w", pady=(6, 0))

        self.photo_oss_frame = ttk.LabelFrame(left_frame, text="OSS 配置", padding=12)
        self.photo_oss_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        self.photo_oss_frame.columnconfigure(1, weight=1)

        self.add_cloud_type_row(self.photo_oss_frame, 0)
        self.add_entry_row(self.photo_oss_frame, 1, "Endpoint / Region", self.endpoint_var)
        self.add_entry_row(self.photo_oss_frame, 2, "AccessKey ID", self.access_key_id_var)
        self.add_entry_row(
            self.photo_oss_frame,
            3,
            "AccessKey Secret",
            self.access_key_secret_var,
            show="*",
        )
        self.add_bucket_picker(self.photo_oss_frame, 4)
        self.add_prefix_picker(self.photo_oss_frame, 5)

        options_frame = ttk.LabelFrame(left_frame, text="选项", padding=12)
        options_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))

        option_specs = [
            ("仅预览，不实际执行", self.dry_run_var),
            ("下载时跳过已存在文件", self.skip_existing_var),
        ]
        for row, (label, variable) in enumerate(option_specs):
            self.add_tick_checkbutton(options_frame, text=label, variable=variable).grid(
                row=row // 2,
                column=row % 2,
                sticky="w",
                padx=(0, 18),
                pady=4,
            )

        action_frame = ttk.Frame(left_frame)
        action_frame.grid(row=5, column=0, sticky="ew", pady=(14, 0))

        self.run_button = ttk.Button(
            action_frame,
            text="下载并生成模板",
            command=self.start_photo_download_run,
        )
        self.run_button.pack(side="left")

        self.photo_classify_button = ttk.Button(
            action_frame,
            text="按模板分类",
            command=self.start_photo_classify_run,
        )
        self.photo_classify_button.pack(side="left", padx=(10, 0))

        self.cancel_button = ttk.Button(
            action_frame,
            text="取消下载",
            command=self.cancel_run,
            state="disabled",
        )
        self.cancel_button.pack(side="left", padx=(10, 0))

        ttk.Label(action_frame, textvariable=self.status_var).pack(side="right")

        progress_frame = ttk.LabelFrame(left_frame, text="执行进度", padding=12)
        progress_frame.grid(row=6, column=0, sticky="ew", pady=(12, 0))
        progress_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        ttk.Label(progress_frame, textvariable=self.progress_text_var).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )

        summary_frame = ttk.LabelFrame(left_frame, text="任务结果", padding=12)
        summary_frame.grid(row=7, column=0, sticky="ew", pady=(12, 0))
        summary_frame.columnconfigure(0, weight=1)

        ttk.Label(
            summary_frame,
            textvariable=self.summary_text_var,
            justify="left",
            wraplength=320,
        ).grid(row=0, column=0, sticky="w")

        self.open_template_button = ttk.Button(
            summary_frame,
            text="打开 Excel 模板",
            command=self.open_template_file,
            state="disabled",
        )
        self.open_template_button.grid(row=1, column=0, sticky="w", pady=(8, 0))

        self.open_photo_report_button = ttk.Button(
            summary_frame,
            text="打开结果清单",
            command=self.open_photo_report_file,
            state="disabled",
        )
        self.open_photo_report_button.grid(row=2, column=0, sticky="w", pady=(8, 0))

        self.photo_browser_frame = ttk.LabelFrame(self.photo_right_frame, text="Bucket 文件夹浏览", padding=12)
        self.photo_browser_frame.grid(row=0, column=0, sticky="nsew")
        self.photo_browser_frame.columnconfigure(0, weight=1)
        self.photo_browser_frame.rowconfigure(2, weight=1)

        browser_toolbar = ttk.Frame(self.photo_browser_frame)
        browser_toolbar.grid(row=0, column=0, sticky="ew")

        ttk.Button(
            browser_toolbar,
            text="加载当前层级",
            command=self.load_bucket_folders,
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            browser_toolbar,
            text="刷新已选文件夹数量",
            command=self.refresh_selected_folder_count,
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Button(
            browser_toolbar,
            text="上一级",
            command=self.go_to_parent_prefix,
        ).grid(row=0, column=2, sticky="w", padx=(8, 0))

        search_toolbar = ttk.Frame(self.photo_browser_frame)
        search_toolbar.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        search_toolbar.columnconfigure(0, weight=1)

        search_entry = self.create_text_entry(search_toolbar, textvariable=self.search_keyword_var)
        search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ttk.Button(
            search_toolbar,
            text="搜索当前层级文件夹",
            command=self.search_bucket_files,
        ).grid(row=0, column=1, sticky="e")

        tree_frame = ttk.Frame(self.photo_browser_frame)
        tree_frame.grid(row=2, column=0, sticky="nsew", pady=(10, 10))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.folder_tree = ttk.Treeview(
            tree_frame,
            columns=("meta",),
            show="tree headings",
            selectmode="browse",
            height=10,
        )
        self.folder_tree.heading("#0", text="名称")
        self.folder_tree.heading("meta", text="信息")
        self.folder_tree.column("#0", width=280, anchor="w")
        self.folder_tree.column("meta", width=140, anchor="center")
        self.folder_tree.grid(row=0, column=0, sticky="nsew")
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.folder_tree.bind("<Double-1>", self.on_tree_double_click)

        tree_scrollbar = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=self.folder_tree.yview,
        )
        tree_scrollbar.grid(row=0, column=1, sticky="ns")
        self.folder_tree.configure(yscrollcommand=tree_scrollbar.set)

        ttk.Label(self.photo_browser_frame, textvariable=self.folder_status_var).grid(
            row=3, column=0, sticky="w"
        )
        ttk.Label(self.photo_browser_frame, textvariable=self.selected_folder_info_var).grid(
            row=4, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Label(self.photo_browser_frame, textvariable=self.search_status_var).grid(
            row=5, column=0, sticky="w", pady=(10, 0)
        )

        certificate_frame = ttk.Frame(notebook, padding=14)
        certificate_frame.columnconfigure(0, weight=1)
        certificate_frame.rowconfigure(0, weight=1)
        notebook.add(certificate_frame, text="证件资料筛选")

        cert_paned = ttk.Panedwindow(certificate_frame, orient="horizontal")
        cert_paned.grid(row=0, column=0, sticky="nsew")
        self.certificate_paned = cert_paned

        cert_left_outer = ttk.Frame(cert_paned)
        cert_left_outer.columnconfigure(0, weight=1)
        cert_left_outer.rowconfigure(0, weight=1)

        cert_left_canvas = tk.Canvas(cert_left_outer, highlightthickness=0)
        cert_left_canvas.grid(row=0, column=0, sticky="nsew")
        cert_left_scrollbar = ttk.Scrollbar(
            cert_left_outer,
            orient="vertical",
            command=cert_left_canvas.yview,
        )
        cert_left_scrollbar.grid(row=0, column=1, sticky="ns")
        cert_left_canvas.configure(yscrollcommand=cert_left_scrollbar.set)

        cert_left_frame = ttk.Frame(cert_left_canvas)
        cert_left_frame.columnconfigure(0, weight=1)
        cert_left_window = cert_left_canvas.create_window((0, 0), window=cert_left_frame, anchor="nw")

        def sync_cert_left_scroll(_event=None) -> None:
            cert_left_canvas.configure(scrollregion=cert_left_canvas.bbox("all"))

        def sync_cert_left_width(event) -> None:
            cert_left_canvas.itemconfigure(cert_left_window, width=event.width)

        cert_left_frame.bind("<Configure>", sync_cert_left_scroll)
        cert_left_canvas.bind("<Configure>", sync_cert_left_width)
        self.bind_mousewheel_to_canvas(cert_left_canvas, cert_left_frame)

        cert_right_frame = ttk.Frame(cert_paned)
        cert_right_frame.columnconfigure(0, weight=1)
        cert_right_frame.rowconfigure(0, weight=1)
        cert_paned.add(cert_left_outer, weight=1)
        cert_paned.add(cert_right_frame, weight=1)

        cert_source_frame = ttk.LabelFrame(cert_left_frame, text="数据来源", padding=12)
        cert_source_frame.grid(row=0, column=0, sticky="ew")
        ttk.Radiobutton(
            cert_source_frame,
            text="从云存储下载后处理",
            variable=self.certificate_source_mode_var,
            value="oss",
            command=self.update_certificate_source_mode_ui,
        ).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(
            cert_source_frame,
            text="直接处理本地目录",
            variable=self.certificate_source_mode_var,
            value="local",
            command=self.update_certificate_source_mode_ui,
        ).pack(side="left")

        self.certificate_oss_frame = ttk.LabelFrame(cert_left_frame, text="证件资料 OSS 配置", padding=12)
        self.certificate_oss_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.certificate_oss_frame.columnconfigure(1, weight=1)
        self.add_cloud_type_row(self.certificate_oss_frame, 0)
        self.add_entry_row(self.certificate_oss_frame, 1, "Endpoint / Region", self.endpoint_var)
        self.add_entry_row(self.certificate_oss_frame, 2, "AccessKey ID", self.access_key_id_var)
        self.add_entry_row(self.certificate_oss_frame, 3, "AccessKey Secret", self.access_key_secret_var, show="*")
        self.add_certificate_bucket_picker(self.certificate_oss_frame, 4)
        self.add_certificate_prefix_picker(self.certificate_oss_frame, 5)

        cert_form = ttk.LabelFrame(cert_left_frame, text="筛选设置", padding=12)
        cert_form.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        cert_form.columnconfigure(1, weight=1)

        self.add_file_row(
            cert_form,
            row=0,
            label="人员模板",
            variable=self.certificate_template_var,
            command=self.choose_certificate_template,
            button_text="选择文件",
        )
        self.add_path_row(
            cert_form,
            row=1,
            label="本地证件资料目录",
            variable=self.certificate_source_dir_var,
        )
        self.add_path_row(
            cert_form,
            row=2,
            label="输出目录",
            variable=self.certificate_output_dir_var,
        )

        ttk.Label(cert_form, text="匹配列", width=16).grid(row=3, column=0, sticky="w", pady=4)
        match_row = ttk.Frame(cert_form)
        match_row.grid(row=3, column=1, columnspan=2, sticky="ew", pady=4)
        match_row.columnconfigure(0, weight=1)
        self.certificate_match_combo = ttk.Combobox(
            match_row,
            textvariable=self.certificate_match_column_var,
            values=self.certificate_headers,
            state="readonly",
        )
        self.certificate_match_combo.grid(row=0, column=0, sticky="ew")
        ttk.Button(match_row, text="加载模板列", command=self.load_certificate_headers).grid(
            row=0, column=1, padx=(8, 0)
        )

        rename_row = ttk.Frame(cert_form)
        rename_row.grid(row=4, column=1, columnspan=2, sticky="ew", pady=4)
        rename_row.columnconfigure(1, weight=1)
        self.add_tick_checkbutton(
            rename_row,
            text="导出后文件夹重命名",
            variable=self.certificate_rename_folder_var,
            command=self.update_certificate_mode_ui,
        ).grid(row=0, column=0, sticky="w")
        self.certificate_folder_name_combo = ttk.Combobox(
            rename_row,
            textvariable=self.certificate_folder_name_column_var,
            values=self.certificate_headers,
            state="readonly",
        )
        self.certificate_folder_name_combo.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(cert_form, text="筛选模式", width=16).grid(row=5, column=0, sticky="w", pady=4)
        mode_row = ttk.Frame(cert_form)
        mode_row.grid(row=5, column=1, columnspan=2, sticky="w", pady=4)
        ttk.Radiobutton(
            mode_row,
            text="复制整个人员文件夹",
            variable=self.certificate_mode_var,
            value="folder",
            command=self.update_certificate_mode_ui,
        ).pack(side="left")
        ttk.Radiobutton(
            mode_row,
            text="只复制关键词文件",
            variable=self.certificate_mode_var,
            value="keyword",
            command=self.update_certificate_mode_ui,
        ).pack(side="left", padx=(10, 0))

        ttk.Label(cert_form, text="文件关键词", width=16).grid(row=6, column=0, sticky="w", pady=4)
        self.certificate_keyword_entry = self.create_text_entry(
            cert_form,
            textvariable=self.certificate_keyword_var,
        )
        self.certificate_keyword_entry.grid(row=6, column=1, sticky="ew", pady=4)

        self.add_tick_checkbutton(
            cert_form,
            text="按分类一/分类二/分类三建立目录",
            variable=self.certificate_classify_var,
        ).grid(row=7, column=1, sticky="w", pady=(6, 0))

        self.add_tick_checkbutton(
            cert_form,
            text="仅预览，不实际执行",
            variable=self.certificate_dry_run_var,
        ).grid(row=8, column=1, sticky="w", pady=(6, 0))

        cert_action = ttk.Frame(cert_left_frame)
        cert_action.grid(row=3, column=0, sticky="ew", pady=(12, 0))

        self.certificate_download_button = ttk.Button(
            cert_action,
            text="下载证件资料",
            command=self.start_certificate_download_run,
        )
        self.certificate_download_button.pack(side="left")

        self.certificate_run_button = ttk.Button(
            cert_action,
            text="开始筛选",
            command=self.start_certificate_run,
        )
        self.certificate_run_button.pack(side="left", padx=(10, 0))

        self.certificate_cancel_button = ttk.Button(
            cert_action,
            text="取消任务",
            command=self.cancel_run,
            state="disabled",
        )
        self.certificate_cancel_button.pack(side="left", padx=(10, 0))

        ttk.Label(cert_action, textvariable=self.certificate_status_var).pack(side="right")

        cert_progress_frame = ttk.LabelFrame(cert_left_frame, text="筛选进度", padding=12)
        cert_progress_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        cert_progress_frame.columnconfigure(0, weight=1)

        self.certificate_progress_bar = ttk.Progressbar(
            cert_progress_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )
        self.certificate_progress_bar.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            cert_progress_frame,
            textvariable=self.certificate_progress_text_var,
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        cert_summary_frame = ttk.LabelFrame(cert_left_frame, text="筛选结果", padding=12)
        cert_summary_frame.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        cert_summary_frame.columnconfigure(0, weight=1)
        ttk.Label(
            cert_summary_frame,
            textvariable=self.certificate_summary_text_var,
            justify="left",
            wraplength=320,
        ).grid(row=0, column=0, sticky="w")

        self.open_certificate_report_button = ttk.Button(
            cert_summary_frame,
            text="打开结果清单",
            command=self.open_certificate_report_file,
            state="disabled",
        )
        self.open_certificate_report_button.grid(row=1, column=0, sticky="w", pady=(8, 0))

        self.certificate_browser_frame = ttk.LabelFrame(
            cert_right_frame,
            text="Bucket 文件夹浏览",
            padding=12,
        )
        self.certificate_browser_frame.grid(row=0, column=0, sticky="nsew")
        self.certificate_browser_frame.columnconfigure(0, weight=1)
        self.certificate_browser_frame.rowconfigure(2, weight=1)

        cert_browser_toolbar = ttk.Frame(self.certificate_browser_frame)
        cert_browser_toolbar.grid(row=0, column=0, sticky="ew")
        ttk.Button(
            cert_browser_toolbar,
            text="加载当前层级",
            command=self.load_certificate_folders,
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            cert_browser_toolbar,
            text="上一级",
            command=self.go_to_certificate_parent_prefix,
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))

        cert_search_toolbar = ttk.Frame(self.certificate_browser_frame)
        cert_search_toolbar.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        cert_search_toolbar.columnconfigure(0, weight=1)

        self.certificate_search_keyword_var = tk.StringVar()
        cert_search_entry = self.create_text_entry(
            cert_search_toolbar,
            textvariable=self.certificate_search_keyword_var,
        )
        cert_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ttk.Button(
            cert_search_toolbar,
            text="搜索当前层级文件夹",
            command=self.search_certificate_files,
        ).grid(row=0, column=1, sticky="e")

        cert_tree_frame = ttk.Frame(self.certificate_browser_frame)
        cert_tree_frame.grid(row=2, column=0, sticky="nsew", pady=(10, 10))
        cert_tree_frame.columnconfigure(0, weight=1)
        cert_tree_frame.rowconfigure(0, weight=1)

        self.certificate_folder_tree = ttk.Treeview(
            cert_tree_frame,
            columns=("meta",),
            show="tree headings",
            selectmode="browse",
            height=10,
        )
        self.certificate_folder_tree.heading("#0", text="名称")
        self.certificate_folder_tree.heading("meta", text="信息")
        self.certificate_folder_tree.column("#0", width=280, anchor="w")
        self.certificate_folder_tree.column("meta", width=140, anchor="center")
        self.certificate_folder_tree.grid(row=0, column=0, sticky="nsew")
        self.certificate_folder_tree.bind("<<TreeviewSelect>>", self.on_certificate_tree_select)
        self.certificate_folder_tree.bind("<Double-1>", self.on_certificate_tree_double_click)

        cert_tree_scrollbar = ttk.Scrollbar(
            cert_tree_frame,
            orient="vertical",
            command=self.certificate_folder_tree.yview,
        )
        cert_tree_scrollbar.grid(row=0, column=1, sticky="ns")
        self.certificate_folder_tree.configure(yscrollcommand=cert_tree_scrollbar.set)

        ttk.Label(self.certificate_browser_frame, textvariable=self.certificate_folder_status_var).grid(
            row=3, column=0, sticky="w"
        )
        ttk.Label(
            self.certificate_browser_frame,
            textvariable=self.certificate_selected_folder_info_var,
        ).grid(row=4, column=0, sticky="w", pady=(6, 0))

        ttk.Label(
            self.certificate_browser_frame,
            textvariable=self.certificate_search_status_var,
        ).grid(row=5, column=0, sticky="w", pady=(10, 0))

        word_frame = ttk.Frame(notebook, padding=14)
        word_frame.columnconfigure(0, weight=1)
        word_frame.rowconfigure(2, weight=1)
        notebook.add(word_frame, text="表样转换")

        word_form = ttk.LabelFrame(word_frame, text="模板转换", padding=12)
        word_form.grid(row=0, column=0, sticky="ew")
        word_form.columnconfigure(1, weight=1)

        self.add_file_row(
            word_form,
            row=0,
            label="表样文件",
            variable=self.word_source_var,
            command=self.choose_word_source,
            button_text="选择文件",
        )

        word_action = ttk.Frame(word_frame)
        word_action.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.word_net_button = ttk.Button(
            word_action,
            text="Net版导出",
            command=lambda: self.start_word_export("net"),
        )
        self.word_net_button.pack(side="left")
        self.word_java_button = ttk.Button(
            word_action,
            text="Java版导出",
            command=lambda: self.start_word_export("java"),
        )
        self.word_java_button.pack(side="left", padx=(10, 0))
        self.word_copy_button = ttk.Button(
            word_action,
            text="复制 HTML",
            command=self.copy_word_html,
            state="disabled",
        )
        self.word_copy_button.pack(side="left", padx=(10, 0))
        self.word_open_browser_button = ttk.Button(
            word_action,
            text="浏览器预览",
            command=self.open_word_preview_in_browser,
            state="disabled",
        )
        self.word_open_browser_button.pack(side="left", padx=(10, 0))
        ttk.Label(word_action, textvariable=self.word_status_var).pack(side="right")

        word_result_frame = ttk.LabelFrame(word_frame, text="导出结果", padding=12)
        word_result_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        word_result_frame.columnconfigure(0, weight=1)
        word_result_frame.rowconfigure(1, weight=1)
        ttk.Label(
            word_result_frame,
            textvariable=self.word_result_var,
            justify="left",
            wraplength=860,
        ).grid(row=0, column=0, sticky="w")

        word_view_notebook = ttk.Notebook(word_result_frame)
        word_view_notebook.grid(row=1, column=0, sticky="nsew", pady=(12, 0))

        word_code_frame = ttk.Frame(word_view_notebook, padding=8)
        word_code_frame.columnconfigure(0, weight=1)
        word_code_frame.rowconfigure(0, weight=1)
        word_view_notebook.add(word_code_frame, text="代码")

        self.word_code_text = tk.Text(
            word_code_frame,
            wrap="none",
            font=("Menlo", 11),
            relief="solid",
            height=18,
        )
        self.word_code_text.grid(row=0, column=0, sticky="nsew")
        self.word_code_text.configure(state="disabled")

        word_code_scroll_y = ttk.Scrollbar(
            word_code_frame,
            orient="vertical",
            command=self.word_code_text.yview,
        )
        word_code_scroll_y.grid(row=0, column=1, sticky="ns")
        self.word_code_text.configure(yscrollcommand=word_code_scroll_y.set)

        word_preview_frame = ttk.Frame(word_view_notebook, padding=8)
        word_preview_frame.columnconfigure(0, weight=1)
        word_preview_frame.rowconfigure(1, weight=1)
        word_view_notebook.add(word_preview_frame, text="预览")

        ttk.Label(
            word_preview_frame,
            textvariable=self.word_preview_status_var,
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.word_preview_container = ttk.Frame(word_preview_frame)
        self.word_preview_container.grid(row=1, column=0, sticky="nsew")
        self.word_preview_container.columnconfigure(0, weight=1)
        self.word_preview_container.rowconfigure(0, weight=1)

        match_frame = ttk.Frame(notebook, padding=14)
        match_frame.columnconfigure(0, weight=1)
        match_frame.rowconfigure(3, weight=1)
        notebook.add(match_frame, text="数据匹配")

        match_form = ttk.LabelFrame(match_frame, text="表格匹配补列", padding=12)
        match_form.grid(row=0, column=0, sticky="ew")
        match_form.columnconfigure(1, weight=1)
        self.add_file_row(
            match_form,
            row=0,
            label="目标表",
            variable=self.match_target_var,
            command=self.choose_match_target,
            button_text="选择文件",
        )
        self.add_file_row(
            match_form,
            row=1,
            label="来源表",
            variable=self.match_source_var,
            command=self.choose_match_source,
            button_text="选择文件",
        )
        self.add_entry_row(match_form, 2, "输出文件", self.match_output_var)
        ttk.Button(match_form, text="自动生成", command=self.fill_match_output_path).grid(
            row=2, column=2, padx=(8, 0), pady=4
        )
        ttk.Label(match_form, text="目标表匹配列", width=16).grid(row=3, column=0, sticky="w", pady=4)
        self.match_target_key_combo = ttk.Combobox(
            match_form,
            textvariable=self.match_target_key_var,
            values=self.match_target_headers,
            state="readonly",
        )
        self.match_target_key_combo.grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(match_form, text="来源表匹配列", width=16).grid(row=4, column=0, sticky="w", pady=4)
        self.match_source_key_combo = ttk.Combobox(
            match_form,
            textvariable=self.match_source_key_var,
            values=self.match_source_headers,
            state="readonly",
        )
        self.match_source_key_combo.grid(row=4, column=1, sticky="ew", pady=4)
        ttk.Button(match_form, text="加载表头", command=self.load_match_headers).grid(
            row=4, column=2, padx=(8, 0), pady=4
        )

        match_lists = ttk.Frame(match_frame)
        match_lists.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        match_lists.columnconfigure(0, weight=1)
        match_lists.columnconfigure(1, weight=1)

        extra_frame = ttk.LabelFrame(match_lists, text="附加匹配列映射", padding=12)
        extra_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        extra_frame.columnconfigure(0, weight=1)
        ttk.Label(extra_frame, text="目标表列").grid(row=0, column=0, sticky="w")
        ttk.Label(extra_frame, text="来源表列").grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.match_extra_target_combo = ttk.Combobox(
            extra_frame,
            textvariable=self.match_extra_target_var,
            values=self.match_target_headers,
            state="readonly",
        )
        self.match_extra_target_combo.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.match_extra_source_combo = ttk.Combobox(
            extra_frame,
            textvariable=self.match_extra_source_var,
            values=self.match_source_headers,
            state="readonly",
        )
        self.match_extra_source_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))
        ttk.Button(extra_frame, text="添加映射", command=self.add_extra_match_mapping).grid(
            row=1, column=2, padx=(8, 0), pady=(4, 0)
        )
        self.match_extra_tree = ttk.Treeview(
            extra_frame,
            columns=("target", "source"),
            show="headings",
            height=5,
        )
        self.match_extra_tree.heading("target", text="目标表列")
        self.match_extra_tree.heading("source", text="来源表列")
        self.match_extra_tree.column("target", width=160, anchor="w")
        self.match_extra_tree.column("source", width=160, anchor="w")
        self.match_extra_tree.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        extra_frame.rowconfigure(2, weight=1)
        ttk.Button(extra_frame, text="删除所选", command=self.remove_extra_match_mapping).grid(
            row=2, column=2, padx=(8, 0), sticky="n"
        )

        transfer_frame = ttk.LabelFrame(match_lists, text="补充列映射", padding=12)
        transfer_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        transfer_frame.columnconfigure(0, weight=1)
        ttk.Label(transfer_frame, text="结果列名").grid(row=0, column=0, sticky="w")
        ttk.Label(transfer_frame, text="来源表列").grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.match_transfer_target_entry = self.create_text_entry(
            transfer_frame,
            textvariable=self.match_transfer_target_var,
        )
        self.match_transfer_target_entry.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.match_transfer_source_combo = ttk.Combobox(
            transfer_frame,
            textvariable=self.match_transfer_source_var,
            values=self.match_source_headers,
            state="readonly",
        )
        self.match_transfer_source_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))
        ttk.Button(transfer_frame, text="添加补充列", command=self.add_transfer_mapping).grid(
            row=1, column=2, padx=(8, 0), pady=(4, 0)
        )
        self.match_transfer_tree = ttk.Treeview(
            transfer_frame,
            columns=("target", "source"),
            show="headings",
            height=5,
        )
        self.match_transfer_tree.heading("target", text="结果列名")
        self.match_transfer_tree.heading("source", text="来源表列")
        self.match_transfer_tree.column("target", width=160, anchor="w")
        self.match_transfer_tree.column("source", width=160, anchor="w")
        self.match_transfer_tree.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        transfer_frame.rowconfigure(2, weight=1)
        ttk.Button(transfer_frame, text="删除所选", command=self.remove_transfer_mapping).grid(
            row=2, column=2, padx=(8, 0), sticky="n"
        )

        match_action = ttk.Frame(match_frame)
        match_action.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self.match_run_button = ttk.Button(match_action, text="开始匹配", command=self.start_match_run)
        self.match_run_button.pack(side="left")
        self.match_open_button = ttk.Button(
            match_action,
            text="打开结果文件",
            command=self.open_match_result_file,
            state="disabled",
        )
        self.match_open_button.pack(side="left", padx=(10, 0))
        ttk.Label(match_action, textvariable=self.match_status_var).pack(side="right")

        match_result_frame = ttk.LabelFrame(match_frame, text="匹配结果", padding=12)
        match_result_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        match_result_frame.columnconfigure(0, weight=1)
        match_result_frame.rowconfigure(0, weight=1)
        self.match_result_text = tk.Text(
            match_result_frame,
            wrap="word",
            height=8,
            relief="solid",
            bd=1,
            bg="#FCFCFC",
            fg="#222222",
            insertbackground="#1677FF",
            padx=10,
            pady=10,
        )
        self.match_result_text.grid(row=0, column=0, sticky="nsew")
        self.match_result_text.configure(state="disabled")
        match_result_scroll = ttk.Scrollbar(
            match_result_frame,
            orient="vertical",
            command=self.match_result_text.yview,
        )
        match_result_scroll.grid(row=0, column=1, sticky="ns")
        self.match_result_text.configure(yscrollcommand=match_result_scroll.set)

        exam_frame = ttk.Frame(notebook, padding=14)
        exam_frame.columnconfigure(0, weight=1)
        exam_frame.rowconfigure(3, weight=1)
        notebook.add(exam_frame, text="考场编排")

        exam_files_frame = ttk.LabelFrame(exam_frame, text="输入文件", padding=12)
        exam_files_frame.grid(row=0, column=0, sticky="ew")
        exam_files_frame.columnconfigure(1, weight=1)
        self.add_file_row(
            exam_files_frame,
            row=0,
            label="考生明细表",
            variable=self.exam_candidate_var,
            command=self.choose_exam_candidate_file,
            button_text="选择文件",
        )
        self.add_file_row(
            exam_files_frame,
            row=1,
            label="岗位归组表",
            variable=self.exam_group_var,
            command=self.choose_exam_group_file,
            button_text="选择文件",
        )
        self.add_file_row(
            exam_files_frame,
            row=2,
            label="编排片段表",
            variable=self.exam_plan_var,
            command=self.choose_exam_plan_file,
            button_text="选择文件",
        )
        self.add_entry_row(exam_files_frame, 3, "输出文件", self.exam_output_var)
        ttk.Button(exam_files_frame, text="自动生成", command=self.fill_exam_output_path).grid(
            row=3, column=2, padx=(8, 0), pady=4
        )

        exam_rule_frame = ttk.LabelFrame(exam_frame, text="考号规则", padding=12)
        exam_rule_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        exam_rule_frame.columnconfigure(0, weight=1)
        exam_digits_row = ttk.Frame(exam_rule_frame)
        exam_digits_row.grid(row=0, column=0, sticky="ew")
        for idx, (label, variable) in enumerate(
            [
                ("考点位数", self.exam_point_digits_var),
                ("考场位数", self.exam_room_digits_var),
                ("座号位数", self.exam_seat_digits_var),
                ("流水号位数", self.exam_serial_digits_var),
            ]
        ):
            ttk.Label(exam_digits_row, text=label).grid(row=0, column=idx * 2, sticky="w")
            entry = self.create_text_entry(exam_digits_row, variable)
            entry.configure(width=6)
            entry.grid(row=0, column=idx * 2 + 1, sticky="w", padx=(6, 12))

        exam_sort_row = ttk.Frame(exam_rule_frame)
        exam_sort_row.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(exam_sort_row, text="组内顺序").pack(side="left")
        ttk.Radiobutton(
            exam_sort_row,
            text="按原顺序",
            variable=self.exam_sort_mode_var,
            value="original",
        ).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(
            exam_sort_row,
            text="随机打乱",
            variable=self.exam_sort_mode_var,
            value="random",
        ).pack(side="left", padx=(10, 0))

        exam_rule_editor = ttk.Frame(exam_rule_frame)
        exam_rule_editor.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(exam_rule_editor, text="项目").grid(row=0, column=0, sticky="w")
        self.exam_rule_type_combo = ttk.Combobox(
            exam_rule_editor,
            textvariable=self.exam_rule_type_var,
            values=self.exam_rule_base_types,
            state="readonly",
        )
        self.exam_rule_type_combo.grid(row=0, column=1, sticky="w", padx=(6, 12))
        ttk.Label(exam_rule_editor, text="自定义内容").grid(row=0, column=2, sticky="w")
        self.exam_rule_custom_entry = self.create_text_entry(exam_rule_editor, self.exam_rule_custom_var)
        self.exam_rule_custom_entry.grid(row=0, column=3, sticky="ew", padx=(6, 12))
        exam_rule_editor.columnconfigure(3, weight=1)
        ttk.Button(exam_rule_editor, text="添加规则", command=self.add_exam_rule_item).grid(
            row=0, column=4, sticky="e"
        )

        self.exam_rule_tree = ttk.Treeview(
            exam_rule_frame,
            columns=("index", "type", "custom"),
            show="headings",
            height=5,
        )
        self.exam_rule_tree.heading("index", text="顺序")
        self.exam_rule_tree.heading("type", text="项目")
        self.exam_rule_tree.heading("custom", text="自定义内容")
        self.exam_rule_tree.column("index", width=60, anchor="center")
        self.exam_rule_tree.column("type", width=120, anchor="w")
        self.exam_rule_tree.column("custom", width=220, anchor="w")
        self.exam_rule_tree.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        exam_rule_frame.rowconfigure(3, weight=1)
        exam_rule_actions = ttk.Frame(exam_rule_frame)
        exam_rule_actions.grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Button(exam_rule_actions, text="上移", command=self.move_exam_rule_up).pack(side="left")
        ttk.Button(exam_rule_actions, text="下移", command=self.move_exam_rule_down).pack(side="left", padx=(8, 0))
        ttk.Button(exam_rule_actions, text="删除所选", command=self.remove_exam_rule_item).pack(side="left", padx=(8, 0))

        exam_action = ttk.Frame(exam_frame)
        exam_action.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self.exam_template_button = ttk.Button(
            exam_action,
            text="导出标准模板",
            command=self.export_exam_templates_from_ui,
        )
        self.exam_template_button.pack(side="left")
        self.exam_run_button = ttk.Button(exam_action, text="开始编排", command=self.start_exam_arrange_run)
        self.exam_run_button.pack(side="left", padx=(10, 0))
        self.exam_open_button = ttk.Button(
            exam_action,
            text="打开结果文件",
            command=self.open_exam_result_file,
            state="disabled",
        )
        self.exam_open_button.pack(side="left", padx=(10, 0))
        ttk.Label(exam_action, textvariable=self.exam_status_var).pack(side="right")

        exam_result_frame = ttk.LabelFrame(exam_frame, text="编排结果", padding=12)
        exam_result_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        exam_result_frame.columnconfigure(0, weight=1)
        exam_result_frame.rowconfigure(0, weight=1)
        self.exam_result_text = tk.Text(
            exam_result_frame,
            wrap="word",
            height=8,
            relief="solid",
            bd=1,
            bg="#FCFCFC",
            fg="#222222",
            insertbackground="#1677FF",
            padx=10,
            pady=10,
        )
        self.exam_result_text.grid(row=0, column=0, sticky="nsew")
        self.exam_result_text.configure(state="disabled")
        exam_result_scroll = ttk.Scrollbar(exam_result_frame, orient="vertical", command=self.exam_result_text.yview)
        exam_result_scroll.grid(row=0, column=1, sticky="ns")
        self.exam_result_text.configure(yscrollcommand=exam_result_scroll.set)

        pack_frame = ttk.Frame(notebook, padding=14)
        pack_frame.columnconfigure(0, weight=1)
        pack_frame.rowconfigure(2, weight=1)
        notebook.add(pack_frame, text="结果打包")

        pack_form = ttk.LabelFrame(pack_frame, text="压缩加密", padding=12)
        pack_form.grid(row=0, column=0, sticky="ew")
        pack_form.columnconfigure(1, weight=1)

        self.add_path_row(
            pack_form,
            row=0,
            label="待打包文件夹",
            variable=self.pack_source_dir_var,
        )
        self.add_path_row(
            pack_form,
            row=1,
            label="输出目录",
            variable=self.pack_output_dir_var,
        )
        self.add_tick_checkbutton(
            pack_form,
            text="手动设置密码",
            variable=self.pack_use_custom_password_var,
            command=self.update_pack_password_mode_ui,
        ).grid(row=2, column=1, sticky="w", pady=(4, 0))
        self.pack_password_entry = self.add_entry_row(
            pack_form,
            3,
            "打包密码",
            self.pack_password_var,
        )

        pack_action = ttk.Frame(pack_frame)
        pack_action.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.pack_run_button = ttk.Button(
            pack_action,
            text="一键打包并加密",
            command=self.start_pack_run,
        )
        self.pack_run_button.pack(side="left")
        self.pack_copy_password_button = ttk.Button(
            pack_action,
            text="复制密码",
            command=self.copy_pack_password,
            state="disabled",
        )
        self.pack_copy_password_button.pack(side="left", padx=(10, 0))
        self.pack_open_button = ttk.Button(
            pack_action,
            text="打开压缩包",
            command=self.open_pack_file,
            state="disabled",
        )
        self.pack_open_button.pack(side="left", padx=(10, 0))
        ttk.Label(pack_action, textvariable=self.pack_status_var).pack(side="right")

        pack_result_frame = ttk.LabelFrame(pack_frame, text="打包结果", padding=12)
        pack_result_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        pack_result_frame.columnconfigure(0, weight=1)
        ttk.Label(
            pack_result_frame,
            textvariable=self.pack_result_var,
            justify="left",
            wraplength=860,
        ).grid(row=0, column=0, sticky="w")

        pack_query_frame = ttk.LabelFrame(pack_frame, text="密码查询", padding=12)
        pack_query_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        pack_query_frame.columnconfigure(0, weight=1)
        pack_query_frame.rowconfigure(1, weight=1)
        self.pack_query_entry = tk.Entry(
            pack_query_frame,
            textvariable=self.pack_query_var,
            font=("Helvetica", 12),
            relief="solid",
            bd=1,
            bg="#FCFCFC",
            fg="#303030",
            insertbackground="#2F6BFF",
            highlightthickness=1,
            highlightbackground="#C9D2E0",
            highlightcolor="#2F6BFF",
        )
        self.pack_query_entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(pack_query_frame, text="查询密码", command=self.run_pack_history_query).grid(
            row=0, column=1, padx=(10, 0)
        )
        self.pack_query_result_text = tk.Text(
            pack_query_frame,
            wrap="word",
            height=6,
            relief="solid",
            bd=1,
            bg="#FCFCFC",
            fg="#222222",
            insertbackground="#1677FF",
            padx=10,
            pady=10,
        )
        self.pack_query_result_text.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        self.pack_query_result_text.configure(state="disabled")
        pack_query_scroll = ttk.Scrollbar(
            pack_query_frame,
            orient="vertical",
            command=self.pack_query_result_text.yview,
        )
        pack_query_scroll.grid(row=1, column=2, sticky="ns", pady=(8, 0))
        self.pack_query_result_text.configure(yscrollcommand=pack_query_scroll.set)

        help_frame = ttk.Frame(notebook, padding=14)
        notebook.add(help_frame, text="使用说明")
        help_frame.columnconfigure(0, weight=1)
        help_frame.rowconfigure(0, weight=1)

        self.help_text = tk.Text(
            help_frame,
            wrap="word",
            font=("Helvetica", 13),
            relief="solid",
            padx=14,
            pady=14,
        )
        self.help_text.grid(row=0, column=0, sticky="nsew")

        help_scrollbar = ttk.Scrollbar(help_frame, orient="vertical", command=self.help_text.yview)
        help_scrollbar.grid(row=0, column=1, sticky="ns")
        self.help_text.configure(yscrollcommand=help_scrollbar.set)
        self.set_help_content()

        log_frame = ttk.Frame(notebook, padding=14)
        notebook.add(log_frame, text="运行日志")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)

        log_toolbar = ttk.Frame(log_frame)
        log_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(log_toolbar, text="清空日志", command=self.clear_log).pack(side="left")

        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            bg="#101826",
            fg="#E5EEF7",
            insertbackground="#E5EEF7",
            font=("Menlo", 12),
            relief="flat",
            padx=12,
            pady=12,
        )
        self.log_text.grid(row=1, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.write_log("桌面工具已启动。")
        self.write_log("建议先勾选“仅预览”，确认下载与分类路径后再正式执行。")
        self.write_log("可以在右侧浏览 bucket 文件夹，双击进入子目录。")
        self.write_log("也可以在右侧按文件夹名称筛选当前层级目录。")
        self.render_word_preview("")
        self.update_photo_source_mode_ui()
        self.update_certificate_mode_ui()
        self.update_certificate_source_mode_ui()
        self.update_pack_password_mode_ui()
        self.set_match_result_text(self.match_result_var.get())
        self.load_exam_group_headers()
        self.refresh_exam_rule_tree()
        self.set_exam_result_text(self.exam_result_var.get())
        self.set_pack_query_result_text("可按文件夹名、压缩包名或密码查询最近打包记录。")
        self.root.after(120, lambda: self.reset_split_default("photo"))

    def set_help_content(self) -> None:
        help_content = f"""报名系统工具箱 使用说明
版本：v{__version__}

一、照片下载与分类

适用场景：
- 批量下载考生照片
- 只下载模板中的人员照片
- 按模板中的分类信息整理照片

操作步骤：
1. 选择数据来源：
   - 从云存储下载后处理
   - 直接处理本地目录
2. 如果使用云存储，填写云类型、Endpoint / Region、AccessKey，并点击“加载 Bucket”。
3. 选择 bucket 后，点击右侧“加载当前层级”，找到照片所在目录。
4. 选择下载目录和分类目录。
5. 如需按名单下载，选择人员模板，点击“加载模板列”，选择匹配列，再勾选“只下载模板中的人员”。
6. 点击“下载并生成模板”。
7. 打开生成的 Excel 模板，填写“分类一 / 分类二 / 分类三 / 修改名称”。
8. 回到程序，点击“按模板分类”。

结果说明：
- 会生成“照片分类模板.xlsx”
- 会生成“照片分类结果清单.xlsx”
- 如果勾选“仅预览，不实际执行”，只显示计划，不真正下载和分类

二、证件资料筛选

适用场景：
- 按模板筛选部分人员的证件资料
- 按分类一、分类二、分类三整理资料
- 只导出某类证件，例如“学历证书”

操作步骤：
1. 选择数据来源：
   - 从云存储下载后处理
   - 直接处理本地目录
2. 选择人员模板，点击“加载模板列”，选择匹配列。
3. 如果需要，勾选“导出后文件夹重命名”，再选择“导出后文件夹名称列”。
4. 如果使用云存储，填写云类型、Endpoint / Region、AccessKey，并点击“加载 Bucket”。
5. 选择 bucket 后，点击右侧“加载当前层级”，找到证件资料目录。
6. 选择证件资料目录。
7. 如果需要先下载资料，可以点击“下载证件资料”。
8. 选择输出目录。
9. 选择筛选模式：
   - 复制整个人员文件夹
   - 只复制关键词文件
10. 如果是关键词模式，填写关键词，例如“学历证书”。
11. 根据需要勾选“按分类一 / 分类二 / 分类三建立目录”或“仅预览，不实际执行”。
12. 点击“开始筛选”。

结果说明：
- 会生成“证件资料筛选结果清单.xlsx”
- 如果模板中有分类信息，可以按分类目录导出
- 如果勾选“导出后文件夹重命名”，输出目录中的人员文件夹名称会按你选择的列来命名

三、表样转换

适用场景：
- 把 Word / Excel 表样转换成 HTML 代码
- 生成 Net 版或 Java 版占位符模板

操作步骤：
1. 选择表样文件（.doc / .docx / .xlsx）。
2. 点击“Net版导出”或“Java版导出”。
3. 在“代码”页查看并复制 HTML。
4. 如需查看页面效果，可以使用“浏览器预览”。

结果说明：
- Net 版占位符示例：{{[#考生表视图.姓名#]}}
- Java 版占位符示例：${{考生.姓名}}
- Excel 默认读取第一个 sheet
- Windows 下建议使用“浏览器预览”

四、数据匹配

适用场景：
- 两个 Excel 表之间按考号、身份证号、姓名等字段补列
- 代替手工写 VLOOKUP、XLOOKUP 或临时写 SQL
- 姓名可能重名时，再叠加单位、岗位等附加列提高匹配准确性

操作步骤：
1. 选择目标表和来源表（支持 .xlsx / .xls）。
2. 点击“加载表头”。
3. 分别选择目标表匹配列和来源表匹配列。
4. 如需提高准确性，在“附加匹配列映射”中设置“目标表列 -> 来源表列”，例如“单位 -> 报考单位”。
5. 在“补充列映射”中设置结果列名和来源表列，例如“身份证号 <- 身份证号”。
6. 确认输出文件路径后，点击“开始匹配”。

结果说明：
- 会生成新的 Excel 结果文件，不改原始目标表
- 结果文件会新增补充列
- 会额外生成一个“匹配结果清单”sheet，方便核对未匹配和重复键
- 如果第一行只是“附件1”这类说明，程序会尽量自动跳过并识别真正表头

五、考场编排

适用场景：
- 根据标准模板给考生批量补充考点、考场、座号、考号
- 考号规则不固定时，在程序里配置拼接顺序
- 混编考场时，通过编排片段表控制同一考场的不同座位区间
- 岗位归组表后续新增字段时，也可以直接作为考号拼接片段使用

操作步骤：
1. 可以先点击“导出标准模板”，生成三张标准表后再补充数据。
2. 准备三张标准表：
   - 考生明细表：姓名、身份证号、招聘单位、岗位名称
   - 岗位归组表：招聘单位、岗位名称、科目组、岗位编码、科目号
   - 编排片段表：科目组、考点、考场号、起始座号、结束座号、人数、起始流水号、结束流水号、备注
3. 在程序中选择这三张表。
4. 设置考点、考场、座号、流水号的位数。
5. 选择同组内顺序：
   - 按原顺序
   - 随机打乱
6. 在“考号规则”里按顺序添加规则，例如：
   - 自定义
   - 考点
   - 考场
   - 座号
7. 点击“开始编排”。

结果说明：
- 会生成新的 Excel 结果文件
- 在原始基础上新增：科目组、岗位编码、科目号、考点、考场、座号、考号、编排备注
- 如果岗位未找到归组，或科目组未找到编排片段，会在“编排备注”里标明原因
- 同一科目组内默认随机打乱后再分配，也可以手动切换为按原顺序

六、结果打包

适用场景：
- 把处理后的结果文件夹直接压缩成加密 zip
- 发送给客户或第三方时提高安全性

操作步骤：
1. 选择“待打包文件夹”。
2. 选择“输出目录”。
3. 如需客户指定密码，可勾选“手动设置密码”并输入密码。
4. 点击“一键打包并加密”。
5. 程序会自动使用文件夹名生成压缩包名。
6. 如果未手动输入密码，程序会自动生成密码，并在结果区显示。
7. 如忘记密码，可在下方“密码查询”中按文件夹名或压缩包名查询。

结果说明：
- 压缩包名称默认和文件夹名称一致
- 自动密码规则：当天日期 + 4位随机字符，例如 260313A7KQ
- 打包记录会保存到本地 JSON 文件，方便后期查询密码
- 可直接点击“复制密码”或“打开压缩包”

七、使用建议

- 第一次处理大批量文件时，建议先用少量数据测试。
- 使用模板前，建议先确认匹配列和分类列填写正确。
- 处理完成后，可以直接点击“打开结果清单”核对结果。
- 切换阿里云 / 腾讯云时，程序会分别记住两套配置。

六、运行日志

运行日志用于查看下载、筛选、分类和导出的详细过程。
如果需要回看处理步骤，可以打开“运行日志”页查看。
"""
        self.help_text.configure(state="normal")
        self.help_text.delete("1.0", tk.END)
        self.help_text.insert("1.0", help_content)
        self.help_text.configure(state="disabled")

    def on_tab_changed(self, _event=None) -> None:
        current_tab = self.notebook.select()
        tab_text = self.notebook.tab(current_tab, "text")
        if tab_text == "照片下载与分类":
            self.reset_split_default("photo")
        elif tab_text == "证件资料筛选":
            self.reset_split_default("certificate")

    def reset_split_default(self, pane_name: str) -> None:
        if not self.default_sash_pending.get(pane_name):
            return

        paned = self.photo_paned if pane_name == "photo" else self.certificate_paned
        try:
            width = paned.winfo_width()
            if width <= 100:
                self.root.after(120, lambda: self.reset_split_default(pane_name))
                return
            paned.sashpos(0, width // 2)
            self.default_sash_pending[pane_name] = False
        except tk.TclError:
            self.root.after(120, lambda: self.reset_split_default(pane_name))

    def add_entry_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        show: Optional[str] = None,
    ):
        ttk.Label(parent, text=label, width=16).grid(row=row, column=0, sticky="w", pady=4)
        entry = self.create_text_entry(parent, textvariable=variable, show=show or "")
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        return entry

    def add_tick_checkbutton(
        self,
        parent,
        text: str,
        variable: tk.BooleanVar,
        command=None,
        row: int = 0,
        column: int = 0,
        padx=(0, 18),
        pady=4,
        sticky: str = "w",
    ):
        background = "#f5f5f5"
        try:
            background = parent.cget("background")
        except tk.TclError:
            try:
                style = ttk.Style()
                background = style.lookup("TFrame", "background") or background
            except tk.TclError:
                pass
        widget = tk.Checkbutton(
            parent,
            text=text,
            variable=variable,
            onvalue=True,
            offvalue=False,
            command=command,
            anchor="w",
            highlightthickness=0,
            relief="flat",
            borderwidth=0,
            background=background,
            activebackground=background,
            selectcolor="#ffffff",
        )
        widget.grid(row=row, column=column, sticky=sticky, padx=padx, pady=pady)
        return widget

    def add_cloud_type_row(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="云类型", width=16).grid(row=row, column=0, sticky="w", pady=4)
        row_frame = ttk.Frame(parent)
        row_frame.grid(row=row, column=1, sticky="w", pady=4)
        ttk.Radiobutton(
            row_frame,
            text="阿里云 OSS",
            variable=self.cloud_type_var,
            value="aliyun",
            command=self.on_cloud_type_changed,
        ).pack(side="left")
        ttk.Radiobutton(
            row_frame,
            text="腾讯云 COS",
            variable=self.cloud_type_var,
            value="tencent",
            command=self.on_cloud_type_changed,
        ).pack(side="left", padx=(10, 0))

    def add_bucket_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Bucket", width=16).grid(row=row, column=0, sticky="w", pady=4)

        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)

        self.bucket_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.bucket_name_var,
            values=self.bucket_values,
            state="readonly",
        )
        self.bucket_combo.grid(row=0, column=0, sticky="ew")

        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        ttk.Button(
            action_frame,
            text="加载 Bucket",
            command=self.load_buckets,
        ).pack(side="left")
        ttk.Button(
            action_frame,
            text="保存配置",
            command=self.save_settings,
        ).pack(side="left", padx=(8, 0))

        ttk.Label(
            picker_frame,
            textvariable=self.bucket_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_prefix_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="下载文件夹", width=16).grid(row=row, column=0, sticky="w", pady=4)

        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)

        self.prefix_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.prefix_var,
            values=self.folder_values,
        )
        self.prefix_combo.grid(row=0, column=0, sticky="ew")

        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        ttk.Button(
            action_frame,
            text="加载文件夹",
            command=self.load_bucket_folders,
        ).pack(side="left")

        ttk.Button(
            action_frame,
            text="上一级",
            command=self.go_to_parent_prefix,
        ).pack(side="left", padx=(8, 0))

        ttk.Label(
            picker_frame,
            textvariable=self.folder_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_certificate_bucket_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Bucket", width=16).grid(row=row, column=0, sticky="w", pady=4)
        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)
        self.certificate_bucket_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.certificate_bucket_name_var,
            values=self.certificate_bucket_values,
            state="readonly",
        )
        self.certificate_bucket_combo.grid(row=0, column=0, sticky="ew")
        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(
            action_frame,
            text="加载 Bucket",
            command=self.load_certificate_buckets,
        ).pack(side="left")
        ttk.Button(
            action_frame,
            text="保存配置",
            command=self.save_settings,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(
            picker_frame,
            textvariable=self.certificate_bucket_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_certificate_prefix_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="证件资料前缀", width=16).grid(row=row, column=0, sticky="w", pady=4)
        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)
        self.certificate_prefix_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.certificate_prefix_var,
            values=self.certificate_folder_values,
        )
        self.certificate_prefix_combo.grid(row=0, column=0, sticky="ew")
        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(
            action_frame,
            text="加载文件夹",
            command=self.load_certificate_folders,
        ).pack(side="left")
        ttk.Button(
            action_frame,
            text="上一级",
            command=self.go_to_certificate_parent_prefix,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(
            picker_frame,
            textvariable=self.certificate_folder_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
    ) -> None:
        ttk.Label(parent, text=label, width=16).grid(row=row, column=0, sticky="w", pady=4)
        entry = self.create_text_entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(
            parent,
            text="选择",
            command=lambda: self.choose_directory(variable),
        ).grid(row=row, column=2, padx=(8, 0), pady=4)

    def add_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
        button_text: str = "选择",
    ) -> None:
        ttk.Label(parent, text=label, width=16).grid(row=row, column=0, sticky="w", pady=4)
        entry = self.create_text_entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(
            parent,
            text=button_text,
            command=command,
        ).grid(row=row, column=2, padx=(8, 0), pady=4)

    def choose_directory(self, variable: tk.StringVar) -> None:
        selected = filedialog.askdirectory(initialdir=variable.get() or str(Path.cwd()))
        if selected:
            variable.set(selected)

    def choose_certificate_template(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.certificate_template_var.get()).parent)
            if self.certificate_template_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if selected:
            self.certificate_template_var.set(selected)
            self.load_certificate_headers()

    def choose_photo_template(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.photo_template_var.get()).parent)
            if self.photo_template_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if selected:
            self.photo_template_var.set(selected)
            self.load_photo_headers()

    def choose_word_source(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.word_source_var.get()).parent)
            if self.word_source_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("表样文件", "*.docx *.doc *.xlsx"), ("所有文件", "*.*")],
        )
        if selected:
            self.word_source_var.set(selected)

    def choose_match_target(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.match_target_var.get()).parent)
            if self.match_target_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if selected:
            self.match_target_var.set(selected)
            self.fill_match_output_path()
            self.load_match_headers()

    def choose_match_source(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.match_source_var.get()).parent)
            if self.match_source_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if selected:
            self.match_source_var.set(selected)
            self.load_match_headers()

    def choose_exam_candidate_file(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.exam_candidate_var.get()).parent)
            if self.exam_candidate_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if selected:
            self.exam_candidate_var.set(selected)
            self.fill_exam_output_path()

    def choose_exam_group_file(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.exam_group_var.get()).parent)
            if self.exam_group_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if selected:
            self.exam_group_var.set(selected)
            self.load_exam_group_headers()

    def choose_exam_plan_file(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.exam_plan_var.get()).parent)
            if self.exam_plan_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if selected:
            self.exam_plan_var.set(selected)

    def export_exam_templates_from_ui(self) -> None:
        output_dir = filedialog.askdirectory(
            initialdir=str(Path(self.exam_candidate_var.get()).parent)
            if self.exam_candidate_var.get().strip()
            else str(Path.cwd()),
            title="选择模板导出目录",
        )
        if not output_dir:
            return
        try:
            summary = export_exam_templates(Path(output_dir))
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            self.exam_status_var.set("失败")
            self.exam_result_var.set(f"模板导出失败：\n{exc}")
            self.set_exam_result_text(self.exam_result_var.get())
            return

        self.last_exam_template_export = summary
        self.exam_candidate_var.set(str(summary.candidate_template_path))
        self.exam_group_var.set(str(summary.group_template_path))
        self.exam_plan_var.set(str(summary.plan_template_path))
        self.load_exam_group_headers()
        self.fill_exam_output_path()
        self.exam_status_var.set("模板已导出")
        self.exam_result_var.set(
            "\n".join(
                [
                    "标准模板已导出：",
                    f"目录：{summary.output_dir}",
                    f"考生明细模板：{summary.candidate_template_path.name}",
                    f"岗位归组模板：{summary.group_template_path.name}",
                    f"编排片段模板：{summary.plan_template_path.name}",
                    "",
                    "建议：先补充三张模板，再回来执行考场编排。",
                ]
            )
        )
        self.set_exam_result_text(self.exam_result_var.get())
        self.write_log(f"考场编排标准模板已导出：{summary.output_dir}")

    def load_exam_group_headers(self) -> None:
        self.exam_group_headers = []
        group_value = self.exam_group_var.get().strip()
        if not group_value:
            self.refresh_exam_rule_type_values()
            return
        group_path = Path(group_value)
        if not group_path.exists():
            self.refresh_exam_rule_type_values()
            return
        try:
            headers = list_exam_headers(group_path, ["招聘单位", "岗位名称", "科目组"])
        except Exception:
            self.refresh_exam_rule_type_values()
            return
        self.exam_group_headers = [
            header for header in headers if header not in {"招聘单位", "岗位名称", "科目组"}
        ]
        self.refresh_exam_rule_type_values()

    def refresh_exam_rule_type_values(self) -> None:
        values = self.exam_rule_base_types + self.exam_group_headers
        if hasattr(self, "exam_rule_type_combo") and self.exam_rule_type_combo is not None:
            self.exam_rule_type_combo.configure(values=values)
        current_value = self.exam_rule_type_var.get().strip()
        if current_value and current_value in values:
            return
        self.exam_rule_type_var.set(values[0] if values else "")

    def fill_exam_output_path(self) -> None:
        value = self.exam_candidate_var.get().strip()
        if not value:
            return
        candidate_path = Path(value)
        self.exam_output_var.set(str(candidate_path.with_name(f"{candidate_path.stem}_考场编排结果.xlsx")))

    def fill_match_output_path(self) -> None:
        target_value = self.match_target_var.get().strip()
        if not target_value:
            return
        target_path = Path(target_value)
        self.match_output_var.set(str(target_path.with_name(f"{target_path.stem}_数据匹配结果.xlsx")))

    def load_match_headers(self) -> None:
        target_value = self.match_target_var.get().strip()
        source_value = self.match_source_var.get().strip()
        self.match_target_headers = []
        self.match_source_headers = []
        self.match_extra_mappings = []
        self.match_transfer_mappings = []
        self.match_target_key_combo.configure(values=[])
        self.match_source_key_combo.configure(values=[])
        self.match_extra_target_combo.configure(values=[])
        self.match_extra_source_combo.configure(values=[])
        self.match_transfer_source_combo.configure(values=[])
        for item_id in self.match_extra_tree.get_children():
            self.match_extra_tree.delete(item_id)
        for item_id in self.match_transfer_tree.get_children():
            self.match_transfer_tree.delete(item_id)
        if not target_value or not source_value:
            return
        target_path = Path(target_value)
        source_path = Path(source_value)
        if not target_path.exists() or not source_path.exists():
            return
        try:
            self.match_target_headers = list_match_headers(target_path)
            self.match_source_headers = list_match_headers(source_path)
        except Exception as exc:
            messagebox.showerror("加载失败", str(exc))
            return
        self.match_target_key_combo.configure(values=self.match_target_headers)
        self.match_source_key_combo.configure(values=self.match_source_headers)
        self.match_extra_target_combo.configure(values=self.match_target_headers)
        self.match_extra_source_combo.configure(values=self.match_source_headers)
        self.match_transfer_source_combo.configure(values=self.match_source_headers)
        if not self.match_target_key_var.get().strip() and self.match_target_headers:
            self.match_target_key_var.set(self.match_target_headers[0])
        if not self.match_source_key_var.get().strip() and self.match_source_headers:
            self.match_source_key_var.set(self.match_source_headers[0])
        if self.match_source_headers and not self.match_transfer_source_var.get().strip():
            self.match_transfer_source_var.set(self.match_source_headers[0])
            self.match_transfer_target_var.set(self.match_source_headers[0])

    def add_extra_match_mapping(self) -> None:
        target_column = self.match_extra_target_var.get().strip()
        source_column = self.match_extra_source_var.get().strip()
        if not target_column or not source_column:
            messagebox.showerror("参数错误", "请选择目标表列和来源表列。")
            return
        mapping = ColumnMapping(target_column=target_column, source_column=source_column)
        if mapping in self.match_extra_mappings:
            return
        self.match_extra_mappings.append(mapping)
        self.match_extra_tree.insert("", tk.END, values=(target_column, source_column))

    def remove_extra_match_mapping(self) -> None:
        selected = self.match_extra_tree.selection()
        if not selected:
            return
        for item_id in selected:
            values = self.match_extra_tree.item(item_id, "values")
            self.match_extra_tree.delete(item_id)
            self.match_extra_mappings = [
                item
                for item in self.match_extra_mappings
                if (item.target_column, item.source_column) != tuple(values)
            ]

    def add_transfer_mapping(self) -> None:
        target_column = self.match_transfer_target_var.get().strip()
        source_column = self.match_transfer_source_var.get().strip()
        if not source_column:
            messagebox.showerror("参数错误", "请选择来源表补充列。")
            return
        if not target_column:
            target_column = source_column
            self.match_transfer_target_var.set(target_column)
        mapping = ColumnMapping(target_column=target_column, source_column=source_column)
        if mapping in self.match_transfer_mappings:
            return
        self.match_transfer_mappings.append(mapping)
        self.match_transfer_tree.insert("", tk.END, values=(target_column, source_column))

    def remove_transfer_mapping(self) -> None:
        selected = self.match_transfer_tree.selection()
        if not selected:
            return
        for item_id in selected:
            values = self.match_transfer_tree.item(item_id, "values")
            self.match_transfer_tree.delete(item_id)
            self.match_transfer_mappings = [
                item
                for item in self.match_transfer_mappings
                if (item.target_column, item.source_column) != tuple(values)
            ]

    def refresh_exam_rule_tree(self) -> None:
        for item_id in self.exam_rule_tree.get_children():
            self.exam_rule_tree.delete(item_id)
        for index, item in enumerate(self.exam_rule_items, start=1):
            self.exam_rule_tree.insert("", tk.END, values=(index, item.item_type, item.custom_text))

    def add_exam_rule_item(self) -> None:
        item_type = self.exam_rule_type_var.get().strip()
        custom_text = self.exam_rule_custom_var.get().strip()
        if not item_type:
            messagebox.showerror("参数错误", "请选择规则项目。")
            return
        if item_type == "自定义" and not custom_text:
            messagebox.showerror("参数错误", "自定义项目需要填写内容。")
            return
        self.exam_rule_items.append(ExamRuleItem(item_type=item_type, custom_text=custom_text))
        self.refresh_exam_rule_tree()
        if item_type == "自定义":
            self.exam_rule_custom_var.set("")

    def remove_exam_rule_item(self) -> None:
        selected = self.exam_rule_tree.selection()
        if not selected:
            return
        indexes = sorted(
            [int(self.exam_rule_tree.item(item_id, "values")[0]) - 1 for item_id in selected],
            reverse=True,
        )
        for index in indexes:
            if 0 <= index < len(self.exam_rule_items):
                self.exam_rule_items.pop(index)
        self.refresh_exam_rule_tree()

    def move_exam_rule_up(self) -> None:
        selected = self.exam_rule_tree.selection()
        if not selected:
            return
        index = int(self.exam_rule_tree.item(selected[0], "values")[0]) - 1
        if index <= 0:
            return
        self.exam_rule_items[index - 1], self.exam_rule_items[index] = (
            self.exam_rule_items[index],
            self.exam_rule_items[index - 1],
        )
        self.refresh_exam_rule_tree()
        self.exam_rule_tree.selection_set(self.exam_rule_tree.get_children()[index - 1])

    def move_exam_rule_down(self) -> None:
        selected = self.exam_rule_tree.selection()
        if not selected:
            return
        index = int(self.exam_rule_tree.item(selected[0], "values")[0]) - 1
        if index >= len(self.exam_rule_items) - 1:
            return
        self.exam_rule_items[index + 1], self.exam_rule_items[index] = (
            self.exam_rule_items[index],
            self.exam_rule_items[index + 1],
        )
        self.refresh_exam_rule_tree()
        self.exam_rule_tree.selection_set(self.exam_rule_tree.get_children()[index + 1])

    def create_text_entry(self, parent, textvariable: tk.StringVar, show: str = ""):
        entry_widget = tk.Entry(
            parent,
            textvariable=textvariable,
            show=show,
            relief="solid",
            bd=1,
            highlightthickness=1,
            bg="#FCFCFC",
            fg="#222222",
            disabledbackground="#F2F2F2",
            disabledforeground="#8A8A8A",
            readonlybackground="#F7F7F7",
            highlightbackground="#CFCFCF",
            highlightcolor="#4A90E2",
            insertbackground="#1677FF",
            insertwidth=2,
            insertborderwidth=0,
            cursor="xterm",
            takefocus=True,
            exportselection=False,
        )
        return entry_widget

    def bind_mousewheel_to_canvas(self, canvas: tk.Canvas, target) -> None:
        def on_mousewheel(event) -> None:
            if event.delta:
                delta = -1 if event.delta > 0 else 1
            elif getattr(event, "num", None) == 4:
                delta = -1
            elif getattr(event, "num", None) == 5:
                delta = 1
            else:
                delta = 0
            if delta != 0:
                canvas.yview_scroll(delta, "units")

        def bind_events(_event=None) -> None:
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)

        def unbind_events(_event=None) -> None:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        target.bind("<Enter>", bind_events)
        target.bind("<Leave>", unbind_events)

    def set_widget_state_recursive(self, widget, state: str) -> None:
        try:
            if isinstance(widget, ttk.Frame) or isinstance(widget, ttk.LabelFrame):
                pass
            elif isinstance(widget, ttk.Combobox):
                widget.configure(state="readonly" if state == "normal" and str(widget.cget("state")) == "readonly" else state)
            else:
                widget.configure(state=state)
        except tk.TclError:
            pass
        for child in widget.winfo_children():
            self.set_widget_state_recursive(child, state)

    def update_photo_source_mode_ui(self) -> None:
        is_oss = self.photo_source_mode_var.get() == "oss"
        self.skip_download_var.set(not is_oss)
        self.set_widget_state_recursive(self.photo_oss_frame, "normal" if is_oss else "disabled")
        self.set_widget_state_recursive(
            self.photo_browser_frame,
            "normal" if is_oss else "disabled",
        )
        if is_oss:
            self.run_button.configure(text="下载并生成模板")
            self.folder_status_var.set("未加载 bucket 文件夹")
            self.search_status_var.set("未搜索文件夹")
            self.selected_folder_info_var.set("当前未选择 bucket 文件夹")
        else:
            self.run_button.configure(text="生成本地模板")
            self.folder_status_var.set("本地模式不使用 bucket 浏览")
            self.search_status_var.set("本地模式不使用 bucket 搜索")
            self.selected_folder_info_var.set("请直接选择本地照片目录")

    def update_certificate_source_mode_ui(self) -> None:
        is_oss = self.certificate_source_mode_var.get() == "oss"
        self.set_widget_state_recursive(
            self.certificate_oss_frame,
            "normal" if is_oss else "disabled",
        )
        self.set_widget_state_recursive(
            self.certificate_browser_frame,
            "normal" if is_oss else "disabled",
        )
        if is_oss:
            self.certificate_download_button.configure(state="normal")
            self.certificate_folder_status_var.set("未加载 bucket 文件夹")
            self.certificate_search_status_var.set("未搜索文件夹")
            self.certificate_selected_folder_info_var.set("当前未选择 bucket 文件夹")
        else:
            self.certificate_download_button.configure(state="disabled")
            self.certificate_folder_status_var.set("本地模式不使用 bucket 浏览")
            self.certificate_search_status_var.set("本地模式不使用 bucket 搜索")
            self.certificate_selected_folder_info_var.set("请直接选择本地证件资料目录")

    def clear_log(self) -> None:
        self.log_text.delete("1.0", tk.END)

    def write_log(self, message: str) -> None:
        if int(self.log_text.index("end-1c").split(".")[0]) > 800:
            self.log_text.delete("1.0", "200.0")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def flush_logs(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            else:
                if isinstance(message, dict) and message.get("type") == "progress":
                    self.update_progress_ui(
                        stage=message.get("stage", ""),
                        current=message.get("current", 0),
                        total=message.get("total", 0),
                        current_file=message.get("current_file", ""),
                    )
                elif isinstance(message, dict) and message.get("type") == "certificate_summary":
                    self.update_certificate_summary_ui(message.get("summary"))
                elif message == "__TASK_DONE__":
                    self.run_button.configure(state="normal")
                    self.photo_classify_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.status_var.set("完成")
                    self.progress_text_var.set("任务完成")
                    if self.progress_bar is not None:
                        self.progress_bar["value"] = self.progress_bar["maximum"]
                elif message == "__CERTIFICATE_TASK_DONE__":
                    self.certificate_download_button.configure(
                        state="normal" if self.certificate_source_mode_var.get() == "oss" else "disabled"
                    )
                    self.certificate_run_button.configure(state="normal")
                    self.certificate_cancel_button.configure(state="disabled")
                    self.certificate_status_var.set("完成")
                    self.certificate_progress_text_var.set("筛选完成")
                    if self.certificate_progress_bar is not None:
                        self.certificate_progress_bar["value"] = self.certificate_progress_bar["maximum"]
                elif isinstance(message, dict) and message.get("type") == "summary":
                    self.update_summary_ui(message.get("summary"))
                elif message == "__TASK_CANCELLED__":
                    self.run_button.configure(state="normal")
                    self.photo_classify_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.status_var.set("已取消")
                    self.progress_text_var.set("任务已取消")
                elif message == "__CERTIFICATE_TASK_CANCELLED__":
                    self.certificate_download_button.configure(
                        state="normal" if self.certificate_source_mode_var.get() == "oss" else "disabled"
                    )
                    self.certificate_run_button.configure(state="normal")
                    self.certificate_cancel_button.configure(state="disabled")
                    self.certificate_status_var.set("已取消")
                    self.certificate_progress_text_var.set("筛选已取消")
                elif isinstance(message, str) and message.startswith("__TASK_FAILED__::"):
                    self.run_button.configure(state="normal")
                    self.photo_classify_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.status_var.set("失败")
                    self.progress_text_var.set("任务失败")
                    self.write_log(message.split("::", 1)[1])
                elif isinstance(message, str) and message.startswith("__CERTIFICATE_TASK_FAILED__::"):
                    self.certificate_download_button.configure(
                        state="normal" if self.certificate_source_mode_var.get() == "oss" else "disabled"
                    )
                    self.certificate_run_button.configure(state="normal")
                    self.certificate_cancel_button.configure(state="disabled")
                    self.certificate_status_var.set("失败")
                    self.certificate_progress_text_var.set("筛选失败")
                    self.write_log(message.split("::", 1)[1])
                elif isinstance(message, dict) and message.get("type") == "word_export":
                    self.update_word_export_ui(message.get("result"))
                elif isinstance(message, dict) and message.get("type") == "pack_summary":
                    self.update_pack_summary_ui(message.get("summary"))
                elif isinstance(message, dict) and message.get("type") == "match_summary":
                    self.update_match_summary_ui(message.get("summary"))
                elif isinstance(message, dict) and message.get("type") == "exam_summary":
                    self.update_exam_summary_ui(message.get("summary"))
                elif message == "__WORD_EXPORT_DONE__":
                    self.word_net_button.configure(state="normal")
                    self.word_java_button.configure(state="normal")
                    self.word_status_var.set("完成")
                elif message == "__MATCH_DONE__":
                    self.match_run_button.configure(state="normal")
                    self.match_status_var.set("完成")
                elif message == "__EXAM_DONE__":
                    self.exam_run_button.configure(state="normal")
                    self.exam_status_var.set("完成")
                elif message == "__PACK_DONE__":
                    self.pack_run_button.configure(state="normal")
                    self.pack_status_var.set("完成")
                elif isinstance(message, str) and message.startswith("__WORD_EXPORT_FAILED__::"):
                    self.word_net_button.configure(state="normal")
                    self.word_java_button.configure(state="normal")
                    self.word_status_var.set("失败")
                    error_text = message.split("::", 1)[1]
                    self.word_result_var.set(f"导出失败：\n{error_text}")
                    self.word_preview_status_var.set("预览不可用")
                    self.set_word_code("")
                    self.render_word_preview("")
                    self.word_copy_button.configure(state="disabled")
                    self.word_open_browser_button.configure(state="disabled")
                    self.write_log(error_text)
                elif isinstance(message, str) and message.startswith("__PACK_FAILED__::"):
                    self.pack_run_button.configure(state="normal")
                    self.pack_status_var.set("失败")
                    self.pack_result_var.set(f"打包失败：\n{message.split('::', 1)[1]}")
                    self.pack_copy_password_button.configure(state="disabled")
                    self.pack_open_button.configure(state="disabled")
                    self.write_log(message.split("::", 1)[1])
                elif isinstance(message, str) and message.startswith("__MATCH_FAILED__::"):
                    self.match_run_button.configure(state="normal")
                    self.match_status_var.set("失败")
                    self.match_result_var.set(f"匹配失败：\n{message.split('::', 1)[1]}")
                    self.match_open_button.configure(state="disabled")
                    self.write_log(message.split("::", 1)[1])
                elif isinstance(message, str) and message.startswith("__EXAM_FAILED__::"):
                    self.exam_run_button.configure(state="normal")
                    self.exam_status_var.set("失败")
                    self.exam_result_var.set(f"编排失败：\n{message.split('::', 1)[1]}")
                    self.set_exam_result_text(self.exam_result_var.get())
                    self.exam_open_button.configure(state="disabled")
                    self.write_log(message.split("::", 1)[1])
                else:
                    self.write_log(message)
        self.root.after(150, self.flush_logs)

    def make_logger(self):
        def logger(message: str) -> None:
            self.log_queue.put(message)

        return logger

    def make_progress_callback(self):
        def progress(stage: str, current: int, total: int, current_file: str) -> None:
            self.log_queue.put(
                {
                    "type": "progress",
                    "stage": stage,
                    "current": current,
                    "total": total,
                    "current_file": current_file,
                }
            )

        return progress

    def update_progress_ui(
        self,
        stage: str,
        current: int,
        total: int,
        current_file: str,
    ) -> None:
        if self.progress_bar is not None:
            self.progress_bar["maximum"] = max(total, 1)
            self.progress_bar["value"] = current

        if stage == "download":
            filename = Path(current_file).name if current_file else ""
            if total > 0:
                self.progress_text_var.set(f"下载进度：{current}/{total} {filename}".strip())
            else:
                self.progress_text_var.set("正在统计下载文件...")
        elif stage == "certificate":
            if self.certificate_progress_bar is not None:
                self.certificate_progress_bar["maximum"] = max(total, 1)
                self.certificate_progress_bar["value"] = current
            if total > 0:
                self.certificate_progress_text_var.set(
                    f"筛选进度：{current}/{total} {current_file}".strip()
                )
            else:
                self.certificate_progress_text_var.set("正在读取模板...")
        elif stage == "certificate_download":
            if self.certificate_progress_bar is not None:
                self.certificate_progress_bar["maximum"] = max(total, 1)
                self.certificate_progress_bar["value"] = current
            filename = Path(current_file).name if current_file else ""
            if total > 0:
                self.certificate_progress_text_var.set(
                    f"下载证件资料：{current}/{total} {filename}".strip()
                )
            else:
                self.certificate_progress_text_var.set("正在统计证件资料文件...")

    def cancel_run(self) -> None:
        if self.worker is None or not self.worker.is_alive():
            return
        self.cancel_event.set()
        self.status_var.set("取消中")
        self.progress_text_var.set("正在取消，请稍候...")
        self.certificate_status_var.set("取消中")
        self.certificate_progress_text_var.set("正在取消，请稍候...")
        self.write_log("用户请求取消当前任务。")

    def update_summary_ui(self, summary: Optional[WorkflowSummary]) -> None:
        self.last_summary = summary
        if summary is None:
            self.summary_text_var.set("任务结果会显示在这里")
            self.open_template_button.configure(state="disabled")
            self.open_photo_report_button.configure(state="disabled")
            return

        lines = []
        if summary.download_result is not None:
            lines.append(
                f"下载任务完成：共找到 {summary.download_result.total_found} 个文件，"
                f"新下载 {summary.download_result.downloaded_count} 个，"
                f"跳过已存在 {summary.download_result.skipped_existing_count} 个。"
            )
        if summary.template_file_count:
            lines.append(f"模板文件统计：当前目录共 {summary.template_file_count} 个文件。")
        if summary.template_created:
            lines.append("已生成 Excel 模板，可以直接打开填写。")
        elif summary.classified_count:
            lines.append(f"分类复制完成，共处理 {summary.classified_count} 个文件。")
            if summary.report_path is not None:
                lines.append(f"已导出结果清单：{summary.report_path.name}")
        elif summary.cancelled:
            lines.append("任务已取消。")
        elif summary.dry_run:
            lines.append("当前为预览模式，没有实际写入文件。")
        if not lines:
            lines.append("任务已完成。")

        self.summary_text_var.set("\n".join(lines))
        if summary.template_path.exists():
            self.open_template_button.configure(state="normal")
        else:
            self.open_template_button.configure(state="disabled")
        if summary.report_path is not None and summary.report_path.exists():
            self.open_photo_report_button.configure(state="normal")
        else:
            self.open_photo_report_button.configure(state="disabled")

    def update_certificate_summary_ui(
        self,
        summary: Optional[CertificateFilterSummary],
    ) -> None:
        self.last_certificate_summary = summary
        if summary is None:
            self.certificate_summary_text_var.set("证件资料筛选结果会显示在这里")
            self.open_certificate_report_button.configure(state="disabled")
            return

        if summary.total_rows == 0 and summary.download_result is not None:
            lines = [
                f"证件资料下载完成：共找到 {summary.download_result.total_found} 个文件，"
                f"新下载 {summary.download_result.downloaded_count} 个，"
                f"跳过已存在 {summary.download_result.skipped_existing_count} 个。",
                f"实际下载目录：{summary.source_dir}",
            ]
            if summary.cancelled:
                lines.append("任务已取消。")
            elif summary.dry_run:
                lines.append("当前为预览模式，没有实际下载文件。")
            else:
                lines.append("请到上面的“证件资料目录”查看下载结果。")
            self.certificate_summary_text_var.set("\n".join(lines))
            self.open_certificate_report_button.configure(state="disabled")
            return

        lines = [
            f"模板有效行数：{summary.total_rows}，匹配到 {summary.matched_people} 人。",
            f"缺失人员文件夹：{summary.missing_people} 人。",
            f"实际复制：{summary.copied_people} 人，{summary.copied_files} 个文件。",
        ]
        if summary.download_result is not None:
            lines.insert(
                0,
                f"证件资料下载完成：共找到 {summary.download_result.total_found} 个文件，"
                f"新下载 {summary.download_result.downloaded_count} 个，"
                f"跳过已存在 {summary.download_result.skipped_existing_count} 个。",
            )
        if summary.classify_output:
            lines.append("输出结构已按 分类一/分类二/分类三 建目录。")
        else:
            lines.append("输出结构未按分类字段建目录。")
        if summary.rename_folder and summary.folder_name_column:
            lines.append(f"导出后文件夹名称列：{summary.folder_name_column}")
        else:
            lines.append("导出后文件夹名称保持匹配列。")
        if summary.keyword:
            lines.append(f"筛选模式：只复制文件名包含“{summary.keyword}”的文件。")
        else:
            lines.append("筛选模式：复制整个人员文件夹。")
        if summary.cancelled:
            lines.append("任务已取消。")
        elif summary.dry_run:
            lines.append("当前为预览模式，没有实际复制文件。")
        elif summary.report_path is not None:
            lines.append(f"已导出结果清单：{summary.report_path.name}")
        self.certificate_summary_text_var.set("\n".join(lines))
        if summary.report_path is not None and summary.report_path.exists():
            self.open_certificate_report_button.configure(state="normal")
        else:
            self.open_certificate_report_button.configure(state="disabled")

    def update_word_export_ui(self, result: Optional[WordExportResult]) -> None:
        self.last_word_export = result
        if result is None:
            self.word_result_var.set("表样转换结果会显示在这里")
            self.word_preview_status_var.set("未生成预览")
            self.set_word_code("")
            self.render_word_preview("")
            self.word_copy_button.configure(state="disabled")
            self.word_open_browser_button.configure(state="disabled")
            return
        variant_label = "Net版" if result.variant == "net" else "Java版"
        self.word_result_var.set(
            f"{variant_label}导出完成：\n源文件：{result.source_path}\nHTML 代码已生成，可直接预览或复制。"
        )
        self.word_preview_status_var.set(f"{variant_label}预览")
        self.set_word_code(result.html_content)
        self.render_word_preview(result.preview_html)
        self.word_copy_button.configure(state="normal")
        self.word_open_browser_button.configure(state="normal")

    def update_pack_summary_ui(self, summary: Optional[PackSummary]) -> None:
        self.last_pack_summary = summary
        if summary is None:
            self.pack_result_var.set("结果打包信息会显示在这里")
            self.pack_copy_password_button.configure(state="disabled")
            self.pack_open_button.configure(state="disabled")
            return

        self.pack_result_var.set(
            "打包完成：\n"
            f"源目录：{summary.source_dir}\n"
            f"压缩包：{summary.output_path}\n"
            f"文件数：{summary.file_count}\n"
            f"时间：{summary.created_at}\n"
            f"密码：{summary.password}"
        )
        self.pack_copy_password_button.configure(state="normal")
        self.pack_open_button.configure(state="normal")

    def update_match_summary_ui(self, summary: Optional[DataMatchSummary]) -> None:
        self.last_match_summary = summary
        if summary is None:
            self.match_result_var.set("数据匹配结果会显示在这里")
            self.match_open_button.configure(state="disabled")
            self.set_match_result_text(self.match_result_var.get())
            return
        self.match_result_var.set(
            "匹配完成：\n"
            f"目标表：{summary.target_path}\n"
            f"来源表：{summary.source_path}\n"
            f"结果文件：{summary.output_path}\n"
            f"目标表表头：第 {summary.target_header_row} 行\n"
            f"来源表表头：第 {summary.source_header_row} 行\n"
            f"总行数：{summary.total_rows}\n"
            f"匹配成功：{summary.matched_rows}\n"
            f"未匹配：{summary.unmatched_rows}\n"
            f"来源重复键：{summary.duplicate_source_keys}\n"
            f"重复未写入：{summary.ambiguous_rows}"
        )
        self.set_match_result_text(self.match_result_var.get())
        self.match_open_button.configure(state="normal")

    def set_match_result_text(self, content: str) -> None:
        if self.match_result_text is None:
            return
        self.match_result_text.configure(state="normal")
        self.match_result_text.delete("1.0", tk.END)
        self.match_result_text.insert("1.0", content)
        self.match_result_text.configure(state="disabled")

    def set_pack_query_result_text(self, content: str) -> None:
        if self.pack_query_result_text is None:
            return
        self.pack_query_result_text.configure(state="normal")
        self.pack_query_result_text.delete("1.0", tk.END)
        self.pack_query_result_text.insert("1.0", content)
        self.pack_query_result_text.configure(state="disabled")

    def update_exam_summary_ui(self, summary: Optional[ExamArrangeSummary]) -> None:
        self.last_exam_summary = summary
        if summary is None:
            self.exam_result_var.set("考场编排结果会显示在这里")
            self.set_exam_result_text(self.exam_result_var.get())
            self.exam_open_button.configure(state="disabled")
            return
        self.exam_result_var.set(
            "编排完成：\n"
            f"考生表：{summary.candidate_path}\n"
            f"岗位归组表：{summary.group_path}\n"
            f"编排片段表：{summary.plan_path}\n"
            f"结果文件：{summary.output_path}\n"
            f"总人数：{summary.total_candidates}\n"
            f"成功编排：{summary.arranged_candidates}\n"
            f"未找到岗位归组：{summary.missing_groups}\n"
            f"未找到编排片段：{summary.missing_plan_groups}\n"
            f"重复岗位归组：{summary.duplicate_group_rows}\n"
            f"剩余空座：{summary.unused_plan_slots}"
        )
        self.set_exam_result_text(self.exam_result_var.get())
        self.exam_open_button.configure(state="normal")

    def set_exam_result_text(self, content: str) -> None:
        if self.exam_result_text is None:
            return
        self.exam_result_text.configure(state="normal")
        self.exam_result_text.delete("1.0", tk.END)
        self.exam_result_text.insert("1.0", content)
        self.exam_result_text.configure(state="disabled")

    def update_pack_password_mode_ui(self) -> None:
        if self.pack_use_custom_password_var.get():
            self.pack_password_entry.configure(state="normal")
        else:
            self.pack_password_var.set("")
            self.pack_password_entry.configure(state="disabled")

    def run_pack_history_query(self) -> None:
        keyword = self.pack_query_var.get().strip()
        records = query_pack_history(keyword)
        if not records:
            self.set_pack_query_result_text("未找到匹配的打包记录。")
            self.pack_status_var.set("未找到记录")
            return
        latest = records[0]
        self.set_pack_query_result_text(
            (
                f"最近匹配记录：\n"
                f"文件夹：{latest.get('source_name', '')}\n"
                f"压缩包：{latest.get('archive_name', '')}\n"
                f"时间：{latest.get('created_at', '')}\n"
                f"密码：{latest.get('password', '')}\n"
                f"路径：{latest.get('output_path', '')}"
            )
        )
        self.pack_status_var.set("已查询密码")

    def set_word_code(self, html_content: str) -> None:
        if self.word_code_text is None:
            return
        self.word_code_text.configure(state="normal")
        self.word_code_text.delete("1.0", tk.END)
        self.word_code_text.insert("1.0", html_content)
        self.word_code_text.configure(state="disabled")

    def render_word_preview(self, html_content: str) -> None:
        for child in self.word_preview_container.winfo_children():
            child.destroy()
        if not html_content:
            ttk.Label(
                self.word_preview_container,
                text="导出后会在这里显示 HTML 预览。",
            ).grid(row=0, column=0, sticky="nw")
            return
        if os.name == "nt":
            ttk.Label(
                self.word_preview_container,
                text="Windows 下已关闭程序内置预览。请点击上方“浏览器预览”查看 HTML 效果。",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, sticky="nw")
            self.word_preview_widget = None
            self.word_preview_status_var.set(self.word_preview_status_var.get() + "（Windows 建议使用浏览器预览）")
            return
        try:
            from tkinterweb import HtmlFrame

            preview = HtmlFrame(
                self.word_preview_container,
                messages_enabled=False,
            )
            preview.load_html(html_content)
            preview.grid(row=0, column=0, sticky="nsew")
            self.word_preview_widget = preview
            self.word_preview_status_var.set(self.word_preview_status_var.get() + "（表格增强预览）")
        except Exception:
            try:
                from tkhtmlview import HTMLScrolledText

                preview = HTMLScrolledText(
                    self.word_preview_container,
                    html=html_content,
                )
                preview.grid(row=0, column=0, sticky="nsew")
                self.word_preview_widget = preview
                self.word_preview_status_var.set(self.word_preview_status_var.get() + "（兼容预览）")
            except Exception:
                ttk.Label(
                    self.word_preview_container,
                    text="当前环境缺少 HTML 预览依赖，请先安装 requirements 后重启程序。",
                    wraplength=760,
                    justify="left",
                ).grid(row=0, column=0, sticky="nw")
                self.word_preview_widget = None

    def copy_word_html(self) -> None:
        if self.last_word_export is None:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_word_export.html_content)
        self.root.update()
        self.word_status_var.set("已复制 HTML")

    def open_word_preview_in_browser(self) -> None:
        if self.last_word_export is None:
            return
        temp_file = Path(tempfile.gettempdir()) / f"word_preview_{self.last_word_export.variant}.html"
        temp_file.write_text(self.last_word_export.html_content, encoding="utf-8")
        try:
            if os.name == "nt":
                os.startfile(str(temp_file))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(temp_file)])
            else:
                subprocess.Popen(["xdg-open", str(temp_file)])
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def copy_pack_password(self) -> None:
        if self.last_pack_summary is None:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_pack_summary.password)
        self.root.update()
        self.pack_status_var.set("已复制密码")

    def open_pack_file(self) -> None:
        if self.last_pack_summary is None:
            return
        self.open_local_file(self.last_pack_summary.output_path, "未找到压缩包")

    def open_match_result_file(self) -> None:
        if self.last_match_summary is None:
            return
        self.open_local_file(self.last_match_summary.output_path, "未找到匹配结果文件")

    def open_exam_result_file(self) -> None:
        if self.last_exam_summary is None:
            return
        self.open_local_file(self.last_exam_summary.output_path, "未找到编排结果文件")

    def open_template_file(self) -> None:
        if self.last_summary is None:
            return
        template_path = self.last_summary.template_path
        self.open_local_file(template_path, "未找到 Excel 模板")

    def open_photo_report_file(self) -> None:
        if self.last_summary is None or self.last_summary.report_path is None:
            return
        self.open_local_file(self.last_summary.report_path, "未找到结果清单")

    def open_certificate_report_file(self) -> None:
        if self.last_certificate_summary is None or self.last_certificate_summary.report_path is None:
            return
        self.open_local_file(self.last_certificate_summary.report_path, "未找到结果清单")

    def open_local_file(self, file_path: Path, not_found_title: str) -> None:
        if not file_path.exists():
            messagebox.showerror("文件不存在", f"{not_found_title}：{file_path}")
            return
        try:
            if os.name == "nt":
                os.startfile(str(file_path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(file_path)])
            else:
                subprocess.Popen(["xdg-open", str(file_path)])
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def build_options(self) -> RunOptions:
        is_oss = self.photo_source_mode_var.get() == "oss"
        return RunOptions(
            prefix=self.prefix_var.get().strip() if is_oss else "",
            download_dir=Path(self.download_dir_var.get().strip()),
            sorted_dir=Path(self.sorted_dir_var.get().strip()),
            skip_download=not is_oss,
            dry_run=self.dry_run_var.get(),
            flat=self.flat_var.get(),
            include_duplicates=self.include_duplicates_var.get(),
            move_sorted_files=self.move_sorted_files_var.get(),
            skip_existing=self.skip_existing_var.get(),
        )

    def build_config(self) -> OssConfig:
        return validate_oss_config(
            OssConfig(
                cloud_type=self.cloud_type_var.get().strip(),
                access_key_id=self.access_key_id_var.get().strip(),
                access_key_secret=self.access_key_secret_var.get().strip(),
                endpoint=self.endpoint_var.get().strip(),
                bucket_name=self.bucket_name_var.get().strip(),
            )
        )

    def build_certificate_config(self) -> OssConfig:
        return validate_oss_config(
            OssConfig(
                cloud_type=self.cloud_type_var.get().strip(),
                access_key_id=self.access_key_id_var.get().strip(),
                access_key_secret=self.access_key_secret_var.get().strip(),
                endpoint=self.endpoint_var.get().strip(),
                bucket_name=self.certificate_bucket_name_var.get().strip(),
            )
        )

    def build_credentials(self) -> tuple[str, str, str, str]:
        return validate_oss_credentials(
            self.cloud_type_var.get().strip(),
            self.access_key_id_var.get().strip(),
            self.access_key_secret_var.get().strip(),
            self.endpoint_var.get().strip(),
        )

    def validate_cloud_endpoint(self, cloud_type: str, endpoint: str) -> None:
        cleaned = endpoint.strip()
        if cloud_type == "aliyun":
            if "aliyuncs.com" not in cleaned:
                raise ValueError(
                    "阿里云 OSS 的 Endpoint 不正确，应类似 `https://oss-cn-hangzhou.aliyuncs.com`。"
                )
            return

        normalized = cleaned
        if normalized.startswith("http://"):
            normalized = normalized[len("http://") :]
        elif normalized.startswith("https://"):
            normalized = normalized[len("https://") :]
        normalized = normalized.split("/")[0]
        if "." in normalized:
            normalized = normalized.split(".", 1)[0]
        if not normalized.startswith("ap-"):
            raise ValueError(
                "腾讯云 COS 的 Endpoint / Region 不正确，应填写 Region（如 `ap-beijing`）或对应 COS Endpoint。"
            )

    def check_endpoint_reachable(self, cloud_type: str, endpoint: str) -> None:
        target = endpoint.strip()
        if cloud_type == "tencent" and not target.startswith(("http://", "https://")):
            target = f"https://cos.{target}.myqcloud.com"

        host = target
        if host.startswith("http://"):
            host = host[len("http://") :]
        elif host.startswith("https://"):
            host = host[len("https://") :]
        host = host.split("/", 1)[0]
        if not host:
            raise ValueError("Endpoint / Region 不能为空。")

        try:
            socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
        except socket.gaierror as exc:
            raise ValueError(f"Endpoint 无法解析，请检查地址是否正确：{host}") from exc

        try:
            with socket.create_connection((host, 443), timeout=2):
                return
        except OSError as exc:
            raise ValueError(f"无法连接到 Endpoint：{host}，请检查地址或网络。") from exc

    def format_cloud_error(self, error: str) -> str:
        lowered = error.lower()
        if "invalidaccesskeyid" in lowered:
            return "AccessKey ID 不存在或填写错误，请检查后重试。"
        if "secretid is forbidden" in lowered or "invalidsecretid" in lowered:
            return "AccessKey ID 不存在或填写错误，请检查后重试。"
        if "secretkey" in lowered and ("invalid" in lowered or "mismatch" in lowered):
            return "AccessKey Secret 不正确，请检查后重试。"
        if "signaturedoesnotmatch" in lowered or "signature not match" in lowered:
            return "AccessKey Secret 不正确，请检查后重试。"
        if "access denied" in lowered or "'status': 403" in lowered or "status code: 403" in lowered:
            return "当前账号没有访问权限，或密钥填写有误。"
        if "nosuchbucket" in lowered or "bucket does not exist" in lowered:
            return "Bucket 不存在，请检查 bucket 名称是否正确。"
        if "invalidbucketname" in lowered:
            return "Bucket 名称格式不正确。"
        if "connection" in lowered or "timeout" in lowered or "name or service not known" in lowered:
            return "无法连接到云存储服务，请检查 Endpoint / Region 和网络。"
        return error

    def set_folder_values(self, folders) -> None:
        values = [""] + list(folders)
        self.folder_values = values
        self.prefix_combo["values"] = values

    def set_bucket_values(self, buckets: List[str]) -> None:
        self.bucket_values = buckets
        self.bucket_combo["values"] = buckets
        current_bucket = self.bucket_name_var.get().strip()
        if current_bucket and current_bucket not in buckets:
            self.bucket_name_var.set("")

    def set_certificate_bucket_values(self, buckets: List[str]) -> None:
        self.certificate_bucket_values = buckets
        self.certificate_bucket_combo["values"] = buckets
        current_bucket = self.certificate_bucket_name_var.get().strip()
        if current_bucket and current_bucket not in buckets:
            self.certificate_bucket_name_var.set("")

    def sync_bucket_values(self, buckets: List[str]) -> None:
        self.set_bucket_values(buckets)
        self.set_certificate_bucket_values(buckets)

    def set_certificate_folder_values(self, folders: List[str]) -> None:
        values = [""] + list(folders)
        self.certificate_folder_values = values
        self.certificate_prefix_combo["values"] = values

    def default_cloud_profile(self, cloud_type: str) -> Dict[str, str]:
        # 两家云厂商的配置分开记忆，切换时不用来回重新输入。
        if cloud_type == "tencent":
            endpoint = "ap-beijing"
        else:
            endpoint = "https://oss-cn-hangzhou.aliyuncs.com"
        return {
            "access_key_id": "",
            "access_key_secret": "",
            "endpoint": endpoint,
            "bucket_name": "",
            "certificate_bucket_name": "",
            "prefix": "",
            "certificate_prefix": "",
        }

    def snapshot_current_cloud_profile(self) -> Dict[str, str]:
        return {
            "access_key_id": self.access_key_id_var.get().strip(),
            "access_key_secret": self.access_key_secret_var.get().strip(),
            "endpoint": self.endpoint_var.get().strip(),
            "bucket_name": self.bucket_name_var.get().strip(),
            "certificate_bucket_name": self.certificate_bucket_name_var.get().strip(),
            "prefix": self.prefix_var.get().strip(),
            "certificate_prefix": self.certificate_prefix_var.get().strip(),
        }

    def apply_cloud_profile(self, cloud_type: str) -> None:
        profile = self.cloud_profiles.get(cloud_type) or self.default_cloud_profile(cloud_type)
        self.access_key_id_var.set(profile.get("access_key_id", ""))
        self.access_key_secret_var.set(profile.get("access_key_secret", ""))
        self.endpoint_var.set(profile.get("endpoint", self.default_cloud_profile(cloud_type)["endpoint"]))
        self.bucket_name_var.set(profile.get("bucket_name", ""))
        self.certificate_bucket_name_var.set(profile.get("certificate_bucket_name", ""))
        self.prefix_var.set(profile.get("prefix", ""))
        self.certificate_prefix_var.set(profile.get("certificate_prefix", ""))

    def load_saved_settings(self) -> None:
        if not self.SETTINGS_FILE.exists():
            return
        try:
            settings = json.loads(self.SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return

        self.prefix_var.set(settings.get("prefix", ""))
        self.photo_template_var.set(settings.get("photo_template", ""))
        self.photo_match_column_var.set(settings.get("photo_match_column", ""))
        self.photo_filter_by_template_var.set(
            settings.get("photo_filter_by_template", self.photo_filter_by_template_var.get())
        )
        self.download_dir_var.set(settings.get("download_dir", self.download_dir_var.get()))
        self.sorted_dir_var.set(settings.get("sorted_dir", self.sorted_dir_var.get()))
        self.cloud_type_var.set(settings.get("cloud_type", self.cloud_type_var.get()))
        saved_profiles = settings.get("cloud_profiles")
        if isinstance(saved_profiles, dict):
            for cloud_type in ("aliyun", "tencent"):
                profile = saved_profiles.get(cloud_type)
                if isinstance(profile, dict):
                    merged = self.default_cloud_profile(cloud_type)
                    merged.update({key: str(value) for key, value in profile.items() if value is not None})
                    self.cloud_profiles[cloud_type] = merged
        else:
            legacy_cloud_type = self.cloud_type_var.get().strip() or "aliyun"
            legacy_profile = self.default_cloud_profile(legacy_cloud_type)
            legacy_profile.update(
                {
                    "access_key_id": settings.get("access_key_id", ""),
                    "access_key_secret": settings.get("access_key_secret", ""),
                    "endpoint": settings.get("endpoint", legacy_profile["endpoint"]),
                    "bucket_name": settings.get("bucket_name", ""),
                    "certificate_bucket_name": settings.get("certificate_bucket_name", ""),
                    "prefix": settings.get("prefix", ""),
                    "certificate_prefix": settings.get("certificate_prefix", ""),
                }
            )
            self.cloud_profiles[legacy_cloud_type] = legacy_profile
        self.apply_cloud_profile(self.cloud_type_var.get().strip() or "aliyun")
        self.photo_source_mode_var.set(settings.get("photo_source_mode", self.photo_source_mode_var.get()))
        self.skip_download_var.set(settings.get("skip_download", self.skip_download_var.get()))
        self.flat_var.set(settings.get("flat", self.flat_var.get()))
        self.dry_run_var.set(settings.get("dry_run", self.dry_run_var.get()))
        self.include_duplicates_var.set(
            settings.get("include_duplicates", self.include_duplicates_var.get())
        )
        self.move_sorted_files_var.set(
            settings.get("move_sorted_files", self.move_sorted_files_var.get())
        )
        self.skip_existing_var.set(settings.get("skip_existing", self.skip_existing_var.get()))
        self.certificate_template_var.set(settings.get("certificate_template", ""))
        self.certificate_source_dir_var.set(
            settings.get("certificate_source_dir", self.certificate_source_dir_var.get())
        )
        self.certificate_output_dir_var.set(
            settings.get("certificate_output_dir", self.certificate_output_dir_var.get())
        )
        self.certificate_match_column_var.set(settings.get("certificate_match_column", ""))
        self.certificate_rename_folder_var.set(
            settings.get("certificate_rename_folder", self.certificate_rename_folder_var.get())
        )
        self.certificate_folder_name_column_var.set(
            settings.get("certificate_folder_name_column", "")
        )
        self.certificate_keyword_var.set(
            settings.get("certificate_keyword", self.certificate_keyword_var.get())
        )
        self.certificate_classify_var.set(
            settings.get("certificate_classify", self.certificate_classify_var.get())
        )
        self.certificate_dry_run_var.set(
            settings.get("certificate_dry_run", self.certificate_dry_run_var.get())
        )
        self.certificate_mode_var.set(
            settings.get("certificate_mode", self.certificate_mode_var.get())
        )
        self.certificate_source_mode_var.set(
            settings.get("certificate_source_mode", self.certificate_source_mode_var.get())
        )
        self.certificate_bucket_name_var.set(settings.get("certificate_bucket_name", ""))
        self.certificate_prefix_var.set(settings.get("certificate_prefix", ""))
        self.word_source_var.set(settings.get("word_source", ""))
        self.pack_source_dir_var.set(settings.get("pack_source_dir", ""))
        self.pack_output_dir_var.set(settings.get("pack_output_dir", self.pack_output_dir_var.get()))
        self.pack_use_custom_password_var.set(
            settings.get("pack_use_custom_password", self.pack_use_custom_password_var.get())
        )
        self.pack_query_var.set(settings.get("pack_query", ""))
        self.match_target_var.set(settings.get("match_target", ""))
        self.match_source_var.set(settings.get("match_source", ""))
        self.match_target_key_var.set(settings.get("match_target_key", ""))
        self.match_source_key_var.set(settings.get("match_source_key", ""))
        self.match_output_var.set(settings.get("match_output", ""))
        self.exam_candidate_var.set(settings.get("exam_candidate", ""))
        self.exam_group_var.set(settings.get("exam_group", ""))
        self.exam_plan_var.set(settings.get("exam_plan", ""))
        self.exam_output_var.set(settings.get("exam_output", ""))
        self.exam_point_digits_var.set(settings.get("exam_point_digits", self.exam_point_digits_var.get()))
        self.exam_room_digits_var.set(settings.get("exam_room_digits", self.exam_room_digits_var.get()))
        self.exam_seat_digits_var.set(settings.get("exam_seat_digits", self.exam_seat_digits_var.get()))
        self.exam_serial_digits_var.set(settings.get("exam_serial_digits", self.exam_serial_digits_var.get()))
        self.exam_sort_mode_var.set(settings.get("exam_sort_mode", self.exam_sort_mode_var.get()))
        self.exam_rule_items = [
            ExamRuleItem(
                item_type=str(item.get("item_type", "")),
                custom_text=str(item.get("custom_text", "")),
            )
            for item in settings.get("exam_rule_items", [])
            if isinstance(item, dict) and str(item.get("item_type", "")).strip()
        ]
        self.load_exam_group_headers()

    def save_settings(self) -> None:
        settings = {
            "prefix": self.prefix_var.get().strip(),
            "photo_template": self.photo_template_var.get().strip(),
            "photo_match_column": self.photo_match_column_var.get().strip(),
            "photo_filter_by_template": self.photo_filter_by_template_var.get(),
            "download_dir": self.download_dir_var.get().strip(),
            "sorted_dir": self.sorted_dir_var.get().strip(),
            "cloud_type": self.cloud_type_var.get().strip(),
            "cloud_profiles": {
                **self.cloud_profiles,
                self.cloud_type_var.get().strip() or "aliyun": self.snapshot_current_cloud_profile(),
            },
            "photo_source_mode": self.photo_source_mode_var.get().strip(),
            "skip_download": self.skip_download_var.get(),
            "flat": self.flat_var.get(),
            "dry_run": self.dry_run_var.get(),
            "include_duplicates": self.include_duplicates_var.get(),
            "move_sorted_files": self.move_sorted_files_var.get(),
            "skip_existing": self.skip_existing_var.get(),
            "certificate_template": self.certificate_template_var.get().strip(),
            "certificate_source_dir": self.certificate_source_dir_var.get().strip(),
            "certificate_output_dir": self.certificate_output_dir_var.get().strip(),
            "certificate_match_column": self.certificate_match_column_var.get().strip(),
            "certificate_rename_folder": self.certificate_rename_folder_var.get(),
            "certificate_folder_name_column": self.certificate_folder_name_column_var.get().strip(),
            "certificate_keyword": self.certificate_keyword_var.get().strip(),
            "certificate_classify": self.certificate_classify_var.get(),
            "certificate_dry_run": self.certificate_dry_run_var.get(),
            "certificate_mode": self.certificate_mode_var.get().strip(),
            "certificate_source_mode": self.certificate_source_mode_var.get().strip(),
            "certificate_bucket_name": self.certificate_bucket_name_var.get().strip(),
            "certificate_prefix": self.certificate_prefix_var.get().strip(),
            "word_source": self.word_source_var.get().strip(),
            "pack_source_dir": self.pack_source_dir_var.get().strip(),
            "pack_output_dir": self.pack_output_dir_var.get().strip(),
            "pack_use_custom_password": self.pack_use_custom_password_var.get(),
            "pack_query": self.pack_query_var.get().strip(),
            "match_target": self.match_target_var.get().strip(),
            "match_source": self.match_source_var.get().strip(),
            "match_target_key": self.match_target_key_var.get().strip(),
            "match_source_key": self.match_source_key_var.get().strip(),
            "match_output": self.match_output_var.get().strip(),
            "exam_candidate": self.exam_candidate_var.get().strip(),
            "exam_group": self.exam_group_var.get().strip(),
            "exam_plan": self.exam_plan_var.get().strip(),
            "exam_output": self.exam_output_var.get().strip(),
            "exam_point_digits": self.exam_point_digits_var.get().strip(),
            "exam_room_digits": self.exam_room_digits_var.get().strip(),
            "exam_seat_digits": self.exam_seat_digits_var.get().strip(),
            "exam_serial_digits": self.exam_serial_digits_var.get().strip(),
            "exam_sort_mode": self.exam_sort_mode_var.get().strip(),
            "exam_rule_items": [
                {"item_type": item.item_type, "custom_text": item.custom_text}
                for item in self.exam_rule_items
            ],
        }
        self.SETTINGS_FILE.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.write_log(f"配置已保存到 {self.SETTINGS_FILE}")

    def update_certificate_mode_ui(self) -> None:
        if self.certificate_mode_var.get() == "keyword":
            self.certificate_keyword_entry.configure(state="normal")
        else:
            self.certificate_keyword_entry.configure(state="disabled")
        if self.certificate_rename_folder_var.get():
            self.certificate_folder_name_combo.configure(state="readonly")
        else:
            self.certificate_folder_name_combo.configure(state="disabled")

    def on_cloud_type_changed(self) -> None:
        previous_type = "tencent" if self.cloud_type_var.get() == "aliyun" else "aliyun"
        self.cloud_profiles[previous_type] = self.snapshot_current_cloud_profile()
        self.apply_cloud_profile(self.cloud_type_var.get().strip())

        # 切换云类型后，旧 bucket/前缀结果不能继续沿用，避免串用阿里云和腾讯云的数据。
        self.sync_bucket_values([])
        self.set_folder_values([])
        self.set_certificate_folder_values([])
        self.bucket_status_var.set("切换云类型后，请重新加载 bucket 列表")
        self.certificate_bucket_status_var.set("切换云类型后，请重新加载 bucket 列表")
        self.folder_status_var.set("未加载 bucket 文件夹")
        self.certificate_folder_status_var.set("未加载 bucket 文件夹")

    def set_certificate_headers(self, headers: List[str]) -> None:
        self.certificate_headers = headers
        self.certificate_match_combo["values"] = headers
        self.certificate_folder_name_combo["values"] = headers
        current_value = self.certificate_match_column_var.get().strip()
        if current_value and current_value not in headers:
            self.certificate_match_column_var.set("")
        if not self.certificate_match_column_var.get().strip() and headers:
            self.certificate_match_column_var.set(headers[0])
        current_folder_name_value = self.certificate_folder_name_column_var.get().strip()
        if current_folder_name_value and current_folder_name_value not in headers:
            self.certificate_folder_name_column_var.set("")

    def set_photo_headers(self, headers: List[str]) -> None:
        self.photo_headers = headers
        self.photo_match_combo["values"] = headers
        current_value = self.photo_match_column_var.get().strip()
        if current_value and current_value not in headers:
            self.photo_match_column_var.set("")
        if not self.photo_match_column_var.get().strip() and headers:
            self.photo_match_column_var.set(headers[0])

    def load_photo_headers(self) -> None:
        template_path = self.photo_template_var.get().strip()
        if not template_path:
            messagebox.showinfo("缺少模板", "请先选择人员模板文件。")
            return
        try:
            headers = list_template_headers(Path(template_path))
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        self.set_photo_headers(headers)
        self.write_log(f"已读取照片模板列：{', '.join(headers) if headers else '无'}")

    def load_certificate_headers(self) -> None:
        template_path = self.certificate_template_var.get().strip()
        if not template_path:
            messagebox.showinfo("缺少模板", "请先选择人员模板文件。")
            return
        try:
            headers = list_template_headers(Path(template_path))
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        self.set_certificate_headers(headers)
        self.write_log(f"已读取模板列：{', '.join(headers) if headers else '无'}")

    def load_buckets(self) -> None:
        try:
            cloud_type, access_key_id, access_key_secret, endpoint = self.build_credentials()
            self.validate_cloud_endpoint(cloud_type, endpoint)
            self.check_endpoint_reachable(cloud_type, endpoint)
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.bucket_status_var.set("正在加载 bucket 列表...")
        self.write_log("开始加载 bucket 列表。")
        self.bucket_load_token += 1
        current_token = self.bucket_load_token
        self.root.after(8000, lambda: self.handle_bucket_load_timeout(current_token))

        def worker() -> None:
            try:
                buckets = list_buckets(access_key_id, access_key_secret, endpoint, cloud_type=cloud_type)
            except Exception as exc:
                self.root.after(
                    0,
                    lambda: self.finish_bucket_load(
                        error=f"{type(exc).__name__}: {exc}",
                        token=current_token,
                    ),
                )
                return
            self.root.after(0, lambda: self.finish_bucket_load(buckets=buckets, token=current_token))

        threading.Thread(target=worker, daemon=True).start()

    def load_certificate_buckets(self) -> None:
        try:
            cloud_type, access_key_id, access_key_secret, endpoint = self.build_credentials()
            self.validate_cloud_endpoint(cloud_type, endpoint)
            self.check_endpoint_reachable(cloud_type, endpoint)
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.certificate_bucket_status_var.set("正在加载 bucket 列表...")
        self.write_log("开始加载证件资料 bucket 列表。")
        self.certificate_bucket_load_token += 1
        current_token = self.certificate_bucket_load_token
        self.root.after(8000, lambda: self.handle_certificate_bucket_load_timeout(current_token))

        def worker() -> None:
            try:
                buckets = list_buckets(access_key_id, access_key_secret, endpoint, cloud_type=cloud_type)
            except Exception as exc:
                self.root.after(
                    0,
                    lambda: self.finish_certificate_bucket_load(
                        error=f"{type(exc).__name__}: {exc}",
                        token=current_token,
                    ),
                )
                return
            self.root.after(
                0,
                lambda: self.finish_certificate_bucket_load(buckets=buckets, token=current_token),
            )

        threading.Thread(target=worker, daemon=True).start()

    def finish_bucket_load(
        self,
        buckets: Optional[List[str]] = None,
        error: Optional[str] = None,
        token: Optional[int] = None,
    ) -> None:
        if token is not None and token != self.bucket_load_token:
            return
        if error is not None:
            self.bucket_status_var.set("加载失败")
            self.write_log(f"加载 bucket 失败：{error}")
            messagebox.showerror("加载 bucket 失败", self.format_cloud_error(error))
            return

        buckets = buckets or []
        self.sync_bucket_values(buckets)
        if buckets:
            if not self.bucket_name_var.get().strip():
                self.bucket_name_var.set(buckets[0])
            if not self.certificate_bucket_name_var.get().strip():
                self.certificate_bucket_name_var.set(buckets[0])
            self.bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
            self.certificate_bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
            self.write_log(f"找到 {len(buckets)} 个 bucket。")
            self.prefix_var.set("")
            self.folder_status_var.set("请选择 bucket 后点击“加载当前层级”")
        else:
            self.bucket_status_var.set("未找到可用 bucket")
            self.certificate_bucket_status_var.set("未找到可用 bucket")
            self.write_log("当前凭证下未找到可用 bucket。")

    def finish_certificate_bucket_load(
        self,
        buckets: Optional[List[str]] = None,
        error: Optional[str] = None,
        token: Optional[int] = None,
    ) -> None:
        if token is not None and token != self.certificate_bucket_load_token:
            return
        if error is not None:
            self.certificate_bucket_status_var.set("加载失败")
            self.write_log(f"加载证件资料 bucket 失败：{error}")
            messagebox.showerror("加载 bucket 失败", self.format_cloud_error(error))
            return

        buckets = buckets or []
        self.sync_bucket_values(buckets)
        if buckets:
            if not self.certificate_bucket_name_var.get().strip():
                self.certificate_bucket_name_var.set(buckets[0])
            if not self.bucket_name_var.get().strip():
                self.bucket_name_var.set(buckets[0])
            self.certificate_bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
            self.bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
            self.write_log(f"证件资料找到 {len(buckets)} 个 bucket。")
            self.certificate_prefix_var.set("")
            self.certificate_folder_status_var.set("请选择 bucket 后点击“加载当前层级”")
        else:
            self.certificate_bucket_status_var.set("未找到可用 bucket")
            self.bucket_status_var.set("未找到可用 bucket")
            self.write_log("当前凭证下未找到可用证件资料 bucket。")

    def handle_bucket_load_timeout(self, token: int) -> None:
        if token != self.bucket_load_token:
            return
        if self.bucket_status_var.get() != "正在加载 bucket 列表...":
            return
        self.bucket_status_var.set("加载失败")
        timeout_message = "加载 bucket 超时，请检查 Endpoint、密钥或网络后重试。"
        self.write_log(timeout_message)
        messagebox.showerror("加载 bucket 失败", timeout_message)
        self.bucket_load_token += 1

    def handle_certificate_bucket_load_timeout(self, token: int) -> None:
        if token != self.certificate_bucket_load_token:
            return
        if self.certificate_bucket_status_var.get() != "正在加载 bucket 列表...":
            return
        self.certificate_bucket_status_var.set("加载失败")
        timeout_message = "加载 bucket 超时，请检查 Endpoint、密钥或网络后重试。"
        self.write_log(timeout_message)
        messagebox.showerror("加载 bucket 失败", timeout_message)
        self.certificate_bucket_load_token += 1

    def go_to_parent_prefix(self) -> None:
        current = self.prefix_var.get().strip().strip("/")
        if not current:
            self.prefix_var.set("")
            return
        parts = current.split("/")
        parent = "/".join(parts[:-1])
        self.prefix_var.set(parent + "/" if parent else "")
        self.load_bucket_folders()

    def go_to_certificate_parent_prefix(self) -> None:
        current = self.certificate_prefix_var.get().strip().strip("/")
        if not current:
            self.certificate_prefix_var.set("")
            return
        parts = current.split("/")
        parent = "/".join(parts[:-1])
        self.certificate_prefix_var.set(parent + "/" if parent else "")
        self.load_certificate_folders()

    def load_bucket_folders(self) -> None:
        if self.photo_source_mode_var.get() != "oss":
            return
        try:
            config = self.build_config()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        current_prefix = self.prefix_var.get().strip()
        self.folder_status_var.set("正在加载 bucket 文件夹...")
        self.write_log(f"加载 bucket 文件夹：{current_prefix or '/'}")

        def worker() -> None:
            try:
                entries = list_browser_entries(config, current_prefix)
            except Exception as exc:
                self.root.after(
                    0,
                    lambda: self.finish_folder_load(error=f"{type(exc).__name__}: {exc}"),
                )
                return
            self.root.after(0, lambda: self.finish_folder_load(entries=entries))

        threading.Thread(target=worker, daemon=True).start()

    def load_certificate_folders(self) -> None:
        try:
            config = self.build_certificate_config()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        current_prefix = self.certificate_prefix_var.get().strip()
        self.certificate_folder_status_var.set("正在加载 bucket 文件夹...")
        self.write_log(f"加载证件资料 bucket 文件夹：{current_prefix or '/'}")

        def worker() -> None:
            try:
                entries = list_browser_entries(config, current_prefix)
            except Exception as exc:
                self.root.after(
                    0,
                    lambda: self.finish_certificate_folder_load(
                        error=f"{type(exc).__name__}: {exc}"
                    ),
                )
                return
            self.root.after(
                0,
                lambda: self.finish_certificate_folder_load(entries=entries),
            )

        threading.Thread(target=worker, daemon=True).start()

    def finish_folder_load(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        if error is not None:
            self.folder_status_var.set("加载失败")
            self.write_log(f"加载文件夹失败：{error}")
            messagebox.showerror("加载失败", self.format_cloud_error(error))
            return

        entries = entries or []
        self.current_folder_entries = entries
        folders = [entry.key for entry in entries if entry.entry_type == "folder"]
        self.set_folder_values(folders)
        self.render_folder_tree(entries)
        folder_count = len([entry for entry in entries if entry.entry_type == "folder"])
        file_count = len([entry for entry in entries if entry.entry_type == "file"])
        if entries:
            self.folder_status_var.set(
                f"已加载 {folder_count} 个文件夹，{file_count} 个文件"
            )
            self.write_log(
                f"当前层级找到 {folder_count} 个文件夹，{file_count} 个文件。"
            )
        else:
            self.folder_status_var.set("当前层级没有子文件夹和文件")
            self.write_log("当前层级没有可显示的文件夹或文件。")

    def finish_certificate_folder_load(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        if error is not None:
            self.certificate_folder_status_var.set("加载失败")
            self.write_log(f"加载证件资料文件夹失败：{error}")
            messagebox.showerror("加载失败", self.format_cloud_error(error))
            return

        entries = entries or []
        self.current_certificate_folder_entries = entries
        folders = [entry.key for entry in entries if entry.entry_type == "folder"]
        self.set_certificate_folder_values(folders)
        self.render_certificate_folder_tree(entries)
        folder_count = len([entry for entry in entries if entry.entry_type == "folder"])
        file_count = len([entry for entry in entries if entry.entry_type == "file"])
        if entries:
            self.certificate_folder_status_var.set(
                f"已加载 {folder_count} 个文件夹，{file_count} 个文件"
            )
            self.write_log(
                f"证件资料当前层级找到 {folder_count} 个文件夹，{file_count} 个文件。"
            )
        else:
            self.certificate_folder_status_var.set("当前层级没有子文件夹和文件")
            self.write_log("证件资料当前层级没有可显示的文件夹或文件。")

    def render_folder_tree(self, entries: List[BrowserEntry]) -> None:
        if self.folder_tree is None:
            return
        for item in self.folder_tree.get_children():
            self.folder_tree.delete(item)
        self.folder_nodes = {}
        for entry in entries:
            meta = "文件夹" if entry.entry_type == "folder" else "文件"
            node_id = self.folder_tree.insert(
                "",
                "end",
                text=entry.display_name,
                values=(meta,),
            )
            self.folder_nodes[node_id] = entry

    def render_certificate_folder_tree(self, entries: List[BrowserEntry]) -> None:
        if self.certificate_folder_tree is None:
            return
        for item in self.certificate_folder_tree.get_children():
            self.certificate_folder_tree.delete(item)
        self.certificate_folder_nodes = {}
        for entry in entries:
            meta = "文件夹" if entry.entry_type == "folder" else "文件"
            node_id = self.certificate_folder_tree.insert(
                "",
                "end",
                text=entry.display_name,
                values=(meta,),
            )
            self.certificate_folder_nodes[node_id] = entry

    def render_search_tree(self, object_keys: List[str]) -> None:
        if self.search_tree is None:
            return
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        self.search_nodes = {}
        for object_key in object_keys:
            filename = Path(object_key).name
            parent_folder = str(Path(object_key).parent)
            if parent_folder == ".":
                parent_folder = "/"
            node_id = self.search_tree.insert(
                "",
                "end",
                text=filename,
                values=(parent_folder,),
            )
            self.search_nodes[node_id] = object_key

    def render_certificate_search_tree(self, object_keys: List[str]) -> None:
        if self.certificate_search_tree is None:
            return
        for item in self.certificate_search_tree.get_children():
            self.certificate_search_tree.delete(item)
        self.certificate_search_nodes = {}
        for object_key in object_keys:
            filename = Path(object_key).name
            parent_folder = str(Path(object_key).parent)
            if parent_folder == ".":
                parent_folder = "/"
            node_id = self.certificate_search_tree.insert(
                "",
                "end",
                text=filename,
                values=(parent_folder,),
            )
            self.certificate_search_nodes[node_id] = object_key

    def on_tree_select(self, _event=None) -> None:
        if self.folder_tree is None:
            return
        selected = self.folder_tree.selection()
        if not selected:
            return
        entry = self.folder_nodes.get(selected[0])
        if entry is None:
            return
        if entry.entry_type == "folder":
            self.prefix_var.set(entry.key)
            self.selected_folder_info_var.set(f"当前已选目录：{entry.key or '/'}")
        else:
            self.selected_folder_info_var.set(f"当前已选文件：{entry.key}")

    def on_tree_double_click(self, _event=None) -> None:
        if self.folder_tree is None:
            return
        selected = self.folder_tree.selection()
        if not selected:
            return
        entry = self.folder_nodes.get(selected[0])
        if entry is None:
            return
        if entry.entry_type == "folder" and entry.key:
            self.prefix_var.set(entry.key)
            self.load_bucket_folders()

    def on_certificate_tree_select(self, _event=None) -> None:
        if self.certificate_folder_tree is None:
            return
        selected = self.certificate_folder_tree.selection()
        if not selected:
            return
        entry = self.certificate_folder_nodes.get(selected[0])
        if entry is None:
            return
        if entry.entry_type == "folder":
            self.certificate_prefix_var.set(entry.key)
            self.certificate_selected_folder_info_var.set(f"当前已选目录：{entry.key or '/'}")
        else:
            self.certificate_selected_folder_info_var.set(f"当前已选文件：{entry.key}")

    def on_certificate_tree_double_click(self, _event=None) -> None:
        if self.certificate_folder_tree is None:
            return
        selected = self.certificate_folder_tree.selection()
        if not selected:
            return
        entry = self.certificate_folder_nodes.get(selected[0])
        if entry is None:
            return
        if entry.entry_type == "folder" and entry.key:
            self.certificate_prefix_var.set(entry.key)
            self.load_certificate_folders()

    def search_bucket_files(self) -> None:
        if self.photo_source_mode_var.get() != "oss":
            messagebox.showinfo("本地模式", "本地模式不使用 bucket 搜索。")
            return

        keyword = self.search_keyword_var.get().strip()
        if not keyword:
            messagebox.showinfo("缺少关键词", "请输入要搜索的文件夹名称关键词。")
            return

        current_prefix = self.prefix_var.get().strip() or "/"
        self.search_status_var.set("正在筛选当前层级文件夹...")
        self.write_log(f"开始筛选当前层级文件夹：{current_prefix}，关键词：{keyword}")
        self.finish_search(entries=self.filter_folder_entries(self.current_folder_entries, keyword))

    def search_certificate_files(self) -> None:
        if self.certificate_source_mode_var.get() != "oss":
            messagebox.showinfo("本地模式", "本地模式不使用 bucket 搜索。")
            return

        keyword = self.certificate_search_keyword_var.get().strip()
        if not keyword:
            messagebox.showinfo("缺少关键词", "请输入要搜索的文件夹名称关键词。")
            return

        current_prefix = self.certificate_prefix_var.get().strip() or "/"
        self.certificate_search_status_var.set("正在筛选当前层级文件夹...")
        self.write_log(f"开始筛选证件资料当前层级文件夹：{current_prefix}，关键词：{keyword}")
        self.finish_certificate_search(
            entries=self.filter_folder_entries(self.current_certificate_folder_entries, keyword)
        )

    def filter_folder_entries(self, entries: List[BrowserEntry], keyword: str) -> List[BrowserEntry]:
        cleaned_keyword = keyword.strip().lower()
        if not cleaned_keyword:
            return entries
        return [
            entry
            for entry in entries
            if entry.entry_type == "folder" and cleaned_keyword in entry.display_name.lower()
        ]

    def finish_search(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        if error is not None:
            self.search_status_var.set("搜索失败")
            self.write_log(f"搜索文件夹失败：{error}")
            messagebox.showerror("搜索失败", self.format_cloud_error(error))
            return

        entries = entries or []
        self.render_folder_tree(entries)
        if entries:
            self.search_status_var.set(f"找到 {len(entries)} 个匹配文件夹")
            self.write_log(f"搜索完成，找到 {len(entries)} 个匹配文件夹。")
        else:
            self.search_status_var.set("没有找到匹配文件夹")
            self.write_log("没有找到匹配文件夹。")

    def finish_certificate_search(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        if error is not None:
            self.certificate_search_status_var.set("搜索失败")
            self.write_log(f"搜索证件资料文件夹失败：{error}")
            messagebox.showerror("搜索失败", self.format_cloud_error(error))
            return

        self.render_certificate_folder_tree(entries)
        entries = entries or []
        if entries:
            self.certificate_search_status_var.set(f"找到 {len(entries)} 个匹配文件夹")
            self.write_log(f"证件资料搜索完成，找到 {len(entries)} 个匹配文件夹。")
        else:
            self.certificate_search_status_var.set("没有找到匹配文件夹")
            self.write_log("没有找到匹配的证件资料文件夹。")

    def on_search_double_click(self, _event=None) -> None:
        if self.search_tree is None:
            return
        selected = self.search_tree.selection()
        if not selected:
            return

        object_key = self.search_nodes.get(selected[0], "")
        if not object_key:
            return

        parent_folder = str(Path(object_key).parent)
        if parent_folder == ".":
            self.prefix_var.set("")
            self.selected_folder_info_var.set(f"当前已选：/，文件：{Path(object_key).name}")
        else:
            normalized_parent = parent_folder.strip("/") + "/"
            self.prefix_var.set(normalized_parent)
            self.selected_folder_info_var.set(
                f"当前已选：{normalized_parent}，文件：{Path(object_key).name}"
            )
        self.write_log(f"已根据搜索结果定位到目录：{self.prefix_var.get() or '/'}")

    def on_certificate_search_double_click(self, _event=None) -> None:
        if self.certificate_search_tree is None:
            return
        selected = self.certificate_search_tree.selection()
        if not selected:
            return

        object_key = self.certificate_search_nodes.get(selected[0], "")
        if not object_key:
            return

        parent_folder = str(Path(object_key).parent)
        if parent_folder == ".":
            self.certificate_prefix_var.set("")
            self.certificate_selected_folder_info_var.set(
                f"当前已选：/，文件：{Path(object_key).name}"
            )
        else:
            normalized_parent = parent_folder.strip("/") + "/"
            self.certificate_prefix_var.set(normalized_parent)
            self.certificate_selected_folder_info_var.set(
                f"当前已选：{normalized_parent}，文件：{Path(object_key).name}"
            )
        self.write_log(
            f"已根据证件资料搜索结果定位到目录：{self.certificate_prefix_var.get() or '/'}"
        )

    def refresh_selected_folder_count(self) -> None:
        if self.photo_source_mode_var.get() != "oss":
            messagebox.showinfo("本地模式", "本地模式不使用 bucket 统计。")
            return
        try:
            config = self.build_config()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        prefix = self.prefix_var.get().strip()
        if not prefix:
            messagebox.showinfo("未选择文件夹", "请先选择 bucket 文件夹。")
            return

        self.folder_status_var.set("正在统计已选文件夹图片数量...")
        self.selected_folder_info_var.set("正在统计图片数量...")
        self.write_log(f"统计图片数量：{prefix}")

        def worker() -> None:
            try:
                count = count_photos_in_prefix(config, prefix)
            except Exception as exc:
                self.root.after(
                    0,
                    lambda: self.finish_count_refresh(error=f"{type(exc).__name__}: {exc}"),
                )
                return
            self.root.after(0, lambda: self.finish_count_refresh(count=count, prefix=prefix))

        threading.Thread(target=worker, daemon=True).start()

    def finish_count_refresh(
        self,
        count: int = 0,
        prefix: str = "",
        error: Optional[str] = None,
    ) -> None:
        if error is not None:
            self.folder_status_var.set("统计失败")
            self.selected_folder_info_var.set("统计失败")
            self.write_log(f"统计图片数量失败：{error}")
            messagebox.showerror("统计失败", self.format_cloud_error(error))
            return

        if self.folder_tree is not None:
            for node_id, entry in self.folder_nodes.items():
                if entry.entry_type == "folder" and entry.key == prefix:
                    self.folder_tree.item(node_id, values=(f"{count} 张图片",))
                    break

        self.folder_status_var.set("统计完成")
        self.selected_folder_info_var.set(f"当前已选：{prefix}，图片数：{count}")
        self.write_log(f"{prefix} 下共有 {count} 张图片。")

    def build_certificate_options(self) -> CertificateFilterOptions:
        template_value = self.certificate_template_var.get().strip()
        match_column = self.certificate_match_column_var.get().strip()

        if not template_value:
            raise ValueError("请选择人员模板文件。")
        if not match_column:
            raise ValueError("请选择用于匹配人员文件夹的模板列。")

        source_dir = self.resolve_certificate_source_dir(require_value=True)
        output_dir = self.resolve_certificate_output_dir()

        keyword = ""
        if self.certificate_mode_var.get() == "keyword":
            keyword = self.certificate_keyword_var.get().strip()
            if not keyword:
                raise ValueError("关键词模式下请输入文件关键词。")
        if self.certificate_rename_folder_var.get() and not self.certificate_folder_name_column_var.get().strip():
            raise ValueError("请选择导出后文件夹名称列。")

        return CertificateFilterOptions(
            template_path=Path(template_value),
            source_dir=source_dir,
            output_dir=output_dir,
            match_column=match_column,
            rename_folder=self.certificate_rename_folder_var.get(),
            folder_name_column=self.certificate_folder_name_column_var.get().strip(),
            classify_output=self.certificate_classify_var.get(),
            keyword=keyword,
            dry_run=self.certificate_dry_run_var.get(),
        )

    def resolve_certificate_source_dir(self, require_value: bool = True) -> Path:
        source_value = self.certificate_source_dir_var.get().strip()
        if not source_value:
            if require_value:
                raise ValueError("请选择证件资料目录。")
            return Path(".").resolve()

        base_dir = Path(source_value)
        if self.certificate_source_mode_var.get() == "oss":
            return build_prefixed_directory(base_dir, self.certificate_prefix_var.get().strip(), "证件资料")
        return base_dir.expanduser().resolve()

    def resolve_certificate_output_dir(self) -> Path:
        output_value = self.certificate_output_dir_var.get().strip()
        if not output_value:
            raise ValueError("请选择输出目录。")
        return Path(output_value).expanduser().resolve()

    def start_photo_download_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        try:
            options = self.build_options()
            oss_config = None if options.skip_download else self.build_config()
            photo_match_values: List[str] = []
            if self.photo_source_mode_var.get() == "oss" and self.photo_filter_by_template_var.get():
                template_value = self.photo_template_var.get().strip()
                match_column = self.photo_match_column_var.get().strip()
                if not template_value:
                    raise ValueError("已勾选按模板名单下载，请先选择人员模板。")
                if not match_column:
                    raise ValueError("已勾选按模板名单下载，请先选择匹配列。")
                template_path = Path(template_value)
                if not template_path.exists():
                    raise ValueError("人员模板文件不存在，请重新选择。")
                photo_match_values = load_match_values(template_path, match_column)
                if not photo_match_values:
                    raise ValueError("照片模板匹配列没有可用数据，无法按名单下载。")
            self.save_settings()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.run_button.configure(state="disabled")
        self.photo_classify_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.status_var.set("运行中")
        self.progress_text_var.set("准备开始...")
        self.update_summary_ui(None)
        self.cancel_event.clear()
        if self.progress_bar is not None:
            self.progress_bar["value"] = 0
        self.write_log("")
        self.write_log("=" * 60)
        self.write_log("启动照片下载/模板任务。")
        if photo_match_values:
            self.write_log(
                f"本次仅下载模板中的照片，匹配列：{self.photo_match_column_var.get().strip()}，共 {len(photo_match_values)} 人。"
            )

        def runner() -> None:
            try:
                if photo_match_values and oss_config is not None:
                    download_dir, sorted_dir = resolve_photo_directories(options)
                    allowed_stems = set(photo_match_values)

                    def key_filter(object_key: str) -> bool:
                        relative_path = object_key
                        normalized_prefix = self.prefix_var.get().strip().strip("/")
                        if normalized_prefix:
                            prefix_with_slash = normalized_prefix + "/"
                            if object_key.startswith(prefix_with_slash):
                                relative_path = object_key[len(prefix_with_slash):]
                        filename_stem = Path(relative_path.lstrip("/")).stem
                        return filename_stem in allowed_stems

                    self.write_log(f"实际下载目录：{download_dir}")
                    download_result = download_objects(
                        config=oss_config,
                        prefix=options.prefix,
                        download_dir=download_dir,
                        dry_run=options.dry_run,
                        skip_existing=options.skip_existing,
                        logger=self.make_logger(),
                        progress_callback=self.make_progress_callback(),
                        cancel_event=self.cancel_event,
                        key_filter=key_filter,
                        file_filter=is_photo_key,
                        stage="download",
                    )
                    if self.cancel_event.is_set():
                        summary = WorkflowSummary(
                            download_dir=download_dir,
                            sorted_dir=sorted_dir,
                            template_path=download_dir / "照片分类模板.xlsx",
                            download_result=download_result,
                            template_file_count=0,
                            classified_count=0,
                            template_created=False,
                            cancelled=True,
                            dry_run=options.dry_run,
                        )
                    else:
                        template_result = generate_template(
                            source_dir=download_dir,
                            dry_run=options.dry_run,
                            logger=self.make_logger(),
                        )
                        summary = WorkflowSummary(
                            download_dir=download_dir,
                            sorted_dir=sorted_dir,
                            template_path=template_result.template_path,
                            download_result=download_result,
                            template_file_count=template_result.file_count,
                            classified_count=0,
                            template_created=template_result.created,
                            cancelled=False,
                            dry_run=options.dry_run,
                        )
                else:
                    summary = run_photo_download_and_template(
                        options=options,
                        oss_config=oss_config,
                        logger=self.make_logger(),
                        progress_callback=self.make_progress_callback(),
                        cancel_event=self.cancel_event,
                    )
            except Exception as exc:
                self.log_queue.put(f"__TASK_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "summary", "summary": summary})
                if self.cancel_event.is_set():
                    self.log_queue.put("__TASK_CANCELLED__")
                else:
                    self.log_queue.put("__TASK_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()

    def start_photo_classify_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        try:
            options = self.build_options()
            self.save_settings()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.run_button.configure(state="disabled")
        self.photo_classify_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.status_var.set("运行中")
        self.progress_text_var.set("准备开始...")
        self.update_summary_ui(None)
        self.cancel_event.clear()
        if self.progress_bar is not None:
            self.progress_bar["value"] = 0
        self.write_log("")
        self.write_log("=" * 60)
        self.write_log("启动照片分类任务。")

        def runner() -> None:
            try:
                summary = run_photo_classification_only(
                    options=options,
                    logger=self.make_logger(),
                    cancel_event=self.cancel_event,
                )
            except Exception as exc:
                self.log_queue.put(f"__TASK_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "summary", "summary": summary})
                if self.cancel_event.is_set():
                    self.log_queue.put("__TASK_CANCELLED__")
                else:
                    self.log_queue.put("__TASK_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()

    def start_certificate_download_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return
        if self.certificate_source_mode_var.get() != "oss":
            messagebox.showinfo("本地模式", "本地模式下不需要下载证件资料。")
            return

        try:
            source_dir = self.resolve_certificate_source_dir(require_value=True)
            certificate_config = self.build_certificate_config()
            template_path = Path(self.certificate_template_var.get().strip())
            match_column = self.certificate_match_column_var.get().strip()
            if not template_path.exists():
                raise ValueError("请先选择有效的人员模板文件。")
            if not match_column:
                raise ValueError("请先选择匹配列。")
            match_values = load_match_values(template_path, match_column)
            if not match_values:
                raise ValueError("模板匹配列没有可用数据，无法下载证件资料。")
            self.save_settings()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.certificate_download_button.configure(state="disabled")
        self.certificate_run_button.configure(state="disabled")
        self.certificate_cancel_button.configure(state="normal")
        self.certificate_status_var.set("运行中")
        self.certificate_progress_text_var.set("准备开始...")
        self.update_certificate_summary_ui(None)
        self.cancel_event.clear()
        if self.certificate_progress_bar is not None:
            self.certificate_progress_bar["value"] = 0
        self.write_log("")
        self.write_log("=" * 60)
        self.write_log("启动证件资料下载任务。")
        self.write_log(f"实际下载目录：{source_dir}")
        self.write_log(f"本次仅下载模板中的人员目录，匹配列：{match_column}，共 {len(match_values)} 人。")

        def runner() -> None:
            try:
                allowed_people = set(match_values)

                def key_filter(object_key: str) -> bool:
                    relative_path = object_key
                    normalized_prefix = self.certificate_prefix_var.get().strip().strip("/")
                    if normalized_prefix:
                        prefix_with_slash = normalized_prefix + "/"
                        if object_key.startswith(prefix_with_slash):
                            relative_path = object_key[len(prefix_with_slash):]
                    relative_path = relative_path.lstrip("/")
                    if not relative_path:
                        return False
                    person_folder = Path(relative_path).parts[0] if Path(relative_path).parts else ""
                    return person_folder in allowed_people

                download_result = download_objects(
                    config=certificate_config,
                    prefix=self.certificate_prefix_var.get().strip(),
                    download_dir=source_dir,
                    dry_run=self.certificate_dry_run_var.get(),
                    skip_existing=self.skip_existing_var.get(),
                    logger=self.make_logger(),
                    progress_callback=self.make_progress_callback(),
                    cancel_event=self.cancel_event,
                    key_filter=key_filter,
                    stage="certificate_download",
                )
                summary = CertificateFilterSummary(
                    template_path=template_path,
                    source_dir=source_dir,
                    output_dir=Path(self.certificate_output_dir_var.get().strip() or source_dir),
                    match_column=match_column,
                    rename_folder=self.certificate_rename_folder_var.get(),
                    folder_name_column=self.certificate_folder_name_column_var.get().strip(),
                    classify_output=self.certificate_classify_var.get(),
                    keyword=self.certificate_keyword_var.get().strip() if self.certificate_mode_var.get() == "keyword" else "",
                    total_rows=0,
                    matched_people=0,
                    missing_people=0,
                    copied_files=0,
                    copied_people=0,
                    download_result=download_result,
                    cancelled=self.cancel_event.is_set(),
                    dry_run=self.certificate_dry_run_var.get(),
                )
            except Exception as exc:
                self.log_queue.put(f"__CERTIFICATE_TASK_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "certificate_summary", "summary": summary})
                if self.cancel_event.is_set():
                    self.log_queue.put("__CERTIFICATE_TASK_CANCELLED__")
                else:
                    self.log_queue.put("__CERTIFICATE_TASK_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()

    def start_certificate_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        try:
            options = self.build_certificate_options()
            self.save_settings()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.certificate_download_button.configure(state="disabled")
        self.certificate_run_button.configure(state="disabled")
        self.certificate_cancel_button.configure(state="normal")
        self.certificate_status_var.set("运行中")
        self.certificate_progress_text_var.set("准备开始...")
        self.update_certificate_summary_ui(None)
        self.cancel_event.clear()
        if self.certificate_progress_bar is not None:
            self.certificate_progress_bar["value"] = 0
        self.write_log("")
        self.write_log("=" * 60)
        self.write_log("启动证件资料筛选任务。")

        def runner() -> None:
            try:
                summary = run_certificate_filter(
                    options=options,
                    logger=self.make_logger(),
                    progress_callback=self.make_progress_callback(),
                    cancel_event=self.cancel_event,
                )
            except Exception as exc:
                self.log_queue.put(f"__CERTIFICATE_TASK_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "certificate_summary", "summary": summary})
                if self.cancel_event.is_set():
                    self.log_queue.put("__CERTIFICATE_TASK_CANCELLED__")
                else:
                    self.log_queue.put("__CERTIFICATE_TASK_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()

    def start_word_export(self, variant: str) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        source_value = self.word_source_var.get().strip()
        if not source_value:
            messagebox.showerror("参数错误", "请选择表样文件。")
            return
        source_path = Path(source_value)
        if not source_path.exists():
            messagebox.showerror("参数错误", f"表样文件不存在：{source_path}")
            self.word_status_var.set("失败")
            self.word_result_var.set(f"导出失败：\n表样文件不存在：{source_path}")
            self.word_preview_status_var.set("预览不可用")
            self.set_word_code("")
            self.render_word_preview("")
            self.word_copy_button.configure(state="disabled")
            self.word_open_browser_button.configure(state="disabled")
            return
        if source_path.suffix.lower() not in {".doc", ".docx", ".xlsx"}:
            messagebox.showerror("参数错误", "仅支持 `.doc`、`.docx` 或 `.xlsx` 文件。")
            self.word_status_var.set("失败")
            self.word_result_var.set("导出失败：\n仅支持 `.doc`、`.docx` 或 `.xlsx` 文件。")
            self.word_preview_status_var.set("预览不可用")
            self.set_word_code("")
            self.render_word_preview("")
            self.word_copy_button.configure(state="disabled")
            self.word_open_browser_button.configure(state="disabled")
            return

        self.save_settings()
        self.word_net_button.configure(state="disabled")
        self.word_java_button.configure(state="disabled")
        self.word_status_var.set("导出中")
        self.word_result_var.set("正在导出 HTML...")
        self.write_log("")
        self.write_log("=" * 60)
        self.write_log(f"启动表样转换任务：{variant}")

        def runner() -> None:
            try:
                result = export_word_to_html(
                    source_path=source_path,
                    variant=variant,
                    logger=self.make_logger(),
                )
            except Exception as exc:
                self.log_queue.put(f"__WORD_EXPORT_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "word_export", "result": result})
                self.log_queue.put("__WORD_EXPORT_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()

    def start_pack_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        source_value = self.pack_source_dir_var.get().strip()
        if not source_value:
            messagebox.showerror("参数错误", "请选择待打包文件夹。")
            return

        source_dir = Path(source_value)
        if not source_dir.exists() or not source_dir.is_dir():
            messagebox.showerror("参数错误", f"待打包文件夹不存在：{source_dir}")
            return

        output_value = self.pack_output_dir_var.get().strip()
        output_dir = Path(output_value) if output_value else source_dir.parent
        self.pack_output_dir_var.set(str(output_dir))
        custom_password = self.pack_password_var.get().strip()
        if self.pack_use_custom_password_var.get() and not custom_password:
            messagebox.showerror("参数错误", "已勾选手动设置密码，请输入打包密码。")
            return
        self.save_settings()

        self.pack_run_button.configure(state="disabled")
        self.pack_copy_password_button.configure(state="disabled")
        self.pack_open_button.configure(state="disabled")
        self.pack_status_var.set("打包中")
        self.pack_result_var.set("正在压缩并加密，请稍候...")
        self.write_log(f"启动结果打包任务：{source_dir}")

        def runner() -> None:
            try:
                summary = pack_encrypted_folder(
                    source_dir=source_dir,
                    output_dir=output_dir,
                    password=custom_password if self.pack_use_custom_password_var.get() else None,
                    logger=self.make_logger(),
                )
            except Exception as exc:
                self.log_queue.put(f"__PACK_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "pack_summary", "summary": summary})
                self.log_queue.put("__PACK_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()

    def start_match_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        target_value = self.match_target_var.get().strip()
        source_value = self.match_source_var.get().strip()
        if not target_value or not source_value:
            messagebox.showerror("参数错误", "请选择目标表和来源表。")
            return
        target_key = self.match_target_key_var.get().strip()
        source_key = self.match_source_key_var.get().strip()
        if not target_key or not source_key:
            messagebox.showerror("参数错误", "请选择目标表匹配列和来源表匹配列。")
            return
        transfer_mappings = list(self.match_transfer_mappings)
        if not transfer_mappings:
            messagebox.showerror("参数错误", "请至少选择一个来源表补充列。")
            return
        extra_mappings = [
            mapping
            for mapping in self.match_extra_mappings
            if not (
                mapping.target_column == target_key
                and mapping.source_column == source_key
            )
        ]
        target_path = Path(target_value)
        source_path = Path(source_value)
        if target_path.suffix.lower() not in {".xlsx", ".xls"} or source_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showerror("参数错误", "数据匹配仅支持 `.xlsx` 或 `.xls` 文件。")
            return
        output_value = self.match_output_var.get().strip()
        output_path = Path(output_value) if output_value else target_path.with_name(
            f"{target_path.stem}_数据匹配结果.xlsx"
        )
        self.match_output_var.set(str(output_path))
        self.save_settings()

        self.match_run_button.configure(state="disabled")
        self.match_open_button.configure(state="disabled")
        self.match_status_var.set("匹配中")
        self.match_result_var.set("正在执行数据匹配，请稍候...")
        self.write_log(f"启动数据匹配任务：{target_path.name} <- {source_path.name}")

        def runner() -> None:
            try:
                summary = run_data_match(
                    DataMatchOptions(
                        target_path=target_path,
                        source_path=source_path,
                        target_key_column=target_key,
                        source_key_column=source_key,
                        extra_match_mappings=extra_mappings,
                        transfer_mappings=transfer_mappings,
                        output_path=output_path,
                    ),
                    logger=self.make_logger(),
                )
            except Exception as exc:
                self.log_queue.put(f"__MATCH_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "match_summary", "summary": summary})
                self.log_queue.put("__MATCH_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()

    def start_exam_arrange_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        candidate_value = self.exam_candidate_var.get().strip()
        group_value = self.exam_group_var.get().strip()
        plan_value = self.exam_plan_var.get().strip()
        if not candidate_value or not group_value or not plan_value:
            messagebox.showerror("参数错误", "请先选择考生明细表、岗位归组表和编排片段表。")
            return
        if not self.exam_rule_items:
            messagebox.showerror("参数错误", "请至少添加一条考号规则。")
            return
        try:
            exam_point_digits = int(self.exam_point_digits_var.get().strip())
            room_digits = int(self.exam_room_digits_var.get().strip())
            seat_digits = int(self.exam_seat_digits_var.get().strip())
            serial_digits = int(self.exam_serial_digits_var.get().strip())
        except ValueError:
            messagebox.showerror("参数错误", "考点、考场、座号、流水号位数必须是整数。")
            return

        candidate_path = Path(candidate_value)
        group_path = Path(group_value)
        plan_path = Path(plan_value)
        output_value = self.exam_output_var.get().strip()
        output_path = Path(output_value) if output_value else candidate_path.with_name(
            f"{candidate_path.stem}_考场编排结果.xlsx"
        )
        self.exam_output_var.set(str(output_path))
        self.save_settings()

        self.exam_run_button.configure(state="disabled")
        self.exam_open_button.configure(state="disabled")
        self.exam_status_var.set("编排中")
        self.exam_result_var.set("正在执行考场编排，请稍候...")
        self.set_exam_result_text(self.exam_result_var.get())
        self.write_log(f"启动考场编排任务：{candidate_path.name}")

        def runner() -> None:
            try:
                summary = run_exam_arrangement(
                    ExamArrangeOptions(
                        candidate_path=candidate_path,
                        group_path=group_path,
                        plan_path=plan_path,
                        output_path=output_path,
                        exam_point_digits=exam_point_digits,
                        room_digits=room_digits,
                        seat_digits=seat_digits,
                        serial_digits=serial_digits,
                        sort_mode=self.exam_sort_mode_var.get().strip() or "original",
                        rule_items=list(self.exam_rule_items),
                    ),
                    logger=self.make_logger(),
                )
            except Exception as exc:
                self.log_queue.put(f"__EXAM_FAILED__::{type(exc).__name__}: {exc}")
            else:
                self.log_queue.put({"type": "exam_summary", "summary": summary})
                self.log_queue.put("__EXAM_DONE__")

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()


def main() -> None:
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
