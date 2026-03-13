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
    run_photo_classification_only,
    run_photo_download_and_template,
)
from .certificate_filter import (
    CertificateFilterOptions,
    CertificateFilterSummary,
    list_template_headers,
    run_certificate_filter,
)
from .config import OssConfig, validate_oss_config, validate_oss_credentials
from .downloader import (
    BrowserEntry,
    count_photos_in_prefix,
    download_objects,
    list_buckets,
    list_browser_entries,
    list_folder_prefixes,
)
from .word_to_html import WordExportResult, export_word_to_html


class App:
    SETTINGS_FILE = Path(__file__).resolve().parents[2] / ".gui_settings.json"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("报名系统工具箱")
        self.root.geometry("1080x760")
        self.root.minsize(920, 680)

        self.log_queue: "queue.Queue[object]" = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.cancel_event = threading.Event()

        self.prefix_var = tk.StringVar()
        self.photo_source_mode_var = tk.StringVar(value="oss")
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
        self.certificate_keyword_var = tk.StringVar(value="学历证书")
        self.certificate_classify_var = tk.BooleanVar(value=True)
        self.certificate_mode_var = tk.StringVar(value="folder")
        self.certificate_bucket_name_var = tk.StringVar()
        self.certificate_prefix_var = tk.StringVar()
        self.word_source_var = tk.StringVar()

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
        self.word_result_var = tk.StringVar(value="Word 转 HTML 结果会显示在这里")
        self.word_preview_status_var = tk.StringVar(value="未生成预览")
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
        self.word_code_text = None
        self.word_preview_widget = None
        self.certificate_headers: List[str] = []
        self.certificate_bucket_values: List[str] = []
        self.certificate_folder_values: List[str] = [""]
        self.default_sash_pending = {"photo": True, "certificate": True}
        self.bucket_load_token = 0
        self.certificate_bucket_load_token = 0

        self.load_saved_settings()
        self.build_ui()
        if self.certificate_template_var.get().strip() and Path(
            self.certificate_template_var.get().strip()
        ).exists():
            self.load_certificate_headers()
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
            text="报名系统工具箱",
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
            text="从 OSS 下载后处理",
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

        self.photo_oss_frame = ttk.LabelFrame(left_frame, text="OSS 配置", padding=12)
        self.photo_oss_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
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
        options_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))

        option_specs = [
            ("仅预览，不实际执行", self.dry_run_var),
            ("下载时跳过已存在文件", self.skip_existing_var),
        ]
        for row, (label, variable) in enumerate(option_specs):
            ttk.Checkbutton(options_frame, text=label, variable=variable).grid(
                row=row // 2,
                column=row % 2,
                sticky="w",
                padx=(0, 18),
                pady=4,
            )

        action_frame = ttk.Frame(left_frame)
        action_frame.grid(row=4, column=0, sticky="ew", pady=(14, 0))

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
        progress_frame.grid(row=5, column=0, sticky="ew", pady=(12, 0))
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
        summary_frame.grid(row=6, column=0, sticky="ew", pady=(12, 0))
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
            text="直接使用本地证件资料目录",
            variable=self.certificate_source_mode_var,
            value="local",
            command=self.update_certificate_source_mode_ui,
        ).pack(side="left")
        ttk.Radiobutton(
            cert_source_frame,
            text="先从 OSS 下载证件资料",
            variable=self.certificate_source_mode_var,
            value="oss",
            command=self.update_certificate_source_mode_ui,
        ).pack(side="left", padx=(10, 0))

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

        ttk.Label(cert_form, text="筛选模式", width=16).grid(row=4, column=0, sticky="w", pady=4)
        mode_row = ttk.Frame(cert_form)
        mode_row.grid(row=4, column=1, columnspan=2, sticky="w", pady=4)
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

        ttk.Label(cert_form, text="文件关键词", width=16).grid(row=5, column=0, sticky="w", pady=4)
        self.certificate_keyword_entry = self.create_text_entry(
            cert_form,
            textvariable=self.certificate_keyword_var,
        )
        self.certificate_keyword_entry.grid(row=5, column=1, sticky="ew", pady=4)

        ttk.Checkbutton(
            cert_form,
            text="按分类一/分类二/分类三建立目录",
            variable=self.certificate_classify_var,
        ).grid(row=6, column=1, sticky="w", pady=(6, 0))

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
        notebook.add(word_frame, text="Word 转 HTML")

        word_form = ttk.LabelFrame(word_frame, text="模板转换", padding=12)
        word_form.grid(row=0, column=0, sticky="ew")
        word_form.columnconfigure(1, weight=1)

        self.add_file_row(
            word_form,
            row=0,
            label="Word 文件",
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
        self.root.after(120, lambda: self.reset_split_default("photo"))

    def set_help_content(self) -> None:
        help_content = """报名系统工具箱使用说明

一、照片下载与分类

这个功能用来下载、整理和分类考生照片。

使用步骤：
1. 先选择数据来源：
   - 从云端下载后处理
   - 直接处理本地目录
2. 如果是云端来源，填写云类型、Endpoint / Region、AccessKey、Bucket。
3. 选择下载目录和分类目录。
4. 点击“下载并生成模板”。
5. 程序会在照片目录里生成一份 Excel 模板。
6. 打开模板，填写“分类一 / 分类二 / 分类三 / 修改名称”。
7. 填写完成后，回到程序点击“按模板分类”。

补充说明：
- 分类一、分类二、分类三都可以不填。
- 修改名称也可以不填。
- 处理完成后，会自动生成“照片分类结果清单.xlsx”。

二、证件资料筛选

这个功能用来按模板筛选人员证件资料。

使用步骤：
1. 选择人员模板。
2. 选择证件资料目录。
3. 选择输出目录。
4. 选择匹配列，例如“身份证号”或“报名序号”。
5. 选择筛选模式：
   - 复制整个人员文件夹
   - 只复制关键词文件
6. 如果选择关键词模式，再填写关键词，例如“学历证书”。
7. 如有需要，可以勾选“按分类一 / 分类二 / 分类三建立目录”。
8. 点击“开始筛选”。

补充说明：
- 模板默认读取第一个 sheet。
- 处理完成后，会自动生成“证件资料筛选结果清单.xlsx”。

三、Word 转 HTML

这个功能用来把 Word 模板转换成 HTML 代码。

使用步骤：
1. 选择 Word 文件。
2. 根据需要点击“Net版导出”或“Java版导出”。
3. 导出后可以在“代码”页复制 HTML。
4. 也可以切到“预览”页查看效果。

补充说明：
- Net 版占位符格式示例：{[#考生表视图.姓名#]}
- Java 版占位符格式示例：${考生.姓名}

四、使用建议

- 第一次处理大批量文件时，建议先用少量文件测试。
- 使用模板时，建议先确认内容填写无误，再进行正式处理。
- 如果页面里已经生成“结果清单”，可以直接点击“打开结果清单”进行核对。

五、运行日志

运行日志主要用来查看处理过程。
如果需要回看详细步骤，可以打开“运行日志”页查看。
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
    ) -> None:
        ttk.Label(parent, text=label, width=16).grid(row=row, column=0, sticky="w", pady=4)
        entry = self.create_text_entry(parent, textvariable=variable, show=show or "")
        entry.grid(row=row, column=1, sticky="ew", pady=4)

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

    def choose_word_source(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path(self.word_source_var.get()).parent)
            if self.word_source_var.get().strip()
            else str(Path.cwd()),
            filetypes=[("Word 文件", "*.docx *.doc"), ("所有文件", "*.*")],
        )
        if selected:
            self.word_source_var.set(selected)

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
        )
        self.enable_entry_cursor(entry_widget)
        return entry_widget

    def enable_entry_cursor(self, entry_widget) -> None:
        def focus_entry(event) -> None:
            try:
                def apply_focus() -> None:
                    try:
                        entry_widget.focus_force()
                        cursor_index = entry_widget.index(f"@{event.x}")
                        entry_widget.icursor(cursor_index)
                        entry_widget.select_clear()
                    except tk.TclError:
                        pass

                self.root.after_idle(apply_focus)
            except tk.TclError:
                pass

        entry_widget.bind("<ButtonRelease-1>", focus_entry, add="+")

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
                elif message == "__WORD_EXPORT_DONE__":
                    self.word_net_button.configure(state="normal")
                    self.word_java_button.configure(state="normal")
                    self.word_status_var.set("完成")
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
            self.word_result_var.set("Word 转 HTML 结果会显示在这里")
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

    def load_saved_settings(self) -> None:
        if not self.SETTINGS_FILE.exists():
            return
        try:
            settings = json.loads(self.SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return

        self.prefix_var.set(settings.get("prefix", ""))
        self.download_dir_var.set(settings.get("download_dir", self.download_dir_var.get()))
        self.sorted_dir_var.set(settings.get("sorted_dir", self.sorted_dir_var.get()))
        self.cloud_type_var.set(settings.get("cloud_type", self.cloud_type_var.get()))
        self.access_key_id_var.set(settings.get("access_key_id", ""))
        self.access_key_secret_var.set(settings.get("access_key_secret", ""))
        self.endpoint_var.set(settings.get("endpoint", self.endpoint_var.get()))
        self.bucket_name_var.set(settings.get("bucket_name", ""))
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
        self.certificate_keyword_var.set(
            settings.get("certificate_keyword", self.certificate_keyword_var.get())
        )
        self.certificate_classify_var.set(
            settings.get("certificate_classify", self.certificate_classify_var.get())
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

    def save_settings(self) -> None:
        settings = {
            "prefix": self.prefix_var.get().strip(),
            "download_dir": self.download_dir_var.get().strip(),
            "sorted_dir": self.sorted_dir_var.get().strip(),
            "cloud_type": self.cloud_type_var.get().strip(),
            "access_key_id": self.access_key_id_var.get().strip(),
            "access_key_secret": self.access_key_secret_var.get().strip(),
            "endpoint": self.endpoint_var.get().strip(),
            "bucket_name": self.bucket_name_var.get().strip(),
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
            "certificate_keyword": self.certificate_keyword_var.get().strip(),
            "certificate_classify": self.certificate_classify_var.get(),
            "certificate_mode": self.certificate_mode_var.get().strip(),
            "certificate_source_mode": self.certificate_source_mode_var.get().strip(),
            "certificate_bucket_name": self.certificate_bucket_name_var.get().strip(),
            "certificate_prefix": self.certificate_prefix_var.get().strip(),
            "word_source": self.word_source_var.get().strip(),
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

    def on_cloud_type_changed(self) -> None:
        current_endpoint = self.endpoint_var.get().strip()
        if self.cloud_type_var.get() == "aliyun":
            if not current_endpoint or "myqcloud.com" in current_endpoint or current_endpoint.startswith("ap-"):
                self.endpoint_var.set("https://oss-cn-hangzhou.aliyuncs.com")
        else:
            if not current_endpoint or "aliyuncs.com" in current_endpoint:
                self.endpoint_var.set("ap-beijing")

        self.bucket_name_var.set("")
        self.certificate_bucket_name_var.set("")
        self.prefix_var.set("")
        self.certificate_prefix_var.set("")
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
        current_value = self.certificate_match_column_var.get().strip()
        if current_value and current_value not in headers:
            self.certificate_match_column_var.set("")
        if not self.certificate_match_column_var.get().strip() and headers:
            self.certificate_match_column_var.set(headers[0])

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
            self.load_bucket_folders()
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
            self.load_certificate_folders()
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
        source_value = self.certificate_source_dir_var.get().strip()
        output_value = self.certificate_output_dir_var.get().strip()
        match_column = self.certificate_match_column_var.get().strip()

        if not template_value:
            raise ValueError("请选择人员模板文件。")
        if not source_value:
            raise ValueError("请选择证件资料目录。")
        if not output_value:
            raise ValueError("请选择输出目录。")
        if not match_column:
            raise ValueError("请选择用于匹配人员文件夹的模板列。")

        keyword = ""
        if self.certificate_mode_var.get() == "keyword":
            keyword = self.certificate_keyword_var.get().strip()
            if not keyword:
                raise ValueError("关键词模式下请输入文件关键词。")

        return CertificateFilterOptions(
            template_path=Path(template_value),
            source_dir=Path(source_value),
            output_dir=Path(output_value),
            match_column=match_column,
            classify_output=self.certificate_classify_var.get(),
            keyword=keyword,
            dry_run=self.dry_run_var.get(),
        )

    def start_photo_download_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("任务执行中", "当前任务还没结束。")
            return

        try:
            options = self.build_options()
            oss_config = None if options.skip_download else self.build_config()
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

        def runner() -> None:
            try:
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
            options = self.build_certificate_options()
            certificate_config = self.build_certificate_config()
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

        def runner() -> None:
            try:
                download_result = download_objects(
                    config=certificate_config,
                    prefix=self.certificate_prefix_var.get().strip(),
                    download_dir=options.source_dir,
                    dry_run=options.dry_run,
                    skip_existing=self.skip_existing_var.get(),
                    logger=self.make_logger(),
                    progress_callback=self.make_progress_callback(),
                    cancel_event=self.cancel_event,
                    stage="certificate_download",
                )
                summary = CertificateFilterSummary(
                    template_path=options.template_path,
                    source_dir=options.source_dir,
                    output_dir=options.output_dir,
                    match_column=options.match_column,
                    classify_output=options.classify_output,
                    keyword=options.keyword.strip(),
                    total_rows=0,
                    matched_people=0,
                    missing_people=0,
                    copied_files=0,
                    copied_people=0,
                    download_result=download_result,
                    cancelled=self.cancel_event.is_set(),
                    dry_run=options.dry_run,
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
            messagebox.showerror("参数错误", "请选择 Word 文件。")
            return
        source_path = Path(source_value)
        if not source_path.exists():
            messagebox.showerror("参数错误", f"Word 文件不存在：{source_path}")
            self.word_status_var.set("失败")
            self.word_result_var.set(f"导出失败：\nWord 文件不存在：{source_path}")
            self.word_preview_status_var.set("预览不可用")
            self.set_word_code("")
            self.render_word_preview("")
            self.word_copy_button.configure(state="disabled")
            self.word_open_browser_button.configure(state="disabled")
            return
        if source_path.suffix.lower() not in {".doc", ".docx"}:
            messagebox.showerror("参数错误", "仅支持 `.doc` 或 `.docx` 文件。")
            self.word_status_var.set("失败")
            self.word_result_var.set("导出失败：\n仅支持 `.doc` 或 `.docx` 文件。")
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
        self.write_log(f"启动 Word 转 HTML 任务：{variant}")

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


def main() -> None:
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
