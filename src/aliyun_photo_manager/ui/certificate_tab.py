from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .common import TOOLBAR_BUTTON_WIDTH


def build_certificate_tab(app, notebook: ttk.Notebook) -> None:
    certificate_frame = ttk.Frame(notebook, padding=14)
    certificate_frame.columnconfigure(0, weight=1)
    certificate_frame.rowconfigure(1, weight=1)
    notebook.add(certificate_frame, text="证件资料筛选")

    intro_frame = tk.Frame(
        certificate_frame,
        bg="#EAF2FF",
        highlightthickness=1,
        highlightbackground="#D1DDF3",
        padx=18,
        pady=16,
    )
    intro_frame.grid(row=0, column=0, sticky="ew", pady=(0, 14))
    intro_frame.grid_columnconfigure(0, weight=1)
    tk.Label(
        intro_frame,
        text="证件资料筛选",
        bg="#EAF2FF",
        fg="#162033",
        font=("Microsoft YaHei UI", 18, "bold"),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        intro_frame,
        text="按模板筛选证件材料，支持云端下载、本地目录处理、关键词筛选和按分类目录导出。",
        bg="#EAF2FF",
        fg="#5A6D83",
        font=("Microsoft YaHei UI", 10),
        anchor="w",
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(8, 0))

    cert_paned = ttk.Panedwindow(certificate_frame, orient="horizontal")
    cert_paned.grid(row=1, column=0, sticky="nsew")
    app.certificate_paned = cert_paned

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
    app.bind_mousewheel_to_canvas(cert_left_canvas, cert_left_frame)

    cert_right_frame = ttk.Frame(cert_paned)
    cert_right_frame.columnconfigure(0, weight=1)
    cert_right_frame.rowconfigure(0, weight=1)
    cert_paned.add(cert_left_outer, weight=1)
    cert_paned.add(cert_right_frame, weight=1)

    cert_source_frame = ttk.LabelFrame(cert_left_frame, text="1. 数据来源", padding=14)
    cert_source_frame.grid(row=0, column=0, sticky="ew")
    ttk.Radiobutton(
        cert_source_frame,
        text="从云存储下载后处理",
        variable=app.certificate_source_mode_var,
        value="oss",
        command=app.update_certificate_source_mode_ui,
    ).pack(side="left")
    ttk.Radiobutton(
        cert_source_frame,
        text="直接处理本地目录",
        variable=app.certificate_source_mode_var,
        value="local",
        command=app.update_certificate_source_mode_ui,
    ).pack(side="left", padx=(10, 0))

    app.certificate_oss_frame = ttk.LabelFrame(cert_left_frame, text="2. 云存储配置", padding=14)
    app.certificate_oss_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
    app.certificate_oss_frame.columnconfigure(1, weight=1)
    app.add_cloud_type_row(app.certificate_oss_frame, 0)
    app.add_entry_row(app.certificate_oss_frame, 1, "Endpoint / Region", app.endpoint_var)
    app.add_entry_row(app.certificate_oss_frame, 2, "AccessKey ID", app.access_key_id_var)
    app.add_entry_row(app.certificate_oss_frame, 3, "AccessKey Secret", app.access_key_secret_var, show="*")
    app.add_certificate_bucket_picker(app.certificate_oss_frame, 4)
    app.add_certificate_prefix_picker(app.certificate_oss_frame, 5)

    cert_form = ttk.LabelFrame(cert_left_frame, text="3. 筛选设置", padding=14)
    cert_form.grid(row=2, column=0, sticky="ew", pady=(12, 0))
    cert_form.columnconfigure(1, weight=1)
    app.add_file_row(
        cert_form,
        row=0,
        label="人员模板",
        variable=app.certificate_template_var,
        command=app.choose_certificate_template,
        button_text="选择文件",
    )
    app.add_path_row(cert_form, row=1, label="本地证件资料目录", variable=app.certificate_source_dir_var)
    app.add_path_row(cert_form, row=2, label="输出目录", variable=app.certificate_output_dir_var)

    ttk.Label(cert_form, text="匹配列", width=16).grid(row=3, column=0, sticky="w", pady=4)
    match_row = ttk.Frame(cert_form)
    match_row.grid(row=3, column=1, columnspan=2, sticky="ew", pady=4)
    match_row.columnconfigure(0, weight=1)
    app.certificate_match_combo = ttk.Combobox(
        match_row,
        textvariable=app.certificate_match_column_var,
        values=app.certificate_headers,
        state="readonly",
    )
    app.certificate_match_combo.grid(row=0, column=0, sticky="ew")
    ttk.Button(match_row, text="加载模板列", command=app.load_certificate_headers).grid(
        row=0, column=1, padx=(8, 0)
    )

    rename_row = ttk.Frame(cert_form)
    rename_row.grid(row=4, column=1, columnspan=2, sticky="ew", pady=4)
    rename_row.columnconfigure(1, weight=1)
    app.add_tick_checkbutton(
        rename_row,
        text="导出后文件夹重命名",
        variable=app.certificate_rename_folder_var,
        command=app.update_certificate_mode_ui,
    ).grid(row=0, column=0, sticky="w")
    app.certificate_folder_name_combo = ttk.Combobox(
        rename_row,
        textvariable=app.certificate_folder_name_column_var,
        values=app.certificate_headers,
        state="readonly",
    )
    app.certificate_folder_name_combo.grid(row=0, column=1, sticky="ew", padx=(8, 0))

    ttk.Label(cert_form, text="筛选模式", width=16).grid(row=5, column=0, sticky="w", pady=4)
    mode_row = ttk.Frame(cert_form)
    mode_row.grid(row=5, column=1, columnspan=2, sticky="w", pady=4)
    ttk.Radiobutton(
        mode_row,
        text="复制整个人员文件夹",
        variable=app.certificate_mode_var,
        value="folder",
        command=app.update_certificate_mode_ui,
    ).pack(side="left")
    ttk.Radiobutton(
        mode_row,
        text="只复制关键词文件",
        variable=app.certificate_mode_var,
        value="keyword",
        command=app.update_certificate_mode_ui,
    ).pack(side="left", padx=(10, 0))

    ttk.Label(cert_form, text="文件关键词", width=16).grid(row=6, column=0, sticky="w", pady=4)
    app.certificate_keyword_entry = app.create_text_entry(cert_form, textvariable=app.certificate_keyword_var)
    app.certificate_keyword_entry.grid(row=6, column=1, sticky="ew", pady=4)
    app.add_tick_checkbutton(
        cert_form,
        text="按分类一/分类二/分类三建立目录",
        variable=app.certificate_classify_var,
    ).grid(row=7, column=1, sticky="w", pady=(6, 0))
    app.add_tick_checkbutton(
        cert_form,
        text="仅预览，不实际执行",
        variable=app.certificate_dry_run_var,
    ).grid(row=8, column=1, sticky="w", pady=(6, 0))

    cert_action = ttk.LabelFrame(cert_left_frame, text="4. 执行操作", padding=14)
    cert_action.grid(row=3, column=0, sticky="ew", pady=(12, 0))
    app.certificate_download_button = ttk.Button(
        cert_action,
        text="下载证件资料",
        command=app.start_certificate_download_run,
        style="Accent.TButton",
        width=14,
    )
    app.certificate_download_button.pack(side="left")
    app.certificate_run_button = ttk.Button(
        cert_action,
        text="开始筛选",
        command=app.start_certificate_run,
        style="Accent.TButton",
        width=12,
    )
    app.certificate_run_button.pack(side="left", padx=(10, 0))
    app.certificate_cancel_button = ttk.Button(
        cert_action,
        text="取消任务",
        command=app.cancel_run,
        state="disabled",
        width=12,
    )
    app.certificate_cancel_button.pack(side="left", padx=(10, 0))
    ttk.Label(cert_action, textvariable=app.certificate_status_var).pack(side="right")

    cert_progress_frame = ttk.LabelFrame(cert_left_frame, text="5. 筛选进度", padding=14)
    cert_progress_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
    cert_progress_frame.columnconfigure(0, weight=1)
    app.certificate_progress_bar = ttk.Progressbar(
        cert_progress_frame,
        orient="horizontal",
        mode="determinate",
        maximum=100,
    )
    app.certificate_progress_bar.grid(row=0, column=0, sticky="ew")
    ttk.Label(cert_progress_frame, textvariable=app.certificate_progress_text_var).grid(
        row=1, column=0, sticky="w", pady=(8, 0)
    )

    cert_summary_frame = ttk.LabelFrame(cert_left_frame, text="6. 筛选结果", padding=14)
    cert_summary_frame.grid(row=5, column=0, sticky="ew", pady=(12, 0))
    cert_summary_frame.columnconfigure(0, weight=1)
    ttk.Label(
        cert_summary_frame,
        textvariable=app.certificate_summary_text_var,
        justify="left",
        wraplength=320,
    ).grid(row=0, column=0, sticky="w")
    app.open_certificate_report_button = ttk.Button(
        cert_summary_frame,
        text="打开结果清单",
        command=app.open_certificate_report_file,
        state="disabled",
        width=16,
    )
    app.open_certificate_report_button.grid(row=1, column=0, sticky="w", pady=(8, 0))

    app.certificate_browser_frame = ttk.LabelFrame(cert_right_frame, text="Bucket 文件夹浏览", padding=14)
    app.certificate_browser_frame.grid(row=0, column=0, sticky="nsew")
    app.certificate_browser_frame.columnconfigure(0, weight=1)
    app.certificate_browser_frame.rowconfigure(2, weight=1)

    ttk.Label(
        app.certificate_browser_frame,
        text="从右侧浏览证件资料目录，定位到正确前缀后再下载或筛选。",
        justify="left",
        wraplength=480,
    ).grid(row=0, column=0, sticky="w", pady=(0, 10))

    cert_browser_toolbar = ttk.Frame(app.certificate_browser_frame)
    cert_browser_toolbar.grid(row=1, column=0, sticky="ew")
    cert_browser_toolbar.columnconfigure(2, weight=1)
    ttk.Button(
        cert_browser_toolbar,
        text="加载当前层级",
        width=TOOLBAR_BUTTON_WIDTH,
        command=app.load_certificate_folders,
    ).grid(row=0, column=0, sticky="w")
    ttk.Button(
        cert_browser_toolbar,
        text="上一级",
        width=TOOLBAR_BUTTON_WIDTH,
        command=app.go_to_certificate_parent_prefix,
    ).grid(row=0, column=1, sticky="w", padx=(8, 0))

    cert_search_toolbar = ttk.Frame(app.certificate_browser_frame)
    cert_search_toolbar.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    cert_search_toolbar.columnconfigure(0, weight=1)
    app.certificate_search_keyword_var = tk.StringVar()
    cert_search_entry = app.create_text_entry(
        cert_search_toolbar,
        textvariable=app.certificate_search_keyword_var,
    )
    cert_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    ttk.Button(
        cert_search_toolbar,
        text="搜索当前层级文件夹",
        width=TOOLBAR_BUTTON_WIDTH,
        command=app.search_certificate_files,
    ).grid(row=0, column=1, sticky="e")

    cert_tree_frame = ttk.Frame(app.certificate_browser_frame)
    cert_tree_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 10))
    cert_tree_frame.columnconfigure(0, weight=1)
    cert_tree_frame.rowconfigure(0, weight=1)
    app.certificate_folder_tree = ttk.Treeview(
        cert_tree_frame,
        columns=("meta",),
        show="tree headings",
        selectmode="browse",
        height=10,
    )
    app.certificate_folder_tree.heading("#0", text="名称")
    app.certificate_folder_tree.heading("meta", text="信息")
    app.certificate_folder_tree.column("#0", width=280, anchor="w")
    app.certificate_folder_tree.column("meta", width=140, anchor="center")
    app.certificate_folder_tree.grid(row=0, column=0, sticky="nsew")
    app.certificate_folder_tree.bind("<<TreeviewSelect>>", app.on_certificate_tree_select)
    app.certificate_folder_tree.bind("<Double-1>", app.on_certificate_tree_double_click)
    cert_tree_scrollbar = ttk.Scrollbar(
        cert_tree_frame,
        orient="vertical",
        command=app.certificate_folder_tree.yview,
    )
    cert_tree_scrollbar.grid(row=0, column=1, sticky="ns")
    app.certificate_folder_tree.configure(yscrollcommand=cert_tree_scrollbar.set)

    ttk.Label(app.certificate_browser_frame, textvariable=app.certificate_folder_status_var).grid(
        row=4, column=0, sticky="w"
    )
    ttk.Label(
        app.certificate_browser_frame,
        textvariable=app.certificate_selected_folder_info_var,
    ).grid(row=5, column=0, sticky="w", pady=(6, 0))
    ttk.Label(app.certificate_browser_frame, textvariable=app.certificate_search_status_var).grid(
        row=6, column=0, sticky="w", pady=(10, 0)
    )
