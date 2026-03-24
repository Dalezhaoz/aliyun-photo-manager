from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .common import TOOLBAR_BUTTON_WIDTH


def build_photo_tab(app, notebook: ttk.Notebook) -> None:
    settings_frame = ttk.Frame(notebook, padding=14)
    settings_frame.columnconfigure(0, weight=1)
    settings_frame.rowconfigure(1, weight=1)
    notebook.add(settings_frame, text="照片下载与分类")

    intro_frame = tk.Frame(
        settings_frame,
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
        text="照片下载与分类",
        bg="#EAF2FF",
        fg="#162033",
        font=("Microsoft YaHei UI", 18, "bold"),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        intro_frame,
        text="适合从云端或本地目录批量下载照片，生成模板后再按分类字段整理结果。建议先预览，再正式执行。",
        bg="#EAF2FF",
        fg="#5A6D83",
        font=("Microsoft YaHei UI", 10),
        anchor="w",
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(8, 0))

    photo_paned = ttk.Panedwindow(settings_frame, orient="horizontal")
    photo_paned.grid(row=1, column=0, sticky="nsew")
    app.photo_paned = photo_paned

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
    app.bind_mousewheel_to_canvas(left_canvas, left_frame)

    app.photo_right_frame = ttk.Frame(settings_frame)
    app.photo_right_frame.columnconfigure(0, weight=1)
    app.photo_right_frame.rowconfigure(0, weight=1)
    photo_paned.add(left_outer, weight=1)
    photo_paned.add(app.photo_right_frame, weight=1)

    photo_source_frame = ttk.LabelFrame(left_frame, text="1. 数据来源", padding=14)
    photo_source_frame.grid(row=0, column=0, sticky="ew")
    ttk.Radiobutton(
        photo_source_frame,
        text="从云存储下载后处理",
        variable=app.photo_source_mode_var,
        value="oss",
        command=app.update_photo_source_mode_ui,
    ).pack(side="left")
    ttk.Radiobutton(
        photo_source_frame,
        text="直接处理本地目录",
        variable=app.photo_source_mode_var,
        value="local",
        command=app.update_photo_source_mode_ui,
    ).pack(side="left", padx=(10, 0))

    paths_frame = ttk.LabelFrame(left_frame, text="2. 输出目录", padding=14)
    paths_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
    paths_frame.columnconfigure(1, weight=1)
    app.add_path_row(paths_frame, row=0, label="下载目录", variable=app.download_dir_var)
    app.add_path_row(paths_frame, row=1, label="分类目录", variable=app.sorted_dir_var)

    photo_template_frame = ttk.LabelFrame(left_frame, text="3. 名单筛选", padding=14)
    photo_template_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
    photo_template_frame.columnconfigure(1, weight=1)
    app.add_file_row(
        photo_template_frame,
        row=0,
        label="人员模板",
        variable=app.photo_template_var,
        command=app.choose_photo_template,
        button_text="选择文件",
    )
    ttk.Label(photo_template_frame, text="匹配列", width=16).grid(row=1, column=0, sticky="w", pady=4)
    photo_match_row = ttk.Frame(photo_template_frame)
    photo_match_row.grid(row=1, column=1, columnspan=2, sticky="ew", pady=4)
    photo_match_row.columnconfigure(0, weight=1)
    app.photo_match_combo = ttk.Combobox(
        photo_match_row,
        textvariable=app.photo_match_column_var,
        values=app.photo_headers,
        state="readonly",
    )
    app.photo_match_combo.grid(row=0, column=0, sticky="ew")
    ttk.Button(photo_match_row, text="加载模板列", command=app.load_photo_headers).grid(
        row=0, column=1, padx=(8, 0)
    )
    app.add_tick_checkbutton(
        photo_template_frame,
        text="只下载模板中的人员",
        variable=app.photo_filter_by_template_var,
    ).grid(row=2, column=1, sticky="w", pady=(6, 0))

    app.photo_oss_frame = ttk.LabelFrame(left_frame, text="4. 云存储配置", padding=14)
    app.photo_oss_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))
    app.photo_oss_frame.columnconfigure(1, weight=1)
    app.add_cloud_type_row(app.photo_oss_frame, 0)
    app.add_entry_row(app.photo_oss_frame, 1, "Endpoint / Region", app.endpoint_var)
    app.add_entry_row(app.photo_oss_frame, 2, "AccessKey ID", app.access_key_id_var)
    app.add_entry_row(
        app.photo_oss_frame,
        3,
        "AccessKey Secret",
        app.access_key_secret_var,
        show="*",
    )
    app.add_bucket_picker(app.photo_oss_frame, 4)
    app.add_prefix_picker(app.photo_oss_frame, 5)

    options_frame = ttk.LabelFrame(left_frame, text="5. 执行选项", padding=14)
    options_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
    for row, (label, variable) in enumerate(
        [
            ("仅预览，不实际执行", app.dry_run_var),
            ("下载时跳过已存在文件", app.skip_existing_var),
        ]
    ):
        app.add_tick_checkbutton(options_frame, text=label, variable=variable).grid(
            row=row // 2,
            column=row % 2,
            sticky="w",
            padx=(0, 18),
            pady=4,
        )

    action_frame = ttk.LabelFrame(left_frame, text="6. 执行操作", padding=14)
    action_frame.grid(row=5, column=0, sticky="ew", pady=(14, 0))
    app.run_button = ttk.Button(
        action_frame,
        text="下载并生成模板",
        command=app.start_photo_download_run,
        style="Accent.TButton",
        width=16,
    )
    app.run_button.pack(side="left")
    app.photo_classify_button = ttk.Button(
        action_frame,
        text="按模板分类",
        command=app.start_photo_classify_run,
        style="Accent.TButton",
        width=14,
    )
    app.photo_classify_button.pack(side="left", padx=(10, 0))
    app.cancel_button = ttk.Button(
        action_frame,
        text="取消下载",
        command=app.cancel_run,
        state="disabled",
        width=12,
    )
    app.cancel_button.pack(side="left", padx=(10, 0))
    ttk.Label(action_frame, textvariable=app.status_var).pack(side="right")

    progress_frame = ttk.LabelFrame(left_frame, text="7. 执行进度", padding=14)
    progress_frame.grid(row=6, column=0, sticky="ew", pady=(12, 0))
    progress_frame.columnconfigure(0, weight=1)
    app.progress_bar = ttk.Progressbar(
        progress_frame,
        orient="horizontal",
        mode="determinate",
        maximum=100,
    )
    app.progress_bar.grid(row=0, column=0, sticky="ew")
    ttk.Label(progress_frame, textvariable=app.progress_text_var).grid(
        row=1, column=0, sticky="w", pady=(8, 0)
    )

    summary_frame = ttk.LabelFrame(left_frame, text="8. 结果与输出", padding=14)
    summary_frame.grid(row=7, column=0, sticky="ew", pady=(12, 0))
    summary_frame.columnconfigure(0, weight=1)
    ttk.Label(
        summary_frame,
        textvariable=app.summary_text_var,
        justify="left",
        wraplength=320,
    ).grid(row=0, column=0, sticky="w")
    app.open_template_button = ttk.Button(
        summary_frame,
        text="打开 Excel 模板",
        command=app.open_template_file,
        state="disabled",
        width=16,
    )
    app.open_template_button.grid(row=1, column=0, sticky="w", pady=(8, 0))
    app.open_photo_report_button = ttk.Button(
        summary_frame,
        text="打开结果清单",
        command=app.open_photo_report_file,
        state="disabled",
        width=16,
    )
    app.open_photo_report_button.grid(row=2, column=0, sticky="w", pady=(8, 0))

    app.photo_browser_frame = ttk.LabelFrame(app.photo_right_frame, text="Bucket 文件夹浏览", padding=14)
    app.photo_browser_frame.grid(row=0, column=0, sticky="nsew")
    app.photo_browser_frame.columnconfigure(0, weight=1)
    app.photo_browser_frame.rowconfigure(2, weight=1)

    ttk.Label(
        app.photo_browser_frame,
        text="从右侧浏览当前 bucket 层级，双击进入子目录，确认目标前缀后再执行下载。",
        justify="left",
        wraplength=480,
    ).grid(row=0, column=0, sticky="w", pady=(0, 10))

    browser_toolbar = ttk.Frame(app.photo_browser_frame)
    browser_toolbar.grid(row=1, column=0, sticky="ew")
    browser_toolbar.columnconfigure(3, weight=1)
    ttk.Button(
        browser_toolbar,
        text="加载当前层级",
        width=TOOLBAR_BUTTON_WIDTH,
        command=app.load_bucket_folders,
    ).grid(row=0, column=0, sticky="w")
    ttk.Button(
        browser_toolbar,
        text="刷新已选文件夹数量",
        command=app.refresh_selected_folder_count,
        width=TOOLBAR_BUTTON_WIDTH,
    ).grid(row=0, column=1, sticky="w", padx=(8, 0))
    ttk.Button(
        browser_toolbar,
        text="上一级",
        width=TOOLBAR_BUTTON_WIDTH,
        command=app.go_to_parent_prefix,
    ).grid(row=0, column=2, sticky="w", padx=(8, 0))

    search_toolbar = ttk.Frame(app.photo_browser_frame)
    search_toolbar.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    search_toolbar.columnconfigure(0, weight=1)
    search_entry = app.create_text_entry(search_toolbar, textvariable=app.search_keyword_var)
    search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    ttk.Button(
        search_toolbar,
        text="搜索当前层级文件夹",
        width=TOOLBAR_BUTTON_WIDTH,
        command=app.search_bucket_files,
    ).grid(row=0, column=1, sticky="e")

    tree_frame = ttk.Frame(app.photo_browser_frame)
    tree_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 10))
    tree_frame.columnconfigure(0, weight=1)
    tree_frame.rowconfigure(0, weight=1)
    app.folder_tree = ttk.Treeview(
        tree_frame,
        columns=("meta",),
        show="tree headings",
        selectmode="browse",
        height=10,
    )
    app.folder_tree.heading("#0", text="名称")
    app.folder_tree.heading("meta", text="信息")
    app.folder_tree.column("#0", width=280, anchor="w")
    app.folder_tree.column("meta", width=140, anchor="center")
    app.folder_tree.grid(row=0, column=0, sticky="nsew")
    app.folder_tree.bind("<<TreeviewSelect>>", app.on_tree_select)
    app.folder_tree.bind("<Double-1>", app.on_tree_double_click)
    tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=app.folder_tree.yview)
    tree_scrollbar.grid(row=0, column=1, sticky="ns")
    app.folder_tree.configure(yscrollcommand=tree_scrollbar.set)

    ttk.Label(app.photo_browser_frame, textvariable=app.folder_status_var).grid(row=4, column=0, sticky="w")
    ttk.Label(app.photo_browser_frame, textvariable=app.selected_folder_info_var).grid(
        row=5, column=0, sticky="w", pady=(6, 0)
    )
    ttk.Label(app.photo_browser_frame, textvariable=app.search_status_var).grid(
        row=6, column=0, sticky="w", pady=(10, 0)
    )
