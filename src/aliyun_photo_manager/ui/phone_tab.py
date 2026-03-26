from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .common import create_scrollable_tab


def build_phone_tab(app, notebook: ttk.Notebook) -> None:
    phone_tab, phone_frame = create_scrollable_tab(notebook)
    phone_frame.columnconfigure(0, weight=1)
    phone_frame.rowconfigure(3, weight=1)
    notebook.add(phone_tab, text="电话解密")

    intro_frame = tk.Frame(
        phone_frame,
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
        text="电话解密",
        bg="#EAF2FF",
        fg="#162033",
        font=("Microsoft YaHei UI", 18, "bold"),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        intro_frame,
        text="按主键编号关联 `web_info.info1`，通过 helper 解密后回写到考生表 `备用3`。支持全部解密或按名单部分解密。",
        bg="#EAF2FF",
        fg="#5A6D83",
        font=("Microsoft YaHei UI", 10),
        anchor="w",
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(8, 0))

    top_grid = ttk.Frame(phone_frame)
    top_grid.grid(row=1, column=0, sticky="ew")
    top_grid.columnconfigure(0, weight=1)
    top_grid.columnconfigure(1, weight=1)

    db_frame = ttk.LabelFrame(top_grid, text="1. 数据库连接", padding=14)
    db_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    db_frame.columnconfigure(1, weight=1)
    db_frame.columnconfigure(3, weight=1)
    ttk.Label(db_frame, text="服务器", width=16).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 10))
    app.phone_server_entry = app.create_text_entry(db_frame, textvariable=app.phone_server_var)
    app.phone_server_entry.grid(row=0, column=1, sticky="ew", pady=6)
    ttk.Label(db_frame, text="端口", width=12).grid(row=0, column=2, sticky="w", pady=6, padx=(18, 10))
    app.phone_port_entry = app.create_text_entry(db_frame, textvariable=app.phone_port_var)
    app.phone_port_entry.grid(row=0, column=3, sticky="ew", pady=6)

    ttk.Label(db_frame, text="用户名", width=16).grid(row=1, column=0, sticky="w", pady=6, padx=(0, 10))
    app.phone_username_entry = app.create_text_entry(db_frame, textvariable=app.phone_username_var)
    app.phone_username_entry.grid(row=1, column=1, sticky="ew", pady=6)
    ttk.Label(db_frame, text="报名数据库", width=12).grid(row=1, column=2, sticky="w", pady=6, padx=(18, 10))
    app.phone_signup_database_entry = app.create_text_entry(
        db_frame, textvariable=app.phone_signup_database_var
    )
    app.phone_signup_database_entry.grid(row=1, column=3, sticky="ew", pady=6)

    ttk.Label(db_frame, text="密码", width=16).grid(row=2, column=0, sticky="w", pady=6, padx=(0, 10))
    app.phone_password_entry = app.create_text_entry(db_frame, textvariable=app.phone_password_var, show="*")
    app.phone_password_entry.grid(row=2, column=1, sticky="ew", pady=6)
    ttk.Label(db_frame, text="电话数据库", width=12).grid(row=2, column=2, sticky="w", pady=6, padx=(18, 10))
    app.phone_info_database_entry = app.create_text_entry(
        db_frame, textvariable=app.phone_info_database_var
    )
    app.phone_info_database_entry.grid(row=2, column=3, sticky="ew", pady=6)

    options_frame = ttk.LabelFrame(top_grid, text="2. 解密参数", padding=14)
    options_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
    options_frame.columnconfigure(1, weight=1)
    options_frame.columnconfigure(3, weight=1)

    app.add_entry_row(options_frame, 0, "考试代码", app.phone_exam_sort_var)
    ttk.Button(options_frame, text="生成表名", command=app.fill_phone_table_name, width=12).grid(
        row=0, column=2, padx=(10, 0), pady=6
    )
    app.add_entry_row(options_frame, 1, "考生表名", app.phone_candidate_table_var)

    ttk.Label(options_frame, text="解密模式", width=16).grid(row=2, column=0, sticky="w", pady=4, padx=(0, 10))
    mode_frame = ttk.Frame(options_frame)
    mode_frame.grid(row=2, column=1, columnspan=3, sticky="w", pady=4)
    ttk.Radiobutton(
        mode_frame,
        text="解密全部电话",
        variable=app.phone_mode_var,
        value="all",
        command=app.update_phone_mode_ui,
    ).pack(side="left")
    ttk.Radiobutton(
        mode_frame,
        text="按名单解密",
        variable=app.phone_mode_var,
        value="partial",
        command=app.update_phone_mode_ui,
    ).pack(side="left", padx=(16, 0))

    ttk.Label(options_frame, text="名单文件", width=16).grid(row=3, column=0, sticky="w", pady=6, padx=(0, 10))
    app.phone_filter_entry = app.create_text_entry(options_frame, textvariable=app.phone_filter_file_var)
    app.phone_filter_entry.grid(row=3, column=1, columnspan=2, sticky="ew", pady=6)
    app.phone_filter_button = ttk.Button(
        options_frame,
        text="选择文件",
        command=app.choose_phone_filter_file,
        width=12,
    )
    app.phone_filter_button.grid(row=3, column=3, sticky="w", pady=6)

    action_frame = ttk.LabelFrame(phone_frame, text="3. 执行", padding=14)
    action_frame.grid(row=2, column=0, sticky="ew", pady=(14, 0))
    app.phone_run_button = ttk.Button(
        action_frame,
        text="开始解密",
        command=app.start_phone_decrypt_run,
        style="Accent.TButton",
        width=12,
    )
    app.phone_run_button.pack(side="left")
    ttk.Label(action_frame, textvariable=app.phone_status_var).pack(side="right")

    result_frame = ttk.LabelFrame(phone_frame, text="4. 解密结果", padding=14)
    result_frame.grid(row=3, column=0, sticky="nsew", pady=(14, 0))
    result_frame.columnconfigure(0, weight=1)
    result_frame.rowconfigure(0, weight=1)
    app.phone_result_text = tk.Text(
        result_frame,
        wrap="word",
        height=10,
        relief="solid",
        bd=1,
        bg="#FCFCFC",
        fg="#222222",
        insertbackground="#1677FF",
        padx=10,
        pady=10,
    )
    app.phone_result_text.grid(row=0, column=0, sticky="nsew")
    app.phone_result_text.configure(state="disabled")
    result_scroll = ttk.Scrollbar(result_frame, orient="vertical", command=app.phone_result_text.yview)
    result_scroll.grid(row=0, column=1, sticky="ns")
    app.phone_result_text.configure(yscrollcommand=result_scroll.set)
