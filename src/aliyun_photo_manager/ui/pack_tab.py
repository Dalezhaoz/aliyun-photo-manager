from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_pack_tab(app, notebook: ttk.Notebook) -> None:
    pack_frame = ttk.Frame(notebook, padding=14)
    pack_frame.columnconfigure(0, weight=1)
    pack_frame.rowconfigure(2, weight=1)
    notebook.add(pack_frame, text="结果打包")

    pack_form = ttk.LabelFrame(pack_frame, text="压缩加密", padding=12)
    pack_form.grid(row=0, column=0, sticky="ew")
    pack_form.columnconfigure(1, weight=1)
    ttk.Label(pack_form, text="待打包对象", width=16).grid(row=0, column=0, sticky="w", pady=6)
    app.pack_source_entry = app.create_text_entry(pack_form, textvariable=app.pack_source_dir_var)
    app.pack_source_entry.grid(row=0, column=1, sticky="ew", pady=6)
    ttk.Button(pack_form, text="选择文件", command=app.choose_pack_source_file, width=12).grid(
        row=0, column=2, padx=(10, 0), pady=6
    )
    ttk.Button(pack_form, text="选择文件夹", command=app.choose_pack_source_directory, width=12).grid(
        row=0, column=3, padx=(10, 0), pady=6
    )
    app.add_path_row(pack_form, row=1, label="输出目录", variable=app.pack_output_dir_var)
    app.add_tick_checkbutton(
        pack_form,
        text="手动设置密码",
        variable=app.pack_use_custom_password_var,
        command=app.update_pack_password_mode_ui,
    ).grid(row=2, column=1, sticky="w", pady=(4, 0))
    app.pack_password_entry = app.add_entry_row(pack_form, 3, "打包密码", app.pack_password_var)

    pack_action = ttk.Frame(pack_frame)
    pack_action.grid(row=1, column=0, sticky="ew", pady=(12, 0))
    app.pack_run_button = ttk.Button(
        pack_action,
        text="一键打包并加密",
        command=app.start_pack_run,
        style="Accent.TButton",
        width=16,
    )
    app.pack_run_button.pack(side="left")
    app.pack_copy_password_button = ttk.Button(
        pack_action,
        text="复制密码",
        command=app.copy_pack_password,
        state="disabled",
        width=12,
    )
    app.pack_copy_password_button.pack(side="left", padx=(10, 0))
    app.pack_open_button = ttk.Button(
        pack_action,
        text="打开压缩包",
        command=app.open_pack_file,
        state="disabled",
        width=12,
    )
    app.pack_open_button.pack(side="left", padx=(10, 0))
    ttk.Label(pack_action, textvariable=app.pack_status_var).pack(side="right")

    pack_result_frame = ttk.LabelFrame(pack_frame, text="打包结果", padding=12)
    pack_result_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
    pack_result_frame.columnconfigure(0, weight=1)
    ttk.Label(
        pack_result_frame,
        textvariable=app.pack_result_var,
        justify="left",
        wraplength=860,
    ).grid(row=0, column=0, sticky="w")

    pack_query_frame = ttk.LabelFrame(pack_frame, text="密码查询", padding=12)
    pack_query_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))
    pack_query_frame.columnconfigure(0, weight=1)
    pack_query_frame.rowconfigure(1, weight=1)
    app.pack_query_entry = tk.Entry(
        pack_query_frame,
        textvariable=app.pack_query_var,
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
    app.pack_query_entry.grid(row=0, column=0, sticky="ew")
    ttk.Button(pack_query_frame, text="查询密码", command=app.run_pack_history_query).grid(
        row=0, column=1, padx=(10, 0)
    )
    app.pack_query_result_text = tk.Text(
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
    app.pack_query_result_text.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
    app.pack_query_result_text.configure(state="disabled")
    pack_query_scroll = ttk.Scrollbar(
        pack_query_frame,
        orient="vertical",
        command=app.pack_query_result_text.yview,
    )
    pack_query_scroll.grid(row=1, column=2, sticky="ns", pady=(8, 0))
    app.pack_query_result_text.configure(yscrollcommand=pack_query_scroll.set)
