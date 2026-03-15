from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_match_tab(app, notebook: ttk.Notebook) -> None:
    match_frame = ttk.Frame(notebook, padding=14)
    match_frame.columnconfigure(0, weight=1)
    match_frame.rowconfigure(3, weight=1)
    notebook.add(match_frame, text="数据匹配")

    match_form = ttk.LabelFrame(match_frame, text="表格匹配补列", padding=12)
    match_form.grid(row=0, column=0, sticky="ew")
    match_form.columnconfigure(1, weight=1)
    app.add_file_row(
        match_form,
        row=0,
        label="目标表",
        variable=app.match_target_var,
        command=app.choose_match_target,
        button_text="选择文件",
    )
    app.add_file_row(
        match_form,
        row=1,
        label="来源表",
        variable=app.match_source_var,
        command=app.choose_match_source,
        button_text="选择文件",
    )
    app.add_entry_row(match_form, 2, "输出文件", app.match_output_var)
    ttk.Button(match_form, text="自动生成", command=app.fill_match_output_path).grid(
        row=2, column=2, padx=(8, 0), pady=4
    )
    ttk.Label(match_form, text="目标表匹配列", width=16).grid(row=3, column=0, sticky="w", pady=4)
    app.match_target_key_combo = ttk.Combobox(
        match_form,
        textvariable=app.match_target_key_var,
        values=app.match_target_headers,
        state="readonly",
    )
    app.match_target_key_combo.grid(row=3, column=1, sticky="ew", pady=4)
    ttk.Label(match_form, text="来源表匹配列", width=16).grid(row=4, column=0, sticky="w", pady=4)
    app.match_source_key_combo = ttk.Combobox(
        match_form,
        textvariable=app.match_source_key_var,
        values=app.match_source_headers,
        state="readonly",
    )
    app.match_source_key_combo.grid(row=4, column=1, sticky="ew", pady=4)
    ttk.Button(match_form, text="加载表头", command=app.load_match_headers).grid(
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
    app.match_extra_target_combo = ttk.Combobox(
        extra_frame,
        textvariable=app.match_extra_target_var,
        values=app.match_target_headers,
        state="readonly",
    )
    app.match_extra_target_combo.grid(row=1, column=0, sticky="ew", pady=(4, 0))
    app.match_extra_source_combo = ttk.Combobox(
        extra_frame,
        textvariable=app.match_extra_source_var,
        values=app.match_source_headers,
        state="readonly",
    )
    app.match_extra_source_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))
    ttk.Button(extra_frame, text="添加映射", command=app.add_extra_match_mapping).grid(
        row=1, column=2, padx=(8, 0), pady=(4, 0)
    )
    app.match_extra_tree = ttk.Treeview(
        extra_frame,
        columns=("target", "source"),
        show="headings",
        height=5,
    )
    app.match_extra_tree.heading("target", text="目标表列")
    app.match_extra_tree.heading("source", text="来源表列")
    app.match_extra_tree.column("target", width=160, anchor="w")
    app.match_extra_tree.column("source", width=160, anchor="w")
    app.match_extra_tree.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
    extra_frame.rowconfigure(2, weight=1)
    ttk.Button(extra_frame, text="删除所选", command=app.remove_extra_match_mapping).grid(
        row=2, column=2, padx=(8, 0), sticky="n"
    )

    transfer_frame = ttk.LabelFrame(match_lists, text="补充列映射", padding=12)
    transfer_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
    transfer_frame.columnconfigure(0, weight=1)
    ttk.Label(transfer_frame, text="结果列名").grid(row=0, column=0, sticky="w")
    ttk.Label(transfer_frame, text="来源表列").grid(row=0, column=1, sticky="w", padx=(8, 0))
    app.match_transfer_target_entry = app.create_text_entry(
        transfer_frame,
        textvariable=app.match_transfer_target_var,
    )
    app.match_transfer_target_entry.grid(row=1, column=0, sticky="ew", pady=(4, 0))
    app.match_transfer_source_combo = ttk.Combobox(
        transfer_frame,
        textvariable=app.match_transfer_source_var,
        values=app.match_source_headers,
        state="readonly",
    )
    app.match_transfer_source_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))
    ttk.Button(transfer_frame, text="添加补充列", command=app.add_transfer_mapping).grid(
        row=1, column=2, padx=(8, 0), pady=(4, 0)
    )
    app.match_transfer_tree = ttk.Treeview(
        transfer_frame,
        columns=("target", "source"),
        show="headings",
        height=5,
    )
    app.match_transfer_tree.heading("target", text="结果列名")
    app.match_transfer_tree.heading("source", text="来源表列")
    app.match_transfer_tree.column("target", width=160, anchor="w")
    app.match_transfer_tree.column("source", width=160, anchor="w")
    app.match_transfer_tree.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
    transfer_frame.rowconfigure(2, weight=1)
    ttk.Button(transfer_frame, text="删除所选", command=app.remove_transfer_mapping).grid(
        row=2, column=2, padx=(8, 0), sticky="n"
    )

    match_action = ttk.Frame(match_frame)
    match_action.grid(row=2, column=0, sticky="ew", pady=(12, 0))
    app.match_run_button = ttk.Button(
        match_action,
        text="开始匹配",
        command=app.start_match_run,
        style="Accent.TButton",
        width=12,
    )
    app.match_run_button.pack(side="left")
    app.match_open_button = ttk.Button(
        match_action,
        text="打开结果文件",
        command=app.open_match_result_file,
        state="disabled",
        width=14,
    )
    app.match_open_button.pack(side="left", padx=(10, 0))
    ttk.Label(match_action, textvariable=app.match_status_var).pack(side="right")

    match_result_frame = ttk.LabelFrame(match_frame, text="匹配结果", padding=12)
    match_result_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
    match_result_frame.columnconfigure(0, weight=1)
    match_result_frame.rowconfigure(0, weight=1)
    app.match_result_text = tk.Text(
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
    app.match_result_text.grid(row=0, column=0, sticky="nsew")
    app.match_result_text.configure(state="disabled")
    match_result_scroll = ttk.Scrollbar(
        match_result_frame,
        orient="vertical",
        command=app.match_result_text.yview,
    )
    match_result_scroll.grid(row=0, column=1, sticky="ns")
    app.match_result_text.configure(yscrollcommand=match_result_scroll.set)
