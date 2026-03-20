from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_status_tab(app, notebook: ttk.Notebook) -> None:
    status_frame = ttk.Frame(notebook, padding=14)
    status_frame.columnconfigure(0, weight=1)
    status_frame.rowconfigure(2, weight=1)
    notebook.add(status_frame, text="项目阶段汇总")

    server_frame = ttk.LabelFrame(status_frame, text="服务器配置", padding=12)
    server_frame.grid(row=0, column=0, sticky="ew")
    server_frame.columnconfigure(0, weight=1)
    server_frame.columnconfigure(1, weight=1)

    list_frame = ttk.Frame(server_frame)
    list_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
    list_frame.columnconfigure(0, weight=1)
    list_frame.rowconfigure(0, weight=1)
    app.status_server_tree = ttk.Treeview(
        list_frame,
        columns=("name", "host", "enabled"),
        show="headings",
        height=6,
    )
    app.status_server_tree.heading("name", text="服务器名称")
    app.status_server_tree.heading("host", text="地址")
    app.status_server_tree.heading("enabled", text="启用")
    app.status_server_tree.column("name", width=140, anchor="w")
    app.status_server_tree.column("host", width=220, anchor="w")
    app.status_server_tree.column("enabled", width=60, anchor="center")
    app.status_server_tree.grid(row=0, column=0, sticky="nsew")
    app.status_server_tree.bind("<<TreeviewSelect>>", app.on_status_server_select)
    server_scroll = ttk.Scrollbar(
        list_frame, orient="vertical", command=app.status_server_tree.yview
    )
    server_scroll.grid(row=0, column=1, sticky="ns")
    app.status_server_tree.configure(yscrollcommand=server_scroll.set)

    form_frame = ttk.Frame(server_frame)
    form_frame.grid(row=0, column=1, sticky="nsew")
    form_frame.columnconfigure(1, weight=1)
    app.add_entry_row(form_frame, 0, "服务器名称", app.status_server_name_var)
    app.add_entry_row(form_frame, 1, "数据库地址", app.status_server_host_var)
    app.add_entry_row(form_frame, 2, "端口", app.status_server_port_var)
    app.add_entry_row(form_frame, 3, "用户名", app.status_server_user_var)
    app.add_entry_row(form_frame, 4, "密码", app.status_server_password_var, show="*")
    app.add_tick_checkbutton(
        form_frame,
        text="启用此服务器",
        variable=app.status_server_enabled_var,
        row=5,
        column=1,
        pady=4,
    )

    server_action_frame = ttk.Frame(server_frame)
    server_action_frame.grid(row=1, column=1, sticky="ew", pady=(12, 0))
    ttk.Button(server_action_frame, text="新增/更新", command=app.save_status_server).pack(side="left")
    ttk.Button(server_action_frame, text="清空表单", command=app.clear_status_server_form).pack(
        side="left", padx=(8, 0)
    )
    ttk.Button(server_action_frame, text="删除所选", command=app.delete_status_server).pack(
        side="left", padx=(8, 0)
    )
    ttk.Button(server_action_frame, text="测试连接", command=app.test_status_server).pack(
        side="left", padx=(8, 0)
    )

    filter_frame = ttk.LabelFrame(status_frame, text="查询条件", padding=12)
    filter_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
    filter_frame.columnconfigure(1, weight=1)
    filter_frame.columnconfigure(3, weight=1)
    ttk.Label(filter_frame, text="状态").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
    app.status_filter_combo = ttk.Combobox(
        filter_frame,
        textvariable=app.status_filter_var,
        state="readonly",
        values=["正在进行 + 即将开始", "全部", "只看正在进行", "只看即将开始"],
    )
    app.status_filter_combo.grid(row=0, column=1, sticky="ew", pady=4)
    ttk.Label(filter_frame, text="阶段关键字").grid(row=0, column=2, sticky="w", padx=(12, 10), pady=4)
    app.status_stage_keyword_entry = app.create_text_entry(
        filter_frame, textvariable=app.status_stage_keyword_var
    )
    app.status_stage_keyword_entry.grid(row=0, column=3, sticky="ew", pady=4)
    ttk.Label(filter_frame, text="项目关键字").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
    app.status_project_keyword_entry = app.create_text_entry(
        filter_frame, textvariable=app.status_project_keyword_var
    )
    app.status_project_keyword_entry.grid(row=1, column=1, sticky="ew", pady=4)

    query_action_frame = ttk.Frame(filter_frame)
    query_action_frame.grid(row=1, column=3, sticky="e", pady=4)
    app.status_run_button = ttk.Button(
        query_action_frame,
        text="开始查询",
        command=app.start_status_query,
        style="Accent.TButton",
        width=12,
    )
    app.status_run_button.pack(side="left")
    app.status_export_button = ttk.Button(
        query_action_frame,
        text="导出 Excel",
        command=app.export_status_result,
        state="disabled",
        width=12,
    )
    app.status_export_button.pack(side="left", padx=(8, 0))

    result_frame = ttk.LabelFrame(status_frame, text="查询结果", padding=12)
    result_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
    result_frame.columnconfigure(0, weight=1)
    result_frame.rowconfigure(1, weight=1)
    ttk.Label(result_frame, textvariable=app.status_query_status_var).grid(row=0, column=0, sticky="w")
    app.status_result_tree = ttk.Treeview(
        result_frame,
        columns=("server", "database", "project", "stage", "start", "end", "status"),
        show="headings",
        height=12,
    )
    app.status_result_tree.heading("server", text="服务器")
    app.status_result_tree.heading("database", text="数据库")
    app.status_result_tree.heading("project", text="项目名称")
    app.status_result_tree.heading("stage", text="阶段名称")
    app.status_result_tree.heading("start", text="开始时间")
    app.status_result_tree.heading("end", text="结束时间")
    app.status_result_tree.heading("status", text="当前状态")
    app.status_result_tree.column("server", width=120, anchor="w")
    app.status_result_tree.column("database", width=140, anchor="w")
    app.status_result_tree.column("project", width=180, anchor="w")
    app.status_result_tree.column("stage", width=180, anchor="w")
    app.status_result_tree.column("start", width=150, anchor="center")
    app.status_result_tree.column("end", width=150, anchor="center")
    app.status_result_tree.column("status", width=90, anchor="center")
    app.status_result_tree.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
    result_scroll = ttk.Scrollbar(
        result_frame, orient="vertical", command=app.status_result_tree.yview
    )
    result_scroll.grid(row=1, column=1, sticky="ns", pady=(8, 0))
    app.status_result_tree.configure(yscrollcommand=result_scroll.set)
