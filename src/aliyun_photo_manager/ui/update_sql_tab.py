from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_update_sql_tab(app, notebook: ttk.Notebook) -> None:
    update_sql_frame = ttk.Frame(notebook, padding=14)
    update_sql_frame.columnconfigure(0, weight=1)
    update_sql_frame.rowconfigure(2, weight=1)
    notebook.add(update_sql_frame, text="更新SQL生成")

    update_sql_form = ttk.LabelFrame(update_sql_frame, text="字段映射更新", padding=12)
    update_sql_form.grid(row=0, column=0, sticky="ew")
    update_sql_form.columnconfigure(1, weight=1)
    app.add_file_row(
        update_sql_form,
        row=0,
        label="映射模板",
        variable=app.update_sql_mapping_var,
        command=app.choose_update_sql_mapping,
        button_text="选择文件",
    )

    update_sql_template_action = ttk.Frame(update_sql_form)
    update_sql_template_action.grid(row=1, column=1, sticky="w", pady=(2, 4))
    ttk.Button(update_sql_template_action, text="加载字段", command=app.load_update_sql_headers).pack(
        side="left"
    )
    ttk.Button(
        update_sql_template_action,
        text="导出模板",
        command=app.export_update_sql_template_file,
    ).pack(side="left", padx=(8, 0))

    app.add_entry_row(update_sql_form, 2, "考生表名称", app.update_sql_target_table_var)
    app.add_entry_row(update_sql_form, 3, "临时表名称", app.update_sql_source_table_var)

    ttk.Label(update_sql_form, text="考生表关联字段", width=16).grid(row=4, column=0, sticky="w", pady=4)
    app.update_sql_target_key_combo = ttk.Combobox(
        update_sql_form,
        textvariable=app.update_sql_target_key_var,
        values=app.update_sql_target_headers,
        state="readonly",
    )
    app.update_sql_target_key_combo.grid(row=4, column=1, sticky="ew", pady=4)

    ttk.Label(update_sql_form, text="临时表关联字段", width=16).grid(row=5, column=0, sticky="w", pady=4)
    app.update_sql_source_key_combo = ttk.Combobox(
        update_sql_form,
        textvariable=app.update_sql_source_key_var,
        values=app.update_sql_source_headers,
        state="readonly",
    )
    app.update_sql_source_key_combo.grid(row=5, column=1, sticky="ew", pady=4)

    app.add_tick_checkbutton(
        update_sql_form,
        text="忽略空值，不覆盖正式表",
        variable=app.update_sql_ignore_empty_var,
        row=6,
        column=1,
        sticky="w",
        padx=(0, 0),
    )

    update_sql_action = ttk.Frame(update_sql_frame)
    update_sql_action.grid(row=1, column=0, sticky="ew", pady=(12, 0))
    app.update_sql_run_button = ttk.Button(
        update_sql_action,
        text="生成 SQL",
        command=app.start_update_sql_render,
        style="Accent.TButton",
        width=12,
    )
    app.update_sql_run_button.pack(side="left")
    app.update_sql_copy_button = ttk.Button(
        update_sql_action,
        text="复制 SQL",
        command=app.copy_update_sql,
        state="disabled",
        width=12,
    )
    app.update_sql_copy_button.pack(side="left", padx=(10, 0))
    ttk.Label(update_sql_action, textvariable=app.update_sql_status_var).pack(side="right")

    update_sql_result_frame = ttk.LabelFrame(update_sql_frame, text="生成结果", padding=12)
    update_sql_result_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
    update_sql_result_frame.columnconfigure(0, weight=1)
    update_sql_result_frame.rowconfigure(0, weight=1)
    app.update_sql_result_text = tk.Text(
        update_sql_result_frame,
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
    app.update_sql_result_text.grid(row=0, column=0, sticky="nsew")
    app.update_sql_result_text.configure(state="disabled")
    update_sql_result_scroll = ttk.Scrollbar(
        update_sql_result_frame,
        orient="vertical",
        command=app.update_sql_result_text.yview,
    )
    update_sql_result_scroll.grid(row=0, column=1, sticky="ns")
    app.update_sql_result_text.configure(yscrollcommand=update_sql_result_scroll.set)
