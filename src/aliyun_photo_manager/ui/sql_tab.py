from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_sql_tab(app, notebook: ttk.Notebook) -> None:
    sql_frame = ttk.Frame(notebook, padding=14)
    sql_frame.columnconfigure(0, weight=1)
    sql_frame.rowconfigure(2, weight=1)
    notebook.add(sql_frame, text="SQL 配置执行")

    intro_frame = tk.Frame(
        sql_frame,
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
        text="SQL 配置执行",
        bg="#EAF2FF",
        fg="#162033",
        font=("Microsoft YaHei UI", 18, "bold"),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        intro_frame,
        text="通过现成 SQL 模板填入考试代码、考试年月和报名审核时间，快速生成一键执行脚本。",
        bg="#EAF2FF",
        fg="#5A6D83",
        font=("Microsoft YaHei UI", 10),
        anchor="w",
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(8, 0))

    sql_form = ttk.LabelFrame(sql_frame, text="1. 模板参数", padding=14)
    sql_form.grid(row=1, column=0, sticky="ew")
    sql_form.columnconfigure(1, weight=1)
    app.add_file_row(
        sql_form,
        row=0,
        label="SQL 模板",
        variable=app.sql_template_path_var,
        command=app.choose_sql_template,
        button_text="选择文件",
    )
    app.add_entry_row(sql_form, 1, "考试代码", app.sql_exam_code_var)
    app.add_entry_row(sql_form, 2, "考试年月", app.sql_exam_date_var)
    app.add_entry_row(sql_form, 3, "报名开始时间", app.sql_signup_start_var)
    app.add_entry_row(sql_form, 4, "报名结束时间", app.sql_signup_end_var)
    app.add_entry_row(sql_form, 5, "审核开始时间", app.sql_audit_start_var)
    app.add_entry_row(sql_form, 6, "审核结束时间", app.sql_audit_end_var)

    sql_action = ttk.LabelFrame(sql_frame, text="2. 生成操作", padding=14)
    sql_action.grid(row=2, column=0, sticky="ew", pady=(14, 0))
    app.sql_generate_button = ttk.Button(
        sql_action,
        text="生成 SQL",
        command=app.start_sql_render,
        style="Accent.TButton",
        width=12,
    )
    app.sql_generate_button.pack(side="left")
    app.sql_copy_button = ttk.Button(
        sql_action,
        text="复制 SQL",
        command=app.copy_sql_text,
        state="disabled",
        width=12,
    )
    app.sql_copy_button.pack(side="left", padx=(10, 0))
    ttk.Label(sql_action, textvariable=app.sql_status_var).pack(side="right")

    sql_result_frame = ttk.LabelFrame(sql_frame, text="3. 生成结果", padding=14)
    sql_result_frame.grid(row=3, column=0, sticky="nsew", pady=(14, 0))
    sql_result_frame.columnconfigure(0, weight=1)
    sql_result_frame.rowconfigure(1, weight=1)
    ttk.Label(
        sql_result_frame,
        textvariable=app.sql_result_var,
        justify="left",
        wraplength=860,
    ).grid(row=0, column=0, sticky="w")

    app.sql_result_text = tk.Text(
        sql_result_frame,
        wrap="none",
        font=("Menlo", 11),
        relief="solid",
        height=18,
    )
    app.sql_result_text.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
    app.sql_result_text.configure(state="disabled")
    sql_scroll_y = ttk.Scrollbar(sql_result_frame, orient="vertical", command=app.sql_result_text.yview)
    sql_scroll_y.grid(row=1, column=1, sticky="ns", pady=(12, 0))
    app.sql_result_text.configure(yscrollcommand=sql_scroll_y.set)
