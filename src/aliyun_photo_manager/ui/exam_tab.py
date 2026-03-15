from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_exam_tab(app, notebook: ttk.Notebook) -> None:
    exam_frame = ttk.Frame(notebook, padding=14)
    exam_frame.columnconfigure(0, weight=1)
    exam_frame.rowconfigure(3, weight=1)
    notebook.add(exam_frame, text="考场编排")

    exam_files_frame = ttk.LabelFrame(exam_frame, text="输入文件", padding=12)
    exam_files_frame.grid(row=0, column=0, sticky="ew")
    exam_files_frame.columnconfigure(1, weight=1)
    app.add_file_row(
        exam_files_frame,
        row=0,
        label="考生明细表",
        variable=app.exam_candidate_var,
        command=app.choose_exam_candidate_file,
        button_text="选择文件",
    )
    app.add_file_row(
        exam_files_frame,
        row=1,
        label="岗位归组表",
        variable=app.exam_group_var,
        command=app.choose_exam_group_file,
        button_text="选择文件",
    )
    app.add_file_row(
        exam_files_frame,
        row=2,
        label="编排片段表",
        variable=app.exam_plan_var,
        command=app.choose_exam_plan_file,
        button_text="选择文件",
    )
    app.add_entry_row(exam_files_frame, 3, "输出文件", app.exam_output_var)
    ttk.Button(exam_files_frame, text="自动生成", command=app.fill_exam_output_path).grid(
        row=3, column=2, padx=(8, 0), pady=4
    )

    exam_rule_frame = ttk.LabelFrame(exam_frame, text="考号规则", padding=12)
    exam_rule_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
    exam_rule_frame.columnconfigure(0, weight=1)
    exam_digits_row = ttk.Frame(exam_rule_frame)
    exam_digits_row.grid(row=0, column=0, sticky="ew")
    for idx, (label, variable) in enumerate(
        [
            ("考点位数", app.exam_point_digits_var),
            ("考场位数", app.exam_room_digits_var),
            ("座号位数", app.exam_seat_digits_var),
            ("流水号位数", app.exam_serial_digits_var),
        ]
    ):
        ttk.Label(exam_digits_row, text=label).grid(row=0, column=idx * 2, sticky="w")
        entry = app.create_text_entry(exam_digits_row, variable)
        entry.configure(width=6)
        entry.grid(row=0, column=idx * 2 + 1, sticky="w", padx=(6, 12))

    exam_sort_row = ttk.Frame(exam_rule_frame)
    exam_sort_row.grid(row=1, column=0, sticky="ew", pady=(10, 0))
    ttk.Label(exam_sort_row, text="组内顺序").pack(side="left")
    ttk.Radiobutton(exam_sort_row, text="按原顺序", variable=app.exam_sort_mode_var, value="original").pack(
        side="left", padx=(10, 0)
    )
    ttk.Radiobutton(exam_sort_row, text="随机打乱", variable=app.exam_sort_mode_var, value="random").pack(
        side="left", padx=(10, 0)
    )

    exam_rule_editor = ttk.Frame(exam_rule_frame)
    exam_rule_editor.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    ttk.Label(exam_rule_editor, text="项目").grid(row=0, column=0, sticky="w")
    app.exam_rule_type_combo = ttk.Combobox(
        exam_rule_editor,
        textvariable=app.exam_rule_type_var,
        values=app.exam_rule_base_types,
        state="readonly",
    )
    app.exam_rule_type_combo.grid(row=0, column=1, sticky="w", padx=(6, 12))
    ttk.Label(exam_rule_editor, text="自定义内容").grid(row=0, column=2, sticky="w")
    app.exam_rule_custom_entry = app.create_text_entry(exam_rule_editor, app.exam_rule_custom_var)
    app.exam_rule_custom_entry.grid(row=0, column=3, sticky="ew", padx=(6, 12))
    exam_rule_editor.columnconfigure(3, weight=1)
    ttk.Button(exam_rule_editor, text="添加规则", command=app.add_exam_rule_item).grid(
        row=0, column=4, sticky="e"
    )

    app.exam_rule_tree = ttk.Treeview(
        exam_rule_frame,
        columns=("index", "type", "custom"),
        show="headings",
        height=5,
    )
    app.exam_rule_tree.heading("index", text="顺序")
    app.exam_rule_tree.heading("type", text="项目")
    app.exam_rule_tree.heading("custom", text="自定义内容")
    app.exam_rule_tree.column("index", width=60, anchor="center")
    app.exam_rule_tree.column("type", width=120, anchor="w")
    app.exam_rule_tree.column("custom", width=220, anchor="w")
    app.exam_rule_tree.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
    exam_rule_frame.rowconfigure(3, weight=1)
    exam_rule_actions = ttk.Frame(exam_rule_frame)
    exam_rule_actions.grid(row=4, column=0, sticky="w", pady=(8, 0))
    ttk.Button(exam_rule_actions, text="上移", command=app.move_exam_rule_up).pack(side="left")
    ttk.Button(exam_rule_actions, text="下移", command=app.move_exam_rule_down).pack(side="left", padx=(8, 0))
    ttk.Button(exam_rule_actions, text="删除所选", command=app.remove_exam_rule_item).pack(side="left", padx=(8, 0))

    exam_action = ttk.Frame(exam_frame)
    exam_action.grid(row=2, column=0, sticky="ew", pady=(12, 0))
    app.exam_template_button = ttk.Button(
        exam_action,
        text="导出标准模板",
        command=app.export_exam_templates_from_ui,
        width=14,
    )
    app.exam_template_button.pack(side="left")
    app.exam_run_button = ttk.Button(
        exam_action,
        text="开始编排",
        command=app.start_exam_arrange_run,
        style="Accent.TButton",
        width=12,
    )
    app.exam_run_button.pack(side="left", padx=(10, 0))
    app.exam_open_button = ttk.Button(
        exam_action,
        text="打开结果文件",
        command=app.open_exam_result_file,
        state="disabled",
        width=14,
    )
    app.exam_open_button.pack(side="left", padx=(10, 0))
    ttk.Label(exam_action, textvariable=app.exam_status_var).pack(side="right")

    exam_result_frame = ttk.LabelFrame(exam_frame, text="编排结果", padding=12)
    exam_result_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
    exam_result_frame.columnconfigure(0, weight=1)
    exam_result_frame.rowconfigure(0, weight=1)
    app.exam_result_text = tk.Text(
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
    app.exam_result_text.grid(row=0, column=0, sticky="nsew")
    app.exam_result_text.configure(state="disabled")
    exam_result_scroll = ttk.Scrollbar(exam_result_frame, orient="vertical", command=app.exam_result_text.yview)
    exam_result_scroll.grid(row=0, column=1, sticky="ns")
    app.exam_result_text.configure(yscrollcommand=exam_result_scroll.set)
