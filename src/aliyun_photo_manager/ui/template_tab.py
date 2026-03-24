from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_template_tab(app, notebook: ttk.Notebook) -> None:
    word_frame = ttk.Frame(notebook, padding=14)
    word_frame.columnconfigure(0, weight=1)
    word_frame.rowconfigure(2, weight=1)
    notebook.add(word_frame, text="表样转换")

    intro_frame = tk.Frame(
        word_frame,
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
        text="表样转换",
        bg="#EAF2FF",
        fg="#162033",
        font=("Microsoft YaHei UI", 18, "bold"),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        intro_frame,
        text="把 Word / Excel 表样模板转换成 HTML，支持 Net 版和 Java 版占位符，并可直接查看代码与预览效果。",
        bg="#EAF2FF",
        fg="#5A6D83",
        font=("Microsoft YaHei UI", 10),
        anchor="w",
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(8, 0))

    word_form = ttk.LabelFrame(word_frame, text="1. 模板文件", padding=14)
    word_form.grid(row=1, column=0, sticky="ew")
    word_form.columnconfigure(1, weight=1)
    app.add_file_row(
        word_form,
        row=0,
        label="表样文件",
        variable=app.word_source_var,
        command=app.choose_word_source,
        button_text="选择文件",
    )

    word_action = ttk.LabelFrame(word_frame, text="2. 导出操作", padding=14)
    word_action.grid(row=2, column=0, sticky="ew", pady=(14, 0))
    app.word_net_button = ttk.Button(
        word_action,
        text="Net版导出",
        command=lambda: app.start_word_export("net"),
        style="Accent.TButton",
        width=12,
    )
    app.word_net_button.pack(side="left")
    app.word_java_button = ttk.Button(
        word_action,
        text="Java版导出",
        command=lambda: app.start_word_export("java"),
        style="Accent.TButton",
        width=12,
    )
    app.word_java_button.pack(side="left", padx=(10, 0))
    app.word_copy_button = ttk.Button(
        word_action,
        text="复制 HTML",
        command=app.copy_word_html,
        state="disabled",
        width=12,
    )
    app.word_copy_button.pack(side="left", padx=(10, 0))
    app.word_open_browser_button = ttk.Button(
        word_action,
        text="浏览器预览",
        command=app.open_word_preview_in_browser,
        state="disabled",
        width=12,
    )
    app.word_open_browser_button.pack(side="left", padx=(10, 0))
    ttk.Label(word_action, textvariable=app.word_status_var).pack(side="right")

    word_result_frame = ttk.LabelFrame(word_frame, text="3. 导出结果", padding=14)
    word_result_frame.grid(row=3, column=0, sticky="nsew", pady=(14, 0))
    word_result_frame.columnconfigure(0, weight=1)
    word_result_frame.rowconfigure(1, weight=1)
    ttk.Label(
        word_result_frame,
        textvariable=app.word_result_var,
        justify="left",
        wraplength=860,
    ).grid(row=0, column=0, sticky="w")

    word_view_notebook = ttk.Notebook(word_result_frame)
    word_view_notebook.grid(row=1, column=0, sticky="nsew", pady=(12, 0))

    word_code_frame = ttk.Frame(word_view_notebook, padding=8)
    word_code_frame.columnconfigure(0, weight=1)
    word_code_frame.rowconfigure(0, weight=1)
    word_view_notebook.add(word_code_frame, text="代码")

    app.word_code_text = tk.Text(
        word_code_frame,
        wrap="none",
        font=("Menlo", 11),
        relief="solid",
        height=18,
    )
    app.word_code_text.grid(row=0, column=0, sticky="nsew")
    app.word_code_text.configure(state="disabled")
    word_code_scroll_y = ttk.Scrollbar(word_code_frame, orient="vertical", command=app.word_code_text.yview)
    word_code_scroll_y.grid(row=0, column=1, sticky="ns")
    app.word_code_text.configure(yscrollcommand=word_code_scroll_y.set)

    word_preview_frame = ttk.Frame(word_view_notebook, padding=8)
    word_preview_frame.columnconfigure(0, weight=1)
    word_preview_frame.rowconfigure(1, weight=1)
    word_view_notebook.add(word_preview_frame, text="预览")
    ttk.Label(word_preview_frame, textvariable=app.word_preview_status_var).grid(
        row=0, column=0, sticky="w", pady=(0, 8)
    )
    app.word_preview_container = ttk.Frame(word_preview_frame)
    app.word_preview_container.grid(row=1, column=0, sticky="nsew")
    app.word_preview_container.columnconfigure(0, weight=1)
    app.word_preview_container.rowconfigure(0, weight=1)
