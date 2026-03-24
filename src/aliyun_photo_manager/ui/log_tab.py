from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_log_tab(app, notebook: ttk.Notebook) -> None:
    log_frame = ttk.Frame(notebook, padding=14)
    notebook.add(log_frame, text="运行日志")
    log_frame.columnconfigure(0, weight=1)
    log_frame.rowconfigure(1, weight=1)

    log_toolbar = ttk.Frame(log_frame)
    log_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    ttk.Button(log_toolbar, text="清空日志", command=app.clear_log).pack(side="left")

    app.log_text = tk.Text(
        log_frame,
        wrap="word",
        bg="#101826",
        fg="#E5EEF7",
        insertbackground="#E5EEF7",
        font=("Menlo", 12),
        relief="flat",
        padx=12,
        pady=12,
    )
    app.log_text.grid(row=1, column=0, sticky="nsew")
    scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=app.log_text.yview)
    scrollbar.grid(row=1, column=1, sticky="ns")
    app.log_text.configure(yscrollcommand=scrollbar.set)
