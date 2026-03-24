from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def build_help_tab(app, notebook: ttk.Notebook) -> None:
    help_frame = ttk.Frame(notebook, padding=14)
    notebook.add(help_frame, text="使用说明")
    help_frame.columnconfigure(0, weight=1)
    help_frame.rowconfigure(0, weight=1)

    app.help_text = tk.Text(
        help_frame,
        wrap="word",
        font=("Helvetica", 13),
        relief="solid",
        padx=14,
        pady=14,
    )
    app.help_text.grid(row=0, column=0, sticky="nsew")
    help_scrollbar = ttk.Scrollbar(help_frame, orient="vertical", command=app.help_text.yview)
    help_scrollbar.grid(row=0, column=1, sticky="ns")
    app.help_text.configure(yscrollcommand=help_scrollbar.set)
    app.set_help_content()
