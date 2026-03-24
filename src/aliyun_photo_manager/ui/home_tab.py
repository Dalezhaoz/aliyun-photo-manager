from __future__ import annotations

import tkinter as tk
from tkinter import ttk

def build_home_tab(app, notebook: ttk.Notebook) -> None:
    home_frame = ttk.Frame(notebook, padding=18)
    notebook.add(home_frame, text="首页")
    home_frame.columnconfigure(0, weight=1)

    shortcut_frame = ttk.LabelFrame(home_frame, text="常用快捷入口", padding=14)
    shortcut_frame.grid(row=0, column=0, sticky="ew")
    for index, shortcut_var in enumerate(app.home_shortcut_vars):
        row = index // 3
        column = index % 3
        shortcut_frame.columnconfigure(column, weight=1)
        card = tk.Frame(
            shortcut_frame,
            bg="#FFFFFF",
            highlightthickness=1,
            highlightbackground="#DDE6F0",
            padx=14,
            pady=14,
        )
        card.grid(row=row, column=column, sticky="ew", padx=(0 if column == 0 else 10, 0), pady=8)
        card.grid_columnconfigure(0, weight=1)
        tk.Label(
            card,
            textvariable=shortcut_var,
            bg="#FFFFFF",
            fg="#162033",
            font=("Microsoft YaHei UI", 11, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            card,
            text="进入",
            command=lambda idx=index: app.open_home_shortcut(idx),
            width=10,
        ).grid(row=1, column=0, sticky="w", pady=(12, 0))
