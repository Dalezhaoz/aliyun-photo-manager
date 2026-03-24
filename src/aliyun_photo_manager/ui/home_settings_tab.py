from __future__ import annotations

from tkinter import ttk


def build_home_settings_tab(app, notebook: ttk.Notebook) -> None:
    frame = ttk.Frame(notebook, padding=18)
    notebook.add(frame, text="首页设置")
    frame.columnconfigure(0, weight=1)

    config_frame = ttk.LabelFrame(frame, text="快捷入口设置", padding=14)
    config_frame.grid(row=0, column=0, sticky="ew")
    for column in range(3):
        config_frame.columnconfigure(column, weight=1)

    for index, shortcut_var in enumerate(app.home_shortcut_vars):
        row = index // 3
        column = index % 3
        cell = ttk.Frame(config_frame)
        cell.grid(row=row, column=column, sticky="ew", padx=(0 if column == 0 else 10, 0), pady=6)
        cell.columnconfigure(0, weight=1)
        ttk.Label(cell, text=f"入口 {index + 1}").grid(row=0, column=0, sticky="w", pady=(0, 6))
        combo = ttk.Combobox(
            cell,
            textvariable=shortcut_var,
            values=app.HOME_SHORTCUT_OPTIONS,
            state="readonly",
        )
        combo.grid(row=1, column=0, sticky="ew")

    ttk.Button(
        config_frame,
        text="保存快捷入口",
        command=app.save_home_shortcuts,
        width=14,
    ).grid(row=2, column=2, sticky="e", pady=(10, 0))
