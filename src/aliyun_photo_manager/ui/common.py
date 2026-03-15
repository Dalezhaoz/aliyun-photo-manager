from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk
from typing import Optional

FORM_LABEL_WIDTH = 16
PICKER_BUTTON_WIDTH = 12
TOOLBAR_BUTTON_WIDTH = 14
ROW_PADY = 6
ENTRY_FONT = ("Helvetica", 11)


def create_text_entry(parent, textvariable: tk.StringVar, show: str = ""):
    return tk.Entry(
        parent,
        textvariable=textvariable,
        show=show,
        relief="solid",
        bd=1,
        highlightthickness=1,
        bg="#FCFCFC",
        fg="#222222",
        disabledbackground="#F2F2F2",
        disabledforeground="#8A8A8A",
        readonlybackground="#F7F7F7",
        highlightbackground="#CFCFCF",
        highlightcolor="#4A90E2",
        insertbackground="#1677FF",
        insertwidth=2,
        insertborderwidth=0,
        cursor="xterm",
        takefocus=True,
        exportselection=False,
        font=ENTRY_FONT,
    )


def add_entry_row(
    app,
    parent: ttk.Frame,
    row: int,
    label: str,
    variable: tk.StringVar,
    show: Optional[str] = None,
):
    ttk.Label(parent, text=label, width=FORM_LABEL_WIDTH).grid(
        row=row, column=0, sticky="w", pady=ROW_PADY, padx=(0, 10)
    )
    entry = create_text_entry(parent, textvariable=variable, show=show or "")
    entry.grid(row=row, column=1, sticky="ew", pady=ROW_PADY)
    return entry


def add_tick_checkbutton(
    parent,
    text: str,
    variable: tk.BooleanVar,
    command=None,
    row: int = 0,
    column: int = 0,
    padx=(0, 18),
    pady=4,
    sticky: str = "w",
):
    background = "#f5f5f5"
    try:
        background = parent.cget("background")
    except tk.TclError:
        try:
            style = ttk.Style()
            background = style.lookup("TFrame", "background") or background
        except tk.TclError:
            pass
    widget = tk.Checkbutton(
        parent,
        text=text,
        variable=variable,
        onvalue=True,
        offvalue=False,
        command=command,
        anchor="w",
        highlightthickness=0,
        relief="flat",
        borderwidth=0,
        background=background,
        activebackground=background,
        selectcolor="#ffffff",
    )
    widget.grid(row=row, column=column, sticky=sticky, padx=padx, pady=(6, 6))
    return widget


def add_path_row(
    app,
    parent: ttk.Frame,
    row: int,
    label: str,
    variable: tk.StringVar,
) -> None:
    ttk.Label(parent, text=label, width=FORM_LABEL_WIDTH).grid(
        row=row, column=0, sticky="w", pady=ROW_PADY, padx=(0, 10)
    )
    entry = create_text_entry(parent, textvariable=variable)
    entry.grid(row=row, column=1, sticky="ew", pady=ROW_PADY)
    ttk.Button(
        parent,
        text="选择",
        width=PICKER_BUTTON_WIDTH,
        command=lambda: app.choose_directory(variable),
    ).grid(row=row, column=2, padx=(10, 0), pady=ROW_PADY)


def add_file_row(
    app,
    parent: ttk.Frame,
    row: int,
    label: str,
    variable: tk.StringVar,
    command,
    button_text: str = "选择",
) -> None:
    ttk.Label(parent, text=label, width=FORM_LABEL_WIDTH).grid(
        row=row, column=0, sticky="w", pady=ROW_PADY, padx=(0, 10)
    )
    entry = create_text_entry(parent, textvariable=variable)
    entry.grid(row=row, column=1, sticky="ew", pady=ROW_PADY)
    ttk.Button(parent, text=button_text, width=PICKER_BUTTON_WIDTH, command=command).grid(
        row=row, column=2, padx=(10, 0), pady=ROW_PADY
    )


def bind_mousewheel_to_canvas(canvas: tk.Canvas, target) -> None:
    def on_mousewheel(event) -> None:
        if event.delta:
            delta = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        else:
            delta = 0
        if delta != 0:
            canvas.yview_scroll(delta, "units")

    def bind_events(_event=None) -> None:
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        canvas.bind_all("<Button-4>", on_mousewheel)
        canvas.bind_all("<Button-5>", on_mousewheel)

    def unbind_events(_event=None) -> None:
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")

    target.bind("<Enter>", bind_events)
    target.bind("<Leave>", unbind_events)
