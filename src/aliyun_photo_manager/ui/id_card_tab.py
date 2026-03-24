from __future__ import annotations

import tkinter as tk
from datetime import date
from tkinter import ttk

from ..id_card_tools import list_provinces


def build_id_card_tab(app, notebook: ttk.Notebook) -> None:
    frame = ttk.Frame(notebook, padding=14)
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(2, weight=1)
    notebook.add(frame, text="身份证工具")

    validate_frame = ttk.LabelFrame(frame, text="输入校验", padding=12)
    validate_frame.grid(row=0, column=0, sticky="ew")
    validate_frame.columnconfigure(1, weight=1)
    ttk.Label(validate_frame, text="身份证号", width=16).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 10))
    app.id_input_entry = app.create_text_entry(validate_frame, textvariable=app.id_input_var)
    app.id_input_entry.grid(row=0, column=1, sticky="ew", pady=6)
    ttk.Button(
        validate_frame,
        text="校验并解析",
        command=app.run_id_card_validate,
        style="Accent.TButton",
        width=12,
    ).grid(row=0, column=2, padx=(10, 0), pady=6)

    generate_frame = ttk.LabelFrame(frame, text="身份证生成", padding=12)
    generate_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
    generate_frame.columnconfigure(1, weight=1)
    generate_frame.columnconfigure(3, weight=1)

    ttk.Label(generate_frame, text="省 / 市 / 县", width=16).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 10))
    region_frame = ttk.Frame(generate_frame)
    region_frame.grid(row=0, column=1, columnspan=3, sticky="ew", pady=6)
    region_frame.columnconfigure(0, weight=1)
    region_frame.columnconfigure(1, weight=1)
    region_frame.columnconfigure(2, weight=1)
    app.id_province_combo = ttk.Combobox(
        region_frame,
        textvariable=app.id_province_var,
        values=list_provinces(),
        state="readonly",
    )
    app.id_province_combo.grid(row=0, column=0, sticky="ew")
    app.id_city_combo = ttk.Combobox(
        region_frame,
        textvariable=app.id_city_var,
        values=[],
        state="readonly",
    )
    app.id_city_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0))
    app.id_county_combo = ttk.Combobox(
        region_frame,
        textvariable=app.id_county_var,
        values=[],
        state="readonly",
    )
    app.id_county_combo.grid(row=0, column=2, sticky="ew", padx=(10, 0))
    app.id_province_combo.bind("<<ComboboxSelected>>", lambda _event: app.update_id_city_values())
    app.id_city_combo.bind("<<ComboboxSelected>>", lambda _event: app.update_id_county_values())
    app.id_county_combo.bind("<<ComboboxSelected>>", lambda _event: app.update_id_region_hint())

    ttk.Label(generate_frame, text="手工区划码", width=16).grid(row=1, column=0, sticky="w", pady=6, padx=(0, 10))
    app.id_custom_region_entry = app.create_text_entry(generate_frame, textvariable=app.id_custom_region_code_var)
    app.id_custom_region_entry.grid(row=1, column=1, sticky="ew", pady=6)
    ttk.Label(generate_frame, textvariable=app.id_region_hint_var).grid(
        row=1, column=2, columnspan=2, sticky="w", pady=6, padx=(12, 0)
    )
    app.id_custom_region_entry.bind("<FocusOut>", lambda _event: app.update_id_region_hint())

    ttk.Label(generate_frame, text="出生日期", width=16).grid(row=2, column=0, sticky="w", pady=6, padx=(0, 10))
    birth_frame = ttk.Frame(generate_frame)
    birth_frame.grid(row=2, column=1, columnspan=3, sticky="w", pady=6)
    current_year = date.today().year
    year_values = [str(year) for year in range(current_year, 1899, -1)]
    month_values = [f"{month:02d}" for month in range(1, 13)]
    day_values = [f"{day:02d}" for day in range(1, 32)]
    app.id_birth_year_combo = ttk.Combobox(
        birth_frame,
        textvariable=app.id_birth_year_var,
        values=year_values,
        state="readonly",
        width=8,
    )
    app.id_birth_year_combo.pack(side="left")
    ttk.Label(birth_frame, text="年").pack(side="left", padx=(6, 12))
    app.id_birth_month_combo = ttk.Combobox(
        birth_frame,
        textvariable=app.id_birth_month_var,
        values=month_values,
        state="readonly",
        width=4,
    )
    app.id_birth_month_combo.pack(side="left")
    ttk.Label(birth_frame, text="月").pack(side="left", padx=(6, 12))
    app.id_birth_day_combo = ttk.Combobox(
        birth_frame,
        textvariable=app.id_birth_day_var,
        values=day_values,
        state="readonly",
        width=4,
    )
    app.id_birth_day_combo.pack(side="left")
    ttk.Label(birth_frame, text="日").pack(side="left", padx=(6, 0))
    app.id_birth_year_combo.bind("<<ComboboxSelected>>", lambda _event: app.update_id_day_values())
    app.id_birth_month_combo.bind("<<ComboboxSelected>>", lambda _event: app.update_id_day_values())

    ttk.Label(generate_frame, text="性别", width=16).grid(row=3, column=0, sticky="w", pady=6, padx=(0, 10))
    gender_frame = ttk.Frame(generate_frame)
    gender_frame.grid(row=3, column=1, columnspan=3, sticky="w", pady=6)
    ttk.Radiobutton(gender_frame, text="男", variable=app.id_gender_var, value="男").pack(side="left")
    ttk.Radiobutton(gender_frame, text="女", variable=app.id_gender_var, value="女").pack(side="left", padx=(16, 0))

    button_frame = ttk.Frame(generate_frame)
    button_frame.grid(row=4, column=1, columnspan=3, sticky="w", pady=(8, 0))
    ttk.Button(
        button_frame,
        text="生成身份证",
        command=app.run_id_card_generate,
        style="Accent.TButton",
        width=12,
    ).pack(side="left")
    ttk.Button(
        button_frame,
        text="复制结果",
        command=app.copy_generated_id_card,
        width=12,
    ).pack(side="left", padx=(10, 0))

    ttk.Label(generate_frame, text="生成结果", width=16).grid(row=5, column=0, sticky="w", pady=6, padx=(0, 10))
    app.id_generated_entry = app.create_text_entry(generate_frame, textvariable=app.id_generated_var)
    app.id_generated_entry.grid(row=5, column=1, columnspan=3, sticky="ew", pady=6)

    result_frame = ttk.LabelFrame(frame, text="结果", padding=12)
    result_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
    result_frame.columnconfigure(0, weight=1)
    result_frame.rowconfigure(0, weight=1)
    app.id_result_text = tk.Text(
        result_frame,
        wrap="word",
        height=12,
        relief="solid",
        bd=1,
        bg="#FCFCFC",
        fg="#222222",
        insertbackground="#1677FF",
        padx=10,
        pady=10,
    )
    app.id_result_text.grid(row=0, column=0, sticky="nsew")
    app.id_result_text.configure(state="disabled")
    scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=app.id_result_text.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    app.id_result_text.configure(yscrollcommand=scrollbar.set)
