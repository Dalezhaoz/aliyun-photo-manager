from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

from ..data_matcher import ColumnMapping, DataMatchOptions, DataMatchSummary, run_data_match
from ..data_matcher import list_headers as list_match_headers


def choose_match_target(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.match_target_var.get()).parent)
        if app.match_target_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
    )
    if selected:
        app.match_target_var.set(selected)
        fill_match_output_path(app)
        load_match_headers(app)


def choose_match_source(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.match_source_var.get()).parent)
        if app.match_source_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
    )
    if selected:
        app.match_source_var.set(selected)
        load_match_headers(app)


def fill_match_output_path(app) -> None:
    target_value = app.match_target_var.get().strip()
    if not target_value:
        return
    target_path = Path(target_value)
    app.match_output_var.set(str(target_path.with_name(f"{target_path.stem}_数据匹配结果.xlsx")))


def load_match_headers(app) -> None:
    target_value = app.match_target_var.get().strip()
    source_value = app.match_source_var.get().strip()
    app.match_target_headers = []
    app.match_source_headers = []
    app.match_extra_mappings = []
    app.match_transfer_mappings = []
    app.match_target_key_combo.configure(values=[])
    app.match_source_key_combo.configure(values=[])
    app.match_extra_target_combo.configure(values=[])
    app.match_extra_source_combo.configure(values=[])
    app.match_transfer_source_combo.configure(values=[])
    for item_id in app.match_extra_tree.get_children():
        app.match_extra_tree.delete(item_id)
    for item_id in app.match_transfer_tree.get_children():
        app.match_transfer_tree.delete(item_id)
    if not target_value or not source_value:
        return
    target_path = Path(target_value)
    source_path = Path(source_value)
    if not target_path.exists() or not source_path.exists():
        return
    try:
        app.match_target_headers = list_match_headers(target_path)
        app.match_source_headers = list_match_headers(source_path)
    except Exception as exc:
        messagebox.showerror("加载失败", str(exc))
        return
    app.match_target_key_combo.configure(values=app.match_target_headers)
    app.match_source_key_combo.configure(values=app.match_source_headers)
    app.match_extra_target_combo.configure(values=app.match_target_headers)
    app.match_extra_source_combo.configure(values=app.match_source_headers)
    app.match_transfer_source_combo.configure(values=app.match_source_headers)
    if not app.match_target_key_var.get().strip() and app.match_target_headers:
        app.match_target_key_var.set(app.match_target_headers[0])
    if not app.match_source_key_var.get().strip() and app.match_source_headers:
        app.match_source_key_var.set(app.match_source_headers[0])
    if app.match_source_headers and not app.match_transfer_source_var.get().strip():
        app.match_transfer_source_var.set(app.match_source_headers[0])
        app.match_transfer_target_var.set(app.match_source_headers[0])


def add_extra_match_mapping(app) -> None:
    target_column = app.match_extra_target_var.get().strip()
    source_column = app.match_extra_source_var.get().strip()
    if not target_column or not source_column:
        messagebox.showerror("参数错误", "请选择目标表列和来源表列。")
        return
    mapping = ColumnMapping(target_column=target_column, source_column=source_column)
    if mapping in app.match_extra_mappings:
        return
    app.match_extra_mappings.append(mapping)
    app.match_extra_tree.insert("", tk.END, values=(target_column, source_column))


def remove_extra_match_mapping(app) -> None:
    selected = app.match_extra_tree.selection()
    if not selected:
        return
    for item_id in selected:
        values = app.match_extra_tree.item(item_id, "values")
        app.match_extra_tree.delete(item_id)
        app.match_extra_mappings = [
            item
            for item in app.match_extra_mappings
            if (item.target_column, item.source_column) != tuple(values)
        ]


def add_transfer_mapping(app) -> None:
    target_column = app.match_transfer_target_var.get().strip()
    source_column = app.match_transfer_source_var.get().strip()
    if not source_column:
        messagebox.showerror("参数错误", "请选择来源表补充列。")
        return
    if not target_column:
        target_column = source_column
        app.match_transfer_target_var.set(target_column)
    mapping = ColumnMapping(target_column=target_column, source_column=source_column)
    if mapping in app.match_transfer_mappings:
        return
    app.match_transfer_mappings.append(mapping)
    app.match_transfer_tree.insert("", tk.END, values=(target_column, source_column))


def remove_transfer_mapping(app) -> None:
    selected = app.match_transfer_tree.selection()
    if not selected:
        return
    for item_id in selected:
        values = app.match_transfer_tree.item(item_id, "values")
        app.match_transfer_tree.delete(item_id)
        app.match_transfer_mappings = [
            item
            for item in app.match_transfer_mappings
            if (item.target_column, item.source_column) != tuple(values)
        ]


def update_match_summary_ui(app, summary: Optional[DataMatchSummary]) -> None:
    app.last_match_summary = summary
    if summary is None:
        app.match_result_var.set("数据匹配结果会显示在这里")
        app.match_open_button.configure(state="disabled")
        set_match_result_text(app, app.match_result_var.get())
        return
    app.match_result_var.set(
        "匹配完成：\n"
        f"目标表：{summary.target_path}\n"
        f"来源表：{summary.source_path}\n"
        f"结果文件：{summary.output_path}\n"
        f"目标表表头：第 {summary.target_header_row} 行\n"
        f"来源表表头：第 {summary.source_header_row} 行\n"
        f"总行数：{summary.total_rows}\n"
        f"匹配成功：{summary.matched_rows}\n"
        f"未匹配：{summary.unmatched_rows}\n"
        f"来源重复键：{summary.duplicate_source_keys}\n"
        f"重复未写入：{summary.ambiguous_rows}"
    )
    set_match_result_text(app, app.match_result_var.get())
    app.match_open_button.configure(state="normal")


def set_match_result_text(app, content: str) -> None:
    if app.match_result_text is None:
        return
    app.match_result_text.configure(state="normal")
    app.match_result_text.delete("1.0", tk.END)
    app.match_result_text.insert("1.0", content)
    app.match_result_text.configure(state="disabled")


def open_match_result_file(app) -> None:
    if app.last_match_summary is None:
        return
    app.open_local_file(app.last_match_summary.output_path, "未找到匹配结果文件")


def start_match_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    target_value = app.match_target_var.get().strip()
    source_value = app.match_source_var.get().strip()
    if not target_value or not source_value:
        messagebox.showerror("参数错误", "请选择目标表和来源表。")
        return
    target_key = app.match_target_key_var.get().strip()
    source_key = app.match_source_key_var.get().strip()
    if not target_key or not source_key:
        messagebox.showerror("参数错误", "请选择目标表匹配列和来源表匹配列。")
        return
    transfer_mappings = list(app.match_transfer_mappings)
    if not transfer_mappings:
        messagebox.showerror("参数错误", "请至少选择一个来源表补充列。")
        return
    extra_mappings = [
        mapping
        for mapping in app.match_extra_mappings
        if not (mapping.target_column == target_key and mapping.source_column == source_key)
    ]
    target_path = Path(target_value)
    source_path = Path(source_value)
    if target_path.suffix.lower() not in {".xlsx", ".xls"} or source_path.suffix.lower() not in {
        ".xlsx",
        ".xls",
    }:
        messagebox.showerror("参数错误", "数据匹配仅支持 `.xlsx` 或 `.xls` 文件。")
        return
    output_value = app.match_output_var.get().strip()
    output_path = Path(output_value) if output_value else target_path.with_name(
        f"{target_path.stem}_数据匹配结果.xlsx"
    )
    app.match_output_var.set(str(output_path))
    app.save_settings()

    app.match_run_button.configure(state="disabled")
    app.match_open_button.configure(state="disabled")
    app.match_status_var.set("匹配中")
    app.match_result_var.set("正在执行数据匹配，请稍候...")
    app.write_log(f"启动数据匹配任务：{target_path.name} <- {source_path.name}")

    def runner() -> None:
        try:
            summary = run_data_match(
                DataMatchOptions(
                    target_path=target_path,
                    source_path=source_path,
                    target_key_column=target_key,
                    source_key_column=source_key,
                    extra_match_mappings=extra_mappings,
                    transfer_mappings=transfer_mappings,
                    output_path=output_path,
                ),
                logger=app.make_logger(),
            )
        except Exception as exc:
            app.log_queue.put(f"__MATCH_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "match_summary", "summary": summary})
            app.log_queue.put("__MATCH_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
