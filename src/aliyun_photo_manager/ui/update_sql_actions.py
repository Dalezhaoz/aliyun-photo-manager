from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from ..update_sql_generator import (
    UpdateSqlResult,
    export_update_sql_template,
    load_update_field_mappings,
    render_update_sql,
)


def choose_update_sql_mapping(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.update_sql_mapping_var.get()).parent)
        if app.update_sql_mapping_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
    )
    if selected:
        app.update_sql_mapping_var.set(selected)
        load_update_sql_headers(app)


def export_update_sql_template_file(app) -> None:
    selected = filedialog.asksaveasfilename(
        initialdir=str(Path(app.update_sql_mapping_var.get()).parent)
        if app.update_sql_mapping_var.get().strip()
        else str(Path.cwd()),
        defaultextension=".xlsx",
        initialfile="更新SQL字段映射模板.xlsx",
        filetypes=[("Excel 文件", "*.xlsx")],
    )
    if not selected:
        return
    try:
        summary = export_update_sql_template(Path(selected))
    except Exception as exc:
        messagebox.showerror("导出失败", str(exc))
        return
    app.update_sql_mapping_var.set(str(summary.output_path))
    load_update_sql_headers(app)
    app.update_sql_status_var.set("模板已导出")
    app.write_log(f"已导出更新SQL字段映射模板：{summary.output_path}")


def load_update_sql_headers(app) -> None:
    mapping_value = app.update_sql_mapping_var.get().strip()
    app.update_sql_target_headers = []
    app.update_sql_source_headers = []
    app.update_sql_target_key_combo.configure(values=[])
    app.update_sql_source_key_combo.configure(values=[])
    if not mapping_value:
        return
    mapping_path = Path(mapping_value)
    if not mapping_path.exists():
        return
    try:
        _, target_values, source_values = load_update_field_mappings(mapping_path)
    except Exception as exc:
        messagebox.showerror("加载失败", str(exc))
        return
    app.update_sql_target_headers = target_values
    app.update_sql_source_headers = source_values
    app.update_sql_target_key_combo.configure(values=target_values)
    app.update_sql_source_key_combo.configure(values=source_values)
    if target_values and not app.update_sql_target_key_var.get().strip():
        app.update_sql_target_key_var.set(target_values[0])
    if source_values and not app.update_sql_source_key_var.get().strip():
        app.update_sql_source_key_var.set(source_values[0])
    app.update_sql_status_var.set("字段已加载")


def set_update_sql_result_text(app, content: str) -> None:
    if app.update_sql_result_text is None:
        return
    app.update_sql_result_text.configure(state="normal")
    app.update_sql_result_text.delete("1.0", tk.END)
    app.update_sql_result_text.insert("1.0", content)
    app.update_sql_result_text.configure(state="disabled")


def copy_update_sql(app) -> None:
    if app.last_update_sql_result is None:
        return
    app.root.clipboard_clear()
    app.root.clipboard_append(app.last_update_sql_result.sql_content)
    app.root.update()
    app.update_sql_status_var.set("已复制 SQL")


def update_update_sql_ui(app, result: UpdateSqlResult | None) -> None:
    app.last_update_sql_result = result
    if result is None:
        app.update_sql_result_var.set("更新 SQL 会显示在这里")
        set_update_sql_result_text(app, app.update_sql_result_var.get())
        app.update_sql_copy_button.configure(state="disabled")
        return

    app.update_sql_result_var.set(
        "SQL 生成完成：\n"
        f"映射模板：{result.mapping_path}\n"
        f"考生表：{result.target_table}\n"
        f"临时表：{result.source_table}\n"
        f"备份表：{result.backup_target_table} / {result.backup_source_table}\n\n"
        f"{result.sql_content}"
    )
    set_update_sql_result_text(app, app.update_sql_result_var.get())
    app.update_sql_copy_button.configure(state="normal")


def start_update_sql_render(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    mapping_value = app.update_sql_mapping_var.get().strip()
    if not mapping_value:
        messagebox.showerror("参数错误", "请选择字段映射模板。")
        return
    mapping_path = Path(mapping_value)
    if not mapping_path.exists():
        messagebox.showerror("参数错误", f"未找到映射模板：{mapping_path}")
        return

    target_table = app.update_sql_target_table_var.get().strip()
    source_table = app.update_sql_source_table_var.get().strip()
    target_key = app.update_sql_target_key_var.get().strip()
    source_key = app.update_sql_source_key_var.get().strip()
    if not target_table or not source_table:
        messagebox.showerror("参数错误", "请输入考生表名称和临时表名称。")
        return
    if not target_key or not source_key:
        messagebox.showerror("参数错误", "请选择关联字段。")
        return

    app.save_settings()
    app.update_sql_run_button.configure(state="disabled")
    app.update_sql_copy_button.configure(state="disabled")
    app.update_sql_status_var.set("生成中")
    app.update_sql_result_var.set("正在生成 SQL，请稍候...")
    set_update_sql_result_text(app, app.update_sql_result_var.get())
    app.write_log(f"启动更新SQL生成任务：{target_table} <- {source_table}")

    def runner() -> None:
        try:
            result = render_update_sql(
                mapping_path=mapping_path,
                target_table=target_table,
                source_table=source_table,
                target_key_column=target_key,
                source_key_column=source_key,
                ignore_empty=app.update_sql_ignore_empty_var.get(),
                logger=app.make_logger(),
            )
        except Exception as exc:
            app.log_queue.put(f"__UPDATE_SQL_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "update_sql_result", "result": result})
            app.log_queue.put("__UPDATE_SQL_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
