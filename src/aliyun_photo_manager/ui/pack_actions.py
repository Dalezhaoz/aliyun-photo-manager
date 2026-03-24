from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

from ..result_packer import PackSummary, pack_encrypted_folder, query_pack_history


def update_pack_summary_ui(app, summary: Optional[PackSummary]) -> None:
    app.last_pack_summary = summary
    if summary is None:
        app.pack_result_var.set("结果打包信息会显示在这里")
        app.pack_copy_password_button.configure(state="disabled")
        app.pack_open_button.configure(state="disabled")
        return

    app.pack_result_var.set(
        "打包完成：\n"
        f"来源对象：{summary.source_path}\n"
        f"压缩包：{summary.output_path}\n"
        f"文件数：{summary.file_count}\n"
        f"时间：{summary.created_at}\n"
        f"密码：{summary.password}"
    )
    app.pack_copy_password_button.configure(state="normal")
    app.pack_open_button.configure(state="normal")


def set_pack_query_result_text(app, content: str) -> None:
    if app.pack_query_result_text is None:
        return
    app.pack_query_result_text.configure(state="normal")
    app.pack_query_result_text.delete("1.0", tk.END)
    app.pack_query_result_text.insert("1.0", content)
    app.pack_query_result_text.configure(state="disabled")


def update_pack_password_mode_ui(app) -> None:
    if app.pack_use_custom_password_var.get():
        app.pack_password_entry.configure(state="normal")
    else:
        app.pack_password_var.set("")
        app.pack_password_entry.configure(state="disabled")


def run_pack_history_query(app) -> None:
    keyword = app.pack_query_var.get().strip()
    records = query_pack_history(keyword)
    if not records:
        set_pack_query_result_text(app, "未找到匹配的打包记录。")
        app.pack_status_var.set("未找到记录")
        return
    latest = records[0]
    set_pack_query_result_text(
        app,
        (
            f"最近匹配记录：\n"
            f"来源名称：{latest.get('source_name', '')}\n"
            f"压缩包：{latest.get('archive_name', '')}\n"
            f"时间：{latest.get('created_at', '')}\n"
            f"密码：{latest.get('password', '')}\n"
            f"路径：{latest.get('output_path', '')}"
        ),
    )
    app.pack_status_var.set("已查询密码")


def copy_pack_password(app) -> None:
    if app.last_pack_summary is None:
        return
    app.root.clipboard_clear()
    app.root.clipboard_append(app.last_pack_summary.password)
    app.root.update()
    app.pack_status_var.set("已复制密码")


def choose_pack_source_file(app) -> None:
    selected = filedialog.askopenfilename(
        title="选择待打包文件",
        initialdir=str(Path(app.pack_source_dir_var.get()).parent)
        if app.pack_source_dir_var.get().strip()
        else str(Path.cwd()),
    )
    if selected:
        app.pack_source_dir_var.set(selected)


def choose_pack_source_directory(app) -> None:
    selected = filedialog.askdirectory(
        title="选择待打包文件夹",
        initialdir=str(Path(app.pack_source_dir_var.get()).parent)
        if app.pack_source_dir_var.get().strip()
        else str(Path.cwd()),
    )
    if selected:
        app.pack_source_dir_var.set(selected)


def open_pack_file(app) -> None:
    if app.last_pack_summary is None:
        return
    app.open_local_file(app.last_pack_summary.output_path, "未找到压缩包")


def start_pack_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    source_value = app.pack_source_dir_var.get().strip()
    if not source_value:
        messagebox.showerror("参数错误", "请选择待打包文件或文件夹。")
        return

    source_dir = Path(source_value)
    if not source_dir.exists():
        messagebox.showerror("参数错误", f"待打包文件或文件夹不存在：{source_dir}")
        return

    output_value = app.pack_output_dir_var.get().strip()
    output_dir = Path(output_value) if output_value else source_dir.parent
    app.pack_output_dir_var.set(str(output_dir))
    custom_password = app.pack_password_var.get().strip()
    if app.pack_use_custom_password_var.get() and not custom_password:
        messagebox.showerror("参数错误", "已勾选手动设置密码，请输入打包密码。")
        return
    app.save_settings()

    app.pack_run_button.configure(state="disabled")
    app.pack_copy_password_button.configure(state="disabled")
    app.pack_open_button.configure(state="disabled")
    app.pack_status_var.set("打包中")
    app.pack_result_var.set("正在压缩并加密，请稍候...")
    app.write_log(f"启动结果打包任务：{source_dir}")

    def runner() -> None:
        try:
            summary = pack_encrypted_folder(
                source_dir=source_dir,
                output_dir=output_dir,
                password=custom_password if app.pack_use_custom_password_var.get() else None,
                logger=app.make_logger(),
            )
        except Exception as exc:
            app.log_queue.put(f"__PACK_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "pack_summary", "summary": summary})
            app.log_queue.put("__PACK_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
