from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

from ..phone_decrypt import PhoneDecryptOptions, PhoneDecryptSummary, load_filter_id_cards, run_phone_decrypt


def choose_phone_filter_file(app) -> None:
    selected = filedialog.askopenfilename(
        title="选择名单文件",
        initialdir=str(Path(app.phone_filter_file_var.get()).parent)
        if app.phone_filter_file_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
    )
    if selected:
        app.phone_filter_file_var.set(selected)


def fill_phone_table_name(app) -> None:
    exam_sort = app.phone_exam_sort_var.get().strip()
    if exam_sort:
        app.phone_candidate_table_var.set(f"考生表{exam_sort}")
        fill_phone_output_path(app)


def fill_phone_output_path(app) -> None:
    table_name = app.phone_candidate_table_var.get().strip()
    if not table_name:
        return
    app.phone_output_path_var.set(str(Path.cwd() / f"{table_name}_电话解密结果.xlsx"))


def update_phone_mode_ui(app) -> None:
    state = "normal" if app.phone_mode_var.get() == "partial" else "disabled"
    app.phone_filter_entry.configure(state=state)
    app.phone_filter_button.configure(state=state)


def set_phone_result_text(app, content: str) -> None:
    if app.phone_result_text is None:
        return
    app.phone_result_text.configure(state="normal")
    app.phone_result_text.delete("1.0", tk.END)
    app.phone_result_text.insert("1.0", content)
    app.phone_result_text.configure(state="disabled")


def update_phone_summary_ui(app, summary: Optional[PhoneDecryptSummary]) -> None:
    app.last_phone_summary = summary
    if summary is None:
        app.phone_result_var.set("电话解密结果会显示在这里")
        set_phone_result_text(app, app.phone_result_var.get())
        app.phone_open_button.configure(state="disabled")
        return

    app.phone_result_var.set(
        "电话解密完成：\n"
        f"报名库：{summary.signup_database}\n"
        f"电话库：{summary.phone_database}\n"
        f"考生表：{summary.candidate_table}\n"
        f"结果文件：{summary.output_path}\n"
        f"解密组件：{summary.backend_name}\n"
        f"总记录：{summary.total_rows}\n"
        f"命中密文：{summary.matched_info_rows}\n"
        f"解密成功：{summary.decrypted_rows}\n"
        f"回写成功：{summary.updated_rows}\n"
        f"跳过：{summary.skipped_rows}\n"
        f"失败：{summary.failed_rows}"
    )
    set_phone_result_text(app, app.phone_result_var.get())
    app.phone_open_button.configure(state="normal")


def open_phone_report_file(app) -> None:
    if app.last_phone_summary is None:
        return
    app.open_local_file(app.last_phone_summary.output_path, "未找到电话解密结果文件")


def start_phone_decrypt_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    server = app.phone_server_var.get().strip()
    username = app.phone_username_var.get().strip()
    password = app.phone_password_var.get().strip()
    signup_database = app.phone_signup_database_var.get().strip()
    phone_database = app.phone_info_database_var.get().strip() or signup_database
    candidate_table = app.phone_candidate_table_var.get().strip()
    exam_sort = app.phone_exam_sort_var.get().strip()
    if exam_sort and not candidate_table:
        candidate_table = f"考生表{exam_sort}"
        app.phone_candidate_table_var.set(candidate_table)
    if not candidate_table:
        messagebox.showerror("参数错误", "请输入考生表名称，或先填写考试代码生成表名。")
        return

    try:
        port = int(app.phone_port_var.get().strip() or "1433")
    except ValueError:
        messagebox.showerror("参数错误", "端口必须是数字。")
        return

    candidate_id_cards = None
    if app.phone_mode_var.get() == "partial":
        filter_path_value = app.phone_filter_file_var.get().strip()
        if not filter_path_value:
            messagebox.showerror("参数错误", "按名单解密时，请选择名单文件。")
            return
        try:
            candidate_id_cards = load_filter_id_cards(Path(filter_path_value))
        except Exception as exc:
            messagebox.showerror("名单文件错误", str(exc))
            return
        if not candidate_id_cards:
            messagebox.showerror("名单文件错误", "名单文件中没有可用的身份证号。")
            return

    output_value = app.phone_output_path_var.get().strip()
    output_path = Path(output_value) if output_value else Path.cwd() / f"{candidate_table}_电话解密结果.xlsx"
    app.phone_output_path_var.set(str(output_path))
    app.save_settings()

    app.phone_run_button.configure(state="disabled")
    app.phone_open_button.configure(state="disabled")
    app.phone_status_var.set("解密中")
    app.phone_result_var.set("正在查询、解密并回写电话，请稍候...")
    set_phone_result_text(app, app.phone_result_var.get())
    app.write_log(
        f"启动电话解密任务：{signup_database}.{candidate_table} -> {phone_database}.web_info"
    )

    options = PhoneDecryptOptions(
        server=server,
        port=port,
        username=username,
        password=password,
        signup_database=signup_database,
        phone_database=phone_database,
        candidate_table=candidate_table,
        candidate_filter_mode=app.phone_mode_var.get().strip(),
        candidate_id_cards=candidate_id_cards,
        output_path=output_path,
    )

    def runner() -> None:
        try:
            summary = run_phone_decrypt(options, logger=app.make_logger())
        except Exception as exc:
            app.log_queue.put(f"__PHONE_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "phone_summary", "summary": summary})
            app.log_queue.put("__PHONE_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
