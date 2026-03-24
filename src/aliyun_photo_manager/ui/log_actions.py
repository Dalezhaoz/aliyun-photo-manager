import queue
from pathlib import Path


def clear_log(app) -> None:
    app.log_text.delete("1.0", "end")


def write_log(app, message: str) -> None:
    if int(app.log_text.index("end-1c").split(".")[0]) > 800:
        app.log_text.delete("1.0", "200.0")
    app.log_text.insert("end", message + "\n")
    app.log_text.see("end")


def flush_logs(app) -> None:
    while True:
        try:
            message = app.log_queue.get_nowait()
        except queue.Empty:
            break
        else:
            if isinstance(message, dict) and message.get("type") == "progress":
                app.update_progress_ui(
                    stage=message.get("stage", ""),
                    current=message.get("current", 0),
                    total=message.get("total", 0),
                    current_file=message.get("current_file", ""),
                )
            elif isinstance(message, dict) and message.get("type") == "certificate_summary":
                app.update_certificate_summary_ui(message.get("summary"))
            elif message == "__TASK_DONE__":
                app.run_button.configure(state="normal")
                app.photo_classify_button.configure(state="normal")
                app.cancel_button.configure(state="disabled")
                app.status_var.set("完成")
                app.progress_text_var.set("任务完成")
                if app.progress_bar is not None:
                    app.progress_bar["value"] = app.progress_bar["maximum"]
            elif message == "__CERTIFICATE_TASK_DONE__":
                app.certificate_download_button.configure(
                    state="normal" if app.certificate_source_mode_var.get() == "oss" else "disabled"
                )
                app.certificate_run_button.configure(state="normal")
                app.certificate_cancel_button.configure(state="disabled")
                app.certificate_status_var.set("完成")
                app.certificate_progress_text_var.set("筛选完成")
                if app.certificate_progress_bar is not None:
                    app.certificate_progress_bar["value"] = app.certificate_progress_bar["maximum"]
            elif isinstance(message, dict) and message.get("type") == "summary":
                app.update_summary_ui(message.get("summary"))
            elif message == "__TASK_CANCELLED__":
                app.run_button.configure(state="normal")
                app.photo_classify_button.configure(state="normal")
                app.cancel_button.configure(state="disabled")
                app.status_var.set("已取消")
                app.progress_text_var.set("任务已取消")
            elif message == "__CERTIFICATE_TASK_CANCELLED__":
                app.certificate_download_button.configure(
                    state="normal" if app.certificate_source_mode_var.get() == "oss" else "disabled"
                )
                app.certificate_run_button.configure(state="normal")
                app.certificate_cancel_button.configure(state="disabled")
                app.certificate_status_var.set("已取消")
                app.certificate_progress_text_var.set("筛选已取消")
            elif isinstance(message, str) and message.startswith("__TASK_FAILED__::"):
                app.run_button.configure(state="normal")
                app.photo_classify_button.configure(state="normal")
                app.cancel_button.configure(state="disabled")
                app.status_var.set("失败")
                app.progress_text_var.set("任务失败")
                app.write_log(message.split("::", 1)[1])
            elif isinstance(message, str) and message.startswith("__CERTIFICATE_TASK_FAILED__::"):
                app.certificate_download_button.configure(
                    state="normal" if app.certificate_source_mode_var.get() == "oss" else "disabled"
                )
                app.certificate_run_button.configure(state="normal")
                app.certificate_cancel_button.configure(state="disabled")
                app.certificate_status_var.set("失败")
                app.certificate_progress_text_var.set("筛选失败")
                app.write_log(message.split("::", 1)[1])
            elif isinstance(message, dict) and message.get("type") == "word_export":
                app.update_word_export_ui(message.get("result"))
            elif isinstance(message, dict) and message.get("type") == "pack_summary":
                app.update_pack_summary_ui(message.get("summary"))
            elif isinstance(message, dict) and message.get("type") == "match_summary":
                app.update_match_summary_ui(message.get("summary"))
            elif isinstance(message, dict) and message.get("type") == "update_sql_result":
                app.update_update_sql_ui(message.get("result"))
            elif isinstance(message, dict) and message.get("type") == "exam_summary":
                app.update_exam_summary_ui(message.get("summary"))
            elif isinstance(message, dict) and message.get("type") == "status_summary":
                app.update_status_summary_ui(message.get("summary"))
            elif message == "__WORD_EXPORT_DONE__":
                app.word_net_button.configure(state="normal")
                app.word_java_button.configure(state="normal")
                app.word_status_var.set("完成")
            elif message == "__MATCH_DONE__":
                app.match_run_button.configure(state="normal")
                app.match_status_var.set("完成")
            elif message == "__EXAM_DONE__":
                app.exam_run_button.configure(state="normal")
                app.exam_status_var.set("完成")
            elif message == "__PACK_DONE__":
                app.pack_run_button.configure(state="normal")
                app.pack_status_var.set("完成")
            elif message == "__UPDATE_SQL_DONE__":
                app.update_sql_run_button.configure(state="normal")
                app.update_sql_status_var.set("完成")
            elif message == "__STATUS_DONE__":
                app.status_run_button.configure(state="normal")
            elif isinstance(message, str) and message.startswith("__WORD_EXPORT_FAILED__::"):
                app.word_net_button.configure(state="normal")
                app.word_java_button.configure(state="normal")
                app.word_status_var.set("失败")
                error_text = message.split("::", 1)[1]
                app.word_result_var.set(f"导出失败：\n{error_text}")
                app.word_preview_status_var.set("预览不可用")
                app.set_word_code("")
                app.render_word_preview("")
                app.word_copy_button.configure(state="disabled")
                app.word_open_browser_button.configure(state="disabled")
                app.write_log(error_text)
            elif isinstance(message, str) and message.startswith("__PACK_FAILED__::"):
                app.pack_run_button.configure(state="normal")
                app.pack_status_var.set("失败")
                app.pack_result_var.set(f"打包失败：\n{message.split('::', 1)[1]}")
                app.pack_copy_password_button.configure(state="disabled")
                app.pack_open_button.configure(state="disabled")
                app.write_log(message.split("::", 1)[1])
            elif isinstance(message, str) and message.startswith("__MATCH_FAILED__::"):
                app.match_run_button.configure(state="normal")
                app.match_status_var.set("失败")
                app.match_result_var.set(f"匹配失败：\n{message.split('::', 1)[1]}")
                app.match_open_button.configure(state="disabled")
                app.write_log(message.split("::", 1)[1])
            elif isinstance(message, str) and message.startswith("__UPDATE_SQL_FAILED__::"):
                error_text = message.split("::", 1)[1]
                app.update_sql_run_button.configure(state="normal")
                app.update_sql_copy_button.configure(state="disabled")
                app.update_sql_status_var.set("失败")
                app.update_sql_result_var.set(f"生成失败：\n{error_text}")
                app.set_update_sql_result_text(app.update_sql_result_var.get())
                app.write_log(error_text)
            elif isinstance(message, str) and message.startswith("__EXAM_FAILED__::"):
                app.exam_run_button.configure(state="normal")
                app.exam_status_var.set("失败")
                app.exam_result_var.set(f"编排失败：\n{message.split('::', 1)[1]}")
                app.set_exam_result_text(app.exam_result_var.get())
                app.exam_open_button.configure(state="disabled")
                app.write_log(message.split("::", 1)[1])
            elif isinstance(message, str) and message.startswith("__STATUS_FAILED__::"):
                app.status_run_button.configure(state="normal")
                app.status_query_status_var.set(f"查询失败：{message.split('::', 1)[1]}")
                app.write_log(message.split("::", 1)[1])
            else:
                app.write_log(message)
    app.root.after(150, app.flush_logs)


def make_logger(app):
    def logger(message: str) -> None:
        app.log_queue.put(message)

    return logger


def make_progress_callback(app):
    def progress(stage: str, current: int, total: int, current_file: str) -> None:
        app.log_queue.put(
            {
                "type": "progress",
                "stage": stage,
                "current": current,
                "total": total,
                "current_file": current_file,
            }
        )

    return progress


def update_progress_ui(app, stage: str, current: int, total: int, current_file: str) -> None:
    if app.progress_bar is not None:
        app.progress_bar["maximum"] = max(total, 1)
        app.progress_bar["value"] = current

    if stage == "download":
        filename = Path(current_file).name if current_file else ""
        if total > 0:
            app.progress_text_var.set(f"下载进度：{current}/{total} {filename}".strip())
        else:
            app.progress_text_var.set("正在统计下载文件...")
    elif stage == "certificate":
        if app.certificate_progress_bar is not None:
            app.certificate_progress_bar["maximum"] = max(total, 1)
            app.certificate_progress_bar["value"] = current
        if total > 0:
            app.certificate_progress_text_var.set(
                f"筛选进度：{current}/{total} {current_file}".strip()
            )
        else:
            app.certificate_progress_text_var.set("正在读取模板...")
    elif stage == "certificate_download":
        if app.certificate_progress_bar is not None:
            app.certificate_progress_bar["maximum"] = max(total, 1)
            app.certificate_progress_bar["value"] = current
        filename = Path(current_file).name if current_file else ""
        if total > 0:
            app.certificate_progress_text_var.set(
                f"下载证件资料：{current}/{total} {filename}".strip()
            )
        else:
            app.certificate_progress_text_var.set("正在统计证件资料文件...")


def cancel_run(app) -> None:
    if app.worker is None or not app.worker.is_alive():
        return
    app.cancel_event.set()
    app.status_var.set("取消中")
    app.progress_text_var.set("正在取消，请稍候...")
    app.certificate_status_var.set("取消中")
    app.certificate_progress_text_var.set("正在取消，请稍候...")
    app.write_log("用户请求取消当前任务。")
