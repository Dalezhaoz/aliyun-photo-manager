from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import List, Optional

from ..certificate_filter import CertificateFilterSummary, list_template_headers, load_match_values, run_certificate_filter
from ..downloader import BrowserEntry, download_objects, list_browser_entries, list_buckets


def choose_certificate_template(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.certificate_template_var.get()).parent)
        if app.certificate_template_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    if selected:
        app.certificate_template_var.set(selected)
        load_certificate_headers(app)


def set_certificate_headers(app, headers: List[str]) -> None:
    app.certificate_headers = headers
    app.certificate_match_combo["values"] = headers
    app.certificate_folder_name_combo["values"] = headers
    current_value = app.certificate_match_column_var.get().strip()
    if current_value and current_value not in headers:
        app.certificate_match_column_var.set("")
    if not app.certificate_match_column_var.get().strip() and headers:
        app.certificate_match_column_var.set(headers[0])
    current_folder_name_value = app.certificate_folder_name_column_var.get().strip()
    if current_folder_name_value and current_folder_name_value not in headers:
        app.certificate_folder_name_column_var.set("")


def load_certificate_headers(app) -> None:
    template_path = app.certificate_template_var.get().strip()
    if not template_path:
        messagebox.showinfo("缺少模板", "请先选择人员模板文件。")
        return
    try:
        headers = list_template_headers(Path(template_path))
    except Exception as exc:
        messagebox.showerror("读取失败", str(exc))
        return
    set_certificate_headers(app, headers)
    app.write_log(f"已读取模板列：{', '.join(headers) if headers else '无'}")


def load_certificate_buckets(app) -> None:
    try:
        cloud_type, access_key_id, access_key_secret, endpoint = app.build_credentials()
        app.validate_cloud_endpoint(cloud_type, endpoint)
        app.check_endpoint_reachable(cloud_type, endpoint)
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    app.certificate_bucket_status_var.set("正在加载 bucket 列表...")
    app.write_log("开始加载证件资料 bucket 列表。")
    app.certificate_bucket_load_token += 1
    current_token = app.certificate_bucket_load_token
    app.root.after(8000, lambda: handle_certificate_bucket_load_timeout(app, current_token))

    def worker() -> None:
        try:
            buckets = list_buckets(access_key_id, access_key_secret, endpoint, cloud_type=cloud_type)
        except Exception as exc:
            app.root.after(
                0,
                lambda: finish_certificate_bucket_load(
                    app,
                    error=f"{type(exc).__name__}: {exc}",
                    token=current_token,
                ),
            )
            return
        app.root.after(
            0,
            lambda: finish_certificate_bucket_load(app, buckets=buckets, token=current_token),
        )

    threading.Thread(target=worker, daemon=True).start()


def finish_certificate_bucket_load(
    app,
    buckets: Optional[List[str]] = None,
    error: Optional[str] = None,
    token: Optional[int] = None,
) -> None:
    if token is not None and token != app.certificate_bucket_load_token:
        return
    if error is not None:
        app.certificate_bucket_status_var.set("加载失败")
        app.write_log(f"加载证件资料 bucket 失败：{error}")
        messagebox.showerror("加载 bucket 失败", app.format_cloud_error(error))
        return

    buckets = buckets or []
    app.sync_bucket_values(buckets)
    if buckets:
        if not app.certificate_bucket_name_var.get().strip():
            app.certificate_bucket_name_var.set(buckets[0])
        if not app.bucket_name_var.get().strip():
            app.bucket_name_var.set(buckets[0])
        app.certificate_bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
        app.bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
        app.write_log(f"证件资料找到 {len(buckets)} 个 bucket。")
        app.certificate_prefix_var.set("")
        app.certificate_folder_status_var.set("请选择 bucket 后点击“加载当前层级”")
    else:
        app.certificate_bucket_status_var.set("未找到可用 bucket")
        app.bucket_status_var.set("未找到可用 bucket")
        app.write_log("当前凭证下未找到可用证件资料 bucket。")


def handle_certificate_bucket_load_timeout(app, token: int) -> None:
    if token != app.certificate_bucket_load_token:
        return
    if app.certificate_bucket_status_var.get() != "正在加载 bucket 列表...":
        return
    app.certificate_bucket_status_var.set("加载失败")
    timeout_message = "加载 bucket 超时，请检查 Endpoint、密钥或网络后重试。"
    app.write_log(timeout_message)
    messagebox.showerror("加载 bucket 失败", timeout_message)
    app.certificate_bucket_load_token += 1


def go_to_certificate_parent_prefix(app) -> None:
    current = app.certificate_prefix_var.get().strip().strip("/")
    if not current:
        app.certificate_prefix_var.set("")
        return
    parts = current.split("/")
    parent = "/".join(parts[:-1])
    app.certificate_prefix_var.set(parent + "/" if parent else "")
    load_certificate_folders(app)


def load_certificate_folders(app) -> None:
    try:
        config = app.build_certificate_config()
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    current_prefix = app.certificate_prefix_var.get().strip()
    app.certificate_folder_status_var.set("正在加载 bucket 文件夹...")
    app.write_log(f"加载证件资料 bucket 文件夹：{current_prefix or '/'}")

    def worker() -> None:
        try:
            entries = list_browser_entries(config, current_prefix)
        except Exception as exc:
            app.root.after(
                0,
                lambda: finish_certificate_folder_load(app, error=f"{type(exc).__name__}: {exc}"),
            )
            return
        app.root.after(0, lambda: finish_certificate_folder_load(app, entries=entries))

    threading.Thread(target=worker, daemon=True).start()


def finish_certificate_folder_load(
    app,
    entries: Optional[List[BrowserEntry]] = None,
    error: Optional[str] = None,
) -> None:
    if error is not None:
        app.certificate_folder_status_var.set("加载失败")
        app.write_log(f"加载证件资料文件夹失败：{error}")
        messagebox.showerror("加载失败", app.format_cloud_error(error))
        return

    entries = entries or []
    app.current_certificate_folder_entries = entries
    folders = [entry.key for entry in entries if entry.entry_type == "folder"]
    app.set_certificate_folder_values(folders)
    render_certificate_folder_tree(app, entries)
    folder_count = len([entry for entry in entries if entry.entry_type == "folder"])
    file_count = len([entry for entry in entries if entry.entry_type == "file"])
    if entries:
        app.certificate_folder_status_var.set(f"已加载 {folder_count} 个文件夹，{file_count} 个文件")
        app.write_log(f"证件资料当前层级找到 {folder_count} 个文件夹，{file_count} 个文件。")
    else:
        app.certificate_folder_status_var.set("当前层级没有子文件夹和文件")
        app.write_log("证件资料当前层级没有可显示的文件夹或文件。")


def render_certificate_folder_tree(app, entries: List[BrowserEntry]) -> None:
    if app.certificate_folder_tree is None:
        return
    for item in app.certificate_folder_tree.get_children():
        app.certificate_folder_tree.delete(item)
    app.certificate_folder_nodes = {}
    for entry in entries:
        meta = "文件夹" if entry.entry_type == "folder" else "文件"
        node_id = app.certificate_folder_tree.insert("", "end", text=entry.display_name, values=(meta,))
        app.certificate_folder_nodes[node_id] = entry


def render_certificate_search_tree(app, object_keys: List[str]) -> None:
    if app.certificate_search_tree is None:
        return
    for item in app.certificate_search_tree.get_children():
        app.certificate_search_tree.delete(item)
    app.certificate_search_nodes = {}
    for object_key in object_keys:
        filename = Path(object_key).name
        parent_folder = str(Path(object_key).parent)
        if parent_folder == ".":
            parent_folder = "/"
        node_id = app.certificate_search_tree.insert("", "end", text=filename, values=(parent_folder,))
        app.certificate_search_nodes[node_id] = object_key


def on_certificate_tree_select(app, _event=None) -> None:
    if app.certificate_folder_tree is None:
        return
    selected = app.certificate_folder_tree.selection()
    if not selected:
        return
    entry = app.certificate_folder_nodes.get(selected[0])
    if entry is None:
        return
    if entry.entry_type == "folder":
        app.certificate_prefix_var.set(entry.key)
        app.certificate_selected_folder_info_var.set(f"当前已选目录：{entry.key or '/'}")
    else:
        app.certificate_selected_folder_info_var.set(f"当前已选文件：{entry.key}")


def on_certificate_tree_double_click(app, _event=None) -> None:
    if app.certificate_folder_tree is None:
        return
    selected = app.certificate_folder_tree.selection()
    if not selected:
        return
    entry = app.certificate_folder_nodes.get(selected[0])
    if entry is None:
        return
    if entry.entry_type == "folder" and entry.key:
        app.certificate_prefix_var.set(entry.key)
        load_certificate_folders(app)


def search_certificate_files(app) -> None:
    if app.certificate_source_mode_var.get() != "oss":
        messagebox.showinfo("本地模式", "本地模式不使用 bucket 搜索。")
        return

    keyword = app.certificate_search_keyword_var.get().strip()
    if not keyword:
        messagebox.showinfo("缺少关键词", "请输入要搜索的文件夹名称关键词。")
        return

    current_prefix = app.certificate_prefix_var.get().strip() or "/"
    app.certificate_search_status_var.set("正在筛选当前层级文件夹...")
    app.write_log(f"开始筛选证件资料当前层级文件夹：{current_prefix}，关键词：{keyword}")
    finish_certificate_search(app, entries=app.filter_folder_entries(app.current_certificate_folder_entries, keyword))


def finish_certificate_search(
    app,
    entries: Optional[List[BrowserEntry]] = None,
    error: Optional[str] = None,
) -> None:
    if error is not None:
        app.certificate_search_status_var.set("搜索失败")
        app.write_log(f"搜索证件资料文件夹失败：{error}")
        messagebox.showerror("搜索失败", app.format_cloud_error(error))
        return

    render_certificate_folder_tree(app, entries or [])
    entries = entries or []
    if entries:
        app.certificate_search_status_var.set(f"找到 {len(entries)} 个匹配文件夹")
        app.write_log(f"证件资料搜索完成，找到 {len(entries)} 个匹配文件夹。")
    else:
        app.certificate_search_status_var.set("没有找到匹配文件夹")
        app.write_log("没有找到匹配的证件资料文件夹。")


def on_certificate_search_double_click(app, _event=None) -> None:
    if app.certificate_search_tree is None:
        return
    selected = app.certificate_search_tree.selection()
    if not selected:
        return

    object_key = app.certificate_search_nodes.get(selected[0], "")
    if not object_key:
        return

    parent_folder = str(Path(object_key).parent)
    if parent_folder == ".":
        app.certificate_prefix_var.set("")
        app.certificate_selected_folder_info_var.set(f"当前已选：/，文件：{Path(object_key).name}")
    else:
        normalized_parent = parent_folder.strip("/") + "/"
        app.certificate_prefix_var.set(normalized_parent)
        app.certificate_selected_folder_info_var.set(f"当前已选：{normalized_parent}，文件：{Path(object_key).name}")
    app.write_log(f"已根据证件资料搜索结果定位到目录：{app.certificate_prefix_var.get() or '/'}")


def update_certificate_summary_ui(app, summary: Optional[CertificateFilterSummary]) -> None:
    app.last_certificate_summary = summary
    if summary is None:
        app.certificate_summary_text_var.set("证件资料筛选结果会显示在这里")
        app.open_certificate_report_button.configure(state="disabled")
        return

    if summary.total_rows == 0 and summary.download_result is not None:
        lines = [
            f"证件资料下载完成：共找到 {summary.download_result.total_found} 个文件，新下载 {summary.download_result.downloaded_count} 个，跳过已存在 {summary.download_result.skipped_existing_count} 个。",
            f"实际下载目录：{summary.source_dir}",
        ]
        if summary.cancelled:
            lines.append("任务已取消。")
        elif summary.dry_run:
            lines.append("当前为预览模式，没有实际下载文件。")
        else:
            lines.append("请到上面的“证件资料目录”查看下载结果。")
        app.certificate_summary_text_var.set("\n".join(lines))
        app.open_certificate_report_button.configure(state="disabled")
        return

    lines = [
        f"模板有效行数：{summary.total_rows}，匹配到 {summary.matched_people} 人。",
        f"缺失人员文件夹：{summary.missing_people} 人。",
        f"实际复制：{summary.copied_people} 人，{summary.copied_files} 个文件。",
    ]
    if summary.download_result is not None:
        lines.insert(
            0,
            f"证件资料下载完成：共找到 {summary.download_result.total_found} 个文件，新下载 {summary.download_result.downloaded_count} 个，跳过已存在 {summary.download_result.skipped_existing_count} 个。",
        )
    if summary.classify_output:
        lines.append("输出结构已按 分类一/分类二/分类三 建目录。")
    else:
        lines.append("输出结构未按分类字段建目录。")
    if summary.rename_folder and summary.folder_name_column:
        lines.append(f"导出后文件夹名称列：{summary.folder_name_column}")
    else:
        lines.append("导出后文件夹名称保持匹配列。")
    if summary.keyword:
        lines.append(f"筛选模式：只复制文件名包含“{summary.keyword}”的文件。")
    else:
        lines.append("筛选模式：复制整个人员文件夹。")
    if summary.cancelled:
        lines.append("任务已取消。")
    elif summary.dry_run:
        lines.append("当前为预览模式，没有实际复制文件。")
    elif summary.report_path is not None:
        lines.append(f"已导出结果清单：{summary.report_path.name}")
    app.certificate_summary_text_var.set("\n".join(lines))
    if summary.report_path is not None and summary.report_path.exists():
        app.open_certificate_report_button.configure(state="normal")
    else:
        app.open_certificate_report_button.configure(state="disabled")


def start_certificate_download_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return
    if app.certificate_source_mode_var.get() != "oss":
        messagebox.showinfo("本地模式", "本地模式下不需要下载证件资料。")
        return

    try:
        source_dir = app.resolve_certificate_source_dir(require_value=True)
        certificate_config = app.build_certificate_config()
        template_path = Path(app.certificate_template_var.get().strip())
        match_column = app.certificate_match_column_var.get().strip()
        if not template_path.exists():
            raise ValueError("请先选择有效的人员模板文件。")
        if not match_column:
            raise ValueError("请先选择匹配列。")
        match_values = load_match_values(template_path, match_column)
        if not match_values:
            raise ValueError("模板匹配列没有可用数据，无法下载证件资料。")
        app.save_settings()
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    app.certificate_download_button.configure(state="disabled")
    app.certificate_run_button.configure(state="disabled")
    app.certificate_cancel_button.configure(state="normal")
    app.certificate_status_var.set("运行中")
    app.certificate_progress_text_var.set("准备开始...")
    update_certificate_summary_ui(app, None)
    app.cancel_event.clear()
    if app.certificate_progress_bar is not None:
        app.certificate_progress_bar["value"] = 0
    app.write_log("")
    app.write_log("=" * 60)
    app.write_log("启动证件资料下载任务。")
    app.write_log(f"实际下载目录：{source_dir}")
    app.write_log(f"本次仅下载模板中的人员目录，匹配列：{match_column}，共 {len(match_values)} 人。")

    def runner() -> None:
        try:
            allowed_people = set(match_values)

            def key_filter(object_key: str) -> bool:
                relative_path = object_key
                normalized_prefix = app.certificate_prefix_var.get().strip().strip("/")
                if normalized_prefix:
                    prefix_with_slash = normalized_prefix + "/"
                    if object_key.startswith(prefix_with_slash):
                        relative_path = object_key[len(prefix_with_slash):]
                relative_path = relative_path.lstrip("/")
                if not relative_path:
                    return False
                person_folder = Path(relative_path).parts[0] if Path(relative_path).parts else ""
                return person_folder in allowed_people

            download_result = download_objects(
                config=certificate_config,
                prefix=app.certificate_prefix_var.get().strip(),
                download_dir=source_dir,
                dry_run=app.certificate_dry_run_var.get(),
                skip_existing=app.skip_existing_var.get(),
                logger=app.make_logger(),
                progress_callback=app.make_progress_callback(),
                cancel_event=app.cancel_event,
                key_filter=key_filter,
                stage="certificate_download",
            )
            summary = CertificateFilterSummary(
                template_path=template_path,
                source_dir=source_dir,
                output_dir=Path(app.certificate_output_dir_var.get().strip() or source_dir),
                match_column=match_column,
                rename_folder=app.certificate_rename_folder_var.get(),
                folder_name_column=app.certificate_folder_name_column_var.get().strip(),
                classify_output=app.certificate_classify_var.get(),
                keyword=app.certificate_keyword_var.get().strip() if app.certificate_mode_var.get() == "keyword" else "",
                total_rows=0,
                matched_people=0,
                missing_people=0,
                copied_files=0,
                copied_people=0,
                download_result=download_result,
                cancelled=app.cancel_event.is_set(),
                dry_run=app.certificate_dry_run_var.get(),
            )
        except Exception as exc:
            app.log_queue.put(f"__CERTIFICATE_TASK_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "certificate_summary", "summary": summary})
            if app.cancel_event.is_set():
                app.log_queue.put("__CERTIFICATE_TASK_CANCELLED__")
            else:
                app.log_queue.put("__CERTIFICATE_TASK_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()


def start_certificate_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    try:
        options = app.build_certificate_options()
        app.save_settings()
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    app.certificate_download_button.configure(state="disabled")
    app.certificate_run_button.configure(state="disabled")
    app.certificate_cancel_button.configure(state="normal")
    app.certificate_status_var.set("运行中")
    app.certificate_progress_text_var.set("准备开始...")
    update_certificate_summary_ui(app, None)
    app.cancel_event.clear()
    if app.certificate_progress_bar is not None:
        app.certificate_progress_bar["value"] = 0
    app.write_log("")
    app.write_log("=" * 60)
    app.write_log("启动证件资料筛选任务。")

    def runner() -> None:
        try:
            summary = run_certificate_filter(
                options=options,
                logger=app.make_logger(),
                progress_callback=app.make_progress_callback(),
                cancel_event=app.cancel_event,
            )
        except Exception as exc:
            app.log_queue.put(f"__CERTIFICATE_TASK_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "certificate_summary", "summary": summary})
            if app.cancel_event.is_set():
                app.log_queue.put("__CERTIFICATE_TASK_CANCELLED__")
            else:
                app.log_queue.put("__CERTIFICATE_TASK_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
