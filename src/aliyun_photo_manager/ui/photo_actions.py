from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import List, Optional

from ..app import WorkflowSummary, resolve_photo_directories, run_photo_classification_only, run_photo_download_and_template
from ..certificate_filter import list_template_headers, load_match_values
from ..downloader import BrowserEntry, count_photos_in_prefix, download_objects, is_photo_key, list_browser_entries
from ..excel_classifier import generate_template


def choose_photo_template(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.photo_template_var.get()).parent)
        if app.photo_template_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    if selected:
        app.photo_template_var.set(selected)
        load_photo_headers(app)


def set_photo_headers(app, headers: List[str]) -> None:
    app.photo_headers = headers
    app.photo_match_combo["values"] = headers
    current_value = app.photo_match_column_var.get().strip()
    if current_value and current_value not in headers:
        app.photo_match_column_var.set("")
    if not app.photo_match_column_var.get().strip() and headers:
        app.photo_match_column_var.set(headers[0])


def load_photo_headers(app) -> None:
    template_path = app.photo_template_var.get().strip()
    if not template_path:
        messagebox.showinfo("缺少模板", "请先选择人员模板文件。")
        return
    try:
        headers = list_template_headers(Path(template_path))
    except Exception as exc:
        messagebox.showerror("读取失败", str(exc))
        return
    set_photo_headers(app, headers)
    app.write_log(f"已读取照片模板列：{', '.join(headers) if headers else '无'}")


def update_photo_source_mode_ui(app) -> None:
    is_oss = app.photo_source_mode_var.get() == "oss"
    app.skip_download_var.set(not is_oss)
    app.set_widget_state_recursive(app.photo_oss_frame, "normal" if is_oss else "disabled")
    app.set_widget_state_recursive(app.photo_browser_frame, "normal" if is_oss else "disabled")
    if is_oss:
        app.run_button.configure(text="下载并生成模板")
        app.folder_status_var.set("未加载 bucket 文件夹")
        app.search_status_var.set("未搜索文件夹")
        app.selected_folder_info_var.set("当前未选择 bucket 文件夹")
    else:
        app.run_button.configure(text="生成本地模板")
        app.folder_status_var.set("本地模式不使用 bucket 浏览")
        app.search_status_var.set("本地模式不使用 bucket 搜索")
        app.selected_folder_info_var.set("请直接选择本地照片目录")


def update_summary_ui(app, summary: Optional[WorkflowSummary]) -> None:
    app.last_summary = summary
    if summary is None:
        app.summary_text_var.set("任务结果会显示在这里")
        app.open_template_button.configure(state="disabled")
        app.open_photo_report_button.configure(state="disabled")
        return

    lines = []
    if summary.download_result is not None:
        lines.append(
            f"下载任务完成：共找到 {summary.download_result.total_found} 个文件，"
            f"新下载 {summary.download_result.downloaded_count} 个，"
            f"跳过已存在 {summary.download_result.skipped_existing_count} 个。"
        )
    if summary.template_file_count:
        lines.append(f"模板文件统计：当前目录共 {summary.template_file_count} 个文件。")
    if summary.template_created:
        lines.append("已生成 Excel 模板，可以直接打开填写。")
    elif summary.classified_count:
        lines.append(f"分类复制完成，共处理 {summary.classified_count} 个文件。")
        if summary.report_path is not None:
            lines.append(f"已导出结果清单：{summary.report_path.name}")
    elif summary.cancelled:
        lines.append("任务已取消。")
    elif summary.dry_run:
        lines.append("当前为预览模式，没有实际写入文件。")
    if not lines:
        lines.append("任务已完成。")

    app.summary_text_var.set("\n".join(lines))
    if summary.template_path.exists():
        app.open_template_button.configure(state="normal")
    else:
        app.open_template_button.configure(state="disabled")
    if summary.report_path is not None and summary.report_path.exists():
        app.open_photo_report_button.configure(state="normal")
    else:
        app.open_photo_report_button.configure(state="disabled")


def open_template_file(app) -> None:
    if app.last_summary is None:
        return
    app.open_local_file(app.last_summary.template_path, "未找到 Excel 模板")


def open_photo_report_file(app) -> None:
    if app.last_summary is None or app.last_summary.report_path is None:
        return
    app.open_local_file(app.last_summary.report_path, "未找到结果清单")


def load_bucket_folders(app) -> None:
    if app.photo_source_mode_var.get() != "oss":
        return
    try:
        config = app.build_config()
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    current_prefix = app.prefix_var.get().strip()
    app.folder_status_var.set("正在加载 bucket 文件夹...")
    app.write_log(f"加载 bucket 文件夹：{current_prefix or '/'}")

    def worker() -> None:
        try:
            entries = list_browser_entries(config, current_prefix)
        except Exception as exc:
            app.root.after(0, lambda: finish_folder_load(app, error=f"{type(exc).__name__}: {exc}"))
            return
        app.root.after(0, lambda: finish_folder_load(app, entries=entries))

    threading.Thread(target=worker, daemon=True).start()


def finish_folder_load(app, entries: Optional[List[BrowserEntry]] = None, error: Optional[str] = None) -> None:
    if error is not None:
        app.folder_status_var.set("加载失败")
        app.write_log(f"加载文件夹失败：{error}")
        messagebox.showerror("加载失败", app.format_cloud_error(error))
        return

    entries = entries or []
    app.current_folder_entries = entries
    folders = [entry.key for entry in entries if entry.entry_type == "folder"]
    app.set_folder_values(folders)
    render_folder_tree(app, entries)
    folder_count = len([entry for entry in entries if entry.entry_type == "folder"])
    file_count = len([entry for entry in entries if entry.entry_type == "file"])
    if entries:
        app.folder_status_var.set(f"已加载 {folder_count} 个文件夹，{file_count} 个文件")
        app.write_log(f"当前层级找到 {folder_count} 个文件夹，{file_count} 个文件。")
    else:
        app.folder_status_var.set("当前层级没有子文件夹和文件")
        app.write_log("当前层级没有可显示的文件夹或文件。")


def render_folder_tree(app, entries: List[BrowserEntry]) -> None:
    if app.folder_tree is None:
        return
    for item in app.folder_tree.get_children():
        app.folder_tree.delete(item)
    app.folder_nodes = {}
    for entry in entries:
        meta = "文件夹" if entry.entry_type == "folder" else "文件"
        node_id = app.folder_tree.insert("", "end", text=entry.display_name, values=(meta,))
        app.folder_nodes[node_id] = entry


def on_tree_select(app, _event=None) -> None:
    if app.folder_tree is None:
        return
    selected = app.folder_tree.selection()
    if not selected:
        return
    entry = app.folder_nodes.get(selected[0])
    if entry is None:
        return
    if entry.entry_type == "folder":
        app.prefix_var.set(entry.key)
        app.selected_folder_info_var.set(f"当前已选目录：{entry.key or '/'}")
    else:
        app.selected_folder_info_var.set(f"当前已选文件：{entry.key}")


def on_tree_double_click(app, _event=None) -> None:
    if app.folder_tree is None:
        return
    selected = app.folder_tree.selection()
    if not selected:
        return
    entry = app.folder_nodes.get(selected[0])
    if entry is None:
        return
    if entry.entry_type == "folder" and entry.key:
        app.prefix_var.set(entry.key)
        load_bucket_folders(app)


def filter_folder_entries(app, entries: List[BrowserEntry], keyword: str) -> List[BrowserEntry]:
    cleaned_keyword = keyword.strip().lower()
    if not cleaned_keyword:
        return entries
    return [
        entry for entry in entries if entry.entry_type == "folder" and cleaned_keyword in entry.display_name.lower()
    ]


def search_bucket_files(app) -> None:
    if app.photo_source_mode_var.get() != "oss":
        messagebox.showinfo("本地模式", "本地模式不使用 bucket 搜索。")
        return

    keyword = app.search_keyword_var.get().strip()
    if not keyword:
        messagebox.showinfo("缺少关键词", "请输入要搜索的文件夹名称关键词。")
        return

    current_prefix = app.prefix_var.get().strip() or "/"
    app.search_status_var.set("正在筛选当前层级文件夹...")
    app.write_log(f"开始筛选当前层级文件夹：{current_prefix}，关键词：{keyword}")
    finish_search(app, entries=filter_folder_entries(app, app.current_folder_entries, keyword))


def finish_search(app, entries: Optional[List[BrowserEntry]] = None, error: Optional[str] = None) -> None:
    if error is not None:
        app.search_status_var.set("搜索失败")
        app.write_log(f"搜索文件夹失败：{error}")
        messagebox.showerror("搜索失败", app.format_cloud_error(error))
        return

    entries = entries or []
    render_folder_tree(app, entries)
    if entries:
        app.search_status_var.set(f"找到 {len(entries)} 个匹配文件夹")
        app.write_log(f"搜索完成，找到 {len(entries)} 个匹配文件夹。")
    else:
        app.search_status_var.set("没有找到匹配文件夹")
        app.write_log("没有找到匹配文件夹。")


def on_search_double_click(app, _event=None) -> None:
    if app.search_tree is None:
        return
    selected = app.search_tree.selection()
    if not selected:
        return

    object_key = app.search_nodes.get(selected[0], "")
    if not object_key:
        return

    parent_folder = str(Path(object_key).parent)
    if parent_folder == ".":
        app.prefix_var.set("")
        app.selected_folder_info_var.set(f"当前已选：/，文件：{Path(object_key).name}")
    else:
        normalized_parent = parent_folder.strip("/") + "/"
        app.prefix_var.set(normalized_parent)
        app.selected_folder_info_var.set(f"当前已选：{normalized_parent}，文件：{Path(object_key).name}")
    app.write_log(f"已根据搜索结果定位到目录：{app.prefix_var.get() or '/'}")


def refresh_selected_folder_count(app) -> None:
    if app.photo_source_mode_var.get() != "oss":
        messagebox.showinfo("本地模式", "本地模式不使用 bucket 统计。")
        return
    try:
        config = app.build_config()
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    prefix = app.prefix_var.get().strip()
    if not prefix:
        messagebox.showinfo("未选择文件夹", "请先选择 bucket 文件夹。")
        return

    app.folder_status_var.set("正在统计已选文件夹图片数量...")
    app.selected_folder_info_var.set("正在统计图片数量...")
    app.write_log(f"统计图片数量：{prefix}")

    def worker() -> None:
        try:
            count = count_photos_in_prefix(config, prefix)
        except Exception as exc:
            app.root.after(0, lambda: finish_count_refresh(app, error=f"{type(exc).__name__}: {exc}"))
            return
        app.root.after(0, lambda: finish_count_refresh(app, count=count, prefix=prefix))

    threading.Thread(target=worker, daemon=True).start()


def finish_count_refresh(app, count: int = 0, prefix: str = "", error: Optional[str] = None) -> None:
    if error is not None:
        app.folder_status_var.set("统计失败")
        app.selected_folder_info_var.set("统计失败")
        app.write_log(f"统计图片数量失败：{error}")
        messagebox.showerror("统计失败", app.format_cloud_error(error))
        return

    if app.folder_tree is not None:
        for node_id, entry in app.folder_nodes.items():
            if entry.entry_type == "folder" and entry.key == prefix:
                app.folder_tree.item(node_id, values=(f"{count} 张图片",))
                break

    app.folder_status_var.set("统计完成")
    app.selected_folder_info_var.set(f"当前已选：{prefix}，图片数：{count}")
    app.write_log(f"{prefix} 下共有 {count} 张图片。")


def start_photo_download_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    try:
        options = app.build_options()
        oss_config = None if options.skip_download else app.build_config()
        photo_match_values: List[str] = []
        if app.photo_source_mode_var.get() == "oss" and app.photo_filter_by_template_var.get():
            template_value = app.photo_template_var.get().strip()
            match_column = app.photo_match_column_var.get().strip()
            if not template_value:
                raise ValueError("已勾选按模板名单下载，请先选择人员模板。")
            if not match_column:
                raise ValueError("已勾选按模板名单下载，请先选择匹配列。")
            template_path = Path(template_value)
            if not template_path.exists():
                raise ValueError("人员模板文件不存在，请重新选择。")
            photo_match_values = load_match_values(template_path, match_column)
            if not photo_match_values:
                raise ValueError("照片模板匹配列没有可用数据，无法按名单下载。")
        app.save_settings()
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    app.run_button.configure(state="disabled")
    app.photo_classify_button.configure(state="disabled")
    app.cancel_button.configure(state="normal")
    app.status_var.set("运行中")
    app.progress_text_var.set("准备开始...")
    update_summary_ui(app, None)
    app.cancel_event.clear()
    if app.progress_bar is not None:
        app.progress_bar["value"] = 0
    app.write_log("")
    app.write_log("=" * 60)
    app.write_log("启动照片下载/模板任务。")
    if photo_match_values:
        app.write_log(
            f"本次仅下载模板中的照片，匹配列：{app.photo_match_column_var.get().strip()}，共 {len(photo_match_values)} 人。"
        )

    def runner() -> None:
        try:
            if photo_match_values and oss_config is not None:
                download_dir, sorted_dir = resolve_photo_directories(options)
                allowed_stems = set(photo_match_values)

                def key_filter(object_key: str) -> bool:
                    relative_path = object_key
                    normalized_prefix = app.prefix_var.get().strip().strip("/")
                    if normalized_prefix:
                        prefix_with_slash = normalized_prefix + "/"
                        if object_key.startswith(prefix_with_slash):
                            relative_path = object_key[len(prefix_with_slash):]
                    filename_stem = Path(relative_path.lstrip("/")).stem
                    return filename_stem in allowed_stems

                app.write_log(f"实际下载目录：{download_dir}")
                download_result = download_objects(
                    config=oss_config,
                    prefix=options.prefix,
                    download_dir=download_dir,
                    dry_run=options.dry_run,
                    skip_existing=options.skip_existing,
                    logger=app.make_logger(),
                    progress_callback=app.make_progress_callback(),
                    cancel_event=app.cancel_event,
                    key_filter=key_filter,
                    file_filter=is_photo_key,
                    stage="download",
                )
                if app.cancel_event.is_set():
                    summary = WorkflowSummary(
                        download_dir=download_dir,
                        sorted_dir=sorted_dir,
                        template_path=download_dir / "照片分类模板.xlsx",
                        download_result=download_result,
                        template_file_count=0,
                        classified_count=0,
                        template_created=False,
                        cancelled=True,
                        dry_run=options.dry_run,
                    )
                else:
                    template_result = generate_template(
                        source_dir=download_dir,
                        dry_run=options.dry_run,
                        logger=app.make_logger(),
                    )
                    summary = WorkflowSummary(
                        download_dir=download_dir,
                        sorted_dir=sorted_dir,
                        template_path=template_result.template_path,
                        download_result=download_result,
                        template_file_count=template_result.file_count,
                        classified_count=0,
                        template_created=template_result.created,
                        cancelled=False,
                        dry_run=options.dry_run,
                    )
            else:
                summary = run_photo_download_and_template(
                    options=options,
                    oss_config=oss_config,
                    logger=app.make_logger(),
                    progress_callback=app.make_progress_callback(),
                    cancel_event=app.cancel_event,
                )
        except Exception as exc:
            app.log_queue.put(f"__TASK_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "summary", "summary": summary})
            if app.cancel_event.is_set():
                app.log_queue.put("__TASK_CANCELLED__")
            else:
                app.log_queue.put("__TASK_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()


def start_photo_classify_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    try:
        options = app.build_options()
        app.save_settings()
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    app.run_button.configure(state="disabled")
    app.photo_classify_button.configure(state="disabled")
    app.cancel_button.configure(state="normal")
    app.status_var.set("运行中")
    app.progress_text_var.set("准备开始...")
    update_summary_ui(app, None)
    app.cancel_event.clear()
    if app.progress_bar is not None:
        app.progress_bar["value"] = 0
    app.write_log("")
    app.write_log("=" * 60)
    app.write_log("启动照片分类任务。")

    def runner() -> None:
        try:
            summary = run_photo_classification_only(
                options=options,
                logger=app.make_logger(),
                cancel_event=app.cancel_event,
            )
        except Exception as exc:
            app.log_queue.put(f"__TASK_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "summary", "summary": summary})
            if app.cancel_event.is_set():
                app.log_queue.put("__TASK_CANCELLED__")
            else:
                app.log_queue.put("__TASK_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
