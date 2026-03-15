from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from ..word_to_html import WordExportResult, export_word_to_html


def choose_word_source(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.word_source_var.get()).parent)
        if app.word_source_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("表样文件", "*.docx *.doc *.xlsx"), ("所有文件", "*.*")],
    )
    if selected:
        app.word_source_var.set(selected)


def update_word_export_ui(app, result: Optional[WordExportResult]) -> None:
    app.last_word_export = result
    if result is None:
        app.word_result_var.set("表样转换结果会显示在这里")
        app.word_preview_status_var.set("未生成预览")
        set_word_code(app, "")
        render_word_preview(app, "")
        app.word_copy_button.configure(state="disabled")
        app.word_open_browser_button.configure(state="disabled")
        return
    variant_label = "Net版" if result.variant == "net" else "Java版"
    app.word_result_var.set(
        f"{variant_label}导出完成：\n源文件：{result.source_path}\nHTML 代码已生成，可直接预览或复制。"
    )
    app.word_preview_status_var.set(f"{variant_label}预览")
    set_word_code(app, result.html_content)
    render_word_preview(app, result.preview_html)
    app.word_copy_button.configure(state="normal")
    app.word_open_browser_button.configure(state="normal")


def set_word_code(app, html_content: str) -> None:
    if app.word_code_text is None:
        return
    app.word_code_text.configure(state="normal")
    app.word_code_text.delete("1.0", "end")
    app.word_code_text.insert("1.0", html_content)
    app.word_code_text.configure(state="disabled")


def render_word_preview(app, html_content: str) -> None:
    for child in app.word_preview_container.winfo_children():
        child.destroy()
    if not html_content:
        ttk.Label(
            app.word_preview_container,
            text="导出后会在这里显示 HTML 预览。",
        ).grid(row=0, column=0, sticky="nw")
        return
    if os.name == "nt":
        ttk.Label(
            app.word_preview_container,
            text="Windows 下已关闭程序内置预览。请点击上方“浏览器预览”查看 HTML 效果。",
            wraplength=760,
            justify="left",
        ).grid(row=0, column=0, sticky="nw")
        app.word_preview_widget = None
        app.word_preview_status_var.set(app.word_preview_status_var.get() + "（Windows 建议使用浏览器预览）")
        return
    try:
        from tkinterweb import HtmlFrame

        preview = HtmlFrame(
            app.word_preview_container,
            messages_enabled=False,
        )
        preview.load_html(html_content)
        preview.grid(row=0, column=0, sticky="nsew")
        app.word_preview_widget = preview
        app.word_preview_status_var.set(app.word_preview_status_var.get() + "（表格增强预览）")
    except Exception:
        try:
            from tkhtmlview import HTMLScrolledText

            preview = HTMLScrolledText(
                app.word_preview_container,
                html=html_content,
            )
            preview.grid(row=0, column=0, sticky="nsew")
            app.word_preview_widget = preview
            app.word_preview_status_var.set(app.word_preview_status_var.get() + "（兼容预览）")
        except Exception:
            ttk.Label(
                app.word_preview_container,
                text="当前环境缺少 HTML 预览依赖，请先安装 requirements 后重启程序。",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, sticky="nw")
            app.word_preview_widget = None


def copy_word_html(app) -> None:
    if app.last_word_export is None:
        return
    app.root.clipboard_clear()
    app.root.clipboard_append(app.last_word_export.html_content)
    app.root.update()
    app.word_status_var.set("已复制 HTML")


def open_word_preview_in_browser(app) -> None:
    if app.last_word_export is None:
        return
    temp_file = Path(tempfile.gettempdir()) / f"word_preview_{app.last_word_export.variant}.html"
    temp_file.write_text(app.last_word_export.html_content, encoding="utf-8")
    try:
        if os.name == "nt":
            os.startfile(str(temp_file))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(temp_file)])
        else:
            subprocess.Popen(["xdg-open", str(temp_file)])
    except Exception as exc:
        messagebox.showerror("打开失败", str(exc))


def start_word_export(app, variant: str) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    source_value = app.word_source_var.get().strip()
    if not source_value:
        messagebox.showerror("参数错误", "请选择表样文件。")
        return
    source_path = Path(source_value)
    if not source_path.exists():
        messagebox.showerror("参数错误", f"表样文件不存在：{source_path}")
        app.word_status_var.set("失败")
        app.word_result_var.set(f"导出失败：\n表样文件不存在：{source_path}")
        app.word_preview_status_var.set("预览不可用")
        set_word_code(app, "")
        render_word_preview(app, "")
        app.word_copy_button.configure(state="disabled")
        app.word_open_browser_button.configure(state="disabled")
        return
    if source_path.suffix.lower() not in {".doc", ".docx", ".xlsx"}:
        messagebox.showerror("参数错误", "仅支持 `.doc`、`.docx` 或 `.xlsx` 文件。")
        app.word_status_var.set("失败")
        app.word_result_var.set("导出失败：\n仅支持 `.doc`、`.docx` 或 `.xlsx` 文件。")
        app.word_preview_status_var.set("预览不可用")
        set_word_code(app, "")
        render_word_preview(app, "")
        app.word_copy_button.configure(state="disabled")
        app.word_open_browser_button.configure(state="disabled")
        return

    app.save_settings()
    app.word_net_button.configure(state="disabled")
    app.word_java_button.configure(state="disabled")
    app.word_status_var.set("导出中")
    app.word_result_var.set("正在导出 HTML...")
    app.write_log("")
    app.write_log("=" * 60)
    app.write_log(f"启动表样转换任务：{variant}")

    def runner() -> None:
        try:
            result = export_word_to_html(
                source_path=source_path,
                variant=variant,
                logger=app.make_logger(),
            )
        except Exception as exc:
            app.log_queue.put(f"__WORD_EXPORT_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "word_export", "result": result})
            app.log_queue.put("__WORD_EXPORT_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
