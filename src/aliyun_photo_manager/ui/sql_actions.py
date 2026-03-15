from __future__ import annotations

from pathlib import Path
from tkinter import filedialog, messagebox

from ..sql_template_executor import render_sql_template


def choose_sql_template(app) -> None:
    selected = filedialog.askopenfilename(
        title="选择 SQL 模板",
        filetypes=[("SQL / 文本模板", "*.sql *.txt"), ("所有文件", "*.*")],
        initialdir=str(Path.cwd()),
    )
    if selected:
        app.sql_template_path_var.set(selected)


def set_sql_result_text(app, content: str) -> None:
    if app.sql_result_text is None:
        return
    app.sql_result_text.configure(state="normal")
    app.sql_result_text.delete("1.0", "end")
    app.sql_result_text.insert("1.0", content)
    app.sql_result_text.configure(state="disabled")


def copy_sql_text(app) -> None:
    if not app.last_sql_result:
        messagebox.showinfo("未生成 SQL", "请先生成 SQL。")
        return
    app.root.clipboard_clear()
    app.root.clipboard_append(app.last_sql_result.sql_content)
    app.sql_status_var.set("已复制 SQL")


def start_sql_render(app) -> None:
    template_path = Path(app.sql_template_path_var.get().strip())
    try:
        result = render_sql_template(
            source_path=template_path,
            exam_code=app.sql_exam_code_var.get(),
            exam_date=app.sql_exam_date_var.get(),
            signup_start=app.sql_signup_start_var.get(),
            signup_end=app.sql_signup_end_var.get(),
            audit_start=app.sql_audit_start_var.get(),
            audit_end=app.sql_audit_end_var.get(),
        )
    except Exception as exc:
        app.last_sql_result = None
        app.sql_status_var.set("生成失败")
        app.sql_result_var.set(str(exc))
        app.set_sql_result_text("")
        app.sql_copy_button.configure(state="disabled")
        messagebox.showerror("SQL 生成失败", str(exc))
        return

    app.last_sql_result = result
    app.sql_status_var.set("SQL 已生成")
    app.sql_result_var.set(
        "\n".join(
            [
                "SQL 模板已生成：",
                f"模板文件：{result.source_path}",
                f"考试代码：{result.exam_code}",
                f"考试年月：{result.exam_date}",
                f"报名开始：{result.signup_start}",
                f"报名结束：{result.signup_end}",
                f"审核开始：{result.audit_start}",
                f"审核结束：{result.audit_end}",
            ]
        )
    )
    app.set_sql_result_text(result.sql_content)
    app.sql_copy_button.configure(state="normal")
