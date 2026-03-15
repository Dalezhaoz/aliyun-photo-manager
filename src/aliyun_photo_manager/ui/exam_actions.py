from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

from ..exam_arranger import (
    ExamArrangeOptions,
    ExamArrangeSummary,
    ExamRuleItem,
    ExamTemplateExportSummary,
    export_exam_templates,
    list_headers as list_exam_headers,
    run_exam_arrangement,
)


def choose_exam_candidate_file(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.exam_candidate_var.get()).parent)
        if app.exam_candidate_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
    )
    if selected:
        app.exam_candidate_var.set(selected)
        fill_exam_output_path(app)


def choose_exam_group_file(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.exam_group_var.get()).parent)
        if app.exam_group_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
    )
    if selected:
        app.exam_group_var.set(selected)
        load_exam_group_headers(app)


def choose_exam_plan_file(app) -> None:
    selected = filedialog.askopenfilename(
        initialdir=str(Path(app.exam_plan_var.get()).parent)
        if app.exam_plan_var.get().strip()
        else str(Path.cwd()),
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
    )
    if selected:
        app.exam_plan_var.set(selected)


def export_exam_templates_from_ui(app) -> None:
    output_dir = filedialog.askdirectory(
        initialdir=str(Path(app.exam_candidate_var.get()).parent)
        if app.exam_candidate_var.get().strip()
        else str(Path.cwd())
    )
    if not output_dir:
        return
    try:
        summary = export_exam_templates(Path(output_dir))
    except Exception as exc:
        messagebox.showerror("导出失败", str(exc))
        app.exam_status_var.set("失败")
        app.exam_result_var.set(f"模板导出失败：\n{exc}")
        set_exam_result_text(app, app.exam_result_var.get())
        return

    app.last_exam_template_export = summary
    app.exam_candidate_var.set(str(summary.candidate_template_path))
    app.exam_group_var.set(str(summary.group_template_path))
    app.exam_plan_var.set(str(summary.plan_template_path))
    load_exam_group_headers(app)
    fill_exam_output_path(app)
    app.exam_status_var.set("模板已导出")
    app.exam_result_var.set(
        "\n".join(
            [
                "标准模板已导出：",
                f"考生明细模板：{summary.candidate_template_path.name}",
                f"岗位归组模板：{summary.group_template_path.name}",
                f"编排片段模板：{summary.plan_template_path.name}",
                f"目录：{summary.output_dir}",
                "建议：先补充三张模板，再回来执行考场编排。",
            ]
        )
    )
    set_exam_result_text(app, app.exam_result_var.get())
    app.write_log(f"考场编排标准模板已导出：{summary.output_dir}")


def load_exam_group_headers(app) -> None:
    app.exam_group_headers = []
    group_value = app.exam_group_var.get().strip()
    if not group_value:
        refresh_exam_rule_type_values(app)
        return
    group_path = Path(group_value)
    if not group_path.exists():
        refresh_exam_rule_type_values(app)
        return
    try:
        headers = list_exam_headers(group_path, ["招聘单位", "岗位名称", "科目组"])
    except Exception:
        refresh_exam_rule_type_values(app)
        return
    app.exam_group_headers = [header for header in headers if header not in {"招聘单位", "岗位名称", "科目组"}]
    refresh_exam_rule_type_values(app)


def refresh_exam_rule_type_values(app) -> None:
    values = app.exam_rule_base_types + app.exam_group_headers
    if hasattr(app, "exam_rule_type_combo") and app.exam_rule_type_combo is not None:
        app.exam_rule_type_combo.configure(values=values)
    current_value = app.exam_rule_type_var.get().strip()
    if current_value in values:
        return
    app.exam_rule_type_var.set(values[0] if values else "")


def fill_exam_output_path(app) -> None:
    value = app.exam_candidate_var.get().strip()
    if not value:
        return
    candidate_path = Path(value)
    app.exam_output_var.set(str(candidate_path.with_name(f"{candidate_path.stem}_考场编排结果.xlsx")))


def refresh_exam_rule_tree(app) -> None:
    for item_id in app.exam_rule_tree.get_children():
        app.exam_rule_tree.delete(item_id)
    for index, item in enumerate(app.exam_rule_items, start=1):
        app.exam_rule_tree.insert("", tk.END, values=(index, item.item_type, item.custom_text))


def add_exam_rule_item(app) -> None:
    item_type = app.exam_rule_type_var.get().strip()
    custom_text = app.exam_rule_custom_var.get().strip()
    if not item_type:
        messagebox.showerror("参数错误", "请选择规则项。")
        return
    if item_type == "自定义" and not custom_text:
        messagebox.showerror("参数错误", "自定义规则项不能为空。")
        return
    app.exam_rule_items.append(ExamRuleItem(item_type=item_type, custom_text=custom_text))
    refresh_exam_rule_tree(app)
    if item_type == "自定义":
        app.exam_rule_custom_var.set("")


def remove_exam_rule_item(app) -> None:
    selected = app.exam_rule_tree.selection()
    if not selected:
        return
    indexes = sorted([int(app.exam_rule_tree.item(item_id, "values")[0]) - 1 for item_id in selected], reverse=True)
    for index in indexes:
        if 0 <= index < len(app.exam_rule_items):
            app.exam_rule_items.pop(index)
    refresh_exam_rule_tree(app)


def move_exam_rule_up(app) -> None:
    selected = app.exam_rule_tree.selection()
    if not selected:
        return
    index = int(app.exam_rule_tree.item(selected[0], "values")[0]) - 1
    if index <= 0:
        return
    app.exam_rule_items[index - 1], app.exam_rule_items[index] = (
        app.exam_rule_items[index],
        app.exam_rule_items[index - 1],
    )
    refresh_exam_rule_tree(app)
    app.exam_rule_tree.selection_set(app.exam_rule_tree.get_children()[index - 1])


def move_exam_rule_down(app) -> None:
    selected = app.exam_rule_tree.selection()
    if not selected:
        return
    index = int(app.exam_rule_tree.item(selected[0], "values")[0]) - 1
    if index >= len(app.exam_rule_items) - 1:
        return
    app.exam_rule_items[index + 1], app.exam_rule_items[index] = (
        app.exam_rule_items[index],
        app.exam_rule_items[index + 1],
    )
    refresh_exam_rule_tree(app)
    app.exam_rule_tree.selection_set(app.exam_rule_tree.get_children()[index + 1])


def update_exam_summary_ui(app, summary: Optional[ExamArrangeSummary]) -> None:
    app.last_exam_summary = summary
    if summary is None:
        app.exam_result_var.set("考场编排结果会显示在这里")
        set_exam_result_text(app, app.exam_result_var.get())
        app.exam_open_button.configure(state="disabled")
        return
    app.exam_result_var.set(
        "编排完成：\n"
        f"考生表：{summary.candidate_path}\n"
        f"岗位归组表：{summary.group_path}\n"
        f"编排片段表：{summary.plan_path}\n"
        f"结果文件：{summary.output_path}\n"
        f"总人数：{summary.total_candidates}\n"
        f"成功编排：{summary.arranged_candidates}\n"
        f"未找到岗位归组：{summary.missing_groups}\n"
        f"未找到编排片段：{summary.missing_plan_groups}\n"
        f"重复岗位归组：{summary.duplicate_group_rows}\n"
        f"剩余空座：{summary.unused_plan_slots}"
    )
    set_exam_result_text(app, app.exam_result_var.get())
    app.exam_open_button.configure(state="normal")


def set_exam_result_text(app, content: str) -> None:
    if app.exam_result_text is None:
        return
    app.exam_result_text.configure(state="normal")
    app.exam_result_text.delete("1.0", tk.END)
    app.exam_result_text.insert("1.0", content)
    app.exam_result_text.configure(state="disabled")


def open_exam_result_file(app) -> None:
    if app.last_exam_summary is None:
        return
    app.open_local_file(app.last_exam_summary.output_path, "未找到编排结果文件")


def start_exam_arrange_run(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return

    candidate_value = app.exam_candidate_var.get().strip()
    group_value = app.exam_group_var.get().strip()
    plan_value = app.exam_plan_var.get().strip()
    if not candidate_value or not group_value or not plan_value:
        messagebox.showerror("参数错误", "请先选择考生明细表、岗位归组表和编排片段表。")
        return
    if not app.exam_rule_items:
        messagebox.showerror("参数错误", "请至少添加一条考号规则。")
        return
    try:
        exam_point_digits = int(app.exam_point_digits_var.get().strip())
        room_digits = int(app.exam_room_digits_var.get().strip())
        seat_digits = int(app.exam_seat_digits_var.get().strip())
        serial_digits = int(app.exam_serial_digits_var.get().strip())
    except ValueError:
        messagebox.showerror("参数错误", "考点、考场、座号、流水号位数必须是整数。")
        return

    candidate_path = Path(candidate_value)
    group_path = Path(group_value)
    plan_path = Path(plan_value)
    output_value = app.exam_output_var.get().strip()
    output_path = Path(output_value) if output_value else candidate_path.with_name(f"{candidate_path.stem}_考场编排结果.xlsx")
    app.exam_output_var.set(str(output_path))
    app.save_settings()

    app.exam_run_button.configure(state="disabled")
    app.exam_open_button.configure(state="disabled")
    app.exam_status_var.set("编排中")
    app.exam_result_var.set("正在执行考场编排，请稍候...")
    set_exam_result_text(app, app.exam_result_var.get())
    app.write_log(f"启动考场编排任务：{candidate_path.name}")

    def runner() -> None:
        try:
            summary = run_exam_arrangement(
                ExamArrangeOptions(
                    candidate_path=candidate_path,
                    group_path=group_path,
                    plan_path=plan_path,
                    output_path=output_path,
                    exam_point_digits=exam_point_digits,
                    room_digits=room_digits,
                    seat_digits=seat_digits,
                    serial_digits=serial_digits,
                    sort_mode=app.exam_sort_mode_var.get().strip() or "original",
                    rule_items=list(app.exam_rule_items),
                ),
                logger=app.make_logger(),
            )
        except Exception as exc:
            app.log_queue.put(f"__EXAM_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "exam_summary", "summary": summary})
            app.log_queue.put("__EXAM_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()
