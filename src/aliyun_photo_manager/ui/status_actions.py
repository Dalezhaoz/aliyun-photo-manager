from __future__ import annotations

import json
import subprocess
import threading
import sys
import tempfile
from dataclasses import asdict
from pathlib import Path
from tkinter import filedialog, messagebox

from ..project_stage_report import (
    ProjectStageSummary,
    StageServerConfig,
    dump_status_query_payload,
    export_project_stages,
    query_project_stages,
    summary_from_dict,
)


def _get_selected_server_index(app) -> int | None:
    selected = app.status_server_tree.selection()
    if not selected:
        return None
    try:
        return int(selected[0])
    except ValueError:
        return None


def refresh_status_server_tree(app) -> None:
    if app.status_server_tree is None:
        return
    for item_id in app.status_server_tree.get_children():
        app.status_server_tree.delete(item_id)
    for index, config in enumerate(app.status_server_configs):
        app.status_server_tree.insert(
            "",
            "end",
            iid=str(index),
            values=(config.name, f"{config.host}:{config.port}", "是" if config.enabled else "否"),
        )


def clear_status_server_form(app) -> None:
    app.status_server_selected_index = None
    app.status_server_name_var.set("")
    app.status_server_host_var.set("")
    app.status_server_port_var.set("1433")
    app.status_server_user_var.set("")
    app.status_server_password_var.set("")
    app.status_server_enabled_var.set(True)
    if app.status_server_tree is not None:
        for item_id in app.status_server_tree.selection():
            app.status_server_tree.selection_remove(item_id)


def on_status_server_select(app, _event=None) -> None:
    index = _get_selected_server_index(app)
    if index is None:
        return
    app.status_server_selected_index = index
    config = app.status_server_configs[index]
    app.status_server_name_var.set(config.name)
    app.status_server_host_var.set(config.host)
    app.status_server_port_var.set(str(config.port))
    app.status_server_user_var.set(config.username)
    app.status_server_password_var.set(config.password)
    app.status_server_enabled_var.set(config.enabled)


def save_status_server(app) -> None:
    name = app.status_server_name_var.get().strip()
    host = app.status_server_host_var.get().strip()
    port_text = app.status_server_port_var.get().strip() or "1433"
    username = app.status_server_user_var.get().strip()
    password = app.status_server_password_var.get().strip()
    if not name or not host or not username or not password:
        messagebox.showerror("参数错误", "请完整填写服务器名称、地址、用户名和密码。")
        return
    try:
        port = int(port_text)
    except ValueError:
        messagebox.showerror("参数错误", "端口必须是数字。")
        return
    config = StageServerConfig(
        name=name,
        host=host,
        port=port,
        username=username,
        password=password,
        enabled=app.status_server_enabled_var.get(),
    )
    index = app.status_server_selected_index
    if index is None:
        app.status_server_configs.append(config)
    else:
        app.status_server_configs[index] = config
    refresh_status_server_tree(app)
    app.save_settings()
    clear_status_server_form(app)


def delete_status_server(app) -> None:
    index = _get_selected_server_index(app)
    if index is None:
        return
    del app.status_server_configs[index]
    refresh_status_server_tree(app)
    app.save_settings()
    clear_status_server_form(app)


def test_status_server(app) -> None:
    name = app.status_server_name_var.get().strip()
    host = app.status_server_host_var.get().strip()
    port_text = app.status_server_port_var.get().strip() or "1433"
    username = app.status_server_user_var.get().strip()
    password = app.status_server_password_var.get().strip()
    if not name or not host or not username or not password:
        messagebox.showerror("参数错误", "请先填写完整服务器信息。")
        return
    try:
        port = int(port_text)
    except ValueError:
        messagebox.showerror("参数错误", "端口必须是数字。")
        return
    try:
        summary = query_project_stages(
            [StageServerConfig(name=name, host=host, port=port, username=username, password=password)],
            status_filter="全部",
        )
    except Exception as exc:
        messagebox.showerror("连接失败", str(exc))
        return
    messagebox.showinfo(
        "连接成功",
        f"服务器可连接。\n共检查数据库 {summary.visited_databases} 个，匹配业务库 {summary.matched_databases} 个。",
    )


def update_status_summary_ui(app, summary: ProjectStageSummary | None) -> None:
    app.last_status_summary = summary
    if app.status_result_tree is None:
        return
    for item_id in app.status_result_tree.get_children():
        app.status_result_tree.delete(item_id)
    if summary is None:
        app.status_query_status_var.set("未开始查询")
        app.status_export_button.configure(state="disabled")
        return
    for item in summary.records:
        app.status_result_tree.insert(
            "",
            "end",
            values=(
                item.server_name,
                item.database_name,
                item.project_name,
                item.stage_name,
                item.start_time.strftime("%Y-%m-%d %H:%M:%S"),
                item.end_time.strftime("%Y-%m-%d %H:%M:%S"),
                item.status,
            ),
        )
    app.status_query_status_var.set(
        f"查询完成：启用服务器 {summary.enabled_servers} 台，遍历数据库 {summary.visited_databases} 个，"
        f"匹配业务库 {summary.matched_databases} 个，正在进行 {summary.ongoing_count} 条，即将开始 {summary.upcoming_count} 条。"
    )
    app.status_export_button.configure(state="normal" if summary.records else "disabled")


def start_status_query(app) -> None:
    if app.worker is not None and app.worker.is_alive():
        messagebox.showinfo("任务执行中", "当前任务还没结束。")
        return
    if not app.status_server_configs:
        messagebox.showerror("参数错误", "请先添加至少一台数据库服务器。")
        return
    app.save_settings()
    app.status_run_button.configure(state="disabled")
    app.status_export_button.configure(state="disabled")
    app.status_query_status_var.set("正在查询项目阶段，请稍候...")
    app.write_log("开始执行项目阶段汇总查询。")

    def runner() -> None:
        try:
            with tempfile.TemporaryDirectory(prefix="project_stage_") as temp_dir:
                temp_path = Path(temp_dir)
                input_path = temp_path / "query.json"
                output_path = temp_path / "result.json"
                dump_status_query_payload(
                    servers=app.status_server_configs,
                    status_filter=app.status_filter_var.get().strip() or "正在进行 + 即将开始",
                    stage_keyword=app.status_stage_keyword_var.get().strip(),
                    project_keyword=app.status_project_keyword_var.get().strip(),
                    output_path=input_path,
                )
                runner_path = Path(__file__).resolve().parents[1] / "project_stage_runner.py"
                result = subprocess.run(
                    [
                        sys.executable,
                        str(runner_path),
                        str(input_path),
                        str(output_path),
                    ],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                if result.returncode != 0:
                    error_text = (result.stderr or result.stdout or "项目阶段汇总查询失败").strip()
                    raise RuntimeError(error_text)
                summary = summary_from_dict(json.loads(output_path.read_text(encoding="utf-8")))
        except Exception as exc:
            app.log_queue.put(f"__STATUS_FAILED__::{type(exc).__name__}: {exc}")
        else:
            app.log_queue.put({"type": "status_summary", "summary": summary})
            app.log_queue.put("__STATUS_DONE__")

    app.worker = threading.Thread(target=runner, daemon=True)
    app.worker.start()


def export_status_result(app) -> None:
    summary = app.last_status_summary
    if summary is None or not summary.records:
        return
    selected = filedialog.asksaveasfilename(
        title="导出项目阶段汇总",
        defaultextension=".xlsx",
        initialfile="项目阶段汇总.xlsx",
        filetypes=[("Excel 文件", "*.xlsx")],
    )
    if not selected:
        return
    output_path = Path(selected)
    try:
        export_project_stages(summary, output_path)
    except Exception as exc:
        messagebox.showerror("导出失败", str(exc))
        return
    app.open_local_file(output_path, "未找到导出的项目阶段汇总文件")


def load_status_settings(app, settings: dict) -> None:
    raw_servers = settings.get("status_servers", [])
    servers = []
    if isinstance(raw_servers, list):
        for item in raw_servers:
            if not isinstance(item, dict):
                continue
            try:
                servers.append(
                    StageServerConfig(
                        name=str(item.get("name", "")).strip(),
                        host=str(item.get("host", "")).strip(),
                        port=int(item.get("port", 1433)),
                        username=str(item.get("username", "")).strip(),
                        password=str(item.get("password", "")).strip(),
                        enabled=bool(item.get("enabled", True)),
                    )
                )
            except Exception:
                continue
    app.status_server_configs = [item for item in servers if item.name and item.host]
    app.status_filter_var.set(settings.get("status_filter", app.status_filter_var.get()))
    app.status_stage_keyword_var.set(settings.get("status_stage_keyword", ""))
    app.status_project_keyword_var.set(settings.get("status_project_keyword", ""))


def dump_status_settings(app) -> dict:
    return {
        "status_servers": [asdict(item) for item in app.status_server_configs],
        "status_filter": app.status_filter_var.get().strip(),
        "status_stage_keyword": app.status_stage_keyword_var.get().strip(),
        "status_project_keyword": app.status_project_keyword_var.get().strip(),
    }
