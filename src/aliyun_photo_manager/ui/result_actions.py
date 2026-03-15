import os
import subprocess
import sys
from pathlib import Path

from tkinter import messagebox


def open_local_file(app, file_path: Path, not_found_title: str) -> None:
    if not file_path.exists():
        messagebox.showerror("文件不存在", f"{not_found_title}：{file_path}")
        return
    try:
        if os.name == "nt":
            os.startfile(str(file_path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(file_path)])
        else:
            subprocess.Popen(["xdg-open", str(file_path)])
    except Exception as exc:
        messagebox.showerror("打开失败", str(exc))


def open_certificate_report_file(app) -> None:
    if app.last_certificate_summary is None or app.last_certificate_summary.report_path is None:
        return
    app.open_local_file(app.last_certificate_summary.report_path, "未找到结果清单")
