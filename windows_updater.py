from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import messagebox, ttk


APP_EXE_NAME = "aliyun_photo_manager_qt.exe"
APP_DIR_NAME = "app"
ProgressCallback = Callable[[str, str, int, int], None]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Launcher and updater for aliyun_photo_manager.")
    parser.add_argument("--apply-update", action="store_true")
    parser.add_argument("--app-dir", default="")
    parser.add_argument("--package", default="")
    parser.add_argument("--wait-pid", type=int, default=0)
    parser.add_argument("--restart-exe", default="")
    return parser.parse_args()


def launcher_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_app_dir() -> Path:
    return launcher_root() / APP_DIR_NAME


def default_app_exe() -> Path:
    return default_app_dir() / APP_EXE_NAME


def launch_app() -> int:
    app_exe = default_app_exe()
    app_dir = app_exe.parent
    if not app_exe.exists():
        messagebox.showerror(
            "启动失败",
            f"未找到主程序：\n{app_exe}",
        )
        return 1
    subprocess.Popen([str(app_exe)], cwd=str(app_dir), close_fds=True)
    return 0


def wait_for_process_exit(pid: int, timeout_seconds: int = 90) -> None:
    if pid <= 0:
        return
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        try:
            import os

            os.kill(pid, 0)
        except OSError:
            return
        time.sleep(0.5)
    raise RuntimeError("等待主程序退出超时。")


def load_manifest(extract_dir: Path) -> dict:
    manifest_path = extract_dir / "update_manifest.json"
    if not manifest_path.exists():
        raise RuntimeError("更新包中缺少 update_manifest.json。")
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def apply_update(app_dir: Path, package_path: Path, progress_callback: ProgressCallback | None = None) -> None:
    extract_dir = Path(tempfile.mkdtemp(prefix="aliyun_photo_manager_patch_"))
    backup_dir = Path(tempfile.mkdtemp(prefix="aliyun_photo_manager_backup_"))
    try:
        if progress_callback:
            progress_callback("正在准备更新包...", "正在解压更新文件", 0, 1)
        with zipfile.ZipFile(package_path) as archive:
            archive.extractall(extract_dir)

        manifest = load_manifest(extract_dir)
        removed = [str(item) for item in manifest.get("removed", [])]
        files = manifest.get("files", [])
        total_steps = max(1, len(removed) + len(files) + 1)
        completed_steps = 0

        def advance(status: str, detail: str) -> None:
            nonlocal completed_steps
            completed_steps += 1
            if progress_callback:
                progress_callback(status, detail, completed_steps, total_steps)

        for relative_path in removed:
            target_path = app_dir / relative_path
            if target_path.exists():
                backup_path = backup_dir / relative_path
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target_path, backup_path)
                if target_path.is_file():
                    target_path.unlink()
            advance("正在移除旧文件...", relative_path)

        for item in files:
            relative_path = str(item.get("path", "")).strip()
            if not relative_path:
                continue
            source_path = extract_dir / "files" / relative_path
            target_path = app_dir / relative_path
            if not source_path.exists():
                raise RuntimeError(f"更新包缺少文件：{relative_path}")
            if target_path.exists():
                backup_path = backup_dir / relative_path
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target_path, backup_path)
            target_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, target_path)
            advance("正在替换文件...", relative_path)

        (app_dir / "installed_manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        advance("正在收尾...", "已写入安装清单")
    except Exception:
        restore_backup(app_dir, backup_dir)
        raise
    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)
        shutil.rmtree(backup_dir, ignore_errors=True)


def restore_backup(app_dir: Path, backup_dir: Path) -> None:
    if not backup_dir.exists():
        return
    for backup_path in sorted(backup_dir.rglob("*")):
        if backup_path.is_dir():
            continue
        relative_path = backup_path.relative_to(backup_dir)
        target_path = app_dir / relative_path
        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(backup_path, target_path)


def restart_app(app_dir: Path, restart_exe: str) -> None:
    exe_name = restart_exe or APP_EXE_NAME
    exe_path = app_dir / exe_name
    if exe_path.exists():
        subprocess.Popen([str(exe_path)], cwd=str(app_dir), close_fds=True)


def run_with_status_window(task: Callable[[ProgressCallback], None], title: str = "在线更新") -> int:
    root = tk.Tk()
    root.title(title)
    root.resizable(False, False)
    root.attributes("-topmost", True)
    root.protocol("WM_DELETE_WINDOW", lambda: None)

    frame = tk.Frame(root, padx=22, pady=18)
    frame.pack(fill="both", expand=True)

    status_var = tk.StringVar(value="正在应用更新，请稍候...")
    detail_var = tk.StringVar(value="更新过程中程序会暂时关闭，请不要重复打开。")
    progress_text_var = tk.StringVar(value="0%")

    tk.Label(frame, textvariable=status_var, font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
    tk.Label(
        frame,
        textvariable=detail_var,
        font=("Microsoft YaHei UI", 9),
        fg="#555555",
        wraplength=360,
        justify="left",
    ).pack(anchor="w", pady=(8, 0))

    progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate", length=360, maximum=100)
    progress.pack(fill="x", pady=(12, 0))
    tk.Label(
        frame,
        textvariable=progress_text_var,
        font=("Microsoft YaHei UI", 9),
        fg="#555555",
    ).pack(anchor="e", pady=(6, 0))

    result: dict[str, str | int | None] = {"code": None, "error": ""}

    def update_progress(status: str, detail: str, current: int, total: int) -> None:
        safe_total = max(1, total)
        percent = max(0, min(100, int(current * 100 / safe_total)))
        status_var.set(status)
        detail_var.set(detail)
        progress.configure(value=percent)
        progress_text_var.set(f"{percent}%")

    def worker() -> None:
        try:
            task(lambda s, d, c, t: root.after(0, update_progress, s, d, c, t))
        except Exception as exc:
            result["code"] = 1
            result["error"] = str(exc)
            root.after(0, lambda: status_var.set("更新失败"))
            root.after(0, lambda: detail_var.set(str(exc) or "更新过程中发生未知错误。"))
            root.after(0, lambda: progress.configure(value=0))
            root.after(0, lambda: progress_text_var.set("0%"))
            root.after(
                0,
                lambda: messagebox.showerror(
                    title,
                    f"更新失败：\n{str(exc) or '未知错误'}",
                    parent=root,
                ),
            )
        else:
            result["code"] = 0
            root.after(0, lambda: status_var.set("更新成功"))
            root.after(0, lambda: detail_var.set("已完成文件替换，程序即将重新打开；如果没有自动打开，请手动启动一次。"))
            root.after(0, lambda: progress.configure(value=100))
            root.after(0, lambda: progress_text_var.set("100%"))
            root.after(
                0,
                lambda: messagebox.showinfo(
                    title,
                    "更新成功，程序即将重新打开。\n如果没有自动打开，请手动启动一次。",
                    parent=root,
                ),
            )
        finally:
            root.after(0, root.destroy)

    threading.Thread(target=worker, daemon=True).start()
    root.mainloop()
    return int(result["code"] or 0)


def run_update(args: argparse.Namespace) -> int:
    app_dir = Path(args.app_dir).resolve() if args.app_dir else default_app_dir().resolve()
    package_path = Path(args.package).resolve()

    def task(progress_callback: ProgressCallback) -> None:
        progress_callback("正在等待程序退出...", f"等待进程 {args.wait_pid} 释放文件", 0, 1)
        wait_for_process_exit(args.wait_pid)
        apply_update(app_dir, package_path, progress_callback=progress_callback)
        progress_callback("正在重新启动程序...", args.restart_exe or APP_EXE_NAME, 1, 1)
        restart_app(app_dir, args.restart_exe)

    return run_with_status_window(task)


def main() -> int:
    args = parse_args()
    if args.apply_update:
        if not args.package:
            messagebox.showerror("更新失败", "缺少更新包路径。")
            return 1
        try:
            return run_update(args)
        except Exception as exc:
            print(f"更新失败：{exc}", file=sys.stderr)
            return 1
    return launch_app()


if __name__ == "__main__":
    raise SystemExit(main())
