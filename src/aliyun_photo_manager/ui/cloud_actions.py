import json
import socket
import threading
from typing import Dict, List, Optional

from tkinter import messagebox

from ..downloader import BrowserEntry, list_buckets
from ..config import validate_oss_credentials
from ..exam_arranger import ExamRuleItem


def build_credentials(app) -> tuple[str, str, str, str]:
    return validate_oss_credentials(
        app.cloud_type_var.get().strip(),
        app.access_key_id_var.get().strip(),
        app.access_key_secret_var.get().strip(),
        app.endpoint_var.get().strip(),
    )


def validate_cloud_endpoint(app, cloud_type: str, endpoint: str) -> None:
    cleaned = endpoint.strip()
    if cloud_type == "aliyun":
        if "aliyuncs.com" not in cleaned:
            raise ValueError(
                "阿里云 OSS 的 Endpoint 不正确，应类似 `https://oss-cn-hangzhou.aliyuncs.com`。"
            )
        return

    normalized = cleaned
    if normalized.startswith("http://"):
        normalized = normalized[len("http://") :]
    elif normalized.startswith("https://"):
        normalized = normalized[len("https://") :]
    normalized = normalized.split("/")[0]
    if "." in normalized:
        normalized = normalized.split(".", 1)[0]
    if not normalized.startswith("ap-"):
        raise ValueError(
            "腾讯云 COS 的 Endpoint / Region 不正确，应填写 Region（如 `ap-beijing`）或对应 COS Endpoint。"
        )


def check_endpoint_reachable(app, cloud_type: str, endpoint: str) -> None:
    target = endpoint.strip()
    if cloud_type == "tencent" and not target.startswith(("http://", "https://")):
        target = f"https://cos.{target}.myqcloud.com"

    host = target
    if host.startswith("http://"):
        host = host[len("http://") :]
    elif host.startswith("https://"):
        host = host[len("https://") :]
    host = host.split("/", 1)[0]
    if not host:
        raise ValueError("Endpoint / Region 不能为空。")

    try:
        socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
    except socket.gaierror as exc:
        raise ValueError(f"Endpoint 无法解析，请检查地址是否正确：{host}") from exc

    try:
        with socket.create_connection((host, 443), timeout=2):
            return
    except OSError as exc:
        raise ValueError(f"无法连接到 Endpoint：{host}，请检查地址或网络。") from exc


def format_cloud_error(app, error: str) -> str:
    lowered = error.lower()
    if "invalidaccesskeyid" in lowered:
        return "AccessKey ID 不存在或填写错误，请检查后重试。"
    if "secretid is forbidden" in lowered or "invalidsecretid" in lowered:
        return "AccessKey ID 不存在或填写错误，请检查后重试。"
    if "secretkey" in lowered and ("invalid" in lowered or "mismatch" in lowered):
        return "AccessKey Secret 不正确，请检查后重试。"
    if "signaturedoesnotmatch" in lowered or "signature not match" in lowered:
        return "AccessKey Secret 不正确，请检查后重试。"
    if "access denied" in lowered or "'status': 403" in lowered or "status code: 403" in lowered:
        return "当前账号没有访问权限，或密钥填写有误。"
    if "nosuchbucket" in lowered or "bucket does not exist" in lowered:
        return "Bucket 不存在，请检查 bucket 名称是否正确。"
    if "invalidbucketname" in lowered:
        return "Bucket 名称格式不正确。"
    if "connection" in lowered or "timeout" in lowered or "name or service not known" in lowered:
        return "无法连接到云存储服务，请检查 Endpoint / Region 和网络。"
    return error


def set_folder_values(app, folders) -> None:
    values = [""] + list(folders)
    app.folder_values = values
    app.prefix_combo["values"] = values


def set_bucket_values(app, buckets: List[str]) -> None:
    app.bucket_values = buckets
    app.bucket_combo["values"] = buckets
    current_bucket = app.bucket_name_var.get().strip()
    if current_bucket and current_bucket not in buckets:
        app.bucket_name_var.set("")


def set_certificate_bucket_values(app, buckets: List[str]) -> None:
    app.certificate_bucket_values = buckets
    app.certificate_bucket_combo["values"] = buckets
    current_bucket = app.certificate_bucket_name_var.get().strip()
    if current_bucket and current_bucket not in buckets:
        app.certificate_bucket_name_var.set("")


def sync_bucket_values(app, buckets: List[str]) -> None:
    set_bucket_values(app, buckets)
    set_certificate_bucket_values(app, buckets)


def set_certificate_folder_values(app, folders: List[str]) -> None:
    values = [""] + list(folders)
    app.certificate_folder_values = values
    app.certificate_prefix_combo["values"] = values


def default_cloud_profile(app, cloud_type: str) -> Dict[str, str]:
    if cloud_type == "tencent":
        endpoint = "ap-beijing"
    else:
        endpoint = "https://oss-cn-hangzhou.aliyuncs.com"
    return {
        "access_key_id": "",
        "access_key_secret": "",
        "endpoint": endpoint,
        "bucket_name": "",
        "certificate_bucket_name": "",
        "prefix": "",
        "certificate_prefix": "",
    }


def snapshot_current_cloud_profile(app) -> Dict[str, str]:
    return {
        "access_key_id": app.access_key_id_var.get().strip(),
        "access_key_secret": app.access_key_secret_var.get().strip(),
        "endpoint": app.endpoint_var.get().strip(),
        "bucket_name": app.bucket_name_var.get().strip(),
        "certificate_bucket_name": app.certificate_bucket_name_var.get().strip(),
        "prefix": app.prefix_var.get().strip(),
        "certificate_prefix": app.certificate_prefix_var.get().strip(),
    }


def apply_cloud_profile(app, cloud_type: str) -> None:
    profile = app.cloud_profiles.get(cloud_type) or default_cloud_profile(app, cloud_type)
    app.access_key_id_var.set(profile.get("access_key_id", ""))
    app.access_key_secret_var.set(profile.get("access_key_secret", ""))
    app.endpoint_var.set(profile.get("endpoint", default_cloud_profile(app, cloud_type)["endpoint"]))
    app.bucket_name_var.set(profile.get("bucket_name", ""))
    app.certificate_bucket_name_var.set(profile.get("certificate_bucket_name", ""))
    app.prefix_var.set(profile.get("prefix", ""))
    app.certificate_prefix_var.set(profile.get("certificate_prefix", ""))


def load_saved_settings(app) -> None:
    if not app.SETTINGS_FILE.exists():
        return
    try:
        settings = json.loads(app.SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return

    app.prefix_var.set(settings.get("prefix", ""))
    app.photo_template_var.set(settings.get("photo_template", ""))
    app.photo_match_column_var.set(settings.get("photo_match_column", ""))
    app.photo_filter_by_template_var.set(
        settings.get("photo_filter_by_template", app.photo_filter_by_template_var.get())
    )
    app.download_dir_var.set(settings.get("download_dir", app.download_dir_var.get()))
    app.sorted_dir_var.set(settings.get("sorted_dir", app.sorted_dir_var.get()))
    app.cloud_type_var.set(settings.get("cloud_type", app.cloud_type_var.get()))
    saved_profiles = settings.get("cloud_profiles")
    if isinstance(saved_profiles, dict):
        for cloud_type in ("aliyun", "tencent"):
            profile = saved_profiles.get(cloud_type)
            if isinstance(profile, dict):
                merged = default_cloud_profile(app, cloud_type)
                merged.update({key: str(value) for key, value in profile.items() if value is not None})
                app.cloud_profiles[cloud_type] = merged
    else:
        legacy_cloud_type = app.cloud_type_var.get().strip() or "aliyun"
        legacy_profile = default_cloud_profile(app, legacy_cloud_type)
        legacy_profile.update(
            {
                "access_key_id": settings.get("access_key_id", ""),
                "access_key_secret": settings.get("access_key_secret", ""),
                "endpoint": settings.get("endpoint", legacy_profile["endpoint"]),
                "bucket_name": settings.get("bucket_name", ""),
                "certificate_bucket_name": settings.get("certificate_bucket_name", ""),
                "prefix": settings.get("prefix", ""),
                "certificate_prefix": settings.get("certificate_prefix", ""),
            }
        )
        app.cloud_profiles[legacy_cloud_type] = legacy_profile
    apply_cloud_profile(app, app.cloud_type_var.get().strip() or "aliyun")
    app.photo_source_mode_var.set(settings.get("photo_source_mode", app.photo_source_mode_var.get()))
    app.skip_download_var.set(settings.get("skip_download", app.skip_download_var.get()))
    app.flat_var.set(settings.get("flat", app.flat_var.get()))
    app.dry_run_var.set(settings.get("dry_run", app.dry_run_var.get()))
    app.include_duplicates_var.set(settings.get("include_duplicates", app.include_duplicates_var.get()))
    app.move_sorted_files_var.set(settings.get("move_sorted_files", app.move_sorted_files_var.get()))
    app.skip_existing_var.set(settings.get("skip_existing", app.skip_existing_var.get()))
    app.certificate_template_var.set(settings.get("certificate_template", ""))
    app.certificate_source_dir_var.set(
        settings.get("certificate_source_dir", app.certificate_source_dir_var.get())
    )
    app.certificate_output_dir_var.set(
        settings.get("certificate_output_dir", app.certificate_output_dir_var.get())
    )
    app.certificate_match_column_var.set(settings.get("certificate_match_column", ""))
    app.certificate_rename_folder_var.set(
        settings.get("certificate_rename_folder", app.certificate_rename_folder_var.get())
    )
    app.certificate_folder_name_column_var.set(settings.get("certificate_folder_name_column", ""))
    app.certificate_keyword_var.set(settings.get("certificate_keyword", app.certificate_keyword_var.get()))
    app.certificate_classify_var.set(
        settings.get("certificate_classify", app.certificate_classify_var.get())
    )
    app.certificate_dry_run_var.set(
        settings.get("certificate_dry_run", app.certificate_dry_run_var.get())
    )
    app.certificate_mode_var.set(settings.get("certificate_mode", app.certificate_mode_var.get()))
    app.certificate_source_mode_var.set(
        settings.get("certificate_source_mode", app.certificate_source_mode_var.get())
    )
    app.certificate_bucket_name_var.set(settings.get("certificate_bucket_name", ""))
    app.certificate_prefix_var.set(settings.get("certificate_prefix", ""))
    app.word_source_var.set(settings.get("word_source", ""))
    app.pack_source_dir_var.set(settings.get("pack_source_dir", ""))
    app.pack_output_dir_var.set(settings.get("pack_output_dir", app.pack_output_dir_var.get()))
    app.pack_use_custom_password_var.set(
        settings.get("pack_use_custom_password", app.pack_use_custom_password_var.get())
    )
    app.pack_query_var.set(settings.get("pack_query", ""))
    app.match_target_var.set(settings.get("match_target", ""))
    app.match_source_var.set(settings.get("match_source", ""))
    app.match_target_key_var.set(settings.get("match_target_key", ""))
    app.match_source_key_var.set(settings.get("match_source_key", ""))
    app.match_output_var.set(settings.get("match_output", ""))
    app.update_sql_mapping_var.set(settings.get("update_sql_mapping", ""))
    app.update_sql_target_table_var.set(settings.get("update_sql_target_table", ""))
    app.update_sql_source_table_var.set(settings.get("update_sql_source_table", ""))
    app.update_sql_target_key_var.set(settings.get("update_sql_target_key", ""))
    app.update_sql_source_key_var.set(settings.get("update_sql_source_key", ""))
    app.update_sql_ignore_empty_var.set(
        settings.get("update_sql_ignore_empty", app.update_sql_ignore_empty_var.get())
    )
    app.exam_candidate_var.set(settings.get("exam_candidate", ""))
    app.exam_group_var.set(settings.get("exam_group", ""))
    app.exam_plan_var.set(settings.get("exam_plan", ""))
    app.exam_output_var.set(settings.get("exam_output", ""))
    app.exam_point_digits_var.set(settings.get("exam_point_digits", app.exam_point_digits_var.get()))
    app.exam_room_digits_var.set(settings.get("exam_room_digits", app.exam_room_digits_var.get()))
    app.exam_seat_digits_var.set(settings.get("exam_seat_digits", app.exam_seat_digits_var.get()))
    app.exam_serial_digits_var.set(settings.get("exam_serial_digits", app.exam_serial_digits_var.get()))
    app.exam_sort_mode_var.set(settings.get("exam_sort_mode", app.exam_sort_mode_var.get()))
    app.exam_rule_items = [
        ExamRuleItem(item_type=str(item.get("item_type", "")), custom_text=str(item.get("custom_text", "")))
        for item in settings.get("exam_rule_items", [])
        if isinstance(item, dict) and str(item.get("item_type", "")).strip()
    ]
    app.load_exam_group_headers()


def save_settings(app) -> None:
    settings = {
        "prefix": app.prefix_var.get().strip(),
        "photo_template": app.photo_template_var.get().strip(),
        "photo_match_column": app.photo_match_column_var.get().strip(),
        "photo_filter_by_template": app.photo_filter_by_template_var.get(),
        "download_dir": app.download_dir_var.get().strip(),
        "sorted_dir": app.sorted_dir_var.get().strip(),
        "cloud_type": app.cloud_type_var.get().strip(),
        "cloud_profiles": {
            **app.cloud_profiles,
            app.cloud_type_var.get().strip() or "aliyun": snapshot_current_cloud_profile(app),
        },
        "photo_source_mode": app.photo_source_mode_var.get().strip(),
        "skip_download": app.skip_download_var.get(),
        "flat": app.flat_var.get(),
        "dry_run": app.dry_run_var.get(),
        "include_duplicates": app.include_duplicates_var.get(),
        "move_sorted_files": app.move_sorted_files_var.get(),
        "skip_existing": app.skip_existing_var.get(),
        "certificate_template": app.certificate_template_var.get().strip(),
        "certificate_source_dir": app.certificate_source_dir_var.get().strip(),
        "certificate_output_dir": app.certificate_output_dir_var.get().strip(),
        "certificate_match_column": app.certificate_match_column_var.get().strip(),
        "certificate_rename_folder": app.certificate_rename_folder_var.get(),
        "certificate_folder_name_column": app.certificate_folder_name_column_var.get().strip(),
        "certificate_keyword": app.certificate_keyword_var.get().strip(),
        "certificate_classify": app.certificate_classify_var.get(),
        "certificate_dry_run": app.certificate_dry_run_var.get(),
        "certificate_mode": app.certificate_mode_var.get().strip(),
        "certificate_source_mode": app.certificate_source_mode_var.get().strip(),
        "certificate_bucket_name": app.certificate_bucket_name_var.get().strip(),
        "certificate_prefix": app.certificate_prefix_var.get().strip(),
        "word_source": app.word_source_var.get().strip(),
        "pack_source_dir": app.pack_source_dir_var.get().strip(),
        "pack_output_dir": app.pack_output_dir_var.get().strip(),
        "pack_use_custom_password": app.pack_use_custom_password_var.get(),
        "pack_query": app.pack_query_var.get().strip(),
        "match_target": app.match_target_var.get().strip(),
        "match_source": app.match_source_var.get().strip(),
        "match_target_key": app.match_target_key_var.get().strip(),
        "match_source_key": app.match_source_key_var.get().strip(),
        "match_output": app.match_output_var.get().strip(),
        "update_sql_mapping": app.update_sql_mapping_var.get().strip(),
        "update_sql_target_table": app.update_sql_target_table_var.get().strip(),
        "update_sql_source_table": app.update_sql_source_table_var.get().strip(),
        "update_sql_target_key": app.update_sql_target_key_var.get().strip(),
        "update_sql_source_key": app.update_sql_source_key_var.get().strip(),
        "update_sql_ignore_empty": app.update_sql_ignore_empty_var.get(),
        "exam_candidate": app.exam_candidate_var.get().strip(),
        "exam_group": app.exam_group_var.get().strip(),
        "exam_plan": app.exam_plan_var.get().strip(),
        "exam_output": app.exam_output_var.get().strip(),
        "exam_point_digits": app.exam_point_digits_var.get().strip(),
        "exam_room_digits": app.exam_room_digits_var.get().strip(),
        "exam_seat_digits": app.exam_seat_digits_var.get().strip(),
        "exam_serial_digits": app.exam_serial_digits_var.get().strip(),
        "exam_sort_mode": app.exam_sort_mode_var.get().strip(),
        "exam_rule_items": [
            {"item_type": item.item_type, "custom_text": item.custom_text}
            for item in app.exam_rule_items
        ],
    }
    app.SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
    app.write_log(f"配置已保存到 {app.SETTINGS_FILE}")


def update_certificate_mode_ui(app) -> None:
    if app.certificate_mode_var.get() == "keyword":
        app.certificate_keyword_entry.configure(state="normal")
    else:
        app.certificate_keyword_entry.configure(state="disabled")
    if app.certificate_rename_folder_var.get():
        app.certificate_folder_name_combo.configure(state="readonly")
    else:
        app.certificate_folder_name_combo.configure(state="disabled")


def on_cloud_type_changed(app) -> None:
    previous_type = "tencent" if app.cloud_type_var.get() == "aliyun" else "aliyun"
    app.cloud_profiles[previous_type] = snapshot_current_cloud_profile(app)
    apply_cloud_profile(app, app.cloud_type_var.get().strip())
    sync_bucket_values(app, [])
    set_folder_values(app, [])
    set_certificate_folder_values(app, [])
    app.bucket_status_var.set("切换云类型后，请重新加载 bucket 列表")
    app.certificate_bucket_status_var.set("切换云类型后，请重新加载 bucket 列表")
    app.folder_status_var.set("未加载 bucket 文件夹")
    app.certificate_folder_status_var.set("未加载 bucket 文件夹")


def load_buckets(app) -> None:
    try:
        cloud_type, access_key_id, access_key_secret, endpoint = build_credentials(app)
        validate_cloud_endpoint(app, cloud_type, endpoint)
        check_endpoint_reachable(app, cloud_type, endpoint)
    except Exception as exc:
        messagebox.showerror("参数错误", str(exc))
        return

    app.bucket_status_var.set("正在加载 bucket 列表...")
    app.write_log("开始加载 bucket 列表。")
    app.bucket_load_token += 1
    current_token = app.bucket_load_token
    app.root.after(8000, lambda: app.handle_bucket_load_timeout(current_token))

    def worker() -> None:
        try:
            buckets = list_buckets(access_key_id, access_key_secret, endpoint, cloud_type=cloud_type)
        except Exception as exc:
            app.root.after(
                0,
                lambda: app.finish_bucket_load(
                    error=f"{type(exc).__name__}: {exc}",
                    token=current_token,
                ),
            )
            return
        app.root.after(0, lambda: app.finish_bucket_load(buckets=buckets, token=current_token))

    threading.Thread(target=worker, daemon=True).start()


def finish_bucket_load(
    app,
    buckets: Optional[List[str]] = None,
    error: Optional[str] = None,
    token: Optional[int] = None,
) -> None:
    if token is not None and token != app.bucket_load_token:
        return
    if error is not None:
        app.bucket_status_var.set("加载失败")
        app.write_log(f"加载 bucket 失败：{error}")
        messagebox.showerror("加载 bucket 失败", format_cloud_error(app, error))
        return

    buckets = buckets or []
    sync_bucket_values(app, buckets)
    if buckets:
        if not app.bucket_name_var.get().strip():
            app.bucket_name_var.set(buckets[0])
        if not app.certificate_bucket_name_var.get().strip():
            app.certificate_bucket_name_var.set(buckets[0])
        app.bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
        app.certificate_bucket_status_var.set(f"已加载 {len(buckets)} 个 bucket")
        app.write_log(f"找到 {len(buckets)} 个 bucket。")
        app.prefix_var.set("")
        app.folder_status_var.set("请选择 bucket 后点击“加载当前层级”")
    else:
        app.bucket_status_var.set("未找到可用 bucket")
        app.certificate_bucket_status_var.set("未找到可用 bucket")
        app.write_log("当前凭证下未找到可用 bucket。")


def handle_bucket_load_timeout(app, token: int) -> None:
    if token != app.bucket_load_token:
        return
    if app.bucket_status_var.get() != "正在加载 bucket 列表...":
        return
    app.bucket_status_var.set("加载失败")
    timeout_message = "加载 bucket 超时，请检查 Endpoint、密钥或网络后重试。"
    app.write_log(timeout_message)
    messagebox.showerror("加载 bucket 失败", timeout_message)
    app.bucket_load_token += 1


def go_to_parent_prefix(app) -> None:
    current = app.prefix_var.get().strip().strip("/")
    if not current:
        app.prefix_var.set("")
        return
    parts = current.split("/")
    parent = "/".join(parts[:-1])
    app.prefix_var.set(parent + "/" if parent else "")
    app.load_bucket_folders()
