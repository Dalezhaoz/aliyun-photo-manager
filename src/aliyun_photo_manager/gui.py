import queue
import threading
import tkinter as tk
import json
import os
import socket
import subprocess
import sys
import tempfile
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from .app import (
    RunOptions,
    WorkflowSummary,
    build_prefixed_directory,
    resolve_photo_directories,
    run_photo_classification_only,
    run_photo_download_and_template,
)
from . import __version__
from .certificate_filter import (
    CertificateFilterOptions,
    CertificateFilterSummary,
    list_template_headers,
    load_match_values,
    run_certificate_filter,
)
from .config import OssConfig, validate_oss_config, validate_oss_credentials
from .downloader import (
    BrowserEntry,
    count_photos_in_prefix,
    download_objects,
    is_photo_key,
    list_buckets,
    list_browser_entries,
    list_folder_prefixes,
)
from .excel_classifier import generate_template
from .result_packer import PackSummary, pack_encrypted_folder, query_pack_history
from .update_sql_generator import UpdateSqlResult
from .data_matcher import (
    ColumnMapping,
    DataMatchOptions,
    DataMatchSummary,
    list_headers as list_match_headers,
    run_data_match,
)
from .word_to_html import WordExportResult, export_word_to_html
from .ui import (
    add_entry_row as ui_add_entry_row,
    add_file_row as ui_add_file_row,
    add_path_row as ui_add_path_row,
    add_tick_checkbutton as ui_add_tick_checkbutton,
    apply_cloud_profile as ui_apply_cloud_profile,
    bind_mousewheel_to_canvas as ui_bind_mousewheel_to_canvas,
    build_credentials as ui_build_credentials,
    check_endpoint_reachable as ui_check_endpoint_reachable,
    choose_certificate_template as ui_choose_certificate_template,
    choose_match_source as ui_choose_match_source,
    choose_match_target as ui_choose_match_target,
    choose_pack_source_directory as ui_choose_pack_source_directory,
    choose_pack_source_file as ui_choose_pack_source_file,
    choose_phone_filter_file as ui_choose_phone_filter_file,
    choose_photo_template as ui_choose_photo_template,
    choose_update_sql_mapping as ui_choose_update_sql_mapping,
    choose_word_source as ui_choose_word_source,
    copy_update_sql as ui_copy_update_sql,
    copy_word_html as ui_copy_word_html,
    export_update_sql_template_file as ui_export_update_sql_template_file,
    fill_match_output_path as ui_fill_match_output_path,
    fill_phone_table_name as ui_fill_phone_table_name,
    finish_count_refresh as ui_finish_count_refresh,
    finish_certificate_bucket_load as ui_finish_certificate_bucket_load,
    finish_certificate_folder_load as ui_finish_certificate_folder_load,
    finish_certificate_search as ui_finish_certificate_search,
    finish_folder_load as ui_finish_folder_load,
    finish_search as ui_finish_search,
    filter_folder_entries as ui_filter_folder_entries,
    go_to_certificate_parent_prefix as ui_go_to_certificate_parent_prefix,
    handle_certificate_bucket_load_timeout as ui_handle_certificate_bucket_load_timeout,
    load_certificate_buckets as ui_load_certificate_buckets,
    load_certificate_folders as ui_load_certificate_folders,
    load_certificate_headers as ui_load_certificate_headers,
    load_match_headers as ui_load_match_headers,
    load_photo_headers as ui_load_photo_headers,
    load_update_sql_headers as ui_load_update_sql_headers,
    load_bucket_folders as ui_load_bucket_folders,
    add_extra_match_mapping as ui_add_extra_match_mapping,
    open_photo_report_file as ui_open_photo_report_file,
    open_template_file as ui_open_template_file,
    remove_extra_match_mapping as ui_remove_extra_match_mapping,
    on_certificate_search_double_click as ui_on_certificate_search_double_click,
    on_certificate_tree_double_click as ui_on_certificate_tree_double_click,
    on_certificate_tree_select as ui_on_certificate_tree_select,
    on_search_double_click as ui_on_search_double_click,
    on_tree_double_click as ui_on_tree_double_click,
    on_tree_select as ui_on_tree_select,
    refresh_selected_folder_count as ui_refresh_selected_folder_count,
    render_certificate_folder_tree as ui_render_certificate_folder_tree,
    render_certificate_search_tree as ui_render_certificate_search_tree,
    render_folder_tree as ui_render_folder_tree,
    add_transfer_mapping as ui_add_transfer_mapping,
    remove_transfer_mapping as ui_remove_transfer_mapping,
    render_word_preview as ui_render_word_preview,
    search_certificate_files as ui_search_certificate_files,
    search_bucket_files as ui_search_bucket_files,
    set_certificate_headers as ui_set_certificate_headers,
    set_photo_headers as ui_set_photo_headers,
    set_word_code as ui_set_word_code,
    update_certificate_summary_ui as ui_update_certificate_summary_ui,
    update_match_summary_ui as ui_update_match_summary_ui,
    update_photo_source_mode_ui as ui_update_photo_source_mode_ui,
    update_summary_ui as ui_update_summary_ui,
    update_word_export_ui as ui_update_word_export_ui,
    set_id_result_text as ui_set_id_result_text,
    set_match_result_text as ui_set_match_result_text,
    set_phone_result_text as ui_set_phone_result_text,
    open_match_result_file as ui_open_match_result_file,
    open_word_preview_in_browser as ui_open_word_preview_in_browser,
    open_certificate_report_file as ui_open_certificate_report_file,
    open_local_file as ui_open_local_file,
    open_pack_file as ui_open_pack_file,
    copy_pack_password as ui_copy_pack_password,
    copy_generated_id_card as ui_copy_generated_id_card,
    run_pack_history_query as ui_run_pack_history_query,
    set_pack_query_result_text as ui_set_pack_query_result_text,
    set_update_sql_result_text as ui_set_update_sql_result_text,
    start_certificate_download_run as ui_start_certificate_download_run,
    start_certificate_run as ui_start_certificate_run,
    start_match_run as ui_start_match_run,
    start_pack_run as ui_start_pack_run,
    run_id_card_generate as ui_run_id_card_generate,
    run_id_card_validate as ui_run_id_card_validate,
    start_phone_decrypt_run as ui_start_phone_decrypt_run,
    start_photo_classify_run as ui_start_photo_classify_run,
    start_photo_download_run as ui_start_photo_download_run,
    start_update_sql_render as ui_start_update_sql_render,
    start_word_export as ui_start_word_export,
    update_update_sql_ui as ui_update_update_sql_ui,
    update_pack_password_mode_ui as ui_update_pack_password_mode_ui,
    update_pack_summary_ui as ui_update_pack_summary_ui,
    update_id_city_values as ui_update_id_city_values,
    update_id_county_values as ui_update_id_county_values,
    update_id_day_values as ui_update_id_day_values,
    update_id_region_hint as ui_update_id_region_hint,
    update_phone_mode_ui as ui_update_phone_mode_ui,
    update_phone_summary_ui as ui_update_phone_summary_ui,
    cancel_run as ui_cancel_run,
    clear_log as ui_clear_log,
    build_certificate_tab,
    build_help_tab,
    build_home_tab,
    build_home_settings_tab,
    build_id_card_tab,
    build_log_tab,
    build_match_tab,
    build_pack_tab,
    build_phone_tab,
    build_photo_tab,
    build_template_tab,
    build_update_sql_tab,
    create_text_entry as ui_create_text_entry,
    default_cloud_profile as ui_default_cloud_profile,
    finish_bucket_load as ui_finish_bucket_load,
    flush_logs as ui_flush_logs,
    format_cloud_error as ui_format_cloud_error,
    go_to_parent_prefix as ui_go_to_parent_prefix,
    handle_bucket_load_timeout as ui_handle_bucket_load_timeout,
    load_buckets as ui_load_buckets,
    load_saved_settings as ui_load_saved_settings,
    make_logger as ui_make_logger,
    make_progress_callback as ui_make_progress_callback,
    on_cloud_type_changed as ui_on_cloud_type_changed,
    save_settings as ui_save_settings,
    set_bucket_values as ui_set_bucket_values,
    set_certificate_bucket_values as ui_set_certificate_bucket_values,
    set_certificate_folder_values as ui_set_certificate_folder_values,
    set_folder_values as ui_set_folder_values,
    snapshot_current_cloud_profile as ui_snapshot_current_cloud_profile,
    update_progress_ui as ui_update_progress_ui,
    sync_bucket_values as ui_sync_bucket_values,
    update_certificate_mode_ui as ui_update_certificate_mode_ui,
    validate_cloud_endpoint as ui_validate_cloud_endpoint,
    write_log as ui_write_log,
)

try:
    from .sql_template_executor import SqlTemplateResult
except Exception:
    class SqlTemplateResult:
        pass

try:
    from .exam_arranger import ExamArrangeSummary, ExamRuleItem, ExamTemplateExportSummary
except Exception:
    class ExamArrangeSummary:
        pass

    class ExamRuleItem:
        def __init__(self, item_type: str = "", custom_text: str = "") -> None:
            self.item_type = item_type
            self.custom_text = custom_text

    class ExamTemplateExportSummary:
        pass

try:
    from .project_stage_report import ProjectStageSummary, StageServerConfig
except Exception:
    class ProjectStageSummary:
        pass

    class StageServerConfig:
        pass


class App:
    SETTINGS_FILE = Path(__file__).resolve().parents[2] / ".gui_settings.json"
    HOME_SHORTCUT_DEFAULTS = [
        "照片下载与分类",
        "证件资料筛选",
        "电话解密",
        "更新SQL生成",
        "身份证工具",
        "运行日志",
    ]
    HOME_SHORTCUT_OPTIONS = [
        "照片下载与分类",
        "证件资料筛选",
        "表样转换",
        "结果打包",
        "数据匹配",
        "更新SQL生成",
        "电话解密",
        "身份证工具",
        "使用说明",
        "运行日志",
    ]
    NAV_GROUPS = [
        (
            "开始使用",
            [
                ("首页", "首页"),
                ("首页设置", "首页设置"),
            ],
        ),
        (
            "文件处理",
            [
                ("照片下载与分类", "照片下载与分类"),
                ("证件资料筛选", "证件资料筛选"),
                ("表样转换", "表样转换"),
                ("结果打包", "结果打包"),
            ],
        ),
        (
            "数据处理",
            [
                ("数据匹配", "数据匹配"),
            ],
        ),
        (
            "数据库工具",
            [
                ("更新 SQL 生成", "更新SQL生成"),
                ("电话解密", "电话解密"),
            ],
        ),
        (
            "查询与辅助",
            [
                ("身份证工具", "身份证工具"),
            ],
        ),
        (
            "说明与日志",
            [
                ("使用说明", "使用说明"),
                ("运行日志", "运行日志"),
            ],
        ),
    ]

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"报名系统工具箱 v{__version__}")
        self.root.geometry("1080x760")
        self.root.minsize(760, 560)

        self.log_queue: "queue.Queue[object]" = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.cancel_event = threading.Event()
        self.tab_ids_by_text: Dict[str, str] = {}
        self.nav_item_buttons: Dict[str, tk.Button] = {}
        self.nav_group_frames: Dict[str, tk.Frame] = {}
        self.nav_group_indicators: Dict[str, ttk.Label] = {}
        self.nav_group_expanded: Dict[str, bool] = {}
        self.nav_selected_text = ""

        self.prefix_var = tk.StringVar()
        self.photo_source_mode_var = tk.StringVar(value="oss")
        self.photo_template_var = tk.StringVar()
        self.photo_match_column_var = tk.StringVar()
        self.photo_filter_by_template_var = tk.BooleanVar(value=False)
        self.search_keyword_var = tk.StringVar()
        self.bucket_values = []
        self.folder_values = [""]
        self.download_dir_var = tk.StringVar(value=str(Path.cwd() / "downloads"))
        self.sorted_dir_var = tk.StringVar(value=str(Path.cwd() / "sorted"))
        self.cloud_type_var = tk.StringVar(value="aliyun")
        self.access_key_id_var = tk.StringVar()
        self.access_key_secret_var = tk.StringVar()
        self.endpoint_var = tk.StringVar(value="https://oss-cn-hangzhou.aliyuncs.com")
        self.bucket_name_var = tk.StringVar()

        self.skip_download_var = tk.BooleanVar(value=False)
        self.flat_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=True)
        self.include_duplicates_var = tk.BooleanVar(value=False)
        self.move_sorted_files_var = tk.BooleanVar(value=False)
        self.skip_existing_var = tk.BooleanVar(value=True)
        self.certificate_template_var = tk.StringVar()
        self.certificate_source_mode_var = tk.StringVar(value="local")
        self.certificate_source_dir_var = tk.StringVar()
        self.certificate_output_dir_var = tk.StringVar()
        self.certificate_match_column_var = tk.StringVar()
        self.certificate_rename_folder_var = tk.BooleanVar(value=False)
        self.certificate_folder_name_column_var = tk.StringVar()
        self.certificate_keyword_var = tk.StringVar(value="学历证书")
        self.certificate_classify_var = tk.BooleanVar(value=True)
        self.certificate_dry_run_var = tk.BooleanVar(value=False)
        self.certificate_mode_var = tk.StringVar(value="folder")
        self.certificate_bucket_name_var = tk.StringVar()
        self.certificate_prefix_var = tk.StringVar()
        self.word_source_var = tk.StringVar()
        self.pack_source_dir_var = tk.StringVar()
        self.pack_output_dir_var = tk.StringVar(value=str(Path.cwd() / "packed_results"))
        self.pack_use_custom_password_var = tk.BooleanVar(value=False)
        self.pack_password_var = tk.StringVar()
        self.pack_query_var = tk.StringVar()
        self.sql_template_path_var = tk.StringVar()
        self.sql_exam_code_var = tk.StringVar()
        self.sql_exam_date_var = tk.StringVar()
        self.sql_signup_start_var = tk.StringVar()
        self.sql_signup_end_var = tk.StringVar()
        self.sql_audit_start_var = tk.StringVar()
        self.sql_audit_end_var = tk.StringVar()
        self.match_target_var = tk.StringVar()
        self.match_source_var = tk.StringVar()
        self.match_target_key_var = tk.StringVar()
        self.match_source_key_var = tk.StringVar()
        self.match_output_var = tk.StringVar()
        self.match_extra_target_var = tk.StringVar()
        self.match_extra_source_var = tk.StringVar()
        self.match_transfer_target_var = tk.StringVar()
        self.match_transfer_source_var = tk.StringVar()
        self.update_sql_mapping_var = tk.StringVar()
        self.update_sql_target_table_var = tk.StringVar()
        self.update_sql_source_table_var = tk.StringVar()
        self.update_sql_target_key_var = tk.StringVar()
        self.update_sql_source_key_var = tk.StringVar()
        self.update_sql_ignore_empty_var = tk.BooleanVar(value=True)
        self.phone_server_var = tk.StringVar()
        self.phone_port_var = tk.StringVar(value="1433")
        self.phone_username_var = tk.StringVar()
        self.phone_password_var = tk.StringVar()
        self.phone_signup_database_var = tk.StringVar()
        self.phone_info_database_var = tk.StringVar()
        self.phone_exam_sort_var = tk.StringVar()
        self.phone_candidate_table_var = tk.StringVar()
        self.phone_mode_var = tk.StringVar(value="all")
        self.phone_filter_file_var = tk.StringVar()
        self.id_input_var = tk.StringVar()
        self.id_province_var = tk.StringVar(value="北京市")
        self.id_city_var = tk.StringVar(value="北京市")
        self.id_county_var = tk.StringVar(value="东城区")
        self.id_custom_region_code_var = tk.StringVar()
        self.id_birth_year_var = tk.StringVar(value="1990")
        self.id_birth_month_var = tk.StringVar(value="01")
        self.id_birth_day_var = tk.StringVar(value="01")
        self.id_gender_var = tk.StringVar(value="男")
        self.id_generated_var = tk.StringVar()
        self.id_region_hint_var = tk.StringVar(value="北京市 东城区（110101）")
        self.home_shortcut_vars = [
            tk.StringVar(value=value) for value in self.HOME_SHORTCUT_DEFAULTS
        ]

        self.exam_candidate_var = tk.StringVar()
        self.exam_group_var = tk.StringVar()
        self.exam_plan_var = tk.StringVar()
        self.exam_output_var = tk.StringVar()
        self.exam_point_digits_var = tk.StringVar(value="2")
        self.exam_room_digits_var = tk.StringVar(value="3")
        self.exam_seat_digits_var = tk.StringVar(value="2")
        self.exam_serial_digits_var = tk.StringVar(value="4")
        self.exam_sort_mode_var = tk.StringVar(value="random")
        self.exam_rule_type_var = tk.StringVar(value="自定义")
        self.exam_rule_custom_var = tk.StringVar()
        self.status_server_name_var = tk.StringVar()
        self.status_server_host_var = tk.StringVar()
        self.status_server_port_var = tk.StringVar(value="1433")
        self.status_server_user_var = tk.StringVar()
        self.status_server_password_var = tk.StringVar()
        self.status_server_enabled_var = tk.BooleanVar(value=True)
        self.status_filter_var = tk.StringVar(value="正在进行 + 即将开始")
        self.status_stage_keyword_var = tk.StringVar()
        self.status_project_keyword_var = tk.StringVar()

        self.status_var = tk.StringVar(value="就绪")
        self.progress_text_var = tk.StringVar(value="未开始")
        self.summary_text_var = tk.StringVar(value="任务结果会显示在这里")
        self.certificate_status_var = tk.StringVar(value="未开始证件资料筛选")
        self.certificate_progress_text_var = tk.StringVar(value="未开始")
        self.certificate_summary_text_var = tk.StringVar(value="证件资料筛选结果会显示在这里")
        self.bucket_status_var = tk.StringVar(value="未加载 bucket 列表")
        self.folder_status_var = tk.StringVar(value="未加载 bucket 文件夹")
        self.search_status_var = tk.StringVar(value="未搜索文件夹")
        self.selected_folder_info_var = tk.StringVar(value="当前未选择 bucket 文件夹")
        self.certificate_bucket_status_var = tk.StringVar(value="未加载 bucket 列表")
        self.certificate_folder_status_var = tk.StringVar(value="未加载 bucket 文件夹")
        self.certificate_search_status_var = tk.StringVar(value="未搜索文件夹")
        self.certificate_selected_folder_info_var = tk.StringVar(value="当前未选择 bucket 文件夹")
        self.word_status_var = tk.StringVar(value="未开始导出")
        self.word_result_var = tk.StringVar(value="表样转换结果会显示在这里")
        self.word_preview_status_var = tk.StringVar(value="未生成预览")
        self.pack_status_var = tk.StringVar(value="未开始打包")
        self.pack_result_var = tk.StringVar(value="结果打包信息会显示在这里")
        self.sql_status_var = tk.StringVar(value="未开始生成")
        self.sql_result_var = tk.StringVar(value="SQL 生成结果会显示在这里")
        self.match_status_var = tk.StringVar(value="未开始匹配")
        self.match_result_var = tk.StringVar(value="数据匹配结果会显示在这里")
        self.update_sql_status_var = tk.StringVar(value="未生成 SQL")
        self.update_sql_result_var = tk.StringVar(value="更新 SQL 会显示在这里")
        self.phone_status_var = tk.StringVar(value="未开始解密")
        self.phone_result_var = tk.StringVar(value="电话解密结果会显示在这里")
        self.id_result_var = tk.StringVar(value="身份证工具结果会显示在这里")
        self.exam_status_var = tk.StringVar(value="未开始编排")
        self.exam_result_var = tk.StringVar(value="考场编排结果会显示在这里")
        self.status_query_status_var = tk.StringVar(value="未开始查询")
        self.folder_tree: Optional[ttk.Treeview] = None
        self.certificate_folder_tree: Optional[ttk.Treeview] = None
        self.folder_nodes: Dict[str, BrowserEntry] = {}
        self.certificate_folder_nodes: Dict[str, BrowserEntry] = {}
        self.current_folder_entries: List[BrowserEntry] = []
        self.current_certificate_folder_entries: List[BrowserEntry] = []
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.certificate_progress_bar: Optional[ttk.Progressbar] = None
        self.last_summary: Optional[WorkflowSummary] = None
        self.last_certificate_summary: Optional[CertificateFilterSummary] = None
        self.last_word_export: Optional[WordExportResult] = None
        self.last_pack_summary: Optional[PackSummary] = None
        self.last_sql_result: Optional[SqlTemplateResult] = None
        self.last_match_summary: Optional[DataMatchSummary] = None
        self.last_update_sql_result: Optional[UpdateSqlResult] = None
        self.last_phone_summary = None
        self.last_exam_summary: Optional[ExamArrangeSummary] = None
        self.last_exam_template_export: Optional[ExamTemplateExportSummary] = None
        self.last_status_summary: Optional[ProjectStageSummary] = None
        self.word_code_text = None
        self.word_preview_widget = None
        self.match_result_text = None
        self.pack_query_result_text = None
        self.sql_result_text = None
        self.update_sql_result_text = None
        self.phone_result_text = None
        self.id_result_text = None
        self.exam_result_text = None
        self.certificate_headers: List[str] = []
        self.photo_headers: List[str] = []
        self.match_target_headers: List[str] = []
        self.match_source_headers: List[str] = []
        self.update_sql_target_headers: List[str] = []
        self.update_sql_source_headers: List[str] = []
        self.match_extra_mappings: List[ColumnMapping] = []
        self.match_transfer_mappings: List[ColumnMapping] = []
        self.exam_rule_items: List[ExamRuleItem] = []
        self.exam_group_headers: List[str] = []
        self.exam_rule_base_types = ["自定义", "考点", "考场", "座号", "流水号"]
        self.status_server_configs: List[StageServerConfig] = []
        self.status_server_selected_index: Optional[int] = None
        self.certificate_bucket_values: List[str] = []
        self.certificate_folder_values: List[str] = [""]
        self.default_sash_pending = {"photo": True, "certificate": True}
        self.bucket_load_token = 0
        self.certificate_bucket_load_token = 0
        self.cloud_profiles: Dict[str, Dict[str, str]] = {
            "aliyun": self.default_cloud_profile("aliyun"),
            "tencent": self.default_cloud_profile("tencent"),
        }

        self.load_saved_settings()
        self.build_ui()
        if self.photo_template_var.get().strip() and Path(
            self.photo_template_var.get().strip()
        ).exists():
            self.load_photo_headers()
        if self.certificate_template_var.get().strip() and Path(
            self.certificate_template_var.get().strip()
        ).exists():
            self.load_certificate_headers()
        if (
            self.match_target_var.get().strip()
            and self.match_source_var.get().strip()
            and Path(self.match_target_var.get().strip()).exists()
            and Path(self.match_source_var.get().strip()).exists()
        ):
            self.load_match_headers()
        if self.update_sql_mapping_var.get().strip() and Path(
            self.update_sql_mapping_var.get().strip()
        ).exists():
            self.load_update_sql_headers()
        self.root.after(150, self.flush_logs)

    def build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.root.configure(background="#EAF0F6")

        base_font = ("Microsoft YaHei UI", 11)
        bold_font = ("Microsoft YaHei UI", 11, "bold")
        title_font = ("Microsoft YaHei UI", 18, "bold")

        style.configure("TFrame", background="#F4F7FB")
        style.configure(
            "TLabelframe",
            background="#F4F7FB",
            borderwidth=1,
            relief="solid",
            padding=12,
            bordercolor="#D7E1EC",
            lightcolor="#D7E1EC",
            darkcolor="#D7E1EC",
        )
        style.configure(
            "TLabelframe.Label",
            background="#F4F7FB",
            foreground="#162033",
            font=bold_font,
        )
        style.configure(
            "TLabel",
            background="#F4F7FB",
            foreground="#243247",
            padding=1,
            font=base_font,
        )
        style.configure(
            "TRadiobutton",
            background="#F4F7FB",
            foreground="#243247",
            font=base_font,
        )
        style.configure(
            "TCheckbutton",
            background="#F4F7FB",
            foreground="#243247",
            font=base_font,
        )
        style.configure(
            "TButton",
            padding=(14, 10),
            font=base_font,
            relief="flat",
            borderwidth=0,
            background="#FFFFFF",
            foreground="#1F3147",
        )
        style.map(
            "TButton",
            background=[
                ("disabled", "#E9EEF4"),
                ("pressed", "#D8E5F6"),
                ("active", "#EDF4FC"),
                ("!disabled", "#FFFFFF"),
            ],
            foreground=[
                ("disabled", "#90A0B3"),
                ("!disabled", "#1F3147"),
            ],
        )
        style.configure(
            "Accent.TButton",
            padding=(14, 10),
            font=bold_font,
            foreground="#FFFFFF",
            background="#3E7BFA",
        )
        style.configure(
            "TCombobox",
            padding=(10, 7),
            arrowsize=16,
            font=base_font,
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            foreground="#162033",
            bordercolor="#D7E1EC",
            lightcolor="#D7E1EC",
            darkcolor="#D7E1EC",
        )
        style.map(
            "Accent.TButton",
            background=[
                ("disabled", "#A9C4FB"),
                ("pressed", "#2D66D4"),
                ("active", "#5A92FA"),
                ("!disabled", "#3E7BFA"),
            ],
            foreground=[
                ("disabled", "#F2F2F2"),
                ("!disabled", "#FFFFFF"),
            ],
        )
        style.configure(
            "TNotebook",
            background="#F4F7FB",
            borderwidth=0,
            tabmargins=(0, 0, 0, 0),
        )
        style.layout("Navless.TNotebook.Tab", [])
        style.configure(
            "TNotebook.Tab",
            padding=(20, 12),
            background="#EEF3F8",
            foreground="#5A6D83",
            font=("Microsoft YaHei UI", 12, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", "#FFFFFF"),
                ("active", "#F6F9FD"),
                ("!selected", "#EEF3F8"),
            ],
            foreground=[
                ("selected", "#1A2940"),
                ("!selected", "#5A6D83"),
            ],
            expand=[("selected", (0, 2, 0, 0))],
        )
        style.configure(
            "Treeview",
            rowheight=32,
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            foreground="#243247",
            bordercolor="#D7E1EC",
            lightcolor="#D7E1EC",
            darkcolor="#D7E1EC",
            font=base_font,
        )
        style.configure(
            "Treeview.Heading",
            padding=(8, 8),
            font=bold_font,
            background="#EEF3F8",
            foreground="#1F3147",
        )
        style.map(
            "Treeview",
            background=[("selected", "#DCE8FF")],
            foreground=[("selected", "#162033")],
        )
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#E7EEF8",
            bordercolor="#D7E1EC",
            background="#3E7BFA",
            lightcolor="#3E7BFA",
            darkcolor="#3E7BFA",
        )

        container = ttk.Frame(self.root, padding=18)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=0)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(1, weight=1)

        title_bar = tk.Frame(container, bg="#F4F7FB")
        title_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        title_bar.grid_columnconfigure(0, weight=1)
        tk.Label(
            title_bar,
            text=f"报名系统工具箱  v{__version__}",
            font=title_font,
            bg="#F4F7FB",
            fg="#162033",
        ).grid(row=0, column=0, sticky="w")
        sidebar = tk.Frame(container, bg="#F8FBFF", width=252, highlightthickness=1, highlightbackground="#D7E1EC")
        sidebar.grid(row=1, column=0, sticky="nsw", pady=(12, 0))
        sidebar.grid_propagate(False)
        self.sidebar = sidebar

        content_frame = ttk.Frame(container)
        content_frame.grid(row=1, column=1, sticky="nsew", pady=(12, 0), padx=(12, 0))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(content_frame, style="Navless.TNotebook")
        notebook.grid(row=0, column=0, sticky="nsew")
        self.notebook = notebook
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        build_home_tab(self, notebook)
        build_home_settings_tab(self, notebook)
        build_photo_tab(self, notebook)
        build_certificate_tab(self, notebook)
        build_template_tab(self, notebook)
        build_match_tab(self, notebook)
        build_update_sql_tab(self, notebook)
        build_phone_tab(self, notebook)
        build_id_card_tab(self, notebook)
        build_pack_tab(self, notebook)
        build_help_tab(self, notebook)
        build_log_tab(self, notebook)
        self.tab_ids_by_text = {
            self.notebook.tab(tab_id, "text"): tab_id for tab_id in self.notebook.tabs()
        }
        self.build_sidebar_navigation(sidebar)
        self.show_tab_by_text("首页")

        self.write_log("桌面工具已启动。")
        self.write_log("建议先勾选“仅预览”，确认下载与分类路径后再正式执行。")
        self.write_log("可以在右侧浏览 bucket 文件夹，双击进入子目录。")
        self.write_log("也可以在右侧按文件夹名称筛选当前层级目录。")
        self.render_word_preview("")
        self.update_photo_source_mode_ui()
        self.update_certificate_mode_ui()
        self.update_certificate_source_mode_ui()
        self.update_pack_password_mode_ui()
        self.set_match_result_text(self.match_result_var.get())
        self.set_update_sql_result_text(self.update_sql_result_var.get())
        self.update_phone_mode_ui()
        self.set_phone_result_text(self.phone_result_var.get())
        self.update_id_day_values()
        self.update_id_city_values()
        self.update_id_county_values()
        self.update_id_region_hint()
        self.set_id_result_text(self.id_result_var.get())
        self.set_pack_query_result_text("可按文件夹名、压缩包名或密码查询最近打包记录。")
        self.root.after(120, lambda: self.reset_split_default("photo"))

    def set_help_content(self) -> None:
        help_content = f"""报名系统工具箱 使用说明
版本：v{__version__}

当前版本已提供以下稳定功能：

1. 照片下载与分类
- 支持阿里云 OSS、腾讯云 COS 和本地目录。
- 可按模板名单下载照片，并按分类一 / 分类二 / 分类三 / 修改名称整理。

2. 证件资料筛选
- 支持云端目录和本地目录。
- 可按身份证号、报名序号等模板列筛选人员资料。
- 支持复制整个人员文件夹，或只导出关键词文件。

3. 表样转换
- 支持 .doc / .docx / .xlsx 转 HTML。
- 提供 Net 版和 Java 版占位符导出。

4. 数据匹配
- 支持两张 Excel 按主键和附加列匹配补列。
- 会输出匹配结果和匹配结果清单。

5. 更新 SQL 生成
- 支持根据字段映射模板生成标准 UPDATE SQL。
- 生成前会备份两张表，并支持忽略空值。

6. 电话解密
- 通过 32 位 PhoneDecryptHelper 调用 DLL 解密 `web_info.info1`。
- 按主键编号关联电话密文，解密后回写到考生表 `备用3`。
- Windows 发版前请先执行 `python build_phone_decrypt_helper.py` 构建 helper。

7. 身份证工具
- 可输入 18 位身份证号校验格式、出生日期、性别、所在地和校验位。
- 可按省 / 市 / 县、出生日期、性别生成合法身份证号。

8. 结果打包
- 支持对文件或文件夹一键压缩并 AES 加密。
- 支持自动密码和历史密码查询。

9. 使用建议
- 第一次处理批量数据时，建议先用少量数据测试。
- 涉及下载、筛选、分类时，建议优先勾选“仅预览”。
- 运行细节和错误会写入“运行日志”页。
"""
        self.help_text.configure(state="normal")
        self.help_text.delete("1.0", tk.END)
        self.help_text.insert("1.0", help_content)
        self.help_text.configure(state="disabled")

    def on_tab_changed(self, _event=None) -> None:
        current_tab = self.notebook.select()
        tab_text = self.notebook.tab(current_tab, "text")
        self.sync_sidebar_selection(tab_text)
        if tab_text == "照片下载与分类":
            self.reset_split_default("photo")
        elif tab_text == "证件资料筛选":
            self.reset_split_default("certificate")

    def build_sidebar_navigation(self, parent: tk.Frame) -> None:
        header = tk.Label(
            parent,
            text="功能导航",
            bg="#F8FBFF",
            fg="#162033",
            font=("Microsoft YaHei UI", 13, "bold"),
            anchor="w",
            padx=16,
            pady=16,
        )
        header.pack(fill="x")

        for group_name, items in self.NAV_GROUPS:
            self.nav_group_expanded[group_name] = True
            group_frame = tk.Frame(parent, bg="#F8FBFF")
            group_frame.pack(fill="x", pady=(0, 6))

            header_row = tk.Frame(group_frame, bg="#F8FBFF")
            header_row.pack(fill="x")
            button = tk.Button(
                header_row,
                text=group_name,
                command=lambda name=group_name: self.toggle_nav_group(name),
                anchor="w",
                relief="flat",
                bd=0,
                bg="#F8FBFF",
                activebackground="#EEF5FF",
                fg="#213349",
                activeforeground="#213349",
                font=("Microsoft YaHei UI", 12, "bold"),
                padx=16,
                pady=10,
                cursor="hand2",
            )
            button.pack(side="left", fill="x", expand=True)
            indicator = ttk.Label(
                header_row,
                text="▾",
                background="#F8FBFF",
                foreground="#678099",
                font=("Microsoft YaHei UI", 12, "bold"),
            )
            indicator.pack(side="right", padx=(0, 16))
            self.nav_group_indicators[group_name] = indicator

            items_frame = tk.Frame(group_frame, bg="#F8FBFF")
            items_frame.pack(fill="x")
            self.nav_group_frames[group_name] = items_frame
            for label, tab_text in items:
                item_button = tk.Button(
                    items_frame,
                    text=label,
                    command=lambda name=tab_text: self.show_tab_by_text(name),
                    anchor="w",
                    relief="flat",
                    bd=0,
                    bg="#F8FBFF",
                    activebackground="#EAF2FF",
                    fg="#40546A",
                    activeforeground="#162033",
                    font=("Microsoft YaHei UI", 11),
                    padx=32,
                    pady=9,
                    cursor="hand2",
                )
                item_button.pack(fill="x")
                self.nav_item_buttons[tab_text] = item_button

        for group_name in list(self.nav_group_frames.keys())[1:]:
            self.toggle_nav_group(group_name, expanded=False)

    def toggle_nav_group(self, group_name: str, expanded: Optional[bool] = None) -> None:
        current = self.nav_group_expanded.get(group_name, True)
        target = (not current) if expanded is None else expanded
        self.nav_group_expanded[group_name] = target
        frame = self.nav_group_frames[group_name]
        indicator = self.nav_group_indicators[group_name]
        if target:
            frame.pack(fill="x")
            indicator.configure(text="▾")
        else:
            frame.pack_forget()
            indicator.configure(text="▸")

    def show_tab_by_text(self, tab_text: str) -> None:
        tab_id = self.tab_ids_by_text.get(tab_text)
        if not tab_id:
            return
        self.notebook.select(tab_id)
        self.sync_sidebar_selection(tab_text)

    def open_home_shortcut(self, index: int) -> None:
        if index < 0 or index >= len(self.home_shortcut_vars):
            return
        tab_text = self.home_shortcut_vars[index].get().strip()
        if tab_text:
            self.show_tab_by_text(tab_text)

    def save_home_shortcuts(self) -> None:
        self.save_settings()
        self.write_log("首页快捷入口配置已保存。")

    def sync_sidebar_selection(self, tab_text: str) -> None:
        self.nav_selected_text = tab_text
        for text, button in self.nav_item_buttons.items():
            is_selected = text == tab_text
            button.configure(
                bg="#DCE8FF" if is_selected else "#F8FBFF",
                fg="#173052" if is_selected else "#40546A",
                font=("Microsoft YaHei UI", 11, "bold") if is_selected else ("Microsoft YaHei UI", 11),
            )
        for group_name, items in self.NAV_GROUPS:
            if any(item_tab_text == tab_text for _, item_tab_text in items):
                if not self.nav_group_expanded.get(group_name, True):
                    self.toggle_nav_group(group_name, expanded=True)
                break

    def reset_split_default(self, pane_name: str) -> None:
        if not self.default_sash_pending.get(pane_name):
            return

        paned = self.photo_paned if pane_name == "photo" else self.certificate_paned
        try:
            width = paned.winfo_width()
            if width <= 100:
                self.root.after(120, lambda: self.reset_split_default(pane_name))
                return
            paned.sashpos(0, width // 2)
            self.default_sash_pending[pane_name] = False
        except tk.TclError:
            self.root.after(120, lambda: self.reset_split_default(pane_name))

    def add_entry_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        show: Optional[str] = None,
    ):
        return ui_add_entry_row(self, parent, row, label, variable, show)

    def add_tick_checkbutton(
        self,
        parent,
        text: str,
        variable: tk.BooleanVar,
        command=None,
        row: int = 0,
        column: int = 0,
        padx=(0, 18),
        pady=4,
        sticky: str = "w",
    ):
        return ui_add_tick_checkbutton(
            parent=parent,
            text=text,
            variable=variable,
            command=command,
            row=row,
            column=column,
            padx=padx,
            pady=pady,
            sticky=sticky,
        )

    def add_cloud_type_row(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="云类型", width=16).grid(
            row=row, column=0, sticky="w", pady=6, padx=(0, 10)
        )
        row_frame = ttk.Frame(parent)
        row_frame.grid(row=row, column=1, sticky="w", pady=6)
        ttk.Radiobutton(
            row_frame,
            text="阿里云 OSS",
            variable=self.cloud_type_var,
            value="aliyun",
            command=self.on_cloud_type_changed,
        ).pack(side="left")
        ttk.Radiobutton(
            row_frame,
            text="腾讯云 COS",
            variable=self.cloud_type_var,
            value="tencent",
            command=self.on_cloud_type_changed,
        ).pack(side="left", padx=(10, 0))

    def add_bucket_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Bucket", width=16).grid(
            row=row, column=0, sticky="w", pady=6, padx=(0, 10)
        )

        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)

        self.bucket_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.bucket_name_var,
            values=self.bucket_values,
            state="readonly",
        )
        self.bucket_combo.grid(row=0, column=0, sticky="ew")

        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        ttk.Button(
            action_frame,
            text="加载 Bucket",
            width=12,
            command=self.load_buckets,
        ).pack(side="left")
        ttk.Button(
            action_frame,
            text="保存配置",
            width=12,
            command=self.save_settings,
        ).pack(side="left", padx=(8, 0))

        ttk.Label(
            picker_frame,
            textvariable=self.bucket_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_prefix_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="下载文件夹", width=16).grid(
            row=row, column=0, sticky="w", pady=6, padx=(0, 10)
        )

        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)

        self.prefix_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.prefix_var,
            values=self.folder_values,
        )
        self.prefix_combo.grid(row=0, column=0, sticky="ew")

        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        ttk.Button(
            action_frame,
            text="加载文件夹",
            width=12,
            command=self.load_bucket_folders,
        ).pack(side="left")

        ttk.Button(
            action_frame,
            text="上一级",
            width=12,
            command=self.go_to_parent_prefix,
        ).pack(side="left", padx=(8, 0))

        ttk.Label(
            picker_frame,
            textvariable=self.folder_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_certificate_bucket_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Bucket", width=16).grid(
            row=row, column=0, sticky="w", pady=6, padx=(0, 10)
        )
        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)
        self.certificate_bucket_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.certificate_bucket_name_var,
            values=self.certificate_bucket_values,
            state="readonly",
        )
        self.certificate_bucket_combo.grid(row=0, column=0, sticky="ew")
        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(
            action_frame,
            text="加载 Bucket",
            width=12,
            command=self.load_certificate_buckets,
        ).pack(side="left")
        ttk.Button(
            action_frame,
            text="保存配置",
            width=12,
            command=self.save_settings,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(
            picker_frame,
            textvariable=self.certificate_bucket_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_certificate_prefix_picker(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="证件资料前缀", width=16).grid(
            row=row, column=0, sticky="w", pady=6, padx=(0, 10)
        )
        picker_frame = ttk.Frame(parent)
        picker_frame.grid(row=row, column=1, sticky="ew", pady=4)
        picker_frame.columnconfigure(0, weight=1)
        self.certificate_prefix_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.certificate_prefix_var,
            values=self.certificate_folder_values,
        )
        self.certificate_prefix_combo.grid(row=0, column=0, sticky="ew")
        action_frame = ttk.Frame(picker_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(
            action_frame,
            text="加载文件夹",
            width=12,
            command=self.load_certificate_folders,
        ).pack(side="left")
        ttk.Button(
            action_frame,
            text="上一级",
            width=12,
            command=self.go_to_certificate_parent_prefix,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(
            picker_frame,
            textvariable=self.certificate_folder_status_var,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def add_path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
    ) -> None:
        ui_add_path_row(self, parent, row, label, variable)

    def add_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
        button_text: str = "选择",
    ) -> None:
        ui_add_file_row(self, parent, row, label, variable, command, button_text)

    def choose_directory(self, variable: tk.StringVar) -> None:
        selected = filedialog.askdirectory(initialdir=variable.get() or str(Path.cwd()))
        if selected:
            variable.set(selected)

    def choose_certificate_template(self) -> None:
        ui_choose_certificate_template(self)

    def choose_photo_template(self) -> None:
        ui_choose_photo_template(self)

    def choose_word_source(self) -> None:
        ui_choose_word_source(self)

    def choose_sql_template(self) -> None:
        ui_choose_sql_template(self)

    def choose_update_sql_mapping(self) -> None:
        ui_choose_update_sql_mapping(self)

    def choose_pack_source_file(self) -> None:
        ui_choose_pack_source_file(self)

    def choose_pack_source_directory(self) -> None:
        ui_choose_pack_source_directory(self)

    def choose_phone_filter_file(self) -> None:
        ui_choose_phone_filter_file(self)

    def clear_status_server_form(self) -> None:
        ui_clear_status_server_form(self)

    def on_status_server_select(self, _event=None) -> None:
        ui_on_status_server_select(self, _event)

    def save_status_server(self) -> None:
        ui_save_status_server(self)

    def delete_status_server(self) -> None:
        ui_delete_status_server(self)

    def test_status_server(self) -> None:
        ui_test_status_server(self)

    def start_status_query(self) -> None:
        ui_start_status_query(self)

    def export_status_result(self) -> None:
        ui_export_status_result(self)

    def refresh_status_server_tree(self) -> None:
        ui_refresh_status_server_tree(self)

    def choose_match_target(self) -> None:
        ui_choose_match_target(self)

    def choose_match_source(self) -> None:
        ui_choose_match_source(self)

    def choose_exam_candidate_file(self) -> None:
        ui_choose_exam_candidate_file(self)

    def choose_exam_group_file(self) -> None:
        ui_choose_exam_group_file(self)

    def choose_exam_plan_file(self) -> None:
        ui_choose_exam_plan_file(self)

    def export_exam_templates_from_ui(self) -> None:
        ui_export_exam_templates_from_ui(self)

    def load_exam_group_headers(self) -> None:
        ui_load_exam_group_headers(self)

    def refresh_exam_rule_type_values(self) -> None:
        ui_refresh_exam_rule_type_values(self)

    def fill_exam_output_path(self) -> None:
        ui_fill_exam_output_path(self)

    def fill_match_output_path(self) -> None:
        ui_fill_match_output_path(self)

    def fill_phone_table_name(self) -> None:
        ui_fill_phone_table_name(self)

    def load_match_headers(self) -> None:
        ui_load_match_headers(self)

    def load_update_sql_headers(self) -> None:
        ui_load_update_sql_headers(self)

    def export_update_sql_template_file(self) -> None:
        ui_export_update_sql_template_file(self)

    def add_extra_match_mapping(self) -> None:
        ui_add_extra_match_mapping(self)

    def remove_extra_match_mapping(self) -> None:
        ui_remove_extra_match_mapping(self)

    def add_transfer_mapping(self) -> None:
        ui_add_transfer_mapping(self)

    def remove_transfer_mapping(self) -> None:
        ui_remove_transfer_mapping(self)

    def refresh_exam_rule_tree(self) -> None:
        ui_refresh_exam_rule_tree(self)

    def add_exam_rule_item(self) -> None:
        ui_add_exam_rule_item(self)

    def remove_exam_rule_item(self) -> None:
        ui_remove_exam_rule_item(self)

    def move_exam_rule_up(self) -> None:
        ui_move_exam_rule_up(self)

    def move_exam_rule_down(self) -> None:
        ui_move_exam_rule_down(self)

    def create_text_entry(self, parent, textvariable: tk.StringVar, show: str = ""):
        return ui_create_text_entry(parent, textvariable, show)

    def bind_mousewheel_to_canvas(self, canvas: tk.Canvas, target) -> None:
        ui_bind_mousewheel_to_canvas(canvas, target)

    def set_widget_state_recursive(self, widget, state: str) -> None:
        try:
            if isinstance(widget, ttk.Frame) or isinstance(widget, ttk.LabelFrame):
                pass
            elif isinstance(widget, ttk.Combobox):
                widget.configure(state="readonly" if state == "normal" and str(widget.cget("state")) == "readonly" else state)
            else:
                widget.configure(state=state)
        except tk.TclError:
            pass
        for child in widget.winfo_children():
            self.set_widget_state_recursive(child, state)

    def update_photo_source_mode_ui(self) -> None:
        ui_update_photo_source_mode_ui(self)

    def update_certificate_source_mode_ui(self) -> None:
        is_oss = self.certificate_source_mode_var.get() == "oss"
        self.set_widget_state_recursive(
            self.certificate_oss_frame,
            "normal" if is_oss else "disabled",
        )
        self.set_widget_state_recursive(
            self.certificate_browser_frame,
            "normal" if is_oss else "disabled",
        )
        if is_oss:
            self.certificate_download_button.configure(state="normal")
            self.certificate_folder_status_var.set("未加载 bucket 文件夹")
            self.certificate_search_status_var.set("未搜索文件夹")
            self.certificate_selected_folder_info_var.set("当前未选择 bucket 文件夹")
        else:
            self.certificate_download_button.configure(state="disabled")
            self.certificate_folder_status_var.set("本地模式不使用 bucket 浏览")
            self.certificate_search_status_var.set("本地模式不使用 bucket 搜索")
            self.certificate_selected_folder_info_var.set("请直接选择本地证件资料目录")

    def clear_log(self) -> None:
        ui_clear_log(self)

    def write_log(self, message: str) -> None:
        ui_write_log(self, message)

    def flush_logs(self) -> None:
        ui_flush_logs(self)

    def make_logger(self):
        return ui_make_logger(self)

    def make_progress_callback(self):
        return ui_make_progress_callback(self)

    def update_progress_ui(
        self,
        stage: str,
        current: int,
        total: int,
        current_file: str,
    ) -> None:
        ui_update_progress_ui(self, stage, current, total, current_file)

    def cancel_run(self) -> None:
        ui_cancel_run(self)

    def update_summary_ui(self, summary: Optional[WorkflowSummary]) -> None:
        ui_update_summary_ui(self, summary)

    def update_certificate_summary_ui(
        self,
        summary: Optional[CertificateFilterSummary],
    ) -> None:
        ui_update_certificate_summary_ui(self, summary)

    def update_word_export_ui(self, result: Optional[WordExportResult]) -> None:
        ui_update_word_export_ui(self, result)

    def update_pack_summary_ui(self, summary: Optional[PackSummary]) -> None:
        ui_update_pack_summary_ui(self, summary)

    def update_match_summary_ui(self, summary: Optional[DataMatchSummary]) -> None:
        ui_update_match_summary_ui(self, summary)

    def update_phone_summary_ui(self, summary) -> None:
        ui_update_phone_summary_ui(self, summary)

    def set_match_result_text(self, content: str) -> None:
        ui_set_match_result_text(self, content)

    def set_phone_result_text(self, content: str) -> None:
        ui_set_phone_result_text(self, content)

    def set_id_result_text(self, content: str) -> None:
        ui_set_id_result_text(self, content)

    def set_pack_query_result_text(self, content: str) -> None:
        ui_set_pack_query_result_text(self, content)

    def set_sql_result_text(self, content: str) -> None:
        ui_set_sql_result_text(self, content)

    def set_update_sql_result_text(self, content: str) -> None:
        ui_set_update_sql_result_text(self, content)

    def update_update_sql_ui(self, result: Optional[UpdateSqlResult]) -> None:
        ui_update_update_sql_ui(self, result)

    def update_exam_summary_ui(self, summary: Optional[ExamArrangeSummary]) -> None:
        ui_update_exam_summary_ui(self, summary)

    def update_status_summary_ui(self, summary: Optional[ProjectStageSummary]) -> None:
        ui_update_status_summary_ui(self, summary)

    def set_exam_result_text(self, content: str) -> None:
        ui_set_exam_result_text(self, content)

    def update_pack_password_mode_ui(self) -> None:
        ui_update_pack_password_mode_ui(self)

    def update_phone_mode_ui(self) -> None:
        ui_update_phone_mode_ui(self)

    def update_id_day_values(self) -> None:
        ui_update_id_day_values(self)

    def update_id_city_values(self) -> None:
        ui_update_id_city_values(self)

    def update_id_county_values(self) -> None:
        ui_update_id_county_values(self)

    def update_id_region_hint(self) -> None:
        ui_update_id_region_hint(self)

    def run_id_card_validate(self) -> None:
        ui_run_id_card_validate(self)

    def run_id_card_generate(self) -> None:
        try:
            ui_run_id_card_generate(self)
        except Exception as exc:
            messagebox.showerror("身份证生成失败", str(exc))

    def copy_generated_id_card(self) -> None:
        ui_copy_generated_id_card(self)

    def run_pack_history_query(self) -> None:
        ui_run_pack_history_query(self)

    def set_word_code(self, html_content: str) -> None:
        ui_set_word_code(self, html_content)

    def render_word_preview(self, html_content: str) -> None:
        ui_render_word_preview(self, html_content)

    def copy_word_html(self) -> None:
        ui_copy_word_html(self)

    def copy_sql_text(self) -> None:
        ui_copy_sql_text(self)

    def copy_update_sql(self) -> None:
        ui_copy_update_sql(self)

    def open_word_preview_in_browser(self) -> None:
        ui_open_word_preview_in_browser(self)

    def copy_pack_password(self) -> None:
        ui_copy_pack_password(self)

    def open_pack_file(self) -> None:
        ui_open_pack_file(self)

    def open_match_result_file(self) -> None:
        ui_open_match_result_file(self)

    def open_exam_result_file(self) -> None:
        ui_open_exam_result_file(self)

    def open_template_file(self) -> None:
        ui_open_template_file(self)

    def open_photo_report_file(self) -> None:
        ui_open_photo_report_file(self)

    def open_certificate_report_file(self) -> None:
        ui_open_certificate_report_file(self)

    def open_local_file(self, file_path: Path, not_found_title: str) -> None:
        ui_open_local_file(self, file_path, not_found_title)

    def build_options(self) -> RunOptions:
        is_oss = self.photo_source_mode_var.get() == "oss"
        return RunOptions(
            prefix=self.prefix_var.get().strip() if is_oss else "",
            download_dir=Path(self.download_dir_var.get().strip()),
            sorted_dir=Path(self.sorted_dir_var.get().strip()),
            skip_download=not is_oss,
            dry_run=self.dry_run_var.get(),
            flat=self.flat_var.get(),
            include_duplicates=self.include_duplicates_var.get(),
            move_sorted_files=self.move_sorted_files_var.get(),
            skip_existing=self.skip_existing_var.get(),
        )

    def build_config(self) -> OssConfig:
        return validate_oss_config(
            OssConfig(
                cloud_type=self.cloud_type_var.get().strip(),
                access_key_id=self.access_key_id_var.get().strip(),
                access_key_secret=self.access_key_secret_var.get().strip(),
                endpoint=self.endpoint_var.get().strip(),
                bucket_name=self.bucket_name_var.get().strip(),
            )
        )

    def build_certificate_config(self) -> OssConfig:
        return validate_oss_config(
            OssConfig(
                cloud_type=self.cloud_type_var.get().strip(),
                access_key_id=self.access_key_id_var.get().strip(),
                access_key_secret=self.access_key_secret_var.get().strip(),
                endpoint=self.endpoint_var.get().strip(),
                bucket_name=self.certificate_bucket_name_var.get().strip(),
            )
        )

    def build_credentials(self) -> tuple[str, str, str, str]:
        return ui_build_credentials(self)

    def validate_cloud_endpoint(self, cloud_type: str, endpoint: str) -> None:
        ui_validate_cloud_endpoint(self, cloud_type, endpoint)

    def check_endpoint_reachable(self, cloud_type: str, endpoint: str) -> None:
        ui_check_endpoint_reachable(self, cloud_type, endpoint)

    def format_cloud_error(self, error: str) -> str:
        return ui_format_cloud_error(self, error)

    def set_folder_values(self, folders) -> None:
        ui_set_folder_values(self, folders)

    def set_bucket_values(self, buckets: List[str]) -> None:
        ui_set_bucket_values(self, buckets)

    def set_certificate_bucket_values(self, buckets: List[str]) -> None:
        ui_set_certificate_bucket_values(self, buckets)

    def sync_bucket_values(self, buckets: List[str]) -> None:
        ui_sync_bucket_values(self, buckets)

    def set_certificate_folder_values(self, folders: List[str]) -> None:
        ui_set_certificate_folder_values(self, folders)

    def default_cloud_profile(self, cloud_type: str) -> Dict[str, str]:
        return ui_default_cloud_profile(self, cloud_type)

    def snapshot_current_cloud_profile(self) -> Dict[str, str]:
        return ui_snapshot_current_cloud_profile(self)

    def apply_cloud_profile(self, cloud_type: str) -> None:
        ui_apply_cloud_profile(self, cloud_type)

    def load_saved_settings(self) -> None:
        ui_load_saved_settings(self)

    def save_settings(self) -> None:
        ui_save_settings(self)

    def update_certificate_mode_ui(self) -> None:
        ui_update_certificate_mode_ui(self)

    def on_cloud_type_changed(self) -> None:
        ui_on_cloud_type_changed(self)

    def set_certificate_headers(self, headers: List[str]) -> None:
        ui_set_certificate_headers(self, headers)

    def set_photo_headers(self, headers: List[str]) -> None:
        ui_set_photo_headers(self, headers)

    def load_photo_headers(self) -> None:
        ui_load_photo_headers(self)

    def load_certificate_headers(self) -> None:
        ui_load_certificate_headers(self)

    def load_buckets(self) -> None:
        ui_load_buckets(self)

    def load_certificate_buckets(self) -> None:
        ui_load_certificate_buckets(self)

    def finish_bucket_load(
        self,
        buckets: Optional[List[str]] = None,
        error: Optional[str] = None,
        token: Optional[int] = None,
    ) -> None:
        ui_finish_bucket_load(self, buckets=buckets, error=error, token=token)

    def finish_certificate_bucket_load(
        self,
        buckets: Optional[List[str]] = None,
        error: Optional[str] = None,
        token: Optional[int] = None,
    ) -> None:
        ui_finish_certificate_bucket_load(self, buckets=buckets, error=error, token=token)

    def handle_bucket_load_timeout(self, token: int) -> None:
        ui_handle_bucket_load_timeout(self, token)

    def handle_certificate_bucket_load_timeout(self, token: int) -> None:
        ui_handle_certificate_bucket_load_timeout(self, token)

    def go_to_parent_prefix(self) -> None:
        ui_go_to_parent_prefix(self)

    def go_to_certificate_parent_prefix(self) -> None:
        ui_go_to_certificate_parent_prefix(self)

    def load_bucket_folders(self) -> None:
        ui_load_bucket_folders(self)

    def load_certificate_folders(self) -> None:
        ui_load_certificate_folders(self)

    def finish_folder_load(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        ui_finish_folder_load(self, entries=entries, error=error)

    def finish_certificate_folder_load(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        ui_finish_certificate_folder_load(self, entries=entries, error=error)

    def render_folder_tree(self, entries: List[BrowserEntry]) -> None:
        ui_render_folder_tree(self, entries)

    def render_certificate_folder_tree(self, entries: List[BrowserEntry]) -> None:
        ui_render_certificate_folder_tree(self, entries)

    def render_search_tree(self, object_keys: List[str]) -> None:
        if self.search_tree is None:
            return
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        self.search_nodes = {}
        for object_key in object_keys:
            filename = Path(object_key).name
            parent_folder = str(Path(object_key).parent)
            if parent_folder == ".":
                parent_folder = "/"
            node_id = self.search_tree.insert(
                "",
                "end",
                text=filename,
                values=(parent_folder,),
            )
            self.search_nodes[node_id] = object_key

    def render_certificate_search_tree(self, object_keys: List[str]) -> None:
        ui_render_certificate_search_tree(self, object_keys)

    def on_tree_select(self, _event=None) -> None:
        ui_on_tree_select(self, _event)

    def on_tree_double_click(self, _event=None) -> None:
        ui_on_tree_double_click(self, _event)

    def on_certificate_tree_select(self, _event=None) -> None:
        ui_on_certificate_tree_select(self, _event)

    def on_certificate_tree_double_click(self, _event=None) -> None:
        ui_on_certificate_tree_double_click(self, _event)

    def search_bucket_files(self) -> None:
        ui_search_bucket_files(self)

    def search_certificate_files(self) -> None:
        ui_search_certificate_files(self)

    def filter_folder_entries(self, entries: List[BrowserEntry], keyword: str) -> List[BrowserEntry]:
        return ui_filter_folder_entries(self, entries, keyword)

    def finish_search(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        ui_finish_search(self, entries=entries, error=error)

    def finish_certificate_search(
        self,
        entries: Optional[List[BrowserEntry]] = None,
        error: Optional[str] = None,
    ) -> None:
        ui_finish_certificate_search(self, entries=entries, error=error)

    def on_search_double_click(self, _event=None) -> None:
        ui_on_search_double_click(self, _event)

    def on_certificate_search_double_click(self, _event=None) -> None:
        ui_on_certificate_search_double_click(self, _event)

    def refresh_selected_folder_count(self) -> None:
        ui_refresh_selected_folder_count(self)

    def finish_count_refresh(
        self,
        count: int = 0,
        prefix: str = "",
        error: Optional[str] = None,
    ) -> None:
        ui_finish_count_refresh(self, count=count, prefix=prefix, error=error)

    def build_certificate_options(self) -> CertificateFilterOptions:
        template_value = self.certificate_template_var.get().strip()
        match_column = self.certificate_match_column_var.get().strip()

        if not template_value:
            raise ValueError("请选择人员模板文件。")
        if not match_column:
            raise ValueError("请选择用于匹配人员文件夹的模板列。")

        source_dir = self.resolve_certificate_source_dir(require_value=True)
        output_dir = self.resolve_certificate_output_dir()

        keyword = ""
        if self.certificate_mode_var.get() == "keyword":
            keyword = self.certificate_keyword_var.get().strip()
            if not keyword:
                raise ValueError("关键词模式下请输入文件关键词。")
        if self.certificate_rename_folder_var.get() and not self.certificate_folder_name_column_var.get().strip():
            raise ValueError("请选择导出后文件夹名称列。")

        return CertificateFilterOptions(
            template_path=Path(template_value),
            source_dir=source_dir,
            output_dir=output_dir,
            match_column=match_column,
            rename_folder=self.certificate_rename_folder_var.get(),
            folder_name_column=self.certificate_folder_name_column_var.get().strip(),
            classify_output=self.certificate_classify_var.get(),
            keyword=keyword,
            dry_run=self.certificate_dry_run_var.get(),
        )

    def resolve_certificate_source_dir(self, require_value: bool = True) -> Path:
        source_value = self.certificate_source_dir_var.get().strip()
        if not source_value:
            if require_value:
                raise ValueError("请选择证件资料目录。")
            return Path(".").resolve()

        base_dir = Path(source_value)
        if self.certificate_source_mode_var.get() == "oss":
            return build_prefixed_directory(base_dir, self.certificate_prefix_var.get().strip(), "证件资料")
        return base_dir.expanduser().resolve()

    def resolve_certificate_output_dir(self) -> Path:
        output_value = self.certificate_output_dir_var.get().strip()
        if not output_value:
            raise ValueError("请选择输出目录。")
        return Path(output_value).expanduser().resolve()

    def start_photo_download_run(self) -> None:
        ui_start_photo_download_run(self)

    def start_photo_classify_run(self) -> None:
        ui_start_photo_classify_run(self)

    def start_certificate_download_run(self) -> None:
        ui_start_certificate_download_run(self)

    def start_certificate_run(self) -> None:
        ui_start_certificate_run(self)

    def start_word_export(self, variant: str) -> None:
        ui_start_word_export(self, variant)

    def start_sql_render(self) -> None:
        ui_start_sql_render(self)

    def start_update_sql_render(self) -> None:
        ui_start_update_sql_render(self)

    def start_pack_run(self) -> None:
        ui_start_pack_run(self)

    def start_phone_decrypt_run(self) -> None:
        ui_start_phone_decrypt_run(self)

    def start_match_run(self) -> None:
        ui_start_match_run(self)

    def start_exam_arrange_run(self) -> None:
        ui_start_exam_arrange_run(self)


def _crash_log_path() -> Path:
    return Path(os.environ.get("APPDATA", Path.home())) / "aliyun_photo_manager_crash.log"


def main() -> None:
    try:
        root = tk.Tk()

        def _on_tk_error(exc_type, exc_value, exc_tb):
            import traceback
            msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
            log_path = _crash_log_path()
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"\n{'='*60}\nTkinter callback exception:\n{msg}\n")
            try:
                messagebox.showerror("程序异常", f"发生错误，已写入日志：\n{log_path}\n\n{msg[:500]}")
            except Exception:
                pass

        root.report_callback_exception = _on_tk_error

        app = App(root)
        root.mainloop()
    except Exception:
        import traceback
        msg = traceback.format_exc()
        log_path = _crash_log_path()
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\nTop-level crash:\n{msg}\n")
        try:
            messagebox.showerror("程序崩溃", f"发生严重错误，已写入日志：\n{log_path}\n\n{msg[:500]}")
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
