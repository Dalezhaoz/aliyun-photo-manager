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
from .sql_template_executor import SqlTemplateResult, render_sql_template
from .update_sql_generator import UpdateSqlResult
from .data_matcher import (
    ColumnMapping,
    DataMatchOptions,
    DataMatchSummary,
    list_headers as list_match_headers,
    run_data_match,
)
from .exam_arranger import (
    ExamArrangeOptions,
    ExamArrangeSummary,
    ExamRuleItem,
    ExamTemplateExportSummary,
    export_exam_templates,
    list_headers as list_exam_headers,
    run_exam_arrangement,
)
from .project_stage_report import ProjectStageSummary, StageServerConfig
from .word_to_html import WordExportResult, export_word_to_html
from .ui import (
    add_entry_row as ui_add_entry_row,
    add_file_row as ui_add_file_row,
    add_path_row as ui_add_path_row,
    add_tick_checkbutton as ui_add_tick_checkbutton,
    add_exam_rule_item as ui_add_exam_rule_item,
    apply_cloud_profile as ui_apply_cloud_profile,
    bind_mousewheel_to_canvas as ui_bind_mousewheel_to_canvas,
    build_credentials as ui_build_credentials,
    check_endpoint_reachable as ui_check_endpoint_reachable,
    choose_certificate_template as ui_choose_certificate_template,
    choose_exam_candidate_file as ui_choose_exam_candidate_file,
    choose_exam_group_file as ui_choose_exam_group_file,
    choose_exam_plan_file as ui_choose_exam_plan_file,
    choose_match_source as ui_choose_match_source,
    choose_match_target as ui_choose_match_target,
    choose_pack_source_directory as ui_choose_pack_source_directory,
    choose_pack_source_file as ui_choose_pack_source_file,
    choose_photo_template as ui_choose_photo_template,
    choose_sql_template as ui_choose_sql_template,
    choose_update_sql_mapping as ui_choose_update_sql_mapping,
    choose_word_source as ui_choose_word_source,
    copy_update_sql as ui_copy_update_sql,
    copy_sql_text as ui_copy_sql_text,
    copy_word_html as ui_copy_word_html,
    export_exam_templates_from_ui as ui_export_exam_templates_from_ui,
    export_update_sql_template_file as ui_export_update_sql_template_file,
    fill_exam_output_path as ui_fill_exam_output_path,
    fill_match_output_path as ui_fill_match_output_path,
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
    load_exam_group_headers as ui_load_exam_group_headers,
    load_match_headers as ui_load_match_headers,
    load_photo_headers as ui_load_photo_headers,
    load_update_sql_headers as ui_load_update_sql_headers,
    load_bucket_folders as ui_load_bucket_folders,
    move_exam_rule_down as ui_move_exam_rule_down,
    move_exam_rule_up as ui_move_exam_rule_up,
    add_extra_match_mapping as ui_add_extra_match_mapping,
    open_exam_result_file as ui_open_exam_result_file,
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
    refresh_exam_rule_tree as ui_refresh_exam_rule_tree,
    refresh_exam_rule_type_values as ui_refresh_exam_rule_type_values,
    remove_exam_rule_item as ui_remove_exam_rule_item,
    add_transfer_mapping as ui_add_transfer_mapping,
    remove_transfer_mapping as ui_remove_transfer_mapping,
    render_word_preview as ui_render_word_preview,
    search_certificate_files as ui_search_certificate_files,
    search_bucket_files as ui_search_bucket_files,
    set_certificate_headers as ui_set_certificate_headers,
    set_exam_result_text as ui_set_exam_result_text,
    set_photo_headers as ui_set_photo_headers,
    set_word_code as ui_set_word_code,
    update_certificate_summary_ui as ui_update_certificate_summary_ui,
    update_match_summary_ui as ui_update_match_summary_ui,
    update_exam_summary_ui as ui_update_exam_summary_ui,
    update_photo_source_mode_ui as ui_update_photo_source_mode_ui,
    update_summary_ui as ui_update_summary_ui,
    update_word_export_ui as ui_update_word_export_ui,
    set_match_result_text as ui_set_match_result_text,
    open_match_result_file as ui_open_match_result_file,
    open_word_preview_in_browser as ui_open_word_preview_in_browser,
    open_certificate_report_file as ui_open_certificate_report_file,
    open_local_file as ui_open_local_file,
    open_pack_file as ui_open_pack_file,
    copy_pack_password as ui_copy_pack_password,
    build_status_tab,
    clear_status_server_form as ui_clear_status_server_form,
    delete_status_server as ui_delete_status_server,
    dump_status_settings as ui_dump_status_settings,
    export_status_result as ui_export_status_result,
    load_status_settings as ui_load_status_settings,
    on_status_server_select as ui_on_status_server_select,
    refresh_status_server_tree as ui_refresh_status_server_tree,
    save_status_server as ui_save_status_server,
    start_status_query as ui_start_status_query,
    test_status_server as ui_test_status_server,
    update_status_summary_ui as ui_update_status_summary_ui,
    run_pack_history_query as ui_run_pack_history_query,
    set_pack_query_result_text as ui_set_pack_query_result_text,
    set_sql_result_text as ui_set_sql_result_text,
    set_update_sql_result_text as ui_set_update_sql_result_text,
    start_certificate_download_run as ui_start_certificate_download_run,
    start_certificate_run as ui_start_certificate_run,
    start_exam_arrange_run as ui_start_exam_arrange_run,
    start_match_run as ui_start_match_run,
    start_pack_run as ui_start_pack_run,
    start_photo_classify_run as ui_start_photo_classify_run,
    start_photo_download_run as ui_start_photo_download_run,
    start_sql_render as ui_start_sql_render,
    start_update_sql_render as ui_start_update_sql_render,
    start_word_export as ui_start_word_export,
    update_update_sql_ui as ui_update_update_sql_ui,
    update_pack_password_mode_ui as ui_update_pack_password_mode_ui,
    update_pack_summary_ui as ui_update_pack_summary_ui,
    cancel_run as ui_cancel_run,
    clear_log as ui_clear_log,
    build_certificate_tab,
    build_exam_tab,
    build_help_tab,
    build_log_tab,
    build_match_tab,
    build_pack_tab,
    build_photo_tab,
    build_sql_tab,
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


class App:
    SETTINGS_FILE = Path(__file__).resolve().parents[2] / ".gui_settings.json"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"报名系统工具箱 v{__version__}")
        self.root.geometry("1080x760")
        self.root.minsize(920, 680)

        self.log_queue: "queue.Queue[object]" = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.cancel_event = threading.Event()

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
        self.last_exam_summary: Optional[ExamArrangeSummary] = None
        self.last_exam_template_export: Optional[ExamTemplateExportSummary] = None
        self.last_status_summary: Optional[ProjectStageSummary] = None
        self.word_code_text = None
        self.word_preview_widget = None
        self.match_result_text = None
        self.pack_query_result_text = None
        self.sql_result_text = None
        self.update_sql_result_text = None
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

        style.configure("TFrame", background="#F6F4EF")
        style.configure("TLabelframe", background="#F6F4EF", borderwidth=1, relief="solid", padding=10)
        style.configure(
            "TLabelframe.Label",
            background="#F6F4EF",
            foreground="#2F2A24",
            font=("Helvetica", 11, "bold"),
        )
        style.configure(
            "TLabel",
            background="#F6F4EF",
            foreground="#2F2A24",
            padding=1,
            font=("Helvetica", 11),
        )
        style.configure(
            "TRadiobutton",
            background="#F6F4EF",
            foreground="#2F2A24",
            font=("Helvetica", 11),
        )
        style.configure(
            "TCheckbutton",
            background="#F6F4EF",
            foreground="#2F2A24",
            font=("Helvetica", 11),
        )
        style.configure(
            "TButton",
            padding=(14, 9),
            font=("Helvetica", 11),
            relief="flat",
        )
        style.map(
            "TButton",
            background=[
                ("disabled", "#E2DED3"),
                ("pressed", "#CFC8B8"),
                ("active", "#E7E0D0"),
                ("!disabled", "#DDD6C5"),
            ],
            foreground=[
                ("disabled", "#9B9588"),
                ("!disabled", "#2F2A24"),
            ],
        )
        style.configure(
            "Accent.TButton",
            padding=(14, 9),
            font=("Helvetica", 11, "bold"),
            foreground="#FFFFFF",
            background="#5D8C63",
        )
        style.configure(
            "TCombobox",
            padding=(8, 6),
            arrowsize=16,
            font=("Helvetica", 11),
        )
        style.map(
            "Accent.TButton",
            background=[
                ("disabled", "#A5B9A8"),
                ("pressed", "#4E7753"),
                ("active", "#6A9971"),
                ("!disabled", "#5D8C63"),
            ],
            foreground=[
                ("disabled", "#F2F2F2"),
                ("!disabled", "#FFFFFF"),
            ],
        )
        style.configure(
            "TNotebook",
            background="#F6F4EF",
            borderwidth=0,
            tabmargins=(0, 0, 0, 0),
        )
        style.configure(
            "TNotebook.Tab",
            padding=(20, 12),
            background="#DDD6C5",
            foreground="#6C6659",
            font=("Helvetica", 12, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", "#F7F4EC"),
                ("active", "#EAE3D3"),
                ("!selected", "#DDD6C5"),
            ],
            foreground=[
                ("selected", "#1F1B17"),
                ("!selected", "#6C6659"),
            ],
            expand=[("selected", (0, 2, 0, 0))],
        )
        style.configure(
            "Treeview",
            rowheight=30,
            fieldbackground="#FCFBF7",
            background="#FCFBF7",
            foreground="#2F2A24",
            bordercolor="#D8D1C3",
            lightcolor="#D8D1C3",
            darkcolor="#D8D1C3",
            font=("Helvetica", 11),
        )
        style.configure(
            "Treeview.Heading",
            padding=(8, 8),
            font=("Helvetica", 11, "bold"),
            background="#E7E0D0",
            foreground="#2F2A24",
        )
        style.map(
            "Treeview",
            background=[("selected", "#DCE8D7")],
            foreground=[("selected", "#1F1B17")],
        )
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#E7E0D0",
            bordercolor="#D8D1C3",
            background="#5D8C63",
            lightcolor="#5D8C63",
            darkcolor="#5D8C63",
        )

        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        title = ttk.Label(
            container,
            text=f"报名系统工具箱 v{__version__}",
            font=("Helvetica", 18, "bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        notebook = ttk.Notebook(container)
        notebook.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.notebook = notebook
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        build_photo_tab(self, notebook)
        build_certificate_tab(self, notebook)
        build_template_tab(self, notebook)
        build_sql_tab(self, notebook)
        build_status_tab(self, notebook)
        build_match_tab(self, notebook)
        build_update_sql_tab(self, notebook)
        build_exam_tab(self, notebook)
        build_pack_tab(self, notebook)
        build_help_tab(self, notebook)
        build_log_tab(self, notebook)

        self.write_log("桌面工具已启动。")
        self.write_log("建议先勾选“仅预览”，确认下载与分类路径后再正式执行。")
        self.write_log("可以在右侧浏览 bucket 文件夹，双击进入子目录。")
        self.write_log("也可以在右侧按文件夹名称筛选当前层级目录。")
        self.render_word_preview("")
        self.update_photo_source_mode_ui()
        self.update_certificate_mode_ui()
        self.update_certificate_source_mode_ui()
        self.update_pack_password_mode_ui()
        self.refresh_status_server_tree()
        self.update_status_summary_ui(None)
        self.set_match_result_text(self.match_result_var.get())
        self.set_sql_result_text("")
        self.set_update_sql_result_text(self.update_sql_result_var.get())
        self.load_exam_group_headers()
        self.refresh_exam_rule_tree()
        self.set_exam_result_text(self.exam_result_var.get())
        self.set_pack_query_result_text("可按文件夹名、压缩包名或密码查询最近打包记录。")
        self.root.after(120, lambda: self.reset_split_default("photo"))

    def set_help_content(self) -> None:
        help_content = f"""报名系统工具箱 使用说明
版本：v{__version__}

一、照片下载与分类

适用场景：
- 批量下载考生照片
- 只下载模板中的人员照片
- 按模板中的分类信息整理照片

操作步骤：
1. 选择数据来源：
   - 从云存储下载后处理
   - 直接处理本地目录
2. 如果使用云存储，填写云类型、Endpoint / Region、AccessKey，并点击“加载 Bucket”。
3. 选择 bucket 后，点击右侧“加载当前层级”，找到照片所在目录。
4. 选择下载目录和分类目录。
5. 如需按名单下载，选择人员模板，点击“加载模板列”，选择匹配列，再勾选“只下载模板中的人员”。
6. 点击“下载并生成模板”。
7. 打开生成的 Excel 模板，填写“分类一 / 分类二 / 分类三 / 修改名称”。
8. 回到程序，点击“按模板分类”。

结果说明：
- 会生成“照片分类模板.xlsx”
- 会生成“照片分类结果清单.xlsx”
- 如果勾选“仅预览，不实际执行”，只显示计划，不真正下载和分类

二、证件资料筛选

适用场景：
- 按模板筛选部分人员的证件资料
- 按分类一、分类二、分类三整理资料
- 只导出某类证件，例如“学历证书”

操作步骤：
1. 选择数据来源：
   - 从云存储下载后处理
   - 直接处理本地目录
2. 选择人员模板，点击“加载模板列”，选择匹配列。
3. 如果需要，勾选“导出后文件夹重命名”，再选择“导出后文件夹名称列”。
4. 如果使用云存储，填写云类型、Endpoint / Region、AccessKey，并点击“加载 Bucket”。
5. 选择 bucket 后，点击右侧“加载当前层级”，找到证件资料目录。
6. 选择证件资料目录。
7. 如果需要先下载资料，可以点击“下载证件资料”。
8. 选择输出目录。
9. 选择筛选模式：
   - 复制整个人员文件夹
   - 只复制关键词文件
10. 如果是关键词模式，填写关键词，例如“学历证书”。
11. 根据需要勾选“按分类一 / 分类二 / 分类三建立目录”或“仅预览，不实际执行”。
12. 点击“开始筛选”。

结果说明：
- 会生成“证件资料筛选结果清单.xlsx”
- 如果模板中有分类信息，可以按分类目录导出
- 如果勾选“导出后文件夹重命名”，输出目录中的人员文件夹名称会按你选择的列来命名

三、表样转换

适用场景：
- 把 Word / Excel 表样转换成 HTML 代码
- 生成 Net 版或 Java 版占位符模板

操作步骤：
1. 选择表样文件（.doc / .docx / .xlsx）。
2. 点击“Net版导出”或“Java版导出”。
3. 在“代码”页查看并复制 HTML。
4. 如需查看页面效果，可以使用“浏览器预览”。

结果说明：
- Net 版占位符示例：{{[#考生表视图.姓名#]}}
- Java 版占位符示例：${{考生.姓名}}
- Excel 默认读取第一个 sheet
- Windows 下建议使用“浏览器预览”

四、SQL 配置执行

适用场景：
- 使用现成 SQL 模板快速生成一键执行脚本
- 把考试代码、考试年月、报名时间、审核时间等关键参数直接填入 SQL
- 避免手工打开 SQL 再逐段修改时间

操作步骤：
1. 选择 SQL 模板文件（支持 .sql / .txt）。
2. 填写考试代码和考试年月。
3. 填写报名开始 / 结束时间、审核开始 / 结束时间。
4. 点击“生成 SQL”。
5. 在下方结果区检查生成内容。
6. 点击“复制 SQL”，粘贴到数据库工具中执行。

结果说明：
- 会直接把模板中的考试代码、考试年月替换成你填写的值
- 会把报名和审核时间替换成明确时间，不再依赖 GETDATE()+N
- 时间支持：
  - YYYY-MM-DD HH:MM[:SS]
  - YYYY/MM/DD HH:MM[:SS]

五、更新SQL生成

适用场景：
- 考生表先导入基础字段，补充字段先导入临时表
- 按字段映射模板生成标准 UPDATE SQL
- 在执行前自动生成两张备份表

操作步骤：
1. 点击“导出模板”，生成“更新SQL字段映射模板.xlsx”。
2. 在模板中填写：
   - 考生表字段名
   - 临时表字段名
   - 是否更新
3. 在程序中选择映射模板并点击“加载字段”。
4. 输入考生表名称和临时表名称。
5. 选择考生表关联字段和临时表关联字段。
6. 如需防止空值覆盖正式表，勾选“忽略空值，不覆盖正式表”。
7. 点击“生成 SQL”。
8. 点击“复制 SQL”，再到数据库工具中执行。

结果说明：
- 生成的 SQL 会先备份考生表和临时表
- 只更新模板中“是否更新”为“是”的字段
- 可直接在结果区查看完整 SQL

六、项目阶段汇总

适用场景：
- 一次查看多台 SQL Server 上所有报名项目阶段的状态
- 自动跳过没有业务表的数据库
- 汇总“正在进行”和“即将开始”的项目阶段

操作步骤：
1. 在“服务器配置”里填写服务器名称、地址、端口、用户名和密码。
2. 点击“新增/更新”保存服务器。
3. 可先点“测试连接”确认该服务器能连接。
4. 根据需要设置：
   - 状态筛选
   - 阶段关键字
   - 项目关键字
5. 点击“开始查询”。
6. 查询完成后，可点击“导出 Excel”。

结果说明：
- 会遍历每台服务器上的在线数据库
- 只有同时存在 EI_ExamTreeDesc、web_SR_CodeItem、WEB_SR_SetTime 三张表的库才参与汇总
- 结果会显示：服务器、数据库、项目名称、阶段名称、开始时间、结束时间、当前状态

七、数据匹配

适用场景：
- 两个 Excel 表之间按考号、身份证号、姓名等字段补列
- 代替手工写 VLOOKUP、XLOOKUP 或临时写 SQL
- 姓名可能重名时，再叠加单位、岗位等附加列提高匹配准确性

操作步骤：
1. 选择目标表和来源表（支持 .xlsx / .xls）。
2. 点击“加载表头”。
3. 分别选择目标表匹配列和来源表匹配列。
4. 如需提高准确性，在“附加匹配列映射”中设置“目标表列 -> 来源表列”，例如“单位 -> 报考单位”。
5. 在“补充列映射”中设置结果列名和来源表列，例如“身份证号 <- 身份证号”。
6. 确认输出文件路径后，点击“开始匹配”。

结果说明：
- 会生成新的 Excel 结果文件，不改原始目标表
- 结果文件会新增补充列
- 会额外生成一个“匹配结果清单”sheet，方便核对未匹配和重复键
- 如果第一行只是“附件1”这类说明，程序会尽量自动跳过并识别真正表头

八、考场编排

适用场景：
- 根据标准模板给考生批量补充考点、考场、座号、考号
- 考号规则不固定时，在程序里配置拼接顺序
- 混编考场时，通过编排片段表控制同一考场的不同座位区间
- 岗位归组表后续新增字段时，也可以直接作为考号拼接片段使用

操作步骤：
1. 可以先点击“导出标准模板”，生成三张标准表后再补充数据。
2. 准备三张标准表：
   - 考生明细表：姓名、身份证号、招聘单位、岗位名称
   - 岗位归组表：招聘单位、岗位名称、科目组、岗位编码、科目号
   - 编排片段表：科目组、考点、考场号、起始座号、结束座号、人数、起始流水号、结束流水号、备注
3. 在程序中选择这三张表。
4. 设置考点、考场、座号、流水号的位数。
5. 选择同组内顺序：
   - 按原顺序
   - 随机打乱
6. 在“考号规则”里按顺序添加规则，例如：
   - 自定义
   - 考点
   - 考场
   - 座号
7. 点击“开始编排”。

结果说明：
- 会生成新的 Excel 结果文件
- 在原始基础上新增：科目组、岗位编码、科目号、考点、考场、座号、考号、编排备注
- 如果岗位未找到归组，或科目组未找到编排片段，会在“编排备注”里标明原因
- 同一科目组内默认随机打乱后再分配，也可以手动切换为按原顺序

九、结果打包

适用场景：
- 把处理后的结果文件或文件夹直接压缩成加密 zip
- 发送给客户或第三方时提高安全性

操作步骤：
1. 选择“待打包文件”或“待打包文件夹”。
2. 选择“输出目录”。
3. 如需客户指定密码，可勾选“手动设置密码”并输入密码。
4. 点击“一键打包并加密”。
5. 程序会自动使用原文件名或文件夹名生成压缩包名。
6. 如果未手动输入密码，程序会自动生成密码，并在结果区显示。
7. 如忘记密码，可在下方“密码查询”中按来源名称或压缩包名查询。

结果说明：
- 压缩包名称默认和文件夹名称一致
- 自动密码规则：当天日期 + 4位随机字符，例如 260313A7KQ
- 打包记录会保存到本地 JSON 文件，方便后期查询密码
- 可直接点击“复制密码”或“打开压缩包”

十、使用建议

- 第一次处理大批量文件时，建议先用少量数据测试。
- 使用模板前，建议先确认匹配列和分类列填写正确。
- 处理完成后，可以直接点击“打开结果清单”核对结果。
- 切换阿里云 / 腾讯云时，程序会分别记住两套配置。

十一、运行日志

运行日志用于查看下载、筛选、分类和导出的详细过程。
如果需要回看处理步骤，可以打开“运行日志”页查看。
"""
        self.help_text.configure(state="normal")
        self.help_text.delete("1.0", tk.END)
        self.help_text.insert("1.0", help_content)
        self.help_text.configure(state="disabled")

    def on_tab_changed(self, _event=None) -> None:
        current_tab = self.notebook.select()
        tab_text = self.notebook.tab(current_tab, "text")
        if tab_text == "照片下载与分类":
            self.reset_split_default("photo")
        elif tab_text == "证件资料筛选":
            self.reset_split_default("certificate")

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

    def set_match_result_text(self, content: str) -> None:
        ui_set_match_result_text(self, content)

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
        settings = {}
        if self.SETTINGS_FILE.exists():
            try:
                settings = json.loads(self.SETTINGS_FILE.read_text(encoding="utf-8"))
            except Exception:
                settings = {}
        ui_load_status_settings(self, settings)

    def save_settings(self) -> None:
        ui_save_settings(self)
        settings = {}
        if self.SETTINGS_FILE.exists():
            try:
                settings = json.loads(self.SETTINGS_FILE.read_text(encoding="utf-8"))
            except Exception:
                settings = {}
        settings.update(ui_dump_status_settings(self))
        self.SETTINGS_FILE.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8"
        )

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

    def start_match_run(self) -> None:
        ui_start_match_run(self)

    def start_exam_arrange_run(self) -> None:
        ui_start_exam_arrange_run(self)


def main() -> None:
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
