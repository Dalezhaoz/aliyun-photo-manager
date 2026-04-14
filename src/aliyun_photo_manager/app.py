from dataclasses import dataclass
from pathlib import Path
from threading import Event
from typing import Callable, Optional

from .certificate_filter import load_column_values
from .config import OssConfig
from .downloader import DownloadResult, download_photos
from .excel_classifier import ClassificationResult, apply_classification_from_template, generate_template


LogFn = Optional[Callable[[str], None]]
ProgressFn = Optional[Callable[[str, int, int, str], None]]


@dataclass
class RunOptions:
    download_dir: Path
    sorted_dir: Path
    prefix: str = ""
    skip_download: bool = False
    dry_run: bool = False
    flat: bool = False
    include_duplicates: bool = False
    move_sorted_files: bool = False
    skip_existing: bool = True
    filter_download: bool = False
    filter_template_path: Optional[Path] = None
    filter_column: str = ""


@dataclass
class WorkflowSummary:
    download_dir: Path
    sorted_dir: Path
    template_path: Path
    download_result: Optional[DownloadResult]
    template_file_count: int
    classified_count: int
    template_created: bool
    cancelled: bool = False
    dry_run: bool = False
    report_path: Optional[Path] = None


def ensure_child_directory(base_dir: Path, child_name: str) -> Path:
    normalized = base_dir.expanduser().resolve()
    if normalized.name == child_name:
        return normalized
    return normalized / child_name


def build_prefixed_directory(base_dir: Path, prefix: str, child_name: str) -> Path:
    normalized = base_dir.expanduser().resolve()
    cleaned_prefix = prefix.strip().strip("/")
    if cleaned_prefix:
        return normalized.joinpath(*cleaned_prefix.split("/"), child_name)
    return ensure_child_directory(normalized, child_name)


def resolve_photo_directories(options: RunOptions) -> tuple[Path, Path]:
    if options.skip_download:
        return (
            options.download_dir.expanduser().resolve(),
            options.sorted_dir.expanduser().resolve(),
        )
    return (
        build_prefixed_directory(options.download_dir, options.prefix, "下载文件"),
        build_prefixed_directory(options.sorted_dir, options.prefix, "分类结果"),
    )


def _build_photo_key_filter(options: RunOptions) -> Optional[Callable[[str], bool]]:
    if not options.filter_download:
        return None
    if options.filter_template_path is None or not options.filter_template_path.exists():
        raise ValueError("已启用按表过滤下载，请先选择有效的过滤表。")
    if not options.filter_column.strip():
        raise ValueError("已启用按表过滤下载，请先选择文件名前缀列。")

    allowed_prefixes = load_column_values(options.filter_template_path, options.filter_column)
    normalized_prefixes = tuple(value.strip() for value in allowed_prefixes if value.strip())
    if not normalized_prefixes:
        raise ValueError("过滤表中的文件名前缀列没有可用数据。")

    def photo_key_filter(object_key: str) -> bool:
        stem = Path(object_key).stem
        return any(stem.startswith(prefix_value) for prefix_value in normalized_prefixes)

    return photo_key_filter


def run_photo_download_and_template(
    options: RunOptions,
    oss_config: Optional[OssConfig] = None,
    logger: LogFn = None,
    progress_callback: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> WorkflowSummary:
    def log(message: str) -> None:
        if logger is not None:
            logger(message)

    download_dir, sorted_dir = resolve_photo_directories(options)
    template_path = download_dir / "照片分类模板.xlsx"

    log("开始执行照片下载/模板任务。")
    log(f"实际下载目录：{download_dir}")
    log(f"实际分类目录：{sorted_dir}")

    download_result: Optional[DownloadResult] = None
    if not options.skip_download:
        if oss_config is None:
            raise ValueError("未提供 OSS 配置。")
        photo_key_filter = _build_photo_key_filter(options)
        if options.filter_download:
            prefixes = load_column_values(options.filter_template_path, options.filter_column)
            log(f"本次仅下载过滤表中的照片，匹配列：{options.filter_column}，共 {len(prefixes)} 个前缀。")
        log("开始下载 OSS 照片。")
        download_result = download_photos(
            config=oss_config,
            prefix=options.prefix,
            download_dir=download_dir,
            dry_run=options.dry_run,
            skip_existing=options.skip_existing,
            logger=logger,
            progress_callback=progress_callback,
            cancel_event=cancel_event,
            key_filter=photo_key_filter,
        )
        if cancel_event is not None and cancel_event.is_set():
            log("下载已取消。")
            return WorkflowSummary(
                download_dir=download_dir,
                sorted_dir=sorted_dir,
                template_path=template_path,
                download_result=download_result,
                template_file_count=0,
                classified_count=0,
                template_created=False,
                cancelled=True,
                dry_run=options.dry_run,
            )
        log("OSS 下载阶段完成。")
    else:
        log("当前为本地模式，直接使用本地照片目录生成模板。")

    if not download_dir.exists() and not options.dry_run:
        raise FileNotFoundError(f"下载目录不存在：{download_dir}")

    log("开始生成 Excel 分类模板。")
    template_result = generate_template(
        source_dir=download_dir,
        dry_run=options.dry_run,
        logger=logger,
    )
    if options.dry_run:
        log("当前为预览模式，不会真正生成 Excel。")
    elif template_result.created:
        log(f"已生成 Excel 模板，共写入 {template_result.file_count} 个文件。")
    else:
        log(f"已更新 Excel 模板，共写入 {template_result.file_count} 个文件。")
    log("下载/模板任务完成。")
    return WorkflowSummary(
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


def run_photo_classification_only(
    options: RunOptions,
    logger: LogFn = None,
    progress_callback: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> WorkflowSummary:
    del progress_callback

    def log(message: str) -> None:
        if logger is not None:
            logger(message)

    download_dir, sorted_dir = resolve_photo_directories(options)
    template_path = download_dir / "照片分类模板.xlsx"

    log("开始执行照片分类任务。")
    log(f"照片来源目录：{download_dir}")
    log(f"分类输出目录：{sorted_dir}")
    if not download_dir.exists() and not options.dry_run:
        raise FileNotFoundError(f"照片目录不存在：{download_dir}")

    if cancel_event is not None and cancel_event.is_set():
        log("任务已取消。")
        return WorkflowSummary(
            download_dir=download_dir,
            sorted_dir=sorted_dir,
            template_path=template_path,
            download_result=None,
            template_file_count=0,
            classified_count=0,
            template_created=False,
            cancelled=True,
            dry_run=options.dry_run,
        )

    classification_result: ClassificationResult = apply_classification_from_template(
        source_dir=download_dir,
        target_dir=sorted_dir,
        dry_run=options.dry_run,
        logger=logger,
    )
    log("照片分类任务完成。")
    return WorkflowSummary(
        download_dir=download_dir,
        sorted_dir=sorted_dir,
        template_path=template_path,
        download_result=None,
        template_file_count=0,
        classified_count=classification_result.processed_count,
        template_created=False,
        cancelled=False,
        dry_run=options.dry_run,
        report_path=classification_result.report_path,
    )


def run_photo_template_only(
    options: RunOptions,
    logger: LogFn = None,
    progress_callback: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> WorkflowSummary:
    del progress_callback

    def log(message: str) -> None:
        if logger is not None:
            logger(message)

    download_dir, sorted_dir = resolve_photo_directories(options)
    template_path = download_dir / "照片分类模板.xlsx"

    log("开始执行照片模板生成任务。")
    log(f"照片来源目录：{download_dir}")
    if not download_dir.exists() and not options.dry_run:
        raise FileNotFoundError(f"照片目录不存在：{download_dir}")

    if cancel_event is not None and cancel_event.is_set():
        log("任务已取消。")
        return WorkflowSummary(
            download_dir=download_dir,
            sorted_dir=sorted_dir,
            template_path=template_path,
            download_result=None,
            template_file_count=0,
            classified_count=0,
            template_created=False,
            cancelled=True,
            dry_run=options.dry_run,
        )

    template_result = generate_template(
        source_dir=download_dir,
        dry_run=options.dry_run,
        logger=logger,
    )
    log("照片模板生成任务完成。")
    return WorkflowSummary(
        download_dir=download_dir,
        sorted_dir=sorted_dir,
        template_path=template_result.template_path,
        download_result=None,
        template_file_count=template_result.file_count,
        classified_count=0,
        template_created=template_result.created,
        cancelled=False,
        dry_run=options.dry_run,
    )


def run_workflow(
    options: RunOptions,
    oss_config: Optional[OssConfig] = None,
    logger: LogFn = None,
    progress_callback: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> WorkflowSummary:
    return run_photo_download_and_template(
        options=options,
        oss_config=oss_config,
        logger=logger,
        progress_callback=progress_callback,
        cancel_event=cancel_event,
    )
