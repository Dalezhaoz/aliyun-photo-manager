from dataclasses import dataclass
from pathlib import Path
from threading import Event
from typing import Callable, Optional

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
    # 云端前缀会映射到本地目录结构里，避免不同批次文件都堆在根目录。
    if cleaned_prefix:
        return normalized.joinpath(*cleaned_prefix.split("/"), child_name)
    return ensure_child_directory(normalized, child_name)


def resolve_photo_directories(options: RunOptions) -> tuple[Path, Path]:
    # 本地模式直接使用用户选择的目录；云端模式会在目录下自动补业务子目录。
    if options.skip_download:
        return (
            options.download_dir.expanduser().resolve(),
            options.sorted_dir.expanduser().resolve(),
        )
    return (
        build_prefixed_directory(options.download_dir, options.prefix, "下载文件"),
        build_prefixed_directory(options.sorted_dir, options.prefix, "分类结果"),
    )


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

    # 模板始终基于“当前层级”的照片生成，给后续人工补分类信息使用。
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

    # 分类阶段严格依赖模板，不再重新扫描云端或本地目录结构。
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


def run_workflow(
    options: RunOptions,
    oss_config: Optional[OssConfig] = None,
    logger: LogFn = None,
    progress_callback: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> WorkflowSummary:
    def log(message: str) -> None:
        if logger is not None:
            logger(message)

    if options.skip_download:
        download_dir = options.download_dir.expanduser().resolve()
        sorted_dir = options.sorted_dir.expanduser().resolve()
    else:
        download_dir = build_prefixed_directory(options.download_dir, options.prefix, "下载文件")
        sorted_dir = build_prefixed_directory(options.sorted_dir, options.prefix, "分类结果")

    log("开始执行任务。")
    log(f"实际下载目录：{download_dir}")
    log(f"实际分类目录：{sorted_dir}")
    if cancel_event is not None and cancel_event.is_set():
        log("任务已取消。")
        return
    download_result: Optional[DownloadResult] = None
    if not options.skip_download:
        if oss_config is None:
            raise ValueError("未提供 OSS 配置。")
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
        )
        if cancel_event is not None and cancel_event.is_set():
            log("下载已取消。")
            return WorkflowSummary(
                download_dir=download_dir,
                sorted_dir=sorted_dir,
                template_path=download_dir / "照片分类模板.xlsx",
                download_result=download_result,
                template_file_count=0,
                classified_count=0,
                template_created=False,
                cancelled=True,
                dry_run=options.dry_run,
            )
        log("OSS 下载阶段完成。")
    else:
        log("已跳过 OSS 下载，直接使用本地目录分类。")

    if not download_dir.exists() and not options.dry_run:
        raise FileNotFoundError(f"下载目录不存在：{download_dir}")

    if cancel_event is not None and cancel_event.is_set():
        log("任务已取消。")
        return WorkflowSummary(
            download_dir=download_dir,
            sorted_dir=sorted_dir,
            template_path=download_dir / "照片分类模板.xlsx",
            download_result=download_result,
            template_file_count=0,
            classified_count=0,
            template_created=False,
            cancelled=True,
            dry_run=options.dry_run,
        )
    log("开始生成 Excel 分类模板。")
    template_result = generate_template(
        source_dir=download_dir,
        dry_run=options.dry_run,
        logger=logger,
    )

    if options.dry_run:
        log("当前为预览模式，不会真正生成 Excel 或复制文件。")
        log("任务完成。")
        return WorkflowSummary(
            download_dir=download_dir,
            sorted_dir=sorted_dir,
            template_path=template_result.template_path,
            download_result=download_result,
            template_file_count=template_result.file_count,
            classified_count=0,
            template_created=template_result.created,
            dry_run=True,
        )

    if template_result.created:
        log(f"已为当前目录生成模板，共写入 {template_result.file_count} 个文件。")
        log("请先填写 Excel 里的分类信息，再次点击开始执行进行复制分类。")
        log("任务完成。")
        return WorkflowSummary(
            download_dir=download_dir,
            sorted_dir=sorted_dir,
            template_path=template_result.template_path,
            download_result=download_result,
            template_file_count=template_result.file_count,
            classified_count=0,
            template_created=True,
        )

    if cancel_event is not None and cancel_event.is_set():
        log("任务已取消。")
        return WorkflowSummary(
            download_dir=download_dir,
            sorted_dir=sorted_dir,
            template_path=template_result.template_path,
            download_result=download_result,
            template_file_count=template_result.file_count,
            classified_count=0,
            template_created=False,
            cancelled=True,
        )
    log("开始按照 Excel 模板复制分类。")
    classified_count = apply_classification_from_template(
        source_dir=download_dir,
        target_dir=sorted_dir,
        dry_run=options.dry_run,
        logger=logger,
    )
    log("任务完成。")
    return WorkflowSummary(
        download_dir=download_dir,
        sorted_dir=sorted_dir,
        template_path=template_result.template_path,
        download_result=download_result,
        template_file_count=template_result.file_count,
        classified_count=classified_count,
        template_created=False,
    )
