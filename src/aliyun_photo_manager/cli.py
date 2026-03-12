import argparse
from pathlib import Path

from .app import RunOptions, run_workflow
from .config import load_oss_config


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从阿里云 OSS 下载文件，生成 Excel 模板，并按模板复制分类。"
    )
    parser.add_argument(
        "--prefix",
        default="",
        help="OSS 对象前缀，例如 photos/2025/",
    )
    parser.add_argument(
        "--download-dir",
        type=Path,
        required=True,
        help="下载目录",
    )
    parser.add_argument(
        "--sorted-dir",
        type=Path,
        required=True,
        help="分类输出目录",
    )
    parser.add_argument(
        "--skip-download",
        action="store_true",
        help="跳过 OSS 下载，只对本地下载目录做分类",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只打印计划，不真正下载、生成 Excel 和复制分类",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    options = RunOptions(
        prefix=args.prefix,
        download_dir=args.download_dir,
        sorted_dir=args.sorted_dir,
        skip_download=args.skip_download,
        dry_run=args.dry_run,
    )

    config = None if args.skip_download else load_oss_config()
    run_workflow(options=options, oss_config=config)


if __name__ == "__main__":
    main()
