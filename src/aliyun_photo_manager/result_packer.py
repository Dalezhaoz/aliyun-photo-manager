from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
import re
import secrets
import shutil
import string
from typing import Callable, Optional


LogFn = Optional[Callable[[str], None]]


@dataclass
class PackSummary:
    source_path: Path
    output_path: Path
    file_count: int
    password: str
    created_at: str


@dataclass
class BatchPackSummary:
    source_dir: Path
    output_dir: Path
    summaries: list[PackSummary]
    password: str
    created_at: str


HISTORY_FILE = Path(__file__).resolve().parents[2] / ".pack_history.json"


def _log(logger: LogFn, message: str) -> None:
    if logger is not None:
        logger(message)


def _load_pyzipper():
    try:
        import pyzipper
    except ImportError as exc:
        raise ImportError(
            "缺少依赖 pyzipper，请先执行 `pip install -r requirements.txt`。"
        ) from exc
    return pyzipper


def build_archive_name_and_password(source_path: Path) -> tuple[str, str]:
    source_name = source_path.name.strip() or "打包结果"
    archive_name = f"{source_name}.zip"
    random_suffix = "".join(secrets.choice(string.ascii_uppercase + string.digits) for _ in range(4))
    password = f"{datetime.now().strftime('%y%m%d')}{random_suffix}"
    return archive_name, password


def load_pack_history() -> list[dict]:
    if not HISTORY_FILE.exists():
        return []
    try:
        data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    return []


def save_pack_history(summary: PackSummary) -> None:
    history = load_pack_history()
    history.insert(
        0,
        {
            "archive_name": summary.output_path.name,
            "output_path": str(summary.output_path),
            "source_path": str(summary.source_path),
            "source_name": summary.source_path.name,
            "password": summary.password,
            "created_at": summary.created_at,
        },
    )
    HISTORY_FILE.write_text(
        json.dumps(history[:500], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def query_pack_history(keyword: str) -> list[dict]:
    normalized = keyword.strip().lower()
    if not normalized:
        return load_pack_history()[:20]
    matches = []
    for item in load_pack_history():
        haystack = " ".join(
            [
                str(item.get("archive_name", "")),
                str(item.get("source_name", "")),
                str(item.get("source_path", "")),
                str(item.get("source_dir", "")),
                str(item.get("output_path", "")),
                str(item.get("password", "")),
            ]
        ).lower()
        if normalized in haystack:
            matches.append(item)
    return matches[:20]


def pack_encrypted_folder(
    source_dir: Path,
    output_dir: Path,
    archive_name: Optional[str] = None,
    password: Optional[str] = None,
    logger: LogFn = None,
) -> PackSummary:
    if not source_dir.exists():
        raise FileNotFoundError(f"待打包文件或文件夹不存在：{source_dir}")

    pyzipper = _load_pyzipper()
    auto_archive_name, auto_password = build_archive_name_and_password(source_dir)
    safe_name = (archive_name or "").strip() or auto_archive_name
    if not safe_name.lower().endswith(".zip"):
        safe_name = f"{safe_name}.zip"
    final_password = (password or "").strip() or auto_password
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / safe_name

    if source_dir.is_file():
        files = [(source_dir, source_dir.name)]
        _log(logger, f"开始打包文件：{source_dir}")
    elif source_dir.is_dir():
        files = [
            (path, str(path.relative_to(source_dir)))
            for path in source_dir.rglob("*")
            if path.is_file()
        ]
        if not files:
            raise ValueError("所选文件夹下没有可打包的文件。")
        _log(logger, f"开始打包文件夹：{source_dir}")
    else:
        raise ValueError(f"不支持的打包对象：{source_dir}")

    _log(logger, f"输出文件：{output_path}")

    with pyzipper.AESZipFile(
        output_path,
        "w",
        compression=pyzipper.ZIP_DEFLATED,
        encryption=pyzipper.WZ_AES,
    ) as zip_file:
        zip_file.setpassword(final_password.encode("utf-8"))
        for file_path, archive_member in files:
            zip_file.write(file_path, arcname=archive_member)

    _log(logger, f"打包完成，共写入 {len(files)} 个文件。")
    summary = PackSummary(
        source_path=source_dir,
        output_path=output_path,
        file_count=len(files),
        password=final_password,
        created_at=created_at,
    )
    save_pack_history(summary)
    return summary


def pack_subfolders_separately(
    source_dir: Path,
    output_dir: Path,
    password: Optional[str] = None,
    logger: LogFn = None,
) -> BatchPackSummary:
    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"待打包目录不存在：{source_dir}")

    subdirs = sorted(
        [child for child in source_dir.iterdir() if child.is_dir()],
        key=lambda p: p.name,
    )
    root_files = sorted(
        [child for child in source_dir.iterdir() if child.is_file()],
        key=lambda p: p.name,
    )

    random_suffix = "".join(secrets.choice(string.ascii_uppercase + string.digits) for _ in range(4))
    shared_password = (password or "").strip() or f"{datetime.now().strftime('%y%m%d')}{random_suffix}"
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 输出到以来源文件夹命名的子目录，避免直接占满输出目录
    output_subdir = output_dir / f"{source_dir.name}_打包结果"
    output_subdir.mkdir(parents=True, exist_ok=True)
    summaries: list[PackSummary] = []

    pyzipper = _load_pyzipper()

    def _pack_one(source: Path, archive_name: str) -> None:
        if source.is_dir():
            files = [
                (path, str(path.relative_to(source)))
                for path in source.rglob("*")
                if path.is_file()
            ]
        else:
            files = [(source, source.name)]

        if not files:
            _log(logger, f"跳过空文件夹：{source.name}" if source.is_dir() else f"跳过：{source.name}")
            return

        output_path = output_subdir / archive_name
        _log(logger, f"打包 {source.name}，共 {len(files)} 个文件 → {archive_name}")

        with pyzipper.AESZipFile(
            output_path,
            "w",
            compression=pyzipper.ZIP_DEFLATED,
            encryption=pyzipper.WZ_AES,
        ) as zip_file:
            zip_file.setpassword(shared_password.encode("utf-8"))
            for file_path, arcname in files:
                zip_file.write(file_path, arcname=arcname)

        summary = PackSummary(
            source_path=source,
            output_path=output_path,
            file_count=len(files),
            password=shared_password,
            created_at=created_at,
        )
        save_pack_history(summary)
        summaries.append(summary)

    # 先打包根级文件
    for file_path in root_files:
        _pack_one(file_path, f"{file_path.stem}.zip")

    # 再打包子文件夹
    for subdir in subdirs:
        _pack_one(subdir, f"{subdir.name}.zip")

    _log(logger, f"子文件夹打包完成：共 {len(summaries)} 个 zip，输出到 {output_subdir}")

    # 最后把整个 _打包结果 目录打成 zip
    final_zip_path = output_dir / f"{source_dir.name}.zip"
    _log(logger, f"正在打包最终结果 → {final_zip_path.name}")
    files_in_subdir = [
        (path, str(path.relative_to(output_subdir)))
        for path in sorted(output_subdir.rglob("*"))
        if path.is_file()
    ]
    with pyzipper.AESZipFile(
        final_zip_path,
        "w",
        compression=pyzipper.ZIP_DEFLATED,
        encryption=pyzipper.WZ_AES,
    ) as zip_file:
        zip_file.setpassword(shared_password.encode("utf-8"))
        for file_path, arcname in files_in_subdir:
            zip_file.write(file_path, arcname=arcname)

    _log(logger, f"全部完成：{final_zip_path.name}，密码：{shared_password}")

    # 清理临时目录
    shutil.rmtree(output_subdir, ignore_errors=True)

    return BatchPackSummary(
        source_dir=source_dir,
        output_dir=final_zip_path.parent,
        summaries=summaries,
        password=shared_password,
        created_at=created_at,
    )
