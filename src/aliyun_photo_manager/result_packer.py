from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
import re
import secrets
import string
from typing import Callable, Optional


LogFn = Optional[Callable[[str], None]]


@dataclass
class PackSummary:
    source_dir: Path
    output_path: Path
    file_count: int
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


def build_archive_name_and_password(source_dir: Path) -> tuple[str, str]:
    folder_name = source_dir.name.strip() or "打包结果"
    archive_name = f"{folder_name}.zip"
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
            "source_dir": str(summary.source_dir),
            "source_name": summary.source_dir.name,
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
    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"待打包目录不存在：{source_dir}")

    pyzipper = _load_pyzipper()
    auto_archive_name, auto_password = build_archive_name_and_password(source_dir)
    safe_name = (archive_name or "").strip() or auto_archive_name
    if not safe_name.lower().endswith(".zip"):
        safe_name = f"{safe_name}.zip"
    final_password = (password or "").strip() or auto_password
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / safe_name

    files = [path for path in source_dir.rglob("*") if path.is_file()]
    if not files:
        raise ValueError("所选目录下没有可打包的文件。")

    _log(logger, f"开始打包目录：{source_dir}")
    _log(logger, f"输出文件：{output_path}")

    with pyzipper.AESZipFile(
        output_path,
        "w",
        compression=pyzipper.ZIP_DEFLATED,
        encryption=pyzipper.WZ_AES,
    ) as zip_file:
        zip_file.setpassword(final_password.encode("utf-8"))
        for file_path in files:
            relative_path = file_path.relative_to(source_dir)
            zip_file.write(file_path, arcname=str(relative_path))

    _log(logger, f"打包完成，共写入 {len(files)} 个文件。")
    summary = PackSummary(
        source_dir=source_dir,
        output_path=output_path,
        file_count=len(files),
        password=final_password,
        created_at=created_at,
    )
    save_pack_history(summary)
    return summary
