from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Sequence


LogFn = Optional[Callable[[str], None]]


@dataclass
class PhoneDecryptOptions:
    server: str
    port: int
    username: str
    password: str
    signup_database: str
    phone_database: str
    candidate_table: str
    candidate_filter_mode: str = "all"
    candidate_id_cards: Optional[Sequence[str]] = None
    output_path: Optional[Path] = None


@dataclass
class PhoneDecryptRecord:
    primary_key: str
    id_card: str
    province: str
    encrypted_phone: str
    decrypted_phone: str
    status: str
    note: str


@dataclass
class PhoneDecryptSummary:
    signup_database: str
    phone_database: str
    candidate_table: str
    output_path: Path
    total_rows: int
    matched_info_rows: int
    decrypted_rows: int
    updated_rows: int
    skipped_rows: int
    failed_rows: int
    backend_name: str
    records: List[PhoneDecryptRecord]


def _log(logger: LogFn, message: str) -> None:
    if logger is not None:
        logger(message)


def _normalize(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


# ── Excel reading (for filter ID card list) ──────────────────────────────

def _load_openpyxl():
    try:
        from openpyxl import Workbook, load_workbook
    except ImportError as exc:
        raise ImportError("缺少依赖 openpyxl，请先执行 `pip install -r requirements.txt`。") from exc
    return Workbook, load_workbook


def _load_xlrd():
    try:
        import xlrd
    except ImportError as exc:
        raise ImportError("缺少依赖 xlrd，请先执行 `pip install -r requirements.txt`。") from exc
    return xlrd


def _read_sheet_matrix(file_path: Path) -> List[List[str]]:
    suffix = file_path.suffix.lower()
    if suffix == ".xlsx":
        _, load_workbook = _load_openpyxl()
        workbook = load_workbook(file_path, data_only=True)
        worksheet = workbook.worksheets[0]
        return [[_normalize(value) for value in row] for row in worksheet.iter_rows(values_only=True)]
    if suffix == ".xls":
        xlrd = _load_xlrd()
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        return [
            [_normalize(sheet.cell_value(row_index, col_index)) for col_index in range(sheet.ncols)]
            for row_index in range(sheet.nrows)
        ]
    raise ValueError("名单文件仅支持 `.xlsx` 或 `.xls`。")


def _trim_row(row: List[str]) -> List[str]:
    trimmed = list(row)
    while trimmed and trimmed[-1] == "":
        trimmed.pop()
    return trimmed


def _detect_header_row(matrix: List[List[str]]) -> int:
    best_index = 0
    best_score = -1
    for row_index in range(min(len(matrix), 10)):
        row = _trim_row(matrix[row_index])
        non_empty = [cell for cell in row if cell]
        if not non_empty:
            continue
        score = len(non_empty) * 10 + len(set(non_empty))
        if len(non_empty) == 1:
            score -= 20
        if score > best_score:
            best_score = score
            best_index = row_index
    return best_index


def load_filter_id_cards(file_path: Path) -> List[str]:
    if not file_path.exists():
        raise FileNotFoundError(f"未找到名单文件：{file_path}")
    matrix = _read_sheet_matrix(file_path)
    if not matrix:
        raise ValueError("名单文件为空。")
    header_index = _detect_header_row(matrix)
    headers = [header or f"未命名列{index + 1}" for index, header in enumerate(_trim_row(matrix[header_index]))]
    candidates = {"身份证号", "证件号码", "身份证", "sfzh"}
    target_index = next((index for index, header in enumerate(headers) if header.strip() in candidates), None)
    if target_index is None:
        raise ValueError("名单文件缺少"身份证号"列。")
    id_cards: List[str] = []
    seen = set()
    width = len(headers)
    for raw_row in matrix[header_index + 1 :]:
        row = list(raw_row[:width])
        if len(row) < width:
            row.extend([""] * (width - len(row)))
        id_card = _normalize(row[target_index])
        if id_card and id_card not in seen:
            id_cards.append(id_card)
            seen.add(id_card)
    return id_cards


# ── Helper path finding ──────────────────────────────────────────────────

def _candidate_helper_paths() -> List[Path]:
    candidates: List[Path] = []
    env_path = os.environ.get("PHONE_DECRYPT_HELPER", "").strip()
    if env_path:
        candidates.append(Path(env_path))

    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        candidates.append(Path(meipass) / "PhoneDecryptHelper.exe")

    project_root = Path(__file__).resolve().parents[2]
    candidates.extend(
        [
            project_root / "PhoneDecryptHelper.exe",
            project_root / "tools" / "PhoneDecryptHelper" / "bin" / "Release" / "net8.0-windows" / "win-x86" / "publish" / "PhoneDecryptHelper.exe",
            project_root / "tools" / "PhoneDecryptHelper" / "bin" / "Release" / "net8.0-windows" / "PhoneDecryptHelper.exe",
            Path.cwd() / "PhoneDecryptHelper.exe",
            Path(sys.executable).resolve().parent / "PhoneDecryptHelper.exe",
        ]
    )
    seen = set()
    result: List[Path] = []
    for item in candidates:
        resolved = item.expanduser().resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        result.append(resolved)
    return result


def _find_helper_path() -> Path | None:
    for candidate in _candidate_helper_paths():
        if candidate.exists():
            return candidate
    return None


# ── Call helper subprocess ───────────────────────────────────────────────

def _call_helper(payload: dict, logger: LogFn = None) -> dict:
    helper_path = _find_helper_path()
    if helper_path is None:
        searched = "，".join(str(item) for item in _candidate_helper_paths())
        raise RuntimeError(
            "未找到 PhoneDecryptHelper.exe。"
            " 请先在 Windows 上编译 32 位 helper（dotnet publish -r win-x86 --self-contained true），"
            " 或设置环境变量 PHONE_DECRYPT_HELPER。"
            f" 已搜索：{searched}"
        )

    with tempfile.TemporaryDirectory(prefix="phone_decrypt_") as temp_dir:
        temp_path = Path(temp_dir)
        input_path = temp_path / "input.json"
        output_path = temp_path / "output.json"
        input_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        _log(logger, f"调用 helper：{helper_path}")

        run_kwargs: dict = dict(
            capture_output=True,
            text=True,
            check=False,
            stdin=subprocess.DEVNULL,
        )
        if sys.platform == "win32":
            run_kwargs["creationflags"] = getattr(
                subprocess, "CREATE_NO_WINDOW", 0x08000000
            )
        result = subprocess.run(
            [str(helper_path), str(input_path), str(output_path)],
            **run_kwargs,
        )
        if result.returncode != 0:
            details = [f"helper 返回码: {result.returncode}"]
            if result.stderr.strip():
                details.append(f"stderr: {result.stderr.strip()}")
            if result.stdout.strip():
                details.append(f"stdout: {result.stdout.strip()}")
            raise RuntimeError("PhoneDecryptHelper 执行失败\n" + "\n".join(details))
        if not output_path.exists():
            raise RuntimeError("PhoneDecryptHelper 未生成输出文件。")
        return json.loads(output_path.read_text(encoding="utf-8"))


# ── Excel export ─────────────────────────────────────────────────────────

def _export_phone_report(records: List[PhoneDecryptRecord], output_path: Path) -> None:
    Workbook, _ = _load_openpyxl()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "电话解密结果"
    sheet.append(["主键编号", "身份证号", "考区", "加密串", "解密后电话", "状态", "备注"])
    for item in records:
        sheet.append(
            [
                item.primary_key,
                item.id_card,
                item.province,
                item.encrypted_phone,
                item.decrypted_phone,
                item.status,
                item.note,
            ]
        )
    workbook.save(output_path)


def _build_output_path(options: PhoneDecryptOptions) -> Path:
    if options.output_path is not None:
        return options.output_path
    return Path.cwd() / f"{options.candidate_table}_电话解密结果.xlsx"


# ── Main entry point ─────────────────────────────────────────────────────

def run_phone_decrypt(options: PhoneDecryptOptions, logger: LogFn = None) -> PhoneDecryptSummary:
    if not options.server.strip():
        raise ValueError("请输入服务器地址。")
    if not options.username.strip() or not options.password.strip():
        raise ValueError("请输入数据库用户名和密码。")
    if not options.signup_database.strip():
        raise ValueError("请输入报名数据库名称。")
    if not options.candidate_table.strip():
        raise ValueError("请输入考生表名称。")

    output_path = _build_output_path(options)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    _log(logger, f"开始电话解密：{options.signup_database}.{options.candidate_table}")

    payload: dict = {
        "mode": "full",
        "server": options.server,
        "port": options.port,
        "username": options.username,
        "password": options.password,
        "signupDatabase": options.signup_database,
        "phoneDatabase": options.phone_database or options.signup_database,
        "candidateTable": options.candidate_table,
        "filterMode": options.candidate_filter_mode,
        "idCards": list(options.candidate_id_cards) if options.candidate_id_cards else [],
        "updateResults": True,
    }

    data = _call_helper(payload, logger=logger)

    # relay logs from helper
    for line in data.get("logs", []):
        _log(logger, f"[helper] {line}")

    if data.get("error"):
        raise RuntimeError(f"helper 报错：{data['error']}")

    backend_name = data.get("backendName", "unknown")
    _log(logger, f"解密组件：{backend_name}")

    records: List[PhoneDecryptRecord] = []
    for item in data.get("records", []):
        records.append(
            PhoneDecryptRecord(
                primary_key=str(item.get("primaryKey", "")),
                id_card=str(item.get("idCard", "")),
                province=str(item.get("province", "")),
                encrypted_phone=str(item.get("encryptedPhone", "")),
                decrypted_phone=str(item.get("decryptedPhone", "")),
                status=str(item.get("status", "")),
                note=str(item.get("note", "")),
            )
        )

    _export_phone_report(records, output_path)
    _log(logger, f"已导出电话解密结果清单：{output_path}")

    return PhoneDecryptSummary(
        signup_database=options.signup_database,
        phone_database=options.phone_database or options.signup_database,
        candidate_table=options.candidate_table,
        output_path=output_path,
        total_rows=int(data.get("totalRows", 0)),
        matched_info_rows=int(data.get("matchedInfoRows", 0)),
        decrypted_rows=int(data.get("decryptedRows", 0)),
        updated_rows=int(data.get("updatedRows", 0)),
        skipped_rows=int(data.get("skippedRows", 0)),
        failed_rows=int(data.get("failedRows", 0)),
        backend_name=backend_name,
        records=records,
    )
