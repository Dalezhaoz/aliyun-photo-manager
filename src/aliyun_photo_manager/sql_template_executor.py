from dataclasses import dataclass
from pathlib import Path
import re
from datetime import datetime


@dataclass
class SqlTemplateResult:
    source_path: Path
    exam_code: str
    exam_date: str
    signup_start: str
    signup_end: str
    audit_start: str
    audit_end: str
    sql_content: str


DATETIME_INPUT_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M",
]


def normalize_datetime_text(value: str) -> str:
    cleaned = value.strip()
    if not cleaned:
        raise ValueError("时间不能为空。")
    for fmt in DATETIME_INPUT_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt).strftime("%Y/%m/%d %H:%M:%S")
        except ValueError:
            continue
    raise ValueError(
        f"无法识别时间格式：{value}。请使用 YYYY-MM-DD HH:MM[:SS] 或 YYYY/MM/DD HH:MM[:SS]。"
    )


def validate_exam_date(value: str) -> str:
    cleaned = value.strip()
    if re.fullmatch(r"\d{6}", cleaned):
        return cleaned
    raise ValueError("考试年月请按 YYYYMM 填写，例如 202603。")


def render_sql_template(
    source_path: Path,
    exam_code: str,
    exam_date: str,
    signup_start: str,
    signup_end: str,
    audit_start: str,
    audit_end: str,
) -> SqlTemplateResult:
    if not source_path.exists():
        raise FileNotFoundError(f"SQL 模板不存在：{source_path}")

    content = source_path.read_text(encoding="utf-8")
    normalized_exam_code = exam_code.strip()
    if not normalized_exam_code:
        raise ValueError("考试代码不能为空。")

    normalized_exam_date = validate_exam_date(exam_date)
    normalized_signup_start = normalize_datetime_text(signup_start)
    normalized_signup_end = normalize_datetime_text(signup_end)
    normalized_audit_start = normalize_datetime_text(audit_start)
    normalized_audit_end = normalize_datetime_text(audit_end)

    # 这类模板原本通过 GETDATE()+N 生成时间，这里直接替换成明确时间，避免后续再手改。
    replacements = {
        "@data1": normalized_signup_start,
        "@data2": normalized_signup_end,
        "@data3": normalized_audit_start,
        "@data4": normalized_audit_end,
    }
    for variable_name, value in replacements.items():
        pattern = rf"select\s+{re.escape(variable_name)}\s*=\s*.*"
        content = re.sub(
            pattern,
            f"select {variable_name}='{value}'",
            content,
            flags=re.IGNORECASE,
        )

    content = content.replace("e!s", normalized_exam_code)
    content = content.replace("e!d", normalized_exam_date)

    return SqlTemplateResult(
        source_path=source_path,
        exam_code=normalized_exam_code,
        exam_date=normalized_exam_date,
        signup_start=normalized_signup_start,
        signup_end=normalized_signup_end,
        audit_start=normalized_audit_start,
        audit_end=normalized_audit_end,
        sql_content=content,
    )
