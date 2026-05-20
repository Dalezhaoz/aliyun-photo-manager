from __future__ import annotations

import html
import json
import math
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable
from urllib.parse import quote

from openpyxl import load_workbook


DOC_TYPE_SEAT = "seat"
DOC_TYPE_DOOR = "door"
DOC_TYPE_DESK = "desk"

PLACEHOLDER_PATTERN = re.compile(r"\$\{([^}]+)\}")
TEMPLATE_STORAGE_NAME = ".exam_print_templates.json"
SUPPORTED_IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}


@dataclass
class PrintTemplate:
    name: str
    doc_type: str
    main_html: str
    item_html: str = ""
    settings: dict[str, int | str] = field(default_factory=dict)
    builtin: bool = False


@dataclass(frozen=True)
class PrintDataConfig:
    excel_path: Path
    room_column: str
    seat_column: str
    exam_no_column: str
    site_column: str
    subject_column: str
    unit_column: str
    job_column: str
    photo_dir: Path | None = None
    photo_match_column: str = ""


def template_storage_path() -> Path:
    if getattr(sys, "frozen", False):
        app_root = Path(sys.executable).resolve().parent
        if app_root.name.lower() == "app" and (app_root.parent / "aliyun_photo_manager_launcher.exe").exists():
            return app_root.parent / TEMPLATE_STORAGE_NAME
        return app_root / TEMPLATE_STORAGE_NAME
    return Path(__file__).resolve().parents[2] / TEMPLATE_STORAGE_NAME


def _escape(value: object) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _coerce_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _to_file_url(path: Path) -> str:
    resolved = path.resolve()
    return "file:///" + quote(str(resolved).replace("\\", "/"), safe="/:")


def available_placeholders(headers: Iterable[str]) -> list[str]:
    base = sorted({str(header).strip() for header in headers if str(header).strip()})
    extras = [
        "考场",
        "考点",
        "考试科目",
        "起始考号",
        "结束考号",
        "单位岗位汇总",
        "考生人数",
        "照片",
        "照片标签",
        "面试报到时间",
        "考生列表",
        "座次表表格",
        "桌贴列表",
    ]
    return [f"${{{item}}}" for item in [*base, *extras]]


def load_excel_headers(excel_path: Path) -> list[str]:
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        header_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        return [_coerce_text(value) for value in header_row if _coerce_text(value)]
    finally:
        workbook.close()


def load_excel_records(excel_path: Path) -> list[dict[str, str]]:
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        rows = worksheet.iter_rows(values_only=True)
        header_row = next(rows, ())
        headers = [_coerce_text(value) for value in header_row]
        records: list[dict[str, str]] = []
        for row in rows:
            record: dict[str, str] = {}
            has_value = False
            for index, header in enumerate(headers):
                if not header:
                    continue
                value = _coerce_text(row[index] if index < len(row) else "")
                record[header] = value
                has_value = has_value or bool(value)
            if has_value:
                records.append(record)
        return records
    finally:
        workbook.close()


def load_templates() -> dict[str, list[PrintTemplate]]:
    templates = default_templates()
    storage_path = template_storage_path()
    if not storage_path.exists():
        return templates
    try:
        payload = json.loads(storage_path.read_text(encoding="utf-8"))
    except Exception:
        return templates
    if not isinstance(payload, dict):
        return templates
    for doc_type, items in payload.items():
        if not isinstance(items, list):
            continue
        loaded: list[PrintTemplate] = []
        for item in items:
            if not isinstance(item, dict):
                continue
            name = _coerce_text(item.get("name"))
            if not name:
                continue
            loaded.append(
                PrintTemplate(
                    name=name,
                    doc_type=doc_type,
                    main_html=str(item.get("main_html", "")),
                    item_html=str(item.get("item_html", "")),
                    settings=dict(item.get("settings", {})),
                    builtin=False,
                )
            )
        if loaded:
            builtin_names = {template.name for template in templates.get(doc_type, [])}
            merged = templates.get(doc_type, [])[:]
            merged.extend(template for template in loaded if template.name not in builtin_names)
            templates[doc_type] = merged
    return templates


def save_template(template: PrintTemplate) -> None:
    storage_path = template_storage_path()
    payload: dict[str, list[dict[str, object]]] = {}
    if storage_path.exists():
        try:
            existing = json.loads(storage_path.read_text(encoding="utf-8"))
            if isinstance(existing, dict):
                payload = existing
        except Exception:
            payload = {}
    items = [item for item in payload.get(template.doc_type, []) if item.get("name") != template.name]
    items.append(
        {
            "name": template.name,
            "main_html": template.main_html,
            "item_html": template.item_html,
            "settings": template.settings,
        }
    )
    payload[template.doc_type] = items
    storage_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def default_templates() -> dict[str, list[PrintTemplate]]:
    seat_main = """
<html>
<head>
<style>
body { font-family: 'Microsoft YaHei UI'; margin: 10px; color: #111827; }
.candidate-grid { width: 100%; border-collapse: collapse; table-layout: fixed; }
.candidate-grid > tr > td, .candidate-grid > tbody > tr > td { border: 1px solid #444; padding: 0; height: 94px; vertical-align: top; }
.candidate-card { width: 100%; height: 94px; border-collapse: collapse; table-layout: fixed; font-size: 12px; line-height: 1.35; }
.candidate-card td { border: none; padding: 0; vertical-align: top; }
.candidate-card .photo-cell { width: 68px; }
.candidate-card .info-cell { padding: 3px 5px; }
.photo { width: 68px; height: 94px; object-fit: cover; display: block; }
.footer-title { margin: 10px 10px 0; padding: 8px 12px; text-align: center; color: #c62828; background: #f5dada; border-radius: 8px; font-size: 18px; }
</style>
</head>
<body>
  ${考生列表}
  <div class="footer-title">2026年度周村区（文昌湖区）事业单位公开招聘综合类岗位人员面试</div>
</body>
</html>
""".strip()
    seat_item = """
<table class="candidate-card">
  <tr>
    <td class="photo-cell">${照片标签}</td>
    <td class="info-cell">
      <div>姓名:${姓名}</div>
      <div>身份证号:</div>
      <div>${身份证号}</div>
      <div>报考单位:${报考单位}</div>
      <div>报考职位:${报考岗位}</div>
    </td>
  </tr>
</table>
""".strip()
    door_main = """
<html>
<head>
<style>
body { font-family: 'Microsoft YaHei UI'; margin: 32px; color: #111827; }
.door { border: 2px solid #cbd5e1; padding: 24px 28px; }
.title { text-align: center; font-size: 34px; font-weight: 700; margin-bottom: 18px; }
.line { font-size: 18px; line-height: 2; }
</style>
</head>
<body>
  <div class="door">
    <div class="title">考场门贴</div>
    <div class="line">考点：${考点}</div>
    <div class="line">考场：${考场}</div>
    <div class="line">考试科目：${考试科目}</div>
    <div class="line">准考证号：${起始考号} - ${结束考号}</div>
    <div class="line">单位岗位：${单位岗位汇总}</div>
    <div class="line">人数：${考生人数}</div>
  </div>
</body>
</html>
""".strip()
    desk_main = """
<html>
<head>
<style>
body { font-family: 'Microsoft YaHei UI'; margin: 18px; color: #111827; }
.page { page-break-after: always; }
.page:last-child { page-break-after: auto; }
.desk-grid { width: 100%; border-collapse: separate; border-spacing: 10px; table-layout: fixed; }
.desk-grid td { border: 1px dashed #94a3b8; padding: 10px; vertical-align: top; height: 100px; }
.desk-card { font-size: 12px; line-height: 1.6; }
.desk-name { font-size: 16px; font-weight: 700; margin-bottom: 6px; }
</style>
</head>
<body>
  ${桌贴列表}
</body>
</html>
""".strip()
    desk_item = """
<div class="desk-card">
  <div class="desk-name">${姓名}</div>
  <div>座号：${座号}</div>
  <div>准考证号：${考号}</div>
  <div>${报考单位}${报考岗位}</div>
</div>
""".strip()
    return {
        DOC_TYPE_SEAT: [
            PrintTemplate("面试表样网格", DOC_TYPE_SEAT, seat_main, seat_item, {"columns": 5, "layout": "candidate_grid"}, builtin=True),
        ],
        DOC_TYPE_DOOR: [
            PrintTemplate("标准门贴", DOC_TYPE_DOOR, door_main, "", {}, builtin=True),
        ],
        DOC_TYPE_DESK: [
            PrintTemplate("桌贴 10 人版", DOC_TYPE_DESK, desk_main, desk_item, {"columns": 2, "items_per_page": 10}, builtin=True),
            PrintTemplate("桌贴 30 人版", DOC_TYPE_DESK, desk_main, desk_item, {"columns": 5, "items_per_page": 30}, builtin=True),
        ],
    }


def room_values(records: Iterable[dict[str, str]], room_column: str) -> list[str]:
    values = sorted({_coerce_text(record.get(room_column, "")) for record in records if _coerce_text(record.get(room_column, ""))})
    return values


def build_room_context(records: list[dict[str, str]], config: PrintDataConfig, room_value: str) -> tuple[dict[str, str], list[dict[str, str]]]:
    if config.room_column:
        room_records = [record for record in records if _coerce_text(record.get(config.room_column, "")) == room_value]
    else:
        room_records = list(records)
    if not room_records:
        raise ValueError("没有匹配到任何考生。")

    sorted_records = sorted(room_records, key=lambda item: _sort_key(item, config))
    for record in sorted_records:
        record["照片"] = _match_photo(record, config)
        record["座号"] = record.get(config.seat_column, "") if config.seat_column else record.get("座号", "")
        record["考号"] = record.get(config.exam_no_column, "") if config.exam_no_column else record.get("考号", "")
        record["报考单位"] = record.get(config.unit_column, "") if config.unit_column else record.get("报考单位", "")
        record["报考岗位"] = record.get(config.job_column, "") if config.job_column else record.get("报考岗位", "")

    exam_numbers = [item.get(config.exam_no_column, "") for item in sorted_records if item.get(config.exam_no_column, "")]
    subjects = _unique_join(item.get(config.subject_column, "") for item in sorted_records)
    units = _unique_join(
        _join_parts(item.get(config.unit_column, ""), item.get(config.job_column, ""), separator="")
        for item in sorted_records
    )
    first = sorted_records[0]
    context = {key: _escape(value) for key, value in first.items()}
    context.update(
        {
            "考场": _escape(room_value),
            "考点": _escape(first.get(config.site_column, "")),
            "考试科目": _escape(subjects),
            "起始考号": _escape(min(exam_numbers, default="")),
            "结束考号": _escape(max(exam_numbers, default="")),
            "单位岗位汇总": _escape(units),
            "考生人数": _escape(len(sorted_records)),
        }
    )
    context["座号"] = _escape(first.get(config.seat_column, ""))
    context["考号"] = _escape(first.get(config.exam_no_column, ""))
    context["报考单位"] = _escape(first.get(config.unit_column, ""))
    context["报考岗位"] = _escape(first.get(config.job_column, ""))
    return context, sorted_records


def render_document(
    template: PrintTemplate,
    records: list[dict[str, str]],
    config: PrintDataConfig,
    room_value: str,
    columns: int,
    items_per_page: int,
) -> str:
    context, room_records = build_room_context(records, config, room_value)
    if template.doc_type == DOC_TYPE_SEAT:
        if str(template.settings.get("layout", "")) == "candidate_grid":
            candidate_grid = _build_candidate_grid(room_records, template.item_html, context, columns)
            context["考生列表"] = candidate_grid
            context["座次表表格"] = candidate_grid
        elif str(template.settings.get("layout", "")) == "signin":
            signin_table = _build_signin_table(room_records, room_context=context)
            context["考生列表"] = signin_table
            context["座次表表格"] = signin_table
        else:
            context["座次表表格"] = _build_seat_table(room_records, template.item_html, context, columns)
    elif template.doc_type == DOC_TYPE_DESK:
        context["桌贴列表"] = _build_desk_pages(room_records, template.item_html, context, columns, items_per_page)
    return render_html_template(template.main_html, context)


def render_html_template(template_html: str, context: dict[str, str]) -> str:
    def replacer(match: re.Match[str]) -> str:
        key = match.group(1).strip()
        return str(context.get(key, ""))

    return PLACEHOLDER_PATTERN.sub(replacer, template_html)


def export_rendered_html(output_path: Path, rendered_html: str) -> None:
    output_path.write_text(rendered_html, encoding="utf-8")


def _join_parts(left: str, right: str, separator: str = " ") -> str:
    parts = [part for part in (left.strip(), right.strip()) if part]
    return separator.join(parts)


def _unique_join(values: Iterable[str]) -> str:
    seen: list[str] = []
    for value in values:
        text = _coerce_text(value)
        if text and text not in seen:
            seen.append(text)
    return "、".join(seen)


def _match_photo(record: dict[str, str], config: PrintDataConfig) -> str:
    if config.photo_dir is None or not config.photo_dir.exists() or not config.photo_match_column:
        return ""
    match_value = _normalize_match_text(record.get(config.photo_match_column, ""))
    if not match_value:
        return ""
    for file_path in sorted(config.photo_dir.rglob("*")):
        if not file_path.is_file() or file_path.suffix.lower() not in SUPPORTED_IMAGE_SUFFIXES:
            continue
        stem = _normalize_match_text(file_path.stem)
        if stem == match_value or stem.startswith(match_value) or match_value in stem:
            return _to_file_url(file_path)
    return ""


def _normalize_match_text(value: object) -> str:
    text = _coerce_text(value).upper()
    return re.sub(r"[^0-9A-Z]", "", text)


def _sort_key(record: dict[str, str], config: PrintDataConfig) -> tuple[int, str, str]:
    seat_text = _coerce_text(record.get(config.seat_column, ""))
    seat_number = int(seat_text) if seat_text.isdigit() else math.inf
    return seat_number, _coerce_text(record.get(config.exam_no_column, "")), _coerce_text(record.get("姓名", ""))


def _merge_context(record: dict[str, str], room_context: dict[str, str]) -> dict[str, str]:
    merged = dict(room_context)
    merged.update({key: _escape(value) for key, value in record.items()})
    merged["座号"] = _escape(record.get("座号") or record.get("seat", ""))
    merged["考号"] = _escape(record.get("考号") or record.get("准考证号") or record.get("exam_no", ""))
    merged["报考单位"] = _escape(record.get("报考单位") or record.get("单位") or "")
    merged["报考岗位"] = _escape(record.get("报考岗位") or record.get("岗位") or "")
    if record.get("照片"):
        merged["照片"] = record["照片"]
        merged["照片标签"] = f'<img class="photo" src="{record["照片"]}" />'
    else:
        merged["照片标签"] = ""
    return merged


def _build_seat_table(records: list[dict[str, str]], item_html: str, room_context: dict[str, str], columns: int) -> str:
    safe_columns = max(1, columns)
    rows: list[str] = []
    cells: list[str] = []
    for record in records:
        merged_context = _merge_context(record, room_context)
        cells.append(f"<td>{render_html_template(item_html, merged_context)}</td>")
        if len(cells) == safe_columns:
            rows.append("<tr>" + "".join(cells) + "</tr>")
            cells = []
    if cells:
        while len(cells) < safe_columns:
            cells.append("<td></td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<table class="seat-table">' + "".join(rows) + "</table>"


def _build_candidate_grid(records: list[dict[str, str]], item_html: str, room_context: dict[str, str], columns: int) -> str:
    safe_columns = max(1, columns)
    rows: list[str] = []
    cells: list[str] = []
    for record in records:
        merged_context = _merge_context(record, room_context)
        cells.append(
            '<td style="border:1px solid #444;padding:0;height:94px;vertical-align:top;">'
            f"{render_html_template(item_html, merged_context)}"
            "</td>"
        )
        if len(cells) == safe_columns:
            rows.append("<tr>" + "".join(cells) + "</tr>")
            cells = []
    if cells:
        while len(cells) < safe_columns:
            cells.append('<td style="border:1px solid #444;padding:0;height:94px;"></td>')
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<table class="candidate-grid">' + "".join(rows) + "</table>"


def _build_signin_table(records: list[dict[str, str]], room_context: dict[str, str]) -> str:
    header = """
<tr>
  <th style="width: 7%;">序号</th>
  <th style="width: 9%;">照片</th>
  <th style="width: 9%;">姓名</th>
  <th style="width: 11%;">考号</th>
  <th style="width: 15%;">报考单位</th>
  <th style="width: 15%;">报考岗位</th>
  <th style="width: 18%;">身份证号</th>
  <th style="width: 9%;">报到时间</th>
  <th style="width: 7%;">签到</th>
</tr>
""".strip()
    rows: list[str] = []
    for index, record in enumerate(records, start=1):
        merged = _merge_context(record, room_context)
        photo_url = record.get("照片", "")
        photo_cell = f'<img class="signin-photo" src="{photo_url}" />' if photo_url else ""
        rows.append(
            "<tr>"
            f"<td>{index}</td>"
            f"<td>{photo_cell}</td>"
            f"<td>{merged.get('姓名', '')}</td>"
            f"<td>{merged.get('考号', '')}</td>"
            f"<td>{merged.get('报考单位', '')}</td>"
            f"<td>{merged.get('报考岗位', '')}</td>"
            f"<td>{merged.get('身份证号', '')}</td>"
            f"<td>{merged.get('面试报到时间', '')}</td>"
            '<td class="signin-signature"></td>'
            "</tr>"
        )
    return '<table class="signin-table">' + header + "".join(rows) + "</table>"


def _build_desk_pages(
    records: list[dict[str, str]],
    item_html: str,
    room_context: dict[str, str],
    columns: int,
    items_per_page: int,
) -> str:
    safe_columns = max(1, columns)
    safe_items = max(1, items_per_page)
    pages: list[str] = []
    for page_start in range(0, len(records), safe_items):
        chunk = records[page_start : page_start + safe_items]
        rows: list[str] = []
        cells: list[str] = []
        for record in chunk:
            merged_context = _merge_context(record, room_context)
            cells.append(f"<td>{render_html_template(item_html, merged_context)}</td>")
            if len(cells) == safe_columns:
                rows.append("<tr>" + "".join(cells) + "</tr>")
                cells = []
        if cells:
            while len(cells) < safe_columns:
                cells.append("<td></td>")
            rows.append("<tr>" + "".join(cells) + "</tr>")
        pages.append(f'<div class="page"><table class="desk-grid">{"".join(rows)}</table></div>')
    return "".join(pages)
