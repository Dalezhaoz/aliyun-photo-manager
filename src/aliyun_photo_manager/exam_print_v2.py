"""考场文件打印 V2 — 纯 reportlab PDF 导出引擎。"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from .exam_printing import (
    _draw_photo,
    _join_parts,
    _layout_positions,
    _natural_sort_key,
    _normalize_match_text,
    _photo_file_map,
    _unique_join,
    _wrap_pdf_text,
    load_excel_headers,
    load_excel_records,
)

PLACEHOLDER_PATTERN = re.compile(r"\$\{([^}]+)\}")


# ── Data Models ──────────────────────────────────────────────────────────────


@dataclass
class ColumnMapping:
    """Excel 列 → 语义角色映射。"""
    room_column: str = ""
    seat_column: str = ""
    exam_no_column: str = ""
    site_column: str = ""
    subject_column: str = ""
    unit_column: str = ""
    job_column: str = ""
    photo_match_column: str = ""
    sort_column: str = ""


@dataclass
class GridSettings:
    """网格布局设置。custom_heights 非空时覆盖 columns + rows。"""
    columns: int = 5
    rows: int = 6
    custom_heights: str = ""
    start_corner: str = "top_left"
    fill_direction: str = "column"
    snake: bool = False


@dataclass
class FieldConfig:
    """考生卡片字段配置。fields 为用户手动输入的列名列表。"""
    fields: list[str] = field(default_factory=list)


@dataclass
class SigninSettings:
    """面试签到表设置。"""
    title: str = ""
    subtitle: str = ""
    title_align: str = "center"
    grid: GridSettings = field(default_factory=GridSettings)
    field_config: FieldConfig = field(default_factory=FieldConfig)
    photo_width: float = 52


@dataclass
class SeatSettings:
    """笔试座次表设置。"""
    title: str = ""
    subtitle: str = ""
    title_align: str = "center"
    show_room_overview: bool = True
    show_proctor_signature: bool = True
    grid: GridSettings = field(default_factory=GridSettings)
    field_config: FieldConfig = field(default_factory=FieldConfig)
    photo_width: float = 52


@dataclass
class DoorSettings:
    """门贴设置。display_fields 为要显示的汇总字段名列表。"""
    display_fields: list[str] = field(
        default_factory=lambda: ["考场", "考点", "考试科目", "准考证号起止", "考生人数"]
    )


@dataclass
class DeskSettings:
    """桌贴设置。"""
    columns: int = 2
    items_per_page: int = 10
    field_config: FieldConfig = field(default_factory=FieldConfig)
    photo_width: float = 52
    per_room: bool = False


# ── Shared Utilities ─────────────────────────────────────────────────────────


def _coerce(value: object) -> str:
    return "" if value is None else str(value).strip()


def _register_font() -> str:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass
    return "STSong-Light"


def find_photo(
    record: dict[str, str],
    match_column: str,
    photo_map: dict[str, Path],
) -> Path | None:
    """根据 match_column 在 photo_map 中查找照片文件。"""
    if not match_column:
        return None
    match_value = _normalize_match_text(record.get(match_column, ""))
    if not match_value:
        return None
    if match_value in photo_map:
        return photo_map[match_value]
    for stem, file_path in photo_map.items():
        if stem.startswith(match_value) or match_value in stem:
            return file_path
    return None


def sort_records(
    records: list[dict[str, str]],
    sort_column: str,
) -> list[dict[str, str]]:
    """按指定列自然排序。"""
    if sort_column:
        return sorted(records, key=lambda r: _natural_sort_key(r.get(sort_column, "")))
    return list(records)


def group_by_room(
    records: list[dict[str, str]],
    room_column: str,
) -> list[tuple[str, list[dict[str, str]]]]:
    """按考场分组，保持原始顺序。返回 [(考场名, 记录列表)]。"""
    if not room_column:
        return [("", records)]
    groups: dict[str, list[dict[str, str]]] = {}
    order: list[str] = []
    for record in records:
        key = _coerce(record.get(room_column, ""))
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(record)
    return [(key, groups[key]) for key in order]


def compute_room_summary(
    records: list[dict[str, str]],
    mapping: ColumnMapping,
) -> dict[str, str]:
    """计算考场汇总信息。"""
    if not records:
        return {}
    first = records[0]
    exam_nos = [
        r.get(mapping.exam_no_column, "")
        for r in records
        if mapping.exam_no_column and r.get(mapping.exam_no_column, "")
    ]
    subjects = _unique_join(
        r.get(mapping.subject_column, "")
        for r in records
        if mapping.subject_column
    )
    units = _unique_join(
        _join_parts(
            r.get(mapping.unit_column, ""),
            r.get(mapping.job_column, ""),
        )
        for r in records
    )
    start_no = min(exam_nos, key=_natural_sort_key) if exam_nos else ""
    end_no = max(exam_nos, key=_natural_sort_key) if exam_nos else ""
    return {
        "考场": _coerce(first.get(mapping.room_column, "")),
        "考点": _coerce(first.get(mapping.site_column, "")),
        "考试科目": subjects,
        "起始考号": start_no,
        "结束考号": end_no,
        "准考证号起止": f"{start_no} - {end_no}" if start_no else "",
        "单位岗位汇总": units,
        "考生人数": str(len(records)),
    }


def parse_custom_heights(text: str) -> list[int]:
    """解析自定义列高字符串，如 '8,7,7,8' → [8, 7, 7, 8]。"""
    text = text.strip()
    if not text:
        return []
    parts = re.split(r"[,\s，]+", text)
    heights: list[int] = []
    for part in parts:
        part = part.strip()
        if part.isdigit() and int(part) > 0:
            heights.append(int(part))
        else:
            return []
    return heights


def resolve_fields(
    record: dict[str, str],
    field_config: FieldConfig,
    context: dict[str, str] | None = None,
) -> list[str]:
    """根据字段配置生成文本行列表。"""
    lines: list[str] = []
    ctx = dict(record)
    if context:
        ctx.update(context)
    for field_name in field_config.fields:
        field_name = field_name.strip()
        if not field_name:
            continue
        # 支持 ${占位符} 语法
        if PLACEHOLDER_PATTERN.search(field_name):
            rendered = PLACEHOLDER_PATTERN.sub(
                lambda m: str(ctx.get(m.group(1).strip(), "")),
                field_name,
            )
            lines.append(rendered.strip())
        else:
            value = _coerce(ctx.get(field_name, ""))
            lines.append(f"{field_name}：{value}" if value else f"{field_name}：")
    return lines


# ── Grid Position Calculator ─────────────────────────────────────────────────


def calculate_grid_positions(
    count: int,
    grid: GridSettings,
) -> list[tuple[int, int]]:
    """计算网格中每个位置的 (row, col)。支持标准 M×N 和自定义列高。"""
    heights = parse_custom_heights(grid.custom_heights)
    if heights:
        return _custom_heights_positions(count, heights, grid)
    return _layout_positions(
        count,
        grid.columns,
        grid.rows,
        start_corner=grid.start_corner,
        fill_direction=grid.fill_direction,
        snake=grid.snake,
    )


def _custom_heights_positions(
    count: int,
    heights: list[int],
    grid: GridSettings,
) -> list[tuple[int, int]]:
    """自定义列高的位置计算。heights 是每列的座位数。"""
    top_first = not grid.start_corner.startswith("bottom")
    left_first = not grid.start_corner.endswith("right")

    col_indices = list(range(len(heights)))
    if not left_first:
        col_indices = list(reversed(col_indices))

    positions: list[tuple[int, int]] = []
    for ci, col in enumerate(col_indices):
        h = heights[col] if col < len(heights) else heights[-1]
        row_order = list(range(h)) if top_first else list(reversed(range(h)))
        if grid.snake and ci % 2 == 1:
            row_order = list(reversed(row_order))
        for row in row_order:
            positions.append((row, col))
            if len(positions) >= count:
                return positions
    return positions


def grid_dimensions(grid: GridSettings) -> tuple[int, int]:
    """返回网格的 (列数, 最大行数)。"""
    heights = parse_custom_heights(grid.custom_heights)
    if heights:
        return len(heights), max(heights)
    return grid.columns, grid.rows


# ── PDF Drawing Helpers ──────────────────────────────────────────────────────


def _draw_title(
    pdf,
    text: str,
    align: str,
    x: float,
    y: float,
    width: float,
    font_name: str,
    font_size: float,
) -> None:
    from reportlab.lib import colors
    if not text:
        return
    pdf.setFillColor(colors.black)
    pdf.setFont(font_name, font_size)
    if align == "left":
        pdf.drawString(x, y, text)
    elif align == "right":
        pdf.drawRightString(x + width, y, text)
    else:
        pdf.drawCentredString(x + width / 2, y, text)


def _draw_field_lines(
    pdf,
    lines: list[str],
    x: float,
    y: float,
    max_width: float,
    bottom_limit: float,
    font_name: str,
    font_size: float,
    leading: float,
) -> None:
    from reportlab.lib import colors
    pdf.setFillColor(colors.black)
    pdf.setFont(font_name, font_size)
    for line in lines:
        if y - leading < bottom_limit:
            break
        wrapped = _wrap_pdf_text(line, max_width, font_name, font_size)
        for wl in wrapped:
            if y - leading < bottom_limit:
                break
            if "签字" in wl or "签名" in wl:
                pdf.drawString(x, y, wl)
                from reportlab.pdfbase import pdfmetrics
                line_end = x + pdfmetrics.stringWidth(wl, font_name, font_size) + 4
                if line_end < x + max_width:
                    pdf.setStrokeColor(colors.HexColor("#333333"))
                    pdf.setLineWidth(0.45)
                    pdf.line(line_end, y - 1, x + max_width, y - 1)
            else:
                pdf.drawString(x, y, wl)
            y -= leading


# ── PDF Export Functions ─────────────────────────────────────────────────────


def export_signin_pdf(
    output_path: Path,
    records: list[dict[str, str]],
    mapping: ColumnMapping,
    settings: SigninSettings,
    photo_dir: Path | None = None,
) -> list[Path]:
    """导出面试签到表 PDF。"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas

    font_name = _register_font()
    sorted_records = sort_records(records, mapping.sort_column)
    if not sorted_records:
        return []
    photo_map = _photo_file_map(photo_dir)

    page_w, page_h = landscape(A4)
    margin = 14
    title_font_size = 24
    subtitle_font_size = 14
    content_width = page_w - margin * 2

    cols, max_rows = grid_dimensions(settings.grid)
    card_w = content_width / cols
    card_h_base = (page_h - margin * 2 - 80) / max_rows
    card_h = max(50, card_h_base)

    safe_photo_w = 0 if settings.photo_width <= 0 else min(settings.photo_width, card_w * 0.5)
    photo_h = max(1, card_h - 12)
    font_size = 10
    leading = font_size + 2

    pdf = canvas.Canvas(str(output_path), pagesize=landscape(A4))
    pdf.setTitle(settings.title or "面试签到表")

    per_page = cols * max_rows
    for page_start in range(0, len(sorted_records), per_page):
        page_records = sorted_records[page_start : page_start + per_page]
        positions = calculate_grid_positions(len(page_records), settings.grid)

        header_y = page_h - margin - 6
        _draw_title(pdf, settings.title, settings.title_align, margin, header_y, content_width, font_name, title_font_size)
        if settings.subtitle:
            _draw_title(pdf, settings.subtitle, settings.title_align, margin, header_y - title_font_size - 4, content_width, font_name, subtitle_font_size)

        grid_top = page_h - margin - 50
        for idx, record in enumerate(page_records):
            if idx >= len(positions):
                break
            row, col = positions[idx]
            x = margin + col * card_w
            y = grid_top - (row + 1) * card_h

            pdf.setStrokeColor(colors.HexColor("#333333"))
            pdf.setLineWidth(0.65)
            pdf.rect(x, y, card_w, card_h, stroke=1, fill=0)

            text_x = x + 6
            text_w = card_w - 12
            if safe_photo_w > 0:
                photo_path = find_photo(record, mapping.photo_match_column, photo_map)
                img_y = y + card_h - photo_h - 6
                if photo_path:
                    _draw_photo(pdf, photo_path, x + 6, img_y, safe_photo_w, photo_h)
                else:
                    pdf.setStrokeColor(colors.HexColor("#aaaaaa"))
                    pdf.rect(x + 6, img_y, safe_photo_w, photo_h, stroke=1, fill=0)
                text_x = x + 6 + safe_photo_w + 4
                text_w = card_w - safe_photo_w - 16

            lines = resolve_fields(record, settings.field_config)
            _draw_field_lines(pdf, lines, text_x, y + card_h - 8, text_w, y + 4, font_name, font_size, leading)

        pdf.showPage()
    pdf.save()
    return [output_path]


def export_seat_pdf(
    output_path: Path,
    records: list[dict[str, str]],
    mapping: ColumnMapping,
    settings: SeatSettings,
    photo_dir: Path | None = None,
) -> list[Path]:
    """导出笔试座次表 PDF。每个考场一页。"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas

    font_name = _register_font()
    photo_map = _photo_file_map(photo_dir)
    rooms = group_by_room(records, mapping.room_column)

    page_w, page_h = landscape(A4)
    margin = 14
    content_width = page_w - margin * 2
    title_font_size = 24
    subtitle_font_size = 14
    overview_font_size = 12

    cols, max_rows = grid_dimensions(settings.grid)
    card_w = content_width / cols
    card_h_base = (page_h - margin * 2 - 100) / max_rows
    card_h = max(50, card_h_base)

    safe_photo_w = 0 if settings.photo_width <= 0 else min(settings.photo_width, card_w * 0.5)
    photo_h = max(1, card_h - 12)
    font_size = 10
    leading = font_size + 2

    pdf = canvas.Canvas(str(output_path), pagesize=landscape(A4))
    pdf.setTitle(settings.title or "笔试座次表")

    for room_name, room_records in rooms:
        sorted_records = sort_records(room_records, mapping.sort_column)
        summary = compute_room_summary(sorted_records, mapping)
        positions = calculate_grid_positions(len(sorted_records), settings.grid)

        y_cursor = page_h - margin - 6
        _draw_title(pdf, settings.title, settings.title_align, margin, y_cursor, content_width, font_name, title_font_size)
        y_cursor -= title_font_size + 4

        if settings.subtitle:
            _draw_title(pdf, settings.subtitle, settings.title_align, margin, y_cursor, content_width, font_name, subtitle_font_size)
            y_cursor -= subtitle_font_size + 4

        if settings.show_room_overview and summary:
            overview_parts = []
            if summary.get("考场"):
                overview_parts.append(f"考场：{summary['考场']}")
            if summary.get("考点"):
                overview_parts.append(f"考点：{summary['考点']}")
            if summary.get("准考证号起止"):
                overview_parts.append(f"准考证号：{summary['准考证号起止']}")
            if summary.get("考生人数"):
                overview_parts.append(f"考生人数：{summary['考生人数']}")
            if summary.get("考试科目"):
                overview_parts.append(f"科目：{summary['考试科目']}")
            overview_text = "  |  ".join(overview_parts)
            pdf.setFont(font_name, overview_font_size)
            pdf.setFillColor(colors.HexColor("#555555"))
            pdf.drawCentredString(page_w / 2, y_cursor, overview_text)
            y_cursor -= overview_font_size + 6

        grid_top = y_cursor - 4
        for idx, record in enumerate(sorted_records):
            if idx >= len(positions):
                break
            row, col = positions[idx]
            x = margin + col * card_w
            y = grid_top - (row + 1) * card_h

            pdf.setStrokeColor(colors.HexColor("#333333"))
            pdf.setLineWidth(0.65)
            pdf.rect(x, y, card_w, card_h, stroke=1, fill=0)

            seat_no = f"{idx + 1:02d}"
            pdf.setFont(font_name, 6)
            pdf.setFillColor(colors.HexColor("#999999"))
            pdf.drawString(x + 2, y + card_h - 8, seat_no)

            text_x = x + 6
            text_w = card_w - 12
            if safe_photo_w > 0:
                photo_path = find_photo(record, mapping.photo_match_column, photo_map)
                img_y = y + 4
                if photo_path:
                    _draw_photo(pdf, photo_path, x + 6, img_y, safe_photo_w, photo_h)
                else:
                    pdf.setStrokeColor(colors.HexColor("#aaaaaa"))
                    pdf.rect(x + 6, img_y, safe_photo_w, photo_h, stroke=1, fill=0)
                text_x = x + 6 + safe_photo_w + 4
                text_w = card_w - safe_photo_w - 16

            lines = resolve_fields(record, settings.field_config)
            _draw_field_lines(pdf, lines, text_x, y + card_h - 12, text_w, y + 4, font_name, font_size, leading)

        if settings.show_proctor_signature:
            sig_y = grid_top - max_rows * card_h - 16
            pdf.setFont(font_name, 10)
            pdf.setFillColor(colors.black)
            pdf.drawString(margin, sig_y, "监考人员签字：________________    ________________")

        pdf.showPage()
    pdf.save()
    return [output_path]


def export_door_pdf(
    output_path: Path,
    records: list[dict[str, str]],
    mapping: ColumnMapping,
    settings: DoorSettings,
    photo_dir: Path | None = None,
) -> list[Path]:
    """导出门贴 PDF。每个考场一页，A4 竖版。"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    font_name = _register_font()
    rooms = group_by_room(records, mapping.room_column)

    page_w, page_h = A4
    margin = 36

    pdf = canvas.Canvas(str(output_path), pagesize=A4)
    pdf.setTitle("考场门贴")

    for room_name, room_records in rooms:
        summary = compute_room_summary(room_records, mapping)
        first = room_records[0] if room_records else {}
        all_values = dict(first)
        all_values.update(summary)

        border_x = margin
        border_y = margin
        border_w = page_w - margin * 2
        border_h = page_h - margin * 2
        pdf.setStrokeColor(colors.HexColor("#cbd5e1"))
        pdf.setLineWidth(2)
        pdf.rect(border_x, border_y, border_w, border_h, stroke=1, fill=0)

        pdf.setFont(font_name, 30)
        pdf.setFillColor(colors.HexColor("#1f2937"))
        pdf.drawCentredString(page_w / 2, page_h - margin - 50, "考场门贴")

        y = page_h - margin - 100
        line_height = 38
        content_width = page_w - margin * 2 - 40
        for field_name in settings.display_fields:
            value = all_values.get(field_name, "")
            if not value and field_name in first:
                value = _coerce(first.get(field_name, ""))
            text = f"{field_name}：{value}" if value else f"{field_name}："
            pdf.setFont(font_name, 18)
            pdf.setFillColor(colors.black)
            wrapped = _wrap_pdf_text(text, content_width, font_name, 18)
            for wl in wrapped:
                pdf.drawString(margin + 20, y, wl)
                y -= line_height

        pdf.showPage()
    pdf.save()
    return [output_path]


def export_desk_pdf(
    output_path: Path,
    records: list[dict[str, str]],
    mapping: ColumnMapping,
    settings: DeskSettings,
    photo_dir: Path | None = None,
) -> list[Path]:
    """导出桌贴 PDF。A4 竖版，虚线边框可裁剪。"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    font_name = _register_font()
    photo_map = _photo_file_map(photo_dir)
    sorted_records = sort_records(records, mapping.sort_column)
    if not sorted_records:
        return []

    page_w, page_h = A4
    margin = 14
    safe_cols = max(1, settings.columns)
    safe_items = max(1, settings.items_per_page)
    rows_per_page = math.ceil(safe_items / safe_cols)
    card_w = (page_w - margin * 2) / safe_cols
    card_h = (page_h - margin * 2) / rows_per_page

    safe_photo_w = 0 if settings.photo_width <= 0 else min(settings.photo_width, card_w * 0.4)
    photo_h = max(1, card_h - 20)
    font_size = 12
    name_font_size = 20
    leading = font_size + 2

    output_paths: list[Path] = []

    if settings.per_room and mapping.room_column:
        rooms = group_by_room(sorted_records, mapping.room_column)
        for room_name, room_records in rooms:
            safe_name = re.sub(r'[\\/:*?"<>|]+', "_", room_name or "未分组")
            room_path = output_path.with_name(f"{output_path.stem}_{safe_name}{output_path.suffix}")
            _write_desk_pdf(
                room_path, room_records, mapping, settings, photo_map,
                font_name, page_w, page_h, margin, safe_cols, safe_items,
                rows_per_page, card_w, card_h, safe_photo_w, photo_h,
                font_size, name_font_size, leading,
            )
            output_paths.append(room_path)
    else:
        _write_desk_pdf(
            output_path, sorted_records, mapping, settings, photo_map,
            font_name, page_w, page_h, margin, safe_cols, safe_items,
            rows_per_page, card_w, card_h, safe_photo_w, photo_h,
            font_size, name_font_size, leading,
        )
        output_paths.append(output_path)

    return output_paths


def _write_desk_pdf(
    output_path, records, mapping, settings, photo_map,
    font_name, page_w, page_h, margin, safe_cols, safe_items,
    rows_per_page, card_w, card_h, safe_photo_w, photo_h,
    font_size, name_font_size, leading,
) -> None:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    pdf = canvas.Canvas(str(output_path), pagesize=A4)
    pdf.setTitle("桌贴")

    for page_start in range(0, len(records), safe_items):
        page_records = records[page_start : page_start + safe_items]
        for idx, record in enumerate(page_records):
            row = idx // safe_cols
            col = idx % safe_cols
            x = margin + col * card_w
            y = page_h - margin - (row + 1) * card_h

            pdf.setStrokeColor(colors.HexColor("#94a3b8"))
            pdf.setLineWidth(0.5)
            pdf.setDash(3, 2)
            pdf.rect(x, y, card_w, card_h, stroke=1, fill=0)
            pdf.setDash()

            text_x = x + 10
            text_w = card_w - 20
            if safe_photo_w > 0:
                photo_path = find_photo(record, mapping.photo_match_column, photo_map)
                img_y = y + card_h - photo_h - 10
                if photo_path:
                    _draw_photo(pdf, photo_path, x + 10, img_y, safe_photo_w, photo_h)
                else:
                    pdf.setStrokeColor(colors.HexColor("#aaaaaa"))
                    pdf.rect(x + 10, img_y, safe_photo_w, photo_h, stroke=1, fill=0)
                text_x = x + 10 + safe_photo_w + 6
                text_w = card_w - safe_photo_w - 26

            name = _coerce(record.get("姓名", ""))
            if name:
                pdf.setFillColor(colors.black)
                pdf.setFont(font_name, name_font_size)
                pdf.drawString(text_x, y + card_h - 14, name)

            lines = resolve_fields(record, settings.field_config)
            lines = [ln for ln in lines if not ln.startswith("姓名：")]
            _draw_field_lines(
                pdf, lines, text_x, y + card_h - 14 - leading * 1.3,
                text_w, y + 10, font_name, font_size, leading,
            )
        pdf.showPage()
    pdf.save()
