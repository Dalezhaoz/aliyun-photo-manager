import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional


LogFn = Optional[Callable[[str], None]]


@dataclass
class WordExportResult:
    source_path: Path
    variant: str
    html_content: str
    preview_html: str


def _log(logger: LogFn, message: str) -> None:
    if logger is not None:
        logger(message)


def _load_mammoth():
    try:
        import mammoth
    except ImportError as exc:
        raise ImportError(
            "缺少依赖 mammoth，请先执行 `pip install -r requirements.txt`。"
        ) from exc
    return mammoth


def _load_bs4():
    try:
        from bs4 import BeautifulSoup
    except ImportError as exc:
        raise ImportError(
            "缺少依赖 beautifulsoup4，请先执行 `pip install -r requirements.txt`。"
        ) from exc
    return BeautifulSoup


def _normalize_label(text: str) -> str:
    cleaned = (
        text.replace("：", "")
        .replace(":", "")
        .replace("\n", "")
        .replace("\r", "")
        .replace("\t", "")
        .strip()
    )
    return "".join(cleaned.split())


def _is_blank_cell(text: str) -> bool:
    return _normalize_label(text) == ""


def _build_placeholder(label: str, variant: str) -> str:
    normalized = _normalize_label(label)
    if not normalized:
        return ""
    # “照片”是固定图片占位符，不走“左标题右空白”那套推断。
    if "照片" in normalized:
        return "{[#考生表视图.照片#]}" if variant == "net" else "${考生.照片}"
    if variant == "net":
        return f"{{[#考生表视图.{normalized}#]}}"
    return f"${{考生.{normalized}}}"


def _convert_doc_to_docx(source_path: Path) -> Path:
    office_binary = shutil.which("soffice") or shutil.which("libreoffice")
    if office_binary is None:
        raise RuntimeError("`.doc` 需要 LibreOffice 支持转换。请先安装 LibreOffice，或先手动另存为 `.docx`。")

    temp_dir = Path(tempfile.mkdtemp(prefix="word_to_html_"))
    command = [
        office_binary,
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        str(temp_dir),
        str(source_path),
    ]
    completed = subprocess.run(command, capture_output=True, text=True)
    if completed.returncode != 0:
        raise RuntimeError(
            "`.doc` 转 `.docx` 失败："
            + (completed.stderr.strip() or completed.stdout.strip() or "未知错误")
        )
    converted_path = temp_dir / f"{source_path.stem}.docx"
    if not converted_path.exists():
        raise RuntimeError("`.doc` 已尝试转换，但未找到生成的 `.docx` 文件。")
    return converted_path


def _prepare_source(source_path: Path) -> Path:
    suffix = source_path.suffix.lower()
    if suffix == ".docx":
        return source_path
    if suffix == ".doc":
        return _convert_doc_to_docx(source_path)
    raise ValueError("仅支持 `.docx` 或 `.doc` 文件。")


def _apply_simple_styles(soup) -> None:
    for table in soup.find_all("table"):
        # 富文本编辑器里更依赖内联样式，所以表格宽度和边框直接写死在标签上。
        table["style"] = "border-collapse:collapse;width:680px;border:1px solid #000;"
        table["border"] = "1"
        table["cellspacing"] = "0"
        table["cellpadding"] = "0"
    for row in soup.find_all("tr"):
        row.attrs.pop("style", None)
    for cell in soup.find_all(["td", "th"]):
        existing = cell.get("style", "")
        parts = [part.strip() for part in existing.split(";") if part.strip()]
        parts.extend(
            [
                "border:1px solid #000",
                "padding:8px",
                "vertical-align:middle",
                "word-break:break-word",
                "text-align:center",
            ]
        )
        cell["style"] = ";".join(dict.fromkeys(parts))
    for paragraph in soup.find_all("p"):
        if not paragraph.get("style"):
            paragraph["style"] = "margin:6px 0;"


def _fill_table_placeholders(soup, variant: str) -> None:
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        previous_row_labels = []
        for row in rows:
            cells = row.find_all(["td", "th"])
            current_row_labels = []
            for index, cell in enumerate(cells):
                text = cell.get_text(" ", strip=True)
                normalized = _normalize_label(text)
                left_label = current_row_labels[index - 1] if index > 0 else ""
                upper_label = previous_row_labels[index] if index < len(previous_row_labels) else ""

                if normalized and "照片" in normalized:
                    cell.clear()
                    cell.append(_build_placeholder("照片", variant))
                    current_row_labels.append("照片")
                    continue

                if _is_blank_cell(text):
                    # 这类报名表通常是“左侧标题、右侧空白值”，优先取左边，再回退到上方标题。
                    label = left_label or upper_label
                    if label:
                        cell.clear()
                        cell.append(_build_placeholder(label, variant))
                        current_row_labels.append(label)
                    else:
                        current_row_labels.append("")
                    continue

                current_row_labels.append(normalized)
            previous_row_labels = current_row_labels


def _wrap_html(body_html: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>Word 转 HTML</title>
  <style>
    html, body {{
      margin: 0;
      padding: 0;
    }}
    body {{
      font-family: SimSun, "Microsoft YaHei", sans-serif;
      font-size: 16px;
      color: #000;
    }}
    .page {{
      width: 680px;
      margin: 16px auto;
      padding: 0;
      box-sizing: border-box;
    }}
    table {{
      border-collapse: collapse;
      width: 100%;
      table-layout: fixed;
    }}
    td, th {{
      border: 1px solid #000;
      padding: 8px;
      vertical-align: middle;
      word-break: break-word;
      text-align: center;
    }}
    p {{
      margin: 6px 0;
    }}
    @media print {{
      .page {{
        width: auto;
        margin: 0;
        padding: 0;
      }}
    }}
  </style>
</head>
<body>
  <div class="page">
{body_html}
  </div>
</body>
</html>
"""


def export_word_to_html(
    source_path: Path,
    variant: str,
    logger: LogFn = None,
) -> WordExportResult:
    if variant not in {"net", "java"}:
        raise ValueError("variant 只支持 `net` 或 `java`。")
    if not source_path.exists():
        raise FileNotFoundError(f"未找到 Word 文件：{source_path}")

    prepared_source = _prepare_source(source_path)
    mammoth = _load_mammoth()
    BeautifulSoup = _load_bs4()

    _log(logger, f"开始转换 Word：{source_path.name}")
    with prepared_source.open("rb") as file_obj:
        result = mammoth.convert_to_html(file_obj)
    soup = BeautifulSoup(result.value, "html.parser")
    _apply_simple_styles(soup)
    _fill_table_placeholders(soup, variant)

    body_html = soup.prettify()
    output_html = _wrap_html(body_html)
    _log(logger, f"已生成 HTML 代码：{source_path.stem}_{variant}.html")
    return WordExportResult(
        source_path=source_path,
        variant=variant,
        html_content=output_html,
        preview_html=body_html,
    )
