from __future__ import annotations

import base64
import json
import re
import uuid
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QByteArray, Qt, QUrl
from PySide6.QtGui import QFont, QImage, QTextCursor, QTextDocument
from PySide6.QtPrintSupport import QPrintDialog, QPrinter
from PySide6.QtWidgets import (
    QDialog,
    QColorDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..exam_printing import (
    DOC_TYPE_SEAT,
    PrintDataConfig,
    PrintTemplate,
    available_placeholders,
    export_interview_signin_pdf,
    export_rendered_html,
    load_excel_headers,
    load_excel_records,
    load_templates,
    render_document,
    room_values,
)
from .base_page import Card
from .common import AppComboBox
from .widgets import FormRow


DATA_IMAGE_PATTERN = re.compile(r'src="data:image/[^;]+;base64,([^"]+)"')
IMAGE_TAG_PATTERN = re.compile(r"<img\\b[^>]*>", re.IGNORECASE)
PHOTO_DATA_URI_PATTERN = re.compile(r"^data:image/[^;]+;base64,(.+)$")
SUPPORTED_PHOTO_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
TOP_SECTION_START = "<!-- TOP_SECTION_START -->"
TOP_SECTION_END = "<!-- TOP_SECTION_END -->"
BOTTOM_SECTION_START = "<!-- BOTTOM_SECTION_START -->"
BOTTOM_SECTION_END = "<!-- BOTTOM_SECTION_END -->"


class HtmlModeEditor(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._syncing = False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        self.tabs = QTabWidget()
        self.visual_edit = QTextEdit()
        self.visual_edit.setAcceptRichText(True)
        self.code_edit = QPlainTextEdit()
        self.code_edit.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.code_edit.setPlaceholderText("在这里编辑 HTML 代码。")

        self.tabs.addTab(self.visual_edit, "显示模式")
        self.tabs.addTab(self.code_edit, "代码模式")
        self.tabs.currentChanged.connect(self._sync_current_tab)
        layout.addWidget(self.tabs)

    def setHtml(self, html_text: str) -> None:
        self._syncing = True
        try:
            self.visual_edit.setHtml(html_text)
            self.code_edit.setPlainText(html_text)
        finally:
            self._syncing = False

    def toHtml(self) -> str:
        if self.tabs.currentWidget() is self.code_edit:
            return self.code_edit.toPlainText()
        return self.visual_edit.toHtml()

    def clear(self) -> None:
        self._syncing = True
        try:
            self.visual_edit.clear()
            self.code_edit.clear()
        finally:
            self._syncing = False

    def insertPlainText(self, text: str) -> None:
        if self.tabs.currentWidget() is self.code_edit:
            self.code_edit.insertPlainText(text)
        else:
            self.visual_edit.insertPlainText(text)

    def active_text_edit(self) -> QTextEdit | None:
        return self.visual_edit if self.tabs.currentWidget() is self.visual_edit else None

    def hasEditorFocus(self) -> bool:
        return self.visual_edit.hasFocus() or self.code_edit.hasFocus()

    def setEnabled(self, enabled: bool) -> None:
        super().setEnabled(enabled)
        self.tabs.setEnabled(enabled)

    def _sync_current_tab(self) -> None:
        if self._syncing:
            return
        self._syncing = True
        try:
            if self.tabs.currentWidget() is self.code_edit:
                self.code_edit.setPlainText(self.visual_edit.toHtml())
            else:
                self.visual_edit.setHtml(self.code_edit.toPlainText())
        finally:
            self._syncing = False


class PreviewDocument(QTextDocument):
    def __init__(self) -> None:
        super().__init__()
        self._images: dict[str, QImage] = {}

    def set_image_html(self, html_text: str) -> None:
        self._images = {}

        def replace_data_image(match: re.Match[str]) -> str:
            try:
                image_data = base64.b64decode(match.group(1))
            except Exception:
                return match.group(0)
            image = QImage.fromData(QByteArray(image_data))
            if image.isNull():
                return match.group(0)
            image_id = f"aliyun-photo-preview://image/{uuid.uuid4().hex}"
            self._images[image_id] = image
            return f'src="{image_id}"'

        self.setHtml(DATA_IMAGE_PATTERN.sub(replace_data_image, html_text))

    def loadResource(self, resource_type: int, name: QUrl):  # noqa: N802
        if resource_type == QTextDocument.ImageResource:
            key = name.toString()
            if key in self._images:
                return self._images[key]
        return super().loadResource(resource_type, name)


class PreviewTextEdit(QTextEdit):
    def __init__(self) -> None:
        super().__init__()
        self._preview_document = PreviewDocument()
        self.setDocument(self._preview_document)

    def set_image_html(self, html_text: str) -> None:
        self._preview_document.set_image_html(html_text)


class ExamPrintPage(QWidget):
    LABEL_WIDTH = 118
    ACTION_WIDTH = 110

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.templates = load_templates()
        self.records: list[dict[str, str]] = []
        self.headers: list[str] = []
        self.rendered_html = ""
        self._build_ui()
        self._refresh_template_combo()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        topbar = Card()
        topbar.setObjectName("ExamPrintTopbar")
        topbar_layout = QHBoxLayout()
        topbar_layout.setContentsMargins(12, 10, 12, 10)
        topbar_layout.setSpacing(8)
        for text, callback in [
            ("预览", self.render_preview),
            ("导出PDF", self.export_pdf),
            ("打印", self.print_preview),
        ]:
            button = QPushButton(text)
            button.setProperty("toolbarButton", True)
            button.clicked.connect(callback)
            topbar_layout.addWidget(button)
        topbar_layout.addStretch(1)
        topbar.body_layout.setContentsMargins(0, 0, 0, 0)
        topbar.body_layout.addLayout(topbar_layout)
        root.addWidget(topbar)

        workspace = QSplitter()
        root.addWidget(workspace, 1)

        center_scroll = QScrollArea()
        center_scroll.setWidgetResizable(True)
        center_scroll.setFrameShape(QFrame.NoFrame)
        center = QWidget()
        center_layout = QVBoxLayout(center)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(10)
        center_scroll.setWidget(center)

        data_card = Card()
        data_layout = data_card.body_layout
        data_layout.setContentsMargins(16, 14, 16, 14)
        data_layout.setSpacing(12)
        data_title = QLabel("面试签到数据与模板")
        data_title.setProperty("sectionTitle", True)
        data_layout.addWidget(data_title)

        file_row = QHBoxLayout()
        self.excel_edit = QLineEdit()
        self.excel_edit.setPlaceholderText("选择 Excel 数据文件")
        excel_button = QPushButton("重新选择")
        excel_button.clicked.connect(self._choose_excel_and_load)
        file_row.addWidget(QLabel("Excel文件："))
        file_row.addWidget(self.excel_edit, 1)
        file_row.addWidget(excel_button)
        data_layout.addLayout(file_row)
        self.record_count_label = QLabel("未加载数据")
        self.record_count_label.setProperty("heroText", True)
        data_layout.addWidget(self.record_count_label)

        mapping_card = QWidget()
        mapping_layout = QVBoxLayout(mapping_card)
        mapping_layout.setContentsMargins(0, 0, 0, 0)
        mapping_layout.setSpacing(8)
        mapping_title = QLabel("打印设置")
        mapping_title.setProperty("formLabel", True)
        mapping_layout.addWidget(mapping_title)

        self.doc_type_combo = AppComboBox()
        self.doc_type_combo.addItem("面试签到表", DOC_TYPE_SEAT)
        self.doc_type_combo.currentIndexChanged.connect(self._refresh_template_combo)
        self.template_combo = AppComboBox()
        self.template_combo.currentIndexChanged.connect(self._apply_selected_template)
        self.template_name_edit = QLineEdit()
        self.template_name_edit.setPlaceholderText("输入模板名称后保存")
        self.room_column_combo = AppComboBox()
        self.room_column_combo.currentIndexChanged.connect(self._refresh_room_values)
        self.seat_column_combo = AppComboBox()
        self.exam_no_column_combo = AppComboBox()
        self.site_column_combo = AppComboBox()
        self.subject_column_combo = AppComboBox()
        self.unit_column_combo = AppComboBox()
        self.job_column_combo = AppComboBox()
        self.photo_match_column_combo = AppComboBox()
        self.room_value_combo = AppComboBox()
        self.sort_column_combo = AppComboBox()
        self.file_group_column_combo = AppComboBox()
        self.page_group_column_combo = AppComboBox()
        self.start_corner_combo = AppComboBox()
        self.start_corner_combo.addItem("左上角", "top_left")
        self.start_corner_combo.addItem("右上角", "top_right")
        self.start_corner_combo.addItem("左下角", "bottom_left")
        self.start_corner_combo.addItem("右下角", "bottom_right")
        self.fill_direction_combo = AppComboBox()
        self.fill_direction_combo.addItem("从左到右、从上到下", "row")
        self.fill_direction_combo.addItem("从上到下、从左到右", "column")
        self.fill_direction_combo.currentIndexChanged.connect(self._update_people_count_label)
        self.snake_combo = AppComboBox()
        self.snake_combo.addItem("普通排序", "0")
        self.snake_combo.addItem("S 型排序", "1")
        fields = [
            ("模板", self.template_combo),
            ("排序列", self.sort_column_combo),
            ("分文件字段", self.file_group_column_combo),
            ("分页字段", self.page_group_column_combo),
            ("第一个考生", self.start_corner_combo),
            ("填充方向", self.fill_direction_combo),
            ("排序方式", self.snake_combo),
        ]
        fields_grid = QGridLayout()
        fields_grid.setContentsMargins(0, 0, 0, 0)
        fields_grid.setHorizontalSpacing(10)
        fields_grid.setVerticalSpacing(6)
        for index, (label, field) in enumerate(fields):
            row = index // 2
            column = (index % 2) * 2
            label_widget = QLabel(label)
            label_widget.setProperty("formLabel", True)
            fields_grid.addWidget(label_widget, row, column)
            fields_grid.addWidget(field, row, column + 1)
        fields_grid.setColumnStretch(1, 1)
        fields_grid.setColumnStretch(3, 1)
        mapping_layout.addLayout(fields_grid)
        photo_row = QHBoxLayout()
        self.photo_dir_edit = QLineEdit()
        self.photo_dir_edit.setPlaceholderText("可选，选择照片文件夹用于照片匹配")
        photo_button = QPushButton("选择")
        photo_button.clicked.connect(self._choose_photo_dir)
        photo_row.addWidget(QLabel("照片目录："))
        photo_row.addWidget(self.photo_dir_edit, 1)
        photo_row.addWidget(photo_button)
        mapping_layout.addLayout(photo_row)

        settings_row = QHBoxLayout()
        self.columns_spin = QSpinBox()
        self.columns_spin.setRange(1, 12)
        self.columns_spin.setValue(5)
        self.items_per_page_spin = QSpinBox()
        self.items_per_page_spin.setRange(1, 60)
        self.items_per_page_spin.setValue(10)
        self.columns_label = QLabel("每行人数")
        settings_row.addWidget(self.columns_label)
        settings_row.addWidget(self.columns_spin)
        self.photo_width_spin = QSpinBox()
        self.photo_width_spin.setRange(0, 90)
        self.photo_width_spin.setValue(52)
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(4, 16)
        self.font_size_spin.setValue(7)
        self.line_spacing_spin = QSpinBox()
        self.line_spacing_spin.setRange(5, 24)
        self.line_spacing_spin.setValue(8)
        settings_row.addWidget(QLabel("照片宽度"))
        settings_row.addWidget(self.photo_width_spin)
        settings_row.addWidget(QLabel("字体"))
        settings_row.addWidget(self.font_size_spin)
        settings_row.addWidget(QLabel("行距"))
        settings_row.addWidget(self.line_spacing_spin)
        settings_row.addStretch(1)
        mapping_layout.addLayout(settings_row)
        data_layout.addWidget(mapping_card)
        center_layout.addWidget(data_card)

        designer_card = Card()
        designer_layout = designer_card.body_layout
        designer_layout.setContentsMargins(16, 14, 16, 14)
        designer_layout.setSpacing(10)
        designer_title = QLabel("模板设计（上栏 + 下栏 + 单个考生）")
        designer_title.setProperty("sectionTitle", True)
        designer_layout.addWidget(designer_title)

        editor_splitter = QSplitter(Qt.Vertical)
        top_card = QWidget()
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.addWidget(QLabel("上栏"))
        self.top_editor = HtmlModeEditor()
        self.top_editor.setMinimumHeight(90)
        top_layout.addWidget(self.top_editor)
        editor_splitter.addWidget(top_card)

        item_card = QWidget()
        item_layout = QVBoxLayout(item_card)
        item_layout.setContentsMargins(0, 0, 0, 0)
        self.item_title = QLabel("单个考生表样")
        item_layout.addWidget(self.item_title)
        self.item_editor = HtmlModeEditor()
        self.item_editor.setMinimumHeight(140)
        item_layout.addWidget(self.item_editor)
        editor_splitter.addWidget(item_card)

        bottom_card = QWidget()
        bottom_layout = QVBoxLayout(bottom_card)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.addWidget(QLabel("下栏"))
        self.bottom_editor = HtmlModeEditor()
        self.bottom_editor.setMinimumHeight(90)
        bottom_layout.addWidget(self.bottom_editor)
        editor_splitter.addWidget(bottom_card)
        designer_layout.addWidget(editor_splitter, 1)
        center_layout.addWidget(designer_card, 1)

        result_card = Card()
        result_layout = result_card.body_layout
        result_layout.setContentsMargins(16, 14, 16, 14)
        result_layout.setSpacing(8)
        result_layout.addWidget(QLabel("预览与输出"))
        self.preview_edit = PreviewTextEdit()
        self.preview_edit.setReadOnly(True)
        self.preview_edit.setMinimumHeight(220)
        result_layout.addWidget(self.preview_edit)
        center_layout.addWidget(result_card)
        workspace.addWidget(center_scroll)

        right = Card()
        right.setObjectName("ExamPrintRightPanel")
        right.setMinimumWidth(230)
        right.setMaximumWidth(290)
        right_layout = right.body_layout
        right_layout.setContentsMargins(12, 12, 12, 12)
        right_layout.setSpacing(10)
        fields_title = QLabel("可用字段")
        fields_title.setProperty("sectionTitle", True)
        right_layout.addWidget(fields_title)
        self.placeholder_text = QPlainTextEdit()
        self.placeholder_text.setReadOnly(True)
        self.placeholder_text.setPlaceholderText("先选择 Excel 并加载列。")
        self.placeholder_text.setMinimumHeight(230)
        right_layout.addWidget(self.placeholder_text, 2)
        color_title = QLabel("颜色选择器")
        color_title.setProperty("sectionTitle", True)
        right_layout.addWidget(color_title)
        color_button = QPushButton("选择颜色")
        color_button.clicked.connect(self._choose_template_color)
        right_layout.addWidget(color_button)
        self.color_swatch = QLabel()
        self.color_swatch.setMinimumHeight(28)
        right_layout.addWidget(self.color_swatch)
        self.color_code_text = QPlainTextEdit()
        self.color_code_text.setReadOnly(True)
        self.color_code_text.setMinimumHeight(120)
        right_layout.addWidget(self.color_code_text, 1)
        self.selected_template_color = "#f5dada"
        self._set_template_color(self.selected_template_color)
        workspace.addWidget(right)

        workspace.setStretchFactor(0, 6)
        workspace.setStretchFactor(1, 1)

    def _choose_excel_and_load(self) -> None:
        self.choose_file(self.excel_edit)
        if self.excel_edit.text().strip():
            self.load_headers()

    def _choose_photo_dir(self) -> None:
        self.choose_folder(self.photo_dir_edit)

    def _new_template(self) -> None:
        template = self._default_local_template()
        self._set_current_template(template)
        self._apply_selected_template()
        template_path = self._local_template_path()
        if template_path is not None:
            try:
                self._write_local_template(template)
                self.log_fn(f"已生成默认模板：{template_path}")
            except Exception as exc:
                QMessageBox.critical(self, "生成模板失败", str(exc))
        self.preview_edit.clear()

    def _refresh_data_preview(self) -> None:
        self.record_count_label.setText(f"共 {len(self.records)} 人" if self.records else "未加载数据")

    def choose_file(self, line_edit: QLineEdit) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if selected:
            line_edit.setText(selected)

    def choose_folder(self, line_edit: QLineEdit) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择目录", str(Path.cwd()))
        if selected:
            line_edit.setText(selected)

    def _refresh_template_combo(self) -> None:
        doc_type = str(self.doc_type_combo.currentData())
        items = self.templates.get(doc_type, [])
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        for item in items:
            suffix = "（内置）" if item.builtin else ""
            self.template_combo.addItem(f"{item.name}{suffix}", item)
        self.template_combo.blockSignals(False)
        self._apply_selected_template()

    def _set_current_template(self, template: PrintTemplate) -> None:
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        self.template_combo.addItem(template.name, template)
        self.template_combo.blockSignals(False)
        self.template_combo.setCurrentIndex(0)

    def _default_local_template(self) -> PrintTemplate:
        default_template = self.templates.get(DOC_TYPE_SEAT, [])[1]
        top_html, bottom_html = self._template_sections(default_template)
        return PrintTemplate(
            name="面试签到表模板",
            doc_type=DOC_TYPE_SEAT,
            main_html=self._compose_main_html(top_html, bottom_html),
            item_html=default_template.item_html,
            settings={
                **default_template.settings,
                "top_html": top_html,
                "bottom_html": bottom_html,
                "columns": 5,
                "sort_column": "",
                "file_group_column": "",
                "page_group_column": "",
                "start_corner": "top_left",
                "fill_direction": "row",
                "snake": "0",
                "photo_width": 52,
                "font_size": 7,
                "line_spacing": 8,
                "section_background": str(default_template.settings.get("section_background", "#f5dada")),
            },
            builtin=False,
        )

    def _local_template_path(self) -> Path | None:
        excel_text = self.excel_edit.text().strip() if hasattr(self, "excel_edit") else ""
        if not excel_text:
            return None
        return Path(excel_text).parent / "面试签到表模板.json"

    def _apply_selected_template(self) -> None:
        template = self.current_template()
        if template is None:
            self.top_editor.clear()
            self.bottom_editor.clear()
            self.item_editor.clear()
            return
        self.template_name_edit.setText(template.name if not template.builtin else "")
        top_html, bottom_html = self._template_sections(template)
        self.top_editor.setHtml(top_html)
        self.bottom_editor.setHtml(bottom_html)
        self.item_editor.setHtml(template.item_html)
        self.columns_spin.setValue(int(template.settings.get("columns", 5)))
        self.items_per_page_spin.setValue(int(template.settings.get("items_per_page", 10)))
        self.photo_width_spin.setValue(int(template.settings.get("photo_width", 52)))
        self.font_size_spin.setValue(int(template.settings.get("font_size", 7)))
        self.line_spacing_spin.setValue(int(template.settings.get("line_spacing", 8)))
        self._set_template_color(str(template.settings.get("section_background", self.selected_template_color)))
        self._set_combo_data(self.start_corner_combo, str(template.settings.get("start_corner", "top_left")))
        self._set_combo_data(self.fill_direction_combo, str(template.settings.get("fill_direction", "row")))
        self._set_combo_data(self.snake_combo, str(template.settings.get("snake", "0")))
        self._set_combo_data(self.sort_column_combo, str(template.settings.get("sort_column", "")))
        self._set_combo_data(self.file_group_column_combo, str(template.settings.get("file_group_column", "")))
        self._set_combo_data(self.page_group_column_combo, str(template.settings.get("page_group_column", "")))
        self._update_people_count_label()
        is_signin = str(template.settings.get("layout", "")) == "signin"
        self.item_title.setText("单个考生表样")
        self.item_editor.setEnabled(not is_signin)
        self.columns_label.setEnabled(not is_signin)
        self.columns_spin.setEnabled(template.doc_type == DOC_TYPE_SEAT and not is_signin)
        self.items_per_page_spin.setEnabled(False)
        if hasattr(self, "template_info_label"):
            type_label = self.doc_type_combo.currentText()
            source = "内置模板" if template.builtin else "自定义模板"
            self.template_info_label.setText(
                f"模板名称：{template.name}\n模板类型：{type_label}\n模板来源：{source}\n页面大小：A4\n纸张方向：纵向"
            )

    def current_template(self) -> PrintTemplate | None:
        data = self.template_combo.currentData()
        return data if isinstance(data, PrintTemplate) else None

    def _compose_main_html(self, top_html: str, bottom_html: str) -> str:
        section_background = getattr(self, "selected_template_color", "#f5dada")
        return f"""
<html>
<head>
<style>
body {{ font-family: 'Microsoft YaHei UI'; margin: 10px; color: #111827; }}
.candidate-grid {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
.candidate-grid > tr > td, .candidate-grid > tbody > tr > td {{ border: 1px solid #444; padding: 0; height: 94px; vertical-align: top; }}
.candidate-card {{ width: 100%; height: 94px; border-collapse: collapse; table-layout: fixed; font-size: 12px; line-height: 1.35; }}
.candidate-card td {{ border: none; padding: 0; vertical-align: top; }}
.candidate-card .photo-cell {{ width: 68px; }}
.candidate-card .info-cell {{ padding: 3px 5px; }}
.photo {{ width: 68px; height: 94px; object-fit: cover; display: block; }}
.print-top {{ margin-bottom: 6px; background-color: {section_background}; }}
.print-bottom {{ margin-top: 8px; background-color: {section_background}; }}
</style>
</head>
<body>
  <div class="print-top">{TOP_SECTION_START}{top_html}{TOP_SECTION_END}</div>
  ${{考生列表}}
  <div class="print-bottom">{BOTTOM_SECTION_START}{bottom_html}{BOTTOM_SECTION_END}</div>
</body>
</html>
""".strip()

    def _current_main_html(self) -> str:
        return self._compose_main_html(
            self._editor_body_fragment(self.top_editor.toHtml()),
            self._editor_body_fragment(self.bottom_editor.toHtml()),
        )

    def _current_top_html(self) -> str:
        return self._editor_body_fragment(self.top_editor.toHtml())

    def _current_bottom_html(self) -> str:
        return self._editor_body_fragment(self.bottom_editor.toHtml())

    @staticmethod
    def _editor_body_fragment(html_text: str) -> str:
        html_text = re.sub(r"(?is)<style\b[^>]*>.*?</style>", "", html_text)
        html_text = re.sub(r"(?is)<script\b[^>]*>.*?</script>", "", html_text)
        body_match = re.search(r"(?is)<body\b[^>]*>(.*?)</body>", html_text)
        if body_match:
            html_text = body_match.group(1)
        html_text = re.sub(r"(?is)<head\b[^>]*>.*?</head>", "", html_text)
        return html_text.strip()

    def _template_sections(self, template: PrintTemplate) -> tuple[str, str]:
        top_html = str(template.settings.get("top_html", ""))
        bottom_html = str(template.settings.get("bottom_html", ""))
        if top_html or bottom_html:
            return top_html, bottom_html
        top_match = re.search(
            re.escape(TOP_SECTION_START) + r"(.*?)" + re.escape(TOP_SECTION_END),
            template.main_html,
            flags=re.DOTALL,
        )
        bottom_match = re.search(
            re.escape(BOTTOM_SECTION_START) + r"(.*?)" + re.escape(BOTTOM_SECTION_END),
            template.main_html,
            flags=re.DOTALL,
        )
        if top_match or bottom_match:
            return (
                top_match.group(1).strip() if top_match else "",
                bottom_match.group(1).strip() if bottom_match else "",
            )
        parts = template.main_html.split("${考生列表}", 1)
        if len(parts) == 2:
            return parts[0].strip(), parts[1].strip()
        return template.main_html, ""

    def _set_combo_data(self, combo: AppComboBox, value: str) -> None:
        index = combo.findData(value)
        if index >= 0:
            combo.setCurrentIndex(index)

    def _update_people_count_label(self) -> None:
        if str(self.fill_direction_combo.currentData() or "row") == "column":
            self.columns_label.setText("每列人数")
        else:
            self.columns_label.setText("每行人数")

    def load_headers(self) -> None:
        excel_path = Path(self.excel_edit.text().strip())
        if not excel_path.exists():
            QMessageBox.critical(self, "缺少 Excel", "请先选择有效的 Excel 数据文件。")
            return
        try:
            self.headers = load_excel_headers(excel_path)
            self.records = load_excel_records(excel_path)
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            return
        self._load_or_create_local_template()
        self._fill_header_combos()
        self.placeholder_text.setPlainText("\n".join(available_placeholders(self.headers)))
        self._refresh_data_preview()
        self.log_fn(f"已加载考场打印数据：{excel_path.name}，共 {len(self.records)} 条。")

    def _load_or_create_local_template(self) -> None:
        template_path = self._local_template_path()
        template = self._default_local_template()
        if template_path and template_path.exists():
            try:
                payload = json.loads(template_path.read_text(encoding="utf-8"))
                template = PrintTemplate(
                    name=str(payload.get("name") or "面试签到表模板"),
                    doc_type=str(payload.get("doc_type") or DOC_TYPE_SEAT),
                    main_html=str(payload.get("main_html") or template.main_html),
                    item_html=str(payload.get("item_html") or template.item_html),
                    settings=dict(payload.get("settings") or template.settings),
                    builtin=False,
                )
            except Exception as exc:
                self.log_fn(f"读取本地模板失败，已使用默认模板：{exc}")
        self._refresh_template_combo()
        if template_path and template_path.exists():
            self.template_combo.blockSignals(True)
            self.template_combo.addItem(template.name, template)
            self.template_combo.setCurrentIndex(self.template_combo.count() - 1)
            self.template_combo.blockSignals(False)
            self._apply_selected_template()
        if template_path and not template_path.exists():
            self._write_local_template(template)

    def _write_local_template(self, template: PrintTemplate) -> Path:
        template_path = self._local_template_path()
        if template_path is None:
            raise ValueError("请先选择 Excel 数据文件。")
        payload = {
            "name": template.name,
            "doc_type": template.doc_type,
            "main_html": template.main_html,
            "item_html": template.item_html,
            "settings": template.settings,
        }
        template_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return template_path

    def _fill_header_combos(self) -> None:
        combos = [
            self.room_column_combo,
            self.seat_column_combo,
            self.exam_no_column_combo,
            self.site_column_combo,
            self.subject_column_combo,
            self.unit_column_combo,
            self.job_column_combo,
            self.photo_match_column_combo,
            self.sort_column_combo,
            self.file_group_column_combo,
            self.page_group_column_combo,
        ]
        for combo in combos:
            combo.clear()
            combo.addItem("未设置", "")
            for header in self.headers:
                combo.addItem(header, header)

        self._select_guess(self.room_column_combo, ["考场", "考场号"])
        self._select_guess(self.seat_column_combo, ["序号", "座号"])
        self._select_guess(self.exam_no_column_combo, ["考号", "准考证号"])
        self._select_guess(self.site_column_combo, ["考点"])
        self._select_guess(self.subject_column_combo, ["考试科目", "科目"])
        self._select_guess(self.unit_column_combo, ["报考单位", "单位"])
        self._select_guess(self.job_column_combo, ["报考岗位", "岗位"])
        self._select_guess(self.photo_match_column_combo, ["身份证号", "证件号码", "身份证", "sfzh", "考号", "准考证号"])
        template = self.current_template()
        saved_sort_column = str((template.settings if template else {}).get("sort_column", ""))
        if saved_sort_column:
            self._set_combo_data(self.sort_column_combo, saved_sort_column)
        else:
            self._select_guess(self.sort_column_combo, ["考号", "序号", "座号", "身份证号"])
        template = self.current_template()
        saved_file_group_column = str((template.settings if template else {}).get("file_group_column", ""))
        if saved_file_group_column:
            self._set_combo_data(self.file_group_column_combo, saved_file_group_column)
        saved_page_group_column = str((template.settings if template else {}).get("page_group_column", ""))
        if saved_page_group_column:
            self._set_combo_data(self.page_group_column_combo, saved_page_group_column)
        self._refresh_room_values()

    def _select_guess(self, combo: AppComboBox, candidates: list[str]) -> None:
        for candidate in candidates:
            index = combo.findData(candidate)
            if index >= 0:
                combo.setCurrentIndex(index)
                return

    def _refresh_room_values(self) -> None:
        room_column = str(self.room_column_combo.currentData() or "")
        self.room_value_combo.clear()
        if not self.records:
            return
        if not room_column:
            self.room_value_combo.addItem("全部考生", "")
            return
        for value in room_values(self.records, room_column):
            self.room_value_combo.addItem(value, value)

    def _build_config(self) -> PrintDataConfig:
        excel_path = Path(self.excel_edit.text().strip())
        if not excel_path.exists():
            raise ValueError("请先选择有效的 Excel 数据文件。")
        room_column = str(self.room_column_combo.currentData() or "")
        photo_dir_text = self.photo_dir_edit.text().strip()
        photo_dir = Path(photo_dir_text) if photo_dir_text else None
        return PrintDataConfig(
            excel_path=excel_path,
            room_column=room_column,
            seat_column=str(self.seat_column_combo.currentData() or ""),
            exam_no_column=str(self.exam_no_column_combo.currentData() or ""),
            site_column=str(self.site_column_combo.currentData() or ""),
            subject_column=str(self.subject_column_combo.currentData() or ""),
            unit_column=str(self.unit_column_combo.currentData() or ""),
            job_column=str(self.job_column_combo.currentData() or ""),
            photo_dir=photo_dir,
            photo_match_column=str(self.photo_match_column_combo.currentData() or ""),
        )

    def save_current_template(self) -> None:
        template = self.current_template()
        doc_type = str(self.doc_type_combo.currentData())
        name = "面试签到表模板"
        saved = PrintTemplate(
            name=name,
            doc_type=doc_type,
            main_html=self._current_main_html(),
            item_html=self.item_editor.toHtml(),
            settings={
                "top_html": self._current_top_html(),
                "bottom_html": self._current_bottom_html(),
                "columns": self.columns_spin.value(),
                "items_per_page": self.items_per_page_spin.value(),
                "layout": str(template.settings.get("layout", "")) if template else "",
                "sort_column": str(self.sort_column_combo.currentData() or ""),
                "file_group_column": str(self.file_group_column_combo.currentData() or ""),
                "page_group_column": str(self.page_group_column_combo.currentData() or ""),
                "start_corner": str(self.start_corner_combo.currentData() or "top_left"),
                "fill_direction": str(self.fill_direction_combo.currentData() or "row"),
                "snake": str(self.snake_combo.currentData() or "0"),
                "photo_width": self.photo_width_spin.value(),
                "font_size": self.font_size_spin.value(),
                "line_spacing": self.line_spacing_spin.value(),
                "section_background": self.selected_template_color,
            },
            builtin=False,
        )
        try:
            template_path = self._write_local_template(saved)
        except Exception as exc:
            QMessageBox.critical(self, "保存模板失败", str(exc))
            return
        self._set_current_template(saved)
        self.log_fn(f"已保存面试签到表模板：{template_path}")
        QMessageBox.information(self, "保存模板", f"已保存到：{template_path}")

    def render_preview(self) -> None:
        if not self.records:
            self.load_headers()
            if not self.records:
                return
        try:
            config = self._build_config()
            room_value = str(self.room_value_combo.currentData() or "")
            if not room_value and str(self.room_column_combo.currentData() or ""):
                self._refresh_room_values()
                room_value = str(self.room_value_combo.currentData() or "")
            if not room_value and str(self.room_column_combo.currentData() or ""):
                raise ValueError("当前没有可预览的考场。")
            template = PrintTemplate(
                name=self.template_name_edit.text().strip() or "当前模板",
                doc_type=str(self.doc_type_combo.currentData()),
                main_html=self._current_main_html(),
                item_html=self.item_editor.toHtml(),
                settings={
                    "columns": self.columns_spin.value(),
                    "items_per_page": self.items_per_page_spin.value(),
                    "layout": str((self.current_template().settings if self.current_template() else {}).get("layout", "")),
                },
            )
            self.rendered_html = render_document(
                template,
                self.records,
                config,
                room_value,
                columns=self.columns_spin.value(),
                items_per_page=self.items_per_page_spin.value(),
                sort_column=str(self.sort_column_combo.currentData() or ""),
            )
        except Exception as exc:
            QMessageBox.critical(self, "预览失败", str(exc))
            return
        self._set_preview_html(self.rendered_html)
        self.log_fn(f"已生成面试签到表预览：{room_value or '全部考生'}")

    def export_html(self) -> None:
        if not self.rendered_html:
            self.render_preview()
            if not self.rendered_html:
                return
        selected, _ = QFileDialog.getSaveFileName(self, "导出 HTML", "考场打印预览.html", "HTML 文件 (*.html)")
        if not selected:
            return
        export_rendered_html(Path(selected), self.rendered_html)
        self.log_fn(f"已导出考场打印 HTML：{selected}")

    def export_pdf(self) -> None:
        if not self.rendered_html:
            self.render_preview()
            if not self.rendered_html:
                return
        selected, _ = QFileDialog.getSaveFileName(self, "导出 PDF", "面试表样打印.pdf", "PDF 文件 (*.pdf)")
        if not selected:
            return
        output_path = Path(selected)
        if output_path.suffix.lower() != ".pdf":
            output_path = output_path.with_suffix(".pdf")
        try:
            config = self._build_config()
            room_value = str(self.room_value_combo.currentData() or "")
            template = self.current_template()
            if (
                str(self.doc_type_combo.currentData()) == DOC_TYPE_SEAT
                and str((template.settings if template else {}).get("layout", "")) == "candidate_grid"
            ):
                top_html = self._current_top_html()
                output_paths = export_interview_signin_pdf(
                    output_path,
                    self.records,
                    config,
                    room_value,
                    title=self._plain_text_from_html(top_html) or "面试表样打印",
                    top_html=top_html,
                    bottom_html=self._current_bottom_html(),
                    item_html=self.item_editor.toHtml(),
                    people_per_line=self.columns_spin.value(),
                    sort_column=str(self.sort_column_combo.currentData() or ""),
                    file_group_column=str(self.file_group_column_combo.currentData() or ""),
                    page_group_column=str(self.page_group_column_combo.currentData() or ""),
                    start_corner=str(self.start_corner_combo.currentData() or "top_left"),
                    fill_direction=str(self.fill_direction_combo.currentData() or "row"),
                    snake=str(self.snake_combo.currentData() or "0") == "1",
                    photo_width=self.photo_width_spin.value(),
                    font_size=self.font_size_spin.value(),
                    line_spacing=self.line_spacing_spin.value(),
                    section_background=self.selected_template_color,
                )
                if not output_paths:
                    raise OSError("PDF 文件没有成功生成。")
            else:
                output_paths = [output_path]
                printer = QPrinter(QPrinter.HighResolution)
                printer.setOutputFormat(QPrinter.PdfFormat)
                printer.setOutputFileName(str(output_path))
                self._print_preview_document(printer)
            if any(not path.exists() or path.stat().st_size == 0 for path in output_paths):
                raise OSError("PDF 文件没有成功生成。")
        except Exception as exc:
            QMessageBox.critical(self, "导出 PDF 失败", str(exc))
            return
        if len(output_paths) == 1:
            message = f"已导出：{output_paths[0]}"
        else:
            message = f"已导出 {len(output_paths)} 个 PDF 到：{output_paths[0].parent}"
        self.log_fn(f"已导出面试表样 PDF：{message}")
        QMessageBox.information(self, "导出 PDF", message)

    def print_preview(self) -> None:
        if not self.rendered_html:
            self.render_preview()
            if not self.rendered_html:
                return
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self._print_preview_document(printer)
        self.log_fn("已发送考场打印任务。")

    def _set_preview_html(self, html_text: str) -> None:
        self.preview_edit.clear()
        preview_html = IMAGE_TAG_PATTERN.sub(
            '<span style="display:inline-block;width:40px;height:58px;border:1px solid #999;text-align:center;color:#777;">照片</span>',
            html_text,
        )
        self.preview_edit.setHtml(preview_html)

    @staticmethod
    def _plain_text_from_html(html_text: str) -> str:
        text = re.sub(r"(?is)<style\b[^>]*>.*?</style>", "", html_text)
        text = re.sub(r"(?is)<script\b[^>]*>.*?</script>", "", text)
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"(?i)</(div|p|tr|li|td|h[1-6])>", "\n", text)
        text = re.sub(r"<[^>]+>", "", text)
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        return lines[0] if lines else ""

    def _print_preview_document(self, printer: QPrinter) -> None:
        print_method = getattr(self.preview_edit, "print_", None) or getattr(self.preview_edit, "print", None)
        if print_method is not None:
            print_method(printer)
            return
        document = self.preview_edit.document()
        document_print = getattr(document, "print_", None) or getattr(document, "print", None)
        if document_print is None:
            raise AttributeError("当前 Qt 版本不支持 PDF 打印方法。")
        document_print(printer)

    def _choose_template_color(self) -> None:
        color = QColorDialog.getColor(parent=self, title="选择模板颜色")
        if color.isValid():
            self._set_template_color(color.name())

    def _set_template_color(self, hex_color: str) -> None:
        self.selected_template_color = hex_color
        self.color_swatch.setStyleSheet(
            f"background-color: {hex_color}; border: 1px solid #94a3b8; border-radius: 4px;"
        )
        self.color_code_text.setPlainText(
            "\n".join(
                [
                    f"颜色值：{hex_color}",
                    "",
                    f"文字颜色：color:{hex_color};",
                    f"背景颜色：background-color:{hex_color};",
                    "",
                    f'<div style="color:{hex_color};">文字</div>',
                    f'<div style="background-color:{hex_color};">文字</div>',
                ]
            )
        )

    def insert_placeholder(self) -> None:
        text = self.placeholder_text.textCursor().selectedText().strip()
        if not text:
            cursor = self.placeholder_text.textCursor()
            cursor.select(QTextCursor.LineUnderCursor)
            text = cursor.selectedText().strip()
        if not text:
            QMessageBox.information(self, "插入占位符", "请先在左侧占位符列表里选中一项。")
            return
        target = self._active_template_editor()
        target.insertPlainText(text)

    def _toggle_weight(self, bold: bool, italic: bool = False) -> None:
        editor = self._active_visual_editor()
        if editor is None:
            return
        if bold:
            editor.setFontWeight(QFont.Normal if editor.fontWeight() > QFont.Normal else QFont.Bold)
        else:
            editor.setFontItalic(not editor.fontItalic() if italic else editor.fontItalic())

    def _toggle_underline(self) -> None:
        editor = self._active_visual_editor()
        if editor is None:
            return
        editor.setFontUnderline(not editor.fontUnderline())

    def _active_template_editor(self) -> HtmlModeEditor:
        if self.item_editor.hasEditorFocus():
            return self.item_editor
        if self.bottom_editor.hasEditorFocus():
            return self.bottom_editor
        return self.top_editor

    def _active_visual_editor(self) -> QTextEdit | None:
        return self._active_template_editor().active_text_edit()
