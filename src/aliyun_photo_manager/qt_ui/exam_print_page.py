from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtPrintSupport import QPrintDialog, QPrinter
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFrame,
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
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..exam_printing import (
    DOC_TYPE_SEAT,
    PrintDataConfig,
    PrintTemplate,
    available_placeholders,
    export_rendered_html,
    load_excel_headers,
    load_excel_records,
    load_templates,
    render_document,
    room_values,
    save_template,
)
from .base_page import Card
from .common import AppComboBox
from .widgets import FormRow


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
            ("新建模板", self._new_template),
            ("打开模板", lambda: self._refresh_template_combo()),
            ("保存模板", self.save_current_template),
            ("打开Excel", self._choose_excel_and_load),
            ("选择照片文件夹", self._choose_photo_dir),
            ("预览", self.render_preview),
            ("导出HTML", self.export_html),
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

        left = Card()
        left.setObjectName("ExamPrintLeftRail")
        left.setMinimumWidth(210)
        left.setMaximumWidth(260)
        left_layout = left.body_layout
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(12)
        for index, title, subtitle in [
            ("1", "数据导入", "导入 Excel 并映射字段"),
            ("2", "照片匹配", "匹配考生照片"),
            ("3", "模板设计", "设计座次表"),
            ("4", "打印输出", "预览并打印/导出"),
        ]:
            step = self._step_item(index, title, subtitle, active=index == "1")
            left_layout.addWidget(step)
        left_layout.addStretch(1)
        info = Card()
        info_layout = info.body_layout
        info_layout.setContentsMargins(12, 12, 12, 12)
        info_layout.setSpacing(8)
        info_title = QLabel("模板信息")
        info_title.setProperty("sectionTitle", True)
        self.template_info_label = QLabel("模板名称：未选择\n模板类型：座次表\n页面大小：A4\n纸张方向：纵向")
        self.template_info_label.setWordWrap(True)
        info_layout.addWidget(info_title)
        info_layout.addWidget(self.template_info_label)
        left_layout.addWidget(info)
        workspace.addWidget(left)

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
        data_title = QLabel("数据导入与字段映射")
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

        data_split = QSplitter()
        mapping_card = QWidget()
        mapping_layout = QVBoxLayout(mapping_card)
        mapping_layout.setContentsMargins(0, 0, 0, 0)
        mapping_layout.setSpacing(8)
        mapping_title = QLabel("字段映射（将 Excel 列映射到系统字段）")
        mapping_title.setProperty("formLabel", True)
        mapping_layout.addWidget(mapping_title)

        self.doc_type_combo = AppComboBox()
        self.doc_type_combo.addItem("座次表", DOC_TYPE_SEAT)
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
        fields = [
            ("打印类型", self.doc_type_combo),
            ("模板", self.template_combo),
            ("模板名称", self.template_name_edit),
            ("考场/分组", self.room_column_combo),
            ("座号", self.seat_column_combo),
            ("准考证号", self.exam_no_column_combo),
            ("考点", self.site_column_combo),
            ("科目", self.subject_column_combo),
            ("单位", self.unit_column_combo),
            ("岗位", self.job_column_combo),
            ("照片匹配", self.photo_match_column_combo),
            ("预览范围", self.room_value_combo),
        ]
        for row, (label, field) in enumerate(fields):
            form_layout_row = FormRow(label, field)
            mapping_layout.addWidget(form_layout_row)
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
        self.columns_spin.setValue(4)
        self.items_per_page_spin = QSpinBox()
        self.items_per_page_spin.setRange(1, 60)
        self.items_per_page_spin.setValue(10)
        settings_row.addWidget(QLabel("每行列数"))
        settings_row.addWidget(self.columns_spin)
        settings_row.addStretch(1)
        mapping_layout.addLayout(settings_row)
        data_split.addWidget(mapping_card)

        preview_card = QWidget()
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(8)
        preview_title = QLabel("数据预览（前 10 条）")
        preview_title.setProperty("formLabel", True)
        preview_layout.addWidget(preview_title)
        self.data_preview_table = QTableWidget(0, 0)
        self.data_preview_table.setMinimumHeight(220)
        self.data_preview_table.setAlternatingRowColors(True)
        preview_layout.addWidget(self.data_preview_table)
        self.record_count_label = QLabel("未加载数据")
        self.record_count_label.setProperty("heroText", True)
        preview_layout.addWidget(self.record_count_label)
        data_split.addWidget(preview_card)
        data_split.setStretchFactor(0, 3)
        data_split.setStretchFactor(1, 5)
        data_layout.addWidget(data_split)
        center_layout.addWidget(data_card)

        designer_card = Card()
        designer_layout = designer_card.body_layout
        designer_layout.setContentsMargins(16, 14, 16, 14)
        designer_layout.setSpacing(10)
        designer_title = QLabel("模板设计（座次表模板）")
        designer_title.setProperty("sectionTitle", True)
        designer_layout.addWidget(designer_title)

        editor_toolbar = QHBoxLayout()
        for text, callback in [
            ("选择", lambda: None),
            ("文本", lambda: None),
            ("图片", lambda: None),
            ("表格", lambda: None),
            ("矩形", lambda: None),
            ("线条", lambda: None),
            ("删除", lambda: None),
            ("撤销", lambda: None),
            ("重做", lambda: None),
        ]:
            button = QPushButton(text)
            button.setProperty("toolbarButton", True)
            button.clicked.connect(callback)
            editor_toolbar.addWidget(button)
        bold_button = QPushButton("B")
        bold_button.clicked.connect(lambda: self._toggle_weight(True))
        italic_button = QPushButton("I")
        italic_button.clicked.connect(lambda: self._toggle_weight(False, italic=True))
        underline_button = QPushButton("U")
        underline_button.clicked.connect(self._toggle_underline)
        insert_button = QPushButton("插入占位符")
        insert_button.clicked.connect(self.insert_placeholder)
        for button in (bold_button, italic_button, underline_button, insert_button):
            button.setProperty("toolbarButton", True)
            editor_toolbar.addWidget(button)
        editor_toolbar.addStretch(1)
        editor_toolbar.addWidget(QLabel("100%"))
        designer_layout.addLayout(editor_toolbar)

        editor_splitter = QSplitter(Qt.Vertical)
        main_card = QWidget()
        main_layout = QVBoxLayout(main_card)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(QLabel("主模板画布"))
        self.main_editor = QTextEdit()
        self.main_editor.setAcceptRichText(True)
        self.main_editor.setMinimumHeight(220)
        main_layout.addWidget(self.main_editor)
        editor_splitter.addWidget(main_card)

        item_card = QWidget()
        item_layout = QVBoxLayout(item_card)
        item_layout.setContentsMargins(0, 0, 0, 0)
        self.item_title = QLabel("考生单元模板")
        item_layout.addWidget(self.item_title)
        self.item_editor = QTextEdit()
        self.item_editor.setAcceptRichText(True)
        self.item_editor.setMinimumHeight(140)
        item_layout.addWidget(self.item_editor)
        editor_splitter.addWidget(item_card)
        designer_layout.addWidget(editor_splitter, 1)
        center_layout.addWidget(designer_card, 1)

        result_card = Card()
        result_layout = result_card.body_layout
        result_layout.setContentsMargins(16, 14, 16, 14)
        result_layout.setSpacing(8)
        result_layout.addWidget(QLabel("预览与输出"))
        self.preview_edit = QTextEdit()
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
        layers_title = QLabel("控件图层")
        layers_title.setProperty("sectionTitle", True)
        right_layout.addWidget(layers_title)
        self.layer_text = QPlainTextEdit()
        self.layer_text.setReadOnly(True)
        self.layer_text.setPlainText("标题文本\n考场信息文本\n考生人数文本\n座次表格")
        right_layout.addWidget(self.layer_text, 1)
        workspace.addWidget(right)

        workspace.setStretchFactor(0, 1)
        workspace.setStretchFactor(1, 6)
        workspace.setStretchFactor(2, 1)

    def _step_item(self, index: str, title: str, subtitle: str, *, active: bool = False) -> QFrame:
        frame = QFrame()
        frame.setProperty("workflowStep", True)
        frame.setProperty("active", active)
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        badge = QLabel(index)
        badge.setFixedSize(28, 28)
        badge.setAlignment(Qt.AlignCenter)
        badge.setProperty("stepBadge", True)
        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(3)
        title_label = QLabel(title)
        title_label.setProperty("formLabel", True)
        subtitle_label = QLabel(subtitle)
        subtitle_label.setWordWrap(True)
        subtitle_label.setProperty("heroText", True)
        text_layout.addWidget(title_label)
        text_layout.addWidget(subtitle_label)
        layout.addWidget(badge)
        layout.addLayout(text_layout, 1)
        return frame

    def _choose_excel_and_load(self) -> None:
        self.choose_file(self.excel_edit)
        if self.excel_edit.text().strip():
            self.load_headers()

    def _choose_photo_dir(self) -> None:
        self.choose_folder(self.photo_dir_edit)

    def _new_template(self) -> None:
        self.template_name_edit.clear()
        self.main_editor.setHtml("<h1 style='text-align:center;'>考场座次表</h1><p>考点：${考点}　考场：${考场}</p><p>${座次表格}</p>")
        self.item_editor.setHtml("<p>${座号}　${姓名}　${准考证号}</p>")
        self.preview_edit.clear()

    def _refresh_data_preview(self) -> None:
        visible_headers = self.headers[:8]
        visible_records = self.records[:10]
        self.data_preview_table.clear()
        self.data_preview_table.setRowCount(len(visible_records))
        self.data_preview_table.setColumnCount(len(visible_headers))
        self.data_preview_table.setHorizontalHeaderLabels(visible_headers)
        for row_index, record in enumerate(visible_records):
            for column_index, header in enumerate(visible_headers):
                self.data_preview_table.setItem(
                    row_index,
                    column_index,
                    QTableWidgetItem(str(record.get(header, ""))),
                )
        self.data_preview_table.resizeColumnsToContents()
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

    def _apply_selected_template(self) -> None:
        template = self.current_template()
        if template is None:
            self.main_editor.clear()
            self.item_editor.clear()
            return
        self.template_name_edit.setText(template.name if not template.builtin else "")
        self.main_editor.setHtml(template.main_html)
        self.item_editor.setHtml(template.item_html)
        self.columns_spin.setValue(int(template.settings.get("columns", 4)))
        self.items_per_page_spin.setValue(int(template.settings.get("items_per_page", 10)))
        self.item_title.setText("考生单元模板")
        self.item_editor.setEnabled(True)
        self.columns_spin.setEnabled(template.doc_type == DOC_TYPE_SEAT)
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
        self._fill_header_combos()
        self.placeholder_text.setPlainText("\n".join(available_placeholders(self.headers)))
        self._refresh_data_preview()
        self.log_fn(f"已加载考场打印数据：{excel_path.name}，共 {len(self.records)} 条。")

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
        ]
        for combo in combos:
            combo.clear()
            combo.addItem("未设置", "")
            for header in self.headers:
                combo.addItem(header, header)

        self._select_guess(self.room_column_combo, ["考场", "考场号"])
        self._select_guess(self.seat_column_combo, ["座号"])
        self._select_guess(self.exam_no_column_combo, ["考号", "准考证号"])
        self._select_guess(self.site_column_combo, ["考点"])
        self._select_guess(self.subject_column_combo, ["考试科目", "科目"])
        self._select_guess(self.unit_column_combo, ["报考单位", "单位"])
        self._select_guess(self.job_column_combo, ["报考岗位", "岗位"])
        self._select_guess(self.photo_match_column_combo, ["身份证号", "证件号码", "身份证", "sfzh", "考号", "准考证号"])
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
        name = self.template_name_edit.text().strip() or (template.name if template and not template.builtin else "")
        if not name:
            QMessageBox.critical(self, "缺少名称", "请输入模板名称后再保存。")
            return
        saved = PrintTemplate(
            name=name,
            doc_type=doc_type,
            main_html=self.main_editor.toHtml(),
            item_html=self.item_editor.toHtml(),
            settings={
                "columns": self.columns_spin.value(),
                "items_per_page": self.items_per_page_spin.value(),
            },
            builtin=False,
        )
        save_template(saved)
        self.templates = load_templates()
        self._refresh_template_combo()
        index = self.template_combo.findText(name)
        if index >= 0:
            self.template_combo.setCurrentIndex(index)
        self.log_fn(f"已保存考场打印模板：{name}")

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
                main_html=self.main_editor.toHtml(),
                item_html=self.item_editor.toHtml(),
                settings={
                    "columns": self.columns_spin.value(),
                    "items_per_page": self.items_per_page_spin.value(),
                },
            )
            self.rendered_html = render_document(
                template,
                self.records,
                config,
                room_value,
                columns=self.columns_spin.value(),
                items_per_page=self.items_per_page_spin.value(),
            )
        except Exception as exc:
            QMessageBox.critical(self, "预览失败", str(exc))
            return
        self.preview_edit.setHtml(self.rendered_html)
        self.log_fn(f"已生成座次表预览：{room_value or '全部考生'}")

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

    def print_preview(self) -> None:
        if not self.rendered_html:
            self.render_preview()
            if not self.rendered_html:
                return
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() != QDialog.Accepted:
            return
        document = self.preview_edit.document().clone()
        document.print(printer)
        self.log_fn("已发送考场打印任务。")

    def insert_placeholder(self) -> None:
        text = self.placeholder_text.textCursor().selectedText().strip()
        if not text:
            cursor = self.placeholder_text.textCursor()
            cursor.select(QTextCursor.LineUnderCursor)
            text = cursor.selectedText().strip()
        if not text:
            QMessageBox.information(self, "插入占位符", "请先在左侧占位符列表里选中一项。")
            return
        target = self.main_editor if self.main_editor.hasFocus() else self.item_editor
        target.insertPlainText(text)

    def _toggle_weight(self, bold: bool, italic: bool = False) -> None:
        editor = self.main_editor if self.main_editor.hasFocus() else self.item_editor
        if bold:
            editor.setFontWeight(QFont.Normal if editor.fontWeight() > QFont.Normal else QFont.Bold)
        else:
            editor.setFontItalic(not editor.fontItalic() if italic else editor.fontItalic())

    def _toggle_underline(self) -> None:
        editor = self.main_editor if self.main_editor.hasFocus() else self.item_editor
        editor.setFontUnderline(not editor.fontUnderline())
