from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtPrintSupport import QPrintDialog, QPrinter
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSpinBox,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..exam_printing import (
    DOC_TYPE_DESK,
    DOC_TYPE_DOOR,
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
from .common import AppComboBox


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
        root.setSpacing(16)

        hero = self._card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        title = QLabel("考场文件打印")
        title.setProperty("heroTitle", True)
        intro = QLabel("从本地 Excel 和照片目录生成座次表、门贴、桌贴。模板使用占位符，可编辑、可保存、可预览。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        body = QSplitter()
        root.addWidget(body, 1)

        left = self._card()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(24, 22, 24, 24)
        left_layout.setSpacing(18)

        data_title = QLabel("数据与模板")
        data_title.setProperty("sectionTitle", True)
        left_layout.addWidget(data_title)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.excel_edit = QLineEdit()
        self._add_row(form, 0, "Excel 数据", self._with_file_button(self.excel_edit))
        self.photo_dir_edit = QLineEdit()
        self._add_row(form, 1, "照片目录", self._with_folder_button(self.photo_dir_edit))

        self.doc_type_combo = AppComboBox()
        self.doc_type_combo.addItem("座次表", DOC_TYPE_SEAT)
        self.doc_type_combo.addItem("门贴", DOC_TYPE_DOOR)
        self.doc_type_combo.addItem("桌贴", DOC_TYPE_DESK)
        self.doc_type_combo.currentIndexChanged.connect(self._refresh_template_combo)
        self._add_row(form, 2, "打印类型", self.doc_type_combo)

        self.template_combo = AppComboBox()
        self.template_combo.currentIndexChanged.connect(self._apply_selected_template)
        self._add_row(form, 3, "模板", self.template_combo)

        self.template_name_edit = QLineEdit()
        self.template_name_edit.setPlaceholderText("输入模板名称后保存")
        self._add_row(form, 4, "模板名称", self.template_name_edit)

        self.room_column_combo = AppComboBox()
        self.room_column_combo.currentIndexChanged.connect(self._refresh_room_values)
        self._add_row(form, 5, "考场列", self.room_column_combo)
        self.seat_column_combo = AppComboBox()
        self._add_row(form, 6, "座号列", self.seat_column_combo)
        self.exam_no_column_combo = AppComboBox()
        self._add_row(form, 7, "考号列", self.exam_no_column_combo)
        self.site_column_combo = AppComboBox()
        self._add_row(form, 8, "考点列", self.site_column_combo)
        self.subject_column_combo = AppComboBox()
        self._add_row(form, 9, "科目列", self.subject_column_combo)
        self.unit_column_combo = AppComboBox()
        self._add_row(form, 10, "单位列", self.unit_column_combo)
        self.job_column_combo = AppComboBox()
        self._add_row(form, 11, "岗位列", self.job_column_combo)
        self.photo_match_column_combo = AppComboBox()
        self._add_row(form, 12, "照片匹配列", self.photo_match_column_combo)
        self.room_value_combo = AppComboBox()
        self._add_row(form, 13, "预览考场", self.room_value_combo)

        self.columns_spin = QSpinBox()
        self.columns_spin.setRange(1, 12)
        self.columns_spin.setValue(4)
        self._add_row(form, 14, "每行列数", self.columns_spin)

        self.items_per_page_spin = QSpinBox()
        self.items_per_page_spin.setRange(1, 60)
        self.items_per_page_spin.setValue(10)
        self._add_row(form, 15, "桌贴每页人数", self.items_per_page_spin)

        left_layout.addLayout(form)

        action_row = QHBoxLayout()
        load_button = QPushButton("加载列")
        load_button.clicked.connect(self.load_headers)
        save_button = QPushButton("保存模板")
        save_button.clicked.connect(self.save_current_template)
        preview_button = QPushButton("生成预览")
        preview_button.setProperty("accent", True)
        preview_button.clicked.connect(self.render_preview)
        action_row.addWidget(load_button)
        action_row.addWidget(save_button)
        action_row.addWidget(preview_button)
        left_layout.addLayout(action_row)

        placeholder_title = QLabel("可用占位符")
        placeholder_title.setProperty("sectionTitle", True)
        left_layout.addWidget(placeholder_title)
        self.placeholder_text = QPlainTextEdit()
        self.placeholder_text.setReadOnly(True)
        self.placeholder_text.setPlaceholderText("先选择 Excel 并加载列。")
        self.placeholder_text.setMinimumHeight(160)
        left_layout.addWidget(self.placeholder_text, 1)

        body.addWidget(left)

        right = self._card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(24, 22, 24, 24)
        right_layout.setSpacing(16)

        editor_title = QLabel("模板编辑与预览")
        editor_title.setProperty("sectionTitle", True)
        right_layout.addWidget(editor_title)

        toolbar = QHBoxLayout()
        bold_button = QPushButton("加粗")
        bold_button.clicked.connect(lambda: self._toggle_weight(True))
        italic_button = QPushButton("斜体")
        italic_button.clicked.connect(lambda: self._toggle_weight(False, italic=True))
        underline_button = QPushButton("下划线")
        underline_button.clicked.connect(self._toggle_underline)
        insert_button = QPushButton("插入占位符")
        insert_button.clicked.connect(self.insert_placeholder)
        toolbar.addWidget(bold_button)
        toolbar.addWidget(italic_button)
        toolbar.addWidget(underline_button)
        toolbar.addWidget(insert_button)
        toolbar.addStretch(1)
        right_layout.addLayout(toolbar)

        editor_splitter = QSplitter()
        right_layout.addWidget(editor_splitter, 1)

        main_card = self._card()
        main_layout = QVBoxLayout(main_card)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.addWidget(QLabel("主模板"))
        self.main_editor = QTextEdit()
        self.main_editor.setAcceptRichText(True)
        main_layout.addWidget(self.main_editor)
        editor_splitter.addWidget(main_card)

        item_card = self._card()
        item_layout = QVBoxLayout(item_card)
        item_layout.setContentsMargins(16, 16, 16, 16)
        self.item_title = QLabel("考生单元模板")
        item_layout.addWidget(self.item_title)
        self.item_editor = QTextEdit()
        self.item_editor.setAcceptRichText(True)
        item_layout.addWidget(self.item_editor)
        editor_splitter.addWidget(item_card)

        preview_card = self._card()
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(16, 16, 16, 16)
        preview_layout.addWidget(QLabel("预览"))
        self.preview_edit = QTextEdit()
        self.preview_edit.setReadOnly(True)
        preview_layout.addWidget(self.preview_edit)
        preview_actions = QHBoxLayout()
        print_button = QPushButton("打印")
        print_button.clicked.connect(self.print_preview)
        export_button = QPushButton("导出 HTML")
        export_button.clicked.connect(self.export_html)
        preview_actions.addWidget(print_button)
        preview_actions.addWidget(export_button)
        preview_actions.addStretch(1)
        preview_layout.addLayout(preview_actions)
        right_layout.addWidget(preview_card, 1)

        body.addWidget(right)
        body.setStretchFactor(0, 2)
        body.setStretchFactor(1, 4)

    def _card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> None:
        label = QLabel(label_text)
        label.setProperty("formLabel", True)
        label.setFixedWidth(self.LABEL_WIDTH)
        layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

    def _with_file_button(self, line_edit: QLineEdit) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton("选择文件")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(lambda: self.choose_file(line_edit))
        row_layout.addWidget(button)
        return row

    def _with_folder_button(self, line_edit: QLineEdit) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        row_layout.addWidget(line_edit, 1)
        button = QPushButton("选择目录")
        button.setFixedWidth(self.ACTION_WIDTH)
        button.clicked.connect(lambda: self.choose_folder(line_edit))
        row_layout.addWidget(button)
        return row

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
        is_door = template.doc_type == DOC_TYPE_DOOR
        self.item_title.setText("考生单元模板" if not is_door else "门贴无需单元模板")
        self.item_editor.setEnabled(not is_door)
        self.columns_spin.setEnabled(template.doc_type in {DOC_TYPE_SEAT, DOC_TYPE_DESK})
        self.items_per_page_spin.setEnabled(template.doc_type == DOC_TYPE_DESK)

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
        self._select_guess(self.photo_match_column_combo, ["考号", "准考证号", "身份证号"])
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
        if not self.records or not room_column:
            return
        for value in room_values(self.records, room_column):
            self.room_value_combo.addItem(value, value)

    def _build_config(self) -> PrintDataConfig:
        excel_path = Path(self.excel_edit.text().strip())
        if not excel_path.exists():
            raise ValueError("请先选择有效的 Excel 数据文件。")
        room_column = str(self.room_column_combo.currentData() or "")
        if not room_column:
            raise ValueError("请先选择考场列。")
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
            if not room_value:
                self._refresh_room_values()
                room_value = str(self.room_value_combo.currentData() or "")
            if not room_value:
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
        self.log_fn(f"已生成考场打印预览：{room_value}")

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
