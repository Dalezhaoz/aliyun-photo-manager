"""考场文件打印 V2 — UI 页面。"""
from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ..exam_print_v2 import (
    ColumnMapping,
    DeskSettings,
    DoorSettings,
    FieldConfig,
    GridSettings,
    SeatSettings,
    SigninSettings,
    export_desk_pdf,
    export_door_pdf,
    export_seat_pdf,
    export_signin_pdf,
    load_excel_headers,
    load_excel_records,
)
from .base_page import Card
from .common import AppComboBox
from .widgets import FormRow, PathInput, PrimaryButton, Section

_DEFAULT_FIELDS = ["姓名", "考号", "身份证号", "报考单位", "报考岗位"]

_COLUMN_GUESS = {
    "room_column": ["考场", "考场号"],
    "seat_column": ["序号", "座号"],
    "exam_no_column": ["考号", "准考证号"],
    "site_column": ["考点"],
    "subject_column": ["考试科目", "科目"],
    "unit_column": ["报考单位", "单位"],
    "job_column": ["报考岗位", "岗位"],
    "photo_match_column": ["身份证号", "证件号码", "身份证", "考号", "准考证号"],
    "sort_column": ["考号", "序号", "座号", "身份证号"],
}

_DOOR_FIELDS = [
    "考场", "考点", "考试科目", "准考证号起止", "考生人数", "单位岗位汇总",
]


class FieldConfigWidget(QWidget):
    """字段配置组件：勾选模式 + 文本模板模式。"""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self._group = QButtonGroup(self)
        mode_row = QHBoxLayout()
        self.radio_checkbox = QRadioButton("勾选字段")
        self.radio_template = QRadioButton("文本模板")
        self.radio_checkbox.setChecked(True)
        self._group.addButton(self.radio_checkbox)
        self._group.addButton(self.radio_template)
        mode_row.addWidget(self.radio_checkbox)
        mode_row.addWidget(self.radio_template)
        mode_row.addStretch(1)
        layout.addLayout(mode_row)

        self.checkbox_area = QScrollArea()
        self.checkbox_area.setWidgetResizable(True)
        self.checkbox_area.setFrameShape(QScrollArea.NoFrame)
        self.checkbox_widget = QWidget()
        self.checkbox_layout = QVBoxLayout(self.checkbox_widget)
        self.checkbox_layout.setContentsMargins(0, 0, 0, 0)
        self.checkbox_layout.setSpacing(4)
        self.checkbox_layout.addStretch(1)
        self.checkbox_area.setWidget(self.checkbox_widget)
        self.checkbox_area.setMinimumHeight(100)
        self.checkbox_area.setMaximumHeight(180)
        layout.addWidget(self.checkbox_area)

        self.template_edit = QPlainTextEdit()
        self.template_edit.setPlaceholderText("例：姓名：${姓名}\n考号：${考号}\n身份证号：${身份证号}")
        self.template_edit.setMinimumHeight(80)
        self.template_edit.setMaximumHeight(140)
        self.template_edit.setVisible(False)
        layout.addWidget(self.template_edit)

        self.radio_checkbox.toggled.connect(self._on_mode_changed)
        self.checkboxes: list[QCheckBox] = []

    def _on_mode_changed(self, checked: bool) -> None:
        self.checkbox_area.setVisible(checked)
        self.template_edit.setVisible(not checked)

    def rebuild(self, headers: list[str]) -> None:
        while self.checkbox_layout.count() > 1:
            item = self.checkbox_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.checkboxes.clear()
        for header in headers:
            cb = QCheckBox(header)
            if header in _DEFAULT_FIELDS:
                cb.setChecked(True)
            self.checkbox_layout.insertWidget(self.checkbox_layout.count() - 1, cb)
            self.checkboxes.append(cb)

    def get_config(self) -> FieldConfig:
        if self.radio_template.isChecked():
            return FieldConfig(mode="template", template_text=self.template_edit.toPlainText())
        fields = [cb.text() for cb in self.checkboxes if cb.isChecked()]
        return FieldConfig(mode="checkbox", checkbox_fields=fields)


class ExamPrintV2Page(QWidget):
    """考场文件打印 V2 页面。"""

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.records: list[dict[str, str]] = []
        self.headers: list[str] = []
        self._build_ui()

    # ── UI Construction ──────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        root.addWidget(self._build_data_card())

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_signin_tab(), "面试签到表")
        self.tabs.addTab(self._build_seat_tab(), "笔试座次表")
        self.tabs.addTab(self._build_door_tab(), "门贴")
        self.tabs.addTab(self._build_desk_tab(), "桌贴")
        root.addWidget(self.tabs, 1)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMaximumHeight(80)
        self.log_edit.setPlaceholderText("操作日志…")
        root.addWidget(self.log_edit)

    def _build_data_card(self) -> Card:
        card = Card("数据源")
        layout = card.body_layout

        self.excel_input = PathInput("选择 Excel 数据文件", "选择文件")
        self.excel_input.button.clicked.connect(self._choose_excel)
        layout.addWidget(self.excel_input)

        self.photo_input = PathInput("选择照片文件夹（可选）", "选择文件夹")
        self.photo_input.button.clicked.connect(self._choose_photo)
        layout.addWidget(self.photo_input)

        self.record_label = QLabel("未加载数据")
        layout.addWidget(self.record_label)

        mapping_title = QLabel("列映射（自动检测，可手动调整）")
        mapping_title.setProperty("formLabel", True)
        layout.addWidget(mapping_title)

        self.mapping_combos: dict[str, AppComboBox] = {}
        grid = QGridLayout()
        grid.setSpacing(6)
        mapping_fields = [
            ("room_column", "考场列"),
            ("seat_column", "座号列"),
            ("exam_no_column", "考号列"),
            ("site_column", "考点列"),
            ("subject_column", "科目列"),
            ("unit_column", "单位列"),
            ("job_column", "岗位列"),
            ("photo_match_column", "照片匹配列"),
            ("sort_column", "排序列"),
        ]
        for idx, (key, label) in enumerate(mapping_fields):
            row = idx // 3
            col = (idx % 3) * 2
            grid.addWidget(QLabel(label), row, col)
            combo = AppComboBox()
            self.mapping_combos[key] = combo
            grid.addWidget(combo, row, col + 1)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        grid.setColumnStretch(5, 1)
        layout.addLayout(grid)
        return card

    # ── Signin Tab ───────────────────────────────────────────────────────────

    def _build_signin_tab(self) -> QWidget:
        widget = QWidget()
        h_layout = QHBoxLayout(widget)
        h_layout.setContentsMargins(0, 0, 0, 0)

        left = QScrollArea()
        left.setWidgetResizable(True)
        left.setFrameShape(QScrollArea.NoFrame)
        left_inner = QWidget()
        left_layout = QVBoxLayout(left_inner)
        left_layout.setSpacing(12)

        title_section = Section("标题设置")
        self.signin_title = QLineEdit()
        self.signin_title.setPlaceholderText("面试签到表")
        title_section.add(FormRow("标题", self.signin_title))
        self.signin_subtitle = QLineEdit()
        title_section.add(FormRow("副标题", self.signin_subtitle))
        self.signin_align = AppComboBox()
        self.signin_align.addItems(["center", "left", "right"])
        title_section.add(FormRow("对齐", self.signin_align))
        left_layout.addWidget(title_section)

        grid_section = Section("网格设置")
        self.signin_cols = QSpinBox()
        self.signin_cols.setRange(1, 12)
        self.signin_cols.setValue(5)
        grid_section.add(FormRow("列数", self.signin_cols))
        self.signin_rows = QSpinBox()
        self.signin_rows.setRange(1, 12)
        self.signin_rows.setValue(6)
        grid_section.add(FormRow("行数", self.signin_rows))
        self.signin_corner = AppComboBox()
        self.signin_corner.addItems(["top_left", "top_right", "bottom_left", "bottom_right"])
        grid_section.add(FormRow("起始角", self.signin_corner))
        self.signin_direction = AppComboBox()
        self.signin_direction.addItems(["row", "column"])
        grid_section.add(FormRow("填充方向", self.signin_direction))
        self.signin_snake = QCheckBox("S 型绕行")
        grid_section.add(self.signin_snake)
        left_layout.addWidget(grid_section)

        field_section = Section("字段设置")
        self.signin_fields = FieldConfigWidget()
        field_section.add(self.signin_fields)
        left_layout.addWidget(field_section)

        photo_section = Section("照片")
        self.signin_photo_w = QSpinBox()
        self.signin_photo_w.setRange(0, 90)
        self.signin_photo_w.setValue(52)
        photo_section.add(FormRow("照片宽度", self.signin_photo_w))
        left_layout.addWidget(photo_section)

        left_layout.addStretch(1)
        left.setWidget(left_inner)
        h_layout.addWidget(left, 3)

        right = Card("导出")
        self.signin_summary = QLabel("请先加载 Excel 数据")
        right.body_layout.addWidget(self.signin_summary)
        signin_export = PrimaryButton("导出面试签到表 PDF")
        signin_export.clicked.connect(self._export_signin)
        right.body_layout.addWidget(signin_export)
        right.body_layout.addStretch(1)
        h_layout.addWidget(right, 1)
        return widget

    # ── Seat Tab ─────────────────────────────────────────────────────────────

    def _build_seat_tab(self) -> QWidget:
        widget = QWidget()
        h_layout = QHBoxLayout(widget)
        h_layout.setContentsMargins(0, 0, 0, 0)

        left = QScrollArea()
        left.setWidgetResizable(True)
        left.setFrameShape(QScrollArea.NoFrame)
        left_inner = QWidget()
        left_layout = QVBoxLayout(left_inner)
        left_layout.setSpacing(12)

        title_section = Section("标题设置")
        self.seat_title = QLineEdit()
        self.seat_title.setPlaceholderText("笔试座次表")
        title_section.add(FormRow("标题", self.seat_title))
        self.seat_subtitle = QLineEdit()
        title_section.add(FormRow("副标题", self.seat_subtitle))
        self.seat_align = AppComboBox()
        self.seat_align.addItems(["center", "left", "right"])
        title_section.add(FormRow("对齐", self.seat_align))
        left_layout.addWidget(title_section)

        display_section = Section("显示选项")
        self.seat_overview = QCheckBox("显示考场概览")
        self.seat_overview.setChecked(True)
        display_section.add(self.seat_overview)
        self.seat_proctor = QCheckBox("显示监考签字栏")
        self.seat_proctor.setChecked(True)
        display_section.add(self.seat_proctor)
        left_layout.addWidget(display_section)

        grid_section = Section("网格设置")
        self.seat_mode_group = QButtonGroup(self)
        self.seat_mode_simple = QRadioButton("M×N 模式")
        self.seat_mode_custom = QRadioButton("自定义列高")
        self.seat_mode_simple.setChecked(True)
        self.seat_mode_group.addButton(self.seat_mode_simple)
        self.seat_mode_group.addButton(self.seat_mode_custom)
        mode_widget = QWidget()
        mode_row = QHBoxLayout(mode_widget)
        mode_row.setContentsMargins(0, 0, 0, 0)
        mode_row.addWidget(self.seat_mode_simple)
        mode_row.addWidget(self.seat_mode_custom)
        mode_row.addStretch(1)
        grid_section.add(mode_widget)

        self.seat_cols = QSpinBox()
        self.seat_cols.setRange(1, 12)
        self.seat_cols.setValue(5)
        self.seat_cols_row = FormRow("列数", self.seat_cols)
        grid_section.add(self.seat_cols_row)
        self.seat_rows = QSpinBox()
        self.seat_rows.setRange(1, 12)
        self.seat_rows.setValue(6)
        self.seat_rows_row = FormRow("行数", self.seat_rows)
        grid_section.add(self.seat_rows_row)
        self.seat_custom_input = QLineEdit()
        self.seat_custom_input.setPlaceholderText("例：8,7,7,8")
        self.seat_custom_row = FormRow("自定义列高", self.seat_custom_input)
        self.seat_custom_row.setVisible(False)
        grid_section.add(self.seat_custom_row)

        self.seat_corner = AppComboBox()
        self.seat_corner.addItems(["top_left", "top_right", "bottom_left", "bottom_right"])
        grid_section.add(FormRow("起始角", self.seat_corner))
        self.seat_direction = AppComboBox()
        self.seat_direction.addItems(["column", "row"])
        grid_section.add(FormRow("填充方向", self.seat_direction))
        self.seat_snake = QCheckBox("S 型绕行")
        self.seat_snake.setChecked(True)
        grid_section.add(self.seat_snake)

        self.seat_mode_simple.toggled.connect(self._on_seat_mode_changed)
        left_layout.addWidget(grid_section)

        field_section = Section("字段设置")
        self.seat_fields = FieldConfigWidget()
        field_section.add(self.seat_fields)
        left_layout.addWidget(field_section)

        photo_section = Section("照片")
        self.seat_photo_w = QSpinBox()
        self.seat_photo_w.setRange(0, 90)
        self.seat_photo_w.setValue(52)
        photo_section.add(FormRow("照片宽度", self.seat_photo_w))
        left_layout.addWidget(photo_section)

        left_layout.addStretch(1)
        left.setWidget(left_inner)
        h_layout.addWidget(left, 3)

        right = Card("导出")
        self.seat_summary = QLabel("请先加载 Excel 数据")
        right.body_layout.addWidget(self.seat_summary)
        seat_export = PrimaryButton("导出笔试座次表 PDF")
        seat_export.clicked.connect(self._export_seat)
        right.body_layout.addWidget(seat_export)
        right.body_layout.addStretch(1)
        h_layout.addWidget(right, 1)
        return widget

    def _on_seat_mode_changed(self, simple_checked: bool) -> None:
        self.seat_cols_row.setVisible(simple_checked)
        self.seat_rows_row.setVisible(simple_checked)
        self.seat_custom_row.setVisible(not simple_checked)

    # ── Door Tab ─────────────────────────────────────────────────────────────

    def _build_door_tab(self) -> QWidget:
        widget = QWidget()
        h_layout = QHBoxLayout(widget)
        h_layout.setContentsMargins(0, 0, 0, 0)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setSpacing(12)

        section = Section("显示字段")
        self.door_checkboxes: list[QCheckBox] = []
        for field_name in _DOOR_FIELDS:
            cb = QCheckBox(field_name)
            if field_name in ["考场", "考点", "考试科目", "准考证号起止", "考生人数"]:
                cb.setChecked(True)
            section.add(cb)
            self.door_checkboxes.append(cb)
        left_layout.addWidget(section)
        left_layout.addStretch(1)
        h_layout.addWidget(left, 3)

        right = Card("导出")
        self.door_summary = QLabel("请先加载 Excel 数据")
        right.body_layout.addWidget(self.door_summary)
        door_export = PrimaryButton("导出所有考场门贴 PDF")
        door_export.clicked.connect(self._export_door)
        right.body_layout.addWidget(door_export)
        right.body_layout.addStretch(1)
        h_layout.addWidget(right, 1)
        return widget

    # ── Desk Tab ─────────────────────────────────────────────────────────────

    def _build_desk_tab(self) -> QWidget:
        widget = QWidget()
        h_layout = QHBoxLayout(widget)
        h_layout.setContentsMargins(0, 0, 0, 0)

        left = QScrollArea()
        left.setWidgetResizable(True)
        left.setFrameShape(QScrollArea.NoFrame)
        left_inner = QWidget()
        left_layout = QVBoxLayout(left_inner)
        left_layout.setSpacing(12)

        layout_section = Section("布局")
        self.desk_cols = QSpinBox()
        self.desk_cols.setRange(1, 6)
        self.desk_cols.setValue(2)
        layout_section.add(FormRow("每页列数", self.desk_cols))
        self.desk_items = QSpinBox()
        self.desk_items.setRange(1, 60)
        self.desk_items.setValue(10)
        layout_section.add(FormRow("每页个数", self.desk_items))
        self.desk_per_room = QCheckBox("按考场分文件")
        layout_section.add(self.desk_per_room)
        left_layout.addWidget(layout_section)

        field_section = Section("字段设置")
        self.desk_fields = FieldConfigWidget()
        field_section.add(self.desk_fields)
        left_layout.addWidget(field_section)

        photo_section = Section("照片")
        self.desk_photo_w = QSpinBox()
        self.desk_photo_w.setRange(0, 90)
        self.desk_photo_w.setValue(52)
        photo_section.add(FormRow("照片宽度", self.desk_photo_w))
        left_layout.addWidget(photo_section)

        left_layout.addStretch(1)
        left.setWidget(left_inner)
        h_layout.addWidget(left, 3)

        right = Card("导出")
        self.desk_summary = QLabel("请先加载 Excel 数据")
        right.body_layout.addWidget(self.desk_summary)
        desk_export = PrimaryButton("导出桌贴 PDF")
        desk_export.clicked.connect(self._export_desk)
        right.body_layout.addWidget(desk_export)
        right.body_layout.addStretch(1)
        h_layout.addWidget(right, 1)
        return widget

    # ── Data Loading ─────────────────────────────────────────────────────────

    def _choose_excel(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if selected:
            self.excel_input.edit.setText(selected)
            self._load_excel(Path(selected))

    def _choose_photo(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择照片文件夹")
        if selected:
            self.photo_input.edit.setText(selected)

    def _load_excel(self, path: Path) -> None:
        try:
            self.headers = load_excel_headers(path)
            self.records = load_excel_records(path)
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            return
        self.record_label.setText(f"共 {len(self.records)} 条记录")
        self._populate_mapping_combos()
        self.signin_fields.rebuild(self.headers)
        self.seat_fields.rebuild(self.headers)
        self.desk_fields.rebuild(self.headers)
        self._update_summaries()
        self.log_fn(f"已加载 Excel：{path.name}，共 {len(self.records)} 条。")

    def _populate_mapping_combos(self) -> None:
        for key, combo in self.mapping_combos.items():
            combo.clear()
            combo.addItem("未设置", "")
            for header in self.headers:
                combo.addItem(header, header)
        for key, candidates in _COLUMN_GUESS.items():
            combo = self.mapping_combos.get(key)
            if not combo:
                continue
            for candidate in candidates:
                idx = combo.findData(candidate)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
                    break

    def _update_summaries(self) -> None:
        n = len(self.records)
        cols = self.signin_cols.value()
        rows = self.signin_rows.value()
        per_page = cols * rows
        pages = (n + per_page - 1) // per_page if per_page else 0
        self.signin_summary.setText(f"共 {n} 条记录\n每页 {per_page} 人\n约 {pages} 页")

        s_cols = self.seat_cols.value()
        s_rows = self.seat_rows.value()
        self.seat_summary.setText(f"共 {n} 条记录\n布局：{s_cols}列×{s_rows}行")

        self.door_summary.setText(f"共 {n} 条记录\n按考场生成门贴")

        d_cols = self.desk_cols.value()
        d_items = self.desk_items.value()
        d_pages = (n + d_items - 1) // d_items if d_items else 0
        self.desk_summary.setText(f"共 {n} 条记录\n每页 {d_items} 个\n约 {d_pages} 页")

    # ── Config Builders ──────────────────────────────────────────────────────

    def _get_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            room_column=str(self.mapping_combos["room_column"].currentData() or ""),
            seat_column=str(self.mapping_combos["seat_column"].currentData() or ""),
            exam_no_column=str(self.mapping_combos["exam_no_column"].currentData() or ""),
            site_column=str(self.mapping_combos["site_column"].currentData() or ""),
            subject_column=str(self.mapping_combos["subject_column"].currentData() or ""),
            unit_column=str(self.mapping_combos["unit_column"].currentData() or ""),
            job_column=str(self.mapping_combos["job_column"].currentData() or ""),
            photo_match_column=str(self.mapping_combos["photo_match_column"].currentData() or ""),
            sort_column=str(self.mapping_combos["sort_column"].currentData() or ""),
        )

    def _get_photo_dir(self) -> Path | None:
        text = self.photo_input.edit.text().strip()
        return Path(text) if text else None

    def _get_signin_settings(self) -> SigninSettings:
        return SigninSettings(
            title=self.signin_title.text().strip() or "面试签到表",
            subtitle=self.signin_subtitle.text().strip(),
            title_align=self.signin_align.currentText(),
            grid=GridSettings(
                columns=self.signin_cols.value(),
                rows=self.signin_rows.value(),
                start_corner=self.signin_corner.currentText(),
                fill_direction=self.signin_direction.currentText(),
                snake=self.signin_snake.isChecked(),
            ),
            field_config=self.signin_fields.get_config(),
            photo_width=self.signin_photo_w.value(),
        )

    def _get_seat_settings(self) -> SeatSettings:
        custom = ""
        if self.seat_mode_custom.isChecked():
            custom = self.seat_custom_input.text().strip()
        return SeatSettings(
            title=self.seat_title.text().strip() or "笔试座次表",
            subtitle=self.seat_subtitle.text().strip(),
            title_align=self.seat_align.currentText(),
            show_room_overview=self.seat_overview.isChecked(),
            show_proctor_signature=self.seat_proctor.isChecked(),
            grid=GridSettings(
                columns=self.seat_cols.value(),
                rows=self.seat_rows.value(),
                custom_heights=custom,
                start_corner=self.seat_corner.currentText(),
                fill_direction=self.seat_direction.currentText(),
                snake=self.seat_snake.isChecked(),
            ),
            field_config=self.seat_fields.get_config(),
            photo_width=self.seat_photo_w.value(),
        )

    def _get_door_settings(self) -> DoorSettings:
        fields = [cb.text() for cb in self.door_checkboxes if cb.isChecked()]
        return DoorSettings(display_fields=fields)

    def _get_desk_settings(self) -> DeskSettings:
        return DeskSettings(
            columns=self.desk_cols.value(),
            items_per_page=self.desk_items.value(),
            field_config=self.desk_fields.get_config(),
            photo_width=self.desk_photo_w.value(),
            per_room=self.desk_per_room.isChecked(),
        )

    # ── Export ────────────────────────────────────────────────────────────────

    def _check_ready(self) -> bool:
        if not self.records:
            QMessageBox.warning(self, "提示", "请先加载 Excel 数据。")
            return False
        return True

    def _export_signin(self) -> None:
        if not self._check_ready():
            return
        default_name = f"{Path(self.excel_input.edit.text()).stem}_签到表.pdf"
        path = self._ask_save_path(default_name)
        if not path:
            return
        try:
            result = export_signin_pdf(path, self.records, self._get_mapping(), self._get_signin_settings(), self._get_photo_dir())
            self.log_fn(f"已导出面试签到表：{result[0]}")
            QMessageBox.information(self, "导出成功", f"已导出到：{result[0]}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def _export_seat(self) -> None:
        if not self._check_ready():
            return
        default_name = f"{Path(self.excel_input.edit.text()).stem}_座次表.pdf"
        path = self._ask_save_path(default_name)
        if not path:
            return
        try:
            result = export_seat_pdf(path, self.records, self._get_mapping(), self._get_seat_settings(), self._get_photo_dir())
            self.log_fn(f"已导出笔试座次表：{result[0]}")
            QMessageBox.information(self, "导出成功", f"已导出到：{result[0]}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def _export_door(self) -> None:
        if not self._check_ready():
            return
        default_name = f"{Path(self.excel_input.edit.text()).stem}_门贴.pdf"
        path = self._ask_save_path(default_name)
        if not path:
            return
        try:
            result = export_door_pdf(path, self.records, self._get_mapping(), self._get_door_settings())
            self.log_fn(f"已导出门贴：{result[0]}")
            QMessageBox.information(self, "导出成功", f"已导出到：{result[0]}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def _export_desk(self) -> None:
        if not self._check_ready():
            return
        default_name = f"{Path(self.excel_input.edit.text()).stem}_桌贴.pdf"
        path = self._ask_save_path(default_name)
        if not path:
            return
        try:
            result = export_desk_pdf(path, self.records, self._get_mapping(), self._get_desk_settings(), self._get_photo_dir())
            if len(result) == 1:
                message = f"已导出到：{result[0]}"
            else:
                message = f"已导出 {len(result)} 个文件到：{result[0].parent}"
            self.log_fn(f"已导出桌贴：{message}")
            QMessageBox.information(self, "导出成功", message)
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def _ask_save_path(self, default_name: str) -> Path | None:
        selected, _ = QFileDialog.getSaveFileName(self, "导出 PDF", default_name, "PDF 文件 (*.pdf)")
        if not selected:
            return None
        path = Path(selected)
        if path.suffix.lower() != ".pdf":
            path = path.with_suffix(".pdf")
        return path
