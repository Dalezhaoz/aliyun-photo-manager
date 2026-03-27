from __future__ import annotations

from datetime import date
from typing import Callable

from PySide6.QtCore import QDate
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDateEdit,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from ..id_card_tools import (
    build_id_card,
    get_region_description,
    list_cities,
    list_counties,
    list_provinces,
    resolve_region_code,
    validate_id_card,
)


class IdCardPage(QWidget):
    LABEL_WIDTH = 132

    def __init__(self, log_fn: Callable[[str], None]) -> None:
        super().__init__()
        self.log_fn = log_fn
        self.generated_values: list[str] = []
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = self._create_card()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)
        title = QLabel("身份证工具")
        title.setProperty("heroTitle", True)
        intro = QLabel("支持 18 位大陆居民身份证校验解析，以及按地区 / 出生日期 / 性别生成合法身份证。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        validate_card = self._create_card()
        validate_layout = QVBoxLayout(validate_card)
        validate_layout.setContentsMargins(24, 22, 24, 24)
        validate_layout.setSpacing(14)
        validate_title = QLabel("输入校验")
        validate_title.setProperty("sectionTitle", True)
        validate_layout.addWidget(validate_title)
        validate_row = QHBoxLayout()
        validate_row.setSpacing(10)
        self.id_input_edit = QLineEdit()
        self.id_input_edit.setPlaceholderText("输入 18 位身份证号")
        validate_row.addWidget(self.id_input_edit, 1)
        validate_button = QPushButton("校验并解析")
        validate_button.setProperty("accent", True)
        validate_button.setFixedWidth(132)
        validate_button.clicked.connect(self.run_validate)
        validate_row.addWidget(validate_button)
        validate_layout.addLayout(validate_row)
        root.addWidget(validate_card)

        generate_card = self._create_card()
        generate_layout = QVBoxLayout(generate_card)
        generate_layout.setContentsMargins(24, 22, 24, 24)
        generate_layout.setSpacing(18)
        generate_title = QLabel("身份证生成")
        generate_title.setProperty("sectionTitle", True)
        generate_layout.addWidget(generate_title)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        province_row = QWidget()
        province_layout = QHBoxLayout(province_row)
        province_layout.setContentsMargins(0, 0, 0, 0)
        province_layout.setSpacing(10)
        self.province_combo = QComboBox()
        self.province_combo.addItems(list_provinces())
        self.province_combo.currentTextChanged.connect(self.update_city_values)
        self.city_combo = QComboBox()
        self.city_combo.currentTextChanged.connect(self.update_county_values)
        self.county_combo = QComboBox()
        self.county_combo.currentTextChanged.connect(self.update_region_hint)
        province_layout.addWidget(self.province_combo, 1)
        province_layout.addWidget(self.city_combo, 1)
        province_layout.addWidget(self.county_combo, 1)
        self._add_row(form, 0, "省 / 市 / 县", province_row)

        self.custom_region_edit = QLineEdit()
        self.custom_region_edit.textChanged.connect(self.update_region_hint)
        custom_row = QWidget()
        custom_layout = QHBoxLayout(custom_row)
        custom_layout.setContentsMargins(0, 0, 0, 0)
        custom_layout.setSpacing(10)
        custom_layout.addWidget(self.custom_region_edit, 1)
        self.region_hint = QLabel("")
        self.region_hint.setWordWrap(True)
        custom_layout.addWidget(self.region_hint, 1)
        self._add_row(form, 1, "手工区划码", custom_row)

        self.birth_edit = QDateEdit()
        self.birth_edit.setCalendarPopup(True)
        self.birth_edit.setDate(QDate.currentDate().addYears(-20))
        self.birth_edit.setMaximumDate(QDate.currentDate())
        self._add_row(form, 2, "出生日期", self.birth_edit)

        gender_row = QWidget()
        gender_layout = QHBoxLayout(gender_row)
        gender_layout.setContentsMargins(0, 0, 0, 0)
        gender_layout.setSpacing(16)
        self.gender_male = QRadioButton("男")
        self.gender_female = QRadioButton("女")
        self.gender_male.setChecked(True)
        gender_layout.addWidget(self.gender_male)
        gender_layout.addWidget(self.gender_female)
        gender_layout.addStretch(1)
        self._add_row(form, 3, "性别", gender_row)

        action_row = QWidget()
        action_layout = QHBoxLayout(action_row)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(10)
        build_button = QPushButton("生成身份证")
        build_button.setProperty("accent", True)
        build_button.setFixedWidth(132)
        build_button.clicked.connect(self.run_generate)
        copy_button = QPushButton("复制结果")
        copy_button.setFixedWidth(132)
        copy_button.clicked.connect(self.copy_generated)
        action_layout.addWidget(build_button)
        action_layout.addWidget(copy_button)
        action_layout.addStretch(1)
        self._add_row(form, 4, "", action_row)

        self.generated_edit = QPlainTextEdit()
        self.generated_edit.setReadOnly(True)
        self.generated_edit.setFixedHeight(124)
        self._add_row(form, 5, "生成结果", self.generated_edit)
        generate_layout.addLayout(form)
        root.addWidget(generate_card)

        result_card = self._create_card()
        result_layout = QVBoxLayout(result_card)
        result_layout.setContentsMargins(24, 22, 24, 24)
        result_layout.setSpacing(14)
        result_title = QLabel("结果")
        result_title.setProperty("sectionTitle", True)
        result_layout.addWidget(result_title)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        result_layout.addWidget(self.result_text, 1)
        root.addWidget(result_card, 1)

        self.update_city_values(self.province_combo.currentText())

    def _create_card(self) -> QFrame:
        frame = QFrame()
        frame.setProperty("pageCard", True)
        return frame

    def _add_row(self, layout: QGridLayout, row: int, label_text: str, field: QWidget) -> None:
        if label_text:
            label = QLabel(label_text)
            label.setProperty("formLabel", True)
            label.setFixedWidth(self.LABEL_WIDTH)
            layout.addWidget(label, row, 0)
        layout.addWidget(field, row, 1)

    def update_city_values(self, province_name: str) -> None:
        self.city_combo.blockSignals(True)
        self.city_combo.clear()
        self.city_combo.addItems(list_cities(province_name))
        self.city_combo.blockSignals(False)
        self.update_county_values(self.city_combo.currentText())

    def update_county_values(self, city_name: str) -> None:
        self.county_combo.clear()
        self.county_combo.addItems(list_counties(self.province_combo.currentText(), city_name))
        self.update_region_hint()

    def update_region_hint(self) -> None:
        area_code = self.custom_region_edit.text().strip() or resolve_region_code(
            self.province_combo.currentText(),
            self.city_combo.currentText(),
            self.county_combo.currentText(),
        )
        self.region_hint.setText(get_region_description(area_code))

    def run_validate(self) -> None:
        result = validate_id_card(self.id_input_edit.text().strip())
        lines = [
            f"身份证号：{result.normalized_id or '无'}",
            f"是否合法：{'是' if result.is_valid else '否'}",
            f"出生日期：{result.birth_date or '无'}",
            f"年龄：{result.age if result.age is not None else '无'}",
            f"性别：{result.gender or '无'}",
            f"所在地：{result.location or '无'}",
            f"校验位：{result.check_code or '无'}",
            f"说明：{result.note}",
        ]
        self.result_text.setPlainText("\n".join(lines))

    def run_generate(self) -> None:
        area_code = self.custom_region_edit.text().strip() or resolve_region_code(
            self.province_combo.currentText(),
            self.city_combo.currentText(),
            self.county_combo.currentText(),
        )
        gender = "男" if self.gender_male.isChecked() else "女"
        try:
            generated_values: list[str] = []
            seen: set[str] = set()
            while len(generated_values) < 10:
                generated = build_id_card(area_code, self.birth_edit.date().toPython(), gender)
                if generated in seen:
                    continue
                seen.add(generated)
                generated_values.append(generated)
        except Exception as exc:
            QMessageBox.critical(self, "生成失败", str(exc))
            return
        self.generated_values = generated_values
        self.generated_edit.setPlainText("\n".join(generated_values))
        self.result_text.setPlainText(
            "已生成身份证：\n"
            + "\n".join(generated_values)
            + f"\n\n所在地：\n{get_region_description(area_code)}"
        )
        self.log_fn(f"已生成 10 个身份证号，首个：{generated_values[0]}")

    def copy_generated(self) -> None:
        value = self.generated_values[0] if self.generated_values else ""
        if not value:
            return
        QApplication.clipboard().setText(value)
        self.log_fn(f"已复制第一个身份证号：{value}")
