from __future__ import annotations

import calendar
from datetime import date

from ..id_card_tools import (
    build_id_card,
    get_region_description,
    list_cities,
    list_counties,
    resolve_region_code,
    validate_id_card,
)


def set_id_result_text(app, content: str) -> None:
    if app.id_result_text is None:
        return
    app.id_result_text.configure(state="normal")
    app.id_result_text.delete("1.0", "end")
    app.id_result_text.insert("1.0", content)
    app.id_result_text.configure(state="disabled")


def update_id_region_hint(app) -> None:
    custom_code = app.id_custom_region_code_var.get().strip()
    selected_code = resolve_region_code(
        app.id_province_var.get(),
        app.id_city_var.get(),
        app.id_county_var.get(),
    )
    region_code = custom_code or selected_code
    if region_code:
        app.id_region_hint_var.set(get_region_description(region_code))
    else:
        app.id_region_hint_var.set("请选择省 / 市 / 县，或手工填写 6 位区划码。")


def update_id_city_values(app) -> None:
    cities = list_cities(app.id_province_var.get())
    app.id_city_combo["values"] = cities
    if not cities:
        app.id_city_var.set("")
        app.id_county_combo["values"] = []
        app.id_county_var.set("")
        app.update_id_region_hint()
        return
    if app.id_city_var.get().strip() not in cities:
        app.id_city_var.set(cities[0])
    update_id_county_values(app)


def update_id_county_values(app) -> None:
    counties = list_counties(app.id_province_var.get(), app.id_city_var.get())
    app.id_county_combo["values"] = counties
    if not counties:
        app.id_county_var.set("")
        app.update_id_region_hint()
        return
    if app.id_county_var.get().strip() not in counties:
        app.id_county_var.set(counties[0])
    app.update_id_region_hint()


def update_id_day_values(app) -> None:
    try:
        year = int(app.id_birth_year_var.get().strip())
        month = int(app.id_birth_month_var.get().strip())
    except ValueError:
        return
    day_count = calendar.monthrange(year, month)[1]
    values = [f"{day:02d}" for day in range(1, day_count + 1)]
    app.id_birth_day_combo["values"] = values
    current_day = app.id_birth_day_var.get().strip()
    if current_day not in values:
        app.id_birth_day_var.set(values[-1])


def run_id_card_validate(app) -> None:
    app.save_settings()
    result = validate_id_card(app.id_input_var.get())
    if result.is_valid:
        text = (
            "身份证校验结果：合法\n"
            f"身份证号：{result.normalized_id}\n"
            f"出生日期：{result.birth_date}\n"
            f"年龄：{result.age}\n"
            f"性别：{result.gender}\n"
            f"所在地：{result.location}\n"
            f"校验位：{result.check_code}"
        )
    else:
        text = (
            "身份证校验结果：不合法\n"
            f"身份证号：{result.normalized_id or '未输入'}\n"
            f"出生日期：{result.birth_date or '无法解析'}\n"
            f"性别：{result.gender or '无法解析'}\n"
            f"所在地：{result.location or '无法解析'}\n"
            f"说明：{result.note}"
        )
    app.id_result_var.set(text)
    set_id_result_text(app, text)


def run_id_card_generate(app) -> None:
    try:
        year = int(app.id_birth_year_var.get().strip())
        month = int(app.id_birth_month_var.get().strip())
        day = int(app.id_birth_day_var.get().strip())
        birthday = date(year, month, day)
    except ValueError:
        raise ValueError("请选择有效的出生日期。")

    custom_code = app.id_custom_region_code_var.get().strip()
    selected_code = resolve_region_code(
        app.id_province_var.get(),
        app.id_city_var.get(),
        app.id_county_var.get(),
    )
    area_code = custom_code or selected_code
    if not area_code:
        raise ValueError("请选择省 / 市 / 县，或填写 6 位区划码。")

    gender = app.id_gender_var.get().strip() or "男"
    generated = build_id_card(area_code=area_code, birthday=birthday, gender=gender)
    result = validate_id_card(generated)
    text = (
        "身份证生成结果：\n"
        f"身份证号：{generated}\n"
        f"出生日期：{result.birth_date}\n"
        f"性别：{result.gender}\n"
        f"所在地：{result.location}\n"
        f"校验位：{result.check_code}"
    )
    app.id_generated_var.set(generated)
    app.id_result_var.set(text)
    set_id_result_text(app, text)
    app.save_settings()


def copy_generated_id_card(app) -> None:
    value = app.id_generated_var.get().strip()
    if not value:
        return
    app.root.clipboard_clear()
    app.root.clipboard_append(value)
    app.write_log("已复制生成的身份证号。")
