from __future__ import annotations

import random
import re
from dataclasses import dataclass
from datetime import date, datetime

from .id_region_data import REGION_TREE


ID_CARD_PATTERN = re.compile(r"^[1-9]\d{16}[\dXx]$")
CHECK_WEIGHTS = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
CHECK_CHARS = "10X98765432"

PROVINCE_NAMES: dict[str, str] = {}
REGION_NAME_BY_CODE: dict[str, tuple[str, str, str]] = {}
for province_name, cities in REGION_TREE.items():
    for city_name, counties in cities.items():
        for county_name, code in counties.items():
            PROVINCE_NAMES.setdefault(code[:2], province_name)
            REGION_NAME_BY_CODE[code] = (province_name, city_name, county_name)


@dataclass
class IdCardValidationResult:
    is_valid: bool
    normalized_id: str
    birth_date: str
    age: int | None
    gender: str
    location: str
    check_code: str
    note: str


def list_provinces() -> list[str]:
    return list(REGION_TREE.keys())


def list_cities(province_name: str) -> list[str]:
    return list(REGION_TREE.get((province_name or "").strip(), {}).keys())


def list_counties(province_name: str, city_name: str) -> list[str]:
    counties = list(
        REGION_TREE.get((province_name or "").strip(), {})
        .get((city_name or "").strip(), {})
        .keys()
    )
    if len(counties) > 1:
        counties = [name for name in counties if name != "市辖区"]
    return counties


def resolve_region_code(province_name: str, city_name: str, county_name: str) -> str:
    return (
        REGION_TREE.get((province_name or "").strip(), {})
        .get((city_name or "").strip(), {})
        .get((county_name or "").strip(), "")
    )


def calculate_check_code(body17: str) -> str:
    total = sum(int(char) * weight for char, weight in zip(body17, CHECK_WEIGHTS))
    return CHECK_CHARS[total % 11]


def _normalize_id_card(value: str) -> str:
    return (value or "").strip().upper()


def _format_location(province_name: str, city_name: str, county_name: str, area_code: str) -> str:
    parts = [province_name]
    if city_name and city_name != province_name:
        parts.append(city_name)
    if county_name and county_name != city_name:
        parts.append(county_name)
    return f"{' '.join(parts)}（{area_code}）"


def _resolve_location(area_code: str) -> str:
    exact = REGION_NAME_BY_CODE.get(area_code)
    if exact:
        return _format_location(*exact, area_code=area_code)
    province = PROVINCE_NAMES.get(area_code[:2])
    if province:
        return f"{province}（区划码 {area_code}）"
    return f"未知地区（{area_code}）"


def _parse_birth_date(raw: str) -> date:
    return datetime.strptime(raw, "%Y%m%d").date()


def _compute_age(birthday: date, today: date | None = None) -> int:
    current = today or date.today()
    age = current.year - birthday.year
    if (current.month, current.day) < (birthday.month, birthday.day):
        age -= 1
    return age


def validate_id_card(value: str) -> IdCardValidationResult:
    normalized = _normalize_id_card(value)
    if not normalized:
        return IdCardValidationResult(False, "", "", None, "", "", "", "请输入身份证号。")
    if not ID_CARD_PATTERN.fullmatch(normalized):
        return IdCardValidationResult(
            False,
            normalized,
            "",
            None,
            "",
            "",
            "",
            "身份证号格式不正确，只支持 18 位大陆居民身份证。",
        )

    area_code = normalized[:6]
    province = PROVINCE_NAMES.get(area_code[:2])
    if not province:
        return IdCardValidationResult(
            False,
            normalized,
            "",
            None,
            "",
            f"未知地区（{area_code}）",
            "",
            "行政区划码前两位无效。",
        )

    birth_raw = normalized[6:14]
    try:
        birthday = _parse_birth_date(birth_raw)
    except ValueError:
        return IdCardValidationResult(
            False,
            normalized,
            birth_raw,
            None,
            "",
            _resolve_location(area_code),
            "",
            "出生日期无效。",
        )

    expected = calculate_check_code(normalized[:17])
    actual = normalized[-1]
    gender = "男" if int(normalized[16]) % 2 == 1 else "女"
    if actual != expected:
        return IdCardValidationResult(
            False,
            normalized,
            birthday.isoformat(),
            _compute_age(birthday),
            gender,
            _resolve_location(area_code),
            expected,
            f"校验位不正确，应为 {expected}。",
        )

    return IdCardValidationResult(
        True,
        normalized,
        birthday.isoformat(),
        _compute_age(birthday),
        gender,
        _resolve_location(area_code),
        expected,
        "身份证号合法。",
    )


def build_id_card(area_code: str, birthday: date, gender: str) -> str:
    normalized_area = (area_code or "").strip()
    if not re.fullmatch(r"\d{6}", normalized_area):
        raise ValueError("区划码必须是 6 位数字。")
    if normalized_area[:2] not in PROVINCE_NAMES:
        raise ValueError("区划码前两位无效。")
    if birthday > date.today():
        raise ValueError("出生日期不能晚于今天。")

    parity_pool = (1, 3, 5, 7, 9) if gender == "男" else (0, 2, 4, 6, 8)
    serial = f"{random.randint(0, 99):02d}{random.choice(parity_pool)}"
    body17 = f"{normalized_area}{birthday:%Y%m%d}{serial}"
    return body17 + calculate_check_code(body17)


def get_region_description(area_code: str) -> str:
    normalized_area = (area_code or "").strip()
    if not normalized_area:
        return ""
    return _resolve_location(normalized_area)
