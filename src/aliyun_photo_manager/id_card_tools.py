from __future__ import annotations

import random
import re
from dataclasses import dataclass
from datetime import date, datetime


ID_CARD_PATTERN = re.compile(r"^[1-9]\d{16}[\dXx]$")
CHECK_WEIGHTS = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
CHECK_CHARS = "10X98765432"

PROVINCE_NAMES = {
    "11": "北京市",
    "12": "天津市",
    "13": "河北省",
    "14": "山西省",
    "15": "内蒙古自治区",
    "21": "辽宁省",
    "22": "吉林省",
    "23": "黑龙江省",
    "31": "上海市",
    "32": "江苏省",
    "33": "浙江省",
    "34": "安徽省",
    "35": "福建省",
    "36": "江西省",
    "37": "山东省",
    "41": "河南省",
    "42": "湖北省",
    "43": "湖南省",
    "44": "广东省",
    "45": "广西壮族自治区",
    "46": "海南省",
    "50": "重庆市",
    "51": "四川省",
    "52": "贵州省",
    "53": "云南省",
    "54": "西藏自治区",
    "61": "陕西省",
    "62": "甘肃省",
    "63": "青海省",
    "64": "宁夏回族自治区",
    "65": "新疆维吾尔自治区",
}

REGION_PRESETS = [
    ("110101", "北京市 东城区"),
    ("120101", "天津市 和平区"),
    ("130102", "河北省 石家庄市 长安区"),
    ("140105", "山西省 太原市 小店区"),
    ("150102", "内蒙古自治区 呼和浩特市 新城区"),
    ("210102", "辽宁省 沈阳市 和平区"),
    ("220102", "吉林省 长春市 南关区"),
    ("230102", "黑龙江省 哈尔滨市 道里区"),
    ("310101", "上海市 黄浦区"),
    ("320102", "江苏省 南京市 玄武区"),
    ("330102", "浙江省 杭州市 上城区"),
    ("340102", "安徽省 合肥市 瑶海区"),
    ("350102", "福建省 福州市 鼓楼区"),
    ("360102", "江西省 南昌市 东湖区"),
    ("370102", "山东省 济南市 历下区"),
    ("410102", "河南省 郑州市 中原区"),
    ("420102", "湖北省 武汉市 江岸区"),
    ("430102", "湖南省 长沙市 芙蓉区"),
    ("440103", "广东省 广州市 荔湾区"),
    ("450102", "广西壮族自治区 南宁市 兴宁区"),
    ("460105", "海南省 海口市 秀英区"),
    ("500103", "重庆市 渝中区"),
    ("510104", "四川省 成都市 锦江区"),
    ("520102", "贵州省 贵阳市 南明区"),
    ("530102", "云南省 昆明市 五华区"),
    ("540102", "西藏自治区 拉萨市 城关区"),
    ("610102", "陕西省 西安市 新城区"),
    ("620102", "甘肃省 兰州市 城关区"),
    ("630103", "青海省 西宁市 城中区"),
    ("640104", "宁夏回族自治区 银川市 兴庆区"),
    ("650102", "新疆维吾尔自治区 乌鲁木齐市 天山区"),
]

REGION_NAME_BY_CODE = {code: name for code, name in REGION_PRESETS}

REGION_TREE: dict[str, dict[str, dict[str, str]]] = {}
for code, name in REGION_PRESETS:
    parts = name.split()
    if len(parts) == 2:
        province_name, county_name = parts
        city_name = province_name
    else:
        province_name, city_name, county_name = parts[0], parts[1], parts[2]
    REGION_TREE.setdefault(province_name, {}).setdefault(city_name, {})[county_name] = code


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
    return list(
        REGION_TREE.get((province_name or "").strip(), {})
        .get((city_name or "").strip(), {})
        .keys()
    )


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


def _resolve_location(area_code: str) -> str:
    exact = REGION_NAME_BY_CODE.get(area_code)
    if exact:
        return f"{exact}（{area_code}）"
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


def extract_region_code(label_or_code: str) -> str:
    cleaned = (label_or_code or "").strip()
    if not cleaned:
        return ""
    token = cleaned.split()[0]
    return token if re.fullmatch(r"\d{6}", token) else ""


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
