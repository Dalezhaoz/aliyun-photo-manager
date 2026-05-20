from __future__ import annotations

import json
import math
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QAction, QBrush, QColor, QDesktopServices, QFont, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressDialog,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from . import __version__
from .qt_ui import (
    CertificatePage,
    ExamPrintPage,
    ExamPage,
    IdCardPage,
    JobCodeAuditPage,
    MatchPage,
    OrderedUnitAuditPage,
    PackPage,
    PhoneDecryptPage,
    PhotoPage,
    ProjectStagePage,
    SqlExecPage,
    TemplateConvertPage,
    UpdateSqlPage,
)
from .release_config import load_release_config
from .update_manager import UpdateError, check_for_updates, download_update_package, launch_windows_updater


@dataclass(frozen=True)
class NavEntry:
    key: str
    label: str
    group: str
    description: str
    migrated: bool = False


NAV_GROUPS: list[tuple[str, str]] = [
    ("start", "开始使用"),
    ("files", "文件处理"),
    ("data", "数据处理"),
    ("db", "数据库工具"),
    ("query", "查询与辅助"),
    ("settings", "设置"),
    ("experimental", "实验功能"),
]


NAV_ENTRIES: list[NavEntry] = [
    NavEntry("home", "首页", "start", "查看常用入口和功能总览。", True),
    NavEntry("certificate", "证件资料筛选", "files", "支持云端按表过滤下载，并按模板动态分层导出。", True),
    NavEntry("template", "表样转换", "files", "将 Word / Excel 表样转换成 HTML。", True),
    NavEntry("photo", "照片下载与分类", "files", "支持云端目录选择、按表过滤下载和按模板分类。", True),
    NavEntry("pack", "结果打包", "files", "对任意结果文件或文件夹压缩并 AES 加密。", True),
    NavEntry("match", "数据匹配", "data", "按主键和附加匹配列补充来源表字段。", True),
    NavEntry("phone", "电话解密", "db", "通过 helper 解密电话并回写备用3。", True),
    NavEntry("update_sql", "更新 SQL 生成", "db", "通过字段映射模板生成标准 UPDATE SQL。", True),
    NavEntry("id_card", "身份证工具", "query", "校验并生成 18 位大陆居民身份证。", True),
    NavEntry("about", "关于", "settings", "查看版本与工具说明。", True),
    NavEntry("exam_print", "面试签到表打印", "experimental", "实验功能：按 Excel 和照片生成面试签到表。", True),
    NavEntry("job_code_audit", "岗位表核对", "experimental", "实验功能：核对地市、主管部门、报考单位、报考岗位编码及顺延关系。", True),
    NavEntry("ordered_name_audit", "单列顺序核对", "experimental", "实验功能：按原表顺序检查指定列的上下相似与分段重复。", True),
    NavEntry("exam", "考场编排", "experimental", "实验功能：按模板和规则生成考号、考场与座号。", True),
    NavEntry("sql_exec", "SQL 配置执行", "experimental", "实验功能：按 SQL 模板参数生成可执行脚本。", True),
    NavEntry("project_stage", "项目阶段汇总", "experimental", "实验功能：汇总多台 SQL Server 上的报名项目阶段状态。", True),
]


DEFAULT_FAVORITES = ["certificate", "template", "phone", "update_sql"]
FAVORITES_FILE = Path(__file__).resolve().parents[3] / ".favorites.json"


def _load_favorites() -> list[str]:
    if not FAVORITES_FILE.exists():
        return list(DEFAULT_FAVORITES)
    try:
        data = json.loads(FAVORITES_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return [item for item in data if isinstance(item, str)]
    except (json.JSONDecodeError, OSError):
        pass
    return list(DEFAULT_FAVORITES)


def _save_favorites(favorites: list[str]) -> None:
    FAVORITES_FILE.write_text(json.dumps(favorites, ensure_ascii=False, indent=2), encoding="utf-8")

MANUAL_TEXTS: dict[str, str] = {
    "photo": "适用场景：从本地目录或云存储批量下载照片，生成模板后再按模板分类。\n\n操作步骤：\n1. 先选择数据来源；云存储模式下填写云类型、Endpoint/Region、AccessKey 和 Bucket。\n2. 如需从云端指定目录下载，先加载 Bucket，再点“加载当前层级”，单击文件夹选中，双击进入子目录。\n3. 选择下载目录和分类输出目录。\n4. 如需只下载名单中的照片，可勾选“云下载时按表过滤”，上传过滤表并选择“前缀列”。\n5. 点击“生成模板”后，在 Excel 中补充分类列、名称等信息。\n6. 回到程序执行按模板分类，结果会输出分类目录和结果清单。",
    "certificate": "适用场景：从本地或云存储筛选证件资料，只导出指定文件夹或指定材料。\n\n操作步骤：\n1. 先选择数据来源；云存储模式下填写云类型、Endpoint/Region、AccessKey 和 Bucket。\n2. 如需从云端指定目录下载，先加载 Bucket，再点“加载当前层级”，单击文件夹选中，双击进入子目录。\n3. 选择处理模板并加载列，确认匹配列、导出名称列和分层列。\n4. 如需只下载表中的人员目录，可勾选“云下载时按表过滤”，上传过滤表并选择“文件夹名列”。\n5. 如需只导出某类材料，可切换为关键词文件模式并填写关键词；如需改名，可勾选导出后文件夹重命名。\n6. 运行后查看结果和结果清单，确认导出人数、文件数、分层列和输出位置。",
    "template": "适用场景：把 Word、Excel 表样模板转换成 HTML 代码。\n\n操作步骤：\n1. 选择表样文件，支持 doc、docx、xls、xlsx。\n2. 根据模板类型点击 Net 版导出或 Java 版导出。\n3. 程序会在结果区展示 HTML 和占位符内容。\n4. 可直接复制 HTML，后续粘贴到系统模板或页面中使用。\n5. 若模板异常，优先检查原始表样中的合并单元格、图片和特殊格式。",
    "match": "适用场景：用来源表字段补充目标表，替代手工 VLOOKUP / XLOOKUP。\n\n操作步骤：\n1. 选择目标表、来源表和输出文件。\n2. 点击加载表头，确认目标表匹配列与来源表匹配列。\n3. 如存在重复姓名或重复主键，可在附加匹配列里继续增加限制条件。\n4. 在补充列映射里填写结果列名，并选择来源表对应列。\n5. 点击开始匹配，程序会输出匹配结果文件和匹配清单。\n6. 若结果不符合预期，优先检查主键列格式、空值和附加匹配列是否一致。",
    "pack": "适用场景：把照片、证件资料或其他结果目录打包并加密交付。\n\n操作步骤：\n1. 选择待打包对象，可选文件或文件夹。\n2. 选择输出目录。\n3. 如需客户指定密码，勾选手动设置密码并输入密码；否则程序自动生成密码。\n4. 点击一键打包并加密，成功后右侧会显示压缩包路径、密码和历史记录。\n5. 复制密码后发给对方，后续可通过查询历史按文件名、来源名或密码回查。",
    "phone": "适用场景：按主键编号关联 web_info，加密电话解密后回写到考生表备用3。\n\n操作步骤：\n1. 填写服务器、数据库账号、报名库名和电话库名。\n2. 选择考生表和需要处理的模式，可全量处理，也可按名单处理。\n3. 程序会根据主键编号、考试代码、考试年月和考区参数调用 helper 解密。\n4. 解密成功后写回备用3，并在结果区显示成功、失败和跳过数量。\n5. 若失败，优先检查 helper 是否已构建、数据库连接是否正常、电话库中是否存在对应密文。",
    "update_sql": "适用场景：根据字段映射模板生成标准 UPDATE SQL，并在执行前自动备份相关表。\n\n操作步骤：\n1. 先准备字段映射模板，或点击导出模板生成标准样例。\n2. 选择映射模板并加载字段。\n3. 填写正式表、临时表以及双方关联字段。\n4. 如需防止空值覆盖正式表，可勾选忽略空值。\n5. 点击生成 SQL，在右侧检查备份语句、更新语句和 where 条件是否正确。\n6. 确认无误后复制 SQL，到数据库工具中执行。",
    "id_card": "适用场景：校验身份证号，或按地区、出生日期、性别生成测试数据。\n\n操作步骤：\n1. 输入校验区可直接输入 18 位身份证号，点击校验并解析查看出生日期、性别和地区。\n2. 生成区先选择省、市、县，也可手工填写 6 位区划码。\n3. 选择出生日期和性别后点击生成身份证。\n4. 程序会一次生成 10 个合法号码，复制结果默认复制第 1 个。\n5. 若用于测试，请不要把生成号码当作真实身份信息使用。",
    "exam": "适用场景：实验功能，用于按规则编排考号、考场和座号。\n\n操作步骤：\n1. 准备考生名单与编排规则模板。\n2. 按页面提示加载规则并设置考场容量、排序字段等参数。\n3. 先用小样本验证规则，再执行完整编排。\n4. 输出结果后重点检查考号连续性、考场容量和特殊考生分配是否正确。",
    "sql_exec": "适用场景：实验功能，用配置模板批量生成 SQL 语句。\n\n操作步骤：\n1. 选择 SQL 模板和参数文件。\n2. 加载后确认变量名、替换值和输出格式。\n3. 先在测试环境预览生成结果，再复制或导出执行。\n4. 如模板中包含删除、更新语句，请务必先备份再执行。",
    "project_stage": "适用场景：实验功能，汇总多台 SQL Server 上的报名项目阶段状态。\n\n操作步骤：\n1. 填写服务器连接信息，先执行测试连接。\n2. 配置需要查询的项目库、阶段表或汇总规则。\n3. 运行后查看结果区输出，确认每台服务器的阶段状态和统计值。\n4. 如查询慢或失败，优先检查网络、SQL Server 权限和超时设置。",
    "about": "当前版本以页面标题显示为准。\n\n本工具覆盖照片下载与分类、证件资料筛选、表样转换、数据匹配、结果打包、电话解密、更新 SQL 生成、身份证工具，以及实验功能页。\n本次已支持云端目录选择、云下载按表过滤，以及证件资料按所选列动态分层导出。"
}


class LogBridge(QObject):
    message = Signal(str)


class PaintedIconButton(QPushButton):
    def __init__(self, kind: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.kind = kind
        self.setCheckable(kind == "star")
        self.setCursor(Qt.PointingHandCursor)

    def set_kind(self, kind: str) -> None:
        self.kind = kind
        self.update()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        rect = self.rect().adjusted(1, 1, -1, -1)
        if self.kind == "star":
            painter.setPen(Qt.NoPen)
            painter.setBrush(Qt.NoBrush)
            self._paint_star(painter, rect)
            return

        bg = QColor("#FFFFFF")
        border = QColor("#CBD5E1")
        painter.setPen(QPen(border, 1.4))
        painter.setBrush(bg)
        painter.drawRoundedRect(rect, 14, 14)

        if self.kind == "menu-open":
            self._paint_bars(painter, rect, vertical=False)
        else:
            self._paint_bars(painter, rect, vertical=True)

    def _paint_bars(self, painter: QPainter, rect, vertical: bool) -> None:
        color = QColor("#2D3F5A")
        painter.setPen(QPen(color, 2.2, Qt.SolidLine, Qt.RoundCap))
        cx = rect.center().x()
        cy = rect.center().y()
        offsets = (-7, 0, 7)
        if vertical:
            for offset in offsets:
                painter.drawLine(cx + offset, cy - 6, cx + offset, cy + 6)
        else:
            for offset in offsets:
                painter.drawLine(cx - 6, cy + offset, cx + 6, cy + offset)

    def _paint_star(self, painter: QPainter, rect) -> None:
        color = QColor("#F5B301") if self.isChecked() else QColor("#7F8EA3")
        painter.setPen(QPen(color, 2))
        painter.setBrush(color if self.isChecked() else Qt.NoBrush)
        cx = rect.center().x()
        cy = rect.center().y()
        outer = 11
        inner = 4.6
        path = QPainterPath()
        for index in range(10):
            radius = outer if index % 2 == 0 else inner
            angle_deg = -90 + index * 36
            angle = angle_deg * 3.141592653589793 / 180
            x = cx + radius * math.cos(angle)
            y = cy + radius * math.sin(angle)
            if index == 0:
                path.moveTo(x, y)
            else:
                path.lineTo(x, y)
        path.closeSubpath()
        painter.drawPath(path)


class InteractiveFrame(QFrame):
    clicked = Signal()

    def __init__(self, object_name: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName(object_name)
        self.setCursor(Qt.PointingHandCursor)
        self.setProperty("pressed", False)
        self.setProperty("hovered", False)

    def enterEvent(self, event) -> None:
        self.setProperty("hovered", True)
        self.style().unpolish(self)
        self.style().polish(self)
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:
        self.setProperty("hovered", False)
        self.setProperty("pressed", False)
        self.style().unpolish(self)
        self.style().polish(self)
        super().leaveEvent(event)

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.LeftButton:
            self.setProperty("pressed", True)
            self.style().unpolish(self)
            self.style().polish(self)
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        was_pressed = bool(self.property("pressed"))
        self.setProperty("pressed", False)
        self.style().unpolish(self)
        self.style().polish(self)
        if was_pressed and self.rect().contains(event.position().toPoint()):
            self.clicked.emit()
        super().mouseReleaseEvent(event)


class PlaceholderPage(QWidget):
    def __init__(self, title: str, description: str) -> None:
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = QFrame()
        hero.setProperty("pageCard", True)
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)

        title_label = QLabel(title)
        title_label.setProperty("heroTitle", True)
        desc_label = QLabel(description)
        desc_label.setWordWrap(True)
        desc_label.setProperty("heroText", True)
        hero_layout.addWidget(title_label)
        hero_layout.addWidget(desc_label)
        root.addWidget(hero)

        info = QFrame()
        info.setProperty("pageCard", True)
        info_layout = QVBoxLayout(info)
        info_layout.setContentsMargins(24, 22, 24, 24)
        info_layout.setSpacing(14)

        section = QLabel("当前状态")
        section.setProperty("sectionTitle", True)
        info_layout.addWidget(section)

        body = QPlainTextEdit()
        body.setReadOnly(True)
        body.setPlainText("这个功能当前使用通用占位页展示，后续会继续补齐交互细节和页面说明。")
        info_layout.addWidget(body)
        root.addWidget(info)


class HomePage(QWidget):
    def __init__(self, open_callback, visible_entries: list[NavEntry], show_experimental: bool) -> None:
        super().__init__()
        self.open_callback = open_callback
        self.visible_entries = visible_entries
        self.show_experimental = show_experimental
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = QFrame()
        hero.setProperty("pageCard", True)
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)

        title = QLabel("报名系统工具箱")
        title.setProperty("heroTitle", True)
        intro = QLabel("左侧支持分组、搜索和常用功能，右侧按页面展开具体业务操作。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        migrated = QFrame()
        migrated.setProperty("pageCard", True)
        migrated_layout = QVBoxLayout(migrated)
        migrated_layout.setContentsMargins(24, 22, 24, 24)
        migrated_layout.setSpacing(14)

        section = QLabel("功能入口")
        section.setProperty("sectionTitle", True)
        migrated_layout.addWidget(section)

        preferred_keys = ["photo", "certificate", "template", "match", "pack", "phone", "update_sql", "id_card"]
        if self.show_experimental:
            preferred_keys.extend(["exam_print", "job_code_audit", "ordered_name_audit", "exam", "sql_exec", "project_stage"])
        visible_map = {entry.key: entry for entry in self.visible_entries}
        for key in preferred_keys:
            entry = visible_map.get(key)
            if entry is None:
                continue
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(12)
            label = QLabel(f"{entry.label}  ·  {entry.description}")
            label.setWordWrap(True)
            row_layout.addWidget(label, 1)
            button = QPushButton("打开")
            button.setFixedWidth(96)
            button.clicked.connect(lambda _=False, target=entry.key: self.open_callback(target))
            row_layout.addWidget(button)
            migrated_layout.addWidget(row)

        root.addWidget(migrated)


class HomeLandingPage(QWidget):
    def __init__(
        self,
        open_callback,
        entries_by_key: dict[str, NavEntry],
        favorites: list[str],
        show_experimental: bool,
    ) -> None:
        super().__init__()
        self.open_callback = open_callback
        self.entries_by_key = entries_by_key
        self.favorites = favorites
        self.show_experimental = show_experimental
        self.current_filter = "全部"
        self.current_view = "grid"
        self.filter_buttons: dict[str, QPushButton] = {}
        self.view_buttons: dict[str, QPushButton] = {}
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = QFrame()
        hero.setProperty("pageCard", True)
        hero.setObjectName("HomeHero")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(32, 28, 32, 28)
        hero_layout.setSpacing(18)

        title = QLabel("欢迎使用 报名系统工具箱")
        title.setProperty("heroTitle", True)
        title.setObjectName("HomeWelcomeTitle")
        intro = QLabel("首页已为您展示常用功能，您也可以通过左侧导航快速访问全部功能。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)

        stats_row = QHBoxLayout()
        stats_row.setSpacing(16)
        stats_row.addWidget(self._build_stat_card("全部功能", len(self._tool_entries()), "#2F73FF", "功"), 1)
        stats_row.addWidget(
            self._build_stat_card(
                "收藏功能",
                len([key for key in self.favorites if key in self.entries_by_key]),
                "#F59E0B",
                "藏",
            ),
            1,
        )
        hero_layout.addLayout(stats_row)
        root.addWidget(hero)

        favorites_card = QFrame()
        favorites_card.setProperty("pageCard", True)
        favorites_card.setObjectName("HomeSectionCard")
        favorites_layout = QVBoxLayout(favorites_card)
        favorites_layout.setContentsMargins(22, 20, 22, 22)
        favorites_layout.setSpacing(16)
        favorites_layout.addWidget(self._build_section_header("收藏功能", "快速访问您收藏的常用功能"))

        favorite_entries = self._favorite_entries()
        if favorite_entries:
            grid = QGridLayout()
            grid.setHorizontalSpacing(14)
            grid.setVerticalSpacing(14)
            cols = min(len(favorite_entries), 3)
            for index, entry in enumerate(favorite_entries):
                grid.addWidget(self._build_feature_card(entry, favorite=True), index // cols, index % cols)
            for column in range(cols):
                grid.setColumnStretch(column, 1)
            favorites_layout.addLayout(grid)
        else:
            empty = QLabel("还没有收藏功能，可以在功能页右上角点击星标加入收藏。")
            empty.setWordWrap(True)
            empty.setProperty("heroText", True)
            favorites_layout.addWidget(empty)
        root.addWidget(favorites_card)

        all_card = QFrame()
        all_card.setProperty("pageCard", True)
        all_card.setObjectName("HomeSectionCard")
        all_layout = QVBoxLayout(all_card)
        all_layout.setContentsMargins(22, 20, 22, 22)
        all_layout.setSpacing(14)
        all_layout.addWidget(self._build_section_header("全部功能", "所有可用功能分类展示"))

        tab_row = QHBoxLayout()
        tab_row.setSpacing(8)
        for text in ("全部", "文件处理", "数据处理", "数据库工具", "查询与辅助", "实验功能"):
            button = QPushButton(text)
            button.setObjectName("HomeFilterButton")
            button.setCheckable(True)
            button.clicked.connect(lambda _checked=False, value=text: self._set_filter(value))
            self.filter_buttons[text] = button
            tab_row.addWidget(button)
        tab_row.addStretch(1)
        grid_button = QPushButton("◫")
        grid_button.setObjectName("HomeViewToggle")
        grid_button.setCheckable(True)
        grid_button.clicked.connect(lambda _checked=False: self._set_view("grid"))
        grid_button.setFixedSize(36, 36)
        list_button = QPushButton("☰")
        list_button.setObjectName("HomeViewToggle")
        list_button.setCheckable(True)
        list_button.clicked.connect(lambda _checked=False: self._set_view("list"))
        list_button.setFixedSize(36, 36)
        self.view_buttons["grid"] = grid_button
        self.view_buttons["list"] = list_button
        tab_row.addWidget(grid_button)
        tab_row.addWidget(list_button)
        all_layout.addLayout(tab_row)

        self.all_content_host = QWidget()
        self.all_content_layout = QVBoxLayout(self.all_content_host)
        self.all_content_layout.setContentsMargins(0, 0, 0, 0)
        self.all_content_layout.setSpacing(0)
        all_layout.addWidget(self.all_content_host)
        self._sync_controls()
        self._rebuild_all_entries()

        root.addWidget(all_card)

    def _tool_entries(self) -> list[NavEntry]:
        keys = [
            "certificate",
            "template",
            "photo",
            "pack",
            "match",
            "phone",
            "update_sql",
            "id_card",
        ]
        if self.show_experimental:
            keys.extend(["exam_print", "job_code_audit", "ordered_name_audit", "exam", "sql_exec", "project_stage"])
        return [self.entries_by_key[key] for key in keys if key in self.entries_by_key]

    def _favorite_entries(self) -> list[NavEntry]:
        keys: list[str] = []
        for key in self.favorites:
            if key in self.entries_by_key and key not in keys:
                keys.append(key)
        if not keys:
            for key in ["certificate", "template", "phone", "update_sql"]:
                if key in self.entries_by_key and key not in keys:
                    keys.append(key)
        return [self.entries_by_key[key] for key in keys]

    def _landing_entries(self) -> list[dict[str, object]]:
        entries: list[dict[str, object]] = []
        mapping = [
            ("certificate", "证件资料筛选", "支持云端按表过滤下载，并按模板动态分层导出。", "文件处理"),
            ("template", "表样转换", "将 Word / Excel 表样转换成 HTML。", "文件处理"),
            ("photo", "照片下载与分类", "批量下载照片并按规则分类整理。", "文件处理"),
            ("pack", "结果打包", "将处理结果按规则打包成压缩文件。", "文件处理"),
            ("match", "数据匹配", "按主键和附加匹配列补充来源表字段。", "数据处理"),
            ("phone", "电话解密", "通过 helper 解密电话并回写备用3。", "数据库工具"),
            ("update_sql", "更新 SQL 生成", "通过字段映射模板生成标准 UPDATE SQL。", "数据库工具"),
            ("id_card", "身份证工具", "校验并生成 18 位大陆居民身份证。", "查询与辅助"),
        ]
        if self.show_experimental:
            mapping.extend([
                ("exam_print", "面试签到表打印", "按 Excel 和照片生成面试签到表。", "实验功能"),
                ("job_code_audit", "岗位表核对", "校验岗位编码格式、绑定关系和顺延规则。", "实验功能"),
                ("ordered_name_audit", "单列顺序核对", "按原表顺序检查指定列的上下相似与分段重复。", "实验功能"),
                ("exam", "考场编排", "按模板和规则编排考号、考场和座号。", "实验功能"),
                ("sql_exec", "SQL 配置执行", "按 SQL 模板参数生成可执行脚本。", "实验功能"),
                ("project_stage", "项目阶段汇总", "汇总多台 SQL Server 的项目阶段状态。", "实验功能"),
            ])
        mapping.append(("about", "关于", f"报名系统工具箱 v{__version__}，查看版本与工具说明。", "设置"))
        for key, title, description, category in mapping:
            entry = self.entries_by_key.get(key)
            if entry is None:
                continue
            entries.append({"entry": entry, "title": title, "description": description, "category": category})
        return entries

    def _filtered_landing_entries(self) -> list[dict[str, object]]:
        entries = self._landing_entries()
        if self.current_filter == "全部":
            return entries
        return [item for item in entries if item["category"] == self.current_filter]

    def _set_filter(self, value: str) -> None:
        self.current_filter = value
        self._sync_controls()
        self._rebuild_all_entries()

    def _set_view(self, value: str) -> None:
        self.current_view = value
        self._sync_controls()
        self._rebuild_all_entries()

    def _sync_controls(self) -> None:
        for key, button in self.filter_buttons.items():
            button.setChecked(key == self.current_filter)
        for key, button in self.view_buttons.items():
            button.setChecked(key == self.current_view)

    def _clear_layout(self, layout: QVBoxLayout) -> None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            child_layout = item.layout()
            if widget is not None:
                widget.deleteLater()
            elif child_layout is not None:
                self._clear_nested_layout(child_layout)

    def _clear_nested_layout(self, layout) -> None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            nested = item.layout()
            if widget is not None:
                widget.deleteLater()
            elif nested is not None:
                self._clear_nested_layout(nested)

    def _rebuild_all_entries(self) -> None:
        self._clear_layout(self.all_content_layout)
        entries = self._filtered_landing_entries()
        if self.current_view == "grid":
            grid = QGridLayout()
            grid.setContentsMargins(0, 0, 0, 0)
            grid.setHorizontalSpacing(14)
            grid.setVerticalSpacing(12)
            for index, item in enumerate(entries):
                grid.addWidget(
                    self._build_feature_row(item["entry"], item["title"], item["description"]),
                    index // 2,
                    index % 2,
                )
            self.all_content_layout.addLayout(grid)
        else:
            column = QVBoxLayout()
            column.setContentsMargins(0, 0, 0, 0)
            column.setSpacing(12)
            for item in entries:
                column.addWidget(self._build_feature_row(item["entry"], item["title"], item["description"]))
            self.all_content_layout.addLayout(column)

    def _build_section_header(self, title_text: str, desc_text: str) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(10)
        marker = QLabel("")
        marker.setObjectName("HomeSectionMarker")
        marker.setFixedSize(10, 18)
        row_layout.addWidget(marker)
        title = QLabel(title_text)
        title.setObjectName("HomeSectionTitle")
        desc = QLabel(desc_text)
        desc.setProperty("heroText", True)
        row_layout.addWidget(title)
        row_layout.addWidget(desc)
        row_layout.addStretch(1)
        return row

    def _build_stat_card(self, label: str, value: int, color: str, icon_text: str) -> QWidget:
        card = QFrame()
        card.setObjectName("HomeStatCard")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QHBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)
        icon = QLabel(icon_text)
        icon.setObjectName("HomeStatIcon")
        icon.setStyleSheet(f"color: {color};")
        icon.setAlignment(Qt.AlignCenter)
        icon.setFixedSize(38, 38)
        icon.setMinimumSize(32, 32)
        layout.addWidget(icon)
        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        value_label = QLabel(str(value))
        value_label.setObjectName("HomeStatValue")
        value_label.setStyleSheet(f"color: {color};")
        label_widget = QLabel(label)
        label_widget.setObjectName("HomeStatLabel")
        text_box.addWidget(value_label)
        text_box.addWidget(label_widget)
        layout.addLayout(text_box, 1)
        return card

    def _build_feature_card(self, entry: NavEntry, *, favorite: bool = False) -> QWidget:
        card = InteractiveFrame("HomeFeatureCard")
        card.clicked.connect(lambda target=entry.key: self.open_callback(target))
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 12)
        layout.setSpacing(10)
        accent = self._entry_accent(entry.key)
        top = QHBoxLayout()
        top.setSpacing(12)
        icon = QLabel(self._entry_icon(entry))
        icon.setObjectName("HomeFeatureIcon")
        icon.setStyleSheet(
            f"color: {accent};"
            f"background: {self._alpha_color(accent, 0.10)};"
            f"border: 1px solid {self._alpha_color(accent, 0.18)};"
        )
        icon.setAlignment(Qt.AlignCenter)
        icon.setFixedSize(42, 42)
        top.addWidget(icon)
        text_block = QVBoxLayout()
        text_block.setSpacing(3)
        title = QLabel(entry.label)
        title.setObjectName("HomeFeatureTitle")
        desc = QLabel(entry.description)
        desc.setWordWrap(True)
        desc.setProperty("heroText", True)
        text_block.addWidget(title)
        text_block.addWidget(desc)
        top.addLayout(text_block, 1)
        if favorite:
            star = QLabel("★")
            star.setObjectName("HomeFavoriteStar")
            top.addWidget(star, 0, Qt.AlignTop)
        layout.addLayout(top)
        button = QPushButton("打开")
        button.setObjectName("HomeOpenButton")
        button.setCursor(Qt.PointingHandCursor)
        button.clicked.connect(lambda _=False, target=entry.key: self.open_callback(target))
        layout.addWidget(button)
        return card

    def _build_feature_row(self, entry: NavEntry, title_text: str | None = None, desc_text: str | None = None) -> QWidget:
        row = InteractiveFrame("HomeFeatureRow")
        row.clicked.connect(lambda target=entry.key: self.open_callback(target))
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(14, 12, 14, 12)
        row_layout.setSpacing(12)
        accent = self._entry_accent(entry.key)
        icon = QLabel(self._entry_icon(entry))
        icon.setObjectName("HomeFeatureIconSmall")
        icon.setStyleSheet(
            f"color: {accent};"
            f"background: {self._alpha_color(accent, 0.10)};"
            f"border: 1px solid {self._alpha_color(accent, 0.18)};"
        )
        icon.setAlignment(Qt.AlignCenter)
        icon.setFixedSize(38, 38)
        row_layout.addWidget(icon)
        text_block = QVBoxLayout()
        text_block.setSpacing(3)
        title = QLabel(title_text or entry.label)
        title.setObjectName("HomeFeatureTitle")
        desc = QLabel(desc_text or entry.description)
        desc.setWordWrap(True)
        desc.setProperty("heroText", True)
        text_block.addWidget(title)
        text_block.addWidget(desc)
        row_layout.addLayout(text_block, 1)
        arrow = QLabel("›")
        arrow.setObjectName("HomeRowArrow")
        row_layout.addWidget(arrow)
        return row

    def _entry_icon(self, entry: NavEntry) -> str:
        icons = {
            "certificate": "证",
            "template": "表",
            "photo": "图",
            "pack": "包",
            "match": "数",
            "phone": "电",
            "update_sql": "SQL",
            "id_card": "证",
            "exam_print": "印",
            "job_code_audit": "核",
            "ordered_name_audit": "序",
            "exam": "考",
            "sql_exec": "库",
            "project_stage": "项",
        }
        return icons.get(entry.key, "工")

    def _entry_accent(self, key: str) -> str:
        accents = {
            "certificate": "#2F73FF",
            "template": "#16A35F",
            "photo": "#3B82F6",
            "pack": "#F59E0B",
            "match": "#2563EB",
            "phone": "#9333EA",
            "update_sql": "#F97316",
            "id_card": "#4F46E5",
            "exam_print": "#0F766E",
            "job_code_audit": "#B45309",
            "ordered_name_audit": "#7C3AED",
            "exam": "#DC2626",
            "sql_exec": "#1D4ED8",
            "project_stage": "#059669",
        }
        return accents.get(key, "#2F73FF")

    def _alpha_color(self, color: str, alpha: float) -> str:
        qcolor = QColor(color)
        qcolor.setAlphaF(alpha)
        return qcolor.name(QColor.HexArgb)


class AboutPage(QWidget):
    def __init__(self, show_experimental: bool, update_callback=None) -> None:
        super().__init__()
        self.show_experimental = show_experimental
        self.update_callback = update_callback
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(16)

        hero = QFrame()
        hero.setProperty("pageCard", True)
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(24, 22, 24, 22)
        hero_layout.setSpacing(8)

        title = QLabel("关于")
        title.setProperty("heroTitle", True)
        intro = QLabel(f"报名系统工具箱 v{__version__}，用于报名业务中的文件处理、数据处理、数据库辅助和查询工具。")
        intro.setWordWrap(True)
        intro.setProperty("heroText", True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(intro)
        root.addWidget(hero)

        if self.update_callback and sys.platform == "win32":
            update_card = QFrame()
            update_card.setProperty("pageCard", True)
            update_layout = QVBoxLayout(update_card)
            update_layout.setContentsMargins(24, 22, 24, 24)
            update_layout.setSpacing(14)
            update_title = QLabel("检查更新")
            update_title.setProperty("sectionTitle", True)
            update_layout.addWidget(update_title)
            update_btn = QPushButton("检查更新")
            update_btn.setProperty("accent", True)
            update_btn.setFixedWidth(150)
            update_btn.clicked.connect(self.update_callback)
            update_layout.addWidget(update_btn)
            update_hint = QLabel("点击检查是否有新版本可用，支持增量更新和整包更新。")
            update_hint.setWordWrap(True)
            update_hint.setProperty("heroText", True)
            update_layout.addWidget(update_hint)
            root.addWidget(update_card)

        capability = QFrame()
        capability.setProperty("pageCard", True)
        capability_layout = QVBoxLayout(capability)
        capability_layout.setContentsMargins(24, 22, 24, 24)
        capability_layout.setSpacing(14)
        capability_title = QLabel("主要功能")
        capability_title.setProperty("sectionTitle", True)
        capability_layout.addWidget(capability_title)

        features = [
            "照片下载与分类：支持本地目录和云存储下载后生成模板、按模板分类。",
            "证件资料筛选：按模板列筛选资料目录，也支持先下载后处理。",
            "表样转换：将 Word / Excel 表样转换成 HTML。",
            "数据匹配：按主键和附加匹配列补充来源表字段。",
            "结果打包：对结果文件或文件夹进行压缩与 AES 加密。",
            "电话解密：通过 helper 解密电话并回写备用3。",
            "更新 SQL 生成：按字段映射模板生成标准 UPDATE SQL。",
            "身份证工具：校验 18 位身份证并批量生成测试号码。",
        ]
        if self.show_experimental:
            features.extend(
                [
                    "面试签到表打印：按 Excel 和照片生成面试签到表。",
                    "考场编排：按模板和规则编排考号、考场和座号。",
                    "SQL 配置执行：按模板参数生成可执行 SQL。",
                    "项目阶段汇总：汇总多台 SQL Server 上的项目阶段状态。",
                    "岗位表核对：校验岗位编码格式和层级顺延关系。",
                    "单列顺序核对：检查指定列上下相似和分段重复。",
                ]
            )
        for feature in features:
            label = QLabel(feature)
            label.setWordWrap(True)
            capability_layout.addWidget(label)
        root.addWidget(capability)


class QtMainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self._initial_window_placed = False
        self.setWindowFlags(
            Qt.Window
            | Qt.WindowTitleHint
            | Qt.WindowSystemMenuHint
            | Qt.WindowMinMaxButtonsHint
            | Qt.WindowCloseButtonHint
        )
        self.release_config = load_release_config()
        self.setWindowTitle(f"报名系统工具箱 v{__version__}")
        self.resize(1220, 760)
        self.setMinimumSize(1120, 700)
        self.setMinimumSize(1180, 760)

        font = QFont("Microsoft YaHei UI", 10)
        QApplication.instance().setFont(font)

        self.log_bridge = LogBridge()
        self.log_bridge.message.connect(self.append_log)

        self.visible_nav_entries = [
            entry for entry in NAV_ENTRIES if self.release_config.show_experimental or entry.group != "experimental"
        ]
        self.visible_nav_groups = [
            group for group in NAV_GROUPS if self.release_config.show_experimental or group[0] != "experimental"
        ]
        self.entries_by_key = {entry.key: entry for entry in self.visible_nav_entries}
        self.favorites: list[str] = _load_favorites()
        self.page_indexes: dict[str, int] = {}
        self.tree_items_by_key: dict[str, QTreeWidgetItem] = {}
        self.page_help_sections: dict[str, QFrame] = {}
        self.page_action_buttons: dict[str, QPushButton] = {}
        self.sidebar_collapsed = False

        central = QWidget()
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(12)
        self.setCentralWidget(central)
        self.central_panel = central

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(10)
        root_layout.addWidget(splitter, 1)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setProperty("collapsed", False)
        self.sidebar.setMinimumWidth(230)
        self.sidebar.setMaximumWidth(280)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(16, 16, 16, 16)
        sidebar_layout.setSpacing(12)
        sidebar_layout.setAlignment(Qt.AlignTop)
        self.sidebar_layout = sidebar_layout

        sidebar_header = QWidget()
        sidebar_header_layout = QVBoxLayout(sidebar_header)
        sidebar_header_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_header_layout.setSpacing(10)

        sidebar_top_row = QHBoxLayout()
        sidebar_top_row.setContentsMargins(0, 0, 0, 0)
        sidebar_top_row.setSpacing(10)
        self.sidebar_brand_icon = QLabel("工")
        self.sidebar_brand_icon.setObjectName("SidebarBrandIcon")
        self.sidebar_brand_icon.setAlignment(Qt.AlignCenter)
        self.sidebar_brand_icon.setFixedSize(56, 56)
        sidebar_top_row.addWidget(self.sidebar_brand_icon, 0, Qt.AlignTop)
        title_block = QVBoxLayout()
        title_block.setSpacing(2)
        self.sidebar_app_title = QLabel("报名系统工具箱")
        self.sidebar_app_title.setProperty("appTitle", True)
        self.sidebar_app_subtitle = QLabel(f"v{__version__}")
        self.sidebar_app_subtitle.setProperty("appSubtitle", True)
        title_block.addWidget(self.sidebar_app_title)
        title_block.addWidget(self.sidebar_app_subtitle)
        sidebar_top_row.addLayout(title_block, 1)
        self.sidebar_toggle_button = PaintedIconButton("menu-open")
        self.sidebar_toggle_button.setObjectName("MiniButton")
        self.sidebar_toggle_button.clicked.connect(self.toggle_sidebar)
        sidebar_top_row.addWidget(self.sidebar_toggle_button, 0, Qt.AlignTop)
        sidebar_header_layout.addLayout(sidebar_top_row)
        self.sidebar_overlay_button = PaintedIconButton("menu-collapsed", self.central_panel)
        self.sidebar_overlay_button.setObjectName("MiniButton")
        self.sidebar_overlay_button.setProperty("collapsed", True)
        self.sidebar_overlay_button.clicked.connect(self.toggle_sidebar)
        self.sidebar_overlay_button.hide()

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索功能，例如：电话 / 模板 / SQL")
        self.search_edit.textChanged.connect(self.filter_navigation)
        sidebar_header_layout.addWidget(self.search_edit)
        sidebar_layout.addWidget(sidebar_header)

        self.sidebar_body = QWidget()
        sidebar_body_layout = QVBoxLayout(self.sidebar_body)
        sidebar_body_layout.setContentsMargins(0, 8, 0, 0)
        sidebar_body_layout.setSpacing(12)

        self.sidebar_title = QLabel("导航")
        self.sidebar_title.setProperty("sectionTitle", True)
        sidebar_body_layout.addWidget(self.sidebar_title)
        self.sidebar_title.hide()

        nav_title_row = QHBoxLayout()
        self.nav_title = QLabel("全部功能")
        self.nav_title.setProperty("sectionTitle", True)
        nav_title_row.addWidget(self.nav_title)
        nav_title_row.addStretch(1)
        sidebar_body_layout.addLayout(nav_title_row)
        self.nav_title.hide()
        self.nav_tree = QTreeWidget()
        self.nav_tree.setHeaderHidden(True)
        self.nav_tree.setIndentation(14)
        self.nav_tree.setObjectName("NavTree")
        self.nav_tree.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.nav_tree.itemPressed.connect(self.on_tree_item_pressed)
        self.nav_tree.itemClicked.connect(self.on_tree_item_clicked)
        sidebar_body_layout.addWidget(self.nav_tree, 1)
        sidebar_layout.addWidget(self.sidebar_body, 1)
        self.sidebar_footer_button = QPushButton("收藏夹")
        self.sidebar_footer_button.setObjectName("SidebarFooterButton")
        self.sidebar_footer_button.clicked.connect(lambda: self.open_entry("home"))
        sidebar_layout.addWidget(self.sidebar_footer_button)

        content_wrapper = QWidget()
        content_layout = QVBoxLayout(content_wrapper)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        self.content_container = QWidget()
        container_layout = QVBoxLayout(self.content_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(12)
        self.header_toolbar = QWidget(self.content_container)
        self.header_toolbar.setObjectName("HeaderToolbar")
        toolbar_layout = QHBoxLayout(self.header_toolbar)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        toolbar_layout.setSpacing(10)
        toolbar_layout.addStretch(1)
        self.header_update_button = QPushButton("检查更新", self.content_container)
        self.header_update_button.setObjectName("HelpButton")
        self.header_update_button.setCursor(Qt.PointingHandCursor)
        self.header_update_button.setFixedSize(104, 32)
        self.header_update_button.clicked.connect(self.run_update_check)
        self.header_back_button = QPushButton("返回首页", self.content_container)
        self.header_back_button.setObjectName("HelpButton")
        self.header_back_button.setCursor(Qt.PointingHandCursor)
        self.header_back_button.setFixedSize(104, 32)
        self.header_back_button.clicked.connect(lambda: self.open_entry("home"))
        self.header_favorite_button = QPushButton("收藏", self.content_container)
        self.header_favorite_button.setObjectName("HelpButton")
        self.header_favorite_button.setCursor(Qt.PointingHandCursor)
        self.header_favorite_button.setFixedSize(84, 32)
        self.header_favorite_button.clicked.connect(self.toggle_current_favorite)
        self.header_settings_button = QPushButton("设置", self.content_container)
        self.header_settings_button.setObjectName("HelpButton")
        self.header_settings_button.setCursor(Qt.PointingHandCursor)
        self.header_settings_button.setFixedSize(84, 32)
        self.header_settings_button.clicked.connect(lambda: self.open_entry("about"))
        toolbar_layout.addWidget(self.header_back_button)
        toolbar_layout.addWidget(self.header_favorite_button)
        toolbar_layout.addWidget(self.header_settings_button)
        toolbar_layout.addWidget(self.header_update_button)
        self.stack = QStackedWidget()
        container_layout.addWidget(self.header_toolbar)
        container_layout.addWidget(self.stack)
        self.header_toolbar.hide()
        content_layout.addWidget(self.content_container, 1)

        splitter.addWidget(self.sidebar)
        splitter.addWidget(content_wrapper)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setHandleWidth(0)
        self.sidebar.hide()
        self.sidebar.setMinimumWidth(0)
        self.sidebar.setMaximumWidth(0)
        splitter.setSizes([0, 1400])
        self.main_splitter = splitter

        self.sidebar_overlay_button.hide()

        self._build_navigation_tree()
        self._build_pages()
        self._build_menu()
        self._apply_styles()

        self.open_entry("home")

    def showEvent(self, event) -> None:
        super().showEvent(event)
        if not self._initial_window_placed:
            self._ensure_reasonable_window_geometry()
            self._initial_window_placed = True

    def _build_pages(self) -> None:
        for entry in self.visible_nav_entries:
            if entry.key == "home":
                page = self._create_home_page()
            elif entry.key == "photo":
                page = PhotoPage(self.emit_log)
            elif entry.key == "certificate":
                page = CertificatePage(self.emit_log)
            elif entry.key == "template":
                page = TemplateConvertPage(self.emit_log)
            elif entry.key == "match":
                page = MatchPage(self.emit_log)
            elif entry.key == "pack":
                page = PackPage(self.emit_log)
            elif entry.key == "phone":
                page = PhoneDecryptPage(self.emit_log)
            elif entry.key == "update_sql":
                page = UpdateSqlPage(self.emit_log)
            elif entry.key == "id_card":
                page = IdCardPage(self.emit_log)
            elif entry.key == "exam_print":
                page = ExamPrintPage(self.emit_log)
            elif entry.key == "job_code_audit":
                page = JobCodeAuditPage(self.emit_log)
            elif entry.key == "ordered_name_audit":
                page = OrderedUnitAuditPage(self.emit_log)
            elif entry.key == "exam":
                page = ExamPage(self.emit_log)
            elif entry.key == "sql_exec":
                page = SqlExecPage(self.emit_log)
            elif entry.key == "project_stage":
                page = ProjectStagePage(self.emit_log)
            elif entry.key == "about":
                page = AboutPage(self.release_config.show_experimental, update_callback=self.run_update_check)
            else:
                page = PlaceholderPage(entry.label, entry.description)
            if hasattr(page, "set_home_callback"):
                page.set_home_callback(lambda target="home": self.open_entry(target))
            if hasattr(page, "set_favorite_callback"):
                page.set_favorite_callback(lambda target=entry.key: self._toggle_favorite_for_key(target))
            if hasattr(page, "set_favorite_state"):
                page.set_favorite_state(entry.key in self.favorites)
            elif entry.key != "home":
                self._attach_page_actions(entry, page)
            self._attach_page_help(entry, page)
            self.page_indexes[entry.key] = self.stack.addWidget(self._wrap_page(page))

    def _create_home_page(self) -> QWidget:
        return HomeLandingPage(
            self.open_entry,
            self.entries_by_key,
            list(self.favorites),
            self.release_config.show_experimental,
        )

    def _replace_home_page(self) -> None:
        home_index = self.page_indexes.get("home")
        if home_index is None:
            return
        old_widget = self.stack.widget(home_index)
        new_widget = self._wrap_page(self._create_home_page())
        self.stack.removeWidget(old_widget)
        old_widget.deleteLater()
        self.stack.insertWidget(home_index, new_widget)
        if self.stack.currentIndex() == home_index:
            self.stack.setCurrentIndex(home_index)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)

    def _ensure_reasonable_window_geometry(self) -> None:
        screen = self.screen() or QApplication.primaryScreen()
        if screen is None:
            return

        available = screen.availableGeometry()
        frame = self.frameGeometry()

        target_width = min(self.width(), max(available.width() - 40, 900))
        target_height = min(self.height(), max(available.height() - 40, 600))

        needs_reset = (
            self.isMaximized()
            or self.isFullScreen()
            or frame.width() > available.width()
            or frame.height() > available.height()
            or frame.right() < available.left()
            or frame.left() > available.right()
            or frame.bottom() < available.top()
            or frame.top() > available.bottom()
        )

        if needs_reset:
            self.showNormal()
            self.resize(target_width, target_height)
            frame = self.frameGeometry()

        x = available.left() + max(0, (available.width() - frame.width()) // 2)
        y = available.top() + max(0, (available.height() - frame.height()) // 2)
        self.move(x, y)

    def _build_navigation_tree(self) -> None:
        self.nav_tree.clear()
        self.tree_items_by_key.clear()
        group_nodes: dict[str, QTreeWidgetItem] = {}
        home_entry = next((entry for entry in self.visible_nav_entries if entry.key == "home"), None)
        if home_entry is not None:
            item = QTreeWidgetItem([self._nav_entry_label(home_entry)])
            item.setData(0, Qt.UserRole, ("entry", home_entry.key))
            self.nav_tree.addTopLevelItem(item)
            self.tree_items_by_key[home_entry.key] = item

        for group_key, group_label in self.visible_nav_groups:
            if group_key == "start":
                continue
            node = QTreeWidgetItem([group_label])
            node.setData(0, Qt.UserRole, ("group", group_key))
            node.setExpanded(True)
            node.setFlags(Qt.ItemIsEnabled)
            node.setForeground(0, QBrush(QColor("#5B6B7E")))
            group_nodes[group_key] = node
            self.nav_tree.addTopLevelItem(node)

        for entry in self.visible_nav_entries:
            if entry.key == "home" or entry.group == "start":
                continue
            item = QTreeWidgetItem([self._nav_entry_label(entry)])
            item.setData(0, Qt.UserRole, ("entry", entry.key))
            if not entry.migrated:
                item.setForeground(0, Qt.gray)
            parent = group_nodes.get(entry.group)
            if parent is not None:
                parent.addChild(item)
            else:
                self.nav_tree.addTopLevelItem(item)
            self.tree_items_by_key[entry.key] = item

    def _nav_entry_label(self, entry: NavEntry) -> str:
        icons = {
            "home": "⌂",
            "certificate": "证",
            "template": "表",
            "photo": "图",
            "pack": "包",
            "match": "数",
            "phone": "电",
            "update_sql": "库",
            "id_card": "查",
            "about": "设",
            "exam_print": "印",
            "job_code_audit": "核",
            "ordered_name_audit": "序",
            "exam": "考",
            "sql_exec": "执",
            "project_stage": "项",
        }
        return f"{icons.get(entry.key, '•')}  {entry.label}"

    def _build_menu(self) -> None:
        menu = self.menuBar().addMenu("工具")
        check_update = QAction("检查更新", self)
        check_update.triggered.connect(self.run_update_check)
        menu.addAction(check_update)

        open_project = QAction("打开项目目录", self)
        open_project.triggered.connect(
            lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path.cwd())))
        )
        menu.addAction(open_project)

        about_action = QAction("关于", self)
        about_action.triggered.connect(lambda: self.open_entry("about"))
        menu.addAction(about_action)

    def run_update_check(self) -> None:
        if sys.platform != "win32":
            QMessageBox.information(self, "检查更新", "增量更新目前只支持 Windows。")
            return
        try:
            result = check_for_updates()
        except UpdateError as exc:
            QMessageBox.warning(self, "检查更新失败", str(exc))
            return

        if not result.has_update or result.package is None:
            QMessageBox.information(
                self,
                "检查更新",
                f"当前已是最新版本：{result.current_version}",
            )
            return

        package = result.package
        package_label = "增量包" if package.package_type == "patch" else "整包"
        message = (
            f"发现新版本：{result.latest_version}\n"
            f"当前版本：{result.current_version}\n"
            f"更新方式：{package_label}\n\n"
            "是否立即下载并应用更新？"
        )
        answer = QMessageBox.question(self, "发现新版本", message)
        if answer != QMessageBox.Yes:
            return

        progress = QProgressDialog("正在下载更新包...", "取消", 0, 100, self)
        progress.setWindowTitle("在线更新")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        progress.setAutoClose(False)
        progress.canceled.connect(lambda: None)

        download_result: dict = {"package_path": None, "error": None}

        def download_thread() -> None:
            try:
                def on_progress(downloaded: int, total: int) -> None:
                    pct = int(downloaded * 100 / total)
                    progress.setLabelText(f"正在下载 {package_label}... {pct}% ({downloaded // 1024}KB / {total // 1024}KB)")
                    progress.setValue(pct)

                package_path = download_update_package(package, progress_callback=on_progress)
                download_result["package_path"] = package_path
            except UpdateError as exc:
                download_result["error"] = str(exc)

        thread = QThread()
        thread.run = download_thread
        thread.finished.connect(lambda: self._on_download_finished(download_result, progress))
        thread.start()

    def _on_download_finished(self, download_result: dict, progress: QProgressDialog) -> None:
        progress.close()
        if download_result["error"]:
            QMessageBox.critical(self, "更新失败", download_result["error"])
            return
        package_path = download_result["package_path"]
        if package_path is None:
            return
        try:
            launch_windows_updater(package_path)
        except UpdateError as exc:
            QMessageBox.critical(self, "更新失败", str(exc))
            return
        QMessageBox.information(self, "开始更新", "更新包已下载，程序退出后会自动替换文件并重新启动。")
        QApplication.instance().quit()

    def _wrap_page(self, page: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        host = QWidget()
        host.setObjectName("ScrollHost")
        host_layout = QVBoxLayout(host)
        host_layout.setContentsMargins(0, 0, 0, 0)
        host_layout.setSpacing(0)
        host_layout.addWidget(page)
        host_layout.addStretch(1)
        scroll.setWidget(host)
        return scroll

    def _position_overlay_buttons(self) -> None:
        return

    def _set_current_tree_item(self, key: str) -> None:
        item = self.tree_items_by_key.get(key)
        if item is None:
            self.nav_tree.clearSelection()
            return
        self.nav_tree.setCurrentItem(item)
        parent = item.parent()
        if parent is not None:
            parent.setExpanded(True)

    def _update_header_state(self, key: str) -> None:
        is_favorite = key in self.favorites
        self.header_favorite_button.setText("已收藏" if is_favorite else "收藏")
        self.header_favorite_button.setEnabled(key not in {"home", "about"})
        self.header_favorite_button.setVisible(key not in {"home", "about"})
        self.header_update_button.setVisible(key in {"home", "about"})
        self.header_settings_button.setVisible(key in {"home", "about"})
        self.header_back_button.setVisible(key != "home")
        self.header_toolbar.hide()

    def open_entry(self, key: str) -> None:
        index = self.page_indexes.get(key)
        if index is None:
            return
        self.stack.setCurrentIndex(index)
        self._set_current_tree_item(key)
        self._update_header_state(key)

    def on_tree_item_clicked(self, item: QTreeWidgetItem) -> None:
        data = item.data(0, Qt.UserRole)
        if not data:
            return
        kind, value = data
        if kind == "group":
            self.nav_tree.clearSelection()
            return
        self.open_entry(value)

    def on_tree_item_pressed(self, item: QTreeWidgetItem, _column: int) -> None:
        data = item.data(0, Qt.UserRole)
        if not data:
            return
        kind, _value = data
        if kind == "group":
            self.nav_tree.clearSelection()

    def toggle_current_favorite(self) -> None:
        current_index = self.stack.currentIndex()
        current_key = next((key for key, index in self.page_indexes.items() if index == current_index), None)
        if current_key is None or current_key in {"home", "about"}:
            return
        self._toggle_favorite_for_key(current_key)

    def _toggle_favorite_for_key(self, key: str) -> None:
        if key in {"home", "about"}:
            return
        if key in self.favorites:
            self.favorites.remove(key)
        else:
            self.favorites.insert(0, key)
        _save_favorites(self.favorites)
        self._replace_home_page()
        self._update_header_state(key)
        self._sync_page_favorite_state(key)

    def _sync_page_favorite_state(self, key: str) -> None:
        favorite_button = self.page_action_buttons.get(key)
        if favorite_button is not None:
            is_favorite = key in self.favorites
            favorite_button.setChecked(is_favorite)
            favorite_button.setText("已收藏" if is_favorite else "收藏")
        index = self.page_indexes.get(key)
        if index is None:
            return
        scroll = self.stack.widget(index)
        if scroll is None:
            return
        host = scroll.widget()
        if host is None or host.layout() is None or host.layout().count() == 0:
            return
        page = host.layout().itemAt(0).widget()
        if page is not None and hasattr(page, "set_favorite_state"):
            page.set_favorite_state(key in self.favorites)

    def _attach_page_actions(self, entry: NavEntry, page: QWidget) -> None:
        layout = page.layout()
        if layout is None or layout.count() == 0:
            return
        hero = layout.itemAt(0).widget()
        if hero is None or hero.layout() is None:
            return

        hero_layout = hero.layout()
        action_row = QWidget()
        action_layout = QHBoxLayout(action_row)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(10)
        action_layout.addStretch(1)

        home_button = QPushButton("返回首页")
        home_button.setObjectName("HelpButton")
        home_button.clicked.connect(lambda: self.open_entry("home"))
        action_layout.addWidget(home_button)

        guide_button = QPushButton("使用指南")
        guide_button.setObjectName("HelpButton")
        guide_button.clicked.connect(lambda _=False, target=entry.key: self._toggle_help_for_key(target))
        action_layout.addWidget(guide_button)

        history_button = QPushButton("历史记录")
        history_button.setObjectName("HelpButton")
        action_layout.addWidget(history_button)

        favorite_button = QPushButton("收藏")
        favorite_button.setObjectName("HelpButton")
        favorite_button.setCheckable(True)
        favorite_button.clicked.connect(lambda: self._toggle_favorite_for_key(entry.key))
        action_layout.addWidget(favorite_button)
        self.page_action_buttons[entry.key] = favorite_button
        self._sync_page_favorite_state(entry.key)

        hero_layout.insertWidget(0, action_row)

    def _attach_page_help(self, entry: NavEntry, page: QWidget) -> None:
        layout = page.layout()
        if layout is None:
            return
        hero_title = None
        hero = None
        if layout.count() > 0:
            hero = layout.itemAt(0).widget()
        if hero is not None and hero.layout() is not None:
            for index in range(hero.layout().count()):
                widget = hero.layout().itemAt(index).widget()
                if isinstance(widget, QLabel) and bool(widget.property("heroTitle")):
                    hero_title = widget
                    break
        banner = QFrame()
        banner.setProperty("pageCard", True)
        banner.hide()
        banner_layout = QVBoxLayout(banner)
        banner_layout.setContentsMargins(20, 16, 20, 16)
        banner_layout.setSpacing(8)
        title = QLabel(f"{entry.label} 操作手册")
        title.setProperty("sectionTitle", True)
        text = QLabel(MANUAL_TEXTS.get(entry.key, entry.description))
        text.setWordWrap(True)
        text.setProperty("heroText", True)
        banner_layout.addWidget(title)
        banner_layout.addWidget(text)
        layout.insertWidget(1, banner)
        self.page_help_sections[entry.key] = banner
        if hero_title is not None:
            hero_title.setCursor(Qt.PointingHandCursor)
            hero_title.setToolTip("点击查看操作手册")
            key = entry.key
            hero_title.mousePressEvent = lambda _event, target=key: self._toggle_help_for_key(target)

    def _toggle_help_for_key(self, key: str) -> None:
        section = self.page_help_sections.get(key)
        if section is None:
            return
        section.setVisible(not section.isVisible())

    def collapse_navigation(self) -> None:
        for index in range(self.nav_tree.topLevelItemCount()):
            item = self.nav_tree.topLevelItem(index)
            data = item.data(0, Qt.UserRole)
            if data and data[1] in {"start", "files", "db"}:
                item.setExpanded(True)
            else:
                item.setExpanded(False)

    def toggle_sidebar(self) -> None:
        self.sidebar_collapsed = not self.sidebar_collapsed
        if self.sidebar_collapsed:
            self.sidebar_brand_icon.hide()
            self.sidebar_app_title.hide()
            self.sidebar_app_subtitle.hide()
            self.search_edit.hide()
            self.sidebar_body.hide()
            self.sidebar_footer_button.hide()
            self.sidebar_toggle_button.hide()
            self.sidebar_overlay_button.show()
            self.sidebar_toggle_button.setProperty("collapsed", True)
            self.sidebar_toggle_button.set_kind("menu-collapsed")
            self.sidebar_overlay_button.set_kind("menu-collapsed")
            self.sidebar.setProperty("collapsed", True)
            self.style().unpolish(self.sidebar_toggle_button)
            self.style().polish(self.sidebar_toggle_button)
            self.style().unpolish(self.sidebar_overlay_button)
            self.style().polish(self.sidebar_overlay_button)
            self.style().unpolish(self.sidebar)
            self.style().polish(self.sidebar)
            self.sidebar_layout.setContentsMargins(0, 12, 0, 12)
            self.sidebar.setMinimumWidth(30)
            self.sidebar.setMaximumWidth(30)
            self.main_splitter.setSizes([30, max(800, self.width() - 30)])
        else:
            self.sidebar_brand_icon.show()
            self.sidebar_app_title.show()
            self.sidebar_app_subtitle.show()
            self.search_edit.show()
            self.sidebar_body.show()
            self.sidebar_footer_button.show()
            self.sidebar_overlay_button.hide()
            self.sidebar_toggle_button.show()
            self.sidebar_toggle_button.setProperty("collapsed", False)
            self.sidebar_toggle_button.set_kind("menu-open")
            self.sidebar.setProperty("collapsed", False)
            self.style().unpolish(self.sidebar_toggle_button)
            self.style().polish(self.sidebar_toggle_button)
            self.style().unpolish(self.sidebar)
            self.style().polish(self.sidebar)
            self.sidebar_layout.setContentsMargins(16, 16, 16, 16)
            self.sidebar.setMinimumWidth(230)
            self.sidebar.setMaximumWidth(280)
            self.main_splitter.setSizes([250, max(760, self.width() - 250)])
        self.sidebar.layout().invalidate()
        self.sidebar.layout().activate()
        self._position_overlay_buttons()

    def filter_navigation(self, text: str) -> None:
        keyword = text.strip().lower()
        for group_index in range(self.nav_tree.topLevelItemCount()):
            group_item = self.nav_tree.topLevelItem(group_index)
            visible_children = 0
            for child_index in range(group_item.childCount()):
                child = group_item.child(child_index)
                data = child.data(0, Qt.UserRole)
                if not data:
                    continue
                _, key = data
                entry = self.entries_by_key[key]
                matched = not keyword or keyword in entry.label.lower() or keyword in entry.description.lower()
                child.setHidden(not matched)
                if matched:
                    visible_children += 1
            group_item.setHidden(visible_children == 0)
            if keyword and visible_children:
                group_item.setExpanded(True)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #EEF3F8; }
            QWidget {
                color: #243247;
                selection-background-color: #DCE8FF;
                selection-color: #173052;
            }
            QMessageBox {
                background: #FFFFFF;
            }
            QMessageBox QLabel {
                color: #172033;
                background: transparent;
                font-size: 13px;
            }
            QMessageBox QPushButton {
                min-width: 88px;
                min-height: 34px;
                padding: 4px 14px;
                color: #243247;
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                font-weight: 600;
            }
            QMessageBox QPushButton:hover {
                background: #F7FAFD;
            }
            #Sidebar, #Card {
                background: #FFFFFF;
                border: 1px solid #D9E2EC;
                border-radius: 14px;
            }
            #Sidebar[collapsed="true"] {
                background: transparent;
                border: none;
            }
            QScrollArea, QScrollArea > QWidget > QWidget, #ScrollHost {
                background: transparent;
                border: none;
            }
            QSplitter::handle {
                background: transparent;
            }
            QLabel {
                color: #243247;
                background: transparent;
            }
            QLabel[appTitle="true"] {
                font-size: 22px;
                font-weight: 700;
                color: #172033;
            }
            QLabel[appSubtitle="true"] {
                color: #607086;
                font-size: 12px;
            }
            QLabel#SidebarBrandIcon {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2F73FF, stop:1 #6AD0FF);
                color: #FFFFFF;
                border-radius: 16px;
                font-size: 24px;
                font-weight: 800;
            }
            QLabel[formLabel="true"] {
                color: #314155;
                font-size: 12px;
                font-weight: 600;
                min-width: 98px;
            }
            QLabel[sectionTitle="true"] {
                color: #172033;
                font-size: 14px;
                font-weight: 700;
            }
            QLabel#SectionTitle {
                color: #172033;
                font-size: 14px;
                font-weight: 700;
                border-left: 4px solid #3E7BFA;
                padding-left: 8px;
            }
            QLabel[heroTitle="true"] {
                font-size: 19px;
                font-weight: 700;
                color: #172033;
            }
            QLabel[heroText="true"] {
                font-size: 11px;
                color: #5B6B7E;
            }
            QPushButton#HelpButton {
                min-width: 72px;
                min-height: 30px;
                padding: 3px 12px;
                color: #314155;
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid #DDE6F2;
                border-radius: 14px;
                font-weight: 600;
            }
            QPushButton#HelpButton:hover {
                background: #F7FAFD;
                border-color: #C7D8EF;
            }
            QPushButton#HelpButton:checked {
                background: #F5B301;
                border-color: #F5B301;
                color: #FFFFFF;
            }
            #HeaderToolbar {
                background: transparent;
                min-height: 36px;
            }
            QFrame[pageCard="true"] {
                background: #FFFFFF;
                border: 1px solid #D9E2EC;
                border-radius: 14px;
            }
            #ExamPrintTopbar {
                border-radius: 10px;
            }
            QLabel#StepItem {
                background: #F4F8FF;
                border: 1px solid #D9E8FF;
                border-radius: 16px;
                padding: 7px 12px;
                color: #2F73FF;
                font-weight: 700;
            }
            #ExamPrintLeftRail, #ExamPrintRightPanel {
                background: #FFFFFF;
                border: 1px solid #D9E2EC;
                border-radius: 10px;
            }
            QFrame[workflowStep="true"] {
                background: transparent;
                border: 1px solid transparent;
                border-radius: 8px;
            }
            QFrame[workflowStep="true"][active="true"] {
                background: #2F80FF;
                border-color: #2F80FF;
            }
            QFrame[workflowStep="true"][active="true"] QLabel {
                color: #FFFFFF;
            }
            QLabel[stepBadge="true"] {
                background: #EEF5FF;
                color: #2F80FF;
                border: 1px solid #CFE0FF;
                border-radius: 14px;
                font-weight: 800;
            }
            QFrame[workflowStep="true"][active="true"] QLabel[stepBadge="true"] {
                background: #FFFFFF;
                color: #2F80FF;
                border-color: #FFFFFF;
            }
            QLineEdit, QComboBox, QDateEdit, QSpinBox {
                color: #172033;
                min-height: 30px;
                padding: 3px 8px;
                border: 1px solid #CBD5E1;
                border-radius: 9px;
                background: #FFFFFF;
            }
            QComboBox {
                padding-right: 30px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 28px;
                border: none;
                background: transparent;
            }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus,
            QPlainTextEdit:focus, QTextEdit:focus {
                border: 1px solid #77A5FF;
            }
            QComboBox QAbstractItemView {
                color: #172033;
                background: #FFFFFF;
                selection-background-color: #DCE8FF;
                selection-color: #173052;
                border: 1px solid #CBD5E1;
                border-radius: 9px;
                padding: 4px;
                outline: none;
            }
            QComboBox QAbstractItemView::item {
                min-height: 28px;
                padding: 3px 8px;
                border-radius: 8px;
            }
            QComboBox QAbstractItemView::item:selected {
                background: #DCE8FF;
                color: #173052;
            }
            QPlainTextEdit, QTextEdit {
                color: #172033;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                background: #FFFFFF;
                selection-background-color: #DCE8FF;
                selection-color: #173052;
            }
            QListWidget, QListView {
                color: #172033;
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                outline: none;
                selection-background-color: #DCE8FF;
                selection-color: #173052;
            }
            QListWidget::item, QListView::item {
                min-height: 24px;
                padding: 4px 8px;
                border-radius: 7px;
            }
            QListWidget::item:selected, QListView::item:selected {
                background: #DCE8FF;
                color: #173052;
            }
            QListWidget::item:disabled, QListView::item:disabled {
                color: #607086;
                background: #FFFFFF;
            }
            QTableWidget {
                color: #172033;
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                gridline-color: #E6EDF5;
                selection-background-color: #DCE8FF;
                selection-color: #173052;
            }
            QTableWidget::item {
                padding: 6px 8px;
                border: none;
            }
            QTableWidget::item:selected {
                background: #DCE8FF;
                color: #173052;
            }
            QHeaderView::section {
                color: #314155;
                background: #F7FAFD;
                border: none;
                border-bottom: 1px solid #D9E2EC;
                padding: 8px 10px;
                font-weight: 700;
            }
            QCheckBox, QRadioButton {
                color: #243247;
                spacing: 8px;
                font-size: 13px;
            }
            QCheckBox::indicator, QRadioButton::indicator {
                width: 18px;
                height: 18px;
            }
            QCheckBox::indicator {
                border: 1px solid #9FB4CC;
                border-radius: 5px;
                background: #FFFFFF;
            }
            QCheckBox::indicator:checked {
                border-color: #3E7BFA;
                background: #3E7BFA;
            }
            QRadioButton::indicator {
                border: 1px solid #9FB4CC;
                border-radius: 9px;
                background: #FFFFFF;
            }
            QRadioButton::indicator:checked {
                border-color: #3E7BFA;
                background: #3E7BFA;
            }
            QPushButton {
                min-height: 32px;
                padding: 0 12px;
                border: 1px solid #CBD5E1;
                border-radius: 9px;
                background: #FFFFFF;
                color: #1F3147;
                font-weight: 600;
            }
            QPushButton[accent="true"] {
                background: #3E7BFA;
                color: #FFFFFF;
                border-color: #3E7BFA;
                font-weight: 700;
            }
            QPushButton[toolbarButton="true"] {
                min-height: 30px;
                padding: 0 10px;
                border: 1px solid transparent;
                border-radius: 8px;
                background: transparent;
                color: #1F3147;
                font-weight: 600;
            }
            QPushButton[toolbarButton="true"]:hover {
                background: #EEF5FF;
                border-color: #CFE0FF;
                color: #1465E8;
            }
            #StarButton {
                min-height: 28px;
                min-width: 28px;
                max-height: 28px;
                max-width: 28px;
                padding: 0;
                border: none;
                background: transparent;
            }
            #MiniButton {
                min-height: 30px;
                min-width: 30px;
                max-height: 30px;
                max-width: 30px;
                padding: 0;
                border-radius: 10px;
                font-size: 16px;
            }
            #MiniButton[collapsed="true"] {
                min-height: 30px;
                min-width: 30px;
                max-height: 30px;
                max-width: 30px;
                border-radius: 10px;
                font-size: 16px;
            }
            #HomeHero {
                border-radius: 18px;
            }
            QLabel#HomeWelcomeTitle {
                font-size: 22px;
                font-weight: 800;
                color: #0F1F3A;
            }
            #HomeSectionCard {
                border-radius: 18px;
            }
            QLabel#HomeSectionMarker {
                background: #2F73FF;
                border-radius: 3px;
            }
            QLabel#HomeSectionTitle {
                font-size: 16px;
                font-weight: 800;
                color: #13223A;
            }
            #HomeStatCard {
                background: #FFFFFF;
                border: 1px solid #E0E8F2;
                border-radius: 16px;
            }
            QLabel#HomeStatIcon {
                background: #EEF5FF;
                border: 1px solid #D9E8FF;
                border-radius: 12px;
                font-size: 17px;
                font-weight: 800;
            }
            QLabel#HomeStatValue {
                font-size: 18px;
                font-weight: 800;
            }
            QLabel#HomeStatLabel {
                color: #506178;
                font-size: 12px;
            }
            #HomeFeatureCard {
                background: #FFFFFF;
                border: 1px solid #DDE7F3;
                border-radius: 16px;
            }
            #HomeFeatureCard[hovered="true"] {
                border-color: #9DBEFF;
                background: #FBFDFF;
            }
            #HomeFeatureCard[pressed="true"] {
                background: #F4F8FF;
                border-color: #7EA8F8;
            }
            QLabel#HomeFeatureIcon {
                border-radius: 12px;
                font-size: 15px;
                font-weight: 800;
            }
            QLabel#HomeFeatureIconSmall {
                border-radius: 11px;
                font-size: 13px;
                font-weight: 800;
            }
            QLabel#HomeFeatureTitle {
                color: #13223A;
                font-size: 14px;
                font-weight: 800;
            }
            QLabel#HomeFavoriteStar {
                color: #F5A400;
                font-size: 16px;
                font-weight: 900;
            }
            QPushButton#HomeOpenButton {
                min-height: 32px;
                color: #2F73FF;
                background: #FFFFFF;
                border: 1px solid #D9E2EC;
                border-radius: 10px;
            }
            QPushButton#HomeOpenButton:hover {
                background: #F4F8FF;
                border-color: #9DBEFF;
            }
            QPushButton#HomeOpenButton:pressed {
                background: #EAF2FF;
                border-color: #7EA8F8;
            }
            QPushButton#HomeFilterButton {
                min-height: 30px;
                padding: 0 14px;
                border-radius: 8px;
                background: #FFFFFF;
                border: 1px solid #D9E2EC;
                color: #506178;
                font-weight: 700;
            }
            QPushButton#HomeFilterButton:hover {
                background: #F8FBFF;
                border-color: #C9D9ED;
            }
            QPushButton#HomeFilterButton:checked {
                background: #2F73FF;
                border-color: #2F73FF;
                color: #FFFFFF;
            }
            QPushButton#HomeViewToggle {
                background: #FFFFFF;
                border: 1px solid #D9E2EC;
                border-radius: 10px;
                color: #7990AA;
                font-size: 15px;
                font-weight: 700;
            }
            QPushButton#HomeViewToggle:hover {
                border-color: #C9D9ED;
                background: #F8FBFF;
            }
            QPushButton#HomeViewToggle:checked {
                background: #EEF5FF;
                border-color: #CFE0FF;
                color: #2F73FF;
            }
            #HomeFeatureRow {
                background: #FFFFFF;
                border: 1px solid #DDE7F3;
                border-radius: 14px;
            }
            #HomeFeatureRow[hovered="true"] {
                background: #FBFDFF;
                border-color: #9DBEFF;
            }
            #HomeFeatureRow[pressed="true"] {
                background: #F4F8FF;
                border-color: #7EA8F8;
            }
            QLabel#HomeRowArrow {
                color: #8DA2BB;
                font-size: 24px;
                font-weight: 300;
            }
            #QuickList {
                border: 1px solid #D9E2EC;
                border-radius: 12px;
                background: #F8FBFF;
                outline: none;
            }
            #QuickList::item {
                margin: 4px;
                padding: 10px 12px;
                border-radius: 10px;
            }
            #QuickList::item:selected {
                background: #DCE8FF;
                color: #173052;
                font-weight: 700;
            }
            #NavTree {
                border: none;
                background: transparent;
                padding: 0;
                outline: none;
                show-decoration-selected: 0;
            }
            #NavTree::item {
                min-height: 34px;
                padding: 4px 10px 4px 12px;
                margin: 2px 0 2px 10px;
                border-radius: 10px;
            }
            #NavTree::item:selected {
                background: #2F73FF;
                color: #FFFFFF;
                font-weight: 700;
            }
            #NavTree::branch:selected {
                background: transparent;
            }
            #NavTree::branch:has-siblings:!adjoins-item,
            #NavTree::branch:has-siblings:adjoins-item,
            #NavTree::branch:closed:has-children,
            #NavTree::branch:open:has-children {
                border-image: none;
                image: none;
            }
            QScrollBar:vertical {
                background: transparent;
                width: 10px;
                margin: 6px 2px 6px 2px;
            }
            QScrollBar::handle:vertical {
                background: #C8D6E8;
                min-height: 48px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: #AFC2DC;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
                background: transparent;
                border: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: transparent;
            }
            QPushButton#SidebarFooterButton {
                min-height: 44px;
                padding: 0 16px;
                text-align: left;
                color: #4E617B;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #F5F8FC, stop:1 #EEF3FA);
                border: 1px solid #DCE5F0;
                border-radius: 14px;
                font-weight: 700;
            }
            QPushButton#SidebarFooterButton:hover {
                border-color: #C9D8EB;
                color: #1B4ECC;
            }
            """
        )

    def emit_log(self, message: str) -> None:
        self.log_bridge.message.emit(message)

    def append_log(self, message: str) -> None:
        pass


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = QtMainWindow()
    window.show()
    sys.exit(app.exec())
