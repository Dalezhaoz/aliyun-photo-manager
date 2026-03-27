from __future__ import annotations

import math
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QUrl, Signal
from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
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
    ExamPage,
    IdCardPage,
    MatchPage,
    PackPage,
    PhoneDecryptPage,
    PhotoPage,
    ProjectStagePage,
    SqlExecPage,
    TemplateConvertPage,
    UpdateSqlPage,
)


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
    NavEntry("home", "首页", "start", "查看收藏入口、最近使用和当前迁移范围。", True),
    NavEntry("certificate", "证件资料筛选", "files", "按模板筛选资料目录，支持本地目录主流程。", True),
    NavEntry("template", "表样转换", "files", "将 Word / Excel 表样转换成 HTML。", True),
    NavEntry("photo", "照片下载与分类", "files", "支持本地目录和云存储下载后生成模板、按模板分类。", True),
    NavEntry("pack", "结果打包", "files", "对任意结果文件或文件夹压缩并 AES 加密。", True),
    NavEntry("match", "数据匹配", "data", "按主键和附加匹配列补充来源表字段。", True),
    NavEntry("phone", "电话解密", "db", "通过 helper 解密电话并回写备用3。", True),
    NavEntry("update_sql", "更新 SQL 生成", "db", "通过字段映射模板生成标准 UPDATE SQL。", True),
    NavEntry("id_card", "身份证工具", "query", "校验并生成 18 位大陆居民身份证。", True),
    NavEntry("about", "关于", "settings", "查看 Qt 预览版说明。", True),
    NavEntry("exam", "考场编排", "experimental", "实验功能：按模板和规则生成考号、考场与座号。", True),
    NavEntry("sql_exec", "SQL 配置执行", "experimental", "实验功能：按 SQL 模板参数生成可执行脚本。", True),
    NavEntry("project_stage", "项目阶段汇总", "experimental", "实验功能：汇总多台 SQL Server 上的报名项目阶段状态。", True),
]


DEFAULT_FAVORITES = ["certificate", "template", "phone"]

MANUAL_TEXTS: dict[str, str] = {
    "photo": "适用场景：下载考生照片、生成分类模板、按模板批量分类。\n\n操作步骤：\n1. 先选择数据来源，本地目录可直接处理，云存储需先填写云类型、Endpoint/Region、AccessKey 和 Bucket。\n2. 如需按名单处理，先选择人员模板并勾选只处理模板中的人员。\n3. 选择下载目录或已有照片目录，再选择分类输出目录。\n4. 需要从云端下载时，先加载 Bucket 和当前层级，再执行下载。\n5. 点击生成模板后，在 Excel 里填写分类一、分类二、分类三、修改名称等字段。\n6. 回到程序执行按模板分类，结果会输出分类目录和结果清单。",
    "certificate": "适用场景：从本地或云存储筛选证件资料，只导出指定人员或指定材料。\n\n操作步骤：\n1. 先选择数据来源，本地目录可直接筛选，云存储模式会先下载后筛选。\n2. 选择人员模板并加载模板列，确认匹配列、名称列等关键字段。\n3. 选择证件资料目录和输出目录。\n4. 如需只导出某类材料，切换筛选模式为关键词文件，并填写关键词。\n5. 如需重命名导出目录，可勾选导出后文件夹重命名并选择名称列。\n6. 运行后查看右侧结果和结果清单，确认导出人数、文件数和输出位置。",
    "template": "适用场景：把 Word、Excel 表样模板转换成 HTML 代码。\n\n操作步骤：\n1. 选择表样文件，支持 doc、docx、xls、xlsx。\n2. 根据模板类型点击 Net 版导出或 Java 版导出。\n3. 程序会在结果区展示 HTML 和占位符内容。\n4. 可直接复制 HTML，后续粘贴到系统模板或页面中使用。\n5. 若模板异常，优先检查原始表样中的合并单元格、图片和特殊格式。",
    "match": "适用场景：用来源表字段补充目标表，替代手工 VLOOKUP / XLOOKUP。\n\n操作步骤：\n1. 选择目标表、来源表和输出文件。\n2. 点击加载表头，确认目标表匹配列与来源表匹配列。\n3. 如存在重复姓名或重复主键，可在附加匹配列里继续增加限制条件。\n4. 在补充列映射里填写结果列名，并选择来源表对应列。\n5. 点击开始匹配，程序会输出匹配结果文件和匹配清单。\n6. 若结果不符合预期，优先检查主键列格式、空值和附加匹配列是否一致。",
    "pack": "适用场景：把照片、证件资料或其他结果目录打包并加密交付。\n\n操作步骤：\n1. 选择待打包对象，可选文件或文件夹。\n2. 选择输出目录。\n3. 如需客户指定密码，勾选手动设置密码并输入密码；否则程序自动生成密码。\n4. 点击一键打包并加密，成功后右侧会显示压缩包路径、密码和历史记录。\n5. 复制密码后发给对方，后续可通过查询历史按文件名、来源名或密码回查。",
    "phone": "适用场景：按主键编号关联 web_info，加密电话解密后回写到考生表备用3。\n\n操作步骤：\n1. 填写服务器、数据库账号、报名库名和电话库名。\n2. 选择考生表和需要处理的模式，可全量处理，也可按名单处理。\n3. 程序会根据主键编号、考试代码、考试年月和考区参数调用 helper 解密。\n4. 解密成功后写回备用3，并在结果区显示成功、失败和跳过数量。\n5. 若失败，优先检查 helper 是否已构建、数据库连接是否正常、电话库中是否存在对应密文。",
    "update_sql": "适用场景：根据字段映射模板生成标准 UPDATE SQL，并在执行前自动备份相关表。\n\n操作步骤：\n1. 先准备字段映射模板，或点击导出模板生成标准样例。\n2. 选择映射模板并加载字段。\n3. 填写正式表、临时表以及双方关联字段。\n4. 如需防止空值覆盖正式表，可勾选忽略空值。\n5. 点击生成 SQL，在右侧检查备份语句、更新语句和 where 条件是否正确。\n6. 确认无误后复制 SQL，到数据库工具中执行。",
    "id_card": "适用场景：校验身份证号，或按地区、出生日期、性别生成测试数据。\n\n操作步骤：\n1. 输入校验区可直接输入 18 位身份证号，点击校验并解析查看出生日期、性别和地区。\n2. 生成区先选择省、市、县，也可手工填写 6 位区划码。\n3. 选择出生日期和性别后点击生成身份证。\n4. 程序会一次生成 10 个合法号码，复制结果默认复制第 1 个。\n5. 若用于测试，请不要把生成号码当作真实身份信息使用。",
    "exam": "适用场景：实验功能，用于按规则编排考号、考场和座号。\n\n操作步骤：\n1. 准备考生名单与编排规则模板。\n2. 按页面提示加载规则并设置考场容量、排序字段等参数。\n3. 先用小样本验证规则，再执行完整编排。\n4. 输出结果后重点检查考号连续性、考场容量和特殊考生分配是否正确。",
    "sql_exec": "适用场景：实验功能，用配置模板批量生成 SQL 语句。\n\n操作步骤：\n1. 选择 SQL 模板和参数文件。\n2. 加载后确认变量名、替换值和输出格式。\n3. 先在测试环境预览生成结果，再复制或导出执行。\n4. 如模板中包含删除、更新语句，请务必先备份再执行。",
    "project_stage": "适用场景：实验功能，汇总多台 SQL Server 上的报名项目阶段状态。\n\n操作步骤：\n1. 填写服务器连接信息，先执行测试连接。\n2. 配置需要查询的项目库、阶段表或汇总规则。\n3. 运行后查看结果区输出，确认每台服务器的阶段状态和统计值。\n4. 如查询慢或失败，优先检查网络、SQL Server 权限和超时设置。",
    "about": "当前是 Qt 预览版，重点在验证新导航、新布局和功能迁移稳定性。\n\n建议优先测试高频页面：证件资料筛选、电话解密、更新 SQL 生成、数据匹配、结果打包和身份证工具。\n若发现样式、布局或交互问题，可直接按页面逐项反馈。"
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
        body.setPlainText("这个功能在 Qt 版主框架里已经预留入口，后续会按优先级逐步迁移。")
        info_layout.addWidget(body)
        root.addWidget(info)


class HomePage(QWidget):
    def __init__(self, open_callback) -> None:
        super().__init__()
        self.open_callback = open_callback
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
        intro = QLabel("Qt 预览版先迁正式功能，实验功能继续保留占位。左侧支持分组、搜索和常用功能。")
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

        section = QLabel("当前已迁移页面")
        section.setProperty("sectionTitle", True)
        migrated_layout.addWidget(section)

        for key in ("photo", "certificate", "template", "match", "pack", "phone", "update_sql", "id_card", "exam", "sql_exec", "project_stage"):
            entry = next(item for item in NAV_ENTRIES if item.key == key)
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


class QtMainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"报名系统工具箱 Qt 预览版 v{__version__}")
        self.resize(1440, 920)
        self.setMinimumSize(1180, 760)

        font = QFont("Microsoft YaHei UI", 10)
        QApplication.instance().setFont(font)

        self.log_bridge = LogBridge()
        self.log_bridge.message.connect(self.append_log)

        self.entries_by_key = {entry.key: entry for entry in NAV_ENTRIES}
        self.favorites: list[str] = list(DEFAULT_FAVORITES)
        self.page_indexes: dict[str, int] = {}
        self.tree_items_by_key: dict[str, QTreeWidgetItem] = {}
        self.page_help_sections: dict[str, QFrame] = {}
        self.sidebar_collapsed = False

        central = QWidget()
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(18, 18, 18, 18)
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
        self.sidebar.setMinimumWidth(280)
        self.sidebar.setMaximumWidth(340)
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
        title_block = QVBoxLayout()
        title_block.setSpacing(2)
        self.sidebar_app_title = QLabel("报名系统工具箱")
        self.sidebar_app_title.setProperty("appTitle", True)
        self.sidebar_app_subtitle = QLabel(f"Qt 预览版 v{__version__}")
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

        self.fav_title = QLabel("常用功能")
        self.fav_title.setProperty("sectionTitle", True)
        sidebar_body_layout.addWidget(self.fav_title)
        self.favorite_list = QListWidget()
        self.favorite_list.setObjectName("QuickList")
        self.favorite_list.setFixedHeight(148)
        self.favorite_list.itemClicked.connect(self.open_from_quick_list)
        sidebar_body_layout.addWidget(self.favorite_list)

        nav_title_row = QHBoxLayout()
        self.nav_title = QLabel("全部功能")
        self.nav_title.setProperty("sectionTitle", True)
        nav_title_row.addWidget(self.nav_title)
        nav_title_row.addStretch(1)
        sidebar_body_layout.addLayout(nav_title_row)
        self.nav_tree = QTreeWidget()
        self.nav_tree.setHeaderHidden(True)
        self.nav_tree.setIndentation(14)
        self.nav_tree.setObjectName("NavTree")
        self.nav_tree.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.nav_tree.itemPressed.connect(self.on_tree_item_pressed)
        self.nav_tree.itemClicked.connect(self.on_tree_item_clicked)
        sidebar_body_layout.addWidget(self.nav_tree, 1)
        sidebar_layout.addWidget(self.sidebar_body, 1)

        content_wrapper = QWidget()
        content_layout = QVBoxLayout(content_wrapper)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        self.content_container = QWidget()
        container_layout = QVBoxLayout(self.content_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(12)
        self.header_help_button = QPushButton("说明", self.content_container)
        self.header_help_button.setObjectName("HelpButton")
        self.header_help_button.setCursor(Qt.PointingHandCursor)
        self.header_help_button.setFixedSize(72, 32)
        self.header_help_button.clicked.connect(self.show_current_help)
        self.header_star_button = PaintedIconButton("star", self.content_container)
        self.header_star_button.setObjectName("StarButton")
        self.header_star_button.clicked.connect(self.toggle_current_favorite)
        self.header_star_button.setFixedSize(28, 28)
        self.stack = QStackedWidget()
        container_layout.addWidget(self.stack)
        self.header_help_button.raise_()
        self.header_star_button.raise_()
        content_layout.addWidget(self.content_container, 1)

        splitter.addWidget(self.sidebar)
        splitter.addWidget(content_wrapper)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([300, 1100])
        self.main_splitter = splitter

        self._build_navigation_tree()
        self._build_pages()
        self._refresh_favorites()
        self._build_menu()
        self._apply_styles()

        self.open_entry("home")

    def _build_pages(self) -> None:
        for entry in NAV_ENTRIES:
            if entry.key == "home":
                page = HomePage(self.open_entry)
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
            elif entry.key == "exam":
                page = ExamPage(self.emit_log)
            elif entry.key == "sql_exec":
                page = SqlExecPage(self.emit_log)
            elif entry.key == "project_stage":
                page = ProjectStagePage(self.emit_log)
            elif entry.key == "about":
                page = PlaceholderPage("关于 Qt 预览版", "当前重点是先把新导航和高频页面做稳，再逐步迁移其他业务页。")
            else:
                page = PlaceholderPage(entry.label, entry.description)
            self._attach_page_help(entry, page)
            self.page_indexes[entry.key] = self.stack.addWidget(self._wrap_page(page))

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._position_overlay_buttons()

    def _build_navigation_tree(self) -> None:
        self.nav_tree.clear()
        self.tree_items_by_key.clear()
        group_nodes: dict[str, QTreeWidgetItem] = {}

        for group_key, group_label in NAV_GROUPS:
            node = QTreeWidgetItem([group_label])
            node.setData(0, Qt.UserRole, ("group", group_key))
            node.setExpanded(group_key in {"start", "files", "db"})
            group_nodes[group_key] = node
            self.nav_tree.addTopLevelItem(node)

        for entry in NAV_ENTRIES:
            item = QTreeWidgetItem([entry.label])
            item.setData(0, Qt.UserRole, ("entry", entry.key))
            if not entry.migrated:
                item.setForeground(0, Qt.gray)
            group_nodes[entry.group].addChild(item)
            self.tree_items_by_key[entry.key] = item

    def _build_menu(self) -> None:
        menu = self.menuBar().addMenu("工具")
        open_project = QAction("打开项目目录", self)
        open_project.triggered.connect(
            lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path.cwd())))
        )
        menu.addAction(open_project)

        about_action = QAction("关于", self)
        about_action.triggered.connect(lambda: self.open_entry("about"))
        menu.addAction(about_action)

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
        if hasattr(self, "content_container"):
            x = self.content_container.width() - self.header_star_button.width() - 16
            y = 18
            self.header_star_button.move(max(0, x), max(0, y))
            help_x = x - self.header_help_button.width() - 10
            help_y = y - 2
            self.header_help_button.move(max(0, help_x), max(0, help_y))
        if hasattr(self, "central_panel"):
            self.sidebar_overlay_button.move(18, 28)

    def _refresh_favorites(self) -> None:
        self.favorite_list.clear()
        for key in self.favorites:
            entry = self.entries_by_key.get(key)
            if entry is None:
                continue
            item = QListWidgetItem(entry.label)
            item.setData(Qt.UserRole, key)
            self.favorite_list.addItem(item)

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
        self.header_star_button.setChecked(is_favorite)
        self.header_star_button.setEnabled(key not in {"home", "about"})
        self.header_help_button.setVisible(key not in {"home"})
        self.header_star_button.update()
        section = self.page_help_sections.get(key)
        self.header_help_button.setText("收起" if section is not None and section.isVisible() else "说明")

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
            item.setExpanded(not item.isExpanded())
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

    def open_from_quick_list(self, item: QListWidgetItem) -> None:
        key = item.data(Qt.UserRole)
        if key:
            self.open_entry(key)

    def toggle_current_favorite(self) -> None:
        current_index = self.stack.currentIndex()
        current_key = next((key for key, index in self.page_indexes.items() if index == current_index), None)
        if current_key is None or current_key in {"home", "about"}:
            return
        if current_key in self.favorites:
            self.favorites.remove(current_key)
        else:
            self.favorites.insert(0, current_key)
        self._refresh_favorites()
        self._update_header_state(current_key)

    def show_current_help(self) -> None:
        current_index = self.stack.currentIndex()
        current_key = next((key for key, index in self.page_indexes.items() if index == current_index), None)
        if current_key is None:
            return
        section = self.page_help_sections.get(current_key)
        if section is None:
            return
        section.setVisible(not section.isVisible())
        self.header_help_button.setText("收起" if section.isVisible() else "说明")

    def _attach_page_help(self, entry: NavEntry, page: QWidget) -> None:
        layout = page.layout()
        if layout is None:
            return
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
            self.sidebar_app_title.hide()
            self.sidebar_app_subtitle.hide()
            self.search_edit.hide()
            self.sidebar_title.hide()
            self.fav_title.hide()
            self.nav_title.hide()
            self.sidebar_body.hide()
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
            self.sidebar_app_title.show()
            self.sidebar_app_subtitle.show()
            self.search_edit.show()
            self.sidebar_title.show()
            self.fav_title.show()
            self.nav_title.show()
            self.sidebar_body.show()
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
            self.sidebar.setMinimumWidth(280)
            self.sidebar.setMaximumWidth(340)
            self.main_splitter.setSizes([300, max(780, self.width() - 300)])
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
                font-size: 26px;
                font-weight: 700;
                color: #172033;
            }
            QLabel[appSubtitle="true"] {
                color: #607086;
                font-size: 12px;
            }
            QLabel[formLabel="true"] {
                color: #314155;
                font-size: 13px;
                font-weight: 600;
                min-width: 112px;
            }
            QLabel[sectionTitle="true"] {
                color: #172033;
                font-size: 14px;
                font-weight: 700;
            }
            QLabel[heroTitle="true"] {
                font-size: 22px;
                font-weight: 700;
                color: #172033;
            }
            QLabel[heroText="true"] {
                font-size: 12px;
                color: #5B6B7E;
            }
            QPushButton#HelpButton {
                min-width: 72px;
                min-height: 32px;
                padding: 4px 14px;
                color: #314155;
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                font-weight: 600;
            }
            QPushButton#HelpButton:hover {
                background: #F7FAFD;
            }
            QFrame[pageCard="true"] {
                background: #FFFFFF;
                border: 1px solid #D9E2EC;
                border-radius: 14px;
            }
            QLineEdit, QComboBox, QDateEdit, QSpinBox {
                color: #172033;
                min-height: 36px;
                padding: 4px 10px;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                background: #FFFFFF;
            }
            QComboBox {
                padding-right: 36px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 34px;
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
                border-radius: 10px;
                padding: 6px;
                outline: none;
            }
            QComboBox QAbstractItemView::item {
                min-height: 34px;
                padding: 4px 10px;
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
                min-height: 36px;
                padding: 0 14px;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
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
                border: 1px solid #D9E2EC;
                border-radius: 14px;
                background: #FFFFFF;
                padding: 10px 8px;
                outline: none;
                show-decoration-selected: 0;
            }
            #NavTree::item {
                min-height: 34px;
                padding: 4px 10px 4px 12px;
                margin: 2px 8px 2px 14px;
                border-radius: 10px;
            }
            #NavTree::item:selected {
                background: #DCE8FF;
                color: #173052;
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
