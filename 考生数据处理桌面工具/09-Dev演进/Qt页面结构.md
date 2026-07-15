# Qt页面结构

## 入口

- `qt_gui.py`
- `app_launcher_qt.py`

## 页面文件

- `qt_ui/photo_page.py`
- `qt_ui/certificate_page.py`
- `qt_ui/template_page.py`
- `qt_ui/match_page.py`
- `qt_ui/pack_page.py`
- `qt_ui/phone_page.py`
- `qt_ui/id_card_page.py`
- `qt_ui/update_sql_page.py`
- `qt_ui/sql_exec_page.py`
- `qt_ui/project_stage_page.py`
- `qt_ui/exam_page.py`
- `qt_ui/manual_page.py`
- `qt_ui/common.py`

## 导航分组

根据 `qt_gui.py` 当前设计，导航大致分为：

- 开始使用
- 文件处理
- 数据处理
- 数据库工具
- 查询与辅助
- 设置
- 实验功能

## 页面状态

- 已迁移页面
  证件资料筛选、表样转换、照片、结果打包、数据匹配、电话解密、更新 SQL、身份证工具等
- 实验功能页
  考场编排、SQL 配置执行、项目阶段汇总
- 说明页
  首页、关于、手册说明

## 设计特点

- 统一导航壳层
- 首页可承载收藏入口
- 多页统一视觉风格
- 适合逐步替换 Tk 页面
