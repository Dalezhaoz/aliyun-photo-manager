# Dev分支概览

## 分支定位

`codex/dev` 是当前开发分支，承担以下目标：

- Qt 界面体系持续迁移
- 首页与导航结构重构
- 实验功能统一纳入新页面框架
- 主线新增功能的预研与整合

## 相比 main 的重点差异

- Qt 预览页覆盖更多功能
- 首页、导航、收藏入口已经成型
- 电话解密能力继续增强
- 身份证号工具提供更完整的地区选择能力
- 新增 SQL 配置执行等实验工具入口

## 当前分支头部

当前看到的 `dev` 新近演进包括：

- `5811796` Add Qt preview pages for all tools
- `d0de299` Redesign main UI shell and home navigation
- `2a8ddae` Add ID card validation and generator tab
- `a6525eb` Add province city county ID region selectors
- `26da714` Use full 2024 region data for ID selectors

## 阅读建议

- 先看 [[Qt页面结构]]
- 再看 [[Dev迭代记录]]
