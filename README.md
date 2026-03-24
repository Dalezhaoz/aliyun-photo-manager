# 报名系统工具箱

当前版本：`v1.3.0`

这是一个桌面工具，主要用于：

- 考生照片下载与分类
- 证件资料下载与筛选
- 表样转换
- 项目阶段汇总
- 项目阶段汇总 Web 服务
- 数据匹配
- 考场编排
- 结果打包
- 更新 SQL 生成

支持两种云存储：

- 阿里云 OSS
- 腾讯云 COS

也支持直接处理本地目录。

## 主要功能

### 1. 照片下载与分类

- 支持从云存储下载考生照片
- 支持直接处理本地照片目录
- 支持只下载模板中的人员照片
- 支持按 Excel 模板中的“分类一 / 分类二 / 分类三 / 修改名称”分类
- 支持仅预览，不实际执行
- 支持跳过已存在文件
- 支持生成“照片分类结果清单.xlsx”

### 2. 证件资料筛选

- 支持从云存储下载证件资料
- 支持直接处理本地证件资料目录
- 支持只下载模板中的人员资料
- 支持按“身份证号”或“报名序号”等任意模板列匹配
- 支持导出后文件夹重命名
- 支持复制整个人员文件夹
- 支持只复制关键词文件，例如“学历证书”
- 支持按分类一 / 分类二 / 分类三建立目录
- 支持仅预览，不实际执行
- 支持生成“证件资料筛选结果清单.xlsx”

### 3. 表样转换

- 支持 `.doc` / `.docx` / `.xlsx`
- 支持 Net 版导出
- 支持 Java 版导出
- 支持代码查看
- 支持复制 HTML
- 支持浏览器预览

### 4. 数据匹配

- 支持两个 Excel 表按主匹配列补列
- 支持 `.xlsx` / `.xls`
- 支持手动设置“目标表列 -> 来源表列”的附加匹配映射
- 支持手动设置“结果列名 <- 来源表列”的补充映射
- 支持生成带“匹配结果清单”sheet 的新结果文件

### 5. 项目阶段汇总

- 支持配置多台 SQL Server 服务器
- 自动遍历服务器上的在线数据库
- 自动忽略不包含业务表的数据库
- 支持查看“正在进行 / 即将开始 / 全部”阶段
- 支持按阶段关键字、项目关键字筛选
- 支持导出 Excel 汇总

### 6. 项目阶段汇总 Web 服务

- 支持在浏览器中访问项目阶段汇总
- 支持本机保存多台服务器配置
- 支持开始查询、测试连接、导出 Excel
- 复用桌面版查询逻辑
- 实际查询在独立子进程执行，降低 ODBC 驱动异常导致服务整体退出的风险

启动方式：

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
source .venv/bin/activate
pip install -r requirements.txt
PYTHONPATH=src python3 -m aliyun_photo_manager.project_stage_web
```

访问地址：

- `http://127.0.0.1:8000`

### 7. 结果打包

- 支持选择任意结果文件或文件夹一键压缩
- 支持 AES 加密 zip
- 压缩包默认使用原文件或文件夹名称
- 支持自动生成密码或手动设置密码
- 支持保存打包记录并查询密码

### 8. 考场编排

- 支持三张标准模板驱动考场编排
- 支持一键导出三张标准模板，补充后直接导入使用
- 支持补充考点、考场、座号、考号
- 支持在程序中配置考号拼接规则，可选岗位编码、科目号等片段
- 支持岗位归组表新增字段后，直接作为考号规则片段使用
- 支持混编考场按座位区间分配

### 9. 更新SQL生成

- 支持导入字段映射模板
- 支持输入考生表名称和临时表名称
- 支持分别选择两边的关联字段
- 支持忽略空值，避免空值覆盖正式表
- 生成的 SQL 会先备份两张表，再执行更新
- 支持导出标准字段映射模板

## 安装依赖

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 启动桌面程序

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
source .venv/bin/activate
PYTHONPATH=src python3 -m aliyun_photo_manager.gui
```

## 使用说明

### 照片下载与分类

适用场景：

- 下载考生照片
- 只下载名单中的照片
- 按模板内容分类整理照片

基本流程：

1. 选择数据来源
2. 如果使用云存储，填写云类型、Endpoint / Region、AccessKey，点击“加载 Bucket”
3. 选择 bucket 后，点击右侧“加载当前层级”
4. 选择下载目录和分类目录
5. 如果只想下载模板中的人员，选择模板并勾选“只下载模板中的人员”
6. 点击“下载并生成模板”
7. 打开模板填写分类信息
8. 点击“按模板分类”

结果文件：

- `照片分类模板.xlsx`
- `照片分类结果清单.xlsx`

### 证件资料筛选

适用场景：

- 下载部分人的证件资料
- 按模板筛选证件资料
- 只导出某类材料

基本流程：

1. 选择数据来源
2. 选择人员模板
3. 点击“加载模板列”
4. 选择匹配列
5. 如果需要，勾选“导出后文件夹重命名”，再选择名称列
6. 如果使用云存储，填写云类型、Endpoint / Region、AccessKey，点击“加载 Bucket”
7. 选择 bucket 后，点击右侧“加载当前层级”
8. 选择证件资料目录
9. 如果要先从云端下载，点击“下载证件资料”
10. 选择输出目录
11. 选择筛选模式
12. 点击“开始筛选”

结果文件：

- `证件资料筛选结果清单.xlsx`

### 表样转换

适用场景：

- 把 Word / Excel 报名表模板转换成 HTML

基本流程：

1. 选择表样文件
2. 点击“Net版导出”或“Java版导出”
3. 在“代码”页复制 HTML
4. 如需看效果，点击“浏览器预览”

占位符示例：

- Net：`{[#考生表视图.姓名#]}`
- Java：`${考生.姓名}`

### 项目阶段汇总

适用场景：

- 一次查看多台数据库服务器中的项目阶段状态
- 汇总当前“正在进行”或“即将开始”的阶段
- 快速定位哪个服务器、哪个数据库、哪个项目正在开放

基本流程：

1. 在“服务器配置”中填写服务器名称、数据库地址、端口、用户名和密码
2. 点击“新增/更新”保存服务器
3. 可先点“测试连接”
4. 设置状态、阶段关键字、项目关键字
5. 点击“开始查询”
6. 如需汇总给同事，可点击“导出 Excel”

结果说明：

- 结果表会显示：服务器、数据库、项目名称、阶段名称、开始时间、结束时间、当前状态
- 只有同时存在 `EI_ExamTreeDesc`、`web_SR_CodeItem`、`WEB_SR_SetTime` 三张表的数据库才会参与统计

### 数据匹配

适用场景：

- 两个 Excel 表之间按考号、身份证号、姓名等字段补列
- 代替手工写 VLOOKUP、XLOOKUP 或临时写 SQL
- 姓名重名时，可以再叠加单位、岗位等附加列做更准确的关联

基本流程：

1. 选择目标表和来源表（支持 `.xlsx` / `.xls`）
2. 点击“加载表头”
3. 选择目标表匹配列和来源表匹配列
4. 如需提高准确性，设置附加匹配映射，例如 `单位 -> 报考单位`
5. 设置补充列映射，例如 `身份证号 <- 身份证号`
6. 点击“开始匹配”

结果说明：

- 会生成一个新的结果 Excel
- 原目标表不改动
- 新文件中会增加补充列
- 结果文件里会带一个“匹配结果清单”sheet
- 如果第一行只是“附件1”之类说明，程序会尽量自动跳过并识别真正表头

### 考场编排

适用场景：

- 根据标准模板给考生批量补充考点、考场、座号、考号
- 考号规则不固定时，在程序中配置拼接顺序
- 混编考场时，通过编排片段表控制同一考场的不同座位区间

基本流程：

1. 可以先点击“导出标准模板”，自动生成三张标准表。
2. 准备三张标准表：
   - 考生明细表：`姓名`、`身份证号`、`招聘单位`、`岗位名称`
   - 岗位归组表：`招聘单位`、`岗位名称`、`科目组`、`岗位编码`、`科目号`
   - 编排片段表：`科目组`、`考点`、`考场号`、`起始座号`、`结束座号`、`人数`、`起始流水号`、`结束流水号`、`备注`
3. 选择三张表
4. 设置考点、考场、座号、流水号位数
5. 选择同组内顺序：按原顺序或随机打乱
6. 按顺序添加考号规则
7. 点击“开始编排”

结果说明：

- 会生成一个新的编排结果 Excel
- 会新增 `科目组`、`岗位编码`、`科目号`、`考点`、`考场`、`座号`、`考号`
- 如果有未归组或未编排成功的人员，会在 `编排备注` 中写明原因
- 同一科目组内默认随机打乱后再分配，也可以手动切换为按原顺序

### 结果打包

适用场景：

- 把照片结果、证件资料结果或其他交付目录打包发给客户

基本流程：

1. 选择待打包文件或文件夹
2. 选择输出目录
3. 如需客户指定密码，可勾选“手动设置密码”并输入密码
4. 点击“一键打包并加密”
5. 复制程序生成或手动输入的密码
6. 把 zip 和密码分别发给对方

密码规则：

- 自动密码：`当天日期 + 4位随机字符`
- 示例：`260313A7KQ`

查询说明：

- 每次打包都会把压缩包路径、来源名称、密码、打包时间写入本地 `.pack_history.json`
- 如果后面忘记密码，可以在“结果打包”页按文件名、文件夹名、压缩包名或密码查询

### 更新SQL生成

适用场景：

- 考生表先导入基础字段，补充字段先导入临时表
- 按字段映射关系生成标准 UPDATE SQL

基本流程：

1. 点击“导出模板”，生成“更新SQL字段映射模板.xlsx”
2. 在模板中填写：
   - `考生表字段名`
   - `临时表字段名`
   - `是否更新`
3. 在程序中选择映射模板并点击“加载字段”
4. 输入考生表名称和临时表名称
5. 选择考生表关联字段和临时表关联字段
6. 如需避免空值覆盖正式表，勾选“忽略空值，不覆盖正式表”
7. 点击“生成 SQL”
8. 点击“复制 SQL”，再到数据库工具中执行

结果说明：

- 生成的 SQL 会先备份考生表和临时表
- 只更新模板中 `是否更新` 为 `是` 的字段

## 云配置说明

程序会分别保存两套云配置：

- 阿里云 OSS
- 腾讯云 COS

切换云类型时，会自动带出各自保存过的：

- AccessKey ID
- AccessKey Secret
- Endpoint / Region
- bucket
- 前缀

## 开发说明

如果后面需要继续维护或加功能，可以先看下面这几个核心模块：

- [app.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/app.py)
  - 负责照片下载、生成模板、按模板分类这几段主流程
- [downloader.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/downloader.py)
  - 负责阿里云 OSS / 腾讯云 COS 的 bucket、目录浏览、对象下载
- [excel_classifier.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/excel_classifier.py)
  - 负责照片分类模板生成、模板校验、按模板复制分类、导出结果清单
- [certificate_filter.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/certificate_filter.py)
  - 负责证件资料按模板筛选、按分类目录导出、导出结果清单
- [word_to_html.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/word_to_html.py)
  - 负责表样转换、占位符生成、预览 HTML 构造
- [result_packer.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/result_packer.py)
  - 负责目录压缩、AES 加密、自动生成压缩包名称和密码
- [data_matcher.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/data_matcher.py)
  - 负责 Excel 表之间的匹配补列和匹配结果清单导出
- [exam_arranger.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/exam_arranger.py)
  - 负责按标准模板进行考场编排和考号生成
- [gui.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/gui.py)
  - 负责桌面界面、参数收集、按钮事件、日志显示、结果展示
- [config.py](/Users/dalezhao/python_learning/aliyun_photo_manager/src/aliyun_photo_manager/config.py)
  - 负责云配置校验和环境变量读取

建议修改顺序：

1. 先确认需求属于哪个业务页
2. 再看对应流程模块
3. 最后再改 GUI 按钮和交互

这样不容易把界面逻辑和业务逻辑混在一起。

## 打包

### macOS

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
source .venv/bin/activate
python3 build_macos_app.py
```

### Windows

```bat
cd aliyun_photo_manager
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python build_windows_app.py
```

## 自定义图标

把图片保存为：

```text
assets/app_icon.png
```

然后执行：

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
source .venv/bin/activate
python3 generate_app_icons.py
```

会生成：

- `assets/app_icon.icns`
- `assets/app_icon.ico`
