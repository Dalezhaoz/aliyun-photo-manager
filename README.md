# 报名系统工具箱

当前版本：`v1.0.0`

这是一个桌面工具，主要用于：

- 考生照片下载与分类
- 证件资料下载与筛选
- Word 模板转 HTML

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

### 3. Word 转 HTML

- 支持 `.doc` / `.docx`
- 支持 Net 版导出
- 支持 Java 版导出
- 支持代码查看
- 支持复制 HTML
- 支持浏览器预览

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

### Word 转 HTML

适用场景：

- 把 Word 报名表模板转换成 HTML

基本流程：

1. 选择 Word 文件
2. 点击“Net版导出”或“Java版导出”
3. 在“代码”页复制 HTML
4. 如需看效果，点击“浏览器预览”

占位符示例：

- Net：`{[#考生表视图.姓名#]}`
- Java：`${考生.姓名}`

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
  - 负责 Word 转 HTML、占位符生成、预览 HTML 构造
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
