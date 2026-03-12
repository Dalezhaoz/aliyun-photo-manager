# aliyun_photo_manager

一个用于从阿里云 OSS 存储桶批量下载文件、生成 Excel 模板并在本地复制分类，同时支持按人员模板筛选证件资料的 Python 项目。

## 功能

- 按前缀批量下载 OSS 对象
- 支持跳过已存在文件
- 在下载目录当前层自动生成 Excel 模板
- 按 Excel 中的“分类一/分类二/分类三/修改名称”复制分类
- 支持 `dry-run` 预览
- 提供桌面 GUI 界面
- 支持树形浏览 bucket 文件夹
- 支持统计已选文件夹下的图片数量
- 支持在当前文件夹内按文件名搜索文件
- 支持下载进度条和大批量下载状态提示
- 自动保存上次使用的界面配置
- 支持按人员模板筛选证件资料
- 支持按模板任意列匹配人员文件夹
- 支持复制整个人员文件夹或只复制指定关键词文件
- 支持 Word 模板转简洁 HTML
- 支持 `Net版导出` 与 `Java版导出`

## 目录结构

```text
aliyun_photo_manager/
├── README.md
├── requirements.txt
└── src/
    └── aliyun_photo_manager/
        ├── app.py
        ├── __init__.py
        ├── cli.py
        ├── config.py
        ├── downloader.py
        ├── gui.py
        └── sorter.py
```

## 安装依赖

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 配置 OSS 凭证

建议使用环境变量：

```bash
export OSS_ACCESS_KEY_ID="你的 AccessKey ID"
export OSS_ACCESS_KEY_SECRET="你的 AccessKey Secret"
export OSS_ENDPOINT="https://oss-cn-hangzhou.aliyuncs.com"
export OSS_BUCKET_NAME="你的 bucket 名称"
```

如果使用内网、区域或自定义域名，替换 `OSS_ENDPOINT` 即可。

## 使用方式

先预览下载和模板生成计划：

```bash
python3 -m aliyun_photo_manager.cli \
  --prefix photos/2025/ \
  --download-dir ./downloads \
  --sorted-dir ./sorted \
  --dry-run
```

正式执行：

```bash
python3 -m aliyun_photo_manager.cli \
  --prefix photos/2025/ \
  --download-dir ./downloads \
  --sorted-dir ./sorted
```

如果只想基于本地下载目录重新生成模板或按模板分类，不下载：

```bash
python3 -m aliyun_photo_manager.cli \
  --skip-download \
  --download-dir ./downloads \
  --sorted-dir ./sorted
```

## 启动桌面界面

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
source .venv/bin/activate
PYTHONPATH=src python3 -m aliyun_photo_manager.gui
```

## 打包成 macOS 应用

先安装依赖：

```bash
cd /Users/dalezhao/python_learning/aliyun_photo_manager
source .venv/bin/activate
pip install -r requirements.txt
```

执行打包：

```bash
python3 build_macos_app.py
```

打包完成后，应用会出现在：

```text
dist/阿里云照片下载与分类.app
```

说明：

- 这是在 macOS 上生成本地 `.app` 的方式
- 第一次打开如果被系统拦截，需要到“系统设置 -> 隐私与安全性”里允许打开
- 如果你修改了代码，重新执行一次 `python3 build_macos_app.py` 就会生成新版本
- 打包脚本会尝试自动移除隔离属性并重新签名

如果已经打包好了但双击打不开，可以手动执行：

```bash
xattr -dr com.apple.quarantine "/Users/dalezhao/python_learning/aliyun_photo_manager/dist/阿里云照片下载与分类.app"
codesign --force --deep --sign - "/Users/dalezhao/python_learning/aliyun_photo_manager/dist/阿里云照片下载与分类.app"
```

## 打包成 Windows 应用

说明：

- `Windows .exe` 建议在 Windows 机器上打包
- 不能在 macOS 上直接可靠地生成可用的 Windows `.exe`

在 Windows 上执行：

```bash
cd aliyun_photo_manager
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python build_windows_app.py
```

打包完成后，应用会出现在：

```text
dist\aliyun_photo_manager\aliyun_photo_manager.exe
```

如果你修改了代码，重新执行一次 `python build_windows_app.py` 就会生成新版本。

界面首页现在叫“报名系统工具箱”，包含两个业务页：

- `照片下载与分类`
- `证件资料筛选`
- `Word 转 HTML`

两个业务页都支持两种来源：

- `OSS`
- `本地`

照片下载页：

- OSS 模式：从 OSS 下载照片后，生成模板，再按模板分类
- 本地模式：直接选择本地照片目录，生成模板，再按模板分类

照片下载页里可以直接填写：

- Endpoint
- AccessKey ID / Secret
- 加载并选择可用 bucket
- 加载并选择 bucket 下的文件夹前缀
- 下载目录
- 分类目录
- 是否只预览、是否跳过下载、是否保留重复文件

目录保护规则：

- 你在界面里选择的“下载目录”会自动再创建一层 `下载文件`
- 你在界面里选择的“分类目录”会自动再创建一层 `分类结果`
- 例如你选择桌面，实际写入会变成：
  - `桌面/下载文件`
  - `桌面/分类结果`

右侧浏览区支持：

- 加载当前层级的 bucket 子文件夹
- 双击进入子文件夹
- 返回上一级前缀
- 刷新当前已选文件夹下的图片数量
- 在当前文件夹内按文件名关键词搜索文件
- 双击搜索结果，自动定位到文件所在目录

桌面工具会把最近一次填写的配置保存到项目根目录下的 `.gui_settings.json`。

## 证件资料筛选

证件资料筛选页适合这种目录结构：

```text
证件资料/
  身份证号或报名序号/
    某个人员的证件文件...
```

使用方式：

1. 选择来源：
   - 直接使用本地证件资料目录
   - 先从 OSS 下载证件资料
2. 如果是 OSS 模式，填写：
   - Endpoint
   - AccessKey ID / Secret
   - Bucket
   - 证件资料前缀
3. 选择人员模板 `.xlsx`
4. 点击“加载模板列”
5. 选择“匹配列”
6. 选择证件资料目录
   - 本地模式：这里就是本地证件资料目录
   - OSS 模式：这里是下载到本地的缓存目录
7. 选择输出目录
8. 选择筛选模式：
   - 复制整个人员文件夹
   - 只复制关键词文件（例如 `学历证书`）
9. 选择是否按 `分类一/分类二/分类三` 建目录
10. 点击“开始筛选证件资料”

输出规则：

- 不分类时：`输出目录/匹配值/`
- 分类时：`输出目录/分类一/分类二/分类三/匹配值/`

## Word 转 HTML

Word 转 HTML 页适合把报名登记表这类 Word 模板转成可嵌入系统的 HTML 模板。

使用方式：

1. 选择 `.docx` 或 `.doc` 文件
2. 点击：
   - `Net版导出`
   - `Java版导出`
3. 导出后可在页面内切换：
   - `代码`
   - `预览`
4. 需要时可直接点击“复制 HTML”

导出规则：

- 尽量保留 Word 表格结构，生成更简洁的 HTML
- 表格中“标题左、值右”的空白单元格，会自动补成占位符
- Java 版变量格式：
  - `${考生.姓名}`
- Net 版变量格式：
  - `{[#考生表视图.姓名#]}`
- 照片字段会直接输出：
  - Java：`${考生.照片}`
  - Net：`{[#考生表视图.照片#]}`

说明：

- `.docx` 可直接处理
- `.doc` 会尝试调用 LibreOffice 转换；如果系统没有 LibreOffice，建议先另存为 `.docx`
- 匹配列可以是模板中的任意列，例如 `身份证号` 或 `报名序号`
- 程序默认只读取模板第一个 sheet

## Excel 模板说明

下载完成后，程序会在下载目录当前层生成：

```text
照片分类模板.xlsx
```

表头固定为：

- `文件名称`
- `分类一`
- `分类二`
- `分类三`
- `修改名称`

使用方式：

1. 第一次执行：下载文件并生成 Excel 模板
2. 打开 Excel，填写分类和重命名信息
3. 第二次执行：程序按 Excel 内容把文件复制到分类目录

规则：

- 只统计下载目录当前这一层的文件，不递归子文件夹
- `分类一/分类二/分类三` 都可以留空
- 空分类会被跳过
- 如果三列分类都为空，文件会直接复制到分类目录根目录
- `修改名称` 可以留空；留空时保留原文件名
- 填写新名称但没写扩展名时，程序会自动保留原扩展名
- 分类动作是复制，不会改动下载目录里的原文件

## 输出结构

例如某一行 Excel 填写为：

```text
文件名称：IMG_001.jpg
分类一：考生照片
分类二：张三
分类三：身份证
修改名称：张三身份证
```

复制后的结果会是：

```text
sorted/考生照片/张三/身份证/张三身份证.jpg
```

## 说明

- 程序不会修改 OSS 桶里的文件，只会读取列表并下载到本地。
- 程序不会移动下载目录里的原文件，分类时只做复制。
