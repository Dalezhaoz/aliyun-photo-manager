# Windows打包说明

## 推荐版本

- 推荐打包版本：`v2.1.0`

## 基本步骤

1. 拉取代码并切到目标 tag
2. 创建虚拟环境
3. 安装依赖
4. 本地运行验证
5. 如需电话解密，先构建 helper
6. 执行 PyInstaller 打包
7. 运行产物自测
8. 压缩后发同事

## 关键命令

```powershell
git fetch --all --tags
git checkout v2.1.0
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:PYTHONPATH="src"
python -m aliyun_photo_manager.gui
python .\build_windows_app.py
```

## Dev 分支补充

若要验证 Qt 预览版，可使用：

```powershell
$env:PYTHONPATH="src"
python -m aliyun_photo_manager.qt_gui
```

## 电话解密附加步骤

```powershell
python .\build_phone_decrypt_helper.py
```

前提：

- 项目根目录已放置 `Interop.DeDll.dll`
- 项目根目录已放置 `DeDLL.dll`
