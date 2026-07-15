# PyInstaller说明

## 用途

用于生成 Windows 与 macOS 桌面发布产物。

## 相关脚本

- `build_windows_app.py`
- `build_macos_app.py`

## 打包特点

- 以 `app_launcher.py` 为入口
- 显式声明隐藏导入
- 支持附加电话解密 helper 二进制
