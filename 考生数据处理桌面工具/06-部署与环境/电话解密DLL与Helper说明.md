# 电话解密DLL与Helper说明

## 组成

- `Interop.DeDll.dll`
- `DeDLL.dll`
- `PhoneDecryptHelper.exe`

## 构建方式

通过 `build_phone_decrypt_helper.py` 调用 `dotnet publish` 生成 helper。

## 注意事项

- 仅 Windows 环境支持构建
- DLL 缺失时电话解密功能不可用
- 打包时 helper 会被附加到发布产物
