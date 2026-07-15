# phone_decrypt.py

## 模块定位

封装电话解密相关业务能力。

## 主要职责

- 组织电话解密输入参数
- 调用 helper 或底层依赖
- 返回解密结果与异常信息

## 风险点

- 依赖 Windows 环境
- 依赖外部 DLL
- 打包时必须保证 helper 与 DLL 一致
