# Dev迭代记录

## 关键迭代

### 1. 页面与框架演进

- `90167a5`
  Refactor dev GUI into modular UI actions
- `d0de299`
  Redesign main UI shell and home navigation
- `5811796`
  Add Qt preview pages for all tools

### 2. 项目阶段能力

- `4325ebb`
  Add project stage summary for dev
- `236548f`
  Isolate stage summary query in subprocess
- `f9c5686`
  Fix Windows stage summary runner startup
- `28628f8`
  Show detailed stage summary subprocess errors
- `f482dbd`
  Add project stage web service

### 3. 电话解密能力演进

- `7877d17`
  Add phone decrypt tab on dev
- `0d381dc`
  Improve phone decrypt DLL loading
- `d0b740b`
  Add x86 phone decrypt helper
- `796c991`
  Build phone decrypt helper self-contained
- `873fca7`
  Fix phone decrypt crash: JSON case mismatch, missing MEIPASS path, and 64-bit DLL guard
- `fba782a`
  Add crash logging and defensive error handling
- `c1a46e1`
  Add faulthandler for native crash capture and fix subprocess in windowed mode
- `c43071e`
  Fix pyodbc access violation crash by probing ODBC drivers in subprocess
- `9f2c34f`
  Filter non-SQL-Server ODBC drivers and improve missing driver error message
- `52dc707`
  Add Encrypt=yes for Driver 18 and log all connection probe attempts
- `6f00d56`
  Add pymssql as fallback when pyodbc crashes on Python 3.14
- `6e614da`
  Don't abort on missing pyodbc if pymssql also unavailable
- `6f20d34`
  Move all DB operations to C# helper, eliminate pyodbc/pymssql dependency
- `c3d2f4f`
  Remove phone decrypt export feature, keep only DB update to 备用3
- `be4075e`
  Wrap phone decrypt updates in transaction

### 4. 身份证工具演进

- `2a8ddae`
  Add ID card validation and generator tab
- `a6525eb`
  Add province city county ID region selectors
- `26da714`
  Use full 2024 region data for ID selectors

### 5. 分支合流背景

- `edb86cd`
  Merge main features into dev

## 结论

`dev` 当前已经不是简单试验分支，而是：

- Qt 页面迁移主战场
- 电话解密稳定性增强分支
- 身份证工具能力增强分支
- 多实验功能整合分支
