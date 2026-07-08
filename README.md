# 物探检测台账管理系统

这是一个用于管理湖南安化抽水蓄能电站施工期物探检测台账资料的本地 Web 系统。系统把分散的 Excel 台账、上传文件、检测记录和统计报表统一管理，方便工程资料查询、质检、统计、导出和后续多人维护。

## 当前功能

- 台账管理：上传 Excel 台账文件，保存文件版本、标段、委托单位、项目名称等信息。
- 原表查看：按原 Excel 工作表在线预览台账内容，支持查看文件和工作表。
- 台账查询：按标段、委托单位、检测类型、文件名、日期范围、报告编号、委托编号、单位工程、工程部位、检测结果等条件筛选，并可导出 Excel。
- 数据质检：检查编号缺失、日期缺失、数量异常、检测结果缺失、文件名与标段或委托类型不一致等问题。报告编号或委托编号重复出现不再视为异常。
- 检测数据统计：按检测项目、标段、委托类别和时间范围生成周报、月报、常规检测数量统计表，支持 Word 和 Excel 导出。
- 单位工程统计：按单位工程筛选并生成季度、年度综合统计表，包含合同量完成情况、单位工程检测覆盖情况和各检测项目统计表，支持 Word 和 Excel 导出。
- 用户管理：支持管理员和普通用户账号，管理员可创建用户、重置密码和管理文件。

当前已覆盖的主要检测类型包括：

- 锚杆无损检测
- 锚杆拉拔试验
- 钻孔摄像检测
- 锚索多循环张拉
- 预应力锚杆张拉
- 洞室松弛圈检测
- 桩身完整性检测
- 回填灌浆质量单孔注浆试验
- 弹性波检测
- 钻孔成像

## 运行环境

需要 Windows 和 Python 3。依赖包见 `requirements.txt`：

```bat
pip install -r requirements.txt
```

主要依赖：

- `openpyxl`：读取和导出 Excel。
- `python-docx`：导出 Word 报表。

## 启动方式

在项目根目录运行：

```bat
work\start_ledger_server.bat
```

脚本会自动按自身位置识别项目目录，不依赖固定电脑路径。默认启动两个本地访问地址：

```text
http://127.0.0.1:8765/
http://127.0.0.1:8766/
```

如果需要修改端口，可以在启动前设置环境变量：

```bat
set LEDGER_WEB_PORT=8765
set LEDGER_WEB_BACKUP_PORT=8766
work\start_ledger_server.bat
```

健康检查地址：

```text
http://127.0.0.1:8765/health
```

## 登录账号

默认管理员账号：

```text
账号：ZhanLin2026
密码：Ahcx@ZL2026
```

也可以通过环境变量覆盖默认账号密码：

```bat
set LEDGER_ADMIN_USER=your_user
set LEDGER_ADMIN_PASSWORD=your_password
work\start_ledger_server.bat
```

首次启动时系统会初始化管理员账号。旧账号 `admin / admin123` 已停用。

## 数据目录

系统运行数据默认保存到：

```text
outputs/ledger_system/ledger_system.db
outputs/ledger_system/storage/
```

说明：

- `ledger_system.db` 是 SQLite 数据库，保存用户、文件信息、Excel 单元格、统计基础数据等。
- `storage/` 保存上传后的台账原始文件和版本文件。
- `server_stdout.log`、`server_stderr.log` 等日志也会写入 `outputs/ledger_system/`。

## 给别人运行时需要发什么

如果只发 GitHub 仓库代码，对方只能拿到系统程序，不能自动拿到你的本地台账数据。

如果希望对方打开后看到和你一样的数据，需要额外发送：

```text
outputs/ledger_system/ledger_system.db
outputs/ledger_system/storage/
```

放置位置必须保持在项目根目录下：

```text
项目根目录/
  work/
  outputs/
    ledger_system/
      ledger_system.db
      storage/
```

如果不发送这些数据，对方也可以运行系统，但需要自己上传台账文件后才会有内容。

## GitHub 协作方式

代码和说明文件由 Git 管理，运行数据默认不纳入 Git。建议多人维护时按以下流程：

1. 从 GitHub 拉取最新代码。
2. 本地修改代码或文档。
3. 运行 `work\start_ledger_server.bat` 验证功能。
4. 使用 Git 提交修改。
5. 推送到 GitHub。
6. 其他人拉取最新代码后重启本地服务。

常用命令：

```bat
git pull origin main
git status
git add .
git commit -m "说明本次修改"
git push origin main
```

代码更新后，本地网页不会自动变化，需要重启服务并刷新浏览器。

## 注意事项

- 当前系统是本地 Web 系统，`127.0.0.1` 只能在本机访问。
- 局域网访问可使用本机 IP 加端口，例如 `http://192.168.x.x:8765/`，但需要对方和你在同一网络，并允许防火墙访问。
- GitHub 仓库不是公共在线网站，它只保存代码；别人修改代码后，只有拉取代码并重启服务，网页内容才会更新。
- 统计表来源于导入的台账数据，不是从年报、季报、月报 Word 文件中复制表格。
- 报告编号或委托编号重复可能是正常业务场景，系统不再把重复编号作为异常。

## 主要文件

```text
work/ledger_web_app.py          Web 系统主程序
work/start_ledger_server.bat    Windows 启动脚本
requirements.txt                Python 依赖
README.md                       项目内容和使用说明
项目需求说明.md                 功能范围、交接说明和后续优化建议
outputs/ledger_system/          本地数据库、上传文件和日志
```
