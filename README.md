# 物探检测台账系统

这是一个用于管理物探检测台账资料的本地 Web 系统。

## 启动

在 Windows 下运行：

```bat
work\start_ledger_server.bat
```

默认访问：

```text
http://127.0.0.1:8765/
```

备用端口：

```text
http://127.0.0.1:8766/
```

## 运行依赖

需要 Python 3，并安装：

```bat
pip install -r requirements.txt
```

## Git 管理范围

Git 只管理代码和说明文档。

以下内容不纳入 Git，需要单独备份或交付：

```text
outputs/ledger_system/ledger_system.db
outputs/ledger_system/storage/
```

详细需求和后续开发建议见：

```text
项目需求说明.md
```
