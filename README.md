# 木材检尺对比系统

一个本地可运行的 Flask 小工具，用于上传 Excel 码单、映射字段、按编号搜索木材记录，并按国标或外标公式重新计算材积差异。

## 功能

- 上传 `.xlsx` / `.xlsm` 码单文件
- 自动识别常见表头，也可手动映射字段
- 每个上传文件单独保存，支持加载、重命名、删除
- 按顺序编号模糊搜索
- 查看木材详情，并对比原材积与重新计算材积

## 启动

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

启动后访问：

```text
http://127.0.0.1:5050
```

## 数据文件

- SQLite 数据库：`checker.db`
- 上传文件目录：`uploads/`

这两个文件/目录会在应用启动时自动创建。

## Excel 要求

第一行需要是表头，数据从第二行开始。至少要映射「编号」列，其余字段可按实际码单选择。

当前支持 openpyxl 可解析的 `.xlsx` 和 `.xlsm`。老式 `.xls` 不是同一种文件格式，需要先用 Excel 或 WPS 另存为 `.xlsx` 后再上传。
