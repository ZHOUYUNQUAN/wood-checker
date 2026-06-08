# 木材检尺对比系统 - Code Wiki

## 项目概述

**项目名称**: 木材检尺对比系统
**项目类型**: Flask Web 应用
**功能**: 上传 Excel 码单、映射字段、按编号搜索木材记录，并按国标或外标公式重新计算材积差异
**依赖**: Flask >=3.1,<4, openpyxl >=3.1,<4

---

## 项目架构

```
/workspace/
├── app.py                    # 应用入口，日志配置，反向代理支持
├── config.py                 # 配置类（数据库、密钥、上传目录）
├── models.py                 # 数据库模型与初始化
├── requirements.txt         # Python 依赖
├── check/                    # 检尺模块（蓝图）
│   ├── __init__.py          # 蓝图定义
│   ├── routes_upload.py     # 文件上传相关路由
│   ├── routes_search.py     # 搜索与计算路由
│   └── utils.py             # 工具函数
├── templates/               # HTML 模板
│   ├── base.html            # 基础模板
│   └── check/
│       └── checker.html     # 检尺页面模板
├── static/
│   └── js/
│       ├── checker.js       # 前端交互逻辑
│       └── vendor-lite.js   # 第三方库（jQuery/Bootstrap）
└── uploads/                 # 上传文件存储目录（运行时自动创建）
```

---

## 主要模块职责

### 1. 应用入口 (app.py)

**职责**: Flask 应用初始化、日志配置、反向代理支持、蓝图注册

**关键功能**:
- 日志系统配置（RotatingFileHandler + ConsoleHandler，10MB 轮转）
- `ReverseProxied` 类：支持反向代理环境下的路径处理
- 蓝图注册：`check_bp` 注册到 `/check` 前缀
- `ensure_runtime()`: 启动时确保上传目录和数据库表存在

**启动方式**:
```bash
python app.py
# 访问 http://127.0.0.1:5050
```

---

### 2. 配置模块 (config.py)

**类**: `Config`

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `BASE_DIR` | 项目根目录 | `os.path.abspath(os.path.dirname(__file__))` |
| `DATABASE` | SQLite 数据库路径 | `checker.db` |
| `SECRET_KEY` | Flask 密钥 | 环境变量 `SECRET_KEY` 或 `'wood-checker-secret-dev'` |
| `UPLOAD_FOLDER` | 上传文件存储目录 | `uploads/` |
| `MAX_CONTENT_LENGTH` | 最大上传大小 | 16MB |

---

### 3. 数据库模型 (models.py)

**数据库**: SQLite (WAL 模式，外键约束开启)

#### 表结构

**code_sheets** (旧表，兼容过渡)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER | 主键 |
| file_name | TEXT | 上传文件名 |
| upload_time | TEXT | 上传时间 |
| no | TEXT | 顺序编号 |
| especie | TEXT | 材种 |
| english_code | TEXT | 英文代码 |
| diameter_1~4 | REAL | 直径 1-4 |
| diameter_avg | REAL | 平均直径 |
| length_m | REAL | 长度(米) |
| volume_m3 | REAL | 材积(立方米) |
| customer | TEXT | 供应商/基地 |
| base_name | TEXT | 基础名称 |
| extra_json | TEXT | 扩展数据(JSON) |
| is_transshipment | INTEGER | 是否转口(0/1) |

**file_registry** (文件注册表)
| 字段 | 类型 | 说明 |
|------|------|------|
| id | TEXT | 主键(文件安全名称) |
| file_name | TEXT | 文件显示名 |
| table_name | TEXT | 独立数据表名 |
| row_count | INTEGER | 记录数 |
| upload_time | TEXT | 上传时间 |

**独立数据表** (每上传一个文件创建一个，如 `sheets_20250608_a1b2c3d4`)
- 结构同 `code_sheets`，存储该文件的全部记录

#### 关键函数

| 函数 | 说明 |
|------|------|
| `get_db()` | Context manager 获取数据库连接 |
| `init_db()` | 初始化数据库表（启动时调用） |
| `_validate_table_name(table_name)` | 校验表名格式，防止 SQL 注入 |
| `_has_column(cursor, table_name, column_name)` | 检查列是否存在 |
| `_add_column_if_missing(cursor, table_name, column_name, column_sql)` | 列不存在时添加 |

---

### 4. 检尺蓝图 (check/)

#### 蓝图初始化 (check/__init__.py)

```python
check_bp = Blueprint('check', __name__,
                     template_folder='../templates/check',
                     static_folder='../static')
```

#### 路由 - 上传 (routes_upload.py)

| 路由 | 方法 | 说明 |
|------|------|------|
| `/check/upload` | POST | 第一步：上传 Excel，返回表头供选择 |
| `/check/upload/preview_sheet` | POST | 切换 sheet 时预览表头 |
| `/check/upload/confirm` | POST | 第二步：确认列映射，解析入库 |
| `/check/files` | GET | 获取已上传文件列表 |
| `/check/files/delete` | POST | 删除文件及数据 |
| `/check/files/rename` | POST | 重命名文件 |

**上传流程**:
1. `upload()`: 保存文件到 `uploads/`，解析 Excel 表头返回前端
2. `upload_confirm()`: 根据用户列映射创建独立数据表，写入 `file_registry`

#### 路由 - 搜索与计算 (routes_search.py)

| 路由 | 方法 | 说明 |
|------|------|------|
| `/check/` | GET | 主页，渲染检尺页面 |
| `/check/search` | GET | 模糊搜索编号，返回匹配记录 |
| `/check/calc` | POST | 按国标/外标计算新材积 |
| `/check/export` | GET | 导出计算结果为 CSV |

**计算逻辑**:
- 根据 `standard` 参数选择公式
- 结果保存到 `extra_json.calc_result`
- 返回原材积、新材积、差异、涨尺率

---

### 5. 工具函数 (check/utils.py)

#### 常量

| 常量 | 说明 |
|------|------|
| `ALLOWED_EXTENSIONS` | 允许的文件扩展名 `{'xlsx', 'xlsm'}` |
| `SKIP_KEYWORDS` | 跳过汇总行的关键词 `('总计', '合计', ...)` |
| `PRESET_LABELS` | 预设列标签候选列表 |

#### 文件处理

| 函数 | 说明 |
|------|------|
| `allowed_file(filename)` | 检查文件扩展名是否允许 |
| `make_safe_upload_name(filename)` | 生成安全的文件名（保留中文） |
| `safe_upload_path(file_key)` | 返回上传文件绝对路径，非法返回 None |
| `get_all_files()` | 获取合并后的文件列表（registry + 旧 code_sheets） |
| `resolve_table(cursor, file_name)` | 根据 file_name 查找对应数据表名 |

#### Excel 解析

| 函数 | 说明 |
|------|------|
| `read_headers(ws)` | 读取 Excel 第一行作为表头列表 |
| `parse_row(row_data, col_map)` | 解析一行数据，跳过汇总行/空行，返回字段字典 |

**parse_row 支持的字段映射**:
- `no`, `especie`, `english_code`
- `diameter_1`, `diameter_2`, `diameter_3`, `diameter_4`, `diameter_avg`
- `length_m`, `volume_m3`, `customer`, `is_transshipment`

#### 材积计算

| 函数 | 公式 | 说明 |
|------|------|------|
| `calc_external_standard(diameter, length)` | `V = π × (D/100)² × L / 4` | 外标公式 |
| `calc_national_standard(diameter, length)` | 条件分支复杂公式 | 国标公式（5种条件） |

**国标公式条件分支**:
- 条件1: `D <= 12 && L <= 10`
- 条件2: `L <= 10 && D >= 14`
- 条件3: `L >= 10.4`
- 条件4: `L == 10.2 && D <= 12`
- 条件5: 兜底

---

## 数据库依赖关系

```
app.py
  ├── config.Config
  ├── models.init_db()
  └── check.check_bp (蓝图)
        ├── routes_upload.py
        │     ├── models.get_db()
        │     ├── models._validate_table_name()
        │     └── utils.py (allowed_file, make_safe_upload_name, ...)
        └── routes_search.py
              ├── models.get_db()
              ├── models._validate_table_name()
              └── utils.py (get_all_files, resolve_table, calc_*, ...)
```

---

## 运行方式

### 环境准备
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 启动应用
```bash
python app.py
# 访问 http://127.0.0.1:5050
```

### 运行时生成的文件/目录
- `checker.db`: SQLite 数据库
- `uploads/`: 上传文件存储
- `logs/`: 应用日志目录

---

## API 路由速查

### 文件管理
- `POST /check/upload` - 上传 Excel 文件
- `POST /check/upload/preview_sheet` - 预览指定 sheet 表头
- `POST /check/upload/confirm` - 确认列映射并入库
- `GET /check/files` - 获取文件列表
- `POST /check/files/delete` - 删除文件
- `POST /check/files/rename` - 重命名文件

### 搜索与计算
- `GET /check/search?q=xxx&file_name=xxx` - 搜索编号
- `POST /check/calc` - 计算材积
- `GET /check/export?file_name=xxx` - 导出 CSV

---

## 关键类与函数说明

### app.py

| 名称 | 类型 | 说明 |
|------|------|------|
| `ReverseProxied` | 类 | WSGI 中间件，处理反向代理路径 |
| `ensure_runtime()` | 函数 | 启动时初始化目录和数据库 |

### models.py

| 名称 | 类型 | 说明 |
|------|------|------|
| `get_db()` | Context Manager | 数据库连接管理 |
| `init_db()` | 函数 | 创建/升级数据库表 |
| `_validate_table_name()` | 函数 | 表名白名单校验 |

### check/utils.py

| 名称 | 类型 | 说明 |
|------|------|------|
| `parse_row()` | 函数 | Excel 行解析，跳过汇总行 |
| `calc_national_standard()` | 函数 | 国标配料公式 |
| `calc_external_standard()` | 函数 | 外标公式 |
| `resolve_table()` | 函数 | file_name → table_name 解析 |

### routes_search.py

| 名称 | 类型 | 说明 |
|------|------|------|
| `search()` | 路由函数 | 模糊搜索编号 |
| `calc()` | 路由函数 | 材积计算与结果存储 |
| `export_csv()` | 路由函数 | CSV 导出 |

### routes_upload.py

| 名称 | 类型 | 说明 |
|------|------|------|
| `upload()` | 路由函数 | Excel 上传与预览 |
| `upload_confirm()` | 路由函数 | 确认导入创建独立表 |
| `delete_file()` | 路由函数 | 删除文件及数据 |
| `rename_file()` | 路由函数 | 重命名文件 |
