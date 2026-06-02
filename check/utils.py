"""
木材检尺对比系统 - 工具函数
包含常量、Excel 解析、材积计算、文件管理等辅助函数
"""
import math
import os
from datetime import datetime
from pathlib import Path

from flask import current_app
from werkzeug.utils import secure_filename
import openpyxl

from models import get_db

# ========== 常量 ==========
ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

# 需要跳过汇总行的关键词
SKIP_KEYWORDS = ('总计', '合计', 'Grand Total', '小计', 'Subtotal')

# 预设列标签候选（按顺序展示，长度 12）
PRESET_LABELS = [
    '顺序编号', '材种编码', '英文代码',
    '直径1', '直径2', '直径3', '直径4',
    '直径(CM)', '长度(M)', '材积(M³)', '客户', '是否转口'
]


def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def make_safe_upload_name(filename):
    """生成适合保存到上传目录的文件名，兼容中文原始文件名。"""
    ext = Path(filename).suffix.lower()
    stem = Path(filename).stem
    safe_stem = secure_filename(stem) or datetime.now().strftime('upload_%Y%m%d_%H%M%S')
    return f'{safe_stem}{ext}'


def safe_upload_path(file_key):
    """返回上传文件绝对路径；非法 file_key 返回 None。"""
    if not file_key or os.path.basename(file_key) != file_key:
        return None
    if secure_filename(Path(file_key).stem) != Path(file_key).stem:
        return None
    if Path(file_key).suffix.lower().lstrip('.') not in ALLOWED_EXTENSIONS:
        return None
    return os.path.join(current_app.config['UPLOAD_FOLDER'], file_key)


def safe_float(value, default=0.0):
    """安全转换为 float，非数值返回默认值"""
    if value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def read_headers(ws):
    """读取 Excel 第一行作为表头列表"""
    row = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    headers = []
    for cell in row:
        if cell is not None:
            headers.append(str(cell).strip())
        else:
            headers.append('')
    return headers


def parse_row(row_data, col_map):
    """
    解析 Excel 一行数据，返回字典或 None（需跳过的行）
    col_map: {field_name: column_index}，例如 {'no': 0, 'especie': 1, ...}
    """
    # 检查是否为汇总行
    for keyword in SKIP_KEYWORDS:
        for cell in row_data:
            if cell is not None and keyword in str(cell):
                return None

    # 检查是否全空行
    if all(cell is None for cell in row_data):
        return None

    no_idx = col_map.get('no')
    if no_idx is None:
        return None
    no_val = row_data[no_idx] if no_idx < len(row_data) else None
    no = str(no_val or '').strip()
    if not no:
        return None

    def _val(field):
        idx = col_map.get(field)
        if idx is not None and idx < len(row_data) and row_data[idx] is not None:
            return row_data[idx]
        return None

    def _str_val(field):
        return str(_val(field) or '').strip()

    return {
        'no': no,
        'especie': _str_val('especie'),
        'english_code': _str_val('english_code'),
        'diameter_1': safe_float(_val('diameter_1')),
        'diameter_2': safe_float(_val('diameter_2')),
        'diameter_3': safe_float(_val('diameter_3')),
        'diameter_4': safe_float(_val('diameter_4')),
        'diameter_avg': safe_float(_val('diameter_avg')),
        'length_m': safe_float(_val('length_m')),
        'volume_m3': safe_float(_val('volume_m3')),
        'customer': _str_val('customer'),
        'is_transshipment': 1 if _str_val('is_transshipment') in (
            '是', 'Y', 'y', 'Yes', 'yes', '1', 'True', 'true'
        ) else 0,
    }


def calc_external_standard(diameter, length):
    """外标公式：V = π × (D/100)² × L / 4"""
    return math.pi * (diameter / 100) ** 2 * length / 4


def calc_national_standard(diameter, length):
    """
    国标公式（转自 Excel IF 嵌套），结果保留 4 位小数
    条件1: D <= 12 && L <= 10
    条件2: L <= 10 && D >= 14
    条件3: L >= 10.4
    条件4: L == 10.2 && D <= 12
    条件5: else（兜底：L == 10.2 && D > 12 及中间状态）
    """
    D = diameter
    L = length

    if D <= 12 and L <= 10:
        # 条件1
        return 0.7854 * L * (D + 0.45 * L + 0.2) ** 2 / 10000
    elif L <= 10 and D >= 14:
        # 条件2
        return 0.7854 * L * (D + 0.5 * L + 0.005 * L ** 2 + 0.000125 * L * (14 - L) ** 2 * (D - 10)) ** 2 / 10000
    elif L >= 10.4:
        # 条件3
        return 0.8 * L * (D + 0.5 * L) ** 2 / 10000
    elif L == 10.2 and D <= 12:
        # 条件4
        return (0.7854 * 10 * (D + 4.7) ** 2 / 10000) / 2 + (0.8 * 10.4 * (D + 5.2) ** 2 / 10000 / 2)
    else:
        # 条件5：兜底（L == 10.2 && D > 12，以及 D=13 且 L<=10 等中间态）
        return 0.8 * 10.4 * (D + 5.2) ** 2 / 20000 + 7.854 * (D + 5.5 + 0.02 * (D - 10)) ** 2 / 20000


def get_all_files():
    """返回合并后的文件列表：优先取 file_registry，再补 code_sheets 中未被迁移的旧文件"""
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('SELECT id, file_name, row_count, upload_time FROM file_registry ORDER BY upload_time DESC')
    registry_files = [dict(row) for row in cursor.fetchall()]
    registry_names = {f['file_name'] for f in registry_files}

    # 补上还在旧 code_sheets 中但未迁移到 registry 的文件
    if registry_names:
        placeholders = ','.join(['?' for _ in registry_names])
        cursor.execute(f'''
            SELECT file_name, COUNT(*) as row_count, MAX(upload_time) as upload_time
            FROM code_sheets
            WHERE file_name NOT IN ({placeholders})
            GROUP BY file_name
            ORDER BY MAX(upload_time) DESC
        ''', tuple(registry_names))
    else:
        cursor.execute('''
            SELECT file_name, COUNT(*) as row_count, MAX(upload_time) as upload_time
            FROM code_sheets
            GROUP BY file_name
            ORDER BY MAX(upload_time) DESC
        ''')
    legacy_files = [dict(row) for row in cursor.fetchall()]

    conn.close()
    return registry_files + legacy_files


def resolve_table(cursor, file_name):
    """
    根据 file_name 查找对应的数据表名。
    先查 file_registry（独立表），找不到则假定在 code_sheets 中。
    返回 (table_name, in_registry) 或 (None, False)。
    """
    cursor.execute('SELECT table_name FROM file_registry WHERE file_name = ?', (file_name,))
    row = cursor.fetchone()
    if row:
        return row['table_name'], True
    return None, False
