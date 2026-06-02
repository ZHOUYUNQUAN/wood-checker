"""
木材检尺对比系统 - 检索与计算路由
处理编号搜索、材积计算等核心功能
"""
import logging

from flask import render_template, request, jsonify

from models import get_db, _validate_table_name
from . import check_bp
from .utils import (
    get_all_files,
    resolve_table,
    calc_external_standard,
    calc_national_standard,
)

logger = logging.getLogger(__name__)


@check_bp.route('/')
def index():
    """主页：渲染 checker 页面，附带已上传文件列表"""
    files = get_all_files()
    return render_template('check/checker.html', files=files, PREFIX='/check')


@check_bp.route('/search')
def search():
    """模糊搜索 no 字段，返回匹配记录。file_name 指向独立表或旧表。"""
    q = request.args.get('q', '').strip()
    file_name = request.args.get('file_name', '').strip()

    if not q:
        return jsonify({'ok': False, 'error': '搜索关键词不能为空'}), 400

    conn = get_db()
    cursor = conn.cursor()

    if file_name:
        table_name, in_registry = resolve_table(cursor, file_name)
        if in_registry and table_name:
            _validate_table_name(table_name)
            cursor.execute(f'SELECT * FROM "{table_name}" WHERE no LIKE ? ORDER BY no LIMIT 200', (f'%{q}%',))
        else:
            cursor.execute('''
                SELECT * FROM code_sheets
                WHERE no LIKE ? AND file_name = ?
                ORDER BY no
                LIMIT 200
            ''', (f'%{q}%', file_name))
    else:
        cursor.execute('''
            SELECT * FROM code_sheets
            WHERE no LIKE ?
            ORDER BY no
            LIMIT 200
        ''', (f'%{q}%',))

    records = [dict(row) for row in cursor.fetchall()]
    conn.close()

    return jsonify({'ok': True, 'records': records, 'count': len(records)})


@check_bp.route('/calc', methods=['POST'])
def calc():
    """计算新材积（国标或外标）。file_name 用于定位独立数据表。"""
    data = request.get_json()
    if not data:
        return jsonify({'ok': False, 'error': '请求数据为空'}), 400

    record_id = data.get('record_id')
    file_name = data.get('file_name', '').strip()
    standard = data.get('standard', '').strip()
    length = data.get('length')
    diameter = data.get('diameter')

    if not record_id:
        return jsonify({'ok': False, 'error': '缺少 record_id'}), 400
    if standard not in ('national', 'external'):
        return jsonify({'ok': False, 'error': '标准参数无效，可选 national 或 external'}), 400

    conn = get_db()
    cursor = conn.cursor()

    if file_name:
        table_name, in_registry = resolve_table(cursor, file_name)
        if in_registry and table_name:
            _validate_table_name(table_name)
            cursor.execute(f'SELECT * FROM "{table_name}" WHERE id = ?', (record_id,))
        else:
            cursor.execute('SELECT * FROM code_sheets WHERE id = ?', (record_id,))
    else:
        cursor.execute('SELECT * FROM code_sheets WHERE id = ?', (record_id,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        return jsonify({'ok': False, 'error': '记录不存在'}), 404

    original_volume = row['volume_m3'] or 0.0

    # 使用传入的直径/长度，若未传则用数据库中的值
    if length is not None:
        try:
            length = float(length)
        except (ValueError, TypeError):
            length = row['length_m'] or 0.0
    else:
        length = row['length_m'] or 0.0

    if diameter is not None:
        try:
            diameter = float(diameter)
        except (ValueError, TypeError):
            diameter = row['diameter_avg'] or 0.0
    else:
        diameter = row['diameter_avg'] or 0.0

    # 根据标准选择公式计算
    if standard == 'external':
        new_volume = calc_external_standard(diameter, length)
    else:
        new_volume = calc_national_standard(diameter, length)

    new_volume = round(new_volume, 4)
    diff = round(new_volume - original_volume, 4)
    if original_volume != 0:
        rate = round((new_volume - original_volume) / original_volume * 100, 2)
    else:
        rate = 0.0

    return jsonify({
        'ok': True,
        'original_volume': original_volume,
        'new_volume': new_volume,
        'diff': diff,
        'rate': rate,
        'standard': standard,
        'diameter_used': diameter,
        'length_used': length,
    })
