"""
木材检尺对比系统 - 检索与计算路由
处理编号搜索、材积计算等核心功能
"""
import json
import logging
from datetime import datetime

from flask import render_template, request, jsonify, Response

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

    with get_db() as conn:
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
    diameters = data.get('diameters')  # 数组，如 [86, 88] 或 [86, 88, 90, 92]
    diameter = data.get('diameter')    # 兼容旧版单值

    if not record_id:
        return jsonify({'ok': False, 'error': '缺少 record_id'}), 400
    if standard not in ('national', 'external'):
        return jsonify({'ok': False, 'error': '标准参数无效，可选 national 或 external'}), 400

    with get_db() as conn:
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

    # 多直径取平均值（兼容旧版单值 diameter 参数）
    if diameters and isinstance(diameters, list) and len(diameters) > 0:
        try:
            vals = [float(d) for d in diameters]
            diameter = sum(vals) / len(vals)
            diameter = round(diameter, 1)
        except (ValueError, TypeError):
            diameter = row['diameter_avg'] or 0.0
    elif diameter is not None:
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

    # ---- 保存计算结果到 extra_json，同一根只保留最后一次 ----
    calc_result = {
        'standard': standard,
        'diameter_used': diameter,
        'length_used': length,
        'original_volume': original_volume,
        'new_volume': new_volume,
        'diff': diff,
        'rate': rate,
        'calc_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    existing_extra = {}
    try:
        raw = row['extra_json']
        if raw:
            existing_extra = json.loads(raw) if isinstance(raw, str) else raw
    except (json.JSONDecodeError, TypeError):
        existing_extra = {}

    existing_extra['calc_result'] = calc_result
    new_extra_json = json.dumps(existing_extra, ensure_ascii=False)

    with get_db() as conn:
        cursor = conn.cursor()
        if file_name:
            table_name, in_registry = resolve_table(cursor, file_name)
            if in_registry and table_name:
                _validate_table_name(table_name)
                cursor.execute(f'UPDATE "{table_name}" SET extra_json = ? WHERE id = ?',
                               (new_extra_json, record_id))
            else:
                cursor.execute('UPDATE code_sheets SET extra_json = ? WHERE id = ?',
                               (new_extra_json, record_id))
        else:
            cursor.execute('UPDATE code_sheets SET extra_json = ? WHERE id = ?',
                           (new_extra_json, record_id))
        conn.commit()

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


@check_bp.route('/export')
def export_csv():
    """导出当前文件中所有计算过的记录为 CSV"""
    import csv
    import io

    file_name = request.args.get('file_name', '').strip()
    if not file_name:
        return jsonify({'ok': False, 'error': '缺少 file_name'}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        table_name, in_registry = resolve_table(cursor, file_name)
        if in_registry and table_name:
            _validate_table_name(table_name)
            cursor.execute(f'SELECT * FROM "{table_name}"')
        else:
            cursor.execute('SELECT * FROM code_sheets WHERE file_name = ?', (file_name,))
        rows = cursor.fetchall()

    # 筛选有 calc_result 的记录
    output_rows = []
    for row in rows:
        extra = {}
        try:
            raw = row['extra_json']
            if raw:
                extra = json.loads(raw) if isinstance(raw, str) else raw
        except (json.JSONDecodeError, TypeError):
            pass
        cr = extra.get('calc_result')
        if not cr:
            continue
        output_rows.append({
            'no': row['no'] or '',
            'especie': row['especie'] or '',
            'english_code': row['english_code'] or '',
            'original_diameter': row['diameter_avg'] or '',
            'original_length': row['length_m'] or '',
            'original_volume': cr.get('original_volume', 0),
            'new_volume': cr.get('new_volume', 0),
            'diff': cr.get('diff', 0),
            'rate': cr.get('rate', 0),
            'standard': '国标' if cr.get('standard') == 'national' else '外标',
            'diameter_used': cr.get('diameter_used', ''),
            'length_used': cr.get('length_used', ''),
        })

    if not output_rows:
        return jsonify({'ok': False, 'error': '没有已计算过的记录'}), 404

    # 生成 CSV
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['编号', '材种', '英文代码', '原始直径(CM)', '原始长度(M)',
                      '原材积', '新材积', '涨尺量', '涨尺率(%)', '标准',
                      '计算直径(CM)', '计算长度(M)'])
    for r in output_rows:
        writer.writerow([
            r['no'], r['especie'], r['english_code'],
            r['original_diameter'], r['original_length'],
            r['original_volume'], r['new_volume'],
            r['diff'], r['rate'], r['standard'],
            r['diameter_used'], r['length_used'],
        ])

    csv_content = output.getvalue()
    output.close()

    from urllib.parse import quote
    safe_name = file_name.rsplit('.', 1)[0] if '.' in file_name else file_name
    filename = f'{safe_name}_计算结果.csv'
    return Response(
        csv_content,
        mimetype='text/csv; charset=utf-8-sig',
        headers={
            'Content-Disposition': f"attachment; filename*=UTF-8''{quote(filename)}",
        },
    )
