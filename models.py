import re
import sqlite3
from contextlib import contextmanager

from config import Config

# 允许的表名格式：code_sheets、sheets_YYYYMMDD_xxxxxxxx、idx_sheets_YYYYMMDD_xxxxxxxx_no
_VALID_TABLE_RE = re.compile(r'^(code_sheets|sheets_\d{8}_[0-9a-f]{8}|idx_sheets_\d{8}_[0-9a-f]{8}_no)$')


def _validate_table_name(table_name):
    """校验表名是否为系统生成的已知格式，防止 SQL 注入"""
    if not _VALID_TABLE_RE.match(table_name):
        raise ValueError(f'无效的表名: {table_name}')


def _has_column(cursor, table_name, column_name):
    _validate_table_name(table_name)
    cursor.execute(f'PRAGMA table_info("{table_name}")')
    return any(row['name'] == column_name for row in cursor.fetchall())


def _add_column_if_missing(cursor, table_name, column_name, column_sql):
    _validate_table_name(table_name)
    if not _has_column(cursor, table_name, column_name):
        cursor.execute(f'ALTER TABLE "{table_name}" ADD COLUMN {column_sql}')


@contextmanager
def get_db():
    """获取数据库连接（context manager，自动关闭连接）"""
    conn = sqlite3.connect(Config.DATABASE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
    finally:
        conn.close()


def init_db():
    """初始化数据库表（由 app.py 启动时显式调用）"""
    with get_db() as conn:
        cursor = conn.cursor()

        # 旧表（保留向后兼容，新上传不再使用）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS code_sheets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                upload_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
                no TEXT,
                especie TEXT,
                english_code TEXT,
                diameter_1 REAL,
                diameter_2 REAL,
                diameter_3 REAL,
                diameter_4 REAL,
                diameter_avg REAL,
                length_m REAL,
                volume_m3 REAL,
                customer TEXT,
                base_name TEXT DEFAULT '',
                extra_json TEXT DEFAULT '{}',
                is_transshipment INTEGER DEFAULT 0
            )
        ''')

        _add_column_if_missing(cursor, 'code_sheets', 'base_name', "base_name TEXT DEFAULT ''")
        _add_column_if_missing(cursor, 'code_sheets', 'extra_json', "extra_json TEXT DEFAULT '{}'")

        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_code_sheets_no
            ON code_sheets(no)
        ''')

        # 文件注册表：每条记录对应一个上传文件的元信息与独立数据表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS file_registry (
                id TEXT PRIMARY KEY,
                file_name TEXT NOT NULL,
                table_name TEXT NOT NULL UNIQUE,
                row_count INTEGER DEFAULT 0,
                upload_time TEXT NOT NULL
            )
        ''')

        cursor.execute('SELECT table_name FROM file_registry')
        for row in cursor.fetchall():
            table_name = row['table_name']
            _add_column_if_missing(cursor, table_name, 'base_name', "base_name TEXT DEFAULT ''")
            _add_column_if_missing(cursor, table_name, 'extra_json', "extra_json TEXT DEFAULT '{}'")

        conn.commit()
