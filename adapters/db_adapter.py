"""
db_adapter.py
SQLiteへの接続・クエリ実行を担う連携レイヤー。
基本設計書 5.1節「モジュール構成」に対応。
"""

import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path

# 環境変数 DB_PATH からDBファイルのパスを取得（基本設計書13.5節）
DB_PATH = os.environ.get("DB_PATH", "./data/app.db")

# schema.sql の場所（models/配下）
_SCHEMA_PATH = Path(__file__).resolve().parent.parent / "models" / "schema.sql"


def get_connection() -> sqlite3.Connection:
    """
    SQLite接続を取得する。
    Row工場をsqlite3.Rowに設定し、辞書的にカラムへアクセスできるようにする。
    """
    # DBファイルの親ディレクトリが存在しない場合は作成する
    db_dir = os.path.dirname(DB_PATH)
    if db_dir:
        os.makedirs(db_dir, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    # 外部キー制約を有効化（SQLiteはデフォルト無効のため）
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def get_cursor(commit: bool = False):
    """
    with文で使えるカーソルのコンテキストマネージャ。
    commit=True の場合、正常終了時にコミットする。
    例外発生時は自動的にロールバックする。

    使用例:
        with get_cursor(commit=True) as cur:
            cur.execute("INSERT INTO users (...) VALUES (...)", (...))
    """
    conn = get_connection()
    cur = conn.cursor()
    try:
        yield cur
        if commit:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    """
    schema.sql を実行し、テーブルが存在しなければ作成する。
    アプリ起動時（app.py）に呼び出す想定。
    """
    if not _SCHEMA_PATH.exists():
        raise FileNotFoundError(f"schema.sql が見つかりません: {_SCHEMA_PATH}")

    schema_sql = _SCHEMA_PATH.read_text(encoding="utf-8")

    conn = get_connection()
    try:
        conn.executescript(schema_sql)
        conn.commit()
    finally:
        conn.close()


def seed_initial_admin(admin_password_hash: str) -> None:
    """
    初期構築時、adminアカウント（employee_id='admin'）が
    存在しない場合のみ1件投入する。
    基本設計書 4.3.1節「初期構築時にDBへ直接1件だけ投入」に対応。

    複数回呼び出されても重複投入されないよう、
    INSERT OR IGNORE + PRIMARY KEY制約で冪等にしている。
    """
    from datetime import datetime, timezone

    now = datetime.now(timezone.utc).isoformat()

    with get_cursor(commit=True) as cur:
        cur.execute(
            """
            INSERT OR IGNORE INTO users
                (employee_id, last_name, first_name, department,
                 password_hash, role, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "admin",
                "管理者",
                "管理者",
                "-",
                admin_password_hash,
                "admin",
                now,
                now,
            ),
        )


def seed_setting_if_absent(key: str, value: str) -> None:
    """
    settings テーブルに指定キーが存在しない場合のみ初期値を投入する。
    基本設計書 4.3.3節「運用開始月は初回起動時のみ使用」に対応。
    """
    from datetime import datetime, timezone

    now = datetime.now(timezone.utc).isoformat()

    with get_cursor(commit=True) as cur:
        cur.execute(
            """
            INSERT OR IGNORE INTO settings (key, value, updated_at)
            VALUES (?, ?, ?)
            """,
            (key, value, now),
        )