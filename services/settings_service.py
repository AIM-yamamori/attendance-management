"""
settings_service.py

【概要】
アプリDBの settings テーブル（基本設計書4.3.3節）から、システム全体で
1件だけ持つ設定値を読み書きするための小さなモジュール。
現時点で使うのは運用開始月（service_start_month）のみだが、
将来的に他の全体設定が増えても同じテーブル・同じ関数群を使い回せる
よう、キーと値の組み合わせを汎用的に扱う作りにしている。
"""

from datetime import datetime, timezone
from typing import Optional

from adapters import db_adapter

KEY_SERVICE_START_MONTH = "service_start_month"


def get_setting(key: str) -> Optional[str]:
    """指定キーの設定値を取得する。存在しなければNoneを返す。"""
    with db_adapter.get_cursor() as cur:
        cur.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cur.fetchone()
    return row["value"] if row else None


def set_setting(key: str, value: str) -> None:
    """指定キーの設定値を作成・更新する。"""
    now = datetime.now(timezone.utc).isoformat()
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            INSERT INTO settings (key, value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
            """,
            (key, value, now),
        )


def get_service_start_month() -> Optional[str]:
    """
    運用開始月（"YYYYMM"形式）を取得する。
    未設定の場合はNoneを返す（呼び出し側の
    attendance_service.build_selectable_months が、その場合は
    安全側に倒して当月のみを選択肢にする。基本設計書3.5.2節）。
    """
    return get_setting(KEY_SERVICE_START_MONTH)