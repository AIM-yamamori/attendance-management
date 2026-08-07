"""
lock_service.py

【概要】
ロック状態の取得・更新を担うモジュール（基本設計書5.2.3節）。

locksテーブルのレコードライフサイクル（基本設計書4.3.2節）：
- 該当月・該当ユーザーのlocksレコードが存在しない場合は「未ロック」
  とみなす（レコードなし＝未ロックがデフォルト）。
- 「まとめてロック」実行時：対象月の全ユーザーについて、レコードが
  存在しなければINSERT（is_locked=1）、存在すればUPDATE
  （is_locked=1, locked_at更新）。
- 「個別に解除」実行時：該当レコードをUPDATE（is_locked=0,
  unlocked_at更新）。レコード自体は削除しない。
"""

from dataclasses import dataclass
from datetime import datetime, timezone

from adapters import db_adapter


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


# ============================================
# データ型（基本設計書5.3節 LockStatus）
# ============================================

@dataclass
class LockStatus:
    employee_id: str
    last_name: str
    first_name: str
    is_locked: bool


# ============================================
# 5.2.3節 関数インターフェース
# ============================================

def is_locked(target_month: str, employee_id: str) -> bool:
    """
    ロック状態確認（3.5.4節・8.2節・8.3節の保存直前チェックで使用）。
    レコードが存在しない場合は未ロックとみなす（4.3.2節）。
    """
    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT is_locked FROM locks WHERE target_month = ? AND employee_id = ?",
            (target_month, employee_id),
        )
        row = cur.fetchone()

    if row is None:
        return False
    return bool(row["is_locked"])


def lock_all(target_month: str) -> int:
    """
    まとめてロック（SC-05）。
    対象月の全ユーザーについて、レコードが存在しなければINSERT
    （is_locked=1）、存在すればUPDATE（is_locked=1, locked_at更新）
    （基本設計書4.3.2節）。

    戻り値：対象件数（ロック対象としたユーザー数）
    """
    from services import user_service

    all_users = user_service.list_users()
    now = _now()

    with db_adapter.get_cursor(commit=True) as cur:
        for u in all_users:
            cur.execute(
                """
                INSERT INTO locks (target_month, employee_id, is_locked, locked_at, updated_at)
                VALUES (?, ?, 1, ?, ?)
                ON CONFLICT(target_month, employee_id) DO UPDATE SET
                    is_locked = 1,
                    locked_at = excluded.locked_at,
                    updated_at = excluded.updated_at
                """,
                (target_month, u.employee_id, now, now),
            )

    return len(all_users)


def lock_one(target_month: str, employee_id: str) -> bool:
    """個別ロック（SC-05の行単位「ロック」ボタン）。"""
    now = _now()
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            INSERT INTO locks (target_month, employee_id, is_locked, locked_at, updated_at)
            VALUES (?, ?, 1, ?, ?)
            ON CONFLICT(target_month, employee_id) DO UPDATE SET
                is_locked = 1,
                locked_at = excluded.locked_at,
                updated_at = excluded.updated_at
            """,
            (target_month, employee_id, now, now),
        )
    return True


def unlock_one(target_month: str, employee_id: str) -> bool:
    """
    個別解除（SC-05の行単位「解除」ボタン）。
    該当レコードをUPDATE（is_locked=0, unlocked_at更新）。
    レコード自体は削除しない（基本設計書4.3.2節）。
    """
    now = _now()
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            INSERT INTO locks (target_month, employee_id, is_locked, unlocked_at, updated_at)
            VALUES (?, ?, 0, ?, ?)
            ON CONFLICT(target_month, employee_id) DO UPDATE SET
                is_locked = 0,
                unlocked_at = excluded.unlocked_at,
                updated_at = excluded.updated_at
            """,
            (target_month, employee_id, now, now),
        )
    return True


def list_lock_status(target_month: str) -> list[LockStatus]:
    """
    月内全ユーザーのロック状態一覧（SC-05表示用）。
    locksレコードが存在しないユーザーは未ロックとして扱う。
    """
    from services import user_service

    all_users = user_service.list_users()

    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT employee_id, is_locked FROM locks WHERE target_month = ?",
            (target_month,),
        )
        rows = cur.fetchall()

    locked_map = {row["employee_id"]: bool(row["is_locked"]) for row in rows}

    return [
        LockStatus(
            employee_id=u.employee_id,
            last_name=u.last_name,
            first_name=u.first_name,
            is_locked=locked_map.get(u.employee_id, False),
        )
        for u in all_users
    ]