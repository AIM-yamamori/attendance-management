"""
user_service.py

【概要】
ユーザーマスタCRUDを担うモジュール（基本設計書5.2.2節）。
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional

from adapters import db_adapter


@dataclass
class User:
    employee_id: str
    last_name: str
    first_name: str
    department: str
    role: str  # "admin" | "general"

    @property
    def display_name(self) -> str:
        return f"{self.last_name} {self.first_name}"

    @property
    def full_name_no_space(self) -> str:
        """ファイル名専用の結合ルール（スペースなし）。"""
        return f"{self.last_name}{self.first_name}"

    @property
    def full_name_with_space(self) -> str:
        """Excel AH5セル用（全角スペース区切り）。"""
        return f"{self.last_name}　{self.first_name}"


def _row_to_user(row) -> User:
    return User(
        employee_id=row["employee_id"],
        last_name=row["last_name"],
        first_name=row["first_name"],
        department=row["department"],
        role=row["role"],
    )


def list_users() -> list[User]:
    """一般ユーザー一覧取得（role='general'のみ。adminは含まない）。"""
    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT employee_id, last_name, first_name, department, role "
            "FROM users WHERE role = 'general' ORDER BY employee_id"
        )
        rows = cur.fetchall()
    return [_row_to_user(r) for r in rows]


def get_user(employee_id: str) -> Optional[User]:
    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT employee_id, last_name, first_name, department, role "
            "FROM users WHERE employee_id = ?",
            (employee_id,),
        )
        row = cur.fetchone()
    return _row_to_user(row) if row else None


def create_user(
    last_name: str, first_name: str, employee_id: str, department: str, initial_password: str
) -> tuple[bool, Optional[str]]:
    """新規追加（社員番号重複チェック含む）。role='general'固定（3.9節）。"""
    from services import auth_service

    if get_user(employee_id) is not None:
        return False, "この社員番号は既に登録されています"

    policy_errors = auth_service.validate_password_policy(initial_password)
    if policy_errors:
        return False, "、".join(policy_errors)

    password_hash = auth_service.hash_password(initial_password)
    now = datetime.now(timezone.utc).isoformat()

    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            INSERT INTO users
                (employee_id, last_name, first_name, department, password_hash, role, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 'general', ?, ?)
            """,
            (employee_id, last_name, first_name, department, password_hash, now, now),
        )
    return True, None


def update_user(employee_id: str, last_name: str, first_name: str, department: str) -> bool:
    """氏名・部署名更新（パスワード・社員番号は対象外）。"""
    now = datetime.now(timezone.utc).isoformat()
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            UPDATE users SET last_name = ?, first_name = ?, department = ?, updated_at = ?
            WHERE employee_id = ?
            """,
            (last_name, first_name, department, now, employee_id),
        )
    return True


def delete_user(employee_id: str) -> bool:
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute("DELETE FROM users WHERE employee_id = ?", (employee_id,))
    return True