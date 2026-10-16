"""
user_service.py

【概要】
ユーザーマスタCRUDを担うモジュール（基本設計書5.2.2節）。

【戻り値の方針】
create_user・update_user・delete_userはいずれも
tuple[bool, Optional[str]]（成功したか, エラーメッセージ or None）を返す。
呼び出し側（_pages/05_admin_user_management.py）は
"ok, error_message = user_service.xxx(...)" という形で受け取る前提のため、
3関数の戻り値シグネチャを統一している（以前はupdate_user・delete_userが
bool単体を返しており、呼び出し側でのアンパックに失敗していた）。
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


def update_user(
    employee_id: str, last_name: str, first_name: str, department: str
) -> tuple[bool, Optional[str]]:
    """
    氏名・部署名更新（パスワード・社員番号は対象外）。

    戻り値: (成功したか, エラーメッセージ or None)
    対象ユーザーが存在しない場合はエラーメッセージ付きでFalseを返す
    （UPDATE自体は対象0件でもSQLite上はエラーにならず成功扱いに
    見えてしまうため、事前にget_userで存在確認する）。
    """
    if get_user(employee_id) is None:
        return False, "ユーザーが見つかりません"

    now = datetime.now(timezone.utc).isoformat()
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            UPDATE users SET last_name = ?, first_name = ?, department = ?, updated_at = ?
            WHERE employee_id = ?
            """,
            (last_name, first_name, department, now, employee_id),
        )
    return True, None


def delete_user(employee_id: str) -> tuple[bool, Optional[str]]:
    """
    一般ユーザーを削除する。

    戻り値: (成功したか, エラーメッセージ or None)

    locksテーブルはemployee_idに外部キー制約（REFERENCES users(employee_id)）
    を持っているため、対象ユーザーのlocksレコードが残った状態でusersから
    削除しようとすると、SQLiteが外部キー制約違反（IntegrityError）を送出する。
    ユーザー削除時はロック履歴を残す必要がない（ユーザー自体が消えるため）
    ので、同一トランザクション内でlocksレコードを先に削除してから
    usersレコードを削除する。

    get_cursor(commit=True)の1つのwithブロック内で両方のDELETEを実行する
    ことで、1トランザクションとして扱われる（どちらかが失敗すれば
    両方ロールバックされ、locksだけ消えてusersが残るような中途半端な
    状態にはならない）。
    """
    if get_user(employee_id) is None:
        return False, "ユーザーが見つかりません"

    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute("DELETE FROM locks WHERE employee_id = ?", (employee_id,))
        cur.execute("DELETE FROM users WHERE employee_id = ?", (employee_id,))
    return True, None