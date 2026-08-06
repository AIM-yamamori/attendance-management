"""
user_service.py

【概要】
ユーザーマスタ（users テーブル）に対するCRUD操作を担うモジュール。
adminがSC-06（ユーザーマスタ管理画面）から一般ユーザーを追加・編集・削除する
際に使う。auth_service がログイン・パスワードに特化しているのに対し、
このモジュールは「誰がどの部署に所属しているか」といった属性情報の
管理に専念する（責務を分けている）。

このモジュールから新規作成できるのは一般ユーザー（role="general"）のみ。
adminアカウントは初期構築時に1件だけ投入する運用のため、
本モジュールにadminを新規作成する関数は用意しない
（基本設計書3.9節「本画面から権限区分adminのユーザーを新規作成する
機能は設けない」に対応）。
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional

from adapters import db_adapter
from services import auth_service


@dataclass
class User:
    employee_id: str
    last_name: str
    first_name: str
    department: str
    role: str

    @property
    def display_name(self) -> str:
        return f"{self.last_name} {self.first_name}"


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _row_to_user(row) -> User:
    return User(
        employee_id=row["employee_id"],
        last_name=row["last_name"],
        first_name=row["first_name"],
        department=row["department"],
        role=row["role"],
    )


def list_users() -> list[User]:
    """一般ユーザー一覧を取得する（adminは一覧に含めない）。"""
    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT * FROM users WHERE role = 'general' ORDER BY employee_id"
        )
        rows = cur.fetchall()
    return [_row_to_user(r) for r in rows]


def get_user(employee_id: str) -> Optional[User]:
    """社員番号からユーザー情報を1件取得する。"""
    with db_adapter.get_cursor() as cur:
        cur.execute("SELECT * FROM users WHERE employee_id = ?", (employee_id,))
        row = cur.fetchone()
    return _row_to_user(row) if row else None


def _validate_employee_id(employee_id: str) -> Optional[str]:
    """
    社員番号の形式チェック。
    要件定義書5.4節：半角数字、桁数可変、0埋めなし。
    「admin」という値は予約済みのため一般ユーザーには使わせない。
    """
    if not employee_id:
        return "社員番号を入力してください"
    if employee_id == "admin":
        return "この社員番号は使用できません"
    if not employee_id.isdigit():
        return "社員番号は半角数字で入力してください"
    if employee_id != str(int(employee_id)):
        # 先頭ゼロ埋めを禁止（"0012"のような値を弾く）
        return "社員番号は0埋めなしで入力してください"
    return None


def create_user(
    last_name: str,
    first_name: str,
    employee_id: str,
    department: str,
    initial_password: str,
) -> tuple[bool, Optional[str]]:
    """
    一般ユーザーを新規作成する。
    戻り値: (成功したか, エラーメッセージ or None)
    """
    if not last_name or not first_name:
        return False, "姓・名を入力してください"
    if not department:
        return False, "部署名を入力してください"

    id_error = _validate_employee_id(employee_id)
    if id_error:
        return False, id_error

    policy_errors = auth_service.validate_password_policy(initial_password)
    if policy_errors:
        return False, "、".join(policy_errors)

    # 社員番号の重複チェック
    if get_user(employee_id) is not None:
        return False, "この社員番号は既に登録されています"

    now = _now_iso()
    password_hash = auth_service.hash_password(initial_password)

    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            INSERT INTO users
                (employee_id, last_name, first_name, department,
                 password_hash, role, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 'general', ?, ?)
            """,
            (employee_id, last_name, first_name, department, password_hash, now, now),
        )

    return True, None


def update_user(
    employee_id: str, last_name: str, first_name: str, department: str
) -> tuple[bool, Optional[str]]:
    """
    氏名・部署名を更新する（社員番号・パスワードはこの関数の対象外）。
    admin自身の情報更新には使わない想定（一般ユーザーのみを対象とする）。
    """
    user = get_user(employee_id)
    if user is None:
        return False, "ユーザーが見つかりません"
    if user.role == "admin":
        return False, "adminアカウントはこの画面から編集できません"

    if not last_name or not first_name:
        return False, "姓・名を入力してください"
    if not department:
        return False, "部署名を入力してください"

    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            UPDATE users
            SET last_name = ?, first_name = ?, department = ?, updated_at = ?
            WHERE employee_id = ?
            """,
            (last_name, first_name, department, _now_iso(), employee_id),
        )

    return True, None


def delete_user(employee_id: str) -> tuple[bool, Optional[str]]:
    """一般ユーザーを削除する。adminアカウントの削除は許可しない。"""
    user = get_user(employee_id)
    if user is None:
        return False, "ユーザーが見つかりません"
    if user.role == "admin":
        return False, "adminアカウントは削除できません"

    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute("DELETE FROM users WHERE employee_id = ?", (employee_id,))

    return True, None