"""
auth_service.py

【概要】
「ログインできるか」「パスワードを変えていいか」を判断する専用モジュール。
画面側（app.py・pages/配下）は、このモジュールの関数を呼ぶだけで済み、
パスワードをどう暗号化するか・DBのどのテーブルを見るか、といった
具体的な処理には関与しない（画面とロジックを分離する設計方針）。

【このモジュールが担う3つの役割】
1. パスワードの暗号化・検証（bcryptライブラリを利用）
   平文パスワードはDBに一切保存せず、常にハッシュ化した値のみ保存する。
   ハッシュ化は不可逆（元のパスワードには戻せない）な変換のため、
   万が一DBが漏洩してもパスワードそのものは判明しない設計。

2. ログイン認証（authenticate）
   社員番号（またはadmin固定ID）とパスワードを受け取り、
   DBに保存されているハッシュと一致するか確認する。
   一致すればUser情報を返し、一致しなければNoneを返す。
   「IDが存在しない」のか「パスワードが違う」のかは区別して返さない
   （区別すると社員番号の存在有無が推測できてしまうため、
   セキュリティ上あえて同じ結果にしている）。

3. パスワード変更・再設定
   - change_password：本人が現在のパスワードを覚えている状態での変更
   - reset_password_by_admin：本人がパスワードを忘れた際、adminが強制的に
     新しいパスワードを設定する（要件定義書5.4節のフローに対応）
   いずれの場合も、新しいパスワードが「8文字以上」「英大文字・英小文字・
   数字のうち2種類以上」等のポリシーを満たしているかを保存前に必ず確認する。

基本設計書 5.2.1節「auth_service」、10章「認証・パスワード設計」に対応。
"""

import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional

import bcrypt

from adapters import db_adapter

# ============================================
# パスワードポリシー（基本設計書 10.3節）
# ============================================
PASSWORD_MIN_LENGTH = 8
PASSWORD_MAX_LENGTH = 64
_ALLOWED_CHARS_PATTERN = re.compile(r"[A-Za-z0-9!#$%&*+\-./:=?@^_]+")


@dataclass
class User:
    """基本設計書 5.3節 DTOイメージに対応"""
    employee_id: str
    last_name: str
    first_name: str
    department: str
    role: str  # "admin" | "general"

    @property
    def full_name(self) -> str:
        """姓名の間にスペースを入れない結合（ファイル名等での用途向け）"""
        return f"{self.last_name}{self.first_name}"

    @property
    def full_name_with_space(self) -> str:
        """姓名の間に全角スペースを入れる結合（Excel AH5セル等での用途向け）"""
        return f"{self.last_name}　{self.first_name}"


def _row_to_user(row) -> User:
    return User(
        employee_id=row["employee_id"],
        last_name=row["last_name"],
        first_name=row["first_name"],
        department=row["department"],
        role=row["role"],
    )


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


# ============================================
# パスワードハッシュ化（bcrypt、基本設計書 10.2節）
# ============================================

def hash_password(plain_password: str) -> str:
    """平文パスワードをbcryptでハッシュ化する"""
    hashed = bcrypt.hashpw(plain_password.encode("utf-8"), bcrypt.gensalt())
    return hashed.decode("utf-8")


def verify_password(plain_password: str, password_hash: str) -> bool:
    """平文パスワードとハッシュを比較検証する"""
    try:
        return bcrypt.checkpw(
            plain_password.encode("utf-8"), password_hash.encode("utf-8")
        )
    except (ValueError, TypeError):
        # 不正なハッシュ形式等の場合は認証失敗として扱う
        return False


# ============================================
# パスワードポリシー検証（基本設計書 10.3節）
# ============================================

def validate_password_policy(password: str) -> list[str]:
    """
    パスワードポリシーに違反している場合、エラーメッセージのリストを返す。
    違反がなければ空リストを返す。
    """
    errors: list[str] = []

    if len(password) < PASSWORD_MIN_LENGTH:
        errors.append(f"{PASSWORD_MIN_LENGTH}文字以上で入力してください")
    if len(password) > PASSWORD_MAX_LENGTH:
        errors.append(f"{PASSWORD_MAX_LENGTH}文字以内で入力してください")
    if not _ALLOWED_CHARS_PATTERN.fullmatch(password or ""):
        errors.append("使用できない文字が含まれています")

    types_used = sum(
        [
            bool(re.search(r"[A-Z]", password)),
            bool(re.search(r"[a-z]", password)),
            bool(re.search(r"[0-9]", password)),
        ]
    )
    if types_used < 2:
        errors.append(
            "英大文字・英小文字・数字のうち2種類以上を組み合わせてください"
        )

    return errors


# ============================================
# 認証（基本設計書 5.2.1節・10.1節）
# ============================================

def authenticate(employee_id: str, password: str) -> Optional[User]:
    """
    社員番号（またはadmin固定ID）とパスワードで認証する。
    成功時はUserを、失敗時はNoneを返す。
    ID存在有無を区別するメッセージは呼び出し側で表示しないこと
    （SC-01「IDまたはパスワードが正しくありません」に統一）。
    """
    if not employee_id or not password:
        return None

    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT * FROM users WHERE employee_id = ?",
            (employee_id,),
        )
        row = cur.fetchone()

    if row is None:
        return None

    if not verify_password(password, row["password_hash"]):
        return None

    return _row_to_user(row)


# ============================================
# パスワード変更（本人による変更、基本設計書 5.2.1節）
# ============================================

def change_password(
    employee_id: str, current_password: str, new_password: str
) -> tuple[bool, Optional[str]]:
    """
    本人によるパスワード変更。現在のパスワードの確認を行う。
    戻り値: (成功したか, エラーメッセージ or None)
    """
    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT password_hash FROM users WHERE employee_id = ?",
            (employee_id,),
        )
        row = cur.fetchone()

    if row is None:
        return False, "ユーザーが見つかりません"

    if not verify_password(current_password, row["password_hash"]):
        return False, "現在のパスワードが正しくありません"

    policy_errors = validate_password_policy(new_password)
    if policy_errors:
        return False, "、".join(policy_errors)

    new_hash = hash_password(new_password)
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            "UPDATE users SET password_hash = ?, updated_at = ? WHERE employee_id = ?",
            (new_hash, _now_iso(), employee_id),
        )

    return True, None


# ============================================
# パスワード再設定（admin用、基本設計書 5.2.1節・10.4節）
# ============================================

def reset_password_by_admin(employee_id: str, new_password: str) -> tuple[bool, Optional[str]]:
    """
    adminによる強制パスワード再設定。現パスワードの確認は行わない。
    「パスワードを忘れた場合の対応」フロー（要件定義書5.4節）に対応。
    戻り値: (成功したか, エラーメッセージ or None)
    """
    policy_errors = validate_password_policy(new_password)
    if policy_errors:
        return False, "、".join(policy_errors)

    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT employee_id FROM users WHERE employee_id = ?",
            (employee_id,),
        )
        if cur.fetchone() is None:
            return False, "ユーザーが見つかりません"

    new_hash = hash_password(new_password)
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            "UPDATE users SET password_hash = ?, updated_at = ? WHERE employee_id = ?",
            (new_hash, _now_iso(), employee_id),
        )

    return True, None


def reset_admin_password_by_recovery(
    recovery_password: str,
    new_password: str,
    configured_recovery_password: str,
) -> tuple[bool, Optional[str]]:
    """
    環境変数 ADMIN_RESET_PASSWORD を利用して
    adminアカウントのパスワードを強制的に再設定する。

    admin本人が現在のパスワードを忘れてログインできない場合に使用する。

    戻り値:
        (True, None)       : 成功
        (False, エラー文) : 失敗
    """

    # 復旧用パスワードが未設定の場合
    if not configured_recovery_password:
        return False, "管理者パスワード復旧機能が設定されていません"

    # 復旧用パスワードを確認
    if recovery_password != configured_recovery_password:
        return False, "復旧用パスワードが正しくありません"

    # 新しいパスワードのポリシー確認
    policy_errors = validate_password_policy(new_password)
    if policy_errors:
        return False, "、".join(policy_errors)

    # adminアカウントの存在確認
    with db_adapter.get_cursor() as cur:
        cur.execute(
            "SELECT employee_id FROM users WHERE employee_id = ?",
            ("admin",),
        )
        if cur.fetchone() is None:
            return False, "adminアカウントが見つかりません"

    # 新しいパスワードをbcryptでハッシュ化
    new_hash = hash_password(new_password)

    # adminのパスワードを更新
    with db_adapter.get_cursor(commit=True) as cur:
        cur.execute(
            """
            UPDATE users
            SET password_hash = ?, updated_at = ?
            WHERE employee_id = ?
            """,
            (new_hash, _now_iso(), "admin"),
        )

    return True, None