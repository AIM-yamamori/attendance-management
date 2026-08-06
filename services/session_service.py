"""
session_service.py

【概要】
「今どのユーザーがログイン中か」をアプリ全体で一貫して参照・更新するための
モジュール。ログイン状態はStreamlitの st.session_state に保持されるが、
st.session_state のキー名を各画面ファイルがバラバラに直接触ると、
キー名のtypoや管理漏れが起きやすい。そのため、st.session_state への
読み書きをこのモジュールに集約し、他のコードは
「is_logged_in()」「get_current_user()」のような意味の分かる関数越しに
アクセスする方針とする。

【セッションが画面遷移で切れない仕組み】
StreamlitはブラウザのタブがWebSocket接続を維持している間、
サーバー側で st.session_state を保持し続ける。
st.navigation / st.Page（app.py参照）を使ってページを切り替えても、
同一のWebSocket接続・同一の st.session_state が使い回されるため、
「勤怠入力画面 → PDF出力画面」のようにページを移動しても
ログイン情報は保持されたままになる。
（ページ遷移のたびに再ログインを求められることはない）

ログイン状態が切れる主なケース：
- ユーザーが明示的に「ログアウト」ボタンを押した場合
- ブラウザのタブを閉じてWebSocket接続が切れた場合
- サーバー側アプリケーションが再起動した場合（st.session_stateはメモリ上のみ）
"""

from dataclasses import dataclass
from typing import Optional

import streamlit as st

# st.session_state で使うキー名をここに集約する
# （他のファイルで直接 "employee_id" のような文字列を書かないようにするため）
_KEY_EMPLOYEE_ID = "employee_id"
_KEY_LAST_NAME = "last_name"
_KEY_FIRST_NAME = "first_name"
_KEY_ROLE = "role"
_KEY_LOGGED_IN = "logged_in"


@dataclass
class SessionUser:
    """ログイン中ユーザーの情報（セッションから読み出した結果を表す入れ物）"""
    employee_id: str
    last_name: str
    first_name: str
    role: str  # "admin" | "general"

    @property
    def display_name(self) -> str:
        return f"{self.last_name} {self.first_name}"

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"


def start_session(employee_id: str, last_name: str, first_name: str, role: str) -> None:
    """
    ログイン成功時に呼び出し、セッションにユーザー情報を保存する。
    以後、ページを移動してもこの情報は保持され続ける。
    """
    st.session_state[_KEY_EMPLOYEE_ID] = employee_id
    st.session_state[_KEY_LAST_NAME] = last_name
    st.session_state[_KEY_FIRST_NAME] = first_name
    st.session_state[_KEY_ROLE] = role
    st.session_state[_KEY_LOGGED_IN] = True


def end_session() -> None:
    """
    ログアウト時に呼び出し、セッション情報を破棄する。

    【なぜログイン関連キーだけでなく st.session_state 全体をクリアするか】
    以前はログイン関連の5キーだけを pop していたが、それだと
    st.navigation() が内部的に保持している「直前に選択していた
    ページ」等の状態が st.session_state 上に残ったままになることがあり、
    ログアウト直後の st.rerun() で一瞬 admin/一般ユーザー用のサイドバーの
    残像が表示されてしまう不具合があった。
    st.session_state は st.session_state.clear() で丸ごと空にできるため、
    ログアウト時は「ログイン関連キーだけ消す」のではなく「セッション全体を
    まっさらにする」方式に変更し、st.navigationの内部状態も含めて
    確実にリセットする。
    """
    st.session_state.clear()


def is_logged_in() -> bool:
    """ログイン中かどうかを返す。"""
    return bool(st.session_state.get(_KEY_LOGGED_IN, False))


def get_current_user() -> Optional[SessionUser]:
    """
    ログイン中ユーザーの情報を返す。未ログインの場合はNoneを返す。
    各画面（pages/配下）は、まずこの関数でユーザー情報を取得し、
    Noneであればログイン画面へ誘導する、という使い方を統一する。
    """
    if not is_logged_in():
        return None

    return SessionUser(
        employee_id=st.session_state.get(_KEY_EMPLOYEE_ID, ""),
        last_name=st.session_state.get(_KEY_LAST_NAME, ""),
        first_name=st.session_state.get(_KEY_FIRST_NAME, ""),
        role=st.session_state.get(_KEY_ROLE, "general"),
    )


def require_login() -> SessionUser:
    """
    ログイン必須の画面の冒頭で呼び出す。
    未ログインの場合はメッセージを表示して処理を中断する（st.stop）。
    基本設計書3.3.1節「ページ側での権限チェック」の実装。
    """
    user = get_current_user()
    if user is None:
        st.error("ログインが必要です。ログイン画面からやり直してください。")
        st.stop()
    return user


def require_admin() -> SessionUser:
    """
    admin専用画面の冒頭で呼び出す。
    ログインしていない、またはadminでない場合は処理を中断する。
    基本設計書3.3.1節「一般ユーザーにはadmin専用画面へのメニュー項目自体を
    表示しない」に加え、URL直接指定等の迂回アクセスに備えた二重チェック。
    """
    user = require_login()
    if not user.is_admin:
        st.error("このページを表示する権限がありません。")
        st.stop()
    return user


def require_general_user() -> SessionUser:
    """
    一般ユーザー専用画面（勤怠入力・閲覧画面）の冒頭で呼び出す。
    adminは自身の勤怠を持たない（要件定義書3.2節）ため、
    adminがこの画面へ迷い込んだ場合は案内して処理を中断する。
    """
    user = require_login()
    if user.is_admin:
        st.info(
            "adminアカウントは勤怠データを持ちません。"
            "「勤怠閲覧・編集」画面をご利用ください。"
        )
        st.stop()
    return user