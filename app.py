"""
app.py

【概要】
Streamlitアプリのエントリポイント。次の2つの役割を持つ。

1. アプリ初回起動時の初期化（DBスキーマ作成・adminアカウント投入）
2. ログイン画面の表示、および ログイン後の画面遷移の制御

【画面遷移とセッションの関係】
st.navigation() / st.Page() というStreamlit標準のページ管理機能を使い、
ログイン中ユーザーのroleに応じて「サイドバーに表示するページの一覧」を
動的に切り替えている（3.3.1節「画面アクセス制御」の実装）。

st.session_state はブラウザのWebSocket接続が続く限りサーバー側に
保持されるため、st.navigation経由でページを移動しても
ログイン情報（session_service参照）は消えない。
ページを移動するたびに再ログインを求められることはない。

【権限によるサイドバー出し分けの考え方】
- 一般ユーザー：勤怠入力画面／パスワード変更のみを一覧に含める
- admin       ：勤怠閲覧・編集／PDF出力／ロック管理／ユーザー管理／
                パスワード変更のみを一覧に含める（勤怠入力は含めない。
                adminは自分の勤怠データを持たないため）
この一覧自体に含まれないページはサイドバーのリンクとして出てこないため、
「一般ユーザーにPDF出力のリンクが見える」という状態は発生しない。
ただし、URLを直接指定するような迂回アクセスに備え、
各ページ側でも session_service.require_admin() 等の権限チェックを行う
（メニュー非表示だけに頼らない、二重の制御。3.3.1節参照）。

【ログイン前にページ一覧を一切出さないための工夫（二重の対策）】
対策1：Streamlitは、アプリ直下に "pages" という名前のディレクトリが
あると、st.navigation()の呼び出しの有無に関わらず自動的にそれらを
マルチページとして検出し、サイドバーに一覧表示してしまう（未ログイン時
も表示されてしまう）。この挙動を避けるため、実ファイルは "_pages/"
（アンダースコア始まり）に置き、Streamlitの自動検出対象から外している。

対策2（保険）：それでも何らかの理由で"pages/"検出が働いてしまう
ケースに備え、ログイン画面の描画時点でも必ず st.navigation() を
（ログイン画面自体を1件のページとして、position="hidden"で）呼び出す。
公式ドキュメントにより「一度でもst.navigation()が呼ばれたセッションは
以後pages/ディレクトリを無視する」という仕様があるため、ログイン画面
の初回表示時点からこれを発火させておくことで、"pages/"自動検出が
働く隙を作らない。
"""

import os

import streamlit as st
from dotenv import load_dotenv

from adapters import db_adapter
from services import auth_service, session_service

load_dotenv()

st.set_page_config(page_title="勤怠管理システム", page_icon="🕒", layout="wide")


def _init_app() -> None:
    """
    アプリ初回起動時の初期化処理。
    何度呼び出しても副作用が出ないよう、各処理は「存在すればスキップ」という
    冪等な作り（db_adapter側で担保）にしてある。
    """
    db_adapter.init_db()

    initial_admin_password = os.environ.get("INITIAL_ADMIN_PASSWORD", "ChangeMe123")
    admin_hash = auth_service.hash_password(initial_admin_password)
    db_adapter.seed_initial_admin(admin_hash)

    service_start_month = os.environ.get("SERVICE_START_MONTH", "")
    if service_start_month:
        db_adapter.seed_setting_if_absent("service_start_month", service_start_month)


def _render_login_page() -> None:
    """
    ログイン画面（SC-01）の中身の描画。
    実際の呼び出しは _run_login_navigation() 経由で行う（下記参照）。
    """
    st.title("勤怠管理システム")
    st.subheader("ログイン")

    with st.form("login_form"):
        employee_id = st.text_input("ログインID（社員番号 / admin）")
        password = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン")

    if submitted:
        user = auth_service.authenticate(employee_id, password)
        if user is None:
            # IDの存在有無を区別しないメッセージ（基本設計書3.4節）
            st.error("IDまたはパスワードが正しくありません")
        else:
            session_service.start_session(
                employee_id=user.employee_id,
                last_name=user.last_name,
                first_name=user.first_name,
                department=user.department,
                role=user.role,
            )
            st.rerun()


def _run_login_navigation() -> None:
    """
    未ログイン時のページ実行。

    【pages/自動検出を確実に無効化するための処理】
    Streamlitは、st.navigation()が一度も呼ばれていない状態だと、
    アプリ直下に "pages" という名前のディレクトリがあった場合に
    自動でそれをサイドバーへ表示してしまう（本アプリは"_pages/"という
    別名にしているため通常は該当しないはずだが、環境によっては
    キャッシュ等の影響で残ってしまうことがある）。

    st.navigation() は「1件以上のページ」を渡す必要があり、空リストは
    エラーになるため、ログイン画面自体を st.Page として1件だけ登録し、
    position="hidden" でサイドバーへの表示を抑止したうえで実行する。
    これにより、ログイン画面の描画時点でも必ず st.navigation() が
    呼ばれるようになり、"pages/"自動検出ロジックは起動しない。
    """
    login_page = st.Page(_render_login_page, title="ログイン")
    nav = st.navigation([login_page], position="hidden")
    nav.run()


def _render_sidebar_user_info() -> None:
    """
    サイドバー上部に、ログイン中ユーザー情報とログアウトボタンを表示する。

    adminアカウントは氏名を持たない特殊アカウント（基本設計書4.3.1節、
    姓・名ともに"管理者"として初期投入される）のため、display_name
    をそのまま表示すると「管理者 管理者（admin）」のように文字列が
    冗長になる。そのため、adminの場合は氏名を出さず「管理者」という
    ラベルのみを表示し、一般ユーザーの場合のみ氏名・社員番号を表示する。
    """
    user = session_service.get_current_user()
    if user is None:
        return
    with st.sidebar:
        if user.is_admin:
            st.write("**管理者**")
        else:
            st.write(f"**{user.display_name}**（{user.employee_id}）")
            st.caption("一般ユーザー")
        if st.button("ログアウト", use_container_width=True):
            session_service.end_session()
            st.rerun()
        st.divider()


def _build_navigation():
    user = session_service.get_current_user()

    password_page = st.Page(
        "_pages/06_password_change.py",
        title="パスワード変更",
        icon="🔑",
    )

    if user is not None and user.is_admin:
        pages = {
            "業務メニュー": [
                st.Page(
                    "_pages/02_admin_attendance_edit.py",
                    title="勤怠閲覧・編集",
                    icon="📋",
                    default=True,
                ),
                st.Page(
                    "_pages/03_admin_pdf_export.py",
                    title="PDF出力",
                    icon="🖨️",
                ),
                st.Page(
                    "_pages/04_admin_lock_management.py",
                    title="ロック管理",
                    icon="🔒",
                ),
                st.Page(
                    "_pages/05_admin_user_management.py",
                    title="ユーザー管理",
                    icon="👤",
                ),
            ],
            "アカウント": [password_page],
        }
    else:
        pages = {
            "業務メニュー": [
                st.Page(
                    "_pages/01_attendance_input.py",
                    title="勤怠入力・閲覧",
                    icon="📝",
                    default=True,
                ),
            ],
            "アカウント": [password_page],
        }

    return st.navigation(pages, position="sidebar")


def main() -> None:
    """
    アプリのエントリーポイント制御。ログイン前・ログイン後いずれの場合も
    必ず st.navigation() を経由してページを実行する（詳細は本ファイル
    冒頭のモジュールdocstring「pages/自動検出を確実に無効化するための
    工夫」を参照）。
    """
    _init_app()

    if not session_service.is_logged_in():
        # ログイン前はサイドバーを一切描画しない。
        # _run_login_navigation() が内部で position="hidden" の
        # st.navigation() を呼ぶため、_pages/配下のページ一覧は
        # サイドバーに表示されず、"pages/"自動検出も発火しない。
        _run_login_navigation()
        return

    _render_sidebar_user_info()
    nav = _build_navigation()
    nav.run()


if __name__ == "__main__":
    main()