"""
03_admin_PDF出力.py

【概要】
管理者用のPDF出力画面（SC-04）。要件定義書7.6節・4.7節、
基本設計書3.7節・9章に基づく。

- 対象月（単一選択）、対象ユーザー（チェックボックスによる複数選択）
  を指定してPDF生成。
- 「すべて選択」チェックボックスで一括選択・一括解除が可能。
- 1名選択時：PDF直接ダウンロード
- 複数名選択時：ZIP生成後、ZIPダウンロード
- OneDrive・サーバーいずれにも保存しない（都度生成・都度破棄）。
"""

import datetime

import streamlit as st

from services import attendance_service, pdf_service, session_service, settings_service, user_service

admin_user = session_service.require_admin()

st.title("PDF出力")

all_users = user_service.list_users()

if not all_users:
    st.info("対象ユーザーが登録されていません。")
    st.stop()

# ============================================
# 対象年月の選択（単一）
# ============================================
service_start_month = settings_service.get_service_start_month()
today = datetime.date.today()
selectable_months = attendance_service.build_selectable_months(
    today=today, service_start_month=service_start_month
)
default_month = attendance_service.default_target_month(today)
default_index = (
    selectable_months.index(default_month)
    if default_month in selectable_months
    else len(selectable_months) - 1
)

target_month = st.selectbox(
    "対象年月",
    options=selectable_months,
    index=default_index,
    format_func=attendance_service.format_month_label,
)

# 対象年月を切り替えたら選択状態をリセットする（月ごとに選び直す想定）
if st.session_state.get("_pdf_target_month") != target_month:
    for u in all_users:
        st.session_state[f"pdf_check_{u.employee_id}"] = False
    st.session_state["pdf_check_all"] = False
    st.session_state["_pdf_target_month"] = target_month
    st.session_state["pdf_output_bytes"] = None

st.divider()

# ============================================
# 対象ユーザーの選択（チェックボックス一覧＋すべて選択）
# ============================================
st.write("対象ユーザー")


def _on_check_all_changed():
    """「すべて選択」の変更を、各ユーザーのチェック状態へ反映する。"""
    check_all = st.session_state["pdf_check_all"]
    for u in all_users:
        st.session_state[f"pdf_check_{u.employee_id}"] = check_all


def _sync_check_all_state():
    """
    各ユーザーの個別チェック状態から、「すべて選択」の見た目を
    整合させる（全員チェック済みならON、1人でも外れていればOFF）。
    """
    all_checked = all(
        st.session_state.get(f"pdf_check_{u.employee_id}", False) for u in all_users
    )
    st.session_state["pdf_check_all"] = all_checked


st.checkbox(
    "すべて選択",
    key="pdf_check_all",
    on_change=_on_check_all_changed,
)

st.divider()

for u in all_users:
    key = f"pdf_check_{u.employee_id}"
    if key not in st.session_state:
        st.session_state[key] = False

    st.checkbox(
        f"{u.last_name} {u.first_name}（{u.employee_id}）",
        key=key,
        on_change=_sync_check_all_state,
    )

selected_employee_ids = [
    u.employee_id for u in all_users if st.session_state.get(f"pdf_check_{u.employee_id}", False)
]

st.divider()

# ============================================
# PDF生成
# ============================================
generate_clicked = st.button(
    "PDF生成",
    type="primary",
    disabled=not selected_employee_ids,
)

if not selected_employee_ids:
    st.caption("対象ユーザーを1名以上選択してください。")

if generate_clicked:
    try:
        with st.spinner("PDFを生成しています..."):
            if len(selected_employee_ids) == 1:
                employee_id = selected_employee_ids[0]
                pdf_bytes = pdf_service.generate_pdf(employee_id, target_month)
                file_name = pdf_service.build_pdf_file_name(employee_id, target_month)

                st.session_state["pdf_output_bytes"] = pdf_bytes
                st.session_state["pdf_output_file_name"] = file_name
                st.session_state["pdf_output_mime"] = "application/pdf"
            else:
                zip_bytes = pdf_service.generate_zip(selected_employee_ids, target_month)
                zip_file_name = f"勤務実績管理表_{target_month}.zip"

                st.session_state["pdf_output_bytes"] = zip_bytes
                st.session_state["pdf_output_file_name"] = zip_file_name
                st.session_state["pdf_output_mime"] = "application/zip"
    except FileNotFoundError as e:
        st.error(f"PDF生成に失敗しました：{e}")
    except Exception as e:
        st.error(f"PDF生成中にエラーが発生しました：{e}")

if st.session_state.get("pdf_output_bytes"):
    label = (
        "PDFをダウンロード"
        if st.session_state["pdf_output_mime"] == "application/pdf"
        else "ZIPをダウンロード"
    )
    st.download_button(
        label=label,
        data=st.session_state["pdf_output_bytes"],
        file_name=st.session_state["pdf_output_file_name"],
        mime=st.session_state["pdf_output_mime"],
        use_container_width=True,
    )