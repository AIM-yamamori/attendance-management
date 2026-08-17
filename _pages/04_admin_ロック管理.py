"""
04_admin_ロック管理.py

【概要】
管理者用のロック管理画面（SC-05）。要件定義書7.7節・11章、
基本設計書3.8節に基づく。

- 対象月を選択すると、その月の全ユーザーのロック状態一覧
  （氏名・社員番号・ロック状態バッジ）を表示する。
- 「まとめてロック」ボタンで一括ロック（実行前に確認ダイアログを表示）。
- 各ユーザー行に「ロック」「解除」の個別操作ボタンを配置する
  （現在の状態と逆の操作のみ活性）。
- 解除は個別解除のみ（要件定義書4.8節。一括解除機能は設けない）。
"""

import datetime

import streamlit as st

from services import attendance_service, lock_service, session_service, settings_service

admin_user = session_service.require_admin()

st.title("ロック管理")

# ============================================
# 対象年月の選択（SC-02と同様の範囲：運用開始月〜当月の翌月）
# ============================================
service_start_month = settings_service.get_service_start_month()
today = datetime.date.today()

selectable_months = attendance_service.build_selectable_months(
    today=today, service_start_month=service_start_month
)
default_month = attendance_service.default_target_month(today)

if default_month in selectable_months:
    default_index = selectable_months.index(default_month)
else:
    default_index = len(selectable_months) - 1

target_month = st.selectbox(
    "対象年月",
    options=selectable_months,
    index=default_index,
    format_func=attendance_service.format_month_label,
)

st.divider()

# ============================================
# 「まとめてロック」（実行前に確認ダイアログを表示。基本設計書3.8節）
# ============================================

@st.dialog("まとめてロックの確認")
def _confirm_lock_all():
    st.write(
        f"{attendance_service.format_month_label(target_month)} の"
        "全ユーザーをロックします。よろしいですか？"
    )
    confirm_col, cancel_col = st.columns(2)
    with confirm_col:
        if st.button("ロックする", use_container_width=True, type="primary"):
            count = lock_service.lock_all(target_month)
            st.session_state["lock_all_result"] = count
            st.rerun()
    with cancel_col:
        if st.button("キャンセル", use_container_width=True):
            st.rerun()


if st.session_state.get("lock_all_result") is not None:
    st.success(f"{st.session_state['lock_all_result']}名をロックしました。")
    st.session_state["lock_all_result"] = None

if st.button("まとめてロック", type="primary"):
    _confirm_lock_all()

st.divider()

# ============================================
# ロック状態一覧
# ============================================
st.subheader("ロック状態一覧")

lock_statuses = lock_service.list_lock_status(target_month)

if not lock_statuses:
    st.info("対象ユーザーが登録されていません。")
else:
    header_col1, header_col2, header_col3, header_col4 = st.columns([2, 2, 2, 2])
    with header_col1:
        st.markdown("**氏名**")
    with header_col2:
        st.markdown("**社員番号**")
    with header_col3:
        st.markdown("**状態**")
    with header_col4:
        st.markdown("**操作**")

    for status in lock_statuses:
        row_col1, row_col2, row_col3, row_col4 = st.columns([2, 2, 2, 2])

        with row_col1:
            st.write(f"{status.last_name} {status.first_name}")
        with row_col2:
            st.write(status.employee_id)
        with row_col3:
            if status.is_locked:
                st.error("● ロック中", icon="🔒")
            else:
                st.success("● 未ロック", icon="🔓")
        with row_col4:
            if status.is_locked:
                if st.button("解除", key=f"unlock_{status.employee_id}", use_container_width=True):
                    lock_service.unlock_one(target_month, status.employee_id)
                    st.rerun()
            else:
                if st.button("ロック", key=f"lock_{status.employee_id}", use_container_width=True):
                    lock_service.lock_one(target_month, status.employee_id)
                    st.rerun()