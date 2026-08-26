"""
06_password_change.py

【概要】
一般ユーザー・admin共通のパスワード変更画面（SC-07）。
auth_service.change_password() が
「現在のパスワードの確認」「新パスワードのポリシー検証」の両方を
担っているため、この画面はフォームの表示とエラーメッセージの表示に専念する
（ロジックをここに書かない、という役割分担）。
"""

import streamlit as st

from services import auth_service, session_service

user = session_service.require_login()

st.title("パスワード変更")
name = "管理者" if user.is_admin else user.display_name
st.write(f"{name}さんのパスワードを変更します。")

with st.form("change_password_form"):
    current_password = st.text_input("現在のパスワード", type="password")
    new_password = st.text_input("新しいパスワード", type="password")
    new_password_confirm = st.text_input("新しいパスワード（確認）", type="password")
    submitted = st.form_submit_button("変更する")

if submitted:
    if new_password != new_password_confirm:
        st.error("新しいパスワードが一致しません")
    else:
        # 現在パスワードの確認・新パスワードのポリシー検証は
        # auth_service.change_password に委譲する
        ok, error_message = auth_service.change_password(
            employee_id=user.employee_id,
            current_password=current_password,
            new_password=new_password,
        )
        if ok:
            st.success("パスワードを変更しました")
        else:
            st.error(error_message)