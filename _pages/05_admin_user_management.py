"""
05_admin_user_management.py

【概要】
admin用のユーザーマスタ管理画面（SC-06）。
一般ユーザーの一覧表示・新規追加・削除・パスワード再設定を行う。
画面はフォームの表示とエラーメッセージの表示に専念し、
実際のバリデーション・DB操作はすべて user_service / auth_service に委譲する
（3.9節「本画面から権限区分adminのユーザーを新規作成する機能は設けない」
に対応し、新規追加フォームには一般ユーザーの登録項目のみを置く）。
"""

import streamlit as st

from services import auth_service, session_service, user_service

user = session_service.require_admin()

st.title("ユーザー管理（admin）")

# ============================================
# 一覧表示
# ============================================
st.subheader("一般ユーザー一覧")

users = user_service.list_users()
if not users:
    st.info("登録されている一般ユーザーはいません。")
else:
    st.dataframe(
        [
            {
                "社員番号": u.employee_id,
                "姓": u.last_name,
                "名": u.first_name,
                "部署名": u.department,
            }
            for u in users
        ],
        use_container_width=True,
        hide_index=True,
    )

st.divider()

# ============================================
# 新規追加
# ============================================
st.subheader("新規追加")

with st.form("create_user_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        new_last_name = st.text_input("姓")
        new_employee_id = st.text_input("社員番号（半角数字、0埋めなし）")
    with col2:
        new_first_name = st.text_input("名")
        new_department = st.text_input("部署名")

    new_initial_password = st.text_input("初期パスワード", type="password")
    st.caption(
        f"{auth_service.PASSWORD_MIN_LENGTH}文字以上"
        f"{auth_service.PASSWORD_MAX_LENGTH}文字以内、"
        "英大文字・英小文字・数字のうち2種類以上を組み合わせてください。"
    )

    create_submitted = st.form_submit_button("追加する")

if create_submitted:
    # バリデーション・重複チェック等は user_service.create_user に委譲する
    ok, error_message = user_service.create_user(
        last_name=new_last_name,
        first_name=new_first_name,
        employee_id=new_employee_id,
        department=new_department,
        initial_password=new_initial_password,
    )
    if ok:
        st.success(f"社員番号 {new_employee_id} のユーザーを追加しました")
        st.rerun()
    else:
        st.error(error_message)

st.divider()

# ============================================
# 編集
# ============================================
st.subheader("編集")

if not users:
    st.info("対象ユーザーがいません。")
else:
    edit_target_options = {
        f"{u.display_name}（{u.employee_id}）": u.employee_id for u in users
    }
    edit_target_label = st.selectbox(
        "編集対象ユーザー", list(edit_target_options.keys()), key="edit_target"
    )
    edit_target_employee_id = edit_target_options[edit_target_label]
    edit_target_user = user_service.get_user(edit_target_employee_id)

    with st.form("update_user_form"):
        # 社員番号はログインIDを兼ねるため編集不可（表示のみ）
        st.caption(f"社員番号（ログインID）: {edit_target_user.employee_id}（変更不可）")
        edit_last_name = st.text_input("姓", value=edit_target_user.last_name)
        edit_first_name = st.text_input("名", value=edit_target_user.first_name)
        edit_department = st.text_input("部署名", value=edit_target_user.department)
        update_submitted = st.form_submit_button("更新する")

    if update_submitted:
        ok, error_message = user_service.update_user(
            employee_id=edit_target_employee_id,
            last_name=edit_last_name,
            first_name=edit_first_name,
            department=edit_department,
        )
        if ok:
            st.success(f"{edit_target_label} の情報を更新しました")
            st.rerun()
        else:
            st.error(error_message)

st.divider()

# ============================================
# 削除 / パスワード再設定
# ============================================
st.subheader("削除・パスワード再設定")

if not users:
    st.info("対象ユーザーがいません。")
else:
    target_options = {f"{u.display_name}（{u.employee_id}）": u.employee_id for u in users}
    target_label = st.selectbox("対象ユーザー", list(target_options.keys()))
    target_employee_id = target_options[target_label]

    col_reset, col_delete = st.columns(2)

    with col_reset:
        st.write("**パスワード再設定**")
        with st.form("reset_password_form"):
            new_password = st.text_input("新しいパスワード", type="password")
            reset_submitted = st.form_submit_button("再設定する")
        if reset_submitted:
            ok, error_message = auth_service.reset_password_by_admin(
                target_employee_id, new_password
            )
            if ok:
                # 設定直後の確認表示のみ（要件定義書5.4節）。
                # 以後この画面上には表示しない（session_stateに保持しない）。
                st.success(
                    f"{target_label} のパスワードを再設定しました。"
                    f"新パスワード「{new_password}」を本人へ直接伝えてください。"
                )
            else:
                st.error(error_message)

    with col_delete:
        st.write("**削除**")
        confirm = st.checkbox(f"{target_label} を削除することを確認しました")
        if st.button("削除する", disabled=not confirm):
            ok, error_message = user_service.delete_user(target_employee_id)
            if ok:
                st.success(f"{target_label} を削除しました")
                st.rerun()
            else:
                st.error(error_message)