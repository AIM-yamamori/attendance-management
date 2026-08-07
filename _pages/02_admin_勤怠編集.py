"""
03_admin_勤怠管理.py

【概要】
管理者用の勤怠閲覧・編集・ロック管理画面（Day8、基本設計書3.6節相当）。

- 対象年月・対象ユーザーを選択して、一般ユーザー画面と同じ
  プレビュー/編集UIを表示する（01_勤怠入力.pyのロジックを再利用）
- ロック管理：月全体の一括ロック/解除、ユーザー個別のロック/解除
- adminは自分自身のロック状態に関わらず常に編集可能とする
  （ロック機能はあくまで一般ユーザーの入力を止めるための機能であり、
  admin自身の編集を妨げるものではないため。基本設計書3.6節想定）

【未確定】全ユーザーの一覧取得方法（社員マスタの取得元）が
本体のユーザー管理実装に依存するため、user_service.list_all_users()
という想定インターフェースを使用している。実装が確定次第、
呼び出し部分のみ差し替えが必要。
"""

import datetime

import streamlit as st

from adapters import excel_adapter
from services import (
    attendance_service,
    lock_service,
    session_service,
    settings_service,
    user_service,
    validation_service,
)

admin_user = session_service.require_admin_user()

st.title("勤怠管理（管理者）")

tab_view, tab_lock = st.tabs(["勤怠閲覧・編集", "ロック管理"])


def _format_time(value) -> str:
    if value is None:
        return "--:--"
    hour, minute = value
    if hour < 0:
        return f"-{abs(hour)}:{minute:02d}"
    return f"{hour}:{minute:02d}"


def _format_plain(value) -> str:
    return "" if value is None else str(value)


def _render_admin_preview_table(attendance_data) -> None:
    st.subheader("勤怠データ（プレビュー）")

    table_rows = []
    for entry in attendance_data.entries:
        if entry.date_value is None:
            continue
        table_rows.append(
            {
                "日": entry.date_value,
                "曜日": entry.weekday or "",
                "休暇種類": entry.leave_type or "",
                "始業": _format_time(entry.start_time),
                "終業": _format_time(entry.end_time),
                "休憩1": _format_time(entry.break_time_1),
                "休憩2": _format_time(entry.break_time_2),
                "休憩3": _format_time(entry.break_time_3),
                "離業": _format_time(entry.leave_time),
                "実働": _format_time(entry.actual_work_time),
                "超勤": _format_time(entry.overtime),
                "休出": _format_time(entry.holiday_work),
                "深夜": _format_time(entry.late_night),
                "自社工数内容": entry.work_note or "",
            }
        )

    if not table_rows:
        st.info("この月のデータはまだありません。")
    else:
        st.dataframe(table_rows, use_container_width=True, hide_index=True)


_HOUR_OPTIONS = ["未入力"] + [str(h) for h in range(24)]
_MINUTE_OPTIONS = ["未入力"] + [f"{m:02d}" for m in range(60)]


def _time_tuple_to_hour_minute_str(value):
    if value is None:
        return "未入力", "未入力"
    hour, minute = value
    return str(hour), f"{minute:02d}"


def _hour_minute_str_to_time_tuple(hour_str: str, minute_str: str):
    if hour_str == "未入力" or minute_str == "未入力":
        return None
    return int(hour_str), int(minute_str)


def _render_admin_edit_form(handle, attendance_data) -> None:
    st.subheader("勤怠データ（編集）")

    editable_entries = [e for e in attendance_data.entries if e.date_value is not None]

    with st.form("admin_attendance_edit_form"):
        edit_rows: list[dict] = []

        for i, entry in enumerate(editable_entries):
            st.markdown(f"**{entry.date_value}（{entry.weekday or ''}）**")
            cols = st.columns([2, 1, 1, 1, 1, 1, 1, 2])
            show_labels = i == 0

            with cols[0]:
                leave_type = st.selectbox(
                    "休暇種類",
                    options=excel_adapter.LEAVE_TYPE_OPTIONS,
                    index=(
                        excel_adapter.LEAVE_TYPE_OPTIONS.index(entry.leave_type)
                        if entry.leave_type in excel_adapter.LEAVE_TYPE_OPTIONS
                        else 0
                    ),
                    key=f"admin_leave_type_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )

            start_h_def, start_m_def = _time_tuple_to_hour_minute_str(entry.start_time)
            with cols[1]:
                start_hour = st.selectbox(
                    "始業(時)", _HOUR_OPTIONS, index=_HOUR_OPTIONS.index(start_h_def),
                    key=f"admin_start_hour_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )
            with cols[2]:
                start_minute = st.selectbox(
                    "始業(分)", _MINUTE_OPTIONS, index=_MINUTE_OPTIONS.index(start_m_def),
                    key=f"admin_start_minute_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )

            end_h_def, end_m_def = _time_tuple_to_hour_minute_str(entry.end_time)
            with cols[3]:
                end_hour = st.selectbox(
                    "終業(時)", _HOUR_OPTIONS, index=_HOUR_OPTIONS.index(end_h_def),
                    key=f"admin_end_hour_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )
            with cols[4]:
                end_minute = st.selectbox(
                    "終業(分)", _MINUTE_OPTIONS, index=_MINUTE_OPTIONS.index(end_m_def),
                    key=f"admin_end_minute_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )

            leave_h_def, leave_m_def = _time_tuple_to_hour_minute_str(entry.leave_time)
            with cols[5]:
                leave_hour = st.selectbox(
                    "離業(時)", _HOUR_OPTIONS, index=_HOUR_OPTIONS.index(leave_h_def),
                    key=f"admin_leave_hour_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )
            with cols[6]:
                leave_minute = st.selectbox(
                    "離業(分)", _MINUTE_OPTIONS, index=_MINUTE_OPTIONS.index(leave_m_def),
                    key=f"admin_leave_minute_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )

            with cols[7]:
                work_note = st.text_input(
                    "自社工数内容",
                    value=entry.work_note or "",
                    key=f"admin_work_note_{entry.row}",
                    label_visibility="visible" if show_labels else "collapsed",
                )

            edit_rows.append(
                {
                    "row": entry.row,
                    "date_value": entry.date_value,
                    "weekday": entry.weekday,
                    "leave_type": leave_type,
                    "start_hour": start_hour,
                    "start_minute": start_minute,
                    "end_hour": end_hour,
                    "end_minute": end_minute,
                    "leave_hour": leave_hour,
                    "leave_minute": leave_minute,
                    "work_note": work_note,
                }
            )

        submitted = st.form_submit_button("保存する", use_container_width=True)

    if submitted:
        edits = [
            attendance_service.DayEditInput(
                row=r["row"],
                date_value=r["date_value"],
                weekday=r["weekday"],
                leave_type=r["leave_type"],
                start_time=_hour_minute_str_to_time_tuple(r["start_hour"], r["start_minute"]),
                end_time=_hour_minute_str_to_time_tuple(r["end_hour"], r["end_minute"]),
                leave_time=_hour_minute_str_to_time_tuple(r["leave_hour"], r["leave_minute"]),
                work_note=r["work_note"],
            )
            for r in edit_rows
        ]

        # ① バリデーション（基本設計書8.2節①）
        validation_errors = validation_service.validate_all(edits)

        if validation_errors:
            st.error("入力内容にエラーがあります。保存されていません。")
            date_labels = {e.row: f"{e.date_value}（{e.weekday or ''}）" for e in editable_entries}
            for err in validation_errors:
                label = date_labels.get(err.row_index, f"{err.row_index}行目")
                st.warning(f"{label}：{err.message}")
        else:
            # ② ロック再チェック → ③ 保存（基本設計書8.2節②③）
            try:
                with st.spinner("保存しています..."):
                    attendance_service.save_attendance(handle, edits, actor_role="general")
            except attendance_service.LockedError as e:
                st.error(str(e))
                # 8.2節：「入力内容は画面上に残す（消去しない）」
                # st.rerun()しないことで、フォームの入力内容はそのまま維持される
            else:
                st.session_state["edit_mode"] = False
                st.success("保存しました。")
                st.rerun()

# ============================================
# タブ1: 勤怠閲覧・編集
# ============================================
with tab_view:
    all_users = user_service.list_all_users()

    col_month, col_user = st.columns(2)

    with col_month:
        service_start_month = settings_service.get_service_start_month()
        today = datetime.date.today()
        selectable_months = attendance_service.build_selectable_months(
            today=today, service_start_month=service_start_month
        )
        target_month = st.selectbox(
            "対象年月",
            options=selectable_months,
            index=len(selectable_months) - 1,
            format_func=attendance_service.format_month_label,
            key="admin_target_month",
        )

    with col_user:
        target_employee = st.selectbox(
            "対象ユーザー",
            options=all_users,
            format_func=lambda u: f"{u.display_name}（{u.employee_id}）",
            key="admin_target_employee",
        )

    # 対象を切り替えたら編集モードをリセット
    _admin_context_key = f"{target_month}_{target_employee.employee_id}"
    if st.session_state.get("_admin_edit_context") != _admin_context_key:
        st.session_state["admin_edit_mode"] = False
        st.session_state["_admin_edit_context"] = _admin_context_key

    is_locked = lock_service.is_locked(
        target_month=target_month, employee_id=target_employee.employee_id
    )

    status_col1, status_col2 = st.columns([3, 1])
    with status_col1:
        st.write(
            f"**{target_employee.display_name}**　"
            f"社員番号：{target_employee.employee_id}　"
            f"部署：{target_employee.department}"
        )
    with status_col2:
        if is_locked:
            st.error("● ロック中", icon="🔒")
        else:
            st.success("● 未ロック", icon="🔓")

    st.caption(
        "管理者はロック中でも編集できます。"
        "一般ユーザーからの入力を止める目的のロックのため、"
        "管理者操作には影響しません。"
    )

    st.divider()

    with st.spinner("勤怠データを読み込んでいます..."):
        handle = attendance_service.get_or_create_monthly_file(
            employee_id=target_employee.employee_id,
            target_month=target_month,
            full_name_no_space=target_employee.full_name_no_space,
            full_name_with_space=target_employee.full_name_with_space,
            department=target_employee.department,
        )
        attendance_data = attendance_service.load_attendance(handle)

    header = attendance_data.header

    st.subheader("勤務情報")
    hcol1, hcol2 = st.columns(2)
    with hcol1:
        st.write(
            f"""
            **派遣先企業名**
            {header.client_company_name}

            **派遣先部署名**
            {header.client_department}

            **社員番号**
            {header.employee_id}

            **氏名**
            {header.employee_name}
            """
        )
    with hcol2:
        st.write(
            f"""
            **深夜開始時間**
            {excel_adapter.format_time_value_no_seconds(header.night_start_time)}

            **所定**
            {excel_adapter.format_time_value_no_seconds(header.scheduled_work_time)}

            **午前**
            {excel_adapter.format_time_value_no_seconds(header.morning_time)}

            **午後**
            {excel_adapter.format_time_value_no_seconds(header.afternoon_time)}

            **休憩時間1**
            {excel_adapter.format_time_value_no_seconds(header.break_time_1_start)} ～ {excel_adapter.format_time_value_no_seconds(header.break_time_1_end)}

            **休憩時間2**
            {excel_adapter.format_time_value_no_seconds(header.break_time_2_start)} ～ {excel_adapter.format_time_value_no_seconds(header.break_time_2_end)}

            **休憩時間3**
            {excel_adapter.format_time_value_no_seconds(header.break_time_3_start)} ～ {excel_adapter.format_time_value_no_seconds(header.break_time_3_end)}
            """
        )

    admin_edit_mode = st.session_state.get("admin_edit_mode", False)

    edit_btn_col, _ = st.columns([1, 4])
    with edit_btn_col:
        if not admin_edit_mode:
            if st.button("編集する", use_container_width=True, key="admin_edit_btn"):
                st.session_state["admin_edit_mode"] = True
                st.rerun()
        else:
            if st.button("編集をやめる", use_container_width=True, key="admin_edit_cancel_btn"):
                st.session_state["admin_edit_mode"] = False
                st.rerun()

    if admin_edit_mode:
        _render_admin_edit_form(handle, attendance_data)
    else:
        _render_admin_preview_table(attendance_data)

# ============================================
# タブ2: ロック管理
# ============================================
with tab_lock:
    st.subheader("ロック管理")

    lock_target_month = st.selectbox(
        "対象年月",
        options=selectable_months,
        index=len(selectable_months) - 1,
        format_func=attendance_service.format_month_label,
        key="lock_target_month",
    )

    lock_state = lock_service.get_lock_summary(lock_target_month)

    st.write("### 月全体ロック")
    if lock_state.is_month_locked:
        st.error(f"{attendance_service.format_month_label(lock_target_month)} は全体ロック中です。", icon="🔒")
        if st.button("月全体ロックを解除する", key="unlock_month_btn"):
            lock_service.unlock_entire_month(lock_target_month)
            st.success("月全体ロックを解除しました。")
            st.rerun()
    else:
        st.info(f"{attendance_service.format_month_label(lock_target_month)} は全体ロックされていません。")
        if st.button("月全体を一括ロックする", key="lock_month_btn"):
            lock_service.lock_entire_month(lock_target_month)
            st.success("月全体をロックしました。")
            st.rerun()

    st.divider()

    st.write("### ユーザー個別のロック")
    all_users_for_lock = user_service.list_all_users()

    for u in all_users_for_lock:
        is_individually_locked = u.employee_id in lock_state.locked_employee_ids
        is_individually_unlocked_from_month_lock = (
            lock_state.is_month_locked and u.employee_id in lock_state.unlocked_employee_ids
        )
        current_locked = lock_service.is_locked(lock_target_month, u.employee_id)

        row_col1, row_col2, row_col3 = st.columns([3, 1, 1])
        with row_col1:
            st.write(f"{u.display_name}（{u.employee_id}）")
        with row_col2:
            if current_locked:
                st.error("ロック中", icon="🔒")
            else:
                st.success("未ロック", icon="🔓")
        with row_col3:
            if current_locked:
                if st.button("解除", key=f"unlock_{u.employee_id}"):
                    lock_service.unlock_employee(lock_target_month, u.employee_id)
                    st.rerun()
            else:
                if st.button("ロック", key=f"lock_{u.employee_id}"):
                    lock_service.lock_employee(lock_target_month, u.employee_id)
                    st.rerun()

                    