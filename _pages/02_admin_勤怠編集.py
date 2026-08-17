"""
02_admin_勤怠編集.py

【概要】
管理者用の勤怠閲覧・編集画面（SC-03）。要件定義書7.5節・4.5節、
基本設計書3.6節に基づく。

**重要**：一般ユーザーとは逆に、adminは「ロック中」のファイルのみ
編集操作へ切り替えられる（要件定義書4.5節・3.2節）。ロックされて
いない（一般ユーザーが編集可能な状態の）ファイルは、adminであっても
編集モードに入れない。この排他条件により、一般ユーザーとadminが
同一ファイルを同時に編集する状況は構造上発生しない
（要件定義書13.1節・基本設計書8.1節）。

プレビュー（読み取り専用表示）は、ロック状態に関わらずadminが
常に行える。
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

admin_user = session_service.require_admin()

st.title("勤怠閲覧・編集")


# ============================================
# 表示用ヘルパー関数
# ============================================

def _format_time(value) -> str:
    """(時, 分) のタプルを "H:MM" 表示に整形する。Noneは "--:--" とする。"""
    if value is None:
        return "--:--"
    hour, minute = value
    if hour < 0:
        return f"-{abs(hour)}:{minute:02d}"
    return f"{hour}:{minute:02d}"


def _format_plain(value) -> str:
    """自動計算項目等、そのまま表示してよい値の整形（Noneは空文字）。"""
    return "" if value is None else str(value)


def _render_preview_table(attendance_data) -> None:
    """
    読み取り専用のプレビューテーブルを表示する（基本設計書3.5.3節）。
    """
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
        st.info(
            "この月のデータはまだありません。"
            "「編集する」から入力を始めてください。"
        )
    else:
        st.dataframe(table_rows, use_container_width=True, hide_index=True)
        st.caption(
            "※プレビュー時は休憩・実働・超勤・休日出勤・深夜・深夜開始時間・"
            "合計欄などの自動計算項目もあわせて表示しています。"
            "編集画面ではこれらは表示されません（基本設計書3.5.3節）。"
        )

        st.markdown("**合計**")
        totals = attendance_data.totals
        totals_row = [
            {
                "定時": _format_time(totals.scheduled_total),
                "休憩1": _format_time(totals.break_time_1_total),
                "休憩2": _format_time(totals.break_time_2_total),
                "休憩3": _format_time(totals.break_time_3_total),
                "離業": _format_time(totals.leave_time_total),
                "実働": _format_time(totals.actual_work_time_total),
                "超勤": _format_time(totals.overtime_total),
                "休出": _format_time(totals.holiday_work_total),
                "超勤+休出": _format_time(totals.overtime_plus_holiday_total),
                "深夜": _format_time(totals.late_night_total),
            }
        ]
        st.dataframe(totals_row, use_container_width=True, hide_index=True)


_HOUR_OPTIONS = ["未入力"] + [str(h) for h in range(24)]
_MINUTE_OPTIONS = ["未入力"] + [f"{m:02d}" for m in range(60)]


def _time_tuple_to_hour_minute_str(value):
    """(時,分)タプルをプルダウン表示用の文字列2つに分解する。"""
    if value is None:
        return "未入力", "未入力"
    hour, minute = value
    return str(hour), f"{minute:02d}"


def _hour_minute_str_to_time_tuple(hour_str: str, minute_str: str):
    """
    プルダウンの選択結果を(時,分)タプルに変換する。
    どちらか一方でも「未入力」の場合はNone（未入力扱い）とする。
    """
    if hour_str == "未入力" or minute_str == "未入力":
        return None
    return int(hour_str), int(minute_str)


def _render_admin_edit_form(handle, attendance_data, target_employee, target_month) -> None:
    """
    admin用の編集フォームを表示する（基本設計書3.6節・8.3節）。

    一般ユーザー用の編集フォーム（01_勤怠入力.py）と同様、
    要件定義書4.6節に基づくリアルタイムバリデーション
    （休暇区分・土日の時刻入力制御と行ハイライト）を適用する。
    保存時は actor_role="admin" を渡し、保存直前チェックは
    「ロックが解除されていたら保存中止」となる（基本設計書8.3節。
    一般ユーザーとは逆の判定）。
    """
    st.subheader("勤怠データ（編集）")
    st.caption(
        "休憩・実働・超勤などの自動計算項目は編集画面には表示されません。"
        "保存後、プレビュー画面で自動計算結果を確認してください。"
    )

    editable_entries = [e for e in attendance_data.entries if e.date_value is not None]

    edit_rows: list[dict] = []
    row_has_error: dict[int, bool] = {}

    for i, entry in enumerate(editable_entries):
        row_key = entry.row

        leave_type_key = f"admin_leave_type_{row_key}"
        start_hour_key = f"admin_start_hour_{row_key}"
        start_minute_key = f"admin_start_minute_{row_key}"
        end_hour_key = f"admin_end_hour_{row_key}"
        end_minute_key = f"admin_end_minute_{row_key}"
        leave_hour_key = f"admin_leave_hour_{row_key}"
        leave_minute_key = f"admin_leave_minute_{row_key}"
        work_note_key = f"admin_work_note_{row_key}"

        current_leave_type = st.session_state.get(
            leave_type_key,
            entry.leave_type if entry.leave_type in excel_adapter.LEAVE_TYPE_OPTIONS else "",
        )

        is_full_day_leave = current_leave_type in validation_service.FULL_DAY_LEAVE_TYPES
        is_weekend_without_leave = (entry.weekday in ("土", "日")) and not current_leave_type
        should_disable_time = is_full_day_leave or is_weekend_without_leave

        if start_hour_key in st.session_state:
            current_start_hour = st.session_state[start_hour_key]
            current_start_minute = st.session_state[start_minute_key]
        else:
            current_start_hour, current_start_minute = _time_tuple_to_hour_minute_str(entry.start_time)

        if end_hour_key in st.session_state:
            current_end_hour = st.session_state[end_hour_key]
            current_end_minute = st.session_state[end_minute_key]
        else:
            current_end_hour, current_end_minute = _time_tuple_to_hour_minute_str(entry.end_time)

        has_time_input = (
            current_start_hour != "未入力" or current_start_minute != "未入力"
            or current_end_hour != "未入力" or current_end_minute != "未入力"
        )

        row_error_message = None
        if is_full_day_leave and has_time_input:
            row_error_message = (
                f"「{current_leave_type}」の日には始業・終業時間を入力できません。"
                "時刻を未入力に戻してください。"
            )
        elif is_weekend_without_leave and has_time_input:
            row_error_message = "休日に始業・終業時間が入力されています。休暇種類を選択するか、時刻を未入力に戻してください。"

        row_has_error[row_key] = row_error_message is not None

        row_container = st.container(border=row_error_message is not None)
        with row_container:
            st.markdown(f"**{entry.date_value}（{entry.weekday or ''}）**")

            if row_error_message:
                st.error(row_error_message, icon="⚠️")

            cols = st.columns([2, 1, 1, 1, 1, 1, 1, 2])
            show_labels = i == 0

            with cols[0]:
                leave_type = st.selectbox(
                    "休暇種類",
                    options=excel_adapter.LEAVE_TYPE_OPTIONS,
                    index=excel_adapter.LEAVE_TYPE_OPTIONS.index(current_leave_type)
                    if current_leave_type in excel_adapter.LEAVE_TYPE_OPTIONS else 0,
                    key=leave_type_key,
                    label_visibility="visible" if show_labels else "collapsed",
                )

            with cols[1]:
                start_hour = st.selectbox(
                    "始業(時)", _HOUR_OPTIONS,
                    index=_HOUR_OPTIONS.index(current_start_hour),
                    key=start_hour_key,
                    disabled=should_disable_time,
                    label_visibility="visible" if show_labels else "collapsed",
                )
            with cols[2]:
                start_minute = st.selectbox(
                    "始業(分)", _MINUTE_OPTIONS,
                    index=_MINUTE_OPTIONS.index(current_start_minute),
                    key=start_minute_key,
                    disabled=should_disable_time,
                    label_visibility="visible" if show_labels else "collapsed",
                )

            with cols[3]:
                end_hour = st.selectbox(
                    "終業(時)", _HOUR_OPTIONS,
                    index=_HOUR_OPTIONS.index(current_end_hour),
                    key=end_hour_key,
                    disabled=should_disable_time,
                    label_visibility="visible" if show_labels else "collapsed",
                )
            with cols[4]:
                end_minute = st.selectbox(
                    "終業(分)", _MINUTE_OPTIONS,
                    index=_MINUTE_OPTIONS.index(current_end_minute),
                    key=end_minute_key,
                    disabled=should_disable_time,
                    label_visibility="visible" if show_labels else "collapsed",
                )

            if leave_hour_key in st.session_state:
                current_leave_hour = st.session_state[leave_hour_key]
                current_leave_minute = st.session_state[leave_minute_key]
            else:
                current_leave_hour, current_leave_minute = _time_tuple_to_hour_minute_str(entry.leave_time)

            with cols[5]:
                leave_hour = st.selectbox(
                    "離業(時)", _HOUR_OPTIONS,
                    index=_HOUR_OPTIONS.index(current_leave_hour),
                    key=leave_hour_key,
                    disabled=should_disable_time,
                    label_visibility="visible" if show_labels else "collapsed",
                )
            with cols[6]:
                leave_minute = st.selectbox(
                    "離業(分)", _MINUTE_OPTIONS,
                    index=_MINUTE_OPTIONS.index(current_leave_minute),
                    key=leave_minute_key,
                    disabled=should_disable_time,
                    label_visibility="visible" if show_labels else "collapsed",
                )

            with cols[7]:
                work_note = st.text_input(
                    "自社工数内容",
                    value=st.session_state.get(work_note_key, entry.work_note or ""),
                    key=work_note_key,
                    disabled=should_disable_time,
                    label_visibility="visible" if show_labels else "collapsed",
                )

        edit_rows.append(
            {
                "row": entry.row,
                "time_row": entry.time_row,
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

    has_any_row_error = any(row_has_error.values())

    if has_any_row_error:
        st.warning("赤枠の行にエラーがあります。修正してから保存してください。")

    submitted = st.button(
        "保存する",
        use_container_width=True,
        disabled=has_any_row_error,
    )

    if submitted:
        edits = [
            attendance_service.DayEditInput(
                row=r["row"],
                time_row=r["time_row"],
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

        validation_errors = validation_service.validate_all(edits)

        if validation_errors:
            st.error("入力内容にエラーがあります。保存されていません。")
            date_labels = {e.row: f"{e.date_value}（{e.weekday or ''}）" for e in editable_entries}
            for err in validation_errors:
                label = date_labels.get(err.row_index, f"{err.row_index}行目")
                st.warning(f"{label}：{err.message}")
        else:
            try:
                with st.spinner("保存しています..."):
                    attendance_service.save_attendance(handle, edits, actor_role="admin")
            except attendance_service.LockedError as e:
                st.error(str(e))
            else:
                st.session_state["admin_edit_mode"] = False
                st.success("保存しました。")
                st.rerun()


def _render_admin_preview_table(attendance_data) -> None:
    """
    読み取り専用のプレビューテーブルを表示する（基本設計書3.6節）。
    ロック状態に関わらず常に表示可能。
    """
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
        st.caption(
            "※プレビュー時は休憩・実働・超勤・休日出勤・深夜等の"
            "自動計算項目もあわせて表示しています。"
        )

        st.markdown("**合計**")
        totals = attendance_data.totals
        totals_row = [
            {
                "定時": _format_time(totals.scheduled_total),
                "休憩1": _format_time(totals.break_time_1_total),
                "休憩2": _format_time(totals.break_time_2_total),
                "休憩3": _format_time(totals.break_time_3_total),
                "離業": _format_time(totals.leave_time_total),
                "実働": _format_time(totals.actual_work_time_total),
                "超勤": _format_time(totals.overtime_total),
                "休出": _format_time(totals.holiday_work_total),
                "超勤+休出": _format_time(totals.overtime_plus_holiday_total),
                "深夜": _format_time(totals.late_night_total),
            }
        ]
        st.dataframe(totals_row, use_container_width=True, hide_index=True)
        

# ============================================
# ここからメイン処理
# ============================================

all_users = user_service.list_users()

if not all_users:
    st.info("対象ユーザーが登録されていません。")
    st.stop()

col_user, col_month = st.columns(2)

with col_user:
    target_employee = st.selectbox(
        "対象ユーザー",
        options=all_users,
        format_func=lambda u: f"{u.last_name} {u.first_name}（{u.employee_id}）",
    )

with col_month:
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

# 対象（ユーザー or 月）を切り替えたら編集モードを解除する
_context_key = f"{target_employee.employee_id}_{target_month}"
if st.session_state.get("_admin_edit_context") != _context_key:
    st.session_state["admin_edit_mode"] = False
    st.session_state["_admin_edit_context"] = _context_key

is_locked = lock_service.is_locked(target_month, target_employee.employee_id)

status_col1, status_col2 = st.columns([3, 1])
with status_col1:
    st.write(
        f"**{target_employee.last_name} {target_employee.first_name}**　"
        f"社員番号：{target_employee.employee_id}　"
        f"部署：{target_employee.department}"
    )
with status_col2:
    if is_locked:
        st.error("● ロック中", icon="🔒")
    else:
        st.success("● 未ロック", icon="🔓")

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

# ============================================
# 編集切り替えボタン
#
# 【重要】一般ユーザーとは逆に、ロック中の場合のみ活性化する
# （要件定義書4.5節・基本設計書3.6節）。
# ============================================
admin_edit_mode = st.session_state.get("admin_edit_mode", False)

edit_btn_col, _ = st.columns([1, 4])
with edit_btn_col:
    if not admin_edit_mode:
        if st.button(
            "編集する",
            use_container_width=True,
            disabled=not is_locked,
        ):
            st.session_state["admin_edit_mode"] = True
            st.rerun()
    else:
        if st.button("編集をやめる", use_container_width=True):
            st.session_state["admin_edit_mode"] = False
            st.rerun()

if not is_locked:
    st.caption(
        "ロックされていないファイルは編集できません。"
        "ロック管理画面から先にロックしてください。"
    )

# ============================================
# プレビュー or 編集
# ============================================
if admin_edit_mode:
    _render_admin_edit_form(handle, attendance_data, target_employee, target_month)
else:
    _render_admin_preview_table(attendance_data)