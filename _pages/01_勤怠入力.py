"""
01_勤怠入力.py

【概要】
一般ユーザー用の勤怠入力・閲覧画面（SC-02）。

Day4: ヘッダー表示、対象年月選択、プレビュー表示（読み取り専用）
Day5: 編集モードへの切り替え・入力・保存（本ファイルで実装）

バリデーション（Day6）、ロックの実処理（Day7）は本ファイルでは扱わない。
"""

import datetime

import streamlit as st

from adapters import excel_adapter
from services import attendance_service, lock_service, session_service, settings_service, validation_service

user = session_service.require_general_user()

st.title("勤怠入力・閲覧")


# ============================================
# 表示用ヘルパー関数（プレビュー・編集共通）
# ============================================

def _format_time(value) -> str:
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

    日付が存在しない行（月末超過分等、Excelテンプレートの31日分確保用の
    空行）はそもそもテーブルに含めない。これは attendance_service 側で
    対象月の実日数分しか date_value が埋まらない設計になっているため、
    ここで除外すれば「存在しない日付」が表示されることはない
    （ダミーの31日分埋めは行わない）。
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


def _render_edit_form(handle, attendance_data) -> None:
    """
    編集フォームを表示する（画面全体をまとめて1回で保存する方式。
    基本設計書3.5.4節）。

    要件定義書4.6節に基づき、休暇種類・時刻入力の都度リアルタイムで
    矛盾を検知する（st.formは使わず、通常のウィジェットで都度rerun
    させることで実現する）。disabled化の対象は始業(時)〜自社工数内容
    の全項目とする（休暇区分の日は時刻・工数内容いずれも入力対象外
    のため）。
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

        leave_type_key = f"leave_type_{row_key}"
        start_hour_key = f"start_hour_{row_key}"
        start_minute_key = f"start_minute_{row_key}"
        end_hour_key = f"end_hour_{row_key}"
        end_minute_key = f"end_minute_{row_key}"
        leave_hour_key = f"leave_hour_{row_key}"
        leave_minute_key = f"leave_minute_{row_key}"
        work_note_key = f"work_note_{row_key}"

        current_leave_type = st.session_state.get(
            leave_type_key,
            entry.leave_type if entry.leave_type in excel_adapter.LEAVE_TYPE_OPTIONS else "",
        )

        is_full_day_leave = current_leave_type in validation_service.FULL_DAY_LEAVE_TYPES
        is_weekend_without_leave = (entry.weekday in ("土", "日")) and not current_leave_type
        should_disable_time = is_full_day_leave or is_weekend_without_leave

        # session_stateに既に値があればそれを使い、無ければentryの初期値を
        # 都度変換する（不要な変換呼び出しを減らすため、無い場合のみ計算）
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
                    attendance_service.save_attendance(handle, edits, actor_role="general")
            except attendance_service.LockedError as e:
                st.error(str(e))
            else:
                st.session_state["edit_mode"] = False
                st.success("保存しました。")
                st.rerun()


# ============================================
# ここからメイン処理
# ============================================

# ============================================
# 対象年月の選択（基本設計書3.5.2節）
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

# 対象年月を切り替えたら編集モードは自動的に解除する
if st.session_state.get("_edit_target_month") != target_month:
    st.session_state["edit_mode"] = False
    st.session_state["_edit_target_month"] = target_month

# ============================================
# ヘッダー（氏名・社員番号・部署名、ロック状態バッジ、ヘルプボタン）
# ============================================
is_locked = lock_service.is_locked(target_month=target_month, employee_id=user.employee_id)

header_col1, header_col2, header_col3 = st.columns([3, 2, 1])
with header_col1:
    st.write(f"**{user.display_name}**　社員番号：{user.employee_id}　部署：{user.department}")
with header_col2:
    if is_locked:
        st.error("● ロック中", icon="🔒")
    else:
        st.success("● 未ロック", icon="🔓")
with header_col3:
    if st.button("ヘルプ", use_container_width=True):
        st.session_state["show_help_notice"] = True

if st.session_state.get("show_help_notice"):
    st.info(
        "記入例ヘルプ（記入例シートの内容表示）は別途実装予定です。"
        "この案内は今後、記入例プレビューのモーダルに置き換わります。"
    )

st.divider()

# ============================================
# 月次ファイルの取得（存在しなければ生成。要件定義書4.2節）
# ============================================
with st.spinner("勤怠データを読み込んでいます..."):
    handle = attendance_service.get_or_create_monthly_file(
        employee_id=user.employee_id,
        target_month=target_month,
        full_name_no_space=user.full_name_no_space,
        full_name_with_space=user.full_name_with_space,
        department=user.department,
    )
    attendance_data = attendance_service.load_attendance(handle)

header = attendance_data.header

# ============================================
# Excelヘッダー情報表示
# ============================================
st.subheader("勤務情報")

col1, col2 = st.columns(2)

with col1:
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

with col2:
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
# ============================================
edit_mode = st.session_state.get("edit_mode", False)

edit_button_col, _ = st.columns([1, 4])
with edit_button_col:
    if not edit_mode:
        if st.button("編集する", disabled=is_locked, use_container_width=True):
            st.session_state["edit_mode"] = True
            st.rerun()
    else:
        if st.button("編集をやめる", use_container_width=True):
            st.session_state["edit_mode"] = False
            st.rerun()

if is_locked:
    st.caption("この月はロックされています。編集する場合は管理者に解除を依頼してください。")

# ============================================
# 勤怠データ（プレビュー or 編集）
# ============================================
if edit_mode:
    _render_edit_form(handle, attendance_data)
else:
    _render_preview_table(attendance_data)