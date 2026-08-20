"""
02_admin_attendance_edit.py

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

【編集フォームのUI方針】
一般ユーザー用の編集フォーム（01_attendance_input.py）と同一方針。
- 1日1行に、日・曜日・休暇種類・始業・終業・離業・自社工数内容を
  すべて並べる（日付列は「日」「曜日」のみ）。
- 時刻入力は "H:MM" 形式のテキストボックス1つに統一し、入力の都度
  形式チェックを行う（プルダウンは使わない）。
- 時刻テキスト入力は on_change のたびにサーバー側でサニタイズする
  （全角数字・全角コロンは半角へ自動変換し、数字とコロン以外の文字は
  除去する）。ブラウザのキー入力段階で物理的にブロックすることは
  st.text_input単体ではできないため、「入力した直後の再描画で
  不正な文字が自動的に消える／半角化される」という体験になる。

  【StreamlitAPIException対策】
  on_changeでsession_state[key]を書き換えるウィジェットには、
  st.text_input(value=..., key=...) のように value と key を同時に
  渡してはいけない（"created with a default value but also had its
  value set via the Session State API" という警告/例外の原因になる）。
  そのため、ウィジェット生成前に st.session_state[key] が未設定なら
  ここで初期値を代入し、st.text_input呼び出し時は key のみを渡す
  （value引数は渡さない）方式に統一する。

【ヘルプ（記入例）表示の方針】
01_attendance_input.pyと同様、attendance_service.load_example を使う。
自動計算列（休憩・実働・超勤等）もPython側で正しく再計算した
AttendanceDataを返すため、プレビュー画面と完全に同じ精度で
記入例を表示できる。
"""

import datetime
import re
import unicodedata

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


def _build_entry_rows(entries) -> list[dict]:
    """
    DayEntryのリストから、プレビュー・記入例表示共通の行データを組み立てる。
    """
    rows = []
    for entry in entries:
        if entry.date_value is None:
            continue

        rows.append(
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
    return rows


def _build_totals_row(attendance_data) -> list[dict]:
    """合計欄1行分のデータを組み立てる（プレビュー共通）。"""
    totals = attendance_data.totals
    return [
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


def _render_admin_preview_table(attendance_data) -> None:
    """
    読み取り専用のプレビューテーブルを表示する（基本設計書3.6節）。
    ロック状態に関わらず常に表示可能。
    """
    st.subheader("勤怠データ（プレビュー）")

    table_rows = _build_entry_rows(attendance_data.entries)

    if not table_rows:
        st.info("この月のデータはまだありません。")
        return

    st.dataframe(table_rows, use_container_width=True, hide_index=True)
    st.caption(
        "※プレビュー時は休憩・実働・超勤・休日出勤・深夜等の"
        "自動計算項目もあわせて表示しています。"
    )

    st.markdown("**合計**")
    st.dataframe(_build_totals_row(attendance_data), use_container_width=True, hide_index=True)


def _render_example_table(example_data) -> None:
    """
    「記入例」シートを、勤怠データ（プレビュー）と同じ
    レイアウト・表示形式（自動計算列込み）で表示する（要件定義書7.4節）。
    """
    table_rows = _build_entry_rows(example_data.entries)

    if not table_rows:
        st.info("記入例シートに表示する内容がありません。")
        return

    st.dataframe(table_rows, use_container_width=True, hide_index=True)


# ============================================
# 編集フォーム用ヘルパー（時刻テキスト入力）
# ============================================

# "H:MM" 形式（半角数字、時は0〜23で先頭ゼロなし、分は0〜59で2桁）
# excel_adapter.validate_excel_time_string と同一の形式を採用し、
# 画面側とExcel保存側で許容形式を一致させる。
_TIME_INPUT_PATTERN = re.compile(r"^([0-9]|1[0-9]|2[0-3]):[0-5][0-9]$")

# サニタイズ後に残してよい文字（半角数字・半角コロンのみ）
_ALLOWED_CHAR_PATTERN = re.compile(r"[0-9:]")


def _sanitize_time_input_text(text: str) -> str:
    """
    時刻入力欄のon_changeで呼ぶサニタイズ処理。
    01_attendance_input.py と同一ロジック（詳細はそちらのdocstring参照）。

    - 全角数字（０-９）・全角コロン（：）は unicodedata.normalize の
      NFKC正規化で半角に変換される。
    - 変換後、半角数字・半角コロン以外の文字はすべて除去する。
    """
    if not text:
        return text
    normalized = unicodedata.normalize("NFKC", text)
    return "".join(ch for ch in normalized if _ALLOWED_CHAR_PATTERN.match(ch))


def _time_tuple_to_input_text(value) -> str:
    """(時,分)タプルをテキスト入力欄の初期表示用文字列に変換する（Noneは空文字＝未入力）。"""
    if value is None:
        return ""
    hour, minute = value
    return f"{hour}:{minute:02d}"


def _parse_time_input_text(text: str):
    """
    テキスト入力欄の文字列を (時, 分) タプルに変換する。

    戻り値: (time_tuple, is_valid)
        - 空文字（前後の空白のみを含む場合も）は未入力として (None, True) を返す。
        - "H:MM" 形式（半角、時0〜23・分0〜59）に一致すれば (time_tuple, True)。
        - それ以外の文字列は不正な形式として (None, False) を返す
          （保存はさせず、呼び出し側で行を赤枠表示させる）。
    """
    stripped = (text or "").strip()
    if stripped == "":
        return None, True

    match = _TIME_INPUT_PATTERN.match(stripped)
    if not match:
        return None, False

    hour_str, minute_str = stripped.split(":")
    return (int(hour_str), int(minute_str)), True


def _make_sanitize_callback(text_key: str):
    """
    st.text_input の on_change に渡すコールバックを、対象ウィジェットの
    keyを束縛した形で生成する。on_changeはコールバック引数を取れないため、
    クロージャでtext_keyを固定する。
    """

    def _callback():
        current = st.session_state.get(text_key, "")
        st.session_state[text_key] = _sanitize_time_input_text(current)

    return _callback


def _render_time_text_input(row_key: str, field_prefix: str, label: str, entry_time_value, disabled: bool, label_visibility: str):
    """
    始業・終業・離業共通の "H:MM" 形式テキスト入力を描画する。
    session_stateのキーには "admin_" プレフィックスを付け、
    一般ユーザー画面（01_attendance_input.py）のキーと衝突しないようにする。

    st.text_input に value と key を同時に渡すと、on_changeで
    session_stateを書き換えるウィジェットの場合に
    StreamlitAPIException（"created with a default value but also
    had its value set via the Session State API"）が発生する。
    そのため、session_stateにまだキーが無い場合のみここで初期値を
    書き込み、st.text_input呼び出し時は value を渡さず key のみで
    値を委ねる（Session State一本化）。

    戻り値: (入力文字列, パース済みtime_tupleまたはNone, 形式が妥当か)
    """
    text_key = f"admin_{field_prefix}_text_{row_key}"

    if text_key not in st.session_state:
        st.session_state[text_key] = _time_tuple_to_input_text(entry_time_value)

    input_text = st.text_input(
        label,
        key=text_key,
        disabled=disabled,
        placeholder="　　:　　",
        label_visibility=label_visibility,
        on_change=_make_sanitize_callback(text_key),
    )

    time_tuple, is_valid = _parse_time_input_text(input_text)
    return input_text, time_tuple, is_valid


def _render_admin_edit_form(handle, attendance_data, target_employee, target_month) -> None:
    """
    admin用の編集フォームを表示する（基本設計書3.6節・8.3節）。

    一般ユーザー用の編集フォーム（01_attendance_input.py）と同様、
    要件定義書4.6節に基づくリアルタイムバリデーション
    （休暇区分・土日の時刻入力制御と行ハイライト）を適用する。
    保存時は actor_role="admin" を渡し、保存直前チェックは
    「ロックが解除されていたら保存中止」となる（基本設計書8.3節。
    一般ユーザーとは逆の判定）。

    1日1行に「日・曜日・休暇種類・始業・終業・離業・自社工数内容」を
    並べる。時刻入力は "H:MM" 形式のテキストボックス1つとし、形式が
    不正な場合も含めて行単位で赤枠・エラー表示する。
    """
    st.subheader("勤怠データ（編集）")
    st.caption(
        "休憩・実働・超勤などの自動計算項目は編集画面には表示されません。"
        "保存後、プレビュー画面で自動計算結果を確認してください。"
    )
    st.caption(
        "始業・終業・離業は「9:00」のように半角数字とコロンで入力してください"
        "（時：0〜23、分：0〜59）。全角で入力しても自動的に半角へ変換されます。"
        "未入力の場合は空欄のままにしてください。"
    )

    editable_entries = [e for e in attendance_data.entries if e.date_value is not None]

    edit_rows: list[dict] = []
    row_has_error: dict[int, bool] = {}

    for i, entry in enumerate(editable_entries):
        row_key = str(entry.row)
        leave_type_key = f"admin_leave_type_{row_key}"
        work_note_key = f"admin_work_note_{row_key}"
        show_labels = i == 0

        current_leave_type = st.session_state.get(
            leave_type_key,
            entry.leave_type if entry.leave_type in excel_adapter.LEAVE_TYPE_OPTIONS else "",
        )

        is_full_day_leave = current_leave_type in validation_service.FULL_DAY_LEAVE_TYPES
        is_weekend_without_leave = (entry.weekday in ("土", "日")) and not current_leave_type
        should_disable_time = is_full_day_leave or is_weekend_without_leave

        row_container = st.container()
        with row_container:
            cols = st.columns([1, 1, 3, 2, 2, 2, 5])
            label_visibility = "visible" if show_labels else "collapsed"

            with cols[0]:
                st.text_input(
                    "日",
                    value=str(entry.date_value.day) if entry.date_value else "",
                    key=f"admin_date_display_{row_key}",
                    disabled=True,
                    label_visibility=label_visibility,
                )
            with cols[1]:
                st.text_input(
                    "曜日",
                    value=entry.weekday or "",
                    key=f"admin_weekday_display_{row_key}",
                    disabled=True,
                    label_visibility=label_visibility,
                )

            with cols[2]:
                leave_type = st.selectbox(
                    "休暇種類",
                    options=excel_adapter.LEAVE_TYPE_OPTIONS,
                    index=excel_adapter.LEAVE_TYPE_OPTIONS.index(current_leave_type)
                    if current_leave_type in excel_adapter.LEAVE_TYPE_OPTIONS else 0,
                    key=leave_type_key,
                    label_visibility=label_visibility,
                )

            with cols[3]:
                start_text, start_time, start_valid = _render_time_text_input(
                    row_key, "start", "始業", entry.start_time, should_disable_time, label_visibility,
                )
            with cols[4]:
                end_text, end_time, end_valid = _render_time_text_input(
                    row_key, "end", "終業", entry.end_time, should_disable_time, label_visibility,
                )
            with cols[5]:
                leave_text, leave_time, leave_valid = _render_time_text_input(
                    row_key, "leave", "離業", entry.leave_time, should_disable_time, label_visibility,
                )

            with cols[6]:
                work_note = st.text_input(
                    "自社工数内容",
                    value=st.session_state.get(work_note_key, entry.work_note or ""),
                    key=work_note_key,
                    disabled=should_disable_time,
                    label_visibility=label_visibility,
                )

            has_time_input = start_time is not None or end_time is not None
            has_format_error = not (start_valid and end_valid and leave_valid)

            row_error_message = None
            if has_format_error:
                bad_fields = []
                if not start_valid:
                    bad_fields.append("始業")
                if not end_valid:
                    bad_fields.append("終業")
                if not leave_valid:
                    bad_fields.append("離業")
                row_error_message = (
                    f"{'・'.join(bad_fields)}の形式が不正です。"
                    "「9:00」のように半角数字とコロンで入力してください。"
                )
            elif is_full_day_leave and has_time_input:
                row_error_message = (
                    f"「{current_leave_type}」の日には始業・終業時間を入力できません。"
                    "時刻を未入力に戻してください。"
                )
            elif is_weekend_without_leave and has_time_input:
                row_error_message = "休日に始業・終業時間が入力されています。休暇種類を選択するか、時刻を未入力に戻してください。"

            row_has_error[entry.row] = row_error_message is not None

            if row_error_message:
                row_container.error(f"{entry.date_value}（{entry.weekday or ''}）：{row_error_message}", icon="⚠️")

        edit_rows.append(
            {
                "row": entry.row,
                "time_row": entry.time_row,
                "date_value": entry.date_value,
                "weekday": entry.weekday,
                "leave_type": leave_type,
                "start_time": start_time if start_valid else None,
                "end_time": end_time if end_valid else None,
                "leave_time": leave_time if leave_valid else None,
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

    if not submitted:
        return

    edits = [
        attendance_service.DayEditInput(
            row=r["row"],
            time_row=r["time_row"],
            date_value=r["date_value"],
            weekday=r["weekday"],
            leave_type=r["leave_type"],
            start_time=r["start_time"],
            end_time=r["end_time"],
            leave_time=r["leave_time"],
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
        return

    try:
        with st.spinner("保存しています..."):
            attendance_service.save_attendance(handle, edits, actor_role="admin")
    except attendance_service.LockedError as e:
        st.error(str(e))
    else:
        st.session_state["admin_edit_mode"] = False
        st.success("保存しました。")
        st.rerun()


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

status_col1, status_col2, status_col3 = st.columns([3, 1, 1])

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

with status_col3:
    if st.button("ヘルプ", use_container_width=True):
        st.session_state["admin_show_help"] = not st.session_state.get("admin_show_help", False)

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

# ============================================
# ヘルプ表示（記入例シート。attendance_service.load_example経由）
# ============================================
if st.session_state.get("admin_show_help", False):
    st.divider()
    st.subheader("📖 記入例")

    try:
        with st.spinner("記入例を読み込んでいます..."):
            example_data = attendance_service.load_example(handle)
        _render_example_table(example_data)
    except FileNotFoundError as e:
        st.error(f"記入例を読み込めませんでした。\n{e}")
    except (ValueError, KeyError) as e:
        st.error(f"記入例シートを読み込めませんでした。\n{e}")

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