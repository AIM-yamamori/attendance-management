"""
02_勤怠入力_編集.py

【概要】
一般ユーザー用の勤怠編集画面（SC-02の編集モード）。
基本設計書3.5.4節に基づき、Day5では以下を実装する。

- 対象年月・対象ユーザーは01_勤怠入力.pyから引き継ぐ
  （st.session_stateで target_month を受け渡す想定）
- 編集フォーム（st.form）で全日分をまとめて入力し、保存ボタン1つで
  一括保存する（画面全体を1回でまとめて保存する方式）
- 時刻項目は時・分を別々のプルダウン（selectbox）で入力し、
  (時,分) タプルに変換してサービス層へ渡す
- 休暇種類は固定の選択肢リスト（excel_adapter.LEAVE_TYPE_OPTIONS）
  からselectboxで選択する
- 自動計算項目（休憩・実働・超勤等）は編集画面では表示しない
  （基本設計書3.5.3節・7.3節）
- ロックされている月は編集不可（Day7で実処理。Day5時点では
  01_勤怠入力.py側で編集ボタンが既にdisabledになっているが、
  URLで直接アクセスされた場合の防御としてこのファイルでも
  ロック確認を行う）

バリデーション（基本設計書3.5.6節等の入力チェック）はDay6で実装する。
Day5時点では最低限の形式チェック（時・分の範囲）のみ行う。
"""

import streamlit as st

from adapters import excel_adapter
from services import attendance_service, lock_service, session_service, settings_service

user = session_service.require_general_user()

st.title("勤怠編集")

# ============================================
# 対象年月の引き継ぎ（01_勤怠入力.pyのselectboxから遷移する想定）
# ============================================
target_month = st.session_state.get("target_month")

if not target_month:
    st.warning("対象年月が選択されていません。勤怠入力画面から操作してください。")
    st.stop()

# ============================================
# ロック確認（URL直接アクセス等への防御。実処理はDay7）
# ============================================
is_locked = lock_service.is_locked(target_month=target_month, employee_id=user.employee_id)
if is_locked:
    st.error("この月はロックされています。編集する場合は管理者に解除を依頼してください。", icon="🔒")
    st.stop()

st.write(f"**{user.display_name}**　社員番号：{user.employee_id}　部署：{user.department}")
st.caption(attendance_service.format_month_label(target_month))
st.divider()

# ============================================
# 月次ファイルの取得
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

# ============================================
# 時刻入力ヘルパー
# ============================================
_HOUR_OPTIONS = ["未入力"] + [str(h) for h in range(24)]
_MINUTE_OPTIONS = ["未入力"] + [f"{m:02d}" for m in range(60)]


def _time_tuple_to_hour_minute_str(value: tuple[int, int] | None) -> tuple[str, str]:
    """(時,分)タプルをプルダウン表示用の文字列2つに分解する。"""
    if value is None:
        return "未入力", "未入力"
    hour, minute = value
    return str(hour), f"{minute:02d}"


def _hour_minute_str_to_time_tuple(hour_str: str, minute_str: str) -> tuple[int, int] | None:
    """
    プルダウンの選択結果を(時,分)タプルに変換する。
    どちらか一方でも「未入力」の場合はNone（未入力扱い）とする。
    """
    if hour_str == "未入力" or minute_str == "未入力":
        return None
    return int(hour_str), int(minute_str)


# ============================================
# 編集フォーム（画面全体をまとめて1回で保存）
# ============================================
st.subheader("勤怠データ（編集）")
st.caption(
    "休憩・実働・超勤などの自動計算項目は編集画面には表示されません。"
    "保存後、プレビュー画面で自動計算結果を確認してください。"
)

with st.form("attendance_edit_form"):
    edit_rows: list[dict] = []

    for entry in attendance_data.entries:
        # 日付が存在しない行（月末超過分等）は編集対象外
        if entry.date_value is None:
            continue

        st.markdown(f"**{entry.date_value}（{entry.weekday or ''}）**")

        cols = st.columns([2, 1, 1, 1, 1, 1, 1, 2])

        with cols[0]:
            leave_type = st.selectbox(
                "休暇種類",
                options=excel_adapter.LEAVE_TYPE_OPTIONS,
                index=(
                    excel_adapter.LEAVE_TYPE_OPTIONS.index(entry.leave_type)
                    if entry.leave_type in excel_adapter.LEAVE_TYPE_OPTIONS
                    else 0
                ),
                key=f"leave_type_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )

        start_hour_default, start_minute_default = _time_tuple_to_hour_minute_str(entry.start_time)
        with cols[1]:
            start_hour = st.selectbox(
                "始業(時)", _HOUR_OPTIONS,
                index=_HOUR_OPTIONS.index(start_hour_default),
                key=f"start_hour_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )
        with cols[2]:
            start_minute = st.selectbox(
                "始業(分)", _MINUTE_OPTIONS,
                index=_MINUTE_OPTIONS.index(start_minute_default),
                key=f"start_minute_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )

        end_hour_default, end_minute_default = _time_tuple_to_hour_minute_str(entry.end_time)
        with cols[3]:
            end_hour = st.selectbox(
                "終業(時)", _HOUR_OPTIONS,
                index=_HOUR_OPTIONS.index(end_hour_default),
                key=f"end_hour_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )
        with cols[4]:
            end_minute = st.selectbox(
                "終業(分)", _MINUTE_OPTIONS,
                index=_MINUTE_OPTIONS.index(end_minute_default),
                key=f"end_minute_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )

        leave_hour_default, leave_minute_default = _time_tuple_to_hour_minute_str(entry.leave_time)
        with cols[5]:
            leave_hour = st.selectbox(
                "離業(時)", _HOUR_OPTIONS,
                index=_HOUR_OPTIONS.index(leave_hour_default),
                key=f"leave_hour_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )
        with cols[6]:
            leave_minute = st.selectbox(
                "離業(分)", _MINUTE_OPTIONS,
                index=_MINUTE_OPTIONS.index(leave_minute_default),
                key=f"leave_minute_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )

        with cols[7]:
            work_note = st.text_input(
                "自社工数内容",
                value=entry.work_note or "",
                key=f"work_note_{entry.row}",
                label_visibility="collapsed" if entry.row != attendance_data.entries[0].row else "visible",
            )

        edit_rows.append(
            {
                "row": entry.row,
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

# ============================================
# 保存処理
# ============================================
if submitted:
    edits = [
        attendance_service.DayEditInput(
            row=r["row"],
            leave_type=r["leave_type"],
            start_time=_hour_minute_str_to_time_tuple(r["start_hour"], r["start_minute"]),
            end_time=_hour_minute_str_to_time_tuple(r["end_hour"], r["end_minute"]),
            leave_time=_hour_minute_str_to_time_tuple(r["leave_hour"], r["leave_minute"]),
            work_note=r["work_note"],
        )
        for r in edit_rows
    ]

    with st.spinner("保存しています..."):
        attendance_service.save_attendance(handle, edits)

    st.success("保存しました。")
    st.rerun()