"""
01_勤怠入力.py

【概要】
一般ユーザー用の勤怠入力・閲覧画面（SC-02）。
基本設計書3.5節に基づき、Day4では以下を実装する。

- ヘッダー（対象年月表示・切替、氏名・社員番号・部署名表示、
  ロック状態バッジ、ヘルプボタンの土台）
- 対象年月の選択制御（当年当月デフォルト・翌月まで選択可・
  運用開始月より前は選択肢に表示しない。基本設計書3.5.2節）
- プレビュー表示（読み取り専用テーブル。休憩・実働・超勤等の
  自動計算項目も含めた全項目表示。基本設計書3.5.3節）

編集モードへの切り替え・入力・保存（Day5）、バリデーション（Day6）、
ロックの実処理（Day7）は本ファイルでは扱わない。編集ボタンは
土台として配置するが、押しても「Day5で実装予定」の案内を出すのみ。
"""

import datetime

import streamlit as st

from services import attendance_service, lock_service, session_service, settings_service

from adapters import excel_adapter

user = session_service.require_general_user()

st.title("勤怠入力・閲覧")

# ============================================
# 対象年月の選択（基本設計書3.5.2節）
# ============================================
service_start_month = settings_service.get_service_start_month()
today = datetime.date.today()

selectable_months = attendance_service.build_selectable_months(
    today=today, service_start_month=service_start_month
)
default_month = attendance_service.default_target_month(today)

# セレクトボックスの初期選択位置。デフォルト月が選択肢に含まれていれば
# その位置を、含まれていなければ末尾（＝最新の選択可能月）を初期値とする。
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
    # ヘルプ（記入例）モーダルはSC-08として別途実装予定。
    # Day4時点ではボタンのみ配置し、押下時は簡易メッセージを表示する。
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

header = attendance_data.header

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
# 編集切り替えボタン（土台のみ。実際の編集はDay5で実装）
# ============================================
edit_button_col, _ = st.columns([1, 4])
with edit_button_col:
    if st.button("編集する", disabled=is_locked, use_container_width=True):
        st.info("編集機能は Day5 で実装予定です。")

if is_locked:
    st.caption("この月はロックされています。編集する場合は管理者に解除を依頼してください。")

# ============================================
# プレビュー表示（読み取り専用テーブル、自動計算項目を含む全項目）
# 基本設計書3.5.3節「プレビュー時」列が○のものすべてを表示する。
# ============================================
st.subheader("勤怠データ（プレビュー）")


def _format_time(value) -> str:
    """(時, 分) のタプルを "H:MM" 表示に整形する。Noneは "--:--" とする。"""
    if value is None:
        return "--:--"
    hour, minute = value
    return f"{hour}:{minute:02d}"


def _format_plain(value) -> str:
    """自動計算項目等、そのまま表示してよい値の整形（Noneは空文字）。"""
    return "" if value is None else str(value)


table_rows = []

for entry in attendance_data.entries:

    # 日付が存在しない行は除外
    # （Excelテンプレートの31日分確保用）
    if entry.date_value is None:
        continue

    table_rows.append(
        {
            "日": entry.date_value,
            "曜日": entry.weekday or "",
            "休暇種類": entry.leave_type or "",
            "始業": _format_time(entry.start_time),
            "終業": _format_time(entry.end_time),
            "休憩1": _format_plain(entry.break_time_1),
            "休憩2": _format_plain(entry.break_time_2),
            "休憩3": _format_plain(entry.break_time_3),
            "離業": _format_time(entry.leave_time),
            "実働": _format_plain(entry.actual_work_time),
            "超勤": _format_plain(entry.overtime),
            "休出": _format_plain(entry.holiday_work),
            "深夜": _format_plain(entry.late_night),
            "自社工数内容": entry.work_note or "",
        }
    )


# ============================================
# 空データでもテーブル表示する
# ============================================

if not table_rows:

    # 想定外にExcelから日付が読めなかった場合でも
    # 表示崩れを防ぐため31日分生成
    for day in range(1, 32):
        table_rows.append(
            {
                "日": day,
                "曜日": "",
                "休暇種類": "",
                "始業": "--:--",
                "終業": "--:--",
                "休憩1": "",
                "休憩2": "",
                "休憩3": "",
                "離業": "--:--",
                "実働": "",
                "超勤": "",
                "休出": "",
                "深夜": "",
                "自社工数内容": "",
            }
        )


st.dataframe(
    table_rows,
    use_container_width=True,
    hide_index=True
)


st.caption(
    "※プレビュー時は休憩・実働・超勤・休日出勤・深夜等の"
    "自動計算項目も表示しています。"
)