"""
excel_adapter.py

【概要】
openpyxlを使ってExcelファイルの読み書きを行うアダプター。
このモジュールが担う役割は大きく3つ。

1. テンプレートファイルから月次ファイルを作る（シートコピー）
   テンプレートの全シートをコピーするのではなく、「原本」シートと
   「記入例」シートだけをコピーし、「変更履歴」シートは含めない
   （要件定義書2章・4.2節）。

2. セルの値だけを読み書きする（書式を絶対に壊さない）
   罫線・条件付き書式・フォント・数式などは、対象セルの `.value`
   だけを書き換える限りopenpyxlによって保持される。行や列を挿入・
   削除したり、シート全体を作り直すような操作は行わない
   （要件定義書4.3節・6章、基本設計書6.4節）。

3. 時刻データの変換（フロントの(時,分)タプル ⇔ Excelの"H:MM"文字列）
   Excel側の時刻セルは "9:30" のような、時の先頭にゼロを付けない
   半角の文字列として扱う（要件定義書4.3節）。フロント側は時・分を
   別々の数値として受け取るため、この変換をここに集約する
   （基本設計書6.6節）。

このモジュールは「1つのExcelファイルをどう読み書きするか」にのみ
関心を持ち、「そのファイルをOneDriveのどこから取得するか」は
関知しない（そちらは onedrive_adapter の責務。呼び出し側の
attendance_service が両者を組み合わせる）。

記入例シートの表示は、本モジュール単体ではなく
attendance_service.load_example（本モジュールのread_day_rows・
read_header_values・calc_auto_valuesを組み合わせ、自動計算列も
正しく再現した AttendanceData を返す）を呼び出し側が使う想定。
本モジュールに記入例専用の簡易読み取り関数は置かない
（原本シートの読み取りと処理経路を分けると、自動計算ロジックの
二重管理になるため）。
"""

import calendar
import datetime
import re
from dataclasses import dataclass
from typing import Optional

import openpyxl
from openpyxl.workbook.workbook import Workbook

# ============================================
# シート名の定数（要件定義書2章の用語定義に対応）
# ============================================
SHEET_HONBUN = "原本"       # 実データ入力シート
SHEET_KINYUREI = "記入例 "    # 入力例シート（保護済み、ヘルプ表示用）
SHEET_HENKOU_REKISHI = "変更履歴"  # コピー対象外

# テンプレートからコピーするシート（変更履歴は含めない）
_SHEETS_TO_COPY = (SHEET_HONBUN, SHEET_KINYUREI)

# ============================================
# セル位置の定数（要件定義書14章の調査結果・4.3節、基本設計書6.5節）
# ============================================
CELL_TARGET_MONTH = "C8"     # 対象月初日（YYYY/MM/01形式）
CELL_EMPLOYEE_ID = "O5"      # 社員番号
CELL_DEPARTMENT = "AH4"      # 部署名
CELL_FULL_NAME = "AH5"       # 氏名（全角スペース区切り）

# B,C,D 開始行, 終了行
DATE_START_ROW = 14
DATE_END_ROW = 44

# N～AP 開始行, 終了行
TIME_START_ROW = 9
TIME_END_ROW = 39

# 日次データの列（要件定義書4.3節・7.4節、基本設計書6.5節）
COL_DATE = "B"          # 日付
COL_WEEKDAY = "C"       # 曜日（Excel側の値をそのまま読み取る）
COL_LEAVE_TYPE = "D"    # 休暇種類
COL_START_TIME = "N"    # 始業時間
COL_END_TIME = "Q"      # 終業時間
COL_LEAVE_TIME = "AC"   # 離業時間（私用外出・自社業務）
COL_WORK_NOTE = "AP"    # 自社工数内容

# 自動計算項目の列（休憩時間・実働時間・超勤・休日出勤・深夜・深夜開始時間）
COL_BREAK_TIME_1 = "T"         # 休憩時間1
COL_BREAK_TIME_2 = "W"         # 休憩時間2
COL_BREAK_TIME_3 = "Z"         # 休憩時間3
COL_ACTUAL_WORK_TIME = "AF"    # 実働時間
COL_OVERTIME = "AI"            # 超勤（超過勤務）
COL_HOLIDAY_WORK = "AL"        # 休日出勤
COL_LATE_NIGHT = "AM"          # 深夜

# フロント入力対象外（要件定義書4.3節「フロント入力対象外とするセル範囲」）
NON_EDITABLE_RANGES = ("F11:G16", "F18:H21")

# 時刻セルの形式チェック用正規表現（基本設計書7.4節）
_TIME_PATTERN = re.compile(r"^([0-9]|1[0-9]|2[0-3]):[0-5][0-9]$")

# 半休扱いの休暇種類（休憩時間1・超勤の計算で共通して参照する）
_HALF_DAY_LEAVE_TYPES = ("午前半休", "午後半休")


# ============================================
# 6. 休暇種類の選択肢（Excel側データ入力規則と対応。要件定義書14章）
# ============================================
LEAVE_TYPE_OPTIONS = (
    "",       # 未選択（通常勤務）
    "有休",
    "午前半休",
    "午後半休",
    "長期連休",
    "特別休暇",
    "欠勤",
    "振替休暇",
    "休日出勤",
    "振替出勤",
    "稼働日",
    "休日",
)


class ExcelFormatError(Exception):
    """Excelファイルの形式が想定と異なる場合に送出する例外。"""


# ============================================
# 1. テンプレートからのシートコピー
# ============================================

def create_workbook_from_template(template_path: str) -> Workbook:
    """
    テンプレートファイルを読み込み、「原本」シートと「記入例」シートのみを
    持つワークブックを作って返す。「変更履歴」シートは含めない
    （要件定義書4.2節）。
    """
    workbook = openpyxl.load_workbook(template_path)

    missing = [s for s in _SHEETS_TO_COPY if s not in workbook.sheetnames]
    if missing:
        raise ExcelFormatError(
            f"テンプレートに必要なシートが見つかりません: {missing}"
        )

    for sheet_name in list(workbook.sheetnames):
        if sheet_name not in _SHEETS_TO_COPY:
            del workbook[sheet_name]

    return workbook


# ============================================
# 2. 月次ファイル生成時の自動セル設定
# ============================================

def set_monthly_header_cells(
    workbook: Workbook,
    target_month_first_day: datetime.date,
    employee_id: str,
    department: str,
    full_name_with_space: str,
) -> None:
    """
    月次ファイル生成時、「原本」シートのヘッダーセルを自動設定する
    （要件定義書4.2節）。
    """
    ws = workbook[SHEET_HONBUN]
    ws[CELL_TARGET_MONTH] = target_month_first_day
    ws[CELL_TARGET_MONTH].number_format = "yyyy/mm/dd"
    ws[CELL_EMPLOYEE_ID] = employee_id
    ws[CELL_DEPARTMENT] = department
    ws[CELL_FULL_NAME] = full_name_with_space


# ============================================
# 3. 時刻データの変換（基本設計書6.6節）
# ============================================

def time_tuple_to_excel_string(time_tuple: Optional[tuple[int, int]]) -> Optional[str]:
    """
    フロントから受け取る (時, 分) のタプルを、Excelの "H:MM" 形式の文字列
    （半角数字、時の先頭ゼロなし）に変換する。
    """
    if time_tuple is None:
        return None
    hour, minute = time_tuple
    if not (0 <= hour <= 23):
        raise ValueError(f"時は0〜23の範囲で指定してください: {hour}")
    if not (0 <= minute <= 59):
        raise ValueError(f"分は0〜59の範囲で指定してください: {minute}")
    return f"{hour}:{minute:02d}"


def excel_string_to_time_tuple(value) -> Optional[tuple[int, int]]:
    """
    Excelセルから読み取った値を (時, 分) のタプルに変換する。
    """
    if value is None:
        return None

    text = str(value).strip()

    if text in ("", ":", "："):
        return None

    if hasattr(value, "hour") and hasattr(value, "minute"):
        return (value.hour, value.minute)

    if isinstance(value, str):
        match = re.match(r"^(\d{1,2}):(\d{1,2})(?::\d{1,2})?$", value.strip())
        if match:
            return (int(match.group(1)), int(match.group(2)))
        raise ExcelFormatError(f"時刻セルの形式が不正です: {value!r}")

    raise ExcelFormatError(f"時刻セルの値を解釈できません: {value!r}")


def validate_excel_time_string(value: str) -> bool:
    """
    保存直前の最終防御チェック（要件定義書4.6節・基本設計書7.4節）。
    """
    return bool(_TIME_PATTERN.match(value))


def time_tuple_to_excel_time(time_tuple: Optional[tuple[int, int]]) -> Optional[datetime.time]:
    """
    フロントから受け取る (時, 分) のタプルを、Excelセルへ直接代入するための
    datetime.time オブジェクトに変換する。
    """
    if time_tuple is None:
        return None
    hour, minute = time_tuple
    if not (0 <= hour <= 23):
        raise ValueError(f"時は0〜23の範囲で指定してください: {hour}")
    if not (0 <= minute <= 59):
        raise ValueError(f"分は0〜59の範囲で指定してください: {minute}")
    return datetime.time(hour, minute)


# ============================================
# 4. 日次データの読み書き（原本シート・記入例シート共通）
# ============================================

@dataclass
class DayCellValues:
    """1日分の「原本シート上」の生データ（セル位置に対応する値そのもの）"""
    row: int
    time_row: int
    date_value: object
    weekday: Optional[str]
    leave_type: Optional[str]
    start_time: Optional[tuple[int, int]]
    end_time: Optional[tuple[int, int]]
    leave_time: Optional[tuple[int, int]]
    work_note: Optional[str]
    break_time_1: object
    break_time_2: object
    break_time_3: object
    actual_work_time: object
    overtime: object
    holiday_work: object
    late_night: object


@dataclass
class AttendanceHeaderValues:
    """勤怠表ヘッダー情報"""

    client_company_name: object
    client_department: object
    employee_id: object
    employee_name: object

    night_start_time: object

    scheduled_work_time: object

    morning_time: object
    afternoon_time: object

    break_time_1_start: object
    break_time_1_end: object

    break_time_2_start: object
    break_time_2_end: object

    break_time_3_start: object
    break_time_3_end: object


_WEEKDAY_LABELS_JA = ("月", "火", "水", "木", "金", "土", "日")

_DATE_STRING_PATTERN = re.compile(r"^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})")


def _calc_weekday_label(date_value) -> Optional[str]:
    """
    B列の日付値から日本語の曜日ラベル（月〜日）を算出する。
    """
    if date_value is None:
        return None

    if hasattr(date_value, "weekday"):
        return _WEEKDAY_LABELS_JA[date_value.weekday()]

    if isinstance(date_value, str):
        match = _DATE_STRING_PATTERN.match(date_value.strip())
        if match:
            year, month, day = (int(g) for g in match.groups())
            try:
                return _WEEKDAY_LABELS_JA[datetime.date(year, month, day).weekday()]
            except ValueError:
                return None

    return None


def read_day_rows(workbook: Workbook, sheet_name: str = SHEET_HONBUN) -> list[DayCellValues]:
    """
    指定シート（デフォルトは「原本」）の14〜44行目を読み取り、
    1日分ずつのDayCellValuesのリストとして返す。
    """
    ws = workbook[sheet_name]
    results: list[DayCellValues] = []

    target_month_first_day = read_target_month_first_day(workbook)

    for day in range(31):
        date_row = DATE_START_ROW + day
        time_row = TIME_START_ROW + day

        date_value = calc_date_for_row(target_month_first_day, date_row)
        leave_type = ws[f"{COL_LEAVE_TYPE}{date_row}"].value

        start_time = None
        end_time = None
        leave_time = None
        work_note = None
        break_time_1 = None
        break_time_2 = None
        break_time_3 = None
        actual_work_time = None
        overtime = None
        holiday_work = None
        late_night = None

        if time_row <= TIME_END_ROW:
            start_time = excel_string_to_time_tuple(ws[f"{COL_START_TIME}{time_row}"].value)
            end_time = excel_string_to_time_tuple(ws[f"{COL_END_TIME}{time_row}"].value)
            leave_time = excel_string_to_time_tuple(ws[f"{COL_LEAVE_TIME}{time_row}"].value)
            work_note = ws[f"{COL_WORK_NOTE}{time_row}"].value
            break_time_1 = ws[f"{COL_BREAK_TIME_1}{time_row}"].value
            break_time_2 = ws[f"{COL_BREAK_TIME_2}{time_row}"].value
            break_time_3 = ws[f"{COL_BREAK_TIME_3}{time_row}"].value
            actual_work_time = ws[f"{COL_ACTUAL_WORK_TIME}{time_row}"].value
            overtime = ws[f"{COL_OVERTIME}{time_row}"].value
            holiday_work = ws[f"{COL_HOLIDAY_WORK}{time_row}"].value
            late_night = ws[f"{COL_LATE_NIGHT}{time_row}"].value

        results.append(
            DayCellValues(
                row=date_row,
                time_row=time_row,
                date_value=date_value,
                weekday=_calc_weekday_label(date_value),
                leave_type=leave_type,
                start_time=start_time,
                end_time=end_time,
                leave_time=leave_time,
                work_note=work_note,
                break_time_1=break_time_1,
                break_time_2=break_time_2,
                break_time_3=break_time_3,
                actual_work_time=actual_work_time,
                overtime=overtime,
                holiday_work=holiday_work,
                late_night=late_night,
            )
        )

    return results


def write_day_cell(
    workbook: Workbook,
    date_row: int,
    time_row: int,
    *,
    leave_type: Optional[str] = None,
    start_time: Optional[tuple[int, int]] = None,
    end_time: Optional[tuple[int, int]] = None,
    leave_time: Optional[tuple[int, int]] = None,
    work_note: Optional[str] = None,
    set_leave_type: bool = False,
    set_start_time: bool = False,
    set_end_time: bool = False,
    set_leave_time: bool = False,
    set_work_note: bool = False,
) -> None:
    """
    「原本」シートの指定行に対し、指定された項目のみを更新する。
    """
    ws = workbook[SHEET_HONBUN]

    if set_leave_type:
        ws[f"{COL_LEAVE_TYPE}{date_row}"] = leave_type

    if set_start_time:
        _validate_time_tuple_format(start_time, "始業時間")
        ws[f"{COL_START_TIME}{time_row}"] = time_tuple_to_excel_time(start_time)

    if set_end_time:
        _validate_time_tuple_format(end_time, "終業時間")
        ws[f"{COL_END_TIME}{time_row}"] = time_tuple_to_excel_time(end_time)

    if set_leave_time:
        _validate_time_tuple_format(leave_time, "離業時間")
        ws[f"{COL_LEAVE_TIME}{time_row}"] = time_tuple_to_excel_time(leave_time)

    if set_work_note:
        ws[f"{COL_WORK_NOTE}{time_row}"] = work_note


def _validate_time_tuple_format(time_tuple: Optional[tuple[int, int]], field_label: str) -> None:
    """
    保存直前の最終防御チェック（要件定義書4.6節・基本設計書7.4節）。
    """
    if time_tuple is None:
        return
    excel_value = time_tuple_to_excel_string(time_tuple)
    if not validate_excel_time_string(excel_value):
        raise ExcelFormatError(f"{field_label}の形式が不正です: {excel_value!r}")


def read_target_month_first_day(workbook: Workbook) -> Optional[datetime.date]:
    """
    C8セル（対象月初日）の値を datetime.date として読み取る。
    """
    ws = workbook[SHEET_HONBUN]
    value = ws[CELL_TARGET_MONTH].value

    if value is None:
        return None

    if isinstance(value, datetime.date):
        if isinstance(value, datetime.datetime):
            return value.date()
        return value

    if isinstance(value, str):
        match = re.match(r"^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})", value.strip())
        if match:
            year, month, day = (int(g) for g in match.groups())
            try:
                return datetime.date(year, month, day)
            except ValueError:
                return None

    return None


def calc_date_for_row(target_month_first_day: datetime.date, row: int) -> Optional[datetime.date]:
    """
    B列の数式と等価な計算をPython側で行う。
    """
    if target_month_first_day is None:
        return None

    day_offset = row - DATE_START_ROW
    if day_offset < 0:
        return None

    days_in_month = calendar.monthrange(
        target_month_first_day.year, target_month_first_day.month
    )[1]
    day_number = day_offset + 1
    if day_number > days_in_month:
        return None

    return datetime.date(
        target_month_first_day.year, target_month_first_day.month, day_number
    )


def read_header_values(
    workbook: Workbook,
    sheet_name: str = SHEET_HONBUN,
) -> AttendanceHeaderValues:
    """
    表外の勤怠情報を読み取る。
    """
    ws = workbook[sheet_name]

    return AttendanceHeaderValues(
        client_company_name=ws["O4"].value,
        client_department=ws["AH4"].value,
        employee_id=ws["O5"].value,
        employee_name=ws["AH5"].value,
        night_start_time=ws["D10"].value,
        scheduled_work_time=ws["G11"].value,
        morning_time=ws["G13"].value,
        afternoon_time=ws["G15"].value,
        break_time_1_start=ws["G19"].value,
        break_time_1_end=ws["H19"].value,
        break_time_2_start=ws["G20"].value,
        break_time_2_end=ws["H20"].value,
        break_time_3_start=ws["G21"].value,
        break_time_3_end=ws["H21"].value,
    )


def format_time_value_no_seconds(value) -> str:
    """
    ヘッダー領域の時刻セル値を、秒なしの "H:MM" 形式の表示用文字列に整形する。
    """
    if value is None:
        return "None"

    if isinstance(value, datetime.timedelta):
        total_minutes = int(value.total_seconds() // 60)
        hour, minute = divmod(total_minutes, 60)
        return f"{hour}:{minute:02d}"

    if hasattr(value, "hour") and hasattr(value, "minute"):
        return f"{value.hour}:{value.minute:02d}"

    if isinstance(value, str):
        text = value.strip().replace("：", ":")
        match = re.match(r"^(\d{1,2}):(\d{1,2})(?::\d{1,2}(?:\.\d+)?)?$", text)
        if match:
            return f"{int(match.group(1))}:{int(match.group(2)):02d}"

    return str(value)


def _time_tuple_to_decimal_hours(value: Optional[tuple[int, int]]) -> Optional[float]:
    """
    (時, 分) タプルを、Excelの時刻シリアル値と同じ考え方の「時間の小数表現」に変換する。
    """
    if value is None:
        return None
    hour, minute = value
    return hour + minute / 60.0


def _decimal_hours_to_time_tuple(value: Optional[float]) -> Optional[tuple[int, int]]:
    """
    小数表現の時間を (時, 分) タプルに戻す（表示用）。負値もそのまま許容する。
    """
    if value is None:
        return None
    sign = -1 if value < 0 else 1
    total_minutes = round(abs(value) * 60)
    hour, minute = divmod(total_minutes, 60)
    if sign < 0 or hour > 0:
        return (sign * hour, minute)
    return (0, minute)


@dataclass
class AutoCalculatedValues:
    """
    自動計算列の再計算結果を表すDTO。
    """
    break_time_1: Optional[tuple[int, int]]
    break_time_2: Optional[tuple[int, int]]
    break_time_3: Optional[tuple[int, int]]
    actual_work_time: Optional[tuple[int, int]]
    overtime: Optional[tuple[int, int]]
    holiday_work: Optional[tuple[int, int]]
    late_night: Optional[tuple[int, int]]


def calc_auto_values(
    leave_type: Optional[str],
    start_time: Optional[tuple[int, int]],
    end_time: Optional[tuple[int, int]],
    break_time_1_input: Optional[tuple[int, int]],
    break_time_2: Optional[tuple[int, int]],
    break_time_3: Optional[tuple[int, int]],
    leave_time: Optional[tuple[int, int]],
    scheduled_work_time: Optional[tuple[int, int]],
    night_start_time: Optional[tuple[int, int]],
    break_1_start: Optional[tuple[int, int]],
    break_1_end: Optional[tuple[int, int]],
) -> AutoCalculatedValues:
    """
    Excel側の自動計算列の数式をPython側で再現し、計算結果を返す。
    """
    start = _time_tuple_to_decimal_hours(start_time)
    end = _time_tuple_to_decimal_hours(end_time)
    b2 = _time_tuple_to_decimal_hours(break_time_2)
    b3 = _time_tuple_to_decimal_hours(break_time_3)
    leave = _time_tuple_to_decimal_hours(leave_time)
    scheduled = _time_tuple_to_decimal_hours(scheduled_work_time)
    night_start = _time_tuple_to_decimal_hours(night_start_time)
    break_1_start_dec = _time_tuple_to_decimal_hours(break_1_start)
    break_1_end_dec = _time_tuple_to_decimal_hours(break_1_end)

    if leave_type in _HALF_DAY_LEAVE_TYPES:
        break_1_decimal = 0.0
    elif start is None or end is None:
        break_1_decimal = None
    elif break_1_start_dec is None or break_1_end_dec is None:
        break_1_decimal = None
    else:
        overlap_end = min(max(end, break_1_start_dec), break_1_end_dec)
        overlap_start = max(min(start, break_1_end_dec), break_1_start_dec)
        break_1_decimal = overlap_end - overlap_start
        if break_1_decimal < 0:
            break_1_decimal = 0.0

    if start is None:
        actual_work_decimal = None
    else:
        deduction = 0.0
        if break_1_decimal is not None:
            deduction += break_1_decimal
        if b2 is not None:
            deduction += b2
        if b3 is not None:
            deduction += b3
        if leave is not None:
            deduction += leave
        actual_work_decimal = (end or 0.0) - start - deduction

    if leave_type in _HALF_DAY_LEAVE_TYPES:
        overtime_decimal = 0.0
    elif leave_type == "休日出勤":
        overtime_decimal = None
    elif start is None:
        overtime_decimal = None
    else:
        overtime_decimal = (
            None if actual_work_decimal is None or scheduled is None
            else actual_work_decimal - scheduled
        )

    if leave_type == "休日出勤":
        holiday_work_decimal = actual_work_decimal
    else:
        holiday_work_decimal = None

    if start is None:
        late_night_decimal = None
    elif end is None or night_start is None or end <= night_start:
        late_night_decimal = None
    else:
        late_night_decimal = end - night_start

    return AutoCalculatedValues(
        break_time_1=_decimal_hours_to_time_tuple(break_1_decimal),
        break_time_2=break_time_2,
        break_time_3=break_time_3,
        actual_work_time=_decimal_hours_to_time_tuple(actual_work_decimal),
        overtime=_decimal_hours_to_time_tuple(overtime_decimal),
        holiday_work=_decimal_hours_to_time_tuple(holiday_work_decimal),
        late_night=_decimal_hours_to_time_tuple(late_night_decimal),
    )


def header_time_value_to_tuple(value) -> Optional[tuple[int, int]]:
    """
    ヘッダー領域の時刻セル値を (時, 分) タプルに変換する。
    """
    if value is None:
        return None

    if isinstance(value, datetime.timedelta):
        total_minutes = int(value.total_seconds() // 60)
        hour, minute = divmod(total_minutes, 60)
        return (hour, minute)

    if hasattr(value, "hour") and hasattr(value, "minute"):
        return (value.hour, value.minute)

    if isinstance(value, str):
        text = value.strip().replace("：", ":")
        match = re.match(r"^(\d{1,2}):(\d{1,2})(?::\d{1,2}(?:\.\d+)?)?$", text)
        if match:
            return (int(match.group(1)), int(match.group(2)))

    return None


# ============================================
# 5. ファイル入出力ヘルパー
# ============================================

def load_workbook_from_path(path: str) -> Workbook:
    """指定パスのExcelファイルを読み込む（数式そのものを保持。書き込み用）。"""
    return openpyxl.load_workbook(path)


def load_workbook_from_path_for_display(path: str) -> Workbook:
    """
    指定パスのExcelファイルを、表示用（計算結果キャッシュ値）として読み込む。
    保存（save）用途には絶対に使わないこと（数式が失われる）。
    """
    return openpyxl.load_workbook(path, data_only=True)


def save_workbook_to_path(workbook: Workbook, path: str) -> None:
    """ワークブックを指定パスへ保存する。"""
    workbook.save(path)