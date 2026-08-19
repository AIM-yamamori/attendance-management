"""
attendance_service.py

【概要】
勤怠データの読み書き（Excel連携含む）を担うモジュール（基本設計書5.2.4節）。
"""

import calendar
import datetime
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from adapters import excel_adapter, onedrive_adapter
from services import lock_service

# ============================================
# ファイル名・フォルダ名の生成規則（要件定義書4.2節・5.1節）
# ============================================

_DEFAULT_COMPANY_NAME = "【会社名】"
TEMPLATE_FILE_NAME_SUFFIX = "勤務実績管理表_テンプレート.xlsx"


def _company_name() -> str:
    return os.environ.get("COMPANY_NAME", _DEFAULT_COMPANY_NAME)


def get_template_relative_path() -> str:
    return f"{_company_name()}{TEMPLATE_FILE_NAME_SUFFIX}"


def build_monthly_file_name(target_month: str, full_name_no_space: str) -> str:
    return f"{_company_name()}勤務実績管理表_{target_month}_{full_name_no_space}.xlsx"


def build_monthly_file_relative_path(target_month: str, full_name_no_space: str) -> str:
    file_name = build_monthly_file_name(target_month, full_name_no_space)
    return f"{target_month}/{file_name}"


def target_month_to_first_day(target_month: str) -> str:
    if len(target_month) != 6 or not target_month.isdigit():
        raise ValueError(f"target_monthは YYYYMM 形式で指定してください: {target_month!r}")
    year = target_month[:4]
    month = target_month[4:6]
    return f"{year}/{month}/01"


def _days_in_month(target_month: str) -> int:
    year = int(target_month[:4])
    month = int(target_month[4:6])
    return calendar.monthrange(year, month)[1]


def target_month_to_date(target_month: str) -> datetime.date:
    if len(target_month) != 6 or not target_month.isdigit():
        raise ValueError(f"target_monthは YYYYMM 形式で指定してください: {target_month!r}")
    year = int(target_month[:4])
    month = int(target_month[4:6])
    return datetime.date(year, month, 1)


# ============================================
# 対象年月の選択制御（基本設計書3.5.2節）
# ============================================

def _add_months(year: int, month: int, delta: int) -> tuple[int, int]:
    total = (year * 12 + (month - 1)) + delta
    new_year, new_month0 = divmod(total, 12)
    return new_year, new_month0 + 1


def build_selectable_months(
    today: datetime.date,
    service_start_month: Optional[str],
) -> list[str]:
    this_year, this_month = today.year, today.month
    default_month = f"{this_year:04d}{this_month:02d}"

    if not service_start_month:
        return [default_month]

    next_year, next_month = _add_months(this_year, this_month, 1)
    upper_bound = f"{next_year:04d}{next_month:02d}"
    lower_bound = service_start_month

    if lower_bound > upper_bound:
        return [upper_bound]

    months: list[str] = []
    year, month = int(lower_bound[:4]), int(lower_bound[4:6])
    current = lower_bound
    while current <= upper_bound:
        months.append(current)
        year, month = _add_months(year, month, 1)
        current = f"{year:04d}{month:02d}"

    return months


def default_target_month(today: datetime.date) -> str:
    return f"{today.year:04d}{today.month:02d}"


def format_month_label(target_month: str) -> str:
    return f"{target_month[:4]}年{target_month[4:6]}月"


# ============================================
# データ型
# ============================================

@dataclass
class WorkbookHandle:
    """「取得・生成した月次Excelファイル」を表すハンドル。"""
    local_path: str
    onedrive_relative_path: str
    target_month: str
    employee_id: str


class LockedError(Exception):
    """
    保存直前のロック再チェックで、ロック状態が保存を許可しない状態に
    なっていた場合に送出する例外（基本設計書8.2節・8.3節）。
    - 一般ユーザーの保存：ロック済み（True）なら送出
    - admin の保存：ロック解除済み（False）なら送出
    """


@dataclass
class DayEntry:
    """フロント（画面）向けの1日分の勤怠データ。基本設計書5.3節DTOに対応。"""
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
class AttendanceTotals:
    """
    月末合計欄（原本シート42・43行目相当）。基本設計書には明記されて
    いないが、要件定義書7.3節「合計欄」に対応する項目。

    Excel側は42行目・43行目に SUM 数式が入っているが、実働時間等の
    自動計算項目と同様、保存直後は数式キャッシュが古いままになる
    問題があるため、Python側で日別のDayEntry（自動計算済み）から
    再集計する（Excelファイル自体の数式・書式には一切触れない）。

    Q42 定時 = AF42(実働合計) - AI43(超勤+休日出勤合計)
    T42 休憩時間1合計 = 各日の休憩時間1（DayEntry.break_time_1）の合計
    W42 休憩時間2合計 = 各日の休憩時間2（DayEntry.break_time_2）の合計
    Z42 休憩時間3合計 = 各日の休憩時間3（DayEntry.break_time_3）の合計
    AC42 離業時間合計 = 各日の離業時間（DayEntry.leave_time）の合計
    AF42 実働時間合計 = 各日の実働時間（DayEntry.actual_work_time）の合計
    AI42 超勤合計 = 各日の超勤（DayEntry.overtime）の合計
    AL42 休日出勤合計 = 各日の休日出勤（DayEntry.holiday_work）の合計
    AI43 超勤+休日出勤合計 = AI42 + AL42
    AM42 深夜合計 = 各日の深夜（DayEntry.late_night）の合計
    """
    scheduled_total: Optional[tuple[int, int]]        # Q42 定時
    break_time_1_total: Optional[tuple[int, int]]      # T42
    break_time_2_total: Optional[tuple[int, int]]      # W42
    break_time_3_total: Optional[tuple[int, int]]      # Z42
    leave_time_total: Optional[tuple[int, int]]        # AC42
    actual_work_time_total: Optional[tuple[int, int]]  # AF42
    overtime_total: Optional[tuple[int, int]]          # AI42
    holiday_work_total: Optional[tuple[int, int]]      # AL42
    overtime_plus_holiday_total: Optional[tuple[int, int]]  # AI43
    late_night_total: Optional[tuple[int, int]]        # AM42


@dataclass
class AttendanceData:
    """1ユーザー・1ヶ月分の勤怠データ全体（基本設計書5.2.4節）。"""
    target_month: str
    employee_id: str
    header: excel_adapter.AttendanceHeaderValues
    entries: list[DayEntry]
    totals: AttendanceTotals


@dataclass
class DayEditInput:
    """
    編集フォームから受け取る1日分の入力値。
    """
    row: int
    time_row: int
    date_value: object
    weekday: Optional[str]
    leave_type: str
    start_time: Optional[tuple[int, int]]
    end_time: Optional[tuple[int, int]]
    leave_time: Optional[tuple[int, int]]
    work_note: str


def _sum_time_tuples(values: list) -> Optional[tuple[int, int]]:
    """
    (時, 分) タプルのリストを合計する。全てNoneなら合計結果もNoneとする
    （Excelの SUM 関数は空セルを0として扱うため、実際には1件でも値が
    あれば合計を返すのが自然だが、月内データが全く無い場合はNoneのまま
    「--:--」表示にする）。
    """
    non_none_values = [v for v in values if v is not None]
    if not non_none_values:
        return None

    total_minutes = 0
    for hour, minute in non_none_values:
        total_minutes += hour * 60 + minute

    sign = -1 if total_minutes < 0 else 1
    abs_minutes = abs(total_minutes)
    hour, minute = divmod(abs_minutes, 60)
    return (sign * hour, minute)


def _calc_totals(entries: list[DayEntry]) -> AttendanceTotals:
    """
    月内全日分のDayEntry（既に自動計算済みの値を持つ）から、
    原本シート42・43行目相当の合計欄を再集計する。
    """
    break_1_total = _sum_time_tuples([e.break_time_1 for e in entries])
    break_2_total = _sum_time_tuples([e.break_time_2 for e in entries])
    break_3_total = _sum_time_tuples([e.break_time_3 for e in entries])
    leave_total = _sum_time_tuples([e.leave_time for e in entries])
    actual_work_total = _sum_time_tuples([e.actual_work_time for e in entries])
    overtime_total = _sum_time_tuples([e.overtime for e in entries])
    holiday_work_total = _sum_time_tuples([e.holiday_work for e in entries])
    late_night_total = _sum_time_tuples([e.late_night for e in entries])

    # AI43 = AI42（超勤合計） + AL42（休日出勤合計）
    overtime_plus_holiday_total = _sum_time_tuples([overtime_total, holiday_work_total])

    # Q42 = AF42（実働合計） - AI43（超勤+休日出勤合計）
    if actual_work_total is None:
        scheduled_total = None
    else:
        deduction = overtime_plus_holiday_total or (0, 0)
        actual_minutes = actual_work_total[0] * 60 + actual_work_total[1]
        deduction_minutes = deduction[0] * 60 + deduction[1]
        diff_minutes = actual_minutes - deduction_minutes
        sign = -1 if diff_minutes < 0 else 1
        abs_minutes = abs(diff_minutes)
        hour, minute = divmod(abs_minutes, 60)
        scheduled_total = (sign * hour, minute)

    return AttendanceTotals(
        scheduled_total=scheduled_total,
        break_time_1_total=break_1_total,
        break_time_2_total=break_2_total,
        break_time_3_total=break_3_total,
        leave_time_total=leave_total,
        actual_work_time_total=actual_work_total,
        overtime_total=overtime_total,
        holiday_work_total=holiday_work_total,
        overtime_plus_holiday_total=overtime_plus_holiday_total,
        late_night_total=late_night_total,
    )


# ============================================
# 月次ファイル生成・取得（6.3節）
# ============================================

def _local_work_dir() -> Path:
    work_dir = Path(tempfile.gettempdir()) / "attendance-system-work"
    work_dir.mkdir(parents=True, exist_ok=True)
    return work_dir


def get_or_create_monthly_file(
    employee_id: str,
    target_month: str,
    full_name_no_space: str,
    full_name_with_space: str,
    department: str,
) -> WorkbookHandle:
    """
    対象ユーザー・対象月の月次Excelファイルを取得する（基本設計書6.3節）。
    既に存在すればそのまま使い、存在しなければテンプレートから新規生成する。
    """
    onedrive_adapter.ensure_folder(target_month)

    relative_path = build_monthly_file_relative_path(target_month, full_name_no_space)
    local_path = str(_local_work_dir() / f"{employee_id}_{target_month}.xlsx")

    if onedrive_adapter.file_exists(relative_path):
        onedrive_adapter.download_file(relative_path, local_path)
        return WorkbookHandle(
            local_path=local_path,
            onedrive_relative_path=relative_path,
            target_month=target_month,
            employee_id=employee_id,
        )

    workbook = _create_new_monthly_workbook(
        target_month=target_month,
        employee_id=employee_id,
        department=department,
        full_name_with_space=full_name_with_space,
    )
    excel_adapter.save_workbook_to_path(workbook, local_path)
    onedrive_adapter.upload_file(local_path, relative_path, overwrite=False)
    onedrive_adapter.download_file(relative_path, local_path)

    return WorkbookHandle(
        local_path=local_path,
        onedrive_relative_path=relative_path,
        target_month=target_month,
        employee_id=employee_id,
    )


def _create_new_monthly_workbook(
    target_month: str,
    employee_id: str,
    department: str,
    full_name_with_space: str,
):
    template_relative_path = get_template_relative_path()

    if not onedrive_adapter.file_exists(template_relative_path):
        raise FileNotFoundError(
            f"テンプレートファイルがOneDrive上に見つかりません: {template_relative_path}"
        )

    local_template_path = str(_local_work_dir() / "_template_download.xlsx")
    onedrive_adapter.download_file(template_relative_path, local_template_path)

    workbook = excel_adapter.create_workbook_from_template(local_template_path)

    target_month_first_day = target_month_to_date(target_month)
    excel_adapter.set_monthly_header_cells(
        workbook,
        target_month_first_day=target_month_first_day,
        employee_id=employee_id,
        department=department,
        full_name_with_space=full_name_with_space,
    )

    return workbook


# ============================================
# データの読み取り（プレビュー・記入例ヘルプ共通）
# ============================================

def _load_attendance_from_sheet(handle: WorkbookHandle, sheet_name: str) -> AttendanceData:
    """
    指定シート（「原本」または「記入例」）の全項目を読み取り、
    表示用データとして返す共通処理（load_attendance / load_example から
    使う。両者はどのシートを読むか以外の処理は完全に同じ）。
    """
    workbook = excel_adapter.load_workbook_from_path_for_display(handle.local_path)

    header = excel_adapter.read_header_values(workbook, sheet_name=sheet_name)
    day_rows = excel_adapter.read_day_rows(workbook, sheet_name=sheet_name)

    entries = [_day_cell_values_to_entry(d, header) for d in day_rows]
    entries = _filter_entries_to_month_days(entries, handle.target_month)

    totals = _calc_totals(entries)

    return AttendanceData(
        target_month=handle.target_month,
        employee_id=handle.employee_id,
        header=header,
        entries=entries,
        totals=totals,
    )


def load_attendance(handle: WorkbookHandle) -> AttendanceData:
    """
    「原本」シートの全項目を読み取り、プレビュー用データとして返す
    （基本設計書5.2.4節）。
    """
    return _load_attendance_from_sheet(handle, excel_adapter.SHEET_HONBUN)


def load_example(handle: WorkbookHandle) -> AttendanceData:
    """「記入例」シートの全項目を、load_attendanceと同一のセル対応表で読み取る（SC-08用）。"""
    return _load_attendance_from_sheet(handle, excel_adapter.SHEET_KINYUREI)


def _filter_entries_to_month_days(entries: list[DayEntry], target_month: str) -> list[DayEntry]:
    """対象月の実日数を超える行を除外する（3.5.4節#5「日付が空欄の行は生成しない」対応）。"""
    max_day = _days_in_month(target_month)
    filtered = []
    for entry in entries:
        day_number = _extract_day_number(entry.date_value)
        if day_number is not None and day_number > max_day:
            continue
        filtered.append(entry)
    return filtered


def _extract_day_number(date_value) -> Optional[int]:
    if date_value is None:
        return None
    if hasattr(date_value, "day"):
        return date_value.day
    if isinstance(date_value, str):
        match = re.match(r"^\d{4}[/\-]\d{1,2}[/\-](\d{1,2})", date_value.strip())
        if match:
            return int(match.group(1))
    return None


def _day_cell_values_to_entry(d, header: excel_adapter.AttendanceHeaderValues) -> DayEntry:
    """
    excel_adapter.DayCellValues を、画面向けの DayEntry へ変換する共通処理。
    自動計算項目（休憩1・実働・超勤・休日出勤・深夜）は、Excelの数式
    キャッシュ値ではなく、excel_adapter.calc_auto_values による
    Python側の再計算結果を使う（保存直後の表示ズレ対策）。
    休憩時間2・3は現時点では自動計算対象外のため、読み取った入力値を
    そのまま使う。
    """
    auto = excel_adapter.calc_auto_values(
        leave_type=d.leave_type,
        start_time=d.start_time,
        end_time=d.end_time,
        break_time_1_input=d.break_time_1,
        break_time_2=d.break_time_2,
        break_time_3=d.break_time_3,
        leave_time=d.leave_time,
        scheduled_work_time=excel_adapter.header_time_value_to_tuple(header.scheduled_work_time),
        night_start_time=excel_adapter.header_time_value_to_tuple(header.night_start_time),
        break_1_start=excel_adapter.header_time_value_to_tuple(header.break_time_1_start),
        break_1_end=excel_adapter.header_time_value_to_tuple(header.break_time_1_end),
    )

    return DayEntry(
        row=d.row,
        time_row=d.time_row,
        date_value=d.date_value,
        weekday=d.weekday,
        leave_type=d.leave_type,
        start_time=d.start_time,
        end_time=d.end_time,
        leave_time=d.leave_time,
        work_note=d.work_note,
        break_time_1=auto.break_time_1,
        break_time_2=auto.break_time_2,
        break_time_3=auto.break_time_3,
        actual_work_time=auto.actual_work_time,
        overtime=auto.overtime,
        holiday_work=auto.holiday_work,
        late_night=auto.late_night,
    )


# ============================================
# 編集・保存（基本設計書5.2.4節・8.2節・8.3節）
# ============================================

def save_attendance(
    handle: WorkbookHandle,
    edits: list[DayEditInput],
    actor_role: str,
) -> None:
    """
    編集フォームの入力内容を「原本」シートへまとめて反映し、OneDriveへ
    保存する。呼び出し前提として、バリデーション（validate_all）は
    画面側で既に実施済みであること（8.2節①・8.3節①）。

    本関数が担うのは②ロック再チェック・③保存の部分（8.2節・8.3節）：
    - actor_role="general"（一般ユーザー）：is_locked()がTrue
      （ロック済み）ならLockedErrorを送出し保存しない。
    - actor_role="admin"：is_locked()がFalse（ロック解除済み）なら
      LockedErrorを送出し保存しない（admin編集はロック中の月のみ
      可能というSC-03の仕様のため。3.6節）。

    引数:
        handle: 保存対象ファイルのハンドル
        edits: 保存する全日分の編集内容
        actor_role: "general" または "admin"
    """
    current_locked = lock_service.is_locked(handle.target_month, handle.employee_id)

    if actor_role == "general" and current_locked:
        raise LockedError(
            "この月はロックされました。入力内容は保存できません。"
        )
    if actor_role == "admin" and not current_locked:
        raise LockedError(
            "この月のロックが解除されています。保存できません。"
        )

    workbook = excel_adapter.load_workbook_from_path(handle.local_path)

    for edit in edits:
        excel_adapter.write_day_cell(
            workbook,
            date_row=edit.row,
            time_row=edit.time_row,
            leave_type=edit.leave_type or None,
            start_time=edit.start_time,
            end_time=edit.end_time,
            leave_time=edit.leave_time,
            work_note=edit.work_note or None,
            set_leave_type=True,
            set_start_time=True,
            set_end_time=True,
            set_leave_time=True,
            set_work_note=True,
        )

    excel_adapter.save_workbook_to_path(workbook, handle.local_path)
    onedrive_adapter.upload_file(
        handle.local_path, handle.onedrive_relative_path, overwrite=True
    )