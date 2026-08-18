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
"""

import re
from dataclasses import dataclass
from typing import Optional

import copy
import calendar
import datetime

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
# 【未確定】要件定義書14章の調査結果には、これら自動計算セルの具体的な
# 列位置が明記されていない（N9:S39・K9:AP39 の範囲に条件付き書式がある
# ことのみ判明している）。そのため、実物のテンプレートを精査するまでの
# 仮の位置として、休暇種類・始業・終業・離業列の近傍に割り当てている。
# プレビュー画面（SC-02）は「原本シートに現在記載されている内容を
# そのまま読み取り専用で表示する」（要件定義書7.4節）ものであり、
# ユーザーが値を入力する項目ではないため、実装初期のテンプレート精査で
# 列位置が確定した時点でこの定数群だけを差し替えれば、読み取りロジック
# （read_day_rowsおよび呼び出し側）に影響なく修正できる。
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


# ============================================
# 6. 休暇種類の選択肢（Excel側データ入力規則と対応。要件定義書14章）
# ============================================
# 【未確定】Excel側のプルダウン（データ入力規則）に設定されている
# 実際の選択肢文言は、テンプレート精査により後日確定する想定。
# それまでの仮の選択肢として、一般的な勤怠管理表で使われる項目を
# 設定している。確定次第この定数だけを差し替えれば、編集画面
# （02_勤怠入力_編集.py等）のロジックには影響しない。
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

    従来はセル単位で新規ブックへ値・書式をコピーする方式を取っていたが、
    その方式では列幅・行高・非表示行/列・グループ化（アウトライン）・
    シートのタブ色・ウィンドウ枠固定・印刷設定・ズーム倍率など、
    openpyxlが個別のプロパティとして明示的にコピーしていない設定が
    漏れ落ちるリスクがあった。

    そのため、テンプレートのワークブック自体をそのまま読み込み、
    コピー対象外のシート（「変更履歴」）だけを削除する方式に変更する。
    この方式であれば、ブック内のすべての書式・表示設定がopenpyxlの
    個別実装に依存せず丸ごと保持される（要件定義書4.2節・6章、
    基本設計書6.4節）。
    """
    workbook = openpyxl.load_workbook(template_path)

    missing = [s for s in _SHEETS_TO_COPY if s not in workbook.sheetnames]
    if missing:
        raise ExcelFormatError(
            f"テンプレートに必要なシートが見つかりません: {missing}"
        )

    # コピー対象（原本・記入例）以外のシートをすべて削除する。
    # 「変更履歴」シートを除外する（要件定義書4.2節）だけでなく、
    # 将来テンプレートに想定外のシートが追加された場合にも安全なよう、
    # _SHEETS_TO_COPY に含まれないシートは一律削除する。
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

    引数:
        target_month_first_day: 対象月の初日を表す datetime.date オブジェクト
            （例: datetime.date(2026, 8, 1)）。文字列ではなくdateオブジェクトを
            渡すことで、Excel側で本来の日付シリアル値として認識され、
            テキスト扱い（'付き文字列相当）になることを防ぐ。
        employee_id: 社員番号（そのまま文字列として設定）
        department: 部署名
        full_name_with_space: 姓名を全角スペースで結合した氏名
    """
    ws = workbook[SHEET_HONBUN]
    ws[CELL_TARGET_MONTH] = target_month_first_day
    # テンプレート側で元々設定されている表示形式（例: "yyyy/mm/dd"）が
    # あればそれを尊重するが、念のため未設定・崩れていた場合に備えて
    # 明示的に指定しておく。
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
    time_tuple が None の場合は None を返す（未入力＝セルを空にする）。

    例: (9, 30) -> "9:30"    (18, 0) -> "18:0"は誤り、"18:00"が正しい
        → 分は必ず2桁ゼロ埋め、時は1桁でもそのまま。
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
    セルが空の場合は None を返す。
    "9:30" のような文字列だけでなく、Excelが内部的に時刻型
    （datetime.time等）として保持している場合にも対応する。

    秒は画面表示・業務データとして不要なため、値に秒が含まれていても
    切り捨てて (時, 分) のみを返す（要件定義書4.3節）。
    """
    if value is None:
        return None

    text = str(value).strip()

    # 空欄として扱う値
    if text in ("", ":", "："):
        return None

    # openpyxlがExcelの時刻フォーマットを自動的にdatetime.time等として
    # 返してくる場合がある（セルの表示形式による）。秒は切り捨てる。
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return (value.hour, value.minute)

    if isinstance(value, str):
        # "9:30" 形式に加え、"9:30:00" のような秒付き形式も許容し、
        # 秒部分は読み捨てる。
        match = re.match(r"^(\d{1,2}):(\d{1,2})(?::\d{1,2})?$", value.strip())
        if match:
            return (int(match.group(1)), int(match.group(2)))
        raise ExcelFormatError(f"時刻セルの形式が不正です: {value!r}")

    raise ExcelFormatError(f"時刻セルの値を解釈できません: {value!r}")


def validate_excel_time_string(value: str) -> bool:
    """
    保存直前の最終防御チェック（要件定義書4.6節・基本設計書7.4節）。
    "H:MM" 形式（時は0〜23、先頭ゼロなし、分は0〜59で2桁固定）に
    一致するかを確認する。
    """
    return bool(_TIME_PATTERN.match(value))


def time_tuple_to_excel_time(time_tuple: Optional[tuple[int, int]]) -> Optional[datetime.time]:
    """
    フロントから受け取る (時, 分) のタプルを、Excelセルへ直接代入するための
    datetime.time オブジェクトに変換する。

    (時, 分) の文字列（"9:30"等）をそのままセルに代入すると、Excel側では
    テキストとして解釈され、'（アポストロフィ）付き文字列と同等の状態に
    なってしまう。datetime.time オブジェクトを代入することで、Excel側で
    本来の時刻シリアル値として認識される（C8への日付代入と同じ理由。
    要件定義書4.2節・4.3節）。

    time_tuple が None の場合は None を返す（未入力＝セルを空にする）。
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
    row: int          # 日付欄の行番号（B・D列。DATE_START_ROW起点）
    time_row: int      # 時刻欄の行番号（N・Q・AC・AP等。TIME_START_ROW起点）
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

    client_company_name: object   # 派遣先企業名 O4
    client_department: object     # 派遣先部署名 AH4
    employee_id: object           # 社員番号 O5
    employee_name: object         # 氏名 AH5

    night_start_time: object      # 深夜開始時間 D10

    scheduled_work_time: object   # 所定 G11

    morning_time: object          # 午前 G13
    afternoon_time: object        # 午後 G15

    break_time_1_start: object    # G19
    break_time_1_end: object      # H19

    break_time_2_start: object    # G20
    break_time_2_end: object      # H20

    break_time_3_start: object    # G21
    break_time_3_end: object      # H21


_WEEKDAY_LABELS_JA = ("月", "火", "水", "木", "金", "土", "日")

_DATE_STRING_PATTERN = re.compile(r"^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})")


def _calc_weekday_label(date_value) -> Optional[str]:
    """
    B列の日付値から日本語の曜日ラベル（月〜日）を算出する。

    B列は datetime.date / datetime.datetime 型で保持されている場合と、
    文字列（例: "2026/08/01"）として保持されている場合の両方が
    あり得るため、どちらにも対応する。
    """
    if date_value is None:
        return None

    # datetime.date / datetime.datetime 型
    if hasattr(date_value, "weekday"):
        return _WEEKDAY_LABELS_JA[date_value.weekday()]

    # 文字列（"2026/08/01" や "2026-08-01" 等）
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

    日付・曜日については、B列・C列の数式やそのキャッシュ値には
    一切依存しない。B列は "=+C8" / "=+B{n-1}+1"、C列は
    "=TEXT(B{n},\"aaa\")" という数式だが、生成直後（Excelで一度も
    開かれていない）ファイルはこれらのキャッシュが存在せず、
    data_only=True読み込みではNoneになってしまう。
    Excelファイルへの書き込みや数式変更は禁止のため、代わりに
    C8セル（Python側が直接書き込む実値）を起点に、Python側だけで
    独立して日付・曜日を計算する（calc_date_for_row /
    _calc_weekday_label）。これによりB列・C列の数式・書式には
    一切触れずに正しい表示を実現する。
    """
    ws = workbook[sheet_name]
    results: list[DayCellValues] = []

    target_month_first_day = read_target_month_first_day(workbook)

    for day in range(31):
        date_row = DATE_START_ROW + day
        time_row = TIME_START_ROW + day

        # B列・C列の値そのものは読み取らず、C8起点の独立計算を使う
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

    時刻項目（始業・終業・離業）はExcel側に本来の時刻シリアル値として
    入力されるよう、文字列ではなく datetime.time オブジェクトを代入する
    （'付き文字列相当になることを防ぐため）。保存直前の形式チェックは、
    文字列変換版（time_tuple_to_excel_string）を使って別途行う
    （validate_excel_time_stringは文字列を前提とした正規表現チェックの
    ため、セルへの代入とは別に検証用の文字列を都度生成する）。
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
    (時, 分) タプルを一旦 "H:MM" 形式の文字列に変換し、正規表現で
    形式が正しいかを確認する。datetime.time自体は時・分の範囲を
    コンストラクタで検証済みだが、既存のvalidate_excel_time_string
    （文字列前提のチェック）をそのまま活かすため、ここで文字列化して
    再確認する。
    """
    if time_tuple is None:
        return
    excel_value = time_tuple_to_excel_string(time_tuple)
    if not validate_excel_time_string(excel_value):
        raise ExcelFormatError(f"{field_label}の形式が不正です: {excel_value!r}")


def read_target_month_first_day(workbook: Workbook) -> Optional[datetime.date]:
    """
    C8セル（対象月初日）の値を datetime.date として読み取る。

    C8はPython側が直接書き込む実値（数式ではない）のため、
    data_only=True 読み込みでも確実に値が取得できる
    （set_monthly_header_cells 参照）。
    """
    ws = workbook[SHEET_HONBUN]
    value = ws[CELL_TARGET_MONTH].value

    if value is None:
        return None
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        # datetime.date / datetime.datetime
        if hasattr(value, "date"):
            return value.date() if not isinstance(value, datetime.date) or isinstance(value, datetime.datetime) else value
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
    B列の数式（"=+C8" → 1日目、"=+B{n-1}+1" → 以降1日ずつ加算）と
    完全に等価な計算をPython側で行う。B列・C列の数式やキャッシュには
    一切触れず、独立して日付を算出する。

    対象月の実日数を超える行（31日がない月の31日目など）は None を返す。
    """
    if target_month_first_day is None:
        return None

    day_offset = row - DATE_START_ROW  # DATE_START_ROW行目が1日目（offset=0）
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
    sheet_name: str = SHEET_HONBUN
) -> AttendanceHeaderValues:
    """
    表外の勤怠情報を読み取る。

    日次データ(DayCellValues)には含めない。
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
    ヘッダー領域の時刻セル値（深夜開始時間・所定・午前・午後・休憩時間等）を、
    秒なしの "H:MM" 形式の表示用文字列に整形する。
    値が None（セルが空）の場合は文字列 "None" を返す。
    """
    if value is None:
        return "None"

    # datetime.timedelta（経過時間書式のセル）
    if isinstance(value, datetime.timedelta):
        total_minutes = int(value.total_seconds() // 60)
        hour, minute = divmod(total_minutes, 60)
        return f"{hour}:{minute:02d}"

    # datetime.time / datetime.datetime
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
    (時, 分) タプルを、Excelの時刻シリアル値と同じ考え方の「時間の小数表現」
    （例：9:30 → 9.5）に変換する。Excelのセル同士の引き算
    （例：Q11-N11）は、内部的にはこの小数表現同士の引き算と等価であるため、
    自動計算列の再現にはこの単位で計算する。
    """
    if value is None:
        return None
    hour, minute = value
    return hour + minute / 60.0


def _decimal_hours_to_time_tuple(value: Optional[float]) -> Optional[tuple[int, int]]:
    """小数表現の時間を (時, 分) タプルに戻す（表示用）。負値もそのまま許容する。"""
    if value is None:
        return None
    sign = -1 if value < 0 else 1
    total_minutes = round(abs(value) * 60)
    hour, minute = divmod(total_minutes, 60)
    return (sign * hour if hour != 0 else (0 if sign > 0 else -0), minute) if False else (
        (sign * hour, minute) if sign < 0 or hour > 0 else (0, minute)
    )


@dataclass
class AutoCalculatedValues:
    """
    自動計算列の再計算結果（基本設計書には存在しない、Python側での
    数式再現の結果を表すDTO）。値は (時, 分) タプル、または
    計算対象外の場合は None。
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

    再現対象の数式（11行目時点でのセル参照例）：

    休憩時間1 = IF(休暇種類 in (午前半休,午後半休), 0,
                    IF(始業と終業が両方入力されている,
                       MIN(MAX(終業,休憩1開始),休憩1終了)
                         - MAX(MIN(始業,休憩1終了),休憩1開始),
                       ""))
        → 勤務時間帯（始業〜終業）と、休憩1の固定時間帯（$G$19〜$H$19）
          との共通区間の長さを休憩時間1とする（区間交差計算）。

    休憩時間2・3 = 手入力値のため計算しない（引数の値をそのまま使う）。

    実働時間 = IF(始業="", "：",
                  終業 - 始業 - (入力されている休憩1・休憩2・休憩3・離業の合計))
        ※休憩1は、上記で計算した値を使う。

    超勤 = IF(休暇種類 in (午前半休,午後半休), 0,
              IF(休暇種類=休日出勤, "：",
              IF(始業が空, "：", 実働時間 - 所定労働時間)))
        → 所定労働時間より実働時間が短い場合は負の値（マイナス）になる。

    休日出勤 = IF(休暇種類=休日出勤, 実働時間, "：")

    深夜 = IF(始業が空, "：",
              IF(終業<=深夜開始時刻, "：", 終業 - 深夜開始時刻))

    引数:
        break_time_1_input: 呼び出し側が保持している休憩1の値
            （読み取り専用の表示にしか使わない。計算結果で上書きされる）
        break_1_start, break_1_end: 休憩時間1の固定時間帯
            （header.break_time_1_start / break_time_1_end に対応）
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

    # ------------------------------------------------
    # 休憩時間1（区間交差計算）
    # ------------------------------------------------
    if leave_type in ("午前半休", "午後半休"):
        break_1_decimal = 0.0
    elif start is None or end is None:
        break_1_decimal = None  # Excel側は "" だが、ここではNoneで「未計算」を表す
    elif break_1_start_dec is None or break_1_end_dec is None:
        # 休憩1の固定時間帯自体がヘッダーに設定されていない場合は計算不能
        break_1_decimal = None
    else:
        overlap_end = min(max(end, break_1_start_dec), break_1_end_dec)
        overlap_start = max(min(start, break_1_end_dec), break_1_start_dec)
        break_1_decimal = overlap_end - overlap_start
        if break_1_decimal < 0:
            break_1_decimal = 0.0

    # ------------------------------------------------
    # 実働時間（休憩1は上記の計算結果を使う）
    # ------------------------------------------------
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

    # ------------------------------------------------
    # 超勤：所定労働時間と実働時間の差。実働が所定に足りなければマイナス。
    # ------------------------------------------------
    if leave_type in ("午前半休", "午後半休"):
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

    # ------------------------------------------------
    # 休日出勤
    # ------------------------------------------------
    if leave_type == "休日出勤":
        holiday_work_decimal = actual_work_decimal
    else:
        holiday_work_decimal = None

    # ------------------------------------------------
    # 深夜
    # ------------------------------------------------
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
    ヘッダー領域の時刻セル値（深夜開始時間・所定等）を (時, 分) タプルに
    変換する。datetime.time / datetime.datetime / datetime.timedelta /
    文字列のいずれにも対応する（format_time_value_no_secondsと対になる
    変換だが、こちらは表示用文字列ではなく計算用のタプルを返す）。
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

    data_only=True を指定すると、セルの `.value` は数式文字列
    （例: "=SUM(N14:N44)"）ではなく、Excelが最後に計算してファイルに
    保存したキャッシュ値（表示上の値）を返すようになる。

    注意点:
    - この方法で開いたワークブックは数式そのものを保持していないため、
      保存（save）用途には絶対に使わないこと（数式が失われる）。
      読み取り専用（プレビュー・記入例表示）専用の関数とする。
    - Excelでファイルが一度も計算・保存されていない場合（例えば
      openpyxlのみで生成し、Excelアプリで開かれたことがないファイル）は、
      計算キャッシュが存在せず .value が None になる点に注意。
      月次ファイル生成直後（テンプレートコピー直後）は、テンプレート自体が
      Excelで保存されている前提であれば通常はキャッシュ値を保持している。
    """
    return openpyxl.load_workbook(path, data_only=True)


def save_workbook_to_path(workbook: Workbook, path: str) -> None:
    """ワークブックを指定パスへ保存する。"""
    workbook.save(path)