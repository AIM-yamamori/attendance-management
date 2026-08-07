"""
attendance_service.py

【概要】
「勤怠Excelファイルをどう取得し、どう読み書きするか」の業務ロジックを
担うモジュール。excel_adapter（Excelファイル単体の読み書き）と
onedrive_adapter（OneDrive上でのファイル入出力）という、それぞれ
別の関心事を持つ2つのアダプターを組み合わせて、「対象ユーザー・
対象月の勤怠データを取得する／保存する」という業務レベルの操作として
まとめて提供する。

このモジュールが担う中心的な機能が get_or_create_monthly_file である。
要件定義書4.2節「月次Excelファイルの新規作成」の仕様（ユーザーが
対象月にアクセスした時点で、未生成であれば生成し、既存ならそれを
使う）を実現する。複数ユーザーがほぼ同時に同じ月へ初回アクセスした
場合でも、最終的にファイル・フォルダが1つに収束するよう、
「存在確認→なければ作る、失敗しても既存を使う」という冪等な手順で
実装している（要件定義書4.2節・基本設計書6.3節・8.4節）。
"""

import datetime
import os
import tempfile
import calendar
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from adapters import excel_adapter, onedrive_adapter

# ============================================
# ファイル名・フォルダ名の生成規則（要件定義書4.2節・5.1節）
# ============================================

# 会社名は固定値。実際の表記は依頼者確認事項のため、
# 環境変数で差し替えられるようにしておく（未確定事項、基本設計書14章）。
_DEFAULT_COMPANY_NAME = "【会社名】"

TEMPLATE_FILE_NAME_SUFFIX = "勤務実績管理表_テンプレート.xlsx"


def _company_name() -> str:
    return os.environ.get("COMPANY_NAME", _DEFAULT_COMPANY_NAME)


def get_template_relative_path() -> str:
    """
    OneDriveルートディレクトリ直下のテンプレートファイルの相対パスを返す。
    例: "【会社名】勤務実績管理表_テンプレート.xlsx"
    """
    return f"{_company_name()}{TEMPLATE_FILE_NAME_SUFFIX}"


def build_monthly_file_name(target_month: str, full_name_no_space: str) -> str:
    """
    月次ファイルのファイル名を組み立てる（要件定義書5.1節）。
    姓名の間にスペースを入れない結合を使う（auth_service.User.full_name
    またはservices.user_service.User.display_nameとは別の、
    ファイル名専用の結合ルール。呼び出し側で "姓+名" を渡すこと）。

    例: build_monthly_file_name("202608", "山田太郎")
        -> "【会社名】勤務実績管理表_202608_山田太郎.xlsx"
    """
    return f"{_company_name()}勤務実績管理表_{target_month}_{full_name_no_space}.xlsx"


def build_monthly_file_relative_path(target_month: str, full_name_no_space: str) -> str:
    """
    月次ファイルの、OneDriveルートディレクトリからの相対パスを返す
    （フォルダ名/ファイル名）。
    例: "202608/【会社名】勤務実績管理表_202608_山田太郎.xlsx"
    """
    file_name = build_monthly_file_name(target_month, full_name_no_space)
    return f"{target_month}/{file_name}"


def target_month_to_first_day(target_month: str) -> str:
    """
    "YYYYMM" 形式の対象月から "YYYY/MM/01" 形式の文字列を作る
    （ファイル名生成など、文字列としての表示が必要な箇所向け）。
    例: "202608" -> "2026/08/01"
    """
    if len(target_month) != 6 or not target_month.isdigit():
        raise ValueError(f"target_monthは YYYYMM 形式で指定してください: {target_month!r}")
    year = target_month[:4]
    month = target_month[4:6]
    return f"{year}/{month}/01"

def _days_in_month(target_month: str) -> int:
    """
    "YYYYMM" 形式の対象月の日数を返す。
    例: "202602" -> 28（うるう年なら29）、"202608" -> 31
    """
    year = int(target_month[:4])
    month = int(target_month[4:6])
    return calendar.monthrange(year, month)[1]


def target_month_to_date(target_month: str) -> datetime.date:
    """
    "YYYYMM" 形式の対象月から、Excel C8セルへ直接代入するための
    datetime.date オブジェクトを作る。

    Excelセルへ日付を "YYYY/MM/01" のような文字列としてそのまま代入すると、
    Excel側ではテキストとして解釈され、セル左上にエラーインジケーターが
    付いたり、'（アポストロフィ）付き文字列と同等の状態になってしまう。
    date オブジェクトを代入することで、Excel側で本来の日付シリアル値として
    認識される（要件定義書4.2節）。
    """
    if len(target_month) != 6 or not target_month.isdigit():
        raise ValueError(f"target_monthは YYYYMM 形式で指定してください: {target_month!r}")
    year = int(target_month[:4])
    month = int(target_month[4:6])
    return datetime.date(year, month, 1)


# ============================================
# 対象年月の選択制御（基本設計書3.5.2節）
# ============================================

def _add_months(year: int, month: int, delta: int) -> tuple[int, int]:
    """年・月に対して delta ヶ月を加算した (年, 月) を返す（負数も可）。"""
    total = (year * 12 + (month - 1)) + delta
    new_year, new_month0 = divmod(total, 12)
    return new_year, new_month0 + 1


def build_selectable_months(
    today: datetime.date,
    service_start_month: Optional[str],
) -> list[str]:
    """
    月選択プルダウンに表示する "YYYYMM" のリストを、
    「運用開始月 〜 当月の翌月」の範囲で古い順に生成する
    （基本設計書3.5.2節）。

    service_start_month が None または空文字の場合（運用開始月が
    アプリDBに未設定の場合）は、安全側に倒して当月のみを選択肢とする
    （運用開始月が未設定のまま過去分まで見えてしまう事故を防ぐため。
    本来は _init_app() で環境変数 SERVICE_START_MONTH から必ず1回
    投入される想定のため、未設定は設定漏れ等の異常系にあたる）。
    """
    this_year, this_month = today.year, today.month
    default_month = f"{this_year:04d}{this_month:02d}"

    if not service_start_month:
        return [default_month]

    next_year, next_month = _add_months(this_year, this_month, 1)
    upper_bound = f"{next_year:04d}{next_month:02d}"
    lower_bound = service_start_month

    if lower_bound > upper_bound:
        # 運用開始月が翌月より後（未来すぎる設定ミス等）の場合は
        # 上限のみを1件返す（空リストにはしない。3.5.2節の趣旨上、
        # 少なくとも当月相当は選べる状態を維持する）
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
    """デフォルト選択月（当年当月）を返す（基本設計書3.5.2節）。"""
    return f"{today.year:04d}{today.month:02d}"


def format_month_label(target_month: str) -> str:
    """
    "YYYYMM" 形式を画面表示用の "YYYY年MM月" 形式に整形する。
    """
    return f"{target_month[:4]}年{target_month[4:6]}月"


# ============================================
# データ型
# ============================================

@dataclass
class WorkbookHandle:
    """
    「取得・生成した月次Excelファイル」を表すハンドル。
    ローカルの一時ファイルパスと、OneDrive上の相対パスの両方を持つ。
    呼び出し側（画面）はこのハンドルを使って読み書きを行い、
    保存時は save_workbook() にこのハンドルを渡す。
    """
    local_path: str
    onedrive_relative_path: str
    target_month: str
    employee_id: str


# ============================================
# 月次ファイル生成・取得（4.2節の中核機能）
# ============================================

def _local_work_dir() -> Path:
    """
    Excelファイルを一時的にローカルへ落とす際の作業ディレクトリ。
    Streamlitはリクエストごとに同一プロセス内で動くため、
    tempfile.gettempdir() 配下にアプリ専用のサブディレクトリを作る。
    """
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
    対象ユーザー・対象月の月次Excelファイルを取得する。
    既に存在すればそれをそのまま使い（再生成しない）、存在しなければ
    テンプレートから新規生成する（要件定義書4.2節・基本設計書6.3節）。

    冪等性について：
    - フォルダ作成は onedrive_adapter.ensure_folder() が「存在すれば
      何もしない」という動作を保証する。
    - ファイル生成についても、アップロード直前に再度存在確認を行い、
      アップロード自体は overwrite=False で行うことで、複数リクエストが
      ほぼ同時に生成を試みても最終的に1つのファイルに収束する
      （基本設計書6.3節・8.4節）。

    引数:
        employee_id: 社員番号
        target_month: "YYYYMM" 形式
        full_name_no_space: ファイル名用の氏名（スペースなし結合）
        full_name_with_space: Excel AH5セル用の氏名（全角スペース結合）
        department: 部署名（AH4セルに自動設定）
    """
    # 1. 対象月フォルダの存在確認・作成（冪等）
    onedrive_adapter.ensure_folder(target_month)

    relative_path = build_monthly_file_relative_path(target_month, full_name_no_space)
    local_path = str(_local_work_dir() / f"{employee_id}_{target_month}.xlsx")

    # 2. 既存ファイルがあれば、それをそのままダウンロードして使う（再生成しない）
    if onedrive_adapter.file_exists(relative_path):
        onedrive_adapter.download_file(relative_path, local_path)
        return WorkbookHandle(
            local_path=local_path,
            onedrive_relative_path=relative_path,
            target_month=target_month,
            employee_id=employee_id,
        )

    # 3. 存在しない場合はテンプレートから新規生成する
    workbook = _create_new_monthly_workbook(
        target_month=target_month,
        employee_id=employee_id,
        department=department,
        full_name_with_space=full_name_with_space,
    )
    excel_adapter.save_workbook_to_path(workbook, local_path)

    # 4. OneDriveへアップロードする。overwrite=False とすることで、
    #    同時アクセスで他のリクエストが先にアップロードを完了させて
    #    いた場合は、そちらを優先しこちらのアップロードは無視される
    #    （基本設計書6.3節「同時アクセスによる重複生成対策」）。
    onedrive_adapter.upload_file(local_path, relative_path, overwrite=False)

    # 5. アップロード後、実際にOneDrive上にある版（＝先着した版かもしれない）
    #    を改めて取得し直す。自分がアップロードしたものと、他リクエストが
    #    先にアップロードしたものが食い違う可能性を考慮した安全策。
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
    """
    （docstring省略・既存のまま）
    """
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

@dataclass
class DayEntry:
    """フロント（画面）向けの1日分の勤怠データ。基本設計書5.3節DTOに対応。"""
    row: int
    date_value: object
    weekday: Optional[str]
    leave_type: Optional[str]
    start_time: Optional[tuple[int, int]]
    end_time: Optional[tuple[int, int]]
    leave_time: Optional[tuple[int, int]]
    work_note: Optional[str]
    # 自動計算項目（プレビュー時のみ表示、編集画面では非表示。
    # 基本設計書3.5.3節・7.3節）
    break_time_1: object
    break_time_2: object
    break_time_3: object
    actual_work_time: object
    overtime: object
    holiday_work: object
    late_night: object

@dataclass
class AttendanceData:
    """1ユーザー・1ヶ月分の勤怠データ全体（基本設計書5.2.4節）。"""
    target_month: str
    employee_id: str
    # 表外情報
    header: excel_adapter.AttendanceHeaderValues
    # 日別勤怠データ
    entries: list[DayEntry]


def load_attendance(handle: WorkbookHandle) -> AttendanceData:
    """
    「原本」シートの全項目を読み取り、プレビュー用データとして返す
    （基本設計書5.2.4節）。
    """
    workbook = excel_adapter.load_workbook_from_path_for_display(handle.local_path)

    header = excel_adapter.read_header_values(
        workbook,
        sheet_name=excel_adapter.SHEET_HONBUN
    )

    day_rows = excel_adapter.read_day_rows(
        workbook,
        sheet_name=excel_adapter.SHEET_HONBUN
    )

    entries = [_day_cell_values_to_entry(d) for d in day_rows]
    entries = _filter_entries_to_month_days(entries, handle.target_month)

    return AttendanceData(
        target_month=handle.target_month,
        employee_id=handle.employee_id,
        header=header,
        entries=entries,
    )


def _filter_entries_to_month_days(
    entries: list[DayEntry], target_month: str
) -> list[DayEntry]:
    """
    対象月の実日数を超える行を除外する。

    テンプレートの日次データ欄は31行分（1〜31日）を固定で持っているため、
    30日までしかない月（4,6,9,11月）や28〜29日までしかない2月では、
    実在しない日（31日や29〜30日）の行が生成されてしまう。
    date_valueがNoneの行（未入力行）は既存の画面側フィルタで除外される
    ため、ここでは「日付は入っているが、対象月の実日数を超える行」
    （例えばテンプレートの自動採番等で32日相当の値が入ってしまった
    ケース）を対象月の日数を基準に除外する。
    """
    max_day = _days_in_month(target_month)
    filtered = []
    for entry in entries:
        day_number = _extract_day_number(entry.date_value)
        if day_number is not None and day_number > max_day:
            continue
        filtered.append(entry)
    return filtered


def _extract_day_number(date_value) -> Optional[int]:
    """
    B列の日付値（datetime型または文字列）から「日」の部分（1〜31）を
    取り出す。判定できない場合はNoneを返す（除外対象にしない、
    安全側に倒す）。
    """
    if date_value is None:
        return None
    if hasattr(date_value, "day"):
        return date_value.day
    if isinstance(date_value, str):
        match = re.match(r"^\d{4}[/\-]\d{1,2}[/\-](\d{1,2})", date_value.strip())
        if match:
            return int(match.group(1))
    return None


def _day_cell_values_to_entry(d) -> DayEntry:
    """
    excel_adapter.DayCellValues を、画面向けの DayEntry へ変換する共通処理。
    load_attendance と load_example の両方から使う
    （対象シートが違うだけで変換ロジックは完全に共通のため、
    基本設計書5.2.4節の方針どおりここに1箇所だけ実装する）。
    """
    return DayEntry(
        row=d.row,
        date_value=d.date_value,
        weekday=d.weekday,
        leave_type=d.leave_type,
        start_time=d.start_time,
        end_time=d.end_time,
        leave_time=d.leave_time,
        work_note=d.work_note,
        break_time_1=d.break_time_1,
        break_time_2=d.break_time_2,
        break_time_3=d.break_time_3,
        actual_work_time=d.actual_work_time,
        overtime=d.overtime,
        holiday_work=d.holiday_work,
        late_night=d.late_night,
    )


def load_example(handle: WorkbookHandle) -> AttendanceData:
    """
    （docstring省略・既存のまま）
    """
    workbook = excel_adapter.load_workbook_from_path_for_display(handle.local_path)

    header = excel_adapter.read_header_values(
        workbook,
        sheet_name=excel_adapter.SHEET_KINYUREI
    )

    day_rows = excel_adapter.read_day_rows(
        workbook,
        sheet_name=excel_adapter.SHEET_KINYUREI
    )

    entries = [_day_cell_values_to_entry(d) for d in day_rows]
    entries = _filter_entries_to_month_days(entries, handle.target_month)

    return AttendanceData(
        target_month=handle.target_month,
        employee_id=handle.employee_id,
        header=header,
        entries=entries,
    )


# ============================================
# 編集・保存（Day5、基本設計書3.5.4節）
# ============================================

@dataclass
class DayEditInput:
    """
    編集フォームから受け取る1日分の入力値。
    自動計算項目（休憩・実働・超勤等）は編集対象外のため含まない
    （基本設計書3.5.3節・7.3節：編集画面では自動計算項目は非表示）。
    """
    row: int
    leave_type: str
    start_time: Optional[tuple[int, int]]
    end_time: Optional[tuple[int, int]]
    leave_time: Optional[tuple[int, int]]
    work_note: str


def save_attendance(handle: WorkbookHandle, edits: list[DayEditInput]) -> None:
    """
    編集フォームの入力内容を「原本」シートへまとめて反映し、OneDriveへ
    保存する（画面全体を1回でまとめて保存する方式。基本設計書3.5.4節）。

    処理の流れ：
    1. 保存用にワークブックを開き直す（data_only=Trueの表示専用版とは
       別に、数式を保持したまま書き込み用として開く必要があるため）。
    2. 各行についてexcel_adapter.write_day_cellで値のみを更新する
       （書式・数式・他のセルには一切触れない）。
    3. ローカルへ保存し、OneDriveへアップロードする（overwrite=True。
       既存ファイルの更新のため、新規生成時のような排他制御は不要）。

    休暇種類が空文字（""）の場合はNoneとしてセルに書き込み、
    「未選択＝通常勤務」を表現する（excel_adapter.LEAVE_TYPE_OPTIONSの
    先頭要素が空文字であることに対応）。
    """
    workbook = excel_adapter.load_workbook_from_path(handle.local_path)

    for edit in edits:
        excel_adapter.write_day_cell(
            workbook,
            row=edit.row,
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