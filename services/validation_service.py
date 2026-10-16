"""
validation_service.py

【概要】
バリデーションルール判定を担うモジュール（基本設計書5.2.5節・7章）。
フロント（入力の都度）・バック（保存直前）の二重チェックに、
同一のロジックを使い回す（7.2節）。

7.1節 ルール一覧：
VAL-01: date is None and leave_type is not None
VAL-02: date is None and (start_time is not None or end_time is not None)
VAL-03: leave_type in [有休,長期連休,特別休暇,欠勤,振替休暇]
        and (start_time is not None or end_time is not None)
VAL-04/05: weekday in [土,日] and leave_type is None
        and (start_time is not None or end_time is not None)
        （休暇種類が選択されていれば土日でも時刻入力を許容する）
"""

from dataclasses import dataclass
from typing import Optional

# 7.1節 VAL-03 対象の休暇種類（この5種類のみ。半休2種は対象外）
FULL_DAY_LEAVE_TYPES = ("有休", "長期連休", "特別休暇", "欠勤", "振替休暇")


@dataclass
class ValidationError:
    """基本設計書5.3節 ValidationError に対応。"""
    row_index: int
    field: str
    message: str


def validate_entry(
    row_index: int,
    date_value,
    weekday: Optional[str],
    leave_type: Optional[str],
    start_time: Optional[tuple[int, int]],
    end_time: Optional[tuple[int, int]],
) -> list[ValidationError]:
    """
    1日分の入力値を検証する（7.1節ルール1〜5）。

    引数:
        row_index: 対象行（Excel行番号や表示上の行インデックス。
            呼び出し側の都合に合わせて渡す）
        date_value: 日付値（None＝日付が存在しない行）
        weekday: 曜日（"月"〜"日"、または未算出ならNone）
        leave_type: 休暇種類（空文字またはNoneは「未選択」として扱う）
        start_time: (時, 分) または None
        end_time: (時, 分) または None
    """
    errors: list[ValidationError] = []

    has_leave_type = bool(leave_type)
    has_start_or_end = start_time is not None or end_time is not None

    # VAL-01: 日付が存在しない行には休暇種類を入力できない
    if date_value is None and has_leave_type:
        errors.append(
            ValidationError(
                row_index=row_index,
                field="leave_type",
                message="日付が存在しない行には休暇種類を入力できません",
            )
        )

    # VAL-02: 日付が存在しない行には始業・終業時間を入力できない
    if date_value is None and has_start_or_end:
        errors.append(
            ValidationError(
                row_index=row_index,
                field="start_time" if start_time is not None else "end_time",
                message="日付が存在しない行には始業・終業時間を入力できません",
            )
        )

    # VAL-03: 休暇区分（有休・長期連休・特別休暇・欠勤・振替休暇）の日には
    # 始業・終業時間を入力できない（逆順操作＝時刻入力済みの状態から
    # 休暇区分へ変更した場合も同一ルールで検知する）
    if leave_type in FULL_DAY_LEAVE_TYPES and has_start_or_end:
        errors.append(
            ValidationError(
                row_index=row_index,
                field="start_time" if start_time is not None else "end_time",
                message="この休暇区分の日には始業・終業時間を入力できません",
            )
        )

    # VAL-04/05: 土日かつ休暇種類未選択の場合に時刻入力があればエラー。
    # 休暇種類が何か選択されていれば（休日出勤・振替出勤等）許容する。
    if weekday in ("土", "日") and not has_leave_type and has_start_or_end:
        errors.append(
            ValidationError(
                row_index=row_index,
                field="start_time" if start_time is not None else "end_time",
                message="休日に始業・終業時間が入力されています",
            )
        )

    return errors


def validate_all(entries: list) -> list[ValidationError]:
    """
    保存直前の一括再チェック（7.2節）。

    引数:
        entries: DayEntry相当のオブジェクトのリスト。
            各要素は row（またはrow_index）、date_value、weekday、
            leave_type、start_time、end_time属性を持つことを想定。
            attendance_service.DayEntry / DayEditInput いずれでも、
            これらの属性名を満たしていれば利用できる。
    """
    errors: list[ValidationError] = []
    for entry in entries:
        row_index = getattr(entry, "row", None)
        if row_index is None:
            row_index = getattr(entry, "row_index", None)

        errors.extend(
            validate_entry(
                row_index=row_index,
                date_value=getattr(entry, "date_value", None),
                weekday=getattr(entry, "weekday", None),
                leave_type=getattr(entry, "leave_type", None),
                start_time=getattr(entry, "start_time", None),
                end_time=getattr(entry, "end_time", None),
            )
        )
    return errors