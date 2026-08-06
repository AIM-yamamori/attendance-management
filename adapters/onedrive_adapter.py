"""
onedrive_adapter.py

【概要】
OneDrive（共有URL経由）とのファイル入出力を担うアダプター。
「フォルダの存在確認」「フォルダ作成」「ファイルの存在確認」
「ファイルのダウンロード」「ファイルのアップロード」という
5つの操作だけを外部（attendance_service等）に公開し、
呼び出し側はOneDriveの内部実装（共有URLの形式やHTTPリクエストの
詳細）を一切意識しなくてよいようにする。

【現時点の実装方式（暫定）】
要件定義書14章・基本設計書14章に記載のとおり、OneDriveの共有URL
経由でのファイル一覧取得・アップロード・上書き時の挙動（エラーに
なるか上書きされるかなど）は本設計時点で未検証である。
そのため、実際のOneDrive共有URL連携を実装する前段階として、
ローカルファイルシステム上のディレクトリを「疑似OneDrive」として
扱う実装をここに用意する。

環境変数 ONEDRIVE_LOCAL_ROOT で指定したローカルディレクトリを
OneDriveのルートディレクトリに見立てて読み書きする。
本番のOneDrive共有URL連携（PoC完了後）に差し替える際は、
このファイル内の各関数の中身だけを実HTTPリクエストに置き換えれば
よく、呼び出し側（attendance_service等）のコードは変更不要となる
よう、関数シグネチャ（引数・戻り値の形）を先に確定させている。

【冪等性についての方針】
フォルダ作成・ファイルアップロードは、「存在すればそれを使う、
なければ作る」という冪等な動作を必ず行う（要件定義書4.2節・
基本設計書6.3節・8.4節）。複数リクエストがほぼ同時に同じフォルダ・
ファイルを作ろうとしても、最終的に1つに収束する必要があるため、
本実装ではファイルシステムの「存在すれば上書きしない」という
性質を利用し、アップロード関数に overwrite=False を渡した場合は
既存ファイルがあれば何もしない（無視する）動作とする。
"""

import os
import shutil
from pathlib import Path


class OneDriveError(Exception):
    """OneDrive操作全般で発生したエラーを表す例外。"""


def _get_root_dir() -> Path:
    """
    疑似OneDriveのルートディレクトリを環境変数から取得する。
    本番のOneDrive共有URL連携に差し替えた際は、この関数（および
    以降の関数の中身）だけを差し替えればよい。
    """
    root = os.environ.get("ONEDRIVE_LOCAL_ROOT", "./data/onedrive_local")
    root_path = Path(root)
    root_path.mkdir(parents=True, exist_ok=True)
    return root_path


def folder_exists(folder_name: str) -> bool:
    """
    ルートディレクトリ直下に、指定した名前のフォルダが存在するかを確認する。
    例: folder_exists("202608") -> ルート直下の"202608"フォルダの有無
    """
    target = _get_root_dir() / folder_name
    return target.is_dir()


def ensure_folder(folder_name: str) -> None:
    """
    指定した名前のフォルダをルートディレクトリ直下に作成する。
    既に存在する場合はエラーにせず、そのまま何もしない（冪等な作成処理。
    要件定義書4.2節「フォルダ作成の競合対策」、基本設計書6.3節参照）。
    """
    target = _get_root_dir() / folder_name
    target.mkdir(parents=True, exist_ok=True)


def file_exists(relative_path: str) -> bool:
    """
    ルートディレクトリからの相対パスでファイルの存在を確認する。
    例: file_exists("202608/【会社名】勤務実績管理表_202608_山田太郎.xlsx")
    例: file_exists("【会社名】勤務実績管理表_テンプレート.xlsx")
    """
    target = _get_root_dir() / relative_path
    return target.is_file()


def download_file(relative_path: str, local_dest_path: str) -> None:
    """
    OneDrive上の指定ファイルを、ローカルの作業用パス（local_dest_path）へ
    コピーする（openpyxlで直接開けるようにするため、一度ローカルへ
    落とす方式に統一する）。

    存在しない場合は OneDriveError を送出する。
    """
    source = _get_root_dir() / relative_path
    if not source.is_file():
        raise OneDriveError(f"ファイルが見つかりません: {relative_path}")

    dest = Path(local_dest_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, dest)


def upload_file(local_source_path: str, relative_path: str, overwrite: bool = True) -> None:
    """
    ローカルの作業用パス（local_source_path）にあるファイルを、
    OneDrive上の指定パス（relative_path）へアップロード（コピー）する。

    overwrite=False の場合、アップロード先に既にファイルが存在すれば
    何もしない（同時アクセスによる重複生成対策。基本設計書6.3節・8.4節の
    「同名ファイルが既に存在する場合はエラーとせず、既存ファイルを使う」
    という冪等な仕様に対応する）。
    overwrite=True の場合は無条件に上書きする（勤怠データの保存時など、
    意図的に内容を更新したい場合に使う）。
    """
    dest = _get_root_dir() / relative_path

    if not overwrite and dest.is_file():
        # 既に存在する場合は何もしない（先着優先、エラーにしない）
        return

    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(local_source_path, dest)


def list_files(folder_name: str) -> list[str]:
    """
    指定フォルダ直下のファイル名一覧を返す（拡張子含むファイル名のみ、
    サブフォルダは含めない）。フォルダが存在しない場合は空リストを返す。
    """
    target = _get_root_dir() / folder_name
    if not target.is_dir():
        return []
    return sorted(p.name for p in target.iterdir() if p.is_file())