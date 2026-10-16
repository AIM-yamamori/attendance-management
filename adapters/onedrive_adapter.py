"""
onedrive_adapter.py

【概要】
OneDrive（共有URL経由）とのファイル入出力を担うアダプター。
「フォルダの存在確認」「フォルダ作成」「ファイルの存在確認」
「ファイルのダウンロード」「ファイルのアップロード」という
5つの操作だけを外部（attendance_service等）に公開し、
呼び出し側はOneDriveの内部実装（Graph APIの認証・エンドポイントの
詳細）を一切意識しなくてよいようにする。

【実装方式：Microsoft Graph API（共有URL経由・クライアントクレデンシャル）】
「リンクを知っている全員が編集可能」なOneDrive共有リンクに対し、
Microsoft Graph APIの /shares エンドポイントでアクセスする
（要件定義書14章・基本設計書14章で「未検証」とされていた部分の本実装）。

認証は client credentials フロー（アプリ自身がテナントに対して自分の
資格情報を提示する方式）を使う。エンドユーザーのMicrosoftアカウント
ログインは発生しない。事前にAzure ADへのアプリ登録・
`Files.ReadWrite.All`（アプリケーション権限）の管理者同意が必要
（詳細は別紙「Azure ADアプリ登録手順書」を参照）。

【必要な環境変数】
- ONEDRIVE_TENANT_ID: Azure AD テナントID
- ONEDRIVE_CLIENT_ID: 登録したアプリのクライアントID
- ONEDRIVE_CLIENT_SECRET: 発行したクライアントシークレット
- ONEDRIVE_SHARE_URL: OneDriveの共有リンク（対象フォルダのルートを指す）

これらが揃っていない場合（ローカル開発時等）は、従来通りローカル
ディレクトリを疑似OneDriveとして使うフォールバック実装を使う
（ONEDRIVE_LOCAL_ROOT。本番運用ではONEDRIVE_TENANT_ID等を設定し、
このフォールバックが使われないようにすること）。

【本モジュールが扱う「パス」の考え方】
呼び出し側（attendance_service等）は、共有リンクが指すフォルダを
ルートとした相対パス（例："202608/【会社名】勤務実績管理表_202608_山田太郎.xlsx"）
だけを意識すればよい。共有リンクの実体（driveId・itemId）の解決や、
サブフォルダ・ファイルのdriveItem ID解決は、すべて本モジュール内で
行い、呼び出し側には一切露出しない。

【冪等性についての方針】
フォルダ作成・ファイルアップロードは、「存在すればそれを使う、
なければ作る」という冪等な動作を必ず行う（要件定義書4.2節・
基本設計書6.3節・8.4節）。Graph API側には「存在チェック→なければ作成」
という単純な逐次処理を実装している。本番運用でリクエストが完全に
同時に競合するケースでは重複作成の可能性がゼロではないが、
フォルダ作成・アップロードいずれも「同名なら上書き／再利用」で
実害が出ない設計にしている（3.2節参照）。
"""

import mimetypes
import os
import shutil
import threading
import time
from base64 import urlsafe_b64encode
from pathlib import Path
from typing import Optional

import requests

_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
_TOKEN_URL_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

# トークンの有効期限に対し、余裕を持って早めに再取得するためのバッファ（秒）
_TOKEN_EXPIRY_BUFFER_SECONDS = 60

_REQUEST_TIMEOUT_SECONDS = 30


class OneDriveError(Exception):
    """OneDrive操作全般で発生したエラーを表す例外。"""


# ============================================
# 実装方式の切り替え（Graph API本番実装 ⇔ ローカル疑似OneDrive）
# ============================================

def _graph_api_configured() -> bool:
    """
    Graph API本番接続に必要な環境変数が揃っているかを判定する。
    揃っていなければローカル疑似OneDrive実装にフォールバックする
    （ローカル開発・単体テスト用）。
    """
    return bool(
        os.environ.get("ONEDRIVE_TENANT_ID")
        and os.environ.get("ONEDRIVE_CLIENT_ID")
        and os.environ.get("ONEDRIVE_CLIENT_SECRET")
        and os.environ.get("ONEDRIVE_SHARE_URL")
    )


# ============================================
# Graph API：認証（クライアントクレデンシャルフロー）
# ============================================

_token_cache_lock = threading.Lock()
_token_cache: dict = {"access_token": None, "expires_at": 0.0}


def _get_access_token() -> str:
    """
    クライアントクレデンシャルフローでアクセストークンを取得する。
    有効期限内であればキャッシュを再利用し、切れていれば再取得する
    （毎回のAPI呼び出しごとにトークンエンドポイントを叩かないため）。
    """
    with _token_cache_lock:
        now = time.monotonic()
        if _token_cache["access_token"] and now < _token_cache["expires_at"]:
            return _token_cache["access_token"]

        tenant_id = os.environ["ONEDRIVE_TENANT_ID"]
        client_id = os.environ["ONEDRIVE_CLIENT_ID"]
        client_secret = os.environ["ONEDRIVE_CLIENT_SECRET"]

        response = requests.post(
            _TOKEN_URL_TEMPLATE.format(tenant_id=tenant_id),
            data={
                "client_id": client_id,
                "client_secret": client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
            timeout=_REQUEST_TIMEOUT_SECONDS,
        )
        if response.status_code != 200:
            raise OneDriveError(
                f"Microsoft Graph APIの認証に失敗しました "
                f"（{response.status_code}）: {response.text}"
            )

        payload = response.json()
        access_token = payload["access_token"]
        expires_in = payload.get("expires_in", 3600)

        _token_cache["access_token"] = access_token
        _token_cache["expires_at"] = now + expires_in - _TOKEN_EXPIRY_BUFFER_SECONDS

        return access_token


def _auth_headers() -> dict:
    return {"Authorization": f"Bearer {_get_access_token()}"}


# ============================================
# Graph API：共有URLのエンコード・ルートdriveItem解決
# ============================================

def _encode_sharing_url(share_url: str) -> str:
    """
    共有URLをGraph APIの /shares エンドポイントで使う形式
    （"u!" + base64url、パディングなし）にエンコードする
    （Microsoft Graph公式ドキュメントの手順に準拠）。
    """
    base64_bytes = urlsafe_b64encode(share_url.encode("utf-8"))
    base64_str = base64_bytes.decode("utf-8").rstrip("=")
    return f"u!{base64_str}"


_root_cache_lock = threading.Lock()
_root_cache: dict = {"drive_id": None, "item_id": None}


def _get_root_drive_item() -> tuple[str, str]:
    """
    環境変数 ONEDRIVE_SHARE_URL が指すフォルダの driveId・itemId を解決する。
    共有リンクの解決結果は実行中変化しない前提のため、プロセス内で
    一度解決したらキャッシュする。

    戻り値: (drive_id, item_id)
    """
    with _root_cache_lock:
        if _root_cache["drive_id"] and _root_cache["item_id"]:
            return _root_cache["drive_id"], _root_cache["item_id"]

        share_url = os.environ["ONEDRIVE_SHARE_URL"]
        encoded = _encode_sharing_url(share_url)

        response = requests.get(
            f"{_GRAPH_BASE_URL}/shares/{encoded}/driveItem",
            headers=_auth_headers(),
            timeout=_REQUEST_TIMEOUT_SECONDS,
        )
        if response.status_code != 200:
            raise OneDriveError(
                f"共有URLの解決に失敗しました（{response.status_code}）: {response.text}"
            )

        payload = response.json()
        drive_id = payload["parentReference"]["driveId"]
        item_id = payload["id"]

        _root_cache["drive_id"] = drive_id
        _root_cache["item_id"] = item_id

        return drive_id, item_id


# ============================================
# Graph API：パス解決の共通処理
# ============================================

def _find_child_by_name(drive_id: str, parent_item_id: str, name: str) -> Optional[dict]:
    """
    指定フォルダ（parent_item_id）直下から、指定名（大小文字区別なし）の
    子アイテム（ファイルまたはフォルダ）を1件探して返す。
    見つからない場合は None を返す。

    Graph APIの children 一覧はページングされる場合があるため、
    @odata.nextLink がある限り辿る。
    """
    url = f"{_GRAPH_BASE_URL}/drives/{drive_id}/items/{parent_item_id}/children"
    name_lower = name.lower()

    while url:
        response = requests.get(url, headers=_auth_headers(), timeout=_REQUEST_TIMEOUT_SECONDS)
        if response.status_code == 404:
            return None
        if response.status_code != 200:
            raise OneDriveError(
                f"フォルダ一覧の取得に失敗しました（{response.status_code}）: {response.text}"
            )

        payload = response.json()
        for item in payload.get("value", []):
            if item.get("name", "").lower() == name_lower:
                return item

        url = payload.get("@odata.nextLink")

    return None


def _resolve_path_to_item(relative_path: str) -> Optional[dict]:
    """
    共有ルートからの相対パス（例: "202608/xxx.xlsx"）を、
    パス構成要素ごとに1階層ずつたどってdriveItemを解決する。
    途中の階層が存在しない、または最終的に対象が見つからない場合は
    None を返す。
    """
    drive_id, root_item_id = _get_root_drive_item()

    parts = [p for p in relative_path.split("/") if p]
    current_item_id = root_item_id

    for part in parts:
        child = _find_child_by_name(drive_id, current_item_id, part)
        if child is None:
            return None
        current_item_id = child["id"]

    # ルート自体（parts が空）を指すケースには本モジュールでは対応しない
    if not parts:
        return None

    return child


# ============================================
# 5.2.3節相当：公開インターフェース（既存シグネチャを維持）
# ============================================

def folder_exists(folder_name: str) -> bool:
    """
    ルートディレクトリ直下に、指定した名前のフォルダが存在するかを確認する。
    例: folder_exists("202608") -> ルート直下の"202608"フォルダの有無
    """
    if not _graph_api_configured():
        return _local_folder_exists(folder_name)

    drive_id, root_item_id = _get_root_drive_item()
    child = _find_child_by_name(drive_id, root_item_id, folder_name)
    return child is not None and "folder" in child


def ensure_folder(folder_name: str) -> None:
    """
    指定した名前のフォルダをルートディレクトリ直下に作成する。
    既に存在する場合はエラーにせず、そのまま何もしない（冪等な作成処理。
    要件定義書4.2節「フォルダ作成の競合対策」、基本設計書6.3節参照）。
    """
    if not _graph_api_configured():
        _local_ensure_folder(folder_name)
        return

    drive_id, root_item_id = _get_root_drive_item()
    existing = _find_child_by_name(drive_id, root_item_id, folder_name)
    if existing is not None and "folder" in existing:
        return

    response = requests.post(
        f"{_GRAPH_BASE_URL}/drives/{drive_id}/items/{root_item_id}/children",
        headers={**_auth_headers(), "Content-Type": "application/json"},
        json={
            "name": folder_name,
            "folder": {},
            # 同時作成の競合時は失敗させず、既存フォルダをそのまま使う
            "@microsoft.graph.conflictBehavior": "fail",
        },
        timeout=_REQUEST_TIMEOUT_SECONDS,
    )
    if response.status_code in (200, 201):
        return
    if response.status_code == 409:
        # 競合（他リクエストが同時に同名フォルダを作成した）は許容する
        return
    raise OneDriveError(
        f"フォルダの作成に失敗しました（{response.status_code}）: {response.text}"
    )


def file_exists(relative_path: str) -> bool:
    """
    ルートディレクトリからの相対パスでファイルの存在を確認する。
    例: file_exists("202608/【会社名】勤務実績管理表_202608_山田太郎.xlsx")
    例: file_exists("【会社名】勤務実績管理表_テンプレート.xlsx")
    """
    if not _graph_api_configured():
        return _local_file_exists(relative_path)

    item = _resolve_path_to_item(relative_path)
    return item is not None and "file" in item


def download_file(relative_path: str, local_dest_path: str) -> None:
    """
    OneDrive上の指定ファイルを、ローカルの作業用パス（local_dest_path）へ
    コピーする（openpyxlで直接開けるようにするため、一度ローカルへ
    落とす方式に統一する）。

    存在しない場合は OneDriveError を送出する。
    """
    if not _graph_api_configured():
        _local_download_file(relative_path, local_dest_path)
        return

    item = _resolve_path_to_item(relative_path)
    if item is None or "file" not in item:
        raise OneDriveError(f"ファイルが見つかりません: {relative_path}")

    drive_id, _ = _get_root_drive_item()
    response = requests.get(
        f"{_GRAPH_BASE_URL}/drives/{drive_id}/items/{item['id']}/content",
        headers=_auth_headers(),
        timeout=_REQUEST_TIMEOUT_SECONDS,
    )
    if response.status_code != 200:
        raise OneDriveError(
            f"ファイルのダウンロードに失敗しました（{response.status_code}）: {relative_path}"
        )

    dest = Path(local_dest_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(response.content)


def upload_file(local_source_path: str, relative_path: str, overwrite: bool = True) -> None:
    """
    ローカルの作業用パス（local_source_path）にあるファイルを、
    OneDrive上の指定パス（relative_path）へアップロードする。

    overwrite=False の場合、アップロード先に既にファイルが存在すれば
    何もしない（同時アクセスによる重複生成対策。基本設計書6.3節・8.4節の
    「同名ファイルが既に存在する場合はエラーとせず、既存ファイルを使う」
    という冪等な仕様に対応する）。
    overwrite=True の場合は無条件に上書きする（勤怠データの保存時など、
    意図的に内容を更新したい場合に使う）。

    アップロード先の親フォルダが存在しない場合は OneDriveError を
    送出する（ensure_folder を先に呼ぶのは呼び出し側の責務。
    attendance_service.get_or_create_monthly_file を参照）。

    このAPIは250MBまでのファイルに対応する（driveItem PUT /content）。
    月次勤怠Excel・PDFはこの上限を大きく下回るため、アップロード
    セッション（分割アップロード）は実装しない。
    """
    if not _graph_api_configured():
        _local_upload_file(local_source_path, relative_path, overwrite=overwrite)
        return

    if not overwrite and file_exists(relative_path):
        # 既に存在する場合は何もしない（先着優先、エラーにしない）
        return

    parts = [p for p in relative_path.split("/") if p]
    if not parts:
        raise OneDriveError(f"不正なアップロード先パスです: {relative_path!r}")

    file_name = parts[-1]
    parent_path = "/".join(parts[:-1])

    drive_id, root_item_id = _get_root_drive_item()

    if parent_path:
        parent_item = _resolve_path_to_item(parent_path)
        if parent_item is None or "folder" not in parent_item:
            raise OneDriveError(
                f"アップロード先の親フォルダが見つかりません: {parent_path}"
            )
        parent_item_id = parent_item["id"]
    else:
        parent_item_id = root_item_id

    content_type = mimetypes.guess_type(file_name)[0] or "application/octet-stream"
    file_bytes = Path(local_source_path).read_bytes()

    response = requests.put(
        f"{_GRAPH_BASE_URL}/drives/{drive_id}/items/{parent_item_id}:/{file_name}:/content",
        headers={**_auth_headers(), "Content-Type": content_type},
        data=file_bytes,
        timeout=_REQUEST_TIMEOUT_SECONDS,
    )
    if response.status_code not in (200, 201):
        raise OneDriveError(
            f"ファイルのアップロードに失敗しました（{response.status_code}）: {response.text}"
        )


def list_files(folder_name: str) -> list[str]:
    """
    指定フォルダ直下のファイル名一覧を返す（拡張子含むファイル名のみ、
    サブフォルダは含めない）。フォルダが存在しない場合は空リストを返す。
    """
    if not _graph_api_configured():
        return _local_list_files(folder_name)

    drive_id, root_item_id = _get_root_drive_item()
    folder_item = _find_child_by_name(drive_id, root_item_id, folder_name)
    if folder_item is None or "folder" not in folder_item:
        return []

    names: list[str] = []
    url = f"{_GRAPH_BASE_URL}/drives/{drive_id}/items/{folder_item['id']}/children"
    while url:
        response = requests.get(url, headers=_auth_headers(), timeout=_REQUEST_TIMEOUT_SECONDS)
        if response.status_code != 200:
            raise OneDriveError(
                f"フォルダ一覧の取得に失敗しました（{response.status_code}）: {response.text}"
            )
        payload = response.json()
        for item in payload.get("value", []):
            if "file" in item:
                names.append(item["name"])
        url = payload.get("@odata.nextLink")

    return sorted(names)


def get_last_modified(relative_path: str) -> Optional[str]:
    """
    指定パスのファイルの最終更新日時を表す文字列（ETag等、
    OneDrive APIが返す一意な識別子）を返す。
    ファイルが存在しない場合は None を返す。

    楽観的排他制御（同時保存対策）で、読み込み時と保存直前でこの値を
    比較し、異なっていれば「他の人が更新した」と判断するために使う
    （基本設計書8.4節相当）。

    【呼び出し側での利用箇所は現時点でまだ存在しない】
    """
    if not _graph_api_configured():
        return None

    item = _resolve_path_to_item(relative_path)
    if item is None:
        return None
    return item.get("eTag")


# ============================================
# ローカル疑似OneDrive実装（環境変数未設定時のフォールバック）
#
# Day3時点の暫定実装をそのまま残してある。ローカル開発・単体テスト用途。
# 本番運用ではONEDRIVE_TENANT_ID等の環境変数を設定し、Graph API実装
# （上記）が使われるようにすること。
# ============================================

def _local_root_dir() -> Path:
    root = os.environ.get("ONEDRIVE_LOCAL_ROOT", "./data/onedrive_local")
    root_path = Path(root)
    root_path.mkdir(parents=True, exist_ok=True)
    return root_path


def _local_folder_exists(folder_name: str) -> bool:
    target = _local_root_dir() / folder_name
    return target.is_dir()


def _local_ensure_folder(folder_name: str) -> None:
    target = _local_root_dir() / folder_name
    target.mkdir(parents=True, exist_ok=True)


def _local_file_exists(relative_path: str) -> bool:
    target = _local_root_dir() / relative_path
    return target.is_file()


def _local_download_file(relative_path: str, local_dest_path: str) -> None:
    source = _local_root_dir() / relative_path
    if not source.is_file():
        raise OneDriveError(f"ファイルが見つかりません: {relative_path}")

    dest = Path(local_dest_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, dest)


def _local_upload_file(local_source_path: str, relative_path: str, overwrite: bool) -> None:
    dest = _local_root_dir() / relative_path

    if not overwrite and dest.is_file():
        return

    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(local_source_path, dest)


def _local_list_files(folder_name: str) -> list[str]:
    target = _local_root_dir() / folder_name
    if not target.is_dir():
        return []
    return sorted(p.name for p in target.iterdir() if p.is_file())