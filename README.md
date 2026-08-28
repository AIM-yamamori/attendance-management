# 勤怠管理システム

Streamlit（Python）一体型の勤怠管理Webアプリケーション。ユーザーはブラウザから月次勤怠を入力・閲覧でき、管理者（admin）は全ユーザーの勤怠閲覧・編集・PDF出力・編集ロック管理を行う。

勤怠データ本体は月次Excelファイルとして OneDrive 上で管理し、ユーザー情報・ロック状態・システム設定は SQLite で管理する。

詳細な仕様は [`勤怠管理システム_基本設計書_v0.7.md`](./勤怠管理システム_基本設計書_v0.7.md)・要件定義書を参照。本READMEはセットアップ・起動手順に特化する。

---

## 目次

- [勤怠管理システム](#勤怠管理システム)
  - [目次](#目次)
  - [1. 前提条件](#1-前提条件)
  - [2. ディレクトリ構成](#2-ディレクトリ構成)
  - [3. セットアップ手順（Docker）](#3-セットアップ手順docker)
  - [4. .env 設定項目](#4-env-設定項目)
    - [4.1 最小構成（ローカル疑似OneDriveで動かす場合）](#41-最小構成ローカル疑似onedriveで動かす場合)
    - [4.2 本番構成（実OneDrive・Graph API連携）](#42-本番構成実onedrivegraph-api連携)
    - [4.3 各項目の詳細](#43-各項目の詳細)
  - [5. テンプレートExcelファイルの設置方法](#5-テンプレートexcelファイルの設置方法)
    - [5.1 ローカル疑似OneDriveの場合](#51-ローカル疑似onedriveの場合)
    - [5.2 実OneDrive（Graph API連携）の場合](#52-実onedrivegraph-api連携の場合)
    - [5.3 テンプレートファイルの中身の要件](#53-テンプレートファイルの中身の要件)
  - [6. Azure ADアプリ登録（本番OneDrive連携時のみ）](#6-azure-adアプリ登録本番onedrive連携時のみ)
  - [7. 初回起動・初期ログイン](#7-初回起動初期ログイン)
  - [8. Dockerを使わないローカル起動（開発用）](#8-dockerを使わないローカル起動開発用)
  - [9. データの永続化](#9-データの永続化)
  - [10. よくあるトラブル](#10-よくあるトラブル)
  - [11. 開発上の注意点](#11-開発上の注意点)

---

## 1. 前提条件

- Docker / Docker Compose が利用できること（推奨。13章参照の通りローカル・本番の環境差異を最小化するため）
- PDF出力機能を使うため、コンテナ内に LibreOffice（`libreoffice-calc` / `libreoffice-core`）・日本語フォント（`fonts-noto-cjk`）が必要（Dockerfileで自動導入される）
- 本番でOneDrive連携を行う場合、Microsoft 365（法人テナント）の管理者権限を持つ担当者による Azure AD アプリ登録が必要（6章参照）

---

## 2. ディレクトリ構成

```text
attendance-management/
├── app.py                          # エントリポイント（ログイン・画面遷移制御）
│
├── _pages/                         # 画面（"pages/"ではなく"_pages/"に注意。8.1節参照）
│   ├── 01_attendance_input.py      # 一般ユーザー：勤怠入力・閲覧
│   ├── 02_admin_attendance_edit.py # admin：勤怠閲覧・編集
│   ├── 03_admin_pdf_export.py      # admin：PDF出力
│   ├── 04_admin_lock_management.py # admin：ロック管理
│   ├── 05_admin_user_management.py # admin：ユーザー管理
│   └── 06_password_change.py       # 全員：パスワード変更
│
├── services/                       # 業務ロジック層
├── adapters/                       # DB・OneDrive・Excel入出力層
├── models/
│   └── schema.sql                  # SQLiteスキーマ定義
│
├── data/                           # 永続化ディレクトリ（Dockerボリューム）
│   ├── app.db                      # SQLiteファイル（自動生成）
│   └── onedrive_local/             # ローカル疑似OneDriveのルート（後述）
│
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .env                            # 実際の設定値（gitには含めない）
└── .env.example                    # 設定項目のひな形（本README対応）
```

---

## 3. セットアップ手順（Docker）

```bash
# 1. リポジトリを取得
git clone <このリポジトリのURL>
cd attendance-management

# 2. .env を作成し、4章の内容に従って値を設定する
cp .env.example .env
vi .env   # または任意のエディタで編集

# 3. テンプレートExcelファイルを配置する（5章を必ず参照）
#    ローカル疑似OneDriveの場合はここでファイルを置く
mkdir -p data/onedrive_local
cp /path/to/【会社名】勤務実績管理表_テンプレート.xlsx \
   "data/onedrive_local/${COMPANY_NAME}勤務実績管理表_テンプレート.xlsx"

# 4. ビルド・起動
docker compose up -d --build

# 5. ブラウザでアクセス
open http://localhost:8501
```

初回起動時、`app.py`の`_init_app()`が以下を自動的に行う（何度起動しても副作用が出ない冪等処理）。

- SQLiteスキーマの作成（`models/schema.sql`を実行）
- adminアカウント（`employee_id="admin"`）が存在しなければ、`.env`の`INITIAL_ADMIN_PASSWORD`で1件だけ投入
- `.env`の`SERVICE_START_MONTH`が設定されていれば、`settings`テーブルに未設定の場合のみ投入

---

## 4. .env 設定項目

`.env.example` をコピーして `.env` を作成し、値を埋める。`.env` は機密情報（クライアントシークレット等）を含むため、**リポジトリにコミットしないこと**（`.gitignore`に追加済みであることを確認）。

### 4.1 最小構成（ローカル疑似OneDriveで動かす場合）

開発・検証時など、実際のOneDriveに接続せずに動かしたい場合は、Graph API関連の4項目（`ONEDRIVE_SHARE_URL` / `ONEDRIVE_TENANT_ID` / `ONEDRIVE_CLIENT_ID` / `ONEDRIVE_CLIENT_SECRET`）を**すべて空欄のまま**にする。1つでも空欄があれば自動的にローカルディレクトリを疑似OneDriveとして使うフォールバック実装に切り替わる（`adapters/onedrive_adapter.py`の`_graph_api_configured()`）。

```dotenv
ONEDRIVE_SHARE_URL=
DB_PATH=./data/app.db
SERVICE_START_MONTH=202608
INITIAL_ADMIN_PASSWORD=ChangeMe123!
ONEDRIVE_TENANT_ID=
ONEDRIVE_CLIENT_ID=
ONEDRIVE_CLIENT_SECRET=
ONEDRIVE_LOCAL_ROOT=./data/onedrive_local
COMPANY_NAME=サンプル株式会社
```

### 4.2 本番構成（実OneDrive・Graph API連携）

4項目すべてを設定して初めてGraph API本番実装が有効になる。1つでも欠けるとローカル疑似OneDriveにフォールバックするため、本番投入時は必ず4つとも埋まっていることを確認する。

```dotenv
ONEDRIVE_SHARE_URL=https://xxxxx.sharepoint.com/:f:/g/xxxxxxxxxxxxxxxxxxxxxxx
DB_PATH=./data/app.db
SERVICE_START_MONTH=202608
INITIAL_ADMIN_PASSWORD=ChangeMe123!
ONEDRIVE_TENANT_ID=00000000-0000-0000-0000-000000000000
ONEDRIVE_CLIENT_ID=11111111-1111-1111-1111-111111111111
ONEDRIVE_CLIENT_SECRET=****************
ONEDRIVE_LOCAL_ROOT=
COMPANY_NAME=株式会社サンプル
```

### 4.3 各項目の詳細

| 変数名 | 必須 | 内容 | 備考 |
|---|---|---|---|
| `ONEDRIVE_SHARE_URL` | 本番連携時のみ | OneDriveの共有URL。**「リンクを知っている全員が編集可能」な、ルートディレクトリ1つに対する共有リンク**を発行し設定する | 共有リンクの発行はOneDrive管理者（依頼者側）が事前に行う。5.2節参照 |
| `DB_PATH` | 任意 | SQLiteファイルのパス | 未設定時は`./data/app.db` |
| `SERVICE_START_MONTH` | 任意（推奨） | サービス開始月（`YYYYMM`形式、例：`202608`）。これより前の月へはアクセスできない | 未設定時、選択可能な対象年月は「当月のみ」になる |
| `INITIAL_ADMIN_PASSWORD` | 任意（推奨） | 初回起動時のみ使用するadminの初期パスワード | 未設定時は`ChangeMe123`が使われる。**初回ログイン後は速やかにSC-07（パスワード変更画面）から変更すること**。パスワードポリシー（8文字以上・英大文字/英小文字/数字のうち2種類以上）に従うこと |
| `ONEDRIVE_TENANT_ID` | 本番連携時のみ | Azure AD テナントID | 6章参照 |
| `ONEDRIVE_CLIENT_ID` | 本番連携時のみ | 登録したアプリのクライアントID | 6章参照 |
| `ONEDRIVE_CLIENT_SECRET` | 本番連携時のみ | 発行したクライアントシークレット | 6章参照。有効期限があるため失効前に更新すること |
| `ONEDRIVE_LOCAL_ROOT` | 任意 | Graph API用の4項目が未設定の場合にのみ使われる、疑似OneDriveのルートディレクトリ | 未設定時は`./data/onedrive_local` |
| `COMPANY_NAME` | 任意（推奨） | 月次ファイル名・テンプレートファイル名の先頭に使う会社名 | 未設定時は`【会社名】`という文字列がそのままファイル名に使われてしまうため、**必ず設定すること** |

> `.env`を変更した場合は `docker compose up -d --build` （または `docker compose restart`）でコンテナに反映させること。

---

## 5. テンプレートExcelファイルの設置方法

月次勤怠ファイルは、ユーザーが対象月へ初回アクセスした時点で、このテンプレートファイルから自動生成される。**運用開始前に必ず配置しておくこと。** 配置されていない場合、月次ファイル未生成のユーザーが対象月にアクセスすると `FileNotFoundError`（「テンプレートファイルがOneDrive上に見つかりません」）が発生する。

テンプレートの相対パス・ファイル名は `COMPANY_NAME` 環境変数から自動的に決まる：

```text
{COMPANY_NAME}勤務実績管理表_テンプレート.xlsx
```

例：`.env`で `COMPANY_NAME=株式会社サンプル` の場合 → `株式会社サンプル勤務実績管理表_テンプレート.xlsx`

### 5.1 ローカル疑似OneDriveの場合

`ONEDRIVE_LOCAL_ROOT`（デフォルト `./data/onedrive_local`）の**直下**にファイルを置く。

```bash
mkdir -p data/onedrive_local
cp /path/to/元ファイル.xlsx \
   "data/onedrive_local/${COMPANY_NAME}勤務実績管理表_テンプレート.xlsx"
```

Docker Composeで `./data` をコンテナの `/app/data` にマウントしているため、ホスト側でこのディレクトリにファイルを置けばコンテナから参照できる。

### 5.2 実OneDrive（Graph API連携）の場合

`ONEDRIVE_SHARE_URL` が指すフォルダ（共有リンクのルートディレクトリ）の**直下**（月フォルダの外側）に配置する。

```text
OneDriveルートディレクトリ/（ONEDRIVE_SHARE_URLの起点）
  ├─ {COMPANY_NAME}勤務実績管理表_テンプレート.xlsx   ← ここに配置
  ├─ 202608/                                            ← 以降は自動生成される月フォルダ
  │   ├─ {COMPANY_NAME}勤務実績管理表_202608_山田太郎.xlsx
  │   └─ ...
  └─ ...
```

配置手順：

1. `.env`の`COMPANY_NAME`と完全に一致するファイル名（全角・半角、記号含め一字一句）でリネームする。
2. OneDrive管理者が、`ONEDRIVE_SHARE_URL`で発行した共有リンクの対象フォルダへ、Webブラウザまたはエクスプローラ（OneDrive同期クライアント）経由でアップロードする。
3. アプリからテンプレートが認識できるかは、SC-02（勤怠入力画面）で未生成の月を選択して確認するのが最も簡単（正常なら空欄の月次ファイルがその場で生成される）。

### 5.3 テンプレートファイルの中身の要件

テンプレートExcelファイルは以下のシート構成を満たしている必要がある（`adapters/excel_adapter.py`が前提とする構造）。

| シート名（Excel上の表記） | 用途 | 備考 |
|---|---|---|
| `原本` | 実データの入力・読み取り対象 | 月次ファイル生成時にこのシートがコピーされる |
| `記入例 ` | ヘルプ機能で参照する入力例（末尾に**半角スペースが1文字**必要） | シート名の末尾スペースが欠けていると読み込みエラーになる。テンプレート作成・リネーム時は要注意 |
| `変更履歴` | コピー対象外（存在してもよいが、月次ファイルには含まれない） | |

主要セル位置（`原本`シート）：

| セル | 内容 |
|---|---|
| C8 | 対象月初日（システムが自動設定） |
| O5 | 社員番号（システムが自動設定） |
| AH4 | 部署名（システムが自動設定） |
| AH5 | 氏名（システムが自動設定、全角スペース区切り） |
| D14:D44 | 休暇種類（プルダウン） |
| N列 | 始業時間 |
| Q列 | 終業時間 |
| AC列 | 離業時間 |
| AP列 | 自社工数内容 |

これら以外の書式・数式・罫線・列幅等はテンプレートの内容がそのまま維持される（システムは対象セルの値のみ書き換え、シート全体の再構築は行わない）。**テンプレート自体の内容変更（書式・数式修正等）は、本システムの画面からは行えない。** OneDrive上のテンプレートファイルを直接Excelで開いて編集すること。

---

## 6. Azure ADアプリ登録（本番OneDrive連携時のみ）

実OneDrive（Graph API）連携を行う場合、事前にAzure ADへのアプリ登録が必要。概要のみ記載する（詳細は組織のAzure AD管理者と調整すること）。

1. Azure Portal → 「Azure Active Directory」→「アプリの登録」→「新規登録」。
2. アプリケーション（クライアント）ID・ディレクトリ（テナント）IDを控える → `.env`の`ONEDRIVE_CLIENT_ID` / `ONEDRIVE_TENANT_ID`に設定。
3. 「証明書とシークレット」からクライアントシークレットを新規作成し、値を控える（**作成直後しか値を確認できない**） → `.env`の`ONEDRIVE_CLIENT_SECRET`に設定。
4. 「APIのアクセス許可」から Microsoft Graph の **アプリケーション権限** `Files.ReadWrite.All` を追加し、テナント管理者の同意を得る（「管理者の同意を与えます」ボタン）。
5. 対象のOneDriveフォルダ（ルートディレクトリ）に「リンクを知っている全員が編集可能」な共有リンクを発行し、そのURLを`.env`の`ONEDRIVE_SHARE_URL`に設定する。

> クライアントシークレットには有効期限がある（最大24ヶ月等）。失効前に更新し、`.env`を差し替えてコンテナを再起動すること。

---

## 7. 初回起動・初期ログイン

1. `docker compose up -d --build` でコンテナを起動する。
2. `http://localhost:8501`（本番では割り当てたドメイン）へアクセスする。
3. ログインID `admin`、パスワードは `.env` の `INITIAL_ADMIN_PASSWORD`（未設定時は `ChangeMe123`）でログインする。
4. 「パスワード変更」画面から、adminパスワードを速やかに変更する。
5. 「ユーザー管理」画面から、一般ユーザー（社員番号・姓名・部署名・初期パスワード）を登録する。
6. 5章の手順でテンプレートファイルが正しく配置されているか、一般ユーザーでログインして勤怠入力画面から確認する。

---

## 8. Dockerを使わないローカル起動（開発用）

Docker環境が使えない場合の参考手順（LibreOfficeが別途必要になるため、PDF出力機能の動作確認にはDocker利用を推奨）。

```bash
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# LibreOfficeをホストOSに別途インストールしておく（PDF出力機能を使う場合）

cp .env.example .env
# .env を編集

mkdir -p data/onedrive_local
# テンプレートファイルを data/onedrive_local/ に配置（5.1節参照）

streamlit run app.py
```

---

## 9. データの永続化

`docker-compose.yml` で `./data` をコンテナの `/app/data` にマウントしている。以下がこのディレクトリに保存される。

- `data/app.db` … SQLite（ユーザーマスタ・ロック状態・システム設定）
- `data/onedrive_local/` … ローカル疑似OneDrive使用時の勤怠Excelファイル群（本番でGraph API連携が有効な場合は使用されない）

**コンテナを再作成（`docker compose down` → `up`）してもこれらのデータは失われない。** バックアップを取る場合はホスト側の `./data` ディレクトリごとコピーすればよい。

実OneDrive連携時は、勤怠Excelファイル本体はOneDrive側で管理されるため、コンテナ側に永続化されるのは `data/app.db`（ユーザーマスタ・ロック状態）のみとなる。

---

## 10. よくあるトラブル

| 症状 | 想定原因 | 対処 |
|---|---|---|
| ログインできない | adminパスワードを忘れた／`INITIAL_ADMIN_PASSWORD`を変更したのに反映されない | `INITIAL_ADMIN_PASSWORD`はadminレコードが**存在しない場合のみ**投入される初期値。既にadminが存在する状態で`.env`を変更しても反映されない。SQLiteを直接操作するか、一度adminレコードを削除して再起動する |
| 勤怠入力画面で「テンプレートファイルがOneDrive上に見つかりません」 | テンプレート未配置、またはファイル名が`COMPANY_NAME`と不一致 | 5章の手順でファイル名・配置場所を確認する |
| 記入例（ヘルプ）表示でエラーになる | テンプレートの`記入例`シート名の末尾に半角スペースがない | シート名を`記入例 `（末尾半角スペース）に修正する（5.3節参照） |
| PDF出力が失敗する | コンテナ内にLibreOfficeが無い、または`soffice`コマンドがPATHにない | `docker compose build --no-cache` で再ビルドし、Dockerfileの`libreoffice-calc`等のインストールが成功しているか確認する |
| adminなのに「編集する」ボタンが押せない | 対象月・対象ユーザーがロックされていない状態 | admin編集は「ロック中」のファイルのみ可能な仕様（8.1節参照）。先にロック管理画面でロックしてから編集する |
| 一般ユーザーなのに「編集する」ボタンが押せない | 対象月がロック済み | adminにロック解除を依頼する |
| OneDrive接続がローカル疑似実装のままになる | Graph API用4環境変数のいずれかが空欄 | `.env`の`ONEDRIVE_SHARE_URL` / `ONEDRIVE_TENANT_ID` / `ONEDRIVE_CLIENT_ID` / `ONEDRIVE_CLIENT_SECRET`が全て設定されているか確認する |

---

## 11. 開発上の注意点

- 画面ファイルは `pages/` ではなく **`_pages/`**（アンダースコア始まり）に配置すること。Streamlitは`pages/`ディレクトリを自動検出してサイドバーに表示してしまうため、意図的に別名にしている（詳細は`app.py`のモジュールdocstring参照）。
- 業務ロジックは画面ファイル（`_pages/`）に書かず、`services/`層に委譲すること。
- Excelファイルへの書き込みは、対象セルの `.value` のみを更新し、書式（罫線・条件付き書式・数式等）を絶対に崩さないこと（`adapters/excel_adapter.py`参照）。
- 時刻入力欄は `H:MM` 形式のテキストボックス1つに統一されている（プルダウン方式ではない）。新規に時刻入力欄を追加する場合は、既存の `_render_time_text_input` パターンを踏襲すること。
- 現時点で admin 間（同一adminの複数タブ操作等）の楽観的排他制御は未実装であることに留意（基本設計書8.4節・16章参照）。