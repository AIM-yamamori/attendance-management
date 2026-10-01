# 勤怠管理システム

Streamlit（Python）一体型の勤怠管理Webアプリケーション。ユーザーはブラウザから月次勤怠を入力・閲覧でき、管理者（admin）は全ユーザーの勤怠閲覧・編集・PDF出力・編集ロック管理を行う。

勤怠データ本体は月次Excelファイルとして OneDrive 上で管理し、ユーザー情報・ロック状態・システム設定は SQLite で管理する。本番環境は Google Cloud Run 上で稼働し、SQLite ファイルは Litestream により Google Cloud Storage（GCS）へ継続的にレプリケーションされる（Cloud Run のファイルシステムはリクエスト間で永続化されないため）。

詳細な仕様は [`勤怠管理システム_基本設計書_v0.9.md`](./勤怠管理システム_基本設計書_v0.9.md)・要件定義書を参照。本READMEはセットアップ・起動手順に特化する。

---

## 目次

- [勤怠管理システム](#勤怠管理システム)
  - [目次](#目次)
  - [1. 前提条件](#1-前提条件)
  - [2. ディレクトリ構成](#2-ディレクトリ構成)
  - [3. アーキテクチャ概要（Cloud Run + Litestream）](#3-アーキテクチャ概要cloud-run--litestream)
  - [4. セットアップ手順（Docker / ローカル）](#4-セットアップ手順docker--ローカル)
  - [5. .env 設定項目](#5-env-設定項目)
    - [5.1 最小構成（ローカル疑似OneDriveで動かす場合）](#51-最小構成ローカル疑似onedriveで動かす場合)
    - [5.2 本番構成（実OneDrive・Graph API連携）](#52-本番構成実onedrivegraph-api連携)
    - [5.3 各項目の詳細](#53-各項目の詳細)
  - [6. テンプレートExcelファイルの設置方法](#6-テンプレートexcelファイルの設置方法)
    - [6.1 ローカル疑似OneDriveの場合](#61-ローカル疑似onedriveの場合)
    - [6.2 実OneDrive（Graph API連携）の場合](#62-実onedrivegraph-api連携の場合)
    - [6.3 テンプレートファイルの中身の要件](#63-テンプレートファイルの中身の要件)
  - [7. Azure ADアプリ登録（本番OneDrive連携時のみ）](#7-azure-adアプリ登録本番onedrive連携時のみ)
  - [8. Google Cloud（Litestream / GCS）側の事前準備](#8-google-cloudlitestream--gcs側の事前準備)
  - [9. 初回起動・初期ログイン](#9-初回起動初期ログイン)
  - [10. Dockerを使わないローカル起動（開発用）](#10-dockerを使わないローカル起動開発用)
  - [11. データの永続化](#11-データの永続化)
    - [ローカル（Docker Compose）](#ローカルdocker-compose)
    - [本番（Google Cloud Run + Litestream）](#本番google-cloud-run--litestream)
  - [12. adminパスワードの復旧（ADMIN\_RESET\_PASSWORD）](#12-adminパスワードの復旧admin_reset_password)
    - [仕組み](#仕組み)
    - [重要な注意点](#重要な注意点)
  - [13. よくあるトラブル](#13-よくあるトラブル)
  - [14. 開発上の注意点](#14-開発上の注意点)

---

## 1. 前提条件

- Docker / Docker Compose が利用できること（推奨。環境差異を最小化するため）
- PDF出力機能を使うため、コンテナ内に LibreOffice（`libreoffice-calc` / `libreoffice-core`）・日本語フォント（`fonts-noto-cjk`）が必要（Dockerfileで自動導入される）
- 本番でOneDrive連携を行う場合、Microsoft 365（法人テナント）の管理者権限を持つ担当者による Azure AD アプリ登録が必要（7章参照）
- 本番デプロイ（Google Cloud Run）を行う場合、GCPプロジェクトへのアクセス権と、SQLiteバックアップ用GCSバケットの作成権限が必要（8章参照）

---

## 2. ディレクトリ構成

```text
attendance-management/
├── app.py                          # エントリポイント（ログイン・画面遷移制御）
│
├── _pages/                         # 画面（"pages/"ではなく"_pages/"に注意。14章参照）
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
├── data/                           # ローカル開発時の永続化ディレクトリ（Dockerボリューム）
│   ├── app.db                      # SQLiteファイル（自動生成。本番は/tmp/app.dbを使用）
│   └── onedrive_local/             # ローカル疑似OneDriveのルート（後述）
│
├── requirements.txt
├── Dockerfile                      # LibreOffice + Litestreamを含むイメージ定義
├── docker-compose.yml
├── entrypoint.sh                   # DBリストア→Litestream replicate→Streamlit起動、SIGTERM時の後処理
├── litestream.yml                  # Litestreamのレプリケーション先（GCS）設定
├── .env                            # 実際の設定値（gitには含めない）
└── .env.example                    # 設定項目のひな形（本README対応）
```

---

## 3. アーキテクチャ概要（Cloud Run + Litestream）

本番環境は Google Cloud Run 上のコンテナとしてデプロイされる想定。Cloud Run のコンテナファイルシステムはリクエスト間・インスタンス再作成時に永続化されないため、SQLiteファイル（ユーザーマスタ・ロック状態・システム設定）は [Litestream](https://litestream.io/) を用いて GCS バケットへ継続的にレプリケーションする（同期間隔は`litestream.yml`の`sync-interval: 60s`、つまり60秒ごと）。

起動時の流れ（`entrypoint.sh`）：

1. コンテナ起動時、`DB_PATH`（デフォルト `/tmp/app.db`）にSQLiteファイルが既に存在するか確認する。
2. 存在しなければ、`litestream restore -if-replica-exists` により GCS 上の最新バックアップから復元を試みる。バックアップが無い場合（初回起動等）は、エラーにせず空の状態から起動を続行する。**復元コマンドが失敗した場合も同様に「新規DBで開始」として起動を続行する**（`|| echo "No existing backup found, starting fresh"`）ため、復元失敗に気づくにはコンテナログの確認が必要（13章参照）。
3. `litestream replicate -exec "streamlit run app.py ..."` の形で Streamlit プロセスを起動し、以後のSQLiteへの書き込みをLitestreamがバックグラウンドで継続的にGCSへレプリケーションする。

停止時の流れ：Cloud RunがコンテナへSIGTERMを送ると、`entrypoint.sh`の`trap`が受け取ってLitestreamプロセス（配下のStreamlitを含む）へSIGTERMを転送し、終了を待ってから`exit 0`する。これにより、停止直前の変更をGCSへ反映する機会をLitestreamに与える。

このため、Cloud Run のインスタンスが再作成されても、直近までの書き込みはGCS上のバックアップから復元される。ただし**厳密な意味での即時整合性・排他制御は提供されない**（Litestreamは非同期レプリケーションであり、異常終了時には最大で同期間隔分（約60秒）の更新が失われ得る。また11章・13章に記載の通り、admin間の楽観的排他制御自体が未実装であることに注意）。

勤怠Excelファイル本体（月次データ）はこの仕組みの対象外であり、引き続き OneDrive（またはローカル疑似OneDrive）側で管理される。

Cloud Run はコンテナが公開する `PORT` 環境変数（Cloud Runが自動的に設定する。未設定時のデフォルトは`8080`）でリッスンするアプリを要求するため、`entrypoint.sh`は`--server.port=${PORT:-8080}`でStreamlitを起動する。

> ローカルのDocker Compose環境（`docker-compose.yml`経由）でも同一のDockerfile・entrypoint.shが使われるため、ローカルでもLitestreamが動作する。GCSの認証情報が無い環境でローカル検証のみ行いたい場合は、Litestreamのレプリカ設定やGCS認証の扱いについて別途調整が必要になる点に留意（本READMEは本番運用を前提とした設定を記載する）。

---

## 4. セットアップ手順（Docker / ローカル）

```bash
# 1. リポジトリを取得
git clone <このリポジトリのURL>
cd attendance-management

# 2. .env を作成し、5章の内容に従って値を設定する
cp .env.example .env
vi .env   # または任意のエディタで編集

# 3. テンプレートExcelファイルを配置する（6章を必ず参照）
#    ローカル疑似OneDriveの場合はここでファイルを置く
mkdir -p data/onedrive_local
cp /path/to/【会社名】勤務実績管理表_テンプレート.xlsx \
   "data/onedrive_local/${COMPANY_NAME}勤務実績管理表_テンプレート.xlsx"

# 4. ビルド・起動
docker compose up -d --build

# 5. ブラウザでアクセス
open http://localhost:8501
```

> **ポートの注意**：`entrypoint.sh`は`PORT`環境変数（未設定時`8080`）でStreamlitを起動する。`docker-compose.yml`の公開ポート（`ports`）と、コンテナ内でStreamlitが待ち受けるポートが一致していないと、上記の`http://localhost:8501`に接続できない。接続できない場合は、`PORT=8501`を`.env`等で設定するか、`ports`を`"8501:8080"`のように調整する（現在の`docker-compose.yml`の設定に合わせること）。

初回起動時、`app.py`の`_init_app()`が以下を自動的に行う（何度起動しても副作用が出ない冪等処理）。

- SQLiteスキーマの作成（`models/schema.sql`を実行）
- adminアカウント（`employee_id="admin"`）が存在しなければ、`.env`の`INITIAL_ADMIN_PASSWORD`で1件だけ投入
- `.env`の`SERVICE_START_MONTH`が設定されていれば、`settings`テーブルに未設定の場合のみ投入

GCSからDBが復元された場合は、上記はいずれも「存在しなければ投入」のため復元済みのデータを上書きしない（復元済みadminのパスワードが`INITIAL_ADMIN_PASSWORD`で書き換わることはない）。

Google Cloud Run へのデプロイ手順（イメージのビルド・push・`gcloud run deploy`等）は組織のCI/CD構成に依存するため本READMEでは扱わない。Cloud Run側で最低限必要な設定は8章を参照。

---

## 5. .env 設定項目

`.env.example` をコピーして `.env` を作成し、値を埋める。`.env` は機密情報（クライアントシークレット等）を含むため、**リポジトリにコミットしないこと**（`.gitignore`に追加済みであることを確認）。

### 5.1 最小構成（ローカル疑似OneDriveで動かす場合）

開発・検証時など、実際のOneDriveに接続せずに動かしたい場合は、Graph API関連の4項目（`ONEDRIVE_SHARE_URL` / `ONEDRIVE_TENANT_ID` / `ONEDRIVE_CLIENT_ID` / `ONEDRIVE_CLIENT_SECRET`）を**すべて空欄のまま**にする。1つでも空欄があれば自動的にローカルディレクトリを疑似OneDriveとして使うフォールバック実装に切り替わる（`adapters/onedrive_adapter.py`の`_graph_api_configured()`）。

```dotenv
ONEDRIVE_SHARE_URL=
DB_PATH=./data/app.db
SERVICE_START_MONTH=202608
INITIAL_ADMIN_PASSWORD=ChangeMe123!
ADMIN_RESET_PASSWORD=
ONEDRIVE_TENANT_ID=
ONEDRIVE_CLIENT_ID=
ONEDRIVE_CLIENT_SECRET=
ONEDRIVE_LOCAL_ROOT=./data/onedrive_local
COMPANY_NAME=サンプル株式会社
```

> 上記の`DB_PATH=./data/app.db`は、Dockerを使わないローカル起動（10章）向けの値。Docker経由（`entrypoint.sh`）で起動する場合、`DB_PATH`が`litestream.yml`の`path`（`/tmp/app.db`）と食い違うとバックアップが取得されないため、`/tmp/app.db`のままにするか、`litestream.yml`側も揃えること（8章参照）。

### 5.2 本番構成（実OneDrive・Graph API連携）

4項目すべてを設定して初めてGraph API本番実装が有効になる。1つでも欠けるとローカル疑似OneDriveにフォールバックするため、本番投入時は必ず4つとも埋まっていることを確認する。本番はCloud Run上での稼働を想定しているため、`DB_PATH`はコンテナの書き込み可能領域である`/tmp`配下を指定する（8章参照）。

```dotenv
ONEDRIVE_SHARE_URL=https://xxxxx.sharepoint.com/:f:/g/xxxxxxxxxxxxxxxxxxxxxxx
DB_PATH=/tmp/app.db
SERVICE_START_MONTH=202608
INITIAL_ADMIN_PASSWORD=ChangeMe123!
ADMIN_RESET_PASSWORD=****************
ONEDRIVE_TENANT_ID=00000000-0000-0000-0000-000000000000
ONEDRIVE_CLIENT_ID=11111111-1111-1111-1111-111111111111
ONEDRIVE_CLIENT_SECRET=****************
ONEDRIVE_LOCAL_ROOT=
COMPANY_NAME=株式会社サンプル
```

### 5.3 各項目の詳細

| 変数名 | 必須 | 内容 | 備考 |
|---|---|---|---|
| `ONEDRIVE_SHARE_URL` | 本番連携時のみ | OneDriveの共有URL。**「リンクを知っている全員が編集可能」な、ルートディレクトリ1つに対する共有リンク**を発行し設定する | 共有リンクの発行はOneDrive管理者（依頼者側）が事前に行う。6.2節参照 |
| `DB_PATH` | 任意 | SQLiteファイルのパス | `entrypoint.sh`経由の起動では未設定時`/tmp/app.db`。アプリ単体（10章）では未設定時`./data/app.db`。**Docker／Cloud Run環境では`litestream.yml`の`path`と必ず一致させる**（3章・8章参照） |
| `SERVICE_START_MONTH` | 任意（推奨） | サービス開始月（`YYYYMM`形式、例：`202608`）。これより前の月へはアクセスできない | 未設定時、選択可能な対象年月は「当月のみ」になる |
| `INITIAL_ADMIN_PASSWORD` | 任意（推奨） | 初回起動時のみ使用するadminの初期パスワード | 未設定時は`ChangeMe123`が使われる。**初回ログイン後は速やかにSC-07（パスワード変更画面）から変更すること**。パスワードポリシー（8文字以上・英大文字/英小文字/数字のうち2種類以上）に従うこと |
| `ADMIN_RESET_PASSWORD` | 任意（推奨） | admin本人が現在のパスワードを忘れてログインできなくなった場合に使う、admin専用の復旧用パスワード | ログイン画面の「adminのパスワードを忘れた場合」から、この値を入力することで新しいadminパスワードを強制的に再設定できる（12章参照）。**この値自体はDBには保存されない**（環境変数として照合にのみ使われる）。未設定の場合はこの復旧機能自体が使用できない（「管理者パスワード復旧機能が設定されていません」エラーになる）。十分にランダムで推測されにくい長い値を設定し、`.env`にコミットしないこと |
| `ONEDRIVE_TENANT_ID` | 本番連携時のみ | Azure AD テナントID | 7章参照 |
| `ONEDRIVE_CLIENT_ID` | 本番連携時のみ | 登録したアプリのクライアントID | 7章参照 |
| `ONEDRIVE_CLIENT_SECRET` | 本番連携時のみ | 発行したクライアントシークレット | 7章参照。有効期限があるため失効前に更新すること |
| `ONEDRIVE_LOCAL_ROOT` | 任意 | Graph API用の4項目が未設定の場合にのみ使われる、疑似OneDriveのルートディレクトリ | 未設定時は`./data/onedrive_local` |
| `COMPANY_NAME` | 任意（推奨） | 月次ファイル名・テンプレートファイル名の先頭に使う会社名 | 未設定時は`【会社名】`という文字列がそのままファイル名に使われてしまうため、**必ず設定すること** |
| `PORT` | 任意（Cloud Runが自動設定） | Streamlitがリッスンするポート番号 | `entrypoint.sh`が参照する。未設定時のデフォルトは`8080`。Cloud Run上ではプラットフォームが自動的に注入するため、通常は手動設定不要。ローカルでは公開ポート設定と合わせること（3章・4章参照） |

> `.env`を変更した場合は `docker compose up -d --build` （または `docker compose restart`）でコンテナに反映させること。Cloud Run環境では、環境変数の変更後に新しいリビジョンをデプロイする必要がある。

---

## 6. テンプレートExcelファイルの設置方法

月次勤怠ファイルは、ユーザーが対象月へ初回アクセスした時点で、このテンプレートファイルから自動生成される。**運用開始前に必ず配置しておくこと。** 配置されていない場合、月次ファイル未生成のユーザーが対象月にアクセスすると `FileNotFoundError`（「テンプレートファイルがOneDrive上に見つかりません」）が発生する。

テンプレートの相対パス・ファイル名は `COMPANY_NAME` 環境変数から自動的に決まる：

```text
{COMPANY_NAME}勤務実績管理表_テンプレート.xlsx
```

例：`.env`で `COMPANY_NAME=株式会社サンプル` の場合 → `株式会社サンプル勤務実績管理表_テンプレート.xlsx`

### 6.1 ローカル疑似OneDriveの場合

`ONEDRIVE_LOCAL_ROOT`（デフォルト `./data/onedrive_local`）の**直下**にファイルを置く。

```bash
mkdir -p data/onedrive_local
cp /path/to/元ファイル.xlsx \
   "data/onedrive_local/${COMPANY_NAME}勤務実績管理表_テンプレート.xlsx"
```

Docker Composeで `./data` をコンテナの `/app/data` にマウントしているため、ホスト側でこのディレクトリにファイルを置けばコンテナから参照できる。

なお、この疑似OneDrive実装はローカル開発・検証用として引き続き利用可能であり、本番のGraph API連携とは独立して動作する（4項目のいずれかが空欄であれば自動的にこちらにフォールバックする。5.1節参照）。**Cloud Run上ではコンテナのローカル領域が消えるため、疑似OneDriveのデータはLitestreamの保護対象にもならない。本番では必ずGraph API連携を有効にすること。**

### 6.2 実OneDrive（Graph API連携）の場合

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

### 6.3 テンプレートファイルの中身の要件

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

## 7. Azure ADアプリ登録（本番OneDrive連携時のみ）

実OneDrive（Graph API）連携を行う場合、事前にAzure ADへのアプリ登録が必要。概要のみ記載する（詳細は組織のAzure AD管理者と調整すること）。

1. Azure Portal → 「Azure Active Directory」→「アプリの登録」→「新規登録」。
2. アプリケーション（クライアント）ID・ディレクトリ（テナント）IDを控える → `.env`の`ONEDRIVE_CLIENT_ID` / `ONEDRIVE_TENANT_ID`に設定。
3. 「証明書とシークレット」からクライアントシークレットを新規作成し、値を控える（**作成直後しか値を確認できない**） → `.env`の`ONEDRIVE_CLIENT_SECRET`に設定。
4. 「APIのアクセス許可」から Microsoft Graph の **アプリケーション権限** `Files.ReadWrite.All` を追加し、テナント管理者の同意を得る（「管理者の同意を与えます」ボタン）。
5. 対象のOneDriveフォルダ（ルートディレクトリ）に「リンクを知っている全員が編集可能」な共有リンクを発行し、そのURLを`.env`の`ONEDRIVE_SHARE_URL`に設定する。

> クライアントシークレットには有効期限がある（最大24ヶ月等）。失効前に更新し、`.env`を差し替えてコンテナを再起動すること（Cloud Run環境では新しいリビジョンをデプロイすること）。

---

## 8. Google Cloud（Litestream / GCS）側の事前準備

本番をGoogle Cloud Run上で稼働させる場合、事前に以下を用意しておく。

1. **GCSバケットの作成**：SQLiteのバックアップ先バケットを作成する（`litestream.yml`の`bucket`および`entrypoint.sh`の`GCS_REPLICA_URL`と一致させる）。

   ```bash
   gsutil mb -l asia-northeast1 gs://aim-kintai-sqlite-backup
   ```

2. **`litestream.yml` と `entrypoint.sh` の確認**：リポジトリ直下の`litestream.yml`でレプリケーション先を定義している。

   ```yaml
   dbs:
     - path: /tmp/app.db
       replicas:
         - type: gcs
           bucket: aim-kintai-sqlite-backup
           path: app-db-backup
           sync-interval: 60s
   ```

   - `path: /tmp/app.db` は`entrypoint.sh`の`DB_PATH`デフォルト値、および`.env`の`DB_PATH`と一致させること（不一致だと、実際のDBとは別のファイルを監視してしまい、バックアップが取得されない）。
   - **バケット名・保存先パスは`litestream.yml`（`bucket` / `path`）と`entrypoint.sh`内の`GCS_REPLICA_URL`（`gcs://<bucket>/<path>`）の2か所に書かれている。** 変更する場合は両方を同じ値に揃えた上でイメージを再ビルドする（片方だけ変えると、バックアップ先と復元元が食い違う）。
   - `sync-interval: 60s` により、GCSへの同期は60秒間隔で行われる。

3. **Cloud Run実行サービスアカウントへの権限付与**：Cloud Runサービスが使用するサービスアカウントに、対象GCSバケットへの読み書き権限（`roles/storage.objectAdmin`など）を付与する。Litestreamの認証はGCPの[Application Default Credentials（ADC）](https://cloud.google.com/docs/authentication/application-default-credentials)経由で行われるため、Cloud Run上ではサービスアカウントの権限設定のみで動作し、`.env`にGCS用の認証キーを別途記載する必要はない。

4. **Cloud Runサービスのデプロイ設定**：
   - コンテナポートはCloud Runが自動設定する`PORT`環境変数に従う（`entrypoint.sh`が`--server.port=${PORT:-8080}`で起動するため、Cloud Run側の追加設定は不要）。
   - `DB_PATH`は`/tmp/app.db`（またはそれに準じる書き込み可能パス。`litestream.yml`と揃えること）を指定する。
   - OneDrive連携用の4環境変数（`ONEDRIVE_SHARE_URL`等）や`INITIAL_ADMIN_PASSWORD`・`ADMIN_RESET_PASSWORD`等は、Cloud Runの環境変数、またはSecret Manager経由で設定する。
   - **最大インスタンス数は1にすることを推奨する。** Litestreamは1つのDBに対する書き込みプロセスが1つであることを前提としており、複数インスタンスが同時に起動する構成ではDBの不整合・レプリカの競合が起こり得る。

5. **動作確認**：デプロイ後、Cloud Runサービスへアクセスしてログインできることを確認する。インスタンスを一度スケールインさせた後に再度アクセスし、直前までのユーザーマスタ・ロック状態の変更がLitestreamのリストアにより復元されていることを確認しておくとよい。あわせてコンテナログに`=== DB RESTORE CHECK ===`以降のリストア結果（復元成功か、`starting fresh`か）が出力されていることを確認する。

---

## 9. 初回起動・初期ログイン

1. `docker compose up -d --build`（ローカル）または Cloud Run へのデプロイ（本番、8章参照）でアプリケーションを起動する。
2. `http://localhost:8501`（ローカル。ポートの注意は4章参照）または Cloud Run が割り当てたURL（本番）へアクセスする。
3. ログインID `admin`、パスワードは `.env`（またはCloud Runの環境変数）の `INITIAL_ADMIN_PASSWORD`（未設定時は `ChangeMe123`）でログインする。
4. 「パスワード変更」画面から、adminパスワードを速やかに変更する。
5. 「ユーザー管理」画面から、一般ユーザー（社員番号・姓名・部署名・初期パスワード）を登録する。
6. 6章の手順でテンプレートファイルが正しく配置されているか、一般ユーザーでログインして勤怠入力画面から確認する。

---

## 10. Dockerを使わないローカル起動（開発用）

Docker環境が使えない場合の参考手順（LibreOfficeが別途必要になるため、PDF出力機能の動作確認にはDocker利用を推奨）。この手順ではLitestreamは使用せず、SQLiteファイルをローカルディスクへ直接読み書きする。

```bash
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# LibreOfficeをホストOSに別途インストールしておく（PDF出力機能を使う場合）

cp .env.example .env
# .env を編集（DB_PATHはローカルパス、例：./data/app.db のままでよい）

mkdir -p data/onedrive_local
# テンプレートファイルを data/onedrive_local/ に配置（6.1節参照）

streamlit run app.py
```

---

## 11. データの永続化

### ローカル（Docker Compose）

`docker-compose.yml` で `./data` をコンテナの `/app/data` にマウントしている。以下がこのディレクトリに保存される。

- `data/app.db` … SQLite（ユーザーマスタ・ロック状態・システム設定）
- `data/onedrive_local/` … ローカル疑似OneDrive使用時の勤怠Excelファイル群（本番でGraph API連携が有効な場合は使用されない）

**コンテナを再作成（`docker compose down` → `up`）してもこれらのデータは失われない。** バックアップを取る場合はホスト側の `./data` ディレクトリごとコピーすればよい。

> `DB_PATH`を`/tmp/app.db`にしている場合（`entrypoint.sh`のデフォルト）、SQLiteは`./data`ではなくコンテナ内の`/tmp`に作られ、コンテナ再作成で失われる（Litestreamでの復元に依存する）。`./data`に永続化したい場合は`DB_PATH=/app/data/app.db`とし、`litestream.yml`の`path`も同じ値にそろえること。

### 本番（Google Cloud Run + Litestream）

Cloud Run上のコンテナファイルシステム（`/tmp`含む）はインスタンスの再作成時に消去されるため、`data/app.db`に相当するファイルはコンテナに永続化されない。代わりに、`entrypoint.sh`が起動時にGCSバケット（`litestream.yml`で指定）からリストアし、稼働中はLitestreamが60秒間隔で同バケットへレプリケーションする（3章・8章参照）。バックアップの実体はGCSバケット側にあるため、バックアップを取得・保全する場合はこのバケットを対象にする。

**復元されるのは最後にGCSへ同期された時点の状態**であり、異常終了時には同期間隔（60秒）以内の更新（ロック操作・ユーザー変更・パスワード変更等）が失われる可能性がある。通常の停止（SIGTERM）では、停止前に同期が行われる。

勤怠Excelファイル本体（月次データ）は、ローカル・本番いずれの場合も上記のLitestream/GCSの仕組みとは独立しており、引き続きOneDrive（またはローカル疑似OneDrive）側で管理される。実OneDrive連携時は、勤怠Excelファイル本体はOneDrive側で管理されるため、この仕組みで永続化する必要があるのはSQLite（ユーザーマスタ・ロック状態）のみとなる。

---

## 12. adminパスワードの復旧（ADMIN_RESET_PASSWORD）

admin本人が通常のパスワード変更を経ずにパスワードを紛失し、ログインできなくなった場合に備えて、環境変数 `ADMIN_RESET_PASSWORD` を用いた復旧経路を用意している。

### 仕組み

1. `.env`（またはCloud Runの環境変数）に、復旧用のパスワードとして `ADMIN_RESET_PASSWORD` を設定しておく。
2. ログイン画面の「adminのパスワードを忘れた場合」を開くと、以下の入力欄が表示される。
   - 復旧用パスワード（`ADMIN_RESET_PASSWORD`と一致する値を入力する）
   - 新しいadminパスワード
   - 新しいadminパスワード（確認）
3. 入力内容が正しければ、`auth_service.reset_admin_password_by_recovery()` が以下を行う。
   - 復旧用パスワードの照合
   - 新しいパスワードのパスワードポリシー確認（8文字以上、英大文字・英小文字・数字のうち2種類以上）
   - 新しいパスワードをbcryptでハッシュ化し、`users`テーブルの`employee_id="admin"`レコードの`password_hash`を更新

この操作はadminアカウント専用であり、一般ユーザーのパスワード復旧には使用できない（一般ユーザーのパスワード紛失時は、admin が「ユーザー管理」画面から再設定する）。

### 重要な注意点

- **`ADMIN_RESET_PASSWORD` の値自体はDBに保存されない。** 照合にのみ使われる環境変数であり、`users`テーブルに書き込まれるのはbcryptハッシュ化された新しいadminパスワードのみ。
- `ADMIN_RESET_PASSWORD` が未設定（空文字列）の場合、この復旧機能自体が使用できず、「管理者パスワード復旧機能が設定されていません」というエラーになる。
- `ADMIN_RESET_PASSWORD` は `INITIAL_ADMIN_PASSWORD`（初回起動時のみ使われる初期パスワード）とは別物であり、いつでも・何度でも使用できる。**十分にランダムで推測されにくい長い値を設定し、`.env`をリポジトリにコミットしないこと。**
- 復旧用パスワードの照合は通常の文字列比較で、**失敗回数の制限（ロックアウト等）は実装されていない**。ログイン画面から誰でも入力を試せるため、短い・推測しやすい値は設定しないこと。
- Cloud Run環境で`ADMIN_RESET_PASSWORD`を変更・追加した場合は、新しいリビジョンをデプロイする必要がある。
- この値を知っている人は誰でもadminパスワードを再設定できてしまうため、値の管理（誰が保持するか、どう伝達するか）は組織のルールに従って厳重に行うこと。

---

## 13. よくあるトラブル

| 症状 | 想定原因 | 対処 |
|---|---|---|
| adminがパスワードを忘れてログインできない | 通常のパスワード変更を経ずにパスワードを紛失した | ログイン画面の「adminのパスワードを忘れた場合」から、`ADMIN_RESET_PASSWORD`（環境変数）を使ってadminパスワードを再設定できる。12章参照。`ADMIN_RESET_PASSWORD`が未設定の場合はこの方法は使えない |
| ログインできない（admin以外、または上記の復旧機能が使えない場合） | `INITIAL_ADMIN_PASSWORD`を変更したのに反映されない、等 | `INITIAL_ADMIN_PASSWORD`はadminレコードが**存在しない場合のみ**投入される初期値。既にadminが存在する状態（GCSから復元された場合を含む）で`.env`（またはCloud Runの環境変数）を変更しても反映されない。SQLiteを直接操作するか、一度adminレコードを削除して再起動する |
| 勤怠入力画面で「テンプレートファイルがOneDrive上に見つかりません」 | テンプレート未配置、またはファイル名が`COMPANY_NAME`と不一致 | 6章の手順でファイル名・配置場所を確認する |
| 記入例（ヘルプ）表示でエラーになる | テンプレートの`記入例`シート名の末尾に半角スペースがない | シート名を`記入例 `（末尾半角スペース）に修正する（6.3節参照） |
| PDF出力が失敗する | コンテナ内にLibreOfficeが無い、または`soffice`コマンドがPATHにない | `docker compose build --no-cache` で再ビルドし、Dockerfileの`libreoffice-calc`等のインストールが成功しているか確認する |
| adminなのに「編集する」ボタンが押せない | 対象月・対象ユーザーがロックされていない状態 | admin編集は「ロック中」のファイルのみ可能な仕様（基本設計書3.5節・8章参照）。先にロック管理画面でロックしてから編集する |
| 一般ユーザーなのに「編集する」ボタンが押せない | 対象月がロック済み | adminにロック解除を依頼する |
| OneDrive接続がローカル疑似実装のままになる | Graph API用4環境変数のいずれかが空欄 | `.env`の`ONEDRIVE_SHARE_URL` / `ONEDRIVE_TENANT_ID` / `ONEDRIVE_CLIENT_ID` / `ONEDRIVE_CLIENT_SECRET`が全て設定されているか確認する |
| Cloud Run再起動後、ユーザー・ロック状態がすべて消えている／adminが初期状態に戻っている | リストア失敗でも`entrypoint.sh`は「新規DBで開始」として起動を続行するため、空のDBで起動している可能性がある | コンテナログで`=== DB RESTORE CHECK ===`以降を確認する（`Attempting restore from GCS...`の後に`No existing backup found, starting fresh`が出ていれば復元できていない）。GCSバケット名・パスが`litestream.yml`と`entrypoint.sh`で一致しているか、Cloud Runサービスアカウントに対象バケットへの読み書き権限があるか（8章参照）を確認する |
| 直前（数十秒以内）のユーザー登録・ロック状態変更だけが消えている | 同期間隔（60秒）以内にコンテナが異常終了し、GCSへ未反映だった | 仕様上の制約（3章・11章参照）。通常の停止（SIGTERM）では停止前に同期される |
| コンテナ起動時に`litestream: replica error`等が出る | GCSバケット名の不一致、認証情報不足、ネットワーク到達性の問題 | `litestream.yml`の`bucket`名が実在のバケットと一致しているか、サービスアカウントの権限（8章参照）を確認する |
| バックアップがGCSに作成されない | `litestream.yml`の`path`と実際のDBのパス（`DB_PATH`）が不一致 | `.env`の`DB_PATH`、`entrypoint.sh`のデフォルト、`litestream.yml`の`path`がすべて同じ値か確認する |
| ローカルで`http://localhost:8501`に接続できない | Streamlitの待ち受けポート（`PORT`、デフォルト8080）と`docker-compose.yml`の公開ポートが不一致 | 4章の「ポートの注意」を参照し、`PORT`または`ports`を調整する |
| ローカルでLitestream関連のエラーが出て起動しない | ローカル環境にGCS認証情報が無い状態でLitestreamがGCSへの接続を試みている | ローカル検証時のLitestream/GCS認証の扱いは別途調整が必要（3章の注記を参照）。応急的にはローカル専用の設定や認証情報のマウントを検討する |

---

## 14. 開発上の注意点

- 画面ファイルは `pages/` ではなく **`_pages/`**（アンダースコア始まり）に配置すること。Streamlitは`pages/`ディレクトリを自動検出してサイドバーに表示してしまうため、意図的に別名にしている（詳細は`app.py`のモジュールdocstring参照）。
- 業務ロジックは画面ファイル（`_pages/`）に書かず、`services/`層に委譲すること。
- Excelファイルへの書き込みは、対象セルの `.value` のみを更新し、書式（罫線・条件付き書式・数式等）を絶対に崩さないこと（`adapters/excel_adapter.py`参照）。
- 時刻入力欄は `H:MM` 形式のテキストボックス1つに統一されている（プルダウン方式ではない）。新規に時刻入力欄を追加する場合は、既存の `_render_time_text_input` パターンを踏襲すること。
- 現時点で admin 間（同一adminの複数タブ操作等）の楽観的排他制御は未実装であることに留意（基本設計書8.4節・16章参照）。Litestreamによるレプリケーションはバックアップ・障害復旧のための仕組みであり、この排他制御の欠如を解消するものではない。
- adminパスワード復旧機能（12章）は `auth_service.py` の `reset_admin_password_by_recovery()` が担う。`app.py`のログイン画面（`_render_login_page()`）から呼び出され、`os.environ.get("ADMIN_RESET_PASSWORD", "")`で取得した値との照合のみを行う設計であり、`db_adapter.py`側の変更は不要（既存の`change_password()` / `reset_password_by_admin()`と同様、`get_cursor(commit=True)`経由で`users.password_hash`を直接UPDATEする）。
- `litestream.yml`・`entrypoint.sh`を変更する場合は、DBパス（`/tmp/app.db`）とGCSのバケット・パスが複数ファイル（`.env` / `entrypoint.sh` / `litestream.yml`）に重複して書かれている点に注意し、すべて同じ値に揃えること（8章参照）。