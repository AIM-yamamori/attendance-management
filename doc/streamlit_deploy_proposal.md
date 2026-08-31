# Streamlitアプリのデプロイ方法について（運用コスト最優先）

## 背景・要件

現在のアプリは以下の構成で稼働している。

- **Streamlit**（Python製のWebアプリフレームワーク）
- **Docker**でコンテナ化
- **SQL**（DBを使用。RDBMS種別は別途確認が必要だが、以下ではPostgreSQL/MySQL相当を想定）
- **LibreOffice**（内部でドキュメント変換処理に使用）

Dockerfileの抜粋:

```dockerfile
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-core \
    fonts-noto-cjk \
    locales \
    && localedef -i ja_JP -c -f UTF-8 -A /usr/share/locale/locale.alias ja_JP.UTF-8 \
    && rm -rf /var/lib/apt/lists/*
```

**最優先事項：運用コストを最小限に抑えること。** 以下の検討・比較は、この前提を軸に整理する。

この構成のため、単純な静的サイトやサーバーレス関数向けのホスティング（Cloudflare Workers/Pages等）は不向きであり、**Dockerコンテナをそのまま動かせる環境**（＋DB）を選定する必要がある。

Cloudflare自体はStreamlit＋LibreOfficeの実行基盤にはならないため、Cloudflareは「独自ドメイン管理・CDN・セキュリティ（WAF等）」の役割として前段に置き、アプリ本体は別のホスティング先で稼働させる構成が現実的。

---

## コスト最優先で考える：Cloudflareを「使う場合」「使わない場合」

運用コストを最小化する観点では、まず「Cloudflareをどう位置づけるか」で3パターンに分かれる。

### パターンA: Cloudflareを一切使わない

ドメイン・DNS・アプリ実行のすべてを1つのサービス（VPS等）に集約する構成。管理対象が減り、最もシンプル。

- **メリット:** 契約先が1つで済み、管理画面・請求も一本化。小規模なら最安になりやすい。
- **デメリット:** CDN・WAF（不正アクセス対策）がない。DDoS攻撃を受けた場合の耐性が低い。SSL証明書は自分で用意（Let's Encrypt等、無料だが更新設定が必要）。

### パターンB: Cloudflareを「窓口」としてのみ使う（無料枠で十分）

ドメイン管理・DNS・CDN・WAFはCloudflareの**無料プラン**を使い、アプリ本体は外部ホスティングで動かす構成。

- **メリット:** CDN・WAF・SSL自動発行が**無料で**付与される。アプリ側の帯域負荷が減り、結果的にアプリ側のコストが下がるケースもある。
- **デメリット:** 特になし（無料プランの範囲で完結するため、コスト増要因はほぼない）。**この位置づけが最もコスパが良い。**

### パターンC: Cloudflareの中でアプリ本体も動かす（Cloudflare Containers）

ドメイン・CDN・WAF・アプリ実行のすべてをCloudflareに集約する構成。

- **メリット:** 契約先を完全に一本化できる。
- **デメリット:** Containersは**従量課金**（Active-CPU課金）で、常時稼働に近い使い方だとコストが読みにくい。DBも別途D1やHyperdrive経由の外部DB接続が必要になり、既存のSQL構成をそのまま持ち込めない可能性がある。新しいサービスで検証コストもかかる。

### コスト最優先での結論

**運用コストを最小化したいなら、パターンB（Cloudflareは無料の窓口としてのみ使う）が最有力。** Cloudflare自体の利用は無料プランの範囲に収め、コストがかかる「アプリ実行部分」を最安のホスティング先に任せるのが合理的。パターンCは魅力的に見えるが、従量課金の性質上、常時稼働するStreamlitアプリでは想定より高くつく可能性があり、コスト最優先の方針とは相性が良くない。

---

一口に「Cloudflare」と言っても、内部には複数のサービスがあり、**Docker対応・非対応がサービスごとに異なる**。まずここを整理する。

| サービス | 役割 | Docker対応 | 備考 |
|---|---|---|---|
| **Cloudflare Pages** | 静的サイト・JAMstackサイトのホスティング | **× 非対応** | 静的ファイル配信＋簡易的な関数（Pages Functions）のみ。Pages Functionsの正体はWorkersであり、Dockerは動かせない |
| **Cloudflare Workers** | サーバーレス関数の実行環境（V8/JSベース） | **× 非対応** | Dockerコンテナは実行不可。JavaScript/TypeScript/Wasmのみ動作する軽量isolate |
| **Cloudflare Containers** | Workers上でDockerコンテナを起動する拡張機能 | **○ 対応**（2026年4月GA） | DockerfileやDocker Hub上のイメージをそのままビルド・実行可能。ただし常駐性・永続化に制限あり（後述） |
| **Cloudflare Workers Sandbox SDK** | AIエージェント向けの隔離実行環境 | **○ 対応** | 内部的にはContainersと同じ基盤。コード実行・ファイル操作に特化しており、Webアプリのホスティング用途ではない |
| **Cloudflare Tunnel** | 自社サーバーとCloudflareを接続する経路 | **該当なし（Docker実行環境ではない）** | Dockerは動かさず、あくまで「外部サーバーへの通信経路」を提供するのみ。アプリ実行は別途VPS等が必要 |
| **Cloudflare R2 / D1 / KV** | ストレージ・データベース | **該当なし（実行環境ではない）** | データの保存先であり、アプリ実行環境ではない |

### ポイント

- **PagesとWorkersは非対応**。これらは「軽量な関数・静的コンテンツ配信」に特化しており、Dockerコンテナという概念自体を扱えない。
- **Containersのみが対応**。しかも2026年4月GAとまだ新しく、実質的に「CloudflareでDockerを動かす唯一の方法」はこのContainersに限られる。
- **Tunnelは実行環境ではなく通信経路**。「Cloudflare Tunnel + VPS」の構成で語られることが多いが、これはDockerをCloudflare上で動かしているのではなく、あくまで自社VPS上でDockerを動かし、Cloudflareは通信の橋渡しをしているだけという点に注意。

つまり、「Cloudflareの中でDockerを動かせるかどうか」は一枚岩ではなく、**Containersという特定のサービスに限られる**、というのが正確な理解になる。

---

## なぜCloudflare上でアプリを実行できないのか（Workers/Pagesの場合）

Cloudflareが提供する実行環境（Workers / Pages）は、Streamlit・LibreOfficeのそれぞれに対して個別に非対応の理由がある。

### Streamlitが動かない理由

| 観点 | Streamlitの要件 | Cloudflare Workersの制約 |
|---|---|---|
| 実行時間 | 常駐プロセスとしてWebサーバー（Tornado）がずっと起動し続ける | 1リクエストごとに起動・終了する短命な実行モデル（無制限の常駐不可） |
| 通信方式 | WebSocketで画面と常時接続し、状態をやり取りする | WebSocketの双方向常時接続を前提とした常駐処理には非対応 |
| 実行言語・ランタイム | CPython（標準のPython処理系）がフル機能で必要 | Workersは軽量なV8（JavaScript/Wasm）ベースで、フルのPython常駐実行は不可 |
| セッション状態 | サーバー側でユーザーごとのセッション状態をメモリ保持 | Workersはステートレス設計が前提で、長時間のメモリ保持に向かない |

→ Streamlitは「サーバーがブラウザとつながりっぱなしで状態を持ち続ける」設計そのものが、Cloudflare Workersの「短時間だけ起動して終了する」実行モデルと根本的に矛盾する。

### LibreOfficeが動かない理由

| 観点 | LibreOfficeの要件 | Cloudflare Workersの制約 |
|---|---|---|
| 実行形態 | OSにインストールされたネイティブバイナリ（Cバイナリ）を`soffice`コマンド等で起動 | Workersはネイティブバイナリの実行不可（JS/Wasmのみ） |
| OS依存パッケージ | `apt-get install`でOSレベルのパッケージ・フォント・ロケールを導入 | WorkersにOSパッケージ管理の概念がなく、`apt`のような導入手段が存在しない |
| メモリ・起動コスト | 起動だけで数百MB規模のメモリを消費し、起動に数秒かかる | Workersは軽量・高速起動が前提で、CPU時間やメモリに厳しい制限がある |
| ファイルシステム | 変換処理で一時ファイルの読み書きが必要 | Workersのファイルシステムアクセスは極めて限定的 |

→ LibreOfficeは「OS上に展開された重量級のネイティブアプリケーション」であり、Cloudflare Workersが想定する「軽量なJS/Wasm関数」の実行モデルとは前提が全く異なる。

### 結論（Cloudflare Workersの場合）

Streamlit単体でもCloudflare Workers（軽量なJS/Wasm実行環境）では動作せず、LibreOfficeを組み合わせるとさらにその差が広がる。ただし、**Cloudflareには2026年4月に正式リリースされた「Cloudflare Containers」という別プロダクトがあり、Dockerイメージをそのまま動かせる**。次章ではCloudflare関連サービスのみで構築する場合の可否を、要素ごとに詳しく検討する。

---

## Cloudflare関連サービスのみで構築したい場合の検討

「外部のCloud Run等を使わず、Cloudflareの製品だけで完結させたい」という要望に対して、要素ごとに障壁と解決可否を整理する。

### 前提：Cloudflare Containersとは

2026年4月13日にGA（正式リリース）された比較的新しいサービスで、Workersから呼び出す形でDockerコンテナを起動できる。**これにより「Dockerが使えない」という制約自体はほぼ解消されている。** ただし、まだ新しいサービスゆえの制約がいくつか残っている。

### 要素ごとの障壁と解決可否

| 要素 | 障壁 | 解決可能か | 詳細 |
|---|---|---|---|
| **Dockerの実行** | 旧来のWorkers/Pagesはコンテナ非対応 | **○ 解決済み** | Cloudflare Containers（2026年4月GA）でDockerfileをそのままビルド・デプロイ可能。Docker Hub等からのイメージ取得にも対応 |
| **LibreOfficeのインストール** | `apt-get install`のようなOS依存パッケージ導入 | **○ 解決可能** | Containersは通常のLinuxコンテナのため、Dockerfile内で`apt-get`を使うこと自体は問題ない |
| **常駐サーバー（Streamlitのステートフルな性質）** | Workersは短命な実行が前提 | **△ 部分的に解決** | ContainersはWorkersとは別に「Durable Object」という仕組み経由で管理され、数十分程度の継続動作は可能。ただしデフォルトで無操作10分後にスリープする仕様があり、完全な「常時起動サーバー」とは性質が異なる |
| **永続ストレージ** | コンテナのディスクは再起動のたびに初期化される（エフェメラル） | **△ 制限あり** | ディスクは基本的に使い捨て。永続化にはR2（オブジェクトストレージ）をFUSEでマウントする方法があるが、SSD相当の速度は出ない。DBが必要ならD1（SQLite）やR2との組み合わせが必要 |
| **オートスケール** | 負荷に応じた自動スケーリングの仕組みが未成熟 | **△ 制限あり** | 現状は「コードで明示的にスケール数を指定する」形式が中心で、Cloud Run等のような自動負荷分散はまだ発展途上 |
| **本番運用の実績・情報量** | 2026年GAの新しいサービス | **△ 要検証** | 情報・実績がCloud Run等に比べて少なく、LibreOfficeのような重量級アプリでの動作実績も限定的。事前の検証（PoC）が必須 |

### 代替手段（Cloudflare関連のみで構築する場合の対応策）

上記の障壁を踏まえ、Cloudflareサービスのみで構築する場合の対応策は以下の通り。

- **常駐性の懸念への対応:** `onActivityExpired()`フックをカスタマイズし、アクセスがある間はスリープさせない設定にする。ただし完全な保証はないため、重要な処理中に停止しないかは要検証。
- **永続化が必要なデータへの対応:** セッション情報やアップロードファイルなど残したいデータはコンテナ内ディスクに置かず、**R2（オブジェクトストレージ）**または**D1（Cloudflareのマネージド SQLite）**に保存する設計に変更する。
- **スケーリングへの対応:** アクセス数が想定できる場合は、コード側で必要数のコンテナインスタンスを明示的に起動する設計にする（自動スケールに依存しない）。
- **未成熟な部分への対応:** 本番投入前に、実際のLibreOffice変換処理を含むPoC（小規模な動作検証）を実施し、10分スリープの影響やメモリ上限（インスタンスタイプにより変動）で問題が出ないか確認する。

### この案のまとめ

**結論として、「Cloudflare関連サービスのみでの構築」は技術的には可能になりつつある**（Cloudflare Containersにより）。ただし、以下の点でCloud Run等の成熟したコンテナサービスと比べるとまだ不利がある。

- サービス自体が新しく（2026年4月GA）、実績・ドキュメントが少ない
- 常駐性やオートスケールの挙動が発展途上
- 永続化にはR2/D1への設計変更が必要になり、既存のDockerfileそのままでは完結しない可能性がある

「Cloudflareのみで完結させたい」という要望が強い場合は、**まず小規模なPoCでCloudflare Containers上にLibreOffice入りのコンテナを立て、実際の変換処理や10分スリープの挙動を検証してから本格導入を判断する**、という進め方を提案したい。

---

## 役割分担の考え方（共通構成）

どの案を選んでも、基本的な役割分担は以下のように整理できる。

| レイヤー | 担当 | 内容 |
|---|---|---|
| ドメイン・DNS | **Cloudflare** | 独自ドメインの管理、DNSレコード設定 |
| CDN・キャッシュ | **Cloudflare** | 静的アセットの配信高速化 |
| セキュリティ（WAF・DDoS対策） | **Cloudflare** | 不正アクセス対策、アクセス制限 |
| SSL証明書 | **Cloudflare**（プロキシ利用時） | HTTPS化 |
| アプリ実行環境 | **各ホスティング先**（Cloud Run等） | Streamlit＋LibreOfficeのコンテナ実行 |
| オートスケール・可用性 | **各ホスティング先** | アクセス増減に応じたリソース調整 |
| サーバー・OSの保守 | **各ホスティング先 or 自社**（VPSの場合） | OSアップデート、障害対応 |

**重要:** Cloudflareは「入口の窓口」であり、Streamlit本体やLibreOfficeの処理そのものは一切担わない。アプリの実行・スケーリング・保守は必ず別のホスティング先が担当する。

---

## 比較表

| 項目 | Cloud Run / Fargate | VPS + Docker + Tunnel | Fly.io / Railway | Streamlit Community Cloud | Cloudflare Containers |
|---|---|---|---|---|---|
| Streamlit | ○ Dockerfileそのまま可 | ○ Dockerfileそのまま可 | ○ Dockerfileそのまま可 | ○（専用サービス） | △ 可能だが常駐性に制限あり（要検証） |
| LibreOffice | ○ apt導入OK | ○ apt導入OK | ○ apt導入OK | △ packages.txtで一部可／重い依存は不安定 | △ apt導入は可能／実績少なく要検証 |
| SQL（DB） | ○ Cloud SQL等と連携、同一コンテナ内も可 | ○ 同一VPS内にDBを同居可能（追加費用なし） | ○ マネージドDBあり（追加費用の場合あり） | △ 外部DB接続のみ（無料枠は制限あり） | △ D1（SQLite）中心／既存SQLの移行要検討 |
| 独自ドメイン | ○ | ○ | ○ | △ プロキシOFF時のみ | ○ Cloudflare上で完結 |
| スケーリング | ○ 自動 | × 手動増強のみ | △ 簡易的に可 | × リソース制限あり | △ 手動スケール中心（自動化は発展途上） |
| **概算月額コスト目安** | 数百〜数千円（アクセス少なければ低額） | 数百〜1,000円台（最安級VPS利用時） | 無料枠あり／超過で数百円〜 | 無料 | 数百円〜（従量課金、稼働時間次第で変動大） |
| 運用負荷 | 低〜中 | 高（自己管理） | 低 | 最低 | 中（新サービスで情報少なく検証コスト高） |
| Cloudflare連携 | CNAME＋プロキシで前段配置可（無料） | Tunnelで直接接続（無料） | CNAME＋プロキシ（無料） | DNSのみ（プロキシ不可） | ネイティブ統合（同一エコシステム） |
| **Cloudflare担当範囲** | ドメイン・CDN・WAF・SSL | ドメイン・トンネル経路・WAF | ドメイン・CDN・WAF・SSL | ドメインのみ | ドメイン・CDN・WAF・SSL・アプリ実行の全て |
| **ホスティング先担当範囲** | アプリ実行・スケーリング・保守 | アプリ実行・DB・OS保守・監視（自社） | アプリ実行・スケーリング | アプリ実行（一部制限あり） | （Cloudflareに統合、外部ホスティング先は不要） |

**概算コストは目安であり、実際のアクセス数・インスタンスサイズによって変動する。本格検討の際は各サービスの料金計算ツールで試算することを推奨。**

---

## 各案の特徴

### 案0: VPS単体（Cloudflare不使用・最安構成）
最も安価なVPS（月数百円〜）にDocker・Streamlit・SQL（PostgreSQL/MySQL等）・LibreOfficeをすべて同居させ、ドメインもVPS契約会社のネームサーバーで管理する構成。**Cloudflareを一切使わないため管理対象が最小**で、コストのみで見れば最有力候補になり得る。ただしCDN・WAFがなく、SSL証明書も自分でLet's Encrypt等を設定・更新する必要がある。小規模な社内利用や、外部公開範囲が限定的な場合に向く。

- **担当:** VPS一本（ドメイン・SSL・アプリ実行・DB・保守のすべてを自社管理）

### 案1: Cloud Run / Fargate（推奨・本番運用向け）
Dockerfileをそのままデプロイ可能で、LibreOfficeのような重い依存関係も問題なく動作する。オートスケールに対応し、アクセスがない時間は課金が発生しにくいため、コストと運用負荷のバランスが良い。SQLはCloud SQL（マネージドDB）と接続するか、小規模ならコンテナ内にDBを同居させる方法もある。

- **Cloudflareの担当:** 独自ドメイン、CDNキャッシュ、WAF/DDoS対策、SSL証明書（**無料プランで十分**）
- **Cloud Run/Fargateの担当:** Streamlit＋LibreOffice＋DBのコンテナ実行、オートスケール、インフラの保守

### 案2: VPS + Docker + Cloudflare Tunnel（コスト最優先ならこれも有力）
自前のVPS（さくら、ConoHa、AWS EC2等）にDockerでそのままデプロイし、Cloudflare Tunnelで外部公開する方法。**Dockerを動かしているのはあくまでVPS側であり、Cloudflare Tunnel自体はDockerを実行する仕組みではなく、VPSとCloudflare網をつなぐ通信経路を提供するだけ**という点に注意。SQLはVPS内にDocker Composeで同居させれば追加費用なし。月額固定費用で予算が読みやすく、**Cloudflareの無料プランと組み合わせることでCDN・WAFも無料で得られるため、コスト最優先の観点では最有力候補の一つ**。サーバーの保守・監視・障害対応を自分たちで行う必要がある。

- **Cloudflareの担当:** 独自ドメイン、Tunnel経由の通信経路、WAF/DDoS対策（**無料プランで十分**）
- **VPS（自社）の担当:** Streamlit＋LibreOffice＋DBのコンテナ実行、OSアップデート、サーバー監視・障害対応

### 案3: Fly.io / Railway
Dockerfileベースで手軽にデプロイでき、設定もシンプル。小〜中規模の用途であれば導入コストが低い。マネージドDB（PostgreSQL等）も用意されているが、規模によっては追加費用が発生する。大規模運用や日本国内リージョンの充実度はCloud Run等に劣る場合がある。

- **Cloudflareの担当:** 独自ドメイン、CDNキャッシュ、WAF/DDoS対策、SSL証明書（**無料プランで十分**）
- **Fly.io/Railwayの担当:** Streamlit＋LibreOffice＋DBのコンテナ実行、簡易スケーリング

### 案4: Streamlit Community Cloud
無料で最も手軽だが、Dockerfileを直接使う運用には対応しておらず、LibreOfficeのような重い依存関係は不安定になりやすい。SQLも外部DB（無料枠のあるSupabase等）への接続が前提となり、構成がやや複雑になる。また独自ドメインを使う場合はCloudflareのプロキシ機能（オレンジクラウド）をOFFにする必要があり、CDN・WAFの恩恵を受けられない。今回の用途にはやや不向き。

- **Cloudflareの担当:** 独自ドメインのDNS転送のみ（CDN・WAFは利用不可）
- **Streamlit Community Cloudの担当:** アプリ実行（LibreOffice等の重い依存関係には制限あり）

### 案5: Cloudflare Containers（Cloudflareのみで完結させたい場合）
2026年4月にGAした新しいサービスで、DockerfileをそのままCloudflare上で動かせる。ドメイン・CDN・セキュリティ・アプリ実行のすべてをCloudflare単体で完結できる点が最大の魅力。**ただし従量課金のため、常時稼働するStreamlitアプリではコストが読みにくく、コスト最優先の方針には現時点では不向き。** 常駐性（無操作10分でスリープ）や永続ストレージ（ディスクは基本エフェメラル）、SQLの扱い（D1はSQLiteベースで既存DBの移行が必要）にも制限がある。

- **Cloudflareの担当:** ドメイン・CDN・WAF・SSL・アプリ実行（コンテナ）のすべて
- **外部ホスティング先の担当:** なし（Cloudflareに統合されるため不要）

---

## 各案の具体的な手順

### 案0: VPS単体（Cloudflare不使用）

1. 最安クラスのVPS（月額数百円〜、さくらの軽量VPSやConoHa VPS等）を契約
2. VPSにDocker / Docker Composeをインストール
3. `docker-compose.yml`でStreamlitアプリ用コンテナとDB用コンテナ（PostgreSQL/MySQL）を定義し、同一VPS内に同居させる
4. ドメインを取得（お名前.com等）し、VPS契約会社が提供するDNS、またはVPSのIPに直接Aレコードを設定
5. Nginx等のリバースプロキシをVPS内に構築し、80/443番ポートでStreamlitへ振り分け
6. Certbot（Let's Encrypt）でSSL証明書を無料取得し、自動更新のcronを設定
7. ファイアウォール（ufw等）で不要なポートを閉じ、最低限のセキュリティ対策を実施

### 案1: Cloud Run / Fargate

**Cloud Runの場合（GCP）**
1. GCPプロジェクトを作成し、Artifact Registry（コンテナイメージ置き場）を有効化
2. 既存のDockerfileを使い `gcloud builds submit` でイメージをビルド・登録
3. `gcloud run deploy` でCloud Runにデプロイ（メモリ・CPUはLibreOffice起動を考慮し2GB以上を推奨）
4. Cloud Runが発行するURL（`https://xxx.run.app`）で動作確認
5. Cloudflareの管理画面でCNAMEレコードを追加し、独自ドメインをCloud RunのURLに向ける（プロキシON推奨）
6. Cloud Run側でカスタムドメインのマッピングを設定し、SSL証明書を発行

**Fargateの場合（AWS）**
1. ECR（コンテナレジストリ）にDockerイメージをプッシュ
2. ECSクラスタを作成し、Fargate起動タイプでタスク定義（メモリ・CPUを設定）
3. ALB（ロードバランサー）を用意しFargateサービスに接続
4. CloudflareでCNAMEをALBのDNS名に向ける
5. Cloudflare側でSSL/TLS設定（Full設定を推奨）

### 案2: VPS + Docker + Cloudflare Tunnel

1. VPS（さくら、ConoHa、AWS EC2等）を契約しサーバーを用意
2. サーバーにDocker / Docker Composeをインストール
3. 既存のDockerfileを使い `docker build` → `docker run` でコンテナを起動（ポートは内部のみで公開、外部には出さない）
4. `cloudflared` をサーバーにインストールし、Cloudflareアカウントと連携（`cloudflared tunnel login`）
5. `cloudflared tunnel create` でトンネルを作成し、Streamlitのポート（デフォルト8501）にルーティング設定
6. Cloudflare側でDNSレコード（CNAME）をトンネルに紐付け
7. `cloudflared` をサービス化（systemd等）し、サーバー再起動時も自動起動するよう設定

### 案3: Fly.io / Railway

**Fly.ioの場合**
1. `flyctl` CLIをインストールし、Fly.ioアカウントでログイン
2. プロジェクトディレクトリで `fly launch` を実行（既存のDockerfileを自動検出）
3. `fly.toml` でメモリ・リージョン（`nrt`＝東京など）を設定
4. `fly deploy` でデプロイ
5. Fly.io管理画面またはCLI（`fly certs add`）で独自ドメインを追加
6. Cloudflareで発行されたCNAME/Aレコードを設定

**Railwayの場合**
1. GitHubリポジトリをRailwayに接続
2. Railwayが自動的にDockerfileを検出しビルド・デプロイ
3. Railway管理画面の「Settings」から独自ドメインを追加
4. 表示されたCNAMEレコードをCloudflareのDNS設定に追加

### 案4: Streamlit Community Cloud

1. GitHubリポジトリにStreamlitアプリのコードを配置
2. `packages.txt` にLibreOffice関連パッケージ（`libreoffice-calc`等）を記載（※フルのDockerfileは使えないため、動作しない依存関係がある可能性に留意）
3. [share.streamlit.io](https://share.streamlit.io) でリポジトリを連携しデプロイ
4. 発行されたURL（`https://xxx.streamlit.app`）で動作確認
5. 独自ドメインを使う場合はCloudflareでCNAMEを設定するが、**プロキシは必ずOFF（DNSのみ）にする必要がある**

### 案5: Cloudflare Containers（Cloudflareのみで完結させる場合）

1. Cloudflareアカウントで有料プラン（Workers Paid）を有効化（Containersは有料プラン限定機能）
2. ローカルにDocker（またはColima等）をインストールし、`wrangler` CLIをセットアップ
3. `wrangler containers build` で既存のDockerfileからイメージをビルド
4. `wrangler.toml`（または`wrangler.jsonc`）でコンテナのインスタンスタイプ（vCPU・メモリ）、`sleepAfter`（スリープまでの時間）を設定
5. コンテナを呼び出すWorkerコード（TypeScript）を作成し、リクエストをコンテナにルーティングするよう実装
6. `wrangler deploy` でWorker＋コンテナを一括デプロイ
7. セッションデータやアップロードファイルなど永続化したいものは、コンテナ内ディスクではなく **R2**（オブジェクトストレージ）や**D1**（マネージドSQLite）に保存するようアプリ側を改修
8. Cloudflareダッシュボードでカスタムドメインを設定（同一エコシステムのため追加のDNS連携作業は不要）
9. 本番投入前に、LibreOfficeによる変換処理を含めた負荷テスト・スリープ挙動の検証（PoC）を実施

---

## まとめ・提案（運用コスト最優先の観点）

- **コストを最優先するなら「案2: VPS + Docker + Cloudflare Tunnel」が最有力。** VPSの固定費（月数百〜1,000円台）のみで、Cloudflareの無料プランを組み合わせればCDN・WAF・SSLも追加費用なしで得られる。DBもVPS内に同居させれば追加コストが発生しない。
- **さらに極限までコストを削るなら「案0: VPS単体（Cloudflare不使用）」も選択肢。** ただしCDN・WAFがない分、セキュリティ面は自己責任になる点は上司に説明しておきたい。
- **本番運用・将来のアクセス増加を見据えるなら「案1: Cloud Run / Fargate」。** アクセスが少ない間は従量課金で安く済み、増えた場合は自動スケールで対応できるため、コストと将来性のバランスが良い。
- 「案3: Fly.io / Railway」「案4: Streamlit Community Cloud」は導入の手軽さはあるが、今回の要件（Docker＋LibreOffice＋SQL）とはやや相性が悪く、優先度は下がる。
- 「案5: Cloudflare Containers」は目新しさはあるが、従量課金の性質上コストが読みにくく、**コスト最優先の方針とは現時点では相性が良くない**ため、優先度は最も低い。

**総合結論:** 運用コストを最優先するなら、**Cloudflareは「無料の窓口」として使い、アプリ本体は安価なVPSかCloud Runで動かす（案2または案1）**構成を提案したい。