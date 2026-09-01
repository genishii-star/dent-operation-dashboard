# Dent Inc. 運営ダッシュボード 仕様書

最終更新: 2026-08-24

---

## 1. 概要

Dent Inc.（民泊運営代行）の**業務可視化・意思決定支援ダッシュボード**。
自社運営データ（Airhost PMS）と外部市場データ（AirDNA・e-Stat）を統合し、KPI・市場比較・将来予測を一画面で確認できる。

- **利用者**: 代表・運営担当
- **デプロイ形態**: Cloudflare Pages（静的サイト）＋ D1 バックエンド（`dent-data-api` Worker 経由）
- **アクセス制限**: Cloudflare Access
- **対象エリア**: 大阪 / 京都 / 東京 ほか主要都市

---

## 2. システム構成

```
┌─────────────────────────────────────────────────────────┐
│ フロントエンド (index.html + app.js)                     │
│  - Cloudflare Pages (op.dent-inc.com) / Access 保護       │
│  - Chart.js + 素のJS (SPA) / LocalStorageキャッシュ30分   │
└──────────┬──────────────────────────────┬───────────────┘
           │ /internal/*.json (読取)      │ GAS (書込・通知)
           ▼                              ▼
┌─────────────────────┐         ┌─────────────────────┐
│ D1  dent-platform-d1│         │ Google Apps Script  │
│  ・reservations     │◀────────│  ・朝の運営サマリー │
│  ・daily_revenue    │  読取   │  ・新法チェック     │
│  (dent-data-api     │         │  ・異常検知         │
│   Worker が front)  │         │  ・物件マスタ書戻し │
└──────────▲──────────┘         └─────────┬───────────┘
           │ POST /cs/ingest/*            │ Slack Webhook
           │                              ▼
┌──────────┴──────────┐         ┌─────────────────────┐
│ Buddy (外注AIワーカー)│        │ Slack通知           │
│  Mac mini常駐        │         └─────────────────────┘
│  ・毎時: 差分        │
│  ・朝8時: 全量       │      ┌─────────────────────┐
└─────────────────────┘      │ Google Sheets       │
                              │  ・シーズンマスタ   │◀ 人が編集
┌─────────────────────┐      │  ・AirDNA / e-Stat  │◀ 拡張/GAS
│ Chrome拡張 (併走中)  │─────▶│  ・清掃依頼管理(LIB)│
│  ・Airhost取込       │      └─────────────────────┘
│  ・AirDNA取込        │
│  ・LIB清掃同期       │  ※Airhost取込は Buddy へ移管中。
└─────────────────────┘    指示書§9 段階4で停止予定
```

> **⚠ 予約・日次データの正本は D1。スプレッドシートではない。**
> Buddy が D1 へ直接書くようになった 2026-08 以降、`予約データ` / `日次データ` シートは
> **Chrome拡張のボタンを押したときしか更新されない**。押さなくなれば静かに凍る。
> 集計・通知を作るときは必ず D1 (`/internal/reservations.json`, `/internal/daily.json`) を読むこと。
> 実際 2026-08-23 に、朝の運営サマリーだけが Sheets を読んでいたため
> 前日の新規予約を 32件 → 13件 と過少に投稿する事故が起きた。

### 主要コンポーネント

| コンポーネント | 実装 | 役割 |
|---|---|---|
| フロントエンド | `index.html` / `app.js` (約8,500行) | 10タブの画面描画・集計・フィルタ |
| Chrome拡張 | `chrome-extension/` (Manifest v3) | AirDNA取込・LIB清掃同期。**Airhost取込は Buddy へ移管中（併走）** |
| Buddy（外注） | Mac mini常駐 / mybrain.tv 三口氏 | Airhostから予約・日次を取得し D1 へ直接投入。仕様は `外注指示書_Buddy_Airhostデータ自動取得.md` |
| GASバックエンド | `gas/` + `update_masters.gs` + `setup_spreadsheet.gs` | スプシ書込API・異常検知・外部統計同期 |
| レビュー自動化 | `review/` (Playwright) | Airbnbレビュー取得／返信投稿 |

---

## 3. ダッシュボード画面（10タブ）

| # | タブ | 主な表示内容 |
|---|---|---|
| 1 | **TOP（日次分析）** | 売上合計 / 受取金 / 新規予約 / ADR / OCC / 平均宿泊日数 / PM・BM売上 など11 KPIカード、月次売上・OCC・ADR推移、OTAチャネル構成 |
| 2 | **オーナー別分析** | オーナーごとの月額目標 vs 実績進捗バー、達成率KPI、ロイヤリティ・OCC・ADR一覧 |
| 3 | **物件別分析** | 物件スコアカード（エリア/期間/種別/間取り/㎡でフィルタ）、シリーズ別集計、物件詳細モーダル（**9項目インライン編集 → GAS経由でマスタ書戻し**） |
| 4 | **予約獲得状況** | 予約件数 / GMV / ADR / 平均泊数 / 平均人数 / リードタイム、物件別・チェックイン月別、予約一覧表（16列） |
| 5 | **売上・稼働** | OCC/ADR/RevPAR/売上/受取金/リードタイム6 KPI、**ペースレポート**（平日/休日別30/60/90日先先行予約・価格推奨）、**4象限スコアカード**（達成率×先行予約）、未来予約分析（90日OCC / 180日ADR予測）、**市場ランク分布**（AirDNAパーセンタイルP25/P50/P75/P90に対する自社物件の位置づけ） |
| 6 | **レビュー** ※モック | 平均★・件数・記録率・ネガ率・自動投稿数、月次トレンド、6軸レーダー、ポジ/ネガキーワード、低評価返信承認キュー、物件スコアカード、レビュー一覧、自動投稿ログ（※現状ダミーデータ。自動化本体は`review/`で稼働中） |
| 7 | **要チェック** | 🆕 新規物件（稼働4ヶ月未満）、⚠️ アラート物件（達成率<70%赤<50% / 直近30日OCC<60% / 直近14日新規予約0） |
| 8 | **新法チェック** | 民泊新法365日上限管理。目標件数 / 超過 / ほぼ満杯(95-100%) / 消化率 / 月次予約・宿泊、物件別 残日数・進捗バー・月次内訳 |
| 9 | **PM/BM分析** | PM（ロイヤリティ）とBM（清掃・サポート）を分離。月次推移（過去3/当月/先3ヶ月）、PM比率、エリア構成、オーナーTOP10（合算/PM/BM）、物件TOP10（PM）、前年比成長・減少 |
| 10 | **マーケット** | 3サブタブ構成。<br>・**TOP**: 需給マトリクス・国別ゲスト構成・CPI vs ADR・予約ペース・24ヶ月季節性ヒートマップ・3データ横断インサイト<br>・**AirDNA**: 大阪/京都/東京の市場OCC/ADR/RevPAR12ヶ月推移、区別ランキング、寝室数別<br>・**e-Stat**: 訪日外客数・国別TOP10・地域構成（東・東南アジア/欧米）・3都市県別稼働・CPI宿泊 |

共通機能:
- 期間フィルタ: 今月 / 先月 / 直近3ヶ月 / 翌月 / 翌2ヶ月 / 前年 / カスタム
- エリアフィルタ: 全体 / 大阪 / 京都 / 東京 / その他
- フィードバックボタン（全タブ右上） → GAS経由でSlackチャンネル `#sync-monitoring` へ送信

---

## 4. データソース

### マスタCSV（スプシと同期）
| ファイル | 内容 |
|---|---|
| `オーナーマスタ.csv` | オーナーID・名称・エリア・月額目標売上 |
| `物件マスタ.csv` | 物件コード・オーナーID・エリア・住所・部屋数・KPI除外・稼働ステータス・閑散/通常/繁忙期目標 |
| `シーズンマスタ.csv` | 月 → シーズン区分（繁忙/通常/閑散） |

> **⚠ facilities YAML `rooms` の書式契約（wiki ↔ operation）**
> 物件マスタの実体は wiki リポジトリの `site/data/facilities/*.yaml`（source of truth）。
> 多部屋物件の `rooms` は **文字列リスト**（例: `- 7104F`）または **`{id, key_box}` オブジェクトリスト**（例: `- id: DIS207`）のどちらか。
> ダッシュボードは `extractRoomCode()` で両形式を吸収し、部屋コードを取得できない物件は `collectMasterSchemaWarnings()` がバナー警告する。
> wiki 側が rooms に新フィールドを足す場合は **必ず `id` を残す**こと（無いと物件コードが空になり脱落する）。
| `予約データ.csv` | 予約36列（予約サイト、Airhost ID、チェックイン/アウト、ゲスト情報、売上、受取金、ステータス、支払状況…） |
| `日次データ.csv` | 物件×日次の売上・清掃費・調整・ペナルティ明細 |

### 予約・日次データ（正本 = D1）
| データ | 実体 | 投入経路 |
|---|---|---|
| 予約 | D1 `reservations` (PK: `airhost_reservation_id`) | Buddy → `POST /cs/ingest/reservations` |
| 日次 | D1 `daily_revenue` (PK: `property_name, room_no, date, reservation_id`) | Buddy → `POST /cs/ingest/daily-revenue` |

読み取りは `dent-data-api`（Cloudflare Access 保護）:

| エンドポイント | 用途 | 絞り込み |
|---|---|---|
| `GET /internal/reservations.json` | 予約 | `since`/`until`（チェックイン日）、`booked_since`/`booked_until`（予約日） |
| `GET /internal/daily.json` | 日次 | `since`/`until`（日付） |
| `GET /internal/data-freshness.json` | 取り込み鮮度 | —（`max(ingested_at)` 等を返す。ヘッダの「最終同期」表示に使用） |
| `GET /internal/facilities.json` | 物件マスタ（YAML） | — |

- **レスポンスは日本語見出しのオブジェクト配列**（`販売` / `合計日数` / `ゲスト数` / `状態` / `部屋番号` / `AirHost予約ID` …）。
  スプシ時代の見出しを踏襲しているので、Sheets読みのコードはURL差し替えだけで移行できる。
  ただし**シートと1文字違いの見出しがいくつかある**ので、必ず上のマップに合わせること
  （`alert-anomaly.gs` は `売上合計`/`泊数`/`Airhost予約ID` のまま書かれており、
  予約IDが空になって全行スキップされ、4ヶ月間1件も検知していなかった）。
- **ゲスト氏名は `reservations.guest_name` に入る**（2026-08-31 方針転換。新法の宿泊者名簿用。
  migration 0026 に守るべき約束事5点あり）。ただし **`/internal/reservations.json` は返さず**、
  オーナー本人が `/owner/guests` で自分の予約分だけを都度取る。**13ヶ月で自動NULL化**。
  Slack通知・明細ファイルには載せないこと。
  2026-09-02 に拡張の「👤 宿泊者名バックフィル」で 2025-08 以降 16,113件を投入済み。
  ⚠ **Buddy はまだ氏名を送っていない**（指示書 §6.2.1 未対応）。そのため本流の取り込みが
  `INSERT OR REPLACE` で **guest_name を NULL で上書きする**。Buddy の窓（過去90日＋未来365日）に
  入る予約の氏名は毎朝8時の全量で消える。窓の外（今日なら 2026-06-03 以前）は消えない。
  Buddy が対応すれば窓の中は翌朝埋まり直すが、対応が遅れるほど「消えたまま埋まらない」
  範囲が1日ずつ広がる。長引くなら `ingest.ts` に空で上書きしないガードを入れること
  （`ingestGuestNames` には既にある）。
- **電話・メール・住所・コメントは D1 に入っていない**（ingest 時に PII として破棄）。通知本文に使えない。
- 物件コード解決は `property_code` 列。マスタに無ければ NULL（orphan）として残す。

### 外部データ（スプシ中継）
| データ | スプシID | 同期方法 |
|---|---|---|
| AirDNA市場データ | 都市別3スプシ（大阪/京都/東京） | Chrome拡張がAirDNA API直叩き |
| e-Stat 訪日・宿泊・CPI | `1d0dfPK…` | GAS (`gas/estat-sync.gs` / `gas/mlit-sync.gs`) 月次 |
| シーズンマスタ | メインスプシ `1C7EiYSz…` | 人が編集。ダッシュボードが唯一まだ読む Sheets |
| 清掃依頼管理(LIB) | `1GqZFC…` | Chrome拡張のボタン → GAS `syncLibCleaning`（**手動**） |
| ~~Airhost 予約・日次~~ | メインスプシ `1C7EiYSz…` | **Buddy → D1 へ移管済み。拡張を押したときだけ更新される凍結気味の経路** |

---

## 5. 自動化・連携

### Chrome拡張（`chrome-extension/`）
- **Airhost取込**: 予約CSV/日次CSVを1クリックでスプシ反映（`doGet?action=morningReport` で Sheets→D1 も走る）。**Buddy へ移管中で併走。停止予定**
- **AirDNA取込**: app.airdna.co閲覧中の市場指標をOCC/ADR/RevPAR/パーセンタイル別に抽出
- **LIB清掃同期**: LIB物件の予約を清掃依頼管理シートへ転記

### Google Apps Script
| GAS | 役割 |
|---|---|
| 物件マスタ書戻し | ダッシュボードの物件詳細編集 → `updatePropertyMaster` dispatcher経由で書込（トークン認証） |
| フィードバックプロキシ | ダッシュボード右上ボタン → Slack Webhookへ転送 |
| 異常検知 (`gas/alert-anomaly.gs`) | 30分トリガ。投稿先は `#alert-異常予約検知`（`C0BT143TY82`。2026-08-27 に `#management` から移設。プロパティ `SLACK_WEBHOOK_ANOMALY`）。**D1読み**（予約日の直近3日を判定、ADR基準はチェックイン直近90日）。低ADR（**その部屋の過去90日最安を下回る** かつ 平均×0.7未満）・長期連泊（14泊以上）・定員超過（棟/フロア貸切 `ALL`・`2F-ALL` は判定対象外）。定員は facilities.json の `room_types[].details.max_guests` / `amenities.max_guests` から取得 |
| 朝の運営サマリー (`gas/morning-report.gs`) | Buddy の全量ジョブ完了フック(`doGet?action=reportOnly`)で 新法通知 → 前日KPI/当月見込みを `#management` に投稿。**D1読み** |
| 新法通知 (`gas/shinpou-report.gs`) | 民泊新法物件の年度営業日数を facilities.json + D1予約から集計しSlack投稿。ダッシュボード「新法チェック」と同一計算 |
| e-Stat / MLIT同期 | 政府統計APIから訪日外客数・宿泊統計・CPIをスプシに定期取得 |

### レビュー自動化（`review/`）
- Playwright製のAirbnb操作スクリプト（取得/投稿/返信の3系統）
- 返信連投で取り違え事故を起こさないため、**1件ごとにリロード＋投稿前検証**を必須化済み
- 検証運用中（OPH等で実績）、ダッシュボードタブ側はUI先行でモック表示

---

## 6. Airhost データ自動取得（Buddy 連携）

2026-08 に、Airhost からの取得を **石井が毎朝 Chrome拡張のボタンを押す** 運用から
**Buddy（外注AIワーカー・Mac mini常駐）が D1 へ直接投入する** 運用へ移管中。
先方向けの正本は `外注指示書_Buddy_Airhostデータ自動取得.md`。

### ジョブ構成

| ジョブ | 頻度 | 取得範囲 | 対象 |
|---|---|---|---|
| **A. 差分** | 毎時1回 | 過去3日 + 未来120日（チェックイン日で切る） | 予約のみ |
| **B. 全量** | 毎朝8:00 | 過去90日 + 未来12ヶ月 | 予約 + 日次 |

- Airhost のエクスポートAPIは `date_type` に **`checkin_date` / `checkout_date` しか受け付けない**
  （`updated_at` / `booking_date` は 400）。よって予約日での差分取得はできず、チェックイン日で切っている。
- `date_range` は UI では3ヶ月上限だが **API は12ヶ月まで通る**（455日は `date_range_too_large`）。
  これで朝の全量が6エクスポート → 3エクスポートに減った。
- エクスポート間隔は **8秒だと429。40秒で安定**。
- 毎時を「未来361日」に広げる案（B案）は見送り。120日より先のチェックインで入る新規予約は
  実測で月9件しかなく、D1 の書き込みが 762行/回 → 約4,000行/回 に増えるだけで見合わない。

### 契約事項

- **冪等**: どちらのジョブも `INSERT OR REPLACE`。A と B がどの順で届いても、何度届いても結果は同じ。
  **失敗した回はリトライせず次の回に任せる**（429/5xx でバックオフして諦める）。
- **PII は送信前に落とす**（電話・メール・住所・コメント）。Dent 側でも破棄するが二重で防ぐ。
  **氏名だけは例外で送ってもらう**（2026-08-31 変更。指示書 §6.2.1）。
- **1リクエスト最大1000行**（推奨200行）。超過は 413。
- **完了フック**: 朝の全量が終わったら **1回だけ** `GET <GAS>/exec?action=reportOnly&token=…`。
  これで新法通知と朝の運営サマリーが最新データで発火する。
  **毎時ジョブでは叩かない**（GAS側に冪等ガードはあるが実行ログが汚れる）。
- **`action=morningReport` と `action=reportOnly` を混同しないこと**。前者は Sheets→D1 の
  取り込み (`d1SyncDaily`) を伴う拡張用で、Buddy が叩くと**古い Sheets で D1 を上書きして
  データが巻き戻る**。移行期に一番踏みやすい地雷。
  （`d1SyncDaily` 側にも「設定」B1 の最終同期時刻が古ければ `stale-sheet` で見送るガードあり）
- **通知**: 失敗は必ず Slack。毎時の成功は通知しない（1日24件で埋もれる）。連続失敗は
  「N回連続で失敗中」に集約。既存の Bot 経路に乗せる。

### 移行段階（指示書 §9）

| 段階 | 内容 | 拡張のポチ | 状況 |
|---|---|---|---|
| 0〜0.5 | 疎通確認・`date_type`/`date_range` 検証 | 継続 | 完了 |
| 1 | 全量Bを手動で1回 → Dent側で突合 | 継続 | 完了（2026-08-21） |
| 2 | 全量Bを毎朝8時で自動化 | 継続 | 完了 |
| 3 | 毎時Aを追加 | 継続 | 完了 |
| 4 | **1週間ズレが出なければ拡張を停止** | 停止 | 判定予定 2026-08-28 |
| 5 | 異常検知のD1移行が済んだら Sheets 経路を撤去 | — | 異常検知は 2026-08-24 に移行済み |

### 物件コード解決と orphan

`property_code` は facilities マスタ（`wiki/site/data/facilities/*.yaml`）から解決する。
解決できない行は **NULL のまま残す**（消さない）。`WHERE property_code IS NULL` で抽出できる。

- `rooms` は「id の文字列配列」と「`{id, entry, sleeping}` のオブジェクト配列」の2形式がある。
  **`String(room)` すると後者が `"[object Object]"` に潰れる。** 必ず `id` を見ること。
  この取り違えが 2026-08 時点で3箇所（Worker のリゾルバ / `backfill-d1/backfill.mjs` / GAS の
  `expandFacilityToMaster`）に存在し、物件コードNULL・OCC過少表示という別々の症状で出ていた。
- 部屋番号が無い/`ALL`/`2F-ALL` の棟単位予約は、**部屋番号が空か ALL 系のときだけ**施設コードに
  フォールバックする。部屋番号があるのに外れた場合は NULL のまま（打ち間違いを検知するため）。
- 旧リスティング名は `NAME_MERGE`（`HGK(旧)`→`HGK`, `NNJ(旧)`→`NNJ`）で寄せる。
  **Worker / backfill / `app.js` の3箇所に同じ表がある**ので、足すときは3つとも直す。

### 動作確認の勘所

```sql
-- 取り込みが生きているか（バケットはUTC。upsertで上書きされるので最新回しか残らない）
SELECT substr(ingested_at,1,13) b, count(*) FROM reservations GROUP BY 1 ORDER BY 1 DESC LIMIT 5;

-- 実物件の orphan（TEST 物件以外が出たらマスタ側の登録漏れ）
SELECT property_name, room_no, count(*) FROM daily_revenue
WHERE property_code IS NULL GROUP BY 1,2 ORDER BY 3 DESC;
```

ダッシュボード右上の「最終同期」は `/internal/data-freshness.json`（D1 の `max(ingested_at)`）。
6時間以上取り込みが無ければ色が変わる。

---

## 7. セキュリティ・運用

- **認証**: トップページパスワード（ログインオーバーレイ）
- **APIキー**: Google Sheets APIキー（読取専用）をフロントに埋込、書込は必ずGAS経由＋トークン
- **キャッシュ**: ブラウザLocalStorageに30分保持、タブ切替・手動更新で再取得
- **通知**: Slack Webhook経由で同期完了・異常検知・フィードバックを `#sync-monitoring` チャンネルに配信

---

## 8. ファイル構成

```
operation/
├─ index.html            … 10タブ定義・モーダル
├─ app.js                … 画面ロジック全般（約8,500行）
├─ style.css
├─ setup_spreadsheet.gs  … 初回スプシ構築
├─ update_masters.gs     … マスタ更新ユーティリティ
├─ gas-feedback-proxy.js … フィードバック→Slack中継GAS
├─ gas/                 … ⚠ .gitignore 対象。ローカルとGASプロジェクトにしか存在しない
│   ├─ alert-anomaly.gs  … 異常検知 + 直前予約通知（+ getProp/parseNum/toDateStr/leadDays を提供）
│   ├─ morning-report.gs … 朝の運営サマリー（+ fetchDataApiJson / fetchReservationsApi / fetchDailyApi）
│   ├─ shinpou-report.gs … 民泊新法 営業日数チェック
│   ├─ d1-sync.gs        … Sheets→D1 取り込み（⚠ 2026-08-24 以降は非常用。定常はBuddy→D1直投入）
│   ├─ estat-sync.gs     … e-Stat同期
│   ├─ mlit-sync.gs      … 国土交通省データ同期
│   └─ meeting_notes.gs  … 議事録
│   ※上記は**単一のGASプロジェクト**に同居。ファイル間で関数を共有しているため分割不可
├─ chrome-extension/     … ⚠ 版管理外 (2026-08-24 に追跡を外した)。Manifest v3拡張（Airhost/AirDNA/LIB）
├─ review/               … Playwrightレビュー自動化
└─ *.csv                 … マスタデータ
```

---

## 9. 今後の拡張予定・既知の制約

- **レビュータブ本番化**: UI完成・自動化本体も稼働済。モックデータを実データ接続に置換する作業が残件。
- **レビュー自動投稿の段階展開**: DOM事故防止のためリロード+検証フローで安全運用中。対象物件を段階的に拡大予定。
- **マーケットTOPのインサイト**: ルールベースで生成中。将来的にLLM要約への差替え余地あり。
- **`gas/` と `chrome-extension/` が版管理外**: 朝レポート・異常検知のロジックが履歴に残らない。
  除外理由は 2026-08-24 に判明した — **このリポジトリは public** (`genishii-star/dent-operation-dashboard`)
  で、`morning-report.gs` の `TRIGGER_TOKEN` を公開できないため。ただし `chrome-extension/` は
  ignore 追加より前から追跡されていて ignore が効かず、**同じトークンが平文で公開されていた**
  (2026-08-24 に追跡を外し、トークンをローテーション)。版管理したいなら private リポジトリへ
  分離するのが本筋。
- **LIB清掃同期が手動依存**: 拡張ポップアップのボタン → GAS `syncLibCleaning` を人が押す方式。
  `syncLibCleaning` のソースがリポジトリに無く、転記元が Sheets か D1 かも未確認。
  拡張を止める段階4の前に確認が必要。
- **低ADR判定のしきい値**: 「過去90日最安を下回る」に変更して 0.4件/日 まで絞ったが、実運用での
  的中率は未評価。数週間ぶんの通知を見てから再調整する。
- **ADR / 予約Window の定義差**: 朝レポートとダッシュボードで除外ステータスの扱いが違い、
  同じ日でも数値が僅かにズレる（例 2026-08-22: 15,266/22日 vs 14,930/24日）。以前からの仕様差。
- **毎時ジョブの窓**: 過去3日+未来120日。120日より先のキャンセル・新規は翌朝の全量まで反映されない。
