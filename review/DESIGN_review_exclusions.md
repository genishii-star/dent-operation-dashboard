# 設計: レビュー依頼の抑制 (review_exclusions)

| | |
|---|---|
| 状態 | **設計のみ / 未実装** |
| 作成 | 2026-07-16 |
| 関連 | `generate-drafts.mjs`, `post-approved.mjs`, `airbnb-helpers.mjs`, D1 `review_drafts` |

## 1. 解きたい問題

レビューパイプラインを完全自動化 (= **オーナー承認の撤廃**) したいが、**レビュー依頼を送りたくないゲストがいる**。この判断はDent側にしか無く、Airhostからフラグとして取得できない。よって自社でDBを持つ必要がある。

### 前提 (要確認)

本設計は「**レビュー依頼 = `review_of_guest` の投稿**」と解釈している。ホストがゲスト評価を投稿するとAirbnbがゲストに通知を出し、リスティングへのレビューを促すため。トラブルのあったゲストを突けば、悪いレビューを自ら呼び込むことになる。

⚠️ Airbnbのレビュー公開仕様 (双方投稿 or 14日で公開、ホスト投稿がゲストへの通知トリガーになる) は**実挙動を確認していない**。§7の遅延投稿案はこの仕様に依存するため、採用前に検証が要る。

## 2. 前提条件: 確定結合への修正 (抑制リストより先)

**現状の結合は信用できない。** `generate-drafts.mjs:172` の `matchReservation()` はゲスト名のあいまい一致で確認コードを引いている:

```js
n === g || n.includes(g) || g.includes(n.split(/\s+/)[0])
```

最後の条件は「レビュー側の名前がテーブル側の姓名の第1トークンを含む」という緩さで、同名ゲストがいれば**先勝ちで別人の行を掴む**。これを鍵に抑制判定すると、除外したゲストに投稿し、無関係のゲストを除外するという最悪の取り違えになる。

### 名前を介さず結合できる (2026-07-16 プローブで確認)

`/hosting/reservations/completed` の**同一行**に両方の識別子が入っている:

```
{"status":"Review guest",                "name":"更紗 辻井", "code":"HMH4YNJ9MN", "reservation_id":"1728066477997763781"}
{"status":"Review guest - Expires soon", "name":"石田 帆未", "code":"HMQYXEYQEH", "reservation_id":"1722969035971353057"}
```

行内の `a[href*="/hosting/reviews/{id}/edit"]` が `scrapePendingGuestReviews()` の返す `reservation_id` と一致する。

**修正内容:**
- `scrapeReservationsIndex()` の返り値に `reservation_id` を追加
- `matchReservation()` を名前一致 → `reservation_id` の完全一致に置換

**副産物:** Airbnbの `Status` 列が `Review guest - Expires soon` を自前で出している。現在の `check_out + 14日` 計算 (owner portalの締切バッジ) より信頼できる一次情報なので、締切判定の根拠を差し替えられる。

## 3. データモデル

```sql
CREATE TABLE review_exclusions (
  confirmation_code TEXT PRIMARY KEY,   -- HM... = D1 reservations.チャンネル予約ID
  scope        TEXT NOT NULL DEFAULT 'review_of_guest',  -- 'review_of_guest' | 'all'
  reason       TEXT,
  created_by   TEXT NOT NULL,           -- CSビューのユーザー識別子 / Access email
  source       TEXT NOT NULL,           -- 'cs-view' | 'dashboard' | 'manual'
  created_at   TEXT NOT NULL,
  released_at  TEXT,                    -- 解除は物理削除しない
  released_by  TEXT
);
CREATE INDEX idx_review_exclusions_active ON review_exclusions(confirmation_code) WHERE released_at IS NULL;
```

### なぜ `reservations` にフラグ列を足さないか

**`d1SyncDaily` が毎朝Airhostから予約を洗い替えるため、列を足すと同期で消える。** 別テーブルなら同期と独立に生存する。予約データは「Airhostのミラー」、抑制フラグは「Dent固有の判断」で、ライフサイクルが違う。混ぜない。

### なぜ鍵を `confirmation_code` にするか

| 候補 | 可否 |
|---|---|
| **`confirmation_code` (HM...)** | ✅ Airhost `チャンネル予約ID` と一致 / CSが画面で見られる / 他OTAにも同概念があり将来拡張できる |
| Airbnb `reservation_id` (19桁) | ❌ Airbnb内部IDで、CSが見る画面のどこにも出ない。フラグを立てる人が指定できない |
| ゲスト名 + チェックアウト日 | ❌ 同名衝突。§2で潰した問題を鍵の側に持ち込む |

## 4. 判定ポイント (二重)

| # | 場所 | 目的 |
|---|---|---|
| 1 | `generate-drafts.mjs` — 生成前 | ドラフト自体を作らない。トークンも節約。誤投稿される文面が存在しない |
| 2 | `post-approved.mjs` — 投稿直前 | **生成後・投稿前にクレームが入る**ケースを捕まえる |

2つ必要な理由は状態が途中で変わるため。承認を撤廃する以上、**gate 2 が最後の砦**になる。

### fail-closed

`draft_type = 'review_of_guest'` かつ `confirmation_code` が null なら**投稿しない** (保留 + 通知)。

理由は損失の非対称性。レビュー1件の取り逃しは軽微だが、悪いレビューは公開されると恒久的に残る。身元を特定できない予約について「抑制リストに載っていないこと」を証明できない以上、投稿してはいけない。§2の確定結合を入れれば null はほぼ出ない。

### scope を分ける理由

**返信 (`reply`) は抑制対象にしない。** 既に公開されたレビューへの応答であり、ゲストを新たに突く効果がない。抑制したいのは「まだレビューを書いていないゲストを起こしてしまう」`review_of_guest` だけ。全接触を断ちたい例外用に `scope='all'` を残す。

### 観測可能性 (silent suppression を作らない)

抑制は**黙って効かせない**。実行ログとSlack通知に「抑制 N件」を出す。全件抑制されているのに「順調に回っている」と誤認する事故を防ぐ。

## 5. フラグの入力経路: CSビュー連携

CSが日常的に使うwebビューからフラグを立てる。連携可能なことは確認済み (2026-07-16)。**Dent側は受け口の契約だけ決め、上流の実装には踏み込まない。**

### 契約

```
POST /internal/review-exclusions        (Cloudflare Access / service token)
  { confirmation_code, scope?, reason?, created_by }
  → 201 { ok: true }
  → 400 { error: "unknown confirmation_code" }   ← 重要 (下記)

DELETE /internal/review-exclusions/:code  → 解除 (released_at を打つ)
GET    /internal/review-exclusions        → 一覧 (ダッシュボード表示用)
```

### 書き込み時に D1 `reservations` と照合して弾く (必須)

**`confirmation_code` が既存予約に存在しなければ 400 で落とす。** 成功扱いにしてはいけない。

理由: 抑制リストは承認撤廃後の**唯一のゲート**であり、キーが1文字ずれれば黙って無効になる。「フラグを立てたつもりで実は誰も守られていない」は、フラグを立て忘れるより悪い — CSは対処済みだと信じてしまう。

これは机上の懸念ではない。既存のCSツール `wiki/site/docs/tools/cs-form.html` では **`予約番号` が手入力のフリーテキスト** (`cs-form.html:195`) で、予約マスタとの照合が一切ない。同じ設計を持ち込むと確実に踏む。上流が選択式なら照合は保険、手入力なら生命線になる。

### 立て忘れは設計で消えない

CSがフラグを立て忘れれば漏れる。ここは仕組みで完全には潰せない。緩和策は §7 の3状態モデル (既定を `hold` にすれば忘れても被害が減る) だが、仕様検証が前提。

## 6. 承認撤廃で何が変わるか

オーナー承認を外すと、抑制リストが解くのは「**誰に**送らないか」だけで、「**何を**公開するか」は誰も見なくなる。この2つは別のリスクで、別の draft_type に効く。

| draft_type | 掲載先 | 撤廃で増えるリスク |
|---|---|---|
| `review_of_guest` | **ゲストのプロフィール** | 文面リスクは小 (定型の称賛)。真のリスクは nudge → 抑制リストが担当 |
| `reply` | **自社リスティング** (将来のゲストが読む) | **文面リスクが大**。低評価レビューへの的外れな返信が公開され、恒久的に残る |

**抑制リストは `reply` の文面リスクを一切カバーしない。** そして承認キューの元々の名前は「★低評価返信 — 承認待ちキュー」で、まさに低評価への返信を人が見るための仕組みだった。それを外すなら、代わりの防御が要る。

### 提案: 全撤廃ではなくリスク階層で切る

| 対象 | 扱い |
|---|---|
| `review_of_guest` | 全自動 (抑制リスト + fail-closed でゲート) |
| `reply` (高評価) | 全自動。定型の礼状で事故りにくい |
| `reply` (低評価) | **人を残す**。件数が少なく、失敗コストが最も高い |

**ただし現状これは実装できない。** `scrapeUnrepliedReviews()` が返すのは `{review_id, guest_name, date, property_name, original_text}` だけで、**★評価を捨てている** (`airbnb-helpers.mjs:155-161`)。カードのDOMには `Overall quality` / `Rating N` があり、本文開始位置の判定に数字行を使っているのに拾っていない。階層化するなら **rating の取得が前提**。

## 7. 検討したが未採用の案

### 遅延投稿 (3状態モデル)

抑制を binary にせず 3状態にする:

| 状態 | 挙動 |
|---|---|
| `normal` | 早期に投稿 (ゲストを突いてレビューを促す = 良いゲストには積極的にやりたい) |
| `hold` | 14日期限のギリギリ (12〜13日目) に投稿。ゲストが反応する時間を最小化しつつ、ゲスト評価自体は残す |
| `never` | 投稿しない |

「良いゲストは早く、危ないゲストは期限ギリギリに」というホストの定石に沿う。**既定を `hold` にすればフラグ立て忘れの被害が減る**ため、§5の弱点への保険にもなる。

**未採用の理由:** §1のAirbnb公開仕様の検証待ち。仕様が想定と違えば無意味。また現状は逆に**期限超過で3件失っている** (NAGAI: Safitri/Matthew/Tricia) ので、遅延を既定にするのは事故を増やしかねない。まず確定結合と `Expires soon` 準拠の締切判定を入れ、期限管理が安定してから再検討する。

## 8. 未決事項

1. **「レビュー依頼」の定義** — `review_of_guest` の投稿でよいか (§1)
2. **Airbnbのレビュー公開/通知仕様の検証** — §7の前提
3. **`reply` の文面を誰が見るか** — 承認撤廃で無防備になる。リスク階層で切るなら rating 取得が前提 (§6)
4. **CSビュー側の `予約番号` が選択式か手入力か** — 手入力なら §5 の照合が生命線
5. **過去分のシード** — 既存の問題ゲストを遡って登録するか。ソースが無いので手入力になる
6. **ゲストが既にレビュー済みかを判定できるか** — できれば「突くリスクが無い」ケースを自動除外でき、抑制対象が減る。`review_id` と `reservation_id` が同じ番号空間かは未調査

## 9. 段階導入 (案)

| 段階 | 内容 | 依存 |
|---|---|---|
| 0 | 確定結合への修正 + 締切を `Expires soon` 準拠に | なし。**単独で価値があり、抑制リストと無関係に今すぐ入れるべき** |
| 1 | `review_exclusions` テーブル + Worker API (照合付き) + gate 2 | 段階0 |
| 2 | CSビュー連携 + ダッシュボードに抑制状態を表示 | 段階1 |
| 3 | gate 1 (生成前スキップ) + 抑制件数の通知 | 段階1 |
| 4 | rating 取得 → `reply` のリスク階層化 | なし (並行可) |
| 5 | 承認撤廃 | 段階1-4が安定してから |

**段階0を先に。** 名前一致の取り違えリスクは、抑制リストの有無に関わらず今そこにある。
