# All Hands デッキの月次自動生成

`op.dent-inc.com/all-hands/YYYY-MM.html` を毎月自動で作って push し、Slack DM で知らせる。

- ファイル名の `YYYY-MM` は **開催月**。中身は前月実績（`2026-08.html` = 8月会 = 7月まとめ）
- スライドの雛形は [`../../all-hands/_template.html`](../../all-hands/_template.html)。数値は AES-256-GCM で暗号化して埋め込むので、public repo でソースを見られても平文の売上は出ない

## 動かし方

```bash
cd operation
node scripts/all-hands/build-ah.mjs              # 前月分を作って push + Slack 通知
node scripts/all-hands/build-ah.mjs 2026-07      # 実績月を明示（作り直し・過去分の再生成）
node scripts/all-hands/build-ah.mjs 2026-07 --dry-run   # 生成と検証だけ。push も通知もしない
```

同じ実績月で再実行すると同じファイルを上書きし、一覧（`index.html`）は二重に追加しない。
JNTO の発表を待って作り直したいときは、そのまま同じコマンドをもう一度叩けばよい。

## スケジュール

`com.dent.all-hands.plist` = **毎月20日 10:07**（このMacの launchd）。
JNTO 訪日外客統計の発表が月によって 15日〜19日とズレるため、20日なら前月分が出揃っている。

```bash
cp scripts/all-hands/com.dent.all-hands.plist ~/Library/LaunchAgents/
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.dent.all-hands.plist
launchctl kickstart -k gui/$(id -u)/com.dent.all-hands   # 手動で即実行
launchctl bootout gui/$(id -u)/com.dent.all-hands        # 停止
```

ログは `~/.dent/all-hands.log`。Mac がスリープしていて20日を跨いだ場合、launchd は起床時に1回だけ遅れて実行する。

## 設定（git に入れない）

`~/.dent/ah.env`（`chmod 600`）:

```
DECK_PASSWORD=...    デッキの復号パスワード（ダッシュボードと同じ）
NOTIFY_URL=https://dent-slack-bot.gen-ishii.workers.dev/api/notify
NOTIFY_TOKEN=...     同エンドポイントの Bearer トークン
```

D1 は wrangler のローカル OAuth、push は `gh` の認証をそのまま使うので追加の資格情報は要らない。
逆に言うと **このMacでしか動かない**。CI やクラウドに移すなら D1 のサービストークンと push 用 PAT を別途用意すること。

## データの出どころ

| 中身 | 出どころ | 備考 |
|---|---|---|
| 売上・予約組数・ゲスト数・チャネル・国籍 | D1 `reservations`（チェックイン月） | 純キャンセル(販売額0)と「ブロックされた」を除外 |
| 室泊数・ADR・稼働率・物件数 | D1 `daily_revenue` | ユニーク prop-day、清掃料のみの行を除外 |
| 物件マスタ | wiki の `origin/main:site/data/facilities` | 作業ツリーは他人のブランチのことがあるので必ず origin/main から |
| 訪日外客数 | JNTO 公式Excel（`_files/<日付>_1615-5.xlsx`） | 翌月中旬発表。国別も最新月まで揃う |

**前月・前年も毎回まるごと再集計する**（like-for-like）。過去月の実績は再同期で少しずつ動くので、
前回デッキの値を据え置くとダッシュボードと食い違う。前月比・前年比は常に同じスコープの比較になる。

集計式は `operation/app.js` の月次KPIに合わせてある。**app.js を変えたら `lib/aggregate.mjs` も直すこと。**

## 壊れたときに気付ける仕組み

- 集計が空 / 稼働物件が50室未満 → その場で異常終了（壊れたデッキを配らない）
- 未分類チャネル・YAMLスキーマの変化・JNTO の国別欠損 → 警告として Slack に併記
- 生成後にヘッドレスで復号→描画まで通し、KPI・グラフ6本・インサイト・各テーブルの中身を確認してから push
- 失敗しても Slack DM に理由を投げる（黙って止まると「今月ぶんが無い」ことに誰も気付かない）

## スライド文言について

インサイト（ハイライト・要観察）は**数値から自動生成**していて、増減の符号で文言が切り替わる。
「減収なのに『回復』と書いてある」事故を防ぐため、テンプレートに結論をハードコードしないこと。
連絡事項スライドは GAS 側に保存され、デッキ上の「編集」ボタンから直接書き換える。
