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

`com.dent.all-hands.plist` = **毎月20日 10:07**（Mac mini `dent-mac` の launchd。2026-10-17 に MacBook Air から移した）。
JNTO 訪日外客統計の発表が月によって 15日〜19日とズレるため、20日なら前月分が出揃っている。

gen ユーザーは画面からログインしないので、LaunchAgent ではなく **LaunchDaemon（`UserName`=gen）** で動かす。登録・変更には `dent` の管理者権限が要る（`buddy` で入って実行）:

```bash
sudo cp /Users/gen/Client/_Dent/claude-code/operation/scripts/all-hands/com.dent.all-hands.plist /Library/LaunchDaemons/
sudo chown root:wheel /Library/LaunchDaemons/com.dent.all-hands.plist
sudo launchctl bootstrap system /Library/LaunchDaemons/com.dent.all-hands.plist
sudo launchctl kickstart -k system/com.dent.all-hands    # 手動で即実行
sudo launchctl bootout system/com.dent.all-hands         # 停止
```

gen のままで手動実行するなら `node scripts/all-hands/build-ah.mjs`（下の「動かし方」）でよい。
Mac mini で入れたもの: `review/` の `npm install`、Playwright の Chrome Headless Shell（Node 26 だと `npx playwright install` の解凍が固まるので、zip を `~/Library/Caches/ms-playwright/chromium_headless_shell-1217/` に手で展開）、`openpyxl`（`/usr/bin/python3` と Homebrew の両方に `--user`）。

launchd からは **常に `origin/main` の最新コードで動く**（`~/.dent/ah-main` に一時 worktree を作って実行し、終わったら消す）。
作業フォルダが別ブランチでも古いスクリプトで走ることはなく、生成物も必ず `main` に push される（op.dent-inc.com は main 配信）。
plist を変えたら上の `cp` → `bootout` → `bootstrap` で入れ直すこと。

ログは Mac mini gen の `~/.dent/all-hands.log`。

## 設定（git に入れない）

Mac mini gen の `~/.dent/ah.env`（`chmod 600`）:

```
DECK_PASSWORD=...    デッキの復号パスワード（ダッシュボードと同じ）
NOTIFY_URL=https://dent-slack-bot.gen-ishii.workers.dev/api/notify
NOTIFY_TOKEN=...     同エンドポイントの Bearer トークン
```

D1 は wrangler のローカル OAuth、push は `gh` の認証をそのまま使うので追加の資格情報は要らない。
ただし gen の `gh` は Tide のアカウント（genshimizu-tide）が既定で、このリポジトリに push できない。plist の中で `GH_TOKEN=$(gh auth token --user genishii-star)` を渡している。手で実行するときも同じものを付けること。
逆に言うと **Mac mini（gen）でしか動かない**。CI やクラウドに移すなら D1 のサービストークンと push 用 PAT を別途用意すること。

## データの出どころ

| 中身 | 出どころ | 備考 |
|---|---|---|
| 売上・予約組数・ゲスト数・チャネル・国籍 | D1 `reservations`（チェックイン月） | 純キャンセル(販売額0)と「ブロックされた」を除外 |
| 室泊数・ADR・稼働率・物件数 | D1 `daily_revenue` | ユニーク prop-day、清掃料のみの行を除外 |
| 物件マスタ | wiki の `origin/main:site/data/facilities` | 作業ツリーは他人のブランチのことがあるので必ず origin/main から |
| Tide 活動実績（清掃件数・物件・当日ターン・スタッフ数など） | Tide アカウントの KV `ops-state`（tide-ops-api の清掃ジョブ台帳） | **売上・原価は載せない**。CO日基準・「済」のみ・スポットは別扱い＝`/api/segment` と同じ数え方。下記 |
| 訪日外客数 | JNTO 公式Excel（`_files/<日付>_1615-5.xlsx`） | 翌月中旬発表。国別も最新月まで揃う |

Tide の KV は Tide の Cloudflare アカウントにあり、gen の wrangler（Dent アカウント）では読めない。
`_Tide/07-deploy/.env` の `CLOUDFLARE_API_TOKEN` を**読むだけ**で使う（`finance(Tide)/deploy.sh` と同じ扱い。場所を変えるなら `ah.env` に `TIDE_ENV=` を書く）。
取れなかったときはデッキ全体は止めず、Tide スライドを「取得できませんでした」にして Slack に警告を出す。
集計は `lib/tide.mjs`。tide-ops-api の数え方（`readState` の差分畳み込み・`isSpotJob`・CO日基準）を変えたらここも直すこと。

**内製化率**（エリア別）＝当月に Tide が清掃した棟 ÷ 物件マスタ「稼働中」の棟のうちオーナー清掃を除いたもの。オーナー清掃の判定は facilities YAML の `cleaning_arranged_by: owner` だけを見る（2026-10-23 石井さん）。マスタ上オーナー手配でも Tide が清掃していれば分母に残す（ETB）。**オーナー清掃の物件が増えたら YAML に `cleaning_arranged_by: owner` を付けること**（付け忘れると分母が膨らんで率が下がる）。

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
