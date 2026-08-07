# Airbnb レビュー自動化 セットアップ手順

GitHub Actions で週次実行する自動化システムのセットアップ手順。

**通常運用は Mac 不要。** セッションが切れた時だけ手元で `login.mjs` を再実行する。
- TOTP 2FA が有効なアカウント → 完全自動 (seed さえ KV にあれば GH Actions 内で再ログインまで完結)
- TOTP 2FA を有効化できないアカウント (Airbnb UI に項目が出ないケース) → セッション切れ時に Slack 通知 → 手動再ログイン

---

## アーキテクチャ概観

```
[GH Actions]                              [Cloudflare]
  cron (Mon 9:00 JST)                       Workers KV
    │                                        ├─ airbnb-session:{account}  (暗号化)
    ▼                                        ├─ airbnb-totp:{account}     (暗号化)
  generate-drafts.mjs ─── Worker ──────────►│
    ↓                     /internal/session/                  D1
    ↓                     /internal/totp/                     ├─ review_drafts
    ↓                     /internal/review-drafts
  Playwright で取得+生成
    ↓
  Slack #review-approval (Block Kit)
                            ↓
                      [人間: ✅承認/❌却下]
                            ↓
                      slack-bot Worker → D1 status 更新

  cron (Mon 18:00 JST)
    │
    ▼
  post-approved.mjs ─── 同 Worker ────────► D1 SELECT WHERE status='approved'
    ↓
  Playwright で投稿
    ↓
  D1 status='posted' / 'failed'
```

---

## Phase 0: 基盤セットアップ (1回だけ)

### 0-1. SESSION_ENC_KEY を生成し Worker Secret に登録

KV に置く Airbnb セッション / TOTP seed の AES-256-GCM 暗号化キー。

```bash
# 64 hex 文字 (32 bytes) のキーを生成
openssl rand -hex 32
# → 例: 8f3a9c1d... (このまま使用)

cd wiki/workers/dent-data-api
npx wrangler secret put SESSION_ENC_KEY
# プロンプトに上の hex 文字列をペースト
```

**⚠️ このキーが失われると全 Airbnb セッション/TOTP seed が復号不能になる**。
1Password など安全な場所にバックアップ。

### 0-2. Worker をデプロイ

```bash
cd wiki/workers/dent-data-api
npx tsc --noEmit                            # 型チェック
npx wrangler d1 migrations apply dent-platform-d1 --remote   # review_drafts 作成
npx wrangler deploy
```

### 0-3. Service Token を確認

GitHub Actions から Cloudflare Access 越しに Worker を呼ぶための Service Token。
既存の GAS 用トークンを流用可（[`cloudflare_access_service_token`](memory)）。
無い場合は Cloudflare ZeroTrust → Access → Service Auth から作成。

---

## Phase 1: アカウント X の追加 (アカウントごと)

以下、アカウント名を `X` とする。アルファベット数字 _ - のみ、最長32字。

### 1-1. (オプション) Airbnb 側で TOTP 2FA を有効化

> **アカウントによっては UI に 2-step verification 項目が出ない** (Airbnb が passkey 推しに切り替え中で、TOTP トグルを段階的に廃止している模様)。
> 出なければ **このセクションはまるごとスキップ** して 1-4 に飛ぶ。コード側は TOTP 未登録を検出すると、セッション切れ時に Slack 通知を出して落ちる挙動 (手動再ログインで復旧)。

1. Airbnb にログイン → Account → Login & Security → 2-step verification
2. Authenticator app を選択
3. QR コード表示画面で **"Can't scan?"** をクリック → base32 形式の seed が表示される
4. その seed をコピー (例: `JBSWY3DPEHPK3PXP`)

### 1-2. (オプション) TOTP seed を KV に登録

```bash
export CF_AID="<CF-Access-Client-Id>"
export CF_SECRET="<CF-Access-Client-Secret>"

curl -X POST https://api.dent-inc.com/internal/totp/X \
  -H "CF-Access-Client-Id: $CF_AID" \
  -H "CF-Access-Client-Secret: $CF_SECRET" \
  -H "Content-Type: application/json" \
  -d '{"seed": "JBSWY3DPEHPK3PXP"}'
```

### 1-3. (オプション) TOTP 動作確認

```bash
curl -H "CF-Access-Client-Id: $CF_AID" \
     -H "CF-Access-Client-Secret: $CF_SECRET" \
     https://api.dent-inc.com/internal/totp/X
# → {"account":"X","code":"123456","expires_in":18}
```

返ってきた 6 桁を Airbnb の 2FA 確認画面に入力して登録を完了。

### 1-4. 初回セッション作成 (Mac でローカル実行)

> TOTP を登録していない場合、初回ログイン時は Airbnb が SMS / メール / passkey などの 2FA チャレンジを送ってくる。**手元で手動でこなせば OK**。「このデバイスを記憶」にチェックを入れれば数週間〜数ヶ月セッションが持つ。

```bash
cd operation/review
# accounts.json に X の email/password を追記しておく
node login.mjs X
# ブラウザが開く → ID/PW 自動入力 → 2FA コード入力 (上記 curl の結果使う)
# 「このデバイスを記憶」を必ずチェック
# 完了したらターミナルで Enter → sessions/X.json が保存される
```

### 1-5. セッションを KV にアップロード

```bash
SESSION=$(cat sessions/X.json)
curl -X POST https://api.dent-inc.com/internal/session/X \
  -H "CF-Access-Client-Id: $CF_AID" \
  -H "CF-Access-Client-Secret: $CF_SECRET" \
  -H "Content-Type: application/json" \
  -d "{\"session\": $SESSION}"
```

### 1-6. 取得検証

```bash
curl -H "CF-Access-Client-Id: $CF_AID" \
     -H "CF-Access-Client-Secret: $CF_SECRET" \
     https://api.dent-inc.com/internal/session/X | jq '.session.cookies | length'
# → cookies の件数 (10〜30 件) が表示されればOK
```

---

## Phase 2: GitHub Actions ワークフロー (Phase 1 実装後)

> ⚠️ Phase 1 (generate-drafts.mjs / post-approved.mjs / Slack handler) が未実装。
> 以下は実装後に有効化する手順。

### 2-1. GitHub Secrets を登録

リポジトリ Settings → Secrets and variables → Actions:

| Secret 名 | 内容 |
|---|---|
| `CF_ACCESS_CLIENT_ID` | Cloudflare Access Service Token (ID) |
| `CF_ACCESS_CLIENT_SECRET` | 同上 (Secret) |
| `ANTHROPIC_API_KEY` | Claude API キー (sk-ant-...) |
| `SLACK_BOT_TOKEN` | xoxb-... (slack-bot Worker と共有) |
| `SLACK_REVIEW_CHANNEL` | `#review-approval` の channel ID (例: `C0...`) |
| `ACCOUNTS_JSON` | アカウント別 email/password の JSON ブロブ。例: `{"X":{"email":"host@example.com","password":"..."}}` セッション切れ時の TOTP 再ログインに必要 |

### 2-2. Slack 設定

1. `#review-approval` チャンネル新設、dent-slack-bot を招待
2. Slack App → Interactivity → Request URL に slack-bot Worker のエンドポイント設定
   (Phase 1 で追加: `/slack/interactions/review-approval`)

### 2-3. ワークフロー有効化

`.github/workflows/review-generate-weekly.yml` `.github/workflows/review-post-weekly.yml` が
自動で有効化される。手動テストは Actions タブから `workflow_dispatch` で。

---

## トラブルシューティング

### TOTP コードが Airbnb に通らない
- Worker と Airbnb サーバの時刻ずれ。再試行で大抵直る (TOTP は前後 1 step を許容するサーバ多い)
- seed の改行/空白混入。`base32` 正規表現で正しいか確認

### セッションが切れた
- `GET /internal/session/X` を叩いて取得 → Playwright で動かしたら 403 や Login wall に飛ぶ
- **TOTP 登録済みアカウント**: GH Actions が `SESSION_EXPIRED_NO_AUTO_RECOVERY` を出さない限り自動復旧されている (本来そう動くはずだが何度も発火するなら TOTP seed の不一致を疑う)
- **TOTP 未登録アカウント**: GH Actions が失敗 → `#review-approval` に「session expired」アラート → **Terminal.app で `relogin.mjs` を実行**（ログイン〜KVアップロード〜検証をワンショット。`login.mjs` + 手動 curl の置き換え）

  ```bash
  cd operation/review
  source ~/.config/dent/review.env   # CF_ACCESS_CLIENT_ID / _SECRET が入っている
  node relogin.mjs X
  # ブラウザが開く → ID/PW自動入力 → 2FA手入力 →「このデバイスを記憶」を必ずチェック
  # → Enter → /hosting 到達を検証 → KVへPOST → 読み戻しでcookie件数確認
  ```

  > `~/.config/dent/review.env` を失った場合、Cloudflare からは再取得できない
  > (Service Token の Client Secret は作成時にしか表示されない)。GAS「運営アラート」
  > プロジェクト > プロジェクトの設定 > スクリプト プロパティ に同じ値がある。
  > **ローテートは避けること** — GAS の Script Property と GitHub Secrets の
  > 両方を更新しないと、朝報・D1同期・レビュー自動化がまとめて止まる。

  復旧後、該当 workflow を再実行:

  ```bash
  gh workflow run review-generate-weekly.yml -f account=X
  ```

  > 「このデバイスを記憶」を飛ばすとセッションが数日で切れて毎週アラートが鳴る。実績では
  > チェックあり = 約8週間持つ (NAGAI: 2026-06-02 取得 → 2026-07-27 失効)。

### 暗号化キーをローテートしたい
- 旧キーで全 KV エントリを decrypt → 新キーで encrypt → KV に上書き
- ローテーションスクリプトは未実装 (将来課題)

---

## セキュリティ整理

| 機密 | 保存場所 | 防御層 |
|---|---|---|
| `SESSION_ENC_KEY` | Worker Secret | Cloudflare 管理。Worker code 内のみ参照可 |
| Airbnb cookie | KV (暗号化済) | KV暗号化 + アプリ層 AES-256-GCM |
| TOTP seed | KV (暗号化済) | 同上 |
| Airbnb ID/PW | `accounts.json` (Mac ローカル) | gitignore済。GH Actionsには不要 (セッションがあれば足りる) |
| Claude API key | GH Secrets | GitHub暗号化。Workflow ログには出さない |
| Service Token | GH Secrets + Mac env | 同上 |

ID/PW を GH Actions に置かない設計：セッション + TOTP だけで再ログイン可能なため。
