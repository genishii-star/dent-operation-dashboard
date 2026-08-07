/**
 * Airbnb セッション切れ復旧ワンショット
 *
 *   node relogin.mjs NAGAI
 *
 * ※ Terminal.app で直接実行してください（ヘッド付きブラウザ + 2FA手入力のため
 *    Claude Code からは実行不可）。
 *
 * login.mjs → cat → curl POST → 検証 curl の4手を1コマンドにまとめたもの。
 *   1. ヘッド付きブラウザでログイン（ID/PWは自動入力、2FAは手動）
 *   2. /hosting に到達できるか検証（ログイン壁に戻されたらやり直し）
 *   3. storageState を sessions/<ACCOUNT>.json に保存
 *   4. Worker (KV) に POST
 *   5. 取得し直して cookie 件数を確認
 *
 * 必要な環境変数 (CF_ACCESS_CLIENT_ID / CF_ACCESS_CLIENT_SECRET) は
 * ~/.config/dent/review.env にあるので、頭に source を付けて実行するだけでよい:
 *
 *   source ~/.config/dent/review.env && node relogin.mjs NAGAI
 *
 * このファイルを失った場合の再取得先は Cloudflare ではなく GAS。
 * Access の Service Token は Client Secret を後から表示できない仕様なので、
 * GAS「運営アラート」プロジェクト > プロジェクトの設定 > スクリプト プロパティ
 * から同じ値をコピーする (ローテートすると朝報と D1同期が両方止まる)。
 */

import { chromium } from "playwright";
import { mkdirSync, readFileSync, writeFileSync, existsSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import { createInterface } from "readline";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_API = "https://api.dent-inc.com";

const ACCOUNT = process.argv.slice(2).find((a) => !a.startsWith("-"));
if (!ACCOUNT || !/^[A-Za-z0-9_-]{1,32}$/.test(ACCOUNT)) {
  console.error("Usage: node relogin.mjs <ACCOUNT>   例: node relogin.mjs NAGAI");
  process.exit(2);
}

const { CF_ACCESS_CLIENT_ID, CF_ACCESS_CLIENT_SECRET } = process.env;
for (const [k, v] of Object.entries({ CF_ACCESS_CLIENT_ID, CF_ACCESS_CLIENT_SECRET })) {
  if (!v) {
    console.error(`missing env: ${k}`);
    console.error("  export CF_ACCESS_CLIENT_ID=... CF_ACCESS_CLIENT_SECRET=... してから再実行してください");
    process.exit(2);
  }
}
const CF_HEADERS = {
  "CF-Access-Client-Id": CF_ACCESS_CLIENT_ID,
  "CF-Access-Client-Secret": CF_ACCESS_CLIENT_SECRET,
};

const accountsPath = join(__dirname, "accounts.json");
if (!existsSync(accountsPath)) {
  console.error("accounts.json が見つかりません");
  process.exit(2);
}
const accounts = JSON.parse(readFileSync(accountsPath, "utf8"));
if (!accounts[ACCOUNT]) {
  console.error(`アカウント "${ACCOUNT}" が accounts.json に見つかりません`);
  process.exit(2);
}
const { email, password } = accounts[ACCOUNT];

const rl = createInterface({ input: process.stdin, output: process.stdout });
const ask = (msg) => new Promise((resolve) => rl.question(msg, resolve));

const sessionDir = join(__dirname, "sessions");
mkdirSync(sessionDir, { recursive: true });
const sessionPath = join(sessionDir, `${ACCOUNT}.json`);

console.log(`\n=== ${ACCOUNT} (${email}) セッション再取得 ===\n`);

const browser = await chromium.launch({ headless: false });
const context = await browser.newContext();
const page = await context.newPage();

// ---------- 1. ログイン ----------
try {
  await page.goto("https://www.airbnb.com/login", { waitUntil: "networkidle", timeout: 30000 });
  await page.waitForTimeout(3000);

  const emailLink = page.locator(
    'div[data-testid="social-auth-button-email"], button:has-text("Continue with email"), a:has-text("Continue with email"), button:has-text("メールで続行")',
  );
  if ((await emailLink.count()) > 0) {
    await emailLink.first().click();
    await page.waitForTimeout(2000);
  }

  await page.locator("input:visible").first().fill(email);
  console.log("  ✓ メール入力完了");
  await page.waitForTimeout(500);
  await page.locator('button[type="submit"]:visible').first().click();
  await page.waitForTimeout(3000);

  const pwInput = page.locator('input[type="password"]:visible');
  if ((await pwInput.count()) > 0) {
    await pwInput.first().fill(password);
    console.log("  ✓ パスワード入力完了");
    await page.waitForTimeout(500);
    await page.locator('button[type="submit"]:visible').first().click();
    console.log("  ✓ ログインボタンクリック済み");
  } else {
    console.log("  パスワード欄が見つからない — 手動で入力してください");
  }
} catch (e) {
  console.log("  自動入力に失敗 — ブラウザで手動ログインしてください");
  console.log(`  (${e.message.split("\n")[0]})`);
}

// ---------- 2. ログイン成立を検証 ----------
// 「このデバイスを記憶」を必ずチェックすること（セッション寿命が延びる）。
console.log("\n  → 2FAが出たらブラウザで入力し、「このデバイスを記憶」をチェックしてください");

let verified = false;
for (let attempt = 1; attempt <= 3 && !verified; attempt++) {
  await ask("  ログイン完了したら Enter → ");
  await page.goto("https://www.airbnb.com/hosting", { waitUntil: "domcontentloaded", timeout: 30000 });
  await page.waitForTimeout(3000);
  if (page.url().includes("/login")) {
    console.log(`  ✗ まだログイン壁に飛ばされます (${page.url()}) — ブラウザで続きを完了してください`);
  } else {
    verified = true;
  }
}
if (!verified) {
  console.error("\n  ログインを検証できませんでした。KVへのアップロードは中止します。");
  rl.close();
  await browser.close();
  process.exit(1);
}
console.log("  ✓ /hosting に到達 — ログイン成立");

// ---------- 3. 保存 ----------
await context.storageState({ path: sessionPath });
const localCount = JSON.parse(readFileSync(sessionPath, "utf8")).cookies?.length ?? 0;
console.log(`  ✓ セッション保存: ${sessionPath} (cookies: ${localCount})`);

rl.close();
await context.close();
await browser.close();

// ---------- 4. Worker (KV) にアップロード ----------
const session = JSON.parse(readFileSync(sessionPath, "utf8"));
const post = await fetch(`${DATA_API}/internal/session/${ACCOUNT}`, {
  method: "POST",
  headers: { ...CF_HEADERS, "Content-Type": "application/json" },
  body: JSON.stringify({ session }),
});
if (!post.ok) {
  console.error(`\n  ✗ アップロード失敗 ${post.status}: ${(await post.text()).slice(0, 300)}`);
  console.error("    302 が返る場合は Service Token が失効/ローテート済み。Cloudflare側を確認してください。");
  process.exit(1);
}
console.log("  ✓ Worker (KV) にアップロード完了");

// ---------- 5. 読み戻し検証 ----------
const get = await fetch(`${DATA_API}/internal/session/${ACCOUNT}`, { headers: CF_HEADERS });
if (!get.ok) {
  console.error(`  ✗ 読み戻し失敗 ${get.status}: ${(await get.text()).slice(0, 300)}`);
  process.exit(1);
}
const remoteCount = (await get.json()).session?.cookies?.length ?? 0;
console.log(`  ✓ 読み戻し確認: cookies ${remoteCount} 件`);
if (remoteCount !== localCount) {
  console.warn(`  ⚠ ローカル(${localCount})と件数が一致しません — 念のため中身を確認してください`);
}

console.log(`
=== 完了 ===
週次ワークフローを手動で回して復旧を確認:

  gh workflow run review-generate-weekly.yml -f account=${ACCOUNT}
  gh run watch
`);
