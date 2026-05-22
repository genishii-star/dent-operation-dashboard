/**
 * Airbnb 半自動ログイン → セッション保存
 *
 * 使い方:
 *   node login.mjs <account-name>        # 1アカウントログイン
 *   node login.mjs --all                  # 全アカウント順次ログイン
 *
 * ※ Terminal.appで直接実行してください（Claude Codeからは実行不可）
 *
 * メール/PWは自動入力。2FAコード入力とログイン完了確認は手動。
 * ログイン完了後、ターミナルで Enter を押すとセッション保存 → 次のアカウントへ。
 */

import { chromium } from 'playwright';
import { mkdirSync, readFileSync, existsSync } from 'fs';
import { createInterface } from 'readline';

const accountsPath = 'accounts.json';
if (!existsSync(accountsPath)) {
  console.error(`accounts.json が見つかりません`);
  process.exit(1);
}
const allAccounts = JSON.parse(readFileSync(accountsPath, 'utf-8'));

const doAll = process.argv.includes('--all');
const accountName = process.argv.filter(a => !a.startsWith('-'))[2];

if (!doAll && !accountName) {
  console.error('Usage: node login.mjs <account-name>  or  node login.mjs --all');
  process.exit(1);
}

const targets = doAll
  ? Object.keys(allAccounts)
  : [accountName];

for (const name of targets) {
  if (!allAccounts[name]) {
    console.error(`アカウント "${name}" が accounts.json に見つかりません`);
    process.exit(1);
  }
}

const sessionDir = 'sessions';
mkdirSync(sessionDir, { recursive: true });

const rl = createInterface({ input: process.stdin, output: process.stdout });
const ask = (msg) => new Promise(resolve => rl.question(msg, resolve));

console.log(`\n=== ${targets.length} アカウントのログイン開始 ===\n`);

const browser = await chromium.launch({ headless: false });

for (let i = 0; i < targets.length; i++) {
  const name = targets[i];
  const { email, password } = allAccounts[name];
  const sessionPath = `${sessionDir}/${name}.json`;

  console.log(`[${i + 1}/${targets.length}] ${name} (${email})`);

  if (existsSync(sessionPath)) {
    const ans = await ask(`  セッション既存。スキップする？ (Y/n): `);
    if (ans.trim().toLowerCase() !== 'n') {
      console.log(`  → スキップ\n`);
      continue;
    }
  }

  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    await page.goto('https://www.airbnb.com/login', { waitUntil: 'networkidle', timeout: 30000 });
    await page.waitForTimeout(3000);

    // "Continue with email" リンクをクリック（ログイン画面にある場合）
    const emailLink = page.locator('div[data-testid="social-auth-button-email"], button:has-text("Continue with email"), a:has-text("Continue with email"), button:has-text("メールで続行")');
    if (await emailLink.count() > 0) {
      await emailLink.first().click();
      await page.waitForTimeout(2000);
    }

    // メールアドレス入力 — 表示されてるinputに入力
    const emailInput = page.locator('input:visible').first();
    await emailInput.fill(email);
    console.log(`  ✓ メール入力完了`);
    await page.waitForTimeout(500);

    // Continue / 続行 ボタンをクリック
    const continueBtn = page.locator('button[type="submit"]:visible');
    await continueBtn.first().click();
    await page.waitForTimeout(3000);

    // パスワード入力
    const pwInput = page.locator('input[type="password"]:visible');
    if (await pwInput.count() > 0) {
      await pwInput.first().fill(password);
      console.log(`  ✓ パスワード入力完了`);
      await page.waitForTimeout(500);

      // Log in ボタンをクリック
      const loginBtn = page.locator('button[type="submit"]:visible');
      await loginBtn.first().click();
      console.log(`  ✓ ログインボタンクリック済み`);
    } else {
      console.log(`  パスワード欄が見つからない — 手動で入力してください`);
    }
  } catch (e) {
    console.log(`  自動入力に失敗 — 手動でログインしてください`);
    console.log(`  (${e.message.split('\n')[0]})`);
  }

  console.log(`  → 2FAが表示されたらブラウザで入力してください`);
  await ask(`  ログイン完了したら Enter → `);

  await context.storageState({ path: sessionPath });
  console.log(`  ✓ セッション保存完了: ${sessionPath}\n`);

  await context.close();
}

rl.close();
await browser.close();

console.log('=== 完了 ===');
