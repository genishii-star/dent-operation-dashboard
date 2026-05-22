/**
 * 系統A: ゲストへのレビュー投稿（★5全軸 + 公開レビュー + 推薦Yes）
 *
 * 使い方:
 *   node post-review.mjs <account-name>              # ドライラン（未投稿一覧を表示）
 *   node post-review.mjs <account-name> --post        # 実際に投稿
 *   node post-review.mjs <account-name> --gui          # GUI表示
 *   node post-review.mjs <account-name> --post --gui   # GUI + 投稿
 *
 * レビューテキスト:
 *   reviews-{account}.json に以下の形式で用意:
 *   { "ゲスト名": "レビュー本文", ... }
 *   ファイルがなければ未投稿一覧を表示して終了。
 */

import { chromium } from 'playwright';
import { existsSync, readFileSync } from 'fs';

const accountName = process.argv[2] || 'OPH';
const sessionPath = `sessions/${accountName}.json`;
const headless = !process.argv.includes('--gui');
const doPost = process.argv.includes('--post');
const reviewsPath = `reviews-${accountName}.json`;
const outputPath = `output-${accountName}.json`;

if (!existsSync(sessionPath)) {
  console.error(`セッションが見つかりません: ${sessionPath}`);
  process.exit(1);
}

if (!existsSync(outputPath)) {
  console.error(`出力ファイルが見つかりません: ${outputPath}`);
  console.error(`先に node fetch-reviews.mjs ${accountName} を実行してください`);
  process.exit(1);
}

// 未投稿レビュー読み込み
const output = JSON.parse(readFileSync(outputPath, 'utf-8'));
const pending = output.pendingReviews || [];
if (pending.length === 0) {
  console.log('未投稿レビューはありません');
  process.exit(0);
}

// ゲスト名をテキストから抽出（"13 days left\nLeave Yours a review\n..." → "Yours"）
function extractGuestName(text) {
  const match = text.match(/Leave\s+(.+?)\s+a review/);
  return match ? match[1] : null;
}

console.log(`\n=== ${accountName}: 未投稿レビュー ${pending.length} 件 ===\n`);
for (const p of pending) {
  const guest = extractGuestName(p.text);
  console.log(`  - ${guest || '(名前不明)'}: ${p.text.split('\n')[0]}`);
}

// レビューテキスト読み込み
let reviewTexts = {};
if (existsSync(reviewsPath)) {
  reviewTexts = JSON.parse(readFileSync(reviewsPath, 'utf-8'));
  console.log(`\nレビューデータ読み込み: ${reviewsPath} (${Object.keys(reviewTexts).length} 件)`);
} else {
  console.log(`\nレビューデータなし: ${reviewsPath}`);
  console.log('上記のゲスト名で reviews-${accountName}.json を作成してください');
  console.log('形式: { "ゲスト名": "レビュー本文", ... }');
  process.exit(0);
}

if (!doPost) {
  console.log('\n--- ドライラン ---');
  for (const p of pending) {
    const guest = extractGuestName(p.text);
    const text = guest ? reviewTexts[guest] : null;
    console.log(`\n[${guest}]`);
    if (text) {
      console.log(`  レビュー: ${text.substring(0, 100)}...`);
      console.log(`  URL: ${p.href}`);
    } else {
      console.log(`  → レビューテキストなし（スキップ）`);
    }
  }
  console.log('\n実投稿するには --post を付けて再実行してください');
  process.exit(0);
}

// === 実投稿 ===
const browser = await chromium.launch({ headless });
const context = await browser.newContext({ storageState: sessionPath });
const page = await context.newPage();

let posted = 0;
let skipped = 0;

for (const p of pending) {
  const guest = extractGuestName(p.text);
  const reviewText = guest ? reviewTexts[guest] : null;

  if (!reviewText) {
    console.log(`\n[${guest || '名前不明'}] レビューテキストなし → スキップ`);
    skipped++;
    continue;
  }

  console.log(`\n[${guest}] 投稿開始...`);
  console.log(`  URL: ${p.href}`);

  try {
    await page.goto(p.href, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(5000);

    // 最大10回ループ（安全策）
    for (let loop = 0; loop < 10; loop++) {
      const bodyText = await page.evaluate(() => document.body.innerText);

      if (bodyText.includes('Your review has been submitted') || bodyText.includes('Thanks for your review')) {
        console.log('  ✓ 投稿完了');
        break;
      }

      // ステップ番号を取得（"step N out of 7"）
      const stepMatch = bodyText.match(/step (\d+) out of (\d+)/);
      const stepNum = stepMatch ? parseInt(stepMatch[1]) : -1;

      if (stepNum <= 0) {
        // Step 0 or イントロ → Continueがあればクリック
        const hasContinue = await page.locator('button:has-text("Continue")').count();
        if (hasContinue > 0) {
          await page.click('button:has-text("Continue")');
          console.log('  Step 0: Continue');
          await page.waitForTimeout(3000);
          continue;
        }
        // Continueがない場合はStep 1と同じ扱い（★評価画面の可能性）
        const hasRadio = await page.locator('input[type="radio"]').count();
        if (hasRadio > 0) {
          // Step 1として処理を続行（stepNum=1にフォールスルー）
        } else {
          console.log(`  (不明なステップ: step=${stepNum})`);
          console.log(`  bodyText: ${bodyText.substring(0, 300)}`);
          break;
        }
      }

      if ((stepNum >= 1 && stepNum <= 3) || stepNum <= 0) {
        // Step 1-3: ★評価 — 最後のradioボタン（★5）をクリック
        const clicked = await page.evaluate(() => {
          const radios = [...document.querySelectorAll('input[type="radio"]')];
          if (radios.length > 0) {
            radios[radios.length - 1].click();
            return `radio-${radios.length}`;
          }
          return null;
        });
        await page.waitForTimeout(500);
        await page.click('button:has-text("Next")');
        console.log(`  Step ${stepNum}/7: ★5 (${clicked})`);
        await page.waitForTimeout(3000);
        continue;
      }

      if (stepNum === 4) {
        // Step 4: 公開レビュー本文
        await page.fill('textarea', reviewText);
        await page.waitForTimeout(500);
        await page.click('button:has-text("Next")');
        console.log(`  Step 4/7: レビュー本文 (${reviewText.length}文字)`);
        await page.waitForTimeout(3000);
        continue;
      }

      if (stepNum === 5) {
        // Step 5: 推薦 → Yes
        await page.click('button:has-text("Yes")');
        await page.waitForTimeout(500);
        await page.click('button:has-text("Next")');
        console.log('  Step 5/7: 推薦 Yes');
        await page.waitForTimeout(3000);
        continue;
      }

      if (stepNum === 6) {
        // Step 6: プライベートノート → Submit
        await page.click('button:has-text("Submit")');
        console.log('  Step 6/7: Submit');
        await page.waitForTimeout(5000);
        continue;
      }

      console.log(`  (不明なステップ: step=${stepNum})`);
      console.log(`  bodyText: ${bodyText.substring(0, 300)}`);
      break;
    }

    posted++;

  } catch (e) {
    console.log(`  ✗ エラー: ${e.message.split('\n')[0]}`);
    skipped++;
  }
}

await browser.close();

console.log(`\n=== 完了: 投稿 ${posted} 件, スキップ ${skipped} 件 ===`);
