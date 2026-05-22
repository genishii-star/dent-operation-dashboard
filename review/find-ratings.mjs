/**
 * ★カテゴリ別スコア + スーパーホスト情報取得
 */

import { chromium } from 'playwright';
import { existsSync } from 'fs';

const accountName = process.argv[2] || 'OPH';
const sessionPath = `sessions/${accountName}.json`;

if (!existsSync(sessionPath)) {
  console.error(`セッションが見つかりません: ${sessionPath}`);
  process.exit(1);
}

const browser = await chromium.launch({ headless: false });
const context = await browser.newContext({ storageState: sessionPath });
const page = await context.newPage();

// --- 1. Insights ページを開く ---
console.log('=== Insights ページ ===');
await page.goto('https://www.airbnb.com/hosting/insights', { waitUntil: 'domcontentloaded', timeout: 30000 });
await page.waitForTimeout(5000);

// --- 2. 各カテゴリタブをクリックしてスコア取得 ---
const categories = ['Overall quality', 'Accuracy', 'Check-in', 'Cleanliness', 'Communication', 'Location', 'Value'];
const scores = {};

for (const cat of categories) {
  console.log(`\n--- ${cat} ---`);
  try {
    // タブをクリック
    await page.click(`text="${cat}"`, { timeout: 5000 });
    await page.waitForTimeout(2000);

    // スコア数値を取得（ページ内の表示から）
    const data = await page.evaluate((category) => {
      const body = document.body.innerText;

      // "Rating X out of 5" パターン
      const ratingPattern = /Rating\s+(\d(?:\.\d+)?)\s+out of\s+5/g;
      const ratings = [];
      let match;
      while ((match = ratingPattern.exec(body)) !== null) {
        ratings.push(parseFloat(match[1]));
      }

      // "X.XX" や "X%" のパターン
      // "5-star ratings" の近くにある数値
      const fiveStarMatch = body.match(/(\d{1,3}(?:\.\d+)?)%?\s*\n?\s*5-star\s+rating/i);
      const overallMatch = body.match(/(\d\.\d{1,2})\s*\n?\s*Overall\s+rating/i);

      // ダッシュ（-）表示の場合（データ不足）
      const noData = body.includes('Your average') && body.includes('will show here once');

      return {
        ratings,
        fiveStar: fiveStarMatch ? fiveStarMatch[1] : null,
        overall: overallMatch ? overallMatch[1] : null,
        noData,
        // メインコンテンツエリアの数値を探す
        bodySnippet: body.substring(0, 3000)
      };
    }, cat);

    scores[cat] = data;
    console.log(`  5-star%: ${data.fiveStar || '-'}`);
    console.log(`  Overall: ${data.overall || '-'}`);
    console.log(`  データなし: ${data.noData}`);
    if (data.ratings.length > 0) console.log(`  個別Rating: ${data.ratings.slice(0, 5).join(', ')}`);

  } catch (e) {
    console.log(`  エラー: ${e.message.substring(0, 80)}`);
  }
}

// --- 3. Superhostタブ ---
console.log('\n\n=== Superhost ===');
try {
  await page.click('text="Superhost"', { timeout: 5000 });
  await page.waitForTimeout(3000);

  await page.screenshot({ path: 'debug-superhost.png', fullPage: true });

  const superhostData = await page.evaluate(() => {
    const body = document.body.innerText;

    // スーパーホスト関連の情報を抽出
    const lines = body.split('\n').map(l => l.trim()).filter(Boolean);
    const shLines = [];
    let capture = false;
    for (const line of lines) {
      if (line === 'Superhost' && !capture) {
        capture = true;
        continue;
      }
      if (capture) {
        shLines.push(line);
        if (shLines.length > 30) break;
      }
    }

    return {
      superhostSection: shLines.join('\n'),
      isSuperhost: body.includes("You're a Superhost") || body.includes('Superhost status'),
      bodySnippet: body.substring(body.indexOf('Superhost'), body.indexOf('Superhost') + 1500)
    };
  });

  console.log(superhostData.superhostSection.substring(0, 1000));

} catch (e) {
  console.log(`エラー: ${e.message.substring(0, 80)}`);
}

// --- 4. "Show 55 reviews" をクリックして個別★取得 ---
console.log('\n\n=== 個別レビュー★ ===');
try {
  // Overall qualityタブに戻る
  await page.click('text="Overall quality"', { timeout: 5000 });
  await page.waitForTimeout(2000);

  // "Show XX reviews" ボタンをクリック
  const showBtn = await page.$('text=/Show \\d+ reviews/');
  if (showBtn) {
    await showBtn.click();
    await page.waitForTimeout(3000);

    // 個別レビューの★を取得
    const reviewRatings = await page.evaluate(() => {
      const body = document.body.innerText;
      const results = [];

      // "名前\n日付 • 物件名\nOverall quality\nRating X out of 5\nX\n本文" パターン
      const blocks = body.split('Overall quality');
      for (let i = 1; i < blocks.length; i++) {
        const block = blocks[i];
        const ratingMatch = block.match(/Rating\s+(\d)\s+out of\s+5/);
        const prevBlock = blocks[i - 1];
        const lines = prevBlock.split('\n').filter(l => l.trim());
        const nameLine = lines[lines.length - 2] || '';
        const dateLine = lines[lines.length - 1] || '';

        if (ratingMatch) {
          results.push({
            guest: nameLine.trim(),
            dateRange: dateLine.trim().substring(0, 80),
            rating: parseInt(ratingMatch[1])
          });
        }
      }
      return results;
    });

    console.log(`${reviewRatings.length} 件のレビュー★取得:`);
    reviewRatings.forEach(r => {
      console.log(`  ★${r.rating} ${r.guest} (${r.dateRange})`);
    });
  } else {
    console.log('"Show reviews" ボタンが見つかりません');
  }

} catch (e) {
  console.log(`エラー: ${e.message.substring(0, 80)}`);
}

console.log('\n15秒後にブラウザを閉じます...');
await page.waitForTimeout(15000);

await browser.close();
console.log('完了');
