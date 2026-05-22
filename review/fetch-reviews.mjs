/**
 * 系統B: Airbnb受信レビュー + ★スコア + Superhost 取得 → スプシ書き込み
 *
 * 使い方:
 *   node fetch-reviews.mjs <account-name>
 *   例: node fetch-reviews.mjs OPH
 */

import { chromium } from 'playwright';
import { existsSync, writeFileSync } from 'fs';

const accountName = process.argv[2] || 'OPH';
const sessionPath = `sessions/${accountName}.json`;
const GAS_URL = 'https://script.google.com/macros/s/AKfycbyKM4IsnGWjpXIYl9HfT6JIBs1zz8fx0ySh-05Mp7XRk3P_DQBwiDqUmUIK_cpmUaJ_uw/exec';

if (!existsSync(sessionPath)) {
  console.error(`セッションが見つかりません: ${sessionPath}`);
  process.exit(1);
}

const headless = !process.argv.includes('--gui');
const browser = await chromium.launch({ headless });
const context = await browser.newContext({ storageState: sessionPath });
const page = await context.newPage();

// ============================================================
// STEP 1: レビュー本文 + プライベートフィードバック取得
// ============================================================
console.log('=== STEP 1: レビュー本文取得 ===');
await page.goto('https://www.airbnb.com/users/reviews', { waitUntil: 'domcontentloaded', timeout: 30000 });
await page.waitForTimeout(5000);

// スクロールして読み込み
let prevHeight = 0;
for (let i = 0; i < 20; i++) {
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await page.waitForTimeout(1500);
  const newHeight = await page.evaluate(() => document.body.scrollHeight);
  if (newHeight === prevHeight) break;
  prevHeight = newHeight;
}

// ページ内レビュー抽出（プロフィールリンク + 日付パターンで検出。ボタン非依存）
async function extractReviewsFromPage() {
  return page.evaluate(() => {
    const results = [];
    const months = 'January|February|March|April|May|June|July|August|September|October|November|December';
    const dateRegex = new RegExp(`(${months})\\s+\\d{4}`);

    // プロフィールリンク（/users/show/）を起点にレビューブロックを検出
    const profileLinks = [...document.querySelectorAll('a[href*="/users/profile/"], a[href*="/users/show/"]')];

    for (const link of profileLinks) {
      const guestName = link.innerText.trim();
      if (!guestName) continue;

      // リンクの親を辿ってレビューコンテナを見つける
      let container = link;
      for (let i = 0; i < 10; i++) {
        container = container.parentElement;
        if (!container) break;
        const text = container.innerText || '';
        // コンテナは日付を含み、適度なサイズ（個別レビュー単位）
        if (text.match(dateRegex) && text.length > 30 && text.length < 2000) {
          const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

          // 日付を探す
          let reviewDate = '';
          let reviewTextStart = -1;
          for (let j = 0; j < lines.length; j++) {
            const m = lines[j].match(dateRegex);
            if (m) {
              reviewDate = m[0];
              reviewTextStart = j + 1;
              break;
            }
          }

          if (!reviewDate) break;

          // レビュー本文: 日付の次の行から、システムテキストまで
          const excludeTexts = ['Leave Public Response', 'Private feedback', 'Read more', 'Complete review', 'Review is hidden'];
          const textLines = [];
          let privateFeedback = '';
          let hasPublicResponse = !text.includes('Leave Public Response');
          let inPrivateFeedback = false;

          for (let j = reviewTextStart; j < lines.length; j++) {
            if (lines[j] === 'Private feedback') {
              inPrivateFeedback = true;
              continue;
            }
            if (inPrivateFeedback) {
              if (!excludeTexts.includes(lines[j]) && !lines[j].match(dateRegex)) {
                privateFeedback = lines[j];
              }
              break;
            }
            if (excludeTexts.includes(lines[j])) continue;
            // 次のゲストの日付行が来たら終了
            if (lines[j].match(dateRegex)) break;
            textLines.push(lines[j]);
          }

          const reviewText = textLines.join(' ');

          if (guestName && (reviewText || reviewDate)) {
            results.push({
              guestName,
              date: reviewDate,
              reviewText: reviewText.substring(0, 2000),
              privateFeedback: privateFeedback.substring(0, 2000),
              hasPublicResponse,
            });
          }
          break;
        }
      }
    }
    return results;
  });
}

// 全ページ巡回（重複排除付き）
const seen = new Set();
const allReviews = [];
let pageNum = 1;

while (true) {
  console.log(`  ページ ${pageNum}...`);
  const pageReviews = await extractReviewsFromPage();

  let newCount = 0;
  for (const r of pageReviews) {
    const key = `${r.guestName}|${r.date}|${r.reviewText.substring(0, 30)}`;
    if (!seen.has(key)) {
      seen.add(key);
      allReviews.push(r);
      newCount++;
    }
  }
  console.log(`  → ${newCount} 件（新規）`);
  if (newCount === 0) break;

  const hasNext = await page.evaluate((current) => {
    const navLinks = [...document.querySelectorAll('nav a, nav button, a[aria-label*="Page"], a[aria-label*="Next"]')];
    for (const el of navLinks) {
      if (el.innerText.trim() === String(current + 1)) { el.click(); return true; }
    }
    for (const el of navLinks) {
      const label = el.getAttribute('aria-label') || el.innerText.trim();
      if (label.includes('Next') || label.includes('次')) { el.click(); return true; }
    }
    return false;
  }, pageNum);

  if (!hasNext) break;
  await page.waitForTimeout(3000);
  pageNum++;
}

console.log(`  合計: ${allReviews.length} 件\n`);

// ============================================================
// STEP 2: カテゴリ別★スコア取得（/performance/quality/*/reviews）
// ============================================================
console.log('=== STEP 2: カテゴリ別★スコア取得 ===');

const categoryPaths = {
  'overall_quality': 'overall',
  'accuracy': 'accuracy',
  'check_in': 'checkin',
  'cleanliness': 'cleanliness',
  'communication': 'communication',
  'location': 'location',
  'value': 'value',
};

const ratingMap = {};

// ページからレビュー★を抽出する関数
async function extractRatingsFromPage(catKey) {
  return page.evaluate((category) => {
    const body = document.body.innerText;
    const results = [];
    const pattern = /Rating\s+(\d)\s+out of\s+5/g;
    let match;
    const indices = [];
    while ((match = pattern.exec(body)) !== null) {
      indices.push({ index: match.index, rating: parseInt(match[1]) });
    }

    for (const { index, rating } of indices) {
      const before = body.substring(Math.max(0, index - 500), index);
      const lines = before.split('\n').map(l => l.trim()).filter(Boolean);

      let guest = '';
      let dateRange = '';
      for (let i = lines.length - 1; i >= 0; i--) {
        if (lines[i].includes('•')) {
          dateRange = lines[i];
          if (i > 0) guest = lines[i - 1];
          break;
        }
      }

      if (guest && guest !== 'View details') {
        results.push({ guest, dateRange: dateRange.substring(0, 80), rating });
      }
    }
    return results;
  }, catKey);
}

for (const [catKey, pathSegment] of Object.entries(categoryPaths)) {
  const catName = catKey.replace(/_/g, ' ');
  process.stdout.write(`  ${catName}...`);

  try {
    const reviewsUrl = `https://www.airbnb.com/performance/quality/${pathSegment}/reviews`;
    await page.goto(reviewsUrl, { waitUntil: 'domcontentloaded', timeout: 20000 });
    await page.waitForTimeout(2000);

    let totalForCat = 0;
    let pageIdx = 1;
    const maxPages = 100;

    while (pageIdx <= maxPages) {
      const ratings = await extractRatingsFromPage(catKey);
      let newCount = 0;
      for (const r of ratings) {
        if (!ratingMap[r.guest]) ratingMap[r.guest] = { dateRange: r.dateRange };
        if (!ratingMap[r.guest][catKey]) {
          ratingMap[r.guest][catKey] = r.rating;
          newCount++;
        }
      }
      totalForCat += newCount;

      // 次のページ番号のリンクをクリック
      const nextPageNum = pageIdx + 1;
      const hasNext = await page.evaluate((nextNum) => {
        const links = [...document.querySelectorAll('a, button')];
        // まず数字ページリンクを探す
        for (const el of links) {
          if (el.innerText && el.innerText.trim() === String(nextNum)) {
            el.click();
            return true;
          }
        }
        // aria-label="Next" も試す
        for (const el of links) {
          const label = (el.getAttribute('aria-label') || '').toLowerCase();
          if (label.includes('next page') || label === 'next') {
            el.click();
            return true;
          }
        }
        return false;
      }, nextPageNum);

      if (!hasNext) break;
      await page.waitForTimeout(1500);
      pageIdx++;
    }

    console.log(` ${totalForCat}件 (${pageIdx}p)`);

  } catch (e) {
    console.log(` エラー: ${e.message.substring(0, 50)}`);
  }
}

console.log(`  ★取得済みゲスト数: ${Object.keys(ratingMap).length}\n`);

// ============================================================
// STEP 3: Superhost + 総合★
// ============================================================
console.log('=== STEP 3: Superhost情報 ===');
let superhostInfo = {};
try {
  await page.click('text="Superhost"', { timeout: 5000 });
  await page.waitForTimeout(3000);

  superhostInfo = await page.evaluate(() => {
    const body = document.body.innerText;

    const isSuperhost = body.includes("You're a Superhost");
    const notSuperhost = body.includes("You didn't get Superhost");

    const ratingMatch = body.match(/([\d.]+)\s+stars?\s[\s\S]*?Overall rating\s[\s\S]*?Criteria:\s*([\d.]+)/);
    const responseMatch = body.match(/([\d.]+)%\s[\s\S]*?Response rate\s[\s\S]*?Criteria:\s*([\d.]+)%/);
    const staysMatch = body.match(/(\d+)\s[\s\S]*?Stays?,\s*(\d+)\s*nights/);
    const cancelMatch = body.match(/([\d.]+)%\s*\n[\s\S]*?Cancellation rate/);

    return {
      isSuperhost: isSuperhost ? true : notSuperhost ? false : null,
      overallRating: ratingMatch ? parseFloat(ratingMatch[1]) : null,
      ratingCriteria: ratingMatch ? parseFloat(ratingMatch[2]) : null,
      responseRate: responseMatch ? parseFloat(responseMatch[1]) : null,
      stays: staysMatch ? parseInt(staysMatch[1]) : null,
      nights: staysMatch ? parseInt(staysMatch[2]) : null,
      cancellationRate: cancelMatch ? parseFloat(cancelMatch[1]) : null,
    };
  });

  console.log(`  Superhost: ${superhostInfo.isSuperhost ? 'Yes' : 'No'}`);
  console.log(`  総合★: ${superhostInfo.overallRating} (基準: ${superhostInfo.ratingCriteria})`);
  console.log(`  レスポンス率: ${superhostInfo.responseRate}%`);
  console.log(`  宿泊数: ${superhostInfo.stays}回 / ${superhostInfo.nights}泊`);
  console.log(`  キャンセル率: ${superhostInfo.cancellationRate}%`);
} catch (e) {
  console.log(`  エラー: ${e.message.substring(0, 60)}`);
}

// ============================================================
// STEP 4: 未投稿レビュー
// ============================================================
console.log('\n=== STEP 4: 未投稿レビュー ===');
let pendingReviews = [];
try {
  await page.goto('https://www.airbnb.com/hosting', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);

  pendingReviews = await page.evaluate(() => {
    return [...document.querySelectorAll('a[href*="/hosting/reviews/"]')].map(a => ({
      text: a.innerText.trim().substring(0, 200),
      href: a.href
    }));
  });
  console.log(`  ${pendingReviews.length} 件`);
} catch (e) {
  console.log(`  エラー: ${e.message.substring(0, 80)}`);
}

// ============================================================
// STEP 4b: 投稿済みレビュー（Reviews by you）
// ============================================================
console.log('\n=== STEP 4b: 投稿済みレビュー（Reviews by you） ===');
await page.goto('https://www.airbnb.com/users/reviews', { waitUntil: 'domcontentloaded', timeout: 30000 });
await page.waitForTimeout(3000);

// Reviews by you タブをクリック
await page.evaluate(() => {
  const btns = [...document.querySelectorAll('button')];
  const tab = btns.find(b => b.innerText.trim() === 'Reviews by you');
  if (tab) tab.click();
});
await page.waitForTimeout(3000);

// スクロールして全件読み込み
prevHeight = 0;
for (let i = 0; i < 20; i++) {
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await page.waitForTimeout(1500);
  const newHeight = await page.evaluate(() => document.body.scrollHeight);
  if (newHeight === prevHeight) break;
  prevHeight = newHeight;
}

const reviewsByYou = await page.evaluate(() => {
  const body = document.body.innerText;

  // 投稿済みレビュー（アポストロフィのバリエーション対応）
  const posted = [];
  const pastMarker = body.match(/Past reviews you.ve written/);
  const pastSection = pastMarker ? body.substring(pastMarker.index + pastMarker[0].length) : null;
  if (pastSection) {
    const expiredSplit = pastSection.includes('Expired Reviews') ? pastSection.split('Expired Reviews')[0] : pastSection;
    const months = 'January|February|March|April|May|June|July|August|September|October|November|December';
    // "SunduibaatarApril 2026\nThank you..." のようなフォーマット
    const entryRegex = new RegExp(`([\\w\\s]+?)((?:${months})\\s+\\d{4})`, 'g');
    const entries = expiredSplit.split('\n').map(l => l.trim()).filter(Boolean);

    let i = 0;
    while (i < entries.length) {
      const line = entries[i];
      const dateMatch = line.match(new RegExp(`(.+?)((?:${months})\\s+\\d{4})$`));
      if (dateMatch) {
        const guestName = dateMatch[1].trim();
        const date = dateMatch[2];
        // 次の行がレビュー本文
        const textLines = [];
        for (let j = i + 1; j < entries.length; j++) {
          if (entries[j].match(new RegExp(`(?:${months})\\s+\\d{4}$`))) break;
          if (entries[j] === 'Expired Reviews') break;
          textLines.push(entries[j]);
        }
        if (guestName) {
          posted.push({ guestName, date, reviewText: textLines.join(' ').substring(0, 2000) });
        }
      }
      i++;
    }
  }

  // 期限切れレビュー
  const expired = [];
  const expiredSection = body.split('Expired Reviews')[1];
  if (expiredSection) {
    const expiredRegex = /your time to write a review for (.+?) has expired/gi;
    let em;
    while ((em = expiredRegex.exec(expiredSection)) !== null) {
      expired.push({ guestName: em[1] });
    }
  }

  return { posted, expired };
});

console.log(`  投稿済み: ${reviewsByYou.posted.length} 件`);
reviewsByYou.posted.forEach(r => console.log(`    - ${r.guestName} (${r.date})`));
console.log(`  期限切れ: ${reviewsByYou.expired.length} 件`);

// ============================================================
// STEP 4c: 未返信レビュー一覧（Leave Public Response があるもの）
// ============================================================
console.log('\n=== STEP 4c: 未返信レビュー一覧 ===');
// Reviews about you タブに戻る
await page.evaluate(() => {
  const btns = [...document.querySelectorAll('button')];
  const tab = btns.find(b => b.innerText.trim() === 'Reviews about you');
  if (tab) tab.click();
});
await page.waitForTimeout(3000);

// スクロール
prevHeight = 0;
for (let i = 0; i < 20; i++) {
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await page.waitForTimeout(1500);
  const newHeight = await page.evaluate(() => document.body.scrollHeight);
  if (newHeight === prevHeight) break;
  prevHeight = newHeight;
}

const unrepliedReviews = await page.evaluate(() => {
  const results = [];
  const buttons = [...document.querySelectorAll('button')].filter(el =>
    el.innerText.trim() === 'Leave Public Response'
  );

  for (const btn of buttons) {
    let container = btn;
    for (let i = 0; i < 10; i++) {
      container = container.parentElement;
      if (!container) break;
      const text = container.innerText || '';
      if (text.includes('Leave Public Response') && text.length > 50 && text.length < 3000) {
        // ゲスト名: プロフィールリンクから取得（言語に依存しない）
        const links = [...container.querySelectorAll('a')];
        const profileLink = links.find(a => {
          const href = a.href || '';
          return href.includes('/users/profile/') || href.includes('/users/show/');
        });
        let guestName = profileLink ? profileLink.innerText.trim() : '';

        // フォールバック: テキストパース
        const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
        const months = 'January|February|March|April|May|June|July|August|September|October|November|December';
        const dateRegex = new RegExp(`(${months})\\s+\\d{4}`);

        if (!guestName) {
          for (let j = 0; j < lines.length; j++) {
            const match = lines[j].match(dateRegex);
            if (match) {
              const beforeDate = lines[j].substring(0, lines[j].indexOf(match[0])).trim();
              guestName = beforeDate || (j > 0 ? lines[j - 1] : '');
              break;
            }
          }
        }

        let reviewDate = '';
        let reviewText = '';
        for (let j = 0; j < lines.length; j++) {
          const match = lines[j].match(dateRegex);
          if (match) {
            reviewDate = match[0];
            const textLines = [];
            for (let k = j + 1; k < lines.length; k++) {
              if (['Leave Public Response', 'Private feedback', 'Read more'].includes(lines[k])) break;
              textLines.push(lines[k]);
            }
            reviewText = textLines.join(' ');
            break;
          }
        }

        if (guestName) {
          results.push({
            guestName,
            date: reviewDate,
            reviewText: reviewText.substring(0, 2000),
          });
        }
        break;
      }
    }
  }
  return results;
});

console.log(`  未返信: ${unrepliedReviews.length} 件`);
unrepliedReviews.forEach(r => console.log(`    - ${r.guestName} (${r.date}): ${r.reviewText.substring(0, 60)}...`));

// ============================================================
// STEP 5: レビュー本文と★スコアをマージ
// ============================================================
console.log('\n=== STEP 5: データマージ ===');
let matched = 0;
for (const review of allReviews) {
  const ratings = ratingMap[review.guestName];
  if (ratings) {
    review.overall_quality = ratings.overall_quality || null;
    review.accuracy = ratings.accuracy || null;
    review.checkin = ratings['check-in'] || ratings.check_in || null;
    review.cleanliness = ratings.cleanliness || null;
    review.communication = ratings.communication || null;
    review.location = ratings.location || null;
    review.value = ratings.value || null;
    matched++;
  }
}
console.log(`  ★マッチ: ${matched}/${allReviews.length} 件`);

await browser.close();

// ============================================================
// STEP 6: スプシ書き込み
// ============================================================
const fetchedAt = new Date().toISOString();
console.log('\n=== STEP 6: スプシ書き込み ===');
try {
  const res = await fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action: 'writeReviews',
      account: accountName,
      fetchedAt,
      reviews: allReviews,
      superhostInfo,
      reviewsByYou,
      unrepliedReviews,
    }),
    redirect: 'follow'
  });
  const result = await res.json();
  if (result.success) {
    console.log(`  書き込み: ${result.written} 件, スキップ: ${result.skipped} 件`);
  } else {
    console.error('  エラー:', result.error);
  }
} catch (err) {
  console.error('  送信エラー:', err.message);
}

// JSON保存
const output = {
  account: accountName,
  fetchedAt,
  totalReviews: allReviews.length,
  reviews: allReviews,
  pendingReviews,
  reviewsByYou,
  unrepliedReviews,
  superhostInfo,
};
writeFileSync(`output-${accountName}.json`, JSON.stringify(output, null, 2));

// サマリー表示
console.log('\n========== サマリー ==========');
console.log(`アカウント: ${accountName}`);
console.log(`受信レビュー: ${allReviews.length} 件`);
console.log(`★スコアマッチ: ${matched}/${allReviews.length} 件`);
console.log(`未投稿レビュー: ${pendingReviews.length} 件`);
console.log(`投稿済みレビュー: ${reviewsByYou.posted.length} 件`);
console.log(`期限切れレビュー: ${reviewsByYou.expired.length} 件`);
console.log(`未返信レビュー: ${unrepliedReviews.length} 件`);
console.log(`Superhost: ${superhostInfo.isSuperhost ? 'Yes' : 'No'} (★${superhostInfo.overallRating})`);
console.log(`保存: output-${accountName}.json`);
console.log('完了');
