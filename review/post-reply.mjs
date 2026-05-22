/**
 * ゲストレビューへの公開返信（Public Response）を投稿
 *
 * 使い方:
 *   node post-reply.mjs <account-name>              # ドライラン（返信内容を表示するだけ）
 *   node post-reply.mjs <account-name> --post        # 実際に投稿
 *   node post-reply.mjs <account-name> --gui          # GUI表示
 *   node post-reply.mjs <account-name> --post --gui   # GUI + 投稿
 *
 * 返信テキスト:
 *   replies-{account}.json に以下の形式で用意:
 *   { "ゲスト名": "返信テキスト", ... }
 *   ファイルがなければ未返信一覧を表示して終了。
 */

import { chromium } from 'playwright';
import { existsSync, readFileSync } from 'fs';

const accountName = process.argv[2] || 'OPH';
const sessionPath = `sessions/${accountName}.json`;
const headless = !process.argv.includes('--gui');
const doPost = process.argv.includes('--post');
const repliesPath = `replies-${accountName}.json`;

if (!existsSync(sessionPath)) {
  console.error(`セッションが見つかりません: ${sessionPath}`);
  process.exit(1);
}

// 返信テキスト読み込み
let replies = {};
if (existsSync(repliesPath)) {
  replies = JSON.parse(readFileSync(repliesPath, 'utf-8'));
  console.log(`返信データ読み込み: ${repliesPath} (${Object.keys(replies).length} 件)\n`);
} else {
  console.log(`返信データなし: ${repliesPath}`);
  console.log('未返信一覧を取得します...\n');
}

const browser = await chromium.launch({ headless });
const context = await browser.newContext({ storageState: sessionPath });
const page = await context.newPage();

// ページ読み込み + スクロール関数
async function loadReviewsPage() {
  await page.goto('https://www.airbnb.com/users/reviews', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);
  let prevHeight = 0;
  for (let i = 0; i < 20; i++) {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(1500);
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === prevHeight) break;
    prevHeight = newHeight;
  }
}

// 未返信レビュー抽出関数（ゲスト名をリンクのテキストから正確に取得）
async function getUnrepliedReviews() {
  return page.evaluate(() => {
    const results = [];
    const buttons = [...document.querySelectorAll('button')].filter(b => b.innerText.trim() === 'Leave Public Response');

    for (const btn of buttons) {
      // ボタンから親をたどってレビューコンテナを特定
      let container = btn;
      for (let i = 0; i < 10; i++) {
        container = container.parentElement;
        if (!container) break;
        const text = container.innerText || '';
        if (text.includes('Leave Public Response') && text.length > 50 && text.length < 1000) {
          // ゲスト名: コンテナ内のリンク（プロフィールリンク）から取得
          const links = [...container.querySelectorAll('a')];
          const profileLink = links.find(a => {
            const href = a.href || '';
            return href.includes('/users/profile/') || href.includes('/users/show/');
          });
          let guestName = profileLink ? profileLink.innerText.trim() : '';

          // フォールバック: テキストパースで取得
          if (!guestName) {
            const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
            const months = 'January|February|March|April|May|June|July|August|September|October|November|December';
            const dateRegex = new RegExp(`(${months})\\s+\\d{4}`);
            for (let j = 0; j < lines.length; j++) {
              const match = lines[j].match(dateRegex);
              if (match) {
                const beforeDate = lines[j].substring(0, lines[j].indexOf(match[0])).trim();
                guestName = beforeDate || (j > 0 ? lines[j - 1] : '');
                break;
              }
            }
          }

          // レビュー日付・本文
          const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
          const months = 'January|February|March|April|May|June|July|August|September|October|November|December';
          const dateRegex = new RegExp(`(${months})\\s+\\d{4}`);
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
            results.push({ guestName, date: reviewDate, reviewText: reviewText.substring(0, 500) });
          }
          break;
        }
      }
    }
    return results;
  });
}

console.log('=== レビュー返信 ===\n');
await loadReviewsPage();

const unrepliedDetails = await getUnrepliedReviews();
console.log(`未返信レビュー: ${unrepliedDetails.length} 件\n`);

if (unrepliedDetails.length === 0) {
  console.log('全件返信済み。');
  await browser.close();
  process.exit(0);
}

// 返信データがない場合: 一覧表示して終了
if (Object.keys(replies).length === 0) {
  console.log('--- 未返信レビュー一覧 ---');
  unrepliedDetails.forEach((r, i) => {
    console.log(`\n[${i + 1}] ${r.guestName} (${r.date})`);
    console.log(`    ${r.reviewText.substring(0, 100)}...`);
  });
  console.log(`\n${repliesPath} を作成して返信テキストを設定してください。`);
  console.log('例: { "Sunduibaatar": "Thank you for your kind words!", ... }');
  await browser.close();
  process.exit(0);
}

// 返信を投稿（1件ずつページリロードして安全に投稿）
let posted = 0;
let skipped = 0;
const errors = [];

for (const detail of unrepliedDetails) {
  const replyText = replies[detail.guestName];
  if (!replyText) {
    console.log(`[SKIP] ${detail.guestName} — 返信テキストなし`);
    skipped++;
    continue;
  }

  console.log(`\n[${detail.guestName}] (${detail.date})`);
  console.log(`  ゲスト: ${detail.reviewText.substring(0, 80)}...`);
  console.log(`  返信: ${replyText.substring(0, 80)}...`);

  if (!doPost) {
    console.log('  → ドライラン（--post で実投稿）');
    continue;
  }

  // === 安全策1: 毎回ページをリロードして最新DOMで操作 ===
  await loadReviewsPage();

  // === 安全策1b: 開いてるtextareaを全て閉じる（Cancelクリック） ===
  await page.evaluate(() => {
    const cancelBtns = [...document.querySelectorAll('button')].filter(b => b.innerText.trim() === 'Cancel');
    cancelBtns.forEach(b => b.click());
  });
  await page.waitForTimeout(1000);

  // === 安全策2: ゲスト名でボタンを特定し、クリック ===
  // コンテナサイズ制限付きで最小のマッチを探す（巨大な親コンテナを避ける）
  const clicked = await page.evaluate((guestName) => {
    const buttons = [...document.querySelectorAll('button')].filter(b => b.innerText.trim() === 'Leave Public Response');
    let bestBtn = null;
    let bestSize = Infinity;

    for (const btn of buttons) {
      let container = btn;
      for (let i = 0; i < 10; i++) {
        container = container.parentElement;
        if (!container) break;
        const textLen = (container.innerText || '').length;
        // コンテナが大きすぎる場合スキップ（複数レビューを含む親コンテナ）
        if (textLen > 1000) break;
        // コンテナ内のプロフィールリンクでゲスト名を確認
        const links = [...container.querySelectorAll('a')];
        const profileLink = links.find(a => {
          const href = a.href || '';
          return (href.includes('/users/profile/') || href.includes('/users/show/')) && a.innerText.trim() === guestName;
        });
        if (profileLink && container.innerText.includes('Leave Public Response')) {
          if (textLen < bestSize) {
            bestBtn = btn;
            bestSize = textLen;
          }
          break;
        }
      }
    }

    if (bestBtn) {
      bestBtn.scrollIntoView({ block: 'center' });
      bestBtn.click();
      return true;
    }
    return false;
  }, detail.guestName);

  if (!clicked) {
    console.log('  → ボタンが見つからない。スキップ。');
    skipped++;
    continue;
  }

  await page.waitForTimeout(2000);

  // === 安全策3: textareaのaria-label/placeholderにゲスト名が含まれるか検証 ===
  const textareaInfo = await page.evaluate((guestName) => {
    const textareas = [...document.querySelectorAll('textarea')];
    for (const ta of textareas) {
      const label = ta.getAttribute('aria-label') || ta.placeholder || '';
      if (label.toLowerCase().includes(guestName.toLowerCase())) {
        return { id: ta.id || null, verified: true, label };
      }
    }
    // ゲスト名がlabelに含まれないtextareaがあれば、検証失敗として返す
    if (textareas.length > 0) {
      const ta = textareas[0];
      const label = ta.getAttribute('aria-label') || ta.placeholder || '';
      return { id: ta.id || null, verified: false, label };
    }
    return null;
  }, detail.guestName);

  if (!textareaInfo) {
    console.log('  → textarea が見つからない。スキップ。');
    skipped++;
    continue;
  }

  if (!textareaInfo.verified) {
    console.log(`  → *** 安全チェック失敗: textareaのラベルにゲスト名が含まれない ***`);
    console.log(`     期待: "${detail.guestName}", 実際: "${textareaInfo.label}"`);
    console.log('  → 投稿中止。');
    errors.push({ guestName: detail.guestName, reason: 'textarea label mismatch', label: textareaInfo.label });
    continue;
  }

  // textarea に入力（IDが数字始まりでCSS selectorに使えないのでevaluateで入力）
  await page.evaluate(({ id, text }) => {
    const ta = id ? document.getElementById(id) : document.querySelector('textarea');
    if (ta) {
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      nativeSetter.call(ta, text);
      ta.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }, { id: textareaInfo.id, text: replyText });
  await page.waitForTimeout(1000);

  // Submit クリック（開いているフォームのSubmitのみ）
  await page.evaluate(() => {
    const btns = [...document.querySelectorAll('button')].filter(b => b.innerText.trim() === 'Submit');
    if (btns.length > 0) btns[0].click();
  });
  await page.waitForTimeout(3000);

  console.log('  → 投稿完了');
  posted++;
}

// サマリー
console.log('\n========== サマリー ==========');
console.log(`アカウント: ${accountName}`);
console.log(`未返信: ${unrepliedDetails.length} 件`);
console.log(`投稿: ${posted} 件`);
console.log(`スキップ: ${skipped} 件`);
if (errors.length > 0) {
  console.log(`エラー: ${errors.length} 件`);
  errors.forEach(e => console.log(`  - ${e.guestName}: ${e.reason}`));
}
console.log(`モード: ${doPost ? '実投稿' : 'ドライラン'}`);

await browser.close();
console.log('完了');
