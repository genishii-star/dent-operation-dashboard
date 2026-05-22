import { chromium } from 'playwright';
const browser = await chromium.launch({ headless: false });
const context = await browser.newContext({ storageState: 'sessions/OPH.json' });
const page = await context.newPage();

// レビュー一覧ページに直接遷移
await page.goto('https://www.airbnb.com/performance/quality/overall/reviews', { waitUntil: 'domcontentloaded', timeout: 30000 });
await page.waitForTimeout(5000);

console.log('URL:', page.url());

// ページネーション要素を調査
const navInfo = await page.evaluate(() => {
  const results = [];
  const els = document.querySelectorAll('a, button');
  for (const el of els) {
    if (!el.innerText) continue;
    const text = (el.innerText || '').trim();
    const ariaLabel = el.getAttribute('aria-label') || '';
    // ページネーション関連
    if (/^\d+$/.test(text) || ariaLabel.toLowerCase().includes('page') ||
        ariaLabel.toLowerCase().includes('next') || ariaLabel.toLowerCase().includes('previous') ||
        text === '…' || text === 'Next' || text === 'Previous') {
      results.push({
        tag: el.tagName,
        text: text.substring(0, 50),
        ariaLabel,
        href: el.href || ''
      });
    }
  }
  return results;
});

console.log('\nページネーション要素:');
navInfo.forEach(n => console.log(`  [${n.tag}] text="${n.text}" aria="${n.ariaLabel}" href="${n.href}"`));

// "Showing X of Y" テキスト
const showingText = await page.evaluate(() => {
  const body = document.body.innerText;
  const match = body.match(/Showing\s+\d+\s+of\s+\d+/i);
  return match ? match[0] : 'not found';
});
console.log('\nShowing:', showingText);

await page.waitForTimeout(10000);
await browser.close();
