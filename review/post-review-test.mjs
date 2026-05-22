/**
 * 系統A検証: レビュー投稿フォーム全ステップ調査（ドライラン）
 * Submitは押さない。
 */

import { chromium } from 'playwright';
import { existsSync } from 'fs';

const accountName = process.argv[2] || 'OPH';
const sessionPath = `sessions/${accountName}.json`;
const headless = !process.argv.includes('--gui');

if (!existsSync(sessionPath)) {
  console.error(`セッションが見つかりません: ${sessionPath}`);
  process.exit(1);
}

const browser = await chromium.launch({ headless });
const context = await browser.newContext({ storageState: sessionPath });
const page = await context.newPage();

const reviewUrl = 'https://www.airbnb.com/hosting/reviews/1659206077832294551/edit?exit_url=/hosting';
console.log('=== レビュー投稿フォーム 全ステップ調査 ===\n');

await page.goto(reviewUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
await page.waitForTimeout(3000);

for (let step = 0; step < 10; step++) {
  const bodyText = await page.evaluate(() => document.body.innerText);
  const progressMatch = bodyText.match(/step (\d+) out of (\d+)/);
  const currentStep = progressMatch ? progressMatch[1] : '?';
  const totalSteps = progressMatch ? progressMatch[2] : '?';

  console.log(`\n========== STEP ${currentStep}/${totalSteps} ==========`);

  // コンテンツ表示
  const lines = bodyText.split('\n').map(l => l.trim()).filter(Boolean);
  const startIdx = lines.findIndex(l => l.includes('Save and exit'));
  const content = lines.slice(startIdx + 1).join('\n');
  console.log(content.substring(0, 500));

  // フォーム要素調査
  const formInfo = await page.evaluate(() => {
    const info = { radios: [], textareas: [], checkboxes: [] };

    // radio inputs
    document.querySelectorAll('input[type="radio"]').forEach(el => {
      info.radios.push({
        id: el.id,
        name: el.name,
        value: el.value,
        label: el.getAttribute('aria-label') || '',
        checked: el.checked
      });
    });

    // textareas
    document.querySelectorAll('textarea').forEach(el => {
      info.textareas.push({
        id: el.id || '',
        name: el.name || '',
        placeholder: el.placeholder || '',
        maxLength: el.maxLength,
        label: el.getAttribute('aria-label') || ''
      });
    });

    // checkboxes
    document.querySelectorAll('input[type="checkbox"]').forEach(el => {
      info.checkboxes.push({
        id: el.id || '',
        label: el.getAttribute('aria-label') || el.closest('label')?.innerText?.trim()?.substring(0, 80) || ''
      });
    });

    // buttons
    info.buttons = [...document.querySelectorAll('button')]
      .filter(el => el.innerText && el.offsetParent !== null)
      .map(el => ({ text: (el.innerText || '').trim().substring(0, 50), disabled: el.disabled }));

    return info;
  });

  if (formInfo.radios.length > 0) {
    console.log(`\nRadio (${formInfo.radios.length}個): ${formInfo.radios.map(r => `${r.value}[${r.name.split('_').pop()}]`).join(', ')}`);
  }
  if (formInfo.textareas.length > 0) {
    console.log(`\nTextarea: maxLength=${formInfo.textareas[0].maxLength}, label="${formInfo.textareas[0].label}"`);
  }
  if (formInfo.checkboxes.length > 0) {
    console.log(`\nCheckbox: ${formInfo.checkboxes.map(c => c.label).join(', ')}`);
  }
  console.log(`Buttons: ${formInfo.buttons.map(b => `${b.text}${b.disabled ? '(disabled)' : ''}`).join(' | ')}`);

  // Submitステップなら停止
  if (formInfo.buttons.some(b => b.text === 'Submit')) {
    console.log('\n*** Submit ステップに到達。投稿しません。 ***');
    break;
  }

  // フォーム入力
  // ★5を選ぶ（value=5のradio）
  const radio5 = formInfo.radios.find(r => r.value === '5');
  if (radio5) {
    await page.evaluate((id) => document.getElementById(id).click(), radio5.id);
    await page.waitForTimeout(500);
    console.log(`→ ★5選択: ${radio5.id}`);
  }

  // textarea に仮テキスト
  if (formInfo.textareas.length > 0) {
    await page.fill('textarea', 'Thank you for your stay! We hope you had a wonderful time.');
    await page.waitForTimeout(500);
    console.log('→ テキスト入力済み');
  }

  // Next/Continue をクリック
  const nextBtn = formInfo.buttons.find(b => (b.text === 'Next' || b.text === 'Continue') && !b.disabled);
  if (!nextBtn) {
    // disabledの場合、500ms待ってリトライ
    await page.waitForTimeout(1000);
    const retryBtn = await page.evaluate(() => {
      const btn = [...document.querySelectorAll('button')].find(b => b.innerText.trim() === 'Next' || b.innerText.trim() === 'Continue');
      return btn ? { text: btn.innerText.trim(), disabled: btn.disabled } : null;
    });
    if (!retryBtn || retryBtn.disabled) {
      console.log('\nNext/Continueがdisabledのまま。停止。');
      break;
    }
  }

  const btnText = nextBtn?.text || 'Next';
  await page.click(`button:has-text("${btnText}")`, { timeout: 5000 });
  await page.waitForTimeout(2000);
}

await page.screenshot({ path: 'debug-review-form.png', fullPage: true });
console.log('\nスクリーンショット: debug-review-form.png');

await browser.close();
console.log('\n完了');
