#!/usr/bin/env node
// 生成したデッキが「実際にブラウザで開いて読める」ことを確かめる。
// 数字は暗号化されているので、復号→描画まで通して初めて壊れていないと言える。
// stdout に JSON を1行出す（build-ah.mjs が読む）。
import { chromium } from '../../review/node_modules/playwright/index.mjs';

const file = process.argv[2];
const pw = process.env.DECK_PASSWORD;
const errors = [];
let result = { ok: false, errors: ['未実行'] };

try {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1440, height: 900 } });
  p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  p.on('pageerror', e => errors.push('PAGEERROR ' + e.message));
  await p.goto('file://' + file);
  await p.fill('#login-pw', pw);
  await p.click('#login-overlay button');
  await p.waitForSelector('#login-overlay.hidden', { state: 'attached', timeout: 60000 });
  await p.waitForTimeout(3000);

  const r = await p.evaluate(() => ({
    sales: document.getElementById('kpi-sales')?.textContent?.trim(),
    occ: document.getElementById('kpi-occ')?.textContent?.trim(),
    charts: [...document.querySelectorAll('canvas')].filter(cv => window.Chart?.getChart(cv)).length,
    canvases: document.querySelectorAll('canvas').length,
    insights: document.querySelectorAll('.insight-item').length,
    areaRows: document.querySelectorAll('#areaTable tbody tr').length,
    jntoRows: document.querySelectorAll('#jntoCountryTable tbody tr').length,
  }));
  await b.close();

  if (!r.sales || r.sales === '—') errors.push('KPI が描画されていません（復号に失敗した可能性）');
  if (r.charts < r.canvases) errors.push(`グラフが ${r.canvases} 個中 ${r.charts} 個しか生成されていません`);
  if (r.insights < 4) errors.push(`インサイトが ${r.insights} 件しかありません`);
  if (r.areaRows < 1) errors.push('エリア別テーブルが空です');
  if (r.jntoRows < 5) errors.push(`JNTO 国別テーブルが ${r.jntoRows} 行しかありません`);
  result = { ok: errors.length === 0, errors, ...r };
} catch (e) {
  result = { ok: false, errors: [...errors, e.message] };
}
console.log(JSON.stringify(result));
