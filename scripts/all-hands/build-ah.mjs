#!/usr/bin/env node
//
// All Hands デッキ（op.dent-inc.com/all-hands/YYYY-MM.html）を月次で自動生成する。
//
//   node scripts/all-hands/build-ah.mjs            # 前月分を作る（毎月20日の launchd 実行がこれ）
//   node scripts/all-hands/build-ah.mjs 2026-07    # 実績月を明示（作り直し・過去分の再生成）
//   node scripts/all-hands/build-ah.mjs --dry-run  # 生成と検証だけして push も Slack も送らない
//
// ファイル名の YYYY-MM は「開催月」で、中身は前月実績。2026-08.html = 8月会 = 7月まとめ。
//
// 認証・秘密は ~/.dent/ah.env から読む（このリポジトリは public なので絶対に埋め込まないこと）:
//   DECK_PASSWORD=...      デッキの復号パスワード（ダッシュボードと同じもの）
//   NOTIFY_URL=...         dent-slack-bot の /api/notify
//   NOTIFY_TOKEN=...       同 Bearer トークン
// D1 は wrangler のローカル OAuth、push は gh の認証をそのまま使う（追加の資格情報は不要）。
import fs from 'fs';
import os from 'os';
import path from 'path';
import { execFileSync } from 'child_process';
import { encrypt } from './lib/crypt.mjs';
import { aggregate, getDaysInMonth } from './lib/aggregate.mjs';

const ROOT = path.resolve(new URL('../..', import.meta.url).pathname);      // operation/
const AH = path.join(ROOT, 'all-hands');
const WIKI = path.resolve(ROOT, '../wiki');
const API_DIR = path.join(WIKI, 'workers/dent-data-api');
const HERE = path.dirname(new URL(import.meta.url).pathname);
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'ah-build-'));

const argv = process.argv.slice(2);
const DRY = argv.includes('--dry-run');
const target = argv.find(a => /^\d{4}-\d{2}$/.test(a));

const log = (...a) => console.log('[ah]', ...a);
const sh = (cmd, args, opts = {}) =>
  execFileSync(cmd, args, { encoding: 'utf8', maxBuffer: 256 * 1024 * 1024, ...opts });

// ---------- 設定 ----------
function loadEnv() {
  const f = path.join(os.homedir(), '.dent/ah.env');
  if (!fs.existsSync(f)) throw new Error(`設定ファイルがありません: ${f}（README 参照）`);
  const env = {};
  fs.readFileSync(f, 'utf8').split('\n').forEach(line => {
    const m = line.match(/^\s*([A-Z_]+)\s*=\s*(.*)$/);
    if (m) env[m[1]] = m[2].trim();
  });
  if (!env.DECK_PASSWORD) throw new Error('ah.env に DECK_PASSWORD がありません');
  return env;
}

// ---------- 月の計算 ----------
const shift = (ym, n) => {
  let [y, m] = ym.split('-').map(Number);
  m += n;
  while (m < 1) { m += 12; y--; }
  while (m > 12) { m -= 12; y++; }
  return `${y}-${String(m).padStart(2, '0')}`;
};
const now = new Date();
const CUR = target || shift(`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`, -1);
const PRV = shift(CUR, -1);
const YA = shift(CUR, -12);
const MEET = shift(CUR, 1);          // 開催月＝ファイル名
const [curY, curM] = CUR.split('-').map(Number);

// ---------- D1 エクスポート ----------
function d1(sql) {
  const out = sh('npx', ['wrangler', 'd1', 'execute', 'dent-platform-d1', '--remote', '--json', '--command', sql],
    { cwd: API_DIR });
  const json = JSON.parse(out.slice(out.indexOf('[')));
  return json[0].results;
}

async function main() {
  const env = loadEnv();
  const warnings = [];
  log(`実績月 ${CUR} / 前月 ${PRV} / 前年 ${YA} → ${MEET}.html`);

  // 物件マスタ。wiki の作業ツリーは他人のブランチにいることがあるので必ず origin/main から取る。
  log('物件マスタを取得中…');
  sh('git', ['fetch', '-q', 'origin', 'main'], { cwd: WIKI });
  const facDir = path.join(TMP, 'facilities');
  fs.mkdirSync(facDir, { recursive: true });
  sh('bash', ['-c', `git archive origin/main site/data/facilities | tar -xf - -C ${JSON.stringify(facDir)} --strip-components=3`], { cwd: WIKI });

  log('D1 から予約・日次を取得中…');
  const reservationRows = d1(
    `SELECT property_name,room_no,check_in,nights,status,guests,nationality,channel,gross_sales,net_received ` +
    `FROM reservations WHERE check_in>='${YA}-01' AND check_in<'${shift(CUR, 1)}-01'`);
  const dailyRows = d1(
    `SELECT property_name,room_no,date,status,channel,gross_sales,net_received,cleaning_fee ` +
    `FROM daily_revenue WHERE (date>='${YA}-01' AND date<'${shift(YA, 1)}-01') ` +
    `OR (date>='${PRV}-01' AND date<'${shift(CUR, 1)}-01')`);
  log(`予約 ${reservationRows.length} 行 / 日次 ${dailyRows.length} 行`);

  const agg = aggregate({
    facilitiesDir: facDir, yamlModulePath: path.join(API_DIR, '/'),
    reservationRows, dailyRows, curYm: CUR, prvYm: PRV, yaYm: YA,
  });
  warnings.push(...agg.warnings);

  // 出来上がりが空・桁違いなら止める（母集団が変わったときに黙って壊れたデッキを配らないため）
  const c = agg.data.cur;
  if (!(c.sales > 0) || !(c.nights > 0)) throw new Error(`${CUR} の集計が空です（sales=${c.sales}, nights=${c.nights}）`);
  if (agg.activeRooms < 50) throw new Error(`稼働物件が ${agg.activeRooms} 室しかありません（マスタ取得の失敗を疑う）`);
  const drift = Math.abs(agg.data.prv.sales / (c.sales || 1) - 1);
  if (drift > 3) warnings.push(`前月比が ${(drift * 100).toFixed(0)}% と極端です。データ欠損の可能性`);

  // ---------- JNTO ----------
  log('JNTO 訪日外客統計を取得中…');
  const page = await (await fetch('https://www.jnto.go.jp/statistics/data/visitors-statistics/')).text();
  // ルート相対パス。日付プレフィックスだけ毎月変わり、コード 1615-5 は固定。
  const m = page.match(/\/statistics\/data\/_files\/(\d{8})_1615-5\.xlsx/);
  if (!m) throw new Error('JNTO: 月次Excelのリンクを見つけられませんでした（ページ構造の変更を疑う）');
  const xlsx = path.join(TMP, 'jnto.xlsx');
  fs.writeFileSync(xlsx, Buffer.from(await (await fetch('https://www.jnto.go.jp' + m[0])).arrayBuffer()));
  const jnto = JSON.parse(sh('python3', [path.join(HERE, 'jnto_parse.py'), xlsx, String(curY), String(curY - 1)]));
  const pub = `${m[1].slice(0, 4)}年${+m[1].slice(4, 6)}月${+m[1].slice(6, 8)}日`;
  if (jnto.missingCountries?.length) warnings.push(`JNTO 国別で取れなかった国: ${jnto.missingCountries.join(', ')}`);
  if (jnto.latestMonth < curM) {
    warnings.push(`JNTO は ${curY}年${jnto.latestMonth}月分が最新で、${curM}月分は未発表です。発表後に再実行すると業界スライドが最新化されます`);
  }
  log(`JNTO ${pub}発表・最新 ${curY}年${jnto.latestMonth}月`);

  // ---------- レンダリング ----------
  const payload = {
    data: agg.data,
    extra: {
      ...agg.extra,
      jnto: { months: jnto.months, cur: jnto.cur, prev: jnto.prev },
      jntoCountries: { latest: jnto.countriesCur, prevYear: jnto.countriesPrev },
    },
  };
  // 大阪・関西万博（2025/4〜10）の特需は前年比を大きく歪めるので、該当する月だけ注記する
  const [yaY, yaM] = YA.split('-').map(Number);
  const expo = yaY === 2025 && yaM >= 4 && yaM <= 10;
  const tokens = {
    MEET_YM_SLASH: MEET.replace('-', '/'), MEET_Y: MEET.split('-')[0], MEET_M: +MEET.split('-')[1],
    CUR_Y: curY, CUR_M: curM, CUR_YM: CUR, CUR_YM_SLASH: CUR.replace('-', '/'),
    PRV_Y: PRV.split('-')[0], PRV_M: +PRV.split('-')[1], YA_Y: yaY,
    ACTIVE_ROOMS: agg.activeRooms, MEETING_KEY: MEET,
    EXPO_NOTE: expo ? `${yaY}年${yaM}月は大阪・関西万博の会期中。` : '',
    EXPO_YEARAGO: String(expo),
    JNTO_PUB: pub, JNTO_Y: curY, JNTO_M: jnto.latestMonth,
    BLOB: encrypt(payload, env.DECK_PASSWORD),
  };
  let html = fs.readFileSync(path.join(AH, '_template.html'), 'utf8');
  for (const [k, v] of Object.entries(tokens)) html = html.split(`{{${k}}}`).join(String(v));
  const leftover = html.match(/\{\{[A-Z_]+\}\}/g);
  if (leftover) throw new Error(`テンプレートの未置換トークン: ${[...new Set(leftover)].join(', ')}`);
  const outFile = path.join(AH, `${MEET}.html`);
  const isNew = !fs.existsSync(outFile);
  fs.writeFileSync(outFile, html);
  log(`${MEET}.html を${isNew ? '作成' : '更新'}`);

  // 一覧に追加（再実行時は既にあるので触らない）
  const idxFile = path.join(AH, 'index.html');
  let idx = fs.readFileSync(idxFile, 'utf8');
  if (!idx.includes(`href="${MEET}.html"`)) {
    const anchor = '  <div class="meeting-list">\n';
    if (!idx.includes(anchor)) throw new Error('index.html に meeting-list が見つかりません');
    const card = `\n    <a href="${MEET}.html" class="meeting-card">\n` +
      `      <div class="date"><span class="y">${MEET.split('-')[0]}</span><span class="m">${+MEET.split('-')[1]}月</span></div>\n` +
      `      <div class="info">\n        <div class="title">${curY}年${curM}月まとめ</div>\n` +
      `        <div class="desc">KPI / チャネル / 国籍 / エリア別 / JNTO業界アップデート</div>\n` +
      `      </div>\n      <div class="arrow">›</div>\n    </a>\n`;
    fs.writeFileSync(idxFile, idx.replace(anchor, anchor + card));
    log('index.html に追加');
  }

  // ---------- 検証（復号・描画が本当に通るか） ----------
  log('ヘッドレスで検証中…');
  const verified = JSON.parse(sh('node', [path.join(HERE, 'verify.mjs'), outFile], { env: { ...process.env, DECK_PASSWORD: env.DECK_PASSWORD } }));
  if (!verified.ok) throw new Error(`検証に失敗: ${verified.errors.join(' / ')}`);
  log(`検証OK（グラフ ${verified.charts}本 / インサイト ${verified.insights}件）`);

  // ---------- 公開 ----------
  const summary = [
    `${curY}年${curM}月まとめ（${MEET.split('-')[0]}年${+MEET.split('-')[1]}月会）`,
    `売上 ${(c.sales / 10000).toFixed(0)}万円（前月比 ${fmtPct(c.sales, agg.data.prv.sales)}・前年比 ${fmtPct(c.sales, agg.data.yearAgo.sales)}）`,
    `稼働率 ${(c.occ * 100).toFixed(1)}% / ADR ${Math.round(c.adr).toLocaleString()}円 / 予約 ${c.count}組 / 稼働 ${agg.activeRooms}室`,
  ];
  if (DRY) {
    log('--dry-run のため push と Slack 通知はしません');
    log(summary.join('\n'));
    if (warnings.length) log('警告:\n- ' + warnings.join('\n- '));
    return;
  }

  sh('git', ['add', path.relative(ROOT, outFile), path.relative(ROOT, idxFile)], { cwd: ROOT });
  const staged = sh('git', ['diff', '--cached', '--name-only'], { cwd: ROOT }).trim();
  if (staged) {
    sh('git', ['commit', '-q', '-m',
      `feat(all-hands): ${MEET.split('-')[0]}年${+MEET.split('-')[1]}月会（${curM}月まとめ）を自動生成\n\n` +
      `build-ah.mjs による定期生成。D1 を現行スコープで全月再集計、JNTO は ${pub} 発表版。`], { cwd: ROOT });
    sh('git', ['push', '-q', 'origin', 'HEAD'], { cwd: ROOT });
    log('push 完了');
  } else {
    log('変更なし（内容が同一）');
  }

  const url = `https://op.dent-inc.com/all-hands/${MEET}.html`;
  const text = `:bar_chart: *All Hands 資料ができました*\n${summary.join('\n')}\n${url}\n` +
    (warnings.length ? `\n:warning: 要確認\n• ${warnings.join('\n• ')}` : '') +
    `\n_インサイト文言は数値から自動生成しています。読み合わせ前に一度目を通してください。_`;
  await notify(env, text);
  log('Slack へ通知しました');
}

function fmtPct(a, b) { const v = (a / b - 1) * 100; return (v >= 0 ? '+' : '') + v.toFixed(1) + '%'; }

async function notify(env, text) {
  if (!env.NOTIFY_URL || !env.NOTIFY_TOKEN) { log('NOTIFY_URL/TOKEN 未設定のため Slack 通知をスキップ'); return; }
  const res = await fetch(env.NOTIFY_URL, {
    method: 'POST',
    headers: { authorization: `Bearer ${env.NOTIFY_TOKEN}`, 'content-type': 'application/json' },
    body: JSON.stringify({ text }),
  });
  if (!res.ok) throw new Error(`Slack 通知に失敗: ${res.status} ${await res.text()}`);
}

main()
  .then(() => { fs.rmSync(TMP, { recursive: true, force: true }); })
  .catch(async e => {
    console.error('[ah] 失敗:', e.message);
    try {
      const env = loadEnv();
      // 失敗も必ず表に出す。黙って止まると「今月ぶんが無い」ことに誰も気付かない。
      if (!DRY) await notify(env, `:x: *All Hands 資料の自動生成に失敗しました*（対象 ${CUR}）\n\`\`\`${e.message}\`\`\`\n手動で \`node scripts/all-hands/build-ah.mjs ${CUR}\` を実行してください。`);
    } catch { /* 通知そのものが死んでいる場合は諦めてログに残す */ }
    process.exit(1);
  });
