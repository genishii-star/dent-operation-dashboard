// Tide（清掃パートナー）の活動実績を、Tide の運用ダッシュボード（tide-ops-api）の KV から集計する。
//
// ⚠ 売上・原価は載せない（Dent のデッキに Tide の金額を出さない＝2026-10-23 石井さん）。件数・物件・人数だけ。
//
// 正本は Tide アカウントの KV `ops-state`（清掃ジョブの台帳）。現場アプリの「済」などは `job-app-patch` に
// 差分で積まれ、worker.js の readState() が読むたびに畳み込む。ここでも同じく畳み込んでから数える。
// 月の基準は CO日（checkout||date）。tide-ops-api の /api/segment（地域別の月次）と同じ＝あちらと件数が揃う。
// スポット（kind="spot"＝セットアップ・共用部など清掃以外の対応）は清掃件数に混ぜない（同じく /api/segment と同じ）。
//
// 認証は Tide の Cloudflare API トークン（_Tide/07-deploy/.env）。gen の wrangler は Dent アカウントなので使えない。
// ここでは読むだけで、トークンを複製しない（finance(Tide)/deploy.sh と同じ扱い）。
import fs from 'fs';

const TIDE_ACCOUNT = '66e3de433a07899fd33073811c136d8d';
const TIDE_KV = '97838151436a431aa70dcba7d2307d31';   // tide-ops-api の STATE
const AREAS = ['大阪', '京都', '東京'];

function readEnvFile(f) {
  const env = {};
  fs.readFileSync(f, 'utf8').split('\n').forEach(line => {
    const m = line.match(/^\s*([A-Z_]+)\s*=\s*(.*)$/);
    if (m) env[m[1]] = m[2].trim().replace(/^["']|["']$/g, '');
  });
  return env;
}

async function kvGet(token, key) {
  const res = await fetch(`https://api.cloudflare.com/client/v4/accounts/${TIDE_ACCOUNT}/storage/kv/namespaces/${TIDE_KV}/values/${encodeURIComponent(key)}`,
    { headers: { authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Tide KV ${key}: HTTP ${res.status}`);
  return res.json();
}

// worker.js applyJobPatch の写し（数えるのに要る項目だけ）
function applyJobPatch(jobs, patch) {
  if (!patch || !Object.keys(patch).length) return jobs;
  return jobs.map(j => {
    const p = patch[j.id];
    if (!p) return j;
    const o = { ...j };
    if (p.status) o.status = p.status;
    if (p.who && !String(o.who || '').trim()) o.who = p.who;
    if (p.report) o.report = p.report;
    return o;
  });
}

const shift = (ym, n) => {
  let [y, m] = ym.split('-').map(Number);
  m += n;
  while (m < 1) { m += 12; y--; }
  while (m > 12) { m -= 12; y++; }
  return `${y}-${String(m).padStart(2, '0')}`;
};

/**
 * @param {object} opts
 * @param {string} opts.envFile   Tide の Cloudflare トークンを持つ .env
 * @param {string} opts.curYm, opts.prvYm
 * @param {object} opts.names     物件コード → 表示名（Dent の物件マスタ）
 * @returns {{ tide: object|null, warnings: string[] }}
 */
export async function aggregateTide({ envFile, curYm, prvYm, names = {} }) {
  const warnings = [];
  if (!fs.existsSync(envFile)) return { tide: null, warnings: [`Tide の認証ファイルがありません（${envFile}）。Tide スライドは空欄です`] };
  const env = readEnvFile(envFile);
  if (!env.CLOUDFLARE_API_TOKEN) return { tide: null, warnings: [`${envFile} に CLOUDFLARE_API_TOKEN がありません。Tide スライドは空欄です`] };

  const st = await kvGet(env.CLOUDFLARE_API_TOKEN, 'ops-state');
  if (!st || !Array.isArray(st.jobs)) throw new Error('Tide KV ops-state に jobs がありません');
  const jobs = applyJobPatch(st.jobs, await kvGet(env.CLOUDFLARE_API_TOKEN, 'job-app-patch'));

  const coOf = j => j.checkout || j.date || '';
  const isSpot = j => j.kind === 'spot';
  const areaOf = j => j.area || '(不明)';
  const unitOf = j => j.unit || j.prop;
  const cleans = jobs.filter(j => !isSpot(j) && coOf(j));

  // 物件ごとの初回清掃月＝「立ち上げ月」
  const firstYm = {};
  cleans.forEach(j => { const ym = coOf(j).slice(0, 7); if (!firstYm[j.prop] || ym < firstYm[j.prop]) firstYm[j.prop] = ym; });

  function month(ym) {
    const all = cleans.filter(j => coOf(j).startsWith(ym));
    const done = all.filter(j => j.status === '済');
    const byArea = {};
    done.forEach(j => {
      const a = byArea[areaOf(j)] || (byArea[areaOf(j)] = { count: 0, sameDay: 0, props: new Set() });
      a.count++; a.props.add(j.prop);
      if (j.nextIn && j.nextIn === (j.date || j.checkout)) a.sameDay++;
    });
    const withReport = done.filter(j => j.report);
    return {
      count: done.length,
      undone: all.length - done.length,
      props: new Set(done.map(j => j.prop)).size,
      units: new Set(done.map(unitOf)).size,
      // 清掃した日にそのまま次のゲストが入る＝待ったなしの清掃（ターン）
      sameDay: done.filter(j => j.nextIn && j.nextIn === (j.date || j.checkout)).length,
      staff: new Set(done.map(j => String(j.who || '').trim()).filter(Boolean)).size,
      reports: withReport.length,
      photos: withReport.reduce((s, j) => s + (+j.report.photos || 0), 0),
      byArea: Object.fromEntries(Object.entries(byArea).map(([a, v]) => [a, { count: v.count, sameDay: v.sameDay, props: v.props.size }])),
      // スポットは作業日基準（CO日を持たない）。件名だけ出す（金額は出さない）
      spots: jobs.filter(j => isSpot(j) && (j.date || '').startsWith(ym) && j.status === '済')
        .map(j => ({ title: String(j.title || 'スポット対応').slice(0, 40), prop: j.prop })),
      newProps: Object.keys(firstYm).filter(p => firstYm[p] === ym).sort()
        .map(p => ({ code: p, name: names[p] || p })),
    };
  }

  const cur = month(curYm), prv = month(prvYm);
  if (cur.undone) warnings.push(`Tide: ${curYm} の清掃のうち ${cur.undone} 件が「済」になっていません（件数に入れていません）`);

  // 推移は立ち上げ月から最大12ヶ月
  const start = Object.values(firstYm).sort()[0] || curYm;
  const months = [];
  for (let ym = curYm; ym >= start && months.length < 12; ym = shift(ym, -1)) months.unshift(ym);
  const trend = {
    months,
    areas: AREAS.filter(a => cleans.some(j => areaOf(j) === a)),
    data: months.map(ym => {
      const o = {};
      cleans.filter(j => j.status === '済' && coOf(j).startsWith(ym)).forEach(j => { o[areaOf(j)] = (o[areaOf(j)] || 0) + 1; });
      return o;
    }),
  };
  // 翌月の予定（台帳に入っている＝Dent の予約から同期済みのぶん。生成時点の見込み）
  const nextYm = shift(curYm, 1);
  const next = cleans.filter(j => coOf(j).startsWith(nextYm)).length;

  return { tide: { cur, prv, trend, next, nextYm, since: start }, warnings };
}
