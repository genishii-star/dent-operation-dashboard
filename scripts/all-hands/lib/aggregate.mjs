// D1 の予約/日次と wiki の物件マスタ(YAML)から、All Hands デッキ用の集計を作る。
//
// ⚠ 集計式は operation/app.js（ダッシュボード）に厳密に合わせてある。
//   app.js 側を変えたらここも変えること。ズレるとデッキとダッシュボードで数字が食い違う。
//   - 売上/予約組数/ゲスト数/チャネル/国籍 → 予約ベース（チェックイン月）
//   - 室泊数/ADR/稼働率/物件数           → 日次ベース（ユニーク prop-day・清掃onlyの行を除外）
import fs from 'fs';
import path from 'path';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);

// ---- app.js から写した判定ロジック ----
export const parseNum = s => { if (!s) return 0; return parseFloat(String(s).replace(/[,¥￥\s]/g, '')) || 0; };
export const normalizeDate = s => (s ? String(s).replace(/\//g, '-').trim() : '');
export const getYearMonth = s => { const d = normalizeDate(s); return d.length >= 7 ? d.substring(0, 7) : ''; };
export const getDaysInMonth = ym => { const [y, m] = ym.split('-').map(Number); return new Date(y, m, 0).getDate(); };
const isCancelStatus = st => st === 'キャンセル' || st === 'システムキャンセル';
// AirHost のカレンダーブロック。売上0・泊数が数年単位の行が混ざるので予約ベース集計から必ず外す。
const isBlockedStatus = st => st === 'ブロックされた';
const deriveArea = a => (!a ? 'その他' : a.includes('東京') ? '東京' : a.includes('京都') ? '京都' : a.includes('大阪') ? '大阪' : 'その他');

const REGION_TO_AREA = { osaka: '大阪', kyoto: '京都', tokyo: '東京', hyogo: '兵庫', okinawa: '沖縄' };
const STATUS_TO_JA = { active: '稼働中', preparing: '準備中', paused: '停止中', ended: '終了' };
const PROPERTY_NAME_MERGE = { 'HGK(旧)': 'HGK', 'HGK旧': 'HGK', Yuan: 'YUAN', 'NNJ(旧)': 'NNJ', NNJ旧: 'NNJ' };

// wiki 側が rooms の書式を変えても静かに壊れないよう、取り出せない場合は空文字を返す（呼び側で検知）
const extractRoomCode = re => (re && typeof re === 'object' ? String(re.id ?? re.room ?? re.code ?? '') : re == null ? '' : String(re));

function facilityToMasterRows(f) {
  const base = {
    name: f.name || String(f.code ?? '').toUpperCase(),
    address: f.address || '',
    area: REGION_TO_AREA[(f.region || '').toLowerCase()] || '',
    excludeKpi: f.kpi_exclude === true || f.kpi_excluded === true,
    status: STATUS_TO_JA[f.status] || (f.status ? f.status : '稼働中'),
  };
  const rooms = Array.isArray(f.rooms) && f.rooms.length > 0 ? f.rooms : null;
  if (!rooms) return [{ ...base, code: String(f.code ?? '').toUpperCase() }];
  return rooms.map(re => ({ ...base, code: String(extractRoomCode(re) ?? '').toUpperCase() }));
}

/**
 * @param {object} opts
 * @param {string} opts.facilitiesDir  facilities YAML を展開したディレクトリ
 * @param {string} opts.yamlModulePath `yaml` パッケージを持つ node_modules の基点
 * @param {Array}  opts.reservationRows D1 reservations の生行
 * @param {Array}  opts.dailyRows       D1 daily_revenue の生行
 * @param {string} opts.curYm, opts.prvYm, opts.yaYm
 */
export function aggregate(opts) {
  const YAML = createRequire(opts.yamlModulePath)('yaml');
  const warnings = [];

  // ---- 物件マスタ ----
  const facilities = fs.readdirSync(opts.facilitiesDir).filter(f => f.endsWith('.yaml'))
    .map(f => {
      try { return YAML.parse(fs.readFileSync(path.join(opts.facilitiesDir, f), 'utf8')); }
      catch (e) { warnings.push(`facilities/${f} をパースできません: ${e.message}`); return null; }
    }).filter(Boolean);
  facilities.forEach(f => {
    if (!Array.isArray(f.rooms)) return;
    f.rooms.forEach((re, i) => { if (!extractRoomCode(re).trim()) warnings.push(`${f.code}: rooms[${i}] から部屋コードを取得できません（YAMLスキーマ変更の疑い）`); });
  });

  let master = facilities.flatMap(facilityToMasterRows);
  const codes = new Set(master.map(m => m.code));
  master = master.filter(m => {
    if (PROPERTY_NAME_MERGE[m.code]) {
      const nc = PROPERTY_NAME_MERGE[m.code];
      if (codes.has(nc)) return false;
      m.code = nc;
    }
    return true;
  });
  master = master.filter(m => m.status === '稼働中');   // 稼働中以外は全画面で非表示

  const activeIds = new Set();
  master.forEach(m => { if (m.code) activeIds.add(m.code); if (m.name) activeIds.add(m.name); });

  // ---- 生データを日本語ヘッダ相当に写して正規化 ----
  const merge = n => PROPERTY_NAME_MERGE[n] || n;
  const rawResv = opts.reservationRows.map(r => ({
    property: merge(r.property_name || ''), roomNum: r.room_no || '',
    checkin: normalizeDate(r.check_in || ''), nights: parseNum(r.nights), status: r.status || '',
    guestCount: parseNum(r.guests), nationality: r.nationality || '', channel: r.channel || '',
    sales: parseNum(r.gross_sales), received: parseNum(r.net_received),
  })).filter(r => activeIds.has(r.property) || activeIds.has(r.property + r.roomNum));

  const rawDaily = opts.dailyRows.map(d => ({
    property: merge(d.property_name || ''), roomNum: d.room_no || '', date: normalizeDate(d.date || ''),
    status: d.status || '', channel: d.channel || '', sales: parseNum(d.gross_sales),
    received: parseNum(d.net_received), cleaning: parseNum(d.cleaning_fee),
  })).filter(d => activeIds.has(d.property) || activeIds.has(d.property + d.roomNum));

  const codeSet = new Set(master.map(m => m.code).filter(Boolean));
  const propCode = (name, room) => {
    if (!name) return '';
    if (room && room !== 'ALL' && name !== room && name.toLowerCase() !== room.toLowerCase()) {
      const cc = name + room;
      if (codeSet.has(cc)) return cc;
      if (codeSet.has(name)) return name;
      return cc;
    }
    return name;
  };

  const properties = master.filter(m => m.code).map(m => ({
    name: m.code, propName: m.name, area: m.area || deriveArea(m.address), rooms: 1,
    excludeKpi: m.excludeKpi, status: m.status,
  }));
  const byCode = {}, byPropName = {};
  properties.forEach(p => { byCode[p.name] = p; if (!byPropName[p.propName]) byPropName[p.propName] = p; });
  const findProp = r => byCode[propCode(r.property, r.roomNum)] || byCode[r.property] || byPropName[r.property] || null;

  const dailyIdx = {};
  rawDaily.forEach(d => { (dailyIdx[propCode(d.property, d.roomNum) + '|' + getYearMonth(d.date)] ||= []).push(d); });

  const today = new Date().toISOString().split('T')[0];
  function propStats(code, ym) {
    const p = byCode[code];
    if (!p) return null;
    // 日次は「今日まで」の実績だけを信頼する（app.js と同じ。過去月だけを扱う限り全件が対象）
    const rows = (dailyIdx[code + '|' + ym] || []).filter(d =>
      d.status !== 'システムキャンセル' && d.status !== 'ブロックされた' && d.date <= today);
    // チェックアウト日に清掃料だけ立つ行は稼働ではないので除く
    const use = rows.filter(d => !(d.cleaning > 0 && Math.abs(d.sales - d.cleaning) < 1));
    const dates = new Set(use.map(d => d.date));
    return {
      nights: dates.size,
      sales: use.reduce((s, d) => s + d.sales, 0),
      received: use.reduce((s, d) => s + d.received, 0),
    };
  }

  function overall(ym, area) {
    const fp = (!area || area === '全体') ? properties : properties.filter(p => p.area === area);
    let nights = 0, sales = 0;
    fp.forEach(p => { const s = propStats(p.name, ym); if (s) { nights += s.nights; sales += s.sales; } });
    const avail = fp.reduce((s, p) => s + getDaysInMonth(ym) * (p.rooms || 1), 0);
    return {
      occ: avail > 0 ? nights / avail : 0,
      adr: nights > 0 ? sales / nights : 0,
      roomRev: sales, nights, properties: fp.length,
    };
  }

  function resvFor(ym, area) {
    return rawResv.filter(r => {
      if (isBlockedStatus(r.status)) return false;
      // 純キャンセル(販売額0)のみ除外。非返金キャンセル(販売額>0=請求対象)は売上に計上する。
      if (isCancelStatus(r.status) && !(r.sales > 0)) return false;
      if (getYearMonth(r.checkin) !== ym) return false;
      if (area && area !== '全体') { const p = findProp(r); if (!p || p.area !== area) return false; }
      return true;
    });
  }

  function monthBlock(ym) {
    const rs = resvFor(ym, '全体');
    const confirmed = rs.filter(r => !isCancelStatus(r.status));   // 稼働量・平均値は確定のみ
    const o = overall(ym, '全体');
    const byChannel = {}, byCountry = {};
    rs.forEach(r => {
      byChannel[r.channel || 'その他'] = (byChannel[r.channel || 'その他'] || 0) + r.sales;
      byCountry[r.nationality || '不明'] = (byCountry[r.nationality || '不明'] || 0) + r.sales;
    });
    const sales = rs.reduce((s, r) => s + r.sales, 0);
    const guests = confirmed.reduce((s, r) => s + r.guestCount, 0);
    const rnights = confirmed.reduce((s, r) => s + r.nights, 0);
    return {
      sales, roomRev: o.roomRev, nights: o.nights, avgDailySales: sales / getDaysInMonth(ym),
      adr: o.adr, properties: o.properties, byChannel, count: confirmed.length, guests,
      avgGuests: confirmed.length ? guests / confirmed.length : 0,
      avgNights: confirmed.length ? rnights / confirmed.length : 0,
      byCountry, occ: o.occ,
    };
  }

  function areaBlock(ym) {
    const out = {};
    [...new Set(properties.map(p => p.area))].forEach(a => {
      const o = overall(ym, a);
      out[a] = { sales: resvFor(ym, a).reduce((s, r) => s + r.sales, 0), roomRev: o.roomRev,
        adr: o.adr, nights: o.nights, properties: o.properties, occ: o.occ };
    });
    return out;
  }

  // ---- チャネル積み上げ（13ヶ月） ----
  const TREND_CHANNELS = ['Airbnb', 'Booking.com', 'Expedia', 'AirHost', 'Agoda', 'manually_created',
    'isCalendar_imported', 'Vacation Stay：楽天', 'Trip.com', 'その他'];
  // Expedia は系列サイトが別名で入ってくるので1本に畳む
  const EXPEDIA = new Set(['Expedia', 'Expedia Affiliate Network', 'Hotels.com', 'American Express Travel',
    'Amex Travel', 'Travelocity', 'Orbitz', 'Wotif']);
  const fold = ch => {
    if (EXPEDIA.has(ch)) return 'Expedia';
    if (ch === '手動作成') return 'manually_created';
    if (ch === 'Rakuten Oyado（旧 Vacation Stay：楽天）') return 'Vacation Stay：楽天';
    if (TREND_CHANNELS.includes(ch)) return ch;
    if (/calendar/i.test(ch)) return 'isCalendar_imported';
    return 'その他';
  };

  const trendMonths = [];
  { let [y, m] = opts.yaYm.split('-').map(Number);
    for (let i = 0; i < 13; i++) { trendMonths.push(`${y}-${String(m).padStart(2, '0')}`); if (++m > 12) { m = 1; y++; } } }
  const trendData = {};
  trendMonths.forEach(ym => {
    const o = Object.fromEntries(TREND_CHANNELS.map(c => [c, 0]));
    resvFor(ym, '全体').forEach(r => { o[fold(r.channel)] += r.sales; });
    trendData[ym] = o;
  });
  // 「その他」に落ちたチャネルは畳み忘れの可能性があるので気付けるようにする
  const otherTotal = trendMonths.reduce((s, m) => s + trendData[m]['その他'], 0);
  if (otherTotal > 0) {
    const unknown = [...new Set(rawResv.map(r => r.channel).filter(ch => fold(ch) === 'その他' && ch))];
    warnings.push(`未分類チャネルが ¥${Math.round(otherTotal).toLocaleString()} 分あります: ${unknown.join(', ')}`);
  }

  return {
    warnings,
    activeRooms: properties.length,
    excludedRooms: properties.filter(p => p.excludeKpi).length,
    data: {
      cur: monthBlock(opts.curYm), prv: monthBlock(opts.prvYm), yearAgo: monthBlock(opts.yaYm),
      trend: { months: trendMonths, channels: TREND_CHANNELS, data: trendData },
    },
    extra: { area: areaBlock(opts.curYm), areaPrev: areaBlock(opts.prvYm), changes: { newIn: [], dropped: [] } },
  };
}
