// ============================================================
// Dent Inc. 運営ダッシュボード
// データソース: Google Sheets API (APIキー認証)
//
// スプレッドシートの共有設定を「リンクを知っている全員が閲覧者」にしてください。
// ============================================================

// ============================================================
// Chart.js Global Defaults
// ============================================================
Chart.defaults.font.family = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
Chart.defaults.font.size = 12;
Chart.defaults.color = '#888';
Chart.defaults.plugins.legend.labels.usePointStyle = true;
Chart.defaults.plugins.legend.labels.pointStyle = 'circle';
Chart.defaults.plugins.legend.labels.padding = 16;
Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(0,0,0,0.8)';
Chart.defaults.plugins.tooltip.cornerRadius = 8;
Chart.defaults.plugins.tooltip.padding = 10;
Chart.defaults.plugins.tooltip.titleFont = { size: 12, weight: '600' };
Chart.defaults.plugins.tooltip.bodyFont = { size: 12 };
Chart.defaults.elements.bar.borderRadius = 4;
Chart.defaults.elements.bar.borderSkipped = false;
Chart.defaults.elements.line.borderWidth = 2.5;
Chart.defaults.elements.point.radius = 4;
Chart.defaults.elements.point.hoverRadius = 6;
Chart.defaults.elements.point.borderWidth = 2;
Chart.defaults.elements.point.backgroundColor = '#fff';
Chart.defaults.scale.grid.color = 'rgba(0,0,0,0.04)';
Chart.defaults.scale.border = { display: false };
Chart.defaults.scale.ticks.padding = 8;

// Palette
const CHART_COLORS = {
  blue: '#4A90D9',
  green: '#50C878',
  orange: '#F5A623',
  purple: '#9B59B6',
  red: '#E74C3C',
  teal: '#1ABC9C',
  gray: '#95A5A6',
  pink: '#E91E90',
  indigo: '#5C6BC0',
  amber: '#FFB300',
};
const PALETTE = Object.values(CHART_COLORS);

// 販売金額/OCC/ADRチャート共通のツールチップフォーマッター
const salesChartTooltip = { callbacks: { label: ctx => { const v = ctx.parsed.y; const lbl = ctx.dataset.label || ''; if (lbl.includes('OCC')) return lbl + ': ' + v.toFixed(1) + '%'; return lbl + ': ¥' + Math.round(v).toLocaleString(); } } };

// 日本の祝日（年ごとに月-日 → 祝日名）
function getJapaneseHolidays(year) {
  const fixed = [
    [1,1,'元日'],[2,11,'建国記念の日'],[2,23,'天皇誕生日'],
    [4,29,'昭和の日'],[5,3,'憲法記念日'],[5,4,'みどりの日'],[5,5,'こどもの日'],
    [8,11,'山の日'],[11,3,'文化の日'],[11,23,'勤労感謝の日'],
  ];
  const holidays = {};
  fixed.forEach(([m, d, name]) => { holidays[`${m}-${d}`] = name; });
  // 成人の日（1月第2月曜）
  const monday2nd = (m) => { let d = new Date(year, m - 1, 1); const dow = d.getDay(); const first = dow === 1 ? 1 : (8 - dow) % 7 + 1; return first + 7; };
  holidays[`1-${monday2nd(1)}`] = '成人の日';
  // 海の日（7月第3月曜）
  const monday3rd = (m) => { let d = new Date(year, m - 1, 1); const dow = d.getDay(); const first = dow === 1 ? 1 : (8 - dow) % 7 + 1; return first + 14; };
  holidays[`7-${monday3rd(7)}`] = '海の日';
  // 敬老の日（9月第3月曜）
  holidays[`9-${monday3rd(9)}`] = '敬老の日';
  // スポーツの日（10月第2月曜）
  holidays[`10-${monday2nd(10)}`] = 'スポーツの日';
  // 春分の日・秋分の日（近似計算）
  const shunbun = Math.floor(20.8431 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  holidays[`3-${shunbun}`] = '春分の日';
  const shubun = Math.floor(23.2488 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  holidays[`9-${shubun}`] = '秋分の日';
  // 振替休日: 祝日が日曜なら翌月曜
  Object.keys(holidays).forEach(key => {
    const [m, d] = key.split('-').map(Number);
    const dt = new Date(year, m - 1, d);
    if (dt.getDay() === 0) {
      let sub = new Date(dt); sub.setDate(sub.getDate() + 1);
      const subKey = `${sub.getMonth() + 1}-${sub.getDate()}`;
      if (!holidays[subKey]) holidays[subKey] = '振替休日';
    }
  });
  return holidays;
}

const SHEET_ID = '1C7EiYSz-3ohjTy3Ul5zdgqL8cDk24jPj3ySTsAcOPhA';
const SHEET_GID_PROPERTY_MASTER = '416395562';
const SHEET_GID_OWNER_MASTER = '907386098';
const API_KEY = 'AIzaSyD_16gkzGw68S4socdFAr5HtIieisPA3uk';

function sheetApiUrl(sheetName) {
  return `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(sheetName)}?key=${API_KEY}`;
}

// ============================================================
// CSV Parser - handles quoted fields with commas and newlines
// ============================================================
function parseCSV(text) {
  const rows = [];
  let i = 0;
  const len = text.length;

  while (i < len) {
    const row = [];
    while (i < len) {
      let value = '';
      if (text[i] === '"') {
        // Quoted field
        i++; // skip opening quote
        while (i < len) {
          if (text[i] === '"') {
            if (i + 1 < len && text[i + 1] === '"') {
              value += '"';
              i += 2;
            } else {
              i++; // skip closing quote
              break;
            }
          } else {
            value += text[i];
            i++;
          }
        }
      } else {
        // Unquoted field
        while (i < len && text[i] !== ',' && text[i] !== '\n' && text[i] !== '\r') {
          value += text[i];
          i++;
        }
      }
      row.push(value.trim());

      if (i < len && text[i] === ',') {
        i++; // skip comma
      } else {
        break; // end of row
      }
    }
    // Skip line endings
    if (i < len && text[i] === '\r') i++;
    if (i < len && text[i] === '\n') i++;

    if (row.length > 1 || (row.length === 1 && row[0] !== '')) {
      rows.push(row);
    }
  }
  return rows;
}

function csvToObjects(text) {
  const rows = parseCSV(text);
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] || ''; });
    return obj;
  });
}

// ============================================================
// Data store
// ============================================================
let rawReservations = [];
let rawDailyData = [];
let propertyMaster = [];
let ownerMaster = [];
let seasonMaster = [];

// Computed data
let reservations = [];
let properties = [];
let owners = [];

// Filter state
let currentFilters = {
  dailyArea: '全体',
  ownerArea: '全体',
  propertyArea: '全体',
  revenueArea: '全体',
  pmbmArea: '全体',
  pmbmPeriod: 'thisMonth',
  dailyPeriod: 'thisMonth',
  ownerPeriod: 'thisMonth',
  propertyPeriod: 'thisMonth',
  reservationPeriod: 'yesterday',
  revenuePeriod: 'thisMonth',
  ownerProgressPeriod: 'thisMonth',
  propertyView: 'all',
  scorecardMode: 'all',
  scorecardDayType: 'total',
  propertyType: '全体',
  propertyLayout: '全体',
  propertySqm: '全体',
  revenueType: '全体',
  revenueLayout: '全体',
  revenueSqm: '全体',
  propDetailPeriod: 'thisMonth',
  grpDrillPeriod: 'thisMonth',
  marketPeriod: 'last3',
  marketCity: '大阪',
  marketSubTab: 'top',
  reservationView: 'all',
};

function setMarketSubTab(el) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.marketSubTab = el.dataset.subtab;
  document.querySelectorAll('.market-subtab').forEach(d => { d.style.display = 'none'; });
  const target = document.getElementById('market-sub-' + el.dataset.subtab);
  if (target) target.style.display = '';
  renderMarketTab();
  setTimeout(() => initMarketCharts(), 50);
}

function setReservationView(el) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.reservationView = el.dataset.view;
  renderReservationTab();
  setTimeout(initReservationCharts, 30);
}

// ============================================================
// Fetch all sheets
// ============================================================
// AirDNAシートを一括取得（AD_で始まるシートを全て読む）
// AirDNA市場データ用スプレッドシート（都市別）
const MARKET_SHEET_IDS = {
  '大阪': '1kPLF2Qq1EqPC7HeG2wPYYPT04mZE69_A-jlyUy-w0dw',
  '京都': '1584vBGDI8AfvoG01Zns1iG5Ivat31azZnu56Uqsx9qY',
  '東京': '17T5gIbfabcVq_IMx0S6rWk6wwEfWpvAKTHCUTuGHUiI',
};

// e-Stat公的統計スプシ
const ESTAT_SHEET_ID = '1d0dfPK79A9wXUyNqgGwbuq1aD5joIhcVvJ2uGiW2WQs';
const ESTAT_SHEETS = [
  'JNTO_訪日外客数', '宿泊統計_定員稼働率', '宿泊統計_延べ宿泊者数',
  'インバウンド消費', 'CPI_宿泊', '国内旅行_消費動向'
];

async function fetchEstatSheets() {
  const result = {};
  await Promise.all(ESTAT_SHEETS.map(async name => {
    try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${ESTAT_SHEET_ID}/values/${encodeURIComponent(name)}?key=${API_KEY}`;
      const resp = await fetch(url);
      if (!resp.ok) return;
      const json = await resp.json();
      const rows = json.values;
      if (!rows || rows.length < 2) return;
      const headers = rows[0];
      result[name] = rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = (row[i] !== undefined ? row[i] : ''); });
        return obj;
      });
    } catch (e) { /* skip */ }
  }));
  return result;
}

async function fetchAirdnaSheets() {
  const result = {};
  // 各都市: メタ取得後、values:batchGet で全シート1リクエストに集約（100レンジまで/req）
  for (const [city, sheetId] of Object.entries(MARKET_SHEET_IDS)) {
    try {
      const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets(properties(title))&key=${API_KEY}`;
      const metaResp = await fetch(metaUrl);
      if (!metaResp.ok) continue;
      const meta = await metaResp.json();
      const adSheetNames = (meta.sheets || [])
        .map(s => s.properties.title)
        .filter(name => name.startsWith('AD_'));
      if (adSheetNames.length === 0) continue;

      // batchGet は1リクエスト上限100レンジ
      for (let i = 0; i < adSheetNames.length; i += 100) {
        const chunk = adSheetNames.slice(i, i + 100);
        const ranges = chunk.map(n => `ranges=${encodeURIComponent(n)}`).join('&');
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchGet?${ranges}&key=${API_KEY}`;
        let resp;
        for (let attempt = 0; attempt < 3; attempt++) {
          resp = await fetch(url);
          if (resp.status !== 429) break;
          await new Promise(r => setTimeout(r, 2000 * (attempt + 1)));
        }
        if (!resp || !resp.ok) continue;
        const json = await resp.json();
        (json.valueRanges || []).forEach((vr, idx) => {
          const name = chunk[idx];
          const rows = vr.values;
          if (!rows || rows.length < 2) return;
          const headers = rows[0];
          result[name] = rows.slice(1).map(row => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = (row[i] !== undefined ? row[i] : ''); });
            return obj;
          });
        });
      }
    } catch (e) { /* skip city */ }
  }
  return result;
}

async function fetchSheet(sheetName) {
  const url = sheetApiUrl(sheetName);
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`Failed to load sheet: ${sheetName} (${resp.status})`);
  const json = await resp.json();
  const rows = json.values;
  if (!rows || rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (row[i] !== undefined ? row[i] : ''); });
    return obj;
  });
}

// ============================================================
// localStorage Cache
// ============================================================
const CACHE_KEY = 'dent_dashboard_cache';
const CACHE_TTL = 30 * 60 * 1000; // 30分

function saveCache(data) {
  try {
    const cache = { timestamp: Date.now(), data };
    localStorage.setItem(CACHE_KEY, JSON.stringify(cache));
  } catch (e) { /* localStorage full or unavailable */ }
}

function loadCache() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return null;
    const cache = JSON.parse(raw);
    return cache.data || null;
  } catch (e) { return null; }
}

function getCacheAge() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return Infinity;
    const cache = JSON.parse(raw);
    return Date.now() - (cache.timestamp || 0);
  } catch (e) { return Infinity; }
}

async function loadAllData() {
  const overlay = document.getElementById('loading-overlay');
  const detail = document.getElementById('loading-detail');
  const errorBanner = document.getElementById('error-banner');
  errorBanner.classList.remove('show');

  // キャッシュがあれば即表示
  const cached = loadCache();
  if (cached) {
    detail.textContent = 'キャッシュから表示中...';
    rawReservations = cached.resv || [];
    rawDailyData = cached.daily || [];
    propertyMaster = cached.propMaster || [];
    ownerMaster = cached.ownMaster || [];
    seasonMaster = cached.seasMaster || [];
    window._airdnaSheets = cached.marketRaw || {};
    window._estatSheets = cached.estatRaw || {};
    processData();
    renderAll();
    updateTimestamp();
    overlay.style.display = 'none';

    // バックグラウンドで最新データを取得（キャッシュ表示後に常に更新）
    fetchAndUpdate();
    return;
  }

  // キャッシュなし → 通常読み込み
  overlay.style.display = 'flex';
  try {
    await fetchAndUpdate();
    overlay.style.display = 'none';
  } catch (err) {
    console.error('Data loading error:', err);
    overlay.style.display = 'none';
    errorBanner.textContent = 'データの読み込みに失敗しました: ' + err.message;
    errorBanner.classList.add('show');
  }
}

async function fetchAndUpdate() {
  const detail = document.getElementById('loading-detail');
  if (detail) detail.textContent = '最新データを取得中...';

  const [resv, daily, propMaster, ownMaster, seasMaster, settingsRaw, marketRaw, estatRaw] = await Promise.all([
    fetchSheet('予約データ'),
    fetchSheet('日次データ'),
    fetchSheet('物件マスタ'),
    fetchSheet('オーナーマスタ'),
    fetchSheet('シーズンマスタ'),
    fetch(sheetApiUrl('設定')).then(r => r.json()).catch(() => ({})),
    fetchAirdnaSheets().catch(() => ({})),
    fetchEstatSheets().catch(() => ({})),
  ]);

  rawReservations = resv;
  rawDailyData = daily;
  propertyMaster = propMaster;
  ownerMaster = ownMaster;
  seasonMaster = seasMaster;
  window._airdnaSheets = marketRaw || {};
  window._estatSheets = estatRaw || {};

  // 最終同期タイムスタンプ（設定シートはキー・バリュー形式: [[key, value], ...]）
  const settingsRows = settingsRaw.values || [];
  const syncRow = settingsRows.find(r => r[0] === '最終同期');
  if (syncRow && syncRow[1]) {
    const el = document.getElementById('lastSynced');
    if (el) el.textContent = '最終同期: ' + syncRow[1];
  }

  processData();
  renderAll();
  updateTimestamp();

  // キャッシュ保存
  saveCache({ resv, daily, propMaster, ownMaster, seasMaster, marketRaw, estatRaw });
}

// ============================================================
// Process raw data into internal structures
// ============================================================
function parseNum(s) {
  if (!s) return 0;
  return parseFloat(String(s).replace(/[,¥￥\s]/g, '')) || 0;
}

function normalizeDate(s) {
  if (!s) return '';
  // Handle various formats: 2026/04/01, 2026-04-01, etc.
  return String(s).replace(/\//g, '-').trim();
}

function getYearMonth(dateStr) {
  const d = normalizeDate(dateStr);
  if (d.length >= 7) return d.substring(0, 7);
  return '';
}

function deriveArea(address) {
  if (!address) return 'その他';
  if (address.includes('東京')) return '東京';
  if (address.includes('京都')) return '京都';
  if (address.includes('大阪')) return '大阪';
  return 'その他';
}

// 住所から区を抽出（例: "大阪市西成区萩之茶屋1-4-1" → "西成区"）
function extractWard(address) {
  if (!address) return null;
  const m = address.match(/([\u4e00-\u9fa5]+区)/);
  return m ? m[1] : null;
}

// 日本語の区名 → AirDNA英語名（大阪・京都・東京の主要区）
const WARD_JP_TO_EN = {
  // 大阪市
  '都島区': 'MiyakojimaKu', '城東区': 'JotoKu', '港区': 'MinatoKu', '生野区': 'IkunoKu',
  '西成区': 'NishinariKu', '中央区': 'ChuoKu', '福島区': 'FukushimaKu',
  '淀川区': 'YodogawaKu', '浪速区': 'NaniwaKu', '西淀川区': 'NishiyodogawaKu',
  '住之江区': 'SuminoeKu', '此花区': 'KonohanaKu', '北区': 'KitaKu',
  '旭区': 'AsahiKu', '阿倍野区': 'AbenoKu', '東住吉区': 'HigashisumiyoshiKu',
  '東淀川区': 'HigashiyodogawaKu', '天王寺区': 'TennojiKu', '大正区': 'TaishoKu',
  '東成区': 'HigashinariKu', '住吉区': 'SumiyoshiKu', '平野区': 'HiranoKu',
  '鶴見区': 'TsurumiKu',
  // 京都市
  '山科区': 'YamashinaKu', '伏見区': 'FushimiKu', '右京区': 'UkyoKu',
  '下京区': 'ShimogyoKu', '南区': 'MinamiKu', '中京区': 'NakagyoKu',
  '東山区': 'HigashiyamaKu', '上京区': 'KamigyoKu', '左京区': 'SakyoKu',
  '西京区': 'NishikyoKu',
  // 東京23区
  '江東区': 'KotoKu', '渋谷区': 'ShibuyaKu', '杉並区': 'SuginamiKu',
  '中野区': 'NakanoKu', '新宿区': 'ShinjukuKu', '台東区': 'TaitoKu',
  '墨田区': 'SumidaKu', '目黒区': 'MeguroKu', '世田谷区': 'SetagayaKu',
  '豊島区': 'ToshimaKu', '葛飾区': 'KatsushikaKu', '荒川区': 'ArakawaKu',
  '千代田区': 'ChiyodaKu', '大田区': 'OtaKu', '足立区': 'AdachiKu',
  '江戸川区': 'EdogawaKu', '板橋区': 'ItabashiKu', '品川区': 'ShinagawaKu',
  '文京区': 'BunkyoKu', '練馬区': 'NerimaKu',
  // 注: 中央区、北区、港区は複数市にあるため文脈依存
};
// 名前衝突する区のエリア別マッピング
const WARD_AMBIGUOUS = {
  '中央区': { '大阪': 'ChuoKu', '東京': 'ChuoKu' },
  '北区': { '大阪': 'KitaKu', '京都': 'KitaKu', '東京': 'KitaKu' },
  '港区': { '大阪': 'MinatoKu', '東京': 'MinatoKu' },
  '西区': { '大阪': 'NishiKu(OsakaCity)', '堺': 'NishiKu(SakaiCity)' },
};

function wardJpToAirdna(wardJp, area, address) {
  if (!wardJp) return null;
  if (WARD_AMBIGUOUS[wardJp]) {
    // 堺市などの判定
    if (address && address.includes('堺市')) return WARD_AMBIGUOUS[wardJp]['堺'] || null;
    return WARD_AMBIGUOUS[wardJp][area] || null;
  }
  return WARD_JP_TO_EN[wardJp] || null;
}

// 物件の分析インサイト（ルールベース）
function buildInsightsHtml(prop, curStats, wdhdStats, paceData) {
  if (!prop || !curStats) return '';
  const insights = []; // {level, icon, title, text, category}

  // ── 価格調整系 ──
  if (wdhdStats) {
    const wd = wdhdStats.weekday;
    const hd = wdhdStats.holiday;

    // 休日値上げ余地
    if (hd.occ >= 85 && hd.nights > 3) {
      const suggestUp = Math.round(hd.adr * 0.12);
      insights.push({ level: 'success', icon: '↑', category: '価格',
        title: '休日値上げ余地あり',
        text: `休日OCC ${fmtPct(hd.occ)}と高水準。休日料金を ¥${suggestUp.toLocaleString()}（+12%）程度上げる余地あり`
      });
    }
    // 平日値下げ推奨
    if (wd.occ > 0 && wd.occ < 40 && wd.nights > 0) {
      insights.push({ level: 'warning', icon: '↓', category: '価格',
        title: '平日値下げ推奨',
        text: `平日OCC ${fmtPct(wd.occ)}と低水準。平日料金を5〜10%下げる or 連泊割引で集客強化`
      });
    }
    // 休日プレミアム不足
    if (wd.adr > 0 && hd.adr > 0) {
      const premium = ((hd.adr - wd.adr) / wd.adr) * 100;
      if (premium < 10) {
        insights.push({ level: 'info', icon: '💡', category: '価格',
          title: '休日プレミアム不足',
          text: `休日ADRと平日ADRの差が${premium.toFixed(0)}%のみ。市場標準は+20〜30%。休日料金を見直す余地あり`
        });
      } else if (premium > 50) {
        insights.push({ level: 'warning', icon: '⚠', category: '価格',
          title: '休日料金が高すぎる可能性',
          text: `休日ADRが平日の+${premium.toFixed(0)}%。休日OCCが${fmtPct(hd.occ)}なので、下げてOCCを取る選択肢も`
        });
      }
    }
  }

  // ── ペース系（先行予約） ──
  if (paceData && paceData.length >= 3) {
    const bucket30 = paceData[0], bucket60 = paceData[1], bucket90 = paceData[2];
    const occFor = (b) => {
      const total = b.weekday.nights + b.holiday.nights;
      const avail = b.weekday.avail + b.holiday.avail;
      return avail > 0 ? (total / avail) * 100 : 0;
    };
    const occ30 = occFor(bucket30), occ60 = occFor(bucket60), occ90 = occFor(bucket90);

    if (occ30 < 40) {
      insights.push({ level: 'warning', icon: '⏰', category: 'ペース',
        title: '直近30日の予約ペース遅い',
        text: `30日先OCC ${fmtPct(occ30)}（市場基準80%）。料金下げ or プロモ強化で需要取り込み`
      });
    } else if (occ30 >= 85) {
      insights.push({ level: 'success', icon: '🔥', category: 'ペース',
        title: '直近30日ほぼ満室',
        text: `30日先OCC ${fmtPct(occ30)}。直近料金を値上げする絶好機`
      });
    }
    if (occ90 >= 30) {
      insights.push({ level: 'success', icon: '📅', category: 'ペース',
        title: '早期予約が順調',
        text: `91日先でもOCC ${fmtPct(occ90)}。長期の値上げ余地あり`
      });
    }
    // 休日ペース集中
    if (bucket30.holiday.occ > 0 && bucket30.weekday.occ > 0) {
      const hdWdRatio = bucket30.holiday.occ / Math.max(bucket30.weekday.occ, 1);
      if (hdWdRatio > 2.5) {
        insights.push({ level: 'info', icon: '⚖', category: 'バランス',
          title: '休日偏重の予約状況',
          text: `休日が平日の${hdWdRatio.toFixed(1)}倍埋まっている。平日プロモで需要分散推奨`
        });
      }
    }
  }

  // ── 市場比較（AirDNAデータがある場合） ──
  const adSheets = window._airdnaSheets || {};
  const ward = extractWard(prop.address || '');
  const wardEn = wardJpToAirdna(ward, prop.area, prop.address);
  const beds = layoutToBedrooms(prop.layout);
  let mktOcc = null, mktAdr = null;

  const findAvgFromSheet = (sheetName, field) => {
    const sheet = adSheets[sheetName];
    if (!sheet) return null;
    const now = new Date();
    const recent = [];
    for (let i = -11; i <= 0; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      recent.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
    }
    const vals = sheet
      .filter(r => recent.includes((r['Date'] || '').slice(0, 7)))
      .map(r => parseFloat(r[field]))
      .filter(v => !isNaN(v) && v > 0);
    return vals.length > 0 ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
  };

  if (wardEn && beds !== null) {
    const bedSuffix = beds === 0 ? 'Studio' : (beds === 4 ? '4BR+' : `${beds}BR`);
    mktOcc = findAvgFromSheet(`AD_${prop.area}_${wardEn}_${bedSuffix}_occupancy`, 'Rate');
  }
  if (mktOcc === null && wardEn) {
    mktOcc = findAvgFromSheet(`AD_${prop.area}_${wardEn}_occupancy`, 'Rate');
  }

  if (mktOcc !== null) {
    const occDiff = curStats.occ - mktOcc;
    if (occDiff > 10) {
      insights.push({ level: 'success', icon: '🏆', category: '競争力',
        title: '市場平均を大きく上回る',
        text: `OCC ${fmtPct(curStats.occ)} vs 市場平均 ${fmtPct(mktOcc)}（+${occDiff.toFixed(1)}pt）。料金を段階的に上げて収益最大化を検討`
      });
    } else if (occDiff < -10) {
      insights.push({ level: 'warning', icon: '📉', category: '競争力',
        title: '市場平均を下回る',
        text: `OCC ${fmtPct(curStats.occ)} vs 市場平均 ${fmtPct(mktOcc)}（${occDiff.toFixed(1)}pt）。リスティング・写真・料金のいずれかに改善余地`
      });
    }
  }

  // ── 目標達成 ──
  // targetLow / targetNormal / targetHigh を使って判定（現状は簡易版）
  if (prop.targetNormal > 0 && curStats.sales > 0) {
    const achievePct = (curStats.sales / prop.targetNormal) * 100;
    if (achievePct >= 120) {
      insights.push({ level: 'success', icon: '🎯', category: '目標',
        title: '目標を大幅超過',
        text: `売上が目標の${achievePct.toFixed(0)}%。来期の目標を引き上げる余地あり`
      });
    } else if (achievePct < 70) {
      insights.push({ level: 'warning', icon: '⚠', category: '目標',
        title: '目標未達リスク',
        text: `売上が目標の${achievePct.toFixed(0)}%。早急に価格見直し or 販促強化が必要`
      });
    }
  }

  // ── 空の場合 ──
  if (insights.length === 0) {
    insights.push({ level: 'info', icon: '✓', category: '総合',
      title: '安定稼働中',
      text: '現時点で特記すべき異常なし。定期的に市場動向をチェックしてください'
    });
  }

  // HTML生成
  const levelColors = {
    success: { bg: '#34C75912', border: '#34C759', text: '#34C759' },
    warning: { bg: '#FF950012', border: '#FF9500', text: '#FF9500' },
    info: { bg: '#007AFF12', border: '#007AFF', text: '#007AFF' },
  };

  const items = insights.map(ins => {
    const c = levelColors[ins.level] || levelColors.info;
    return `<div style="background:${c.bg};border-left:3px solid ${c.border};border-radius:6px;padding:10px 14px;margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
        <span style="font-size:16px;">${ins.icon}</span>
        <span style="font-weight:700;font-size:13px;color:${c.text};">${ins.title}</span>
        <span style="font-size:10px;color:#86868b;margin-left:auto;background:rgba(0,0,0,0.04);padding:2px 6px;border-radius:4px;">${ins.category}</span>
      </div>
      <div style="font-size:12px;color:#1d1d1f;line-height:1.5;">${ins.text}</div>
    </div>`;
  }).join('');

  return `<div style="margin-bottom:20px;">
    <div style="font-size:13px;font-weight:600;color:#1d1d1f;margin-bottom:10px;">
      🔍 分析インサイト <span style="font-size:11px;font-weight:400;color:#86868b;">（${insights.length}件）</span>
    </div>
    ${items}
  </div>`;
}

// シリーズまとめの分析インサイト（ルールベース）
function buildGroupInsightsHtml(seriesBase, seriesProps, curAgg, months) {
  if (!seriesProps || !seriesProps.length) return '';
  const insights = [];
  const propNames = seriesProps.map(p => p.name);

  // 1. 室間パフォーマンス格差
  const perRoom = seriesProps.map(p => {
    let nights = 0, sales = 0;
    months.forEach(ym => {
      const s = computePropertyStats(p.name, ym);
      if (s) { nights += s.nights; sales += s.sales; }
    });
    const days = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);
    const avail = days * (p.rooms || 1);
    return { name: p.name, occ: avail > 0 ? (nights / avail) * 100 : 0, sales };
  });
  if (perRoom.length >= 2) {
    const occs = perRoom.map(r => r.occ);
    const maxOcc = Math.max(...occs), minOcc = Math.min(...occs);
    if (maxOcc - minOcc >= 30) {
      const top = perRoom.reduce((a, b) => a.occ > b.occ ? a : b);
      const bot = perRoom.reduce((a, b) => a.occ < b.occ ? a : b);
      insights.push({ level: 'warning', icon: '⚠', category: '室間格差',
        title: '室間の稼働差が大きい',
        text: `最高 ${top.name}（${fmtPct(top.occ)}） vs 最低 ${bot.name}（${fmtPct(bot.occ)}） ＝ ${(maxOcc - minOcc).toFixed(0)}pt差。低稼働室の料金見直し・写真刷新を検討`
      });
    }
  }

  // 2. 不稼働室検出
  const recentMonth = months[months.length - 1];
  const inactive = seriesProps.filter(p => {
    const s = computePropertyStats(p.name, recentMonth);
    return s && s.nights === 0;
  });
  if (inactive.length > 0 && inactive.length < seriesProps.length) {
    insights.push({ level: 'warning', icon: '🚫', category: '不稼働',
      title: `${inactive.length}室が直近月ゼロ稼働`,
      text: `${inactive.map(p => p.name).join(', ')} が${recentMonth}に0泊。リスティング停止・価格帯ミスマッチの可能性`
    });
  }

  // 3. ペース（シリーズ全体）
  const paceData = computePaceReport(propNames);
  if (paceData && paceData.length >= 3) {
    const occFor = (b) => {
      const total = b.weekday.nights + b.holiday.nights;
      const avail = b.weekday.avail + b.holiday.avail;
      return avail > 0 ? (total / avail) * 100 : 0;
    };
    const occ30 = occFor(paceData[0]);
    const occ90 = occFor(paceData[2]);
    if (occ30 < 40) {
      insights.push({ level: 'warning', icon: '⏰', category: 'ペース',
        title: 'シリーズ全体で30日先ペース遅い',
        text: `集約OCC ${fmtPct(occ30)}（基準80%）。全室で料金見直し or 販促強化を検討`
      });
    } else if (occ30 >= 85) {
      insights.push({ level: 'success', icon: '🔥', category: 'ペース',
        title: 'シリーズ全体が30日先で高稼働',
        text: `集約OCC ${fmtPct(occ30)}。段階的な値上げ余地あり`
      });
    }
    if (occ90 >= 30) {
      insights.push({ level: 'success', icon: '📅', category: 'ペース',
        title: 'シリーズ全体で早期予約が順調',
        text: `91日先でも集約OCC ${fmtPct(occ90)}。長期の値上げ余地あり`
      });
    }
  }

  // 4. 目標達成（シリーズ合計）
  const targetSum = seriesProps.reduce((s, p) => s + (p.targetNormal || 0) * months.length, 0);
  if (targetSum > 0 && curAgg && curAgg.sales > 0) {
    const achieve = (curAgg.sales / targetSum) * 100;
    if (achieve >= 120) {
      insights.push({ level: 'success', icon: '🎯', category: '目標',
        title: 'シリーズ合計で目標大幅超過',
        text: `合計売上が目標の${achieve.toFixed(0)}%。来期の目標引き上げを検討`
      });
    } else if (achieve < 70) {
      insights.push({ level: 'warning', icon: '⚠', category: '目標',
        title: 'シリーズ合計で目標未達リスク',
        text: `合計売上が目標の${achieve.toFixed(0)}%。価格・販促の見直しが急務`
      });
    }
  }

  // 5. 市場比較（エリア全域 vs シリーズ集約）
  const area = seriesProps[0] && seriesProps[0].area;
  if (area && curAgg) {
    const mk = resolveAreaMarketLookup(area);
    if (mk.hasData) {
      const now = new Date();
      const recentYms = [];
      for (let i = -11; i <= 0; i++) {
        const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
        recentYms.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
      }
      const mktOccs = recentYms.map(ym => mk.occ(ym)).filter(v => v !== null);
      const mktAvg = mktOccs.length ? mktOccs.reduce((s, v) => s + v, 0) / mktOccs.length : null;
      if (mktAvg !== null) {
        const diff = curAgg.occ - mktAvg;
        if (diff > 10) {
          insights.push({ level: 'success', icon: '🏆', category: '競争力',
            title: `${area}全域平均を大きく上回る`,
            text: `シリーズOCC ${fmtPct(curAgg.occ)} vs ${area}全域平均 ${fmtPct(mktAvg)}（+${diff.toFixed(1)}pt）。値上げで収益最大化の余地`
          });
        } else if (diff < -10) {
          insights.push({ level: 'warning', icon: '📉', category: '競争力',
            title: `${area}全域平均を下回る`,
            text: `シリーズOCC ${fmtPct(curAgg.occ)} vs ${area}全域平均 ${fmtPct(mktAvg)}（${diff.toFixed(1)}pt）。リスティング・料金・写真のいずれかに改善余地`
          });
        }
      }
    }
  }

  if (insights.length === 0) {
    insights.push({ level: 'info', icon: '✓', category: '総合',
      title: 'シリーズ安定稼働中',
      text: `${seriesProps.length}室すべて特記事項なし。市場動向を継続監視`
    });
  }

  const levelColors = {
    success: { bg: '#34C75912', border: '#34C759', text: '#34C759' },
    warning: { bg: '#FF950012', border: '#FF9500', text: '#FF9500' },
    info: { bg: '#007AFF12', border: '#007AFF', text: '#007AFF' },
  };
  const items = insights.map(ins => {
    const c = levelColors[ins.level] || levelColors.info;
    return `<div style="background:${c.bg};border-left:3px solid ${c.border};border-radius:6px;padding:10px 14px;margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
        <span style="font-size:16px;">${ins.icon}</span>
        <span style="font-weight:700;font-size:13px;color:${c.text};">${ins.title}</span>
        <span style="font-size:10px;color:#86868b;margin-left:auto;background:rgba(0,0,0,0.04);padding:2px 6px;border-radius:4px;">${ins.category}</span>
      </div>
      <div style="font-size:12px;color:#1d1d1f;line-height:1.5;">${ins.text}</div>
    </div>`;
  }).join('');
  return `<div style="margin-bottom:20px;">
    <div style="font-size:13px;font-weight:600;color:#1d1d1f;margin-bottom:10px;">
      🔍 シリーズ分析インサイト <span style="font-size:11px;font-weight:400;color:#86868b;">（${insights.length}件）</span>
    </div>
    ${items}
  </div>`;
}

// 物件の市場比較HTMLを生成
function buildMarketCompareHtml(prop, curStats) {
  if (!prop || !curStats) return '';
  const adSheets = window._airdnaSheets || {};
  const ward = extractWard(prop.address || '');
  const wardEn = wardJpToAirdna(ward, prop.area, prop.address);
  const beds = layoutToBedrooms(prop.layout);

  // 探索順: 区×間取り → 区全体 → エリア全体
  let occSheet = null, adrSheet = null, revSheet = null;
  let matchedLevel = '';

  if (wardEn && beds !== null) {
    // 区 × 間取りのシートを探す
    const bedSuffix = beds === 0 ? 'Studio' : (beds === 4 ? '4BR+' : `${beds}BR`);
    const occName = `AD_${prop.area}_${wardEn}_${bedSuffix}_occupancy`;
    if (adSheets[occName]) {
      occSheet = adSheets[occName];
      adrSheet = adSheets[`AD_${prop.area}_${wardEn}_${bedSuffix}_rates_summary`] || null;
      revSheet = adSheets[`AD_${prop.area}_${wardEn}_${bedSuffix}_revenue_summary`] || null;
      matchedLevel = `${ward} × ${prop.layout || bedSuffix}`;
    }
  }
  if (!occSheet && wardEn) {
    // 区全体
    occSheet = adSheets[`AD_${prop.area}_${wardEn}_occupancy`];
    adrSheet = adSheets[`AD_${prop.area}_${wardEn}_rates_summary`];
    revSheet = adSheets[`AD_${prop.area}_${wardEn}_revenue_summary`];
    if (occSheet) matchedLevel = ward;
  }
  if (!occSheet) {
    // エリア全体
    occSheet = adSheets[`AD_${prop.area}全域_occupancy`];
    adrSheet = adSheets[`AD_${prop.area}全域_rates_summary`];
    if (occSheet) matchedLevel = `${prop.area}全域`;
  }

  if (!occSheet) {
    return `<div style="margin-bottom:20px;max-width:520px;font-size:12px;color:#86868b;">
      市場比較: AirDNAデータが未取込です${ward ? `（${ward}）` : ''}
    </div>`;
  }

  // 直近12ヶ月の平均を算出
  const now = new Date();
  const recentMonths = [];
  for (let i = -11; i <= 0; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    recentMonths.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  const avgField = (sheet, field) => {
    if (!sheet) return null;
    const vals = sheet
      .filter(r => recentMonths.includes((r['Date'] || '').slice(0, 7)))
      .map(r => parseFloat(r[field]))
      .filter(v => !isNaN(v) && v > 0);
    return vals.length > 0 ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
  };
  const mktOcc = avgField(occSheet, 'Rate'); // occupancy sheet has Rate column (we renamed 'rate' to 'Rate' or 'Value')
  // Actually the flatten maps 'rate' → 'Rate' (capitalized). Let me try multiple field names.
  const tryFields = (sheet, fields) => {
    for (const f of fields) {
      const v = avgField(sheet, f);
      if (v !== null) return v;
    }
    return null;
  };
  const mOcc = tryFields(occSheet, ['Rate', 'rate', 'Value', 'Occupancy']);
  const mAdr = tryFields(adrSheet, ['Average daily rate', 'Daily rate', 'Daily Rate', 'Rate', 'daily_rate', 'rate']);
  const mRevenue = tryFields(revSheet, ['Average annual revenue', 'Revenue', 'revenue', 'Rate']);

  // 自社の直近（curStatsから）
  const myOcc = curStats.occ;
  const myAdr = curStats.adr;
  const myRevpar = curStats.revpar;
  const mRevpar = (mOcc && mAdr) ? (mOcc / 100) * mAdr : null;

  const diffPct = (my, mkt) => (mkt && mkt > 0) ? Math.round(((my - mkt) / mkt) * 100) : null;
  const idxColor = (v) => v >= 0 ? '#34C759' : '#FF3B30';
  const idxSign = (v) => v > 0 ? '+' : '';

  const occDiff = diffPct(myOcc, mOcc);
  const adrDiff = diffPct(myAdr, mAdr);
  const revparDiff = diffPct(myRevpar, mRevpar);

  const html = `<div>
    <div style="font-size:13px;font-weight:600;color:#1d1d1f;margin-bottom:8px;">
      市場比較（AirDNA） <span style="font-size:11px;font-weight:400;color:#86868b;">基準: ${matchedLevel}（直近12ヶ月平均）</span>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed;">
      <thead><tr style="border-bottom:2px solid #e5e5ea;">
        <th style="text-align:left;padding:5px 4px;color:#86868b;font-weight:500;width:24%;"></th>
        <th style="text-align:right;padding:5px 4px;color:#86868b;font-weight:500;width:26%;">自社</th>
        <th style="text-align:right;padding:5px 4px;color:#86868b;font-weight:500;width:26%;">市場平均</th>
        <th style="text-align:right;padding:5px 4px;color:#86868b;font-weight:500;width:24%;">差分</th>
      </tr></thead>
      <tbody>
        <tr style="border-bottom:1px solid #f0f0f0;">
          <td style="padding:5px 4px;font-weight:500;">OCC</td>
          <td style="text-align:right;padding:5px 4px;font-weight:600;">${fmtPct(myOcc)}</td>
          <td style="text-align:right;padding:5px 4px;">${mOcc !== null ? fmtPct(mOcc) : '-'}</td>
          <td style="text-align:right;padding:5px 4px;color:${occDiff !== null ? idxColor(occDiff) : '#999'};font-weight:600;">${occDiff !== null ? idxSign(occDiff) + (myOcc - mOcc).toFixed(1) + 'pt' : '-'}</td>
        </tr>
        <tr style="border-bottom:1px solid #f0f0f0;">
          <td style="padding:5px 4px;font-weight:500;">ADR</td>
          <td style="text-align:right;padding:5px 4px;font-weight:600;">${fmtYenFull(Math.round(myAdr))}</td>
          <td style="text-align:right;padding:5px 4px;">${mAdr !== null ? fmtYenFull(Math.round(mAdr)) : '-'}</td>
          <td style="text-align:right;padding:5px 4px;color:${adrDiff !== null ? idxColor(adrDiff) : '#999'};font-weight:600;">${adrDiff !== null ? idxSign(adrDiff) + adrDiff + '%' : '-'}</td>
        </tr>
        <tr>
          <td style="padding:5px 4px;font-weight:500;">RevPAR</td>
          <td style="text-align:right;padding:5px 4px;font-weight:600;">${fmtYenFull(Math.round(myRevpar))}</td>
          <td style="text-align:right;padding:5px 4px;">${mRevpar !== null ? fmtYenFull(Math.round(mRevpar)) : '-'}</td>
          <td style="text-align:right;padding:5px 4px;color:${revparDiff !== null ? idxColor(revparDiff) : '#999'};font-weight:600;">${revparDiff !== null ? idxSign(revparDiff) + revparDiff + '%' : '-'}</td>
        </tr>
      </tbody>
    </table>
  </div>`;
  return html;
}

// 物件に対応するAirDNA月次参照を構築（ward×bed → ward → area全域の順に探索）
// 戻り値: { occ(ym), adr(ym), matched } — 未来月は同月前年にフォールバック
function resolveMarketLookup(prop) {
  const adSheets = window._airdnaSheets || {};
  const ward = extractWard(prop && prop.address || '');
  const wardEn = wardJpToAirdna(ward, prop && prop.area, prop && prop.address);
  const beds = layoutToBedrooms(prop && prop.layout);
  let occSheet = null, adrSheet = null, matched = '';

  if (wardEn && beds !== null) {
    const bedSuffix = beds === 0 ? 'Studio' : (beds === 4 ? '4BR+' : `${beds}BR`);
    const o = adSheets[`AD_${prop.area}_${wardEn}_${bedSuffix}_occupancy`];
    if (o) {
      occSheet = o;
      adrSheet = adSheets[`AD_${prop.area}_${wardEn}_${bedSuffix}_rates_summary`] || null;
      matched = `${ward} × ${prop.layout || bedSuffix}`;
    }
  }
  if (!occSheet && wardEn) {
    occSheet = adSheets[`AD_${prop.area}_${wardEn}_occupancy`] || null;
    adrSheet = adSheets[`AD_${prop.area}_${wardEn}_rates_summary`] || null;
    if (occSheet) matched = ward;
  }
  if (!occSheet && prop && prop.area) {
    occSheet = adSheets[`AD_${prop.area}全域_occupancy`] || null;
    adrSheet = adSheets[`AD_${prop.area}全域_rates_summary`] || null;
    if (occSheet) matched = `${prop.area}全域`;
  }

  const pick = (sheet, fields, ym) => {
    if (!sheet) return null;
    // 優先: 指定YMに完全一致、なければ前年同月
    const candidates = [ym];
    const [yStr, mStr] = ym.split('-');
    candidates.push(`${parseInt(yStr, 10) - 1}-${mStr}`);
    for (const targetYm of candidates) {
      const row = sheet.find(r => (r['Date'] || '').slice(0, 7) === targetYm);
      if (!row) continue;
      for (const f of fields) {
        const v = parseFloat(row[f]);
        if (!isNaN(v) && v > 0) return v;
      }
    }
    return null;
  };

  return {
    matched,
    hasData: !!occSheet,
    occ: (ym) => pick(occSheet, ['Rate', 'rate', 'Value', 'Occupancy'], ym),
    adr: (ym) => pick(adrSheet, ['Average daily rate', 'Daily rate', 'Daily Rate', 'Rate', 'daily_rate', 'rate'], ym),
  };
}

// エリア全域の月次参照（エリアフィルタ用）
function resolveAreaMarketLookup(area) {
  const adSheets = window._airdnaSheets || {};
  const pick = (sheet, fields, ym) => {
    if (!sheet) return null;
    const [yStr, mStr] = ym.split('-');
    const candidates = [ym, `${parseInt(yStr, 10) - 1}-${mStr}`];
    for (const t of candidates) {
      const row = sheet.find(r => (r['Date'] || '').slice(0, 7) === t);
      if (!row) continue;
      for (const f of fields) {
        const v = parseFloat(row[f]);
        if (!isNaN(v) && v > 0) return v;
      }
    }
    return null;
  };

  if (area && area !== '全体') {
    const o = adSheets[`AD_${area}全域_occupancy`] || null;
    const a = adSheets[`AD_${area}全域_rates_summary`] || null;
    return {
      matched: `${area}全域`,
      hasData: !!o,
      occ: (ym) => pick(o, ['Rate', 'rate', 'Value', 'Occupancy'], ym),
      adr: (ym) => pick(a, ['Average daily rate', 'Daily rate', 'Daily Rate', 'Rate', 'daily_rate', 'rate'], ym),
    };
  }
  // 全体: 3都市を単純平均
  const areas = ['大阪', '京都', '東京'];
  const occSheets = areas.map(ar => adSheets[`AD_${ar}全域_occupancy`]).filter(Boolean);
  const adrSheets = areas.map(ar => adSheets[`AD_${ar}全域_rates_summary`]).filter(Boolean);
  const avg = (sheets, fields, ym) => {
    const vals = sheets.map(s => pick(s, fields, ym)).filter(v => v !== null);
    return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
  };
  return {
    matched: '主要3都市平均',
    hasData: occSheets.length > 0,
    occ: (ym) => avg(occSheets, ['Rate', 'rate', 'Value', 'Occupancy'], ym),
    adr: (ym) => avg(adrSheets, ['Average daily rate', 'Daily rate', 'Daily Rate', 'Rate', 'daily_rate', 'rate'], ym),
  };
}

// 間取りからbedroom数を推定
function layoutToBedrooms(layout) {
  if (!layout) return null;
  const s = layout.toUpperCase();
  // 1R, 1K, スタジオ → 0 (AirDNAの0BR=studio)
  if (/^1R$|^1K$|スタジオ|STUDIO/.test(s)) return 0;
  const m = s.match(/^(\d+)/);
  if (!m) return null;
  const n = parseInt(m[1], 10);
  return n >= 4 ? 4 : n; // 4LDK以上は4+にまとめる
}

function generatePropCode(propName, roomNum) {
  if (!propName) return '';
  // 物件コードは「物件名+ルーム番号」(多部屋) または「物件名」(単独) の連結形式。
  // マスタの物件コードSetと突き合わせて解決する。
  const codeSet = window._propCodeSet;
  if (roomNum && roomNum !== 'ALL' && propName !== roomNum && propName.toLowerCase() !== roomNum.toLowerCase()) {
    const concatenated = propName + roomNum;
    if (codeSet && codeSet.has(concatenated)) return concatenated;
    if (codeSet && codeSet.has(propName)) return propName;
    return concatenated;
  }
  if (codeSet && codeSet.has(propName)) return propName;
  return propName;
}

function processData() {
  invalidatePropStatsCache();
  // 物件名のマージ（旧名 → 新名）
  const PROPERTY_NAME_MERGE = {
    'HGK(旧)': 'HGK',
    'HGK旧': 'HGK',
  };
  rawReservations.forEach(r => {
    if (PROPERTY_NAME_MERGE[r['物件名']]) r['物件名'] = PROPERTY_NAME_MERGE[r['物件名']];
  });
  rawDailyData.forEach(d => {
    if (PROPERTY_NAME_MERGE[d['物件名']]) d['物件名'] = PROPERTY_NAME_MERGE[d['物件名']];
  });
  // 物件マスタ側もマージ：旧コード行は新コード行が存在すればドロップ、無ければ rename
  const _masterCodes = new Set(propertyMaster.map(pm => pm['物件コード'] || ''));
  propertyMaster = propertyMaster.filter(pm => {
    const code = pm['物件コード'] || '';
    if (PROPERTY_NAME_MERGE[code]) {
      const newCode = PROPERTY_NAME_MERGE[code];
      if (_masterCodes.has(newCode)) return false; // 重複→ドロップ
      pm['物件コード'] = newCode;
    }
    return true;
  });

  // 物件コードSetを構築（マスタの物件コード列がそのまま正規化キー）
  const codeSet = new Set();
  propertyMaster.forEach(pm => {
    const code = pm['物件コード'] || '';
    if (code) codeSet.add(code);
  });
  window._propCodeSet = codeSet;

  // Map reservations
  reservations = rawReservations.map(r => {
    const propName = r['物件名'] || '';
    const roomNum = r['部屋番号'] || '';
    return {
      id: r['AirHost予約ID'] || '',
      channel: r['予約サイト'] || '',
      channelId: r['チャンネル予約ID'] || '',
      date: normalizeDate(r['予約日']),
      property: propName,
      propCode: generatePropCode(propName, roomNum),
      roomNum: roomNum,
      guest: r['ゲスト名'] || '',
      nationality: r['国籍'] || '',
      guestCount: parseNum(r['ゲスト数']),
      checkin: normalizeDate(r['チェックイン']),
      checkout: normalizeDate(r['チェックアウト']),
      nights: parseNum(r['合計日数']),
      status: r['状態'] || '',
      sales: parseNum(r['販売']),
      received: parseNum(r['受取金']),
      otaFee: parseNum(r['OTA サービス料']),
      cleaningFee: parseNum(r['クリーニング代']),
      paid: r['支払い済み'] || '',
      roomTag: r['物件タグ'] || '',
    };
  });

  // Sort reservations by date descending
  reservations.sort((a, b) => b.date.localeCompare(a.date));

  // Populate status filter options from actual data
  const statuses = [...new Set(reservations.map(r => r.status).filter(Boolean))];
  const statusSelect = document.getElementById('resv-status-filter');
  statusSelect.innerHTML = '<option value="">すべて</option>';
  statuses.forEach(st => {
    statusSelect.innerHTML += `<option value="${st}">${st}</option>`;
  });

  // Build owner lookup（オーナーIDがそのまま表示名）
  // ロイヤリティ計算優先順位: 計算用ロイヤリティ列 → ロイヤリティ列
  // 階段式などフリーテキストで parseNum が失敗するオーナーは「計算用ロイヤリティ」列で固定%を指定
  const ownerMap = {};
  ownerMaster.forEach(om => {
    const id = om['オーナーID'] || '';
    const royaltyText = om['ロイヤリティ'] || '';
    const overrideText = om['計算用ロイヤリティ'] || '';
    const overridePct = parseNum(overrideText);
    const fallbackPct = parseNum(royaltyText);
    // パース失敗判定: 元テキストが空でも「運営費のみ」でも「0%」でもないのに 0 になるケース
    const isIntentionalZero = !royaltyText
      || /運営費のみ/.test(royaltyText)
      || /^0\s*%?$/.test(royaltyText.trim());
    const fallbackParseFailed = !overrideText && fallbackPct === 0 && !isIntentionalZero;
    let pct = 0;
    if (overrideText) pct = overridePct;
    else if (!fallbackParseFailed) pct = fallbackPct;
    ownerMap[id] = {
      id: id,
      name: id,
      royalty: royaltyText,
      royaltyOverride: overrideText,
      royaltyPct: pct,
      royaltyParseFailed: fallbackParseFailed,
    };
  });

  // Build season lookup (month number -> season type)
  const seasonMap = {};
  seasonMaster.forEach(sm => {
    const monthNum = parseNum(sm['月']);
    seasonMap[monthNum] = sm['シーズン'] || '通常期';
  });

  // Build properties from property master (master-based, not daily-data-based)
  properties = propertyMaster.map(pm => {
    const code = pm['物件コード'] || '';
    if (!code) return null;
    const ownerInfo = ownerMap[pm['オーナーID']] || {};
    const address = pm['住所'] || '';
    const area = pm['エリア'] || deriveArea(address);
    return {
      name: code,
      propName: pm['物件名'] || code,
      code: code,
      ownerId: pm['オーナーID'] || '',
      ownerName: ownerInfo.name || '',
      royalty: ownerInfo.royalty || '',
      royaltyPct: ownerInfo.royaltyPct || 0,
      royaltyParseFailed: !!ownerInfo.royaltyParseFailed,
      area: area,
      rooms: parseNum(pm['部屋数']) || 1,
      excludeKpi: (pm['KPI除外'] || '') === 'TRUE' || (pm['KPI除外'] || '') === '1',
      status: pm['ステータス'] || '稼働中',
      targetLow: parseNum(pm['閑散期目標']),
      targetNormal: parseNum(pm['通常期目標']),
      targetHigh: parseNum(pm['繁忙期目標']),
      airbnbAccount: pm['airbnbアカウント'] || '',
      airbnbListingId: pm['airbnbリスティングID'] || '',
      licenseType: pm['許可種類'] || '',
      operationLimitDays: parseNum(pm['営業日数上限']) || 0,
      startDate: normalizeDate(pm['運用開始日'] || ''),
      propType: pm['タイプ'] || '',
      layout: pm['間取り'] || '',
      sqm: parseNum(pm['平米数']) || 0,
    };
  }).filter(Boolean);

  // Build property lookup maps for fast access (avoid O(N) find per reservation)
  window._propByName = {};
  window._propByPropName = {};
  properties.forEach(p => {
    window._propByName[p.name] = p;
    if (p.propName) window._propByPropName[p.propName] = p;
  });

  // Build daily data index by propCode+ym for fast lookup
  window._dailyByPropYm = {};
  rawDailyData.forEach(d => {
    const date = normalizeDate(d['日付']);
    const code = generatePropCode(d['物件名'] || '', d['ルーム番号'] || '');
    const ym = getYearMonth(date);
    const key = code + '|' + ym;
    if (!window._dailyByPropYm[key]) window._dailyByPropYm[key] = [];
    window._dailyByPropYm[key].push(d);
  });

  // Build reservation index by propCode for fast lookup
  window._resvByProp = {};
  reservations.forEach(r => {
    const keys = new Set([r.propCode, r.property].filter(Boolean));
    keys.forEach(k => {
      if (!window._resvByProp[k]) window._resvByProp[k] = [];
      window._resvByProp[k].push(r);
    });
  });

  // Build owners array
  const ownerIds = [...new Set(propertyMaster.map(pm => pm['オーナーID']).filter(Boolean))];
  owners = ownerIds.map(oid => {
    const info = ownerMap[oid] || {};
    const ownerProps = properties.filter(p => p.ownerId === oid);
    return {
      id: oid,
      name: info.name || oid,
      royalty: info.royalty || '',
      royaltyOverride: info.royaltyOverride || '',
      royaltyPct: info.royaltyPct || 0,
      royaltyParseFailed: !!info.royaltyParseFailed,
      properties: ownerProps.map(p => p.name),
    };
  });
}

// ============================================================
// Aggregation helpers
// ============================================================
function getSelectedMonths(tabId) {
  const period = currentFilters[tabId + 'Period'] || 'thisMonth';
  const now = new Date();
  const thisYm = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  if (period === 'thisMonth') {
    return [thisYm];
  } else if (period === 'lastMonth') {
    const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
  } else if (period === 'last3Months') {
    const months = [];
    for (let i = 2; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'));
    }
    return months;
  } else if (period === 'nextMonth') {
    const d = new Date(now.getFullYear(), now.getMonth() + 1, 1);
    return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
  } else if (period === 'next2Month') {
    const d = new Date(now.getFullYear(), now.getMonth() + 2, 1);
    return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
  } else if (period === 'lastYear') {
    const d = new Date(now.getFullYear() - 1, now.getMonth(), 1);
    return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
  }
  return [thisYm];
}

// Backwards compat: single month for cases that need it
function getSelectedMonth(tabId) {
  const months = getSelectedMonths(tabId);
  return months[months.length - 1];
}

function getSelectedMonths_custom(period) {
  const now = new Date();
  const thisYm = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  if (period === 'thisMonth') return [thisYm];
  if (period === 'lastMonth') {
    const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
  }
  if (period === 'last3Month') {
    const d = new Date(now.getFullYear(), now.getMonth() - 3, 1);
    return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
  }
  if (period === 'lastYear') {
    const d = new Date(now.getFullYear() - 1, now.getMonth(), 1);
    return [d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')];
  }
  return [thisYm];
}

function getDaysInMonth(ym) {
  const [y, m] = ym.split('-').map(Number);
  return new Date(y, m, 0).getDate();
}

function getMonthNumber(ym) {
  return parseInt(ym.split('-')[1], 10);
}

function getSeasonForMonth(monthNum) {
  const s = seasonMaster.find(sm => parseNum(sm['月']) === monthNum);
  return s ? s['シーズン'] : '通常期';
}

function findPropByReservation(r) {
  return window._propByName[r.propCode] || window._propByName[r.property] || window._propByPropName[r.property] || null;
}

function findPropByName(name) {
  return window._propByName[name] || window._propByPropName[name] || null;
}

// マスタに存在しない物件名を検出（予約データ + 日次データ両方をスキャン）
function findOrphanProperties() {
  const isTest = name => /TEST/i.test(name || '');
  const orphans = new Map(); // key: 表示名, value: { source: Set, count }
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    if (findPropByReservation(r)) return;
    if (isTest(r.property) || isTest(r.roomNum)) return;
    const key = r.property + (r.roomNum ? ` (部屋${r.roomNum})` : '');
    if (!orphans.has(key)) orphans.set(key, { sources: new Set(), count: 0 });
    const o = orphans.get(key);
    o.sources.add('予約');
    o.count++;
  });
  rawDailyData.forEach(d => {
    const propName = d['物件名'] || '';
    const roomNum = d['ルーム番号'] || '';
    if (!propName) return;
    if (isTest(propName) || isTest(roomNum)) return;
    const code = generatePropCode(propName, roomNum);
    if (findPropByName(code)) return;
    const key = propName + (roomNum && roomNum !== 'ALL' ? ` (部屋${roomNum})` : '');
    if (!orphans.has(key)) orphans.set(key, { sources: new Set(), count: 0 });
    const o = orphans.get(key);
    o.sources.add('日次');
    o.count++;
  });
  return [...orphans.entries()].map(([name, info]) => ({ name, sources: [...info.sources], count: info.count }));
}

function renderOrphanAlert(containerId) {
  const el = document.getElementById(containerId);
  if (!el) return;
  const orphans = findOrphanProperties();
  if (orphans.length === 0) {
    el.innerHTML = '';
    return;
  }
  const items = orphans.map(o => `<li>${o.name} <span style="color:#999;">(${o.sources.join('/')}データ・${o.count}件)</span></li>`).join('');
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${SHEET_GID_PROPERTY_MASTER}`;
  el.innerHTML = `<div class="alert-orphan">
    <div class="alert-title">⚠ マスタ未登録の物件が ${orphans.length} 件あります</div>
    物件マスタへの追加または名称統一が必要です。<a href="${sheetUrl}" target="_blank" rel="noopener" style="color:#ff3b30;font-weight:600;text-decoration:underline;">物件マスタを開く ↗</a>
    <ul>${items}</ul>
  </div>`;
}

function getTargetForProperty(prop, monthNum) {
  const season = getSeasonForMonth(monthNum);
  if (season === '閑散期') return prop.targetLow;
  if (season === '繁忙期') return prop.targetHigh;
  return prop.targetNormal;
}

function filterPropertiesByArea(area, extraFilters) {
  let result = (!area || area === '全体') ? properties : properties.filter(p => p.area === area);
  if (extraFilters) {
    if (extraFilters.type && extraFilters.type !== '全体') result = result.filter(p => p.propType === extraFilters.type);
    if (extraFilters.layout && extraFilters.layout !== '全体') result = result.filter(p => p.layout === extraFilters.layout);
    if (extraFilters.sqm && extraFilters.sqm !== '全体') result = result.filter(p => matchSqmRange(p.sqm, extraFilters.sqm));
  }
  return result;
}

function matchSqmRange(sqm, range) {
  if (range === '〜20m²') return sqm > 0 && sqm <= 20;
  if (range === '20〜40m²') return sqm > 20 && sqm <= 40;
  if (range === '40〜60m²') return sqm > 40 && sqm <= 60;
  if (range === '60〜80m²') return sqm > 60 && sqm <= 80;
  if (range === '80m²〜') return sqm > 80;
  return true;
}

function getExtraFilters(tabId) {
  return {
    type: currentFilters[tabId + 'Type'] || '全体',
    layout: currentFilters[tabId + 'Layout'] || '全体',
    sqm: currentFilters[tabId + 'Sqm'] || '全体',
  };
}

function setTypeFilter(el, tabId) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters[tabId + 'Type'] = el.dataset.type;
  renderAll();
  setTimeout(() => initChartsForTab(tabId), 50);
}

function setLayoutFilter(el, tabId) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters[tabId + 'Layout'] = el.dataset.layout;
  renderAll();
  setTimeout(() => initChartsForTab(tabId), 50);
}

function setSqmFilter(el, tabId) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters[tabId + 'Sqm'] = el.dataset.sqm;
  renderAll();
  setTimeout(() => initChartsForTab(tabId), 50);
}

function buildLayoutPills(tabId) {
  const layouts = [...new Set(properties.map(p => p.layout).filter(Boolean))].sort();
  const container = document.getElementById(tabId + '-layout-filter');
  if (!container) return;
  const current = currentFilters[tabId + 'Layout'] || '全体';
  container.innerHTML = `<span class="pill ${current === '全体' ? 'active' : ''}" data-layout="全体" onclick="setLayoutFilter(this,'${tabId}')">全体</span>` +
    layouts.map(l => `<span class="pill ${current === l ? 'active' : ''}" data-layout="${l}" onclick="setLayoutFilter(this,'${tabId}')">${l}</span>`).join('');
}

function aggregateDailyForMonth(ym, areaFilter, excludeKpi, extraFilters) {
  const filteredProps = filterPropertiesByArea(areaFilter, extraFilters);
  const propNames = new Set(filteredProps.filter(p => !excludeKpi || !p.excludeKpi).map(p => p.name));

  // Filter daily data for this month and these properties
  const monthData = rawDailyData.filter(d => {
    const date = normalizeDate(d['日付']);
    const propCode = generatePropCode(d['物件名'] || '', d['ルーム番号'] || '');
    const status = d['状態'] || '';
    return getYearMonth(date) === ym && propNames.has(propCode) && status !== 'システムキャンセル';
  });

  return { monthData, filteredProps: filteredProps.filter(p => !excludeKpi || !p.excludeKpi), propNames };
}

const _propStatsCache = {};
let _propStatsCacheVer = 0;

function invalidatePropStatsCache() {
  _propStatsCacheVer++;
  Object.keys(_propStatsCache).forEach(k => delete _propStatsCache[k]);
}

function computePropertyStats(propName, ym) {
  const cacheKey = propName + '|' + ym;
  if (_propStatsCache[cacheKey]) return _propStatsCache[cacheKey];

  const prop = findPropByName(propName);
  if (!prop) return null;

  const daysInMonth = getDaysInMonth(ym);
  const totalAvailableDays = daysInMonth * (prop.rooms || 1);

  // 日次データ: 過去〜今日分の実績のみ信頼（未来分は予約データで補完するため除外）
  const today = new Date().toISOString().split('T')[0];
  const propDaily = (window._dailyByPropYm[propName + '|' + ym] || []).filter(d => {
    const status = d['状態'] || '';
    if (status === 'システムキャンセル' || status === 'ブロックされた') return false;
    const date = normalizeDate(d['日付']);
    if (date > today) return false; // 未来分は予約データで補完
    // クリーニング代のみの行を除外（チェックアウト日に清掃料だけ計上される）
    const cleaningFee = parseNum(d['清掃料']);
    const sales = parseNum(d['売上合計']);
    if (cleaningFee > 0 && Math.abs(sales - cleaningFee) < 1) return false;
    return true;
  });

  // 日次データに含まれる日付のセット（重複防止用 + ユニーク日数カウント）
  const dailyDates = new Set();
  propDaily.forEach(d => {
    const date = normalizeDate(d['日付']);
    dailyDates.add(date);
  });

  let bookedDays = dailyDates.size;
  let totalSales = propDaily.reduce((s, d) => s + parseNum(d['売上合計']), 0);
  let totalReceived = propDaily.reduce((s, d) => s + parseNum(d['受取金合計']), 0);

  // 予約データ: 日次データにない未来分の確定予約を補完
  const [ymY, ymM] = ym.split('-').map(Number);
  const monthStart = ym + '-01';
  const monthEnd = ym + '-' + String(daysInMonth).padStart(2, '0');

  // 予約データから未来分を追加（インデックス使用）
  const _resvCandidates = new Set();
  (window._resvByProp[propName] || []).forEach(r => _resvCandidates.add(r));
  const propReservations = [..._resvCandidates].filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
    return r.propCode === propName || r.property === propName;
  });

  let futureNights = 0;
  let futureSales = 0;
  let futureReceived = 0;

  propReservations.forEach(r => {
    if (!r.checkin || !r.checkout) return;
    // チェックイン〜チェックアウト間の各日を確認し、この月の未来泊数をカウント
    const ci = new Date(r.checkin);
    const co = new Date(r.checkout);
    let monthNights = 0;
    for (let d = new Date(ci); d < co; d.setDate(d.getDate() + 1)) {
      const ds = d.toISOString().split('T')[0];
      if (ds >= monthStart && ds <= monthEnd && ds > today && !dailyDates.has(ds)) {
        monthNights++;
        dailyDates.add(ds);
      }
    }
    futureNights += monthNights;
    // 売上・受取金を月内泊数で按分（清掃料を除外）
    if (monthNights > 0 && r.nights > 0) {
      const netSales = (r.sales || 0) - (r.cleaningFee || 0);
      futureSales += netSales * (monthNights / r.nights);
      futureReceived += (r.received || 0) * (monthNights / r.nights);
    }
  });

  bookedDays += futureNights;
  totalSales += futureSales;
  totalReceived += futureReceived;

  const occ = totalAvailableDays > 0 ? (bookedDays / totalAvailableDays) * 100 : 0;
  const adr = bookedDays > 0 ? totalSales / bookedDays : 0;
  const revpar = adr * (occ / 100);

  // Channel breakdown from daily data
  const channels = {};
  propDaily.forEach(d => {
    const ch = d['予約サイト'] || 'その他';
    if (!channels[ch]) channels[ch] = { count: 0, sales: 0 };
    channels[ch].count++;
    channels[ch].sales += parseNum(d['売上合計']);
  });

  const result = {
    name: propName,
    ownerId: prop.ownerId,
    ownerName: prop.ownerName,
    area: prop.area,
    rooms: prop.rooms,
    excludeKpi: prop.excludeKpi,
    status: prop.status,
    occ: occ,
    adr: adr,
    revpar: revpar,
    nights: bookedDays,
    sales: totalSales,
    received: totalReceived,
    channels: channels,
  };
  _propStatsCache[cacheKey] = result;
  return result;
}

// 先行予約ペースレポート（平日/休日 × バケット別）
// propNames: 物件コード配列, buckets: [{label, minDay, maxDay}]
function computePaceReport(propNames, buckets) {
  if (!buckets) buckets = [
    { label: '0〜30日', minDay: 0, maxDay: 30 },
    { label: '31〜60日', minDay: 31, maxDay: 60 },
    { label: '61〜90日', minDay: 61, maxDay: 90 },
  ];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayStr = today.toISOString().split('T')[0];

  // バケットごとの日付範囲を事前計算
  const bucketRanges = buckets.map(b => {
    const start = new Date(today); start.setDate(start.getDate() + b.minDay);
    const end = new Date(today); end.setDate(end.getDate() + b.maxDay);
    return { ...b, startStr: start.toISOString().split('T')[0], endStr: end.toISOString().split('T')[0] };
  });

  // 祝日キャッシュ（年をまたぐ可能性あり）
  const holidayCache = {};
  function getHolidays(y) { if (!holidayCache[y]) holidayCache[y] = getJapaneseHolidays(y); return holidayCache[y]; }

  function isHolidayNight(date) {
    const dow = date.getDay();
    if (dow === 5 || dow === 6) return true; // 金・土泊
    // 祝前日: 翌日が祝日
    const next = new Date(date); next.setDate(next.getDate() + 1);
    const hols = getHolidays(next.getFullYear());
    if (hols[`${next.getMonth() + 1}-${next.getDate()}`]) return true;
    return false;
  }

  // バケットごとに平日/休日のavailable日数をカウント
  const propSet = new Set(propNames);
  const totalRooms = propNames.reduce((s, pn) => { const p = findPropByName(pn); return s + (p ? (p.rooms || 1) : 1); }, 0);

  const results = bucketRanges.map(br => {
    let wdAvail = 0, hdAvail = 0;
    for (let d = new Date(br.startStr); d <= new Date(br.endStr); d.setDate(d.getDate() + 1)) {
      if (isHolidayNight(d)) { hdAvail += totalRooms; } else { wdAvail += totalRooms; }
    }
    return { ...br, wdAvail, hdAvail, wdNights: 0, hdNights: 0, wdSales: 0, hdSales: 0 };
  });

  // 予約データから未来の宿泊日を振り分け
  const processedDates = new Set(); // propCode|date で重複防止
  propNames.forEach(propName => {
    const cands = new Set();
    (window._resvByProp[propName] || []).forEach(r => cands.add(r));
    [...cands].filter(r => {
      if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
      return r.propCode === propName || r.property === propName;
    }).forEach(r => {
      if (!r.checkin || !r.checkout || !r.nights || r.nights <= 0) return;
      const netSales = (r.sales || 0) - (r.cleaningFee || 0);
      const dailyRate = netSales / r.nights;
      const ci = new Date(r.checkin);
      const co = new Date(r.checkout);
      for (let d = new Date(ci); d < co; d.setDate(d.getDate() + 1)) {
        const ds = d.toISOString().split('T')[0];
        if (ds <= todayStr) continue;
        const dedupKey = propName + '|' + ds;
        if (processedDates.has(dedupKey)) continue;
        processedDates.add(dedupKey);
        const isHol = isHolidayNight(d);
        // どのバケットに入るか
        const daysAhead = Math.floor((d - today) / 86400000);
        const bucket = results.find(br => daysAhead >= br.minDay && daysAhead <= br.maxDay);
        if (!bucket) continue;
        if (isHol) { bucket.hdNights++; bucket.hdSales += dailyRate; }
        else { bucket.wdNights++; bucket.wdSales += dailyRate; }
      }
    });
  });

  return results.map(br => ({
    label: br.label,
    weekday: {
      occ: br.wdAvail > 0 ? (br.wdNights / br.wdAvail) * 100 : 0,
      adr: br.wdNights > 0 ? Math.round(br.wdSales / br.wdNights) : 0,
      nights: br.wdNights, avail: br.wdAvail,
    },
    holiday: {
      occ: br.hdAvail > 0 ? (br.hdNights / br.hdAvail) * 100 : 0,
      adr: br.hdNights > 0 ? Math.round(br.hdSales / br.hdNights) : 0,
      nights: br.hdNights, avail: br.hdAvail,
    },
  }));
}

// ペースレポートHTML生成（物件詳細・全体横断の両方で使用）
function renderPaceReportHtml(paceData, title) {
  // 閾値: 休日OCC 85%超→値上げ余地, 平日OCC 30%未満→値下げ検討（0-30日バケット）
  const thresholds = [
    { hdHigh: 85, wdLow: 30 },  // 0-30日
    { hdHigh: 70, wdLow: 20 },  // 31-60日
    { hdHigh: 40, wdLow: 10 },  // 61-90日
  ];
  function badge(occ, idx, isHoliday) {
    const th = thresholds[idx] || thresholds[2];
    if (isHoliday && occ >= th.hdHigh) return ' <span style="color:#ff9500;font-size:10px;font-weight:600;">値上げ余地</span>';
    if (!isHoliday && occ > 0 && occ < th.wdLow) return ' <span style="color:#5856d6;font-size:10px;font-weight:600;">値下げ検討</span>';
    return '';
  }

  let html = `<div>
    <div style="font-size:13px;font-weight:600;color:#1d1d1f;margin-bottom:8px;">${title || '先行予約ペース'} <span style="font-size:11px;font-weight:400;color:#86868b;">（本日起点）</span></div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed;">
      <thead><tr style="border-bottom:2px solid #e5e5ea;">
        <th style="text-align:left;padding:5px 4px;color:#86868b;font-weight:500;width:22%;"></th>`;
  paceData.forEach(b => { html += `<th style="text-align:center;padding:5px 4px;color:#86868b;font-weight:500;" colspan="1">${b.label}</th>`; });
  html += `</tr></thead><tbody>`;

  // 休日OCC行
  html += `<tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:5px 4px;font-weight:500;">休日 OCC</td>`;
  paceData.forEach((b, i) => {
    const v = b.holiday.occ;
    const color = v >= (thresholds[i]||thresholds[2]).hdHigh ? '#ff9500' : '#1d1d1f';
    html += `<td style="text-align:center;padding:5px 4px;font-weight:600;color:${color};">${fmtPct(v)}${badge(v, i, true)}</td>`;
  });
  html += `</tr>`;

  // 平日OCC行
  html += `<tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:5px 4px;font-weight:500;">平日 OCC</td>`;
  paceData.forEach((b, i) => {
    const v = b.weekday.occ;
    const color = v > 0 && v < (thresholds[i]||thresholds[2]).wdLow ? '#5856d6' : '#1d1d1f';
    html += `<td style="text-align:center;padding:5px 4px;font-weight:600;color:${color};">${fmtPct(v)}${badge(v, i, false)}</td>`;
  });
  html += `</tr>`;

  // 休日ADR行
  html += `<tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:5px 4px;font-weight:500;">休日 ADR</td>`;
  paceData.forEach(b => { html += `<td style="text-align:center;padding:5px 4px;">${b.holiday.adr > 0 ? fmtYenFull(b.holiday.adr) : '-'}</td>`; });
  html += `</tr>`;

  // 平日ADR行
  html += `<tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:5px 4px;font-weight:500;">平日 ADR</td>`;
  paceData.forEach(b => { html += `<td style="text-align:center;padding:5px 4px;">${b.weekday.adr > 0 ? fmtYenFull(b.weekday.adr) : '-'}</td>`; });
  html += `</tr>`;

  // 泊数行
  html += `<tr><td style="padding:5px 4px;font-weight:500;">泊数</td>`;
  paceData.forEach(b => {
    const total = b.weekday.nights + b.holiday.nights;
    const avail = b.weekday.avail + b.holiday.avail;
    html += `<td style="text-align:center;padding:5px 4px;">${total}<span style="color:#86868b;font-size:10px;">/${avail}</span></td>`;
  });
  html += `</tr>`;

  html += `</tbody></table></div>`;
  return html;
}

// 平日/休日別のOCC・ADR・RevPAR を算出
// 休日 = 金曜泊・土曜泊 + 祝前日泊、平日 = それ以外
function computeWeekdayHolidayStats(propName, months) {
  const prop = findPropByName(propName);
  if (!prop) return null;
  const rooms = prop.rooms || 1;

  let wdNights = 0, wdSales = 0, wdAvail = 0;
  let hdNights = 0, hdSales = 0, hdAvail = 0;

  months.forEach(ym => {
    const [y, m] = ym.split('-').map(Number);
    const daysInMonth = getDaysInMonth(ym);
    const holidays = getJapaneseHolidays(y);
    // 翌月1日が祝日かチェック（月末が祝前日になる可能性）
    const nextMonthHolidays = m === 12 ? getJapaneseHolidays(y + 1) : holidays;
    const today = new Date().toISOString().split('T')[0];
    const monthStart = ym + '-01';
    const monthEnd = ym + '-' + String(daysInMonth).padStart(2, '0');

    // 各日が休日泊かどうかを判定
    function isHolidayNight(day) {
      const dt = new Date(y, m - 1, day);
      const dow = dt.getDay(); // 0=Sun
      // 金曜泊(dow=5)・土曜泊(dow=6)
      if (dow === 5 || dow === 6) return true;
      // 祝前日泊: 翌日が祝日
      const nextDay = day + 1;
      if (nextDay <= daysInMonth) {
        if (holidays[`${m}-${nextDay}`]) return true;
      } else {
        // 月末 → 翌月1日が祝日かチェック
        const nm = m === 12 ? 1 : m + 1;
        if (nextMonthHolidays[`${nm}-1`]) return true;
      }
      return false;
    }

    // 日ごとのavailableを平日/休日に振り分け
    for (let day = 1; day <= daysInMonth; day++) {
      if (isHolidayNight(day)) { hdAvail += rooms; } else { wdAvail += rooms; }
    }

    // 日次データから日別売上を取得（過去分）
    const dailySales = {}; // day -> sales
    const dailyDates = new Set();
    const propDailyData = (window._dailyByPropYm[propName + '|' + ym] || []).filter(d => {
      const status = d['状態'] || '';
      if (status === 'システムキャンセル' || status === 'ブロックされた') return false;
      const date = normalizeDate(d['日付']);
      if (date > today) return false;
      const cf = parseNum(d['清掃料']);
      const sl = parseNum(d['売上合計']);
      if (cf > 0 && Math.abs(sl - cf) < 1) return false;
      return true;
    });
    propDailyData.forEach(d => {
      const date = normalizeDate(d['日付']);
      const day = parseInt(date.split('-')[2], 10);
      const sales = parseNum(d['売上合計']);
      if (!dailySales[day]) dailySales[day] = 0;
      dailySales[day] += sales;
      dailyDates.add(date);
    });

    // 予約データから未来分を補完
    const _cands = new Set();
    (window._resvByProp[propName] || []).forEach(r => _cands.add(r));
    [..._cands].filter(r => {
      if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
      return r.propCode === propName || r.property === propName;
    }).forEach(r => {
      if (!r.checkin || !r.checkout || !r.nights || r.nights <= 0) return;
      const netSales = (r.sales || 0) - (r.cleaningFee || 0);
      const dailyRate = netSales / r.nights;
      const ci = new Date(r.checkin);
      const co = new Date(r.checkout);
      for (let d = new Date(ci); d < co; d.setDate(d.getDate() + 1)) {
        const ds = d.toISOString().split('T')[0];
        if (ds < monthStart || ds > monthEnd || ds <= today || dailyDates.has(ds)) continue;
        const day = d.getDate();
        if (!dailySales[day]) dailySales[day] = 0;
        dailySales[day] += dailyRate;
        dailyDates.add(ds);
      }
    });

    // 平日/休日に振り分け
    dailyDates.forEach(ds => {
      const day = parseInt(ds.split('-')[2], 10);
      const s = dailySales[day] || 0;
      if (isHolidayNight(day)) { hdNights++; hdSales += s; } else { wdNights++; wdSales += s; }
    });
  });

  const wdOcc = wdAvail > 0 ? (wdNights / wdAvail) * 100 : 0;
  const wdAdr = wdNights > 0 ? wdSales / wdNights : 0;
  const wdRevpar = wdAdr * (wdOcc / 100);
  const hdOcc = hdAvail > 0 ? (hdNights / hdAvail) * 100 : 0;
  const hdAdr = hdNights > 0 ? hdSales / hdNights : 0;
  const hdRevpar = hdAdr * (hdOcc / 100);

  return {
    weekday: { occ: wdOcc, adr: Math.round(wdAdr), revpar: Math.round(wdRevpar), nights: wdNights, avail: wdAvail },
    holiday: { occ: hdOcc, adr: Math.round(hdAdr), revpar: Math.round(hdRevpar), nights: hdNights, avail: hdAvail },
  };
}

function computeOverallStats(ym, areaFilter, excludeKpi) {
  return computeOverallStatsMulti([ym], areaFilter, excludeKpi);
}

function computeOverallStatsMulti(months, areaFilter, excludeKpi, extraFilters) {
  const { filteredProps } = aggregateDailyForMonth(months[0], areaFilter, excludeKpi, extraFilters);

  // Merge stats across months per property
  const propStatsMap = {};
  months.forEach(ym => {
    filteredProps.forEach(p => {
      const s = computePropertyStats(p.name, ym);
      if (!s) return;
      if (!propStatsMap[p.name]) {
        propStatsMap[p.name] = { ...s, _months: 1 };
      } else {
        propStatsMap[p.name].nights += s.nights;
        propStatsMap[p.name].sales += s.sales;
        propStatsMap[p.name].received += s.received;
        propStatsMap[p.name]._months++;
      }
    });
  });

  const stats = Object.values(propStatsMap);
  // Recalculate OCC/ADR/RevPAR from merged totals
  const totalDays = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);
  stats.forEach(s => {
    const totalAvail = totalDays * (filteredProps.find(p => p.name === s.name)?.rooms || 1);
    s.occ = totalAvail > 0 ? (s.nights / totalAvail) * 100 : 0;
    s.adr = s.nights > 0 ? s.sales / s.nights : 0;
    s.revpar = s.adr * (s.occ / 100);
  });

  const totalNights = stats.reduce((s, p) => s + p.nights, 0);
  const totalSales = stats.reduce((s, p) => s + p.sales, 0);
  const totalReceived = stats.reduce((s, p) => s + p.received, 0);
  const totalAvailable = filteredProps.reduce((s, p) => s + totalDays * (p.rooms || 1), 0);

  const occ = totalAvailable > 0 ? (totalNights / totalAvailable) * 100 : 0;
  const adr = totalNights > 0 ? totalSales / totalNights : 0;
  const revpar = adr * (occ / 100);

  return { occ, adr, revpar, totalSales, totalReceived, totalNights, totalAvailable, stats, propertyCount: filteredProps.length };
}

// ============================================================
// Format helpers
// ============================================================
function fmtYen(n) {
  if (n >= 1000000) return '¥' + (n / 1000000).toFixed(2) + 'M';
  if (n >= 10000) return '¥' + (n / 10000).toFixed(1) + '万';
  return '¥' + Math.round(n).toLocaleString();
}

function fmtYenFull(n) {
  return '¥' + Math.round(n).toLocaleString();
}

function fmtPct(n) {
  return n.toFixed(1) + '%';
}

// ============================================================
// Tab switching
// ============================================================
function switchTab(id) {
  currentTabId = id;
  document.querySelectorAll('.tab-btn').forEach((btn, i) => {
    const tabIds = ['daily','owner','property','reservation','revenue','review','watchlist','shinpou','pmbm','market'];
    btn.classList.toggle('active', tabIds[i] === id);
  });
  document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  // If this tab is dirty (data/filters changed since last render), re-render now
  if (dirtyTabs.has(id)) {
    const fn = tabRenderers[id];
    if (fn) {
      try { fn(); } catch (e) { console.error('[render]', id, e); }
    }
    dirtyTabs.delete(id);
    setTimeout(initSortableHeaders, 50);
  }
  setTimeout(() => initChartsForTab(id), 50);
}

// ============================================================
// Filter handling
// ============================================================
function setAreaFilter(el, tabId) {
  const pills = el.parentElement.querySelectorAll('.pill');
  pills.forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters[tabId + 'Area'] = el.dataset.area;
  renderAll();
  setTimeout(() => initChartsForTab(tabId), 50);
}

function setPropertyView(el) {
  const pills = el.parentElement.querySelectorAll('.pill');
  pills.forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.propertyView = el.dataset.view;
  document.getElementById('property-view-all').style.display = el.dataset.view === 'all' ? '' : 'none';
  document.getElementById('property-view-grouped').style.display = el.dataset.view === 'grouped' ? '' : 'none';
  renderAll();
}

function setPeriodFilter(el, tabId) {
  const pills = el.parentElement.querySelectorAll('.pill');
  pills.forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters[tabId + 'Period'] = el.dataset.period;
  renderAll();
  setTimeout(() => initChartsForTab(tabId), 50);
}

function setPropDetailPeriod(el, containerSelector, propertyName, prefix) {
  currentFilters.propDetailPeriod = el.dataset.period;
  const container = document.querySelector(containerSelector);
  if (container) renderPropertyDetail(container, propertyName, prefix);
}

// ============================================================
// Table sort
// ============================================================
const sortState = {};

function initSortableHeaders() {
  document.querySelectorAll('table thead th').forEach(th => {
    if (th.querySelector('.sort-icon')) return;
    const icon = document.createElement('span');
    icon.className = 'sort-icon';
    icon.textContent = '▲▼';
    th.appendChild(icon);
    th.addEventListener('click', () => handleSort(th));
  });
}

function handleSort(th) {
  const table = th.closest('table');
  const tbody = table.querySelector('tbody');
  if (!tbody || tbody.rows.length === 0) return;

  const colIdx = Array.from(th.parentElement.children).indexOf(th);
  const tableId = tbody.id || table.id || '';
  const stateKey = tableId + '_' + colIdx;

  // Toggle sort direction
  if (sortState[stateKey] === 'asc') {
    sortState[stateKey] = 'desc';
  } else if (sortState[stateKey] === 'desc') {
    sortState[stateKey] = null;
  } else {
    sortState[stateKey] = 'asc';
  }

  // Update icons
  th.parentElement.querySelectorAll('.sort-icon').forEach(ic => {
    ic.classList.remove('active');
    ic.textContent = '▲▼';
  });
  const icon = th.querySelector('.sort-icon');
  if (sortState[stateKey]) {
    icon.classList.add('active');
    icon.textContent = sortState[stateKey] === 'asc' ? '▲' : '▼';
  }

  if (!sortState[stateKey]) {
    // Reset: re-render to restore original order
    renderAll();
    setTimeout(initSortableHeaders, 50);
    return;
  }

  const rows = Array.from(tbody.querySelectorAll('tr'));
  // Skip totals row
  const totalsRow = rows.find(r => r.classList.contains('totals-row'));
  const sortableRows = rows.filter(r => !r.classList.contains('totals-row'));

  sortableRows.sort((a, b) => {
    const aCell = a.cells[colIdx];
    const bCell = b.cells[colIdx];
    if (!aCell || !bCell) return 0;
    let aVal = (aCell.textContent || '').trim();
    let bVal = (bCell.textContent || '').trim();

    // Try numeric comparison (strip ¥, %, 件, 泊, 名, M, 万, commas)
    const numA = parseFloat(aVal.replace(/[¥%件泊名万M,]/g, ''));
    const numB = parseFloat(bVal.replace(/[¥%件泊名万M,]/g, ''));
    if (!isNaN(numA) && !isNaN(numB)) {
      // Handle 万 and M multipliers
      let realA = numA, realB = numB;
      if (aVal.includes('万')) realA = numA * 10000;
      if (aVal.includes('M')) realA = numA * 1000000;
      if (bVal.includes('万')) realB = numB * 10000;
      if (bVal.includes('M')) realB = numB * 1000000;
      return sortState[stateKey] === 'asc' ? realA - realB : realB - realA;
    }

    // String comparison
    return sortState[stateKey] === 'asc' ? aVal.localeCompare(bVal, 'ja') : bVal.localeCompare(aVal, 'ja');
  });

  // Re-append in sorted order
  sortableRows.forEach(r => tbody.appendChild(r));
  if (totalsRow) tbody.appendChild(totalsRow);
}

// ============================================================
// Refresh / timestamp
// ============================================================
function updateTimestamp() {
  const now = new Date();
  const ts = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0') + '-' + String(now.getDate()).padStart(2,'0') + ' ' + String(now.getHours()).padStart(2,'0') + ':' + String(now.getMinutes()).padStart(2,'0');
  document.getElementById('lastUpdated').textContent = '最終更新: ' + ts;
}

function refreshData() {
  loadAllData();
}

// ============================================================
// KPI除外 toggle
// ============================================================
function toggleKpiExclude() {
  renderAll();
}

// ============================================================
// Render all (lazy: active tab now, others marked dirty + idle background)
// ============================================================
let currentTabId = 'daily'; // initial active tab in index.html
const dirtyTabs = new Set();
const tabRenderers = {
  daily: renderDailyTab,
  owner: renderOwnerTab,
  property: renderPropertyTab,
  reservation: renderReservationTab,
  revenue: renderRevenueTab,
  review: renderReviewTab,
  watchlist: renderWatchlistTab,
  shinpou: renderShinpouTab,
  pmbm: renderPmbmTab,
  market: renderMarketTab,
};
const ALL_TAB_IDS = Object.keys(tabRenderers);

function renderAll() {
  // 1. Render active tab immediately
  const fn = tabRenderers[currentTabId];
  if (fn) {
    try { fn(); } catch (e) { console.error('[render]', currentTabId, e); }
  }
  // 2. Mark all others dirty so they re-render on activation
  ALL_TAB_IDS.forEach(id => { if (id !== currentTabId) dirtyTabs.add(id); });
  setTimeout(initSortableHeaders, 50);
  // 3. Background-render dirty tabs during idle time
  scheduleIdleRender();
}

let idleRenderHandle = null;
function scheduleIdleRender() {
  if (idleRenderHandle != null) return;
  if (dirtyTabs.size === 0) return;
  const cb = (deadline) => {
    idleRenderHandle = null;
    const ids = Array.from(dirtyTabs);
    for (const id of ids) {
      // Yield if running out of idle budget (Chart.js renders are heavy)
      if (deadline && deadline.timeRemaining && deadline.timeRemaining() < 8) break;
      if (id === currentTabId) { dirtyTabs.delete(id); continue; }
      const fn = tabRenderers[id];
      if (fn) {
        try { fn(); } catch (e) { console.error('[idle render]', id, e); }
      }
      dirtyTabs.delete(id);
    }
    if (dirtyTabs.size > 0) scheduleIdleRender();
  };
  if (typeof requestIdleCallback === 'function') {
    idleRenderHandle = requestIdleCallback(cb, { timeout: 3000 });
  } else {
    idleRenderHandle = setTimeout(() => cb({ timeRemaining: () => 50 }), 300);
  }
}

// ============================================================
// Tab 1: TOP
// ============================================================
function shiftMonths(months, deltaMonths) {
  return months.map(ym => {
    const [y, m] = ym.split('-').map(Number);
    const d = new Date(y, m - 1 + deltaMonths, 1);
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  });
}

function computeDailyMetrics(months, area) {
  const monthSet = new Set(months);
  const totalDays = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);

  // チェックイン月ベース予約
  const monthResvs = reservations.filter(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return false;
    if (!monthSet.has(getYearMonth(r.checkin))) return false;
    if (area !== '全体') {
      const prop = findPropByReservation(r);
      if (!prop || prop.area !== area) return false;
    }
    return true;
  });
  const sales = monthResvs.reduce((s, r) => s + (r.sales || 0), 0);
  const received = monthResvs.reduce((s, r) => s + (r.received || 0), 0);
  const avgDailySales = totalDays > 0 ? sales / totalDays : 0;
  const avgNights = monthResvs.length > 0 ? monthResvs.reduce((s, r) => s + r.nights, 0) / monthResvs.length : 0;

  // PM売上: (販売 - OTA手数料 - 清掃費) × オーナーロイヤリティ%
  let pmSales = 0;
  monthResvs.forEach(r => {
    const prop = findPropByReservation(r);
    if (!prop) return;
    const royaltyPct = prop.royaltyPct || 0;
    if (royaltyPct === 0) return;
    const base = (r.sales || 0) - (r.otaFee || 0) - (r.cleaningFee || 0);
    pmSales += base * (royaltyPct / 100);
  });

  // BM売上: 対象月チェックアウト予約の清掃費合計
  let bmSales = 0;
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    if (!monthSet.has(getYearMonth(r.checkout))) return;
    if (area !== '全体') {
      const prop = findPropByReservation(r);
      if (!prop || prop.area !== area) return;
    }
    bmSales += (r.cleaningFee || 0);
  });

  const overall = computeOverallStatsMulti(months, area, false);

  // 目標達成物件数 / オーナー数
  const filteredProps = filterPropertiesByArea(area).filter(p => !p.excludeKpi && p.status === '稼働中');
  let hitCount = 0;
  const ownerAgg = {};
  filteredProps.forEach(p => {
    let actual = 0, target = 0;
    months.forEach(ym => {
      const s = computePropertyStats(p.name, ym);
      if (s) actual += s.sales;
      const monthNum = parseInt(ym.split('-')[1], 10);
      target += getTargetForProperty(p, monthNum) || 0;
    });
    if (target > 0 && actual >= target) hitCount++;
    if (p.ownerId) {
      if (!ownerAgg[p.ownerId]) ownerAgg[p.ownerId] = { actual: 0, target: 0 };
      ownerAgg[p.ownerId].actual += actual;
      ownerAgg[p.ownerId].target += target;
    }
  });
  let ownerHit = 0;
  Object.values(ownerAgg).forEach(o => {
    if (o.target > 0 && o.actual >= o.target) ownerHit++;
  });

  return {
    sales, received, avgDailySales, avgNights,
    pmSales, bmSales,
    adr: overall.adr, occ: overall.occ,
    hitCount, totalProps: filteredProps.length,
    ownerHit, totalOwners: Object.keys(ownerAgg).length,
  };
}

// 直近N日の新規予約数（予約日ベース）
function countNewBookings(area, daysBack, endDateStr) {
  const end = new Date(endDateStr);
  const start = new Date(end);
  start.setDate(start.getDate() - daysBack + 1);
  const startStr = localDateStr(start);
  let count = 0;
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    if (!r.date || r.date < startStr || r.date > endDateStr) return;
    if (area !== '全体') {
      const prop = findPropByReservation(r);
      if (!prop || prop.area !== area) return;
    }
    count++;
  });
  return count;
}

// vs ratio formatter (% diff)
function fmtVsLine(cur, py, pm, formatter) {
  const fmtOne = (label, prev) => {
    if (prev == null || prev === 0) return `${label} -`;
    const diff = cur - prev;
    const pct = (diff / prev) * 100;
    const sign = diff >= 0 ? '+' : '';
    const cls = diff >= 0 ? 'positive' : 'negative';
    return `<span class="${cls}">${label} ${sign}${pct.toFixed(1)}%</span>`;
  };
  return `${fmtOne('YoY', py)} / ${fmtOne('MoM', pm)}`;
}

// pt diff formatter (for OCC etc)
function fmtVsLinePt(cur, py, pm) {
  const fmtOne = (label, prev) => {
    if (prev == null) return `${label} -`;
    const diff = cur - prev;
    const sign = diff >= 0 ? '+' : '';
    const cls = diff >= 0 ? 'positive' : 'negative';
    return `<span class="${cls}">${label} ${sign}${diff.toFixed(1)}pt</span>`;
  };
  return `${fmtOne('YoY', py)} / ${fmtOne('MoM', pm)}`;
}

function renderDailyTab() {
  const months = getSelectedMonths('daily');
  const area = currentFilters.dailyArea;

  // Date range display
  const firstMonth = months[0];
  const lastMonth = months[months.length - 1];
  const firstDay = firstMonth + '-01';
  const lastDaysInMonth = getDaysInMonth(lastMonth);
  const lastDay = lastMonth + '-' + String(lastDaysInMonth).padStart(2, '0');
  document.getElementById('daily-date-range').textContent = firstDay + ' ~ ' + lastDay;

  const cur = computeDailyMetrics(months, area);
  const py = computeDailyMetrics(shiftMonths(months, -12), area);
  const pm = computeDailyMetrics(shiftMonths(months, -1), area);

  // Primary KPIs
  document.getElementById('kpi-daily-sales').textContent = fmtYen(cur.sales);
  document.getElementById('kpi-daily-sales-vs').innerHTML = fmtVsLine(cur.sales, py.sales, pm.sales);
  document.getElementById('kpi-daily-received').textContent = fmtYen(cur.received);
  document.getElementById('kpi-daily-received-vs').innerHTML = fmtVsLine(cur.received, py.received, pm.received);
  document.getElementById('kpi-daily-pm').textContent = fmtYen(cur.pmSales);
  document.getElementById('kpi-daily-pm-vs').innerHTML = fmtVsLine(cur.pmSales, py.pmSales, pm.pmSales);
  document.getElementById('kpi-daily-bm').textContent = fmtYen(cur.bmSales);
  document.getElementById('kpi-daily-bm-vs').innerHTML = fmtVsLine(cur.bmSales, py.bmSales, pm.bmSales);
  document.getElementById('kpi-daily-avg').textContent = fmtYenFull(Math.round(cur.avgDailySales));
  document.getElementById('kpi-daily-avg-vs').innerHTML = fmtVsLine(cur.avgDailySales, py.avgDailySales, pm.avgDailySales);
  document.getElementById('kpi-daily-adr').textContent = fmtYenFull(Math.round(cur.adr));
  document.getElementById('kpi-daily-adr-vs').innerHTML = fmtVsLine(cur.adr, py.adr, pm.adr);
  document.getElementById('kpi-daily-occ').textContent = fmtPct(cur.occ);
  document.getElementById('kpi-daily-occ-vs').innerHTML = fmtVsLinePt(cur.occ, py.occ, pm.occ);
  document.getElementById('kpi-daily-nights').textContent = cur.avgNights.toFixed(1) + '泊';
  document.getElementById('kpi-daily-nights-vs').innerHTML = fmtVsLine(cur.avgNights, py.avgNights, pm.avgNights);

  // 目標達成物件数 / オーナー数
  document.getElementById('kpi-daily-target-hit').textContent = cur.hitCount + '件';
  document.getElementById('kpi-daily-target-hit-sub').textContent = cur.totalProps > 0 ? `${cur.totalProps}件中 (${(cur.hitCount / cur.totalProps * 100).toFixed(0)}%)` : '-';
  document.getElementById('kpi-daily-target-hit-vs').innerHTML = fmtVsLine(cur.hitCount, py.hitCount, pm.hitCount);
  document.getElementById('kpi-daily-target-hit-owner').textContent = cur.ownerHit + '名';
  document.getElementById('kpi-daily-target-hit-owner-sub').textContent = cur.totalOwners > 0 ? `${cur.totalOwners}名中 (${(cur.ownerHit / cur.totalOwners * 100).toFixed(0)}%)` : '-';
  document.getElementById('kpi-daily-target-hit-owner-vs').innerHTML = fmtVsLine(cur.ownerHit, py.ownerHit, pm.ownerHit);

  // 新規予約 (直近7日) - 対前週(直前7日) / 対前年(同期間1年前)
  const todayStr = localDateStr(new Date());
  const newBookCount = countNewBookings(area, 7, todayStr);
  // 前週(直前7日: 14日前〜8日前)
  const prevWeekEnd = new Date();
  prevWeekEnd.setDate(prevWeekEnd.getDate() - 7);
  const prevWeekCount = countNewBookings(area, 7, localDateStr(prevWeekEnd));
  // 前年同期(7日前と同じ7日窓を1年シフト)
  const prevYearEnd = new Date();
  prevYearEnd.setFullYear(prevYearEnd.getFullYear() - 1);
  const prevYearWeekCount = countNewBookings(area, 7, localDateStr(prevYearEnd));
  document.getElementById('kpi-daily-newbooking').textContent = newBookCount + '件';
  document.getElementById('kpi-daily-newbooking-vs').innerHTML = (function () {
    const fmtOne = (label, prev) => {
      if (prev == null || prev === 0) return `${label} -`;
      const diff = newBookCount - prev;
      const pct = (diff / prev) * 100;
      const sign = diff >= 0 ? '+' : '';
      const cls = diff >= 0 ? 'positive' : 'negative';
      return `<span class="${cls}">${label} ${sign}${pct.toFixed(0)}%</span>`;
    };
    return `${fmtOne('YoY', prevYearWeekCount)} / ${fmtOne('WoW', prevWeekCount)}`;
  })();

  // Charts
  initDailyCharts();
}

// ============================================================
// Tab 2: オーナー別分析
// ============================================================
function renderOwnerTab() {
  renderOrphanAlert('owner-orphan-alert');
  const months = getSelectedMonths('owner');
  const area = currentFilters.ownerArea;

  // Filter owners by area
  let filteredOwners = owners;
  if (area !== '全体') {
    filteredOwners = owners.filter(o => {
      return o.properties.some(pn => {
        const p = findPropByName(pn);
        return p && p.area === area;
      });
    });
  }

  // Compute stats per owner (merged across months)
  const ownerStats = filteredOwners.map(o => {
    const ownerProps = properties.filter(p => p.ownerId === o.id);
    const relevantProps = area !== '全体' ? ownerProps.filter(p => p.area === area) : ownerProps;

    let totalSales = 0, totalNights = 0, target = 0;
    const allPropStats = {};
    months.forEach(ym => {
      const monthNum = getMonthNumber(ym);
      relevantProps.forEach(p => {
        const s = computePropertyStats(p.name, ym);
        if (!s) return;
        if (!allPropStats[p.name]) {
          allPropStats[p.name] = { ...s };
        } else {
          allPropStats[p.name].sales += s.sales;
          allPropStats[p.name].nights += s.nights;
          allPropStats[p.name].received += s.received;
        }
      });
      target += relevantProps.reduce((s, p) => s + getTargetForProperty(p, monthNum), 0);
    });

    const propStats = Object.values(allPropStats);
    totalSales = propStats.reduce((s, p) => s + p.sales, 0);
    totalNights = propStats.reduce((s, p) => s + p.nights, 0);

    // Recalc OCC per property across months
    const totalDays = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);
    propStats.forEach(ps => {
      const prop = relevantProps.find(p => p.name === ps.name);
      const avail = totalDays * (prop?.rooms || 1);
      ps.occ = avail > 0 ? (ps.nights / avail) * 100 : 0;
      ps.adr = ps.nights > 0 ? ps.sales / ps.nights : 0;
      ps.revpar = ps.adr * (ps.occ / 100);
    });

    const avgOcc = propStats.length > 0 ? propStats.reduce((s, p) => s + p.occ, 0) / propStats.length : 0;
    const avgAdr = totalNights > 0 ? totalSales / totalNights : 0;
    const rate = target > 0 ? (totalSales / target) * 100 : 0;

    return {
      ...o,
      propCount: relevantProps.length,
      target,
      actual: totalSales,
      rate,
      avgOcc,
      avgAdr,
      propStats,
    };
  });

  // 目標未設定オーナーアラート
  const noTargetAlert = document.getElementById('owner-no-target-alert');
  if (noTargetAlert) {
    const noTargetOwners = ownerStats.filter(o => o.target === 0 && o.propCount > 0);
    if (noTargetOwners.length > 0) {
      const items = noTargetOwners.map(o => `<li>${o.name} <span style="color:#999;">(${o.propCount}物件)</span></li>`).join('');
      const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${SHEET_GID_PROPERTY_MASTER}`;
      noTargetAlert.innerHTML = `<div class="alert-orphan">
        <div class="alert-title">⚠ 目標が未設定のオーナーが ${noTargetOwners.length} 名います</div>
        物件マスタで目標売上（閑散期・通常期・繁忙期）を設定してください。<a href="${sheetUrl}" target="_blank" rel="noopener" style="color:#ff3b30;font-weight:600;text-decoration:underline;">物件マスタを開く ↗</a>
        <ul>${items}</ul>
      </div>`;
    } else {
      noTargetAlert.innerHTML = '';
    }
  }

  // KPIs
  document.getElementById('kpi-owner-count').textContent = filteredOwners.length + '名';
  const avgRate = ownerStats.length > 0 ? ownerStats.reduce((s, o) => s + o.rate, 0) / ownerStats.length : 0;
  document.getElementById('kpi-owner-avg-rate').textContent = fmtPct(avgRate);
  const underCount = ownerStats.filter(o => o.rate < 100).length;
  document.getElementById('kpi-owner-under').innerHTML = underCount + '名' + (underCount > 0 ? ' <span class="badge-orange">要確認</span>' : '');

  // 達成率プログレスバー一覧（独立した期間ピルで集計、達成率の低い順）
  const progressList = document.getElementById('owner-progress-list');
  if (progressList) {
    const progressMonths = getSelectedMonths('ownerProgress');
    const progressStats = filteredOwners.map(o => {
      const ownerProps2 = properties.filter(p => p.ownerId === o.id);
      const relevantProps2 = area !== '全体' ? ownerProps2.filter(p => p.area === area) : ownerProps2;
      let totalSales2 = 0, target2 = 0;
      const allPropStats2 = {};
      progressMonths.forEach(ym => {
        const monthNum = getMonthNumber(ym);
        relevantProps2.forEach(p => {
          const s = computePropertyStats(p.name, ym);
          if (!s) return;
          if (!allPropStats2[p.name]) allPropStats2[p.name] = { ...s };
          else allPropStats2[p.name].sales += s.sales;
        });
        target2 += relevantProps2.reduce((s, p) => s + getTargetForProperty(p, monthNum), 0);
      });
      totalSales2 = Object.values(allPropStats2).reduce((s, p) => s + p.sales, 0);
      const rate2 = target2 > 0 ? (totalSales2 / target2) * 100 : 0;
      return { id: o.id, name: o.name, propCount: relevantProps2.length, target: target2, actual: totalSales2, rate: rate2 };
    });
    const sortedByRate = [...progressStats].sort((a, b) => a.rate - b.rate);
    progressList.innerHTML = sortedByRate.map(o => {
      const barClass = o.rate >= 100 ? 'progress-green' : o.rate >= 70 ? 'progress-orange' : 'progress-red';
      const barWidth = Math.min(o.rate, 100);
      const rateClass = o.rate >= 100 ? 'positive' : o.rate >= 70 ? '' : 'negative';
      return `<div class="clickable" onclick="toggleOwnerDrill('${o.id}', null)" style="display:flex;align-items:center;gap:12px;padding:8px 0;border-bottom:1px solid #f0f0f0;">
        <div style="width:140px;font-size:13px;font-weight:600;flex-shrink:0;">${o.name}</div>
        <div style="width:50px;font-size:11px;color:#999;flex-shrink:0;">${o.propCount}物件</div>
        <div style="flex:1;min-width:120px;"><div class="progress-bar-bg">${barWidth > 0 ? `<div class="progress-bar-fill ${barClass}" style="width:${barWidth}%"></div>` : ''}</div></div>
        <div class="${rateClass}" style="width:60px;text-align:right;font-size:13px;font-weight:600;flex-shrink:0;">${fmtPct(o.rate)}</div>
        <div style="width:200px;text-align:right;font-size:11px;color:#666;flex-shrink:0;">${fmtYen(o.actual)} / ${fmtYen(o.target)}</div>
      </div>`;
    }).join('') || '<div style="color:#999;text-align:center;padding:12px;">対象オーナーがありません</div>';
  }

  // Table
  const tbody = document.getElementById('owner-table');
  tbody.innerHTML = ownerStats.map(o => {
    const rateClass = o.rate >= 100 ? 'positive' : o.rate >= 70 ? '' : 'negative';
    return `<tr class="clickable" onclick="toggleOwnerDrill('${o.id}', this)">
      <td>${o.name}</td><td>${o.propCount}物件</td><td>${o.royalty}</td>
      <td>${fmtYen(o.target)}</td><td>${fmtYen(o.actual)}</td>
      <td class="${rateClass}">${fmtPct(o.rate)}</td>
      <td>${fmtPct(o.avgOcc)}</td><td>${fmtYenFull(Math.round(o.avgAdr))}</td>
    </tr>`;
  }).join('');
}

// ============================================================
// Owner drill-down
// ============================================================
let activeOwnerDrill = null;
function toggleOwnerDrill(ownerId, clickedRow) {
  // 既存のドリルダウン行を削除
  const existing = document.getElementById('owner-drill-row');
  if (existing) existing.remove();

  // 同じオーナーをクリック → 閉じるだけ
  if (activeOwnerDrill === ownerId) {
    activeOwnerDrill = null;
    return;
  }
  activeOwnerDrill = ownerId;

  const months = getSelectedMonths('owner');
  const owner = owners.find(o => o.id === ownerId);
  if (!owner) return;

  const ownerProps = properties.filter(p => p.ownerId === ownerId);
  // Merge stats across months
  const propStatsMap = {};
  let target = 0;
  months.forEach(ym => {
    const monthNum = getMonthNumber(ym);
    ownerProps.forEach(p => {
      const s = computePropertyStats(p.name, ym);
      if (!s) return;
      if (!propStatsMap[p.name]) {
        propStatsMap[p.name] = { ...s };
      } else {
        propStatsMap[p.name].sales += s.sales;
        propStatsMap[p.name].nights += s.nights;
        propStatsMap[p.name].received += s.received;
      }
    });
    target += ownerProps.reduce((s, p) => s + getTargetForProperty(p, monthNum), 0);
  });
  const totalDays = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);
  const propStats = Object.values(propStatsMap);
  propStats.forEach(ps => {
    const prop = ownerProps.find(p => p.name === ps.name);
    const avail = totalDays * (prop?.rooms || 1);
    ps.occ = avail > 0 ? (ps.nights / avail) * 100 : 0;
    ps.adr = ps.nights > 0 ? ps.sales / ps.nights : 0;
    ps.revpar = ps.adr * (ps.occ / 100);
  });
  const totalSales = propStats.reduce((s, p) => s + p.sales, 0);
  const rate = target > 0 ? (totalSales / target) * 100 : 0;
  const barColor = rate >= 100 ? 'progress-green' : rate >= 70 ? 'progress-orange' : 'progress-red';
  const barWidth = Math.min(rate, 100);

  let propRows = propStats.map(p => {
    const prop = findPropByName(p.name);
    return `<tr class="clickable" onclick="event.stopPropagation();toggleOwnerPropertyDrill('${p.name}')">
      <td>${p.name}</td><td>${p.area}</td><td>${fmtPct(p.occ)}</td><td>${fmtYenFull(Math.round(p.adr))}</td><td>${fmtYenFull(Math.round(p.revpar))}</td><td>${fmtYenFull(p.sales)}</td><td>${fmtYenFull(p.received)}</td><td>${prop && prop.excludeKpi ? '<span class="badge-gray">除外</span>' : '-'}</td>
    </tr>`;
  }).join('');

  // クリックした行のすぐ下にドリルダウン行を挿入
  const drillRow = document.createElement('tr');
  drillRow.id = 'owner-drill-row';
  const drillCell = document.createElement('td');
  drillCell.colSpan = 8;
  drillCell.style.padding = '0';
  drillRow.appendChild(drillCell);

  if (clickedRow) {
    clickedRow.insertAdjacentElement('afterend', drillRow);
  } else {
    document.getElementById('owner-table').appendChild(drillRow);
  }

  // 物件詳細と同じスタイルで合算表示
  // オーナー所属の全予約を取得
  const ownerPropNames = new Set(ownerProps.map(p => p.name));
  const ownerPropPropNames = new Set(ownerProps.map(p => p.propName).filter(Boolean));
  const ownerResvAll = reservations.filter(r => ownerPropNames.has(r.propCode) || ownerPropNames.has(r.property) || ownerPropPropNames.has(r.property));
  const ownerResv = ownerResvAll.slice(0, 10);
  let ownerResvRows = ownerResv.map(r => `<tr><td>${(r.date || '').slice(0, 10)}</td><td>${r.channel}</td><td>${r.property}</td><td>${r.guest}</td><td>${r.checkin}</td><td>${r.checkout}</td><td>${r.nights}泊</td><td>${fmtYenFull(r.sales)}</td><td>${r.status}</td></tr>`).join('');
  if (!ownerResvRows) ownerResvRows = '<tr><td colspan="9" style="color:#999;text-align:center;">データなし</td></tr>';

  destroyDrillCharts('own');

  drillCell.innerHTML = `<div class="drill-down show" style="margin-top:12px;">
    <h3>${owner.name} <span style="font-size:13px;color:#666;font-weight:400;">(${ownerProps.length}物件)</span></h3>
    <div class="progress-bar-wrap">
      <div class="progress-bar-label"><span>目標: ${fmtYen(target)}</span><span>実績: ${fmtYen(totalSales)} (${fmtPct(rate)})</span></div>
      <div class="progress-bar-bg"><div class="progress-bar-fill ${barColor}" style="width:${barWidth}%"></div></div>
    </div>
    <div class="chart-grid">
      <div class="card"><h2>月別 販売金額/OCC推移（合算）</h2><canvas id="ownChartSalesOcc"></canvas></div>
      <div class="card"><h2>月別 販売金額/ADR推移（合算）</h2><canvas id="ownChartSalesAdr"></canvas></div>
      <div class="card"><h2>チャネル別売上構成比（合算）</h2><canvas id="ownChartChannel"></canvas></div>
      <div class="card"><h2>ゲスト国籍別（合算）</h2><canvas id="ownChartNationality"></canvas></div>
      <div class="card" id="ownRecentBookings"></div>
    </div>
    <div class="card"><h2>予約一覧（直近10件・合算）</h2><div class="table-wrap"><table>
      <thead><tr><th>予約日</th><th>予約サイト</th><th>物件名</th><th>ゲスト名</th><th>チェックイン</th><th>チェックアウト</th><th>泊数</th><th>販売金額</th><th>状態</th></tr></thead>
      <tbody>${ownerResvRows}</tbody>
    </table></div></div>
    <div class="card"><h2>物件別内訳</h2><div class="table-wrap"><table>
      <thead><tr><th>物件名</th><th>エリア</th><th>OCC</th><th>ADR</th><th>RevPAR</th><th>販売金額</th><th>受取金</th><th>KPI除外</th></tr></thead>
      <tbody>${propRows}</tbody>
    </table></div></div>
    <div id="owner-property-drill-container"></div>
  </div>`;

  renderOwnerDetailCharts(ownerProps, ownerResvAll, 'own');

  setTimeout(initSortableHeaders, 50);
  drillRow.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// オーナー合算チャート（物件詳細と同じ構成、複数物件の合算）
function renderOwnerDetailCharts(ownerProps, ownerResvAll, prefix) {
  setTimeout(() => {
    const now = new Date();
    const curMonth = now.getMonth() + 1;
    const curYear = now.getFullYear();
    const monthLabels = [];
    const occData = [];
    const adrData = [];
    const salesData = [];
    const targetData = [];
    const totalRooms = ownerProps.reduce((s, p) => s + (p.rooms || 1), 0);

    for (let i = -5; i <= 6; i++) {
      const d = new Date(curYear, curMonth - 1 + i, 1);
      const m = d.getMonth() + 1;
      const y = d.getFullYear();
      const mym = `${y}-${String(m).padStart(2, '0')}`;
      monthLabels.push(`${m}月`);

      // 月内の全物件合算
      let mNights = 0, mSales = 0;
      ownerProps.forEach(p => {
        const s = computePropertyStats(p.name, mym);
        if (!s) return;
        mNights += s.nights;
        mSales += s.sales;
      });
      const days = getDaysInMonth(mym);
      const avail = days * totalRooms;
      const occ = avail > 0 ? (mNights / avail) * 100 : 0;
      const adr = mNights > 0 ? mSales / mNights : 0;
      occData.push(Math.round(occ * 10) / 10);
      adrData.push(Math.round(adr));
      salesData.push(Math.round(mSales));
      // 目標売上もオーナー全物件合算
      targetData.push(ownerProps.reduce((s, p) => s + getTargetForProperty(p, m), 0));
    }
    const targetLineDataset = {
      type: 'line', label: '目標売上', data: targetData,
      borderColor: '#ff3b30', backgroundColor: 'transparent',
      borderDash: [6, 4], borderWidth: 2, pointRadius: 0,
      pointHoverRadius: 4, pointBackgroundColor: '#ff3b30',
      tension: 0, fill: false, yAxisID: 'y',
    };
    const currentIdx = 5;

    const blueBarColors = salesData.map((_, i) => i < currentIdx ? 'rgba(74,144,217,0.2)' : i === currentIdx ? 'rgba(74,144,217,0.5)' : 'rgba(74,144,217,0.1)');
    const orangeBarColors = salesData.map((_, i) => i < currentIdx ? 'rgba(245,166,35,0.2)' : i === currentIdx ? 'rgba(245,166,35,0.5)' : 'rgba(245,166,35,0.1)');
    const barBorders = salesData.map((_, i) => i > currentIdx ? 'rgba(0,0,0,0.06)' : 'transparent');
    const barBorderWidths = salesData.map((_, i) => i > currentIdx ? 1 : 0);

    // Sales/OCC chart
    const ctx1 = document.getElementById(prefix + 'ChartSalesOcc');
    if (ctx1) {
      chartInstances[prefix + 'SalesOcc'] = new Chart(ctx1, {
        type: 'bar',
        data: {
          labels: monthLabels,
          datasets: [
            { type: 'line', label: 'OCC (%)', data: occData, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.08)', fill: true, yAxisID: 'y1', tension: 0.4, pointBorderColor: CHART_COLORS.blue },
            { type: 'bar', label: '販売金額', data: salesData, backgroundColor: blueBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' },
            { ...targetLineDataset }
          ]
        },
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true }, tooltip: salesChartTooltip }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v / 10000).toFixed(0) + '万' } }, y1: { position: 'right', min: 0, max: 110, grid: { drawOnChartArea: false }, title: { display: true, text: 'OCC (%)', font: { size: 11 } } } } }
      });
    }

    // Sales/ADR chart
    const ctx1b = document.getElementById(prefix + 'ChartSalesAdr');
    if (ctx1b) {
      chartInstances[prefix + 'SalesAdr'] = new Chart(ctx1b, {
        type: 'bar',
        data: {
          labels: monthLabels,
          datasets: [
            { type: 'line', label: 'ADR (¥)', data: adrData, borderColor: CHART_COLORS.orange, backgroundColor: 'rgba(245,166,35,0.08)', fill: true, yAxisID: 'y1', tension: 0.4, pointBorderColor: CHART_COLORS.orange },
            { type: 'bar', label: '販売金額', data: salesData, backgroundColor: orangeBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' },
            { ...targetLineDataset }
          ]
        },
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true }, tooltip: salesChartTooltip }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v / 10000).toFixed(0) + '万' } }, y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
      });
    }

    // Channel breakdown (current month, aggregated)
    const ym = getSelectedMonth('owner');
    const channelAgg = {};
    ownerProps.forEach(p => {
      const s = computePropertyStats(p.name, ym);
      if (!s || !s.channels) return;
      Object.entries(s.channels).forEach(([ch, data]) => {
        if (!channelAgg[ch]) channelAgg[ch] = { count: 0, sales: 0 };
        channelAgg[ch].count += data.count;
        channelAgg[ch].sales += data.sales;
      });
    });
    const channelColors = PALETTE;
    const totalChSales = Object.values(channelAgg).reduce((s, c) => s + c.sales, 0);
    const sortedCh = Object.entries(channelAgg).sort((a, b) => b[1].sales - a[1].sales);
    const channelLabels = sortedCh.map(([k]) => k);
    const channelPct = sortedCh.map(([, v]) => totalChSales > 0 ? (v.sales / totalChSales) * 100 : 0);
    const ctx2 = document.getElementById(prefix + 'ChartChannel');
    if (ctx2 && channelLabels.length > 0) {
      chartInstances[prefix + 'Channel'] = new Chart(ctx2, {
        type: 'bar',
        data: { labels: channelLabels, datasets: [{ label: '売上構成比', data: channelPct, backgroundColor: channelColors.slice(0, channelLabels.length).map(c => c + 'CC'), hoverBackgroundColor: channelColors.slice(0, channelLabels.length) }] },
        options: { indexAxis: 'y', responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.parsed.x.toFixed(1) + '%' } } }, scales: { x: { beginAtZero: true, max: 100, grid: { display: false }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false } } } }
      });
    }

    // Nationality breakdown (aggregated)
    const natAgg = {};
    ownerResvAll.forEach(r => {
      if (r.status === 'システムキャンセル') return;
      const nat = r.nationality || '不明';
      if (!natAgg[nat]) natAgg[nat] = { count: 0, sales: 0 };
      natAgg[nat].count++;
      natAgg[nat].sales += r.sales || 0;
    });
    const natTotal = Object.values(natAgg).reduce((s, v) => s + v.count, 0);
    const natSorted = Object.entries(natAgg).sort((a, b) => b[1].count - a[1].count);
    const natTop = natSorted.slice(0, 5);
    const natOthers = natSorted.slice(5);
    const natOtherCount = natOthers.reduce((s, [, v]) => s + v.count, 0);
    if (natOtherCount > 0) natTop.push(['その他', { count: natOtherCount, sales: 0 }]);
    const natLabels = natTop.map(([k]) => k);
    const natPct = natTop.map(([, v]) => natTotal > 0 ? (v.count / natTotal) * 100 : 0);
    const ctxNat = document.getElementById(prefix + 'ChartNationality');
    if (ctxNat && natLabels.length > 0) {
      chartInstances[prefix + 'Nationality'] = new Chart(ctxNat, {
        type: 'bar',
        data: { labels: natLabels, datasets: [{ label: '予約数構成比', data: natPct, backgroundColor: channelColors.slice(0, natLabels.length).map(c => c + 'CC'), hoverBackgroundColor: channelColors.slice(0, natLabels.length) }] },
        options: { indexAxis: 'y', responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: false }, tooltip: { callbacks: {
          title: ctx => ctx[0].label,
          afterBody: ctx => {
            if (ctx[0].label === 'その他') {
              return natOthers.map(([k, v]) => `  ${k}: ${(natTotal > 0 ? (v.count / natTotal) * 100 : 0).toFixed(1)}% (${v.count}件)`);
            }
            return [];
          },
          label: ctx => ctx.parsed.x.toFixed(1) + '% (' + natTop[ctx.dataIndex][1].count + '件)'
        } } }, scales: { x: { beginAtZero: true, max: 100, grid: { display: false }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false } } } }
      });
    }

    // Recent bookings (last 30 days, aggregated)
    const recentBox = document.getElementById(prefix + 'RecentBookings');
    if (recentBox) {
      const nowDate = new Date();
      const recent = ownerResvAll.filter(r => {
        if (r.status === 'システムキャンセル') return false;
        if (!r.date) return false;
        const diffMs = nowDate - new Date(r.date);
        return diffMs >= 0 && diffMs < 30 * 86400000;
      });
      const buckets = [
        { label: '直近3日', min: 0, max: 3, count: 0 },
        { label: '4〜7日前', min: 3, max: 7, count: 0 },
        { label: '8〜14日前', min: 7, max: 14, count: 0 },
        { label: '15〜30日前', min: 14, max: 30, count: 0 },
      ];
      recent.forEach(r => {
        const days = Math.floor((nowDate - new Date(r.date)) / 86400000);
        buckets.forEach(b => { if (days >= b.min && days < b.max) b.count++; });
      });
      const total = recent.length;
      // 物件数で閾値を比例調整（物件数が多いオーナーは閾値も高くなる）
      const propCount = ownerProps.length;
      const threshGood = Math.max(5, propCount * 2);
      const threshWarn = Math.max(2, propCount);
      const statusColor = total >= threshGood ? '#34c759' : total >= threshWarn ? '#ff9500' : '#ff3b30';
      const statusText = total >= threshGood ? '好調' : total >= threshWarn ? '注意' : '要確認';
      const barsHtml = buckets.map(b => {
        const pct = total > 0 ? (b.count / total) * 100 : 0;
        const barColor = b.count > 0 ? CHART_COLORS.blue : '#e5e5e5';
        return `<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
          <span style="font-size:12px;color:#666;width:80px;text-align:right;flex-shrink:0;">${b.label}</span>
          <div style="flex:1;background:#f0f0f0;border-radius:4px;height:24px;position:relative;overflow:hidden;">
            <div style="width:${pct}%;background:${barColor};height:100%;border-radius:4px;transition:width 0.3s;min-width:${b.count > 0 ? '2px' : '0'}"></div>
          </div>
          <span style="font-size:13px;font-weight:600;width:40px;text-align:right;flex-shrink:0;">${b.count}件</span>
        </div>`;
      }).join('');
      recentBox.innerHTML = `
        <h2>直近予約数分析 <span style="font-size:11px;color:#999;">（過去30日間・合算）</span></h2>
        <div style="display:flex;align-items:center;gap:16px;margin-bottom:16px;">
          <div style="font-size:36px;font-weight:700;">${total}<span style="font-size:14px;color:#666;">件</span></div>
          <span style="background:${statusColor}20;color:${statusColor};padding:4px 12px;border-radius:12px;font-size:12px;font-weight:600;">${statusText}</span>
        </div>
        ${barsHtml}
      `;
    }
  }, 100);
}

// ============================================================
// Owner -> Property drill-down
// ============================================================
let activeOwnerPropertyDrill = null;
function toggleOwnerPropertyDrill(propertyName) {
  const container = document.getElementById('owner-property-drill-container');
  if (!container) return;
  if (activeOwnerPropertyDrill === propertyName) {
    destroyDrillCharts('owp');
    container.innerHTML = '';
    activeOwnerPropertyDrill = null;
    return;
  }
  activeOwnerPropertyDrill = propertyName;
  renderPropertyDetail(container, propertyName, 'owp');
  setTimeout(initSortableHeaders, 50);
}

// ============================================================
// Property drill-down (Tab 3)
// ============================================================
let activePropertyDrill = null;
const chartInstances = {};

function togglePropertyDrill(propertyName, clickedRow) {
  // 既存のドリルダウン行を削除
  const existing = document.getElementById('property-drill-row');
  if (existing) {
    destroyDrillCharts('prop');
    existing.remove();
  }

  // 同じ物件をクリック → 閉じるだけ
  if (activePropertyDrill === propertyName) {
    activePropertyDrill = null;
    return;
  }

  activePropertyDrill = propertyName;

  // クリックした行のすぐ下にドリルダウン行を挿入
  const drillRow = document.createElement('tr');
  drillRow.id = 'property-drill-row';
  const drillCell = document.createElement('td');
  drillCell.colSpan = 12;
  drillCell.style.padding = '0';
  drillRow.appendChild(drillCell);

  if (clickedRow) {
    clickedRow.insertAdjacentElement('afterend', drillRow);
  } else {
    document.getElementById('property-table').appendChild(drillRow);
  }

  renderPropertyDetail(drillCell, propertyName, 'prop');
  setTimeout(initSortableHeaders, 50);
  drillRow.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// スコアカード表示モード切替
function setScorecardMode(el) {
  const pills = el.parentElement.querySelectorAll('.pill');
  pills.forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.scorecardMode = el.dataset.mode;
  _scorecardCurrentProp = null;
  initRevenueCharts();
}
function setScorecardDayType(el) {
  const pills = el.parentElement.querySelectorAll('.pill');
  pills.forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.scorecardDayType = el.dataset.daytype;
  _scorecardCurrentProp = null;
  initRevenueCharts();
}

// スコアカードの物件クリック → 行直下で展開/閉じる
let _scorecardCurrentProp = null;
let _scorecardCurrentIdx = null;
function toggleScorecardDetail(propertyName, idx) {
  // 同じ物件再クリック → 閉じる
  if (_scorecardCurrentProp === propertyName) {
    destroyDrillCharts('sc');
    const prevRow = document.getElementById('sc-detail-row-' + _scorecardCurrentIdx);
    const prevInner = document.getElementById('sc-detail-inner-' + _scorecardCurrentIdx);
    if (prevRow) prevRow.style.display = 'none';
    if (prevInner) prevInner.innerHTML = '';
    _scorecardCurrentProp = null;
    _scorecardCurrentIdx = null;
    return;
  }
  // 既に別物件が開いていれば閉じる
  if (_scorecardCurrentIdx !== null) {
    destroyDrillCharts('sc');
    const prevRow = document.getElementById('sc-detail-row-' + _scorecardCurrentIdx);
    const prevInner = document.getElementById('sc-detail-inner-' + _scorecardCurrentIdx);
    if (prevRow) prevRow.style.display = 'none';
    if (prevInner) prevInner.innerHTML = '';
  }
  _scorecardCurrentProp = propertyName;
  _scorecardCurrentIdx = idx;
  currentFilters.propDetailPeriod = 'thisMonth';
  const row = document.getElementById('sc-detail-row-' + idx);
  const inner = document.getElementById('sc-detail-inner-' + idx);
  if (row && inner) {
    row.style.display = '';
    renderPropertyDetail(inner, propertyName, 'sc');
    row.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function destroyDrillCharts(prefix) {
  if (chartInstances[prefix+'SalesOcc']) { chartInstances[prefix+'SalesOcc'].destroy(); delete chartInstances[prefix+'SalesOcc']; }
  if (chartInstances[prefix+'SalesAdr']) { chartInstances[prefix+'SalesAdr'].destroy(); delete chartInstances[prefix+'SalesAdr']; }
  if (chartInstances[prefix+'Channel']) { chartInstances[prefix+'Channel'].destroy(); delete chartInstances[prefix+'Channel']; }
  if (chartInstances[prefix+'Nationality']) { chartInstances[prefix+'Nationality'].destroy(); delete chartInstances[prefix+'Nationality']; }
  if (chartInstances[prefix+'DailyOcc']) { chartInstances[prefix+'DailyOcc'].destroy(); delete chartInstances[prefix+'DailyOcc']; }
}

// ============================================================
// Shared property detail renderer
// ============================================================
function renderPropertyDetail(container, propertyName, prefix) {
  const prop = findPropByName(propertyName);
  if (!prop) return;

  // Use propDetailPeriod for month selection
  const detailPeriod = currentFilters.propDetailPeriod || 'thisMonth';
  const detailMonths = getSelectedMonths_custom(detailPeriod);
  const ym = detailMonths[detailMonths.length - 1];

  // Get reservations for this property
  const propObj = prop;
  const propResvAll = reservations.filter(r => r.propCode === propertyName || r.property === propertyName);
  const propResv = propResvAll.slice(0, 10);
  let resvRows = propResv.map(r => `<tr><td>${(r.date || '').slice(0, 10)}</td><td>${r.channel}</td><td>${r.guest}</td><td>${r.checkin}</td><td>${r.checkout}</td><td>${r.nights}泊</td><td>${fmtYenFull(r.sales)}</td><td>${r.status}</td></tr>`).join('');
  if (!resvRows) resvRows = '<tr><td colspan="8" style="color:#999;text-align:center;">データなし</td></tr>';

  // KPI: aggregate across selected months
  function aggPropStats(propName, months) {
    let totalSales = 0, totalNights = 0, totalAvailable = 0;
    months.forEach(m => {
      const s = computePropertyStats(propName, m);
      const avail = getDaysInMonth(m) * (prop.rooms || 1);
      if (s) {
        totalSales += s.sales;
        totalNights += s.nights;
      }
      totalAvailable += avail;
    });
    if (!totalAvailable) return null;
    const occ = (totalNights / totalAvailable) * 100;
    const adr = totalNights > 0 ? totalSales / totalNights : 0;
    const revpar = adr * (occ / 100);
    return { sales: totalSales, occ, adr, revpar };
  }

  const curStats = aggPropStats(propertyName, detailMonths);
  // YoY: shift each month by -1 year
  const yoyMonths = detailMonths.map(m => { const [y, mo] = m.split('-'); return `${Number(y) - 1}-${mo}`; });
  // MoM: shift each month by -1 month
  const momMonths = detailMonths.map(m => {
    const [y, mo] = m.split('-').map(Number);
    const d = new Date(y, mo - 2, 1);
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  });
  const yoyStats = aggPropStats(propertyName, yoyMonths);
  const momStats = aggPropStats(propertyName, momMonths);

  const cOcc = curStats ? curStats.occ : 0;
  const cAdr = curStats ? curStats.adr : 0;
  const cRevpar = curStats ? curStats.revpar : 0;
  const cSales = curStats ? curStats.sales : 0;

  const occVs = fmtVsLinePt(cOcc, yoyStats ? yoyStats.occ : null, momStats ? momStats.occ : null);
  const adrVs = fmtVsLine(cAdr, yoyStats ? yoyStats.adr : null, momStats ? momStats.adr : null);
  const revparVs = fmtVsLine(cRevpar, yoyStats ? yoyStats.revpar : null, momStats ? momStats.revpar : null);
  const salesVs = fmtVsLine(cSales, yoyStats ? yoyStats.sales : null, momStats ? momStats.sales : null);

  // Booking window (lead time): average days between booking date and check-in date
  const detailMonthSet = new Set(detailMonths);
  const activeResv = propResvAll.filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた' || !r.date || !r.checkin) return false;
    const ciYm = r.checkin.slice(0, 7);
    return detailMonthSet.has(ciYm);
  });
  let avgLeadTime = null;
  if (activeResv.length > 0) {
    const totalLead = activeResv.reduce((sum, r) => {
      const bookDate = new Date(r.date);
      const ciDate = new Date(r.checkin);
      return sum + Math.max(0, Math.floor((ciDate - bookDate) / 86400000));
    }, 0);
    avgLeadTime = Math.round(totalLead / activeResv.length);
  }

  const kpiHtml = `<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px;">
    <div class="kpi-card"><div class="label">総販売額</div><div class="value">${fmtYen(cSales)}</div><div class="sub">${salesVs}</div></div>
    <div class="kpi-card"><div class="label">OCC</div><div class="value">${fmtPct(cOcc)}</div><div class="sub">${occVs}</div></div>
    <div class="kpi-card"><div class="label">ADR</div><div class="value">${fmtYenFull(Math.round(cAdr))}</div><div class="sub">${adrVs}</div></div>
    <div class="kpi-card"><div class="label">RevPAR</div><div class="value">${fmtYenFull(Math.round(cRevpar))}</div><div class="sub">${revparVs}</div></div>
    <div class="kpi-card"><div class="label">予約Window</div><div class="value">${avgLeadTime !== null ? avgLeadTime + '日' : '-'}</div><div class="sub">予約〜チェックイン平均</div></div>
  </div>`;

  // 平日/休日比較テーブル
  const wdhdStats = computeWeekdayHolidayStats(propertyName, detailMonths);
  let wdhdHtml = '';
  if (wdhdStats) {
    const wd = wdhdStats.weekday;
    const hd = wdhdStats.holiday;
    const occDiff = hd.occ - wd.occ;
    const adrDiffPct = wd.adr > 0 ? ((hd.adr - wd.adr) / wd.adr * 100) : 0;
    const revparDiffPct = wd.revpar > 0 ? ((hd.revpar - wd.revpar) / wd.revpar * 100) : 0;
    const diffColor = v => v > 0 ? '#34c759' : v < 0 ? '#ff3b30' : '#999';
    const diffSign = v => v > 0 ? '+' : '';
    wdhdHtml = `<div>
      <div style="font-size:13px;font-weight:600;color:#1d1d1f;margin-bottom:8px;">平日 vs 休日 <span style="font-size:11px;font-weight:400;color:#86868b;">（休日＝金土泊＋祝前日泊）</span></div>
      <table style="width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed;">
        <thead><tr style="border-bottom:2px solid #e5e5ea;">
          <th style="text-align:left;padding:5px 4px;color:#86868b;font-weight:500;width:20%;"></th>
          <th style="text-align:right;padding:5px 4px;color:#86868b;font-weight:500;width:28%;">平日</th>
          <th style="text-align:right;padding:5px 4px;color:#86868b;font-weight:500;width:28%;">休日</th>
          <th style="text-align:right;padding:5px 4px;color:#86868b;font-weight:500;width:24%;">差分</th>
        </tr></thead>
        <tbody>
          <tr style="border-bottom:1px solid #f0f0f0;">
            <td style="padding:5px 4px;font-weight:500;">OCC</td>
            <td style="text-align:right;padding:5px 4px;">${fmtPct(wd.occ)}</td>
            <td style="text-align:right;padding:5px 4px;font-weight:600;">${fmtPct(hd.occ)}</td>
            <td style="text-align:right;padding:5px 4px;color:${diffColor(occDiff)};font-weight:600;">${diffSign(occDiff)}${occDiff.toFixed(1)}pt</td>
          </tr>
          <tr style="border-bottom:1px solid #f0f0f0;">
            <td style="padding:5px 4px;font-weight:500;">ADR</td>
            <td style="text-align:right;padding:5px 4px;">${fmtYenFull(wd.adr)}</td>
            <td style="text-align:right;padding:5px 4px;font-weight:600;">${fmtYenFull(hd.adr)}</td>
            <td style="text-align:right;padding:5px 4px;color:${diffColor(adrDiffPct)};font-weight:600;">${diffSign(adrDiffPct)}${Math.round(adrDiffPct)}%</td>
          </tr>
          <tr style="border-bottom:1px solid #f0f0f0;">
            <td style="padding:5px 4px;font-weight:500;">RevPAR</td>
            <td style="text-align:right;padding:5px 4px;">${fmtYenFull(wd.revpar)}</td>
            <td style="text-align:right;padding:5px 4px;font-weight:600;">${fmtYenFull(hd.revpar)}</td>
            <td style="text-align:right;padding:5px 4px;color:${diffColor(revparDiffPct)};font-weight:600;">${diffSign(revparDiffPct)}${Math.round(revparDiffPct)}%</td>
          </tr>
          <tr>
            <td style="padding:5px 4px;font-weight:500;">泊数</td>
            <td style="text-align:right;padding:5px 4px;">${wd.nights}泊<span style="color:#86868b;font-size:10px;">/${wd.avail}日</span></td>
            <td style="text-align:right;padding:5px 4px;font-weight:600;">${hd.nights}泊<span style="color:#86868b;font-size:10px;">/${hd.avail}日</span></td>
            <td style="text-align:right;padding:5px 4px;"></td>
          </tr>
        </tbody>
      </table>
    </div>`;
  }

  // 先行予約ペースレポート
  const paceData = computePaceReport([propertyName]);
  const paceHtml = renderPaceReportHtml(paceData, '先行予約ペース');

  // 市場比較（AirDNA）
  const marketCompareHtml = buildMarketCompareHtml(prop, curStats);

  // 分析インサイト（ルールベース）
  const insightsHtml = buildInsightsHtml(prop, curStats, wdhdStats, paceData);

  // 未来予約分析（自社実データ vs AirDNA市場データ）
  const futureAnalysisHtml = `<div class="card" style="margin-bottom:20px;">
    <h2>未来予約分析 <span id="${prefix}FutureMarketBasis" style="font-size:11px;color:#86868b;font-weight:400;"></span></h2>
    <div class="chart-grid">
      <div class="card"><h2>未来OCC推移（次90日）</h2><canvas id="${prefix}ChartFutureOcc"></canvas></div>
      <div class="card"><h2>未来ADR推移（次180日）</h2><canvas id="${prefix}ChartFutureAdr"></canvas></div>
    </div>
    <div class="card"><h2>リードタイム分布</h2><canvas id="${prefix}ChartLeadTime" height="180"></canvas></div>
  </div>`;

  // Period pills for detail view
  const dp = currentFilters.propDetailPeriod || 'thisMonth';
  const periodPillsHtml = `<div class="filter-pills" style="margin-bottom:16px;" id="${prefix}DetailPeriodPills">
    <span class="pill${dp === 'thisMonth' ? ' active' : ''}" data-period="thisMonth" onclick="setPropDetailPeriod(this,'#${prefix}DetailContainer','${propertyName}','${prefix}')">今月</span>
    <span class="pill${dp === 'lastMonth' ? ' active' : ''}" data-period="lastMonth" onclick="setPropDetailPeriod(this,'#${prefix}DetailContainer','${propertyName}','${prefix}')">前月</span>
    <span class="pill${dp === 'last3Month' ? ' active' : ''}" data-period="last3Month" onclick="setPropDetailPeriod(this,'#${prefix}DetailContainer','${propertyName}','${prefix}')">3ヶ月前</span>
    <span class="pill${dp === 'lastYear' ? ' active' : ''}" data-period="lastYear" onclick="setPropDetailPeriod(this,'#${prefix}DetailContainer','${propertyName}','${prefix}')">前年</span>
  </div>`;

  destroyDrillCharts(prefix);

  container.id = prefix + 'DetailContainer';
  container.innerHTML = `<div class="drill-down show" style="margin-top:12px;">
    <h3>${prop.name} <span style="font-size:13px;color:#666;font-weight:400;">(${prop.ownerName} / ${prop.area})</span></h3>
    ${periodPillsHtml}
    ${kpiHtml}
    ${insightsHtml}
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:20px;margin-bottom:20px;">
      ${wdhdHtml}
      ${paceHtml}
      ${marketCompareHtml}
    </div>
    <div class="chart-grid">
      <div class="card"><h2>月別 販売金額/OCC推移</h2><canvas id="${prefix}ChartSalesOcc"></canvas></div>
      <div class="card"><h2>月別 販売金額/ADR推移</h2><canvas id="${prefix}ChartSalesAdr"></canvas></div>
      <div class="card"><h2>チャネル別売上構成比</h2><canvas id="${prefix}ChartChannel"></canvas></div>
      <div class="card"><h2>ゲスト国籍別</h2><canvas id="${prefix}ChartNationality"></canvas></div>
      <div class="card" id="${prefix}RecentBookings"></div>
    </div>
    <div class="card" id="${prefix}CalendarCard">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;">
        <h2 style="margin-bottom:0;">宿泊単価カレンダー（今年 vs 前年）</h2>
        <div style="display:flex;align-items:center;gap:8px;">
          <button onclick="shiftPropertyCalendar('${prefix}','${propertyName}',-1)" style="border:1px solid #ddd;background:white;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:13px;">◀</button>
          <span id="${prefix}CalMonth" style="font-size:14px;font-weight:600;min-width:80px;text-align:center;"></span>
          <button onclick="shiftPropertyCalendar('${prefix}','${propertyName}',1)" style="border:1px solid #ddd;background:white;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:13px;">▶</button>
        </div>
      </div>
      <div id="${prefix}CalGrid"></div>
      <div style="display:flex;gap:16px;margin-top:12px;font-size:11px;color:#999;">
        <span>🟦 今年</span><span>🟧 前年</span><span style="color:#34c759;">▲ 前年比UP</span><span style="color:#ff3b30;">▼ 前年比DOWN</span>
      </div>
    </div>
    ${futureAnalysisHtml}
    <div class="card"><h2>予約一覧</h2><div class="table-wrap"><table>
      <thead><tr><th>予約日</th><th>予約サイト</th><th>ゲスト名</th><th>チェックイン</th><th>チェックアウト</th><th>泊数</th><th>販売金額</th><th>状態</th></tr></thead>
      <tbody>${resvRows}</tbody>
    </table></div></div>
  </div>`;

  setTimeout(() => {
    // Compute past 6 months + future 6 months (12 months total)
    const now = new Date();
    const curMonth = now.getMonth() + 1;
    const curYear = now.getFullYear();
    const monthLabels = [];
    const occData = [];
    const adrData = [];
    const salesData = [];
    const targetData = [];
    for (let i = -5; i <= 6; i++) {
      const d = new Date(curYear, curMonth - 1 + i, 1);
      const m = d.getMonth() + 1;
      const y = d.getFullYear();
      const mym = `${y}-${String(m).padStart(2, '0')}`;
      monthLabels.push(`${m}月`);
      const stats = computePropertyStats(propertyName, mym);
      occData.push(stats ? Math.round(stats.occ * 10) / 10 : 0);
      adrData.push(stats ? Math.round(stats.adr) : 0);
      salesData.push(stats ? Math.round(stats.sales) : 0);
      targetData.push(getTargetForProperty(prop, m));
    }
    const targetLineDataset = {
      type: 'line',
      label: '目標売上',
      data: targetData,
      borderColor: '#ff3b30',
      backgroundColor: 'transparent',
      borderDash: [6, 4],
      borderWidth: 2,
      pointRadius: 0,
      pointHoverRadius: 4,
      pointBackgroundColor: '#ff3b30',
      tension: 0,
      fill: false,
      yAxisID: 'y',
    };
    const currentIdx = 5; // index of current month

    // Bar colors: past=light, current=solid, future=dashed
    const blueBarColors = salesData.map((_, i) => i < currentIdx ? 'rgba(74,144,217,0.2)' : i === currentIdx ? 'rgba(74,144,217,0.5)' : 'rgba(74,144,217,0.1)');
    const orangeBarColors = salesData.map((_, i) => i < currentIdx ? 'rgba(245,166,35,0.2)' : i === currentIdx ? 'rgba(245,166,35,0.5)' : 'rgba(245,166,35,0.1)');
    const barBorders = salesData.map((_, i) => i > currentIdx ? 'rgba(0,0,0,0.06)' : 'transparent');
    const barBorderWidths = salesData.map((_, i) => i > currentIdx ? 1 : 0);

    // Chart 1: Sales bar + OCC line
    const ctx1 = document.getElementById(prefix+'ChartSalesOcc');
    if (ctx1) {
      chartInstances[prefix+'SalesOcc'] = new Chart(ctx1, {
        type: 'bar',
        data: {
          labels: monthLabels,
          datasets: [
            { type: 'line', label: 'OCC (%)', data: occData, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.08)', fill: true, yAxisID: 'y1', tension: 0.4, pointBorderColor: CHART_COLORS.blue },
            { type: 'bar', label: '販売金額', data: salesData, backgroundColor: blueBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' },
            { ...targetLineDataset }
          ]
        },
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => { const v = ctx.parsed.y; if (ctx.dataset.yAxisID === 'y1') return ctx.dataset.label + ': ' + v.toFixed(1) + '%'; return ctx.dataset.label + ': ¥' + Math.round(v).toLocaleString(); } } } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', min: 0, max: 110, grid: { drawOnChartArea: false }, title: { display: true, text: 'OCC (%)', font: { size: 11 } } } } }
      });
    }

    // Chart 2: Sales bar + ADR line
    const ctx1b = document.getElementById(prefix+'ChartSalesAdr');
    if (ctx1b) {
      chartInstances[prefix+'SalesAdr'] = new Chart(ctx1b, {
        type: 'bar',
        data: {
          labels: monthLabels,
          datasets: [
            { type: 'line', label: 'ADR (¥)', data: adrData, borderColor: CHART_COLORS.orange, backgroundColor: 'rgba(245,166,35,0.08)', fill: true, yAxisID: 'y1', tension: 0.4, pointBorderColor: CHART_COLORS.orange },
            { type: 'bar', label: '販売金額', data: salesData, backgroundColor: orangeBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' },
            { ...targetLineDataset }
          ]
        },
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => { const v = ctx.parsed.y; if (ctx.dataset.yAxisID === 'y1') return ctx.dataset.label + ': ¥' + Math.round(v).toLocaleString(); return ctx.dataset.label + ': ¥' + Math.round(v).toLocaleString(); } } } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
      });
    }

    // Channel breakdown from current month stats (horizontal bar, %)
    const stats = computePropertyStats(propertyName, ym);
    const channelColors = PALETTE;
    if (stats && stats.channels) {
      const totalChSales = Object.values(stats.channels).reduce((s, c) => s + c.sales, 0);
      const sorted = Object.entries(stats.channels).sort((a, b) => b[1].sales - a[1].sales);
      const channelLabels = sorted.map(([k]) => k);
      const channelPct = sorted.map(([, v]) => totalChSales > 0 ? (v.sales / totalChSales) * 100 : 0);

      const ctx2 = document.getElementById(prefix+'ChartChannel');
      if (ctx2) {
        chartInstances[prefix+'Channel'] = new Chart(ctx2, {
          type: 'bar',
          data: { labels: channelLabels, datasets: [{ label: '売上構成比', data: channelPct, backgroundColor: channelColors.slice(0, channelLabels.length).map(c => c + 'CC'), hoverBackgroundColor: channelColors.slice(0, channelLabels.length) }] },
          options: { indexAxis: 'y', responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.parsed.x.toFixed(1) + '%' } } }, scales: { x: { beginAtZero: true, max: 100, grid: { display: false }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false } } } }
        });
      }
    }

    // Nationality breakdown
    const natAgg = {};
    propResvAll.forEach(r => {
      if (r.status === 'システムキャンセル') return;
      const nat = r.nationality || '不明';
      if (!natAgg[nat]) natAgg[nat] = { count: 0, sales: 0 };
      natAgg[nat].count++;
      natAgg[nat].sales += r.sales || 0;
    });
    const natTotal = Object.values(natAgg).reduce((s, v) => s + v.count, 0);
    const natSorted = Object.entries(natAgg).sort((a, b) => b[1].count - a[1].count);
    // 上位5件 + その他
    const natTop = natSorted.slice(0, 5);
    const natOthers = natSorted.slice(5);
    const natOtherCount = natOthers.reduce((s, [, v]) => s + v.count, 0);
    if (natOtherCount > 0) natTop.push(['その他', { count: natOtherCount, sales: 0 }]);
    const natLabels = natTop.map(([k]) => k);
    const natPct = natTop.map(([, v]) => natTotal > 0 ? (v.count / natTotal) * 100 : 0);

    const ctxNat = document.getElementById(prefix+'ChartNationality');
    if (ctxNat && natLabels.length > 0) {
      chartInstances[prefix+'Nationality'] = new Chart(ctxNat, {
        type: 'bar',
        data: { labels: natLabels, datasets: [{ label: '予約数構成比', data: natPct, backgroundColor: channelColors.slice(0, natLabels.length).map(c => c + 'CC'), hoverBackgroundColor: channelColors.slice(0, natLabels.length) }] },
        options: { indexAxis: 'y', responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: false }, tooltip: { callbacks: {
          title: ctx => ctx[0].label,
          afterBody: ctx => {
            if (ctx[0].label === 'その他') {
              return natOthers.map(([k, v]) => `  ${k}: ${(natTotal > 0 ? (v.count / natTotal) * 100 : 0).toFixed(1)}% (${v.count}件)`);
            }
            return [];
          },
          label: ctx => ctx.parsed.x.toFixed(1) + '% (' + natTop[ctx.dataIndex][1].count + '件)'
        } } }, scales: { x: { beginAtZero: true, max: 100, grid: { display: false }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false } } } }
      });
    }

    // Recent bookings analysis (last 30 days)
    const recentBox = document.getElementById(prefix+'RecentBookings');
    if (recentBox) {
      const now = new Date();
      const propResv30 = reservations.filter(r => {
        if (r.status === 'システムキャンセル') return false;
        if (r.propCode !== propertyName && r.property !== propertyName) {
          if (!propObj || r.property !== propObj.propName) return false;
        }
        if (!r.date) return false;
        const diffMs = now - new Date(r.date);
        return diffMs >= 0 && diffMs < 30 * 86400000;
      });

      // Period buckets
      const buckets = [
        { label: '直近3日', min: 0, max: 3, count: 0 },
        { label: '4〜7日前', min: 3, max: 7, count: 0 },
        { label: '8〜14日前', min: 7, max: 14, count: 0 },
        { label: '15〜30日前', min: 14, max: 30, count: 0 },
      ];
      propResv30.forEach(r => {
        const days = Math.floor((now - new Date(r.date)) / 86400000);
        buckets.forEach(b => { if (days >= b.min && days < b.max) b.count++; });
      });

      const total = propResv30.length;
      const statusColor = total >= 5 ? '#34c759' : total >= 2 ? '#ff9500' : '#ff3b30';
      const statusText = total >= 5 ? '好調' : total >= 2 ? '注意' : '要確認';

      let barsHtml = buckets.map(b => {
        const pct = total > 0 ? (b.count / total) * 100 : 0;
        const barColor = b.count > 0 ? CHART_COLORS.blue : '#e5e5e5';
        return `<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
          <span style="font-size:12px;color:#666;width:80px;text-align:right;flex-shrink:0;">${b.label}</span>
          <div style="flex:1;background:#f0f0f0;border-radius:4px;height:24px;position:relative;overflow:hidden;">
            <div style="width:${pct}%;background:${barColor};height:100%;border-radius:4px;transition:width 0.3s;min-width:${b.count > 0 ? '2px' : '0'}"></div>
          </div>
          <span style="font-size:13px;font-weight:600;width:40px;text-align:right;flex-shrink:0;">${b.count}件</span>
        </div>`;
      }).join('');

      recentBox.innerHTML = `
        <h2>直近予約数分析 <span style="font-size:11px;color:#999;">（過去30日間）</span></h2>
        <div style="display:flex;align-items:center;gap:16px;margin-bottom:16px;">
          <div style="font-size:36px;font-weight:700;">${total}<span style="font-size:14px;color:#666;">件</span></div>
          <span style="background:${statusColor}20;color:${statusColor};padding:4px 12px;border-radius:12px;font-size:12px;font-weight:600;">${statusText}</span>
          <span style="position:relative;cursor:help;font-size:13px;color:#999;" onmouseenter="this.querySelector('.tip').style.display='block'" onmouseleave="this.querySelector('.tip').style.display='none'">&#63;
            <span class="tip" style="display:none;position:absolute;left:50%;transform:translateX(-50%);bottom:24px;background:#1d1d1f;color:white;font-size:11px;padding:8px 12px;border-radius:8px;white-space:nowrap;z-index:100;font-weight:400;">好調: 5件以上 / 注意: 2〜4件 / 要確認: 0〜1件<br>過去30日間の新規予約獲得数で判定</span>
          </span>
        </div>
        ${barsHtml}
      `;
    }

    renderPropertyCalendar(prefix, propertyName, 0);
    renderPropertyFutureAnalysis(prefix, propertyName);
  }, 100);
}

// 物件詳細の未来予約分析（自社実データ vs AirDNA市場データ）
function renderPropertyFutureAnalysis(prefix, propertyName) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // このプロパティの予約（未来分）
  const propResv = reservations.filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
    return r.propCode === propertyName || r.property === propertyName;
  });

  // 市場データ参照（AirDNA）— 物件マスタから該当物件を特定
  const propObj = (propertyMaster || []).find(p => p.propCode === propertyName || p.propName === propertyName) || null;
  const marketLookup = propObj ? resolveMarketLookup(propObj) : { hasData: false, matched: '' };
  const basisEl = document.getElementById(prefix + 'FutureMarketBasis');
  if (basisEl) {
    basisEl.textContent = marketLookup.hasData
      ? `（市場データ基準: ${marketLookup.matched} / AirDNA月次・未来月は前年同月）`
      : '（市場データ未取込）';
  }

  // 1. 未来OCC（次90日）
  destroyChart(prefix + 'FutureOcc');
  const occLabels = [];
  const myOccFuture = [];
  const mktOccFuture = [];
  for (let i = 0; i < 90; i++) {
    const d = new Date(today); d.setDate(d.getDate() + i);
    occLabels.push(`${d.getMonth() + 1}/${d.getDate()}`);
    const ds = d.toISOString().split('T')[0];
    // 自社: この日が予約されているか（0 or 100）
    const isBooked = propResv.some(r => r.checkin <= ds && ds < r.checkout) ? 100 : 0;
    myOccFuture.push(isBooked);
    // 市場: 該当月の月次OCC（%）
    const ym = ds.slice(0, 7);
    const mv = marketLookup.hasData ? marketLookup.occ(ym) : null;
    mktOccFuture.push(mv !== null ? Math.round(mv) : null);
  }
  // 自社のOCCは1日単位で0/100になるため、7日移動平均でならす
  const myOccSmoothed = myOccFuture.map((_, i) => {
    const start = Math.max(0, i - 3);
    const end = Math.min(myOccFuture.length, i + 4);
    const slice = myOccFuture.slice(start, end);
    return Math.round(slice.reduce((s, v) => s + v, 0) / slice.length);
  });

  const ctxFO = document.getElementById(prefix + 'ChartFutureOcc');
  if (ctxFO) {
    allCharts[prefix + 'FutureOcc'] = new Chart(ctxFO, {
      type: 'line',
      data: { labels: occLabels, datasets: [
        { label: '自社OCC (7日移動平均)', data: myOccSmoothed, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.1)', fill: true, tension: 0.3, pointRadius: 0 },
        { label: '市場平均OCC', data: mktOccFuture, borderColor: CHART_COLORS.orange, borderDash: [6, 3], backgroundColor: 'transparent', tension: 0.3, pointRadius: 0 },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y + '%' } } },
        scales: { x: { ticks: { maxTicksLimit: 12 } }, y: { beginAtZero: true, max: 100, ticks: { callback: v => v + '%' } } }
      }
    });
  }

  // 2. 未来ADR（次180日）
  destroyChart(prefix + 'FutureAdr');
  const adrLabels = [];
  const myAdrFuture = [];
  const mktAdrFuture = [];
  for (let i = 0; i < 180; i++) {
    const d = new Date(today); d.setDate(d.getDate() + i);
    adrLabels.push(`${d.getMonth() + 1}/${d.getDate()}`);
    const ds = d.toISOString().split('T')[0];
    const dow = d.getDay();
    // 自社: この日が予約されていれば、そのADR、なければnull
    const resv = propResv.find(r => r.checkin <= ds && ds < r.checkout);
    if (resv && resv.nights > 0) {
      const netSales = (resv.sales || 0) - (resv.cleaningFee || 0);
      myAdrFuture.push(Math.round(netSales / resv.nights));
    } else {
      myAdrFuture.push(null);
    }
    // 市場: 月次ADRに曜日バンプを適用（AirDNAは月次のみ）
    const ym = ds.slice(0, 7);
    const baseAdr = marketLookup.hasData ? marketLookup.adr(ym) : null;
    if (baseAdr !== null) {
      const weekendBump = (dow === 5 || dow === 6) ? 1.25 : 0.95;
      mktAdrFuture.push(Math.round(baseAdr * weekendBump));
    } else {
      mktAdrFuture.push(null);
    }
  }
  const ctxFA = document.getElementById(prefix + 'ChartFutureAdr');
  if (ctxFA) {
    allCharts[prefix + 'FutureAdr'] = new Chart(ctxFA, {
      type: 'line',
      data: { labels: adrLabels, datasets: [
        { label: '自社ADR', data: myAdrFuture, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.2)', tension: 0.3, pointRadius: 2, spanGaps: false },
        { label: '市場平均ADR', data: mktAdrFuture, borderColor: CHART_COLORS.orange, borderDash: [6, 3], backgroundColor: 'transparent', tension: 0.3, pointRadius: 0 },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + (ctx.parsed.y !== null ? '¥' + ctx.parsed.y.toLocaleString() : '未予約') } } },
        scales: { x: { ticks: { maxTicksLimit: 15 } }, y: { beginAtZero: true, ticks: { callback: v => '¥' + (v / 1000).toFixed(0) + 'k' } } }
      }
    });
  }

  // 3. リードタイム分布（過去1年の確定予約から実データ）
  destroyChart(prefix + 'LeadTime');
  const oneYearAgo = new Date(today); oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
  const pastResv = propResv.filter(r => r.date && r.checkin && new Date(r.checkin) >= oneYearAgo && new Date(r.checkin) <= today);
  const buckets = [
    { label: '0-7日', min: 0, max: 7, count: 0 },
    { label: '8-14日', min: 8, max: 14, count: 0 },
    { label: '15-30日', min: 15, max: 30, count: 0 },
    { label: '31-60日', min: 31, max: 60, count: 0 },
    { label: '61-90日', min: 61, max: 90, count: 0 },
    { label: '91日〜', min: 91, max: Infinity, count: 0 },
  ];
  pastResv.forEach(r => {
    const lead = Math.floor((new Date(r.checkin) - new Date(r.date)) / 86400000);
    const b = buckets.find(b => lead >= b.min && lead <= b.max);
    if (b) b.count++;
  });
  const totalLead = pastResv.length || 1;
  const myLtPct = buckets.map(b => Math.round((b.count / totalLead) * 100));

  const ctxLT = document.getElementById(prefix + 'ChartLeadTime');
  if (ctxLT) {
    allCharts[prefix + 'LeadTime'] = new Chart(ctxLT, {
      type: 'bar',
      data: { labels: buckets.map(b => b.label), datasets: [
        { label: '自社（過去1年）', data: myLtPct, backgroundColor: CHART_COLORS.blue + 'CC' },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y + '%' } } },
        scales: { x: { grid: { display: false } }, y: { beginAtZero: true, ticks: { callback: v => v + '%' } } }
      }
    });
  }
}

// 物件カレンダー: 日別宿泊単価（今年 vs 前年）
const _calendarState = {};

function renderPropertyCalendar(prefix, propertyName, offsetMonth) {
  if (offsetMonth === undefined) offsetMonth = 0;
  _calendarState[prefix] = { propertyName, offsetMonth };

  const now = new Date();
  const target = new Date(now.getFullYear(), now.getMonth() + offsetMonth, 1);
  const year = target.getFullYear();
  const month = target.getMonth() + 1;
  const ym = `${year}-${String(month).padStart(2, '0')}`;
  const prevYm = `${year - 1}-${String(month).padStart(2, '0')}`;

  // ヘッダー更新
  const labelEl = document.getElementById(prefix + 'CalMonth');
  if (labelEl) labelEl.textContent = `${year}年${month}月`;

  // 日次データ + 予約データから日別単価を集計
  function getDailyAdr(propName, targetYm) {
    const prop = findPropByName(propName);
    const daysInM = getDaysInMonth(targetYm);
    const monthStart = targetYm + '-01';
    const monthEnd = targetYm + '-' + String(daysInM).padStart(2, '0');
    const today = new Date().toISOString().split('T')[0];

    const dailyMap = {};
    const coveredDates = new Set();

    // 1) 日次データ（過去〜今日の実績のみ。未来分は予約データで補完）
    rawDailyData.forEach(d => {
      const date = normalizeDate(d['日付']);
      if (date > today) return; // 未来の日次データは信頼しない
      const code = generatePropCode(d['物件名'] || '', d['ルーム番号'] || '');
      const status = d['状態'] || '';
      if (code !== propName || getYearMonth(date) !== targetYm || status === 'システムキャンセル' || status === 'ブロックされた') return;
      // クリーニング代のみの行を除外（チェックアウト日に清掃料だけ計上される）
      const cleaningFee = parseNum(d['清掃料']);
      const sales = parseNum(d['売上合計']);
      if (cleaningFee > 0 && Math.abs(sales - cleaningFee) < 1) return;
      const day = parseInt(date.split('-')[2], 10);
      if (!dailyMap[day]) dailyMap[day] = { sales: 0, count: 0 };
      dailyMap[day].sales += sales;
      dailyMap[day].count += 1;
      coveredDates.add(date);
    });

    // 2) 予約データから未来分を補完（日次データにない日のみ）
    const propResv = reservations.filter(r => {
      if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
      return r.propCode === propName || r.property === propName;
    });
    propResv.forEach(r => {
      if (!r.checkin || !r.checkout || !r.nights || r.nights <= 0) return;
      const netSales = (r.sales || 0) - (r.cleaningFee || 0);
      const dailyRate = Math.round(netSales / r.nights);
      const ci = new Date(r.checkin);
      const co = new Date(r.checkout);
      for (let d = new Date(ci); d < co; d.setDate(d.getDate() + 1)) {
        const ds = d.toISOString().split('T')[0];
        if (ds < monthStart || ds > monthEnd || coveredDates.has(ds)) continue;
        const day = d.getDate();
        if (!dailyMap[day]) dailyMap[day] = { sales: 0, count: 0 };
        dailyMap[day].sales += dailyRate;
        dailyMap[day].count += 1;
        coveredDates.add(ds);
      }
    });

    const result = {};
    Object.keys(dailyMap).forEach(day => {
      const e = dailyMap[day];
      result[day] = e.count > 0 ? Math.round(e.sales / e.count) : 0;
    });
    return result;
  }

  const thisYearAdr = getDailyAdr(propertyName, ym);
  const lastYearAdr = getDailyAdr(propertyName, prevYm);

  const daysInMonth = new Date(year, month, 0).getDate();
  const firstDow = new Date(year, month - 1, 1).getDay(); // 0=Sun

  const holidays = getJapaneseHolidays(year);
  const dowLabels = ['日', '月', '火', '水', '木', '金', '土'];
  let html = '<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:4px;font-size:12px;">';
  // 曜日ヘッダー
  dowLabels.forEach((d, i) => {
    const color = i === 0 ? '#ff3b30' : i === 6 ? '#007aff' : '#666';
    html += `<div style="text-align:center;font-weight:600;color:${color};padding:4px 0;">${d}</div>`;
  });
  // 空セル
  for (let i = 0; i < firstDow; i++) html += '<div></div>';
  // 日セル
  for (let day = 1; day <= daysInMonth; day++) {
    const thisAdr = thisYearAdr[day] || 0;
    const prevAdr = lastYearAdr[day] || 0;
    const dow = (firstDow + day - 1) % 7;
    const holidayName = holidays[`${month}-${day}`] || null;
    const dayColor = (dow === 0 || holidayName) ? '#ff3b30' : dow === 6 ? '#007aff' : '#333';

    let diffHtml = '';
    if (thisAdr > 0 && prevAdr > 0) {
      const diff = thisAdr - prevAdr;
      const pct = ((diff / prevAdr) * 100).toFixed(0);
      const color = diff > 0 ? '#34c759' : diff < 0 ? '#ff3b30' : '#999';
      const arrow = diff > 0 ? '▲' : diff < 0 ? '▼' : '→';
      diffHtml = `<div style="font-size:9px;color:${color};font-weight:600;">${arrow}${Math.abs(pct)}%</div>`;
    } else if (thisAdr > 0 && prevAdr === 0) {
      diffHtml = `<div style="font-size:9px;color:#999;">前年なし</div>`;
    }

    const bgColor = thisAdr > 0 ? 'rgba(74,144,217,0.06)' : '#fafafa';
    html += `<div style="background:${bgColor};border-radius:6px;padding:4px 2px;text-align:center;min-height:60px;">
      <div style="font-weight:600;color:${dayColor};margin-bottom:2px;">${day}${holidayName ? `<span style="font-size:7px;font-weight:400;display:block;line-height:1;">${holidayName}</span>` : ''}</div>
      ${thisAdr > 0 ? `<div style="font-size:10px;color:#007aff;font-weight:600;">¥${thisAdr.toLocaleString()}</div>` : `<div style="font-size:10px;color:#ccc;">-</div>`}
      ${prevAdr > 0 ? `<div style="font-size:9px;color:#ff9500;">¥${prevAdr.toLocaleString()}</div>` : ''}
      ${diffHtml}
    </div>`;
  }
  html += '</div>';

  const gridEl = document.getElementById(prefix + 'CalGrid');
  if (gridEl) gridEl.innerHTML = html;
}

function shiftPropertyCalendar(prefix, propertyName, delta) {
  const state = _calendarState[prefix] || { propertyName, offsetMonth: 0 };
  renderPropertyCalendar(prefix, propertyName, state.offsetMonth + delta);
}

// ============================================================
// Tab 3: 物件別分析
// ============================================================
function renderPropertyTab() {
  renderOrphanAlert('property-orphan-alert');
  buildLayoutPills('property');
  const months = getSelectedMonths('property');
  const area = currentFilters.propertyArea;
  const excludeKpi = document.getElementById('excludeKpiToggle') && document.getElementById('excludeKpiToggle').checked;
  const extra = getExtraFilters('property');

  const overall = computeOverallStatsMulti(months, area, excludeKpi, extra);

  // KPIs
  document.getElementById('kpi-prop-count').textContent = overall.propertyCount + '件';
  document.getElementById('kpi-prop-occ').textContent = fmtPct(overall.occ);
  document.getElementById('kpi-prop-adr').textContent = fmtYenFull(Math.round(overall.adr));
  document.getElementById('kpi-prop-revpar').textContent = fmtYenFull(Math.round(overall.revpar));
  document.getElementById('kpi-prop-sales').textContent = fmtYen(overall.totalSales);

  // Table - use merged stats from overall
  const tbody = document.getElementById('property-table');
  let filteredProps = filterPropertiesByArea(area, extra);
  if (excludeKpi) filteredProps = filteredProps.filter(p => !p.excludeKpi);

  // Build latest reservation lookup per property
  const todayMs = new Date().setHours(0,0,0,0);
  const latestResvMap = {};
  const todayStr = new Date().toISOString().split('T')[0];
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル') return;
    if (!r.date || r.date > todayStr) return; // 未来の予約日は除外
    const key = r.propCode || r.property;
    if (!key) return;
    if (!latestResvMap[key] || r.date > latestResvMap[key]) {
      latestResvMap[key] = r.date;
    }
  });

  tbody.innerHTML = filteredProps.map(p => {
    const stats = overall.stats.find(s => s.name === p.name);
    if (!stats) return '';
    const rowClass = p.excludeKpi ? 'clickable row-excluded' : 'clickable';
    const lastDate = latestResvMap[p.name];
    let recentLabel = '-';
    if (lastDate) {
      const diffDays = Math.floor((todayMs - new Date(lastDate).getTime()) / 86400000);
      if (diffDays === 0) recentLabel = '<span class="badge-green">今日</span>';
      else if (diffDays > 0) recentLabel = `<span class="${diffDays <= 3 ? 'badge-green' : diffDays <= 7 ? 'badge-blue' : diffDays <= 14 ? 'badge-orange' : 'badge-red'}">${diffDays}日前</span>`;
      else recentLabel = '<span class="badge-green">今日</span>';
    }
    return `<tr class="${rowClass}" data-exclude="${p.excludeKpi}" onclick="togglePropertyDrill('${p.name}', this)">
      <td>${p.name}</td><td>${p.ownerName}</td><td>${p.area}</td><td>${recentLabel}</td>
      <td>${fmtPct(stats.occ)}</td><td>${fmtYenFull(Math.round(stats.adr))}</td><td>${fmtYenFull(Math.round(stats.revpar))}</td>
      <td>${stats.nights}泊</td><td>${fmtYenFull(stats.sales)}</td><td>${fmtYenFull(stats.received)}</td>
      <td>${p.excludeKpi ? '<span class="badge-gray">除外</span>' : '-'}</td>
      <td><span class="badge-${p.status === '稼働中' ? 'green' : 'orange'}">${p.status}</span></td>
    </tr>`;
  }).join('');

  // Grouped view
  renderGroupedPropertyView(months, area, excludeKpi, overall);
}

// ============================================================
// 複数物件まとめ表示
// ============================================================
function getSeriesBase(propCode) {
  // Remove trailing digits to get series base (e.g. TYB203 -> TYB, ENK101 -> ENK, FFO1 -> FFO)
  // But keep codes like H2H, FC7 intact if they don't have room suffixes
  const match = propCode.match(/^([A-Za-z]+)/);
  return match ? match[1] : propCode;
}

let activeGroupedDrill = null;

function renderGroupedPropertyView(months, area, excludeKpi, overall) {
  const tbody = document.getElementById('grouped-property-table');
  if (!tbody) return;

  let filteredProps = filterPropertiesByArea(area);
  if (excludeKpi) filteredProps = filteredProps.filter(p => !p.excludeKpi);

  // Group by series base
  const groups = {};
  filteredProps.forEach(p => {
    const base = getSeriesBase(p.name);
    if (!groups[base]) groups[base] = [];
    groups[base].push(p);
  });

  // Only show groups with 2+ properties (複数物件)
  const multiGroups = Object.entries(groups).filter(([_, props]) => props.length >= 2);

  tbody.innerHTML = multiGroups.map(([base, props]) => {
    // Aggregate stats
    let totalNights = 0, totalSales = 0, totalReceived = 0;
    const totalDays = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);
    let totalAvailable = 0;

    props.forEach(p => {
      const stats = overall.stats.find(s => s.name === p.name);
      if (stats) {
        totalNights += stats.nights;
        totalSales += stats.sales;
        totalReceived += stats.received;
      }
      totalAvailable += totalDays * (p.rooms || 1);
    });

    const occ = totalAvailable > 0 ? (totalNights / totalAvailable) * 100 : 0;
    const adr = totalNights > 0 ? totalSales / totalNights : 0;
    const revpar = adr * (occ / 100);

    const ownerName = props[0].ownerName || '';
    const areaName = props[0].area || '';

    return `<tr class="clickable" onclick="toggleGroupedDrill('${base}', this)">
      <td>${base}（${props.length}室）</td><td>${ownerName}</td><td>${areaName}</td><td>${props.length}</td>
      <td>${fmtPct(occ)}</td><td>${fmtYenFull(Math.round(adr))}</td><td>${fmtYenFull(Math.round(revpar))}</td>
      <td>${totalNights}泊</td><td>${fmtYenFull(totalSales)}</td><td>${fmtYenFull(totalReceived)}</td>
    </tr>`;
  }).join('');
}

let _lastGroupedDrillClickedRow = null;

function setGrpDrillPeriod(el) {
  currentFilters.grpDrillPeriod = el.dataset.period;
  if (activeGroupedDrill) {
    toggleGroupedDrill(activeGroupedDrill, _lastGroupedDrillClickedRow, true);
  }
}

function toggleGroupedDrill(seriesBase, clickedRow, isRefresh) {
  if (clickedRow) _lastGroupedDrillClickedRow = clickedRow;
  // 既存のドリルダウン行を削除
  const existing = document.getElementById('grouped-drill-row');
  if (existing) {
    destroyDrillCharts('grp');
    destroyDrillCharts('grpProp');
    _activeGroupedPropDetail = null;
    existing.remove();
  }

  if (activeGroupedDrill === seriesBase && !isRefresh) {
    activeGroupedDrill = null;
    return;
  }
  activeGroupedDrill = seriesBase;

  const grpPeriod = currentFilters.grpDrillPeriod || 'thisMonth';
  const months = getSelectedMonths_custom(grpPeriod);
  const area = currentFilters.propertyArea;
  const excludeKpi = document.getElementById('excludeKpiToggle') && document.getElementById('excludeKpiToggle').checked;
  const overall = computeOverallStatsMulti(months, area, excludeKpi);

  let filteredProps = filterPropertiesByArea(area);
  if (excludeKpi) filteredProps = filteredProps.filter(p => !p.excludeKpi);
  const seriesProps = filteredProps.filter(p => getSeriesBase(p.name) === seriesBase);

  const rows = seriesProps.map(p => {
    const stats = overall.stats.find(s => s.name === p.name);
    if (!stats) return '';
    return `<tr class="clickable" onclick="openGroupedPropertyDetail('${p.name}', this)">
      <td style="color:#007aff;text-decoration:underline;cursor:pointer;">${p.name}</td><td>${fmtPct(stats.occ)}</td><td>${fmtYenFull(Math.round(stats.adr))}</td>
      <td>${fmtYenFull(Math.round(stats.revpar))}</td><td>${stats.nights}泊</td>
      <td>${fmtYenFull(stats.sales)}</td><td>${fmtYenFull(stats.received)}</td>
    </tr>`;
  }).join('');

  // Aggregate totals for current / YoY / MoM periods
  function aggSeriesStats(props, targetMonths) {
    const td = targetMonths.reduce((s, ym) => s + getDaysInMonth(ym), 0);
    let nights = 0, sales = 0, received = 0, avail = 0;
    props.forEach(p => {
      targetMonths.forEach(m => {
        const s = computePropertyStats(p.name, m);
        if (s) { nights += s.nights; sales += s.sales; received += s.received || 0; }
      });
      avail += td * (p.rooms || 1);
    });
    const occ = avail > 0 ? (nights / avail) * 100 : 0;
    const adr = nights > 0 ? sales / nights : 0;
    const revpar = avail > 0 ? sales / avail : 0;
    return { nights, sales, received, occ, adr, revpar };
  }

  const curAgg = aggSeriesStats(seriesProps, months);
  const totalNights = curAgg.nights, totalSales = curAgg.sales, totalReceived = curAgg.received;
  const aggOcc = curAgg.occ;
  const aggAdr = curAgg.adr;
  const aggRevpar = curAgg.revpar;

  // YoY
  const yoyMonths = months.map(m => { const [y, mo] = m.split('-'); return `${Number(y) - 1}-${mo}`; });
  const yoyAgg = aggSeriesStats(seriesProps, yoyMonths);
  // MoM
  const momMonths = months.map(m => {
    const [y, mo] = m.split('-').map(Number);
    const d = new Date(y, mo - 2, 1);
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  });
  const momAgg = aggSeriesStats(seriesProps, momMonths);

  const salesVs = fmtVsLine(totalSales, yoyAgg.sales || null, momAgg.sales || null);
  const occVs = fmtVsLinePt(aggOcc, yoyAgg.occ || null, momAgg.occ || null);
  const adrVs = fmtVsLine(aggAdr, yoyAgg.adr || null, momAgg.adr || null);
  const revparVs = fmtVsLine(aggRevpar, yoyAgg.revpar || null, momAgg.revpar || null);

  // 予約Window (lead time) — チェックイン月が選択期間内の予約のみ
  const seriesPropNames = new Set(seriesProps.map(p => p.name));
  const grpMonthSet = new Set(months);
  const seriesResv = reservations.filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた' || !r.date || !r.checkin) return false;
    if (!seriesPropNames.has(r.propCode) && !seriesPropNames.has(r.property) && !seriesProps.some(p => r.property === p.propName)) return false;
    const ciYm = r.checkin.slice(0, 7);
    return grpMonthSet.has(ciYm);
  });
  let avgLeadTime = null;
  if (seriesResv.length > 0) {
    const totalLead = seriesResv.reduce((sum, r) => {
      const bookDate = new Date(r.date);
      const ciDate = new Date(r.checkin);
      return sum + Math.max(0, Math.floor((ciDate - bookDate) / 86400000));
    }, 0);
    avgLeadTime = Math.round(totalLead / seriesResv.length);
  }

  const ownerName = seriesProps[0] ? seriesProps[0].ownerName || '' : '';
  const areaName = seriesProps[0] ? seriesProps[0].area || '' : '';

  // クリックした行の下にドリルダウン行を挿入
  const drillRow = document.createElement('tr');
  drillRow.id = 'grouped-drill-row';
  const drillCell = document.createElement('td');
  drillCell.colSpan = 10;
  drillCell.style.padding = '0';
  drillRow.appendChild(drillCell);

  if (clickedRow) {
    clickedRow.insertAdjacentElement('afterend', drillRow);
  } else {
    document.getElementById('grouped-property-table').appendChild(drillRow);
  }

  drillCell.innerHTML = `<div class="drill-down show">
    <h3>${seriesBase} シリーズ <span style="font-size:13px;color:#666;font-weight:400;">(${ownerName} / ${areaName} / ${seriesProps.length}室)</span></h3>
    <div class="filter-pills" style="margin-bottom:16px;">
      <span class="pill${grpPeriod === 'thisMonth' ? ' active' : ''}" data-period="thisMonth" onclick="setGrpDrillPeriod(this)">今月</span>
      <span class="pill${grpPeriod === 'lastMonth' ? ' active' : ''}" data-period="lastMonth" onclick="setGrpDrillPeriod(this)">前月</span>
      <span class="pill${grpPeriod === 'last3Month' ? ' active' : ''}" data-period="last3Month" onclick="setGrpDrillPeriod(this)">3ヶ月前</span>
      <span class="pill${grpPeriod === 'lastYear' ? ' active' : ''}" data-period="lastYear" onclick="setGrpDrillPeriod(this)">前年</span>
    </div>
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:16px;">
      <div class="kpi-card"><div class="label">総販売額</div><div class="value">${fmtYen(totalSales)}</div><div class="sub">${salesVs}</div></div>
      <div class="kpi-card"><div class="label">OCC</div><div class="value">${fmtPct(aggOcc)}</div><div class="sub">${occVs}</div></div>
      <div class="kpi-card"><div class="label">ADR</div><div class="value">${fmtYenFull(Math.round(aggAdr))}</div><div class="sub">${adrVs}</div></div>
      <div class="kpi-card"><div class="label">RevPAR</div><div class="value">${fmtYenFull(Math.round(aggRevpar))}</div><div class="sub">${revparVs}</div></div>
      <div class="kpi-card"><div class="label">予約Window</div><div class="value">${avgLeadTime !== null ? avgLeadTime + '日' : '-'}</div><div class="sub">予約〜チェックイン平均</div></div>
    </div>
    ${buildGroupInsightsHtml(seriesBase, seriesProps, curAgg, months)}
    <div class="chart-grid">
      <div class="card"><h2>月別 販売金額/OCC推移</h2><canvas id="grpChartSalesOcc"></canvas></div>
      <div class="card"><h2>月別 販売金額/ADR推移</h2><canvas id="grpChartSalesAdr"></canvas></div>
      <div class="card"><h2>チャネル別売上構成比</h2><canvas id="grpChartChannel"></canvas></div>
      <div class="card"><h2>ゲスト国籍別</h2><canvas id="grpChartNationality"></canvas></div>
    </div>
    <div class="card">
      <h2>日別稼働率（過去30日〜次90日） <span style="font-size:11px;color:#86868b;font-weight:400;">シリーズ全室合計</span></h2>
      <canvas id="grpChartDailyOcc" height="140"></canvas>
    </div>
    <div class="card"><h2>部屋別内訳</h2><div class="table-wrap"><table>
      <thead><tr><th>部屋</th><th>OCC</th><th>ADR</th><th>RevPAR</th><th>販売泊数</th><th>販売金額</th><th>受取金</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div></div>
    <div id="grouped-property-detail-container"></div>
  </div>`;

  setTimeout(() => {
    const now = new Date();
    const curMonth = now.getMonth() + 1;
    const curYear = now.getFullYear();
    const monthLabels = [];
    const occData = [];
    const adrData = [];
    const salesData = [];

    for (let i = -5; i <= 6; i++) {
      const d = new Date(curYear, curMonth - 1 + i, 1);
      const m = d.getMonth() + 1;
      const y = d.getFullYear();
      const mym = `${y}-${String(m).padStart(2, '0')}`;
      monthLabels.push(`${m}月`);
      // Aggregate stats across all series properties
      let mNights = 0, mSales = 0, mAvail = 0;
      seriesProps.forEach(p => {
        const s = computePropertyStats(p.name, mym);
        if (s) { mNights += s.nights; mSales += s.sales; }
        mAvail += getDaysInMonth(mym) * (p.rooms || 1);
      });
      const mOcc = mAvail > 0 ? (mNights / mAvail) * 100 : 0;
      const mAdr = mNights > 0 ? mSales / mNights : 0;
      occData.push(Math.round(mOcc * 10) / 10);
      adrData.push(Math.round(mAdr));
      salesData.push(Math.round(mSales));
    }
    const currentIdx = 5;

    const blueBarColors = salesData.map((_, i) => i < currentIdx ? 'rgba(74,144,217,0.2)' : i === currentIdx ? 'rgba(74,144,217,0.5)' : 'rgba(74,144,217,0.1)');
    const orangeBarColors = salesData.map((_, i) => i < currentIdx ? 'rgba(245,166,35,0.2)' : i === currentIdx ? 'rgba(245,166,35,0.5)' : 'rgba(245,166,35,0.1)');
    const barBorders = salesData.map((_, i) => i > currentIdx ? 'rgba(0,0,0,0.06)' : 'transparent');
    const barBorderWidths = salesData.map((_, i) => i > currentIdx ? 1 : 0);

    // Chart 1: Sales + OCC
    const ctx1 = document.getElementById('grpChartSalesOcc');
    if (ctx1) {
      chartInstances['grpSalesOcc'] = new Chart(ctx1, {
        type: 'bar',
        data: { labels: monthLabels, datasets: [
          { type: 'line', label: 'OCC (%)', data: occData, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.08)', fill: true, yAxisID: 'y1', tension: 0.4, pointBorderColor: CHART_COLORS.blue },
          { type: 'bar', label: '販売金額', data: salesData, backgroundColor: blueBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' }
        ]},
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true }, tooltip: salesChartTooltip }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', min: 0, max: 110, grid: { drawOnChartArea: false }, title: { display: true, text: 'OCC (%)', font: { size: 11 } } } } }
      });
    }

    // Chart 2: Sales + ADR
    const ctx2 = document.getElementById('grpChartSalesAdr');
    if (ctx2) {
      chartInstances['grpSalesAdr'] = new Chart(ctx2, {
        type: 'bar',
        data: { labels: monthLabels, datasets: [
          { type: 'line', label: 'ADR (¥)', data: adrData, borderColor: CHART_COLORS.orange, backgroundColor: 'rgba(245,166,35,0.08)', fill: true, yAxisID: 'y1', tension: 0.4, pointBorderColor: CHART_COLORS.orange },
          { type: 'bar', label: '販売金額', data: salesData, backgroundColor: orangeBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' }
        ]},
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true }, tooltip: salesChartTooltip }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
      });
    }

    // Chart 3: Channel breakdown (aggregate across series)
    const ym = getSelectedMonth('property');
    const channelAgg = {};
    seriesProps.forEach(p => {
      const s = computePropertyStats(p.name, ym);
      if (s && s.channels) {
        Object.entries(s.channels).forEach(([ch, v]) => {
          if (!channelAgg[ch]) channelAgg[ch] = 0;
          channelAgg[ch] += v.sales;
        });
      }
    });
    const totalChSales = Object.values(channelAgg).reduce((s, v) => s + v, 0);
    const sorted = Object.entries(channelAgg).sort((a, b) => b[1] - a[1]);
    const channelLabels = sorted.map(([k]) => k);
    const channelPct = sorted.map(([, v]) => totalChSales > 0 ? (v / totalChSales) * 100 : 0);
    const channelColors = PALETTE;

    const ctx3 = document.getElementById('grpChartChannel');
    if (ctx3 && channelLabels.length > 0) {
      chartInstances['grpChannel'] = new Chart(ctx3, {
        type: 'bar',
        data: { labels: channelLabels, datasets: [{ label: '売上構成比', data: channelPct, backgroundColor: channelColors.slice(0, channelLabels.length).map(c => c + 'CC'), hoverBackgroundColor: channelColors.slice(0, channelLabels.length) }] },
        options: { indexAxis: 'y', responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.parsed.x.toFixed(1) + '%' } } }, scales: { x: { beginAtZero: true, max: 100, grid: { display: false }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false } } } }
      });
    }

    // Nationality breakdown (aggregate across series)
    const grpNatAgg = {};
    seriesProps.forEach(p => {
      const propObj = findPropByName(p.name);
      reservations.filter(r => {
        if (r.status === 'システムキャンセル') return false;
        return r.propCode === p.name || r.property === p.name;
      }).forEach(r => {
        const nat = r.nationality || '不明';
        if (!grpNatAgg[nat]) grpNatAgg[nat] = { count: 0, sales: 0 };
        grpNatAgg[nat].count++;
        grpNatAgg[nat].sales += r.sales || 0;
      });
    });
    const grpNatTotal = Object.values(grpNatAgg).reduce((s, v) => s + v.count, 0);
    const grpNatSorted = Object.entries(grpNatAgg).sort((a, b) => b[1].count - a[1].count);
    const grpNatTop = grpNatSorted.slice(0, 5);
    const grpNatOthers = grpNatSorted.slice(5);
    const grpNatOtherCount = grpNatOthers.reduce((s, [, v]) => s + v.count, 0);
    if (grpNatOtherCount > 0) grpNatTop.push(['その他', { count: grpNatOtherCount, sales: 0 }]);
    const grpNatLabels = grpNatTop.map(([k]) => k);
    const grpNatPct = grpNatTop.map(([, v]) => grpNatTotal > 0 ? (v.count / grpNatTotal) * 100 : 0);

    const ctxNat = document.getElementById('grpChartNationality');
    if (ctxNat && grpNatLabels.length > 0) {
      chartInstances['grpNationality'] = new Chart(ctxNat, {
        type: 'bar',
        data: { labels: grpNatLabels, datasets: [{ label: '予約数構成比', data: grpNatPct, backgroundColor: channelColors.slice(0, grpNatLabels.length).map(c => c + 'CC'), hoverBackgroundColor: channelColors.slice(0, grpNatLabels.length) }] },
        options: { indexAxis: 'y', responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: false }, tooltip: { callbacks: {
          title: ctx => ctx[0].label,
          afterBody: ctx => {
            if (ctx[0].label === 'その他') {
              return grpNatOthers.map(([k, v]) => `  ${k}: ${(grpNatTotal > 0 ? (v.count / grpNatTotal) * 100 : 0).toFixed(1)}% (${v.count}件)`);
            }
            return [];
          },
          label: ctx => ctx.parsed.x.toFixed(1) + '% (' + grpNatTop[ctx.dataIndex][1].count + '件)'
        } } }, scales: { x: { beginAtZero: true, max: 100, grid: { display: false }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false } } } }
      });
    }

    // 日別稼働率（次90日）— シリーズ全室の合計ベース
    const todayD = new Date(); todayD.setHours(0, 0, 0, 0);
    const totalRooms = seriesProps.reduce((s, p) => s + (p.rooms || 1), 0) || 1;
    const dailyLabels = [];
    const dailyOcc = [];
    const dailyBooked = [];
    const dailyBlocked = [];
    const dailySales = [];
    // 各物件の未来予約（キャンセル除く）をキャッシュ
    const propResvMap = seriesProps.map(p => {
      const all = reservations.filter(r =>
        r.status !== 'キャンセル' && r.status !== 'システムキャンセル' &&
        (r.propCode === p.name || r.property === p.name)
      );
      return {
        propCode: p.name,
        rooms: p.rooms || 1,
        resv: all.filter(r => r.status !== 'ブロックされた'),
        blocked: all.filter(r => r.status === 'ブロックされた'),
      };
    });
    let todayIdx = 0;
    for (let i = -30; i < 90; i++) {
      const d = new Date(todayD); d.setDate(d.getDate() + i);
      dailyLabels.push(`${d.getMonth() + 1}/${d.getDate()}`);
      if (i === 0) todayIdx = dailyLabels.length - 1;
      const ds = d.toISOString().split('T')[0];
      let bookedRooms = 0, blockedRooms = 0, daySales = 0;
      propResvMap.forEach(pp => {
        const hits = pp.resv.filter(r => r.checkin <= ds && ds < r.checkout);
        bookedRooms += Math.min(hits.length, pp.rooms);
        hits.forEach(r => {
          if (r.nights > 0) {
            const net = (r.sales || 0) - (r.cleaningFee || 0);
            daySales += net / r.nights;
          }
        });
        const blk = pp.blocked.filter(r => r.checkin <= ds && ds < r.checkout).length;
        blockedRooms += Math.min(blk, pp.rooms);
      });
      dailyOcc.push(Math.round((bookedRooms / totalRooms) * 1000) / 10);
      dailyBooked.push(bookedRooms);
      dailyBlocked.push(blockedRooms);
      dailySales.push(Math.round(daySales));
    }
    const ctxDaily = document.getElementById('grpChartDailyOcc');
    if (ctxDaily) {
      const todayLinePlugin = {
        id: 'todayLine',
        afterDatasetsDraw(chart) {
          const { ctx, chartArea: { top, bottom }, scales: { x } } = chart;
          const xPos = x.getPixelForValue(todayIdx);
          ctx.save();
          ctx.strokeStyle = '#ff3b30';
          ctx.lineWidth = 1.5;
          ctx.setLineDash([5, 4]);
          ctx.beginPath();
          ctx.moveTo(xPos, top);
          ctx.lineTo(xPos, bottom);
          ctx.stroke();
          ctx.fillStyle = '#ff3b30';
          ctx.font = '10px sans-serif';
          ctx.fillText('今日', xPos + 4, top + 12);
          ctx.restore();
        }
      };
      chartInstances['grpDailyOcc'] = new Chart(ctxDaily, {
        type: 'line',
        plugins: [todayLinePlugin],
        data: { labels: dailyLabels, datasets: [
          { type: 'line', label: '日別OCC (%)', data: dailyOcc, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.15)', fill: true, tension: 0.25, pointRadius: 0, yAxisID: 'y',
            segment: {
              borderColor: c => c.p1DataIndex < todayIdx ? 'rgba(74,144,217,0.35)' : CHART_COLORS.blue,
              backgroundColor: c => c.p1DataIndex < todayIdx ? 'rgba(74,144,217,0.05)' : 'rgba(74,144,217,0.15)',
            }
          },
          { type: 'line', label: '残室数', data: dailyBooked.map((b, i) => totalRooms - b - dailyBlocked[i]), borderColor: CHART_COLORS.orange, borderDash: [4, 3], backgroundColor: 'transparent', tension: 0.25, pointRadius: 0, yAxisID: 'y1',
            segment: {
              borderColor: c => c.p1DataIndex < todayIdx ? 'rgba(245,166,35,0.3)' : CHART_COLORS.orange,
            }
          },
        ]},
        options: {
          responsive: true,
          animation: { duration: 600, easing: 'easeOutQuart' },
          interaction: { mode: 'index', intersect: false },
          plugins: { legend: { display: true }, tooltip: { callbacks: {
            label: ctx => {
              const i = ctx.dataIndex;
              const booked = dailyBooked[i];
              const blocked = dailyBlocked[i];
              const remaining = totalRooms - booked - blocked;
              const sales = dailySales[i];
              if (ctx.datasetIndex === 0) {
                return [
                  `OCC: ${ctx.parsed.y}%`,
                  `予約: ${booked}室 / ブロック: ${blocked}室 / 残: ${remaining}室（全${totalRooms}室）`,
                  `売上: ¥${sales.toLocaleString()}`,
                ];
              }
              return null;
            }
          } } },
          scales: {
            x: { ticks: { maxTicksLimit: 12 }, grid: { display: false } },
            y: { position: 'left', beginAtZero: true, max: 100, title: { display: true, text: 'OCC (%)', font: { size: 11 } }, ticks: { callback: v => v + '%' } },
            y1: { position: 'right', beginAtZero: true, max: totalRooms, grid: { drawOnChartArea: false }, title: { display: true, text: '残室数', font: { size: 11 } }, ticks: { stepSize: 1, precision: 0 } }
          }
        }
      });
    }
  }, 100);
  setTimeout(initSortableHeaders, 150);
  drillRow.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

let _activeGroupedPropDetail = null;

function openGroupedPropertyDetail(propertyName, clickedRow) {
  const container = document.getElementById('grouped-property-detail-container');
  if (!container) return;

  // 同じ物件をクリック → 閉じる
  if (_activeGroupedPropDetail === propertyName) {
    container.innerHTML = '';
    _activeGroupedPropDetail = null;
    destroyDrillCharts('grpProp');
    return;
  }

  _activeGroupedPropDetail = propertyName;
  destroyDrillCharts('grpProp');
  renderPropertyDetail(container, propertyName, 'grpProp');
  container.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ============================================================
// Tab 4: 予約管理
// ============================================================
function localDateStr(d) {
  // ローカルタイムゾーン基準で YYYY-MM-DD 文字列を返す（toISOStringはUTC変換でズレるため）
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function getReservationPeriodPredicate() {
  return getReservationPeriodInfo().current;
}

function getReservationPeriodInfo() {
  // 全ピル共通: 予約日 r.date ベース。current/previous は同じ長さの直前期間。
  const period = currentFilters.reservationPeriod || 'yesterday';
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // [start, end) の半開区間を返すヘルパー
  const range = (() => {
    const dayRange = (days, label) => {
      const start = new Date(today); start.setDate(start.getDate() - days);
      return { start, end: new Date(today), label };
    };
    if (period === 'yesterday') return dayRange(1, '対前日');
    if (period === 'last3Days') return dayRange(3, '対前3日');
    if (period === 'last7Days') return dayRange(7, '対前7日');
    if (period === 'thisMonth') {
      const start = new Date(today.getFullYear(), today.getMonth(), 1);
      const end = new Date(today.getFullYear(), today.getMonth() + 1, 1);
      return { start, end, label: '対前月' };
    }
    if (period === 'lastMonth') {
      const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      const end = new Date(today.getFullYear(), today.getMonth(), 1);
      return { start, end, label: '対前月' };
    }
    if (period === 'last3Months') {
      const start = new Date(today.getFullYear(), today.getMonth() - 2, 1);
      const end = new Date(today.getFullYear(), today.getMonth() + 1, 1);
      return { start, end, label: '対前3ヶ月' };
    }
    return dayRange(1, '対前日');
  })();

  // 前期間 = 同じ長さで直前にスライド
  const lenMs = range.end - range.start;
  const prevStart = new Date(range.start.getTime() - lenMs);
  const prevEnd = new Date(range.start);

  const cs = localDateStr(range.start), ce = localDateStr(range.end);
  const ps = localDateStr(prevStart), pe = localDateStr(prevEnd);

  return {
    current: r => r.date && r.date >= cs && r.date < ce,
    previous: r => r.date && r.date >= ps && r.date < pe,
    vsLabel: range.label,
    _start: range.start,
    _end: range.end,
  };
}

function fmtVsPct(cur, prev) {
  if (prev === 0) {
    if (cur === 0) return '±0%';
    return '前期0';
  }
  const pct = ((cur - prev) / prev) * 100;
  const sign = pct >= 0 ? '+' : '';
  return `${sign}${pct.toFixed(1)}%`;
}

function renderReservationTab() {
  const statusFilter = document.getElementById('resv-status-filter').value;
  const periodInfo = getReservationPeriodInfo();

  // ステータスのプレフィルタを共通関数化（current/previous両方に適用）
  const applySelectFilters = arr => {
    if (statusFilter) return arr.filter(r => r.status === statusFilter);
    return arr;
  };

  const filtered = applySelectFilters(reservations).filter(periodInfo.current);

  // YoY / MoM / DoD comparison filters
  const shiftDateFilter = (filterFn, shiftDays) => {
    return r => {
      const shifted = { ...r, date: r.date ? localDateStr(new Date(new Date(r.date).getTime() + shiftDays * 86400000)) : r.date };
      return filterFn(shifted);
    };
  };
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const buildYoyFilter = () => {
    // Same period last year
    const curStart = periodInfo._start;
    const curEnd = periodInfo._end;
    if (!curStart || !curEnd) return () => false;
    const yoyStart = new Date(curStart); yoyStart.setFullYear(yoyStart.getFullYear() - 1);
    const yoyEnd = new Date(curEnd); yoyEnd.setFullYear(yoyEnd.getFullYear() - 1);
    const ys = localDateStr(yoyStart), ye = localDateStr(yoyEnd);
    return r => r.date && r.date >= ys && r.date < ye;
  };
  const buildMomFilter = () => {
    const curStart = periodInfo._start;
    const curEnd = periodInfo._end;
    if (!curStart || !curEnd) return () => false;
    const momStart = new Date(curStart); momStart.setMonth(momStart.getMonth() - 1);
    const momEnd = new Date(curEnd); momEnd.setMonth(momEnd.getMonth() - 1);
    const ms = localDateStr(momStart), me = localDateStr(momEnd);
    return r => r.date && r.date >= ms && r.date < me;
  };
  const buildDodFilter = () => {
    const curStart = periodInfo._start;
    const curEnd = periodInfo._end;
    if (!curStart || !curEnd) return () => false;
    const lenMs = curEnd - curStart;
    const dodStart = new Date(curStart.getTime() - 86400000);
    const dodEnd = new Date(dodStart.getTime() + lenMs);
    const ds = localDateStr(dodStart), de = localDateStr(dodEnd);
    return r => r.date && r.date >= ds && r.date < de;
  };

  const base = applySelectFilters(reservations);
  const yoyFiltered = base.filter(buildYoyFilter());
  const momFiltered = base.filter(buildMomFilter());
  const dodFiltered = base.filter(buildDodFilter());

  // KPI計算ヘルパー
  const calcKpis = (arr) => {
    const count = arr.length;
    const cancel = arr.filter(r => r.status === 'システムキャンセル').length;
    const confirmed = arr.filter(r => r.status !== 'システムキャンセル');
    const nights = confirmed.reduce((s, r) => s + r.nights, 0);
    const sales = confirmed.reduce((s, r) => s + (r.sales || 0), 0);
    const cancelSales = arr.filter(r => r.status === 'システムキャンセル').reduce((s, r) => s + (r.sales || 0), 0);
    const avgNights = confirmed.length > 0 ? nights / confirmed.length : 0;
    const avgGuests = confirmed.length > 0 ? confirmed.reduce((s, r) => s + r.guestCount, 0) / confirmed.length : 0;
    const adr = nights > 0 ? sales / nights : 0;
    const validW = confirmed.filter(r => r.date && r.checkin);
    const window = validW.length > 0 ? Math.round(validW.reduce((s, r) => s + Math.max(0, Math.floor((new Date(r.checkin) - new Date(r.date)) / 86400000)), 0) / validW.length) : null;
    return { count, cancel, cancelSales, sales, adr, avgNights, avgGuests, window };
  };

  const cur = calcKpis(filtered);
  const yoy = calcKpis(yoyFiltered);
  const mom = calcKpis(momFiltered);
  const dod = calcKpis(dodFiltered);

  const fmtVs3 = (curVal, yoyVal, momVal, dodVal) => {
    const fmt1 = (label, prev) => {
      if (prev == null || prev === 0) return `${label} -`;
      const pct = ((curVal - prev) / prev) * 100;
      const sign = pct >= 0 ? '+' : '';
      const cls = pct >= 0 ? 'positive' : 'negative';
      return `<span class="${cls}">${label} ${sign}${pct.toFixed(1)}%</span>`;
    };
    return `${fmt1('YoY', yoyVal)} / ${fmt1('MoM', momVal)} / ${fmt1('DoD', dodVal)}`;
  };

  document.getElementById('kpi-resv-count').textContent = cur.count + '件';
  document.getElementById('kpi-resv-count-vs-cnt').innerHTML = fmtVs3(cur.count, yoy.count, mom.count, dod.count);
  document.getElementById('kpi-resv-sales').textContent = fmtYenFull(cur.sales);
  document.getElementById('kpi-resv-count-vs').innerHTML = fmtVs3(cur.sales, yoy.sales, mom.sales, dod.sales);
  document.getElementById('kpi-resv-adr').textContent = fmtYenFull(Math.round(cur.adr));
  document.getElementById('kpi-resv-adr-vs').innerHTML = fmtVs3(cur.adr, yoy.adr, mom.adr, dod.adr);
  document.getElementById('kpi-resv-nights').textContent = cur.avgNights.toFixed(1) + '泊';
  document.getElementById('kpi-resv-nights-vs').innerHTML = fmtVs3(cur.avgNights, yoy.avgNights, mom.avgNights, dod.avgNights);
  document.getElementById('kpi-resv-guests').textContent = cur.avgGuests.toFixed(1) + '名';
  document.getElementById('kpi-resv-guests-vs').innerHTML = fmtVs3(cur.avgGuests, yoy.avgGuests, mom.avgGuests, dod.avgGuests);
  document.getElementById('kpi-resv-window').textContent = cur.window !== null ? cur.window + '日' : '-';
  const windowVsEl = document.getElementById('kpi-resv-window-vs');
  if (windowVsEl) {
    windowVsEl.innerHTML = fmtVs3(cur.window || 0, yoy.window, mom.window, dod.window);
  }

  // Build watchlist lookup sets for badges
  const newPropNames = new Set(properties.filter(p => isNewProperty(p) && p.status === '稼働中' && !p.excludeKpi).map(p => p.name));
  const watchPropNames = new Set();
  properties.forEach(p => {
    if (p.status !== '稼働中' || p.excludeKpi || isNewProperty(p)) return;
    if (getWatchlistReasons(p).length > 0) watchPropNames.add(p.name);
  });

  // Table（キャンセルは表示から除外）
  const tbody = document.getElementById('reservation-table');
  const thead = tbody && tbody.parentElement.querySelector('thead');
  const validResv = filtered.filter(r => r.status !== 'システムキャンセル' && r.status !== 'キャンセル');
  const viewMode = currentFilters.reservationView || 'all';

  if (viewMode === 'grouped') {
    // シリーズ集計テーブル
    const agg = {};
    validResv.forEach(r => {
      const code = r.propCode || r.property || '';
      const base = code ? getSeriesBase(code) : (r.property || 'その他');
      if (!agg[base]) agg[base] = { count: 0, sales: 0, nights: 0, guests: 0, received: 0 };
      agg[base].count++;
      agg[base].sales += r.sales || 0;
      agg[base].nights += r.nights || 0;
      agg[base].guests += r.guestCount || 0;
      agg[base].received += r.received || 0;
    });
    const rows = Object.entries(agg).sort((a, b) => b[1].count - a[1].count);
    if (thead) thead.innerHTML = `<tr><th>シリーズ</th><th>予約数</th><th>GMV</th><th>ADR</th><th>平均泊数</th><th>平均ゲスト数</th><th>総受取金</th></tr>`;
    tbody.innerHTML = rows.map(([base, s]) => {
      const adr = s.nights > 0 ? Math.round(s.sales / s.nights) : 0;
      const avgNights = s.count > 0 ? (s.nights / s.count).toFixed(1) : '-';
      const avgGuests = s.count > 0 ? (s.guests / s.count).toFixed(1) : '-';
      return `<tr><td style="font-weight:600;">${base}</td><td>${s.count}件</td><td>${fmtYenFull(s.sales)}</td><td>${fmtYenFull(adr)}</td><td>${avgNights}泊</td><td>${avgGuests}名</td><td>${fmtYenFull(s.received)}</td></tr>`;
    }).join('');
    return;
  }

  // 通常モード: 個別予約リスト
  if (thead) thead.innerHTML = `<tr><th>予約日</th><th>予約サイト</th><th>物件名</th><th>販売額</th><th>チェックイン</th><th>泊数</th><th>ゲスト数</th><th>チェックアウト</th><th>ゲスト名</th><th>国籍</th><th>状態</th><th>受取金</th><th>支払済み</th><th>AirHost予約ID</th></tr>`;
  const displayResv = validResv.slice(0, 100);
  tbody.innerHTML = displayResv.map(r => {
    const statusBadge = r.status === '確認済み' ? 'badge-green' : r.status === 'システムキャンセル' ? 'badge-red' : 'badge-orange';
    const prop = findPropByReservation(r);
    const propName = prop ? prop.name : (r.propCode || r.property);
    let propMarks = '';
    if (prop && newPropNames.has(prop.name)) propMarks += ' <span class="badge-blue" style="font-size:10px;">🆕</span>';
    if (prop && watchPropNames.has(prop.name)) propMarks += ' <span class="badge-orange" style="font-size:10px;">⚠️</span>';
    return `<tr>
      <td>${(r.date || '').slice(0, 10)}</td><td>${r.channel}</td><td>${r.property}${propMarks}</td><td>${fmtYenFull(r.sales)}</td><td>${r.checkin}</td><td>${r.nights}泊</td><td>${r.guestCount}名</td><td>${r.checkout}</td><td>${r.guest}</td><td>${r.nationality}</td><td><span class="${statusBadge}">${r.status}</span></td><td>${fmtYenFull(r.received)}</td><td>${r.paid}</td><td>${r.id}</td>
    </tr>`;
  }).join('');
}

// ============================================================
// Tab 5: 売上・稼働
// ============================================================
function renderRevenueTab() {
  buildLayoutPills('revenue');
  const months = getSelectedMonths('revenue');
  const monthSet = new Set(months);
  const area = currentFilters.revenueArea;
  const excludeKpi = document.getElementById('excludeKpiToggleRev') && document.getElementById('excludeKpiToggleRev').checked;
  const extra = getExtraFilters('revenue');

  const overall = computeOverallStatsMulti(months, area, excludeKpi, extra);

  // YoY / MoM comparison
  const yoyMonths = months.map(ym => { const [y, m] = ym.split('-'); return `${Number(y) - 1}-${m}`; });
  const momMonths = months.map(ym => { const [y, m] = ym.split('-').map(Number); const d = new Date(y, m - 2, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; });
  const yoyOverall = computeOverallStatsMulti(yoyMonths, area, excludeKpi, extra);
  const momOverall = computeOverallStatsMulti(momMonths, area, excludeKpi, extra);

  document.getElementById('kpi-rev-occ').textContent = fmtPct(overall.occ);
  document.getElementById('kpi-rev-occ-vs').innerHTML = fmtVsLinePt(overall.occ, yoyOverall.occ, momOverall.occ);
  document.getElementById('kpi-rev-adr').textContent = fmtYenFull(Math.round(overall.adr));
  document.getElementById('kpi-rev-adr-vs').innerHTML = fmtVsLine(overall.adr, yoyOverall.adr, momOverall.adr);
  document.getElementById('kpi-rev-revpar').textContent = fmtYenFull(Math.round(overall.revpar));
  document.getElementById('kpi-rev-revpar-vs').innerHTML = fmtVsLine(overall.revpar, yoyOverall.revpar, momOverall.revpar);
  document.getElementById('kpi-rev-sales').textContent = fmtYen(overall.totalSales);
  document.getElementById('kpi-rev-sales-vs').innerHTML = fmtVsLine(overall.totalSales, yoyOverall.totalSales, momOverall.totalSales);
  document.getElementById('kpi-rev-received').textContent = fmtYen(overall.totalReceived);
  document.getElementById('kpi-rev-received-vs').innerHTML = fmtVsLine(overall.totalReceived, yoyOverall.totalReceived, momOverall.totalReceived);

  // 予約Window (booking lead time)
  const extraFilteredNames = new Set(filterPropertiesByArea(area, extra).map(p => p.name));
  const calcRevWindowForMonths = (ms) => {
    const mSet = new Set(ms);
    const rvs = reservations.filter(r => {
      if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
      if (!r.date || !r.checkin) return false;
      if (!mSet.has(getYearMonth(r.checkin))) return false;
      const prop = findPropByReservation(r);
      if (!prop) return false;
      if (!extraFilteredNames.has(prop.name)) return false;
      if (excludeKpi && prop.excludeKpi) return false;
      return true;
    });
    if (rvs.length === 0) return null;
    return Math.round(rvs.reduce((s, r) => s + Math.max(0, Math.floor((new Date(r.checkin) - new Date(r.date)) / 86400000)), 0) / rvs.length);
  };
  const curWindow = calcRevWindowForMonths(months);
  const yoyWindow = calcRevWindowForMonths(yoyMonths);
  const momWindow = calcRevWindowForMonths(momMonths);
  document.getElementById('kpi-rev-window').textContent = curWindow !== null ? curWindow + '日' : '-';
  document.getElementById('kpi-rev-window-vs').innerHTML = fmtVsLine(curWindow || 0, yoyWindow, momWindow);

  // Channel performance table
  const confirmedResv = reservations.filter(r => {
    if (r.status === 'システムキャンセル') return false;
    if (!monthSet.has(getYearMonth(r.checkin))) return false;
    const prop = findPropByReservation(r);
    if (!prop) return area === '全体';
    if (!extraFilteredNames.has(prop.name)) return false;
    if (excludeKpi && prop.excludeKpi) return false;
    return true;
  });

  const channelMap = {};
  confirmedResv.forEach(r => {
    const ch = r.channel || 'その他';
    if (!channelMap[ch]) channelMap[ch] = { count: 0, nights: 0, sales: 0 };
    channelMap[ch].count++;
    channelMap[ch].nights += r.nights;
    channelMap[ch].sales += r.sales;
  });

  const totalChSales = Object.values(channelMap).reduce((s, c) => s + c.sales, 0);
  const chEntries = Object.entries(channelMap).sort((a, b) => b[1].sales - a[1].sales);

  const chTbody = document.getElementById('channel-perf-table');
  let totalCount = 0, totalNights = 0;
  chTbody.innerHTML = chEntries.map(([ch, data]) => {
    totalCount += data.count;
    totalNights += data.nights;
    const adr = data.nights > 0 ? data.sales / data.nights : 0;
    const share = totalChSales > 0 ? (data.sales / totalChSales) * 100 : 0;
    return `<tr><td>${ch}</td><td>${data.count}</td><td>${data.nights}</td><td class="text-right">${fmtYenFull(data.sales)}</td><td class="text-right">${fmtYenFull(Math.round(adr))}</td><td>${fmtPct(share)}</td></tr>`;
  }).join('');

  const overallAdr = totalNights > 0 ? totalChSales / totalNights : 0;
  chTbody.innerHTML += `<tr class="totals-row"><td>合計</td><td>${totalCount}</td><td>${totalNights}</td><td class="text-right">${fmtYenFull(totalChSales)}</td><td class="text-right">${fmtYenFull(Math.round(overallAdr))}</td><td>100%</td></tr>`;
}

// ============================================================
// Charts
// ============================================================
const allCharts = {};

function destroyChart(key) {
  if (allCharts[key]) { allCharts[key].destroy(); delete allCharts[key]; }
}

function initChartsForTab(tabId) {
  if (tabId === 'pmbm') initPmbmCharts();
  if (tabId === 'daily') initDailyCharts();
  if (tabId === 'reservation') initReservationCharts();
  if (tabId === 'revenue') initRevenueCharts();
  if (tabId === 'review') initReviewCharts();
  if (tabId === 'watchlist') initWatchlistCharts();
  if (tabId === 'market') initMarketCharts();
}

function initDailyCharts() {
  const area = currentFilters.dailyArea;
  const now = new Date();

  // Past 3 months + current + future 3 months = 7 months (always, regardless of period filter)
  const chartMonths = [];
  const chartLabels = [];
  for (let i = -3; i <= 3; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const m = d.getMonth() + 1;
    const y = d.getFullYear();
    chartMonths.push(`${y}-${String(m).padStart(2, '0')}`);
    chartLabels.push(`${m}月`);
  }
  const currentIdx = 3;

  // Collect channel sales per month (checkin-based, from reservation data)
  const channelSet = new Set();
  const monthChannelSales = {};
  chartMonths.forEach(ym => { monthChannelSales[ym] = {}; });

  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    const ciYm = getYearMonth(r.checkin);
    if (!monthChannelSales[ciYm]) return;
    if (area !== '全体') {
      const prop = findPropByReservation(r);
      if (!prop || prop.area !== area) return;
    }
    const ch = r.channel || 'その他';
    channelSet.add(ch);
    monthChannelSales[ciYm][ch] = (monthChannelSales[ciYm][ch] || 0) + (r.sales || 0);
  });

  const channels = [...channelSet].sort((a, b) => {
    const totalA = chartMonths.reduce((s, ym) => s + (monthChannelSales[ym][a] || 0), 0);
    const totalB = chartMonths.reduce((s, ym) => s + (monthChannelSales[ym][b] || 0), 0);
    return totalB - totalA;
  });

  const channelColors = PALETTE;

  // Chart 1: Stacked bar - monthly sales by channel
  destroyChart('dailyMonthlySales');
  const ctx1 = document.getElementById('chartDailyMonthlySales');
  if (ctx1) {
    const datasets = channels.map((ch, idx) => ({
      label: ch,
      data: chartMonths.map(ym => monthChannelSales[ym][ch] || 0),
      backgroundColor: channelColors[idx % channelColors.length] + 'CC',
      hoverBackgroundColor: channelColors[idx % channelColors.length],
    }));
    // 総額はチャネル合計（バーの高さと同じソース）で計算
    const monthTotals = chartMonths.map(ym => {
      return channels.reduce((s, ch) => s + (monthChannelSales[ym][ch] || 0), 0);
    });
    const stackedTotalPlugin = {
      id: 'stackedTotalLabel',
      afterDatasetsDraw(chart) {
        const { ctx: c, scales: { x, y } } = chart;
        c.save();
        c.font = 'bold 11px -apple-system, sans-serif';
        c.fillStyle = '#1d1d1f';
        c.textAlign = 'center';
        c.textBaseline = 'bottom';
        monthTotals.forEach((total, i) => {
          if (total <= 0) return;
          const xPos = x.getPixelForValue(i);
          // 棒グラフの実際の高さ（チャネル合計）を使って位置決め
          const barTotal = channels.reduce((s, ch) => s + (monthChannelSales[chartMonths[i]][ch] || 0), 0);
          const yPos = y.getPixelForValue(barTotal);
          const label = total >= 100000000 ? (total / 100000000).toFixed(1) + '億' : Math.round(total / 10000).toLocaleString() + '万';
          c.fillText(label, xPos, yPos - 4);
        });
        c.restore();
      }
    };
    allCharts['dailyMonthlySales'] = new Chart(ctx1, {
      type: 'bar',
      data: { labels: chartLabels, datasets },
      plugins: [stackedTotalPlugin],
      options: {
        responsive: true,
        animation: { duration: 600, easing: 'easeOutQuart' },
        plugins: {
          legend: { position: 'top', labels: { font: { size: 11 }, padding: 20 } },
          tooltip: {
            mode: 'index',
            callbacks: { label: ctx => ctx.dataset.label + ': ¥' + Math.round(ctx.parsed.y).toLocaleString() }
          }
        },
        scales: {
          x: { stacked: true, grid: { display: false } },
          y: { stacked: true, beginAtZero: true, ticks: { callback: v => (v / 10000).toFixed(0) + '万円' } }
        }
      }
    });
  }

  // Monthly OCC/ADR data
  const occArr = [];
  const adrArr = [];
  chartMonths.forEach(ym => {
    const stats = computeOverallStats(ym, area, false);
    occArr.push(stats.occ);
    adrArr.push(Math.round(stats.adr));
  });

  // Chart 2: OCC line + ADR bar
  destroyChart('dailyOccAdr');
  const ctx2 = document.getElementById('chartDailyOccAdr');
  if (ctx2) {
    const barColors = adrArr.map((_, i) => i < currentIdx ? 'rgba(74,144,217,0.2)' : i === currentIdx ? 'rgba(74,144,217,0.5)' : 'rgba(74,144,217,0.1)');
    allCharts['dailyOccAdr'] = new Chart(ctx2, {
      type: 'bar',
      data: {
        labels: chartLabels,
        datasets: [
          { type: 'line', label: 'OCC (%)', data: occArr, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.08)', fill: true, yAxisID: 'y1', tension: 0.4, pointBorderColor: CHART_COLORS.blue },
          { type: 'bar', label: 'ADR (¥)', data: adrArr, backgroundColor: barColors, yAxisID: 'y' }
        ]
      },
      options: {
        responsive: true,
        animation: { duration: 600, easing: 'easeOutQuart' },
        plugins: { legend: { display: true } },
        scales: {
          x: { grid: { display: false } },
          y: { position: 'left', beginAtZero: true, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } },
          y1: { position: 'right', min: 0, max: 110, grid: { drawOnChartArea: false }, title: { display: true, text: 'OCC (%)', font: { size: 11 } } }
        }
      }
    });
  }

  // Chart 3: Channel share horizontal bar
  destroyChart('dailyChannelShare');
  const ctx3 = document.getElementById('chartDailyChannelShare');
  if (ctx3) {
    const totalAllCh = channels.reduce((s, ch) => s + chartMonths.reduce((ss, ym) => ss + (monthChannelSales[ym][ch] || 0), 0), 0);
    const chPct = channels.map(ch => {
      const total = chartMonths.reduce((s, ym) => s + (monthChannelSales[ym][ch] || 0), 0);
      return totalAllCh > 0 ? (total / totalAllCh) * 100 : 0;
    });
    allCharts['dailyChannelShare'] = new Chart(ctx3, {
      type: 'bar',
      data: {
        labels: channels,
        datasets: [{ label: '売上構成比', data: chPct, backgroundColor: channelColors.slice(0, channels.length).map(c => c + 'CC'), hoverBackgroundColor: channelColors.slice(0, channels.length) }]
      },
      options: {
        indexAxis: 'y', responsive: true,
        animation: { duration: 600, easing: 'easeOutQuart' },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.parsed.x.toFixed(1) + '%' } } },
        scales: { x: { beginAtZero: true, max: 100, grid: { display: false }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false } } }
      }
    });
  }
}

function initReservationCharts() {
  const periodPred = getReservationPeriodPredicate();
  const statusFilter = document.getElementById('resv-status-filter').value;

  let filtered = [...reservations];
  if (statusFilter) filtered = filtered.filter(r => r.status === statusFilter);
  filtered = filtered.filter(periodPred);

  // チェックイン月別予約数（縦棒、月昇順）
  destroyChart('checkinMonthBD');
  const monthMap = {};
  const monthSalesMap = {};
  filtered.forEach(r => {
    const ym = getYearMonth(r.checkin);
    if (!ym) return;
    monthMap[ym] = (monthMap[ym] || 0) + 1;
    monthSalesMap[ym] = (monthSalesMap[ym] || 0) + (r.sales || 0);
  });
  const ctxM = document.getElementById('chartCheckinMonthBreakdown');
  if (ctxM) {
    const sortedMonths = Object.keys(monthMap).sort();
    const mLabels = sortedMonths.map(ym => {
      const [y, m] = ym.split('-');
      return `${parseInt(y, 10)}/${parseInt(m, 10)}月`;
    });
    const mData = sortedMonths.map(ym => monthMap[ym]);
    const mSales = sortedMonths.map(ym => monthSalesMap[ym] || 0);
    const mTotal = mData.reduce((s, v) => s + v, 0);
    allCharts['checkinMonthBD'] = new Chart(ctxM, {
      type: 'bar',
      data: {
        labels: mLabels,
        datasets: [
          {
            type: 'bar',
            label: 'GMV',
            data: mSales,
            backgroundColor: CHART_COLORS.blue + 'CC',
            hoverBackgroundColor: CHART_COLORS.blue,
            yAxisID: 'y1',
            order: 2,
          },
          {
            type: 'line',
            label: '予約数',
            data: mData,
            borderColor: CHART_COLORS.orange,
            backgroundColor: 'rgba(255,159,64,0.08)',
            tension: 0.4,
            yAxisID: 'y',
            order: 1,
            pointBackgroundColor: CHART_COLORS.orange,
          },
        ]
      },
      options: {
        responsive: true,
        animation: { duration: 600, easing: 'easeOutQuart' },
        plugins: {
          legend: { display: true, position: 'top', labels: { font: { size: 11 }, padding: 12 } },
          tooltip: {
            mode: 'index',
            callbacks: {
              label: ctx => {
                if (ctx.dataset.label === '予約数') {
                  const v = ctx.parsed.y;
                  const pct = mTotal > 0 ? ((v / mTotal) * 100).toFixed(1) : '0.0';
                  return `予約数: ${v}件 (${pct}%)`;
                }
                return `GMV: ${fmtYenFull(ctx.parsed.y)}`;
              }
            }
          }
        },
        scales: {
          x: { grid: { display: false } },
          y: { position: 'left', beginAtZero: true, title: { display: true, text: '予約数', font: { size: 11 } }, ticks: { precision: 0 } },
          y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'GMV (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v / 10000).toFixed(0) + '万' } }
        }
      }
    });
  }

  // 物件別予約数（横棒、件数降順）— グループ表示対応
  destroyChart('propertyBD');
  const propMap = {};
  const propSalesMap = {};
  const viewMode = currentFilters.reservationView || 'all';
  filtered.forEach(r => {
    let key;
    if (viewMode === 'grouped') {
      const code = r.propCode || r.property || '';
      const base = code ? getSeriesBase(code) : '';
      key = base || r.property || 'その他';
    } else {
      key = r.property || 'その他';
    }
    propMap[key] = (propMap[key] || 0) + 1;
    propSalesMap[key] = (propSalesMap[key] || 0) + (r.sales || 0);
  });
  const ctx4 = document.getElementById('chartPropertyBreakdown');
  if (ctx4) {
    const sortedAll = Object.entries(propMap).sort((a, b) => b[1] - a[1]);
    const TOP_N = 15;
    const top = sortedAll.slice(0, TOP_N);
    const rest = sortedAll.slice(TOP_N);
    const restCount = rest.reduce((s, [, v]) => s + v, 0);
    const restSales = rest.reduce((s, [k]) => s + (propSalesMap[k] || 0), 0);
    const sorted = restCount > 0 ? [...top, ['その他', restCount]] : top;
    const labels = sorted.map(([k]) => k);
    const data = sorted.map(([, v]) => v);
    const propSales = sorted.map(([k]) => k === 'その他' ? restSales : (propSalesMap[k] || 0));
    const total = data.reduce((s, v) => s + v, 0);
    // 全物件名を表示するためにキャンバス高さを物件数に応じて調整
    ctx4.parentElement.style.height = Math.max(240, labels.length * 22) + 'px';
    allCharts['propertyBD'] = new Chart(ctx4, {
      type: 'bar',
      data: {
        labels,
        datasets: [{
          label: '予約数',
          data,
          backgroundColor: CHART_COLORS.blue + 'CC',
          hoverBackgroundColor: CHART_COLORS.blue,
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 600, easing: 'easeOutQuart' },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => {
                const v = ctx.parsed.x;
                const pct = total > 0 ? ((v / total) * 100).toFixed(1) : '0.0';
                return [`${v}件 (${pct}%)`, `販売額: ${fmtYenFull(propSales[ctx.dataIndex])}`];
              }
            }
          }
        },
        scales: {
          x: { beginAtZero: true, grid: { display: false }, ticks: { precision: 0 } },
          y: { grid: { display: false }, ticks: { font: { size: 11 }, autoSkip: false } }
        }
      }
    });
  }
}

// 市場データ欠損チェック（物件マスタで使われている {area, ward, bedSuffix} × 3メトリクス）
function renderMarketDataCoverage(areaFilter) {
  const el = document.getElementById('market-data-coverage');
  if (!el) return;
  const adSheets = window._airdnaSheets || {};
  const bedLabel = (n) => n === 0 ? 'Studio' : (n === 4 ? '4BR+' : `${n}BR`);

  // 物件マスタから使用中の組み合わせを抽出
  const combos = new Map(); // key: area|ward|bed → {area, ward, wardEn, bed, bedSuffix, propNames}
  (propertyMaster || []).forEach(p => {
    if (p.excludeKpi) return;
    if (!p.area || !p.address) return;
    if (areaFilter && areaFilter !== '全体' && p.area !== areaFilter) return;
    const ward = extractWard(p.address);
    const wardEn = wardJpToAirdna(ward, p.area, p.address);
    const bed = layoutToBedrooms(p.layout);
    if (!ward || !wardEn || bed === null) return;
    const key = `${p.area}|${ward}|${bed}`;
    if (!combos.has(key)) combos.set(key, { area: p.area, ward, wardEn, bed, bedSuffix: bedLabel(bed), propNames: [] });
    combos.get(key).propNames.push(p.name);
  });

  // 各組で3シートの有無を確認
  const metrics = ['occupancy', 'rates_summary', 'revenue_summary'];
  const rows = [];
  combos.forEach(c => {
    const sheets = metrics.map(m => `AD_${c.area}_${c.wardEn}_${c.bedSuffix}_${m}`);
    const missing = sheets.filter(s => !adSheets[s]);
    if (missing.length === 0) return;
    // フォールバック判定
    const wardOnly = adSheets[`AD_${c.area}_${c.wardEn}_occupancy`];
    const fallback = wardOnly ? `${c.ward}全体` : `${c.area}全域`;
    rows.push({ ...c, missing, missingCount: missing.length, fallback });
  });

  if (combos.size === 0) {
    el.innerHTML = '';
    return;
  }
  // coverage情報がある場合はセクション自体も見えるように
  const section = document.getElementById('market-compare-section');
  if (section) section.style.display = '';
  if (rows.length === 0) {
    el.innerHTML = `<div style="background:#34C75912;border-left:3px solid #34C759;border-radius:6px;padding:8px 12px;font-size:12px;color:#1d1d1f;">
      ✓ 市場データ完備: 使用中の ${combos.size} 組（区×間取り）すべて取得済み
    </div>`;
    return;
  }

  // 欠損一覧テーブル（折りたたみ）
  rows.sort((a, b) => (a.area + a.ward).localeCompare(b.area + b.ward));
  const totalMissing = rows.reduce((s, r) => s + r.missingCount, 0);
  const tableRows = rows.map(r => `<tr style="border-bottom:1px solid #f0f0f0;">
    <td style="padding:4px 8px;">${r.area}</td>
    <td style="padding:4px 8px;">${r.ward}</td>
    <td style="padding:4px 8px;">${r.bedSuffix}</td>
    <td style="padding:4px 8px;text-align:right;color:#FF9500;font-weight:600;">${r.missingCount}/3</td>
    <td style="padding:4px 8px;font-size:11px;color:#86868b;">${r.missing.map(s => s.split('_').pop()).join(', ')}</td>
    <td style="padding:4px 8px;color:#86868b;">→ ${r.fallback}</td>
    <td style="padding:4px 8px;font-size:11px;color:#86868b;">${r.propNames.length}室</td>
  </tr>`).join('');

  el.innerHTML = `<details style="background:#FF950012;border-left:3px solid #FF9500;border-radius:6px;padding:8px 12px;">
    <summary style="font-size:12px;cursor:pointer;font-weight:600;color:#FF9500;">
      ⚠ 要取得データあり: ${rows.length}組 / ${combos.size}組（計${totalMissing}シート欠損）— 展開して詳細表示
    </summary>
    <div style="margin-top:8px;font-size:11px;color:#86868b;">
      対処: Chrome拡張で該当都市の「🛏 間取り別取得」を実行。既に存在するシートはスキップされます。
    </div>
    <div style="max-height:400px;overflow-y:auto;margin-top:8px;">
      <table style="width:100%;border-collapse:collapse;font-size:12px;background:white;border-radius:4px;">
        <thead><tr style="background:#f5f5f7;border-bottom:2px solid #e5e5ea;">
          <th style="padding:6px 8px;text-align:left;">エリア</th>
          <th style="padding:6px 8px;text-align:left;">区</th>
          <th style="padding:6px 8px;text-align:left;">間取り</th>
          <th style="padding:6px 8px;text-align:right;">欠損</th>
          <th style="padding:6px 8px;text-align:left;">欠損メトリクス</th>
          <th style="padding:6px 8px;text-align:left;">フォールバック先</th>
          <th style="padding:6px 8px;text-align:left;">該当物件</th>
        </tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>
  </details>`;
}

// 未来予約分析（AirDNA市場データ vs 自社実予約）
function renderFutureBookingAnalysis() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const area = currentFilters.revenueArea;

  // 対象物件
  const targetProps = (propertyMaster || []).filter(p => {
    if (area && area !== '全体' && p.area !== area) return false;
    if (p.excludeKpi) return false;
    return true;
  });
  const targetCodes = new Set(targetProps.map(p => p.propCode));
  const targetNames = new Set(targetProps.map(p => p.propName));
  const futureResv = reservations.filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
    return targetCodes.has(r.propCode) || targetNames.has(r.property);
  });

  const marketLookup = resolveAreaMarketLookup(area);

  // 1. 未来OCC推移（次90日）: 自社（日別の埋まり率%）vs 市場（月次OCC）
  destroyChart('futureOcc');
  const occLabels = [];
  const myOccFuture = [];
  const mktOccFuture = [];
  const propCount = targetProps.length || 1;
  for (let i = 0; i < 90; i++) {
    const d = new Date(today); d.setDate(d.getDate() + i);
    occLabels.push(`${d.getMonth() + 1}/${d.getDate()}`);
    const ds = d.toISOString().split('T')[0];
    // 自社: その日に予約が入っている物件数 / 総物件数
    const bookedCount = targetProps.reduce((s, p) => {
      const has = futureResv.some(r =>
        (r.propCode === p.propCode || r.property === p.propName) &&
        r.checkin <= ds && ds < r.checkout
      );
      return s + (has ? 1 : 0);
    }, 0);
    myOccFuture.push(Math.round((bookedCount / propCount) * 1000) / 10);
    // 市場: 月次OCC
    const ym = ds.slice(0, 7);
    const mv = marketLookup.hasData ? marketLookup.occ(ym) : null;
    mktOccFuture.push(mv !== null ? Math.round(mv) : null);
  }
  // 自社OCCを7日移動平均でならす
  const mySmoothed = myOccFuture.map((_, i) => {
    const start = Math.max(0, i - 3);
    const end = Math.min(myOccFuture.length, i + 4);
    const slice = myOccFuture.slice(start, end);
    return Math.round(slice.reduce((s, v) => s + v, 0) / slice.length);
  });
  const ctxFO = document.getElementById('chartFutureOcc');
  if (ctxFO) {
    allCharts['futureOcc'] = new Chart(ctxFO, {
      type: 'line',
      data: { labels: occLabels, datasets: [
        { label: '自社OCC (7日移動平均)', data: mySmoothed, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.1)', fill: true, tension: 0.3, pointRadius: 0 },
        { label: `市場平均OCC (${marketLookup.matched || '—'})`, data: mktOccFuture, borderColor: CHART_COLORS.orange, borderDash: [6, 3], backgroundColor: 'transparent', tension: 0.3, pointRadius: 0, spanGaps: true },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + (ctx.parsed.y !== null ? ctx.parsed.y + '%' : '—') } } },
        scales: { x: { ticks: { maxTicksLimit: 12 } }, y: { beginAtZero: true, max: 100, ticks: { callback: v => v + '%' } } }
      }
    });
  }

  // 2. 未来ADR推移（次180日）: 自社予約実績ADR vs 市場平均（月次×曜日バンプ）
  destroyChart('futureAdr');
  const adrLabels = [];
  const myAdrFuture = [];
  const mktAdrFuture = [];
  for (let i = 0; i < 180; i++) {
    const d = new Date(today); d.setDate(d.getDate() + i);
    adrLabels.push(`${d.getMonth() + 1}/${d.getDate()}`);
    const ds = d.toISOString().split('T')[0];
    const dow = d.getDay();

    // 自社: その日を含む予約のADR平均
    const dayResv = futureResv.filter(r => r.checkin <= ds && ds < r.checkout && r.nights > 0);
    if (dayResv.length > 0) {
      const adrs = dayResv.map(r => Math.round(((r.sales || 0) - (r.cleaningFee || 0)) / r.nights)).filter(v => v > 0);
      myAdrFuture.push(adrs.length ? Math.round(adrs.reduce((s, v) => s + v, 0) / adrs.length) : null);
    } else {
      myAdrFuture.push(null);
    }

    // 市場: 月次ADRに曜日バンプ
    const ym = ds.slice(0, 7);
    const baseAdr = marketLookup.hasData ? marketLookup.adr(ym) : null;
    if (baseAdr !== null) {
      const weekendBump = (dow === 5 || dow === 6) ? 1.25 : 0.95;
      mktAdrFuture.push(Math.round(baseAdr * weekendBump));
    } else {
      mktAdrFuture.push(null);
    }
  }
  const ctxFA = document.getElementById('chartFutureAdr');
  if (ctxFA) {
    allCharts['futureAdr'] = new Chart(ctxFA, {
      type: 'line',
      data: { labels: adrLabels, datasets: [
        { label: '自社ADR', data: myAdrFuture, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.15)', tension: 0.3, pointRadius: 1, spanGaps: false },
        { label: `市場平均ADR (${marketLookup.matched || '—'})`, data: mktAdrFuture, borderColor: CHART_COLORS.orange, borderDash: [6, 3], backgroundColor: 'transparent', tension: 0.3, pointRadius: 0, spanGaps: true },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + (ctx.parsed.y !== null ? '¥' + ctx.parsed.y.toLocaleString() : '—') } } },
        scales: { x: { ticks: { maxTicksLimit: 15 } }, y: { beginAtZero: true, ticks: { callback: v => '¥' + (v / 1000).toFixed(0) + 'k' } } }
      }
    });
  }

  // 3. リードタイム分布: 自社の過去1年の実データ（市場側はAirDNAで未取得のため自社のみ）
  destroyChart('leadTime');
  const ltLabels = ['0-7日', '8-14日', '15-30日', '31-60日', '61-90日', '91日〜'];
  const ltBuckets = [
    { min: 0, max: 7 }, { min: 8, max: 14 }, { min: 15, max: 30 },
    { min: 31, max: 60 }, { min: 61, max: 90 }, { min: 91, max: Infinity },
  ];
  const ltBucketCounts = ltBuckets.map(() => 0);
  const oneYearAgo = new Date(today); oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
  const pastResvForLt = reservations.filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
    if (!(targetCodes.has(r.propCode) || targetNames.has(r.property))) return false;
    if (!r.date || !r.checkin) return false;
    const ci = new Date(r.checkin);
    return ci >= oneYearAgo && ci <= today;
  });
  pastResvForLt.forEach(r => {
    const lead = Math.floor((new Date(r.checkin) - new Date(r.date)) / 86400000);
    const idx = ltBuckets.findIndex(b => lead >= b.min && lead <= b.max);
    if (idx >= 0) ltBucketCounts[idx]++;
  });
  const ltTotal = pastResvForLt.length || 1;
  const myLt = ltBucketCounts.map(c => Math.round((c / ltTotal) * 100));
  const ctxLT = document.getElementById('chartLeadTime');
  if (ctxLT) {
    allCharts['leadTime'] = new Chart(ctxLT, {
      type: 'bar',
      data: { labels: ltLabels, datasets: [
        { label: '自社（過去1年 / 該当エリア）', data: myLt, backgroundColor: CHART_COLORS.blue + 'CC' },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y + '%' } } },
        scales: { x: { grid: { display: false } }, y: { beginAtZero: true, ticks: { callback: v => v + '%' } } }
      }
    });
  }
}

function initRevenueCharts() {
  const months = getSelectedMonths('revenue');
  const monthSet = new Set(months);
  const area = currentFilters.revenueArea;
  const excludeKpi = document.getElementById('excludeKpiToggleRev') && document.getElementById('excludeKpiToggleRev').checked;

  // Channel revenue bar
  destroyChart('channelRev');
  const confirmedResv = reservations.filter(r => {
    if (r.status === 'システムキャンセル') return false;
    if (!monthSet.has(getYearMonth(r.checkin))) return false;
    if (area !== '全体') {
      const prop = findPropByReservation(r);
      if (!prop || prop.area !== area) return false;
    }
    return true;
  });
  const channelMap = {};
  confirmedResv.forEach(r => {
    const ch = r.channel || 'その他';
    channelMap[ch] = (channelMap[ch] || 0) + r.sales;
  });

  const ctx2 = document.getElementById('chartChannelRevenue');
  if (ctx2) {
    const labels = Object.keys(channelMap).sort((a, b) => channelMap[b] - channelMap[a]);
    const data = labels.map(l => channelMap[l]);
    const colors = PALETTE;
    allCharts['channelRev'] = new Chart(ctx2, {
      type: 'bar',
      data: { labels, datasets: [{ label: '販売金額', data, backgroundColor: colors.slice(0, labels.length) }] },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } } } }
    });
  }

  // Area RevPAR
  destroyChart('areaRevpar');
  const areas = ['大阪', '京都', '東京'];
  const areaRevpars = areas.map(a => {
    const stats = computeOverallStatsMulti(months, a, excludeKpi);
    return Math.round(stats.revpar);
  });

  const ctx4 = document.getElementById('chartAreaRevpar');
  if (ctx4) {
    allCharts['areaRevpar'] = new Chart(ctx4, {
      type: 'bar',
      data: { labels: areas, datasets: [{ label: 'RevPAR', data: areaRevpars, backgroundColor: PALETTE.slice(0, 3) }] },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
    });
  }

  // 予約Window別 ADR & 件数
  destroyChart('windowAdr');
  destroyChart('windowCount');
  const extra = getExtraFilters('revenue');
  const extraFilteredNames = new Set(filterPropertiesByArea(area, extra).map(p => p.name));
  const windowResv = reservations.filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return false;
    if (!r.date || !r.checkin || !r.nights || r.nights <= 0) return false;
    if (!monthSet.has(getYearMonth(r.checkin))) return false;
    const prop = findPropByReservation(r);
    if (!prop) return false;
    if (!extraFilteredNames.has(prop.name)) return false;
    if (excludeKpi && prop.excludeKpi) return false;
    return true;
  });

  const windowBuckets = [
    { label: '0〜3日', min: 0, max: 3 },
    { label: '4〜7日', min: 4, max: 7 },
    { label: '8〜14日', min: 8, max: 14 },
    { label: '15〜30日', min: 15, max: 30 },
    { label: '31〜60日', min: 31, max: 60 },
    { label: '61〜90日', min: 61, max: 90 },
    { label: '91日〜', min: 91, max: Infinity },
  ];
  const bucketData = windowBuckets.map(() => ({ sales: 0, nights: 0, count: 0 }));
  windowResv.forEach(r => {
    const lead = Math.max(0, Math.floor((new Date(r.checkin) - new Date(r.date)) / 86400000));
    const idx = windowBuckets.findIndex(b => lead >= b.min && lead <= b.max);
    if (idx >= 0) {
      bucketData[idx].sales += r.sales || 0;
      bucketData[idx].nights += r.nights || 0;
      bucketData[idx].count++;
    }
  });

  const wLabels = windowBuckets.map(b => b.label);
  const wAdr = bucketData.map(d => d.nights > 0 ? Math.round(d.sales / d.nights) : 0);
  const wCount = bucketData.map(d => d.count);

  const ctxWA = document.getElementById('chartWindowAdr');
  if (ctxWA) {
    allCharts['windowAdr'] = new Chart(ctxWA, {
      type: 'bar',
      data: { labels: wLabels, datasets: [
        { type: 'bar', label: 'ADR', data: wAdr, backgroundColor: PALETTE.slice(0, wLabels.length).map(c => c + 'CC'), yAxisID: 'y' },
        { type: 'line', label: '予約件数', data: wCount, borderColor: CHART_COLORS.orange, backgroundColor: 'transparent', yAxisID: 'y1', tension: 0.3, pointRadius: 4 },
      ] },
      options: { responsive: true, plugins: { legend: { display: true } }, scales: {
        x: { grid: { display: false } },
        y: { position: 'left', beginAtZero: true, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } },
        y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: '予約件数', font: { size: 11 } } },
      } }
    });
  }

  const ctxWC = document.getElementById('chartWindowCount');
  if (ctxWC) {
    const avgSales = bucketData.map(d => d.count > 0 ? Math.round(d.sales / d.count) : 0);
    allCharts['windowCount'] = new Chart(ctxWC, {
      type: 'bar',
      data: { labels: wLabels, datasets: [
        { type: 'bar', label: '予約件数', data: wCount, backgroundColor: CHART_COLORS.blue + '99' },
        { type: 'line', label: '平均売上/件', data: avgSales, borderColor: CHART_COLORS.orange, backgroundColor: 'transparent', yAxisID: 'y1', tension: 0.3, pointRadius: 4 },
      ] },
      options: { responsive: true, plugins: { legend: { display: true } }, scales: {
        x: { grid: { display: false } },
        y: { position: 'left', beginAtZero: true, title: { display: true, text: '予約件数', font: { size: 11 } } },
        y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: '平均売上 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } },
      } }
    });
  }

  // ── 未来予約分析（モックデータ） ──
  renderFutureBookingAnalysis();

  // ── 市場比較（AirDNA） ──
  const marketSection = document.getElementById('market-compare-section');
  const adSheets = window._airdnaSheets || {};
  const adSheetNames = Object.keys(adSheets);

  // 市場データ欠損チェック
  renderMarketDataCoverage(area);

  // エリアフィルタに対応するシートを選択（AD_{エリア}_xxx）
  const mktArea = area === '全体' ? null : area;
  const matchingNames = adSheetNames.filter(name => {
    if (!mktArea) return true;
    return name.includes(`_${mktArea}_`);
  });

  // OCC/ADR/Revenue シートを探す
  const findSheet = (keyword) => {
    const name = matchingNames.find(n => n.toLowerCase().includes(keyword));
    return name ? adSheets[name] : null;
  };
  const occSheet = findSheet('occupancy_');  // occupancy（futureではなく過去実績）
  const adrSheet = findSheet('ratebydailyaverage');
  const revSheet = findSheet('revenueaverage');

  if (marketSection && (occSheet || adrSheet)) {
    marketSection.style.display = '';

    // 月別データを構築
    const mktByMonth = {};
    if (occSheet) occSheet.forEach(r => {
      const ym = (r['Date'] || '').slice(0, 7);
      if (!ym) return;
      if (!mktByMonth[ym]) mktByMonth[ym] = {};
      mktByMonth[ym].occ = parseFloat(r['Occupancy']) || 0;
    });
    if (adrSheet) adrSheet.forEach(r => {
      const ym = (r['Date'] || '').slice(0, 7);
      if (!ym) return;
      if (!mktByMonth[ym]) mktByMonth[ym] = {};
      mktByMonth[ym].adr = parseFloat(r['Daily Rate']) || 0;
    });
    if (revSheet) revSheet.forEach(r => {
      const ym = (r['Date'] || '').slice(0, 7);
      if (!ym) return;
      if (!mktByMonth[ym]) mktByMonth[ym] = {};
      mktByMonth[ym].revenue = parseFloat(r['Revenue']) || 0;
    });

    // 12ヶ月の軸
    const now = new Date();
    const chartMonths = [];
    for (let i = -11; i <= 0; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      chartMonths.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
    }
    const chartLabels = chartMonths.map(ym => parseInt(ym.split('-')[1]) + '月');

    const myOcc = [], myAdr = [];
    chartMonths.forEach(ym => {
      const s = computeOverallStatsMulti([ym], area, excludeKpi, extra);
      myOcc.push(Math.round(s.occ * 10) / 10);
      myAdr.push(Math.round(s.adr));
    });
    const mktOcc = chartMonths.map(ym => mktByMonth[ym]?.occ != null ? Math.round(mktByMonth[ym].occ * 10) / 10 : null);
    const mktAdr = chartMonths.map(ym => mktByMonth[ym]?.adr != null ? Math.round(mktByMonth[ym].adr) : null);

    // OCC比較チャート
    destroyChart('marketOcc');
    const ctxMO = document.getElementById('chartMarketOcc');
    if (ctxMO) {
      allCharts['marketOcc'] = new Chart(ctxMO, {
        type: 'line',
        data: { labels: chartLabels, datasets: [
          { label: '自社OCC', data: myOcc, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.1)', fill: true, tension: 0.3 },
          { label: '市場平均OCC', data: mktOcc, borderColor: CHART_COLORS.orange, borderDash: [6, 3], backgroundColor: 'transparent', tension: 0.3 },
        ]},
        options: { responsive: true, plugins: { legend: { display: true }, tooltip: salesChartTooltip }, scales: { y: { beginAtZero: true, max: 100, ticks: { callback: v => v + '%' } } } }
      });
    }

    // ADR比較チャート
    destroyChart('marketAdr');
    const ctxMA = document.getElementById('chartMarketAdr');
    if (ctxMA) {
      allCharts['marketAdr'] = new Chart(ctxMA, {
        type: 'line',
        data: { labels: chartLabels, datasets: [
          { label: '自社ADR', data: myAdr, borderColor: CHART_COLORS.blue, backgroundColor: 'rgba(74,144,217,0.1)', fill: true, tension: 0.3 },
          { label: '市場平均ADR', data: mktAdr, borderColor: CHART_COLORS.orange, borderDash: [6, 3], backgroundColor: 'transparent', tension: 0.3 },
        ]},
        options: { responsive: true, plugins: { legend: { display: true }, tooltip: salesChartTooltip }, scales: { y: { beginAtZero: true, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
      });
    }

    // 直近月のインデックスカード
    const latestYm = chartMonths[chartMonths.length - 1];
    const latestMkt = mktByMonth[latestYm];
    const latestMy = computeOverallStatsMulti([latestYm], area, excludeKpi, extra);
    const indexEl = document.getElementById('market-index-cards');
    if (indexEl && latestMkt) {
      const mOcc = latestMkt.occ || 0;
      const mAdr = latestMkt.adr || 0;
      const occIdx = mOcc > 0 ? Math.round((latestMy.occ / mOcc) * 100) : '-';
      const adrIdx = mAdr > 0 ? Math.round((latestMy.adr / mAdr) * 100) : '-';
      const mRevpar = mAdr * mOcc / 100;
      const revparIdx = mRevpar > 0 ? Math.round((latestMy.revpar / mRevpar) * 100) : '-';
      const idxColor = v => v >= 100 ? '#34C759' : v >= 80 ? '#FF9500' : '#FF3B30';
      indexEl.innerHTML = `<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:12px;">
        <div class="kpi-card"><div class="label">OCCインデックス</div><div class="value" style="color:${idxColor(occIdx)}">${occIdx}</div><div class="sub">市場平均=100</div></div>
        <div class="kpi-card"><div class="label">ADRインデックス</div><div class="value" style="color:${idxColor(adrIdx)}">${adrIdx}</div><div class="sub">市場平均=100</div></div>
        <div class="kpi-card"><div class="label">RevPARインデックス</div><div class="value" style="color:${idxColor(revparIdx)}">${revparIdx}</div><div class="sub">市場平均=100</div></div>
      </div>`;
    }
  } else if (marketSection) {
    marketSection.style.display = 'none';
  }

  // ── 物件スコアカード ──
  const scorecardEl = document.getElementById('prop-scorecard');
  if (scorecardEl) {
    const scoreProps = filterPropertiesByArea(area, extra)
      .filter(p => !excludeKpi || !p.excludeKpi)
      .filter(p => p.status === '稼働中');
    const curMonths = getSelectedMonths('revenue');
    const scMode = currentFilters.scorecardMode || 'all';
    const scDayType = currentFilters.scorecardDayType || 'total';

    // スコア算出ヘルパー
    function calcScore(propNames, label, areaName) {
      let totalSales = 0, totalTarget = 0;
      propNames.forEach(pn => {
        const p = findPropByName(pn);
        curMonths.forEach(m => {
          const s = computePropertyStats(pn, m);
          if (s) totalSales += s.sales;
          if (p) totalTarget += getTargetForProperty(p, parseInt(m.split('-')[1], 10));
        });
      });
      const targetPct = totalTarget > 0 ? (totalSales / totalTarget) * 100 : null;
      const pace = computePaceReport(propNames);
      const occFor = (b) => {
        if (!b) return 0;
        if (scDayType === 'weekday') return b.weekday.avail > 0 ? (b.weekday.nights / b.weekday.avail) * 100 : 0;
        if (scDayType === 'holiday') return b.holiday.avail > 0 ? (b.holiday.nights / b.holiday.avail) * 100 : 0;
        return (b.weekday.nights + b.holiday.nights) / Math.max(b.weekday.avail + b.holiday.avail, 1) * 100;
      };
      const occ30 = occFor(pace[0]), occ60 = occFor(pace[1]), occ90 = occFor(pace[2]);
      function grade(val, green, yellow) { return val >= green ? 2 : val >= yellow ? 1 : 0; }
      const tGrade = targetPct !== null ? grade(targetPct, 100, 80) : -1;
      const g30 = grade(occ30, 80, 50), g60 = grade(occ60, 60, 30), g90 = grade(occ90, 20, 10);
      const redCount = [tGrade, g30, g60, g90].filter(g => g === 0).length;
      const allGreen = [tGrade, g30, g60, g90].every(g => g === 2 || g === -1);
      const noRed = redCount === 0;
      const overall = allGreen ? '◎' : noRed ? '○' : redCount <= 2 ? '△' : '✕';
      const overallColor = allGreen ? '#34C759' : noRed ? '#007AFF' : redCount <= 2 ? '#FF9500' : '#FF3B30';
      return { name: label, propNames, area: areaName, sales: totalSales, targetPct, occ30, occ60, occ90, tGrade, g30, g60, g90, overall, overallColor };
    }

    let scoreRows;
    if (scMode === 'grouped') {
      const groups = {};
      scoreProps.forEach(p => {
        const base = getSeriesBase(p.name);
        if (!groups[base]) groups[base] = { props: [], area: p.area };
        groups[base].props.push(p);
      });
      scoreRows = Object.entries(groups)
        .filter(([_, g]) => g.props.length >= 2)
        .map(([base, g]) => calcScore(g.props.map(p => p.name), base + '（' + g.props.length + '室）', g.area));
    } else {
      scoreRows = scoreProps.map(p => calcScore([p.name], p.name, p.area));
    }

    const overallOrder = { '✕': 0, '△': 1, '○': 2, '◎': 3 };
    scoreRows.sort((a, b) => (overallOrder[a.overall] || 0) - (overallOrder[b.overall] || 0));

    const gColors = ['#FF3B30', '#FF9500', '#34C759'];
    const gBg = ['#FF3B3015', '#FF950015', '#34C75915'];
    function cellStyle(g) { return `color:${gColors[g]};font-weight:600;background:${gBg[g]};`; }

    let scHtml = `<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:12px;font-size:12px;">
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#34C759;"></span> 好調</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#FF9500;"></span> 注意</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#FF3B30;"></span> 要改善</span>
      <span style="color:#86868b;">基準: 目標≧100%/80% ｜ 30日≧80%/50% ｜ 60日≧60%/30% ｜ 90日≧20%/10%</span>
    </div>`;
    const dayLabel = scDayType === 'weekday' ? '平日' : scDayType === 'holiday' ? '休日' : '';
    const dayTag = dayLabel ? ` <span style="font-size:10px;color:#5856D6;">${dayLabel}</span>` : '';
    scHtml += `<div class="table-wrap"><table style="width:100%;border-collapse:collapse;font-size:12px;">
      <thead><tr style="border-bottom:2px solid #e5e5ea;">
        <th style="text-align:left;padding:6px 4px;">物件</th>
        <th style="text-align:left;padding:6px 4px;">エリア</th>
        <th style="text-align:right;padding:6px 4px;">目標達成</th>
        <th style="text-align:right;padding:6px 4px;">売上</th>
        <th style="text-align:center;padding:6px 4px;">30日${dayTag}</th>
        <th style="text-align:center;padding:6px 4px;">60日${dayTag}</th>
        <th style="text-align:center;padding:6px 4px;">90日${dayTag}</th>
        <th style="text-align:center;padding:6px 4px;">総合</th>
      </tr></thead><tbody>`;

    scoreRows.forEach((r, idx) => {
      const clickName = r.propNames.length === 1 ? r.propNames[0] : r.propNames[0];
      const safeName = clickName.replace(/'/g, "\\'");
      scHtml += `<tr style="border-bottom:1px solid #f0f0f0;">
        <td style="padding:5px 4px;font-weight:600;"><a href="#" style="color:#007aff;text-decoration:none;" onclick="event.preventDefault();toggleScorecardDetail('${safeName}', ${idx})">${r.name}</a></td>
        <td style="padding:5px 4px;color:#86868b;">${r.area}</td>
        <td style="text-align:right;padding:5px 4px;${r.tGrade >= 0 ? cellStyle(r.tGrade) : ''}border-radius:4px;">${r.targetPct !== null ? Math.round(r.targetPct) + '%' : '-'}</td>
        <td style="text-align:right;padding:5px 4px;">${fmtYen(r.sales)}</td>
        <td style="text-align:center;padding:5px 4px;${cellStyle(r.g30)}border-radius:4px;">${r.occ30.toFixed(0)}%</td>
        <td style="text-align:center;padding:5px 4px;${cellStyle(r.g60)}border-radius:4px;">${r.occ60.toFixed(0)}%</td>
        <td style="text-align:center;padding:5px 4px;${cellStyle(r.g90)}border-radius:4px;">${r.occ90.toFixed(0)}%</td>
        <td style="text-align:center;padding:5px 4px;font-size:16px;font-weight:700;color:${r.overallColor};">${r.overall}</td>
      </tr>
      <tr class="sc-detail-row" id="sc-detail-row-${idx}" style="display:none;"><td colspan="8" style="padding:0;background:#fafafa;"><div id="sc-detail-inner-${idx}"></div></td></tr>`;
    });
    scHtml += '</tbody></table></div>';
    scorecardEl.innerHTML = scHtml;
    setTimeout(initSortableHeaders, 50);
  }

  // ── 先行予約ペースレポート（全物件横断） ──
  const paceAllEl = document.getElementById('pace-report-all');
  const paceByPropEl = document.getElementById('pace-report-by-prop');
  if (paceAllEl) {
    const activePropNames = filterPropertiesByArea(area, extra)
      .filter(p => !excludeKpi || !p.excludeKpi)
      .filter(p => p.status === '稼働中')
      .map(p => p.name);
    const paceAll = computePaceReport(activePropNames);
    paceAllEl.innerHTML = renderPaceReportHtml(paceAll, '全体');

    // ペース4象限（0-30日バケット基準）
    if (paceByPropEl) {
      const hdThresh = 70; // 休日OCC高の閾値
      const wdThresh = 40; // 平日OCC高の閾値
      const quadrants = {
        hotBoth:  { label: '全体好調', icon: '🔥', action: '全体値上げ検討', color: '#34C759', bg: '#34C75912', items: [] },
        hotWdOnly:{ label: '平日だけ好調', icon: '🟡', action: '休日値下げ検討', color: '#FF9500', bg: '#FF950012', items: [] },
        hotHdOnly:{ label: '休日偏り', icon: '⚠', action: '休日値上げ＋平日値下げ', color: '#5856D6', bg: '#5856D612', items: [] },
        cold:     { label: '予約不足', icon: '❄', action: 'リスティング改善', color: '#FF3B30', bg: '#FF3B3012', items: [] },
      };
      activePropNames.forEach(pn => {
        const pace = computePaceReport([pn]);
        const b0 = pace[0];
        const hdHigh = b0.holiday.occ >= hdThresh;
        const wdHigh = b0.weekday.occ >= wdThresh;
        const qId = hdHigh && wdHigh ? 'hotBoth' : !hdHigh && wdHigh ? 'hotWdOnly' : hdHigh && !wdHigh ? 'hotHdOnly' : 'cold';
        const prop = findPropByName(pn);
        quadrants[qId].items.push({
          name: pn, area: prop ? prop.area : '',
          hdOcc: b0.holiday.occ, wdOcc: b0.weekday.occ,
          hdAdr: b0.holiday.adr, wdAdr: b0.weekday.adr,
        });
      });

      let qHtml = `<div style="margin-top:16px;font-size:13px;font-weight:600;color:#1d1d1f;margin-bottom:10px;">ペース4象限 <span style="font-size:11px;font-weight:400;color:#86868b;">（0〜30日 / 休日≧${hdThresh}% 平日≧${wdThresh}%）</span></div>`;
      qHtml += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">';
      Object.values(quadrants).forEach(q => {
        const items = q.items.sort((a, b) => (b.hdOcc + b.wdOcc) - (a.hdOcc + a.wdOcc));
        qHtml += `<div style="background:${q.bg};border:1px solid ${q.color}25;border-radius:10px;padding:12px 14px;">
          <div style="font-weight:700;font-size:13px;color:${q.color};margin-bottom:2px;">${q.icon} ${q.label}</div>
          <div style="font-size:11px;color:#86868b;margin-bottom:8px;">→ ${q.action}</div>
          <div style="font-size:18px;font-weight:700;color:#1d1d1f;margin-bottom:6px;">${items.length}件</div>`;
        if (items.length > 0) {
          qHtml += `<table style="width:100%;border-collapse:collapse;font-size:11px;">
            <thead><tr style="border-bottom:1px solid ${q.color}20;">
              <th style="text-align:left;padding:3px 2px;color:#86868b;font-weight:500;">物件</th>
              <th style="text-align:right;padding:3px 2px;color:#86868b;font-weight:500;">休日</th>
              <th style="text-align:right;padding:3px 2px;color:#86868b;font-weight:500;">平日</th>
            </tr></thead><tbody>`;
          items.slice(0, 8).forEach(it => {
            qHtml += `<tr style="border-bottom:1px solid ${q.color}10;">
              <td style="padding:3px 2px;">${it.name}<span style="color:#aaa;font-size:9px;"> ${it.area}</span></td>
              <td style="text-align:right;padding:3px 2px;font-weight:600;">${it.hdOcc.toFixed(0)}%</td>
              <td style="text-align:right;padding:3px 2px;">${it.wdOcc.toFixed(0)}%</td>
            </tr>`;
          });
          if (items.length > 8) qHtml += `<tr><td colspan="3" style="padding:3px 2px;color:#86868b;text-align:center;">他${items.length - 8}件</td></tr>`;
          qHtml += `</tbody></table>`;
        }
        qHtml += `</div>`;
      });
      qHtml += '</div>';
      paceByPropEl.innerHTML = qHtml;
    }
  }

  // ── OCC × ADR 4象限マトリクス ──
  destroyChart('occAdrMatrix');
  const overall = computeOverallStatsMulti(months, area, excludeKpi, extra);
  const propStats = (overall.stats || []).filter(s => s.nights > 0);
  const avgOcc = overall.occ;
  const avgAdr = overall.adr;

  // Quadrant classification
  const QUADRANT_DEFS = [
    { id: 'star',    label: '★ スター（高OCC × 高ADR）',     color: '#34C759', bg: '#34C75918', check: (o, a) => o >= avgOcc && a >= avgAdr },
    { id: 'raise',   label: '↑ 値上げ余地（高OCC × 低ADR）', color: '#FF9500', bg: '#FF950018', check: (o, a) => o >= avgOcc && a < avgAdr },
    { id: 'promo',   label: '↓ 値下げ検討（低OCC × 高ADR）',   color: '#5856D6', bg: '#5856D618', check: (o, a) => o < avgOcc && a >= avgAdr },
    { id: 'problem', label: '⚠ 要改善（低OCC × 低ADR）',     color: '#FF3B30', bg: '#FF3B3018', check: (o, a) => o < avgOcc && a < avgAdr },
  ];

  const quadrants = { star: [], raise: [], promo: [], problem: [] };
  propStats.forEach(s => {
    const q = QUADRANT_DEFS.find(d => d.check(s.occ, s.adr));
    if (q) quadrants[q.id].push(s);
  });

  // Area color mapping for scatter
  const areaColorMap = { '大阪': CHART_COLORS.blue, '京都': CHART_COLORS.green, '東京': CHART_COLORS.orange, 'その他': CHART_COLORS.purple };
  const maxSales = Math.max(...propStats.map(s => s.sales), 1);

  const ctxMatrix = document.getElementById('chartOccAdrMatrix');
  if (ctxMatrix && propStats.length > 0) {
    // Group by area for legend
    const areaGroups = {};
    propStats.forEach(s => {
      const prop = findPropByName(s.name);
      const a = prop ? prop.area : 'その他';
      if (!areaGroups[a]) areaGroups[a] = [];
      areaGroups[a].push({ x: s.occ, y: Math.round(s.adr), r: Math.max(5, Math.sqrt(s.sales / maxSales) * 30), name: s.name, sales: s.sales, occ: s.occ, adr: s.adr });
    });

    const datasets = Object.entries(areaGroups).map(([a, pts]) => ({
      label: a,
      data: pts,
      backgroundColor: (areaColorMap[a] || CHART_COLORS.purple) + '88',
      borderColor: areaColorMap[a] || CHART_COLORS.purple,
      borderWidth: 1.5,
    }));

    // Quadrant line plugin
    const quadrantPlugin = {
      id: 'quadrantLines',
      beforeDraw(chart) {
        const { ctx: c, chartArea: { left, right, top, bottom }, scales: { x, y } } = chart;
        const xPx = x.getPixelForValue(avgOcc);
        const yPx = y.getPixelForValue(avgAdr);
        c.save();
        c.setLineDash([6, 4]);
        c.strokeStyle = '#86868b';
        c.lineWidth = 1;
        // Vertical line (avg OCC)
        c.beginPath(); c.moveTo(xPx, top); c.lineTo(xPx, bottom); c.stroke();
        // Horizontal line (avg ADR)
        c.beginPath(); c.moveTo(left, yPx); c.lineTo(right, yPx); c.stroke();
        c.setLineDash([]);
        // Labels
        c.font = '10px -apple-system, sans-serif';
        c.fillStyle = '#86868b';
        c.textAlign = 'center';
        c.fillText('平均OCC ' + avgOcc.toFixed(1) + '%', xPx, bottom + 14);
        c.save();
        c.translate(left - 8, yPx);
        c.rotate(-Math.PI / 2);
        c.fillText('平均ADR ¥' + Math.round(avgAdr).toLocaleString(), 0, 0);
        c.restore();
        c.restore();
      }
    };

    allCharts['occAdrMatrix'] = new Chart(ctxMatrix, {
      type: 'bubble',
      data: { datasets },
      plugins: [quadrantPlugin],
      options: {
        responsive: true,
        animation: { duration: 600, easing: 'easeOutQuart' },
        plugins: {
          legend: { position: 'top', labels: { font: { size: 11 }, padding: 16 } },
          tooltip: {
            callbacks: {
              label: ctx => {
                const d = ctx.raw;
                return `${d.name}  OCC: ${d.occ.toFixed(1)}%  ADR: ¥${Math.round(d.adr).toLocaleString()}  売上: ${fmtYen(d.sales)}`;
              }
            }
          }
        },
        scales: {
          x: { title: { display: true, text: 'OCC（稼働率 %）', font: { size: 12 } }, min: 0, max: 100, ticks: { callback: v => v + '%' } },
          y: { title: { display: true, text: 'ADR（平均客室単価 ¥）', font: { size: 12 } }, beginAtZero: true, ticks: { callback: v => '¥' + v.toLocaleString() } },
        }
      }
    });
  }

  // Quadrant summary cards
  const summaryEl = document.getElementById('occ-adr-quadrant-summary');
  if (summaryEl) {
    summaryEl.innerHTML = QUADRANT_DEFS.map(q => {
      const items = quadrants[q.id];
      const names = items.map(s => s.name).join(', ') || '-';
      return `<div style="background:${q.bg};border:1px solid ${q.color}30;border-radius:8px;padding:10px 14px;">
        <div style="font-weight:600;font-size:13px;color:${q.color};margin-bottom:4px;">${q.label}</div>
        <div style="font-size:20px;font-weight:700;color:#1d1d1f;">${items.length}件</div>
        <div style="font-size:11px;color:#86868b;margin-top:4px;line-height:1.4;word-break:break-all;">${names}</div>
      </div>`;
    }).join('');
  }

  // Quadrant detail table
  const occAdrTable = document.getElementById('occ-adr-table');
  if (occAdrTable) {
    const labelMap = { star: '★ スター', raise: '↑ 値上げ余地', promo: '↓ 値下げ検討', problem: '⚠ 要改善' };
    const colorMap = { star: '#34C759', raise: '#FF9500', promo: '#5856D6', problem: '#FF3B30' };
    const sorted = propStats.slice().sort((a, b) => b.revpar - a.revpar);
    occAdrTable.innerHTML = sorted.map(s => {
      const prop = findPropByName(s.name);
      const areaName = prop ? prop.area : '-';
      const qId = QUADRANT_DEFS.find(d => d.check(s.occ, s.adr))?.id || 'problem';
      return `<tr>
        <td>${s.name}</td><td>${areaName}</td>
        <td>${fmtPct(s.occ)}</td>
        <td class="text-right">${fmtYenFull(Math.round(s.adr))}</td>
        <td class="text-right">${fmtYenFull(Math.round(s.revpar))}</td>
        <td class="text-right">${fmtYen(s.sales)}</td>
        <td><span style="color:${colorMap[qId]};font-weight:600;font-size:12px;">${labelMap[qId]}</span></td>
      </tr>`;
    }).join('');
  }
}

// ============================================================
// Init
// ============================================================
// ── Feedback ──
// GAS Feedback Proxy URL（デプロイ後にここに設定）
const GAS_FEEDBACK_URL = 'https://script.google.com/macros/s/AKfycbwdA4HjP05m1Y0IQ7lZEVMWfVuJw3kbDlTjUqwVsTBv5aZOa-ya2VLsXtizcd2XDLZ8_w/exec';

function openFeedback() {
  document.getElementById('feedback-modal').classList.add('show');
  document.getElementById('fb-result').innerHTML = '';
  document.getElementById('fb-send').disabled = false;
}
function closeFeedback() {
  document.getElementById('feedback-modal').classList.remove('show');
}
document.getElementById('feedback-modal').addEventListener('click', e => {
  if (e.target.id === 'feedback-modal') closeFeedback();
});

async function sendFeedback() {
  const name = document.getElementById('fb-name').value.trim() || '匿名';
  const type = document.getElementById('fb-type').value;
  const message = document.getElementById('fb-message').value.trim();
  if (!message) { alert('内容を入力してください'); return; }

  const btn = document.getElementById('fb-send');
  btn.disabled = true;
  btn.textContent = '送信中...';

  const fbTab = document.getElementById('fb-tab').value;
  const activeTab = document.querySelector('.tab-btn.active');
  const tabName = fbTab || (activeTab ? activeTab.textContent : '-');

  const payload = {
    text: `*[${type}]* from ${name}\nタブ: ${tabName}\n\n${message}`
  };

  try {
    if (!GAS_FEEDBACK_URL) { alert('フィードバックURLが未設定です'); btn.disabled = false; btn.textContent = '送信'; return; }
    const resp = await fetch(GAS_FEEDBACK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload),
    });
    const result = await resp.json();
    if (result.success) {
      document.getElementById('fb-result').innerHTML = '<div class="feedback-sent">送信しました ✓</div>';
      document.getElementById('fb-message').value = '';
      setTimeout(closeFeedback, 1500);
    } else {
      document.getElementById('fb-result').innerHTML = '<div style="color:#ff3b30;font-size:13px;text-align:center;">送信失敗: ' + (result.error || '') + '</div>';
    }
  } catch (e) {
    document.getElementById('fb-result').innerHTML = '<div style="color:#ff3b30;font-size:13px;text-align:center;">送信失敗</div>';
  }
  btn.disabled = false;
  btn.textContent = '送信';
}

// ============================================================
// Tab 6: レビュー（仮実装・モックデータ）
// ============================================================
const REVIEW_MOCK = (function generateMockReviews() {
  const propsBase = [
    { code: 'ENK801', name: 'ENK', room: '801', area: '大阪' },
    { code: 'ENK902', name: 'ENK', room: '902', area: '大阪' },
    { code: 'ENK302', name: 'ENK', room: '302', area: '大阪' },
    { code: 'ENK602', name: 'ENK', room: '602', area: '大阪' },
    { code: 'ENK502', name: 'ENK', room: '502', area: '大阪' },
    { code: 'ENK101', name: 'ENK', room: '101', area: '大阪' },
    { code: 'ENK201', name: 'ENK', room: '201', area: '大阪' },
    { code: 'WID101', name: 'WID', room: '101', area: '京都' },
    { code: 'WID201', name: 'WID', room: '201', area: '京都' },
    { code: 'HTL布団6', name: 'HTL', room: '布団6', area: '東京' },
    { code: 'HTL布団3', name: 'HTL', room: '布団3', area: '東京' },
    { code: 'YUAN401', name: 'Yuan', room: '401', area: '京都' },
  ];
  const guestsByLang = {
    ja: ['田中 健', '佐藤 美咲', '鈴木 大輔', '高橋 由紀', '伊藤 翔'],
    en: ['John Smith', 'Emma Wilson', 'Michael Brown', 'Sarah Davis', 'Victor Sobczak'],
    zh: ['王 偉', '李 娜', '張 偉', '劉 洋', '陳 静']
  };
  const positives = {
    ja: ['とても清潔で快適でした。立地も良く再訪したいです。', '部屋が広く、設備も充実していました。', 'スタッフの対応が丁寧で安心して滞在できました。'],
    en: ['Spotless apartment, perfect location! Will book again.', 'Comfortable stay, host was very responsive.', 'Excellent value for money, highly recommended.'],
    zh: ['房间很干净，位置也很好，下次还会再来。', '设施齐全，体验很好。', '主人很热情，推荐！']
  };
  const negatives = {
    ja: ['Wi-Fiが不安定で仕事に支障が出ました。改善希望です。', 'チェックインの説明が分かりにくかったです。'],
    en: ['Wi-Fi was unreliable, made remote work difficult.', 'The room was smaller than expected based on photos.'],
    zh: ['Wi-Fi信号不稳定。', '房间比照片看起来小。']
  };
  const privates = [
    'シャワーの水圧が弱い気がしました。',
    'エアコンのリモコンの電池が切れていました。',
    '冷蔵庫の中に前のゲストの食品が残っていました。',
    'ベッドのスプリングが少し気になりました。',
    'キッチンの包丁が切れにくかったです。',
    '玄関の鍵の操作がやや分かりにくかったです。'
  ];
  const reviews = [];
  const now = new Date();
  let id = 1;
  for (let i = 0; i < 180; i++) {
    const daysAgo = Math.floor(Math.random() * 360);
    const date = new Date(now.getTime() - daysAgo * 86400000);
    const prop = propsBase[Math.floor(Math.random() * propsBase.length)];
    const langs = ['ja', 'en', 'zh'];
    const lang = langs[Math.floor(Math.random() * langs.length)];
    const guests = guestsByLang[lang];
    const guest = guests[Math.floor(Math.random() * guests.length)];
    const r = Math.random();
    let star;
    if (r < 0.7) star = 5;
    else if (r < 0.88) star = 4;
    else if (r < 0.96) star = 3;
    else if (r < 0.99) star = 2;
    else star = 1;
    const sentiment = star >= 4 ? 'pos' : star === 3 ? 'neu' : 'neg';
    const body = sentiment === 'neg'
      ? negatives[lang][Math.floor(Math.random() * negatives[lang].length)]
      : positives[lang][Math.floor(Math.random() * positives[lang].length)];
    const hasPrivate = Math.random() < 0.18;
    const replyStatus = star <= 3
      ? (Math.random() < 0.4 ? 'pending' : (Math.random() < 0.5 ? 'approved' : 'posted'))
      : (Math.random() < 0.7 ? 'posted' : 'none');
    reviews.push({
      id: 'rv_' + (id++),
      date: date.toISOString().slice(0, 10),
      property: prop,
      guest, lang, star, body, sentiment,
      stars: {
        cleanliness: Math.max(1, star - (Math.random() < 0.3 ? 1 : 0)),
        accuracy: Math.max(1, star - (Math.random() < 0.2 ? 1 : 0)),
        checkin: Math.max(1, star - (Math.random() < 0.2 ? 1 : 0)),
        communication: Math.max(1, star),
        location: Math.max(1, star + (Math.random() < 0.3 ? 0 : 0)),
        value: Math.max(1, star - (Math.random() < 0.2 ? 1 : 0)),
      },
      privateFeedback: hasPrivate ? privates[Math.floor(Math.random() * privates.length)] : null,
      replyStatus,
      replyDraft: star <= 3 ? 'このたびはご不便をおかけし申し訳ございません。いただいたご指摘を真摯に受け止め、改善に努めてまいります。' : null
    });
  }
  // Mock execution log
  const logs = [];
  for (let i = 0; i < 20; i++) {
    const daysAgo = Math.floor(Math.random() * 14);
    const dt = new Date(now.getTime() - daysAgo * 86400000 - Math.random() * 86400000);
    const prop = propsBase[Math.floor(Math.random() * propsBase.length)];
    const langs = ['ja', 'en', 'zh'];
    const lang = langs[Math.floor(Math.random() * langs.length)];
    const guests = guestsByLang[lang];
    const r = Math.random();
    const status = r < 0.85 ? 'success' : (r < 0.93 ? 'skipped' : 'failed');
    logs.push({
      datetime: dt.toISOString().replace('T', ' ').slice(0, 16),
      system: Math.random() < 0.85 ? 'A_post' : 'B_reply',
      property: prop,
      guest: guests[Math.floor(Math.random() * guests.length)],
      lang, status,
      note: status === 'skipped' ? '催促OFF' : (status === 'failed' ? 'モーダル開けず' : '★5固定')
    });
  }
  logs.sort((a, b) => b.datetime.localeCompare(a.datetime));
  return { reviews, logs };
})();

let _reviewFilters = { period: 30, area: '全体', lang: 'all', star: 'all' };
let _rvCharts = {};

function _filterReviews() {
  const now = new Date();
  const cutoff = new Date(now.getTime() - _reviewFilters.period * 86400000);
  return REVIEW_MOCK.reviews.filter(r => {
    if (new Date(r.date) < cutoff) return false;
    if (_reviewFilters.area !== '全体' && r.property.area !== _reviewFilters.area) return false;
    if (_reviewFilters.lang !== 'all' && r.lang !== _reviewFilters.lang) return false;
    return true;
  });
}

function setReviewFilter(el, key) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  if (key === 'period') _reviewFilters.period = parseInt(el.dataset.period, 10);
  if (key === 'area') _reviewFilters.area = el.dataset.area;
  if (key === 'lang') _reviewFilters.lang = el.dataset.lang;
  if (key === 'star') _reviewFilters.star = el.dataset.star;
  renderReviewTab();
  initReviewCharts();
}

// ============================================================
// Tab: 要チェック（新着 + パフォーマンス警告）
// ============================================================
const WATCHLIST_NEW_MONTHS = 4;
const WATCHLIST_SALES_YELLOW = 70; // 達成率<70% → 黄
const WATCHLIST_SALES_RED = 50;    // 達成率<50% → 赤
const WATCHLIST_OCC_THRESHOLD = 60; // 直近30日稼働率<60% → 警告
const WATCHLIST_BOOKING_DAYS = 14;  // 直近14日

function isNewProperty(prop) {
  if (!prop || !prop.startDate) return false;
  const start = new Date(prop.startDate);
  if (isNaN(start)) return false;
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - WATCHLIST_NEW_MONTHS);
  return start >= cutoff;
}

function getMonthsSinceStart(prop) {
  if (!prop || !prop.startDate) return null;
  const start = new Date(prop.startDate);
  if (isNaN(start)) return null;
  const now = new Date();
  return (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth());
}

function getCurrentMonthSalesAchievement(prop) {
  const now = new Date();
  const ym = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  const daysInMonth = getDaysInMonth(ym);
  const dayOfMonth = now.getDate();
  const stats = computePropertyStats(prop.name, ym);
  if (!stats) return null;
  const monthTarget = getTargetForProperty(prop, now.getMonth() + 1);
  if (!monthTarget) return null;
  const proratedTarget = monthTarget * (dayOfMonth / daysInMonth);
  const pct = proratedTarget > 0 ? (stats.sales / proratedTarget) * 100 : 0;
  return { actual: stats.sales, target: monthTarget, proratedTarget, pct };
}

function getRecent30DayOcc(prop) {
  const today = new Date();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  const cutoffStr = localDateStr(cutoff);
  const todayStr = localDateStr(today);

  const dailyDates = new Set();
  rawDailyData.forEach(d => {
    const date = normalizeDate(d['日付']);
    const code = generatePropCode(d['物件名'] || '', d['ルーム番号'] || '');
    const status = d['状態'] || '';
    if (code !== prop.name) return;
    if (status === 'システムキャンセル') return;
    if (date >= cutoffStr && date < todayStr) dailyDates.add(date);
  });
  const totalAvail = 30 * (prop.rooms || 1);
  return totalAvail > 0 ? (dailyDates.size / totalAvail) * 100 : 0;
}

function getRecent14DayBookingCount(prop) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - WATCHLIST_BOOKING_DAYS);
  const cutoffStr = localDateStr(cutoff);
  let count = 0;
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    if (!r.date || r.date < cutoffStr) return;
    if (r.propCode !== prop.name && r.property !== prop.name && r.property !== prop.propName) return;
    count++;
  });
  return count;
}

function getWatchlistReasons(prop) {
  const reasons = [];
  const sales = getCurrentMonthSalesAchievement(prop);
  if (sales && sales.target > 0) {
    if (sales.pct < WATCHLIST_SALES_RED) {
      reasons.push({ type: 'sales', severity: 'red', label: `売上達成率 ${sales.pct.toFixed(0)}%`, detail: `${fmtYen(sales.actual)} / 想定 ${fmtYen(sales.proratedTarget)}` });
    } else if (sales.pct < WATCHLIST_SALES_YELLOW) {
      reasons.push({ type: 'sales', severity: 'yellow', label: `売上達成率 ${sales.pct.toFixed(0)}%`, detail: `${fmtYen(sales.actual)} / 想定 ${fmtYen(sales.proratedTarget)}` });
    }
  }
  const occ = getRecent30DayOcc(prop);
  if (occ < WATCHLIST_OCC_THRESHOLD) {
    reasons.push({ type: 'occ', severity: 'yellow', label: `稼働率 ${occ.toFixed(0)}%`, detail: '直近30日' });
  }
  const bookings = getRecent14DayBookingCount(prop);
  if (bookings === 0) {
    reasons.push({ type: 'booking', severity: 'red', label: '新規予約 0件', detail: '直近14日' });
  }
  return reasons;
}

function getWatchlistChartData(prop) {
  const labels = [];
  const sales = [];
  const targets = [];
  const occs = [];
  const now = new Date();
  for (let i = -3; i <= 3; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
    labels.push((d.getMonth() + 1) + '月');
    const s = computePropertyStats(prop.name, ym);
    sales.push(s ? Math.round(s.sales) : 0);
    targets.push(getTargetForProperty(prop, d.getMonth() + 1) || 0);
    occs.push(s ? Math.round(s.occ * 10) / 10 : 0);
  }
  return { labels, sales, targets, occs };
}

let _watchlistCounts = { newCount: 0, watchCount: 0 };

function renderWatchlistTab() {
  if (!document.getElementById('tab-watchlist')) return;
  _activeWatchlistDetail = null;

  // 新着物件
  const newProps = properties.filter(p => isNewProperty(p) && p.status === '稼働中' && !p.excludeKpi);
  // パフォーマンス警告（新着除外、稼働中のみ、KPI除外も除外）
  const watchProps = [];
  properties.forEach(p => {
    if (p.status !== '稼働中') return;
    if (p.excludeKpi) return;
    if (isNewProperty(p)) return;
    const reasons = getWatchlistReasons(p);
    if (reasons.length > 0) watchProps.push({ prop: p, reasons });
  });
  // 重要度順（赤の数 → 理由数）
  watchProps.sort((a, b) => {
    const ar = a.reasons.filter(r => r.severity === 'red').length;
    const br = b.reasons.filter(r => r.severity === 'red').length;
    if (br !== ar) return br - ar;
    return b.reasons.length - a.reasons.length;
  });

  _watchlistCounts = { newCount: newProps.length, watchCount: watchProps.length };

  // KPI
  const kpiNew = document.getElementById('kpi-watchlist-new');
  const kpiWatch = document.getElementById('kpi-watchlist-watch');
  if (kpiNew) kpiNew.textContent = newProps.length + '件';
  if (kpiWatch) kpiWatch.textContent = watchProps.length + '件';

  // Render new section
  const newSec = document.getElementById('watchlist-new-list');
  if (newSec) {
    if (newProps.length === 0) {
      newSec.innerHTML = '<div style="color:#999;font-size:13px;padding:12px;">新着物件はありません</div>';
    } else {
      const d7ago = new Date(); d7ago.setDate(d7ago.getDate() - 7);
      const d7agoStr = localDateStr(d7ago);
      const todayStr = localDateStr(new Date());
      newSec.innerHTML = newProps.map((p, i) => {
        const months = getMonthsSinceStart(p);
        let bkCount = 0;
        reservations.forEach(r => {
          if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
          if (r.propCode !== p.name && r.property !== p.name && r.property !== p.propName) return;
          if (r.date && r.date >= d7agoStr && r.date <= todayStr) bkCount++;
        });
        const bkPill = bkCount > 0
          ? `<span class="badge-green" style="margin-left:6px;">直近7日 ${bkCount}件予約</span>`
          : `<span class="badge-orange" style="margin-left:6px;">直近7日 0件</span>`;
        return `<div class="watchlist-card clickable" onclick="toggleWatchlistDetail('${p.name}', event)">
          <div class="watchlist-card-header">
            <div class="watchlist-card-title">${p.name}</div>
            <div><span class="badge-blue">🆕 新着 ${months !== null ? months + 'ヶ月目' : ''}</span>${bkPill}</div>
          </div>
          <div class="watchlist-card-meta">運用開始: ${p.startDate || '-'} / オーナー: ${p.ownerId || '-'}</div>
          <div style="height:120px;"><canvas id="watchlist-chart-new-${i}"></canvas></div>
        </div>`;
      }).join('');
    }
  }

  // Render watch section
  const watchSec = document.getElementById('watchlist-watch-list');
  if (watchSec) {
    if (watchProps.length === 0) {
      watchSec.innerHTML = '<div style="color:#999;font-size:13px;padding:12px;">要チェック物件はありません 🎉</div>';
    } else {
      watchSec.innerHTML = watchProps.map(({ prop, reasons }, i) => {
        const badges = reasons.map(r => {
          const cls = r.severity === 'red' ? 'badge-red' : 'badge-orange';
          return `<span class="${cls}" title="${r.detail}">${r.label}</span>`;
        }).join(' ');
        const sales = getCurrentMonthSalesAchievement(prop);
        let salesLine = '';
        if (sales && sales.target > 0) {
          const pctColor = sales.pct < WATCHLIST_SALES_RED ? '#ff3b30' : (sales.pct < WATCHLIST_SALES_YELLOW ? '#ff9500' : '#34c759');
          salesLine = `<div style="font-size:12px;margin:6px 0 4px;">当月売上達成率: <strong style="color:${pctColor};">${sales.pct.toFixed(0)}%</strong> <span style="color:#999;">(${fmtYen(sales.actual)} / 想定 ${fmtYen(sales.proratedTarget)})</span></div>`;
        }
        return `<div class="watchlist-card clickable" onclick="toggleWatchlistDetail('${prop.name}', event)">
          <div class="watchlist-card-header">
            <div class="watchlist-card-title">${prop.name}</div>
            <div>${badges}</div>
          </div>
          <div class="watchlist-card-meta">オーナー: ${prop.ownerId || '-'} / エリア: ${prop.area || '-'}</div>
          ${salesLine}
          <div style="height:120px;"><canvas id="watchlist-chart-watch-${i}"></canvas></div>
        </div>`;
      }).join('');
    }
  }

  // Update summary badge in daily tab
  const badge = document.getElementById('watchlist-summary-badge');
  if (badge) {
    if (newProps.length + watchProps.length === 0) {
      badge.style.display = 'none';
    } else {
      badge.style.display = '';
      badge.innerHTML = `⚠️ 要チェック: <strong>${watchProps.length}件</strong> &nbsp;|&nbsp; 🆕 新着: <strong>${newProps.length}件</strong> &nbsp; <a href="#" onclick="switchTab('watchlist');return false;" style="color:#007aff;text-decoration:none;font-weight:600;">→ タブを開く</a>`;
    }
  }
}

let _activeWatchlistDetail = null;

function toggleWatchlistDetail(propertyName, event) {
  // 詳細パネル内のクリックは無視
  if (event.target.closest('#wl-detail-panel')) return;

  // 既存の詳細パネルを削除
  const existing = document.getElementById('wl-detail-panel');
  if (existing) {
    destroyDrillCharts('wl');
    existing.remove();
  }

  // 同じ物件をクリック → 閉じる
  if (_activeWatchlistDetail === propertyName) {
    _activeWatchlistDetail = null;
    return;
  }
  _activeWatchlistDetail = propertyName;

  // クリックされたカードのすぐ下に詳細パネルを挿入
  const card = event.target.closest('.watchlist-card');
  if (!card) return;
  const panel = document.createElement('div');
  panel.id = 'wl-detail-panel';
  panel.className = 'card';
  panel.style.cssText = 'margin-top:8px; grid-column: 1 / -1;';
  card.insertAdjacentElement('afterend', panel);

  // renderPropertyDetail が期間ピルで #wlDetailContainer を参照するのでラッパーを用意
  const inner = document.createElement('div');
  inner.id = 'wlDetailContainer';
  panel.appendChild(inner);

  currentFilters.propDetailPeriod = 'thisMonth';
  renderPropertyDetail(inner, propertyName, 'wl');
  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function initWatchlistCharts() {
  // Destroy existing
  Object.keys(window).forEach(k => {
    if (k.startsWith('_watchlistChart_') && window[k]) {
      try { window[k].destroy(); } catch (e) {}
      window[k] = null;
    }
  });

  const drawMini = (canvasId, prop) => {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const { labels, sales, targets, occs } = getWatchlistChartData(prop);
    const ctx = canvas.getContext('2d');
    const chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: '売上', data: sales, backgroundColor: CHART_COLORS.blue, borderRadius: 4, order: 3, yAxisID: 'y' },
          { label: '目標', data: targets, type: 'line', borderColor: CHART_COLORS.red, borderDash: [4, 4], borderWidth: 2, pointRadius: 0, fill: false, order: 2, yAxisID: 'y' },
          { label: 'OCC', data: occs, type: 'line', borderColor: CHART_COLORS.green, borderWidth: 2, pointRadius: 2, fill: false, order: 1, yAxisID: 'y1' },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (c) => c.dataset.label === 'OCC' ? `OCC: ${c.parsed.y}%` : `${c.dataset.label}: ${fmtYen(c.parsed.y)}`,
            },
          },
        },
        scales: {
          y: { beginAtZero: true, ticks: { callback: (v) => fmtYen(v), font: { size: 10 } } },
          y1: { beginAtZero: true, max: 100, position: 'right', grid: { drawOnChartArea: false }, ticks: { callback: (v) => v + '%', font: { size: 10 } } },
          x: { ticks: { font: { size: 10 } } },
        },
      },
    });
    window['_watchlistChart_' + canvasId] = chart;
  };

  const newProps = properties.filter(p => isNewProperty(p) && p.status === '稼働中' && !p.excludeKpi);
  newProps.forEach((p, i) => drawMini('watchlist-chart-new-' + i, p));

  const watchProps = [];
  properties.forEach(p => {
    if (p.status !== '稼働中') return;
    if (p.excludeKpi) return;
    if (isNewProperty(p)) return;
    const reasons = getWatchlistReasons(p);
    if (reasons.length > 0) watchProps.push({ prop: p, reasons });
  });
  watchProps.sort((a, b) => {
    const ar = a.reasons.filter(r => r.severity === 'red').length;
    const br = b.reasons.filter(r => r.severity === 'red').length;
    if (br !== ar) return br - ar;
    return b.reasons.length - a.reasons.length;
  });
  watchProps.forEach(({ prop }, i) => drawMini('watchlist-chart-watch-' + i, prop));
}

// ============================================================
// Tab: PM/BM 分析
// ============================================================
function computePmBmDetail(months, area) {
  const monthSet = new Set(months);
  const byOwner = new Map();
  const byProp = new Map();
  const byArea = new Map();
  let pmSales = 0, bmSales = 0, totalSales = 0;
  let cleaningCount = 0;

  const ensureOwner = (prop) => {
    const oid = prop.ownerId || '_none';
    if (!byOwner.has(oid)) byOwner.set(oid, { id: oid, name: prop.ownerName || oid, pm: 0, bm: 0 });
    return byOwner.get(oid);
  };
  const ensureProp = (prop) => {
    if (!byProp.has(prop.name)) byProp.set(prop.name, {
      name: prop.name, ownerName: prop.ownerName || '', area: prop.area || '',
      royaltyPct: prop.royaltyPct || 0, pm: 0, bm: 0,
    });
    return byProp.get(prop.name);
  };
  const ensureArea = (a) => {
    if (!byArea.has(a)) byArea.set(a, { area: a, pm: 0, bm: 0 });
    return byArea.get(a);
  };

  // PM (and total sales): checkin-month basis
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    if (!monthSet.has(getYearMonth(r.checkin))) return;
    const prop = findPropByReservation(r);
    if (!prop) return;
    if (area !== '全体' && prop.area !== area) return;

    const sale = r.sales || 0;
    totalSales += sale;
    const royaltyPct = prop.royaltyPct || 0;
    let pm = 0;
    if (royaltyPct > 0) {
      pm = ((sale - (r.otaFee || 0) - (r.cleaningFee || 0)) * royaltyPct) / 100;
    }
    pmSales += pm;

    ensureOwner(prop).pm += pm;
    ensureProp(prop).pm += pm;
    ensureArea(prop.area || '不明').pm += pm;
  });

  // BM: checkout-month basis
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    if (!monthSet.has(getYearMonth(r.checkout))) return;
    const prop = findPropByReservation(r);
    if (!prop) return;
    if (area !== '全体' && prop.area !== area) return;

    const bm = r.cleaningFee || 0;
    if (bm <= 0) return;
    bmSales += bm;
    cleaningCount++;

    ensureOwner(prop).bm += bm;
    ensureProp(prop).bm += bm;
    ensureArea(prop.area || '不明').bm += bm;
  });

  return {
    pmSales, bmSales, totalSales,
    pmRate: totalSales > 0 ? (pmSales / totalSales) * 100 : 0,
    cleaningCount,
    avgCleaningUnit: cleaningCount > 0 ? bmSales / cleaningCount : 0,
    byOwner: Array.from(byOwner.values()),
    byProp: Array.from(byProp.values()),
    byArea: Array.from(byArea.values()),
  };
}

function renderPmbmTab() {
  if (!document.getElementById('tab-pmbm')) return;
  const months = getSelectedMonths('pmbm');
  const area = currentFilters.pmbmArea;

  const cur = computePmBmDetail(months, area);
  const py = computePmBmDetail(shiftMonths(months, -12), area);
  const pm = computePmBmDetail(shiftMonths(months, -1), area);

  // 物件数 (稼働中) — area で絞った稼働中物件数で平均単価を算出
  const propCount = filterPropertiesByArea(area).filter(p => p.status === '稼働中').length;
  const avgPmPerProp = propCount > 0 ? cur.pmSales / propCount : 0;
  const avgBmPerProp = propCount > 0 ? cur.bmSales / propCount : 0;
  const pyAvgPmPerProp = propCount > 0 ? py.pmSales / propCount : 0;
  const pyAvgBmPerProp = propCount > 0 ? py.bmSales / propCount : 0;
  const pmAvgPmPerProp = propCount > 0 ? pm.pmSales / propCount : 0;
  const pmAvgBmPerProp = propCount > 0 ? pm.bmSales / propCount : 0;

  // KPI cards
  const setText = (id, v) => { const el = document.getElementById(id); if (el) el.innerHTML = v; };
  setText('kpi-pmbm-pm', fmtYen(cur.pmSales));
  setText('kpi-pmbm-pm-vs', fmtVsLine(cur.pmSales, py.pmSales, pm.pmSales));
  setText('kpi-pmbm-bm', fmtYen(cur.bmSales));
  setText('kpi-pmbm-bm-vs', fmtVsLine(cur.bmSales, py.bmSales, pm.bmSales));
  setText('kpi-pmbm-total', fmtYen(cur.pmSales + cur.bmSales));
  setText('kpi-pmbm-total-vs', fmtVsLine(cur.pmSales + cur.bmSales, py.pmSales + py.bmSales, pm.pmSales + pm.bmSales));
  setText('kpi-pmbm-rate', cur.pmRate.toFixed(1) + '%');
  setText('kpi-pmbm-rate-vs', fmtVsLinePt(cur.pmRate, py.pmRate, pm.pmRate));
  setText('kpi-pmbm-pm-avg', fmtYen(avgPmPerProp));
  setText('kpi-pmbm-pm-avg-vs', fmtVsLine(avgPmPerProp, pyAvgPmPerProp, pmAvgPmPerProp));
  setText('kpi-pmbm-bm-avg', fmtYen(avgBmPerProp));
  setText('kpi-pmbm-bm-avg-vs', fmtVsLine(avgBmPerProp, pyAvgBmPerProp, pmAvgBmPerProp));

  // ロイヤリティ計算不能オーナーの警告
  // 物件マスタに紐づいているオーナーのうち、royaltyParseFailed のものを抽出
  const usedOwnerIds = new Set(properties.map(p => p.ownerId).filter(Boolean));
  const failedOwners = owners.filter(o => o.royaltyParseFailed && usedOwnerIds.has(o.id));
  const warnEl = document.getElementById('pmbm-royalty-warning');
  if (warnEl) {
    if (failedOwners.length === 0) {
      warnEl.style.display = 'none';
      warnEl.innerHTML = '';
    } else {
      const items = failedOwners.map(o => {
        const preview = (o.royalty || '').replace(/\s+/g, ' ').slice(0, 60);
        return `<li><strong>${o.id}</strong> — ロイヤリティ: 「${preview}${o.royalty.length > 60 ? '…' : ''}」</li>`;
      }).join('');
      warnEl.style.display = 'block';
      warnEl.innerHTML =
        `<div style="font-weight:600;margin-bottom:6px;">⚠️ ロイヤリティが自動計算できないオーナーが ${failedOwners.length} 件あります（PM売上に未集計）</div>` +
        `<ul style="margin:4px 0 8px 20px;padding:0;">${items}</ul>` +
        `<div style="font-size:12px;">対応: オーナーマスタの <strong>計算用ロイヤリティ</strong> 列に固定% (例: <code>18</code>) を入力してください。</div>`;
    }
  }

  // Owner combined top 10
  const ownerCombined = [...cur.byOwner].map(o => ({ ...o, total: o.pm + o.bm }))
    .sort((a, b) => b.total - a.total).slice(0, 10);
  const tbodyCombined = document.getElementById('pmbm-owner-combined-tbody');
  if (tbodyCombined) {
    tbodyCombined.innerHTML = ownerCombined.map((o, i) =>
      `<tr><td>${i + 1}</td><td>${o.name}</td><td class="text-right">${fmtYen(o.pm)}</td><td class="text-right">${fmtYen(o.bm)}</td><td class="text-right"><strong>${fmtYen(o.total)}</strong></td></tr>`
    ).join('') || '<tr><td colspan="5" style="color:#999;text-align:center;">データなし</td></tr>';
  }

  // Owner PM top 10
  const ownerPm = [...cur.byOwner].sort((a, b) => b.pm - a.pm).slice(0, 10);
  const tbodyOwnerPm = document.getElementById('pmbm-owner-pm-tbody');
  if (tbodyOwnerPm) {
    tbodyOwnerPm.innerHTML = ownerPm.map((o, i) =>
      `<tr><td>${i + 1}</td><td>${o.name}</td><td class="text-right">${fmtYen(o.pm)}</td></tr>`
    ).join('') || '<tr><td colspan="3" style="color:#999;text-align:center;">データなし</td></tr>';
  }

  // Owner BM top 10
  const ownerBm = [...cur.byOwner].sort((a, b) => b.bm - a.bm).slice(0, 10);
  const tbodyOwnerBm = document.getElementById('pmbm-owner-bm-tbody');
  if (tbodyOwnerBm) {
    tbodyOwnerBm.innerHTML = ownerBm.map((o, i) =>
      `<tr><td>${i + 1}</td><td>${o.name}</td><td class="text-right">${fmtYen(o.bm)}</td></tr>`
    ).join('') || '<tr><td colspan="3" style="color:#999;text-align:center;">データなし</td></tr>';
  }

  // Property PM top 10
  const propPm = [...cur.byProp].sort((a, b) => b.pm - a.pm).slice(0, 10);
  const tbodyPropPm = document.getElementById('pmbm-prop-pm-tbody');
  if (tbodyPropPm) {
    tbodyPropPm.innerHTML = propPm.map((p, i) =>
      `<tr><td>${i + 1}</td><td>${p.name}</td><td>${p.ownerName}</td><td>${p.area}</td><td class="text-right">${p.royaltyPct}%</td><td class="text-right">${fmtYen(p.pm)}</td></tr>`
    ).join('') || '<tr><td colspan="6" style="color:#999;text-align:center;">データなし</td></tr>';
  }

  // YoY PM up/down top 5
  const pyPropMap = new Map(py.byProp.map(p => [p.name, p.pm]));
  const propDeltas = cur.byProp
    .map(p => ({ name: p.name, cur: p.pm, prev: pyPropMap.get(p.name) || 0 }))
    .filter(p => p.cur > 0 || p.prev > 0)
    .map(p => ({ ...p, diff: p.cur - p.prev }));
  const ups = [...propDeltas].sort((a, b) => b.diff - a.diff).slice(0, 5);
  const downs = [...propDeltas].sort((a, b) => a.diff - b.diff).slice(0, 5);
  const fillDelta = (id, rows) => {
    const tbody = document.getElementById(id);
    if (!tbody) return;
    tbody.innerHTML = rows.map(r => {
      const cls = r.diff >= 0 ? 'positive' : 'negative';
      const sign = r.diff >= 0 ? '+' : '';
      return `<tr><td>${r.name}</td><td class="text-right">${fmtYen(r.prev)}</td><td class="text-right">${fmtYen(r.cur)}</td><td class="text-right ${cls}">${sign}${fmtYen(r.diff)}</td></tr>`;
    }).join('') || '<tr><td colspan="4" style="color:#999;text-align:center;">データなし</td></tr>';
  };
  fillDelta('pmbm-prop-up-tbody', ups);
  fillDelta('pmbm-prop-down-tbody', downs);

  // Royalty bucket distribution
  const buckets = [
    { label: '0% (運営費のみ等)', min: 0, max: 0 },
    { label: '1〜10%', min: 1, max: 10 },
    { label: '11〜15%', min: 11, max: 15 },
    { label: '16〜20%', min: 16, max: 20 },
    { label: '21〜25%', min: 21, max: 25 },
    { label: '26%以上', min: 26, max: 999 },
  ];
  const propPmMap = new Map(cur.byProp.map(p => [p.name, p.pm]));
  const filteredProps = filterPropertiesByArea(area).filter(p => p.status === '稼働中');
  const bucketAgg = buckets.map(b => ({ ...b, count: 0, pm: 0 }));
  filteredProps.forEach(p => {
    const r = p.royaltyPct || 0;
    const b = bucketAgg.find(x => r >= x.min && r <= x.max);
    if (!b) return;
    b.count++;
    b.pm += propPmMap.get(p.name) || 0;
  });
  const tbodyRoy = document.getElementById('pmbm-royalty-tbody');
  if (tbodyRoy) {
    tbodyRoy.innerHTML = bucketAgg.map(b =>
      `<tr><td>${b.label}</td><td class="text-right">${b.count}件</td><td class="text-right">${fmtYen(b.pm)}</td></tr>`
    ).join('');
  }

  // BM breakdown
  setText('kpi-pmbm-clean-count', cur.cleaningCount + '件');
  setText('kpi-pmbm-clean-avg', fmtYen(cur.avgCleaningUnit));
  setText('kpi-pmbm-clean-total', fmtYen(cur.bmSales));
}

let _pmbmCharts = { monthly: null, rate: null, area: null };
// ============================================================
// マーケットタブ（AirDNA市場データの俯瞰・都市比較・インサイト）
// ============================================================
const MKT_CITIES = ['大阪', '京都', '東京'];
const MKT_CITY_COLORS = { '大阪': '#007AFF', '京都': '#FF9500', '東京': '#34C759' };

function setMarketPeriod(el) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.marketPeriod = el.dataset.period;
  renderMarketTab();
  setTimeout(() => initMarketCharts(), 50);
}

function setMarketCity(el) {
  el.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.marketCity = el.dataset.city;
  renderMarketTab();
  setTimeout(() => initMarketCharts(), 50);
}

function mktSheet(name) { return (window._airdnaSheets || {})[name]; }
function mktFirstValidField(sheet, fields) {
  if (!sheet || !sheet.length) return null;
  const keys = Object.keys(sheet[0]);
  for (const f of fields) if (keys.includes(f)) return f;
  return null;
}
function mktLatestMonth(city) {
  const s = mktSheet(`AD_${city}全域_occupancy`);
  if (!s) return null;
  const yms = s.map(r => (r['Date'] || '').slice(0, 7)).filter(Boolean).sort();
  return yms[yms.length - 1] || null;
}
function mktMonthsEndingAt(ym, count) {
  if (!ym) return [];
  const [y, m] = ym.split('-').map(Number);
  const out = [];
  for (let i = count - 1; i >= 0; i--) {
    const d = new Date(y, m - 1 - i, 1);
    out.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  return out;
}
function mktValue(sheet, field, ym) {
  if (!sheet || !field) return null;
  const row = sheet.find(r => (r['Date'] || '').slice(0, 7) === ym);
  if (!row) return null;
  const v = parseFloat(row[field]);
  return isNaN(v) ? null : v;
}
function mktAvg(sheet, field, yms) {
  if (!sheet || !field) return null;
  const set = new Set(yms);
  const vals = sheet.filter(r => set.has((r['Date'] || '').slice(0, 7)))
    .map(r => parseFloat(r[field]))
    .filter(v => !isNaN(v) && v > 0);
  return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
}
function mktCityKpi(city, yms) {
  const occS = mktSheet(`AD_${city}全域_occupancy`);
  const adrS = mktSheet(`AD_${city}全域_rates_summary`);
  const revS = mktSheet(`AD_${city}全域_revenue_summary`);
  const occF = mktFirstValidField(occS, ['Rate', 'Occupancy', 'rate']);
  const adrF = mktFirstValidField(adrS, ['Average daily rate', 'Daily rate', 'Rate', 'daily_rate']);
  const revF = mktFirstValidField(revS, ['Average annual revenue', 'Revenue', 'revenue']);
  const occ = mktAvg(occS, occF, yms);
  const adr = mktAvg(adrS, adrF, yms);
  const revpar = (occ !== null && adr !== null) ? (occ / 100) * adr : null;
  const revenue = mktAvg(revS, revF, yms);
  return { occ, adr, revpar, revenue };
}

function mktEnumerateWards(city) {
  const adSheets = window._airdnaSheets || {};
  const prefix = `AD_${city}_`;
  const wards = new Set();
  Object.keys(adSheets).forEach(name => {
    if (!name.startsWith(prefix)) return;
    if (!name.endsWith('_occupancy')) return;
    const rest = name.slice(prefix.length, -'_occupancy'.length);
    // ward のみ（bedSuffix含まない）= パーツが1つ
    if (!rest.includes('_')) wards.add(rest);
  });
  // エリア全域は除外
  wards.delete('全域');
  return [...wards];
}

// WARD_JP_TO_EN の逆引き（あいまい含む）
function mktWardEnToJp(en, city) {
  for (const [jp, m] of Object.entries(WARD_AMBIGUOUS)) {
    if (m[city] === en) return jp;
  }
  for (const [jp, e] of Object.entries(WARD_JP_TO_EN)) {
    if (e === en) return jp;
  }
  return en;
}

// ============================================================
// e-Stat データアクセスヘルパ
// ============================================================
function estatSheet(name) { return (window._estatSheets || {})[name]; }
function estatFindCol(row, candidates) {
  const keys = Object.keys(row || {});
  for (const c of candidates) {
    const match = keys.find(k => k && k.indexOf(c) >= 0);
    if (match) return match;
  }
  return null;
}
function estatExtractYm(s) {
  if (!s) return '';
  const m = String(s).match(/(\d{4})[^\d]?(\d{1,2})/);
  if (!m) return '';
  const y = m[1], mm = String(parseInt(m[2], 10)).padStart(2, '0');
  return `${y}-${mm}`;
}
function estatParseVal(v) {
  if (v === null || v === undefined) return NaN;
  const n = parseFloat(String(v).replace(/[,\s]/g, ''));
  return n;
}
function estatRecentYms(count, endYm) {
  // endYm='2026-02' なら直近count月のYM配列（古い→新しい）
  if (!endYm) return [];
  const [y, m] = endYm.split('-').map(Number);
  const out = [];
  for (let i = count - 1; i >= 0; i--) {
    const d = new Date(y, m - 1 - i, 1);
    out.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  return out;
}
function estatLatestYm(sheetName) {
  const sheet = estatSheet(sheetName);
  if (!sheet || !sheet.length) return null;
  const timeCol = estatFindCol(sheet[0], ['時間軸', '時点', '月']);
  if (!timeCol) return null;
  const yms = sheet.map(r => estatExtractYm(r[timeCol])).filter(Boolean);
  return yms.sort().pop() || null;
}

// 国籍英→地域マッピング（主要国）
const COUNTRY_REGION = {
  '韓国': '東アジア', '中国': '東アジア', '台湾': '東アジア', '香港': '東アジア',
  'タイ': '東南アジア', 'シンガポール': '東南アジア', 'マレーシア': '東南アジア', 'インドネシア': '東南アジア',
  'フィリピン': '東南アジア', 'ベトナム': '東南アジア',
  'インド': '南アジア',
  '米国': '欧米豪', 'アメリカ': '欧米豪', 'カナダ': '欧米豪', '英国': '欧米豪', 'イギリス': '欧米豪',
  'フランス': '欧米豪', 'ドイツ': '欧米豪', 'イタリア': '欧米豪', 'スペイン': '欧米豪',
  'オーストラリア': '欧米豪',
};
function classifyRegion(country) {
  if (!country) return 'その他';
  for (const key of Object.keys(COUNTRY_REGION)) {
    if (country.indexOf(key) >= 0) return COUNTRY_REGION[key];
  }
  return 'その他';
}

function renderMarketTab() {
  const sub = currentFilters.marketSubTab || 'top';
  if (sub === 'top') return renderMarketTopTab();
  if (sub === 'macro') return renderMarketMacroTab();
  return renderMarketAirdnaTab();
}

function renderMarketAirdnaTab() {
  const basis = document.getElementById('mkt-data-basis');
  const kpiEl = document.getElementById('mkt-kpi-grid');
  const insightsEl = document.getElementById('mkt-insights');
  const wardEl = document.getElementById('mkt-ward-rank');
  if (!kpiEl) return;

  // AirDNAデータが無い場合
  const anyCityData = Object.keys(window._airdnaSheets || {}).some(k => /^AD_(大阪|京都|東京)全域_/.test(k));
  if (!anyCityData) {
    kpiEl.innerHTML = `<div style="background:#FF950012;border-left:3px solid #FF9500;border-radius:6px;padding:12px 16px;font-size:13px;">
      AirDNA市場データが未取得です。Chrome拡張の「🏙 3都市一括取得」で取得してください。
    </div>`;
    if (basis) basis.textContent = '';
    if (insightsEl) insightsEl.innerHTML = '';
    if (wardEl) wardEl.innerHTML = '';
    return;
  }

  // 期間決定（各都市の最新月の最も古いもの基準に揃える）
  const latestPerCity = MKT_CITIES.map(c => mktLatestMonth(c)).filter(Boolean);
  const baseLatest = latestPerCity.sort()[0]; // 3都市共通の最新月
  const period = currentFilters.marketPeriod || 'last3';
  const count = period === 'latest' ? 1 : period === 'last3' ? 3 : period === 'last6' ? 6 : 12;
  const selYms = mktMonthsEndingAt(baseLatest, count);
  const prevYms = selYms.map(ym => { const [y, m] = ym.split('-'); return `${+y - 1}-${m}`; });

  if (basis) basis.textContent = `基準月: ${baseLatest || '—'} / 集計: ${selYms[0]}〜${selYms[selYms.length - 1]}`;

  // KPIカード（3都市 × OCC/ADR/RevPAR）
  const cityKpis = {};
  MKT_CITIES.forEach(c => { cityKpis[c] = { cur: mktCityKpi(c, selYms), prev: mktCityKpi(c, prevYms) }; });

  const fmtDelta = (cur, prev, isPct) => {
    if (cur === null || prev === null || prev === 0) return '';
    const d = cur - prev;
    const pct = (d / prev) * 100;
    const sign = d >= 0 ? '+' : '';
    const col = d >= 0 ? '#34C759' : '#FF3B30';
    const valStr = isPct ? `${sign}${d.toFixed(1)}pt` : `${sign}${pct.toFixed(1)}%`;
    return `<span style="color:${col};font-weight:600;">${valStr}</span>`;
  };

  kpiEl.innerHTML = `<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;">
    ${MKT_CITIES.map(c => {
      const { cur, prev } = cityKpis[c];
      const color = MKT_CITY_COLORS[c];
      return `<div class="card" style="border-top:3px solid ${color};">
        <div style="font-size:16px;font-weight:700;margin-bottom:12px;color:${color};">${c}</div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;">
          <div>
            <div style="font-size:11px;color:#86868b;">OCC</div>
            <div style="font-size:20px;font-weight:700;">${cur.occ !== null ? cur.occ.toFixed(1) + '%' : '—'}</div>
            <div style="font-size:11px;">YoY ${fmtDelta(cur.occ, prev.occ, true)}</div>
          </div>
          <div>
            <div style="font-size:11px;color:#86868b;">ADR</div>
            <div style="font-size:20px;font-weight:700;">${cur.adr !== null ? '¥' + Math.round(cur.adr).toLocaleString() : '—'}</div>
            <div style="font-size:11px;">YoY ${fmtDelta(cur.adr, prev.adr, false)}</div>
          </div>
          <div>
            <div style="font-size:11px;color:#86868b;">RevPAR</div>
            <div style="font-size:20px;font-weight:700;">${cur.revpar !== null ? '¥' + Math.round(cur.revpar).toLocaleString() : '—'}</div>
            <div style="font-size:11px;">YoY ${fmtDelta(cur.revpar, prev.revpar, false)}</div>
          </div>
        </div>
      </div>`;
    }).join('')}
  </div>`;

  // 区別ランキング
  const city = currentFilters.marketCity || '大阪';
  document.getElementById('mktWardRankTitle').textContent = `${city} 区別ランキング（RevPAR）`;
  const wards = mktEnumerateWards(city);
  const wardStats = wards.map(w => {
    const occS = mktSheet(`AD_${city}_${w}_occupancy`);
    const adrS = mktSheet(`AD_${city}_${w}_rates_summary`);
    const occF = mktFirstValidField(occS, ['Rate', 'Occupancy', 'rate']);
    const adrF = mktFirstValidField(adrS, ['Average daily rate', 'Daily rate', 'Rate', 'daily_rate']);
    const occ = mktAvg(occS, occF, selYms);
    const adr = mktAvg(adrS, adrF, selYms);
    const revpar = (occ !== null && adr !== null) ? (occ / 100) * adr : null;
    return { en: w, jp: mktWardEnToJp(w, city), occ, adr, revpar };
  }).filter(x => x.revpar !== null).sort((a, b) => b.revpar - a.revpar);

  if (wardEl) {
    if (wardStats.length === 0) {
      wardEl.innerHTML = `<div style="font-size:12px;color:#86868b;">${city}の区別データが未取得です</div>`;
    } else {
      const top5 = wardStats.slice(0, 5);
      const bot5 = wardStats.slice(-5).reverse();
      const row = (r, rank) => `<tr style="border-bottom:1px solid #f0f0f0;">
        <td style="padding:4px 6px;color:#86868b;width:28px;">${rank}</td>
        <td style="padding:4px 6px;font-weight:600;">${r.jp}</td>
        <td style="padding:4px 6px;text-align:right;">${r.occ !== null ? r.occ.toFixed(1) + '%' : '—'}</td>
        <td style="padding:4px 6px;text-align:right;">${r.adr !== null ? '¥' + Math.round(r.adr).toLocaleString() : '—'}</td>
        <td style="padding:4px 6px;text-align:right;font-weight:600;">¥${Math.round(r.revpar).toLocaleString()}</td>
      </tr>`;
      wardEl.innerHTML = `<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
        <div>
          <div style="font-size:11px;font-weight:600;color:#34C759;margin-bottom:4px;">▲ 上位5区</div>
          <table style="width:100%;border-collapse:collapse;font-size:12px;">
            <thead><tr style="color:#86868b;font-size:10px;"><th></th><th style="text-align:left;padding:2px 6px;">区</th><th style="text-align:right;padding:2px 6px;">OCC</th><th style="text-align:right;padding:2px 6px;">ADR</th><th style="text-align:right;padding:2px 6px;">RevPAR</th></tr></thead>
            <tbody>${top5.map((r, i) => row(r, i + 1)).join('')}</tbody>
          </table>
        </div>
        <div>
          <div style="font-size:11px;font-weight:600;color:#FF3B30;margin-bottom:4px;">▼ 下位5区</div>
          <table style="width:100%;border-collapse:collapse;font-size:12px;">
            <thead><tr style="color:#86868b;font-size:10px;"><th></th><th style="text-align:left;padding:2px 6px;">区</th><th style="text-align:right;padding:2px 6px;">OCC</th><th style="text-align:right;padding:2px 6px;">ADR</th><th style="text-align:right;padding:2px 6px;">RevPAR</th></tr></thead>
            <tbody>${bot5.map((r, i) => row(r, wardStats.length - i)).join('')}</tbody>
          </table>
        </div>
      </div>
      <div style="font-size:11px;color:#86868b;margin-top:8px;">※ 全${wardStats.length}区を集計</div>`;
    }
  }

  // 間取り別比較タイトル
  const bedTitle = document.getElementById('mktBedCompareTitle');
  if (bedTitle) bedTitle.textContent = `${city} 間取り別比較（OCC / ADR）`;

  // インサイト
  if (insightsEl) {
    const insights = buildMarketInsights(cityKpis, city, wardStats, selYms);
    insightsEl.innerHTML = insights;
  }
}

// ============================================================
// 観光統計（e-Stat）サブタブ
// ============================================================
function renderMarketMacroTab() {
  const statusEl = document.getElementById('mkt-macro-data-status');
  const hasData = Object.keys(window._estatSheets || {}).length > 0;
  if (!hasData) {
    if (statusEl) statusEl.innerHTML = `<div style="background:#FF950012;border-left:3px solid #FF9500;border-radius:6px;padding:12px 16px;font-size:13px;">観光統計データ未読込です。GASで <code>runAll()</code> を実行してください。</div>`;
    return;
  }
  if (statusEl) {
    const sheetCount = Object.keys(window._estatSheets).length;
    statusEl.innerHTML = `<div style="font-size:11px;color:#86868b;">📊 読込済み ${sheetCount} シート / 最終同期: e-Stat 公的統計</div>`;
  }
}

// 訪日月次集計（JNTO_訪日外客数シートから国籍別）
function estatBuildVisitorsByMonth() {
  const sheet = estatSheet('JNTO_訪日外客数');
  if (!sheet || !sheet.length) return { yms: [], total: {}, byCountry: {} };
  const countryCol = estatFindCol(sheet[0], ['国籍', '国']);
  const timeCol = estatFindCol(sheet[0], ['時間軸', '時点', '月']);
  const valCol = estatFindCol(sheet[0], ['値']);
  if (!countryCol || !timeCol || !valCol) return { yms: [], total: {}, byCountry: {} };

  const total = {}; // ym -> sum
  const byCountry = {}; // country -> { ym -> val }
  const ymSet = new Set();
  sheet.forEach(r => {
    const country = String(r[countryCol] || '').trim();
    const ym = estatExtractYm(r[timeCol]);
    const v = estatParseVal(r[valCol]);
    if (!country || !ym || isNaN(v)) return;
    // 「総数」系および地域集計は除外して国別のみ集計
    if (country === '総数' || country === '総計' || country === '計' || country.indexOf('総数') >= 0) return;
    const REGION_AGG = [
      'アジア', '東アジア', '東南アジア', '南アジア',
      '北アメリカ', '北米', '南アメリカ', '中南米',
      'ヨーロッパ', '欧州', '西欧', '東欧', '北欧', '南欧',
      'オセアニア', 'アフリカ', '中東',
      'その他', 'その他の地域', '無国籍・その他'
    ];
    if (REGION_AGG.indexOf(country) >= 0) return;
    ymSet.add(ym);
    if (!byCountry[country]) byCountry[country] = {};
    byCountry[country][ym] = (byCountry[country][ym] || 0) + v;
    total[ym] = (total[ym] || 0) + v;
  });
  return { yms: [...ymSet].sort(), total, byCountry };
}

// 都道府県別OCC（宿泊統計_定員稼働率）
function estatBuildPrefOcc() {
  const sheet = estatSheet('宿泊統計_定員稼働率');
  if (!sheet || !sheet.length) return { yms: [], byPref: {} };
  const prefCol = estatFindCol(sheet[0], ['所在地', '都道府県']);
  const typeCol = estatFindCol(sheet[0], ['宿泊施設タイプ', '施設タイプ']);
  const timeCol = estatFindCol(sheet[0], ['時間軸', '時点', '月']);
  const valCol = estatFindCol(sheet[0], ['値']);
  if (!prefCol || !timeCol || !valCol) return { yms: [], byPref: {} };

  // 「全施設タイプ」or 最初のタイプのみ集計（平均近似）
  const byPref = {}; // pref -> {ym -> {sum, count}}
  const ymSet = new Set();
  sheet.forEach(r => {
    const pref = String(r[prefCol] || '').trim();
    const ym = estatExtractYm(r[timeCol]);
    const v = estatParseVal(r[valCol]);
    if (!pref || !ym || isNaN(v) || v <= 0) return;
    ymSet.add(ym);
    if (!byPref[pref]) byPref[pref] = {};
    if (!byPref[pref][ym]) byPref[pref][ym] = { sum: 0, count: 0 };
    byPref[pref][ym].sum += v;
    byPref[pref][ym].count++;
  });
  // 平均化
  const byPrefAvg = {};
  Object.keys(byPref).forEach(p => {
    byPrefAvg[p] = {};
    Object.keys(byPref[p]).forEach(ym => {
      const { sum, count } = byPref[p][ym];
      byPrefAvg[p][ym] = count > 0 ? sum / count : 0;
    });
  });
  return { yms: [...ymSet].sort(), byPref: byPrefAvg };
}

// 日本人国内旅行者数（国内旅行_消費動向）
function estatBuildDomesticTravel() {
  const sheet = estatSheet('国内旅行_消費動向');
  if (!sheet || !sheet.length) return { yms: [], travelers: {}, spend: {} };
  const timeCol = estatFindCol(sheet[0], ['時間軸', '時点', '月']);
  const valCol = estatFindCol(sheet[0], ['値']);
  const indicatorCol = estatFindCol(sheet[0], ['指標', '項目', '表側']);
  if (!timeCol || !valCol) return { yms: [], travelers: {}, spend: {} };

  const travelers = {}, spend = {};
  const ymSet = new Set();
  sheet.forEach(r => {
    const ym = estatExtractYm(r[timeCol]);
    const v = estatParseVal(r[valCol]);
    if (!ym || isNaN(v)) return;
    ymSet.add(ym);
    const indicator = indicatorCol ? String(r[indicatorCol] || '') : '';
    if (indicator.indexOf('延べ旅行者') >= 0 || indicator.indexOf('旅行者数') >= 0) {
      travelers[ym] = (travelers[ym] || 0) + v;
    }
    if (indicator.indexOf('消費額') >= 0 || indicator.indexOf('消費単価') >= 0) {
      spend[ym] = Math.max(spend[ym] || 0, v); // 単価は最大値を採用
    }
  });
  return { yms: [...ymSet].sort(), travelers, spend };
}

// CPI宿泊料（CPI_宿泊）
function estatBuildCpi() {
  const sheet = estatSheet('CPI_宿泊');
  if (!sheet || !sheet.length) return { yms: [], values: {} };
  const timeCol = estatFindCol(sheet[0], ['時間軸', '時点', '月']);
  const valCol = estatFindCol(sheet[0], ['値']);
  if (!timeCol || !valCol) return { yms: [], values: {} };

  const values = {};
  const ymSet = new Set();
  sheet.forEach(r => {
    const ym = estatExtractYm(r[timeCol]);
    const v = estatParseVal(r[valCol]);
    if (!ym || isNaN(v) || v <= 0) return;
    ymSet.add(ym);
    values[ym] = v;
  });
  return { yms: [...ymSet].sort(), values };
}

// 延べ宿泊者数 外国人/日本人内訳
function estatBuildForeignRatio() {
  const sheet = estatSheet('宿泊統計_延べ宿泊者数');
  if (!sheet || !sheet.length) return { yms: [], byPref: {} };
  const prefCol = estatFindCol(sheet[0], ['所在地', '都道府県']);
  const timeCol = estatFindCol(sheet[0], ['時間軸', '時点', '月']);
  const valCol = estatFindCol(sheet[0], ['値']);
  if (!prefCol || !timeCol || !valCol) return { yms: [], byPref: {} };

  // 「外国人」「日本人」を含む列（表側分類）を探す
  const catCols = Object.keys(sheet[0]).filter(k => {
    const anyHas = sheet.slice(0, 20).some(r => {
      const v = String(r[k] || '');
      return v.indexOf('外国人') >= 0 || v.indexOf('日本人') >= 0;
    });
    return anyHas;
  });
  const catCol = catCols[0] || null;
  if (!catCol) return { yms: [], byPref: {} };

  const byPref = {}; // pref -> {ym -> {jp, fr}}
  const ymSet = new Set();
  sheet.forEach(r => {
    const pref = String(r[prefCol] || '').trim();
    const ym = estatExtractYm(r[timeCol]);
    const v = estatParseVal(r[valCol]);
    const cat = String(r[catCol] || '');
    if (!pref || !ym || isNaN(v)) return;
    ymSet.add(ym);
    if (!byPref[pref]) byPref[pref] = {};
    if (!byPref[pref][ym]) byPref[pref][ym] = { jp: 0, fr: 0 };
    if (cat.indexOf('外国人') >= 0) byPref[pref][ym].fr += v;
    else if (cat.indexOf('日本人') >= 0) byPref[pref][ym].jp += v;
  });
  return { yms: [...ymSet].sort(), byPref };
}

// ============================================================
// TOPサブタブ
// ============================================================
function renderMarketTopTab() {
  renderMarketTopKpi();
  renderMarketTopInsights();
}

function renderMarketTopKpi() {
  const el = document.getElementById('mkt-macro-kpi');
  if (!el) return;
  const v = estatBuildVisitorsByMonth();
  const latest = v.yms[v.yms.length - 1];
  const prev = latest ? `${parseInt(latest.slice(0, 4), 10) - 1}-${latest.slice(5)}` : null;
  const totalLatest = latest ? v.total[latest] : null;
  const totalPrev = prev ? v.total[prev] : null;
  const yoy = (totalLatest && totalPrev) ? ((totalLatest - totalPrev) / totalPrev) * 100 : null;

  const dom = estatBuildDomesticTravel();
  const domLatest = dom.yms[dom.yms.length - 1];
  const domPrev = domLatest ? `${parseInt(domLatest.slice(0, 4), 10) - 1}-${domLatest.slice(5)}` : null;
  const domYoy = (dom.travelers[domLatest] && dom.travelers[domPrev]) ? ((dom.travelers[domLatest] - dom.travelers[domPrev]) / dom.travelers[domPrev]) * 100 : null;

  const cpi = estatBuildCpi();
  const cpiLatest = cpi.yms[cpi.yms.length - 1];
  const cpiPrev = cpiLatest ? `${parseInt(cpiLatest.slice(0, 4), 10) - 1}-${cpiLatest.slice(5)}` : null;
  const cpiYoy = (cpi.values[cpiLatest] && cpi.values[cpiPrev]) ? cpi.values[cpiLatest] - cpi.values[cpiPrev] : null;

  const deltaHtml = (v, suffix, good) => {
    if (v === null || v === undefined) return '—';
    const col = (good ? v >= 0 : v < 0) ? '#34C759' : '#FF3B30';
    const sign = v >= 0 ? '+' : '';
    return `<span style="color:${col};font-weight:600;">${sign}${v.toFixed(1)}${suffix}</span>`;
  };

  el.innerHTML = `<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;">
    <div class="kpi-card">
      <div class="label">訪日客（最新月）</div>
      <div class="value">${totalLatest ? Math.round(totalLatest).toLocaleString() : '—'}</div>
      <div class="sub">YoY ${deltaHtml(yoy, '%', true)} <span style="color:#86868b;">📊e-Stat</span></div>
    </div>
    <div class="kpi-card">
      <div class="label">日本人旅行者（最新月）</div>
      <div class="value">${dom.travelers[domLatest] ? Math.round(dom.travelers[domLatest]).toLocaleString() : '—'}</div>
      <div class="sub">YoY ${deltaHtml(domYoy, '%', true)} <span style="color:#86868b;">📊e-Stat</span></div>
    </div>
    <div class="kpi-card">
      <div class="label">宿泊料CPI</div>
      <div class="value">${cpi.values[cpiLatest] ? cpi.values[cpiLatest].toFixed(1) : '—'}</div>
      <div class="sub">YoY ${deltaHtml(cpiYoy, 'pt', true)} <span style="color:#86868b;">📊e-Stat</span></div>
    </div>
    <div class="kpi-card">
      <div class="label">3都市 市場OCC</div>
      <div class="value" id="topKpiAirdnaOcc">—</div>
      <div class="sub"><span style="color:#86868b;">📊AirDNA</span></div>
    </div>
  </div>`;

  // AirDNA 3都市OCC平均（最新月）
  const cityOccs = MKT_CITIES.map(c => {
    const latest = mktLatestMonth(c);
    if (!latest) return null;
    const sheet = mktSheet(`AD_${c}全域_occupancy`);
    const field = mktFirstValidField(sheet, ['Rate', 'Occupancy', 'rate']);
    return mktValue(sheet, field, latest);
  }).filter(v => v !== null);
  if (cityOccs.length) {
    const avg = cityOccs.reduce((s, v) => s + v, 0) / cityOccs.length;
    const kpiEl = document.getElementById('topKpiAirdnaOcc');
    if (kpiEl) kpiEl.textContent = avg.toFixed(1) + '%';
  }
}

function renderMarketTopInsights() {
  const el = document.getElementById('mkt-top-insights');
  if (!el) return;
  const insights = [];

  const v = estatBuildVisitorsByMonth();
  const latest = v.yms[v.yms.length - 1];
  const prev = latest ? `${parseInt(latest.slice(0, 4), 10) - 1}-${latest.slice(5)}` : null;

  // ① マクロ需要 vs 市場OCC
  if (latest && prev) {
    MKT_CITIES.forEach(c => {
      const adLatest = mktLatestMonth(c);
      if (!adLatest) return;
      const sheet = mktSheet(`AD_${c}全域_occupancy`);
      const f = mktFirstValidField(sheet, ['Rate', 'Occupancy', 'rate']);
      const curOcc = mktValue(sheet, f, adLatest);
      const [ay, am] = adLatest.split('-');
      const prevYm = `${+ay - 1}-${am}`;
      const prevOcc = mktValue(sheet, f, prevYm);
      if (curOcc === null || prevOcc === null) return;
      const occDiff = curOcc - prevOcc;

      const visitorYoy = (v.total[latest] && v.total[prev]) ? ((v.total[latest] - v.total[prev]) / v.total[prev]) * 100 : null;
      if (visitorYoy === null) return;
      if (visitorYoy > 10 && occDiff < 0) {
        insights.push({ level: 'warning', icon: '⚠', category: `${c}需要×供給`,
          title: `${c}: 需要増(+${visitorYoy.toFixed(1)}%)に市場OCCが追従していない`,
          text: `訪日客YoY +${visitorYoy.toFixed(1)}% に対し${c}市場OCCは ${occDiff.toFixed(1)}pt減。供給過多or価格抵抗の可能性`
        });
      } else if (visitorYoy > 10 && occDiff > 5) {
        insights.push({ level: 'success', icon: '🔥', category: `${c}需要×供給`,
          title: `${c}: 需要増×OCC伸長の追い風`,
          text: `訪日客YoY +${visitorYoy.toFixed(1)}% かつ${c}OCC +${occDiff.toFixed(1)}pt。値上げで収益最大化の好機`
        });
      }
    });
  }

  // ② 宿泊料CPI vs AirDNA ADR
  const cpi = estatBuildCpi();
  const cpiLatest = cpi.yms[cpi.yms.length - 1];
  const cpiPrev = cpiLatest ? `${parseInt(cpiLatest.slice(0, 4), 10) - 1}-${cpiLatest.slice(5)}` : null;
  if (cpiLatest && cpiPrev && cpi.values[cpiLatest] && cpi.values[cpiPrev]) {
    const cpiYoy = ((cpi.values[cpiLatest] - cpi.values[cpiPrev]) / cpi.values[cpiPrev]) * 100;
    // 自社ADR YoY
    const now = new Date();
    const curYm = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const prevYm = `${now.getFullYear() - 1}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const curStats = computeOverallStatsMulti([curYm], '全体', true);
    const prevStats = computeOverallStatsMulti([prevYm], '全体', true);
    if (curStats.adr > 0 && prevStats.adr > 0) {
      const ourYoy = ((curStats.adr - prevStats.adr) / prevStats.adr) * 100;
      const gap = ourYoy - cpiYoy;
      if (gap < -3) {
        insights.push({ level: 'warning', icon: '📉', category: '価格追従',
          title: '自社ADRが宿泊料インフレに追従できていない',
          text: `宿泊料CPI YoY +${cpiYoy.toFixed(1)}% vs 自社ADR YoY ${ourYoy >= 0 ? '+' : ''}${ourYoy.toFixed(1)}%（${gap.toFixed(1)}pt遅れ）。市場価格帯へ追従の余地`
        });
      } else if (gap > 5) {
        insights.push({ level: 'success', icon: '🏆', category: '価格追従',
          title: '自社ADRは市場インフレを先行',
          text: `宿泊料CPI +${cpiYoy.toFixed(1)}% / 自社ADR +${ourYoy.toFixed(1)}%（+${gap.toFixed(1)}pt先行）`
        });
      }
    }
  }

  // ③ インバウンド依存度
  const dom = estatBuildDomesticTravel();
  const domLatest = dom.yms[dom.yms.length - 1];
  if (domLatest && dom.travelers[domLatest] && v.total[latest]) {
    const ratio = v.total[latest] / (v.total[latest] + dom.travelers[domLatest]) * 100;
    insights.push({ level: 'info', icon: '🌏', category: '依存度',
      title: `国内宿泊市場のインバウンド比率: ${ratio.toFixed(1)}%`,
      text: `訪日客 ${Math.round(v.total[latest]).toLocaleString()} / 日本人旅行者 ${Math.round(dom.travelers[domLatest]).toLocaleString()}。自社ポートフォリオのインバウンド依存度と照らして方向性判断を`
    });
  }

  // ④ 国別伸び率トップ
  if (latest && prev) {
    const growth = [];
    Object.keys(v.byCountry).forEach(country => {
      const cur = v.byCountry[country][latest];
      const pv = v.byCountry[country][prev];
      if (!cur || !pv || pv < 1000) return;
      growth.push({ country, yoy: ((cur - pv) / pv) * 100, cur });
    });
    growth.sort((a, b) => b.yoy - a.yoy);
    const top3 = growth.slice(0, 3);
    if (top3.length > 0) {
      insights.push({ level: 'info', icon: '📈', category: '国別トレンド',
        title: '訪日客YoY伸長トップ3',
        text: top3.map(g => `${g.country} +${g.yoy.toFixed(0)}%`).join(' / ') + '。自社ゲスト構成比と比較して狙う国を検討'
      });
    }
  }

  // ⑤ デフォルト
  if (insights.length === 0) {
    insights.push({ level: 'info', icon: '✓', category: '総合', title: 'データ不足', text: '十分な比較データが揃うまでお待ちください' });
  }

  const levelColors = {
    success: { bg: '#34C75912', border: '#34C759', text: '#34C759' },
    warning: { bg: '#FF950012', border: '#FF9500', text: '#FF9500' },
    info: { bg: '#007AFF12', border: '#007AFF', text: '#007AFF' },
  };
  el.innerHTML = insights.map(ins => {
    const c = levelColors[ins.level] || levelColors.info;
    return `<div style="background:${c.bg};border-left:3px solid ${c.border};border-radius:6px;padding:10px 14px;margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
        <span style="font-size:16px;">${ins.icon}</span>
        <span style="font-weight:700;font-size:13px;color:${c.text};">${ins.title}</span>
        <span style="font-size:10px;color:#86868b;margin-left:auto;background:rgba(0,0,0,0.04);padding:2px 6px;border-radius:4px;">${ins.category}</span>
      </div>
      <div style="font-size:12px;color:#1d1d1f;line-height:1.5;">${ins.text}</div>
    </div>`;
  }).join('');
}

function buildMarketInsights(cityKpis, focusCity, wardStats, selYms) {
  const list = [];
  // YoY 各都市のRevPAR伸び率
  const yoys = MKT_CITIES.map(c => {
    const { cur, prev } = cityKpis[c];
    if (!cur.revpar || !prev.revpar) return null;
    return { city: c, pct: ((cur.revpar - prev.revpar) / prev.revpar) * 100, cur: cur.revpar };
  }).filter(Boolean);
  if (yoys.length >= 2) {
    const sorted = [...yoys].sort((a, b) => b.pct - a.pct);
    list.push({ level: sorted[0].pct >= 0 ? 'success' : 'warning', icon: '📈', category: 'YoY',
      title: `最もRevPAR伸長: ${sorted[0].city}（${sorted[0].pct >= 0 ? '+' : ''}${sorted[0].pct.toFixed(1)}%）`,
      text: yoys.map(y => `${y.city}: ${y.pct >= 0 ? '+' : ''}${y.pct.toFixed(1)}%`).join(' / ')
    });
  }
  // 3都市のOCC比較
  const occs = MKT_CITIES.map(c => ({ city: c, v: cityKpis[c].cur.occ })).filter(x => x.v !== null);
  if (occs.length >= 2) {
    const max = occs.reduce((a, b) => a.v > b.v ? a : b);
    const min = occs.reduce((a, b) => a.v < b.v ? a : b);
    const gap = max.v - min.v;
    if (gap >= 5) {
      list.push({ level: 'info', icon: '🏙', category: '都市間OCC',
        title: `${max.city}が最高稼働（${max.v.toFixed(1)}%）`,
        text: `${min.city}（${min.v.toFixed(1)}%）との差 ${gap.toFixed(1)}pt。${max.city}は供給増のタイミングの可能性、${min.city}は価格調整の余地`
      });
    }
  }
  // 3都市のADR比較
  const adrs = MKT_CITIES.map(c => ({ city: c, v: cityKpis[c].cur.adr })).filter(x => x.v !== null);
  if (adrs.length >= 2) {
    const max = adrs.reduce((a, b) => a.v > b.v ? a : b);
    const min = adrs.reduce((a, b) => a.v < b.v ? a : b);
    const gapPct = ((max.v - min.v) / min.v) * 100;
    list.push({ level: 'info', icon: '💴', category: '都市間ADR',
      title: `最高ADR: ${max.city}（¥${Math.round(max.v).toLocaleString()}）`,
      text: `最安 ${min.city}（¥${Math.round(min.v).toLocaleString()}）比 +${gapPct.toFixed(1)}%`
    });
  }
  // 区ランキング（トップ格差）
  if (wardStats && wardStats.length >= 5) {
    const top = wardStats[0], bot = wardStats[wardStats.length - 1];
    const gapPct = ((top.revpar - bot.revpar) / bot.revpar) * 100;
    list.push({ level: 'info', icon: '📍', category: `${focusCity}区別`,
      title: `最高 ${top.jp}（¥${Math.round(top.revpar).toLocaleString()}） vs 最低 ${bot.jp}（¥${Math.round(bot.revpar).toLocaleString()}）`,
      text: `RevPAR差 ${gapPct.toFixed(0)}%。立地・価格帯戦略の参考に`
    });
  }
  // 自社 vs 市場（全体OCC）
  try {
    const ourOcc = MKT_CITIES.map(c => {
      const props = (propertyMaster || []).filter(p => p.area === c && !p.excludeKpi);
      let nights = 0, avail = 0;
      selYms.forEach(ym => {
        const days = getDaysInMonth(ym);
        props.forEach(p => {
          const s = computePropertyStats(p.name, ym);
          if (s) nights += s.nights;
          avail += days * (p.rooms || 1);
        });
      });
      return { city: c, occ: avail > 0 ? (nights / avail) * 100 : null };
    });
    ourOcc.forEach(o => {
      if (o.occ === null) return;
      const mktOcc = cityKpis[o.city].cur.occ;
      if (mktOcc === null) return;
      const diff = o.occ - mktOcc;
      if (Math.abs(diff) >= 5) {
        list.push({
          level: diff >= 0 ? 'success' : 'warning',
          icon: diff >= 0 ? '🏆' : '⚠',
          category: '自社vs市場',
          title: `${o.city}: 自社 ${o.occ.toFixed(1)}% vs 市場 ${mktOcc.toFixed(1)}%（${diff >= 0 ? '+' : ''}${diff.toFixed(1)}pt）`,
          text: diff >= 0 ? '市場平均を上回る。ADR見直しで収益最大化可能' : 'リスティング・価格帯・写真の改善余地'
        });
      }
    });
  } catch (e) {}

  if (list.length === 0) {
    list.push({ level: 'info', icon: '✓', category: '総合', title: '3都市とも安定推移', text: '特記すべき乖離・トレンド転換は見られません' });
  }

  const levelColors = {
    success: { bg: '#34C75912', border: '#34C759', text: '#34C759' },
    warning: { bg: '#FF950012', border: '#FF9500', text: '#FF9500' },
    info: { bg: '#007AFF12', border: '#007AFF', text: '#007AFF' },
  };
  return list.map(ins => {
    const c = levelColors[ins.level] || levelColors.info;
    return `<div style="background:${c.bg};border-left:3px solid ${c.border};border-radius:6px;padding:10px 14px;margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
        <span style="font-size:16px;">${ins.icon}</span>
        <span style="font-weight:700;font-size:13px;color:${c.text};">${ins.title}</span>
        <span style="font-size:10px;color:#86868b;margin-left:auto;background:rgba(0,0,0,0.04);padding:2px 6px;border-radius:4px;">${ins.category}</span>
      </div>
      <div style="font-size:12px;color:#1d1d1f;line-height:1.5;">${ins.text}</div>
    </div>`;
  }).join('');
}

const _mktCharts = {};
function initMarketCharts() {
  Object.keys(_mktCharts).forEach(k => {
    if (_mktCharts[k]) { try { _mktCharts[k].destroy(); } catch (e) {} _mktCharts[k] = null; }
  });
  const sub = currentFilters.marketSubTab || 'top';
  if (sub === 'top') return initMarketTopCharts();
  if (sub === 'macro') return initMarketMacroCharts();
  return initMarketAirdnaCharts();
}

// ============================================================
// TOPサブタブのチャート
// ============================================================
function initMarketTopCharts() {
  // ① 需要×供給マトリクス
  const ctxMatrix = document.getElementById('mktTopMatrix');
  if (ctxMatrix) {
    const v = estatBuildVisitorsByMonth();
    const latest = v.yms[v.yms.length - 1];
    const prev = latest ? `${parseInt(latest.slice(0, 4), 10) - 1}-${latest.slice(5)}` : null;
    const points = [];
    // 3都市同じX軸（全国YoY）で重なるので、都市別booking_demand YoYをX軸に使う
    MKT_CITIES.forEach(c => {
      const occLatest = mktLatestMonth(c);
      if (!occLatest) return;
      const occSheet = mktSheet(`AD_${c}全域_occupancy`);
      const occ = mktValue(occSheet, mktFirstValidField(occSheet, ['Rate', 'Occupancy', 'rate']), occLatest);
      // 都市別需要YoY: booking_demand の nights booked 前年同月比
      const bdSheet = mktSheet(`AD_${c}全域_booking_demand`);
      const bdField = mktFirstValidField(bdSheet, ['Booking demand nights booked', 'Nights booked', 'nights_booked']);
      const prevYm = `${parseInt(occLatest.slice(0, 4), 10) - 1}-${occLatest.slice(5)}`;
      const bdCur = mktValue(bdSheet, bdField, occLatest);
      const bdPrev = mktValue(bdSheet, bdField, prevYm);
      const demandYoy = (bdCur && bdPrev) ? ((bdCur - bdPrev) / bdPrev) * 100 : null;
      if (occ === null || demandYoy === null) return;
      points.push({ x: demandYoy, y: occ, city: c });
    });
    _mktCharts.topMatrix = new Chart(ctxMatrix, {
      type: 'scatter',
      data: { datasets: points.map(p => ({
        label: p.city, data: [{ x: p.x, y: p.y }],
        backgroundColor: MKT_CITY_COLORS[p.city], pointRadius: 14, pointHoverRadius: 18,
      })) },
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: 需要YoY ${ctx.parsed.x.toFixed(1)}% / 市場OCC ${ctx.parsed.y.toFixed(1)}%` } } },
        scales: {
          x: { title: { display: true, text: '都市別 需要YoY伸び率 (%)（AirDNA booking_demand）' }, ticks: { callback: v => (v >= 0 ? '+' : '') + v + '%' } },
          y: { title: { display: true, text: '市場OCC (%)（AirDNA）' }, beginAtZero: true, max: 100, ticks: { callback: v => v + '%' } }
        }
      }
    });
  }

  // ② 国別訪日構成 vs 自社ゲスト構成
  const ctxCountry = document.getElementById('mktTopCountryCompare');
  if (ctxCountry) {
    const v = estatBuildVisitorsByMonth();
    const recent = v.yms.slice(-3);
    const estatAgg = {};
    recent.forEach(ym => {
      Object.keys(v.byCountry).forEach(c => { estatAgg[c] = (estatAgg[c] || 0) + (v.byCountry[c][ym] || 0); });
    });
    const estatTotal = Object.values(estatAgg).reduce((s, x) => s + x, 0) || 1;
    const estatRanked = Object.entries(estatAgg).sort((a, b) => b[1] - a[1]).slice(0, 10);
    // 自社国別（直近3ヶ月・日本人除外で訪日外国人と同列比較）
    const now = new Date();
    const selfCountries = {};
    const threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 3, 1);
    const NAT_ALIAS = {
      'South Korea': '韓国', 'Korea': '韓国', '大韓民国': '韓国',
      'Taiwan': '台湾', '中華民国': '台湾',
      'China': '中国', '中華人民共和国': '中国', 'PRC': '中国',
      'United States': '米国', 'USA': '米国', 'US': '米国', 'アメリカ': '米国',
      'Hong Kong': '中国〔香港〕', '香港': '中国〔香港〕',
      'Thailand': 'タイ',
      'Australia': 'オーストラリア', '豪州': 'オーストラリア',
      'Philippines': 'フィリピン',
      'Singapore': 'シンガポール',
      'Malaysia': 'マレーシア',
    };
    const normalizeNat = (n) => {
      const s = String(n || '').trim();
      if (NAT_ALIAS[s]) return NAT_ALIAS[s];
      return s;
    };
    reservations.forEach(r => {
      if (r.status === 'キャンセル' || r.status === 'システムキャンセル' || r.status === 'ブロックされた') return;
      if (!r.checkin) return;
      const ci = new Date(r.checkin);
      if (ci < threeMonthsAgo || ci > now) return;
      const nat = normalizeNat(r.nationality);
      if (!nat || nat === '日本' || nat === 'Japan' || nat === '日本国') return; // 外国人のみ
      selfCountries[nat] = (selfCountries[nat] || 0) + 1;
    });
    const selfTotal = Object.values(selfCountries).reduce((s, x) => s + x, 0) || 1;
    const labels = estatRanked.map(([c]) => c);
    const estatPcts = estatRanked.map(([, v]) => (v / estatTotal) * 100);
    const selfPcts = labels.map(l => {
      let sum = 0;
      Object.keys(selfCountries).forEach(sc => {
        if (sc === l || sc.indexOf(l) >= 0 || l.indexOf(sc) >= 0) sum += selfCountries[sc];
      });
      return (sum / selfTotal) * 100;
    });
    _mktCharts.topCountry = new Chart(ctxCountry, {
      type: 'bar',
      data: { labels, datasets: [
        { label: '訪日構成比（e-Stat・外国人のみ）', data: estatPcts, backgroundColor: '#007AFFCC' },
        { label: '自社外国人ゲスト構成比', data: selfPcts, backgroundColor: '#FF9500CC' },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.parsed.y.toFixed(1)}%` } } },
        scales: { x: { grid: { display: false } }, y: { beginAtZero: true, title: { display: true, text: '構成比 (%)' }, ticks: { callback: v => v + '%' } } }
      }
    });
  }

  // ③ 宿泊料CPI vs AirDNA ADR vs 自社ADR
  const ctxPrice = document.getElementById('mktTopPriceCompare');
  if (ctxPrice) {
    const cpi = estatBuildCpi();
    const yms = cpi.yms.slice(-24);
    const labels = yms.map(ym => ym.slice(2).replace('-', '/'));
    const cpiData = yms.map(ym => cpi.values[ym] || null);
    // AirDNA の rates_summary は未取得のため使用不可（occupancy/booking_demand のみ取得中）
    // 自社ADR
    const selfData = yms.map(ym => {
      const s = computeOverallStatsMulti([ym], '全体', true);
      return s.adr > 0 ? s.adr : null;
    });
    // インデックス化（基準=最初のnull以外の値）
    const norm = (arr) => {
      const base = arr.find(v => v !== null);
      return base ? arr.map(v => v !== null ? (v / base) * 100 : null) : arr;
    };
    _mktCharts.topPrice = new Chart(ctxPrice, {
      type: 'line',
      data: { labels, datasets: [
        { label: '宿泊料CPI（e-Stat, 指数）', data: norm(cpiData), borderColor: '#007AFF', backgroundColor: 'transparent', tension: 0.3, spanGaps: true, pointRadius: 0 },
        { label: '自社ADR（指数化）', data: norm(selfData), borderColor: '#34C759', backgroundColor: 'transparent', tension: 0.3, spanGaps: true, pointRadius: 2 },
      ]},
      options: {
        responsive: true,
        interaction: { mode: 'index', intersect: false },
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.parsed.y !== null ? ctx.parsed.y.toFixed(1) : '—'}` } } },
        scales: { x: { ticks: { maxTicksLimit: 12 }, grid: { display: false } }, y: { title: { display: true, text: '指数（開始時=100）' } } }
      }
    });
  }

  // ④ 訪日客数 vs 日本人旅行者数
  const ctxInDom = document.getElementById('mktTopInboundVsDomestic');
  if (ctxInDom) {
    const v = estatBuildVisitorsByMonth();
    const dom = estatBuildDomesticTravel();
    const yms = [...new Set([...v.yms, ...dom.yms])].sort().slice(-24);
    const labels = yms.map(ym => ym.slice(2).replace('-', '/'));
    _mktCharts.topInDom = new Chart(ctxInDom, {
      type: 'line',
      data: { labels, datasets: [
        { label: '訪日客数（e-Stat）', data: yms.map(ym => v.total[ym] || null), borderColor: '#007AFF', backgroundColor: '#007AFF22', fill: true, tension: 0.3, yAxisID: 'y', spanGaps: true, pointRadius: 0 },
        { label: '日本人延べ旅行者数（e-Stat）', data: yms.map(ym => dom.travelers[ym] || null), borderColor: '#FF9500', backgroundColor: 'transparent', tension: 0.3, yAxisID: 'y1', spanGaps: true, pointRadius: 0, borderDash: [4, 3] },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.parsed.y !== null ? Math.round(ctx.parsed.y).toLocaleString() : '—'}` } } },
        scales: {
          x: { ticks: { maxTicksLimit: 12 }, grid: { display: false } },
          y: { position: 'left', beginAtZero: true, title: { display: true, text: '訪日客数' } },
          y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: '日本人旅行者数' } }
        }
      }
    });
  }

  // ⑤ 先行予約ペース（AirDNA booking_demand + 自社）
  const ctxBD = document.getElementById('mktTopBookingDemand');
  if (ctxBD) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const labels = [];
    const dates = [];
    for (let i = 0; i < 90; i++) {
      const d = new Date(today); d.setDate(d.getDate() + i);
      labels.push(`${d.getMonth() + 1}/${d.getDate()}`);
      dates.push(d.toISOString().split('T')[0]);
    }
    // AirDNA booking_demand は月次なので日次には展開できない → 月次値を日次に適用（単位: 予約済物件数）
    const airdnaSeries = MKT_CITIES.map(c => {
      const sheet = mktSheet(`AD_${c}全域_booking_demand`);
      const f = mktFirstValidField(sheet, ['Booking demand booked properties', 'Booked properties', 'booked_properties']);
      return {
        city: c,
        data: dates.map(d => {
          const ym = d.slice(0, 7);
          return sheet ? mktValue(sheet, f, ym) : null;
        })
      };
    });
    // 自社先行OCC（properties 配列から、KPI対象のみ）
    const targetProps = (typeof properties !== 'undefined' ? properties : []).filter(p => !p.excludeKpi);
    const totalRooms = targetProps.reduce((s, p) => s + (p.rooms || 1), 0) || 1;
    const selfData = dates.map(ds => {
      let booked = 0;
      targetProps.forEach(p => {
        const has = reservations.some(r =>
          r.status !== 'キャンセル' && r.status !== 'システムキャンセル' && r.status !== 'ブロックされた' &&
          (r.propCode === p.name || r.property === p.propName) &&
          r.checkin <= ds && ds < r.checkout
        );
        if (has) booked += (p.rooms || 1);
      });
      return Math.round((booked / totalRooms) * 1000) / 10;
    });
    // 7日移動平均
    const selfSmoothed = selfData.map((_, i) => {
      const slice = selfData.slice(Math.max(0, i - 3), Math.min(selfData.length, i + 4));
      return Math.round(slice.reduce((s, v) => s + v, 0) / slice.length);
    });

    _mktCharts.topBD = new Chart(ctxBD, {
      type: 'line',
      data: { labels, datasets: [
        { label: '自社先行OCC（7日MA, 左軸%）', data: selfSmoothed, borderColor: '#34C759', backgroundColor: '#34C75922', fill: true, tension: 0.3, pointRadius: 0, yAxisID: 'y' },
        ...airdnaSeries.filter(s => s.data.some(v => v !== null)).map(s => ({
          label: `${s.city} 市場予約済物件数（AirDNA, 右軸）`, data: s.data, borderColor: MKT_CITY_COLORS[s.city], backgroundColor: 'transparent', tension: 0.3, pointRadius: 0, borderDash: [4, 3], spanGaps: true, yAxisID: 'y1'
        }))
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => {
          const isSelf = ctx.dataset.yAxisID === 'y';
          return `${ctx.dataset.label}: ${ctx.parsed.y !== null ? (isSelf ? ctx.parsed.y + '%' : Math.round(ctx.parsed.y).toLocaleString()) : '—'}`;
        } } } },
        scales: {
          x: { ticks: { maxTicksLimit: 12 }, grid: { display: false } },
          y: { position: 'left', beginAtZero: true, max: 100, title: { display: true, text: '自社OCC (%)' }, ticks: { callback: v => v + '%' } },
          y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: '市場予約済物件数' } }
        }
      }
    });
  }

  // ⑥ 季節性ヒートマップ（24ヶ月×3都市）
  const heatEl = document.getElementById('mktTopSeasonHeat');
  if (heatEl) {
    const v = estatBuildVisitorsByMonth();
    const yms = v.yms.slice(-24);
    // 月ごとの訪日客総数を、過去2年の月別で集計。都市粒度は無いので全国総数
    const vals = yms.map(ym => v.total[ym] || 0);
    const max = Math.max(...vals, 1);
    const colorFor = (val) => {
      const intensity = Math.min(1, val / max);
      const r = Math.round(255 - intensity * 100);
      const g = Math.round(255 - intensity * 170);
      const b = Math.round(255 - intensity * 50);
      return `rgb(${r},${g},${b})`;
    };
    heatEl.innerHTML = `<div style="overflow-x:auto;"><table style="border-collapse:collapse;font-size:11px;">
      <thead><tr>${yms.map(ym => `<th style="padding:4px 6px;text-align:center;font-weight:400;color:#86868b;">${ym.slice(2).replace('-', '/')}</th>`).join('')}</tr></thead>
      <tbody><tr>${yms.map((ym, i) => `<td title="${Math.round(vals[i]).toLocaleString()}" style="padding:10px 6px;text-align:center;background:${colorFor(vals[i])};color:#1d1d1f;min-width:50px;border:1px solid #fff;">${Math.round(vals[i] / 10000)}万</td>`).join('')}</tr></tbody>
    </table></div>
    <div style="font-size:11px;color:#86868b;margin-top:6px;">全国訪日客数（万人/月）— 色濃い=多い</div>`;
  }
}

// ============================================================
// 観光統計サブタブのチャート
// ============================================================
function initMarketMacroCharts() {
  const v = estatBuildVisitorsByMonth();
  const yms12 = v.yms.slice(-12);
  const labels = yms12.map(ym => ym.slice(2).replace('-', '/'));

  // 訪日月次推移（YoY棒+線）
  const ctxVT = document.getElementById('macroVisitorsTrend');
  if (ctxVT) {
    const totals = yms12.map(ym => v.total[ym] || 0);
    const yoys = yms12.map(ym => {
      const [y, m] = ym.split('-'); const pv = `${+y - 1}-${m}`;
      const cur = v.total[ym], prev = v.total[pv];
      return (cur && prev) ? ((cur - prev) / prev) * 100 : null;
    });
    _mktCharts.macroVT = new Chart(ctxVT, {
      type: 'bar',
      data: { labels, datasets: [
        { type: 'bar', label: '訪日客数', data: totals, backgroundColor: '#007AFFCC', yAxisID: 'y' },
        { type: 'line', label: 'YoY (%)', data: yoys, borderColor: '#FF3B30', backgroundColor: 'transparent', tension: 0.3, yAxisID: 'y1', spanGaps: true, pointRadius: 3 },
      ]},
      options: {
        responsive: true,
        plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => ctx.dataset.yAxisID === 'y' ? `${Math.round(ctx.parsed.y).toLocaleString()}人` : `YoY ${ctx.parsed.y !== null ? (ctx.parsed.y >= 0 ? '+' : '') + ctx.parsed.y.toFixed(1) + '%' : '—'}` } } },
        scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '訪日客数' } }, y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'YoY (%)' }, ticks: { callback: val => (val >= 0 ? '+' : '') + val + '%' } } }
      }
    });
  }

  // 国別Top10（最新月）
  const ctxT10 = document.getElementById('macroCountryTop10');
  if (ctxT10) {
    const latest = v.yms[v.yms.length - 1];
    if (latest) {
      const ranks = Object.entries(v.byCountry).map(([c, map]) => [c, map[latest] || 0])
        .filter(([, n]) => n > 0).sort((a, b) => b[1] - a[1]).slice(0, 10);
      _mktCharts.macroT10 = new Chart(ctxT10, {
        type: 'bar',
        data: { labels: ranks.map(r => r[0]), datasets: [{ label: latest, data: ranks.map(r => r[1]), backgroundColor: PALETTE.slice(0, 10) }] },
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${Math.round(ctx.parsed.x).toLocaleString()}人` } } }, scales: { x: { beginAtZero: true, ticks: { callback: val => Math.round(val / 1000) + 'k' } }, y: { grid: { display: false } } } }
      });
    }
  }

  // 国別YoY
  const ctxYoy = document.getElementById('macroCountryYoy');
  if (ctxYoy) {
    const latest = v.yms[v.yms.length - 1];
    if (latest) {
      const [y, m] = latest.split('-');
      const prevYm = `${+y - 1}-${m}`;
      const yoys = Object.entries(v.byCountry).map(([c, map]) => {
        const cur = map[latest], prev = map[prevYm];
        if (!cur || !prev || prev < 500) return null;
        return { country: c, yoy: ((cur - prev) / prev) * 100, cur };
      }).filter(Boolean).sort((a, b) => b.cur - a.cur).slice(0, 10);
      _mktCharts.macroYoy = new Chart(ctxYoy, {
        type: 'bar',
        data: { labels: yoys.map(y => y.country), datasets: [{ label: 'YoY', data: yoys.map(y => y.yoy), backgroundColor: yoys.map(y => y.yoy >= 0 ? '#34C759CC' : '#FF3B30CC') }] },
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => (ctx.parsed.x >= 0 ? '+' : '') + ctx.parsed.x.toFixed(1) + '%' } } }, scales: { x: { title: { display: true, text: 'YoY (%)' }, ticks: { callback: v => (v >= 0 ? '+' : '') + v + '%' } }, y: { grid: { display: false } } } }
      });
    }
  }

  // 地域別構成比推移（直近12ヶ月）
  const ctxRegion = document.getElementById('macroRegionMix');
  if (ctxRegion) {
    const regions = ['東アジア', '東南アジア', '南アジア', '欧米豪', 'その他'];
    const datasets = regions.map(r => {
      const data = yms12.map(ym => {
        let sum = 0;
        Object.keys(v.byCountry).forEach(c => {
          if (classifyRegion(c) === r) sum += (v.byCountry[c][ym] || 0);
        });
        return sum;
      });
      return { label: r, data, backgroundColor: null, stack: 'region' };
    });
    const palette = ['#007AFFCC', '#FF9500CC', '#34C759CC', '#AF52DECC', '#86868bCC'];
    datasets.forEach((d, i) => d.backgroundColor = palette[i]);
    _mktCharts.macroRegion = new Chart(ctxRegion, {
      type: 'bar', data: { labels, datasets },
      options: { responsive: true, plugins: { legend: { display: true } }, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true, title: { display: true, text: '訪日客数（人）' } } } }
    });
  }

  // 日本人 延べ旅行者数 & 消費単価
  const dom = estatBuildDomesticTravel();
  const domYms = dom.yms.slice(-12);
  const domLabels = domYms.map(ym => ym.slice(2).replace('-', '/'));
  const ctxDT = document.getElementById('macroDomesticTravelers');
  if (ctxDT) {
    _mktCharts.macroDT = new Chart(ctxDT, {
      type: 'bar',
      data: { labels: domLabels, datasets: [{ label: '延べ旅行者数', data: domYms.map(ym => dom.travelers[ym] || 0), backgroundColor: '#FF9500CC' }] },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true } } }
    });
  }
  const ctxDS = document.getElementById('macroDomesticSpend');
  if (ctxDS) {
    _mktCharts.macroDS = new Chart(ctxDS, {
      type: 'line',
      data: { labels: domLabels, datasets: [{ label: '消費単価', data: domYms.map(ym => dom.spend[ym] || null), borderColor: '#34C759', backgroundColor: '#34C75922', fill: true, tension: 0.3, spanGaps: true }] },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: false } } }
    });
  }

  // 都道府県別OCC（3都市）
  const ctxPO = document.getElementById('macroPrefOcc');
  if (ctxPO) {
    const pref = estatBuildPrefOcc();
    const prefYms = pref.yms.slice(-12);
    const prefLabels = prefYms.map(ym => ym.slice(2).replace('-', '/'));
    const targets = ['大阪府', '京都府', '東京都'];
    const colors = ['#007AFF', '#FF9500', '#34C759'];
    const datasets = targets.map((p, i) => {
      const data = prefYms.map(ym => pref.byPref[p] ? pref.byPref[p][ym] || null : null);
      return { label: p, data, borderColor: colors[i], backgroundColor: colors[i] + '22', tension: 0.3, pointRadius: 2, spanGaps: true };
    });
    _mktCharts.macroPO = new Chart(ctxPO, {
      type: 'line', data: { labels: prefLabels, datasets },
      options: { responsive: true, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true, max: 100, title: { display: true, text: 'OCC (%)' }, ticks: { callback: v => v + '%' } } } }
    });
  }

  // 外国人比率（3都市）
  const ctxFR = document.getElementById('macroForeignRatio');
  if (ctxFR) {
    const fr = estatBuildForeignRatio();
    const frYms = fr.yms.slice(-12);
    const frLabels = frYms.map(ym => ym.slice(2).replace('-', '/'));
    const targets = ['大阪府', '京都府', '東京都'];
    const colors = ['#007AFF', '#FF9500', '#34C759'];
    const datasets = targets.map((p, i) => {
      const data = frYms.map(ym => {
        const r = fr.byPref[p] && fr.byPref[p][ym];
        if (!r) return null;
        const total = r.jp + r.fr;
        return total > 0 ? (r.fr / total) * 100 : null;
      });
      return { label: `${p} 外国人比率`, data, borderColor: colors[i], backgroundColor: 'transparent', tension: 0.3, pointRadius: 2, spanGaps: true };
    });
    _mktCharts.macroFR = new Chart(ctxFR, {
      type: 'line', data: { labels: frLabels, datasets },
      options: { responsive: true, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true, max: 100, title: { display: true, text: '外国人比率 (%)' }, ticks: { callback: v => v + '%' } } } }
    });
  }

  // インバウンド消費
  const ctxIS = document.getElementById('macroInboundSpend');
  if (ctxIS) {
    const s = estatSheet('インバウンド消費');
    if (s && s.length) {
      const countryCol = estatFindCol(s[0], ['国籍', '国']);
      const valCol = estatFindCol(s[0], ['値']);
      const agg = {};
      s.forEach(r => {
        const c = String(r[countryCol] || '');
        const v = estatParseVal(r[valCol]);
        if (!c || isNaN(v) || v <= 0) return;
        agg[c] = Math.max(agg[c] || 0, v);
      });
      const top = Object.entries(agg).sort((a, b) => b[1] - a[1]).slice(0, 10);
      _mktCharts.macroIS = new Chart(ctxIS, {
        type: 'bar', data: { labels: top.map(t => t[0]), datasets: [{ label: '消費単価', data: top.map(t => t[1]), backgroundColor: PALETTE.slice(0, 10) }] },
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => '¥' + Math.round(ctx.parsed.x).toLocaleString() } } }, scales: { x: { beginAtZero: true, ticks: { callback: v => '¥' + (v / 1000).toFixed(0) + 'k' } }, y: { grid: { display: false } } } }
      });
    }
  }

  // CPI宿泊料
  const ctxCPI = document.getElementById('macroLodgingCpi');
  if (ctxCPI) {
    const cpi = estatBuildCpi();
    const cpiYms = cpi.yms.slice(-24);
    const cpiLabels = cpiYms.map(ym => ym.slice(2).replace('-', '/'));
    _mktCharts.macroCPI = new Chart(ctxCPI, {
      type: 'line', data: { labels: cpiLabels, datasets: [{ label: '宿泊料CPI (2020=100)', data: cpiYms.map(ym => cpi.values[ym] || null), borderColor: '#AF52DE', backgroundColor: '#AF52DE22', fill: true, tension: 0.3, pointRadius: 0, spanGaps: true }] },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { x: { grid: { display: false }, ticks: { maxTicksLimit: 12 } }, y: { title: { display: true, text: '指数' } } } }
    });
  }
}

// ============================================================
// AirDNAサブタブのチャート（従来のinitMarketChartsの中身）
// ============================================================
function initMarketAirdnaCharts() {
  const anyCityData = Object.keys(window._airdnaSheets || {}).some(k => /^AD_(大阪|京都|東京)全域_/.test(k));
  if (!anyCityData) return;

  const latestPerCity = MKT_CITIES.map(c => mktLatestMonth(c)).filter(Boolean);
  const baseLatest = latestPerCity.sort()[0];
  const yms12 = mktMonthsEndingAt(baseLatest, 12);
  const labels = yms12.map(ym => { const [y, m] = ym.split('-'); return `${+m}月`; });

  const buildSeries = (field, sheetSuffix, srcFields) => {
    return MKT_CITIES.map(c => {
      const sheet = mktSheet(`AD_${c}全域_${sheetSuffix}`);
      const f = mktFirstValidField(sheet, srcFields);
      return { city: c, data: yms12.map(ym => mktValue(sheet, f, ym)) };
    });
  };

  const occSeries = buildSeries('occ', 'occupancy', ['Rate', 'Occupancy', 'rate']);
  const adrSeries = buildSeries('adr', 'rates_summary', ['Daily rate', 'Rate', 'daily_rate']);
  const revparSeries = occSeries.map(os => {
    const adr = adrSeries.find(a => a.city === os.city).data;
    return { city: os.city, data: os.data.map((v, i) => (v !== null && adr[i] !== null) ? Math.round((v / 100) * adr[i]) : null) };
  });

  const lineDataset = (s, y = '%') => ({
    label: s.city, data: s.data, borderColor: MKT_CITY_COLORS[s.city],
    backgroundColor: MKT_CITY_COLORS[s.city] + '22', tension: 0.3, pointRadius: 2, spanGaps: true,
  });

  const trendOpts = (yLabel, yFmt) => ({
    responsive: true, interaction: { mode: 'index', intersect: false },
    plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${yFmt(ctx.parsed.y)}` } } },
    scales: {
      x: { grid: { display: false } },
      y: { beginAtZero: yLabel === 'OCC', title: { display: true, text: yLabel, font: { size: 11 } }, ticks: { callback: yFmt } }
    }
  });

  const ctxOcc = document.getElementById('mktChartOccTrend');
  if (ctxOcc) _mktCharts.OccTrend = new Chart(ctxOcc, {
    type: 'line', data: { labels, datasets: occSeries.map(s => lineDataset(s)) },
    options: trendOpts('OCC (%)', v => (v === null ? '—' : v.toFixed(1) + '%')),
  });
  const ctxAdr = document.getElementById('mktChartAdrTrend');
  if (ctxAdr) _mktCharts.AdrTrend = new Chart(ctxAdr, {
    type: 'line', data: { labels, datasets: adrSeries.map(s => lineDataset(s)) },
    options: trendOpts('ADR (¥)', v => (v === null ? '—' : '¥' + Math.round(v).toLocaleString())),
  });
  const ctxRevpar = document.getElementById('mktChartRevparTrend');
  if (ctxRevpar) _mktCharts.RevparTrend = new Chart(ctxRevpar, {
    type: 'line', data: { labels, datasets: revparSeries.map(s => lineDataset(s)) },
    options: trendOpts('RevPAR (¥)', v => (v === null ? '—' : '¥' + Math.round(v).toLocaleString())),
  });

  // YoY OCC差分（pt）
  const yoySeries = MKT_CITIES.map(c => {
    const sheet = mktSheet(`AD_${c}全域_occupancy`);
    const f = mktFirstValidField(sheet, ['Rate', 'Occupancy', 'rate']);
    return {
      city: c,
      data: yms12.map(ym => {
        const cur = mktValue(sheet, f, ym);
        const [y, m] = ym.split('-'); const prevYm = `${+y - 1}-${m}`;
        const prev = mktValue(sheet, f, prevYm);
        if (cur === null || prev === null) return null;
        return +(cur - prev).toFixed(1);
      })
    };
  });
  const ctxYoy = document.getElementById('mktChartYoy');
  if (ctxYoy) _mktCharts.Yoy = new Chart(ctxYoy, {
    type: 'bar',
    data: { labels, datasets: yoySeries.map(s => ({ label: s.city, data: s.data, backgroundColor: MKT_CITY_COLORS[s.city] + 'CC' })) },
    options: {
      responsive: true, plugins: { legend: { display: true }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.parsed.y >= 0 ? '+' : ''}${ctx.parsed.y}pt` } } },
      scales: { x: { grid: { display: false } }, y: { title: { display: true, text: 'YoY OCC差分 (pt)' }, ticks: { callback: v => (v >= 0 ? '+' : '') + v + 'pt' } } }
    }
  });

  // 間取り別比較（選択都市）
  const city = currentFilters.marketCity || '大阪';
  const bedSuffixes = ['Studio', '1BR', '2BR', '3BR', '4BR+'];
  const period = currentFilters.marketPeriod || 'last3';
  const count = period === 'latest' ? 1 : period === 'last3' ? 3 : period === 'last6' ? 6 : 12;
  const selYms = mktMonthsEndingAt(baseLatest, count);
  const bedOcc = bedSuffixes.map(bs => {
    const sheet = mktSheet(`AD_${city}全域_${bs}_occupancy`);
    const f = mktFirstValidField(sheet, ['Rate', 'Occupancy', 'rate']);
    return sheet ? mktAvg(sheet, f, selYms) : null;
  });
  const bedAdr = bedSuffixes.map(bs => {
    const sheet = mktSheet(`AD_${city}全域_${bs}_rates_summary`);
    const f = mktFirstValidField(sheet, ['Daily rate', 'Rate', 'daily_rate']);
    return sheet ? mktAvg(sheet, f, selYms) : null;
  });
  const ctxBeds = document.getElementById('mktChartBeds');
  if (ctxBeds) _mktCharts.Beds = new Chart(ctxBeds, {
    type: 'bar',
    data: { labels: bedSuffixes, datasets: [
      { type: 'bar', label: 'OCC (%)', data: bedOcc.map(v => v !== null ? +v.toFixed(1) : null), backgroundColor: CHART_COLORS.blue + 'CC', yAxisID: 'y' },
      { type: 'line', label: 'ADR (¥)', data: bedAdr.map(v => v !== null ? Math.round(v) : null), borderColor: CHART_COLORS.orange, backgroundColor: 'transparent', tension: 0.3, yAxisID: 'y1', pointRadius: 4 },
    ]},
    options: {
      responsive: true, plugins: { legend: { display: true }, tooltip: { callbacks: {
        label: ctx => ctx.dataset.yAxisID === 'y' ? `OCC: ${ctx.parsed.y !== null ? ctx.parsed.y + '%' : '—'}` : `ADR: ${ctx.parsed.y !== null ? '¥' + ctx.parsed.y.toLocaleString() : '—'}`
      } } },
      scales: {
        x: { grid: { display: false } },
        y: { position: 'left', beginAtZero: true, max: 100, title: { display: true, text: 'OCC (%)' }, ticks: { callback: v => v + '%' } },
        y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'ADR (¥)' }, ticks: { callback: v => '¥' + (v / 1000).toFixed(0) + 'k' } }
      }
    }
  });
}

function initPmbmCharts() {
  const area = currentFilters.pmbmArea;
  Object.keys(_pmbmCharts).forEach(k => {
    if (_pmbmCharts[k]) { try { _pmbmCharts[k].destroy(); } catch (e) {} _pmbmCharts[k] = null; }
  });

  // Build 7-month series (-3 to +3)
  const now = new Date();
  const labels = [];
  const pmSeries = [];
  const bmSeries = [];
  const rateSeries = [];
  for (let i = -3; i <= 3; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
    labels.push((d.getMonth() + 1) + '月');
    const det = computePmBmDetail([ym], area);
    pmSeries.push(Math.round(det.pmSales));
    bmSeries.push(Math.round(det.bmSales));
    rateSeries.push(Math.round(det.pmRate * 10) / 10);
  }

  const monthlyCtx = document.getElementById('chartPmbmMonthly');
  if (monthlyCtx) {
    _pmbmCharts.monthly = new Chart(monthlyCtx.getContext('2d'), {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'PM売上', data: pmSeries, backgroundColor: CHART_COLORS.blue, borderRadius: 4 },
          { label: 'BM売上', data: bmSeries, backgroundColor: CHART_COLORS.green, borderRadius: 4 },
        ],
      },
      options: {
        responsive: true,
        plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${fmtYen(c.parsed.y)}` } } },
        scales: { y: { beginAtZero: true, ticks: { callback: (v) => fmtYen(v) } } },
      },
    });
  }

  const rateCtx = document.getElementById('chartPmbmRate');
  if (rateCtx) {
    _pmbmCharts.rate = new Chart(rateCtx.getContext('2d'), {
      type: 'line',
      data: { labels, datasets: [{ label: 'PM率', data: rateSeries, borderColor: CHART_COLORS.purple, backgroundColor: 'rgba(155,89,182,0.1)', fill: true, tension: 0.3 }] },
      options: {
        responsive: true,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => `PM率: ${c.parsed.y}%` } } },
        scales: { y: { beginAtZero: true, ticks: { callback: (v) => v + '%' } } },
      },
    });
  }

  // Area breakdown (current period)
  const months = getSelectedMonths('pmbm');
  const cur = computePmBmDetail(months, area);
  const areaSorted = [...cur.byArea].sort((a, b) => (b.pm + b.bm) - (a.pm + a.bm));
  const areaCtx = document.getElementById('chartPmbmArea');
  if (areaCtx) {
    _pmbmCharts.area = new Chart(areaCtx.getContext('2d'), {
      type: 'bar',
      data: {
        labels: areaSorted.map(a => a.area),
        datasets: [
          { label: 'PM', data: areaSorted.map(a => Math.round(a.pm)), backgroundColor: CHART_COLORS.blue },
          { label: 'BM', data: areaSorted.map(a => Math.round(a.bm)), backgroundColor: CHART_COLORS.green },
        ],
      },
      options: {
        responsive: true,
        plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${fmtYen(c.parsed.y)}` } } },
        scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, ticks: { callback: (v) => fmtYen(v) } } },
      },
    });
  }
}

// ============================================================
// Tab 7: 新法チェック
// ============================================================
function getCurrentFiscalYear() {
  // 4/1 〜 翌3/31。今が4月以降ならその年が起点、3月以前なら前年が起点。
  const now = new Date();
  const startYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  const start = new Date(startYear, 3, 1); // April 1
  const end = new Date(startYear + 1, 3, 1); // April 1 of next year (exclusive)
  return { start, end, startYear };
}

function countFyNightsForProperty(propName, fyStart, fyEnd) {
  // 該当物件の予約から、FY期間内の泊数を算出（キャンセル除外、30泊以上(マンスリー)除外）
  // また、マンスリー（30泊以上）の件数・FY内泊数を別途集計
  const prop = findPropByName(propName);
  let nights = 0;
  let monthlyCount = 0;
  let monthlyNights = 0;
  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    if (r.propCode !== propName && r.property !== propName && (!prop || r.property !== prop.propName)) return;
    if (!r.checkin || !r.checkout) return;
    const ci = new Date(r.checkin);
    const co = new Date(r.checkout);
    if (isNaN(ci) || isNaN(co)) return;
    // FY範囲内の泊数を算出（[ci, co) ∩ [fyStart, fyEnd)）
    const overlapStart = ci > fyStart ? ci : fyStart;
    const overlapEnd = co < fyEnd ? co : fyEnd;
    const ms = overlapEnd - overlapStart;
    if (ms <= 0) return;
    const fyNights = Math.round(ms / 86400000);
    if (r.nights >= 30) {
      // マンスリー予約は新法カウントから除外
      monthlyCount++;
      monthlyNights += fyNights;
    } else {
      nights += fyNights;
    }
  });
  return { nights, monthlyCount, monthlyNights };
}

function renderShinpouTab() {
  const { start, end, startYear } = getCurrentFiscalYear();
  const fmtFy = d => `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
  const endDisplay = new Date(end.getTime() - 86400000); // 表示用に末日へ
  const fyEl = document.getElementById('shinpou-fy-range');
  if (fyEl) fyEl.textContent = `年度: ${fmtFy(start)} 〜 ${fmtFy(endDisplay)}`;

  // 対象物件: 許可種類が「民泊新法」
  const targetProps = properties.filter(p => p.licenseType === '民泊新法');

  let overCount = 0, warnCount = 0, sumPct = 0;
  let totalMonthlyCount = 0, totalMonthlyNights = 0;

  const rows = targetProps.map(p => {
    const { nights, monthlyCount, monthlyNights } = countFyNightsForProperty(p.name, start, end);
    const limit = p.operationLimitDays || 180;
    const remaining = Math.max(limit - nights, 0);
    const pct = limit > 0 ? (nights / limit) * 100 : 0;
    sumPct += pct;
    totalMonthlyCount += monthlyCount;
    totalMonthlyNights += monthlyNights;

    let status, statusBadge, barClass;
    if (pct > 100) {
      status = '超過';
      statusBadge = 'badge-red';
      barClass = 'progress-red';
      overCount++;
    } else if (pct >= 90) {
      status = '残りわずか';
      statusBadge = 'badge-orange';
      barClass = 'progress-orange';
      warnCount++;
    } else if (pct >= 70) {
      status = '注意';
      statusBadge = 'badge-orange';
      barClass = 'progress-orange';
    } else {
      status = 'OK';
      statusBadge = 'badge-green';
      barClass = 'progress-green';
    }
    const barWidth = Math.min(pct, 100);
    return `<tr>
      <td>${p.name}</td>
      <td>${p.licenseType}</td>
      <td class="text-right">${limit}日</td>
      <td class="text-right">${nights}日</td>
      <td class="text-right">${remaining}日</td>
      <td><div class="progress-bar-bg" style="min-width:120px;"><div class="progress-bar-fill ${barClass}" style="width:${barWidth}%"></div></div></td>
      <td class="text-right">${pct.toFixed(1)}%</td>
      <td><span class="${statusBadge}">${status}</span></td>
      <td class="text-right">${monthlyCount}件</td>
      <td class="text-right">${monthlyNights}日</td>
    </tr>`;
  }).join('');

  const tbody = document.getElementById('shinpou-table');
  if (tbody) {
    tbody.innerHTML = rows || '<tr><td colspan="10" style="color:#999;text-align:center;">対象物件がありません</td></tr>';
  }

  const avgPct = targetProps.length > 0 ? sumPct / targetProps.length : 0;
  document.getElementById('kpi-shinpou-count').textContent = targetProps.length + '件';
  document.getElementById('kpi-shinpou-over').textContent = overCount + '件';
  document.getElementById('kpi-shinpou-warn').textContent = warnCount + '件';
  document.getElementById('kpi-shinpou-avg').textContent = avgPct.toFixed(1) + '%';
  document.getElementById('kpi-shinpou-monthly-count').textContent = totalMonthlyCount + '件';
  document.getElementById('kpi-shinpou-monthly-nights').textContent = totalMonthlyNights + '日';
}

function renderReviewTab() {
  if (!document.getElementById('rv-kpi-avg')) return;
  const filtered = _filterReviews();

  // KPI
  const avg = filtered.length ? (filtered.reduce((s, r) => s + r.star, 0) / filtered.length) : 0;
  const negCount = filtered.filter(r => r.sentiment === 'neg').length;
  const posted = REVIEW_MOCK.logs.filter(l => l.system === 'A_post' && l.status === 'success').length;
  document.getElementById('rv-kpi-avg').textContent = avg.toFixed(2) + ' ★';
  document.getElementById('rv-kpi-avg-sub').textContent = filtered.length ? `n=${filtered.length}` : '';
  document.getElementById('rv-kpi-count').textContent = filtered.length;
  document.getElementById('rv-kpi-count-sub').textContent = `直近${_reviewFilters.period}日`;
  // 記載率: mock = received / (received + 推定未記載)
  const totalReservations = Math.floor(filtered.length / 0.62);
  const rate = totalReservations ? (filtered.length / totalReservations * 100) : 0;
  document.getElementById('rv-kpi-rate').textContent = rate.toFixed(0) + '%';
  const negRate = filtered.length ? (negCount / filtered.length * 100) : 0;
  document.getElementById('rv-kpi-neg').textContent = negRate.toFixed(1) + '%';
  document.getElementById('rv-kpi-neg-sub').textContent = `${negCount} 件`;
  document.getElementById('rv-kpi-posted').textContent = posted;
  document.getElementById('rv-kpi-posted-sub').textContent = '直近14日';

  // 物件別テーブル
  const byProp = {};
  filtered.forEach(r => {
    const k = r.property.code;
    if (!byProp[k]) byProp[k] = { prop: r.property, list: [] };
    byProp[k].list.push(r);
  });
  const propRows = Object.values(byProp).map(o => {
    const cnt = o.list.length;
    const avgS = o.list.reduce((s, r) => s + r.star, 0) / cnt;
    const neg = o.list.filter(r => r.sentiment === 'neg').length;
    const latest = o.list.map(r => r.date).sort().slice(-1)[0];
    const rateP = Math.min(100, 50 + Math.random() * 40);
    return { prop: o.prop, cnt, avg: avgS, neg, latest, rate: rateP };
  }).sort((a, b) => b.cnt - a.cnt);

  // 物件マスタからAirbnb情報をルックアップ（コード正規化マッチ）
  const airbnbByCode = {};
  (propertyMaster || []).forEach(pm => {
    const code = pm['物件コード'] || '';
    if (!code) return;
    airbnbByCode[code] = {
      account: pm['airbnbアカウント'] || '',
      listingId: pm['airbnbリスティングID'] || ''
    };
  });

  const tbody = document.querySelector('#rv-property-table tbody');
  tbody.innerHTML = propRows.map(r => {
    const starColor = r.avg >= 4.7 ? 'positive' : r.avg < 4.3 ? 'negative' : '';
    const negPct = (r.neg / r.cnt * 100).toFixed(0);
    const trend = r.avg >= 4.7 ? '▲' : r.avg < 4.3 ? '▼' : '→';
    const ab = airbnbByCode[r.prop.code] || {};
    const link = ab.listingId
      ? ` <a href="https://www.airbnb.com/rooms/${ab.listingId}" target="_blank" style="font-size:11px;color:#007aff;text-decoration:none;">↗</a>`
      : '';
    const acctBadge = ab.account ? ` <span class="badge-blue">${ab.account}</span>` : '';
    return `<tr>
      <td>${r.prop.name} #${r.prop.room}${link}${acctBadge}</td>
      <td>${r.prop.area}</td>
      <td class="text-right">${r.cnt}</td>
      <td class="text-right ${starColor}">${r.avg.toFixed(2)}</td>
      <td class="text-right">${r.rate.toFixed(0)}%</td>
      <td class="text-right ${r.neg > 0 ? 'negative' : ''}">${negPct}%</td>
      <td>${r.latest}</td>
      <td>${trend}</td>
    </tr>`;
  }).join('');

  // 承認待ち
  const pending = filtered.filter(r => r.replyStatus === 'pending');
  document.getElementById('rv-approval-count').textContent = pending.length;
  document.getElementById('rv-approval-list').innerHTML = pending.length ? pending.map(r => `
    <div style="border:1px solid #ffe1c2;background:#fffbf5;border-radius:8px;padding:14px 16px;margin-bottom:10px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
        <div style="font-size:13px;font-weight:600;">${r.property.name} #${r.property.room} <span style="color:#999;font-weight:400;margin-left:8px;">${r.guest} / ${r.date}</span></div>
        <div><span class="badge-red">★${r.star}</span></div>
      </div>
      <div style="font-size:12px;color:#444;background:#f5f5f7;padding:8px 12px;border-radius:6px;margin-bottom:8px;">
        <strong>レビュー本文:</strong> ${r.body}
      </div>
      <div style="font-size:12px;color:#444;background:#eef5ff;padding:8px 12px;border-radius:6px;margin-bottom:10px;">
        <strong>提案返信（Claude生成）:</strong> ${r.replyDraft}
      </div>
      <div style="display:flex;gap:8px;">
        <button class="refresh-btn" style="background:#34c759;color:white;border:none;">承認</button>
        <button class="refresh-btn" style="background:#ff9500;color:white;border:none;">編集</button>
        <button class="refresh-btn" style="background:#ff3b30;color:white;border:none;">却下</button>
      </div>
    </div>
  `).join('') : '<div style="color:#999;font-size:13px;padding:20px 0;text-align:center;">承認待ちはありません</div>';

  // プライベートFB
  const withPrivate = filtered.filter(r => r.privateFeedback);
  const privByProp = {};
  withPrivate.forEach(r => {
    const k = r.property.code;
    if (!privByProp[k]) privByProp[k] = { prop: r.property, list: [] };
    privByProp[k].list.push(r);
  });
  document.getElementById('rv-private-list').innerHTML = Object.values(privByProp).length
    ? Object.values(privByProp).slice(0, 8).map(o => `
      <div style="border-left:3px solid #ff9500;padding:8px 14px;margin-bottom:10px;background:#fafafa;border-radius:0 6px 6px 0;">
        <div style="font-weight:600;font-size:12px;margin-bottom:6px;">${o.prop.name} #${o.prop.room} <span style="color:#999;font-weight:400;">（${o.list.length}件）</span></div>
        ${o.list.slice(0, 3).map(r => `<div style="font-size:12px;color:#555;margin:3px 0;">・ ${r.privateFeedback} <span style="color:#999;">(${r.date} / ${r.guest})</span></div>`).join('')}
      </div>
    `).join('')
    : '<div style="color:#999;font-size:13px;padding:20px 0;text-align:center;">プライベートフィードバックはありません</div>';

  // レビュー一覧
  let listFiltered = filtered;
  if (_reviewFilters.star !== 'all') {
    if (_reviewFilters.star === 'lt3') listFiltered = listFiltered.filter(r => r.star < 3);
    else listFiltered = listFiltered.filter(r => r.star === parseInt(_reviewFilters.star, 10));
  }
  listFiltered = listFiltered.sort((a, b) => b.date.localeCompare(a.date)).slice(0, 30);
  const listBody = document.querySelector('#rv-list-table tbody');
  listBody.innerHTML = listFiltered.map(r => {
    const sentBadge = r.sentiment === 'pos' ? '<span class="badge-green">ポジ</span>' :
                      r.sentiment === 'neg' ? '<span class="badge-red">ネガ</span>' :
                      '<span class="badge-gray">中立</span>';
    const replyBadge = {
      none: '<span class="badge-gray">未返信</span>',
      draft: '<span class="badge-blue">下書き</span>',
      pending: '<span class="badge-orange">承認待ち</span>',
      approved: '<span class="badge-blue">承認済</span>',
      posted: '<span class="badge-green">投稿済</span>',
      rejected: '<span class="badge-gray">却下</span>'
    }[r.replyStatus];
    const starColor = r.star >= 4 ? 'positive' : r.star <= 2 ? 'negative' : '';
    return `<tr>
      <td>${r.date}</td>
      <td>${r.property.name} #${r.property.room}</td>
      <td>${r.guest}</td>
      <td>${r.lang.toUpperCase()}</td>
      <td class="text-right ${starColor}">★${r.star}</td>
      <td style="max-width:380px;white-space:normal;font-size:12px;">${r.body}</td>
      <td>${sentBadge}</td>
      <td>${replyBadge}</td>
    </tr>`;
  }).join('');

  // ★フィルタのpill clickをbind
  document.querySelectorAll('#rv-list-star-filter .pill').forEach(p => {
    p.onclick = () => setReviewFilter(p, 'star');
  });

  // 実行ログ
  const logBody = document.querySelector('#rv-log-table tbody');
  logBody.innerHTML = REVIEW_MOCK.logs.map(l => {
    const stBadge = l.status === 'success' ? '<span class="badge-green">成功</span>' :
                    l.status === 'skipped' ? '<span class="badge-gray">スキップ</span>' :
                    '<span class="badge-red">失敗</span>';
    return `<tr>
      <td style="font-size:12px;">${l.datetime}</td>
      <td><span class="badge-blue">${l.system}</span></td>
      <td>${l.property.name} #${l.property.room}</td>
      <td>${l.guest}</td>
      <td>${l.lang.toUpperCase()}</td>
      <td>${stBadge}</td>
      <td style="font-size:12px;color:#666;">${l.note}</td>
    </tr>`;
  }).join('');

  // キーワード
  const kwPos = [
    { w: '清潔', n: 42 }, { w: '立地', n: 38 }, { w: '快適', n: 31 },
    { w: '広い', n: 24 }, { w: 'スタッフ', n: 19 }, { w: 'recommended', n: 17 },
    { w: 'comfortable', n: 14 }, { w: '便利', n: 12 }
  ];
  const kwNeg = [
    { w: 'Wi-Fi', n: 8 }, { w: '狭い', n: 5 }, { w: '臭い', n: 3 },
    { w: '騒音', n: 3 }, { w: 'shower', n: 2 }
  ];
  const kwHtml = arr => arr.map(k => {
    const sz = 11 + Math.min(k.n / 4, 8);
    return `<span style="display:inline-block;margin:3px;padding:3px 10px;background:#f0f0f0;border-radius:12px;font-size:${sz}px;">${k.w} <span style="color:#999;font-size:10px;">×${k.n}</span></span>`;
  }).join('');
  document.getElementById('rv-kw-pos').innerHTML = kwHtml(kwPos);
  document.getElementById('rv-kw-neg').innerHTML = kwHtml(kwNeg);

  // フィルタpillのbind（毎回再bind）
  document.querySelectorAll('#review-period-filter .pill').forEach(p => p.onclick = () => setReviewFilter(p, 'period'));
  document.querySelectorAll('#review-area-filter .pill').forEach(p => p.onclick = () => setReviewFilter(p, 'area'));
  document.querySelectorAll('#review-lang-filter .pill').forEach(p => p.onclick = () => setReviewFilter(p, 'lang'));
}

function initReviewCharts() {
  if (typeof Chart === 'undefined') return;
  const filtered = _filterReviews();

  // 月別集計
  const byMonth = {};
  filtered.forEach(r => {
    const m = r.date.slice(0, 7);
    if (!byMonth[m]) byMonth[m] = { stars: [], pos: 0, neu: 0, neg: 0 };
    byMonth[m].stars.push(r.star);
    byMonth[m][r.sentiment]++;
  });
  const months = Object.keys(byMonth).sort();
  const trendData = months.map(m => {
    const ss = byMonth[m].stars;
    return ss.reduce((a, b) => a + b, 0) / ss.length;
  });
  const posData = months.map(m => byMonth[m].pos);
  const neuData = months.map(m => byMonth[m].neu);
  const negData = months.map(m => byMonth[m].neg);

  if (_rvCharts.trend) _rvCharts.trend.destroy();
  const trendCtx = document.getElementById('rv-trend-chart');
  if (trendCtx) {
    _rvCharts.trend = new Chart(trendCtx, {
      type: 'line',
      data: {
        labels: months,
        datasets: [{
          label: '平均★',
          data: trendData,
          borderColor: '#007aff',
          backgroundColor: 'rgba(0,122,255,0.1)',
          tension: 0.3, fill: true
        }]
      },
      options: {
        responsive: true,
        scales: { y: { min: 1, max: 5 } },
        plugins: { legend: { display: false } }
      }
    });
  }

  if (_rvCharts.sent) _rvCharts.sent.destroy();
  const sentCtx = document.getElementById('rv-sentiment-chart');
  if (sentCtx) {
    _rvCharts.sent = new Chart(sentCtx, {
      type: 'bar',
      data: {
        labels: months,
        datasets: [
          { label: 'ポジ', data: posData, backgroundColor: '#34c759' },
          { label: '中立', data: neuData, backgroundColor: '#aaa' },
          { label: 'ネガ', data: negData, backgroundColor: '#ff3b30' }
        ]
      },
      options: {
        responsive: true,
        scales: { x: { stacked: true }, y: { stacked: true } }
      }
    });
  }

  // レーダー（カテゴリ別★平均）
  const cats = ['cleanliness', 'accuracy', 'checkin', 'communication', 'location', 'value'];
  const catLabels = ['清潔さ', '正確さ', 'チェックイン', 'コミュニケーション', 'ロケーション', '価値'];
  const catAvg = cats.map(c => {
    const vals = filtered.map(r => r.stars[c]);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
  });
  if (_rvCharts.radar) _rvCharts.radar.destroy();
  const radarCtx = document.getElementById('rv-radar-chart');
  if (radarCtx) {
    _rvCharts.radar = new Chart(radarCtx, {
      type: 'radar',
      data: {
        labels: catLabels,
        datasets: [{
          label: '平均★',
          data: catAvg,
          borderColor: '#007aff',
          backgroundColor: 'rgba(0,122,255,0.2)',
          pointBackgroundColor: '#007aff'
        }]
      },
      options: {
        responsive: true,
        scales: { r: { min: 0, max: 5, ticks: { stepSize: 1 } } },
        plugins: { legend: { display: false } }
      }
    });
  }
}

// ── Auth ──
const LOGIN_PW = 'TDGnq!uPu@KVM!EcZ3';
let _dataReady = false;
// ログイン画面表示中にバックグラウンドでデータ読み込み開始
loadAllData().then(() => { _dataReady = true; });

function tryLogin() {
  const pw = document.getElementById('login-pw').value;
  if (pw === LOGIN_PW) {
    document.getElementById('login-overlay').classList.add('hidden');
    if (!_dataReady) {
      document.getElementById('loading-overlay').style.display = 'flex';
    }
  } else {
    document.getElementById('login-error').style.display = 'block';
  }
}
document.getElementById('login-pw').addEventListener('keydown', e => { if (e.key === 'Enter') tryLogin(); });
