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
};

// ============================================================
// Fetch all sheets
// ============================================================
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

  const [resv, daily, propMaster, ownMaster, seasMaster] = await Promise.all([
    fetchSheet('予約データ'),
    fetchSheet('日次データ'),
    fetchSheet('物件マスタ'),
    fetchSheet('オーナーマスタ'),
    fetchSheet('シーズンマスタ'),
  ]);

  rawReservations = resv;
  rawDailyData = daily;
  propertyMaster = propMaster;
  ownerMaster = ownMaster;
  seasonMaster = seasMaster;

  processData();
  renderAll();
  updateTimestamp();

  // キャッシュ保存
  saveCache({ resv, daily, propMaster, ownMaster, seasMaster });
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
    };
  }).filter(Boolean);

  // Build property lookup maps for fast access (avoid O(N) find per reservation)
  window._propByName = {};
  window._propByPropName = {};
  properties.forEach(p => {
    window._propByName[p.name] = p;
    if (p.propName) window._propByPropName[p.propName] = p;
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

function filterPropertiesByArea(area) {
  if (!area || area === '全体') return properties;
  return properties.filter(p => p.area === area);
}

function aggregateDailyForMonth(ym, areaFilter, excludeKpi) {
  const filteredProps = filterPropertiesByArea(areaFilter);
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

function computePropertyStats(propName, ym) {
  const prop = findPropByName(propName);
  if (!prop) return null;

  const daysInMonth = getDaysInMonth(ym);
  const totalAvailableDays = daysInMonth * (prop.rooms || 1);

  // 日次データ: 過去〜今日分の実績
  const propDaily = rawDailyData.filter(d => {
    const date = normalizeDate(d['日付']);
    const code = generatePropCode(d['物件名'] || '', d['ルーム番号'] || '');
    const status = d['状態'] || '';
    return code === propName && getYearMonth(date) === ym && status !== 'システムキャンセル';
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
  const today = new Date().toISOString().split('T')[0];
  const [ymY, ymM] = ym.split('-').map(Number);
  const monthStart = ym + '-01';
  const monthEnd = ym + '-' + String(daysInMonth).padStart(2, '0');

  // 予約データから未来分を追加（propCode/property/propNameでマッチ）
  const propReservations = reservations.filter(r => {
    if (r.status === 'キャンセル' || r.status === 'システムキャンセル') return false;
    return r.propCode === propName || r.property === propName || (prop && r.property === prop.propName);
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
    // 売上・受取金を月内泊数で按分
    if (monthNights > 0 && r.nights > 0) {
      futureSales += (r.sales || 0) * (monthNights / r.nights);
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

  return {
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
}

function computeOverallStats(ym, areaFilter, excludeKpi) {
  return computeOverallStatsMulti([ym], areaFilter, excludeKpi);
}

function computeOverallStatsMulti(months, areaFilter, excludeKpi) {
  const { filteredProps } = aggregateDailyForMonth(months[0], areaFilter, excludeKpi);

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
    const tabIds = ['daily','owner','property','reservation','revenue','review','watchlist','shinpou','pmbm'];
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
      occData.push(occ);
      adrData.push(adr);
      salesData.push(mSales);
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
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v / 10000).toFixed(0) + '万' } }, y1: { position: 'right', min: 0, max: 110, grid: { drawOnChartArea: false }, title: { display: true, text: 'OCC (%)', font: { size: 11 } } } } }
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
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v / 10000).toFixed(0) + '万' } }, y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
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

function destroyDrillCharts(prefix) {
  if (chartInstances[prefix+'SalesOcc']) { chartInstances[prefix+'SalesOcc'].destroy(); delete chartInstances[prefix+'SalesOcc']; }
  if (chartInstances[prefix+'SalesAdr']) { chartInstances[prefix+'SalesAdr'].destroy(); delete chartInstances[prefix+'SalesAdr']; }
  if (chartInstances[prefix+'Channel']) { chartInstances[prefix+'Channel'].destroy(); delete chartInstances[prefix+'Channel']; }
  if (chartInstances[prefix+'Nationality']) { chartInstances[prefix+'Nationality'].destroy(); delete chartInstances[prefix+'Nationality']; }
}

// ============================================================
// Shared property detail renderer
// ============================================================
function renderPropertyDetail(container, propertyName, prefix) {
  const prop = findPropByName(propertyName);
  if (!prop) return;

  const ym = getSelectedMonth('property');

  // Get reservations for this property
  const propObj = prop;
  const propResvAll = reservations.filter(r => r.propCode === propertyName || r.property === propertyName || (propObj && r.property === propObj.propName));
  const propResv = propResvAll.slice(0, 10);
  let resvRows = propResv.map(r => `<tr><td>${(r.date || '').slice(0, 10)}</td><td>${r.channel}</td><td>${r.guest}</td><td>${r.checkin}</td><td>${r.checkout}</td><td>${r.nights}泊</td><td>${fmtYenFull(r.sales)}</td><td>${r.status}</td></tr>`).join('');
  if (!resvRows) resvRows = '<tr><td colspan="8" style="color:#999;text-align:center;">データなし</td></tr>';

  // KPI: current, YoY, MoM
  const curStats = computePropertyStats(propertyName, ym);
  const [ymY, ymM] = ym.split('-').map(Number);
  const momYm = `${ymM === 1 ? ymY - 1 : ymY}-${String(ymM === 1 ? 12 : ymM - 1).padStart(2, '0')}`;
  const yoyYm = `${ymY - 1}-${String(ymM).padStart(2, '0')}`;
  const momStats = computePropertyStats(propertyName, momYm);
  const yoyStats = computePropertyStats(propertyName, yoyYm);

  const cOcc = curStats ? curStats.occ : 0;
  const cAdr = curStats ? curStats.adr : 0;
  const cRevpar = curStats ? curStats.revpar : 0;
  const cSales = curStats ? curStats.sales : 0;

  const occVs = fmtVsLinePt(cOcc, yoyStats ? yoyStats.occ : null, momStats ? momStats.occ : null);
  const adrVs = fmtVsLine(cAdr, yoyStats ? yoyStats.adr : null, momStats ? momStats.adr : null);
  const revparVs = fmtVsLine(cRevpar, yoyStats ? yoyStats.revpar : null, momStats ? momStats.revpar : null);
  const salesVs = fmtVsLine(cSales, yoyStats ? yoyStats.sales : null, momStats ? momStats.sales : null);

  // Booking window (lead time): average days between booking date and check-in date
  const activeResv = propResvAll.filter(r => r.status !== 'キャンセル' && r.status !== 'システムキャンセル' && r.date && r.checkin);
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

  destroyDrillCharts(prefix);

  container.innerHTML = `<div class="drill-down show" style="margin-top:12px;">
    <h3>${prop.name} <span style="font-size:13px;color:#666;font-weight:400;">(${prop.ownerName} / ${prop.area})</span></h3>
    ${kpiHtml}
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
      occData.push(stats ? stats.occ : 0);
      adrData.push(stats ? stats.adr : 0);
      salesData.push(stats ? stats.sales : 0);
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
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', min: 0, max: 110, grid: { drawOnChartArea: false }, title: { display: true, text: 'OCC (%)', font: { size: 11 } } } } }
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
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
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
  }, 100);
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

    // 1) 日次データ（実績）
    rawDailyData.forEach(d => {
      const date = normalizeDate(d['日付']);
      const code = generatePropCode(d['物件名'] || '', d['ルーム番号'] || '');
      const status = d['状態'] || '';
      if (code !== propName || getYearMonth(date) !== targetYm || status === 'システムキャンセル') return;
      const day = parseInt(date.split('-')[2], 10);
      const sales = parseNum(d['売上合計']);
      if (!dailyMap[day]) dailyMap[day] = { sales: 0, count: 0 };
      dailyMap[day].sales += sales;
      dailyMap[day].count += 1;
      coveredDates.add(date);
    });

    // 2) 予約データから未来分を補完（日次データにない日のみ）
    const propResv = reservations.filter(r => {
      if (r.status === 'キャンセル' || r.status === 'システムキャンセル') return false;
      return r.propCode === propName || r.property === propName || (prop && r.property === prop.propName);
    });
    propResv.forEach(r => {
      if (!r.checkin || !r.checkout || !r.nights || r.nights <= 0) return;
      const dailyRate = Math.round((r.sales || 0) / r.nights);
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
    const dayColor = dow === 0 ? '#ff3b30' : dow === 6 ? '#007aff' : '#333';

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
      <div style="font-weight:600;color:${dayColor};margin-bottom:2px;">${day}</div>
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
  const months = getSelectedMonths('property');
  const area = currentFilters.propertyArea;
  const excludeKpi = document.getElementById('excludeKpiToggle') && document.getElementById('excludeKpiToggle').checked;

  const overall = computeOverallStatsMulti(months, area, excludeKpi);

  // KPIs
  document.getElementById('kpi-prop-count').textContent = overall.propertyCount + '件';
  document.getElementById('kpi-prop-occ').textContent = fmtPct(overall.occ);
  document.getElementById('kpi-prop-adr').textContent = fmtYenFull(Math.round(overall.adr));
  document.getElementById('kpi-prop-revpar').textContent = fmtYenFull(Math.round(overall.revpar));
  document.getElementById('kpi-prop-sales').textContent = fmtYen(overall.totalSales);

  // Table - use merged stats from overall
  const tbody = document.getElementById('property-table');
  let filteredProps = filterPropertiesByArea(area);
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

function toggleGroupedDrill(seriesBase, clickedRow) {
  // 既存のドリルダウン行を削除
  const existing = document.getElementById('grouped-drill-row');
  if (existing) {
    destroyDrillCharts('grp');
    existing.remove();
  }

  if (activeGroupedDrill === seriesBase) {
    activeGroupedDrill = null;
    return;
  }
  activeGroupedDrill = seriesBase;

  const months = getSelectedMonths('property');
  const area = currentFilters.propertyArea;
  const excludeKpi = document.getElementById('excludeKpiToggle') && document.getElementById('excludeKpiToggle').checked;
  const overall = computeOverallStatsMulti(months, area, excludeKpi);

  let filteredProps = filterPropertiesByArea(area);
  if (excludeKpi) filteredProps = filteredProps.filter(p => !p.excludeKpi);
  const seriesProps = filteredProps.filter(p => getSeriesBase(p.name) === seriesBase);

  const rows = seriesProps.map(p => {
    const stats = overall.stats.find(s => s.name === p.name);
    if (!stats) return '';
    return `<tr>
      <td>${p.name}</td><td>${fmtPct(stats.occ)}</td><td>${fmtYenFull(Math.round(stats.adr))}</td>
      <td>${fmtYenFull(Math.round(stats.revpar))}</td><td>${stats.nights}泊</td>
      <td>${fmtYenFull(stats.sales)}</td><td>${fmtYenFull(stats.received)}</td>
    </tr>`;
  }).join('');

  // Aggregate totals
  const totalDays = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);
  let totalNights = 0, totalSales = 0, totalReceived = 0, totalAvailable = 0;
  seriesProps.forEach(p => {
    const stats = overall.stats.find(s => s.name === p.name);
    if (stats) { totalNights += stats.nights; totalSales += stats.sales; totalReceived += stats.received; }
    totalAvailable += totalDays * (p.rooms || 1);
  });
  const aggOcc = totalAvailable > 0 ? (totalNights / totalAvailable) * 100 : 0;
  const aggAdr = totalNights > 0 ? totalSales / totalNights : 0;

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
    <div class="kpi-grid-5" style="margin-bottom:16px;">
      <div class="kpi-card"><div class="label">OCC</div><div class="value">${fmtPct(aggOcc)}</div></div>
      <div class="kpi-card"><div class="label">ADR</div><div class="value">${fmtYenFull(Math.round(aggAdr))}</div></div>
      <div class="kpi-card"><div class="label">販売泊数</div><div class="value">${totalNights}泊</div></div>
      <div class="kpi-card"><div class="label">販売金額</div><div class="value">${fmtYen(totalSales)}</div></div>
      <div class="kpi-card"><div class="label">受取金</div><div class="value">${fmtYen(totalReceived)}</div></div>
    </div>
    <div class="chart-grid">
      <div class="card"><h2>月別 販売金額/OCC推移</h2><canvas id="grpChartSalesOcc"></canvas></div>
      <div class="card"><h2>月別 販売金額/ADR推移</h2><canvas id="grpChartSalesAdr"></canvas></div>
      <div class="card"><h2>チャネル別売上構成比</h2><canvas id="grpChartChannel"></canvas></div>
      <div class="card"><h2>ゲスト国籍別</h2><canvas id="grpChartNationality"></canvas></div>
    </div>
    <div class="card"><h2>部屋別内訳</h2><div class="table-wrap"><table>
      <thead><tr><th>部屋</th><th>OCC</th><th>ADR</th><th>RevPAR</th><th>販売泊数</th><th>販売金額</th><th>受取金</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div></div>
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
      occData.push(mOcc);
      adrData.push(mAdr);
      salesData.push(mSales);
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
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', min: 0, max: 110, grid: { drawOnChartArea: false }, title: { display: true, text: 'OCC (%)', font: { size: 11 } } } } }
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
        options: { responsive: true, animation: { duration: 600, easing: 'easeOutQuart' }, plugins: { legend: { display: true } }, scales: { x: { grid: { display: false } }, y: { position: 'left', beginAtZero: true, title: { display: true, text: '販売金額 (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + (v/10000).toFixed(0) + '万' } }, y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'ADR (¥)', font: { size: 11 } }, ticks: { callback: v => '¥' + v.toLocaleString() } } } }
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
        return r.propCode === p.name || r.property === p.name || (propObj && r.property === propObj.propName);
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
  }, 100);
  setTimeout(initSortableHeaders, 150);
  drillRow.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
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
  const prevFiltered = applySelectFilters(reservations).filter(periodInfo.previous);

  // KPIs - 当期
  const totalCount = filtered.length;
  const cancelCount = filtered.filter(r => r.status === 'システムキャンセル').length;
  const confirmedOnly = filtered.filter(r => r.status !== 'システムキャンセル');
  const totalNights = confirmedOnly.reduce((s, r) => s + r.nights, 0);
  const avgNights = confirmedOnly.length > 0 ? totalNights / confirmedOnly.length : 0;
  const avgGuests = confirmedOnly.length > 0 ? confirmedOnly.reduce((s, r) => s + r.guestCount, 0) / confirmedOnly.length : 0;
  const totalSales = confirmedOnly.reduce((s, r) => s + (r.sales || 0), 0);
  const adr = totalNights > 0 ? totalSales / totalNights : 0;
  const cancelSales = filtered.filter(r => r.status === 'システムキャンセル').reduce((s, r) => s + (r.sales || 0), 0);

  // KPIs - 前期間
  const prevTotalCount = prevFiltered.length;
  const prevCancelCount = prevFiltered.filter(r => r.status === 'システムキャンセル').length;
  const prevConfirmed = prevFiltered.filter(r => r.status !== 'システムキャンセル');
  const prevTotalNights = prevConfirmed.reduce((s, r) => s + r.nights, 0);
  const prevTotalSales = prevConfirmed.reduce((s, r) => s + (r.sales || 0), 0);
  const prevAdr = prevTotalNights > 0 ? prevTotalSales / prevTotalNights : 0;
  const prevCancelSales = prevFiltered.filter(r => r.status === 'システムキャンセル').reduce((s, r) => s + (r.sales || 0), 0);

  document.getElementById('kpi-resv-count').textContent = totalCount + '件';
  document.getElementById('kpi-resv-count-vs-cnt').textContent = `${periodInfo.vsLabel} ${fmtVsPct(totalCount, prevTotalCount)}`;
  document.getElementById('kpi-resv-sales').textContent = fmtYenFull(totalSales);
  document.getElementById('kpi-resv-count-vs').textContent = `${periodInfo.vsLabel} ${fmtVsPct(totalSales, prevTotalSales)}`;
  document.getElementById('kpi-resv-adr').textContent = fmtYenFull(Math.round(adr));
  document.getElementById('kpi-resv-adr-vs').textContent = `${periodInfo.vsLabel} ${fmtVsPct(adr, prevAdr)}`;
  document.getElementById('kpi-resv-cancel').textContent = cancelCount + '件';
  document.getElementById('kpi-resv-cancel-rate').textContent = totalCount > 0 ? 'キャンセル率 ' + fmtPct((cancelCount / totalCount) * 100) : '-';
  document.getElementById('kpi-resv-cancel-sales').textContent = fmtYenFull(cancelSales);
  document.getElementById('kpi-resv-cancel-vs').textContent = `${periodInfo.vsLabel} ${fmtVsPct(cancelSales, prevCancelSales)}`;
  document.getElementById('kpi-resv-nights').textContent = avgNights.toFixed(1) + '泊';
  document.getElementById('kpi-resv-guests').textContent = avgGuests.toFixed(1) + '名';

  // Table（キャンセルは表示から除外）
  const tbody = document.getElementById('reservation-table');
  const displayResv = filtered.filter(r => r.status !== 'システムキャンセル' && r.status !== 'キャンセル').slice(0, 100);
  tbody.innerHTML = displayResv.map(r => {
    const statusBadge = r.status === '確認済み' ? 'badge-green' : r.status === 'システムキャンセル' ? 'badge-red' : 'badge-orange';
    return `<tr>
      <td>${(r.date || '').slice(0, 10)}</td><td>${r.channel}</td><td>${r.property}</td><td>${fmtYenFull(r.sales)}</td><td>${r.checkin}</td><td>${r.nights}泊</td><td>${r.guestCount}名</td><td>${r.checkout}</td><td>${r.guest}</td><td>${r.nationality}</td><td><span class="${statusBadge}">${r.status}</span></td><td>${fmtYenFull(r.received)}</td><td>${r.paid}</td><td>${r.id}</td>
    </tr>`;
  }).join('');
}

// ============================================================
// Tab 5: 売上・稼働
// ============================================================
function renderRevenueTab() {
  const months = getSelectedMonths('revenue');
  const monthSet = new Set(months);
  const area = currentFilters.revenueArea;
  const excludeKpi = document.getElementById('excludeKpiToggleRev') && document.getElementById('excludeKpiToggleRev').checked;

  const overall = computeOverallStatsMulti(months, area, excludeKpi);

  document.getElementById('kpi-rev-occ').textContent = fmtPct(overall.occ);
  document.getElementById('kpi-rev-adr').textContent = fmtYenFull(Math.round(overall.adr));
  document.getElementById('kpi-rev-revpar').textContent = fmtYenFull(Math.round(overall.revpar));
  document.getElementById('kpi-rev-sales').textContent = fmtYen(overall.totalSales);
  document.getElementById('kpi-rev-received').textContent = fmtYen(overall.totalReceived);

  // Channel performance table
  const confirmedResv = reservations.filter(r => {
    if (r.status === 'システムキャンセル') return false;
    if (!monthSet.has(getYearMonth(r.checkin))) return false;
    if (area !== '全体') {
      const prop = findPropByReservation(r);
      if (!prop || prop.area !== area) return false;
    }
    if (excludeKpi) {
      const prop = findPropByReservation(r);
      if (prop && prop.excludeKpi) return false;
    }
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

  // 物件別予約数（横棒、件数降順）
  destroyChart('propertyBD');
  const propMap = {};
  const propSalesMap = {};
  filtered.forEach(r => {
    const p = r.property || 'その他';
    propMap[p] = (propMap[p] || 0) + 1;
    propSalesMap[p] = (propSalesMap[p] || 0) + (r.sales || 0);
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
}

// ============================================================
// Init
// ============================================================
// ── Feedback ──
function getFeedbackWebhook() {
  return localStorage.getItem('feedbackWebhookUrl') || '';
}
function saveFeedbackWebhook(url) {
  localStorage.setItem('feedbackWebhookUrl', url);
}

function openFeedback() {
  document.getElementById('feedback-modal').classList.add('show');
  document.getElementById('fb-result').innerHTML = '';
  document.getElementById('fb-send').disabled = false;
  // Webhook URL設定欄の初期値
  const webhookInput = document.getElementById('fb-webhook');
  if (webhookInput) webhookInput.value = getFeedbackWebhook();
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
    const webhookUrl = document.getElementById('fb-webhook').value.trim();
    if (webhookUrl) saveFeedbackWebhook(webhookUrl);
    if (!webhookUrl) { alert('Webhook URLを設定してください'); btn.disabled = false; btn.textContent = '送信'; return; }
    await fetch(webhookUrl, { method: 'POST', mode: 'no-cors', body: 'payload=' + encodeURIComponent(JSON.stringify(payload)) });
    document.getElementById('fb-result').innerHTML = '<div class="feedback-sent">送信しました</div>';
    document.getElementById('fb-message').value = '';
    setTimeout(closeFeedback, 1500);
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
        return `<div class="watchlist-card">
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
        return `<div class="watchlist-card">
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
    } else if (pct >= 95) {
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
