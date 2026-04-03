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
  dailyPeriod: 'thisMonth',
  ownerPeriod: 'thisMonth',
  propertyPeriod: 'thisMonth',
  reservationPeriod: 'thisMonth',
  revenuePeriod: 'thisMonth',
  propertyView: 'all',
  recentDays: 1,
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
  // Use master lookup if available
  if (window._propCodeLookup) {
    const key = propName + '|' + (roomNum || '');
    if (window._propCodeLookup[key]) return window._propCodeLookup[key];
  }
  if (!roomNum || roomNum === 'ALL') return propName;
  if (propName === roomNum || propName.toLowerCase() === roomNum.toLowerCase()) return propName;
  return propName + roomNum;
}

function processData() {
  // Build propCode lookup first (before any generatePropCode calls)
  const propCodeLookup = {};
  propertyMaster.forEach(pm => {
    const code = pm['物件コード'] || '';
    const pn = pm['物件名'] || '';
    const rn = pm['ルーム番号'] || '';
    if (code && pn) {
      propCodeLookup[pn + '|' + rn] = code;
    }
  });
  window._propCodeLookup = propCodeLookup;

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
      paid: r['支払い済み'] || '',
      roomTag: r['物件タグ'] || '',
    };
  });

  // Sort reservations by date descending
  reservations.sort((a, b) => b.date.localeCompare(a.date));

  // Populate channel filter options
  const channels = [...new Set(reservations.map(r => r.channel).filter(Boolean))];
  const channelSelect = document.getElementById('resv-channel-filter');
  channelSelect.innerHTML = '<option value="">すべて</option>';
  channels.forEach(ch => {
    channelSelect.innerHTML += `<option value="${ch}">${ch}</option>`;
  });

  // Populate status filter options from actual data
  const statuses = [...new Set(reservations.map(r => r.status).filter(Boolean))];
  const statusSelect = document.getElementById('resv-status-filter');
  statusSelect.innerHTML = '<option value="">すべて</option>';
  statuses.forEach(st => {
    statusSelect.innerHTML += `<option value="${st}">${st}</option>`;
  });

  // Build owner lookup
  const ownerMap = {};
  ownerMaster.forEach(om => {
    ownerMap[om['オーナーID']] = {
      id: om['オーナーID'] || '',
      name: om['オーナー名'] || '',
      royalty: om['ロイヤリティ'] || '',
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
    const area = deriveArea(address);
    return {
      name: code,
      propName: pm['物件名'] || code,
      code: code,
      ownerId: pm['オーナーID'] || '',
      ownerName: ownerInfo.name || '',
      royalty: ownerInfo.royalty || '',
      area: area,
      rooms: parseNum(pm['部屋数']) || 1,
      excludeKpi: (pm['KPI除外'] || '') === 'TRUE' || (pm['KPI除外'] || '') === '1',
      status: pm['ステータス'] || '稼働中',
      targetLow: parseNum(pm['閑散期目標']),
      targetNormal: parseNum(pm['通常期目標']),
      targetHigh: parseNum(pm['繁忙期目標']),
    };
  }).filter(Boolean);

  // Build owners array
  const ownerIds = [...new Set(propertyMaster.map(pm => pm['オーナーID']).filter(Boolean))];
  owners = ownerIds.map(oid => {
    const info = ownerMap[oid] || {};
    const ownerProps = properties.filter(p => p.ownerId === oid);
    return {
      id: oid,
      name: info.name || oid,
      royalty: info.royalty || '',
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
  const prop = properties.find(p => p.name === propName);
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

  let bookedDays = propDaily.length;
  let totalSales = propDaily.reduce((s, d) => s + parseNum(d['売上合計']), 0);
  let totalReceived = propDaily.reduce((s, d) => s + parseNum(d['受取金合計']), 0);

  // 予約データ: 日次データにない未来分の確定予約を補完
  const today = new Date().toISOString().split('T')[0];
  const [ymY, ymM] = ym.split('-').map(Number);
  const monthStart = ym + '-01';
  const monthEnd = ym + '-' + String(daysInMonth).padStart(2, '0');

  // 日次データに含まれる日付のセット（重複防止用）
  const dailyDates = new Set();
  propDaily.forEach(d => {
    const date = normalizeDate(d['日付']);
    dailyDates.add(date);
  });

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
  document.querySelectorAll('.tab-btn').forEach((btn, i) => {
    const tabIds = ['daily','owner','property','reservation','revenue','review'];
    btn.classList.toggle('active', tabIds[i] === id);
  });
  document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
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

function setRecentFilter(el) {
  const pills = el.parentElement.querySelectorAll('.pill');
  pills.forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters.recentDays = parseInt(el.dataset.days, 10);
  renderAll();
}

function setPeriodFilter(el, tabId) {
  const pills = el.parentElement.querySelectorAll('.pill');
  pills.forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  currentFilters[tabId + 'Period'] = el.dataset.period;
  renderAll();
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
// Render all
// ============================================================
function renderAll() {
  renderDailyTab();
  renderOwnerTab();
  renderPropertyTab();
  renderReservationTab();
  renderRevenueTab();
  setTimeout(initSortableHeaders, 50);
}

// ============================================================
// Tab 1: 全体概況
// ============================================================
function renderDailyTab() {
  const months = getSelectedMonths('daily');
  const area = currentFilters.dailyArea;

  // Date range display
  const monthSet = new Set(months);
  const firstMonth = months[0];
  const lastMonth = months[months.length - 1];
  const firstDay = firstMonth + '-01';
  const lastDaysInMonth = getDaysInMonth(lastMonth);
  const lastDay = lastMonth + '-' + String(lastDaysInMonth).padStart(2, '0');
  document.getElementById('daily-date-range').textContent = firstDay + ' ~ ' + lastDay;

  // Overall stats for selected period
  const overall = computeOverallStatsMulti(months, area, false);

  // Compute total days and avg daily sales
  const totalDays = months.reduce((s, ym) => s + getDaysInMonth(ym), 0);
  const avgDailySales = totalDays > 0 ? overall.totalSales / totalDays : 0;

  // Average nights per reservation
  const monthResvs = reservations.filter(r => {
    if (r.status === 'システムキャンセル') return false;
    if (!monthSet.has(getYearMonth(r.checkin))) return false;
    if (area !== '全体') {
      const prop = properties.find(p => p.name === r.propCode || p.name === r.property || p.propName === r.property);
      if (!prop || prop.area !== area) return false;
    }
    return true;
  });
  const avgNights = monthResvs.length > 0 ? monthResvs.reduce((s, r) => s + r.nights, 0) / monthResvs.length : 0;

  // KPIs
  document.getElementById('kpi-daily-sales').textContent = fmtYen(overall.totalSales);
  document.getElementById('kpi-daily-avg').textContent = fmtYenFull(Math.round(avgDailySales));
  document.getElementById('kpi-daily-received').textContent = fmtYen(overall.totalReceived);
  document.getElementById('kpi-daily-adr').textContent = fmtYenFull(Math.round(overall.adr));
  document.getElementById('kpi-daily-occ').textContent = fmtPct(overall.occ);
  document.getElementById('kpi-daily-nights').textContent = avgNights.toFixed(1) + '泊';

  // Recent reservations
  const recentDays = currentFilters.recentDays;
  const today = new Date();
  const cutoff = new Date(today);
  cutoff.setDate(cutoff.getDate() - recentDays);
  const cutoffStr = cutoff.toISOString().substring(0, 10);
  const todayStr = today.toISOString().substring(0, 10);

  let recentResv = reservations.filter(r => {
    if (area !== '全体') {
      const prop = properties.find(p => p.name === r.propCode || p.name === r.property || p.propName === r.property);
      if (!prop || prop.area !== area) return false;
    }
    return r.date >= cutoffStr && r.date <= todayStr;
  }).slice(0, 20);

  const tbody = document.getElementById('daily-reservations');
  if (recentResv.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" style="color:#999;text-align:center;">該当する予約はありません</td></tr>';
  } else {
    tbody.innerHTML = recentResv.map(r => `<tr>
      <td>${r.date}</td><td>${r.channel}</td><td>${r.property}</td><td>${r.guest}</td><td>${r.checkin}</td><td>${r.checkout}</td><td>${r.nights}泊</td><td>${fmtYenFull(r.sales)}</td>
    </tr>`).join('');
  }

  // Charts
  initDailyCharts();
}

// ============================================================
// Tab 2: オーナー別分析
// ============================================================
function renderOwnerTab() {
  const months = getSelectedMonths('owner');
  const area = currentFilters.ownerArea;

  // Filter owners by area
  let filteredOwners = owners;
  if (area !== '全体') {
    filteredOwners = owners.filter(o => {
      return o.properties.some(pn => {
        const p = properties.find(pp => pp.name === pn);
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

  // KPIs
  document.getElementById('kpi-owner-count').textContent = filteredOwners.length + '名';
  const avgRate = ownerStats.length > 0 ? ownerStats.reduce((s, o) => s + o.rate, 0) / ownerStats.length : 0;
  document.getElementById('kpi-owner-avg-rate').textContent = fmtPct(avgRate);
  const underCount = ownerStats.filter(o => o.rate < 100).length;
  document.getElementById('kpi-owner-under').innerHTML = underCount + '名' + (underCount > 0 ? ' <span class="badge-orange">要確認</span>' : '');

  // Table
  const tbody = document.getElementById('owner-table');
  tbody.innerHTML = ownerStats.map(o => {
    const rateClass = o.rate >= 100 ? 'positive' : o.rate >= 70 ? '' : 'negative';
    return `<tr class="clickable" onclick="toggleOwnerDrill('${o.id}')">
      <td>${o.name}</td><td>${o.propCount}件</td><td>${o.royalty}</td>
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
function toggleOwnerDrill(ownerId) {
  const container = document.getElementById('owner-drill-container');
  if (activeOwnerDrill === ownerId) {
    container.innerHTML = '';
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
    const prop = properties.find(pp => pp.name === p.name);
    return `<tr class="clickable" onclick="event.stopPropagation();toggleOwnerPropertyDrill('${p.name}')">
      <td>${p.name}</td><td>${p.area}</td><td>${fmtPct(p.occ)}</td><td>${fmtYenFull(Math.round(p.adr))}</td><td>${fmtYenFull(Math.round(p.revpar))}</td><td>${fmtYenFull(p.sales)}</td><td>${fmtYenFull(p.received)}</td><td>${prop && prop.excludeKpi ? '<span class="badge-gray">除外</span>' : '-'}</td>
    </tr>`;
  }).join('');

  container.innerHTML = `<div class="drill-down show">
    <h3>${owner.name}</h3>
    <div class="progress-bar-wrap">
      <div class="progress-bar-label"><span>目標: ${fmtYen(target)}</span><span>実績: ${fmtYen(totalSales)} (${fmtPct(rate)})</span></div>
      <div class="progress-bar-bg"><div class="progress-bar-fill ${barColor}" style="width:${barWidth}%"></div></div>
    </div>
    <div class="table-wrap"><table>
      <thead><tr><th>物件名</th><th>エリア</th><th>今月OCC</th><th>ADR</th><th>RevPAR</th><th>販売金額</th><th>受取金</th><th>KPI除外</th></tr></thead>
      <tbody>${propRows}</tbody>
    </table></div>
    <div id="owner-property-drill-container"></div>
  </div>`;
  setTimeout(initSortableHeaders, 50);
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
  const prop = properties.find(p => p.name === propertyName);
  if (!prop) return;

  const ym = getSelectedMonth('property');

  // Get reservations for this property
  const propObj = properties.find(p => p.name === propertyName);
  const propResvAll = reservations.filter(r => r.propCode === propertyName || r.property === propertyName || (propObj && r.property === propObj.propName));
  const propResv = propResvAll.slice(0, 10);
  let resvRows = propResv.map(r => `<tr><td>${r.date}</td><td>${r.channel}</td><td>${r.guest}</td><td>${r.checkin}</td><td>${r.checkout}</td><td>${r.nights}泊</td><td>${fmtYenFull(r.sales)}</td><td>${r.status}</td></tr>`).join('');
  if (!resvRows) resvRows = '<tr><td colspan="8" style="color:#999;text-align:center;">データなし</td></tr>';

  destroyDrillCharts(prefix);

  container.innerHTML = `<div class="drill-down show" style="margin-top:12px;">
    <h3>${prop.name} <span style="font-size:13px;color:#666;font-weight:400;">(${prop.ownerName} / ${prop.area})</span></h3>
    <div class="chart-grid">
      <div class="card"><h2>月別 販売金額/OCC推移</h2><canvas id="${prefix}ChartSalesOcc"></canvas></div>
      <div class="card"><h2>月別 販売金額/ADR推移</h2><canvas id="${prefix}ChartSalesAdr"></canvas></div>
      <div class="card"><h2>チャネル別売上構成比</h2><canvas id="${prefix}ChartChannel"></canvas></div>
      <div class="card"><h2>ゲスト国籍別</h2><canvas id="${prefix}ChartNationality"></canvas></div>
      <div class="card" id="${prefix}RecentBookings"></div>
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
    }
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
            { type: 'bar', label: '販売金額', data: salesData, backgroundColor: blueBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' }
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
            { type: 'bar', label: '販売金額', data: salesData, backgroundColor: orangeBarColors, borderColor: barBorders, borderWidth: barBorderWidths, yAxisID: 'y' }
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
  }, 100);
}

// ============================================================
// Tab 3: 物件別分析
// ============================================================
function renderPropertyTab() {
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
      const propObj = properties.find(pp => pp.name === p.name);
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
function renderReservationTab() {
  const channelFilter = document.getElementById('resv-channel-filter').value;
  const statusFilter = document.getElementById('resv-status-filter').value;
  const months = getSelectedMonths('reservation');
  const monthSet = new Set(months);

  let filtered = [...reservations];

  // Channel filter
  if (channelFilter) filtered = filtered.filter(r => r.channel === channelFilter);
  // Status filter
  if (statusFilter) filtered = filtered.filter(r => r.status === statusFilter);

  // Period filter (based on checkin date)
  filtered = filtered.filter(r => monthSet.has(getYearMonth(r.checkin)));

  // KPIs
  const totalCount = filtered.length;
  const cancelCount = filtered.filter(r => r.status === 'システムキャンセル').length;
  const confirmedOnly = filtered.filter(r => r.status !== 'システムキャンセル');
  const avgNights = confirmedOnly.length > 0 ? confirmedOnly.reduce((s, r) => s + r.nights, 0) / confirmedOnly.length : 0;
  const avgGuests = confirmedOnly.length > 0 ? confirmedOnly.reduce((s, r) => s + r.guestCount, 0) / confirmedOnly.length : 0;

  document.getElementById('kpi-resv-count').textContent = totalCount + '件';
  document.getElementById('kpi-resv-cancel').textContent = cancelCount + '件';
  document.getElementById('kpi-resv-cancel-rate').textContent = totalCount > 0 ? 'キャンセル率 ' + fmtPct((cancelCount / totalCount) * 100) : '-';
  document.getElementById('kpi-resv-nights').textContent = avgNights.toFixed(1) + '泊';
  document.getElementById('kpi-resv-guests').textContent = avgGuests.toFixed(1) + '名';

  // Table
  const tbody = document.getElementById('reservation-table');
  const displayResv = filtered.slice(0, 100);
  tbody.innerHTML = displayResv.map(r => {
    const statusBadge = r.status === '確認済み' ? 'badge-green' : r.status === 'システムキャンセル' ? 'badge-red' : 'badge-orange';
    return `<tr>
      <td>${r.id}</td><td>${r.channel}</td><td>${r.date}</td><td>${r.property}</td><td>${r.guest}</td><td>${r.nationality}</td><td>${r.guestCount}名</td><td>${r.checkin}</td><td>${r.checkout}</td><td>${r.nights}泊</td><td><span class="${statusBadge}">${r.status}</span></td><td>${fmtYenFull(r.sales)}</td><td>${fmtYenFull(r.received)}</td><td>${r.paid}</td>
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
      const prop = properties.find(p => p.name === r.propCode || p.name === r.property || p.propName === r.property);
      if (!prop || prop.area !== area) return false;
    }
    if (excludeKpi) {
      const prop = properties.find(p => p.name === r.propCode || p.name === r.property || p.propName === r.property);
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
  if (tabId === 'daily') initDailyCharts();
  if (tabId === 'reservation') initReservationCharts();
  if (tabId === 'revenue') initRevenueCharts();
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

  // Collect channel sales per month
  const channelSet = new Set();
  const monthChannelSales = {};
  chartMonths.forEach(ym => { monthChannelSales[ym] = {}; });

  reservations.forEach(r => {
    if (r.status === 'システムキャンセル' || r.status === 'キャンセル') return;
    const ciYm = getYearMonth(r.checkin);
    if (!monthChannelSales[ciYm]) return;
    if (area !== '全体') {
      const prop = properties.find(p => p.name === r.propCode || p.name === r.property || p.propName === r.property);
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
  const months = getSelectedMonths('reservation');
  const monthSet = new Set(months);
  const channelFilter = document.getElementById('resv-channel-filter').value;
  const statusFilter = document.getElementById('resv-status-filter').value;

  let filtered = [...reservations];
  if (channelFilter) filtered = filtered.filter(r => r.channel === channelFilter);
  if (statusFilter) filtered = filtered.filter(r => r.status === statusFilter);
  filtered = filtered.filter(r => monthSet.has(getYearMonth(r.checkin)));

  // Channel breakdown doughnut
  destroyChart('channelBD');
  const channelMap = {};
  filtered.forEach(r => {
    const ch = r.channel || 'その他';
    channelMap[ch] = (channelMap[ch] || 0) + 1;
  });
  const ctx2 = document.getElementById('chartChannelBreakdown');
  if (ctx2) {
    const labels = Object.keys(channelMap);
    const data = Object.values(channelMap);
    const colors = PALETTE;
    allCharts['channelBD'] = new Chart(ctx2, {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: colors.slice(0, labels.length) }] },
      options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } } }
    });
  }

  // Nationality breakdown doughnut
  destroyChart('nationalityBD');
  const natMap = {};
  filtered.forEach(r => {
    const nat = r.nationality || '不明';
    natMap[nat] = (natMap[nat] || 0) + 1;
  });
  const ctx3 = document.getElementById('chartNationalityBreakdown');
  if (ctx3) {
    const labels = Object.keys(natMap).sort((a, b) => natMap[b] - natMap[a]).slice(0, 10);
    const data = labels.map(l => natMap[l]);
    const colors = PALETTE;
    allCharts['nationalityBD'] = new Chart(ctx3, {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: colors.slice(0, labels.length) }] },
      options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } } }
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
      const prop = properties.find(p => p.name === r.propCode || p.name === r.property || p.propName === r.property);
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
const FEEDBACK_WEBHOOK = 'https://hooks.slack.com/services/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';

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
    await fetch(FEEDBACK_WEBHOOK, { method: 'POST', body: JSON.stringify(payload) });
    document.getElementById('fb-result').innerHTML = '<div class="feedback-sent">送信しました</div>';
    document.getElementById('fb-message').value = '';
    setTimeout(closeFeedback, 1500);
  } catch (e) {
    document.getElementById('fb-result').innerHTML = '<div style="color:#ff3b30;font-size:13px;text-align:center;">送信失敗</div>';
  }
  btn.disabled = false;
  btn.textContent = '送信';
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
