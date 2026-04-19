/**
 * Background Service Worker
 * ポップアップからのCSVデータをGoogle Sheetsに書き込む
 */

const SPREADSHEET_ID_DEFAULT = '1C7EiYSz-3ohjTy3Ul5zdgqL8cDk24jPj3ySTsAcOPhA';

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'writeToSheet') {
    handleWriteToSheet(msg.sheetName, msg.rows)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }
  if (msg.action === 'sendSlack') {
    fetch(msg.webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: msg.text })
    })
      .then(res => {
        if (res.ok) sendResponse({ success: true });
        else sendResponse({ success: false, error: 'HTTP ' + res.status });
      })
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }
  if (msg.action === 'postJson') {
    // GAS Web App など任意のエンドポイントにJSONをPOSTしてJSONで返す
    fetch(msg.url, { method: 'POST', body: JSON.stringify(msg.body) })
      .then(async res => {
        const text = await res.text();
        try {
          sendResponse({ success: true, status: res.status, data: JSON.parse(text) });
        } catch (e) {
          sendResponse({ success: false, error: 'Response not JSON (' + res.status + '): ' + text.slice(0, 2000) });
        }
      })
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }
  if (msg.action === 'writeCell') {
    handleWriteCell(msg.sheetName, msg.range, msg.values)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }
  if (msg.action === 'readSheet') {
    handleReadSheet(msg.sheetName)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }
  if (msg.action === 'listSheets') {
    handleListSheets(msg.spreadsheetId)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ error: err.message, sheetNames: [] }));
    return true;
  }
  if (msg.action === 'openTab') {
    chrome.tabs.create({ url: msg.url, active: false }, (tab) => {
      if (chrome.runtime.lastError) sendResponse({ error: chrome.runtime.lastError.message });
      else sendResponse({ success: true, tabId: tab.id });
    });
    return true;
  }
  if (msg.action === 'ensureAndWriteSheet') {
    handleEnsureAndWrite(msg.spreadsheetId, msg.sheetName, msg.rows)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }
});

async function handleReadSheet(sheetName) {
  try {
    const token = await getAuthToken();
    const stored = await chrome.storage.local.get(['sheetId']);
    const spreadsheetId = stored.sheetId || SPREADSHEET_ID_DEFAULT;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}`;
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(`読み込み失敗: ${err.error.message}`);
    }
    const data = await res.json();
    return { success: true, values: data.values || [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function handleWriteToSheet(sheetName, rows) {
  try {
    const token = await getAuthToken();
    const stored = await chrome.storage.local.get(['sheetId']);
    const spreadsheetId = stored.sheetId || SPREADSHEET_ID_DEFAULT;

    await ensureSheetSize(token, spreadsheetId, sheetName, rows.length, rows[0]?.length || 42);
    await clearSheet(token, spreadsheetId, sheetName);
    await writeData(token, spreadsheetId, sheetName, rows);

    const now = new Date();
    const syncTime = now.getFullYear() + '-' +
      (now.getMonth() + 1).toString().padStart(2, '0') + '-' +
      now.getDate().toString().padStart(2, '0') + ' ' +
      now.getHours().toString().padStart(2, '0') + ':' +
      now.getMinutes().toString().padStart(2, '0');
    chrome.storage.local.set({ lastSync: syncTime });

    return { success: true, rowCount: rows.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function handleWriteCell(sheetName, range, values) {
  try {
    const token = await getAuthToken();
    const stored = await chrome.storage.local.get(['sheetId']);
    const spreadsheetId = stored.sheetId || SPREADSHEET_ID_DEFAULT;
    const fullRange = `${sheetName}!${range}`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(fullRange)}?valueInputOption=RAW`;
    const res = await fetch(url, {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ range: fullRange, majorDimension: 'ROWS', values })
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error.message);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getAuthToken() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(token);
      }
    });
  });
}

async function ensureSheetSize(token, spreadsheetId, sheetName, rowCount, colCount) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties)`;
  const res = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
  if (!res.ok) throw new Error('シート情報取得失敗');
  const data = await res.json();

  const sheet = data.sheets.find(s => s.properties.title === sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません`);

  const sheetIdNum = sheet.properties.sheetId;
  const currentRows = sheet.properties.gridProperties.rowCount;
  const currentCols = sheet.properties.gridProperties.columnCount;
  const requests = [];

  if (rowCount > currentRows) {
    requests.push({
      updateSheetProperties: {
        properties: { sheetId: sheetIdNum, gridProperties: { rowCount: rowCount + 100 } },
        fields: 'gridProperties.rowCount'
      }
    });
  }
  if (colCount > currentCols) {
    requests.push({
      updateSheetProperties: {
        properties: { sheetId: sheetIdNum, gridProperties: { columnCount: colCount + 5 } },
        fields: 'gridProperties.columnCount'
      }
    });
  }

  if (requests.length > 0) {
    const batchUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
    const batchRes = await fetch(batchUrl, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests })
    });
    if (!batchRes.ok) {
      const err = await batchRes.json();
      throw new Error(`シートサイズ調整失敗: ${err.error.message}`);
    }
  }
}

async function clearSheet(token, sheetId, sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(sheetName)}:clear`;
  const res = await fetch(url, { method: 'POST', headers: { 'Authorization': `Bearer ${token}` } });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(`クリア失敗: ${err.error.message}`);
  }
}

async function writeData(token, sheetId, sheetName, rows) {
  const BATCH_SIZE = 1000;
  for (let i = 0; i < rows.length; i += BATCH_SIZE) {
    const batch = rows.slice(i, i + BATCH_SIZE);
    const startRow = i + 1;
    const endRow = startRow + batch.length - 1;
    const range = `${sheetName}!A${startRow}:ZZ${endRow}`;

    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
    const res = await fetch(url, {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: batch })
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(`書き込み失敗: ${err.error.message}`);
    }
  }
}

// シート名一覧を取得
async function handleListSheets(spreadsheetId) {
  const token = await getAuthToken();
  const sid = spreadsheetId || SPREADSHEET_ID_DEFAULT;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sid}?fields=sheets(properties(title))`;
  const res = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
  if (!res.ok) throw new Error('シート一覧取得失敗');
  const data = await res.json();
  return { sheetNames: (data.sheets || []).map(s => s.properties.title) };
}

// AirDNA用: シートがなければ作成してからデータ書き込み
async function handleEnsureAndWrite(spreadsheetId, sheetName, rows) {
  const token = await getAuthToken();
  const sid = spreadsheetId || SPREADSHEET_ID_DEFAULT;

  // Check if sheet exists
  const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sid}?fields=sheets(properties)`;
  const metaRes = await fetch(metaUrl, { headers: { 'Authorization': `Bearer ${token}` } });
  if (!metaRes.ok) throw new Error('シート情報取得失敗');
  const meta = await metaRes.json();
  const exists = meta.sheets.some(s => s.properties.title === sheetName);

  // Create if not exists
  if (!exists) {
    const batchUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sid}:batchUpdate`;
    const batchRes = await fetch(batchUrl, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests: [{ addSheet: { properties: { title: sheetName, gridProperties: { rowCount: Math.max(rows.length + 10, 40), columnCount: Math.max((rows[0]?.length || 5) + 2, 10) } } } }] })
    });
    if (!batchRes.ok) {
      const err = await batchRes.json();
      throw new Error(`シート作成失敗: ${err.error?.message || 'unknown'}`);
    }
  }

  // Resize, clear, write
  await ensureSheetSize(token, sid, sheetName, rows.length, rows[0]?.length || 10);
  await clearSheet(token, sid, sheetName);
  await writeData(token, sid, sheetName, rows);

  return { success: true, sheetName, rowCount: rows.length };
}
