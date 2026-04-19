/**
 * Airhost管理画面にワンクリック同期パネルを表示
 * ページ内にスクリプトを注入してAPIを叩き、結果をGoogle Sheetsに書き込む
 *
 * 構成:
 *   content.js (ISOLATED world) — UI、Chrome API、Sheets書き込み
 *   inject.js  (MAIN world)     — Airhost APIコール（ページの認証コンテキスト使用）
 */

(function () {
  'use strict';

  let logPanel = null;
  let syncBtn = null;
  let isSyncing = false;

  // ======== UI ========

  function createSyncUI() {
    if (document.getElementById('dent-sync-panel')) return;

    const panel = document.createElement('div');
    panel.id = 'dent-sync-panel';
    panel.innerHTML = `
      <div class="dent-header" id="dent-toggle">
        <span>📊 Dent データ同期</span>
        <span class="dent-toggle-icon" id="dent-toggle-icon">−</span>
      </div>
      <div class="dent-body" id="dent-body">
        <button class="dent-btn" id="dent-sync-btn">🔄 ワンクリック同期</button>
        <div class="dent-info">日次(過去1年) + 予約(過去1年+未来6ヶ月)</div>
        <div class="dent-log" id="dent-log"></div>
      </div>
    `;
    document.body.appendChild(panel);

    syncBtn = document.getElementById('dent-sync-btn');
    logPanel = document.getElementById('dent-log');

    syncBtn.addEventListener('click', () => {
      if (!isSyncing) startSync();
    });

    document.getElementById('dent-toggle').addEventListener('click', () => {
      const body = document.getElementById('dent-body');
      const icon = document.getElementById('dent-toggle-icon');
      if (body.style.display === 'none') {
        body.style.display = 'block';
        icon.textContent = '−';
      } else {
        body.style.display = 'none';
        icon.textContent = '+';
      }
    });
  }

  function log(msg) {
    if (!logPanel) return;
    logPanel.style.display = 'block';
    const time = new Date().toLocaleTimeString('ja-JP', {
      hour: '2-digit', minute: '2-digit', second: '2-digit'
    });
    logPanel.innerHTML += `<div>[${time}] ${msg}</div>`;
    logPanel.scrollTop = logPanel.scrollHeight;
    console.log(`[Dent Sync] ${msg}`);
  }

  // ======== ページ内スクリプト注入 ========

  function injectPageScript() {
    if (document.getElementById('dent-inject-script')) return;

    // bootstrap.jsはmanifestのcontent_scripts(MAIN world)で既に注入済み
    // inject.jsを追加で読み込む
    const script = document.createElement('script');
    script.id = 'dent-inject-script';
    script.src = chrome.runtime.getURL('inject.js');
    (document.head || document.documentElement).appendChild(script);
  }

  // ページスクリプトへコマンド送信 → 結果をPromiseで受け取る
  let pendingRequests = {};
  let requestId = 0;

  function sendToPage(command, data) {
    return new Promise((resolve, reject) => {
      const id = ++requestId;
      const timeout = setTimeout(() => {
        delete pendingRequests[id];
        reject(new Error('ページスクリプト応答タイムアウト (330秒)'));
      }, 330000);

      pendingRequests[id] = { resolve, reject, timeout };

      window.postMessage({
        source: 'dent-content',
        id: id,
        command: command,
        data: data
      }, '*');
    });
  }

  // ページスクリプトからの応答を受信
  window.addEventListener('message', (event) => {
    if (event.source !== window) return;
    if (!event.data || event.data.source !== 'dent-inject') return;

    const msg = event.data;

    // ログメッセージ
    if (msg.type === 'log') {
      log(msg.message);
      return;
    }

    // リクエスト応答
    if (msg.type === 'response' && pendingRequests[msg.id]) {
      const { resolve, reject, timeout } = pendingRequests[msg.id];
      clearTimeout(timeout);
      delete pendingRequests[msg.id];

      if (msg.error) {
        reject(new Error(msg.error));
      } else {
        resolve(msg.result);
      }
    }
  });

  // ======== CSV Parser ========

  function parseCSV(text) {
    const rows = [];
    let current = '';
    let inQuotes = false;
    let row = [];

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      const next = text[i + 1];

      if (inQuotes) {
        if (ch === '"' && next === '"') {
          current += '"';
          i++;
        } else if (ch === '"') {
          inQuotes = false;
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          row.push(current);
          current = '';
        } else if (ch === '\n' || (ch === '\r' && next === '\n')) {
          row.push(current);
          current = '';
          if (row.length > 1 || row[0] !== '') rows.push(row);
          row = [];
          if (ch === '\r') i++;
        } else {
          current += ch;
        }
      }
    }
    if (current !== '' || row.length > 0) {
      row.push(current);
      rows.push(row);
    }
    return rows;
  }

  // ======== Date Helpers ========

  function formatDate(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function getDateChunks(startDate, endDate, monthsPerChunk) {
    monthsPerChunk = monthsPerChunk || 3;
    const chunks = [];
    const current = new Date(startDate);
    const end = new Date(endDate);

    while (current < end) {
      const chunkEnd = new Date(current);
      chunkEnd.setMonth(chunkEnd.getMonth() + monthsPerChunk);
      chunkEnd.setDate(chunkEnd.getDate() - 1);

      const actualEnd = chunkEnd > end ? new Date(end) : chunkEnd;

      chunks.push({
        start_date: formatDate(current),
        end_date: formatDate(actualEnd)
      });

      const next = new Date(actualEnd);
      next.setDate(next.getDate() + 1);
      current.setTime(next.getTime());
    }

    return chunks;
  }

  // ======== Sync Flow ========

  async function startSync() {
    if (isSyncing) return;
    isSyncing = true;
    syncBtn.disabled = true;
    syncBtn.textContent = '⏳ 同期中...';
    logPanel.innerHTML = '';
    logPanel.style.display = 'block';

    const today = new Date();
    const oneYearAgo = new Date(today);
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    oneYearAgo.setDate(1);
    const sixMonthsLater = new Date(today);
    sixMonthsLater.setMonth(sixMonthsLater.getMonth() + 6);

    try {
      // ---- 日次データ（過去1年） ----
      log('📅 日次データ取得開始（過去1年）');
      const dailyRows = await fetchAllChunks('daily', oneYearAgo, today);
      log(`日次データ合計: ${dailyRows.length}行`);

      // ---- 予約データ（過去1年+未来6ヶ月） ----
      log('📋 予約データ取得開始（過去1年+未来6ヶ月）- 35秒待機...');
      await new Promise(r => setTimeout(r, 35000));
      const resRows = await fetchAllChunks('reservation', oneYearAgo, sixMonthsLater);
      log(`予約データ合計: ${resRows.length}行`);

      // ---- Google Sheets に書き込み ----
      log('📝 スプレッドシートに書き込み中...');

      if (dailyRows.length > 0) {
        log(`  日次データ: ${dailyRows.length}行を書き込み中...`);
        const r1 = await sendToBackground('writeToSheet', {
          sheetName: '日次データ', rows: dailyRows
        });
        if (!r1.success) throw new Error(`日次データ書き込み失敗: ${r1.error}`);
        log(`  日次データ: 完了`);
      }

      if (resRows.length > 0) {
        log(`  予約データ: ${resRows.length}行を書き込み中...`);
        const r2 = await sendToBackground('writeToSheet', {
          sheetName: '予約データ', rows: resRows
        });
        if (!r2.success) throw new Error(`予約データ書き込み失敗: ${r2.error}`);
        log(`  予約データ: 完了`);
      }

      // 最終同期タイムスタンプをSheetsに書き込み
      const syncNow = new Date();
      const syncTs = syncNow.getFullYear() + '-' +
        String(syncNow.getMonth() + 1).padStart(2, '0') + '-' +
        String(syncNow.getDate()).padStart(2, '0') + ' ' +
        String(syncNow.getHours()).padStart(2, '0') + ':' +
        String(syncNow.getMinutes()).padStart(2, '0');
      await sendToBackground('writeCell', {
        sheetName: '設定',
        range: 'A1:B1',
        values: [['最終同期', syncTs]]
      });
      log('✅ 同期完了！');

      // 新法ステータス計算（民泊新法物件の年度内稼働日数）
      let shinpouLines = '';
      let shinpouCalcError = null;
      try {
        shinpouLines = await computeShinpouStatus(resRows);
      } catch (e) {
        shinpouCalcError = e.message;
        log('  新法ステータス計算失敗: ' + e.message);
      }

      // Slack通知（新法ステータス: 別チャネル）— 先に送って結果を取得
      let shinpouNotifyStatus;
      if (shinpouCalcError) {
        shinpouNotifyStatus = { status: 'error', message: '計算失敗: ' + shinpouCalcError };
      } else if (!shinpouLines) {
        shinpouNotifyStatus = { status: 'skipped', message: '対象物件なし' };
      } else {
        shinpouNotifyStatus = await sendShinpouSlackNotification(shinpouLines);
      }

      // LIB清掃依頼シート同期（GAS経由）
      let libCleaningResult = null;
      try {
        libCleaningResult = await syncLibCleaning();
      } catch (e) {
        libCleaningResult = { error: e.message };
        log('  LIB清掃同期失敗: ' + e.message);
      }

      // Slack通知（同期完了）— 新法通知・LIB清掃結果も含める
      const startTime = new Date(today);
      const elapsed = Math.round((Date.now() - startTime.getTime()) / 60000);
      await sendSlackNotification(true, {
        dailyRows: dailyRows.length,
        resRows: resRows.length,
        elapsed: elapsed,
        shinpouNotifyStatus: shinpouNotifyStatus,
        libCleaningResult: libCleaningResult,
      });

    } catch (err) {
      log(`❌ エラー: ${err.message}`);
      console.error('[Dent Sync]', err);
      await sendSlackNotification(false, { error: err.message });
    } finally {
      isSyncing = false;
      syncBtn.disabled = false;
      syncBtn.textContent = '🔄 ワンクリック同期';
    }
  }

  async function fetchAllChunks(type, startDate, endDate) {
    const chunks = getDateChunks(startDate, endDate);
    const allRows = [];
    let headers = null;

    for (let ci = 0; ci < chunks.length; ci++) {
      const chunk = chunks[ci];
      log(`  [${ci + 1}/${chunks.length}] ${chunk.start_date} ~ ${chunk.end_date}`);

      // レート制限対策: 2つ目以降のチャンクは35秒待機
      if (ci > 0) {
        log(`  レート制限待機中（35秒）...`);
        await new Promise(r => setTimeout(r, 35000));
      }

      // ページスクリプト経由でCSVテキストを取得
      const csvText = await sendToPage('exportCSV', {
        type: type,
        dateRange: chunk
      });

      const cleanText = csvText.replace(/^\uFEFF/, '');
      const rows = parseCSV(cleanText);

      if (rows.length === 0) {
        log(`  データなし`);
        continue;
      }

      if (!headers) {
        headers = rows[0];
        allRows.push(rows[0]);
      }
      const dataRows = rows.slice(1);
      allRows.push(...dataRows);
      log(`  ${dataRows.length}行取得`);
    }

    return allRows;
  }

  // ======== Background通信 ========

  function sendToBackground(action, data) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({ action, ...data }, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(response);
        }
      });
    });
  }

  // ======== 新法ステータス計算 ========

  // ======== LIB清掃依頼シート同期（GAS経由） ========
  const GAS_WRITE_URL = 'https://script.google.com/macros/s/AKfycbyiSBemvDdrNFmUdXTdNoK8TBV1oz8AANeIbDY6Qd7DNt8ZPC51Ej9rLr9CjSS4zldI2g/exec';
  const GAS_WRITE_TOKEN = 'dent_dashboard_2026';

  async function syncLibCleaning() {
    log('  LIB清掃同期中...');
    const result = await sendToBackground('postJson', {
      url: GAS_WRITE_URL,
      body: { action: 'syncLibCleaning', token: GAS_WRITE_TOKEN },
    });
    if (!result || !result.success) {
      throw new Error(result?.error || 'GAS call failed');
    }
    const json = result.data;
    if (!json || !json.ok) {
      throw new Error((json && json.error) || 'GAS returned non-ok');
    }
    log(`  LIB清掃: +${json.added}件追加 / ${json.cancelled}件キャンセル`);
    return json;
  }

  async function computeShinpouStatus(resRows) {
    // 物件マスタを取得
    const masterRes = await sendToBackground('readSheet', { sheetName: '物件マスタ' });
    if (!masterRes.success) {
      log('  物件マスタ読み込み失敗: ' + masterRes.error);
      return '';
    }
    const values = masterRes.values;
    if (!values || values.length < 2) return '';
    const header = values[0];
    const rows = values.slice(1);
    const idxCode = header.indexOf('物件コード');
    const idxName = header.indexOf('物件名');
    const idxLicense = header.indexOf('許可種類');
    const idxLimit = header.indexOf('営業日数上限');
    if (idxCode < 0 || idxName < 0 || idxLicense < 0 || idxLimit < 0) {
      log('  物件マスタの必須列が見つかりません');
      return '';
    }

    // 民泊新法物件のみ抽出
    const shinpouProps = rows
      .filter(r => (r[idxLicense] || '') === '民泊新法')
      .map(r => ({
        code: (r[idxCode] || '').trim(),
        name: (r[idxName] || '').trim(),
        limit: parseInt(r[idxLimit], 10) || 180,
      }))
      .filter(p => p.code);
    if (shinpouProps.length === 0) return '';

    // コード集合（マッチング用）
    const codeSet = new Set(shinpouProps.map(p => p.code));

    // 予約データのヘッダーから列インデックスを取得
    if (!resRows || resRows.length < 2) return '';
    const rHeader = resRows[0];
    const rIdxName = rHeader.indexOf('物件名');
    const rIdxRoom = rHeader.indexOf('部屋番号');
    const rIdxCheckin = rHeader.indexOf('チェックイン');
    const rIdxCheckout = rHeader.indexOf('チェックアウト');
    const rIdxStatus = rHeader.indexOf('状態');
    const rIdxNights = rHeader.indexOf('合計日数');
    if (rIdxName < 0 || rIdxCheckin < 0 || rIdxCheckout < 0 || rIdxStatus < 0) {
      log('  予約データの必須列が見つかりません');
      return '';
    }

    // 年度範囲: 4/1 〜 翌3/31
    const now = new Date();
    const startYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
    const fyStart = new Date(startYear, 3, 1);
    const fyEnd = new Date(startYear + 1, 3, 1);

    // 物件コード解決（マスタ全体ではなく shinpou 用の簡易ロジック：連結 → 物件名フォールバック）
    const resolveCode = (propName, roomNum) => {
      if (!propName) return '';
      if (roomNum && propName !== roomNum) {
        const concat = propName + roomNum;
        if (codeSet.has(concat)) return concat;
      }
      if (codeSet.has(propName)) return propName;
      return '';
    };

    // 物件ごとの確定日数を集計
    const nightsByCode = {};
    shinpouProps.forEach(p => { nightsByCode[p.code] = 0; });

    for (let i = 1; i < resRows.length; i++) {
      const row = resRows[i];
      const status = row[rIdxStatus] || '';
      if (status === 'システムキャンセル' || status === 'キャンセル') continue;
      const totalNights = rIdxNights >= 0 ? parseInt(row[rIdxNights], 10) || 0 : 0;
      if (totalNights >= 30) continue; // マンスリー除外
      const code = resolveCode((row[rIdxName] || '').trim(), rIdxRoom >= 0 ? (row[rIdxRoom] || '').trim() : '');
      if (!code || !(code in nightsByCode)) continue;

      const ci = new Date(row[rIdxCheckin]);
      const co = new Date(row[rIdxCheckout]);
      if (isNaN(ci) || isNaN(co)) continue;
      const overlapStart = ci > fyStart ? ci : fyStart;
      const overlapEnd = co < fyEnd ? co : fyEnd;
      const ms = overlapEnd - overlapStart;
      if (ms <= 0) continue;
      nightsByCode[code] += Math.round(ms / 86400000);
    }

    // 行を組み立て（物件名/コード順）
    const lines = shinpouProps
      .sort((a, b) => a.code.localeCompare(b.code))
      .map(p => `${p.code}：${nightsByCode[p.code]}/${p.limit}`)
      .join('\n');
    return lines;
  }

  // ======== Slack通知 ========

  async function sendSlackNotification(success, data) {
    try {
      const stored = await chrome.storage.local.get(['slackWebhookUrl']);
      const webhookUrl = stored.slackWebhookUrl;
      if (!webhookUrl) {
        log('  Slack通知: Webhook未設定（設定画面から設定してください）');
        return;
      }

      let text;
      if (success) {
        text = `✅ Airhost同期完了\n` +
               `日次データ: ${data.dailyRows.toLocaleString()}行\n` +
               `予約データ: ${data.resRows.toLocaleString()}行\n` +
               `所要時間: 約${data.elapsed}分`;
        // 新法通知の結果を追記
        const sn = data.shinpouNotifyStatus;
        if (sn) {
          if (sn.status === 'success') text += `\n📋 新法通知: ✅ 送信完了`;
          else if (sn.status === 'skipped') text += `\n📋 新法通知: ⏭ スキップ（${sn.message}）`;
          else if (sn.status === 'error') text += `\n📋 新法通知: ❌ ${sn.message}`;
        }
        // LIB清掃依頼の結果を追記
        const lc = data.libCleaningResult;
        if (lc) {
          if (lc.error) text += `\n🧹 LIB清掃: ❌ ${lc.error}`;
          else if (!lc.added && !lc.cancelled) text += `\n🧹 LIB清掃: 更新なし`;
          else text += `\n🧹 LIB清掃: +${lc.added || 0}件追加 / ${lc.cancelled || 0}件キャンセル`;
        }
      } else {
        text = `❌ Airhost同期失敗\nエラー: ${data.error}`;
      }

      // CORSを避けるためbackground.js経由で送信
      const result = await sendToBackground('sendSlack', { webhookUrl, text });
      if (result.success) {
        log('  Slack通知送信完了');
      } else {
        log('  Slack通知失敗: ' + (result.error || '不明なエラー'));
      }
    } catch (e) {
      log('  Slack通知失敗: ' + e.message);
    }
  }

  async function sendShinpouSlackNotification(shinpouLines) {
    try {
      const stored = await chrome.storage.local.get(['slackWebhookUrlShinpou']);
      const webhookUrl = stored.slackWebhookUrlShinpou;
      if (!webhookUrl) {
        log('  新法Slack通知: Webhook未設定（設定画面から設定してください）');
        return { status: 'skipped', message: 'Webhook未設定' };
      }
      const now = new Date();
      const dateStr = `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()}`;
      const text = `${dateStr}\n${shinpouLines}`;
      const result = await sendToBackground('sendSlack', { webhookUrl, text });
      if (result.success) {
        log('  新法Slack通知送信完了');
        return { status: 'success' };
      } else {
        const msg = result.error || '不明なエラー';
        log('  新法Slack通知失敗: ' + msg);
        return { status: 'error', message: msg };
      }
    } catch (e) {
      log('  新法Slack通知失敗: ' + e.message);
      return { status: 'error', message: e.message };
    }
  }

  // ======== ポップアップからのトリガー ========

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.action === 'startSync') {
      if (!isSyncing) {
        startSync();
        sendResponse({ ok: true });
      } else {
        sendResponse({ ok: false, error: '同期中です' });
      }
    }
  });

  // ======== Initialize ========

  // document_start で即座にinject.jsを注入（ページJSより先に）
  if (window.location.hostname.includes('airhost.co')) {
    injectPageScript();
  }

  // UIはDOM準備後に作成
  function initUI() {
    if (window.location.hostname.includes('airhost.co')) {
      createSyncUI();
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initUI);
  } else {
    initUI();
  }

  // SPA遷移に対応
  let lastUrl = location.href;
  new MutationObserver(function () {
    if (location.href !== lastUrl) {
      lastUrl = location.href;
      setTimeout(initUI, 1000);
    }
  }).observe(document, { subtree: true, childList: true });

})();
