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
        <div class="dent-info">日次(過去90日) + 予約(過去90日+未来365日)</div>
        <details class="dent-backfill">
          <summary>⏳ 過去データ取込（バックフィル）</summary>
          <div class="dent-info">D1に無い過去期間を一度だけ取り込む。定常同期とは別。</div>
          <label class="dent-field">開始 <input type="date" id="dent-bf-start" value="2024-01-01" min="2020-01-01" max="2026-12-31"></label>
          <label class="dent-field">終了 <input type="date" id="dent-bf-end" value="2024-03-31" min="2020-01-01" max="2026-12-31"></label>
          <button class="dent-btn dent-btn-sub" id="dent-backfill-btn">過去データを取り込む</button>
        </details>
        <div class="dent-log" id="dent-log"></div>
      </div>
    `;
    document.body.appendChild(panel);

    syncBtn = document.getElementById('dent-sync-btn');
    logPanel = document.getElementById('dent-log');

    syncBtn.addEventListener('click', () => {
      if (!isSyncing) startSync();
    });

    document.getElementById('dent-backfill-btn').addEventListener('click', () => {
      if (isSyncing) return;
      const start = document.getElementById('dent-bf-start').value;
      const end = document.getElementById('dent-bf-end').value;
      if (!start || !end || start > end) { log('⚠️ 開始/終了日が不正です'); return; }
      startBackfill(start, end);
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
    // 取得ウィンドウ: 過去90日 + 未来365日。
    // 確定済みの過去予約は不変なので、それより前は取得しない（毎日の取得負荷削減）。
    // 全履歴はD1 (dent-data-api) がupsertで永続蓄積するため、Sheetsは直近窓のバッファでよい。
    // 過去90日は alert-anomaly.gs のADRベースライン(BASELINE_WINDOW_DAYS=90)が
    // Sheetsを読むための下限。Sheets読みconsumerをD1移行すれば更に縮められる。
    const windowStart = new Date(today);
    windowStart.setDate(windowStart.getDate() - 90);
    const oneYearLater = new Date(today);
    oneYearLater.setFullYear(oneYearLater.getFullYear() + 1);

    try {
      // ---- 日次データ（過去90日） ----
      log('📅 日次データ取得開始（過去90日）');
      const dailyRows = await fetchAllChunks('daily', windowStart, today);
      log(`日次データ合計: ${dailyRows.length}行`);

      // ---- 予約データ（過去90日+未来365日） ----
      log('📋 予約データ取得開始（過去90日+未来365日）- 35秒待機...');
      await new Promise(r => setTimeout(r, 35000));
      const resRows = await fetchAllChunks('reservation', windowStart, oneYearLater);
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

      // 最新シートで朝の運営サマリーを発火（イベント駆動）。
      // 固定9:00トリガーの代わりに、同期完了時＝データが最新の瞬間に投稿させる。
      // GAS側の冪等ガードで、同日に複数回同期しても本編は1回だけ投稿される。
      // 民泊新法の営業日数チェックも同じ呼び出しでGAS側が投稿する
      // （拡張側での計算は 2026-07-27 に廃止。理由は gas/shinpou-report.gs 冒頭を参照）。
      const triggerResult = await triggerMorningReport();

      // Slack通知（同期完了）— 新法通知の結果も添える
      const startTime = new Date(today);
      const elapsed = Math.round((Date.now() - startTime.getTime()) / 60000);
      await sendSlackNotification(true, {
        dailyRows: dailyRows.length,
        resRows: resRows.length,
        elapsed: elapsed,
        shinpouNotifyStatus: triggerResult && triggerResult.shinpou,
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

  // ======== バックフィル（過去データ一括取込） ========
  //
  // 定常同期(過去90日)とは別に、D1に無い過去期間を一度だけ取り込む。
  // 設計上の要点:
  //  - Sheetsは直近窓バッファなので、全期間を溜めると溢れる。よって
  //    「3ヶ月チャンクごとに Sheets を上書き → 即 D1 へ流す」を繰り返し、
  //    Sheetsには常に1チャンク分しか乗せない。D1はupsertで永続蓄積。
  //  - D1送信は拡張から直接ではなくGAS(d1SyncOnly)経由。CF-Accessトークンを
  //    各ブラウザの拡張に持たせないため（GAS側が既に認証を持っている）。
  //  - 朝レポートは発火しない（過去日のサマリーを投稿されると困る）。
  async function startBackfill(startStr, endStr) {
    if (isSyncing) return;
    isSyncing = true;
    syncBtn.disabled = true;
    const bfBtn = document.getElementById('dent-backfill-btn');
    if (bfBtn) bfBtn.disabled = true;
    logPanel.innerHTML = '';
    logPanel.style.display = 'block';

    const start = new Date(startStr + 'T00:00:00');
    const end = new Date(endStr + 'T00:00:00');
    // 予約・日次とも3ヶ月チャンク。両タイプを1チャンク内で取ってからD1へ流す。
    const chunks = getDateChunks(start, end, 3);
    log(`⏳ バックフィル開始: ${startStr} 〜 ${endStr}（${chunks.length}チャンク）`);
    log(`  ※ 定常同期(過去90日)とは別。朝レポートは発火しません。`);

    let okChunks = 0, totalDaily = 0, totalRes = 0;
    try {
      for (let ci = 0; ci < chunks.length; ci++) {
        const ch = chunks[ci];
        log(`━━ [${ci + 1}/${chunks.length}] ${ch.start_date} 〜 ${ch.end_date} ━━`);

        // このチャンクの日次＋予約をAirhostから取得（内部でレート制限待機あり）
        const dailyRows = await fetchChunk('daily', ch);
        await new Promise(r => setTimeout(r, 35000));
        const resRows = await fetchChunk('reservation', ch);

        if (dailyRows.length <= 1 && resRows.length <= 1) {
          log(`  データなし（この期間はAirhostに無し）— スキップ`);
          continue;
        }

        // Sheetsを上書き（ensureAndWriteSheet はヘッダ込みで全置換）
        if (dailyRows.length > 1) {
          const r = await sendToBackground('writeToSheet', { sheetName: '日次データ', rows: dailyRows });
          if (!r.success) throw new Error(`日次書込失敗: ${r.error}`);
          totalDaily += dailyRows.length - 1;
        }
        if (resRows.length > 1) {
          const r = await sendToBackground('writeToSheet', { sheetName: '予約データ', rows: resRows });
          if (!r.success) throw new Error(`予約書込失敗: ${r.error}`);
          totalRes += resRows.length - 1;
        }
        log(`  Sheets書込: 日次${Math.max(0, dailyRows.length - 1)} / 予約${Math.max(0, resRows.length - 1)}`);

        // このチャンク分をD1へ流す（朝レポートは出さない専用action）
        const d1ok = await triggerD1SyncOnly();
        if (!d1ok) throw new Error('D1同期に失敗（このチャンクを中断）');
        log(`  D1取込: ✅`);
        okChunks++;

        if (ci < chunks.length - 1) {
          log(`  次チャンクまで35秒待機...`);
          await new Promise(r => setTimeout(r, 35000));
        }
      }
      log(`✅ バックフィル完了: ${okChunks}/${chunks.length}チャンク / 日次${totalDaily}行 予約${totalRes}行 をD1へ`);
    } catch (err) {
      log(`❌ 中断: ${err.message}`);
      log(`  ここまでのチャンクはD1に取込済み。日付を調整して再実行すれば続きから可能。`);
      console.error('[Dent Backfill]', err);
    } finally {
      // Sheetsに過去チャンクが残ったままだと、翌朝の定常同期の前提(直近窓バッファ)が
      // 崩れる。最後に直近90日で上書きし直して現状復帰する。D1は既に永続化済みなので
      // これで失われるものは無い。
      try {
        log(`🔄 Sheetsを直近窓に復帰中...`);
        await restoreRecentWindow();
        log(`  復帰完了`);
      } catch (e) {
        log(`  ⚠️ 復帰失敗: ${e.message} — 手動で「ワンクリック同期」を1回押してください`);
      }
      isSyncing = false;
      syncBtn.disabled = false;
      if (bfBtn) bfBtn.disabled = false;
    }
  }

  // 定常同期と同じ直近窓でSheetsを上書きし直す（D1へは流さない＝朝レポートも出さない）
  async function restoreRecentWindow() {
    const today = new Date();
    const windowStart = new Date(today);
    windowStart.setDate(windowStart.getDate() - 90);
    const oneYearLater = new Date(today);
    oneYearLater.setFullYear(oneYearLater.getFullYear() + 1);

    const dailyRows = await fetchAllChunks('daily', windowStart, today);
    if (dailyRows.length > 1) {
      await sendToBackground('writeToSheet', { sheetName: '日次データ', rows: dailyRows });
    }
    await new Promise(r => setTimeout(r, 35000));
    const resRows = await fetchAllChunks('reservation', windowStart, oneYearLater);
    if (resRows.length > 1) {
      await sendToBackground('writeToSheet', { sheetName: '予約データ', rows: resRows });
    }
  }

  // 1チャンク分を取得（fetchAllChunks の単一チャンク版。全期間を溜めない）
  async function fetchChunk(type, chunk) {
    const csvText = await sendToPage('exportCSV', { type, dateRange: chunk });
    const rows = parseCSV(csvText.replace(/^﻿/, ''));
    return rows;
  }

  async function triggerD1SyncOnly() {
    try {
      const stored = await chrome.storage.local.get(['morningReportUrl']);
      const baseUrl = stored.morningReportUrl;
      if (!baseUrl) { log('  D1取込: ❌ URL未設定（朝レポートURLと同じものを設定画面で）'); return false; }
      const url = baseUrl + (baseUrl.includes('?') ? '&' : '?') +
        'action=d1SyncOnly&token=' + encodeURIComponent(MORNING_REPORT_TOKEN);
      const result = await sendToBackground('getUrl', { url });
      if (result.success && result.data && result.data.ok) return true;
      log('  D1取込: ❌ ' + JSON.stringify((result.data && result.data.error) || result.error || result).slice(0, 200));
      return false;
    } catch (e) {
      log('  D1取込: ❌ ' + e.message);
      return false;
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

  // 新法ステータスの計算・通知は GAS (gas/shinpou-report.gs) に移管済み (2026-07-27)。
  // 旧実装はスプシ「物件マスタ」+ 過去90日の予約しか見ておらず、YAML移管後の新規物件
  // (SKA等) が通知から消え、年度累計も過小になっていた。詳細は shinpou-report.gs 冒頭。


  // ======== 朝の運営サマリー発火 ========

  // GAS morning-report.gs の MORNING_REPORT_CONFIG.TRIGGER_TOKEN と一致させること
  const MORNING_REPORT_TOKEN = 'dent_morning_2026';

  // GAS の doGet(?action=morningReport) は D1同期 → 新法通知 → 朝レポート を実行する。
  // 戻り値の shinpou を呼び出し側に返して、同期完了通知に結果を載せる。
  async function triggerMorningReport() {
    try {
      const stored = await chrome.storage.local.get(['morningReportUrl']);
      const baseUrl = stored.morningReportUrl;
      if (!baseUrl) {
        log('  朝レポート発火: URL未設定（設定画面から設定してください）');
        return { shinpou: { error: 'GAS URL未設定' } };
      }
      const url = baseUrl +
        (baseUrl.indexOf('?') >= 0 ? '&' : '?') +
        'action=morningReport&token=' + encodeURIComponent(MORNING_REPORT_TOKEN);
      const result = await sendToBackground('getUrl', { url });
      // ok=false（朝レポート失敗）でも新法通知は独立して成功していることがあるので
      // shinpou は ok に関わらず拾う。
      const shinpou = (result.data && result.data.shinpou) || null;
      if (shinpou) {
        if (shinpou.posted) log('  新法通知: ✅ 投稿完了（' + shinpou.count + '件）');
        else if (shinpou.skipped === 'already-posted') log('  新法通知: ⏭ 本日投稿済み');
        else if (shinpou.skipped === 'no-target') log('  新法通知: ⏭ 対象物件なし');
        else if (shinpou.error) log('  新法通知: ❌ ' + shinpou.error);
      }
      if (result.success && result.data && result.data.ok) {
        const r = result.data.result || {};
        if (r.posted) log('  朝レポート発火: ✅ 投稿完了');
        else if (r.skipped === 'already-posted') log('  朝レポート発火: ⏭ 本日投稿済み');
        else if (r.skipped === 'not-synced') log('  朝レポート発火: ⏭ 同期未完了扱い');
        else log('  朝レポート発火: 完了');
      } else {
        const msg = (result.data && result.data.error) || result.error || result.raw || '不明なエラー';
        log('  朝レポート発火: ❌ ' + msg);
      }
      return { shinpou: shinpou };
    } catch (e) {
      log('  朝レポート発火: ❌ ' + e.message);
      return { shinpou: { error: e.message } };
    }
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
        // 新法通知の結果を追記（投稿はGAS shinpou-report.gs 側）
        const sn = data.shinpouNotifyStatus;
        if (sn) {
          if (sn.posted) text += `\n📋 新法通知: ✅ 送信完了（${sn.count}件）`;
          else if (sn.skipped) text += `\n📋 新法通知: ⏭ スキップ（${sn.skipped}）`;
          else if (sn.error) text += `\n📋 新法通知: ❌ ${sn.error}`;
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
