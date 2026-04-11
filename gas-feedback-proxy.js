/**
 * Google Apps Script — Feedback → Slack Proxy
 *
 * セットアップ手順:
 * 1. https://script.google.com で新しいプロジェクトを作成
 * 2. このコードを貼り付け
 * 3. SLACK_WEBHOOK_URL を自分のSlack Webhook URLに変更
 * 4. デプロイ → 新しいデプロイ → ウェブアプリ
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイURLをコピーして app.js の GAS_FEEDBACK_URL に設定
 */

const SLACK_WEBHOOK_URL = 'YOUR_SLACK_WEBHOOK_URL_HERE';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const payload = {
      text: data.text || 'No message',
    };

    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
