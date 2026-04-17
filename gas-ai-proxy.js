/**
 * Google Apps Script — AI Chat Proxy (Claude API)
 *
 * セットアップ手順:
 * 1. https://script.google.com で新しいプロジェクトを作成
 * 2. このコードを貼り付け
 * 3. スクリプトプロパティに ANTHROPIC_API_KEY を設定
 *    (プロジェクトの設定 → スクリプトプロパティ → 行を追加)
 * 4. デプロイ → 新しいデプロイ → ウェブアプリ
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイURLをコピーして app.js の GAS_AI_URL に設定
 */

function doPost(e) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if (!apiKey) {
      return jsonResponse({ error: 'ANTHROPIC_API_KEY が未設定です' });
    }

    const req = JSON.parse(e.postData.contents);
    const messages = req.messages || [];
    const systemPrompt = req.system || '';

    const body = {
      model: 'claude-sonnet-4-20250514',
      max_tokens: 2048,
      system: systemPrompt,
      messages: messages,
    };

    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });

    const status = resp.getResponseCode();
    const result = JSON.parse(resp.getContentText());

    if (status !== 200) {
      return jsonResponse({ error: `Claude API error (${status}): ${result.error?.message || 'unknown'}` });
    }

    const text = result.content?.[0]?.text || '';
    return jsonResponse({ reply: text });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doGet() {
  return jsonResponse({ status: 'ok', service: 'ai-chat-proxy' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
