# dent-data-api: `/internal/reviews.json` 仕様

レビュー実データを Cloudflare Worker 経由でダッシュボードに配信するエンドポイント。
operation 側 (`app.js`) は `fetchDataApi('/internal/reviews.json')` で呼ぶ前提。

## ソース

- スプシ ID: `1uBghVbJw50HN3C3ehhtYKwNHPRCHhOwvJZxy_U83HwE`
- タイトル: `test2026年度LookerStudio用売上データ`
- タブ名: 「レビュー」（gid=491525933）
- 認証: 既存の Sheets API キー（シートはリンク共有で閲覧可能に設定済み）

### 元シートのカラム

```
チャンネル名 | 施設名 | 予約番号 | 作成日 | 総合評価 | 清潔さ | コミュニケーション
| チェックイン | 正確さ | ロケーション | コスパ | 公開レビュー | プライベートフィードバック | 応答 | レビューア種別
```

`チャンネル名` は `"Airbnb - NPA"` 形式。末尾の識別子（NPA/OPH/PHM等）が物件マスタ `airbnbアカウント` と対応。

## エンドポイント

```
GET https://api.dent-inc.com/internal/reviews.json
```

- 認証: 既存の Cloudflare Access (Google account)
- CORS: 既存の `op.dent-inc.com` origin 許可と同じ
- キャッシュ: 数分 (KV or in-memory)。レビューは即時性不要

## レスポンス形状

```json
{
  "fetchedAt": "2026-05-13T12:34:56Z",
  "reviews": [
    {
      "channel": "Airbnb",
      "account": "NPA",
      "listing": "8 min to downtown, 2 min to the sta.【Free-Wifi】",
      "reservationId": "HM2TEJZ43K",
      "date": "2026-03-06",
      "stars": {
        "overall": 5,
        "cleanliness": 4,
        "communication": 5,
        "checkin": 5,
        "accuracy": 5,
        "location": 5,
        "value": 5
      },
      "publicReview": "鄰近車站旁，走路只需要2分鐘即可到達...",
      "privateFeedback": "最裡面的燈壞掉了",
      "response": "",
      "reviewerType": "guest"
    }
  ]
}
```

### 正規化ルール

- `チャンネル名` を `"Airbnb - NPA"` → `{ channel: "Airbnb", account: "NPA" }` に分解。`" - "` で split、末尾を account に。
- 各 ★ カラムは数値化。空欄は `null` ではなく **行ごと除外** （無評価行はノイズ）
- `作成日` は `YYYY-MM-DD` ISO 形式に正規化
- 行は `作成日` 降順で返す
- `レビューア種別` が `guest` 以外（ホスト→ゲスト等）の行はオプションで含める（クエリ `?type=guest` でフィルタ可、デフォルト全件）

## エラーハンドリング

- 認証失敗: 既存パターン同様 401/302（`fetchDataApi` 側で `requiresAuth` フラグ付与）
- Sheets API レート制限: 30s リトライ × 3 回
- シート構造変化（カラム名変更）: 5xx + ログに警告

## 実装メモ

- 既存の `fetchEstatSheets` / `fetchAirdnaSheets` (app.js:248-) と同じく `sheets.googleapis.com/v4/spreadsheets/{ID}/values/{range}` を叩く実装で十分
- レンジは `'レビュー'!A:O` 等で指定
- スプシのカラム順が変わる可能性に備え、ヘッダー行を読んで動的にマッピング

## TODO（運用フェーズ）

- スプシの更新トリガ（手動/GASスケジュール/外部スクレイパー）を確定
- 「応答」列の更新を operation 側からも書き戻すか（Phase 4 で検討）
