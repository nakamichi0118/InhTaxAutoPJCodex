# チャンクアップロード仕様

## 背景

VBA（`VBA/transaction_import.bas`）は `WinHttp.WinHttpRequest.5.1` を使用しており、
HTTP/1.1 のみ対応している。Cloud Run はリクエストボディを 32MB に制限しているため、
ブラウザ版（HTTP/2）では問題ないが、VBA から大きな PDF を送ると
`413 Request Entity Too Large` エラーになる。

この問題を回避するため、PDF を 20MB 単位に分割して複数リクエストで送信する
チャンクアップロード機構を実装している。

## API エンドポイント

### 1. セッション開始

```
POST /api/upload/init
```

レスポンス:
```json
{
  "upload_id": "abc123...",
  "max_chunk_bytes": 26214400
}
```

### 2. チャンク送信

```
POST /api/upload/{upload_id}/chunk
Content-Type: multipart/form-data

chunk_index: <int>   (0始まり)
chunk: <binary>      (application/octet-stream)
```

レスポンス:
```json
{
  "received": 0,
  "size": 20971520
}
```

### 3. ジョブ作成（全チャンク送信後）

```
POST /api/upload/{upload_id}/jobs
Content-Type: multipart/form-data

file_name: <string>           (必須)
document_type: <string>       (任意)
date_format: <string>         (任意: auto / western / wareki)
processing_mode: <string>     (任意: gemini)
gemini_model: <string>        (任意)
start_date: <YYYY-MM-DD>      (任意)
end_date: <YYYY-MM-DD>        (任意)
```

レスポンス: `/api/jobs` と同じ `JobCreateResponse` (HTTP 202)
```json
{
  "status": "accepted",
  "job_id": "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
}
```

以降のジョブ状態確認・結果取得は通常の `/api/jobs/{job_id}` エンドポイントを使用する。

## VBA 側の動作

`CreateAnalysisJob` の冒頭で `FileLen(pdfPath)` を確認し、
25MB（= 26,214,400 バイト）を超える場合は自動的に `CreateAnalysisJobChunked` に委譲する。
ユーザーには `Application.StatusBar` でチャンク進捗を表示する。

## Cloud Run デプロイ時の制約

**チャンクデータはサーバーのメモリ内 dict で管理している。**

Cloud Run でインスタンスが複数起動していると、init・chunk・finalize が
異なるインスタンスに振り分けられてデータが見つからず 404 エラーになる。

### 必須対応

以下のいずれかを設定すること:

**オプション A（推奨・シンプル）: インスタンス数を 1 に固定**

```bash
gcloud run deploy sorobocr-api \
  --max-instances=1 \
  ...
```

**オプション B: セッションアフィニティを有効化**

```bash
gcloud run deploy sorobocr-api \
  --session-affinity \
  ...
```

セッションアフィニティを使う場合、同一クライアントの後続リクエストが
同一インスタンスにルーティングされるため複数インスタンス運用が可能になる。
ただし Cloud Run のセッションアフィニティはベストエフォートであり、
インスタンス置き換え時（デプロイ、スケールイン等）にセッションが失われる場合がある。

## アップロードセッションの TTL

セッションは init から 30 分間有効。30 分以内に finalize されない場合は
次回リクエスト時に自動削除される。
