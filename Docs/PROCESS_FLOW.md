# Document Processing Flow (Draft)

## 1. Frontend UX Outline
- **Step 1: 書類アップロード**
  - ユーザーは PDF / 画像をドラッグ＆ドロップまたはファイル選択で投入。
  - 書類種別を自動推定（断定できない場合はユーザー選択）。
- **Step 2: OCR/LLM 呼び出し**
  - バックエンド `/api/documents/analyze` にアップロードし、Azure Document Intelligence → Gemini (必要に応じて) を実行。
- **Step 3: 取得データのレビュー**
  - 取引明細・評価額などをテーブル表示。
  - 書類ごとに区分/所有者/メモ等を編集可能。
- **Step 4: CSV／PDF エクスポート**
  - 正規化 JSON を保存しつつ、`/api/export` で CSV を生成。
  - リネーム済み PDF をダウンロード可能にする予定。

## 2. Backend Interaction
- `POST /api/documents/analyze`
  - Body: multipart/form-data (`file`, `document_type` optional)。
  - Response: `{ status, document_type, raw_layout, normalized_assets[] }`。
- `POST /api/documents/commit`
  - Body: `{ assets: [...], source_files: [...] }` save normalized payload, return CSV (delegate to existing exporter).
- `GET /api/ping`
  - 既存のヘルスチェック。

## 3. Data Mapping (最初の対応範囲)
- **通帳 (bank_deposit)**
  - レイアウトから支店名/口座番号/残高欄を抽出。
  - Gemini で取引明細を JSON 化（既存 GAS ロジックを Python に移植）。
- **固定資産税通知書 (land/building)**
  - Azure レイアウトのキー-バリューを正規表現で抽出。
  - 所在地・地目・評価額を `asset` レコードにマッピング。

## 4. 次の開発タスク
1. バックエンドに `parser` モジュールを追加し、レイアウト→正規化 JSON 変換を実装。
2. `/api/documents/analyze` エンドポイントで PDF を受け取り、変換結果を返却。
3. フロントエンドに書類アップローダーとレビュー UI を追加。
4. 正規化結果を既存 `/api/export` に渡して CSV をダウンロード。
