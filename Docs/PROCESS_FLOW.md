# Document Processing Flow (Draft)

## 1. Frontend UX Outline
- **Step 1: アップロード**
  - ユーザーが PDF / 画像をドラッグ＆ドロップまたはファイル選択する。
  - ページ順の確認や削除が必要な場合は、アップロード時に提示する。
- **Step 2: LLM 解析を実行**
  - バックエンド `/api/documents/analyze` にアップロードし、Gemini API を用いたレイアウト解析を呼び出す。
- **Step 3: 抽出データのレビュー**
  - 取引明細・残高などをテーブル表示で確認。
  - 必要に応じて項目名・摘要・補足コメントを編集可能にする。
- **Step 4: CSV / PDF エクスポート**
  - 正規化 JSON を保存し、`/api/export` で CSV を生成。
  - リネーム済み PDF をダウンロード可能にする。

## 2. Backend Interaction
- `POST /api/documents/analyze`
  - Body: multipart/form-data (`file`, `document_type` optional)。
  - Response: `{ status, document_type, raw_layout, normalized_assets[] }`。
- `POST /api/documents/commit`
  - Body: `{ assets: [...], source_files: [...] }` を保存し、CSV 出力を委譲。
- `GET /api/ping`
  - ヘルスチェック用のエンドポイント。

## 3. Data Mapping (初期対応範囲)
- **通帳 (bank_deposit)**
  - 金融機関・支店・口座番号・残高を抽出。
  - Gemini が返した JSON を Python で正規化。
- **土地・建物 (land/building)**
  - レイアウトから所在地や評価額を抽出し `asset` レコードにマッピング。

## 4. 今後のタスク
1. バックエンドの `parser` を拡張し、Gemini 出力を正規化 JSON へ変換。
2. `/api/documents/analyze` を通じて PDF → レイアウト変換を行い、UI に返す。
3. フロントエンドにアップローダーとレビュー UI を実装。
4. 正規化結果を `/api/export` に渡し、CSV を生成・ダウンロード可能にする。
