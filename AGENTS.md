# Repository Guidelines

## Project Structure & Module Organization
The CLI exporter lives in `src/export_csv.py` and turns normalized JSON into `dist/assets.csv` and `dist/bank_transactions.csv`. FastAPI endpoints and Gemini-driven orchestration sit in `backend/app` (notably `main.py`, `parser.py`, `exporter.py`), while shared scripts reside under `backend/scripts`. Specifications are in `Docs/`, canonical payloads in `examples/`, web assets in `webapp/`, and sample PDFs for manual validation in `test/`. Generated CSVs and other artifacts belong in `dist/`.

## Build, Test, and Development Commands
Create an isolated environment with `python -m venv .venv` and activate it before running `pip install -r requirements.txt`. Regenerate CSV fixtures via `python src/export_csv.py examples/sample_assets.json --output-dir dist --force`, then inspect the new files in `dist/`. Start the API locally with `uvicorn backend.app.main:app --reload` and verify readiness at `GET /api/ping`. To inspect Gemini layout analysis, run `python backend/scripts/analyze_pdf.py "test/1/touki_tate1.pdf"`.

## Coding Style & Naming Conventions
Target Python 3.11+, use four-space indentation, and prefer explicit type hints. Follow `snake_case` for functions, variables, and modules; reserve `PascalCase` for Pydantic models and dataclasses. Keep CSV formatting logic inside `src/export_csv.py`, parser changes in `backend/app/parser.py`, and environment-specific settings in `backend/app/config.py`. Add succinct comments only when logic is non-obvious.

## Testing Guidelines
No automated suite exists yet, so rely on fixtures inside `examples/` and PDFs under `test/`. After parser or exporter changes, regenerate CSVs into `dist/` and confirm headers, row counts, and BOM preservation. For API updates, post a sample payload to `/api/export`, verify the response writes valid CSV files, and attach representative snippets to reviews.

## Commit & Pull Request Guidelines
Write imperative, present-tense commit messages (e.g., `Adjust bankbook parser for multi-page statements`). Pull requests should explain motivation, summarize behavioral changes, reference related tickets, and list manual checks performed. Include CSV diffs or API response samples whenever behavior shifts, exclude `.env` or credentials, and clean temporary files before pushing.

## Security & Document Intelligence Tips
Store secrets in environment variables or secret managers compatible with your deployment target. Override Gemini behavior with `GEMINI_MODEL`, `GEMINI_DOCUMENT_MAX_MB`, and `GEMINI_CHUNK_PAGE_LIMIT`; a 413 response indicates the PDF still exceeds the configured chunk size. Ensure `GEMINI_API_KEY` stays valid, rotate it when needed, and verify rate limits before large batches.

## Progress Log

### 2025-11-17
- READMEとDocs/CSV_SPEC.mdを精読し、CLIエクスポーター(`src/export_csv.py`)が正規化済みJSONを`assets.csv`と`bank_transactions.csv`へ整形する中心ロジックである点を把握。UTF-8 BOM付き出力やUUIDベースの`record_id`生成、取引IDの決定論的生成を確認。
- FastAPIバックエンド(`backend/app`)のレイヤー構成を調査。`main.py`でCORS許可済みAPIを公開し、PDFアップロードをGemini/Azureに渡す処理フロー、PdfChunk計画による分割制御、Geminiフォールバック戦略、Bankbook解析後の`exporter.py`経由CSV生成を理解。
- `parser.py`と`azure_analyzer.py`で銀行通帳OCR行から口座情報・取引を抽出し、`models.py`がPydanticでJSONペイロードを定義している点を整理。`job_manager.py`がバックグラウンド処理+一時ファイル管理と進捗追跡を担当することを把握。
- `webapp/index.html`でのシングルページUIがRailway上のAPIにPOSTしCSVダウンロードをトリガー、`backend/scripts/analyze_pdf.py`でGemini解析を単体で確認できることを理解。環境変数の設定(`backend/app/config.py`)とGemini APIキーのローテーションロジックも確認済み。
- Azure生データの確認を容易にするため、ジョブ完了時に`azure_raw_transactions.csv`を新たに返却するよう`backend/app/main.py`を拡張。ページ番号と行番号付きでAzureが抽出した取引そのままをCSV化し、Web UIのダウンロード一覧から取得できるようにした。
- Gemini単体モードをジョブAPI・Web UIに追加し、`processing_mode=gemini`と`gemini-2.5-flash/pro`の切替をサポート。UIで解析エンジンを選ぶとFastAPI側がGeminiのみでCSVを生成し、Azureを経由せずモデル別の比較が可能になった。
- Gemini単体実行時に1ページずつ解析+処理進捗を返すよう改善し、タイムアウト後の再解析による遅延を緩和。`backend/app/main.py`でGeminiチャンク進捗をジョブステータスに反映し、検証時の安定性を高めた。
- Gemini 2.5 Proを標準解析フローに設定し、UIの解析エンジン選択もPro単体が初期値となるよう更新。残高と入出金の符号チェックを追加してGemini出力の逆転誤りを自動補正するよう`backend/app/main.py`を拡張。
- 入出金逆転の再発に備え、摘要キーワード（振込資金/手数料など）と残高推移の両方から補正を行うよう`_enforce_continuity`を拡張。Web UIには読み取り完了秒数を表示するサマリを追加し、検証時の速度比較が容易になった。
