# Repository Guidelines

## Project Structure & Module Organization
The CLI exporter lives in `src/export_csv.py` and turns normalized JSON into `dist/assets.csv` and `dist/bank_transactions.csv`. FastAPI endpoints and Gemini-driven orchestration sit in `backend/app` (notably `main.py`, `parser.py`, `exporter.py`), while shared scripts reside under `backend/scripts`. Specifications are in `Docs/`, canonical payloads in `examples/`, web assets in `webapp/`, and sample PDFs for manual validation in `test/`. Generated CSVs and other artifacts belong in `dist/`.

## Build, Test, and Development Commands
Create an isolated environment with `python -m venv .venv` and activate it before running `pip install -r requirements.txt`. Regenerate CSV fixtures via `python src/export_csv.py examples/sample_assets.json --output-dir dist --force`, then inspect the new files in `dist/`. Start the API locally with `uvicorn backend.app.main:app --reload` and verify readiness at `GET /api/ping`. To inspect Gemini layout analysis, run `python backend/scripts/analyze_pdf.py "test/1��/touki_tate1.pdf"`.

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
