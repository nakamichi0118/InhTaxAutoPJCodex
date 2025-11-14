# InhTaxAutoPJ CSV Exporter

## Components
- `src/export_csv.py` - CLI utility that converts normalised JSON into `assets.csv` / `bank_transactions.csv`.
- `Docs/CSV_SPEC.md` - specification of the CSV schema and the expected JSON payload.
- `examples/sample_assets.json` - sample JSON matching the spec.
- `backend/app` - FastAPI backend (Gemini-powered document analysis + CSV export API).
- `backend/scripts/analyze_pdf.py` - CLI helper to run the Gemini layout flow against local PDFs.
- `webapp/index.html`
 - Static Web UI that talks to the deployed API and downloads CSVs.

## CLI usage
```bash
python src/export_csv.py examples/sample_assets.json --output-dir dist --force
```

## Backend API (local)
```bash
pip install -r backend/requirements.txt
uvicorn backend.app.main:app --reload
```
Endpoints:
- `GET /api/ping`
- `POST /api/analyze/pdf`
- `POST /api/export`

環境変数は `.env` を活用してください。ローカルで PDF を確認する場合:
```bash
python backend/scripts/analyze_pdf.py test/1号/touki_tate1.pdf
```

## Web フロー
1. `webapp/index.html` をブラウザで起動（Cloudflare Pages などでのホスティング想定）。
2. API エンドポイントは Railway などの URL（例: `https://inhtaxautopjcodex-production.up.railway.app/api`）を指定。
3. JSON をアップロードして「バックエンドでCSV生成」を実行すると、API 経由で CSV が生成・ダウンロードできます。

備考:
- API パラメータはクエリ `?api=` で切り替え可能。
- CSV は UTF-8 BOM 付きで出力され、Excel で文字化けしません。
- `webapp/index.html` の UI ではサンプル JSON を読み込みテストできます。


Large PDF uploads are automatically split before they hit Gemini. Control chunking with `GEMINI_DOCUMENT_MAX_MB` and per-chunk page count via `GEMINI_CHUNK_PAGE_LIMIT`.
Gemini-based analysis requires `GEMINI_API_KEY`. Override the model with `GEMINI_MODEL` (default `gemini-1.5-flash-latest`).
Gemini handles oversized PDFs by uploading them through the Files API automatically; no manual preprocessing is needed.
