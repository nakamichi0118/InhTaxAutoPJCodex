# InhTaxAutoPJ CSV Exporter

## Components
- `src/export_csv.py` - CLI utility that converts normalised JSON into `assets.csv` / `bank_transactions.csv`.
- `Docs/CSV_SPEC.md` - specification of the CSV schema and the expected JSON payload.
- `examples/sample_assets.json` - sample JSON matching the spec.
- `backend/app` - FastAPI backend (Gemini-powered document analysis + CSV export API).
- `backend/scripts/analyze_pdf.py` - CLI helper to run the Gemini layout flow against local PDFs.
- `webapp/index.html`
 - Static Web UI that talks to the deployed API and downloads CSVs.
- `ledger_frontend/`
 - React + Vite 入出金検討表ツール。`npm run build` で成果物を `webapp/ledger/` に書き出し、Railway 上の FastAPI (`/api/ledger` 系) と連携して勘定科目のCRUDやインポート/エクスポートを行います。
 - OCRジョブが完了すると `webapp/index.html` が通帳データをブラウザに退避し、Ledger 画面に「未登録の口座候補」が表示されます。案件を選んで「取り込み」を押すと一括で口座/取引が登録されます。
 - Ledger画面のURLは常に `/ledger/` 固定で、案件切替・新規作成は画面上部のドロップダウン＋ボタンで行います。Railway APIの公開URLは `window.__ledger_api_base` で上書きできます。
- `docs/USAGE.md`
 - How-to guide (usage + FAQ).

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

ϐ `.env` pĂB[J PDF mFꍇ:
```bash
python backend/scripts/analyze_pdf.py test/1/touki_tate1.pdf
```

## Web t[
1. `webapp/index.html` uEUŋNiCloudflare Pages Ȃǂł̃zXeBOzjB
2. API Gh|Cg Railway Ȃǂ URLi: `https://inhtaxautopjcodex-production.up.railway.app/api`jwB
3. JSON Abv[hāuobNGhCSVvsƁAAPI oR CSV E_E[hł܂B

l:
- API p[^̓NG `?api=` Ő؂ւ\B
- CSV  UTF-8 BOM tŏo͂AExcel ŕ܂B
- `webapp/index.html`  UI ł̓Tv JSON ǂݍ݃eXgł܂B


Large PDF uploads are automatically split before they hit Gemini. Control chunking with `GEMINI_DOCUMENT_MAX_MB` and per-chunk page count via `GEMINI_CHUNK_PAGE_LIMIT`.
Gemini-based analysis requires `GEMINI_API_KEY`. Override the model with `GEMINI_MODEL` (default `gemini-1.5-flash-latest`).
Gemini handles oversized PDFs by uploading them through the Files API automatically; no manual preprocessing is needed.

## Ledger API

The `/api/ledger/*` endpoints power the React 入出金検討表ツール. Data is persisted to a lightweight SQLite database whose location defaults to `data/ledger.db`. Override it with the `LEDGER_DB_PATH` environment variable when deploying (e.g. mounted volume or managed database path).
