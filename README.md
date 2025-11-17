# InhTaxAutoPJ CSV Exporter

## Components
- `src/export_csv.py` - CLI utility that converts normalised JSON into `assets.csv` / `bank_transactions.csv`.
- `Docs/CSV_SPEC.md` - specification of the CSV schema and the expected JSON payload.
- `examples/sample_assets.json` - sample JSON matching the spec.
- `backend/app` - FastAPI backend (Gemini-powered document analysis + CSV export API).
- `backend/scripts/analyze_pdf.py` - CLI helper to run the Gemini layout flow against local PDFs.
- `webapp/index.html`
 - Static Web UI that talks to the deployed API and downloads CSVs.
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
