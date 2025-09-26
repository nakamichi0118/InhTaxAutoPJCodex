# InhTaxAutoPJ CSV Exporter

## Components
- src/export_csv.py - command line exporter that reads normalised JSON and writes ssets.csv plus optional ank_transactions.csv.
- Docs/CSV_SPEC.md - CSV schema and the intermediate JSON contract.
- examples/sample_assets.json - fixture matching the expected JSON payload.
- webapp/index.html - static web demo for browser based conversion.
- ackend/app - FastAPI backend with PDF→Azure Document Intelligence integration and CSV export API.
- ackend/scripts/analyze_pdf.py - CLI helper to run the Azure layout model against local PDFs.
- Docs/ - background documentation for the broader system.

## CLI usage
`ash
python src/export_csv.py examples/sample_assets.json --output-dir dist --force
`
This command writes dist/assets.csv and dist/bank_transactions.csv. The script accepts either a single JSON file or a directory containing multiple JSON files.

### Input contract
The input JSON must expose an ssets array. Each entry follows the fields described in Docs/CSV_SPEC.md. Unknown properties are ignored during export so the schema can grow without breaking the tool.

### Options
- --output-dir (default ./output): destination directory for the CSV files.
- --force: overwrite existing CSV files in the destination folder.

## Backend API
`ash
pip install -r backend/requirements.txt
uvicorn backend.app.main:app --reload
`
Endpoints:
- GET /api/ping – health check.
- POST /api/analyze/pdf – upload a PDF, Azure Document Intelligence (prebuilt-layout) extracts text lines per page.
- POST /api/export – submit the normalised JSON payload and receive base64 encoded CSV strings.

Environment variables are loaded from .env (Azure/Gemini keys). A helper CLI can test PDFs locally without running the server:
`ash
python backend/scripts/analyze_pdf.py test/1組/touki_tate1.pdf --out tmp.json
`

## Web demo
1. Open webapp/index.html in a browser (double click or serve the folder with python -m http.server).
2. Paste the normalised JSON into the textarea, or click the sample button.
3. Press the generate button to preview and download the CSV files.

Notes:
- CSV files are encoded in UTF-8 with BOM for Excel compatibility.
- Record IDs are deterministic when the source document, asset name, and identifiers stay the same.
- Date normalisation covers ISO strings, YYYYMMDD, YYYY-MM-DD, Japanese style YYYY年M月D日, and era notation (Reiwa, Heisei, Showa, Taisho).
