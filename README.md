# InhTaxAutoPJ CSV Exporter

## Components
- `src/export_csv.py` - CLI utility that converts normalised JSON into `assets.csv` / `bank_transactions.csv`.
- `Docs/CSV_SPEC.md` - specification of the CSV schema and the expected JSON payload.
- `examples/sample_assets.json` - sample JSON matching the spec.
- `backend/app` - FastAPI backend (Azure Document Intelligence integration + CSV export API).
- `backend/scripts/analyze_pdf.py` - CLI helper to run the Azure layout model against local PDFs.
- `webapp/index.html` - Static Web UI that talks to the deployed API and downloads CSVs.

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

環境変数は `.env` を利用します。ローカルで PDF を試す場合:
```bash
python backend/scripts/analyze_pdf.py test/1組/touki_tate1.pdf --out tmp.json
```

## Web デモ
1. `webapp/index.html` をブラウザで開く（Cloudflare Pages でも同じ）。
2. API エンドポイントに Railway の URL（例: `https://inhtaxautopjcodex-production.up.railway.app/api`）を設定。
3. JSON を貼り付けて「バックエンドでCSV生成」を実行すると、API 経由で CSV を生成しダウンロードできます。

メモ:
- API パラメータはクエリ `?api=` でも差し替え可能。
- CSV は UTF-8 BOM 付きで出力され、Excel でも文字化けしません。
- `webapp/index.html` の UI からサンプル JSON を読み込んで動作確認できます。
