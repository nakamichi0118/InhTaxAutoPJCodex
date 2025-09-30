# Repository Guidelines

## Project Structure & Module Organization
The CLI exporter `src/export_csv.py` turns normalised JSON into `dist/assets.csv` and `dist/bank_transactions.csv`. FastAPI code lives in `backend/app` (`main.py`, `parser.py`, `exporter.py`), while shared Azure helpers sit in `backend/scripts`. Specifications are kept under `Docs/`, canonical payloads under `examples/`, static assets in `webapp/`, and sample PDFs inside `test/`.

## Build, Test, and Development Commands
Create a virtual environment with `python -m venv .venv` and install dependencies via `pip install -r requirements.txt`. Validate CSV generation by running `python src/export_csv.py examples/sample_assets.json --output-dir dist --force`. Start the API locally with `uvicorn backend.app.main:app --reload` and confirm readiness at `GET /api/ping`. To inspect Azure layout output, execute `python backend/scripts/analyze_pdf.py "test/1号/touki_tate1.pdf" --out tmp.json`.

## Coding Style & Naming Conventions
Target Python 3.11+, use four-space indentation, and follow `snake_case` for functions, variables, and modules. Reserve `PascalCase` for Pydantic models. Keep CSV formatting logic inside `export_csv.py`, parser improvements in `backend/app/parser.py`, and load environment-specific values from `backend/app/config.py`. Prefer explicit type hints and dataclasses where they clarify intent.

## Testing Guidelines
No automated suite exists yet. Use JSON fixtures in `examples/` and PDFs in `test/` for manual verification. After parser or exporter changes, regenerate CSVs into `dist/` and check headers, row counts, and BOM encoding. For API updates, post the sample payload to `/api/export` and ensure the response writes to valid CSV files.

## Commit & Pull Request Guidelines
Write imperative, present-tense commits (e.g., `Adjust bankbook parser for multi-page statements`). Pull requests should call out motivation, summarize behaviour changes, list manual or automated checks, and reference related tickets. Include representative CSV or API snippets when behaviour changes, and never commit `.env` or credential material.

## Security & Configuration Tips
Store secrets in environment variables or Azure Key Vault instead of source control. Point the web client at alternate endpoints via the `?api=` query parameter during testing. Keep generated artifacts in `dist/` and clean temporary files before pushing. Always push your branch once assigned work is complete.
## Document Intelligence Handling
Large PDFs are automatically split before hitting Azure. Tweak the thresholds via `AZURE_DOCUMENT_MAX_MB` (default 4 MB) and `AZURE_CHUNK_PAGE_LIMIT` (default 20 pages). If Azure still rejects a single page, the API returns HTTP 413 so we can consider compressing or routing to Gemini.
