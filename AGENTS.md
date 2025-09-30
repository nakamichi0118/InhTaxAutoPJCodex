# Repository Guidelines

## Project Structure & Module Organization
Core CSV generation logic lives in src/export_csv.py, which accepts normalized JSON and emits ssets.csv and ank_transactions.csv. The FastAPI service is under ackend/app; main.py wires routing, parser.py handles Document Intelligence payloads, and exporter.py bridges to the CLI module. API helpers and Azure integration scripts sit in ackend/scripts. Specification assets reside in Docs/CSV_SPEC.md, while canonical payloads are stored in examples/. Web delivery is handled by the static client in webapp/, and sample source documents live in 	est/. Generated CSVs should go to dist/.

## Build, Test, and Development Commands
Install dependencies with python -m venv .venv followed by pip install -r requirements.txt. Convert sample data locally via python src/export_csv.py examples/sample_assets.json --output-dir dist --force. Launch the API during development with uvicorn backend.app.main:app --reload; confirm readiness at GET /api/ping. Use python backend/scripts/analyze_pdf.py test/1号/touki_tate1.pdf --out tmp.json when validating the Azure layout pipeline against local PDFs.

## Coding Style & Naming Conventions
Use Python 3.11 or newer with 4-space indentation, snake_case for functions and variables, and PascalCase only for pydantic models. Keep modules cohesive—CSV formatting logic stays in export_csv.py; OCR parsing enhancements belong in ackend/app/parser.py. Prefer explicit type hints and dataclasses where appropriate, and keep environment-dependent values inside config.py.

## Testing Guidelines
There is no dedicated test harness yet; rely on fixture JSON in examples/ and test PDFs under 	est/ to exercise new parsing rules. When adjusting the exporter, regenerate CSVs into dist/ and diff against expected headers and row counts. For API work, hit /api/export with the sample payload and ensure the response bytes produce valid CSV when written to disk. Document any manual verification steps in the pull request.

## Commit & Pull Request Guidelines
Follow the existing history by writing present-tense, imperative commit titles (e.g., "Adjust bankbook parser for multi-page statements"). Each pull request should describe the motivation, summarize functional changes, and list manual or automated checks performed. Link to tracking tickets where available, and attach before/after CSV snippets or API responses when behavior changes. Never commit .env or credential-bearing files.

## 作業ログ
- 2025-09-28 銀行通帳（きのくに）向けのOCR出力に対応するため、ackend/app/parser.pyを改修しました。
- 作業完了後はgit pushまで実施してください。
- 2025-09-30 通帳PDFの縦レイアウト（複数列が1行に集約されるケース）に対応できるようackend/app/parser.pyの行分割・日付復元ロジックを強化しました。
