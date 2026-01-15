# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

InhTaxAutoPJ (SOROBOCR) is a Japanese inheritance tax document processing system. It extracts data from bank passbooks (通帳) and other financial documents via Gemini AI OCR, then exports structured CSVs for import into Excel-based inventory workbooks.

## Development Commands

```bash
# Backend setup
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements.txt

# Run API server locally
uvicorn backend.app.main:app --reload

# Test Gemini PDF analysis directly
python backend/scripts/analyze_pdf.py "test/1/touki_tate1.pdf"

# Generate CSVs from normalized JSON (CLI tool)
python src/export_csv.py examples/sample_assets.json --output-dir dist --force

# Run unit tests
pytest tests/

# Ledger frontend development
cd ledger_frontend
npm install
npm run dev      # Dev server
npm run build    # Build to webapp/ledger/
npm run lint     # ESLint check
```

## Architecture

### Core Processing Pipeline
1. **Frontend** (`webapp/index.html`) - Static SPA that uploads PDFs to the API
2. **FastAPI Backend** (`backend/app/main.py`) - Orchestrates document analysis via async job system
3. **Gemini Processing** (`backend/app/gemini.py`) - Page-by-page parallel OCR with `ThreadPoolExecutor(max_workers=4)`
4. **Transaction Correction** - Multi-stage pipeline: balance reconciliation, deposit/withdrawal direction fixing via `_enforce_continuity`, `_finalize_transaction_directions`
5. **CSV Export** (`backend/app/exporter.py`) - UTF-8 BOM output for Excel

### Key Backend Modules
- `main.py` - FastAPI routes, job lifecycle (`/api/analyze/pdf`, `/api/jobs/{id}`), correction pipelines
- `gemini.py` - Gemini API integration, chunk-based PDF processing, Files API for large PDFs
- `parser.py` - Document type detection, asset building from OCR lines
- `date_inference/engine.py` - Smart 2-digit year interpretation (handles 令和/平成/昭和 prefixes)
- `ledger_router.py` + `ledger_store.py` - SQLite-backed ledger CRUD API (`/api/ledger/*`)
- `job_manager.py` - Background job handling with temp file cleanup

### API Endpoints
- `GET /api/ping` - Health check
- `POST /api/analyze/pdf` - Upload PDF, returns job ID
- `GET /api/jobs/{id}` - Poll job status and progress
- `GET /api/jobs/{id}/result` - Fetch completed results (JSON or CSV via `?format=csv`)

### VBA Integration
`VBA/transaction_import.bas` imports JSON results directly into Excel via ADODB.Stream Base64 decoding. The macro calls the API, parses `bank_transactions.json`, and populates the 財産目録 workbook.

### Ledger Frontend
React + Vite app in `ledger_frontend/`. Builds to `webapp/ledger/` and communicates with `/api/ledger/*` endpoints.

## Environment Variables

Required:
- `GEMINI_API_KEY` - Google AI API key for document analysis

Optional:
- `GEMINI_MODEL` - Override model (default: `gemini-2.5-pro`)
- `GEMINI_DOCUMENT_MAX_MB` / `GEMINI_CHUNK_PAGE_LIMIT` - PDF chunking controls
- `LEDGER_DB_PATH` - SQLite database location (default: `data/ledger.db`)

## Testing

Run unit tests with `pytest tests/`. Manual validation:
1. Regenerate CSVs to `dist/` and compare headers/row counts
2. Test API endpoints via `POST /api/analyze/pdf`
3. Use sample PDFs in `test/` directory

## Code Conventions

- Python 3.11+, 4-space indentation, explicit type hints
- `snake_case` for functions/variables, `PascalCase` for Pydantic models
- CSVs output with UTF-8 BOM for Windows Excel compatibility
- All user-facing responses in Japanese

## Key Specifications

- `Docs/CSV_SPEC.md` - CSV schema and intermediate JSON format
- `Docs/USAGE.md` - User guide and FAQ
- `AGENTS.md` - Progress log and historical context
