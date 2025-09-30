from __future__ import annotations

import base64
import logging
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from .config import get_settings
from .exporter import export_to_csv_strings
from .models import DocumentAnalyzeResponse, DocumentType
from .ocr import AzureFormRecognizerClient, AzureFormRecognizerError
from .parser import build_assets, detect_document_type
from .pdf_utils import PdfChunkingError, PdfChunkingPlan, chunk_pdf_by_limits

logger = logging.getLogger(__name__)

app = FastAPI(title="InhTaxAutoPJ Backend", version="0.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


async def _load_file_bytes(file: UploadFile) -> tuple[bytes, str]:
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Empty file uploaded")
    content_type = file.content_type or "application/pdf"
    if content_type not in {"application/pdf"}:
        logger.warning("Unexpected content type %s; defaulting to application/pdf", content_type)
        content_type = "application/pdf"
    return contents, content_type


def _analyze_layout(contents: bytes, content_type: str) -> List[str]:
    client = AzureFormRecognizerClient()

    def run_single(payload: bytes) -> List[str]:
        raw_result = client.analyze_layout(payload, content_type=content_type)
        pages = extract_layout_pages(raw_result)
        return [line for page in pages for line in page["lines"]]

    if content_type == "application/pdf":
        plan = PdfChunkingPlan(
            max_bytes=client.max_upload_bytes,
            max_pages=client.chunk_page_limit,
        )
        try:
            if len(contents) <= client.max_upload_bytes:
                try:
                    return run_single(contents)
                except AzureFormRecognizerError as exc:
                    if "InvalidContentLength" not in str(exc):
                        raise
                    logger.warning(
                        "Azure rejected PDF due to size; falling back to chunked upload: %s",
                        exc,
                    )
                    chunks = chunk_pdf_by_limits(contents, plan)
                    return [line for chunk in chunks for line in run_single(chunk)]

            chunks = chunk_pdf_by_limits(contents, plan)
            results: List[str] = []
            for chunk in chunks:
                results.extend(run_single(chunk))
            return results
        except PdfChunkingError as exc:
            logger.error("PDF chunking failed: %s", exc)
            raise HTTPException(status_code=413, detail=str(exc)) from exc
        except AzureFormRecognizerError as exc:
            logger.exception("Azure Document Intelligence error: %s", exc)
            raise HTTPException(status_code=502, detail=str(exc)) from exc

    try:
        return run_single(contents)
    except AzureFormRecognizerError as exc:
        logger.exception("Azure Document Intelligence error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc)) from exc


@app.get("/api/ping")
def ping() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/api/export")
async def export_csv(payload: Dict[str, Any]) -> Dict[str, Any]:
    try:
        csv_map = export_to_csv_strings(payload)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    encoded = {
        name: base64.b64encode(content.encode("utf-8-sig")).decode("ascii")
        for name, content in csv_map.items()
    }
    return {"status": "ok", "files": encoded}


@app.post("/api/analyze/pdf")
async def analyze_pdf(file: UploadFile = File(...)) -> Dict[str, Any]:
    contents, content_type = await _load_file_bytes(file)
    lines = _analyze_layout(contents, content_type)
    return {
        "status": "ok",
        "line_count": len(lines),
        "lines": lines,
    }


@app.post("/api/documents/analyze", response_model=DocumentAnalyzeResponse)
async def analyze_document(
    file: UploadFile = File(...),
    document_type: Optional[DocumentType] = Form(None),
) -> DocumentAnalyzeResponse:
    contents, content_type = await _load_file_bytes(file)
    lines = _analyze_layout(contents, content_type)
    detected_type = document_type or detect_document_type(lines)
    assets = build_assets(detected_type, lines, source_name=file.filename or "uploaded.pdf")
    return DocumentAnalyzeResponse(
        status="ok",
        document_type=detected_type,
        raw_lines=lines,
        assets=assets,
    )


def extract_layout_pages(result: Dict[str, Any]) -> list[Dict[str, Any]]:
    analyze = result.get("analyzeResult") or result
    pages = []
    for page in analyze.get("pages", []):
        lines = [line.get("content", "") for line in page.get("lines", [])]
        pages.append({
            "page_number": page.get("pageNumber"),
            "unit": page.get("unit"),
            "width": page.get("width"),
            "height": page.get("height"),
            "lines": lines,
        })
    if not pages and "content" in analyze:
        pages.append({"page_number": 1, "lines": analyze.get("content", "").splitlines()})
    return pages


@app.exception_handler(AzureFormRecognizerError)
async def azure_error_handler(_, exc: AzureFormRecognizerError):
    logger.exception("Unhandled Azure error: %s", exc)
    return JSONResponse(status_code=502, content={"detail": str(exc)})


@app.on_event("startup")
def log_startup() -> None:
    settings = get_settings()
    logger.info("Azure endpoint: %s", settings.azure_endpoint)
