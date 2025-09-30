from __future__ import annotations

import base64
import logging
from typing import Any, Callable, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from .config import get_settings
from .exporter import export_to_csv_strings
from .gemini import GeminiClient, GeminiError
from .models import DocumentAnalyzeResponse, DocumentType
from .ocr import AzureFormRecognizerClient, AzureFormRecognizerError
from .parser import build_assets, detect_document_type
from .pdf_utils import PdfChunkingError, PdfChunkingPlan, chunk_pdf_by_limits

logger = logging.getLogger(__name__)

app = FastAPI(title="InhTaxAutoPJ Backend", version="0.3.0")

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


def _with_pdf_chunks(
    payload: bytes,
    plan: PdfChunkingPlan,
    analyzer: Callable[[bytes], List[str]],
) -> List[str]:
    if len(payload) <= plan.max_bytes:
        return analyzer(payload)
    chunks = chunk_pdf_by_limits(payload, plan)
    lines: List[str] = []
    for chunk in chunks:
        lines.extend(analyzer(chunk))
    return lines


def _analyze_with_gemini(contents: bytes, content_type: str, settings) -> List[str]:
    if content_type != "application/pdf":
        raise GeminiError("Gemini analysis currently supports PDF documents only")
    if not settings.gemini_api_key:
        raise GeminiError("Gemini API key not configured")

    client = GeminiClient(api_key=settings.gemini_api_key, model=settings.gemini_model)
    plan = PdfChunkingPlan(
        max_bytes=settings.gemini_max_document_bytes,
        max_pages=settings.gemini_chunk_page_limit,
    )

    def analyzer(blob: bytes) -> List[str]:
        return client.extract_lines_from_pdf(blob)

    return _with_pdf_chunks(contents, plan, analyzer)


def _analyze_with_azure(contents: bytes, content_type: str, settings) -> List[str]:
    client = AzureFormRecognizerClient()

    def analyzer(blob: bytes) -> List[str]:
        raw_result = client.analyze_layout(blob, content_type=content_type)
        pages = extract_layout_pages(raw_result)
        return [line for page in pages for line in page["lines"]]

    plan = PdfChunkingPlan(
        max_bytes=settings.azure_max_document_bytes,
        max_pages=settings.azure_chunk_page_limit,
    )

    try:
        if content_type == "application/pdf":
            try:
                return _with_pdf_chunks(contents, plan, analyzer)
            except AzureFormRecognizerError as exc:
                if "InvalidContentLength" not in str(exc):
                    raise
                logger.warning("Azure reported InvalidContentLength; retrying with smaller chunks")
                chunks = chunk_pdf_by_limits(contents, plan)
                lines: List[str] = []
                for chunk in chunks:
                    lines.extend(analyzer(chunk))
                return lines
        return analyzer(contents)
    except PdfChunkingError as exc:
        logger.error("PDF chunking failed: %s", exc)
        raise HTTPException(status_code=413, detail=str(exc)) from exc
    except AzureFormRecognizerError as exc:
        logger.exception("Azure Document Intelligence error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc)) from exc


def _analyze_layout(contents: bytes, content_type: str) -> List[str]:
    settings = get_settings()

    if settings.gemini_api_key:
        try:
            return _analyze_with_gemini(contents, content_type, settings)
        except (GeminiError, PdfChunkingError) as exc:
            logger.warning("Gemini analysis failed, falling back to Azure: %s", exc)

    return _analyze_with_azure(contents, content_type, settings)


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
    if settings.gemini_api_key:
        logger.info("Gemini model: %s", settings.gemini_model)
