from __future__ import annotations

import base64
import logging
from typing import Any, Callable, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from .azure_analyzer import AzureAnalysisError, AzureAnalysisResult, AzureTransactionAnalyzer
from .config import get_settings
from .exporter import export_to_csv_strings
from .gemini import GeminiClient, GeminiError
from .models import AssetRecord, DocumentAnalyzeResponse, DocumentType
from .parser import build_assets, detect_document_type
from .pdf_utils import PdfChunkingError, PdfChunkingPlan, chunk_pdf_by_limits

logger = logging.getLogger(__name__)

app = FastAPI(title="InhTaxAutoPJ Backend", version="0.5.0")

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
    if content_type != "application/pdf":
        logger.warning("Unexpected content type %s; defaulting to application/pdf", content_type)
        content_type = "application/pdf"
    return contents, content_type


def _with_pdf_chunks(
    payload: bytes,
    plan: PdfChunkingPlan,
    analyzer: Callable[[bytes], List[str]],
) -> List[str]:
    current_plan = plan
    while True:
        chunks = chunk_pdf_by_limits(payload, current_plan)
        try:
            lines: List[str] = []
            for chunk in chunks:
                lines.extend(analyzer(chunk))
            return lines
        except GeminiError as exc:
            if current_plan.max_pages <= 1:
                raise
            logger.warning(
                "Gemini processing timed out for plan max_pages=%s; retrying with smaller chunks",
                current_plan.max_pages,
            )
            current_plan = PdfChunkingPlan(
                max_bytes=current_plan.max_bytes,
                max_pages=max(1, current_plan.max_pages // 2),
            )


def _analyze_with_gemini(contents: bytes, settings) -> List[str]:
    client = GeminiClient(api_key=settings.gemini_api_key, model=settings.gemini_model)
    plan = PdfChunkingPlan(
        max_bytes=settings.gemini_max_document_bytes,
        max_pages=settings.gemini_chunk_page_limit,
    )

    def analyzer(blob: bytes) -> List[str]:
        return client.extract_lines_from_pdf(blob)

    return _with_pdf_chunks(contents, plan, analyzer)


def _analyze_with_azure(contents: bytes, settings, source_name: str, *, date_format: str) -> AzureAnalysisResult:
    if not settings.azure_form_recognizer_endpoint or not settings.azure_form_recognizer_key:
        raise HTTPException(status_code=503, detail="Azure Form Recognizer is not configured")
    analyzer = AzureTransactionAnalyzer(
        endpoint=settings.azure_form_recognizer_endpoint,
        api_key=settings.azure_form_recognizer_key,
    )
    plan = PdfChunkingPlan(
        max_bytes=min(settings.gemini_max_document_bytes, 2_000_000),
        max_pages=1,
    )
    chunks = chunk_pdf_by_limits(contents, plan)

    combined_lines: List[str] = []
    combined_transactions: List[Any] = []

    for chunk in chunks:
        try:
            result = analyzer.analyze_pdf(chunk, source_name=source_name, date_format=date_format)
        except AzureAnalysisError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc
        combined_lines.extend(result.raw_lines)
        for asset in result.assets:
            combined_transactions.extend(asset.transactions)

    asset = AssetRecord(
        category="bank_deposit",
        type="transaction_history",
        source_document=source_name,
        asset_name="預金取引推移表",
        transactions=combined_transactions,
    )

    return AzureAnalysisResult(raw_lines=combined_lines, assets=[asset])


def _analyze_layout(contents: bytes, content_type: str) -> List[str]:
    if content_type != "application/pdf":
        raise HTTPException(status_code=415, detail="Only PDF documents are supported")

    settings = get_settings()
    try:
        return _analyze_with_gemini(contents, settings)
    except PdfChunkingError as exc:
        logger.error("PDF chunking failed: %s", exc)
        raise HTTPException(status_code=413, detail=str(exc)) from exc
    except GeminiError as exc:
        logger.error("Gemini analysis failed: %s", exc)
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
    date_format: Optional[str] = Form("auto"),
) -> DocumentAnalyzeResponse:
    contents, content_type = await _load_file_bytes(file)
    settings = get_settings()
    source_name = file.filename or "uploaded.pdf"
    date_format_normalized = (date_format or "auto").lower()
    if date_format_normalized not in {"auto", "western", "wareki"}:
        date_format_normalized = "auto"

    if document_type == "transaction_history":
        azure_result = _analyze_with_azure(contents, settings, source_name, date_format=date_format_normalized)
        return DocumentAnalyzeResponse(
            status="ok",
            document_type="transaction_history",
            raw_lines=azure_result.raw_lines,
            assets=azure_result.assets,
        )

    lines = _analyze_layout(contents, content_type)
    detected_type = document_type or detect_document_type(lines)
    if detected_type == "transaction_history":
        azure_result = _analyze_with_azure(contents, settings, source_name, date_format=date_format_normalized)
        return DocumentAnalyzeResponse(
            status="ok",
            document_type=detected_type,
            raw_lines=azure_result.raw_lines,
            assets=azure_result.assets,
        )

    assets = build_assets(detected_type, lines, source_name=source_name)
    return DocumentAnalyzeResponse(
        status="ok",
        document_type=detected_type,
        raw_lines=lines,
        assets=assets,
    )


@app.on_event("startup")
def log_startup() -> None:
    settings = get_settings()
    logger.info("Gemini model: %s", settings.gemini_model)
    logger.info(
        "Gemini chunking: max_bytes=%s, max_pages=%s",
        settings.gemini_max_document_bytes,
        settings.gemini_chunk_page_limit,
    )
