from __future__ import annotations

import base64
import logging
from typing import Any, Callable, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from .azure_analyzer import (
    AzureAnalysisError,
    AzureAnalysisResult,
    AzureTransactionAnalyzer,
    build_transactions_from_lines,
    merge_transactions,
    post_process_transactions,
)
from .config import get_settings
from .exporter import export_to_csv_strings
from .gemini import GeminiClient, GeminiError
from .job_manager import JobHandle, JobManager, JobRecord
from .models import (
    AssetRecord,
    DocumentAnalyzeResponse,
    DocumentType,
    JobCreateResponse,
    JobResultResponse,
    JobStatusResponse,
)
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
    client = GeminiClient(api_keys=settings.gemini_api_keys, model=settings.gemini_model)
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
        max_bytes=settings.azure_chunk_max_bytes,
        max_pages=1,
    )
    try:
        chunks = chunk_pdf_by_limits(contents, plan)
    except PdfChunkingError as exc:
        logger.warning(
            "Azure chunking failed (max_bytes=%s): %s. Falling back to single-chunk upload.",
            plan.max_bytes,
            exc,
        )
        chunks = [contents]

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

    azure_line_transactions = build_transactions_from_lines(combined_lines, date_format=date_format)
    if azure_line_transactions:
        combined_transactions = merge_transactions(combined_transactions, azure_line_transactions)

    gemini_lines: List[str] = []
    try:
        gemini_lines = _analyze_with_gemini(contents, settings)
    except (GeminiError, PdfChunkingError) as exc:
        logger.warning("Gemini補完の取得に失敗しました: %s", exc)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Gemini補完処理で予期しないエラーが発生しました")
    else:
        supplementary_transactions = build_transactions_from_lines(gemini_lines, date_format=date_format)
        if supplementary_transactions:
            combined_transactions = merge_transactions(combined_transactions, supplementary_transactions)
        combined_lines = _merge_line_lists(combined_lines, gemini_lines)
    combined_transactions = post_process_transactions(combined_transactions)

    asset = AssetRecord(
        category="bank_deposit",
        type="transaction_history",
        source_document=source_name,
        asset_name="預金取引推移表",
        transactions=combined_transactions,
    )

    return AzureAnalysisResult(raw_lines=combined_lines, assets=[asset])


def _merge_line_lists(primary: List[str], supplementary: List[str]) -> List[str]:
    if not supplementary:
        return primary
    merged = list(primary)
    existing = {line.strip() for line in primary if line and line.strip()}
    for line in supplementary:
        if not line:
            continue
        key = line.strip()
        if not key or key in existing:
            continue
        merged.append(line)
        existing.add(key)
    return merged


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


def _resolve_document_assets(
    contents: bytes,
    content_type: str,
    document_type: Optional[DocumentType],
    date_format_normalized: str,
    settings,
    source_name: str,
) -> tuple[DocumentType, List[AssetRecord], List[str]]:
    if document_type == "transaction_history":
        azure_result = _analyze_with_azure(contents, settings, source_name, date_format=date_format_normalized)
        return "transaction_history", azure_result.assets, azure_result.raw_lines

    lines = _analyze_layout(contents, content_type)
    detected_type = document_type or detect_document_type(lines)
    if detected_type == "transaction_history":
        azure_result = _analyze_with_azure(contents, settings, source_name, date_format=date_format_normalized)
        return detected_type, azure_result.assets, azure_result.raw_lines

    assets = build_assets(detected_type, lines, source_name=source_name)
    return detected_type, assets, lines


def _process_job_record(job: JobRecord, handle: JobHandle) -> None:
    settings = get_settings()
    handle.update(stage="analyzing", detail="レイアウト解析中")
    with open(job.file_path, "rb") as stream:
        contents = stream.read()
    source_name = job.file_name or "uploaded.pdf"
    doc_type, assets, _ = _resolve_document_assets(
        contents,
        job.content_type,
        job.document_type_hint,
        job.date_format,
        settings,
        source_name,
    )
    handle.update(stage="exporting", detail="CSV 生成中")
    payload = {"assets": [asset.to_export_payload() for asset in assets]}
    csv_map = export_to_csv_strings(payload)
    encoded = {
        name: base64.b64encode(content.encode("utf-8-sig")).decode("ascii")
        for name, content in csv_map.items()
    }
    if not encoded:
        raise ValueError("CSV出力が空です")
    handle.update(
        status="completed",
        stage="completed",
        detail="完了",
        document_type=doc_type,
        result_files=encoded,
    )


job_manager = JobManager(_process_job_record)


@app.get("/api/ping")
def ping() -> Dict[str, str]:
    return {"status": "ok"}


@app.get("/api/meta/limits")
def get_limits() -> Dict[str, Any]:
    settings = get_settings()
    return {
        "azure": {
            "chunk_max_mb": round(settings.azure_chunk_max_bytes / (1024 * 1024), 2),
            "chunk_max_bytes": settings.azure_chunk_max_bytes,
            "chunk_max_pages": 1,
        },
        "gemini": {
            "document_max_mb": round(settings.gemini_max_document_bytes / (1024 * 1024), 2),
            "document_max_bytes": settings.gemini_max_document_bytes,
            "chunk_page_limit": settings.gemini_chunk_page_limit,
        },
    }


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
    doc_type, assets, raw_lines = _resolve_document_assets(
        contents,
        content_type,
        document_type,
        date_format_normalized,
        settings,
        source_name,
    )
    return DocumentAnalyzeResponse(
        status="ok",
        document_type=doc_type,
        raw_lines=raw_lines,
        assets=assets,
    )


@app.post("/api/documents/analyze-export")
async def analyze_document_and_export(
    file: UploadFile = File(...),
    document_type: Optional[DocumentType] = Form(None),
    date_format: Optional[str] = Form("auto"),
) -> Dict[str, Any]:
    contents, content_type = await _load_file_bytes(file)
    settings = get_settings()
    source_name = file.filename or "uploaded.pdf"
    date_format_normalized = (date_format or "auto").lower()
    if date_format_normalized not in {"auto", "western", "wareki"}:
        date_format_normalized = "auto"

    doc_type, assets, _ = _resolve_document_assets(
        contents,
        content_type,
        document_type,
        date_format_normalized,
        settings,
        source_name,
    )
    payload = {"assets": [asset.to_export_payload() for asset in assets]}
    csv_map = export_to_csv_strings(payload)
    encoded = {
        name: base64.b64encode(content.encode("utf-8-sig")).decode("ascii")
        for name, content in csv_map.items()
    }
    return {
        "status": "ok",
        "document_type": doc_type,
        "files": encoded,
    }


@app.post("/api/jobs", response_model=JobCreateResponse, status_code=202)
async def enqueue_document_job(
    file: UploadFile = File(...),
    document_type: Optional[DocumentType] = Form(None),
    date_format: Optional[str] = Form("auto"),
) -> JobCreateResponse:
    contents, content_type = await _load_file_bytes(file)
    source_name = file.filename or "uploaded.pdf"
    date_format_normalized = (date_format or "auto").lower()
    if date_format_normalized not in {"auto", "western", "wareki"}:
        date_format_normalized = "auto"
    job = job_manager.submit(
        contents,
        content_type,
        source_name,
        document_type,
        date_format_normalized,
    )
    return JobCreateResponse(status="accepted", job_id=job.job_id)


@app.get("/api/jobs/{job_id}", response_model=JobStatusResponse)
async def get_job_status(job_id: str) -> JobStatusResponse:
    job = job_manager.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    return JobStatusResponse(
        job_id=job.job_id,
        status=job.status,
        stage=job.stage,
        detail=job.detail,
        document_type=job.document_type,
        created_at=job.created_at,
        updated_at=job.updated_at,
    )


@app.get("/api/jobs/{job_id}/result", response_model=JobResultResponse)
async def get_job_result(job_id: str) -> JobResultResponse:
    job = job_manager.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job.status != "completed" or not job.result_files:
        raise HTTPException(status_code=409, detail="Job is not completed")
    document_type = job.document_type or "unknown"
    return JobResultResponse(status="ok", job_id=job.job_id, document_type=document_type, files=job.result_files)


@app.on_event("startup")
def log_startup() -> None:
    settings = get_settings()
    logger.info("Gemini model: %s", settings.gemini_model)
    logger.info(
        "Gemini chunking: max_bytes=%s, max_pages=%s",
        settings.gemini_max_document_bytes,
        settings.gemini_chunk_page_limit,
    )
