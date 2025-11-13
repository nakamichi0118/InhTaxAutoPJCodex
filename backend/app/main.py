from __future__ import annotations

import base64
import logging
from collections import OrderedDict
from datetime import datetime
from typing import Any, Callable, Dict, Iterable, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from .azure_analyzer import (
    AzureAnalysisError,
    AzureAnalysisResult,
    AzureTransactionAnalyzer,
    build_transactions_from_lines,
    merge_transactions,
    post_process_transactions,
    _reconcile_transactions,
)
from .config import get_settings
from .exporter import export_to_csv_strings
from .gemini import GeminiClient, GeminiError, GeminiExtraction
from .job_manager import JobHandle, JobManager, JobRecord
from .models import (
    AssetRecord,
    DocumentAnalyzeResponse,
    DocumentType,
    JobCreateResponse,
    JobResultResponse,
    JobStatusResponse,
    TransactionLine,
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
    analyzer: Callable[[bytes], GeminiExtraction],
) -> GeminiExtraction:
    current_plan = plan
    while True:
        chunks = chunk_pdf_by_limits(payload, current_plan)
        try:
            aggregated = GeminiExtraction(lines=[], transactions=[])
            for chunk in chunks:
                extraction = analyzer(chunk)
                aggregated.extend(extraction)
            return aggregated
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


def _analyze_with_gemini(contents: bytes, settings) -> GeminiExtraction:
    client = GeminiClient(api_keys=settings.gemini_api_keys, model=settings.gemini_model)
    plan = PdfChunkingPlan(
        max_bytes=settings.gemini_max_document_bytes,
        max_pages=settings.gemini_chunk_page_limit,
    )

    def analyzer(blob: bytes) -> GeminiExtraction:
        return client.extract_lines_from_pdf(blob)

    return _with_pdf_chunks(contents, plan, analyzer)


def _build_gemini_transaction_result(
    contents: bytes,
    settings,
    source_name: str,
    *,
    date_format: str,
) -> AzureAnalysisResult:
    extraction = _analyze_with_gemini(contents, settings)
    transactions = _convert_gemini_structured_transactions(extraction.transactions, date_format=date_format)
    if not transactions:
        transactions = build_transactions_from_lines(extraction.lines, date_format=date_format)
    asset = AssetRecord(
        category="bank_deposit",
        type="transaction_history",
        source_document=source_name,
        asset_name="預金取引推移表",
        transactions=transactions,
    )
    return AzureAnalysisResult(raw_lines=extraction.lines, assets=[asset])


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
            "Azure chunking failed (max_bytes=%s): %s. Falling back to Gemini-only analysis.",
            plan.max_bytes,
            exc,
        )
        return _build_gemini_transaction_result(contents, settings, source_name, date_format=date_format)

    combined_lines: List[str] = []
    combined_transactions: List[Any] = []

    for chunk in chunks:
        try:
            result = analyzer.analyze_pdf(chunk, source_name=source_name, date_format=date_format)
        except AzureAnalysisError as exc:
            message = str(exc)
            lowered = message.lower()
            if "out of call volume quota" in lowered or "quota" in lowered:
                logger.warning("Azure quota exhausted; falling back to Gemini-only analysis: %s", message)
                return _build_gemini_transaction_result(contents, settings, source_name, date_format=date_format)
            raise HTTPException(status_code=502, detail=message) from exc
        combined_lines.extend(result.raw_lines)
        for asset in result.assets:
            combined_transactions.extend(asset.transactions)

    azure_line_transactions = build_transactions_from_lines(combined_lines, date_format=date_format)
    if azure_line_transactions:
        combined_transactions = merge_transactions(combined_transactions, azure_line_transactions)

    gemini_transactions: List[TransactionLine] = []
    try:
        gemini_extraction = _analyze_with_gemini(contents, settings)
    except (GeminiError, PdfChunkingError) as exc:
        logger.warning("Gemini補完の取得に失敗しました: %s", exc)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Gemini補完処理で予期しないエラーが発生しました")
    else:
        gemini_transactions = _convert_gemini_structured_transactions(
            gemini_extraction.transactions,
            date_format=date_format,
        )
        if not gemini_transactions:
            gemini_transactions = build_transactions_from_lines(
                gemini_extraction.lines,
                date_format=date_format,
            )
        combined_lines = _merge_line_lists(combined_lines, gemini_extraction.lines)

    combined_transactions = post_process_transactions(combined_transactions)
    if gemini_transactions:
        gemini_transactions = post_process_transactions(gemini_transactions)
        combined_transactions = _reconcile_transactions(
            combined_transactions,
            gemini_transactions,
            None,
        )

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


def _aggregate_assets(
    store: "OrderedDict[str, AssetRecord]",
    additions: Iterable[AssetRecord],
) -> None:
    for asset in additions:
        key = _asset_aggregation_key(asset)
        if key not in store:
            store[key] = asset.model_copy(deep=True)
        else:
            target = store[key]
            if asset.transactions:
                target.transactions.extend(asset.transactions)


def _asset_aggregation_key(asset: AssetRecord) -> str:
    owners = ";".join(sorted(asset.owner_name)) if asset.owner_name else ""
    return "|".join(
        [
            asset.category or "",
            asset.type or "",
            asset.asset_name or "",
            owners,
        ]
    )


def _sort_asset_transactions(assets: Iterable[AssetRecord]) -> None:
    for asset in assets:
        if not asset.transactions:
            continue
        asset.transactions.sort(
            key=lambda txn: (
                txn.transaction_date or "",
                txn.balance if txn.balance is not None else 0.0,
            )
        )


def _convert_gemini_structured_transactions(
    items: List[Dict[str, Any]],
    *,
    date_format: str,
) -> List[TransactionLine]:
    transactions: List[TransactionLine] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        date_value = item.get("date") or item.get("transaction_date")
        transaction_date = _normalize_gemini_date(date_value)
        description = str(item.get("description") or item.get("memo") or "").strip() or None
        withdrawal = _parse_gemini_amount(
            item.get("withdrawal")
            or item.get("withdraw")
            or item.get("debit")
            or item.get("withdrawal_amount")
        )
        deposit = _parse_gemini_amount(
            item.get("deposit")
            or item.get("credit")
            or item.get("deposit_amount")
        )
        balance = _parse_gemini_amount(item.get("balance") or item.get("current_balance"))
        if not any([transaction_date, description, withdrawal, deposit, balance]):
            continue
        transactions.append(
            TransactionLine(
                transaction_date=transaction_date,
                description=description,
                withdrawal_amount=withdrawal,
                deposit_amount=deposit,
                balance=balance,
            )
        )
    return transactions


def _normalize_gemini_date(value: Any) -> Optional[str]:
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    candidates = [text, text.replace("/", "-"), text.replace(".", "-")]
    for candidate in candidates:
        normalized = candidate.replace(" ", "-")
        try:
            return datetime.fromisoformat(normalized).date().isoformat()
        except ValueError:
            pass
        parts = normalized.split("-")
        if len(parts) == 3 and all(part.isdigit() for part in parts):
            year, month, day = map(int, parts)
            try:
                return datetime(year, month, day).date().isoformat()
            except ValueError:
                continue
    return None


def _parse_gemini_amount(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    text = text.replace(",", "").replace("円", "").replace("¥", "")
    text = text.replace("＋", "+").replace("ー", "-")
    try:
        value = float(text)
        return -value if negative else value
    except ValueError:
        return None


def _analyze_layout(contents: bytes, content_type: str) -> List[str]:
    if content_type != "application/pdf":
        raise HTTPException(status_code=415, detail="Only PDF documents are supported")

    settings = get_settings()
    try:
        extraction = _analyze_with_gemini(contents, settings)
        return extraction.lines
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
    with open(job.file_path, "rb") as stream:
        contents = stream.read()
    source_name = job.file_name or "uploaded.pdf"

    plan = PdfChunkingPlan(
        max_bytes=settings.azure_chunk_max_bytes,
        max_pages=1,
    )
    try:
        chunks = chunk_pdf_by_limits(contents, plan)
    except PdfChunkingError as exc:
        raise ValueError(str(exc)) from exc

    if not chunks:
        chunks = [contents]
    total_chunks = len(chunks)
    handle.update(
        stage="analyzing",
        detail="PDFを分割しています…",
        total_chunks=total_chunks,
        processed_chunks=0,
    )

    aggregated_assets: "OrderedDict[str, AssetRecord]" = OrderedDict()
    document_type: Optional[DocumentType] = None
    latest_files: Optional[Dict[str, str]] = None

    for index, chunk in enumerate(chunks, start=1):
        handle.update(stage="analyzing", detail=f"{index}/{total_chunks} ページ解析中")
        doc_type, chunk_assets, _ = _resolve_document_assets(
            chunk,
            "application/pdf",
            job.document_type_hint,
            job.date_format,
            settings,
            source_name,
        )
        document_type = doc_type
        _aggregate_assets(aggregated_assets, chunk_assets)
        _sort_asset_transactions(aggregated_assets.values())

        payload = {"assets": [asset.to_export_payload() for asset in aggregated_assets.values()]}
        csv_map = export_to_csv_strings(payload)
        latest_files = {
            name: base64.b64encode(content.encode("utf-8-sig")).decode("ascii")
            for name, content in csv_map.items()
        }
        handle.update(
            detail=f"{index}/{total_chunks} ページ完了",
            processed_chunks=index,
            total_chunks=total_chunks,
            document_type=document_type,
            partial_files=latest_files,
        )

    if not aggregated_assets or not latest_files:
        raise ValueError("CSV出力が空です")

    handle.update(stage="exporting", detail="CSV 生成中")
    handle.update(
        status="completed",
        stage="completed",
        detail="完了",
        document_type=document_type,
        result_files=latest_files,
        partial_files=latest_files,
        processed_chunks=total_chunks,
        total_chunks=total_chunks,
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
        processed_chunks=job.processed_chunks or None,
        total_chunks=job.total_chunks or None,
        files=job.partial_files,
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
