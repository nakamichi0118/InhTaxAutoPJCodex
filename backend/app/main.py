from __future__ import annotations

import base64
import csv
import logging
from datetime import datetime
from io import StringIO
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
    BALANCE_TOLERANCE,
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

CHUNK_RESIDUAL_TOLERANCE = 500.0
SUPPORTED_GEMINI_MODELS = {"gemini-2.5-flash", "gemini-2.5-pro"}
WITHDRAWAL_DESC_HINTS = (
    "振込資金",
    "振込手数料",
    "振込料",
    "送金",
    "資金移動",
    "手数料",
    "カード",
    "ＡＴＭ",
    "ATM",
    "取扱店",
)
DEPOSIT_DESC_HINTS = (
    "入金",
    "預入",
    "配当",
    "給与",
    "お利息",
)
DEPOSIT_NOTE_KEYWORDS = (
    "入金額を再算出",
    "入金扱い",
    "入金を前行",
)
WITHDRAWAL_NOTE_KEYWORDS = (
    "出金額を再算出",
    "出金扱い",
    "出金を前行",
)

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
    *,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> GeminiExtraction:
    current_plan = plan
    while True:
        chunks = chunk_pdf_by_limits(payload, current_plan)
        try:
            aggregated = GeminiExtraction(lines=[], transactions=[])
            total_chunks = len(chunks)
            for index, chunk in enumerate(chunks, start=1):
                extraction = analyzer(chunk)
                if progress_callback:
                    progress_callback(index, total_chunks or 1)
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


def _analyze_with_gemini(
    contents: bytes,
    settings,
    *,
    model_override: Optional[str] = None,
    chunk_page_limit_override: Optional[int] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> GeminiExtraction:
    model = model_override or settings.gemini_model
    client = GeminiClient(api_keys=settings.gemini_api_keys, model=model)
    max_pages = chunk_page_limit_override or settings.gemini_chunk_page_limit
    if max_pages <= 0:
        max_pages = 1
    plan = PdfChunkingPlan(
        max_bytes=settings.gemini_max_document_bytes,
        max_pages=max_pages,
    )

    def analyzer(blob: bytes) -> GeminiExtraction:
        return client.extract_lines_from_pdf(blob)

    return _with_pdf_chunks(contents, plan, analyzer, progress_callback=progress_callback)


def _build_gemini_transaction_result(
    contents: bytes,
    settings,
    source_name: str,
    *,
    date_format: str,
    model_override: Optional[str] = None,
    chunk_page_limit_override: Optional[int] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> AzureAnalysisResult:
    extraction = _analyze_with_gemini(
        contents,
        settings,
        model_override=model_override,
        chunk_page_limit_override=chunk_page_limit_override,
        progress_callback=progress_callback,
    )
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


def _analyze_with_azure(
    contents: bytes,
    settings,
    source_name: str,
    *,
    date_format: str,
    progress_reporter: Optional[Callable[[int, int], None]] = None,
    perform_global_reconciliation: bool = True,
    gemini_cache: Optional[Dict[int, List[TransactionLine]]] = None,
) -> AzureAnalysisResult:
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

    total_chunks = len(chunks) or 1
    for index, chunk in enumerate(chunks, start=1):
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
        if progress_reporter:
            progress_reporter(index, total_chunks)

    azure_line_transactions = build_transactions_from_lines(combined_lines, date_format=date_format)
    if azure_line_transactions:
        combined_transactions = merge_transactions(combined_transactions, azure_line_transactions)

    combined_lines = list(combined_lines)

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


def _append_note(existing: Optional[str], additions: List[str]) -> Optional[str]:
    parts: List[str] = []
    if existing:
        parts.append(existing)
    parts.extend(note for note in additions if note)
    combined = "; ".join(part for part in parts if part)
    return combined or None


def _note_mentions_deposit(text: Optional[str]) -> bool:
    if not text:
        return False
    return any(keyword in text for keyword in DEPOSIT_NOTE_KEYWORDS)


def _note_mentions_withdrawal(text: Optional[str]) -> bool:
    if not text:
        return False
    return any(keyword in text for keyword in WITHDRAWAL_NOTE_KEYWORDS)


def _collect_diagnostics(
    transactions: List[TransactionLine],
    start_balance: Optional[float],
    store: List[Dict[str, Any]],
    *,
    stage: str,
) -> None:
    running = start_balance
    for idx, txn in enumerate(transactions, start=1):
        withdrawal = txn.withdrawal_amount or 0.0
        deposit = txn.deposit_amount or 0.0
        if running is None:
            running = txn.balance
        else:
            running = running - withdrawal + deposit
        diff = None
        if txn.balance is not None and running is not None:
            diff = txn.balance - running
        store.append(
            {
                "stage": stage,
                "transaction_date": txn.transaction_date,
                "description": txn.description,
                "withdrawal": txn.withdrawal_amount,
                "deposit": txn.deposit_amount,
                "azure_balance": txn.balance,
                "computed_balance": running,
                "difference": diff,
            }
        )


def _build_diagnostics_csv(rows: List[Dict[str, Any]]) -> str:
    headers = [
        "stage",
        "transaction_date",
        "description",
        "withdrawal",
        "deposit",
        "azure_balance",
        "computed_balance",
        "difference",
    ]
    buffer = StringIO()
    writer = csv.writer(buffer)
    writer.writerow(headers)
    for row in rows:
        writer.writerow([row.get(col, "") for col in headers])
    return buffer.getvalue()


def _build_azure_raw_row(
    page_number: int,
    row_number: int,
    txn: TransactionLine,
) -> Dict[str, Any]:
    return {
        "page_number": page_number,
        "row_number": row_number,
        "transaction_date": txn.transaction_date,
        "description": txn.description,
        "withdrawal_amount": txn.withdrawal_amount,
        "deposit_amount": txn.deposit_amount,
        "balance": txn.balance,
        "line_confidence": txn.line_confidence,
    }


def _build_azure_raw_transactions_csv(rows: List[Dict[str, Any]]) -> str:
    headers = [
        "page_number",
        "row_number",
        "transaction_date",
        "description",
        "withdrawal_amount",
        "deposit_amount",
        "balance",
        "line_confidence",
    ]
    buffer = StringIO()
    writer = csv.writer(buffer)
    writer.writerow(headers)
    for row in rows:
        writer.writerow([row.get(col, "") for col in headers])
    return buffer.getvalue()


def _extract_transactions_with_gemini_chunk(
    chunk_bytes: bytes,
    settings,
    source_name: str,
    *,
    date_format: str,
    model_override: Optional[str] = None,
) -> List[TransactionLine]:
    result = _build_gemini_transaction_result(
        chunk_bytes,
        settings,
        source_name,
        date_format=date_format,
        model_override=model_override,
        chunk_page_limit_override=1,
    )
    transactions: List[TransactionLine] = []
    for asset in result.assets:
        transactions.extend(asset.transactions)
    return transactions


def _compute_chunk_residuals(transactions: List[TransactionLine]) -> List[float]:
    residuals: List[float] = []
    if not transactions:
        return residuals
    running_balance = transactions[0].balance
    if running_balance is None:
        running_balance = (transactions[0].deposit_amount or 0.0) - (transactions[0].withdrawal_amount or 0.0)
    residuals.append(_balance_difference(transactions[0].balance, running_balance))
    for txn in transactions[1:]:
        withdrawal = txn.withdrawal_amount or 0.0
        deposit = txn.deposit_amount or 0.0
        running_balance = running_balance - withdrawal + deposit if running_balance is not None else None
        expected = running_balance if running_balance is not None else 0.0
        residuals.append(_balance_difference(txn.balance, expected))
    return residuals


def _balance_difference(actual: Optional[float], expected: float) -> float:
    if actual is None:
        return 0.0
    return float(actual) - expected


def _needs_balance_fix(transactions: List[TransactionLine]) -> bool:
    if not transactions:
        return False
    for diff in _compute_chunk_residuals(transactions):
        if abs(diff) > CHUNK_RESIDUAL_TOLERANCE:
            return True
    return False


def _gemini_refine_chunk(
    chunk_bytes: bytes,
    azure_transactions: List[TransactionLine],
    initial_balance: Optional[float],
    settings,
    *,
    date_format: str,
    model_override: Optional[str] = None,
    chunk_page_limit_override: Optional[int] = None,
) -> Optional[List[TransactionLine]]:
    try:
        extraction = _analyze_with_gemini(
            chunk_bytes,
            settings,
            model_override=model_override,
            chunk_page_limit_override=chunk_page_limit_override,
        )
    except (GeminiError, PdfChunkingError) as exc:
        logger.warning("Gemini補正に失敗しました: %s", exc)
        return None
    except Exception as exc:  # noqa: BLE001
        logger.exception("Gemini補正処理で予期しないエラーが発生しました")
        return None

    gemini_transactions = _convert_gemini_structured_transactions(
        extraction.transactions,
        date_format=date_format,
    )
    if not gemini_transactions:
        gemini_transactions = build_transactions_from_lines(
            extraction.lines,
            date_format=date_format,
        )
    if not gemini_transactions:
        return None

    gemini_transactions = post_process_transactions(gemini_transactions)
    try:
        reconciled = _reconcile_transactions(
            azure_transactions,
            gemini_transactions,
            None,
            initial_balance=initial_balance,
        )
    except Exception as exc:  # noqa: BLE001
        logger.exception("Gemini補正の適用に失敗しました: %s", exc)
        return None
    return reconciled


def _enforce_continuity(
    prev_balance: Optional[float],
    transactions: List[TransactionLine],
) -> tuple[List[TransactionLine], Optional[float]]:
    if not transactions:
        return [], prev_balance
    updated: List[TransactionLine] = []
    running_balance = prev_balance
    for txn in transactions:
        original_withdrawal = txn.withdrawal_amount
        original_deposit = txn.deposit_amount
        withdrawal = original_withdrawal or 0.0
        deposit = original_deposit or 0.0
        notes: List[str] = []
        actual_balance = txn.balance
        description = (txn.description or "").lower()
        has_withdrawal = original_withdrawal is not None and original_withdrawal != 0
        has_deposit = original_deposit is not None and original_deposit != 0

        if has_deposit and not has_withdrawal:
            if any(hint in description for hint in WITHDRAWAL_DESC_HINTS):
                withdrawal = deposit
                deposit = 0.0
                has_withdrawal = True
                has_deposit = False
                notes.append("摘要から出金扱いに補正しました")
        if has_withdrawal and not has_deposit:
            if any(hint in description for hint in DEPOSIT_DESC_HINTS):
                deposit = withdrawal
                withdrawal = 0.0
                has_withdrawal = False
                has_deposit = True
                notes.append("摘要から入金扱いに補正しました")
        if running_balance is not None and actual_balance is not None:
            balance_change = actual_balance - running_balance
            if balance_change < 0 and not has_withdrawal and has_deposit:
                withdrawal = deposit
                deposit = 0.0
                has_withdrawal = True
                has_deposit = False
                notes.append("残高推移に合わせて入出金を入れ替えました")
            elif balance_change > 0 and has_withdrawal and not has_deposit:
                deposit = withdrawal
                withdrawal = 0.0
                has_withdrawal = False
                has_deposit = True
                notes.append("残高推移に合わせて入出金を入れ替えました")
        if running_balance is not None:
            if actual_balance is not None:
                balance_delta = running_balance - actual_balance
                if abs(balance_delta) > BALANCE_TOLERANCE:
                    if balance_delta > 0:
                        withdrawal = abs(balance_delta)
                        deposit = 0.0
                        notes.append("残高差から出金額を再算出しました")
                    elif balance_delta < 0:
                        deposit = abs(balance_delta)
                        withdrawal = 0.0
                        notes.append("残高差から入金額を再算出しました")
                    actual_balance = running_balance - withdrawal + deposit
            if actual_balance is None:
                actual_balance = running_balance - withdrawal + deposit
                notes.append("残高を前行から補完しました")

        running_balance = actual_balance
        update_fields: Dict[str, Any] = {}
        if actual_balance is not None and actual_balance != txn.balance:
            update_fields["balance"] = actual_balance
        if (withdrawal or None) != txn.withdrawal_amount:
            update_fields["withdrawal_amount"] = withdrawal or None
        if (deposit or None) != txn.deposit_amount:
            update_fields["deposit_amount"] = deposit or None
        if notes:
            update_fields["correction_note"] = _append_note(txn.correction_note, notes)
        if update_fields:
            txn = txn.model_copy(update=update_fields)

        note_text = txn.correction_note or ""
        if (
            txn.withdrawal_amount
            and (txn.deposit_amount is None or txn.deposit_amount == 0)
            and _note_mentions_deposit(note_text)
        ):
            txn = txn.model_copy(
                update={
                    "deposit_amount": txn.withdrawal_amount,
                    "withdrawal_amount": None,
                    "correction_note": _append_note(note_text, ["入出金欄を残高整合の結果として入れ替えました"]),
                }
            )
        elif (
            txn.deposit_amount
            and (txn.withdrawal_amount is None or txn.withdrawal_amount == 0)
            and _note_mentions_withdrawal(note_text)
        ):
            txn = txn.model_copy(
                update={
                    "withdrawal_amount": txn.deposit_amount,
                    "deposit_amount": None,
                    "correction_note": _append_note(note_text, ["入出金欄を残高整合の結果として入れ替えました"]),
                }
            )

        updated.append(txn)

    return updated, running_balance


def _finalize_transaction_directions(transactions: List[TransactionLine]) -> List[TransactionLine]:
    finalized: List[TransactionLine] = []
    prev_balance: Optional[float] = None
    for txn in transactions:
        current = txn
        deposit = txn.deposit_amount or 0.0
        withdrawal = txn.withdrawal_amount or 0.0
        note_text = txn.correction_note or ""

        def _swap(to_deposit: bool) -> None:
            nonlocal current, deposit, withdrawal, note_text
            if to_deposit:
                deposit = withdrawal
                withdrawal = 0.0
            else:
                withdrawal = deposit
                deposit = 0.0
            note_text = _append_note(note_text, ["残高整合に合わせて入出金を入れ替えました"]) or ""
            current = current.model_copy(
                update={
                    "deposit_amount": deposit or None,
                    "withdrawal_amount": withdrawal or None,
                    "correction_note": note_text,
                }
            )

        if prev_balance is not None and current.balance is not None:
            delta = current.balance - prev_balance
            recorded_delta = deposit - withdrawal
            if abs(delta - recorded_delta) > BALANCE_TOLERANCE:
                if delta > BALANCE_TOLERANCE:
                    deposit = float(delta)
                    withdrawal = 0.0
                    note_text = _append_note(note_text, ["残高の増加量に合わせて入金額を補正しました"]) or ""
                    current = current.model_copy(
                        update={
                            "deposit_amount": deposit,
                            "withdrawal_amount": None,
                            "correction_note": note_text,
                        }
                    )
                elif delta < -BALANCE_TOLERANCE:
                    withdrawal = float(-delta)
                    deposit = 0.0
                    note_text = _append_note(note_text, ["残高の減少量に合わせて出金額を補正しました"]) or ""
                    current = current.model_copy(
                        update={
                            "withdrawal_amount": withdrawal,
                            "deposit_amount": None,
                            "correction_note": note_text,
                        }
                    )
                elif withdrawal and not deposit:
                    _swap(to_deposit=True)
                elif deposit and not withdrawal:
                    _swap(to_deposit=False)
        else:
            if withdrawal and not deposit and _note_mentions_deposit(note_text):
                _swap(to_deposit=True)
            elif deposit and not withdrawal and _note_mentions_withdrawal(note_text):
                _swap(to_deposit=False)

        if current.balance is not None:
            prev_balance = current.balance
        elif prev_balance is not None:
            # 保守的に残高は変更しない
            pass

        finalized.append(current)

    return finalized


def _recompute_balances(transactions: List[TransactionLine]) -> List[TransactionLine]:
    recomputed: List[TransactionLine] = []
    running: Optional[float] = None
    initialized = False
    for txn in transactions:
        withdrawal = txn.withdrawal_amount or 0.0
        deposit = txn.deposit_amount or 0.0
        update_fields: Dict[str, Any] = {}
        if not initialized and txn.balance is not None:
            running = float(txn.balance)
            initialized = True
        else:
            if running is None:
                running = deposit - withdrawal
            else:
                running = running - withdrawal + deposit
            update_fields["balance"] = running
        if update_fields:
            txn = txn.model_copy(update=update_fields)
        recomputed.append(txn)
    return recomputed


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
    *,
    progress_reporter: Optional[Callable[[int, int], None]] = None,
) -> tuple[DocumentType, List[AssetRecord], List[str]]:
    if document_type == "transaction_history":
        azure_result = _analyze_with_azure(
            contents,
            settings,
            source_name,
            date_format=date_format_normalized,
            progress_reporter=progress_reporter,
        )
        return "transaction_history", azure_result.assets, azure_result.raw_lines

    lines = _analyze_layout(contents, content_type)
    detected_type = document_type or detect_document_type(lines)
    if detected_type == "transaction_history":
        azure_result = _analyze_with_azure(
            contents,
            settings,
            source_name,
            date_format=date_format_normalized,
            progress_reporter=progress_reporter,
        )
        return detected_type, azure_result.assets, azure_result.raw_lines

    assets = build_assets(detected_type, lines, source_name=source_name)
    return detected_type, assets, lines


def _process_job_record(job: JobRecord, handle: JobHandle) -> None:
    settings = get_settings()
    handle.update(stage="queued", detail="ジョブを初期化しています…")
    with open(job.file_path, "rb") as stream:
        contents = stream.read()
    source_name = job.file_name or "uploaded.pdf"

    if job.processing_mode == "gemini":
        model_label = job.gemini_model or settings.gemini_model
        handle.update(
            stage="analyzing",
            detail=f"Gemini ({model_label}) で解析しています…",
            processed_chunks=0,
            total_chunks=1,
        )
        progress_state = {"total": 0}

        def _report_progress(current: int, total: int) -> None:
            progress_state["total"] = total
            handle.update(
                stage="analyzing",
                detail=f"Gemini ({model_label}) 解析中… {current}/{total}",
                processed_chunks=current,
                total_chunks=total,
            )

        try:
            gemini_result = _build_gemini_transaction_result(
                contents,
                settings,
                source_name,
                date_format=job.date_format,
                model_override=job.gemini_model,
                chunk_page_limit_override=1,
                progress_callback=_report_progress,
            )
        except Exception as exc:  # noqa: BLE001
            raise ValueError(str(exc)) from exc
        assets = gemini_result.assets or []
        if not assets:
            raise ValueError("取引が抽出できませんでした")

        export_assets: List[dict] = []
        document_type = assets[0].category if assets else "transaction_history"
        for asset in assets:
            transactions = asset.transactions or []
            transactions, _ = _enforce_continuity(None, transactions)
            transactions = post_process_transactions(transactions)
            transactions = _finalize_transaction_directions(transactions)
            transactions = _recompute_balances(transactions)
            asset.transactions = transactions
            export_assets.append(asset.to_export_payload())

        payload = {"assets": export_assets}
        handle.update(stage="exporting", detail="CSV を書き出しています…")
        csv_map = export_to_csv_strings(payload)
        encoded = {
            name: base64.b64encode(content.encode("utf-8-sig")).decode("ascii")
            for name, content in csv_map.items()
        }
        gemini_total_chunks = progress_state["total"] or 1
        handle.update(
            status="completed",
            stage="completed",
            detail="完了",
            document_type=document_type,
            result_files=encoded,
            partial_files=encoded,
            processed_chunks=gemini_total_chunks,
            total_chunks=gemini_total_chunks,
            assets_payload=export_assets,
        )
        return

    plan = PdfChunkingPlan(
        max_bytes=settings.azure_chunk_max_bytes,
        max_pages=1,
    )
    try:
        chunks = chunk_pdf_by_limits(contents, plan)
    except PdfChunkingError as exc:
        raise ValueError(str(exc)) from exc

    if not chunks:
        raise ValueError("PDFにページが含まれていません")

    total_chunks = len(chunks)
    handle.update(
        stage="analyzing",
        detail="レイアウト解析を開始します…",
        processed_chunks=0,
        total_chunks=total_chunks,
    )

    all_transactions: List[TransactionLine] = []
    diagnostics: List[Dict[str, Any]] = []
    prev_balance: Optional[float] = None
    document_type: Optional[DocumentType] = None
    azure_raw_rows: List[Dict[str, Any]] = []

    for index, chunk in enumerate(chunks, start=1):
        handle.update(
            stage="analyzing",
            detail=f"{index}/{total_chunks} ページ解析中…",
            processed_chunks=index - 1,
            total_chunks=total_chunks,
        )
        chunk_result = _analyze_with_azure(
            chunk,
            settings,
            source_name,
            date_format=job.date_format,
            perform_global_reconciliation=False,
        )
        document_type = chunk_result.assets[0].category if chunk_result.assets else "transaction_history"
        raw_transactions = build_transactions_from_lines(
            chunk_result.raw_lines,
            date_format=job.date_format,
        )
        _collect_diagnostics(raw_transactions, prev_balance, diagnostics, stage="azure_raw")

        chunk_transactions: List[TransactionLine] = []
        chunk_row_number = 0
        for asset in chunk_result.assets:
            for txn in asset.transactions:
                chunk_row_number += 1
                azure_raw_rows.append(_build_azure_raw_row(index, chunk_row_number, txn))
                chunk_transactions.append(txn)
        if not chunk_transactions:
            fallback_txns = _extract_transactions_with_gemini_chunk(
                chunk,
                settings,
                f"{source_name}#chunk{index}",
                date_format=job.date_format,
                model_override=job.gemini_model,
            )
            for txn in fallback_txns:
                chunk_row_number += 1
                azure_raw_rows.append(_build_azure_raw_row(index, chunk_row_number, txn))
            chunk_transactions = fallback_txns
        chunk_start_balance = prev_balance
        chunk_transactions, prev_balance = _enforce_continuity(prev_balance, chunk_transactions)
        _collect_diagnostics(chunk_transactions, chunk_start_balance, diagnostics, stage="adjusted")

        if _needs_balance_fix(chunk_transactions):
            handle.update(
                stage="analyzing",
                detail=f"{index}/{total_chunks} ページのAI補正を適用しています…",
                processed_chunks=index - 1,
                total_chunks=total_chunks,
            )
            refined = _gemini_refine_chunk(
                chunk,
                chunk_transactions,
                chunk_start_balance,
                settings,
                date_format=job.date_format,
                model_override=job.gemini_model,
                chunk_page_limit_override=1 if job.gemini_model else None,
            )
            if refined:
                chunk_transactions = refined
                chunk_transactions, prev_balance = _enforce_continuity(chunk_start_balance, chunk_transactions)

        all_transactions.extend(chunk_transactions)
        handle.update(
            stage="analyzing",
            detail=f"{index}/{total_chunks} ページ完了",
            processed_chunks=index,
            total_chunks=total_chunks,
        )

    if not all_transactions:
        raise ValueError("取引が抽出できませんでした")

    handle.update(stage="analyzing", detail="残高を整合しています…")
    reconciled_transactions, _ = _enforce_continuity(None, all_transactions)
    reconciled_transactions = post_process_transactions(reconciled_transactions)
    reconciled_transactions = _finalize_transaction_directions(reconciled_transactions)
    reconciled_transactions = _recompute_balances(reconciled_transactions)

    asset = AssetRecord(
        category=document_type or "transaction_history",
        type="transaction_history",
        source_document=source_name,
        asset_name="預金取引推移表",
        transactions=reconciled_transactions,
    )

    asset_payload = asset.to_export_payload()
    payload = {"assets": [asset_payload]}
    handle.update(stage="exporting", detail="CSV を書き出しています…")
    csv_map = export_to_csv_strings(payload)
    if not csv_map:
        raise ValueError("CSV出力が空です")
    encoded = {
        name: base64.b64encode(content.encode("utf-8-sig")).decode("ascii")
        for name, content in csv_map.items()
    }
    if diagnostics:
        debug_csv = _build_diagnostics_csv(diagnostics)
        encoded["azure_balance_diagnostics.csv"] = base64.b64encode(
            debug_csv.encode("utf-8-sig")
        ).decode("ascii")
    if azure_raw_rows:
        raw_csv = _build_azure_raw_transactions_csv(azure_raw_rows)
        encoded["azure_raw_transactions.csv"] = base64.b64encode(
            raw_csv.encode("utf-8-sig")
        ).decode("ascii")
    assets_payload = [asset_payload]

    handle.update(
        status="completed",
        stage="completed",
        detail="完了",
        document_type=asset.category,
        result_files=encoded,
        partial_files=encoded,
        processed_chunks=total_chunks,
        total_chunks=total_chunks,
        assets_payload=assets_payload,
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
    processing_mode: Optional[str] = Form("gemini"),
    gemini_model: Optional[str] = Form("gemini-2.5-pro"),
) -> JobCreateResponse:
    contents, content_type = await _load_file_bytes(file)
    source_name = file.filename or "uploaded.pdf"
    date_format_normalized = (date_format or "auto").lower()
    if date_format_normalized not in {"auto", "western", "wareki"}:
        date_format_normalized = "auto"
    processing_mode_normalized = (processing_mode or "hybrid").lower()
    if processing_mode_normalized not in {"hybrid", "gemini"}:
        processing_mode_normalized = "hybrid"
    gemini_model_normalized: Optional[str] = None
    if gemini_model:
        candidate = gemini_model.strip()
        if candidate and candidate not in SUPPORTED_GEMINI_MODELS:
            raise HTTPException(status_code=400, detail="Unsupported Gemini model specified")
        if candidate:
            gemini_model_normalized = candidate
    if processing_mode_normalized != "gemini":
        gemini_model_normalized = None
    job = job_manager.submit(
        contents,
        content_type,
        source_name,
        document_type,
        date_format_normalized,
        processing_mode=processing_mode_normalized,
        gemini_model=gemini_model_normalized,
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
    assets_payload = job.assets_payload or []
    assets: List[AssetRecord] = [AssetRecord.model_validate(item) for item in assets_payload]
    return JobResultResponse(
        status="ok",
        job_id=job.job_id,
        document_type=document_type,
        files=job.result_files,
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
