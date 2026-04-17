"""チャンクアップロード用ルーター。

VBA（WinHttp.WinHttpRequest.5.1）はHTTP/1.1のみ対応のため、
Cloud Runの32MB制限を回避するために複数リクエストに分割して
PDFを送信する仕組みを提供する。

注意: チャンクデータをメモリ内dictで管理するため、
Cloud Run複数インスタンス構成では動作しない。
運用要件: --max-instances=1 または --session-affinity を有効化すること。
詳細: Docs/CHUNKED_UPLOAD.md 参照。
"""
from __future__ import annotations

import logging
import threading
import time
import uuid
from typing import Dict, Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from .config import get_settings
from .models import DocumentType, JobCreateResponse
from .pdf_utils import compress_pdf

logger = logging.getLogger("uvicorn.error")

router = APIRouter()

# メモリ内チャンクストア
_upload_store: Dict[str, Dict] = {}  # upload_id -> {"chunks": {int: bytes}, "created": float}
_upload_lock = threading.Lock()
_UPLOAD_TTL_SECONDS = 1800  # 30分で自動失効
_MAX_CHUNK_BYTES = 26_214_400  # 25MB（1チャンク上限）
_MAX_TOTAL_BYTES = 500 * 1024 * 1024  # 500MB（総サイズ上限）


def _collect_expired_ids() -> list:
    """ロック内で呼ぶ: 期限切れ upload_id のリストを返す（削除はしない）。"""
    now = time.time()
    return [
        uid for uid, entry in _upload_store.items()
        if now - entry["created"] > _UPLOAD_TTL_SECONDS
    ]


def _cleanup_expired() -> None:
    """期限切れアップロードエントリを削除する。

    ロック保持時間を最小化するため2段階で処理する:
    1. ロック内で期限切れIDをスナップショット取得
    2. ロック内で dict から pop（バイト配列の GC はロック外で発生する）
    """
    with _upload_lock:
        expired_ids = _collect_expired_ids()
        evicted = {uid: _upload_store.pop(uid) for uid in expired_ids}
    # ロック外で大きなバイト配列を GC に返す
    for uid, entry in evicted.items():
        entry["chunks"].clear()
        logger.info("Expired chunk upload removed: %s", uid)


def _process_pdf_bytes(contents: bytes) -> bytes:
    """PDFバイト列を受け取り、必要に応じて圧縮して返す。

    main.py の _load_file_bytes から分割した共通ロジック。
    """
    if not contents:
        raise HTTPException(status_code=400, detail="Empty file data")

    settings = get_settings()
    max_upload = settings.gemini_max_document_bytes
    if len(contents) > max_upload:
        original_mb = len(contents) / (1024 * 1024)
        logger.info(
            "PDF too large (%.1fMB > %dMB limit), attempting compression...",
            original_mb,
            max_upload // (1024 * 1024),
        )
        contents = compress_pdf(contents, max_bytes=max_upload)
        if len(contents) > max_upload:
            raise HTTPException(
                status_code=413,
                detail=(
                    f"ファイルサイズが大きすぎます（{original_mb:.0f}MB）。"
                    f"圧縮後も{len(contents) / (1024 * 1024):.0f}MBあり、"
                    f"上限{max_upload // (1024 * 1024)}MBを超えています。"
                    f"スキャン解像度を下げて再度お試しください。"
                ),
            )
        logger.info(
            "PDF compressed: %.1fMB -> %.1fMB",
            original_mb,
            len(contents) / (1024 * 1024),
        )
    return contents


@router.post("/api/upload/init")
async def init_upload() -> Dict[str, object]:
    """チャンクアップロードセッションを開始する。

    Returns:
        upload_id: 後続リクエストで使用するID
        max_chunk_bytes: 1チャンクの最大バイト数（25MB）
        max_total_bytes: 全チャンク合計の最大バイト数（500MB）
    """
    # _cleanup_expired は内部でロックを取得する
    _cleanup_expired()
    with _upload_lock:
        upload_id = uuid.uuid4().hex
        _upload_store[upload_id] = {"chunks": {}, "created": time.time()}

    logger.info("Chunk upload session initialized: %s", upload_id)
    return {
        "upload_id": upload_id,
        "max_chunk_bytes": _MAX_CHUNK_BYTES,
        "max_total_bytes": _MAX_TOTAL_BYTES,
    }


@router.post("/api/upload/{upload_id}/chunk")
async def upload_chunk(
    upload_id: str,
    chunk_index: int = Form(...),
    chunk: UploadFile = File(...),
) -> Dict[str, object]:
    """PDFの1チャンクを受信して保存する。

    Args:
        upload_id: init エンドポイントで発行されたID
        chunk_index: チャンク番号（0始まり）
        chunk: バイナリデータ
    """
    # _cleanup_expired は内部でロックを取得する（ロック外で呼ぶ）
    _cleanup_expired()
    with _upload_lock:
        if upload_id not in _upload_store:
            raise HTTPException(status_code=404, detail="Upload session not found or expired")

    data = await chunk.read()
    size = len(data)

    # Fix #5a: 1チャンクサイズ上限チェック（25MB）
    if size > _MAX_CHUNK_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Chunk size {size} exceeds limit of {_MAX_CHUNK_BYTES} bytes (25MB)",
        )

    logger.info(
        "Received chunk %d for upload %s (%.2fMB)",
        chunk_index,
        upload_id,
        size / (1024 * 1024),
    )

    with _upload_lock:
        if upload_id not in _upload_store:
            raise HTTPException(status_code=404, detail="Upload session not found or expired")

        # Fix #5b: 総サイズ上限チェック（500MB）
        current_total = sum(len(b) for b in _upload_store[upload_id]["chunks"].values())
        if current_total + size > _MAX_TOTAL_BYTES:
            raise HTTPException(
                status_code=413,
                detail="Total upload size exceeds 500MB limit",
            )

        _upload_store[upload_id]["chunks"][chunk_index] = data

    return {"received": chunk_index, "size": size}


@router.post("/api/upload/{upload_id}/jobs", response_model=JobCreateResponse, status_code=202)
async def finalize_upload_and_create_job(
    upload_id: str,
    file_name: str = Form(...),
    document_type: Optional[DocumentType] = Form(None),
    date_format: Optional[str] = Form("auto"),
    processing_mode: Optional[str] = Form("gemini"),
    gemini_model: Optional[str] = Form("gemini-2.5-pro"),
    start_date: Optional[str] = Form(None),
    end_date: Optional[str] = Form(None),
) -> JobCreateResponse:
    """チャンク群を結合してジョブを作成する。

    全チャンク受信済みであることを確認し、index順に連結してPDFを組み立てる。
    組み立て後にアップロードセッションを削除し、既存の /api/jobs と同等の
    処理フローに引き渡す。
    """
    # job_manager と SUPPORTED_GEMINI_MODELS は main.py に定義されているため遅延 import
    from .main import job_manager, SUPPORTED_GEMINI_MODELS  # noqa: PLC0415

    # _cleanup_expired は内部でロックを取得する（ロック外で呼ぶ）
    _cleanup_expired()
    with _upload_lock:
        if upload_id not in _upload_store:
            raise HTTPException(status_code=404, detail="Upload session not found or expired")
        entry = _upload_store.pop(upload_id)

    chunks: Dict[int, bytes] = entry["chunks"]
    if not chunks:
        raise HTTPException(status_code=400, detail="No chunks received")

    # Fix #1: チャンク欠番検証（0..N-1 の連続性を確認）
    sorted_indices = sorted(chunks.keys())
    max_index = sorted_indices[-1]
    expected = list(range(max_index + 1))
    if sorted_indices != expected:
        raise HTTPException(
            status_code=400,
            detail=f"Missing chunks: expected 0..{max_index}, got {sorted_indices}",
        )

    # index 順に連結
    assembled = b"".join(chunks[i] for i in sorted_indices)
    logger.info(
        "Assembled %d chunks -> %.2fMB for upload %s",
        len(sorted_indices),
        len(assembled) / (1024 * 1024),
        upload_id,
    )

    # 圧縮・サイズチェック（main.py の _process_pdf_bytes と同じロジック）
    contents = _process_pdf_bytes(assembled)
    content_type = "application/pdf"
    source_name = file_name or "uploaded.pdf"

    # date_format 正規化
    date_format_normalized = (date_format or "auto").lower()
    if date_format_normalized not in {"auto", "western", "wareki"}:
        date_format_normalized = "auto"

    # processing_mode 正規化
    incoming_mode = (processing_mode or "").strip().lower()
    if incoming_mode and incoming_mode != "gemini":
        logger.warning(
            "Unsupported processing_mode '%s' was requested; forcing gemini-only flow.",
            incoming_mode,
        )

    # gemini_model 正規化
    gemini_model_normalized: Optional[str] = None
    if gemini_model:
        candidate = gemini_model.strip()
        if candidate and candidate not in SUPPORTED_GEMINI_MODELS:
            raise HTTPException(status_code=400, detail="Unsupported Gemini model specified")
        if candidate:
            gemini_model_normalized = candidate

    # start_date / end_date 正規化
    start_date_normalized = start_date.strip() if start_date else None
    if start_date_normalized == "":
        start_date_normalized = None
    end_date_normalized = end_date.strip() if end_date else None
    if end_date_normalized == "":
        end_date_normalized = None

    job = job_manager.submit(
        contents,
        content_type,
        source_name,
        document_type,
        date_format_normalized,
        processing_mode="gemini",
        gemini_model=gemini_model_normalized,
        start_date=start_date_normalized,
        end_date=end_date_normalized,
    )
    logger.info("Job created from chunked upload: %s", job.job_id)
    return JobCreateResponse(status="accepted", job_id=job.job_id)
