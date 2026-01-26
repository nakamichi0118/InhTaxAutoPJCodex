"""FastAPI router for JON API endpoints."""
from __future__ import annotations

import asyncio
import logging
import uuid
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException

from ..config import get_settings
from .client import JonApiClient, JonApiError
from .models import (
    AnalyzeRequest,
    JonBatchItem,
    JonBatchItemResult,
    JonBatchRequest,
    JonBatchResponse,
    JonStatusResponse,
    LocationResult,
    LocatingRequest,
    RegistrationRequest,
    RegistrationResult,
    RosenImageResult,
    RosenkaRequest,
)

logger = logging.getLogger("uvicorn.error")
router = APIRouter(prefix="/api/jon", tags=["jon"])

# バッチ処理の状態を保持（本番環境ではRedis等を使用）
_batch_jobs: Dict[str, JonBatchResponse] = {}


def _get_client() -> JonApiClient:
    """JON APIクライアントを取得"""
    settings = get_settings()
    if not settings.jon_client_id or not settings.jon_client_secret:
        raise HTTPException(
            status_code=503,
            detail="JON API is not configured. Set JON_CLIENT_ID and JON_CLIENT_SECRET."
        )
    return JonApiClient(
        client_id=settings.jon_client_id,
        client_secret=settings.jon_client_secret,
        base_url=settings.jon_api_base_url,
        touki_login_id=settings.touki_login_id,
        touki_password=settings.touki_password,
    )


@router.get("/status", response_model=JonStatusResponse)
async def get_status() -> JonStatusResponse:
    """JON API設定状態を確認"""
    settings = get_settings()
    return JonStatusResponse(
        configured=bool(settings.jon_client_id and settings.jon_client_secret),
        has_touki_credentials=bool(settings.touki_login_id and settings.touki_password),
        api_base_url=settings.jon_api_base_url,
    )


@router.post("/locating", response_model=LocationResult)
async def locating(request: LocatingRequest) -> LocationResult:
    """位置特定API: 住所から座標を取得"""
    client = _get_client()
    try:
        result = await client.locating_async(
            address=request.address,
            address_type=request.address_type,
        )
        return result
    except JonApiError as e:
        raise HTTPException(status_code=502, detail=str(e))


@router.post("/rosenka", response_model=RosenImageResult)
async def rosenka(request: RosenkaRequest) -> RosenImageResult:
    """路線価図API: 座標から路線価図を取得"""
    client = _get_client()
    try:
        result = await client.rosen_image_async(
            lat=request.lat,
            long=request.long,
            length=request.length,
            width=request.width,
        )
        return result
    except JonApiError as e:
        raise HTTPException(status_code=502, detail=str(e))


@router.post("/registration", response_model=RegistrationResult)
async def registration(request: RegistrationRequest) -> RegistrationResult:
    """登記取得API: 登記情報PDFを取得"""
    client = _get_client()
    settings = get_settings()
    if not settings.touki_login_id or not settings.touki_password:
        raise HTTPException(
            status_code=503,
            detail="Touki credentials not configured. Set TOUKI_LOGIN_ID and TOUKI_PASSWORD."
        )
    try:
        result = await client.get_registration_async(
            v1_code=request.v1_code,
            number=request.number,
            number_type=request.number_type,
            pdf_type=request.pdf_type,
        )
        return result
    except JonApiError as e:
        raise HTTPException(status_code=502, detail=str(e))


@router.post("/analyze")
async def analyze(request: AnalyzeRequest) -> Dict[str, Any]:
    """登記解析API: 登記情報PDFを解析（非同期ポーリング）"""
    client = _get_client()
    try:
        result = await client.analyze_registration_async(
            pdf_id=request.pdf_id,
            fields=request.fields,
        )
        return {"status": "ok", "result": result}
    except JonApiError as e:
        raise HTTPException(status_code=502, detail=str(e))


@router.post("/batch", response_model=JonBatchResponse, status_code=202)
async def batch_process(request: JonBatchRequest) -> JonBatchResponse:
    """バッチ処理API: 複数の不動産情報を一括取得"""
    batch_id = str(uuid.uuid4())

    # 初期状態を作成
    results = [
        JonBatchItemResult(id=item.id, status="pending")
        for item in request.properties
    ]

    response = JonBatchResponse(
        batch_id=batch_id,
        status="processing",
        total=len(request.properties),
        completed=0,
        results=results,
    )
    _batch_jobs[batch_id] = response

    # バックグラウンドでバッチ処理を開始
    asyncio.create_task(_process_batch(batch_id, request.properties))

    return response


@router.get("/batch/{batch_id}", response_model=JonBatchResponse)
async def get_batch_status(batch_id: str) -> JonBatchResponse:
    """バッチ処理状態を確認"""
    if batch_id not in _batch_jobs:
        raise HTTPException(status_code=404, detail="Batch job not found")
    return _batch_jobs[batch_id]


async def _process_batch(batch_id: str, items: List[JonBatchItem]) -> None:
    """バッチ処理を実行"""
    client = _get_client()
    job = _batch_jobs[batch_id]

    for i, item in enumerate(items):
        job.results[i].status = "processing"

        try:
            result = job.results[i]

            # 1. 位置特定
            try:
                location = await client.locating_async(
                    address=item.address,
                    address_type="地番" if item.property_type == "land" else "家屋番号",
                )
                result.location = location

                # Google Map URL生成
                if "google_map" in item.acquisitions:
                    result.google_map_url = f"https://www.google.com/maps?q={location.lat},{location.long}"

                # 2. 路線価図取得
                if "rosenka" in item.acquisitions:
                    try:
                        rosenka = await client.rosen_image_async(
                            lat=location.lat,
                            long=location.long,
                        )
                        result.rosenka_image = rosenka.image_base64
                    except JonApiError as e:
                        logger.warning(f"路線価図取得エラー: {e}")

                # 3. 登記取得
                if any(acq in item.acquisitions for acq in ["touki", "kozu", "chiseki", "tatemono"]):
                    settings = get_settings()
                    if settings.touki_login_id and settings.touki_password:
                        # pdf_type の決定
                        pdf_types_to_fetch = []
                        if "touki" in item.acquisitions:
                            pdf_types_to_fetch.append(1)  # 全部事項
                        if "kozu" in item.acquisitions:
                            pdf_types_to_fetch.append(3)  # 地図
                        if "chiseki" in item.acquisitions and item.property_type == "land":
                            pdf_types_to_fetch.append(4)  # 地積測量図
                        if "tatemono" in item.acquisitions and item.property_type == "building":
                            pdf_types_to_fetch.append(6)  # 建物図面

                        # 最初の登記情報を取得（複数取得は別途実装）
                        if pdf_types_to_fetch:
                            try:
                                number_type = 1 if item.property_type == "land" else 2
                                reg = await client.get_registration_async(
                                    v1_code=location.v1_code,
                                    number=location.number,
                                    number_type=number_type,
                                    pdf_type=pdf_types_to_fetch[0],
                                )
                                result.registration = reg

                                # 解析も実行
                                try:
                                    analyze_result = await client.analyze_registration_async(
                                        pdf_id=reg.pdf_id,
                                    )
                                    result.analyze_result = analyze_result
                                except JonApiError as e:
                                    logger.warning(f"登記解析エラー: {e}")

                            except JonApiError as e:
                                logger.warning(f"登記取得エラー: {e}")

            except JonApiError as e:
                result.error = str(e)

            result.status = "completed"

        except Exception as e:
            logger.exception(f"バッチ処理エラー (item {item.id}): {e}")
            job.results[i].status = "failed"
            job.results[i].error = str(e)

        job.completed = i + 1

    job.status = "completed"
