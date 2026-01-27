"""FastAPI router for JON API endpoints."""
from __future__ import annotations

import asyncio
import logging
import uuid
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException

from ..config import get_settings
from .client import JonApiClient, JonApiError
from .rosenka_lookup import lookup_rosenka_urls
from .models import (
    AnalyzeRequest,
    ForceRegistrationRequest,
    ForceRegistrationResponse,
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


@router.post("/force-registration", response_model=ForceRegistrationResponse)
async def force_registration(request: ForceRegistrationRequest) -> ForceRegistrationResponse:
    """
    強制登記取得API: 精度が低い場合でも登記情報を取得
    ユーザーが明示的に「取得を試みる」を選択した場合に使用
    """
    # バッチジョブから対象アイテムを取得
    if request.batch_id not in _batch_jobs:
        raise HTTPException(status_code=404, detail="Batch job not found")

    job = _batch_jobs[request.batch_id]
    item_result = next((r for r in job.results if r.id == request.item_id), None)
    if not item_result:
        raise HTTPException(status_code=404, detail="Item not found in batch")

    if not item_result.location:
        raise HTTPException(status_code=400, detail="位置情報がありません。先に位置特定を実行してください。")

    client = _get_client()
    settings = get_settings()

    if not settings.touki_login_id or not settings.touki_password:
        raise HTTPException(
            status_code=503,
            detail="登記情報提供サービスの認証情報が設定されていません。"
        )

    response = ForceRegistrationResponse(success=False, item_id=request.item_id)
    location = item_result.location

    # バッチアイテムから物件種別を推定（locationのaddress_type_resultで判定）
    is_land = location.address_type_result != "家屋番号"
    number_type = 1 if is_land else 2

    try:
        # 登記簿（全部事項）
        if 1 in request.pdf_types:
            try:
                reg = await client.get_registration_async(
                    v1_code=location.v1_code,
                    number=location.number,
                    number_type=number_type,
                    pdf_type=1,
                )
                response.registration_pdf_url = reg.pdf_url
                item_result.registration = reg
                logger.info(f"強制取得成功（登記簿）: {request.item_id}")
            except JonApiError as e:
                logger.warning(f"強制取得エラー（登記簿）: {e}")
                response.error = f"登記簿取得エラー: {e}"

        # 公図
        if 3 in request.pdf_types:
            try:
                kozu_reg = await client.get_registration_async(
                    v1_code=location.v1_code,
                    number=location.number,
                    number_type=number_type,
                    pdf_type=3,
                )
                response.kozu_pdf_url = kozu_reg.pdf_url
                item_result.kozu_pdf_url = kozu_reg.pdf_url
                logger.info(f"強制取得成功（公図）: {request.item_id}")
            except JonApiError as e:
                logger.warning(f"強制取得エラー（公図）: {e}")
                if response.error:
                    response.error += f" / 公図取得エラー: {e}"
                else:
                    response.error = f"公図取得エラー: {e}"

        # 地積測量図（土地のみ）
        if 4 in request.pdf_types and is_land:
            try:
                chiseki_reg = await client.get_registration_async(
                    v1_code=location.v1_code,
                    number=location.number,
                    number_type=number_type,
                    pdf_type=4,
                )
                response.chiseki_pdf_url = chiseki_reg.pdf_url
                logger.info(f"強制取得成功（地積測量図）: {request.item_id}")
            except JonApiError as e:
                logger.warning(f"強制取得エラー（地積測量図）: {e}")

        # 建物図面（建物のみ）
        if 6 in request.pdf_types and not is_land:
            try:
                tatemono_reg = await client.get_registration_async(
                    v1_code=location.v1_code,
                    number=location.number,
                    number_type=number_type,
                    pdf_type=6,
                )
                response.tatemono_pdf_url = tatemono_reg.pdf_url
                logger.info(f"強制取得成功（建物図面）: {request.item_id}")
            except JonApiError as e:
                logger.warning(f"強制取得エラー（建物図面）: {e}")

        # 警告メッセージを更新
        item_result.accuracy_warning = None
        response.success = True

    except Exception as e:
        logger.exception(f"強制取得エラー: {e}")
        response.error = str(e)

    return response


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
                result.locating_level = location.locating_level
                result.accuracy_label = location.accuracy_label
                logger.info(f"位置特定結果: address={item.address}, level={location.locating_level}, v1_code={location.v1_code}, raw_locating={location.raw_response.get('result', {}).get('locating', {})}")

                # 精度チェック
                is_high_accuracy = location.is_high_accuracy
                # 登記関連項目が選択されている場合のみ精度警告を表示
                has_registration_items = any(acq in item.acquisitions for acq in ["touki", "kozu", "chiseki", "tatemono"])
                if not is_high_accuracy and has_registration_items:
                    result.accuracy_warning = f"位置特定の精度が低いため（レベル{location.locating_level}）、登記情報の自動取得をスキップしました。住所を確認してください。"
                    logger.warning(f"位置精度低: {item.address} (level={location.locating_level})")

                # Google Map URL生成
                if "google_map" in item.acquisitions:
                    result.google_map_url = f"https://www.google.com/maps?q={location.lat},{location.long}"

                # 2. 路線価図取得（精度に関係なく取得）
                if "rosenka" in item.acquisitions:
                    # JON APIから路線価図画像を取得
                    try:
                        rosenka = await client.rosen_image_async(
                            lat=location.lat,
                            long=location.long,
                        )
                        result.rosenka_image = rosenka.image_base64
                    except JonApiError as e:
                        logger.warning(f"路線価図取得エラー: {e}")

                    # GCSから路線価図URLを検索
                    try:
                        search_district = location.small_section or location.large_section
                        logger.info(f"路線価URL検索: pref={location.pref}, city={location.city}, district={search_district}")
                        rosenka_urls = await lookup_rosenka_urls(
                            prefecture=location.pref,
                            city=location.city,
                            district=search_district,
                        )
                        if rosenka_urls:
                            result.rosenka_urls = rosenka_urls
                            logger.info(f"路線価URL取得成功: {len(rosenka_urls)}件 ({location.pref}/{location.city}/{search_district})")
                        else:
                            logger.info(f"路線価URL見つからず: {location.pref}/{location.city}/{search_district}")
                    except Exception as e:
                        logger.warning(f"路線価URL検索エラー: {e}")

                # 3. 登記取得（高精度の場合のみ）
                if any(acq in item.acquisitions for acq in ["touki", "kozu", "chiseki", "tatemono"]):
                    if not is_high_accuracy:
                        # 精度が低い場合はスキップ（警告は既に設定済み）
                        logger.info(f"登記取得スキップ（精度不足）: {item.address}")
                    else:
                        settings = get_settings()
                        if settings.touki_login_id and settings.touki_password:
                            number_type = 1 if item.property_type == "land" else 2

                            # 登記簿（全部事項）取得
                            if "touki" in item.acquisitions:
                                try:
                                    reg = await client.get_registration_async(
                                        v1_code=location.v1_code,
                                        number=location.number,
                                        number_type=number_type,
                                        pdf_type=1,  # 全部事項
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
                                    logger.warning(f"登記簿取得エラー: {e}")

                            # 公図取得
                            if "kozu" in item.acquisitions:
                                try:
                                    kozu_reg = await client.get_registration_async(
                                        v1_code=location.v1_code,
                                        number=location.number,
                                        number_type=number_type,
                                        pdf_type=3,  # 地図（公図）
                                    )
                                    result.kozu_pdf_url = kozu_reg.pdf_url
                                except JonApiError as e:
                                    logger.warning(f"公図取得エラー: {e}")

                            # 地積測量図取得（土地のみ）
                            if "chiseki" in item.acquisitions and item.property_type == "land":
                                try:
                                    await client.get_registration_async(
                                        v1_code=location.v1_code,
                                        number=location.number,
                                        number_type=number_type,
                                        pdf_type=4,  # 地積測量図
                                    )
                                    # TODO: 結果をresultに格納
                                except JonApiError as e:
                                    logger.warning(f"地積測量図取得エラー: {e}")

                            # 建物図面取得（建物のみ）
                            if "tatemono" in item.acquisitions and item.property_type == "building":
                                try:
                                    await client.get_registration_async(
                                        v1_code=location.v1_code,
                                        number=location.number,
                                        number_type=number_type,
                                        pdf_type=6,  # 建物図面
                                    )
                                    # TODO: 結果をresultに格納
                                except JonApiError as e:
                                    logger.warning(f"建物図面取得エラー: {e}")

            except JonApiError as e:
                result.error = str(e)

            result.status = "completed"

        except Exception as e:
            logger.exception(f"バッチ処理エラー (item {item.id}): {e}")
            job.results[i].status = "failed"
            job.results[i].error = str(e)

        job.completed = i + 1

    job.status = "completed"
