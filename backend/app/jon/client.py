"""
JON API クライアント
土地資料収集自動化のためのAPIクライアント
"""
from __future__ import annotations

import base64
import logging
import time
from typing import Any, Dict, List, Literal, Optional

import httpx

from .models import (
    BuildingNumber,
    LocationResult,
    RegistrationResult,
    RosenImageResult,
)

logger = logging.getLogger("uvicorn.error")


class JonApiError(Exception):
    """JON API エラー"""
    def __init__(self, message: str, status_code: Optional[int] = None, raw_response: Optional[Dict] = None):
        super().__init__(message)
        self.status_code = status_code
        self.raw_response = raw_response


class JonApiClient:
    """JON API クライアント"""

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        base_url: str = "https://jon-api.com/api/v1",
        touki_login_id: Optional[str] = None,
        touki_password: Optional[str] = None,
        timeout: float = 120.0,
    ):
        """
        クライアントを初期化

        Args:
            client_id: JON API クライアントID
            client_secret: JON API クライアントシークレット
            base_url: API ベースURL
            touki_login_id: 登記情報提供サービス ログインID
            touki_password: 登記情報提供サービス パスワード
            timeout: リクエストタイムアウト（秒）
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.base_url = base_url.rstrip("/")
        self.touki_login_id = touki_login_id
        self.touki_password = touki_password
        self.timeout = timeout

    def _make_request(self, endpoint: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """APIリクエストを送信"""
        params["client_id"] = self.client_id
        params["client_secret"] = self.client_secret

        url = f"{self.base_url}/{endpoint}"

        with httpx.Client(timeout=self.timeout) as client:
            response = client.get(url, params=params)
            return response.json()

    async def _make_request_async(self, endpoint: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """非同期APIリクエストを送信"""
        params["client_id"] = self.client_id
        params["client_secret"] = self.client_secret

        url = f"{self.base_url}/{endpoint}"

        async with httpx.AsyncClient(timeout=self.timeout) as client:
            response = await client.get(url, params=params)
            return response.json()

    def locating(
        self,
        address: str,
        address_type: Literal["地番", "家屋番号", "住居表示番号", "不明"] = "不明"
    ) -> LocationResult:
        """
        位置特定API: 住所文字列から位置座標を特定

        Args:
            address: 住所文字列（例: 東京都渋谷区渋谷２丁目8-8）
            address_type: 住所種類（地番/家屋番号/住居表示番号/不明）

        Returns:
            LocationResult: 位置特定結果
        """
        params = {
            "address": address,
            "address_type": address_type
        }

        result = self._make_request("locating", params)

        if result.get("status_code") != 200:
            raise JonApiError(
                f"位置特定API エラー: {result}",
                status_code=result.get("status_code"),
                raw_response=result
            )

        locating = result["result"]["locating"]
        analysis = result["result"]["address_analysis"]

        return LocationResult(
            lat=float(locating["lat"]),
            long=float(locating["long"]),
            locating_level=locating.get("locating_level", 0),
            located_by=locating.get("located_by", ""),
            v1_code=analysis["v1_code"],
            pref=analysis.get("pref", ""),
            city=analysis.get("city", ""),
            large_section=analysis.get("large_section", ""),
            small_section=analysis.get("small_section", ""),
            number=analysis.get("number", ""),
            address_type_result=analysis.get("address_type_result", ""),
            raw_response=result
        )

    async def locating_async(
        self,
        address: str,
        address_type: Literal["地番", "家屋番号", "住居表示番号", "不明"] = "不明"
    ) -> LocationResult:
        """位置特定API（非同期版）"""
        params = {
            "address": address,
            "address_type": address_type
        }

        result = await self._make_request_async("locating", params)

        if result.get("status_code") != 200:
            raise JonApiError(
                f"位置特定API エラー: {result}",
                status_code=result.get("status_code"),
                raw_response=result
            )

        locating = result["result"]["locating"]
        analysis = result["result"]["address_analysis"]

        return LocationResult(
            lat=float(locating["lat"]),
            long=float(locating["long"]),
            locating_level=locating.get("locating_level", 0),
            located_by=locating.get("located_by", ""),
            v1_code=analysis["v1_code"],
            pref=analysis.get("pref", ""),
            city=analysis.get("city", ""),
            large_section=analysis.get("large_section", ""),
            small_section=analysis.get("small_section", ""),
            number=analysis.get("number", ""),
            address_type_result=analysis.get("address_type_result", ""),
            raw_response=result
        )

    def rosen_image(
        self,
        lat: float,
        long: float,
        length: int = 250,
        width: int = 250,
    ) -> RosenImageResult:
        """
        路線価図API: 指定座標の路線価図画像を取得

        Args:
            lat: 緯度
            long: 経度
            length: 出力画像の縦幅（1〜250m）
            width: 出力画像の横幅（1〜250m）

        Returns:
            RosenImageResult: 路線価図画像（Base64）
        """
        params = {
            "response_type": "json",
            "lat": lat,
            "long": long,
            "length": length,
            "width": width
        }

        result = self._make_request("rosen_image", params)

        if result.get("status_code") != 200:
            raise JonApiError(
                f"路線価図API エラー: {result}",
                status_code=result.get("status_code"),
                raw_response=result
            )

        # Base64データを取得
        image_data = result["image"]
        if image_data.startswith("data:image/png;base64,"):
            image_data = image_data[len("data:image/png;base64,"):]

        return RosenImageResult(
            image_base64=image_data,
            lat=lat,
            long=long,
        )

    async def rosen_image_async(
        self,
        lat: float,
        long: float,
        length: int = 250,
        width: int = 250,
    ) -> RosenImageResult:
        """路線価図API（非同期版）"""
        params = {
            "response_type": "json",
            "lat": lat,
            "long": long,
            "length": length,
            "width": width
        }

        result = await self._make_request_async("rosen_image", params)

        if result.get("status_code") != 200:
            raise JonApiError(
                f"路線価図API エラー: {result}",
                status_code=result.get("status_code"),
                raw_response=result
            )

        image_data = result["image"]
        if image_data.startswith("data:image/png;base64,"):
            image_data = image_data[len("data:image/png;base64,"):]

        return RosenImageResult(
            image_base64=image_data,
            lat=lat,
            long=long,
        )

    def convert_number(
        self,
        v1_code: str,
        number: str,
        number_type: Literal[1, 2] = 1,
        search_times: int = 1
    ) -> List[BuildingNumber]:
        """
        家屋番号API: 地番から家屋番号、家屋番号から地番を取得

        Args:
            v1_code: 登記地名コード（13桁）
            number: 地番または家屋番号
            number_type: 1=地番, 2=家屋番号
            search_times: 最大検索回数（0〜100）

        Returns:
            list[BuildingNumber]: 家屋番号リスト
        """
        params = {
            "response_type": "json",
            "v1_code": v1_code,
            "number": number,
            "number_type": number_type,
            "search_times": search_times
        }

        result = self._make_request("convert_number", params)

        if result.get("status_code") != 200:
            error_msg = result.get("error", {})
            if isinstance(error_msg, dict):
                error_msg = error_msg.get("message", str(result))
            raise JonApiError(
                f"家屋番号API エラー: {error_msg}",
                status_code=result.get("status_code"),
                raw_response=result
            )

        buildings = []
        for item in result.get("result", []):
            bn_data = item.get("building_numbers", {})
            if isinstance(bn_data, dict):
                if bn_data.get("building_number"):
                    buildings.append(BuildingNumber(
                        building_number=bn_data["building_number"],
                        v1_code=bn_data.get("v1_code", ""),
                        pref=bn_data.get("pref", ""),
                        city=bn_data.get("city", ""),
                        large_section=bn_data.get("large_section", ""),
                        small_section=bn_data.get("small_section", "")
                    ))
            elif isinstance(bn_data, list):
                for bn in bn_data:
                    buildings.append(BuildingNumber(
                        building_number=bn["building_number"],
                        v1_code=bn.get("v1_code", ""),
                        pref=bn.get("pref", ""),
                        city=bn.get("city", ""),
                        large_section=bn.get("large_section", ""),
                        small_section=bn.get("small_section", "")
                    ))

        return buildings

    def get_registration(
        self,
        v1_code: str,
        number: str,
        number_type: Literal[1, 2] = 1,
        pdf_type: Literal[1, 2, 3, 4, 5, 6] = 1
    ) -> RegistrationResult:
        """
        登記取得API: 登記情報PDFを取得

        Args:
            v1_code: 登記地名コード（13桁）
            number: 地番または家屋番号
            number_type: 1=地番, 2=家屋番号
            pdf_type: 1=全部事項, 2=所有者事項, 3=地図, 4=地積測量図, 5=地役権図面, 6=建物図面

        Returns:
            RegistrationResult: 登記取得結果
        """
        if not self.touki_login_id or not self.touki_password:
            raise JonApiError("TOUKI_LOGIN_ID と TOUKI_PASSWORD が設定されていません")

        params = {
            "response_type": "json",
            "v1_code": v1_code,
            "number": number,
            "number_type": number_type,
            "pdf_type": pdf_type,
            "login_id": self.touki_login_id,
            "password": self.touki_password
        }

        result = self._make_request("hudousan_get_single", params)

        if result.get("status_code") != 200:
            error_msg = result.get("error", {})
            if isinstance(error_msg, dict):
                error_msg = error_msg.get("message", str(result))
            raise JonApiError(
                f"登記取得API エラー: {error_msg}",
                status_code=result.get("status_code"),
                raw_response=result
            )

        r = result["result"]
        return RegistrationResult(
            pdf_id=r["pdf_id"],
            pdf_url=r["pdf_url"],
            v1_code=r["v1_code"],
            number=r["number"],
            number_type=r["number_type"],
            pdf_type=r["pdf_type"],
            based_at=r["based_at"],
            raw_response=result
        )

    async def get_registration_async(
        self,
        v1_code: str,
        number: str,
        number_type: Literal[1, 2] = 1,
        pdf_type: Literal[1, 2, 3, 4, 5, 6] = 1
    ) -> RegistrationResult:
        """登記取得API（非同期版）"""
        if not self.touki_login_id or not self.touki_password:
            raise JonApiError("TOUKI_LOGIN_ID と TOUKI_PASSWORD が設定されていません")

        params = {
            "response_type": "json",
            "v1_code": v1_code,
            "number": number,
            "number_type": number_type,
            "pdf_type": pdf_type,
            "login_id": self.touki_login_id,
            "password": self.touki_password
        }

        result = await self._make_request_async("hudousan_get_single", params)

        if result.get("status_code") != 200:
            error_msg = result.get("error", {})
            if isinstance(error_msg, dict):
                error_msg = error_msg.get("message", str(result))
            raise JonApiError(
                f"登記取得API エラー: {error_msg}",
                status_code=result.get("status_code"),
                raw_response=result
            )

        r = result["result"]
        return RegistrationResult(
            pdf_id=r["pdf_id"],
            pdf_url=r["pdf_url"],
            v1_code=r["v1_code"],
            number=r["number"],
            number_type=r["number_type"],
            pdf_type=r["pdf_type"],
            based_at=r["based_at"],
            raw_response=result
        )

    def analyze_registration(
        self,
        pdf_id: int,
        fields: Optional[List[str]] = None,
        max_retries: int = 30,
        retry_interval: float = 2.0
    ) -> Dict[str, Any]:
        """
        登記解析API: 登記情報PDFを解析

        Args:
            pdf_id: PDFのID（登記取得APIで取得）
            fields: 出力項目（省略時は全項目）
            max_retries: 解析完了待ちの最大リトライ回数
            retry_interval: リトライ間隔（秒）

        Returns:
            dict: 解析結果
        """
        params: Dict[str, Any] = {
            "response_type": "json",
            "pdf_id": pdf_id
        }

        if fields:
            for i, field in enumerate(fields):
                params[f"fields[{i}]"] = field

        for attempt in range(max_retries):
            result = self._make_request("hudousan_analyze_single", params)
            status_code = result.get("status_code")

            # 解析完了
            if status_code == 200:
                return result["result"]

            # 解析中（10001〜10163）
            if 10001 <= status_code <= 10163:
                logger.info(f"登記解析中... (status: {status_code}, attempt: {attempt + 1}/{max_retries})")
                time.sleep(retry_interval)
                continue

            # エラー
            error_msg = result.get("error", {})
            if isinstance(error_msg, dict):
                error_msg = error_msg.get("message", str(result))
            raise JonApiError(
                f"登記解析API エラー: {error_msg}",
                status_code=status_code,
                raw_response=result
            )

        raise JonApiError(f"登記解析がタイムアウトしました（{max_retries}回リトライ）")

    async def analyze_registration_async(
        self,
        pdf_id: int,
        fields: Optional[List[str]] = None,
        max_retries: int = 30,
        retry_interval: float = 2.0
    ) -> Dict[str, Any]:
        """登記解析API（非同期版）"""
        import asyncio

        params: Dict[str, Any] = {
            "response_type": "json",
            "pdf_id": pdf_id
        }

        if fields:
            for i, field in enumerate(fields):
                params[f"fields[{i}]"] = field

        for attempt in range(max_retries):
            result = await self._make_request_async("hudousan_analyze_single", params)
            status_code = result.get("status_code")

            if status_code == 200:
                return result["result"]

            if 10001 <= status_code <= 10163:
                logger.info(f"登記解析中... (status: {status_code}, attempt: {attempt + 1}/{max_retries})")
                await asyncio.sleep(retry_interval)
                continue

            error_msg = result.get("error", {})
            if isinstance(error_msg, dict):
                error_msg = error_msg.get("message", str(result))
            raise JonApiError(
                f"登記解析API エラー: {error_msg}",
                status_code=status_code,
                raw_response=result
            )

        raise JonApiError(f"登記解析がタイムアウトしました（{max_retries}回リトライ）")
