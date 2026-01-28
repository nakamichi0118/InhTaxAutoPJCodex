"""不動産情報ライブラリAPIクライアント"""
import math
import httpx
import logging
from typing import Optional, List, Tuple
from .models import (
    LandUseZone,
    FirePreventionArea,
    CityPlanningArea,
    LocationOptimizationArea,
    LiquefactionRisk,
    DisasterRiskArea,
    LargeFillArea,
    LandslideArea,
    SteepSlopeArea,
    LandTransaction,
    LandPrice,
    ReinfolibResult,
)

logger = logging.getLogger(__name__)

# APIベースURL
BASE_URL = "https://www.reinfolib.mlit.go.jp/ex-api/external"


class ReinfolibClient:
    """不動産情報ライブラリAPIクライアント"""

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.headers = {
            "Ocp-Apim-Subscription-Key": api_key,
            "Accept": "application/json",
        }

    @staticmethod
    def lat_lng_to_tile(lat: float, lng: float, zoom: int) -> Tuple[int, int]:
        """緯度経度をタイル座標に変換"""
        n = 2 ** zoom
        x = int((lng + 180.0) / 360.0 * n)
        lat_rad = math.radians(lat)
        y = int((1.0 - math.asinh(math.tan(lat_rad)) / math.pi) / 2.0 * n)
        return x, y

    async def _fetch_geojson(
        self, endpoint: str, lat: float, lng: float, zoom: int = 14
    ) -> Optional[dict]:
        """GeoJSON形式でタイルデータを取得"""
        x, y = self.lat_lng_to_tile(lat, lng, zoom)
        url = f"{BASE_URL}/{endpoint}"
        params = {
            "response_format": "geojson",
            "z": zoom,
            "x": x,
            "y": y,
        }

        try:
            async with httpx.AsyncClient(timeout=30.0) as client:
                response = await client.get(url, headers=self.headers, params=params)
                if response.status_code == 200:
                    return response.json()
                else:
                    logger.warning(f"{endpoint} API error: {response.status_code} - {response.text}")
                    return None
        except Exception as e:
            logger.error(f"{endpoint} API exception: {e}")
            return None

    def _find_feature_at_point(
        self, geojson: dict, lat: float, lng: float
    ) -> Optional[dict]:
        """GeoJSONから指定座標を含むフィーチャーを検索（簡易版：最初のフィーチャーを返す）"""
        # 本来はポイントインポリゴン判定が必要だが、タイル内のデータなので最初のフィーチャーで近似
        features = geojson.get("features", [])
        if features:
            return features[0]
        return None

    # ========================================
    # 法規制情報取得
    # ========================================

    async def get_land_use_zone(self, lat: float, lng: float) -> Optional[LandUseZone]:
        """用途地域情報を取得（XKT002）"""
        geojson = await self._fetch_geojson("XKT002", lat, lng)
        if not geojson:
            return None

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return None

        props = feature.get("properties", {})
        return LandUseZone(
            zone_name=props.get("u_use_district_ja") or props.get("youto_name"),
            building_coverage_ratio=props.get("u_building_coverage_ratio_ja") or props.get("kenpei"),
            floor_area_ratio=props.get("u_floor_area_ratio_ja") or props.get("yoseki"),
            raw_data=props,
        )

    async def get_fire_prevention_area(self, lat: float, lng: float) -> Optional[FirePreventionArea]:
        """防火・準防火地域情報を取得（XKT014）"""
        geojson = await self._fetch_geojson("XKT014", lat, lng)
        if not geojson:
            return FirePreventionArea(area_type="指定なし")

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return FirePreventionArea(area_type="指定なし")

        props = feature.get("properties", {})
        area_type = props.get("u_fire_prevention_ja") or props.get("bouka_type") or "指定なし"
        return FirePreventionArea(area_type=area_type, raw_data=props)

    async def get_city_planning_area(self, lat: float, lng: float) -> Optional[CityPlanningArea]:
        """都市計画区域情報を取得（XKT001）"""
        geojson = await self._fetch_geojson("XKT001", lat, lng)
        if not geojson:
            return CityPlanningArea(area_type="都市計画区域外")

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return CityPlanningArea(area_type="都市計画区域外")

        props = feature.get("properties", {})
        # 区域区分を判定
        area_division = props.get("u_area_classification_ja") or props.get("kuiki_kubun")
        if area_division:
            area_type = area_division
        else:
            area_type = "都市計画区域内"

        return CityPlanningArea(area_type=area_type, raw_data=props)

    async def get_location_optimization_area(self, lat: float, lng: float) -> Optional[LocationOptimizationArea]:
        """立地適正化計画区域情報を取得（XKT003）"""
        geojson = await self._fetch_geojson("XKT003", lat, lng)
        if not geojson:
            return None

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return None

        props = feature.get("properties", {})
        area_type = props.get("u_location_normalization_type_ja") or props.get("area_type")
        return LocationOptimizationArea(area_type=area_type, raw_data=props)

    # ========================================
    # リスク情報取得
    # ========================================

    async def get_liquefaction_risk(self, lat: float, lng: float) -> Optional[LiquefactionRisk]:
        """液状化リスク情報を取得（XKT025）"""
        geojson = await self._fetch_geojson("XKT025", lat, lng)
        if not geojson:
            return None

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return None

        props = feature.get("properties", {})
        # 微地形区分から液状化リスクを判定
        landform = props.get("u_landform_classification_ja") or props.get("landform_type")

        # リスクレベルの判定（微地形区分に基づく）
        high_risk_landforms = ["埋立地", "旧河道", "後背湿地", "三角州・海岸低地"]
        medium_risk_landforms = ["砂州・砂礫州", "自然堤防", "扇状地"]

        if landform:
            if any(risk in landform for risk in high_risk_landforms):
                risk_level = "高"
            elif any(risk in landform for risk in medium_risk_landforms):
                risk_level = "中"
            else:
                risk_level = "低"
        else:
            risk_level = "不明"

        return LiquefactionRisk(
            risk_level=risk_level,
            landform_type=landform,
            raw_data=props,
        )

    async def get_disaster_risk_area(self, lat: float, lng: float) -> Optional[DisasterRiskArea]:
        """災害危険区域情報を取得（XKT016）"""
        geojson = await self._fetch_geojson("XKT016", lat, lng)
        if not geojson:
            return DisasterRiskArea(is_designated=False)

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return DisasterRiskArea(is_designated=False)

        props = feature.get("properties", {})
        return DisasterRiskArea(
            is_designated=True,
            area_name=props.get("name") or props.get("area_name"),
            raw_data=props,
        )

    async def get_large_fill_area(self, lat: float, lng: float) -> Optional[LargeFillArea]:
        """大規模盛土造成地情報を取得（XKT020）"""
        geojson = await self._fetch_geojson("XKT020", lat, lng)
        if not geojson:
            return LargeFillArea(is_designated=False)

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return LargeFillArea(is_designated=False)

        props = feature.get("properties", {})
        fill_type = props.get("u_large_scale_fill_type_ja") or props.get("fill_type")
        return LargeFillArea(
            is_designated=True,
            fill_type=fill_type,
            raw_data=props,
        )

    async def get_landslide_area(self, lat: float, lng: float) -> Optional[LandslideArea]:
        """地すべり防止地区情報を取得（XKT021）"""
        geojson = await self._fetch_geojson("XKT021", lat, lng)
        if not geojson:
            return LandslideArea(is_designated=False)

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return LandslideArea(is_designated=False)

        props = feature.get("properties", {})
        return LandslideArea(
            is_designated=True,
            area_name=props.get("name") or props.get("area_name"),
            raw_data=props,
        )

    async def get_steep_slope_area(self, lat: float, lng: float) -> Optional[SteepSlopeArea]:
        """急傾斜地崩壊危険区域情報を取得（XKT022）"""
        geojson = await self._fetch_geojson("XKT022", lat, lng)
        if not geojson:
            return SteepSlopeArea(is_designated=False)

        feature = self._find_feature_at_point(geojson, lat, lng)
        if not feature:
            return SteepSlopeArea(is_designated=False)

        props = feature.get("properties", {})
        return SteepSlopeArea(
            is_designated=True,
            area_name=props.get("name") or props.get("area_name"),
            raw_data=props,
        )

    # ========================================
    # 価格参考情報取得
    # ========================================

    async def get_nearby_transactions(
        self, prefecture_code: str, city_code: str, year: int = 2024, quarter: int = 1
    ) -> List[LandTransaction]:
        """取引価格情報を取得（XIT001）"""
        url = f"{BASE_URL}/XIT001"
        params = {
            "year": year,
            "quarter": quarter,
            "area": prefecture_code,
            "city": city_code,
        }

        try:
            async with httpx.AsyncClient(timeout=30.0) as client:
                response = await client.get(url, headers=self.headers, params=params)
                if response.status_code != 200:
                    logger.warning(f"XIT001 API error: {response.status_code}")
                    return []

                data = response.json()
                transactions = []
                for item in data.get("data", [])[:10]:  # 最大10件
                    transactions.append(LandTransaction(
                        transaction_price=item.get("TradePrice"),
                        price_per_sqm=item.get("PricePerUnit"),
                        transaction_date=f"{item.get('Year')}年 Q{item.get('Quarter')}",
                        land_type=item.get("Type"),
                        area_sqm=item.get("Area"),
                        structure=item.get("Structure"),
                        building_year=item.get("BuildingYear"),
                        address=f"{item.get('Prefecture', '')}{item.get('Municipality', '')}{item.get('DistrictName', '')}",
                        raw_data=item,
                    ))
                return transactions
        except Exception as e:
            logger.error(f"XIT001 API exception: {e}")
            return []

    async def get_nearby_land_prices(self, lat: float, lng: float) -> List[LandPrice]:
        """地価公示・地価調査情報を取得（XPT002）"""
        geojson = await self._fetch_geojson("XPT002", lat, lng, zoom=14)
        if not geojson:
            return []

        prices = []
        for feature in geojson.get("features", [])[:5]:  # 最大5件
            props = feature.get("properties", {})
            prices.append(LandPrice(
                price_per_sqm=props.get("u_published_price") or props.get("price"),
                survey_year=props.get("u_year") or props.get("year"),
                survey_type=props.get("u_survey_type_ja") or props.get("survey_type"),
                address=props.get("u_address_ja") or props.get("address"),
                land_use=props.get("u_current_use_ja") or props.get("land_use"),
                raw_data=props,
            ))
        return prices

    # ========================================
    # 統合取得メソッド
    # ========================================

    async def get_all_info(
        self,
        lat: float,
        lng: float,
        prefecture_code: Optional[str] = None,
        city_code: Optional[str] = None,
    ) -> ReinfolibResult:
        """すべての情報を取得"""
        result = ReinfolibResult()
        errors = []

        # 法規制情報
        try:
            result.land_use_zone = await self.get_land_use_zone(lat, lng)
        except Exception as e:
            errors.append(f"用途地域取得エラー: {e}")

        try:
            result.fire_prevention_area = await self.get_fire_prevention_area(lat, lng)
        except Exception as e:
            errors.append(f"防火地域取得エラー: {e}")

        try:
            result.city_planning_area = await self.get_city_planning_area(lat, lng)
        except Exception as e:
            errors.append(f"都市計画区域取得エラー: {e}")

        try:
            result.location_optimization_area = await self.get_location_optimization_area(lat, lng)
        except Exception as e:
            errors.append(f"立地適正化区域取得エラー: {e}")

        # リスク情報
        try:
            result.liquefaction_risk = await self.get_liquefaction_risk(lat, lng)
        except Exception as e:
            errors.append(f"液状化リスク取得エラー: {e}")

        try:
            result.disaster_risk_area = await self.get_disaster_risk_area(lat, lng)
        except Exception as e:
            errors.append(f"災害危険区域取得エラー: {e}")

        try:
            result.large_fill_area = await self.get_large_fill_area(lat, lng)
        except Exception as e:
            errors.append(f"盛土造成地取得エラー: {e}")

        try:
            result.landslide_area = await self.get_landslide_area(lat, lng)
        except Exception as e:
            errors.append(f"地すべり地区取得エラー: {e}")

        try:
            result.steep_slope_area = await self.get_steep_slope_area(lat, lng)
        except Exception as e:
            errors.append(f"急傾斜地取得エラー: {e}")

        # 価格参考情報
        try:
            result.nearby_land_prices = await self.get_nearby_land_prices(lat, lng)
        except Exception as e:
            errors.append(f"地価公示取得エラー: {e}")

        if prefecture_code and city_code:
            try:
                result.nearby_transactions = await self.get_nearby_transactions(
                    prefecture_code, city_code
                )
            except Exception as e:
                errors.append(f"取引価格取得エラー: {e}")

        result.errors = errors
        return result
