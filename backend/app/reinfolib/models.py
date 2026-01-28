"""不動産情報ライブラリAPIのデータモデル"""
from typing import Optional, List
from pydantic import BaseModel


# ========================================
# 法規制情報
# ========================================

class LandUseZone(BaseModel):
    """用途地域情報（XKT002）"""
    zone_name: Optional[str] = None  # 用途地域名（第一種住居地域など）
    building_coverage_ratio: Optional[str] = None  # 建ぺい率
    floor_area_ratio: Optional[str] = None  # 容積率
    raw_data: Optional[dict] = None  # 生データ


class FirePreventionArea(BaseModel):
    """防火・準防火地域情報（XKT014）"""
    area_type: Optional[str] = None  # 防火地域、準防火地域、指定なし
    raw_data: Optional[dict] = None


class CityPlanningArea(BaseModel):
    """都市計画区域情報（XKT001）"""
    area_type: Optional[str] = None  # 市街化区域、市街化調整区域、非線引き、都市計画区域外
    raw_data: Optional[dict] = None


class LocationOptimizationArea(BaseModel):
    """立地適正化計画区域情報（XKT003）"""
    area_type: Optional[str] = None  # 居住誘導区域、都市機能誘導区域など
    raw_data: Optional[dict] = None


# ========================================
# リスク情報
# ========================================

class LiquefactionRisk(BaseModel):
    """液状化リスク情報（XKT025）"""
    risk_level: Optional[str] = None  # リスクレベル
    landform_type: Optional[str] = None  # 微地形区分
    raw_data: Optional[dict] = None


class DisasterRiskArea(BaseModel):
    """災害危険区域情報（XKT016）"""
    is_designated: bool = False  # 指定区域内かどうか
    area_name: Optional[str] = None
    raw_data: Optional[dict] = None


class LargeFillArea(BaseModel):
    """大規模盛土造成地情報（XKT020）"""
    is_designated: bool = False
    fill_type: Optional[str] = None  # 谷埋め型、腹付け型など
    raw_data: Optional[dict] = None


class LandslideArea(BaseModel):
    """地すべり防止地区情報（XKT021）"""
    is_designated: bool = False
    area_name: Optional[str] = None
    raw_data: Optional[dict] = None


class SteepSlopeArea(BaseModel):
    """急傾斜地崩壊危険区域情報（XKT022）"""
    is_designated: bool = False
    area_name: Optional[str] = None
    raw_data: Optional[dict] = None


# ========================================
# 価格参考情報
# ========================================

class LandTransaction(BaseModel):
    """取引価格情報（XIT001）"""
    transaction_price: Optional[int] = None  # 取引価格
    price_per_sqm: Optional[int] = None  # 平米単価
    transaction_date: Optional[str] = None  # 取引時期
    land_type: Optional[str] = None  # 土地の種類
    area_sqm: Optional[float] = None  # 面積
    structure: Optional[str] = None  # 構造
    building_year: Optional[str] = None  # 建築年
    address: Optional[str] = None  # 所在地
    raw_data: Optional[dict] = None


class LandPrice(BaseModel):
    """地価公示・地価調査情報（XPT002）"""
    price_per_sqm: Optional[int] = None  # 価格（円/㎡）
    survey_year: Optional[int] = None  # 調査年
    survey_type: Optional[str] = None  # 公示/調査
    address: Optional[str] = None  # 所在地
    land_use: Optional[str] = None  # 利用現況
    raw_data: Optional[dict] = None


# ========================================
# 統合結果モデル
# ========================================

class ReinfolibResult(BaseModel):
    """不動産情報ライブラリからの取得結果"""
    # 法規制情報
    land_use_zone: Optional[LandUseZone] = None
    fire_prevention_area: Optional[FirePreventionArea] = None
    city_planning_area: Optional[CityPlanningArea] = None
    location_optimization_area: Optional[LocationOptimizationArea] = None

    # リスク情報
    liquefaction_risk: Optional[LiquefactionRisk] = None
    disaster_risk_area: Optional[DisasterRiskArea] = None
    large_fill_area: Optional[LargeFillArea] = None
    landslide_area: Optional[LandslideArea] = None
    steep_slope_area: Optional[SteepSlopeArea] = None

    # 価格参考情報
    nearby_transactions: List[LandTransaction] = []
    nearby_land_prices: List[LandPrice] = []

    # エラー情報
    errors: List[str] = []
