"""Pydantic models for JON API."""
from __future__ import annotations

from typing import Any, Dict, List, Literal, Optional
from pydantic import BaseModel, Field


class LocationResult(BaseModel):
    """位置特定APIの結果"""
    lat: float
    long: float
    locating_level: int = 0
    located_by: str = ""
    v1_code: str
    pref: str = ""
    city: str = ""
    large_section: str = ""
    small_section: str = ""
    number: str = ""
    address_type_result: str = ""
    raw_response: Dict[str, Any] = Field(default_factory=dict)

    @property
    def accuracy_label(self) -> str:
        """
        精度レベルの説明ラベル
        JON API locating_level: 1=完全一致, 2=号・枝番不一致, 3=番地不一致,
        4=丁目・小字不一致, 5=大字・町名不一致, 6=都道府県一致のみ
        ※数字が小さいほど高精度
        """
        if self.locating_level == 0:
            return "精度不明"
        elif self.locating_level <= 2:
            return "高精度"
        elif self.locating_level <= 4:
            return "中精度"
        else:
            return "低精度"

    @property
    def is_high_accuracy(self) -> bool:
        """登記取得に十分な精度か（レベル1-2：完全一致または号・枝番のみ不一致）"""
        return 1 <= self.locating_level <= 2


class BuildingNumber(BaseModel):
    """家屋番号"""
    building_number: str
    v1_code: str = ""
    pref: str = ""
    city: str = ""
    large_section: str = ""
    small_section: str = ""


class RegistrationResult(BaseModel):
    """登記取得APIの結果"""
    pdf_id: int
    pdf_url: str
    v1_code: str
    number: str
    number_type: int
    pdf_type: int
    based_at: str
    raw_response: Dict[str, Any] = Field(default_factory=dict)


class RosenImageResult(BaseModel):
    """路線価図APIの結果"""
    image_base64: str
    lat: float
    long: float


class AnalyzeResult(BaseModel):
    """登記解析APIの結果"""
    pdf_id: int
    status: str
    data: Optional[Dict[str, Any]] = None
    raw_response: Dict[str, Any] = Field(default_factory=dict)


# Batch processing models
class JonBatchItem(BaseModel):
    """バッチ処理の個別アイテム"""
    id: str
    address: str
    property_type: Literal["land", "building"]
    acquisitions: List[Literal["google_map", "rosenka", "touki", "kozu", "chiseki", "tatemono"]]


class JonBatchRequest(BaseModel):
    """バッチ処理リクエスト"""
    properties: List[JonBatchItem]
    skip_kozu_for_bairitsu: bool = False  # 倍率地域の場合は公図をスキップ


class JonBatchItemResult(BaseModel):
    """バッチ処理の個別結果"""
    id: str
    status: Literal["pending", "processing", "completed", "failed"]
    location: Optional[LocationResult] = None
    locating_level: int = 0  # 位置特定精度レベル
    accuracy_label: str = ""  # 精度ラベル（高精度/中精度/低精度）
    accuracy_warning: Optional[str] = None  # 精度警告メッセージ
    google_map_url: Optional[str] = None
    rosenka_image: Optional[str] = None  # base64
    rosenka_urls: List[str] = Field(default_factory=list)  # 国税庁路線価図URL
    is_bairitsu: bool = False  # 倍率地域かどうか
    registration: Optional[RegistrationResult] = None  # 登記簿PDF
    kozu_pdf_url: Optional[str] = None  # 公図PDF URL
    chiseki_pdf_url: Optional[str] = None  # 地積測量図PDF URL
    tatemono_pdf_url: Optional[str] = None  # 建物図面PDF URL
    analyze_result: Optional[Dict[str, Any]] = None
    error: Optional[str] = None


class JonBatchResponse(BaseModel):
    """バッチ処理レスポンス"""
    batch_id: str
    status: Literal["pending", "processing", "completed", "failed"]
    total: int
    completed: int
    results: List[JonBatchItemResult]


# Request models for individual API calls
class LocatingRequest(BaseModel):
    """位置特定APIリクエスト"""
    address: str
    address_type: Literal["地番", "家屋番号", "住居表示番号", "不明"] = "不明"


class RosenkaRequest(BaseModel):
    """路線価図APIリクエスト"""
    lat: float
    long: float
    length: int = Field(default=250, ge=1, le=250)
    width: int = Field(default=250, ge=1, le=250)


class RegistrationRequest(BaseModel):
    """登記取得APIリクエスト"""
    v1_code: str
    number: str
    number_type: Literal[1, 2] = 1  # 1=地番, 2=家屋番号
    pdf_type: Literal[1, 2, 3, 4, 5, 6] = 1  # 1=全部事項, 2=所有者事項, 3=地図, 4=地積測量図, 5=地役権図面, 6=建物図面


class AnalyzeRequest(BaseModel):
    """登記解析APIリクエスト"""
    pdf_id: int
    fields: Optional[List[str]] = None


# Force registration request
class ForceRegistrationRequest(BaseModel):
    """強制登記取得リクエスト（精度チェックをスキップ）"""
    batch_id: str
    item_id: str
    pdf_types: List[Literal[1, 3, 4, 6]] = Field(default_factory=lambda: [1])  # 1=登記簿, 3=公図, 4=地積測量図, 6=建物図面


class ForceRegistrationResponse(BaseModel):
    """強制登記取得レスポンス"""
    success: bool
    item_id: str
    registration_pdf_url: Optional[str] = None
    kozu_pdf_url: Optional[str] = None
    chiseki_pdf_url: Optional[str] = None
    tatemono_pdf_url: Optional[str] = None
    error: Optional[str] = None


# Status response
class JonStatusResponse(BaseModel):
    """JON API設定状態"""
    configured: bool
    has_touki_credentials: bool
    api_base_url: str
