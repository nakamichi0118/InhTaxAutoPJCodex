"""Parser for 名寄帳 (property tax roll) documents."""
from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from ..models import AssetRecord

logger = logging.getLogger(__name__)

JSON_PATTERN = re.compile(r"\{.*\}", re.DOTALL)


@dataclass
class NayoseProperty:
    """Represents a single property extracted from 名寄帳."""
    property_type: str  # "land" or "building"
    location: Optional[str] = None
    lot_number: Optional[str] = None
    land_category: Optional[str] = None
    area: Optional[float] = None
    valuation_amount: Optional[float] = None
    structure: Optional[str] = None
    built_year: Optional[str] = None
    floors: Optional[str] = None
    notes: Optional[str] = None


def parse_nayose_response(
    response_text: str,
    source_name: str,
    municipality: Optional[str] = None,
) -> List[AssetRecord]:
    """Parse Gemini response for 名寄帳 into AssetRecord list.

    Args:
        response_text: Raw text response from Gemini API
        source_name: Name of the source document
        municipality: Optional municipality name (extracted from response if not provided)

    Returns:
        List of AssetRecord objects for each property
    """
    # Extract JSON from response
    match = JSON_PATTERN.search(response_text)
    if not match:
        logger.warning("No JSON found in nayose response")
        return []

    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError as exc:
        logger.error("Failed to parse nayose JSON: %s", exc)
        return []

    # Get municipality from response if not provided
    if not municipality:
        municipality = data.get("municipality", "")

    properties = data.get("properties", [])
    if not properties:
        logger.warning("No properties found in nayose response")
        return []

    assets: List[AssetRecord] = []
    for prop in properties:
        asset = _build_asset_from_property(prop, source_name, municipality)
        if asset:
            assets.append(asset)

    return assets


def _build_asset_from_property(
    prop: Dict[str, Any],
    source_name: str,
    municipality: str,
) -> Optional[AssetRecord]:
    """Convert a property dict to an AssetRecord."""
    property_type = prop.get("property_type", "").lower()

    if property_type == "land":
        return _build_land_asset(prop, source_name, municipality)
    elif property_type == "building":
        return _build_building_asset(prop, source_name, municipality)
    else:
        logger.warning("Unknown property type: %s", property_type)
        return None


def _build_land_asset(
    prop: Dict[str, Any],
    source_name: str,
    municipality: str,
) -> AssetRecord:
    """Build AssetRecord for land (土地)."""
    location = prop.get("location", "")
    lot_number = prop.get("lot_number", "")
    land_category = prop.get("land_category", "")
    area = _parse_number(prop.get("area"))
    valuation = _parse_number(prop.get("valuation_amount"))
    notes_parts = []

    # Build notes with area and land category
    if land_category:
        notes_parts.append(f"課税地目: {land_category}")
    if area is not None:
        notes_parts.append(f"地積: {area}㎡")
    if prop.get("notes"):
        notes_parts.append(prop["notes"])

    # Map Japanese land categories to standardized types
    asset_type = _normalize_land_type(land_category)

    return AssetRecord(
        category="land",
        type=asset_type,
        source_document=source_name,
        location_municipality=municipality,
        location_detail=f"{location} {lot_number}".strip() if location else lot_number,
        identifier_primary=lot_number,
        valuation_basis="固定資産税評価額",
        valuation_amount=valuation,
        notes="\n".join(notes_parts) if notes_parts else None,
    )


def _build_building_asset(
    prop: Dict[str, Any],
    source_name: str,
    municipality: str,
) -> AssetRecord:
    """Build AssetRecord for building (家屋)."""
    location = prop.get("location", "")
    lot_number = prop.get("lot_number", "")  # 家屋番号
    structure = prop.get("structure", "")
    area = _parse_number(prop.get("area"))
    valuation = _parse_number(prop.get("valuation_amount"))
    built_year = prop.get("built_year", "")
    floors = prop.get("floors", "")

    notes_parts = []
    if structure:
        notes_parts.append(f"構造: {structure}")
    if floors:
        notes_parts.append(f"階数: {floors}")
    if area is not None:
        notes_parts.append(f"床面積: {area}㎡")
    if built_year:
        notes_parts.append(f"建築年: {built_year}")
    if prop.get("notes"):
        notes_parts.append(prop["notes"])

    # Map structure to standardized building type
    asset_type = _normalize_building_type(structure)

    return AssetRecord(
        category="building",
        type=asset_type,
        source_document=source_name,
        location_municipality=municipality,
        location_detail=f"{location} 家屋番号{lot_number}".strip() if location else f"家屋番号{lot_number}",
        identifier_primary=lot_number,
        valuation_basis="固定資産税評価額",
        valuation_amount=valuation,
        notes="\n".join(notes_parts) if notes_parts else None,
    )


def _parse_number(value: Any) -> Optional[float]:
    """Parse a number from various input formats."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        # Remove commas, spaces, and unit suffixes
        cleaned = value.replace(",", "").replace(" ", "").replace("㎡", "").replace("円", "")
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def _normalize_land_type(land_category: Optional[str]) -> str:
    """Normalize Japanese land category to standardized type code."""
    if not land_category:
        return "other"

    category_map = {
        "宅地": "residential",
        "田": "agricultural",
        "畑": "agricultural",
        "山林": "forest",
        "原野": "wasteland",
        "雑種地": "miscellaneous",
        "池沼": "pond",
        "牧場": "pasture",
        "鉱泉地": "mineral_spring",
        "墓地": "cemetery",
        "境内地": "temple_grounds",
        "運河用地": "canal",
        "水道用地": "waterworks",
        "用悪水路": "drainage",
        "ため池": "reservoir",
        "堤": "embankment",
        "井溝": "well_ditch",
        "保安林": "protection_forest",
        "公衆用道路": "public_road",
        "公園": "park",
    }

    for key, value in category_map.items():
        if key in land_category:
            return value

    return "other"


def _normalize_building_type(structure: Optional[str]) -> str:
    """Normalize Japanese building structure/usage to standardized type code."""
    if not structure:
        return "other"

    # Check for building usage keywords
    usage_map = {
        "居宅": "residence",
        "住宅": "residence",
        "共同住宅": "apartment",
        "店舗": "store",
        "事務所": "office",
        "倉庫": "warehouse",
        "工場": "factory",
        "車庫": "garage",
        "物置": "storage",
        "作業所": "workshop",
        "診療所": "clinic",
        "病院": "hospital",
        "旅館": "inn",
        "ホテル": "hotel",
        "劇場": "theater",
        "映画館": "cinema",
    }

    for key, value in usage_map.items():
        if key in structure:
            return value

    return "other"
