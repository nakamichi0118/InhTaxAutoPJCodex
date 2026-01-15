"""Thin wrapper around CSV export utility."""
from __future__ import annotations

import re
from typing import Any, Dict, List

from src.export_csv import (
    ASSET_EXPORT_COLUMNS,
    BUILDING_EXPORT_COLUMNS,
    LAND_EXPORT_COLUMNS,
    TRANSACTION_EXPORT_COLUMNS,
    build_csv,
    convert_assets_payload,
)


def export_to_csv_strings(payload) -> Dict[str, str]:
    assets, bank_transactions = convert_assets_payload(payload)
    csv_map: Dict[str, str] = {}
    include_assets = any(row.get("asset_category") != "bank_deposit" for row in assets)
    if include_assets:
        csv_map["assets.csv"] = build_csv(ASSET_EXPORT_COLUMNS, assets)
    if bank_transactions:
        csv_map["bank_transactions.csv"] = build_csv(TRANSACTION_EXPORT_COLUMNS, bank_transactions)

    # 土地・家屋の専用CSVを出力
    land_rows = _extract_land_rows(assets)
    building_rows = _extract_building_rows(assets)

    if land_rows:
        csv_map["land.csv"] = build_csv(LAND_EXPORT_COLUMNS, land_rows)
    if building_rows:
        csv_map["building.csv"] = build_csv(BUILDING_EXPORT_COLUMNS, building_rows)

    return csv_map


def _extract_land_rows(assets: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """Extract land-specific rows from assets."""
    rows: List[Dict[str, str]] = []
    for asset in assets:
        if asset.get("asset_category") != "land":
            continue
        notes = asset.get("notes", "") or ""
        rows.append({
            "location_municipality": asset.get("location_municipality", ""),
            "location_detail": asset.get("location_detail", ""),
            "land_category_tax": _extract_from_notes(notes, "課税地目"),
            "land_category_registry": "",  # 手入力欄
            "area": _extract_from_notes(notes, "地積"),
            "valuation_amount": str(int(asset.get("valuation_amount") or 0)) if asset.get("valuation_amount") else "",
            "ownership_share": "",  # 手入力欄
            "notes": _clean_notes_for_export(notes),
        })
    return rows


def _extract_building_rows(assets: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """Extract building-specific rows from assets."""
    rows: List[Dict[str, str]] = []
    for asset in assets:
        if asset.get("asset_category") != "building":
            continue
        notes = asset.get("notes", "") or ""
        rows.append({
            "location_municipality": asset.get("location_municipality", ""),
            "location_detail": asset.get("location_detail", ""),
            "structure": _extract_from_notes(notes, "構造"),
            "floors": _extract_from_notes(notes, "階数"),
            "area": _extract_from_notes(notes, "床面積"),
            "built_year": _extract_from_notes(notes, "建築年"),
            "valuation_amount": str(int(asset.get("valuation_amount") or 0)) if asset.get("valuation_amount") else "",
            "ownership_share": "",  # 手入力欄
            "notes": _clean_notes_for_export(notes),
        })
    return rows


def _extract_from_notes(notes: str, key: str) -> str:
    """Extract a value from notes by key pattern (e.g., '課税地目: 宅地')."""
    pattern = rf"{key}[:：]?\s*(.+?)(?:\n|$)"
    match = re.search(pattern, notes)
    if match:
        value = match.group(1).strip()
        # Remove unit suffixes for cleaner values
        value = re.sub(r"[㎡]$", "", value)
        return value
    return ""


def _clean_notes_for_export(notes: str) -> str:
    """Remove structured fields from notes, keeping only free-text content."""
    # Remove common structured fields
    patterns = [
        r"課税地目[:：]?\s*.+?(?:\n|$)",
        r"地積[:：]?\s*.+?(?:\n|$)",
        r"構造[:：]?\s*.+?(?:\n|$)",
        r"階数[:：]?\s*.+?(?:\n|$)",
        r"床面積[:：]?\s*.+?(?:\n|$)",
        r"建築年[:：]?\s*.+?(?:\n|$)",
    ]
    result = notes
    for pattern in patterns:
        result = re.sub(pattern, "", result)
    return result.strip()
