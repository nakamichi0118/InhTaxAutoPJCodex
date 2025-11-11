"""Thin wrapper around CSV export utility."""
from __future__ import annotations

from typing import Dict

from src.export_csv import (
    ASSET_EXPORT_COLUMNS,
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
    return csv_map
