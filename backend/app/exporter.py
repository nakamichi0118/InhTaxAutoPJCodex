"""Thin wrapper around CSV export utility."""
from __future__ import annotations

from typing import Dict

from src.export_csv import ASSET_COLUMNS, BANK_TRANSACTION_COLUMNS, build_csv, convert_assets_payload


def export_to_csv_strings(payload) -> Dict[str, str]:
    assets, bank_transactions = convert_assets_payload(payload)
    csv_map: Dict[str, str] = {}
    csv_map["assets.csv"] = build_csv(ASSET_COLUMNS, assets)
    if bank_transactions:
        csv_map["bank_transactions.csv"] = build_csv(BANK_TRANSACTION_COLUMNS, bank_transactions)
    return csv_map
