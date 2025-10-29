"""Utility to render inheritance asset CSV files from normalised JSON."""
from __future__ import annotations

import argparse
import csv
import json
import re
import uuid
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

ASSET_COLUMNS: Sequence[str] = (
    "record_id",
    "source_document",
    "asset_category",
    "asset_type",
    "owner_name",
    "asset_name",
    "location_prefecture",
    "location_municipality",
    "location_detail",
    "identifier_primary",
    "identifier_secondary",
    "valuation_basis",
    "valuation_currency",
    "valuation_amount",
    "valuation_date",
    "ownership_share",
    "notes",
)

BANK_TRANSACTION_COLUMNS: Sequence[str] = (
    "record_id",
    "transaction_id",
    "transaction_date",
    "value_date",
    "description",
    "withdrawal_amount",
    "deposit_amount",
    "balance",
    "line_confidence",
)

JAPANESE_ERA_BASE_YEAR = {
    "令和": 2018,
    "平成": 1988,
    "昭和": 1925,
    "大正": 1911,
}

UUID_NAMESPACE = uuid.uuid5(uuid.NAMESPACE_URL, "https://inhtaxautopj-codex/export")


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export inheritance CSV files from JSON inputs.")
    parser.add_argument("input", type=Path, help="Path to a JSON file or directory containing JSON files.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Directory where CSV files will be written (default: ./output).",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing CSV files if they already exist in the output directory.",
    )
    return parser.parse_args(argv)


@dataclass
class AssetRecord:
    data: Dict[str, Any]

    @property
    def record_id(self) -> str:
        existing = self.data.get("record_id")
        if existing:
            return str(existing)
        fingerprint_parts = [
            str(self.data.get("source_document", "")),
            str(self.data.get("asset_name", "")),
            json.dumps(self.data.get("identifiers", {}), ensure_ascii=False, sort_keys=True),
            str(self.data.get("category", "")),
        ]
        fingerprint = "|".join(fingerprint_parts)
        return str(uuid.uuid5(UUID_NAMESPACE, fingerprint))

    def to_asset_row(self) -> Dict[str, str]:
        location = self.data.get("location") or {}
        identifiers = self.data.get("identifiers") or {}
        valuation = self.data.get("valuation") or {}

        return {
            "record_id": self.record_id,
            "source_document": clean_text(self.data.get("source_document")),
            "asset_category": clean_text(self.data.get("category")),
            "asset_type": clean_text(self.data.get("type")),
            "owner_name": join_names(self.data.get("owner_name")),
            "asset_name": clean_text(self.data.get("asset_name")),
            "location_prefecture": clean_text(location.get("prefecture")),
            "location_municipality": clean_text(location.get("municipality")),
            "location_detail": clean_text(location.get("detail")),
            "identifier_primary": clean_text(identifiers.get("primary")),
            "identifier_secondary": clean_text(identifiers.get("secondary")),
            "valuation_basis": clean_text(valuation.get("basis")),
            "valuation_currency": clean_text(valuation.get("currency") or "JPY"),
            "valuation_amount": normalize_decimal(valuation.get("amount")),
            "valuation_date": normalize_date(valuation.get("date")),
            "ownership_share": normalize_decimal(self.data.get("ownership_share")),
            "notes": clean_text(self.data.get("notes")),
        }

    def iter_transactions(self) -> Iterator[Dict[str, str]]:
        transactions = self.data.get("transactions")
        if not transactions:
            return iter(())

        def _generator() -> Iterator[Dict[str, str]]:
            for index, txn in enumerate(transactions, start=1):
                if not isinstance(txn, dict):
                    continue
                transaction_id = f"{self.record_id}-{index:04d}"
                yield {
                    "record_id": self.record_id,
                    "transaction_id": transaction_id,
                    "transaction_date": normalize_date(txn.get("transaction_date")),
                    "value_date": normalize_date(txn.get("value_date")),
                    "description": clean_text(txn.get("description")),
                    "withdrawal_amount": normalize_decimal(txn.get("withdrawal_amount")),
                    "deposit_amount": normalize_decimal(txn.get("deposit_amount")),
                    "balance": normalize_decimal(txn.get("balance")),
                    "line_confidence": normalize_decimal(txn.get("confidence")),
                }

        return _generator()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def join_names(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple, set)):
        cleaned = [clean_text(v) for v in value if clean_text(v)]
        return ";".join(cleaned)
    return clean_text(value)


def normalize_decimal(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, (int, float, Decimal)):
        decimal_value = Decimal(str(value))
        return format_decimal(decimal_value)
    try:
        raw = str(value).replace(",", "").strip()
        if not raw:
            return ""
        decimal_value = Decimal(raw)
        return format_decimal(decimal_value)
    except (InvalidOperation, ValueError):
        return clean_text(value)


def format_decimal(value: Decimal) -> str:
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def normalize_date(value: Any) -> str:
    if value is None:
        return ""
    text = clean_text(value)
    if not text:
        return ""
    try:
        return datetime.fromisoformat(text).date().isoformat()
    except ValueError:
        pass

    digits_only = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", text)
    if digits_only:
        year, month, day = digits_only.groups()
        try:
            dt = datetime(int(year), int(month), int(day))
            return dt.date().isoformat()
        except ValueError:
            return text

    sep_match = re.fullmatch(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", text)
    if sep_match:
        year, month, day = map(int, sep_match.groups())
        try:
            return datetime(year, month, day).date().isoformat()
        except ValueError:
            return text

    kanji_match = re.fullmatch(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    if kanji_match:
        year, month, day = map(int, kanji_match.groups())
        try:
            return datetime(year, month, day).date().isoformat()
        except ValueError:
            return text

    era_match = re.fullmatch(r"(令和|平成|昭和|大正)(\d{1,2})年?(\d{1,2})月?(\d{1,2})日?", text)
    if era_match:
        era, era_year, month, day = era_match.groups()
        base_year = JAPANESE_ERA_BASE_YEAR[era]
        try:
            gregorian_year = base_year + int(era_year)
            return datetime(gregorian_year, int(month), int(day)).date().isoformat()
        except ValueError:
            return text

    return text


def load_assets(path: Path) -> List[AssetRecord]:
    if not path.exists():
        raise FileNotFoundError(f"Input path not found: {path}")

    assets: List[AssetRecord] = []

    def _parse_json_file(file_path: Path) -> None:
        try:
            with file_path.open("r", encoding="utf-8-sig") as fh:
                payload = json.load(fh)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Failed to parse JSON file: {file_path}") from exc
        items = payload.get("assets")
        if items is None:
            raise ValueError(f"JSON file does not contain 'assets' key: {file_path}")
        for item in items:
            if not isinstance(item, dict):
                raise ValueError(f"Asset entry is not an object in file: {file_path}")
            assets.append(AssetRecord(item))

    if path.is_file():
        _parse_json_file(path)
    else:
        for file_path in sorted(path.glob("*.json")):
            _parse_json_file(file_path)

    return assets


def ensure_output_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def write_csv(path: Path, columns: Sequence[str], rows: Iterable[Dict[str, str]], *, overwrite: bool) -> bool:
    if path.exists() and not overwrite:
        raise FileExistsError(f"Output file already exists: {path}")
    wrote_any = False
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(columns), extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
            wrote_any = True
    return wrote_any


def export_csv_files(input_path: Path, output_dir: Path, *, overwrite: bool = False) -> Dict[str, Path]:
    payload_assets = load_assets(input_path)
    ensure_output_dir(output_dir)

    asset_rows = [asset.to_asset_row() for asset in payload_assets]
    assets_csv_path = output_dir / "assets.csv"
    write_csv(assets_csv_path, ASSET_COLUMNS, asset_rows, overwrite=overwrite)

    transaction_rows: List[Dict[str, str]] = []
    for asset in payload_assets:
        transaction_rows.extend(list(asset.iter_transactions()))

    exported_files: Dict[str, Path] = {"assets": assets_csv_path}

    if transaction_rows:
        bank_csv_path = output_dir / "bank_transactions.csv"
        write_csv(bank_csv_path, BANK_TRANSACTION_COLUMNS, transaction_rows, overwrite=overwrite)
        exported_files["bank_transactions"] = bank_csv_path

    return exported_files


def convert_assets_payload(payload: Dict[str, Any]) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    if not isinstance(payload, dict):
        raise ValueError("Payload must be a dictionary with an 'assets' key")
    assets = payload.get("assets")
    if not isinstance(assets, list):
        raise ValueError("Payload missing 'assets' array")

    records = [AssetRecord(item) for item in assets if isinstance(item, dict)]
    asset_rows = [record.to_asset_row() for record in records]
    transaction_rows: List[Dict[str, str]] = []
    for record in records:
        transaction_rows.extend(list(record.iter_transactions()))
    return asset_rows, transaction_rows


def build_csv(columns: Sequence[str], rows: Iterable[Dict[str, str]]) -> str:
    lines = [",".join(columns)]
    for row in rows:
        values = [csv_escape(row.get(column, "")) for column in columns]
        lines.append(",".join(values))
    return "\r\n".join(lines) + "\r\n"


def csv_escape(value: str) -> str:
    text = str(value)
    if text == "":
        return ""
    if any(ch in text for ch in ('"', ',', '\n')):
        return '"' + text.replace('"', '""') + '"'
    return text


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = parse_args(argv)
    exported = export_csv_files(args.input, args.output_dir, overwrite=args.force)
    relative = {key: str(path) for key, path in exported.items()}
    print(json.dumps({"exported": relative}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
