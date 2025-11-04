from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Dict, Iterable, List, Optional

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import HttpResponseError

from .models import AssetRecord, TransactionLine


@dataclass
class AzureAnalysisResult:
    raw_lines: List[str]
    assets: List[AssetRecord]


HEADER_ALIASES: Dict[str, str] = {
    "transaction_date": "取引日",
    "description": "摘要",
    "withdrawal": "支払金額",
    "deposit": "入金金額",
    "balance": "残高",
}

HEADER_KEYWORDS: Dict[str, Iterable[str]] = {
    "transaction_date": ("取引日", "年月日", "日付"),
    "description": ("摘要", "内容", "件名"),
    "withdrawal": ("支払", "出金", "支払金額"),
    "deposit": ("入金", "預入", "入金金額"),
    "balance": ("残高", "差引残高", "残高金額", "残", "高"),
}


class AzureAnalysisError(RuntimeError):
    pass


class AzureTransactionAnalyzer:
    def __init__(self, *, endpoint: str, api_key: str) -> None:
        credential = AzureKeyCredential(api_key)
        self.client = DocumentAnalysisClient(endpoint=endpoint, credential=credential)

    def analyze_pdf(self, pdf_bytes: bytes, *, source_name: str) -> AzureAnalysisResult:
        try:
            poller = self.client.begin_analyze_document("prebuilt-document", pdf_bytes)
            result = poller.result()
        except HttpResponseError as exc:
            raise AzureAnalysisError(str(exc)) from exc

        raw_lines = result.content.splitlines() if result.content else []
        transactions: List[TransactionLine] = []
        for table in result.tables or []:
            transactions.extend(_extract_transactions_from_table(table))

        asset = AssetRecord(
            category="bank_deposit",
            type="transaction_history",
            source_document=source_name,
            asset_name="預金取引推移表",
            transactions=transactions,
        )

        return AzureAnalysisResult(raw_lines=raw_lines, assets=[asset])


def _extract_transactions_from_table(table, pages=None) -> List[TransactionLine]:
    if not table.cells:
        return []

    rows: Dict[int, Dict[int, str]] = {}
    for cell in table.cells:
        span = getattr(cell, "column_span", 1) or 1
        row = rows.setdefault(cell.row_index, {})
        for offset in range(span):
            row[cell.column_index + offset] = cell.content or ""

    header_row = min(rows.keys())
    header_map = _map_headers(rows.get(header_row, {}))
    if not header_map:
        return []

    transactions: List[TransactionLine] = []
    for row_index in sorted(rows.keys()):
        if row_index == header_row:
            continue
        row_cells = rows[row_index]
        values: Dict[str, List[Optional[str]]] = {}
        for column_index, field in header_map.items():
            values.setdefault(field, []).append(row_cells.get(column_index))
        txn = _build_transaction(values)
        if txn:
            transactions.append(txn)
    return transactions


def _map_headers(columns: Dict[int, str]) -> Dict[int, str]:
    mapping: Dict[int, str] = {}
    for column_index, text in columns.items():
        normalized = _normalize_header(text)
        for field, keywords in HEADER_KEYWORDS.items():
            if any(keyword in normalized for keyword in keywords):
                mapping[column_index] = field
                break
    if not mapping and columns:
        sorted_items = sorted(columns.items(), key=lambda item: item[0])
        slots = ["transaction_date", "description", "withdrawal", "deposit", "balance"]
        for column_index, _ in sorted_items:
            if not slots:
                break
            mapping[column_index] = slots.pop(0)
    else:
        last_field: Optional[str] = None
        for column_index in sorted(columns.keys()):
            if column_index in mapping:
                last_field = mapping[column_index]
                continue
            text = _normalize_header(columns.get(column_index, ""))
            if last_field in {"withdrawal", "deposit"} and text in {"", "金額"}:
                mapping[column_index] = last_field
            elif last_field == "balance" and text in {"", "金額", "残", "高"}:
                mapping[column_index] = last_field
    return mapping


def _normalize_header(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def _merge_parts(parts: Optional[List[Optional[str]]]) -> Optional[str]:
    if not parts:
        return None
    merged = " ".join(part for part in parts if part)
    merged = merged.strip()
    return merged or None


def _build_transaction(values: Dict[str, List[Optional[str]]]) -> Optional[TransactionLine]:
    date_text = _merge_parts(values.get("transaction_date"))
    description = _clean_text(_merge_parts(values.get("description")))
    withdrawal = _parse_amount(_merge_parts(values.get("withdrawal")))
    deposit = _parse_amount(_merge_parts(values.get("deposit")))
    balance = _parse_amount(_merge_parts(values.get("balance")))

    if not any([date_text, description, withdrawal, deposit, balance]):
        return None

    transaction_date = _parse_date(date_text)
    return TransactionLine(
        transaction_date=transaction_date,
        description=description or None,
        withdrawal_amount=withdrawal,
        deposit_amount=deposit,
        balance=balance,
    )


def _clean_text(text: Optional[str]) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", text.strip())


def _parse_amount(text: Optional[str]) -> Optional[float]:
    if not text:
        return None
    cleaned = text.replace("円", "").replace("¥", "")
    cleaned = cleaned.replace(",", "").replace(" ", "")
    if "." in cleaned and cleaned.count(".") == 1 and len(cleaned.replace(".", "")) > 4:
        cleaned = cleaned.replace(".", "")
    cleaned = cleaned.replace("△", "-")
    cleaned = cleaned.strip()
    if not cleaned or cleaned == "-":
        return None
    negative = cleaned.startswith("(") and cleaned.endswith(")")
    cleaned = cleaned.strip("()").strip()
    try:
        value = Decimal(cleaned)
    except (InvalidOperation, ValueError):
        return None
    if negative:
        value = -value
    return float(value)


DATE_PATTERN = re.compile(r"(\d{4}|\d{2})-(\d{1,2})-(\d{1,2})")


def _parse_date(text: Optional[str]) -> Optional[str]:
    if not text:
        return None
    cleaned = text.strip().replace(":", "-").replace("/", "-").replace(",", "-")
    cleaned = re.sub(r"\s+", "-", cleaned)
    match = DATE_PATTERN.search(cleaned)
    if not match:
        return None
    year, month, day = match.groups()
    year_value = int(year)
    if year_value < 100:
        if 6 <= year_value <= 31:
            year_value = 1988 + year_value
        elif 32 <= year_value <= 64:
            year_value = 1925 + year_value
        else:
            year_value += 2000 if year_value < 50 else 1900
    try:
        return f"{year_value:04d}-{int(month):02d}-{int(day):02d}"
    except ValueError:
        return None
