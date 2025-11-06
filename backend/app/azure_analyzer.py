from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, Iterable, List, Optional

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

BALANCE_ONLY_KEYWORDS = (
    "繰越",
    "繰り越し",
    "前日残高",
)

DESCRIPTION_CLEANUPS = (
    (":selected:", ""),
)


class AzureAnalysisError(RuntimeError):
    pass


class AzureTransactionAnalyzer:
    def __init__(self, *, endpoint: str, api_key: str) -> None:
        credential = AzureKeyCredential(api_key)
        self.client = DocumentAnalysisClient(endpoint=endpoint, credential=credential)

    def analyze_pdf(self, pdf_bytes: bytes, *, source_name: str, date_format: str = "auto") -> AzureAnalysisResult:
        try:
            poller = self.client.begin_analyze_document("prebuilt-document", pdf_bytes)
            result = poller.result()
        except HttpResponseError as exc:
            raise AzureAnalysisError(str(exc)) from exc

        raw_lines = result.content.splitlines() if result.content else []
        raw_transactions: List[TransactionLine] = []
        for table in result.tables or []:
            raw_transactions.extend(_extract_transactions_from_table(table, date_format=date_format))

        transactions = _post_process_transactions(raw_transactions)

        asset = AssetRecord(
            category="bank_deposit",
            type="transaction_history",
            source_document=source_name,
            asset_name="預金取引推移表",
            transactions=transactions,
        )

        return AzureAnalysisResult(raw_lines=raw_lines, assets=[asset])


def _extract_transactions_from_table(table, *, date_format: str) -> List[TransactionLine]:
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
        txn = _build_transaction(values, date_format=date_format)
        if txn:
            transactions.append(txn)
    return transactions



def _post_process_transactions(raw_transactions: List[TransactionLine]) -> List[TransactionLine]:
    enriched: List[TransactionLine] = []
    last_balance: Optional[float] = None

    for txn in raw_transactions:
        adjusted = _impute_missing_values(txn, last_balance)
        if not adjusted:
            continue

        base_updates: Dict[str, Any] = {}

        desc_cleaned = _clean_description(adjusted.description)
        if desc_cleaned != (adjusted.description or ""):
            base_updates["description"] = desc_cleaned or None

        if not adjusted.transaction_date and enriched:
            base_updates.setdefault("transaction_date", enriched[-1].transaction_date)

        candidate = adjusted.model_copy(update=base_updates) if base_updates else adjusted

        normalized_balance = _normalize_amount(candidate.balance, reference=last_balance)
        normalized_withdrawal = _normalize_amount(candidate.withdrawal_amount, reference=last_balance)
        normalized_deposit = _normalize_amount(candidate.deposit_amount, reference=last_balance)

        normalization_updates: Dict[str, Any] = {}
        if normalized_balance != candidate.balance:
            normalization_updates["balance"] = normalized_balance
        if normalized_withdrawal != candidate.withdrawal_amount:
            normalization_updates["withdrawal_amount"] = normalized_withdrawal
        if normalized_deposit != candidate.deposit_amount:
            normalization_updates["deposit_amount"] = normalized_deposit

        if desc_cleaned and any(keyword in desc_cleaned for keyword in BALANCE_ONLY_KEYWORDS):
            normalization_updates["withdrawal_amount"] = None
            normalization_updates["deposit_amount"] = None

        candidate = candidate.model_copy(update=normalization_updates) if normalization_updates else candidate

        balance_value = candidate.balance
        if balance_value is None and last_balance is not None:
            projected = last_balance
            if candidate.withdrawal_amount is not None:
                projected -= candidate.withdrawal_amount
            if candidate.deposit_amount is not None:
                projected += candidate.deposit_amount
            if abs(projected - last_balance) > 1e-6:
                candidate = candidate.model_copy(update={"balance": projected})
                balance_value = projected

        delta_updates: Dict[str, Any] = {}
        if balance_value is not None and last_balance is not None:
            delta = round(balance_value - last_balance, 2)
            if abs(delta) <= 0.01:
                delta = 0.0
            expected_withdrawal: Optional[float] = None
            expected_deposit: Optional[float] = None
            if delta > 0:
                expected_deposit = float(delta)
            elif delta < 0:
                expected_withdrawal = float(-delta)

            if _needs_amount_override(candidate, expected_withdrawal, expected_deposit):
                delta_updates["withdrawal_amount"] = expected_withdrawal
                delta_updates["deposit_amount"] = expected_deposit

        candidate = candidate.model_copy(update=delta_updates) if delta_updates else candidate

        if not any(
            [
                candidate.description,
                candidate.withdrawal_amount,
                candidate.deposit_amount,
                candidate.balance,
            ]
        ):
            continue

        enriched.append(candidate)
        if candidate.balance is not None:
            last_balance = candidate.balance

    return enriched



def _impute_missing_values(txn: TransactionLine, last_balance: Optional[float]) -> Optional[TransactionLine]:
    data = txn.model_dump()
    withdrawal = data.get("withdrawal_amount")
    deposit = data.get("deposit_amount")
    balance = data.get("balance")

    if balance is not None and last_balance is not None and withdrawal is None and deposit is None:
        delta = round(balance - last_balance, 2)
        if abs(delta) >= 0.01:
            if delta > 0:
                deposit = float(delta)
            else:
                withdrawal = float(-delta)

    if balance is None and last_balance is not None:
        projected = last_balance
        if withdrawal is not None:
            projected -= withdrawal
        if deposit is not None:
            projected += deposit
        if abs(projected - last_balance) >= 0.01:
            balance = projected

    if withdrawal is None and deposit is None and balance is None:
        return None

    data.update(
        {
            "withdrawal_amount": withdrawal,
            "deposit_amount": deposit,
            "balance": balance,
        }
    )
    return TransactionLine(**data)


def _normalize_amount(value: Optional[float], *, reference: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    normalized = float(value)
    # Trim obviously duplicated magnitude (e.g. 10^13)
    while abs(normalized) >= 1_000_000_000:
        normalized /= 10
    if reference is not None and abs(reference) > 0:
        limit = max(abs(reference) * 5, 10_000_000)
        while abs(normalized) > limit:
            normalized /= 10
    return round(normalized, 2)


def _needs_amount_override(
    candidate: TransactionLine,
    expected_withdrawal: Optional[float],
    expected_deposit: Optional[float],
) -> bool:
    current_withdrawal = None if candidate.withdrawal_amount is None else round(candidate.withdrawal_amount, 2)
    current_deposit = None if candidate.deposit_amount is None else round(candidate.deposit_amount, 2)

    if expected_withdrawal is None and expected_deposit is None:
        return False

    if expected_withdrawal is not None:
        if current_withdrawal is None or abs(current_withdrawal - expected_withdrawal) > 0.5:
            return True
        if current_deposit not in (None, 0.0):
            return True

    if expected_deposit is not None:
        if current_deposit is None or abs(current_deposit - expected_deposit) > 0.5:
            return True
        if current_withdrawal not in (None, 0.0):
            return True

    if expected_withdrawal is None and current_withdrawal not in (None, 0.0):
        return True
    if expected_deposit is None and current_deposit not in (None, 0.0):
        return True

    return False


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


def _build_transaction(values: Dict[str, List[Optional[str]]], *, date_format: str) -> Optional[TransactionLine]:
    date_text = _merge_parts(values.get("transaction_date"))
    description = _clean_text(_merge_parts(values.get("description")))
    withdrawal = _parse_amount(_merge_parts(values.get("withdrawal")))
    deposit = _parse_amount(_merge_parts(values.get("deposit")))
    balance = _parse_amount(_merge_parts(values.get("balance")))

    if not any([date_text, description, withdrawal, deposit, balance]):
        return None

    if withdrawal is None and deposit is None and balance is None:
        return None

    transaction_date = _parse_date(date_text, date_format=date_format)
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


def _clean_description(text: Optional[str]) -> str:
    cleaned = _clean_text(text)
    if not cleaned:
        return ""
    for target, replacement in DESCRIPTION_CLEANUPS:
        cleaned = cleaned.replace(target, replacement)
    return cleaned.strip()


def _parse_amount(text: Optional[str]) -> Optional[float]:
    if not text:
        return None
    cleaned = text.replace("円", "").replace("¥", "")
    cleaned = cleaned.replace(",", "").replace(" ", "")
    if "." in cleaned and len(cleaned.replace(".", "")) > 4:
        cleaned = cleaned.replace(".", "")
    cleaned = cleaned.replace("△", "-")
    cleaned = cleaned.strip()
    if not cleaned or cleaned == "-":
        return None
    negative = cleaned.startswith("(") and cleaned.endswith(")")
    cleaned = cleaned.strip("()").strip()
    sign = ""
    if cleaned.startswith("-"):
        sign = "-"
        cleaned = cleaned[1:]
    digits_only = cleaned.replace(".", "")
    if digits_only.isdigit() and len(digits_only) > 9:
        digits_only = digits_only[-9:]
        cleaned = digits_only
    cleaned = sign + cleaned
    try:
        value = Decimal(cleaned)
    except (InvalidOperation, ValueError):
        return None
    if negative:
        value = -value
    return float(value)


WAREKI_BOUNDARIES = [
    (5, 2018),   # Reiwa 1-5 -> 2019-2023
    (31, 1988),  # Heisei 1-31 -> 1989-2019
    (64, 1925),  # Showa 1-64 -> 1926-1989
]


def _parse_date(text: Optional[str], *, date_format: str) -> Optional[str]:
    if not text:
        return None
    cleaned = text.strip()
    cleaned = cleaned.replace("：", "-").replace(":", "-")
    cleaned = cleaned.replace("/", "-").replace(".", "-").replace(",", "-")
    cleaned = re.sub(r"\s+", "-", cleaned)
    cleaned = re.sub(r"[^0-9-]", "", cleaned)
    cleaned = cleaned.strip("-")
    if not cleaned:
        return None

    year: Optional[int] = None
    month: Optional[int] = None
    day: Optional[int] = None

    def build(year_value: int, month_value: int, day_value: int) -> Optional[str]:
        try:
            return f"{year_value:04d}-{month_value:02d}-{day_value:02d}"
        except ValueError:
            return None

    # Try YYYY-MM-DD first
    match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", cleaned)
    if match:
        year, month, day = map(int, match.groups())
        return build(year, month, day)

    match = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", cleaned)
    if match:
        year, month, day = map(int, match.groups())
        return build(year, month, day)

    # Handle two-digit years
    match = re.fullmatch(r"(\d{1,2})-(\d{1,2})-(\d{1,2})", cleaned)
    if match:
        y, m, d = match.groups()
        return build(_resolve_two_digit_year(int(y), date_format), int(m), int(d))

    match = re.fullmatch(r"(\d{2})(\d{2})(\d{2})", cleaned)
    if match:
        y, m, d = match.groups()
        return build(_resolve_two_digit_year(int(y), date_format), int(m), int(d))

    return None


def _resolve_two_digit_year(two_digit: int, date_format: str) -> int:
    if date_format == "western":
        return 2000 + two_digit if two_digit < 50 else 1900 + two_digit
    # default or wareki
    for boundary, offset in WAREKI_BOUNDARIES:
        if two_digit <= boundary:
            return offset + two_digit
    return 2000 + two_digit
