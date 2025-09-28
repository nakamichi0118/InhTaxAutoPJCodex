"""Document parsers converting OCR output into normalized asset structures."""
from __future__ import annotations

import math
import re
from datetime import datetime
from typing import Iterable, List, Optional

from .models import AssetRecord, DocumentType, TransactionLine

DATE_PATTERN = re.compile(r"(\d{1,2})[-/](\d{1,2})[-/](\d{1,2})")
AMOUNT_PATTERN = re.compile(r"([+-]?[0-9][0-9,]*)")
ACCOUNT_PATTERN = re.compile(r"([0-9]{6,10})")
BRANCH_PATTERN = re.compile(r"([\w\u3040-\u30ff\u4e00-\u9faf]+)支店")
OWNER_PATTERN = re.compile(r"\b(\S+\s*\S*)サマ")
BALANCE_KEYWORDS = ["残高", "繰越残高"]

HYPHENS = str.maketrans({
    "−": "-",
    "―": "-",
    "ー": "-",
    "–": "-",
    "—": "-",
    "／": "/",
    "・": "",
})


def detect_document_type(lines: Iterable[str]) -> DocumentType:
    joined = "\n".join(lines)
    if any(keyword in joined for keyword in ("普通預金", "通帳", "預金", "入出金")):
        return "bank_deposit"
    if any(keyword in joined for keyword in ("固定資産税", "地番", "家屋")):
        return "land"
    return "unknown"


def parse_bankbook(lines: List[str], source_name: str) -> List[AssetRecord]:
    normalized_lines = [normalize_line(line) for line in lines if line.strip()]
    branch = find_first_match(BRANCH_PATTERN, normalized_lines)
    account = find_first_match(ACCOUNT_PATTERN, normalized_lines)
    owner = find_first_match(OWNER_PATTERN, normalized_lines)
    balance_line = next((line for line in normalized_lines if any(key in line for key in BALANCE_KEYWORDS)), None)
    balance_amount = extract_amount(balance_line) if balance_line else None

    transactions: List[TransactionLine] = []
    for line in normalized_lines:
        txn = parse_transaction_line(line)
        if txn:
            transactions.append(txn)

    asset = AssetRecord(
        category="bank_deposit",
        type="ordinary_deposit",
        source_document=source_name,
        owner_name=[owner] if owner else [],
        asset_name="普通預金",
        identifier_primary=account,
        identifier_secondary=branch,
        valuation_basis="通帳残高" if balance_amount is not None else None,
        valuation_amount=balance_amount,
        notes="\n".join(normalized_lines[:30]),
        transactions=transactions,
    )
    return [asset]


def normalize_line(line: str) -> str:
    text = line.translate(HYPHENS)
    text = text.replace("*", "")
    text = re.sub(r"\s+", " ", text)
    text = text.replace(" -", "-").replace("- ", "-")
    text = text.replace("--", "-")
    text = text.replace(" :", ":")
    return text.strip()


def parse_transaction_line(line: str) -> Optional[TransactionLine]:
    date_iso = extract_date(line)
    if not date_iso:
        return None
    amounts = [normalize_amount(match) for match in AMOUNT_PATTERN.findall(line)]
    amounts = [amt for amt in amounts if amt is not None]
    withdrawal = deposit = balance = None
    if amounts:
        balance = amounts[-1]
        if len(amounts) >= 2:
            primary = amounts[0]
            if any(keyword in line for keyword in ("振込", "入金", "預入", "配当")):
                deposit = primary
            else:
                withdrawal = primary
    description = line.strip()
    return TransactionLine(
        transaction_date=date_iso,
        description=description,
        withdrawal_amount=withdrawal,
        deposit_amount=deposit,
        balance=balance,
    )


def extract_date(text: str) -> Optional[str]:
    match = DATE_PATTERN.search(text)
    if not match:
        return None
    y, m, d = map(int, match.groups())
    if m < 1 or m > 12 or d < 1 or d > 31:
        return None
    year = convert_year(y)
    try:
        return datetime(year, m, d).date().isoformat()
    except ValueError:
        return None


def convert_year(value: int) -> int:
    if value >= 63:
        return 1925 + value  # 昭和を想定
    if value >= 30:
        return 1988 + value  # 平成
    if value <= 5:
        return 2018 + value  # 令和
    return 2000 + value  # fallback


def normalize_amount(raw: str) -> Optional[float]:
    if not raw:
        return None
    cleaned = raw.replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def extract_amount(text: Optional[str]) -> Optional[float]:
    if not text:
        return None
    match = AMOUNT_PATTERN.search(text)
    if match:
        return normalize_amount(match.group(1))
    return None


def find_first_match(pattern: re.Pattern[str], lines: Iterable[str]) -> Optional[str]:
    for line in lines:
        match = pattern.search(line)
        if match:
            return match.group(1)
    return None


def build_assets(document_type: DocumentType, lines: List[str], source_name: str) -> List[AssetRecord]:
    if document_type == "bank_deposit":
        return parse_bankbook(lines, source_name)
    asset = AssetRecord(
        category=document_type,
        source_document=source_name,
        notes="\n".join(lines[:50]),
    )
    return [asset]
