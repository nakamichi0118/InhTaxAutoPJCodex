"""Document parsers converting OCR output into normalized asset structures."""
from __future__ import annotations

import re
from typing import Iterable, List, Optional

from .models import AssetRecord, DocumentType, TransactionLine

DATE_PATTERNS = [
    re.compile(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})"),
    re.compile(r"(\d{2})/(\d{2})"),
    re.compile(r"(\d{1,2})月(\d{1,2})日"),
]
AMOUNT_PATTERN = re.compile(r"([+-]?[0-9][0-9,]*)(?:円)?")
ACCOUNT_PATTERN = re.compile(r"([0-9]{6,8})")
BRANCH_PATTERN = re.compile(r"([\w\u3040-\u30ff\u4e00-\u9faf]+)支店")
OWNER_PATTERN = re.compile(r"\b(\S+\s*\S*)名義")
BALANCE_KEYWORDS = ["残高", "残高合計"]


def detect_document_type(lines: Iterable[str]) -> DocumentType:
    joined = "\n".join(lines)
    if any(keyword in joined for keyword in ("普通預金", "通帳", "預金", "入出金")):
        return "bank_deposit"
    if any(keyword in joined for keyword in ("固定資産税", "地番", "家屋")):
        return "land"
    return "unknown"


def parse_bankbook(lines: List[str], source_name: str) -> List[AssetRecord]:
    branch = find_first_match(BRANCH_PATTERN, lines)
    account = find_first_match(ACCOUNT_PATTERN, lines)
    owner = find_first_match(OWNER_PATTERN, lines)
    balance_line = next((line for line in lines if any(key in line for key in BALANCE_KEYWORDS)), None)
    balance_amount = extract_amount(balance_line) if balance_line else None

    transactions = []
    for line in lines:
        transaction = parse_transaction_line(line)
        if transaction:
            transactions.append(transaction)

    asset = AssetRecord(
        category="bank_deposit",
        type="ordinary_deposit",
        source_document=source_name,
        owner_name=[owner] if owner else [],
        asset_name=f"{branch or ''} 普通預金".strip(),
        identifier_primary=account,
        valuation_basis="通帳残高",
        valuation_amount=balance_amount,
        notes="\n".join(lines[:20]),
        transactions=transactions,
    )
    return [asset]


def parse_transaction_line(line: str) -> Optional[TransactionLine]:
    date_iso = extract_date(line)
    if not date_iso:
        return None
    amounts = AMOUNT_PATTERN.findall(line)
    withdrawal = deposit = balance = None
    if len(amounts) >= 1:
        withdrawal = normalize_amount(amounts[0])
    if len(amounts) >= 2:
        deposit = normalize_amount(amounts[1])
    if len(amounts) >= 3:
        balance = normalize_amount(amounts[2])
    description = re.sub(r"\s+", " ", line)
    return TransactionLine(
        transaction_date=date_iso,
        description=description,
        withdrawal_amount=withdrawal,
        deposit_amount=deposit,
        balance=balance,
    )


def extract_date(text: str) -> Optional[str]:
    for pattern in DATE_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        groups = match.groups()
        if len(groups) == 3 and len(groups[0]) == 4:
            year, month, day = map(int, groups)
        elif len(groups) == 2:
            month, day = map(int, groups)
            year = guess_year()
        else:
            continue
        return f"{year:04d}-{month:02d}-{day:02d}"
    kanji = re.search(r"(令和|平成|昭和|大正)(\d{1,2})年(\d{1,2})月(\d{1,2})日", text)
    if kanji:
        base = {"令和": 2018, "平成": 1988, "昭和": 1925, "大正": 1911}[kanji.group(1)]
        year = base + int(kanji.group(2))
        month = int(kanji.group(3))
        day = int(kanji.group(4))
        return f"{year:04d}-{month:02d}-{day:02d}"
    return None


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


def guess_year() -> int:
    from datetime import datetime

    return datetime.utcnow().year


def build_assets(document_type: DocumentType, lines: List[str], source_name: str) -> List[AssetRecord]:
    if document_type == "bank_deposit":
        return parse_bankbook(lines, source_name)
    # Fallback: return single unknown asset with notes
    asset = AssetRecord(
        category=document_type,
        type=None,
        source_document=source_name,
        notes="\n".join(lines[:50]),
    )
    return [asset]
