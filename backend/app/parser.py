"""Document parsers converting OCR output into normalized asset structures."""
from __future__ import annotations

import re
from datetime import datetime
from typing import Iterable, List, Optional, Tuple

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

DEPOSIT_KEYWORDS = ("振込", "入金", "預入", "配当")


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
    i = 0
    while i < len(normalized_lines):
        date_iso, remainder = extract_date_and_remainder(normalized_lines[i])
        if not date_iso:
            i += 1
            continue
        segments = [remainder.strip()] if remainder.strip() else []
        j = i + 1
        while j < len(normalized_lines):
            next_date, _ = extract_date_and_remainder(normalized_lines[j])
            if next_date:
                break
            segments.append(normalized_lines[j].strip())
            j += 1
        txn = build_transaction(date_iso, segments)
        if txn:
            transactions.append(txn)
        i = j

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
    text = text.replace(", ", ",")
    return text.strip()


def build_transaction(date_iso: str, segments: List[str]) -> Optional[TransactionLine]:
    filtered_segments = [seg for seg in segments if seg and not seg.isdigit()]
    if not filtered_segments:
        return None
    description = " ".join(filtered_segments).strip()
    numbers = [normalize_amount(match) for match in AMOUNT_PATTERN.findall(description)]
    numbers = [amt for amt in numbers if amt is not None]
    numbers = [amt for amt in numbers if abs(amt) >= 10]

    withdrawal = deposit = balance = None
    if numbers:
        balance = numbers[-1]
        if len(numbers) >= 2:
            primary = numbers[0]
            if any(keyword in description for keyword in DEPOSIT_KEYWORDS):
                deposit = primary
            else:
                withdrawal = primary
    if not numbers and not description:
        return None
    return TransactionLine(
        transaction_date=date_iso,
        description=description or date_iso,
        withdrawal_amount=withdrawal,
        deposit_amount=deposit,
        balance=balance,
    )


def extract_date_and_remainder(text: str) -> Tuple[Optional[str], str]:
    match = DATE_PATTERN.search(text)
    if not match:
        return None, text
    date_iso = convert_to_iso(match.groups())
    remainder = text[match.end():]
    return date_iso, remainder


def convert_to_iso(groups: Tuple[str, str, str]) -> Optional[str]:
    y, m, d = map(int, groups)
    if m < 1 or m > 12 or d < 1 or d > 31:
        return None
    year = convert_year(y)
    try:
        return datetime(year, m, d).date().isoformat()
    except ValueError:
        return None


def convert_year(value: int) -> int:
    if value >= 63:
        return 1925 + value  # 昭和
    if value >= 30:
        return 1988 + value  # 平成
    if value <= 5:
        return 2018 + value  # 令和
    return 2000 + value


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
