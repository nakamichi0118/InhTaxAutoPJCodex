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
ACCOUNT_HINTS = ('口座番号', '店番号', '座番号')
OWNER_SUFFIXES = ('様', 'サマ', 'さま')
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
    branch = extract_branch_name(normalized_lines)
    account = extract_account_number(normalized_lines)
    owner = extract_owner_name(normalized_lines)
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

    final_balance = balance_amount
    if final_balance is None:
        final_balance = next((txn.balance for txn in reversed(transactions) if txn.balance is not None), None)

    asset = AssetRecord(
        category="bank_deposit",
        type="ordinary_deposit",
        source_document=source_name,
        owner_name=[owner] if owner else [],
        asset_name="普通預金",
        identifier_primary=account,
        identifier_secondary=branch,
        valuation_basis="最終残高" if final_balance is not None else None,
        valuation_amount=final_balance,
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
    if value >= 2000:
        return value
    if 32 <= value <= 64:
        return 1925 + value  # 昭和
    if 6 <= value <= 31:
        return 1988 + value  # 平成
    if 1 <= value <= 5:
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



def extract_branch_name(lines: List[str]) -> Optional[str]:
    branch = find_first_match(BRANCH_PATTERN, lines)
    if branch:
        return branch
    for idx, line in enumerate(lines):
        if "支店" not in line:
            continue
        cleaned = line.replace("支店", "").strip()
        if cleaned and cleaned != "お取引店":
            return f"{cleaned}支店"
        for back in range(idx - 1, -1, -1):
            candidate = lines[back].strip()
            if not candidate or candidate in {"お取引店", "店番号", "電話"} or candidate.isdigit():
                continue
            if candidate.endswith("支店") or "支店" in candidate:
                return candidate
            return f"{candidate}支店"
    return None


def extract_account_number(lines: List[str]) -> Optional[str]:
    candidates: List[str] = []
    for hint in ACCOUNT_HINTS:
        for idx, line in enumerate(lines):
            if hint not in line:
                continue
            for candidate in lines[idx: idx + 5]:
                match = ACCOUNT_PATTERN.search(candidate)
                if match:
                    value = match.group(1)
                    candidates.append(value)
            if candidates:
                break
        if candidates:
            break
    if candidates:
        for value in candidates:
            if len(value) in (7, 8):
                return value
        return max(candidates, key=len)
    return find_first_match(ACCOUNT_PATTERN, lines)


def extract_owner_name(lines: List[str]) -> Optional[str]:
    owner = find_first_match(OWNER_PATTERN, lines)
    owner = owner.strip() if owner else None

    candidates: List[str] = []
    for line in lines:
        for suffix in OWNER_SUFFIXES:
            if suffix not in line:
                continue
            candidate = line.rsplit(suffix, 1)[0].strip()
            if candidate:
                candidates.append(candidate)
                break
    if candidates:
        best = max(candidates, key=len)
        if not owner or len(best) >= len(owner):
            return best
    return owner

def build_assets(document_type: DocumentType, lines: List[str], source_name: str) -> List[AssetRecord]:
    if document_type == "bank_deposit":
        return parse_bankbook(lines, source_name)
    asset = AssetRecord(
        category=document_type,
        source_document=source_name,
        notes="\n".join(lines[:50]),
    )
    return [asset]
