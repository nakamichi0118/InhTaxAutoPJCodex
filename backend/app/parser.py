"""Document parsers converting OCR output into normalized asset structures."""
from __future__ import annotations

import re
from datetime import datetime
from typing import Iterable, List, Optional, Tuple

from .models import AssetRecord, DocumentType, TransactionLine

DATE_PATTERN = re.compile(r"(\d{1,4})[-/](\d{1,2})[-/](\d{1,2})")
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
    "⁄": "/",
    "・": "",
    "：": "-",
    ";": "-",
    ":": "-",
})

DEPOSIT_KEYWORDS = ("振込", "入金", "預入", "配当", "振込入金", "定期積金")

ROW_CODE_PATTERN = re.compile(r"^\d{3}$")
NON_DATE_TOKEN = re.compile(r"[^\d\s:/\-.;]")


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

    transactions = build_transactions_from_rows(normalized_lines)
    if not transactions:
        transactions = build_transactions_from_entries(normalized_lines)

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
    cleaned_segments = [seg.strip() for seg in segments if seg and seg.strip()]
    if not cleaned_segments:
        return None

    text_segments = [seg for seg in cleaned_segments if not seg.isdigit()]
    description = " ".join(text_segments).strip()
    if not description:
        description = " ".join(cleaned_segments).strip()

    number_source = " ".join(cleaned_segments)
    numbers = [normalize_amount(match) for match in AMOUNT_PATTERN.findall(number_source)]
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


def build_transactions_from_rows(lines: List[str]) -> List[TransactionLine]:
    rows = extract_transaction_rows(lines)
    transactions: List[TransactionLine] = []
    for row in rows:
        if not row:
            continue
        code = row[0].strip()
        if not ROW_CODE_PATTERN.fullmatch(code):
            continue
        date_iso, consumed = parse_row_date(row[1:])
        if not date_iso:
            continue
        remainder = [part.strip() for part in row[1 + consumed:] if part.strip()]
        if not remainder:
            continue
        txn = build_transaction(date_iso, remainder)
        if txn:
            transactions.append(txn)
    return transactions



def extract_transaction_rows(lines: List[str]) -> List[List[str]]:
    rows: List[List[str]] = []
    current: List[str] = []
    for line in lines:
        if ROW_CODE_PATTERN.fullmatch(line.strip()):
            if current:
                rows.append(current)
            current = [line]
        else:
            if current:
                current.append(line)
    if current:
        rows.append(current)
    return rows



def parse_row_date(tokens: List[str]) -> Tuple[Optional[str], int]:
    consumed = 0
    pieces: List[str] = []
    for token in tokens:
        cleaned = token.strip()
        consumed += 1
        if not cleaned:
            continue
        if NON_DATE_TOKEN.search(cleaned):
            if not pieces:
                return None, 0
            consumed -= 1
            break
        pieces.append(cleaned)
        digit_groups = re.findall(r"\d+", ''.join(pieces))
        digits = ''.join(digit_groups)
        if len(digit_groups) < 3 and len(digits) < 6:
            continue
        date_iso = parse_compact_date_digits(digits)
        if date_iso:
            return date_iso, consumed
    return None, 0


def parse_compact_date_digits(digits: str) -> Optional[str]:
    digits = digits.strip()
    if len(digits) < 3:
        return None
    for year_len in (4, 3, 2):
        if len(digits) <= year_len:
            continue
        year_part = digits[:year_len]
        rest = digits[year_len:]
        for month_len in (2, 1):
            if len(rest) < month_len + 1:
                continue
            month_part = rest[:month_len]
            day_part = rest[month_len:]
            if not day_part:
                continue
            day_part = day_part[:2]
            try:
                year_value = int(year_part)
                if len(year_part) > 2 and year_value < 1900:
                    year_value = int(year_part[-2:])
                month_value = int(month_part)
                day_value = int(day_part)
            except ValueError:
                continue
            if year_value == 0 or not (1 <= month_value <= 12 and 1 <= day_value <= 31):
                continue
            iso = convert_to_iso((str(year_value), str(month_value), str(day_value)))
            if iso:
                return iso
    return None


def build_transactions_from_entries(lines: List[str]) -> List[TransactionLine]:
    entries: List[tuple[Optional[str], str]] = []
    for line in lines:
        entries.extend(split_line_segments(line))

    transactions: List[TransactionLine] = []
    current_date: Optional[str] = None
    segments: List[str] = []
    for date_iso, fragment in entries:
        fragment = fragment.strip()
        if date_iso:
            if current_date:
                txn = build_transaction(current_date, segments)
                if txn:
                    transactions.append(txn)
            current_date = date_iso
            segments = [fragment] if fragment else []
            continue
        if current_date and fragment:
            segments.append(fragment)

    if current_date:
        txn = build_transaction(current_date, segments)
        if txn:
            transactions.append(txn)
    return transactions


def split_line_segments(line: str) -> List[tuple[Optional[str], str]]:
    matches = list(DATE_PATTERN.finditer(line))
    if not matches:
        return [(None, line)]
    entries: List[tuple[Optional[str], str]] = []
    prefix = line[: matches[0].start()].strip()
    if prefix:
        entries.append((None, prefix))
    for idx, match in enumerate(matches):
        date_iso = convert_to_iso(match.groups())
        next_start = matches[idx + 1].start() if idx + 1 < len(matches) else len(line)
        remainder = line[match.end():next_start]
        if date_iso:
            entries.append((date_iso, remainder))
        else:
            segment = line[match.start():next_start]
            entries.append((None, segment))
    return entries

def extract_branch_name(lines: List[str]) -> Optional[str]:
    branch = find_first_match(BRANCH_PATTERN, lines)
    if branch:
        return branch.replace(' ', '')
    for idx, line in enumerate(lines):
        normalized = line.replace(' ', '')
        if '支店' not in normalized:
            continue
        cleaned = normalized.replace('支店', '')
        if cleaned and cleaned != 'お取引店':
            return f"{cleaned}支店"
        for back in range(idx - 1, -1, -1):
            candidate = lines[back].replace(' ', '').strip()
            if not candidate or candidate in {'お取引店', '店番号', '電話'} or candidate.isdigit():
                continue
            if candidate.endswith('支店') or '支店' in candidate:
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
    if not owner:
        owner = find_owner_from_labels(lines)
    return owner



def find_owner_from_labels(lines: List[str]) -> Optional[str]:
    skip_tokens = {'住所', '電話', '顧客番号', '受付番号', '支店', '信用金庫'}

    def next_candidate(start: int) -> Optional[str]:
        for candidate in lines[start:start + 5]:
            stripped = candidate.strip()
            if not stripped:
                continue
            if stripped.isdigit():
                continue
            if any(token in stripped for token in skip_tokens):
                continue
            if '非会員' in stripped:
                continue
            return stripped
        return None

    for idx, line in enumerate(lines):
        if '非会員' in line:
            candidate = next_candidate(idx + 1)
            if candidate:
                return candidate
    for idx, line in enumerate(lines):
        if '氏名' in line or '名義人' in line:
            candidate = next_candidate(idx + 1)
            if candidate:
                return candidate
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
