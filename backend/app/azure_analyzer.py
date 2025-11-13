from __future__ import annotations

import logging
import re
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import HttpResponseError

from .models import AssetRecord, TransactionLine

logger = logging.getLogger(__name__)


ProgressCallback = Callable[[str, str], None]


@dataclass
class AzureAnalysisResult:
    raw_lines: List[str]
    assets: List[AssetRecord]


def build_transactions_from_lines(lines: Iterable[str], *, date_format: str) -> List[TransactionLine]:
    """Convert free-form OCR lines into post-processed transactions."""
    raw = _extract_transactions_from_lines(lines, date_format=date_format)
    if not raw:
        return []
    return _post_process_transactions(raw)


def merge_transactions(
    primary: List[TransactionLine],
    supplementary: List[TransactionLine],
) -> List[TransactionLine]:
    """Merge and sort transactions, avoiding duplicates."""
    if not supplementary:
        return list(primary)
    return _merge_transactions(primary, supplementary)


def post_process_transactions(transactions: List[TransactionLine]) -> List[TransactionLine]:
    """Run post-processing pipeline on finalized transaction list."""
    if not transactions:
        return []
    return _post_process_transactions(transactions)


HEADER_ALIASES: Dict[str, str] = {
    "transaction_date": "取引日",
    "description": "摘要",
    "withdrawal": "支払金額",
    "deposit": "入金金額",
    "balance": "残高",
}

HEADER_KEYWORDS: Dict[str, Iterable[str]] = {
    "transaction_date": ("取引日", "年月日", "日付"),
    "description": ("摘要", "内容", "件名", "備考"),
    "withdrawal": ("支払", "出金", "出", "支払金額"),
    "deposit": ("入金", "預入", "入金金額", "預り", "お預り"),
    "balance": ("残高", "差引残高", "残高金額", "残", "高"),
}

DATE_TOKEN_RE = re.compile(r"(?:19|20)\d{2}[./-]\d{1,2}[./-]\d{1,2}")

BALANCE_ONLY_KEYWORDS = (
    "繰越",
    "繰り越し",
    "前日残高",
)

DESCRIPTION_CLEANUPS = (
    (":selected:", ""),
)

BALANCE_TOLERANCE = 1.0
DELTA_OVERRIDE_REL_TOLERANCE = 0.02  # 2%
BALANCE_REBALANCE_TOLERANCE = 999.0
MAX_REBALANCE_PASSES = 10
SEGMENT_WINDOW = 10
MAX_SEGMENT_ATTEMPTS = 10
AMOUNT_DISCREPANCY_THRESHOLD = 1_000.0
AMOUNT_RATIO_THRESHOLD = 8.0


DEPOSIT_KEYWORDS = (
    "入金",
    "振込",
    "振替",
    "利息",
    "配当",
    "給付",
    "カンプ",
    "キャンプ",
    "返金",
    "返戻",
    "還付",
    "補助",
    "保険金",
    "保険料",
    "入庫",
    "ATM入金",
    "預り",
    "お預り",
)


WITHDRAWAL_KEYWORDS = (
    "支払",
    "支払い",
    "引落",
    "引き落と",
    "引去",
    "ガス",
    "水道",
    "電気",
    "オリコ",
    "カード",
    "手数料",
    "公共料金",
    "納付",
    "ATM支払",
    "ATM出金",
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

    field_columns: Dict[str, List[int]] = {}
    for column_index, field in header_map.items():
        field_columns.setdefault(field, []).append(column_index)

    transactions: List[TransactionLine] = []
    for row_index in sorted(rows.keys()):
        if row_index == header_row:
            continue
        row_cells = rows[row_index]
        values: Dict[str, List[Optional[str]]] = {}
        for field, columns in field_columns.items():
            if not columns:
                continue
            sorted_columns = sorted(set(columns))
            clusters: List[List[int]] = []
            current_cluster: List[int] = [sorted_columns[0]]
            for column_index in sorted_columns[1:]:
                if column_index - current_cluster[-1] <= 1:
                    current_cluster.append(column_index)
                else:
                    clusters.append(current_cluster)
                    current_cluster = [column_index]
            clusters.append(current_cluster)

            selected_values: List[Optional[str]] = []
            for cluster in clusters:
                block_values = [
                    row_cells.get(column_index)
                    for column_index in range(cluster[0], cluster[-1] + 1)
                ]
                if any(block_values):
                    selected_values = block_values
                    break
                if not selected_values:
                    selected_values = block_values
            if selected_values:
                values.setdefault(field, []).extend(selected_values)
        txn = _build_transaction(values, date_format=date_format)
        if txn:
            transactions.append(txn)
    return transactions


def _extract_transactions_from_lines(lines: Iterable[str], *, date_format: str) -> List[TransactionLine]:
    transactions: List[TransactionLine] = []
    for line in lines:
        normalized = _clean_text(line)
        if not normalized:
            continue
        date_match = DATE_TOKEN_RE.search(normalized)
        if not date_match:
            continue
        date_raw = normalized[date_match.start():date_match.end()]
        transaction_date = _parse_date(date_raw, date_format=date_format)
        if not transaction_date:
            continue
        remainder = normalized[date_match.end():].strip()
        if not remainder:
            continue
        tokens = remainder.split()
        value_tokens: List[str] = []
        while tokens and _is_numeric_token(tokens[-1]):
            value_tokens.append(tokens.pop())
        description = " ".join(tokens).strip()
        if not description:
            description = remainder

        numeric_values = [
            _parse_numeric_token(token)
            for token in reversed(value_tokens)
            if _parse_numeric_token(token) is not None
        ]

        balance = numeric_values[-1] if numeric_values else None
        withdrawal = None
        deposit = None
        classification = _classify_description(description)
        if len(numeric_values) >= 2:
            amount = numeric_values[-2]
            if classification == "deposit":
                deposit = amount
            elif classification == "withdrawal":
                withdrawal = amount
        if len(numeric_values) >= 3 and classification == "unknown":
            withdrawal = numeric_values[-3]
            deposit = numeric_values[-2]

        transactions.append(
            TransactionLine(
                transaction_date=transaction_date,
                description=description or None,
                withdrawal_amount=withdrawal,
                deposit_amount=deposit,
                balance=balance,
            )
        )
    return transactions


def _merge_transactions(
    primary: List[TransactionLine],
    supplementary: List[TransactionLine],
) -> List[TransactionLine]:
    merged = list(primary)
    for txn in supplementary:
        match_index = _find_matching_transaction_index(merged, txn)
        if match_index is not None:
            merged[match_index] = _combine_transactions(merged[match_index], txn)
            continue
        merged.append(txn)
    merged.sort(key=_transaction_sort_key)
    return merged


def _transaction_signature(txn: TransactionLine) -> Tuple[str, str, Optional[float], Optional[float], Optional[float]]:
    return (
        txn.transaction_date or "",
        (txn.description or "").replace(" ", ""),
        round(txn.withdrawal_amount, 2) if txn.withdrawal_amount is not None else None,
        round(txn.deposit_amount, 2) if txn.deposit_amount is not None else None,
        round(txn.balance, 2) if txn.balance is not None else None,
    )


def _transaction_sort_key(txn: TransactionLine) -> Tuple[str, float]:
    date_key = txn.transaction_date or ""
    balance_key = txn.balance if txn.balance is not None else 0.0
    return (date_key, balance_key)


def _find_matching_transaction_index(
    existing: List[TransactionLine],
    candidate: TransactionLine,
) -> Optional[int]:
    for index, txn in enumerate(existing):
        if _transactions_equivalent(txn, candidate):
            return index
    return None


def _transactions_equivalent(lhs: TransactionLine, rhs: TransactionLine) -> bool:
    return (
        (lhs.transaction_date or "") == (rhs.transaction_date or "")
        and _float_equal(lhs.withdrawal_amount, rhs.withdrawal_amount)
        and _float_equal(lhs.deposit_amount, rhs.deposit_amount)
        and _float_equal(lhs.balance, rhs.balance)
    )


def _combine_transactions(primary: TransactionLine, supplementary: TransactionLine) -> TransactionLine:
    updates: Dict[str, Any] = {}
    if (not primary.description) and supplementary.description:
        updates["description"] = supplementary.description
    if primary.withdrawal_amount is None and supplementary.withdrawal_amount is not None:
        updates["withdrawal_amount"] = supplementary.withdrawal_amount
    if primary.deposit_amount is None and supplementary.deposit_amount is not None:
        updates["deposit_amount"] = supplementary.deposit_amount
    if primary.balance is None and supplementary.balance is not None:
        updates["balance"] = supplementary.balance
    if not updates:
        return primary
    return primary.model_copy(update=updates)


def _float_equal(lhs: Optional[float], rhs: Optional[float]) -> bool:
    if lhs is None and rhs is None:
        return True
    if lhs is None or rhs is None:
        return False
    return abs(lhs - rhs) <= 0.01


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
        normalized_withdrawal = _normalize_magnitude(candidate.withdrawal_amount)
        normalized_deposit = _normalize_magnitude(candidate.deposit_amount)

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
        projected_balance: Optional[float] = None
        if last_balance is not None and (
            candidate.withdrawal_amount is not None or candidate.deposit_amount is not None
        ):
            projected_balance = last_balance
            if candidate.withdrawal_amount is not None:
                projected_balance -= candidate.withdrawal_amount
            if candidate.deposit_amount is not None:
                projected_balance += candidate.deposit_amount

        if balance_value is None and projected_balance is not None:
            candidate = candidate.model_copy(update={"balance": projected_balance})
            balance_value = projected_balance
        elif balance_value is not None and projected_balance is not None:
            if abs(projected_balance - balance_value) > BALANCE_TOLERANCE:
                candidate = candidate.model_copy(update={"balance": projected_balance})
                balance_value = projected_balance

        classification = _classify_description(desc_cleaned)

        delta_updates: Dict[str, Any] = {}
        delta = None
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

            if classification == "deposit" and delta is not None and delta != 0:
                expected_deposit = float(abs(delta))
                expected_withdrawal = None
            elif classification == "withdrawal" and delta is not None and delta != 0:
                expected_withdrawal = float(abs(delta))
                expected_deposit = None

            if expected_withdrawal is not None and _should_override_amount(
                candidate.withdrawal_amount, expected_withdrawal
            ):
                delta_updates["withdrawal_amount"] = expected_withdrawal
                if candidate.deposit_amount not in (None, 0.0):
                    delta_updates["deposit_amount"] = None
            elif expected_deposit is not None and _should_override_amount(
                candidate.deposit_amount, expected_deposit
            ):
                delta_updates["deposit_amount"] = expected_deposit
                if candidate.withdrawal_amount not in (None, 0.0):
                    delta_updates["withdrawal_amount"] = None

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

        if not candidate.transaction_date:
            continue

        enriched.append(candidate)
        if candidate.balance is not None:
            last_balance = candidate.balance

    return enriched


def _reconcile_transactions(
    azure_transactions: List[TransactionLine],
    gemini_transactions: List[TransactionLine],
    progress_callback: Optional[ProgressCallback] = None,
    *,
    initial_balance: Optional[float] = None,
) -> List[TransactionLine]:
    if not azure_transactions:
        return []

    def notify(stage: str, detail: str) -> None:
        if progress_callback:
            progress_callback(stage, detail)

    if not gemini_transactions:
        notify("balance_probe", "Gemini結果が得られなかったためAzure結果を採用します。")
        return azure_transactions

    gemini_index = _build_gemini_index(gemini_transactions)
    corrected: List[TransactionLine] = []
    prev_balance = (
        initial_balance
        if initial_balance is not None
        else _opening_balance_before_first(azure_transactions)
    )

    for azure_txn in azure_transactions:
        candidate = _select_gemini_candidate(azure_txn, prev_balance, gemini_index)
        gemini_txn = candidate["txn"] if candidate else None
        prefer_gemini = _has_significant_amount_gap(azure_txn, gemini_txn)
        chosen, prev_balance = _choose_best_transaction(
            prev_balance,
            azure_txn,
            gemini_txn,
            prefer_gemini=prefer_gemini,
        )
        corrected.append(chosen)

    expected_balances, residuals = _compute_balance_residuals(corrected)
    worst_idx, worst_residual = _find_worst_residual(residuals)
    if worst_idx is None:
        notify("balance_refine", "取引結果の検算が完了しました。")
        return corrected
    if abs(worst_residual) <= BALANCE_REBALANCE_TOLERANCE:
        notify(
            "balance_refine",
            f"残高差 {abs(worst_residual):,.0f} 円以内に収まりました。",
        )
        return corrected

    notify(
        "balance_probe",
        f"{worst_idx + 1} 行目付近で {abs(worst_residual):,.0f} 円の残高差を検出しました",
    )

    passes = 0
    while passes < MAX_REBALANCE_PASSES and abs(worst_residual) > BALANCE_REBALANCE_TOLERANCE:
        segments = _build_residual_segments(
            residuals,
            tolerance=BALANCE_REBALANCE_TOLERANCE,
            window=SEGMENT_WINDOW,
        )
        if not segments:
            break

        changed = False
        attempts = 0
        for start_idx, end_idx in segments:
            notify(
                "balance_refine",
                f"{start_idx + 1}〜{end_idx + 1} 行の差分をGeminiと照合しています…",
            )
            if _apply_segment_corrections(
                corrected,
                gemini_index,
                start_idx,
                end_idx,
                expected_balances,
            ):
                changed = True
                break
            attempts += 1
            if attempts >= MAX_SEGMENT_ATTEMPTS:
                break

        if not changed:
            break

        passes += 1
        expected_balances, residuals = _compute_balance_residuals(corrected)
        worst_idx, worst_residual = _find_worst_residual(residuals)

    if worst_idx is None or abs(worst_residual) <= BALANCE_REBALANCE_TOLERANCE:
        notify("balance_refine", "AI補正が完了しました。")
    else:
        notify(
            "balance_refine",
            f"残高差 {abs(worst_residual):,.0f} 円 (行 {worst_idx + 1}) が残っています。CSVでご確認ください。",
        )
    return corrected


def _build_gemini_index(transactions: List[TransactionLine]) -> Dict[str, List[TransactionLine]]:
    index: Dict[str, List[TransactionLine]] = defaultdict(list)
    for txn in transactions:
        for key in _build_lookup_keys(txn):
            index[key].append(txn)
        index["__all__"].append(txn)
    return index


def _build_lookup_keys(txn: TransactionLine) -> List[str]:
    keys: List[str] = []
    date_value = txn.transaction_date
    if isinstance(date_value, str):
        date_key = date_value
    elif date_value:
        date_key = date_value.isoformat()
    else:
        date_key = ""
    desc_key = _normalize_desc_token(txn.description)
    amount_key: Optional[str] = None
    if txn.deposit_amount:
        amount_key = f"D:{round(txn.deposit_amount)}"
    elif txn.withdrawal_amount:
        amount_key = f"W:{round(txn.withdrawal_amount)}"

    if date_key and desc_key:
        keys.append(f"pair:{date_key}:{desc_key}")
    if date_key and amount_key:
        keys.append(f"date_amount:{date_key}:{amount_key}")
    if desc_key and amount_key:
        keys.append(f"desc_amount:{desc_key}:{amount_key}")
    if date_key:
        keys.append(f"date:{date_key}")
    if desc_key:
        keys.append(f"desc:{desc_key}")
    if amount_key:
        keys.append(f"amount:{amount_key}")
    return keys or ["__all__"]


def _normalize_desc_token(text: Optional[str]) -> str:
    if not text:
        return ""
    cleaned = _clean_description(text)
    normalized = re.sub(r"\s+", "", cleaned).lower()
    return normalized


def _find_worst_residual(residuals: List[float]) -> Tuple[Optional[int], float]:
    worst_idx: Optional[int] = None
    worst_value = 0.0
    for idx, value in enumerate(residuals):
        if abs(value) > abs(worst_value):
            worst_value = value
            worst_idx = idx
    return worst_idx, worst_value


def _build_residual_segments(
    residuals: List[float], *, tolerance: float, window: int
) -> List[Tuple[int, int]]:
    segments: List[Tuple[int, int]] = []
    idx = 0
    total = len(residuals)
    while idx < total:
        if abs(residuals[idx]) > tolerance:
            start = max(0, idx - window // 2)
            end = min(total - 1, start + window - 1)
            segments.append((start, end))
            idx = end + 1
        else:
            idx += 1
    return segments


def _apply_segment_corrections(
    transactions: List[TransactionLine],
    gemini_index: Dict[str, List[TransactionLine]],
    start_idx: int,
    end_idx: int,
    expected_balances: List[float],
) -> bool:
    changed = False
    prev_balance = _balance_before_index(transactions, expected_balances, start_idx)

    for idx in range(start_idx, min(end_idx + 1, len(transactions))):
        azure_txn = transactions[idx]
        azure_candidate = _prepare_candidate(prev_balance, azure_txn, source="azure")
        gemini_candidate = _select_gemini_candidate(
            azure_txn,
            prev_balance,
            gemini_index,
        )

        if gemini_candidate and abs(gemini_candidate["residual"]) + BALANCE_TOLERANCE < abs(azure_candidate["residual"]):
            notes = list(gemini_candidate.get("notes", []))
            notes.extend(
                [
                    "Geminiの結果で補正しました",
                    f"残高差 {azure_candidate['residual']:,.0f}→{gemini_candidate['residual']:,.0f} 円",
                ]
            )
            transactions[idx] = gemini_candidate["txn"].model_copy(
                update={
                    "balance": gemini_candidate["next_balance"],
                    "correction_note": _merge_notes(gemini_candidate["txn"].correction_note, notes),
                }
            )
            prev_balance = gemini_candidate["next_balance"]
            changed = True
        else:
            prev_balance = azure_candidate["next_balance"]

    return changed


def _balance_before_index(
    transactions: List[TransactionLine], expected_balances: List[float], idx: int
) -> float:
    if not transactions:
        return 0.0
    if idx <= 0:
        first = transactions[0]
        baseline = expected_balances[0] if expected_balances else 0.0
        withdrawal = first.withdrawal_amount or 0.0
        deposit = first.deposit_amount or 0.0
        return baseline + withdrawal - deposit
    if idx - 1 < len(expected_balances):
        return expected_balances[idx - 1]
    return expected_balances[-1] if expected_balances else 0.0


def _opening_balance_before_first(transactions: List[TransactionLine]) -> float:
    if not transactions:
        return 0.0
    first = transactions[0]
    withdrawal = first.withdrawal_amount or 0.0
    deposit = first.deposit_amount or 0.0
    if first.balance is None:
        return deposit - withdrawal
    return float(first.balance) + withdrawal - deposit


def _select_gemini_candidate(
    azure_txn: TransactionLine,
    prev_balance: float,
    gemini_index: Dict[str, List[TransactionLine]],
) -> Optional[Dict[str, Any]]:
    lookup_keys = _build_lookup_keys(azure_txn) + ["__all__"]
    seen: set[int] = set()
    candidates: List[Dict[str, Any]] = []

    for key in lookup_keys:
        for txn in gemini_index.get(key, []):
            txn_id = id(txn)
            if txn_id in seen:
                continue
            seen.add(txn_id)
            prepared = _prepare_candidate(prev_balance, txn, source="gemini")
            penalty = _candidate_penalty(azure_txn, txn)
            prepared["penalty"] = penalty
            candidates.append(prepared)
        if candidates:
            break

    if not candidates:
        return None

    candidates.sort(key=lambda item: (abs(item["residual"]), item.get("penalty", 0)))
    return candidates[0]


def _candidate_penalty(azure_txn: TransactionLine, gemini_txn: TransactionLine) -> int:
    penalty = 0
    if azure_txn.transaction_date and gemini_txn.transaction_date:
        if azure_txn.transaction_date != gemini_txn.transaction_date:
            penalty += 1
    if _normalize_desc_token(azure_txn.description) and _normalize_desc_token(gemini_txn.description):
        if _normalize_desc_token(azure_txn.description) != _normalize_desc_token(gemini_txn.description):
            penalty += 1
    return penalty


def _has_significant_amount_gap(
    azure_txn: Optional[TransactionLine],
    gemini_txn: Optional[TransactionLine],
) -> bool:
    if not azure_txn or not gemini_txn:
        return False
    gap = _amount_gap(azure_txn, gemini_txn)
    if gap >= AMOUNT_DISCREPANCY_THRESHOLD:
        return True
    ratio = _amount_ratio(azure_txn, gemini_txn)
    if ratio >= AMOUNT_RATIO_THRESHOLD:
        return True
    missing_withdrawal = azure_txn.withdrawal_amount and gemini_txn.withdrawal_amount is None
    missing_deposit = azure_txn.deposit_amount and gemini_txn.deposit_amount is None
    if missing_withdrawal and abs(azure_txn.withdrawal_amount or 0.0) >= AMOUNT_DISCREPANCY_THRESHOLD:
        return True
    if missing_deposit and abs(azure_txn.deposit_amount or 0.0) >= AMOUNT_DISCREPANCY_THRESHOLD:
        return True
    return False


def _amount_gap(a: TransactionLine, b: TransactionLine) -> float:
    diffs: List[float] = []
    if a.withdrawal_amount is not None and b.withdrawal_amount is not None:
        diffs.append(abs(a.withdrawal_amount - b.withdrawal_amount))
    if a.deposit_amount is not None and b.deposit_amount is not None:
        diffs.append(abs(a.deposit_amount - b.deposit_amount))
    return max(diffs) if diffs else 0.0


def _amount_ratio(a: TransactionLine, b: TransactionLine) -> float:
    ratios: List[float] = []
    if a.withdrawal_amount and b.withdrawal_amount:
        ratios.append(_safe_ratio(a.withdrawal_amount, b.withdrawal_amount))
    if a.deposit_amount and b.deposit_amount:
        ratios.append(_safe_ratio(a.deposit_amount, b.deposit_amount))
    return max(ratios) if ratios else 1.0


def _safe_ratio(a: float, b: float) -> float:
    if b == 0:
        return abs(a)
    return max(abs(a / b), abs(b / a))


SOURCE_PRIORITY = {"azure": 0, "gemini": 1}


def _choose_best_transaction(
    prev_balance: float,
    azure_txn: Optional[TransactionLine],
    gemini_txn: Optional[TransactionLine],
    *,
    prefer_gemini: bool = False,
) -> tuple[TransactionLine, float]:
    candidates: List[Dict[str, Any]] = []
    if azure_txn:
        candidates.append(_prepare_candidate(prev_balance, azure_txn, source="azure"))
    if gemini_txn:
        candidates.append(_prepare_candidate(prev_balance, gemini_txn, source="gemini"))
    if not candidates:
        placeholder = TransactionLine(
            transaction_date=None,
            description=None,
            withdrawal_amount=None,
            deposit_amount=None,
            balance=prev_balance,
        )
        return placeholder, prev_balance

    forced_source = "gemini" if prefer_gemini and gemini_txn else None
    if forced_source:
        preferred = [candidate for candidate in candidates if candidate["source"] == forced_source]
        if preferred:
            best = preferred[0]
        else:
            best = candidates[0]
    else:
        candidates.sort(key=lambda item: (abs(item["residual"]), SOURCE_PRIORITY.get(item["source"], 99)))
        best = candidates[0]
    notes = list(best["notes"])
    if best["source"] == "gemini":
        notes.append("Geminiの読み取り結果を採用しました")
    if abs(best["residual"]) > BALANCE_TOLERANCE:
        notes.append(f"残高差 {best['residual']:,.0f} 円が残っています")
    updated = best["txn"].model_copy(
        update={"correction_note": _merge_notes(best["txn"].correction_note, notes)}
    )
    return updated, best["next_balance"]


def _prepare_candidate(
    prev_balance: float,
    txn: TransactionLine,
    *,
    source: str,
) -> Dict[str, Any]:
    notes: List[str] = []
    withdrawal = txn.withdrawal_amount or 0.0
    deposit = txn.deposit_amount or 0.0
    expected_balance = prev_balance - withdrawal + deposit
    actual_balance = txn.balance
    updates: Dict[str, Any] = {}
    if actual_balance is None:
        actual_balance = expected_balance
        updates["balance"] = actual_balance
        notes.append("残高が欠損していたため再計算しました")
    residual = actual_balance - expected_balance
    normalized = txn.model_copy(update=updates) if updates else txn
    return {
        "source": source,
        "txn": normalized,
        "residual": residual,
        "next_balance": actual_balance,
        "notes": notes,
    }


def _compute_balance_residuals(transactions: List[TransactionLine]) -> Tuple[List[float], List[float]]:
    expected_balances: List[float] = []
    residuals: List[float] = []
    if not transactions:
        return expected_balances, residuals

    first = transactions[0]
    running_balance = first.balance
    if running_balance is None:
        running_balance = (first.deposit_amount or 0.0) - (first.withdrawal_amount or 0.0)
    expected_balances.append(running_balance)
    residuals.append(_balance_residual(first.balance, running_balance))

    for txn in transactions[1:]:
        withdrawal = txn.withdrawal_amount or 0.0
        deposit = txn.deposit_amount or 0.0
        running_balance = running_balance - withdrawal + deposit
        expected_balances.append(running_balance)
        residuals.append(_balance_residual(txn.balance, running_balance))

    return expected_balances, residuals


def _balance_residual(actual: Optional[float], expected: float) -> float:
    if actual is None:
        return 0.0
    return float(actual) - expected


def _merge_notes(existing: Optional[str], notes: List[str]) -> Optional[str]:
    parts: List[str] = []
    if existing:
        parts.append(existing)
    parts.extend(note for note in notes if note)
    merged = "; ".join(part for part in parts if part)
    return merged or None


def _should_override_amount(current: Optional[float], expected: Optional[float]) -> bool:
    if expected is None:
        return False
    if current is None or abs(current) < 1e-6:
        return True
    tolerance = max(1.0, abs(expected) * DELTA_OVERRIDE_REL_TOLERANCE)
    return abs(current - expected) <= tolerance



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


def _normalize_magnitude(value: Optional[float]) -> Optional[float]:
    """Normalize withdrawal/deposit values without balance reference."""
    return _normalize_amount(value, reference=None)


def _map_headers(columns: Dict[int, str]) -> Dict[int, str]:
    mapping: Dict[int, str] = {}
    for column_index, text in columns.items():
        normalized = _normalize_header(text)
        if not normalized:
            continue
        if "前回" in normalized or "端末" in normalized:
            continue
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
            if last_field in {"withdrawal", "deposit"} and text in {"", "金額", "金(円)", "円", "金"}:
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
    withdrawal_text = _merge_parts(values.get("withdrawal"))
    deposit_text = _merge_parts(values.get("deposit"))
    balance_text = _merge_parts(values.get("balance"))

    if not date_text and any(_looks_like_annotation(text) for text in (withdrawal_text, deposit_text)):
        return None

    withdrawal = _parse_amount(withdrawal_text)
    deposit = _parse_amount(deposit_text)
    balance = _parse_amount(balance_text)

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


def _looks_like_annotation(text: Optional[str]) -> bool:
    if not text:
        return False
    stripped = text.strip()
    return "(" in stripped or ")" in stripped or stripped.startswith("(") or stripped.endswith(")")


def _is_numeric_token(token: str) -> bool:
    if not token:
        return False
    token = token.strip()
    if not token:
        return False
    token = token.replace(",", "")
    if token.startswith("+") or token.startswith("-"):
        token = token[1:]
    return token.isdigit()


def _parse_numeric_token(token: str) -> Optional[float]:
    token = token.strip().replace(",", "")
    if not token:
        return None
    try:
        return float(token)
    except ValueError:
        return None


def _classify_description(text: str) -> str:
    if not text:
        return "unknown"
    lowered = text.lower()
    for keyword in DEPOSIT_KEYWORDS:
        if keyword.lower() in lowered:
            return "deposit"
    for keyword in WITHDRAWAL_KEYWORDS:
        if keyword.lower() in lowered:
            return "withdrawal"
    return "unknown"


def _parse_amount(text: Optional[str]) -> Optional[float]:
    if not text:
        return None
    cleaned = text.replace("円", "").replace("¥", "")
    cleaned = cleaned.replace(",", "").replace(" ", "")
    cleaned = cleaned.replace(":", "").replace("|", "")
    cleaned = re.sub(r"[^\d\-\.\(\)]", "", cleaned)
    if "." in cleaned and len(cleaned.replace(".", "")) > 4:
        cleaned = cleaned.replace(".", "")
    cleaned = cleaned.replace("△", "-")
    cleaned = cleaned.strip()
    if not cleaned or cleaned == "-":
        return None
    negative = cleaned.startswith("(") and cleaned.endswith(")")
    cleaned = cleaned.strip("()").strip()
    number_match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if number_match:
        cleaned = number_match.group(0)
    else:
        return None
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
    normalized = _normalize_amount(float(value), reference=None)
    return normalized if normalized is not None else float(value)


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
    cleaned = re.sub(r"-+", "-", cleaned)
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

    match = re.fullmatch(r"(\d{2})(\d{2})-(\d{2})", cleaned)
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
