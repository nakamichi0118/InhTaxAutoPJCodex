from __future__ import annotations

import threading
import uuid
from datetime import datetime, timezone
from typing import Dict, List, Optional, Set, Tuple

from fastapi import APIRouter, Depends, Header, HTTPException, Query

from . import job_registry
from .config import get_settings
from .ledger_store import LedgerScope, LedgerStore
from .models import (
    LedgerAccountCreateRequest,
    LedgerAccountPayload,
    LedgerAccountReorderRequest,
    LedgerAccountUpdateRequest,
    LedgerCaseCreateRequest,
    LedgerCasePayload,
    LedgerExportResponse,
    LedgerImportRequest,
    LedgerJobImportRequest,
    LedgerJobPreviewAccount,
    LedgerJobPreviewResponse,
    LedgerSessionRequest,
    LedgerSessionResponse,
    LedgerStateResponse,
    LedgerTransactionCreateRequest,
    LedgerTransactionPayload,
    LedgerTransactionUpdateRequest,
    LedgerTransactionsReorderRequest,
)

router = APIRouter(prefix="/api/ledger", tags=["ledger"])

_store: Optional[LedgerStore] = None
_store_lock = threading.Lock()


def get_ledger_store() -> LedgerStore:
    global _store
    if _store is None:
        with _store_lock:
            if _store is None:
                settings = get_settings()
                _store = LedgerStore(settings.ledger_db_path)
    return _store


def get_scope(
    ledger_token: str = Header(..., alias="X-Ledger-Token"),
    app_id_header: Optional[str] = Header(None, alias="X-Ledger-App"),
    app_id_query: Optional[str] = Query(None, alias="app_id"),
) -> LedgerScope:
    app_id = (app_id_header or app_id_query or "ledger-app").strip()
    token = (ledger_token or "").strip()
    if not token:
        raise HTTPException(status_code=400, detail="Missing ledger token")
    return LedgerScope(app_id=app_id, user_id=token)


def _resolve_cases(scope: LedgerScope, store: LedgerStore, preferred_id: Optional[str] = None) -> tuple[dict, List[dict]]:
    cases = store.list_cases(scope)
    if not cases:
        default_case = store.get_or_create_default_case(scope)
        cases = [default_case]
    if preferred_id:
        try:
            selected_case = store.get_case(scope, preferred_id)
        except KeyError:
            selected_case = cases[0]
    else:
        selected_case = cases[0]
    return selected_case, cases


def _serialize_cases(cases: List[dict]) -> List[LedgerCasePayload]:
    return [LedgerCasePayload(**case) for case in cases]


def _encode_tags(tags: Optional[List[str]]) -> str:
    if not tags:
        return ""
    return ",".join([tag.strip() for tag in tags if tag and tag.strip()])


def _decode_tags(raw: Optional[str]) -> List[str]:
    if not raw:
        return []
    return [part.strip() for part in raw.split(",") if part.strip()]


@router.post("/session", response_model=LedgerSessionResponse)
def create_session(payload: LedgerSessionRequest) -> LedgerSessionResponse:
    app_id = (payload.app_id or "ledger-app").strip()
    session_token = payload.session_token or uuid.uuid4().hex
    return LedgerSessionResponse(status="ok", app_id=app_id, user_id=session_token, session_token=session_token)


@router.get("/cases")
def list_cases(scope: LedgerScope = Depends(get_scope), store: LedgerStore = Depends(get_ledger_store)) -> dict:
    _, cases = _resolve_cases(scope, store)
    return {"status": "ok", "cases": _serialize_cases(cases)}


@router.post("/cases", response_model=LedgerCasePayload)
def create_case(
    payload: LedgerCaseCreateRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerCasePayload:
    case = store.create_case(scope, payload.name)
    return LedgerCasePayload(**case)


@router.get("/state", response_model=LedgerStateResponse)
def get_state(
    case_id: Optional[str] = Query(None, alias="case_id"),
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerStateResponse:
    selected_case, cases = _resolve_cases(scope, store, case_id)
    snapshot = store.snapshot(scope, selected_case["id"])
    return LedgerStateResponse(
        status="ok",
        case=LedgerCasePayload(**selected_case),
        cases=_serialize_cases(cases),
        accounts=[LedgerAccountPayload(**account) for account in snapshot["accounts"]],
        transactions=[LedgerTransactionPayload(**txn) for txn in snapshot["transactions"]],
    )


def _resolve_case_id(scope: LedgerScope, store: LedgerStore, case_id: Optional[str]) -> dict:
    if case_id:
        try:
            return store.get_case(scope, case_id)
        except KeyError:
            raise HTTPException(status_code=404, detail="Case not found") from None
    selected_case, _ = _resolve_cases(scope, store)
    return selected_case


@router.post("/accounts", response_model=LedgerAccountPayload)
def create_account(
    payload: LedgerAccountCreateRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerAccountPayload:
    case = _resolve_case_id(scope, store, getattr(payload, "case_id", None))
    account = store.create_account(
        scope,
        case_id=case["id"],
        name=payload.name,
        number=payload.number,
        holder_name=payload.holder_name,
        order=payload.order,
    )
    return LedgerAccountPayload(**account)


@router.patch("/accounts/{account_id}", response_model=LedgerAccountPayload)
def update_account(
    account_id: str,
    payload: LedgerAccountUpdateRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerAccountPayload:
    try:
        account = store.update_account(
            scope,
            account_id,
            name=payload.name,
            number=payload.number,
            holder_name=payload.holder_name,
            order=payload.order,
        )
    except KeyError:
        raise HTTPException(status_code=404, detail="Account not found") from None
    return LedgerAccountPayload(**account)


@router.delete("/accounts/{account_id}")
def delete_account(
    account_id: str,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> dict:
    try:
        store.delete_account(scope, account_id)
    except KeyError:
        raise HTTPException(status_code=404, detail="Account not found") from None
    return {"status": "ok"}


@router.post("/accounts/reorder")
def reorder_accounts(
    payload: LedgerAccountReorderRequest,
    case_id: str = Query(..., alias="case_id"),
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> dict:
    if not payload.items:
        return {"status": "ok", "updated": 0}
    store.reorder_accounts(scope, case_id, [(item.account_id, item.order) for item in payload.items])
    return {"status": "ok", "updated": len(payload.items)}


@router.post("/transactions", response_model=LedgerTransactionPayload)
def create_transaction(
    payload: LedgerTransactionCreateRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerTransactionPayload:
    withdrawal = payload.withdrawal or 0
    deposit = payload.deposit or 0
    if withdrawal < 0 or deposit < 0:
        raise HTTPException(status_code=400, detail="withdrawal/deposit must be positive")
    if withdrawal == 0 and deposit == 0:
        raise HTTPException(status_code=400, detail="Either withdrawal or deposit is required")
    try:
        transaction = store.create_transaction(
            scope,
            account_id=payload.account_id,
            date=payload.date,
            withdrawal=withdrawal,
            deposit=deposit,
            memo=payload.memo,
            txn_type=payload.type,
            user_order=payload.user_order,
            row_color=payload.row_color,
            tags=_encode_tags(payload.tags),
        )
    except KeyError:
        raise HTTPException(status_code=404, detail="Account not found") from None
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from None
    return LedgerTransactionPayload(**transaction)


@router.patch("/transactions/{transaction_id}", response_model=LedgerTransactionPayload)
def update_transaction(
    transaction_id: str,
    payload: LedgerTransactionUpdateRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerTransactionPayload:
    withdrawal = payload.withdrawal
    deposit = payload.deposit
    if withdrawal is not None and withdrawal < 0:
        raise HTTPException(status_code=400, detail="withdrawal must be positive")
    if deposit is not None and deposit < 0:
        raise HTTPException(status_code=400, detail="deposit must be positive")
    if withdrawal and deposit:
        raise HTTPException(status_code=400, detail="withdrawal and deposit cannot both be set")
    try:
        transaction = store.update_transaction(
            scope,
            transaction_id,
            date=payload.date,
            withdrawal=withdrawal,
            deposit=deposit,
            memo=payload.memo,
            txn_type=payload.type,
            row_color=payload.row_color,
            user_order=payload.user_order,
            account_id=payload.account_id,
            tags=_encode_tags(payload.tags) if payload.tags is not None else None,
        )
    except KeyError:
        raise HTTPException(status_code=404, detail="Transaction not found") from None
    return LedgerTransactionPayload(**transaction)


@router.delete("/transactions/{transaction_id}")
def delete_transaction(
    transaction_id: str,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> dict:
    try:
        store.delete_transaction(scope, transaction_id)
    except KeyError:
        raise HTTPException(status_code=404, detail="Transaction not found") from None
    return {"status": "ok"}


@router.post("/transactions/reorder")
def reorder_transactions(
    payload: LedgerTransactionsReorderRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> dict:
    if not payload.items:
        return {"status": "ok", "updated": 0}
    updates = [(item.transaction_id, item.user_order) for item in payload.items]
    store.reorder_transactions(scope, updates)
    return {"status": "ok", "updated": len(updates)}


@router.post("/import", response_model=LedgerStateResponse)
def import_ledger(
    payload: LedgerImportRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerStateResponse:
    case = _resolve_case_id(scope, store, payload.case_id)
    store.replace_all(
        scope,
        case_id=case["id"],
        accounts=[item.model_dump(by_alias=True) for item in payload.accounts],
        transactions=[item.model_dump(by_alias=True) for item in payload.transactions],
    )
    snapshot = store.snapshot(scope, case["id"])
    cases = store.list_cases(scope)
    return LedgerStateResponse(
        status="ok",
        case=LedgerCasePayload(**case),
        cases=_serialize_cases(cases),
        accounts=[LedgerAccountPayload(**account) for account in snapshot["accounts"]],
        transactions=[LedgerTransactionPayload(**txn) for txn in snapshot["transactions"]],
    )


@router.get("/export", response_model=LedgerExportResponse)
def export_ledger(
    case_id: Optional[str] = Query(None, alias="case_id"),
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerExportResponse:
    case = _resolve_case_id(scope, store, case_id)
    snapshot = store.snapshot(scope, case["id"])
    exported_at = datetime.now(timezone.utc).isoformat()
    cases = store.list_cases(scope)
    return LedgerExportResponse(
        status="ok",
        exported_at=exported_at,
        case=LedgerCasePayload(**case),
        cases=_serialize_cases(cases),
        accounts=[LedgerAccountPayload(**account) for account in snapshot["accounts"]],
        transactions=[LedgerTransactionPayload(**txn) for txn in snapshot["transactions"]],
    )


def _load_job_assets(job_id: str) -> List[dict]:
    manager = job_registry.job_manager
    if not manager:
        raise HTTPException(status_code=503, detail="Job manager is not available")
    record = manager.get(job_id)
    if not record or record.status != "completed":
        raise HTTPException(status_code=404, detail="Job not found or not completed")
    assets_payload = record.assets_payload or []
    return assets_payload


def _build_preview_accounts(assets_payload: List[dict]) -> List[dict]:
    preview_accounts: List[dict] = []
    for index, asset in enumerate(assets_payload, start=1):
        transactions = asset.get("transactions") or []
        total_withdrawal = sum(int(txn.get("withdrawal_amount") or 0) for txn in transactions)
        total_deposit = sum(int(txn.get("deposit_amount") or 0) for txn in transactions)
        asset_id = (
            asset.get("record_id")
            or (asset.get("identifiers") or {}).get("primary")
            or f"asset_{index:04d}"
        )
        # Use source_document filename (without extension) as default account name
        source_doc = asset.get("source_document") or ""
        if source_doc:
            # Remove file extension from filename
            account_name = source_doc.rsplit(".", 1)[0] if "." in source_doc else source_doc
        else:
            account_name = asset.get("asset_name") or "預金口座"
        preview_accounts.append(
            {
                "assetId": str(asset_id),
                "accountName": account_name,
                "accountNumber": (asset.get("identifiers") or {}).get("primary"),
                "ownerName": asset.get("owner_name") or [],
                "transactionCount": len(transactions),
                "totalWithdrawal": total_withdrawal,
                "totalDeposit": total_deposit,
                "sampleTransactions": [
                    {
                        "transaction_date": txn.get("transaction_date"),
                        "description": txn.get("description"),
                        "withdrawal_amount": int(txn.get("withdrawal_amount") or 0),
                        "deposit_amount": int(txn.get("deposit_amount") or 0),
                        "memo": txn.get("correction_note") or txn.get("memo"),
                    }
                    for txn in transactions[:3]
                ],
                "_transactions": transactions,
            }
        )
    return preview_accounts


@router.get("/jobs/{job_id}/preview", response_model=LedgerJobPreviewResponse)
def preview_job(
    job_id: str,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerJobPreviewResponse:
    _ = store  # store is unused but kept for dependency symmetry
    assets_payload = _load_job_assets(job_id)
    preview_accounts = _build_preview_accounts(assets_payload)
    response_accounts = [
        LedgerJobPreviewAccount(**{k: v for k, v in account.items() if not k.startswith("_")})
        for account in preview_accounts
    ]
    return LedgerJobPreviewResponse(status="ok", job_id=job_id, accounts=response_accounts)


@router.post("/jobs/{job_id}/import")
def import_job_assets(
    job_id: str,
    payload: LedgerJobImportRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> Dict[str, str]:
    assets_payload = _load_job_assets(job_id)
    preview_accounts = _build_preview_accounts(assets_payload)
    preview_map = {account["assetId"]: account for account in preview_accounts}
    if not payload.mappings:
        raise HTTPException(status_code=400, detail="No mappings provided")

    if payload.case_id:
        case = _resolve_case_id(scope, store, payload.case_id)
    elif payload.new_case_name:
        case = store.create_case(scope, payload.new_case_name)
    else:
        case = _resolve_case_id(scope, store, None)

    existing_accounts = {acc["id"]: acc for acc in store.list_accounts(scope, case["id"])}
    created_account_ids: List[str] = []
    merged_account_ids: List[str] = []

    def serialize_transactions(source: dict) -> List[Tuple[Optional[str], int, int, Optional[str], Optional[str], Optional[str]]]:
        transactions = source.get("_transactions") or []
        return [
            (
                txn.get("transaction_date"),
                int(txn.get("withdrawal_amount") or 0),
                int(txn.get("deposit_amount") or 0),
                txn.get("description") or txn.get("memo"),
                txn.get("description"),
                _encode_tags(txn.get("tags")) if txn.get("tags") else "",
            )
            for txn in transactions
        ]

    grouped_map: Dict[str, List[LedgerJobImportMapping]] = {}
    ordered_entries: List[Tuple[str, object]] = []

    for mapping in payload.mappings:
        group_key = None
        if getattr(mapping, "group_key", None):
            group_key = mapping.group_key.strip()
        if group_key and mapping.mode != "merge":
            if group_key not in grouped_map:
                grouped_map[group_key] = []
                ordered_entries.append(("group", group_key))
            grouped_map[group_key].append(mapping)
        else:
            ordered_entries.append(("single", mapping))

    processed_groups: Set[str] = set()

    for entry_type, entry_value in ordered_entries:
        if entry_type == "group":
            group_key = entry_value  # type: ignore[assignment]
            if group_key in processed_groups:
                continue
            processed_groups.add(group_key)
            group_mappings = grouped_map.get(group_key) or []
            if not group_mappings:
                continue
            combined_transactions: List[Tuple[Optional[str], int, int, Optional[str], Optional[str]]] = []
            account_name = None
            account_number = None
            holder_name = None
            for group_mapping in group_mappings:
                source = preview_map.get(group_mapping.asset_id)
                if not source:
                    continue
                if account_name is None:
                    account_name = (
                        group_mapping.group_name
                        or group_mapping.account_name
                        or source.get("accountName")
                        or "預金口座"
                    )
                if account_number is None:
                    account_number = group_mapping.group_number or group_mapping.account_number or source.get("accountNumber")
                if holder_name is None:
                    holder_name = (
                        group_mapping.group_holder_name
                        or group_mapping.holder_name
                        or (" / ".join(source.get("ownerName") or []) or None)
                    )
                combined_transactions.extend(serialize_transactions(source))
            if not combined_transactions:
                continue
            new_account = store.create_account(
                scope,
                case_id=case["id"],
                name=account_name or "預金口座",
                number=account_number,
                holder_name=holder_name,
            )
            store.bulk_insert_transactions(scope, new_account["id"], combined_transactions)
            created_account_ids.append(new_account["id"])
            continue

        mapping = entry_value  # type: ignore[assignment]
        source = preview_map.get(mapping.asset_id)
        if not source:
            continue
        serialized_transactions = serialize_transactions(source)
        if mapping.mode == "merge":
            target_id = mapping.target_account_id
            if not target_id or target_id not in existing_accounts:
                raise HTTPException(status_code=400, detail="Invalid target account for merge")
            store.bulk_insert_transactions(scope, target_id, serialized_transactions)
            merged_account_ids.append(target_id)
        else:
            account_name = mapping.account_name or source.get("accountName") or "預金口座"
            account_number = mapping.account_number or source.get("accountNumber")
            holder_name = mapping.holder_name or (" / ".join(source.get("ownerName") or []) or None)
            new_account = store.create_account(
                scope,
                case_id=case["id"],
                name=account_name,
                number=account_number,
                holder_name=holder_name,
            )
            store.bulk_insert_transactions(scope, new_account["id"], serialized_transactions)
            created_account_ids.append(new_account["id"])

    return {
        "status": "ok",
        "caseId": case["id"],
        "createdAccountIds": created_account_ids,
        "mergedAccountIds": merged_account_ids,
    }


@router.post("/analyze")
def analyze_transactions(
    case_id: Optional[str] = Query(None, alias="case_id"),
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> dict:
    """Analyze transactions using Gemini AI for inheritance tax insights."""
    from .gemini import GeminiClient
    from .config import get_settings

    case = _resolve_case_id(scope, store, case_id)
    snapshot = store.snapshot(scope, case["id"])
    accounts = snapshot["accounts"]
    transactions = snapshot["transactions"]

    if not transactions:
        return {
            "status": "ok",
            "findings": [],
            "summary": "取引データがありません。",
        }

    # Build account lookup
    account_map = {acc["id"]: acc for acc in accounts}

    # Format transactions for analysis
    txn_lines = []
    for txn in transactions:
        acc = account_map.get(txn.get("account_id"), {})
        holder = acc.get("holder_name") or "不明"
        acc_name = acc.get("name") or "不明"
        date = txn.get("date") or "不明"
        withdrawal = txn.get("withdrawal") or 0
        deposit = txn.get("deposit") or 0
        memo = txn.get("memo") or txn.get("type") or ""
        if withdrawal > 0:
            txn_lines.append(f"{date} | {holder} | {acc_name} | 出金 {withdrawal:,}円 | {memo}")
        if deposit > 0:
            txn_lines.append(f"{date} | {holder} | {acc_name} | 入金 {deposit:,}円 | {memo}")

    txn_text = "\n".join(txn_lines[:500])  # Limit to 500 transactions

    prompt = f"""あなたは相続税申告の専門家です。以下の銀行取引履歴を分析し、相続税申告において注意すべき点を指摘してください。

## 分析観点
1. **贈与税の検討**: 個人名義への入金で年間110万円を超えるものがあれば指摘
2. **保険関連**: 保険料の支払いや保険金の受取があれば指摘（生命保険契約に関する権利として申告が必要な可能性）
3. **不動産関連**: 固定資産税、管理費、賃料収入などがあれば不動産の有無を確認
4. **有価証券**: 配当金、株式売買などがあれば有価証券の有無を確認
5. **定期的な大口入出金**: パターンのある大口取引があれば背景を確認
6. **死亡直前の出金**: 被相続人の死亡直前の大口出金は手許現金として申告が必要

## 取引データ
{txn_text}

## 出力形式
以下のJSON形式で回答してください。findingsは最大10件まで。
```json
{{
  "findings": [
    {{
      "category": "カテゴリ名（贈与税の検討/保険関連/不動産関連/有価証券/定期的取引/その他）",
      "severity": "重要度（high/medium/low）",
      "title": "タイトル（20文字以内）",
      "description": "詳細説明（100文字以内）",
      "relatedTransactions": ["関連する取引の日付と内容（最大3件）"]
    }}
  ],
  "summary": "総評（200文字以内）"
}}
```
"""

    settings = get_settings()
    if not settings.gemini_api_key:
        return {
            "status": "error",
            "message": "Gemini APIキーが設定されていません。",
            "findings": [],
            "summary": "",
        }

    try:
        client = GeminiClient(api_keys=settings.gemini_api_key, model=settings.gemini_model)
        response_text = client.analyze_text(prompt)

        # Extract JSON from response
        import json
        import re
        json_match = re.search(r"```json\s*(.*?)\s*```", response_text, re.DOTALL)
        if json_match:
            result = json.loads(json_match.group(1))
        else:
            # Try parsing the whole response as JSON
            result = json.loads(response_text)

        return {
            "status": "ok",
            "findings": result.get("findings", []),
            "summary": result.get("summary", ""),
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"分析中にエラーが発生しました: {str(e)}",
            "findings": [],
            "summary": "",
        }


__all__ = ["router", "get_ledger_store"]
