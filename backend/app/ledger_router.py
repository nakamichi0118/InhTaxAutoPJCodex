from __future__ import annotations

import threading
import uuid
from datetime import datetime, timezone
from typing import Optional

from fastapi import APIRouter, Depends, Header, HTTPException, Query

from .config import get_settings
from .ledger_store import LedgerScope, LedgerStore
from .models import (
    LedgerAccountCreateRequest,
    LedgerAccountPayload,
    LedgerAccountReorderRequest,
    LedgerAccountUpdateRequest,
    LedgerExportResponse,
    LedgerImportRequest,
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


@router.post("/session", response_model=LedgerSessionResponse)
def create_session(payload: LedgerSessionRequest) -> LedgerSessionResponse:
    app_id = (payload.app_id or "ledger-app").strip()
    session_token = payload.session_token or uuid.uuid4().hex
    return LedgerSessionResponse(status="ok", app_id=app_id, user_id=session_token, session_token=session_token)


@router.get("/state", response_model=LedgerStateResponse)
def get_state(scope: LedgerScope = Depends(get_scope), store: LedgerStore = Depends(get_ledger_store)) -> LedgerStateResponse:
    snapshot = store.snapshot(scope)
    return LedgerStateResponse(status="ok", accounts=snapshot["accounts"], transactions=snapshot["transactions"])


@router.post("/accounts", response_model=LedgerAccountPayload)
def create_account(
    payload: LedgerAccountCreateRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerAccountPayload:
    account = store.create_account(scope, name=payload.name, number=payload.number, order=payload.order)
    return LedgerAccountPayload(**account)


@router.patch("/accounts/{account_id}", response_model=LedgerAccountPayload)
def update_account(
    account_id: str,
    payload: LedgerAccountUpdateRequest,
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerAccountPayload:
    try:
        account = store.update_account(scope, account_id, name=payload.name, number=payload.number, order=payload.order)
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
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> dict:
    if not payload.items:
        return {"status": "ok", "updated": 0}
    updates = [(item.account_id, item.order) for item in payload.items]
    store.reorder_accounts(scope, updates)
    return {"status": "ok", "updated": len(updates)}


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
    store.replace_all(scope, accounts=[item.model_dump() for item in payload.accounts], transactions=[item.model_dump() for item in payload.transactions])
    snapshot = store.snapshot(scope)
    return LedgerStateResponse(status="ok", accounts=snapshot["accounts"], transactions=snapshot["transactions"])


@router.get("/export", response_model=LedgerExportResponse)
def export_ledger(
    scope: LedgerScope = Depends(get_scope),
    store: LedgerStore = Depends(get_ledger_store),
) -> LedgerExportResponse:
    snapshot = store.snapshot(scope)
    exported_at = datetime.now(timezone.utc).isoformat()
    return LedgerExportResponse(status="ok", exported_at=exported_at, accounts=snapshot["accounts"], transactions=snapshot["transactions"])


__all__ = ["router", "get_ledger_store"]
