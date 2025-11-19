from __future__ import annotations

import sqlite3
import threading
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional


@dataclass(frozen=True)
class LedgerScope:
    app_id: str
    user_id: str


class LedgerStore:
    """Lightweight SQLite-backed storage for ledger accounts and transactions."""

    def __init__(self, db_path: Path) -> None:
        self._db_path = Path(db_path)
        self._db_path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._db_path, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    def _initialize(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS ledger_accounts (
                    id TEXT PRIMARY KEY,
                    app_id TEXT NOT NULL,
                    user_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    number TEXT,
                    display_order REAL NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS ledger_transactions (
                    id TEXT PRIMARY KEY,
                    app_id TEXT NOT NULL,
                    user_id TEXT NOT NULL,
                    account_id TEXT NOT NULL,
                    date TEXT NOT NULL,
                    withdrawal INTEGER NOT NULL DEFAULT 0,
                    deposit INTEGER NOT NULL DEFAULT 0,
                    memo TEXT,
                    type TEXT,
                    row_color TEXT,
                    user_order REAL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(account_id) REFERENCES ledger_accounts(id) ON DELETE CASCADE
                );
                CREATE INDEX IF NOT EXISTS idx_ledger_accounts_scope ON ledger_accounts(app_id, user_id);
                CREATE INDEX IF NOT EXISTS idx_ledger_transactions_scope ON ledger_transactions(app_id, user_id);
                CREATE INDEX IF NOT EXISTS idx_ledger_transactions_account ON ledger_transactions(account_id);
                """
            )

    # ------------------------------------------------------------------
    # Account operations
    # ------------------------------------------------------------------

    def list_accounts(self, scope: LedgerScope) -> List[dict]:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, name, number, display_order, created_at, updated_at
                FROM ledger_accounts
                WHERE app_id = ? AND user_id = ?
                ORDER BY display_order ASC, created_at ASC
                """,
                (scope.app_id, scope.user_id),
            )
            return [self._serialize_account(row, scope) for row in cursor.fetchall()]

    def _next_account_order(self, conn: sqlite3.Connection, scope: LedgerScope) -> float:
        cursor = conn.execute(
            "SELECT COALESCE(MAX(display_order), 0) FROM ledger_accounts WHERE app_id = ? AND user_id = ?",
            (scope.app_id, scope.user_id),
        )
        current = cursor.fetchone()[0]
        return float(current or 0) + 1000.0

    def create_account(self, scope: LedgerScope, *, name: str, number: Optional[str], order: Optional[float] = None) -> dict:
        account_id = uuid.uuid4().hex
        with self._connect() as conn:
            conn.execute("BEGIN")
            resolved_order = order if order is not None else self._next_account_order(conn, scope)
            conn.execute(
                """
                INSERT INTO ledger_accounts (id, app_id, user_id, name, number, display_order)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (account_id, scope.app_id, scope.user_id, name.strip(), (number or "").strip(), resolved_order),
            )
            conn.commit()
        return self.get_account(scope, account_id)

    def get_account(self, scope: LedgerScope, account_id: str) -> dict:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, name, number, display_order, created_at, updated_at
                FROM ledger_accounts
                WHERE app_id = ? AND user_id = ? AND id = ?
                """,
                (scope.app_id, scope.user_id, account_id),
            )
            row = cursor.fetchone()
            if not row:
                raise KeyError("account_not_found")
            return self._serialize_account(row, scope)

    def update_account(self, scope: LedgerScope, account_id: str, *, name: Optional[str] = None, number: Optional[str] = None, order: Optional[float] = None) -> dict:
        updates = []
        params: List[object] = []
        if name is not None:
            updates.append("name = ?")
            params.append(name.strip())
        if number is not None:
            updates.append("number = ?")
            params.append(number.strip())
        if order is not None:
            updates.append("display_order = ?")
            params.append(order)
        if not updates:
            return self.get_account(scope, account_id)
        updates.append("updated_at = CURRENT_TIMESTAMP")
        params.extend([scope.app_id, scope.user_id, account_id])
        query = f"UPDATE ledger_accounts SET {', '.join(updates)} WHERE app_id = ? AND user_id = ? AND id = ?"
        with self._connect() as conn:
            cursor = conn.execute(query, tuple(params))
            if cursor.rowcount == 0:
                raise KeyError("account_not_found")
        return self.get_account(scope, account_id)

    def delete_account(self, scope: LedgerScope, account_id: str) -> None:
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM ledger_accounts WHERE app_id = ? AND user_id = ? AND id = ?",
                (scope.app_id, scope.user_id, account_id),
            )
            if cursor.rowcount == 0:
                raise KeyError("account_not_found")

    def reorder_accounts(self, scope: LedgerScope, items: Iterable[tuple[str, float]]) -> None:
        with self._connect() as conn:
            conn.execute("BEGIN")
            for account_id, order in items:
                conn.execute(
                    """
                    UPDATE ledger_accounts
                    SET display_order = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE app_id = ? AND user_id = ? AND id = ?
                    """,
                    (order, scope.app_id, scope.user_id, account_id),
                )
            conn.commit()

    # ------------------------------------------------------------------
    # Transaction operations
    # ------------------------------------------------------------------

    def list_transactions(self, scope: LedgerScope) -> List[dict]:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, account_id, date, withdrawal, deposit, memo, type, row_color, user_order, created_at, updated_at
                FROM ledger_transactions
                WHERE app_id = ? AND user_id = ?
                ORDER BY COALESCE(user_order, CAST(strftime('%s', date || ' 00:00:00') AS REAL)), date ASC, id ASC
                """,
                (scope.app_id, scope.user_id),
            )
            return [self._serialize_transaction(row, scope) for row in cursor.fetchall()]

    def _account_exists(self, conn: sqlite3.Connection, scope: LedgerScope, account_id: str) -> bool:
        cursor = conn.execute(
            "SELECT 1 FROM ledger_accounts WHERE app_id = ? AND user_id = ? AND id = ?",
            (scope.app_id, scope.user_id, account_id),
        )
        return cursor.fetchone() is not None

    def create_transaction(
        self,
        scope: LedgerScope,
        *,
        account_id: str,
        date: str,
        withdrawal: int,
        deposit: int,
        memo: Optional[str],
        txn_type: Optional[str],
        user_order: Optional[float] = None,
        row_color: Optional[str] = None,
    ) -> dict:
        if withdrawal and deposit:
            raise ValueError("withdrawal_and_deposit_conflict")
        txn_id = uuid.uuid4().hex
        with self._connect() as conn:
            conn.execute("BEGIN")
            if not self._account_exists(conn, scope, account_id):
                raise KeyError("account_not_found")
            conn.execute(
                """
                INSERT INTO ledger_transactions (id, app_id, user_id, account_id, date, withdrawal, deposit, memo, type, user_order, row_color)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    txn_id,
                    scope.app_id,
                    scope.user_id,
                    account_id,
                    date,
                    withdrawal,
                    deposit,
                    (memo or "").strip(),
                    (txn_type or "").strip(),
                    user_order,
                    (row_color or None),
                ),
            )
            conn.commit()
        return self.get_transaction(scope, txn_id)

    def get_transaction(self, scope: LedgerScope, transaction_id: str) -> dict:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, account_id, date, withdrawal, deposit, memo, type, row_color, user_order, created_at, updated_at
                FROM ledger_transactions
                WHERE app_id = ? AND user_id = ? AND id = ?
                """,
                (scope.app_id, scope.user_id, transaction_id),
            )
            row = cursor.fetchone()
            if not row:
                raise KeyError("transaction_not_found")
            return self._serialize_transaction(row, scope)

    def update_transaction(
        self,
        scope: LedgerScope,
        transaction_id: str,
        *,
        date: Optional[str] = None,
        withdrawal: Optional[int] = None,
        deposit: Optional[int] = None,
        memo: Optional[str] = None,
        txn_type: Optional[str] = None,
        row_color: Optional[str] = None,
        user_order: Optional[float] = None,
        account_id: Optional[str] = None,
    ) -> dict:
        updates = []
        params: List[object] = []
        if account_id is not None:
            updates.append("account_id = ?")
            params.append(account_id)
        if date is not None:
            updates.append("date = ?")
            params.append(date)
        if withdrawal is not None:
            updates.append("withdrawal = ?")
            params.append(withdrawal)
        if deposit is not None:
            updates.append("deposit = ?")
            params.append(deposit)
        if memo is not None:
            updates.append("memo = ?")
            params.append(memo.strip())
        if txn_type is not None:
            updates.append("type = ?")
            params.append(txn_type.strip())
        if row_color is not None:
            updates.append("row_color = ?")
            params.append(row_color)
        if user_order is not None:
            updates.append("user_order = ?")
            params.append(user_order)
        if not updates:
            return self.get_transaction(scope, transaction_id)
        updates.append("updated_at = CURRENT_TIMESTAMP")
        params.extend([scope.app_id, scope.user_id, transaction_id])
        query = f"UPDATE ledger_transactions SET {', '.join(updates)} WHERE app_id = ? AND user_id = ? AND id = ?"
        with self._connect() as conn:
            cursor = conn.execute(query, tuple(params))
            if cursor.rowcount == 0:
                raise KeyError("transaction_not_found")
        return self.get_transaction(scope, transaction_id)

    def delete_transaction(self, scope: LedgerScope, transaction_id: str) -> None:
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM ledger_transactions WHERE app_id = ? AND user_id = ? AND id = ?",
                (scope.app_id, scope.user_id, transaction_id),
            )
            if cursor.rowcount == 0:
                raise KeyError("transaction_not_found")

    def reorder_transactions(self, scope: LedgerScope, items: Iterable[tuple[str, float]]) -> None:
        with self._connect() as conn:
            conn.execute("BEGIN")
            for transaction_id, order in items:
                conn.execute(
                    """
                    UPDATE ledger_transactions
                    SET user_order = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE app_id = ? AND user_id = ? AND id = ?
                    """,
                    (order, scope.app_id, scope.user_id, transaction_id),
                )
            conn.commit()

    # ------------------------------------------------------------------
    # Import / export helpers
    # ------------------------------------------------------------------

    def replace_all(self, scope: LedgerScope, *, accounts: List[dict], transactions: List[dict]) -> None:
        with self._connect() as conn:
            conn.execute("BEGIN")
            conn.execute("DELETE FROM ledger_transactions WHERE app_id = ? AND user_id = ?", (scope.app_id, scope.user_id))
            conn.execute("DELETE FROM ledger_accounts WHERE app_id = ? AND user_id = ?", (scope.app_id, scope.user_id))
            for account in accounts:
                order_value = account.get("order")
                if order_value is None:
                    order_value = account.get("display_order")
                conn.execute(
                    """
                    INSERT INTO ledger_accounts (id, app_id, user_id, name, number, display_order)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        account.get("id") or uuid.uuid4().hex,
                        scope.app_id,
                        scope.user_id,
                        (account.get("name") or "").strip(),
                        (account.get("number") or "").strip(),
                        float(order_value or 0),
                    ),
                )
            for txn in transactions:
                account_id = txn.get("accountId") or txn.get("account_id")
                date_value = txn.get("date")
                if not account_id or not date_value:
                    continue
                conn.execute(
                    """
                    INSERT INTO ledger_transactions (id, app_id, user_id, account_id, date, withdrawal, deposit, memo, type, user_order, row_color)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        txn.get("id") or uuid.uuid4().hex,
                        scope.app_id,
                        scope.user_id,
                        account_id,
                        date_value,
                        int(txn.get("withdrawal") or txn.get("withdrawal_amount") or 0),
                        int(txn.get("deposit") or txn.get("deposit_amount") or 0),
                        (txn.get("memo") or "").strip(),
                        (txn.get("type") or "").strip(),
                        txn.get("userOrder") or txn.get("user_order"),
                        txn.get("rowColor") or txn.get("row_color"),
                    ),
                )
            conn.commit()

    def snapshot(self, scope: LedgerScope) -> dict:
        return {
            "accounts": self.list_accounts(scope),
            "transactions": self.list_transactions(scope),
        }

    # ------------------------------------------------------------------
    # Serialization helpers
    # ------------------------------------------------------------------

    def _serialize_account(self, row: sqlite3.Row, scope: LedgerScope) -> dict:
        return {
            "id": row["id"],
            "name": row["name"],
            "number": row["number"],
            "order": int(row["display_order"] or 0),
            "user_id": scope.user_id,
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
        }

    def _serialize_transaction(self, row: sqlite3.Row, scope: LedgerScope) -> dict:
        return {
            "id": row["id"],
            "account_id": row["account_id"],
            "date": row["date"],
            "withdrawal": int(row["withdrawal"] or 0),
            "deposit": int(row["deposit"] or 0),
            "memo": row["memo"],
            "type": row["type"],
            "row_color": row["row_color"],
            "user_order": row["user_order"],
            "user_id": scope.user_id,
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
        }


__all__ = ["LedgerStore", "LedgerScope"]
