from __future__ import annotations

import sqlite3
import threading
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple


@dataclass(frozen=True)
class LedgerScope:
    app_id: str
    user_id: str


class LedgerStore:
    """SQLite-backed storage for ledger cases, accounts, and transactions."""

    def __init__(self, db_path: Path) -> None:
        self._db_path = Path(db_path)
        self._db_path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._initialize()

    # ------------------------------------------------------------------
    # Low-level helpers
    # ------------------------------------------------------------------

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
                CREATE TABLE IF NOT EXISTS ledger_cases (
                    id TEXT PRIMARY KEY,
                    app_id TEXT NOT NULL,
                    user_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS ledger_accounts (
                    id TEXT PRIMARY KEY,
                    app_id TEXT NOT NULL,
                    user_id TEXT NOT NULL,
                    case_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    number TEXT,
                    display_order REAL NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(case_id) REFERENCES ledger_cases(id) ON DELETE CASCADE
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
                CREATE INDEX IF NOT EXISTS idx_cases_scope ON ledger_cases(app_id, user_id);
                CREATE INDEX IF NOT EXISTS idx_accounts_scope_case ON ledger_accounts(app_id, user_id, case_id);
                CREATE INDEX IF NOT EXISTS idx_transactions_scope ON ledger_transactions(app_id, user_id);
                CREATE INDEX IF NOT EXISTS idx_transactions_account ON ledger_transactions(account_id);
                """
            )
            self._ensure_case_column(conn)

    def _ensure_case_column(self, conn: sqlite3.Connection) -> None:
        cursor = conn.execute("PRAGMA table_info(ledger_accounts)")
        columns = {row[1] for row in cursor.fetchall()}
        if "case_id" not in columns:
            conn.execute("ALTER TABLE ledger_accounts ADD COLUMN case_id TEXT")
            conn.execute("UPDATE ledger_accounts SET case_id = '' WHERE case_id IS NULL")

    # ------------------------------------------------------------------
    # Case operations
    # ------------------------------------------------------------------

    def list_cases(self, scope: LedgerScope) -> List[dict]:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, name, created_at, updated_at
                FROM ledger_cases
                WHERE app_id = ? AND user_id = ?
                ORDER BY created_at ASC
                """,
                (scope.app_id, scope.user_id),
            )
            return [self._serialize_case(row, scope) for row in cursor.fetchall()]

    def _serialize_case(self, row: sqlite3.Row, scope: LedgerScope) -> dict:
        return {
            "id": row["id"],
            "name": row["name"],
            "user_id": scope.user_id,
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
        }

    def create_case(self, scope: LedgerScope, name: str) -> dict:
        case_id = uuid.uuid4().hex
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO ledger_cases (id, app_id, user_id, name)
                VALUES (?, ?, ?, ?)
                """,
                (case_id, scope.app_id, scope.user_id, name.strip() or "案件"),
            )
        return self.get_case(scope, case_id)

    def get_case(self, scope: LedgerScope, case_id: str) -> dict:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, name, created_at, updated_at
                FROM ledger_cases
                WHERE app_id = ? AND user_id = ? AND id = ?
                """,
                (scope.app_id, scope.user_id, case_id),
            )
            row = cursor.fetchone()
            if not row:
                raise KeyError("case_not_found")
            return self._serialize_case(row, scope)

    def get_or_create_default_case(self, scope: LedgerScope) -> dict:
        cases = self.list_cases(scope)
        if cases:
            return cases[0]
        return self.create_case(scope, "案件")

    # ------------------------------------------------------------------
    # Account operations
    # ------------------------------------------------------------------

    def list_accounts(self, scope: LedgerScope, case_id: str) -> List[dict]:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, name, number, case_id, display_order, created_at, updated_at
                FROM ledger_accounts
                WHERE app_id = ? AND user_id = ? AND case_id = ?
                ORDER BY display_order ASC, created_at ASC
                """,
                (scope.app_id, scope.user_id, case_id),
            )
            return [self._serialize_account(row, scope) for row in cursor.fetchall()]

    def _next_account_order(self, conn: sqlite3.Connection, scope: LedgerScope, case_id: str) -> float:
        cursor = conn.execute(
            """
            SELECT COALESCE(MAX(display_order), 0)
            FROM ledger_accounts
            WHERE app_id = ? AND user_id = ? AND case_id = ?
            """,
            (scope.app_id, scope.user_id, case_id),
        )
        current = cursor.fetchone()[0]
        return float(current or 0) + 1000.0

    def create_account(
        self,
        scope: LedgerScope,
        *,
        case_id: str,
        name: str,
        number: Optional[str],
        order: Optional[float] = None,
    ) -> dict:
        account_id = uuid.uuid4().hex
        with self._connect() as conn:
            conn.execute("BEGIN")
            resolved_order = order if order is not None else self._next_account_order(conn, scope, case_id)
            conn.execute(
                """
                INSERT INTO ledger_accounts (id, app_id, user_id, case_id, name, number, display_order)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (account_id, scope.app_id, scope.user_id, case_id, name.strip(), (number or "").strip(), resolved_order),
            )
            conn.commit()
        return self.get_account(scope, account_id)

    def get_account(self, scope: LedgerScope, account_id: str) -> dict:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT id, name, number, case_id, display_order, created_at, updated_at
                FROM ledger_accounts
                WHERE app_id = ? AND user_id = ? AND id = ?
                """,
                (scope.app_id, scope.user_id, account_id),
            )
            row = cursor.fetchone()
            if not row:
                raise KeyError("account_not_found")
            return self._serialize_account(row, scope)

    def update_account(
        self,
        scope: LedgerScope,
        account_id: str,
        *,
        name: Optional[str] = None,
        number: Optional[str] = None,
        order: Optional[float] = None,
    ) -> dict:
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

    def reorder_accounts(self, scope: LedgerScope, case_id: str, items: Iterable[tuple[str, float]]) -> None:
        with self._connect() as conn:
            conn.execute("BEGIN")
            for account_id, order in items:
                conn.execute(
                    """
                    UPDATE ledger_accounts
                    SET display_order = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE app_id = ? AND user_id = ? AND id = ? AND case_id = ?
                    """,
                    (order, scope.app_id, scope.user_id, account_id, case_id),
                )
            conn.commit()

    # ------------------------------------------------------------------
    # Transaction operations
    # ------------------------------------------------------------------

    def list_transactions(self, scope: LedgerScope, case_id: str) -> List[dict]:
        with self._connect() as conn:
            cursor = conn.execute(
                """
                SELECT t.id, t.account_id, t.date, t.withdrawal, t.deposit, t.memo, t.type,
                       t.row_color, t.user_order, t.created_at, t.updated_at
                FROM ledger_transactions AS t
                INNER JOIN ledger_accounts AS a ON t.account_id = a.id
                WHERE t.app_id = ? AND t.user_id = ? AND a.case_id = ?
                ORDER BY COALESCE(t.user_order, CAST(strftime('%s', t.date || ' 00:00:00') AS REAL)), t.date ASC, t.id ASC
                """,
                (scope.app_id, scope.user_id, case_id),
            )
            return [self._serialize_transaction(row, scope) for row in cursor.fetchall()]

    def _account_exists(self, conn: sqlite3.Connection, scope: LedgerScope, account_id: str) -> bool:
        cursor = conn.execute(
            "SELECT 1 FROM ledger_accounts WHERE app_id = ? AND user_id = ? AND id = ?",
            (scope.app_id, scope.user_id, account_id),
        )
        return cursor.fetchone() is not None

    def _max_user_order_for_account(self, conn: sqlite3.Connection, scope: LedgerScope, account_id: str) -> float:
        cursor = conn.execute(
            """
            SELECT COALESCE(MAX(user_order), 0)
            FROM ledger_transactions
            WHERE app_id = ? AND user_id = ? AND account_id = ?
            """,
            (scope.app_id, scope.user_id, account_id),
        )
        return float(cursor.fetchone()[0] or 0)

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
            resolved_order = user_order
            if resolved_order is None:
                resolved_order = self._max_user_order_for_account(conn, scope, account_id) + 1000.0
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
                    resolved_order,
                    (row_color or None),
                ),
            )
            conn.commit()
        return self.get_transaction(scope, txn_id)

    def bulk_insert_transactions(
        self,
        scope: LedgerScope,
        account_id: str,
        records: Sequence[Tuple[str, int, int, Optional[str], Optional[str]]],
    ) -> None:
        if not records:
            return
        with self._connect() as conn:
            conn.execute("BEGIN")
            if not self._account_exists(conn, scope, account_id):
                raise KeyError("account_not_found")
            base_order = self._max_user_order_for_account(conn, scope, account_id) + 1000.0
            for index, (date, withdrawal, deposit, memo, txn_type) in enumerate(records, start=1):
                txn_id = uuid.uuid4().hex
                conn.execute(
                    """
                    INSERT INTO ledger_transactions (id, app_id, user_id, account_id, date, withdrawal, deposit, memo, type, user_order)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                        base_order + index,
                    ),
                )
            conn.commit()

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

    def replace_all(self, scope: LedgerScope, *, case_id: str, accounts: List[dict], transactions: List[dict]) -> None:
        with self._connect() as conn:
            conn.execute("BEGIN")
            conn.execute(
                "DELETE FROM ledger_transactions WHERE app_id = ? AND user_id = ? AND account_id IN (SELECT id FROM ledger_accounts WHERE case_id = ?)",
                (scope.app_id, scope.user_id, case_id),
            )
            conn.execute(
                "DELETE FROM ledger_accounts WHERE app_id = ? AND user_id = ? AND case_id = ?",
                (scope.app_id, scope.user_id, case_id),
            )
            for account in accounts:
                conn.execute(
                    """
                    INSERT INTO ledger_accounts (id, app_id, user_id, case_id, name, number, display_order)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        account.get("id") or uuid.uuid4().hex,
                        scope.app_id,
                        scope.user_id,
                        case_id,
                        (account.get("name") or "").strip(),
                        (account.get("number") or "").strip(),
                        float(account.get("order") or 0),
                    ),
                )
            for txn in transactions:
                account_id = txn.get("accountId") or txn.get("account_id")
                if not account_id:
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
                        txn.get("date"),
                        int(txn.get("withdrawal") or txn.get("withdrawal_amount") or 0),
                        int(txn.get("deposit") or txn.get("deposit_amount") or 0),
                        (txn.get("memo") or "").strip(),
                        (txn.get("type") or "").strip(),
                        txn.get("userOrder") or txn.get("user_order"),
                        txn.get("rowColor") or txn.get("row_color"),
                    ),
                )
            conn.commit()

    def snapshot(self, scope: LedgerScope, case_id: str) -> dict:
        return {
            "accounts": self.list_accounts(scope, case_id),
            "transactions": self.list_transactions(scope, case_id),
        }

    # ------------------------------------------------------------------
    # Serialization helpers
    # ------------------------------------------------------------------

    def _serialize_account(self, row: sqlite3.Row, scope: LedgerScope) -> dict:
        return {
            "id": row["id"],
            "case_id": row["case_id"],
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
            "rowColor": row["row_color"],
            "userOrder": row["user_order"],
            "user_id": scope.user_id,
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
        }


__all__ = ["LedgerStore", "LedgerScope"]
