"""SQLite-backed storage for access logs."""
from __future__ import annotations

import sqlite3
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional


class AnalyticsStore:
    """SQLite-backed storage for API access logs."""

    def __init__(self, db_path: Path) -> None:
        self._db_path = Path(db_path)
        self._db_path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._db_path, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn

    def _initialize(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS access_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    endpoint TEXT NOT NULL,
                    method TEXT NOT NULL,
                    client_type TEXT NOT NULL DEFAULT 'web',
                    doc_type TEXT,
                    status_code INTEGER,
                    duration_ms INTEGER,
                    user_agent TEXT,
                    ip_address TEXT,
                    extra TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_logs_timestamp ON access_logs(timestamp);
                CREATE INDEX IF NOT EXISTS idx_logs_endpoint ON access_logs(endpoint);
                CREATE INDEX IF NOT EXISTS idx_logs_client_type ON access_logs(client_type);
                """
            )

    def log_access(
        self,
        endpoint: str,
        method: str,
        client_type: str = "web",
        doc_type: Optional[str] = None,
        status_code: Optional[int] = None,
        duration_ms: Optional[int] = None,
        user_agent: Optional[str] = None,
        ip_address: Optional[str] = None,
        extra: Optional[str] = None,
    ) -> int:
        """Log an API access event."""
        with self._lock:
            with self._connect() as conn:
                cursor = conn.execute(
                    """
                    INSERT INTO access_logs
                    (timestamp, endpoint, method, client_type, doc_type, status_code, duration_ms, user_agent, ip_address, extra)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        datetime.utcnow().isoformat(),
                        endpoint,
                        method,
                        client_type,
                        doc_type,
                        status_code,
                        duration_ms,
                        user_agent,
                        ip_address,
                        extra,
                    ),
                )
                return cursor.lastrowid or 0

    def get_daily_stats(self, start_date: str, end_date: str) -> List[Dict[str, Any]]:
        """Get daily access counts between dates."""
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    date(timestamp) as date,
                    client_type,
                    COUNT(*) as count
                FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
                GROUP BY date(timestamp), client_type
                ORDER BY date(timestamp)
                """,
                (start_date, end_date),
            ).fetchall()
            return [dict(row) for row in rows]

    def get_endpoint_stats(self, start_date: str, end_date: str) -> List[Dict[str, Any]]:
        """Get endpoint usage counts between dates."""
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    endpoint,
                    client_type,
                    COUNT(*) as count,
                    AVG(duration_ms) as avg_duration_ms
                FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
                GROUP BY endpoint, client_type
                ORDER BY count DESC
                """,
                (start_date, end_date),
            ).fetchall()
            return [dict(row) for row in rows]

    def get_doc_type_stats(self, start_date: str, end_date: str) -> List[Dict[str, Any]]:
        """Get document type usage counts between dates."""
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    doc_type,
                    client_type,
                    COUNT(*) as count
                FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
                  AND doc_type IS NOT NULL
                GROUP BY doc_type, client_type
                ORDER BY count DESC
                """,
                (start_date, end_date),
            ).fetchall()
            return [dict(row) for row in rows]

    def get_recent_logs(self, limit: int = 100) -> List[Dict[str, Any]]:
        """Get recent access logs."""
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT * FROM access_logs
                ORDER BY timestamp DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
            return [dict(row) for row in rows]

    def get_summary(self, start_date: str, end_date: str) -> Dict[str, Any]:
        """Get summary statistics between dates."""
        with self._connect() as conn:
            # Total counts by client type
            client_counts = conn.execute(
                """
                SELECT client_type, COUNT(*) as count
                FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
                GROUP BY client_type
                """,
                (start_date, end_date),
            ).fetchall()

            # Total unique endpoints
            endpoint_count = conn.execute(
                """
                SELECT COUNT(DISTINCT endpoint) as count
                FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
                """,
                (start_date, end_date),
            ).fetchone()

            # Average duration
            avg_duration = conn.execute(
                """
                SELECT AVG(duration_ms) as avg_ms
                FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
                  AND duration_ms IS NOT NULL
                """,
                (start_date, end_date),
            ).fetchone()

            return {
                "client_counts": {row["client_type"]: row["count"] for row in client_counts},
                "total_requests": sum(row["count"] for row in client_counts),
                "unique_endpoints": endpoint_count["count"] if endpoint_count else 0,
                "avg_duration_ms": round(avg_duration["avg_ms"], 2) if avg_duration and avg_duration["avg_ms"] else None,
            }
