"""SQLite-backed storage for access logs."""
from __future__ import annotations

import sqlite3
import threading
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

# JST timezone (UTC+9)
JST = timezone(timedelta(hours=9))


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
                        datetime.now(JST).strftime("%Y-%m-%dT%H:%M:%S"),
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

            # Session & user estimation
            session_data = self._estimate_sessions(conn, start_date, end_date)

            # Cost and time estimates
            cost_time = self._estimate_cost_and_time(
                conn, start_date, end_date, session_data
            )

            return {
                "client_counts": {row["client_type"]: row["count"] for row in client_counts},
                "total_requests": sum(row["count"] for row in client_counts),
                "unique_endpoints": endpoint_count["count"] if endpoint_count else 0,
                "avg_duration_ms": round(avg_duration["avg_ms"], 2) if avg_duration and avg_duration["avg_ms"] else None,
                **session_data,
                **cost_time,
            }

    # ------------------------------------------------------------------
    # セッション推定 / ユニークユーザー数
    # ------------------------------------------------------------------

    def _estimate_sessions(
        self, conn: sqlite3.Connection, start_date: str, end_date: str
    ) -> Dict[str, Any]:
        """Estimate actual usage sessions and unique users.

        VBA 1回の実行 = POST(1) + GET polling(10-20) + GET result(1) なので、
        生リクエスト数ではなく「セッション」で実利用回数を推定する。
        """

        # ユニークIP数（≒ユーザー数）per client type
        unique_ips = conn.execute(
            """
            SELECT client_type, COUNT(DISTINCT ip_address) as cnt
            FROM access_logs
            WHERE date(timestamp) >= ? AND date(timestamp) <= ?
              AND ip_address IS NOT NULL
            GROUP BY client_type
            """,
            (start_date, end_date),
        ).fetchall()

        # セッション推定: 同一IP + 30分枠でグループ化
        sessions = conn.execute(
            """
            SELECT client_type, COUNT(*) as cnt
            FROM (
                SELECT DISTINCT
                    client_type,
                    COALESCE(ip_address, 'unknown') || '_' ||
                    strftime('%Y-%m-%d-%H-', timestamp) ||
                    CAST(CAST(strftime('%M', timestamp) AS INTEGER) / 30 AS TEXT)
                    as session_key
                FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
            )
            GROUP BY client_type
            """,
            (start_date, end_date),
        ).fetchall()

        # 完了解析数: ジョブ結果取得のユニーク数
        # GET /api/jobs/{id}/result の成功数 = 実際にGeminiを呼んだ回数
        completed = conn.execute(
            """
            SELECT client_type, COUNT(DISTINCT endpoint) as cnt
            FROM access_logs
            WHERE endpoint LIKE '/api/jobs/%/result'
              AND method = 'GET'
              AND (status_code IS NULL OR status_code < 400)
              AND date(timestamp) >= ? AND date(timestamp) <= ?
            GROUP BY client_type
            """,
            (start_date, end_date),
        ).fetchall()

        total_completed = sum(row["cnt"] for row in completed)

        return {
            "unique_users": {row["client_type"]: row["cnt"] for row in unique_ips},
            "sessions": {row["client_type"]: row["cnt"] for row in sessions},
            "completed_analyses": total_completed,
            "completed_by_client": {row["client_type"]: row["cnt"] for row in completed},
        }

    # ------------------------------------------------------------------
    # API料金概算 / 削減時間の推計
    # ------------------------------------------------------------------
    _COST_PER_PDF_ANALYSIS = 15     # Gemini 2.5 Pro: 約¥15/リクエスト
    _COST_PER_JON_BATCH = 5         # JON API: 約¥5/バッチ
    _COST_PER_JON_SINGLE = 2        # JON 個別API: 約¥2/リクエスト

    _MANUAL_MINUTES_PER_PDF = 30    # 通帳1冊の目視確認+入力: 約30分
    _MANUAL_MINUTES_PER_JON = 15    # 不動産1物件の登記取得+路線価確認: 約15分

    def _estimate_cost_and_time(
        self,
        conn: sqlite3.Connection,
        start_date: str,
        end_date: str,
        session_data: Dict[str, Any],
    ) -> Dict[str, Any]:
        """Estimate API costs and time savings."""

        def _count_endpoint(pattern: str, method: str = "POST") -> int:
            row = conn.execute(
                """
                SELECT COUNT(*) as cnt FROM access_logs
                WHERE date(timestamp) >= ? AND date(timestamp) <= ?
                  AND endpoint LIKE ? AND method = ?
                  AND (status_code IS NULL OR status_code < 400)
                """,
                (start_date, end_date, pattern, method),
            ).fetchone()
            return row["cnt"] if row else 0

        pdf_post_count = _count_endpoint("%/analyze%")
        jon_batch_count = _count_endpoint("%/jon/batch%")
        jon_single_count = (
            _count_endpoint("%/jon/locating%")
            + _count_endpoint("%/jon/rosenka%")
            + _count_endpoint("%/jon/registration%")
        )

        # completed_analyses（ジョブ結果取得数）をフォールバックに使用
        # Volume取り付け前のPOSTが消失していてもコスト計算が動く
        completed = session_data.get("completed_analyses", 0)
        pdf_count = max(pdf_post_count, completed)

        # 概算API料金（円）
        estimated_cost_yen = (
            pdf_count * self._COST_PER_PDF_ANALYSIS
            + jon_batch_count * self._COST_PER_JON_BATCH
            + jon_single_count * self._COST_PER_JON_SINGLE
        )

        # 削減時間（分）
        saved_minutes = (
            pdf_count * self._MANUAL_MINUTES_PER_PDF
            + jon_batch_count * self._MANUAL_MINUTES_PER_JON
        )

        # Per-client completed analyses
        completed_by_client = session_data.get("completed_by_client", {})

        return {
            "pdf_analysis_count": pdf_count,
            "jon_batch_count": jon_batch_count,
            "jon_single_count": jon_single_count,
            "estimated_cost_yen": estimated_cost_yen,
            "saved_minutes": saved_minutes,
            "excel_analysis_count": completed_by_client.get("excel", 0),
            "web_analysis_count": completed_by_client.get("web", 0),
        }
