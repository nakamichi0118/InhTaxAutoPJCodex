"""Analytics API router."""
from __future__ import annotations

import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import HTMLResponse
from pydantic import BaseModel

logger = logging.getLogger(__name__)

from ..config import Settings, get_settings
from .store import AnalyticsStore

router = APIRouter(prefix="/api/analytics", tags=["analytics"])
page_router = APIRouter(tags=["analytics-page"])

# Path to analytics HTML page
ANALYTICS_HTML_PATH = Path(__file__).resolve().parents[3] / "webapp" / "analytics.html"


@page_router.get("/analytics", response_class=HTMLResponse)
def serve_analytics_page() -> HTMLResponse:
    """Serve the analytics dashboard HTML page."""
    if not ANALYTICS_HTML_PATH.exists():
        raise HTTPException(status_code=404, detail="Analytics page not found")
    return HTMLResponse(content=ANALYTICS_HTML_PATH.read_text(encoding="utf-8"))

# Global store instance (initialized in main.py)
_store: Optional[AnalyticsStore] = None


def set_store(store: AnalyticsStore) -> None:
    """Set the global analytics store instance."""
    global _store
    _store = store


def get_store() -> AnalyticsStore:
    """Get the analytics store instance."""
    if _store is None:
        raise HTTPException(status_code=500, detail="Analytics store not initialized")
    return _store


def verify_password(
    password: str = Query(..., description="Analytics password"),
    settings: Settings = Depends(get_settings),
) -> bool:
    """Verify the analytics password."""
    if not settings.analytics_password:
        raise HTTPException(status_code=403, detail="Analytics password not configured")
    if password != settings.analytics_password:
        raise HTTPException(status_code=401, detail="Invalid password")
    return True


class DailyStats(BaseModel):
    date: str
    client_type: str
    count: int


class EndpointStats(BaseModel):
    endpoint: str
    client_type: str
    count: int
    avg_duration_ms: Optional[float]


class DocTypeStats(BaseModel):
    doc_type: Optional[str]
    client_type: str
    count: int


class SummaryResponse(BaseModel):
    client_counts: dict
    total_requests: int
    unique_endpoints: int
    avg_duration_ms: Optional[float]
    # Session & user estimation
    unique_users: dict = {}
    sessions: dict = {}
    completed_analyses: int = 0
    completed_by_client: dict = {}
    # Cost & time savings estimates
    pdf_analysis_count: int = 0
    total_pages_processed: int = 0
    jon_batch_count: int = 0
    jon_single_count: int = 0
    estimated_cost_yen: int = 0
    saved_minutes: int = 0
    # Per-client analysis counts
    excel_analysis_count: int = 0
    web_analysis_count: int = 0


class LogEntry(BaseModel):
    id: int
    timestamp: str
    endpoint: str
    method: str
    client_type: str
    doc_type: Optional[str]
    status_code: Optional[int]
    duration_ms: Optional[int]
    user_agent: Optional[str]
    ip_address: Optional[str]


@router.get("/summary")
def get_summary(
    start_date: Optional[str] = Query(None, description="Start date (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="End date (YYYY-MM-DD)"),
    _: bool = Depends(verify_password),
    store: AnalyticsStore = Depends(get_store),
) -> SummaryResponse:
    """Get summary statistics."""
    if not start_date:
        start_date = "2025-02-01"
    if not end_date:
        end_date = datetime.utcnow().strftime("%Y-%m-%d")

    return SummaryResponse(**store.get_summary(start_date, end_date))


@router.get("/daily")
def get_daily_stats(
    start_date: Optional[str] = Query(None, description="Start date (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="End date (YYYY-MM-DD)"),
    _: bool = Depends(verify_password),
    store: AnalyticsStore = Depends(get_store),
) -> list[DailyStats]:
    """Get daily access statistics."""
    if not start_date:
        start_date = "2025-02-01"
    if not end_date:
        end_date = datetime.utcnow().strftime("%Y-%m-%d")

    rows = store.get_daily_stats(start_date, end_date)
    return [DailyStats(**row) for row in rows]


@router.get("/endpoints")
def get_endpoint_stats(
    start_date: Optional[str] = Query(None, description="Start date (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="End date (YYYY-MM-DD)"),
    _: bool = Depends(verify_password),
    store: AnalyticsStore = Depends(get_store),
) -> list[EndpointStats]:
    """Get endpoint usage statistics."""
    if not start_date:
        start_date = "2025-02-01"
    if not end_date:
        end_date = datetime.utcnow().strftime("%Y-%m-%d")

    rows = store.get_endpoint_stats(start_date, end_date)
    return [EndpointStats(**row) for row in rows]


@router.get("/doc-types")
def get_doc_type_stats(
    start_date: Optional[str] = Query(None, description="Start date (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="End date (YYYY-MM-DD)"),
    _: bool = Depends(verify_password),
    store: AnalyticsStore = Depends(get_store),
) -> list[DocTypeStats]:
    """Get document type usage statistics."""
    if not start_date:
        start_date = "2025-02-01"
    if not end_date:
        end_date = datetime.utcnow().strftime("%Y-%m-%d")

    rows = store.get_doc_type_stats(start_date, end_date)
    return [DocTypeStats(**row) for row in rows]


class ImportResult(BaseModel):
    inserted: int
    skipped: int
    total: int


@router.post("/import")
def import_logs(
    payload: List[Dict[str, Any]],
    _: bool = Depends(verify_password),
    store: AnalyticsStore = Depends(get_store),
) -> ImportResult:
    """Bulk import access logs from external source (e.g., Railway export)."""
    # 1リクエスト最大10,000件まで（メモリ保護）
    MAX_BATCH_SIZE = 10_000
    if len(payload) > MAX_BATCH_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"Batch too large: {len(payload)} rows exceeds limit of {MAX_BATCH_SIZE}. "
                   "Split into smaller batches.",
        )
    inserted = 0
    skipped = 0
    for row in payload:
        try:
            if store.exists_log(
                row.get("timestamp"),
                row.get("endpoint"),
                row.get("ip_address"),
            ):
                skipped += 1
                continue
            store.insert_log_raw(
                timestamp=row.get("timestamp") or "",
                endpoint=row.get("endpoint") or "",
                method=row.get("method", "GET") or "GET",
                client_type=row.get("client_type", "web") or "web",
                doc_type=row.get("doc_type"),
                status_code=row.get("status_code"),
                duration_ms=row.get("duration_ms"),
                user_agent=row.get("user_agent"),
                ip_address=row.get("ip_address"),
                extra=row.get("extra"),
            )
            inserted += 1
        except Exception as e:
            logger.warning(f"Skipped log row: {e}")
            skipped += 1
    return ImportResult(inserted=inserted, skipped=skipped, total=len(payload))


@router.get("/logs")
def get_recent_logs(
    limit: int = Query(100, ge=1, le=100000, description="Number of logs to return"),
    _: bool = Depends(verify_password),
    store: AnalyticsStore = Depends(get_store),
) -> list[LogEntry]:
    """Get recent access logs."""
    rows = store.get_recent_logs(limit)
    return [LogEntry(**row) for row in rows]
