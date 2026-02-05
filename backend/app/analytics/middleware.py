"""Middleware for automatic access logging."""
from __future__ import annotations

import time
from typing import Callable, Optional

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import Response

from .store import AnalyticsStore


class AccessLogMiddleware(BaseHTTPMiddleware):
    """Middleware to log all API requests."""

    # Endpoints to skip logging (health checks, static files, etc.)
    SKIP_ENDPOINTS = {
        "/api/ping",
        "/favicon.ico",
        "/analytics",
    }

    # Endpoints that are typically static file requests or analytics itself
    SKIP_PREFIXES = (
        "/ledger/",
        "/assets/",
        "/api/analytics/",  # Don't log analytics API calls
    )

    def __init__(self, app, store: AnalyticsStore) -> None:
        super().__init__(app)
        self.store = store

    async def dispatch(self, request: Request, call_next: Callable) -> Response:
        # Skip logging for certain endpoints
        path = request.url.path
        if path in self.SKIP_ENDPOINTS or path.startswith(self.SKIP_PREFIXES):
            return await call_next(request)

        # Only log API endpoints
        if not path.startswith("/api/"):
            return await call_next(request)

        start_time = time.time()

        # Determine client type from headers
        client_type = self._detect_client_type(request)

        # Extract doc_type if available (from form data or query params)
        doc_type = request.query_params.get("doc_type")

        # Process the request
        response = await call_next(request)

        # Calculate duration
        duration_ms = int((time.time() - start_time) * 1000)

        # Get user agent and IP
        user_agent = request.headers.get("user-agent", "")[:500]  # Truncate long UAs
        ip_address = self._get_client_ip(request)

        # Log the access
        try:
            self.store.log_access(
                endpoint=path,
                method=request.method,
                client_type=client_type,
                doc_type=doc_type,
                status_code=response.status_code,
                duration_ms=duration_ms,
                user_agent=user_agent,
                ip_address=ip_address,
            )
        except Exception:
            # Don't let logging failures break the API
            pass

        return response

    def _detect_client_type(self, request: Request) -> str:
        """Detect client type from request headers."""
        # Check custom header first (for VBA/Excel clients)
        custom_client = request.headers.get("x-client-type", "").lower()
        if custom_client in ("vba", "excel"):
            return "excel"

        # Check user agent for common patterns
        user_agent = request.headers.get("user-agent", "").lower()

        # VBA/Excel typically use MSXML or WinHTTP
        if any(pattern in user_agent for pattern in ("msxml", "winhttp", "excel", "vba")):
            return "excel"

        # Python requests (could be scripts or automated)
        if "python" in user_agent:
            return "script"

        # Default to web
        return "web"

    def _get_client_ip(self, request: Request) -> Optional[str]:
        """Get client IP address, handling proxies."""
        # Check X-Forwarded-For header (for reverse proxies)
        forwarded_for = request.headers.get("x-forwarded-for")
        if forwarded_for:
            # Take the first IP in the chain
            return forwarded_for.split(",")[0].strip()

        # Check X-Real-IP header
        real_ip = request.headers.get("x-real-ip")
        if real_ip:
            return real_ip

        # Fall back to direct client IP
        if request.client:
            return request.client.host

        return None
