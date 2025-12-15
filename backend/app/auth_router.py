from __future__ import annotations

import logging
import secrets
import smtplib
import ssl
from email.message import EmailMessage
from typing import Optional

from fastapi import APIRouter, HTTPException

from .config import get_settings
from .ledger_store import LedgerStore
from .models import AuthRequestCode, AuthVerifyRequest, AuthVerifyResponse

router = APIRouter(prefix="/api/auth", tags=["auth"])
logger = logging.getLogger("uvicorn.error")
ALLOWED_EMAIL_DOMAIN = "taxlawyer328.com"
CODE_TTL_SECONDS = 15 * 60
CODE_LENGTH = 6


def _send_email_code(email: str, code: str, settings) -> None:
    if not settings.smtp_host or not settings.smtp_from:
        raise HTTPException(status_code=500, detail="SMTP settings are not configured")

    message = EmailMessage()
    message["Subject"] = "SOROBOCR 認証コード"
    message["From"] = settings.smtp_from
    message["To"] = email
    message.set_content(
        f"""以下の認証コードを入力してください（15分以内に有効です）。

認証コード: {code}
発行元: SOROBOCR 入出金検討表
"""
    )

    port = settings.smtp_port or 587
    if settings.smtp_use_tls:
        context = ssl.create_default_context()
        with smtplib.SMTP(settings.smtp_host, port) as server:
            server.starttls(context=context)
            if settings.smtp_user:
                server.login(settings.smtp_user, settings.smtp_password or "")
            server.send_message(message)
    else:
        with smtplib.SMTP(settings.smtp_host, port) as server:
            if settings.smtp_user:
                server.login(settings.smtp_user, settings.smtp_password or "")
            server.send_message(message)


def _generate_code(length: int = 6) -> str:
    upper = 10**length - 1
    return f"{secrets.randbelow(upper + 1):0{length}d}"


@router.post("/request_code")
def request_code(payload: AuthRequestCode) -> dict:
    email = (payload.email or "").strip()
    if not email:
        raise HTTPException(status_code=400, detail="Email is required")
    email_lower = email.lower()
    if not email_lower.endswith(f"@{ALLOWED_EMAIL_DOMAIN}"):
        raise HTTPException(status_code=400, detail="This email domain is not allowed")
    settings = get_settings()
    store = LedgerStore(settings.ledger_db_path)
    code = _generate_code(CODE_LENGTH)
    record = store.issue_auth_code(email_lower, code, ttl_seconds=CODE_TTL_SECONDS)
    try:
        _send_email_code(email, code, settings)
    except Exception as exc:  # pragma: no cover - SMTP errors
        logger.exception("Failed to send auth code email")
        raise HTTPException(status_code=500, detail="Failed to send verification email") from exc
    return {"status": "ok", "email": email_lower, "expires_at": record["expires_at"]}


@router.post("/verify", response_model=AuthVerifyResponse)
def verify_code(payload: AuthVerifyRequest) -> AuthVerifyResponse:
    email = (payload.email or "").strip()
    code = (payload.code or "").strip()
    if not email or not code:
        raise HTTPException(status_code=400, detail="Email and code are required")
    email_lower = email.lower()
    if not email_lower.endswith(f"@{ALLOWED_EMAIL_DOMAIN}"):
        raise HTTPException(status_code=400, detail="This email domain is not allowed")
    if not code.isdigit() or len(code) != CODE_LENGTH:
        raise HTTPException(status_code=400, detail="Code must be 6 digits")
    settings = get_settings()
    store = LedgerStore(settings.ledger_db_path)
    verified = store.verify_auth_code(email_lower, code)
    if not verified:
        raise HTTPException(status_code=400, detail="Invalid or expired code")
    return AuthVerifyResponse(status="ok", verified=True, email=email_lower)
