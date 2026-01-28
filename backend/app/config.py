"""Application settings."""
from __future__ import annotations

import os
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

_ENV_PATH = Path(__file__).resolve().parents[2] / ".env"
load_dotenv(_ENV_PATH)


@dataclass(frozen=True)
class Settings:
    gemini_api_key: str
    gemini_api_keys: tuple[str, ...]
    gemini_model: str
    gemini_max_document_bytes: int
    gemini_chunk_page_limit: int
    azure_chunk_max_bytes: int
    azure_form_recognizer_endpoint: Optional[str]
    azure_form_recognizer_key: Optional[str]
    ledger_db_path: Path
    cors_allow_origins: tuple[str, ...]
    # JON API settings
    jon_client_id: Optional[str]
    jon_client_secret: Optional[str]
    jon_api_base_url: str
    touki_login_id: Optional[str]
    touki_password: Optional[str]
    # 不動産情報ライブラリAPI settings
    reinfolib_api_key: Optional[str]


@lru_cache()
def get_settings() -> Settings:
    gemini_api_keys_env = os.getenv("GEMINI_API_KEYS")
    if gemini_api_keys_env:
        raw_keys = [key.strip() for key in gemini_api_keys_env.split(",")]
        gemini_api_keys = tuple(key for key in raw_keys if key)
    else:
        gemini_api_keys = tuple()

    gemini_api_key = os.getenv("GEMINI_API_KEY")
    if gemini_api_key:
        gemini_api_key = gemini_api_key.strip()

    if not gemini_api_keys:
        if not gemini_api_key:
            raise RuntimeError("Missing required environment variable: GEMINI_API_KEY")
        gemini_api_keys = (gemini_api_key,)
    elif gemini_api_key and gemini_api_key not in gemini_api_keys:
        # Ensure GEMINI_API_KEY is always considered primary if provided separately.
        gemini_api_keys = (gemini_api_key,) + tuple(key for key in gemini_api_keys if key != gemini_api_key)
    else:
        gemini_api_key = gemini_api_keys[0]

    gemini_model = os.getenv("GEMINI_MODEL", "gemini-2.5-pro")
    gemini_max_mb = int(os.getenv("GEMINI_DOCUMENT_MAX_MB", "20"))
    gemini_max_bytes = gemini_max_mb * 1024 * 1024
    gemini_chunk_page_limit = int(os.getenv("GEMINI_CHUNK_PAGE_LIMIT", "2"))
    azure_chunk_max_mb = int(os.getenv("AZURE_CHUNK_MAX_MB", "6"))
    azure_chunk_max_bytes = azure_chunk_max_mb * 1024 * 1024

    ledger_db_path_env = os.getenv("LEDGER_DB_PATH")
    if ledger_db_path_env:
        ledger_db_path = Path(ledger_db_path_env).expanduser().resolve()
    else:
        ledger_db_path = Path(__file__).resolve().parents[2] / "data" / "ledger.db"

    cors_env = os.getenv("CORS_ALLOW_ORIGINS", "")
    if cors_env.strip():
        cors_allow_origins = tuple([origin.strip() for origin in cors_env.split(",") if origin.strip()])
    else:
        cors_allow_origins = (
            "https://inhtaxautopjcodex.pages.dev",
            "https://www.sorobocr.taxlawyer328.jp",
            "http://localhost:5173",
            "http://127.0.0.1:5173",
        )

    # JON API settings
    jon_client_id = os.getenv("JON_CLIENT_ID")
    jon_client_secret = os.getenv("JON_CLIENT_SECRET")
    jon_api_base_url = os.getenv("JON_API_BASE_URL", "https://jon-api.com/api/v1")
    touki_login_id = os.getenv("TOUKI_LOGIN_ID")
    touki_password = os.getenv("TOUKI_PASSWORD")

    return Settings(
        gemini_api_key=gemini_api_key,
        gemini_api_keys=gemini_api_keys,
        gemini_model=gemini_model,
        gemini_max_document_bytes=gemini_max_bytes,
        gemini_chunk_page_limit=gemini_chunk_page_limit,
        azure_chunk_max_bytes=azure_chunk_max_bytes,
        azure_form_recognizer_endpoint=os.getenv("AZURE_FORM_RECOGNIZER_ENDPOINT"),
        azure_form_recognizer_key=os.getenv("AZURE_FORM_RECOGNIZER_KEY"),
        ledger_db_path=ledger_db_path,
        cors_allow_origins=cors_allow_origins,
        jon_client_id=jon_client_id,
        jon_client_secret=jon_client_secret,
        jon_api_base_url=jon_api_base_url,
        touki_login_id=touki_login_id,
        touki_password=touki_password,
        reinfolib_api_key=os.getenv("REINFOLIB_API_KEY"),
    )
