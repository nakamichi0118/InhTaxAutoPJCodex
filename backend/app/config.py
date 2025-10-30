"""Application settings."""
from __future__ import annotations

import os
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from dotenv import load_dotenv

_ENV_PATH = Path(__file__).resolve().parents[2] / ".env"
load_dotenv(_ENV_PATH)


@dataclass(frozen=True)
class Settings:
    gemini_api_key: str
    gemini_model: str
    gemini_max_document_bytes: int
    gemini_chunk_page_limit: int


@lru_cache()
def get_settings() -> Settings:
    gemini_api_key = os.getenv("GEMINI_API_KEY")
    if not gemini_api_key:
        raise RuntimeError("Missing required environment variable: GEMINI_API_KEY")

    gemini_model = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
    gemini_max_mb = int(os.getenv("GEMINI_DOCUMENT_MAX_MB", "20"))
    gemini_max_bytes = gemini_max_mb * 1024 * 1024
    gemini_chunk_page_limit = int(os.getenv("GEMINI_CHUNK_PAGE_LIMIT", "5"))

    return Settings(
        gemini_api_key=gemini_api_key,
        gemini_model=gemini_model,
        gemini_max_document_bytes=gemini_max_bytes,
        gemini_chunk_page_limit=gemini_chunk_page_limit,
    )
