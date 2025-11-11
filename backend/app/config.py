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
    azure_form_recognizer_endpoint: Optional[str]
    azure_form_recognizer_key: Optional[str]


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

    gemini_model = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
    gemini_max_mb = int(os.getenv("GEMINI_DOCUMENT_MAX_MB", "20"))
    gemini_max_bytes = gemini_max_mb * 1024 * 1024
    gemini_chunk_page_limit = int(os.getenv("GEMINI_CHUNK_PAGE_LIMIT", "2"))

    return Settings(
        gemini_api_key=gemini_api_key,
        gemini_api_keys=gemini_api_keys,
        gemini_model=gemini_model,
        gemini_max_document_bytes=gemini_max_bytes,
        gemini_chunk_page_limit=gemini_chunk_page_limit,
        azure_form_recognizer_endpoint=os.getenv("AZURE_FORM_RECOGNIZER_ENDPOINT"),
        azure_form_recognizer_key=os.getenv("AZURE_FORM_RECOGNIZER_KEY"),
    )
