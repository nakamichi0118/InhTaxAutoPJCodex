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
    azure_endpoint: str
    azure_key: str
    gemini_api_key: Optional[str]
    azure_max_document_bytes: int
    azure_chunk_page_limit: int
    gemini_model: str
    gemini_max_document_bytes: int
    gemini_chunk_page_limit: int


@lru_cache()
def get_settings() -> Settings:
    endpoint = os.getenv("AZURE_FORM_RECOGNIZER_ENDPOINT")
    key = os.getenv("AZURE_FORM_RECOGNIZER_KEY")
    gemini = os.getenv("GEMINI_API_KEY")

    max_mb = int(os.getenv("AZURE_DOCUMENT_MAX_MB", "4"))
    max_bytes = max_mb * 1024 * 1024
    chunk_page_limit = int(os.getenv("AZURE_CHUNK_PAGE_LIMIT", "20"))

    gemini_model = os.getenv("GEMINI_MODEL", "gemini-1.5-flash-latest")
    gemini_max_mb = int(os.getenv("GEMINI_DOCUMENT_MAX_MB", "20"))
    gemini_max_bytes = gemini_max_mb * 1024 * 1024
    gemini_chunk_page_limit = int(os.getenv("GEMINI_CHUNK_PAGE_LIMIT", str(chunk_page_limit)))

    if not endpoint or not key:
        missing = [name for name, value in [
            ("AZURE_FORM_RECOGNIZER_ENDPOINT", endpoint),
            ("AZURE_FORM_RECOGNIZER_KEY", key),
        ] if not value]
        raise RuntimeError(f"Missing required environment variables: {', '.join(missing)}")
    return Settings(
        azure_endpoint=endpoint.rstrip("/"),
        azure_key=key,
        gemini_api_key=gemini,
        azure_max_document_bytes=max_bytes,
        azure_chunk_page_limit=chunk_page_limit,
        gemini_model=gemini_model,
        gemini_max_document_bytes=gemini_max_bytes,
        gemini_chunk_page_limit=gemini_chunk_page_limit,
    )
