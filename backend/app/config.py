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


@lru_cache()
def get_settings() -> Settings:
    endpoint = os.getenv("AZURE_FORM_RECOGNIZER_ENDPOINT")
    key = os.getenv("AZURE_FORM_RECOGNIZER_KEY")
    gemini = os.getenv("GEMINI_API_KEY")
    if not endpoint or not key:
        missing = [name for name, value in [
            ("AZURE_FORM_RECOGNIZER_ENDPOINT", endpoint),
            ("AZURE_FORM_RECOGNIZER_KEY", key),
        ] if not value]
        raise RuntimeError(f"Missing required environment variables: {', '.join(missing)}")
    return Settings(azure_endpoint=endpoint.rstrip("/"), azure_key=key, gemini_api_key=gemini)
