"""Gemini API helper for PDF analysis."""
from __future__ import annotations

import base64
import json
import logging
import re
from typing import Any, Dict, List

import requests

logger = logging.getLogger(__name__)

JSON_PATTERN = re.compile(r"\{.*\}", re.DOTALL)


class GeminiError(RuntimeError):
    """Raised when the Gemini service returns an error payload."""


class GeminiClient:
    def __init__(self, *, api_key: str, model: str) -> None:
        self.api_key = api_key
        self.model = model
        self.endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

    def extract_lines_from_pdf(self, pdf_bytes: bytes) -> List[str]:
        payload = self._build_request_payload(pdf_bytes)
        response = requests.post(
            f"{self.endpoint}?key={self.api_key}",
            json=payload,
            timeout=120,
        )
        if response.status_code >= 400:
            raise GeminiError(self._extract_error(response))
        data = response.json()
        return self._parse_response(data)

    @staticmethod
    def _build_request_payload(pdf_bytes: bytes) -> Dict[str, Any]:
        encoded = base64.b64encode(pdf_bytes).decode("ascii")
        prompt = (
            "You are assisting with converting Japanese bank documents into plain text. "
            "Read the attached PDF and return a JSON object with a single key `lines` containing an array "
            "of strings. Preserve the reading order and include blank lines only when they are meaningful. "
            "Do not add any explanations or markdown—return raw JSON only."
        )
        return {
            "contents": [
                {
                    "role": "user",
                    "parts": [
                        {"text": prompt},
                        {
                            "inline_data": {
                                "mime_type": "application/pdf",
                                "data": encoded,
                            }
                        },
                    ],
                }
            ]
        }

    @staticmethod
    def _extract_error(response: requests.Response) -> str:
        try:
            payload = response.json()
        except ValueError:
            return response.text
        return json.dumps(payload)

    @staticmethod
    def _parse_response(data: Dict[str, Any]) -> List[str]:
        candidates = data.get("candidates") or []
        if not candidates:
            raise GeminiError("No candidates returned from Gemini API")
        parts = candidates[0].get("content", {}).get("parts", [])
        if not parts:
            raise GeminiError("Candidate contains no parts")
        text = parts[0].get("text", "").strip()
        if not text:
            raise GeminiError("Gemini response did not contain text")
        match = JSON_PATTERN.search(text)
        if not match:
            raise GeminiError("Unable to locate JSON payload in Gemini response")
        try:
            parsed = json.loads(match.group(0))
        except json.JSONDecodeError as exc:
            raise GeminiError(f"Failed to decode JSON: {exc}") from exc
        lines = parsed.get("lines")
        if not isinstance(lines, list):
            raise GeminiError("Gemini JSON payload missing `lines` array")
        return [str(line) for line in lines]
