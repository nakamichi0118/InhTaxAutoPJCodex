"""Gemini API helper for PDF analysis."""
from __future__ import annotations

import base64
import json
import logging
import re
import uuid
from typing import Any, Dict, List

import requests

logger = logging.getLogger(__name__)

JSON_PATTERN = re.compile(r"\{.*\}", re.DOTALL)
INLINE_LIMIT_BYTES = 4 * 1024 * 1024
UPLOAD_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/files:upload"
FILE_BASE_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/"


class GeminiError(RuntimeError):
    """Raised when the Gemini service returns an error payload."""


class GeminiClient:
    def __init__(self, *, api_key: str, model: str) -> None:
        self.api_key = api_key
        self.model = model
        self.endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

    def extract_lines_from_pdf(self, pdf_bytes: bytes) -> List[str]:
        if len(pdf_bytes) <= INLINE_LIMIT_BYTES:
            payload = self._build_inline_payload(pdf_bytes)
            data = self._invoke_generate(payload)
            return self._parse_response(data)

        file_name = self._upload_file(pdf_bytes, mime_type="application/pdf")
        try:
            payload = self._build_file_payload(file_name)
            data = self._invoke_generate(payload)
            return self._parse_response(data)
        finally:
            self._delete_file(file_name)

    def _invoke_generate(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        response = requests.post(
            f"{self.endpoint}?key={self.api_key}",
            json=payload,
            timeout=120,
        )
        if response.status_code >= 400:
            raise GeminiError(self._extract_error(response))
        return response.json()

    def _upload_file(self, pdf_bytes: bytes, *, mime_type: str) -> str:
        metadata = {
            "mimeType": mime_type,
            "displayName": f"document-{uuid.uuid4().hex}.pdf",
        }
        files = {
            "metadata": ("metadata.json", json.dumps(metadata), "application/json; charset=UTF-8"),
            "file": (metadata["displayName"], pdf_bytes, mime_type),
        }
        response = requests.post(
            f"{UPLOAD_ENDPOINT}?key={self.api_key}",
            files=files,
            timeout=120,
        )
        if response.status_code >= 400:
            raise GeminiError(self._extract_error(response))
        payload = response.json()
        file_name = payload.get("name")
        if not file_name:
            raise GeminiError("Gemini upload did not return a file name")
        return file_name

    def _delete_file(self, file_name: str) -> None:
        try:
            requests.delete(
                f"{FILE_BASE_ENDPOINT}{file_name}?key={self.api_key}",
                timeout=30,
            )
        except requests.RequestException as exc:
            logger.debug("Failed to delete Gemini file %s: %s", file_name, exc)

    def _build_inline_payload(self, pdf_bytes: bytes) -> Dict[str, Any]:
        encoded = base64.b64encode(pdf_bytes).decode("ascii")
        return self._base_prompt(parts=[
            {"text": self._prompt_text()},
            {
                "inline_data": {
                    "mime_type": "application/pdf",
                    "data": encoded,
                }
            },
        ])

    def _build_file_payload(self, file_name: str) -> Dict[str, Any]:
        return self._base_prompt(parts=[
            {"text": self._prompt_text()},
            {
                "file_data": {
                    "mime_type": "application/pdf",
                    "file_uri": file_name,
                }
            },
        ])

    @staticmethod
    def _base_prompt(parts: List[Dict[str, Any]]) -> Dict[str, Any]:
        return {
            "contents": [
                {
                    "role": "user",
                    "parts": parts,
                }
            ]
        }

    @staticmethod
    def _prompt_text() -> str:
        return (
            "You are assisting with converting Japanese bank and financial documents into plain text. "
            "Read the attached PDF and return a JSON object with a single key `lines` containing an array "
            "of strings. Preserve the reading order and include blank lines only when they are meaningful. "
            "Do not add explanations or markdown. Return raw JSON only."
        )

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
        if match:
            try:
                parsed = json.loads(match.group(0))
            except json.JSONDecodeError as exc:
                raise GeminiError(f"Failed to decode JSON: {exc}") from exc
            lines = parsed.get("lines")
            if isinstance(lines, list):
                return [str(line) for line in lines]
            raise GeminiError("Gemini JSON payload missing `lines` array")
        logger.warning("Gemini response lacked JSON payload; falling back to raw text split")
        return [line for line in text.splitlines() if line]
