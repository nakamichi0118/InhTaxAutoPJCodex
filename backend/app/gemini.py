"""Gemini API helper for PDF analysis."""
from __future__ import annotations

import base64
import json
import logging
import re
import time
import uuid
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Sequence

import requests
from requests import RequestException

from .prompts.nayose_prompt import NAYOSE_PROMPT

logger = logging.getLogger(__name__)

JSON_PATTERN = re.compile(r"\{.*\}", re.DOTALL)
INLINE_LIMIT_BYTES = 4 * 1024 * 1024
UPLOAD_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/files:upload"
FILE_BASE_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/"


@dataclass
class GeminiExtraction:
    lines: List[str]
    transactions: List[Dict[str, Any]]

    def extend(self, other: "GeminiExtraction") -> None:
        self.lines.extend(other.lines)
        self.transactions.extend(other.transactions)


class GeminiError(RuntimeError):
    """Raised when the Gemini service returns an error payload."""

    def __init__(self, message: str, *, status_code: int | None = None, can_retry_key: bool = False) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.can_retry_key = can_retry_key


class GeminiClient:
    def __init__(self, *, api_keys: Sequence[str] | str, model: str) -> None:
        if isinstance(api_keys, str):
            keys = [api_keys]
        else:
            keys = list(api_keys)
        sanitized = [key.strip() for key in keys if key and key.strip()]
        if not sanitized:
            raise ValueError("At least one Gemini API key must be provided")

        self.api_keys = sanitized
        self.model = model
        self.endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

    def extract_lines_from_pdf(self, pdf_bytes: bytes) -> GeminiExtraction:
        last_error: GeminiError | None = None
        for index, api_key in enumerate(self.api_keys):
            try:
                return self._extract_with_key(pdf_bytes, api_key)
            except GeminiError as exc:
                last_error = exc
                if exc.can_retry_key and index < len(self.api_keys) - 1:
                    logger.warning("Gemini API key rejected (reason: %s); trying next key", exc)
                    continue
                raise
        if last_error:
            raise last_error
        raise GeminiError("Gemini API key configuration is empty")

    def _extract_with_key(self, pdf_bytes: bytes, api_key: str) -> GeminiExtraction:
        if len(pdf_bytes) <= INLINE_LIMIT_BYTES:
            payload = self._build_inline_payload(pdf_bytes)
            data = self._invoke_generate(payload, api_key)
            return self._parse_response(data)

        file_name = self._upload_file(pdf_bytes, api_key=api_key, mime_type="application/pdf")
        try:
            payload = self._build_file_payload(file_name)
            data = self._invoke_generate(payload, api_key)
            return self._parse_response(data)
        finally:
            self._delete_file(file_name, api_key=api_key)

    def _invoke_generate(self, payload: Dict[str, Any], api_key: str) -> Dict[str, Any]:
        last_error: GeminiError | None = None
        for attempt in range(3):
            try:
                response = requests.post(
                    f"{self.endpoint}?key={api_key}",
                    json=payload,
                    timeout=300,
                )
            except RequestException as exc:
                last_error = GeminiError(f"Gemini request failed: {exc}", can_retry_key=False)
                sleep_for = 1.5 * (attempt + 1)
                time.sleep(sleep_for)
                continue

            if response.status_code < 400:
                return response.json()

            error_text, error_json = self._extract_error(response)
            can_retry_key = self._should_rotate_key(response.status_code, error_json or error_text)
            if response.status_code in {429, 500, 502, 503} and attempt < 2 and not can_retry_key:
                time.sleep(1.5 * (attempt + 1))
                continue

            raise GeminiError(error_text, status_code=response.status_code, can_retry_key=can_retry_key)

        if last_error:
            raise last_error
        raise GeminiError("Gemini request failed without response")

    def _upload_file(self, pdf_bytes: bytes, *, api_key: str, mime_type: str) -> str:
        metadata = {
            "mimeType": mime_type,
            "displayName": f"document-{uuid.uuid4().hex}.pdf",
        }
        files = {
            "metadata": ("metadata.json", json.dumps(metadata), "application/json; charset=UTF-8"),
            "file": (metadata["displayName"], pdf_bytes, mime_type),
        }
        response = requests.post(
            f"{UPLOAD_ENDPOINT}?key={api_key}",
            files=files,
            timeout=300,
        )
        if response.status_code >= 400:
            error_text, error_json = self._extract_error(response)
            raise GeminiError(
                error_text,
                status_code=response.status_code,
                can_retry_key=self._should_rotate_key(response.status_code, error_json or error_text),
            )
        payload = response.json()
        file_name = payload.get("name")
        if not file_name:
            raise GeminiError("Gemini upload did not return a file name")
        return file_name

    def _delete_file(self, file_name: str, *, api_key: str) -> None:
        try:
            requests.delete(
                f"{FILE_BASE_ENDPOINT}{file_name}?key={api_key}",
                timeout=30,
            )
        except requests.RequestException as exc:
            logger.debug("Failed to delete Gemini file %s: %s", file_name, exc)

    def _build_inline_payload(self, pdf_bytes: bytes) -> Dict[str, Any]:
        encoded = base64.b64encode(pdf_bytes).decode("ascii")
        return self._base_prompt(
            parts=[
                {"text": self._prompt_text()},
                {
                    "inline_data": {
                        "mime_type": "application/pdf",
                        "data": encoded,
                    }
                },
            ]
        )

    def _build_file_payload(self, file_name: str) -> Dict[str, Any]:
        return self._base_prompt(
            parts=[
                {"text": self._prompt_text()},
                {
                    "file_data": {
                        "mime_type": "application/pdf",
                        "file_uri": file_name,
                    }
                },
            ]
        )

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
            "You are assisting with converting Japanese bank and financial documents into structured text. "
            "Read the attached PDF and return JSON with two keys: `lines` containing an array of strings "
            "in reading order, and `transactions` containing an array of objects with the fields "
            "`date`, `description`, `withdrawal`, `deposit`, `balance`, and `account_type`. "
            "IMPORTANT: Japanese passbooks use 2-digit Japanese era years (和暦). "
            "Return dates EXACTLY as shown in the document without converting the year. "
            "For example: if the document shows '01-12-06', return '01-12-06' (not '2001-12-06' or '2019-12-06'). "
            "If it shows '17-11-24', return '17-11-24' (not '2005-11-24' or '2017-11-24'). "
            "Use null for unknown numeric values, keep amounts as plain numbers (no commas), and do not add "
            "explanations or markdown. Return raw JSON only.\n\n"
            "ACCOUNT TYPE DETECTION: For combined passbooks (総合口座通帳), detect and return `account_type` for each transaction:\n"
            "- 'ordinary_deposit': Transactions on pages labeled 普通預金, 普通口座, or the front side of the passbook\n"
            "- 'time_deposit': Transactions on pages labeled 定期預金, 定期積金, 定期, or the back side showing fixed deposits\n"
            "If account type cannot be determined, return null for `account_type`.\n"
            "IMPORTANT: Always extract the description (摘要) for ALL transactions, regardless of amount."
        )

    @staticmethod
    def _extract_error(response: requests.Response) -> tuple[str, Dict[str, Any] | None]:
        try:
            payload = response.json()
        except ValueError:
            return response.text, None
        return json.dumps(payload), payload

    @staticmethod
    def _should_rotate_key(status_code: int, error_payload: Dict[str, Any] | str) -> bool:
        if status_code != 403:
            return False
        message_candidates: Iterable[str] = ()
        status = ""
        if isinstance(error_payload, dict):
            error = error_payload.get("error") or {}
            message = error.get("message", "")
            status = error.get("status", "")
            message_candidates = (message, status)
        else:
            message_candidates = (str(error_payload),)

        for message in message_candidates:
            lowered = (message or "").lower()
            if "reported as leaked" in lowered:
                return True
            if "api key not valid" in lowered:
                return True
            if "permission_denied" in lowered or "permission denied" in lowered:
                return True
        return False

    @staticmethod
    def _parse_response(data: Dict[str, Any]) -> GeminiExtraction:
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
            if not isinstance(lines, list):
                raise GeminiError("Gemini JSON payload missing `lines` array")
            structured = parsed.get("transactions")
            transactions: List[Dict[str, Any]] = []
            if isinstance(structured, list):
                for item in structured:
                    if isinstance(item, dict):
                        transactions.append(item)
            return GeminiExtraction(lines=[str(line) for line in lines], transactions=transactions)
        logger.warning("Gemini response lacked JSON payload; falling back to raw text split")
        return GeminiExtraction(lines=[line for line in text.splitlines() if line], transactions=[])

    def analyze_text(self, prompt: str) -> str:
        """Send a text prompt to Gemini and return the response text."""
        payload = {
            "contents": [
                {
                    "role": "user",
                    "parts": [{"text": prompt}],
                }
            ]
        }
        last_error: GeminiError | None = None
        for index, api_key in enumerate(self.api_keys):
            try:
                data = self._invoke_generate(payload, api_key)
                candidates = data.get("candidates") or []
                if not candidates:
                    raise GeminiError("No candidates returned from Gemini API")
                parts = candidates[0].get("content", {}).get("parts", [])
                if not parts:
                    raise GeminiError("Candidate contains no parts")
                text = parts[0].get("text", "").strip()
                if not text:
                    raise GeminiError("Gemini response did not contain text")
                return text
            except GeminiError as exc:
                last_error = exc
                if exc.can_retry_key and index < len(self.api_keys) - 1:
                    logger.warning("Gemini API key rejected (reason: %s); trying next key", exc)
                    continue
                raise
        if last_error:
            raise last_error
        raise GeminiError("Gemini API key configuration is empty")

    def extract_nayose_from_pdf(self, pdf_bytes: bytes) -> str:
        """Extract property information from 名寄帳/固定資産評価証明書 PDF.

        Returns the raw JSON response text from Gemini.
        """
        last_error: GeminiError | None = None
        for index, api_key in enumerate(self.api_keys):
            try:
                return self._extract_nayose_with_key(pdf_bytes, api_key)
            except GeminiError as exc:
                last_error = exc
                if exc.can_retry_key and index < len(self.api_keys) - 1:
                    logger.warning("Gemini API key rejected (reason: %s); trying next key", exc)
                    continue
                raise
        if last_error:
            raise last_error
        raise GeminiError("Gemini API key configuration is empty")

    def _extract_nayose_with_key(self, pdf_bytes: bytes, api_key: str) -> str:
        """Extract nayose data using a specific API key."""
        if len(pdf_bytes) <= INLINE_LIMIT_BYTES:
            payload = self._build_nayose_inline_payload(pdf_bytes)
        else:
            file_name = self._upload_file(pdf_bytes, api_key=api_key, mime_type="application/pdf")
            try:
                payload = self._build_nayose_file_payload(file_name)
            finally:
                self._delete_file(file_name, api_key=api_key)

        data = self._invoke_generate(payload, api_key)
        candidates = data.get("candidates") or []
        if not candidates:
            raise GeminiError("No candidates returned from Gemini API")
        parts = candidates[0].get("content", {}).get("parts", [])
        if not parts:
            raise GeminiError("Candidate contains no parts")
        text = parts[0].get("text", "").strip()
        if not text:
            raise GeminiError("Gemini response did not contain text")
        return text

    def _build_nayose_inline_payload(self, pdf_bytes: bytes) -> Dict[str, Any]:
        """Build inline payload for nayose extraction."""
        encoded = base64.b64encode(pdf_bytes).decode("ascii")
        return self._base_prompt(
            parts=[
                {"text": NAYOSE_PROMPT},
                {
                    "inline_data": {
                        "mime_type": "application/pdf",
                        "data": encoded,
                    }
                },
            ]
        )

    def _build_nayose_file_payload(self, file_name: str) -> Dict[str, Any]:
        """Build file-based payload for nayose extraction."""
        return self._base_prompt(
            parts=[
                {"text": NAYOSE_PROMPT},
                {
                    "file_data": {
                        "mime_type": "application/pdf",
                        "file_uri": file_name,
                    }
                },
            ]
        )
