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
from .prompts.generic_ocr_prompt import GENERIC_OCR_PROMPT
from .pdf_utils import enhance_scanned_pdf

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

    def extract_lines_from_pdf(self, pdf_bytes: bytes, *, page_context: str = "") -> GeminiExtraction:
        last_error: GeminiError | None = None
        for index, api_key in enumerate(self.api_keys):
            try:
                return self._extract_with_key(pdf_bytes, api_key, page_context=page_context)
            except GeminiError as exc:
                last_error = exc
                if exc.can_retry_key and index < len(self.api_keys) - 1:
                    logger.warning("Gemini API key rejected (reason: %s); trying next key", exc)
                    continue
                raise
        if last_error:
            raise last_error
        raise GeminiError("Gemini API key configuration is empty")

    def _extract_with_key(self, pdf_bytes: bytes, api_key: str, *, page_context: str = "") -> GeminiExtraction:
        if len(pdf_bytes) <= INLINE_LIMIT_BYTES:
            payload = self._build_inline_payload(pdf_bytes, page_context=page_context)
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

    def _build_inline_payload(self, pdf_bytes: bytes, *, page_context: str = "") -> Dict[str, Any]:
        # Try to enhance scanned images for better OCR accuracy
        data_bytes, mime_type = enhance_scanned_pdf(pdf_bytes)
        encoded = base64.b64encode(data_bytes).decode("ascii")
        prompt = self._prompt_text()
        if page_context:
            prompt = f"{page_context}\n\n{prompt}"
        return self._base_prompt(
            parts=[
                {"text": prompt},
                {
                    "inline_data": {
                        "mime_type": mime_type,
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
            ],
            "generationConfig": {
                "temperature": 0,
            },
        }

    @staticmethod
    def _prompt_text() -> str:
        return (
            "あなたは日本の銀行通帳・取引明細を構造化データに変換するAIアシスタントです。\n"
            "添付画像を読み取り、以下のJSON形式で返してください:\n"
            '{"lines": ["行1", "行2", ...], "transactions": [...]}\n\n'
            "transactionsの各要素:\n"
            '{"date": "YY-M-D", "description": "摘要", "withdrawal": 数値orNull, '
            '"deposit": 数値orNull, "balance": 数値orNull, "account_type": "ordinary_deposit"orNull}\n\n'
            "【日付ルール（最重要）】\n"
            "- 日付は印字された数字をそのまま「YY-M-D」形式で返すこと。\n"
            "- 絶対に年号変換をしないこと。22は22、24は24、25は25として返す。\n"
            "  例: 通帳に「22-4-15」と印字 → \"22-4-15\" （「2010-4-15」や「2022-4-15」にしない）\n"
            "  例: 通帳に「25-12-15」と印字 → \"25-12-15\" （「2013-12-15」にしない）\n"
            "- 通帳では同じ日の2行目以降は日付欄が空白になることがある。\n"
            "  空白の場合はnullにせず、直前の取引と同じ日付をコピーしてよい。\n"
            "- ただし、異なる月/日の取引には必ず異なる日付が印字されている。\n"
            "  ページ内で新しい日付が現れたら、必ずその日付を正確に読み取ること。\n"
            "- 各ページの最初の取引には必ず日付が印字されている。見落とさないこと。\n\n"
            "【金額ルール（極めて重要）】\n"
            "- 金額はカンマなしの数値で返すこと\n"
            "- 不明な場合はnull\n"
            "- 出金=withdrawal, 入金=deposit, 残高=balance\n"
            "- 数字を1桁ずつ慎重に読むこと。特に 2/5/6/8 など似た形の数字は要注意。\n"
            "- 隣接する列（残高列など）の数字を絶対に金額列に混ぜないこと。\n"
            "  例: 「200,000」と「9,158,028」が並んでいたら別々の値として読む。\n"
            "  「2,809」を「2,809,421」と読まない（次のセルを巻き込まない）。\n"
            "- 同じ行の他のセル(残高など)に書かれた数字を金額列に転記しないこと。\n\n"
            "【取引明細表の2行構造（流動性預金取引明細表など）】\n"
            "- 一部の取引明細表では、1取引が2行に分かれている：\n"
            "  - 1行目: 取扱日 / 取引区分 / 金額 / 残高\n"
            "  - 2行目: 摘要(取引相手名)詳細\n"
            "- 2行を1取引としてまとめて返すこと。2取引に分けないこと。\n"
            "- 摘要(description)は1行目と2行目の両方の文字を結合する。\n\n"
            "【完全性チェック】\n"
            "- ページ内の全取引を漏れなく抽出すること。1行も飛ばさない。\n"
            "- 残高の連続性を確認: 各取引で『前の残高 ± 金額 = 現在の残高』が成立するはず。\n"
            "  成立しない場合は、金額の桁を読み間違えている可能性が高い。再確認すること。\n\n"
            "【厳守事項】\n"
            "- 画像に実際に存在するデータのみ抽出すること。推測・捏造は厳禁。\n"
            "- 銀行取引以外の書類（投資信託、有価証券等）の場合は "
            '{\"lines\": [], \"transactions\": []} を返すこと。\n'
            "- 摘要（description）は金額の有無にかかわらず必ず抽出すること。\n"
            "- 「未記帳分合算」の行も必ず含めること。\n\n"
            "【口座種別】\n"
            "- 普通預金ページ → account_type: \"ordinary_deposit\"\n"
            "- 定期預金ページ → account_type: \"time_deposit\"\n"
            "- 判別不能 → null\n\n"
            "JSONのみ返すこと（説明不要）。"
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
        logger.info("Gemini raw response: length=%d, first 300 chars: %s", len(text), text[:300])
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
        logger.info("Gemini nayose raw response: length=%d, first 300 chars: %s", len(text), text[:300])
        return text

    def _build_nayose_inline_payload(self, pdf_bytes: bytes) -> Dict[str, Any]:
        """Build inline payload for nayose extraction."""
        data_bytes, mime_type = enhance_scanned_pdf(pdf_bytes)
        encoded = base64.b64encode(data_bytes).decode("ascii")
        return self._base_prompt(
            parts=[
                {"text": NAYOSE_PROMPT},
                {
                    "inline_data": {
                        "mime_type": mime_type,
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

    # ── Generic OCR ──────────────────────────────────────────

    def extract_generic_from_pdf(self, pdf_bytes: bytes) -> str:
        """Extract structured data from any document PDF.

        Returns the raw JSON response text from Gemini.
        """
        last_error: GeminiError | None = None
        for index, api_key in enumerate(self.api_keys):
            try:
                return self._extract_generic_with_key(pdf_bytes, api_key)
            except GeminiError as exc:
                last_error = exc
                if exc.can_retry_key and index < len(self.api_keys) - 1:
                    logger.warning("Gemini API key rejected (reason: %s); trying next key", exc)
                    continue
                raise
        if last_error:
            raise last_error
        raise GeminiError("Gemini API key configuration is empty")

    def _extract_generic_with_key(self, pdf_bytes: bytes, api_key: str) -> str:
        """Extract generic OCR data using a specific API key."""
        if len(pdf_bytes) <= INLINE_LIMIT_BYTES:
            payload = self._build_generic_inline_payload(pdf_bytes)
        else:
            file_name = self._upload_file(pdf_bytes, api_key=api_key, mime_type="application/pdf")
            try:
                payload = self._build_generic_file_payload(file_name)
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
        logger.info("Gemini generic_ocr raw response: length=%d, first 300 chars: %s", len(text), text[:300])
        return text

    def _build_generic_inline_payload(self, pdf_bytes: bytes) -> Dict[str, Any]:
        data_bytes, mime_type = enhance_scanned_pdf(pdf_bytes)
        encoded = base64.b64encode(data_bytes).decode("ascii")
        return self._base_prompt(
            parts=[
                {"text": GENERIC_OCR_PROMPT},
                {
                    "inline_data": {
                        "mime_type": mime_type,
                        "data": encoded,
                    }
                },
            ]
        )

    def _build_generic_file_payload(self, file_name: str) -> Dict[str, Any]:
        return self._base_prompt(
            parts=[
                {"text": GENERIC_OCR_PROMPT},
                {
                    "file_data": {
                        "mime_type": "application/pdf",
                        "file_uri": file_name,
                    }
                },
            ]
        )
