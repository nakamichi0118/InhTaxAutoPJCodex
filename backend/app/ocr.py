"""Azure Document Intelligence integration."""
from __future__ import annotations

import os
import time
from typing import Any, Dict, Iterable, Optional

import requests

from .config import get_settings
from .pdf_utils import PdfChunkingError, PdfChunkingPlan, chunk_pdf_by_limits

DEFAULT_API_VERSIONS = [
    "2024-02-29-preview",
    "2023-10-31-preview",
    "2023-07-31",
    "2022-08-31",
]
MODEL_ID = "prebuilt-layout"
POLL_INTERVAL_SECONDS = 1.0
POLL_TIMEOUT_SECONDS = 60.0
PATH_TEMPLATES = (
    "formrecognizer/documentModels/{model_id}:analyze?api-version={api_version}",
    "documentintelligence/documentModels/{model_id}:analyze?api-version={api_version}",
)


def _load_api_versions() -> Iterable[str]:
    raw = os.getenv("AZURE_DOCINTELLIGENCE_API_VERSIONS")
    if not raw:
        return DEFAULT_API_VERSIONS
    return [token.strip() for token in raw.split(",") if token.strip()]


class AzureFormRecognizerError(RuntimeError):
    """Raised when the Azure service returns an error payload."""


class AzureFormRecognizerClient:
    def __init__(self, endpoint: Optional[str] = None, api_key: Optional[str] = None) -> None:
        settings = get_settings()
        self.endpoint = (endpoint or settings.azure_endpoint).rstrip("/")
        self.api_key = api_key or settings.azure_key
        self.api_versions = list(_load_api_versions())
        self.max_upload_bytes = settings.azure_max_document_bytes
        self.chunk_page_limit = settings.azure_chunk_page_limit

    def analyze_layout(self, file_bytes: bytes, *, content_type: str = "application/pdf") -> Dict[str, Any]:
        last_error: Optional[str] = None
        for api_version in self.api_versions:
            for path_template in PATH_TEMPLATES:
                url = f"{self.endpoint}/{path_template.format(model_id=MODEL_ID, api_version=api_version)}"
                headers = {
                    "Ocp-Apim-Subscription-Key": self.api_key,
                    "Content-Type": content_type,
                }
                response = requests.post(url, headers=headers, data=file_bytes, timeout=30)
                if response.status_code == 202:
                    operation_url = response.headers.get("operation-location")
                    if not operation_url:
                        raise AzureFormRecognizerError("operation-location header missing in async response")
                    return self._poll_operation(operation_url)
                if response.status_code == 404:
                    last_error = self._extract_error(response)
                    continue
                if response.status_code >= 400:
                    raise AzureFormRecognizerError(self._extract_error(response))
                return response.json()
        raise AzureFormRecognizerError(last_error or "Azure Document Intelligence resource not found")

    def _poll_operation(self, url: str) -> Dict[str, Any]:
        headers = {"Ocp-Apim-Subscription-Key": self.api_key}
        deadline = time.monotonic() + POLL_TIMEOUT_SECONDS
        while True:
            if time.monotonic() > deadline:
                raise AzureFormRecognizerError("Polling timed out waiting for operation to complete")
            response = requests.get(url, headers=headers, timeout=15)
            if response.status_code == 404:
                raise AzureFormRecognizerError(self._extract_error(response))
            data = response.json()
            status = data.get("status")
            if status in {"succeeded", "failed"}:
                if status == "failed":
                    raise AzureFormRecognizerError(self._extract_error(response))
                return data
            time.sleep(POLL_INTERVAL_SECONDS)

    @staticmethod
    def _extract_error(response: requests.Response) -> str:
        try:
            payload = response.json()
            error = payload.get("error") or payload.get("errors") or payload
            return str(error)
        except ValueError:
            return response.text

