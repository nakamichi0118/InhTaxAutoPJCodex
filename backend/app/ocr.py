"""Azure Document Intelligence integration."""
from __future__ import annotations

import time
from typing import Any, Dict, Optional

import requests

from .config import get_settings

API_VERSION = "2024-02-29-preview"
MODEL_ID = "prebuilt-layout"
POLL_INTERVAL_SECONDS = 1.0
POLL_TIMEOUT_SECONDS = 60.0


class AzureFormRecognizerError(RuntimeError):
    """Raised when the Azure service returns an error payload."""


class AzureFormRecognizerClient:
    def __init__(self, endpoint: Optional[str] = None, api_key: Optional[str] = None) -> None:
        settings = get_settings()
        self.endpoint = (endpoint or settings.azure_endpoint).rstrip("/")
        self.api_key = api_key or settings.azure_key

    def analyze_layout(self, file_bytes: bytes, *, content_type: str = "application/pdf") -> Dict[str, Any]:
        url = f"{self.endpoint}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version={API_VERSION}"
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
        if response.status_code >= 400:
            raise AzureFormRecognizerError(self._extract_error(response))
        return response.json()

    def _poll_operation(self, url: str) -> Dict[str, Any]:
        headers = {"Ocp-Apim-Subscription-Key": self.api_key}
        deadline = time.monotonic() + POLL_TIMEOUT_SECONDS
        while True:
            if time.monotonic() > deadline:
                raise AzureFormRecognizerError("Polling timed out waiting for operation to complete")
            response = requests.get(url, headers=headers, timeout=15)
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
