from __future__ import annotations

import base64
import logging
from typing import Any, Dict

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from .config import get_settings
from .exporter import export_to_csv_strings
from .ocr import AzureFormRecognizerClient, AzureFormRecognizerError

logger = logging.getLogger(__name__)

app = FastAPI(title="InhTaxAutoPJ Backend", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/ping")
def ping() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/api/export")
async def export_csv(payload: Dict[str, Any]) -> Dict[str, Any]:
    try:
        csv_map = export_to_csv_strings(payload)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    encoded = {
        name: base64.b64encode(content.encode("utf-8-sig")).decode("ascii")
        for name, content in csv_map.items()
    }
    return {"status": "ok", "files": encoded}


@app.post("/api/analyze/pdf")
async def analyze_pdf(file: UploadFile = File(...)) -> Dict[str, Any]:
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Empty file uploaded")
    content_type = file.content_type or "application/pdf"
    if content_type not in {"application/pdf"}:
        logger.warning("Unexpected content type %s; defaulting to application/pdf", content_type)
        content_type = "application/pdf"

    client = AzureFormRecognizerClient()
    try:
        raw_result = client.analyze_layout(contents, content_type=content_type)
    except AzureFormRecognizerError as exc:
        logger.exception("Azure Document Intelligence error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    pages = extract_layout_pages(raw_result)
    plain_text = ["\n".join(page["lines"]) for page in pages]
    return {
        "status": "ok",
        "page_count": len(pages),
        "pages": pages,
        "plain_text": plain_text,
    }


def extract_layout_pages(result: Dict[str, Any]) -> list[Dict[str, Any]]:
    analyze = result.get("analyzeResult") or result
    pages = []
    for page in analyze.get("pages", []):
        lines = [line.get("content", "") for line in page.get("lines", [])]
        pages.append({
            "page_number": page.get("pageNumber"),
            "unit": page.get("unit"),
            "width": page.get("width"),
            "height": page.get("height"),
            "lines": lines,
        })
    if not pages and "content" in analyze:
        pages.append({"page_number": 1, "lines": analyze.get("content", "").splitlines()})
    return pages


@app.exception_handler(AzureFormRecognizerError)
async def azure_error_handler(_, exc: AzureFormRecognizerError):
    logger.exception("Unhandled Azure error: %s", exc)
    return JSONResponse(status_code=502, content={"detail": str(exc)})


@app.on_event("startup")
def log_startup() -> None:
    settings = get_settings()
    logger.info("Azure endpoint: %s", settings.azure_endpoint)
