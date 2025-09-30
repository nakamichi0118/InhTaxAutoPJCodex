"""PDF helper utilities for preprocessing large documents."""
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Iterable, List

from pypdf import PdfReader, PdfWriter


class PdfChunkingError(RuntimeError):
    """Raised when the PDF cannot be reduced to acceptable chunks."""


@dataclass(frozen=True)
class PdfChunkingPlan:
    max_bytes: int
    max_pages: int


def chunk_pdf_by_limits(pdf_bytes: bytes, plan: PdfChunkingPlan) -> List[bytes]:
    """Split a PDF into byte-limited chunks while preserving page order.

    Args:
        pdf_bytes: Raw PDF content.
        plan: Chunking constraints.

    Returns:
        List of PDF byte blobs, each within the configured limits.

    Raises:
        PdfChunkingError: If a single page exceeds the size limit and cannot be chunked.
    """

    if len(pdf_bytes) <= plan.max_bytes:
        return [pdf_bytes]

    reader = PdfReader(BytesIO(pdf_bytes))
    total_pages = len(reader.pages)
    if total_pages == 0:
        return [pdf_bytes]

    def write_pages(pages: Iterable) -> bytes:
        writer = PdfWriter()
        for page in pages:
            writer.add_page(page)
        buffer = BytesIO()
        writer.write(buffer)
        return buffer.getvalue()

    chunks: List[bytes] = []
    current_pages: List = []

    for page_index, page in enumerate(reader.pages):
        tentative_pages = current_pages + [page]
        tentative_blob = write_pages(tentative_pages)

        if len(tentative_blob) > plan.max_bytes and current_pages:
            current_blob = write_pages(current_pages)
            chunks.append(current_blob)
            current_pages = [page]
            single_blob = write_pages(current_pages)
            if len(single_blob) > plan.max_bytes:
                raise PdfChunkingError(
                    "Single page exceeds Azure upload limit; compression required."
                )
            tentative_blob = single_blob
        elif len(tentative_blob) > plan.max_bytes:
            raise PdfChunkingError(
                "Single page exceeds Azure upload limit; compression required."
            )

        current_pages.append(page)

        if len(current_pages) >= plan.max_pages:
            chunks.append(tentative_blob)
            current_pages = []

    if current_pages:
        chunks.append(write_pages(current_pages))

    return chunks
