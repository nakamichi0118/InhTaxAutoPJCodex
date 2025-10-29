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

    reader = PdfReader(BytesIO(pdf_bytes))
    total_pages = len(reader.pages)
    if total_pages == 0:
        return [pdf_bytes]

    if len(pdf_bytes) <= plan.max_bytes and total_pages <= plan.max_pages:
        return [pdf_bytes]

    def write_pages(indices: Iterable[int]) -> bytes:
        writer = PdfWriter()
        for idx in indices:
            writer.add_page(reader.pages[idx])
        buffer = BytesIO()
        writer.write(buffer)
        return buffer.getvalue()

    chunks: List[bytes] = []
    current_indices: List[int] = []

    for page_index in range(total_pages):
        current_indices.append(page_index)
        tentative_blob = write_pages(current_indices)

        if len(tentative_blob) > plan.max_bytes:
            if len(current_indices) == 1:
                raise PdfChunkingError(
                    "Single page exceeds analysis upload limit; compression required."
                )
            last_index = current_indices.pop()
            chunks.append(write_pages(current_indices))
            current_indices = [last_index]
            tentative_blob = write_pages(current_indices)
            if len(tentative_blob) > plan.max_bytes:
                raise PdfChunkingError(
                    "Single page exceeds analysis upload limit; compression required."
                )

        if len(current_indices) >= plan.max_pages:
            chunks.append(tentative_blob)
            current_indices = []

    if current_indices:
        chunks.append(write_pages(current_indices))

    return chunks
