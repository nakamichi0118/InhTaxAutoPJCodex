"""PDF helper utilities for preprocessing large documents."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from io import BytesIO
from typing import Iterable, List, Optional, Tuple

from pypdf import PdfReader, PdfWriter

logger = logging.getLogger(__name__)


def enhance_scanned_page(pdf_bytes: bytes, page_index: int = 0, dpi: int = 300) -> Optional[bytes]:
    """Render and enhance a PDF page for better OCR accuracy.

    Renders the full page as a pixmap (correctly composites all embedded images
    and vector content), then applies grayscale + autocontrast + sharpening.
    Returns enhanced PNG bytes.
    """
    try:
        import fitz  # PyMuPDF
        from PIL import Image, ImageEnhance, ImageOps
        import io
    except ImportError:
        return None

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_index >= len(doc):
            doc.close()
            return None

        page = doc[page_index]
        # Render the full page at specified DPI (composites all layers)
        pix = page.get_pixmap(dpi=dpi)
        raw_png = pix.tobytes("png")
        raw_size = len(raw_png)
        doc.close()

        img = Image.open(io.BytesIO(raw_png))

        # Convert to grayscale for cleaner OCR
        img = ImageOps.grayscale(img)
        # Auto-contrast: stretch histogram to use full range
        img = ImageOps.autocontrast(img, cutoff=1)
        # Sharpen to make blurry dot-printer text clearer
        img = ImageEnhance.Sharpness(img).enhance(3.0)
        # Boost contrast to make faint text stand out
        img = ImageEnhance.Contrast(img).enhance(2.0)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        enhanced_bytes = buf.getvalue()

        logger.info(
            "Enhanced page %d: rendered %dx%d @ %ddpi, %dKB -> %dKB",
            page_index,
            img.width, img.height,
            dpi,
            raw_size // 1024,
            len(enhanced_bytes) // 1024,
        )
        return enhanced_bytes

    except Exception as exc:
        logger.warning("Failed to enhance page %d: %s", page_index, exc)
        return None


def enhance_scanned_pdf(pdf_bytes: bytes) -> Tuple[bytes, str]:
    """Try to enhance a single-page scanned PDF for OCR.

    Returns (image_bytes, mime_type) if enhancement succeeded,
    or (pdf_bytes, "application/pdf") if not applicable.
    """
    enhanced = enhance_scanned_page(pdf_bytes, page_index=0)
    if enhanced is not None:
        return enhanced, "image/png"
    return pdf_bytes, "application/pdf"


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


def compress_pdf(pdf_bytes: bytes, max_bytes: int, dpi: int = 200, quality: int = 80) -> bytes:
    """Compress a PDF by re-rendering pages as JPEG images at reduced DPI.

    This handles scanned PDFs with huge embedded PNG images (e.g. 101MB for 15 pages).
    Each page is rendered to a JPEG image and reassembled into a new PDF.

    Args:
        pdf_bytes: Original PDF content.
        max_bytes: Target maximum size. If already under this, returns unchanged.
        dpi: Render resolution (default 200 — good balance for OCR).
        quality: JPEG quality 1-100 (default 80).

    Returns:
        Compressed PDF bytes, or the original if already small enough.
    """
    if len(pdf_bytes) <= max_bytes:
        return pdf_bytes

    try:
        import fitz  # PyMuPDF
    except ImportError:
        logger.warning("PyMuPDF not available; cannot compress PDF (%d bytes)", len(pdf_bytes))
        return pdf_bytes

    original_mb = len(pdf_bytes) / (1024 * 1024)
    logger.info("Compressing PDF: %.1fMB -> target %.1fMB (dpi=%d, quality=%d)",
                original_mb, max_bytes / (1024 * 1024), dpi, quality)

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    new_doc = fitz.open()

    for page_index in range(len(doc)):
        page = doc[page_index]
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("jpeg", jpg_quality=quality)

        # Create a new page with the same dimensions and insert the JPEG
        img_doc = fitz.open(stream=img_bytes, filetype="jpeg")
        pdf_page_bytes = img_doc.convert_to_pdf()
        img_doc.close()

        img_pdf = fitz.open(stream=pdf_page_bytes, filetype="pdf")
        new_doc.insert_pdf(img_pdf)
        img_pdf.close()

    result = new_doc.tobytes(deflate=True)
    new_doc.close()
    doc.close()

    compressed_mb = len(result) / (1024 * 1024)
    logger.info("PDF compressed: %.1fMB -> %.1fMB (%d pages)",
                original_mb, compressed_mb, page_index + 1)

    return result
