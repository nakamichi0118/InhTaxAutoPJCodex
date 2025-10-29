import csv
import io
import json
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from backend.app.parser import parse_bankbook
from backend.app.pdf_utils import PdfChunkingPlan, chunk_pdf_by_limits
from backend.app.exporter import export_to_csv_strings


FIXTURE_DIR = Path("test/fixtures/ocr_lines")


EXPECTED_TRANSACTION_COUNTS = {
    "三井住友銀行": 293,
    "取引履歴　きのくに": 5,
    "取引履歴　きのくに2": 0,
    "取引履歴（ゆうちょ銀行／14340-84250031）": 15,
    "通帳（南都銀行）": 152,
}


def load_fixture(stem: str) -> list[str]:
    fixture_path = FIXTURE_DIR / f"{stem}.json"
    return json.loads(fixture_path.read_text(encoding="utf-8"))


@pytest.mark.parametrize("stem,expected", EXPECTED_TRANSACTION_COUNTS.items())
def test_bankbook_transaction_counts(stem: str, expected: int) -> None:
    lines = load_fixture(stem)
    asset = parse_bankbook(lines, f"{stem}.pdf")[0]
    assert len(asset.transactions) == expected


def test_pdf_chunking_respects_page_limit() -> None:
    source = Path("test/通帳/三井住友銀行.pdf")
    pdf_bytes = source.read_bytes()
    plan = PdfChunkingPlan(max_bytes=10**9, max_pages=4)
    chunks = chunk_pdf_by_limits(pdf_bytes, plan)

    assert len(chunks) == 3

    from pypdf import PdfReader

    total_pages = 0
    original_page_count = len(PdfReader(io.BytesIO(pdf_bytes)).pages)
    for blob in chunks:
        reader = PdfReader(io.BytesIO(blob))
        total_pages += len(reader.pages)
        assert len(reader.pages) <= plan.max_pages
    assert total_pages == original_page_count


def test_exporter_preserves_multiline_notes() -> None:
    lines = load_fixture("三井住友銀行")
    asset = parse_bankbook(lines, "三井住友銀行.pdf")[0]
    payload = {"assets": [asset.to_export_payload()]}
    csv_map = export_to_csv_strings(payload)

    assets_csv = csv_map["assets.csv"]
    reader = csv.DictReader(io.StringIO(assets_csv))
    row = next(reader)
    assert row["notes"].count("\n") > 1

    bank_csv = csv_map["bank_transactions.csv"]
    transactions = list(csv.DictReader(io.StringIO(bank_csv)))
    assert len(transactions) == len(asset.transactions)
