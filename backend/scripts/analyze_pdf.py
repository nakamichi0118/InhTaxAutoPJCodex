import argparse
import json
from pathlib import Path

from backend.app.ocr import AzureFormRecognizerClient
from backend.app.main import extract_layout_pages


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Analyze a PDF with Azure Document Intelligence")
    parser.add_argument("pdf", type=Path, help="Path to the PDF file")
    parser.add_argument("--out", type=Path, default=None, help="Optional path to save raw JSON response")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.pdf.exists():
        raise SystemExit(f"File not found: {args.pdf}")
    data = args.pdf.read_bytes()
    client = AzureFormRecognizerClient()
    result = client.analyze_layout(data)
    pages = extract_layout_pages(result)
    for page in pages:
        print(f"--- page {page.get('page_number')} ---")
        for line in page.get("lines", []):
            print(line)
        print()
    if args.out:
        args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
