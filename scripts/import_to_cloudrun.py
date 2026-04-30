"""ローカルで実行してJSONLファイルを Cloud Run にPOSTで送り込むスクリプト。

使い方:
  python scripts/import_to_cloudrun.py /path/to/railway_logs.jsonl

オプション:
  --base-url    既定: https://sorobocr-441106995152.asia-northeast1.run.app
  --password    既定: 環境変数 ANALYTICS_PASSWORD または sorobocr2024
  --batch-size  既定: 500
  --start-batch 既定: 1（1始まり、途中失敗時のリジューム用）
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.parse
import urllib.request
from pathlib import Path

DEFAULT_BASE = "https://sorobocr-441106995152.asia-northeast1.run.app"
DEFAULT_PASSWORD = os.getenv("ANALYTICS_PASSWORD", "sorobocr2024")


def main() -> int:
    p = argparse.ArgumentParser(
        description="JSONLファイルを Cloud Run の /api/analytics/import エンドポイントにインポートする"
    )
    p.add_argument("jsonl", type=Path, help="JSONLファイルパス（export_railway_logs.py の出力）")
    p.add_argument("--base-url", default=DEFAULT_BASE, help="Cloud Run サービスURL")
    p.add_argument("--password", default=DEFAULT_PASSWORD, help="アナリティクス認証パスワード")
    p.add_argument("--batch-size", type=int, default=500, help="1回に送るレコード数")
    p.add_argument("--start-batch", type=int, default=1, help="開始バッチ番号（1始まり、リジューム用）")
    args = p.parse_args()

    if not args.jsonl.exists():
        print(f"File not found: {args.jsonl}", file=sys.stderr)
        return 1

    lines = [line for line in args.jsonl.read_text(encoding="utf-8").splitlines() if line.strip()]
    rows: list[dict] = []
    for line in lines:
        try:
            rows.append(json.loads(line))
        except json.JSONDecodeError as e:
            print(f"JSONパースエラー（スキップ）: {e}", file=sys.stderr)

    print(f"Loaded {len(rows)} rows from {args.jsonl}", file=sys.stderr)

    url = f"{args.base_url}/api/analytics/import?password={urllib.parse.quote(args.password)}"
    inserted_total = 0
    skipped_total = 0

    total_batches = (len(rows) + args.batch_size - 1) // args.batch_size
    start_offset = (args.start_batch - 1) * args.batch_size
    if start_offset > 0:
        print(f"Resuming from batch {args.start_batch}/{total_batches} (skipping {start_offset} rows)", file=sys.stderr)

    for i in range(start_offset, len(rows), args.batch_size):
        batch_num = i // args.batch_size + 1
        batch = rows[i : i + args.batch_size]
        body = json.dumps(batch, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=body,
            method="POST",
            headers={"Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(req, timeout=120) as res:
                result: dict = json.loads(res.read().decode("utf-8"))
                inserted_total += result.get("inserted", 0)
                skipped_total += result.get("skipped", 0)
                print(
                    f"  Batch {batch_num}/{total_batches}: "
                    f"inserted={result.get('inserted')} skipped={result.get('skipped')}",
                    file=sys.stderr,
                )
        except Exception as e:
            print(
                f"\n*** Batch {batch_num}/{total_batches} FAILED: {e}\n"
                f"*** リトライ時は --start-batch {batch_num} を指定して再実行してください\n",
                file=sys.stderr,
            )
            return 1

    print(f"Done: inserted={inserted_total} skipped={skipped_total}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
