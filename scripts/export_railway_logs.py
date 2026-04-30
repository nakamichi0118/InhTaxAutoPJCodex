"""Railwayコンテナ内で実行してanalytics.dbを丸ごとJSONLで出力するスクリプト。

使い方:
  Railway CLI で:
    railway login
    railway link  # プロジェクト選択
    railway run python scripts/export_railway_logs.py > /tmp/railway_logs.jsonl

stdoutにはJSONLのみ出力し、件数などの情報はstderrに出力する（リダイレクト想定）。
"""
from __future__ import annotations

import json
import os
import sqlite3
import sys
from pathlib import Path

DB_PATH = Path(os.getenv("ANALYTICS_DB_PATH", "data/analytics.db"))


def main() -> int:
    if not DB_PATH.exists():
        print(f"DB not found: {DB_PATH}", file=sys.stderr)
        return 1

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM access_logs ORDER BY id ASC").fetchall()
    conn.close()

    for row in rows:
        sys.stdout.write(json.dumps(dict(row), ensure_ascii=False) + "\n")

    print(f"Exported {len(rows)} rows", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
