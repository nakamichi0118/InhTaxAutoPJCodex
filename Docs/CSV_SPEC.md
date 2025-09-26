# Inheritance CSV Specification

This project exports structured CSV outputs that can be imported into the existing Excel-based inheritance inventory workbook. Two files are generated per processing run:

1. `assets.csv` – one row per asset that was recognised from the document set.
2. `bank_transactions.csv` – optional file that lists detailed transactions for bankbooks when available.

Both files are encoded in UTF-8 with BOM to remain compatible with Microsoft Excel on Windows.

---

## 1. `assets.csv`

| Column | Required | Description |
| --- | --- | --- |
| `record_id` | Yes | Stable identifier for the row (UUID v4). |
| `source_document` | Yes | File name (including extension) that produced this record. |
| `asset_category` | Yes | High level bucket: `land`, `building`, `bank_deposit`, `listed_security`, `bond`, `insurance`, `mutual_fund`, `other`. |
| `asset_type` | Optional | Sub classification (for example `jotochi`, `ordinary_deposit`, `time_deposit`, `public_bond`). |
| `owner_name` | Optional | Name of the asset owner if it can be detected. Multiple owners are joined with `;`. |
| `asset_name` | Optional | Human readable label (e.g. property name, account nickname, security name). |
| `location_prefecture` | Optional | Prefecture (JP47 code or name) for real estate. |
| `location_municipality` | Optional | Municipality / ward / city. |
| `location_detail` | Optional | Remaining address details such as chome, banchi, building, parcel numbers. |
| `identifier_primary` | Optional | Key identifier (parcel number, account number, ISIN, certificate number, policy number). |
| `identifier_secondary` | Optional | Secondary identifier (branch name, registry office, sub-account id). |
| `valuation_basis` | Optional | Basis used to compute amount (e.g. "固定資産税評価額", "残高証明"). |
| `valuation_currency` | Yes | ISO currency code; defaults to `JPY`. |
| `valuation_amount` | Optional | Numeric amount in the specified currency. |
| `valuation_date` | Optional | ISO 8601 date (`YYYY-MM-DD`) that the valuation is effective. |
| `ownership_share` | Optional | Percentage (0-100) represented as decimal (e.g. 50.0). |
| `notes` | Optional | Free text remarks (OCR confidence, manual adjustments, memo). |

### Normalisation rules
- Any missing value is emitted as an empty cell.
- Numeric fields (`valuation_amount`, `ownership_share`) are exported without thousands separators.
- Strings are trimmed and internal new lines are collapsed into single spaces.

---

## 2. `bank_transactions.csv`

A supplementary table for detailed bankbook entries. The file is only produced when at least one bank transaction is extracted.

| Column | Required | Description |
| --- | --- | --- |
| `record_id` | Yes | UUID v4 that matches the `assets.csv` record representing the parent bank account. |
| `transaction_id` | Yes | Deterministic identifier `(record_id + transaction_index)` to help de-duplication. |
| `transaction_date` | Optional | ISO 8601 date. |
| `value_date` | Optional | Date of value if different from transaction date. |
| `description` | Optional | Narration / 摘要. |
| `withdrawal_amount` | Optional | 出金額 (numeric, JPY). |
| `deposit_amount` | Optional | 入金額 (numeric, JPY). |
| `balance` | Optional | 残高 after the transaction (numeric, JPY). |
| `line_confidence` | Optional | Floating point score (0.0–1.0) returned by OCR/LLM. |

Rules mirror those of `assets.csv` (UTF-8 BOM, empty string for missing text, plain numbers for amounts).

---

## 3. Intermediate JSON schema

The pipeline converts recognised documents into the following normalised JSON before rendering CSVs. Scripts in this repository expect this shape.

```json
{
  "assets": [
    {
      "category": "bank_deposit",
      "type": "ordinary_deposit",
      "source_document": "doc001.pdf",
      "owner_name": ["山田 太郎"],
      "asset_name": "三菱UFJ銀行 普通預金",
      "location": {
        "prefecture": null,
        "municipality": null,
        "detail": null
      },
      "identifiers": {
        "primary": "1234567",
        "secondary": "渋谷支店"
      },
      "valuation": {
        "basis": "残高証明",
        "currency": "JPY",
        "amount": 1234567,
        "date": "2025-01-31"
      },
      "ownership_share": 100.0,
      "notes": "OCR confidence 0.92",
      "transactions": [
        {
          "transaction_date": "2025-01-05",
          "value_date": null,
          "description": "給与振込",
          "withdrawal_amount": 0,
          "deposit_amount": 350000,
          "balance": 1250000,
          "confidence": 0.88
        }
      ]
    }
  ]
}
```

---

## 4. Validation checklist

- `record_id` is generated per asset when exporting.
- `bank_transactions.csv` rows must reference an existing `record_id`.
- Scripts should deduplicate whitespace and normalise calendar conversions (e.g. 令和→西暦) before export.
- Downstream tooling can join the two CSVs on `record_id` to obtain summary + transaction detail.

This specification intentionally keeps the surface small so that we can evolve the schema without breaking Excel import macros: adding columns is safe; renaming requires coordinated updates.
