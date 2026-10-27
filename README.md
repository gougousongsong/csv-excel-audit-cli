# CSV/Excel Audit CLI

A small, runnable Python CLI for a common operations problem: monthly order exports contain duplicate rows or missing customer fields. It produces an XLSX workbook and a JSON audit instead of silently changing data.

## Run

Requires Python 3.10+ and `openpyxl`.

```bash
python -m pip install -r requirements.txt
python clean_csv.py sample_orders.csv output --duplicate-columns customer_email --missing placeholder
```

The output directory must not already exist. It contains `cleaned.xlsx` and `audit.json`. The included sample has 4 input rows, one later duplicate, and one missing email.

## What it does

- reads ordinary UTF-8 CSV input;
- normalizes surrounding whitespace;
- identifies later duplicates using selected columns;
- records source row numbers and missing fields;
- writes XLSX output and a machine-readable audit with row counts and duplicate rows.

## Exact limits

This demo does not provide PDF OCR, macros/VBA, pivot tables, database access, customer-system writes, ongoing monitoring, or personal-data redaction. Use synthetic data for this demo. Customer-specific processing is a separate scoped service.

## Links

- Free runnable sample: https://2477537799006.gumroad.com/l/ebhbcm?utm_source=github_demo&utm_medium=repository&utm_campaign=csv_audit
- Full toolkit: https://2477537799006.gumroad.com/l/csv-excel-cleaning-audit-toolkit?utm_source=github_demo&utm_medium=repository&utm_campaign=csv_audit

AI-assisted delivery is disclosed.
