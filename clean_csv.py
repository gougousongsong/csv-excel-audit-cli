#!/usr/bin/env python3
"""CSV/XLSX cleaning CLI for the CSV/Excel Cleaning & Audit Toolkit."""
from __future__ import annotations
import argparse
import csv
import json
import hashlib
from pathlib import Path
from openpyxl import Workbook, load_workbook


def clean_csv(source: Path, output: Path, duplicate_columns: list[str], missing: str) -> dict:
    if output.exists():
        raise SystemExit(f"Output already exists: {output}")
    with source.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("CSV must contain a header row")
        fields = list(reader.fieldnames)
        rows = list(reader)
    keys = duplicate_columns or fields
    missing_cells = 0
    seen = set()
    result = []
    duplicate_rows = []
    for source_row, row in enumerate(rows, 2):
        normalized = {k: " ".join((row.get(k) or "").split()) for k in fields}
        missing_fields = [k for k in fields if not normalized[k]]
        missing_cells += len(missing_fields)
        if missing == "placeholder":
            for k in missing_fields:
                normalized[k] = "[MISSING]"
        key = tuple(normalized[k] for k in keys)
        duplicate = bool(any(key) and key in seen)
        if duplicate:
            duplicate_rows.append({"source_row": source_row, "reason": "duplicate"})
        else:
            seen.add(key)
            result.append({"source_row": source_row, **normalized, "duplicate": duplicate, "missing_fields": missing_fields})
    output.mkdir(parents=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Cleaned data"
    out_fields = ["source_row", *fields, "duplicate", "missing_fields"]
    ws.append(out_fields)
    for row in result:
        ws.append([row.get(k, "") if k != "missing_fields" else ", ".join(row[k]) for k in out_fields])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(output / "cleaned.xlsx")
    wb.close()
    audit = {"source": str(source.name), "input_rows": len(rows), "output_rows": len(result), "duplicates_removed": len(duplicate_rows), "missing_cells": missing_cells, "duplicate_rows": duplicate_rows, "verified": True, "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest()}
    (output / "audit.json").write_text(json.dumps(audit, indent=2), encoding="utf-8")
    return audit


def main() -> int:
    ap = argparse.ArgumentParser(description="Clean a CSV and produce an auditable XLSX result")
    ap.add_argument("source", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--duplicate-columns", nargs="+", help="Columns used to identify duplicates")
    ap.add_argument("--missing", choices=["keep", "placeholder"], default="keep")
    args = ap.parse_args()
    audit = clean_csv(args.source, args.output, args.duplicate_columns or [], args.missing)
    print(json.dumps(audit, indent=2))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
