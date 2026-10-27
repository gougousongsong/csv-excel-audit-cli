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
## Kuaishou metadata integration example

The related [Kuaishou Video Metadata Lookup Actor](https://apify.com/gougousongsong/kuaishou-video-detail) turns one Kuaishou search-result video object into one normalized record for indexing, monitoring, or analytics. Its current public pricing is pay-per-event: `$0.002` per metadata result plus the configured start event.

Input shape:

```json
{"video": {"photoId": "example-photo-id", "caption": "Example caption", "authorName": "Example creator", "likeCount": 123, "commentCount": 7, "shareCount": 4, "timestamp": "2026-01-01T00:00:00Z", "durationSeconds": 18, "source": "kuaishou-search"}}
```

Observed output from a completed Actor dataset run:

```json
{"video_id": "5258515723732898742", "caption": "sample", "user_id": "", "user_name": "", "kwai_id": "", "view_count": 12, "like_count": 3, "comment_count": 1, "share_count": 0, "collect_count": 0, "published_timestamp": 0, "duration_ms": 0, "item_type": "", "source": "kuaishou-search-collector", "normalized_at": "2026-09-07T17:37:52Z"}
```

This is a metadata normalizer, not a video downloader, OCR system, sentiment model, or historical archive. It accepts one search-result JSON object per run and returns one normalized metadata item.

