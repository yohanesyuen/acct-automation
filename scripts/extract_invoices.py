"""
Extract invoice data from PDFs in a folder.

PDFs are discovered recursively under 'pdf_dir'. By default the vendor is
auto-detected from each PDF's content ('auto_detect_vendor': True). Disable
auto-detection — via the GUI checkbox, the YAML, or by passing '--vendor' on
the CLI — to force a single vendor's extractor for every PDF.

Extracted data is saved as CSV under the task root for later processing:
  - '<report_file>.csv': one row per PDF with the scalar invoice fields.
  - '<report_file>_line_items.csv': one row per line item per invoice
    (source_file, vendor, line_no, then each vendor's line-item columns;
    columns are the union across vendors, blank where not applicable).

Usage:
    python scripts/extract_invoices.py
    python scripts/extract_invoices.py --no-gui --pdf-dir /path/to/invoices
    python scripts/extract_invoices.py --no-gui --pdf-dir /path --vendor "MJM Services"
"""

import csv
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.extraction import (
    UIEIndustrialExtractor,
    LPConstructionExtractor,
    MJMServicesExtractor,
    find_pdfs,
    guess_extractor,
)
from lib.task_config import parse_task_args, get_report_path

# Vendor display name -> extractor class. Display names mirror the GUI dropdown
# options in FIELD_TYPES['vendor'].
VENDOR_EXTRACTORS = {
    'UIE Industrial':  UIEIndustrialExtractor,
    'LP Construction': LPConstructionExtractor,
    'MJM Services':    MJMServicesExtractor,
}

_EXTRACTOR_LABELS = {cls: name for name, cls in VENDOR_EXTRACTORS.items()}


def _print_result(result: dict) -> None:
    line_items = result.pop('line_items', None)

    print("\n--- Invoice fields ---")
    for field, value in result.items():
        print(f"  {field:20s}: {value}")

    if line_items:
        print("\n--- Line items ---")
        headers, *rows = line_items
        col_widths = [max(len(str(r[i])) for r in [headers] + rows) for i in range(len(headers))]
        fmt = '  '.join(f'{{:<{w}}}' for w in col_widths)
        print(fmt.format(*headers))
        print(fmt.format(*['-' * w for w in col_widths]))
        for row in rows:
            print(fmt.format(*row))
    else:
        print("\n  No line items found.")


def _write_csv(path: str, fieldnames: list, rows: list) -> None:
    """Write rows (list of dicts) to a CSV. utf-8-sig opens cleanly in Excel."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, restval="", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _save_reports(config, summary_rows, field_keys, line_item_rows, line_item_cols) -> None:
    """Persist extracted fields and line items as CSVs under the task root."""
    if not summary_rows:
        print("\nNothing extracted — no report written.")
        return

    # Force a .csv extension regardless of the configured report_file name.
    stem = os.path.splitext(config.get("report_file") or "")[0] or "Invoice_Extract_Report"

    summary_path = get_report_path(config, f"{stem}.csv")
    _write_csv(summary_path, ["source_file", "vendor"] + field_keys, summary_rows)
    print(f"\nSaved invoice fields → {summary_path}")

    if line_item_rows:
        line_items_path = get_report_path(config, f"{stem}_line_items.csv")
        _write_csv(
            line_items_path,
            ["source_file", "vendor", "line_no"] + line_item_cols,
            line_item_rows,
        )
        print(f"Saved line items    → {line_items_path}")


def extract_invoices(config):
    pdf_dir = (config.get("pdf_dir") or "").strip()
    auto_detect = bool(config.get("auto_detect_vendor", True))
    vendor = (config.get("vendor") or "").strip()

    if not pdf_dir or not os.path.isdir(pdf_dir):
        print(f"PDF directory not found: {pdf_dir or '(not set)'}")
        return

    # When auto-detection is off, force one vendor's extractor for every PDF.
    forced_cls = None
    if not auto_detect:
        forced_cls = VENDOR_EXTRACTORS.get(vendor)
        if forced_cls is None:
            print(f"Unknown vendor: {vendor!r}. "
                  f"Choose one of: {', '.join(VENDOR_EXTRACTORS)}.")
            return

    pdfs = list(find_pdfs(pdf_dir))
    if not pdfs:
        print(f"No PDF files found in: {pdf_dir}")
        return

    mode = "auto-detect" if auto_detect else f"forced vendor: {vendor}"
    print(f"Scanning {len(pdfs)} PDF(s) in {pdf_dir} ({mode})\n")

    summary_rows = []     # one dict per PDF (scalar invoice fields)
    field_keys = []       # ordered union of scalar field names across PDFs
    line_item_rows = []   # one dict per line item per invoice
    line_item_cols = []   # ordered union of line-item column names across PDFs

    for pdf_path in pdfs:
        print(f"\n{'='*60}")
        print(f"  {pdf_path.name}")

        extractor_cls = forced_cls or guess_extractor(str(pdf_path))
        if extractor_cls is None:
            print("  No matching extractor found — skipping.")
            continue

        label = _EXTRACTOR_LABELS.get(extractor_cls, extractor_cls.__name__)
        print(f"  Vendor: {label}")
        print(f"{'='*60}")

        try:
            result = extractor_cls().extract(str(pdf_path))
        except Exception as e:
            print(f"  Error: {e}")
            continue

        line_items = result.get("line_items")
        _print_result(result)  # pops 'line_items'; leaves scalar fields

        summary = {"source_file": pdf_path.name, "vendor": label}
        summary.update(result)
        summary_rows.append(summary)
        for key in result:
            if key not in field_keys:
                field_keys.append(key)

        if line_items:
            headers, *rows = line_items
            for col in headers:
                if col not in line_item_cols:
                    line_item_cols.append(col)
            for line_no, row in enumerate(rows, 1):
                item = {"source_file": pdf_path.name, "vendor": label, "line_no": line_no}
                item.update(dict(zip(headers, row)))
                line_item_rows.append(item)

    _save_reports(config, summary_rows, field_keys, line_item_rows, line_item_cols)


if __name__ == '__main__':
    # Remember whether the vendor was explicitly passed before argv is consumed.
    _vendor_on_cli = any(a == '--vendor' or a.startswith('--vendor=') for a in sys.argv[1:])

    config = parse_task_args(
        description="Extract invoice data from PDFs in a folder.",
        default_task="extract_invoices",
        config_keys=['pdf_dir', 'auto_detect_vendor', 'vendor', 'report_file'],
    )

    # If the vendor was explicitly passed as a CLI arg, force it.
    if _vendor_on_cli:
        config['auto_detect_vendor'] = False

    extract_invoices(config)
