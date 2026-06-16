"""
Extract invoice data from PDFs in a folder.

PDFs are discovered recursively under 'pdf_dir'. By default the vendor is
auto-detected from each PDF's content ('auto_detect_vendor': True). Disable
auto-detection — via the GUI checkbox, the YAML, or by passing '--vendor' on
the CLI — to force a single vendor's extractor for every PDF.

Usage:
    python scripts/extract_invoices.py
    python scripts/extract_invoices.py --no-gui --pdf-dir /path/to/invoices
    python scripts/extract_invoices.py --no-gui --pdf-dir /path --vendor "MJM Services"
"""

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
from lib.task_config import parse_task_args

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
            _print_result(result)
        except Exception as e:
            print(f"  Error: {e}")


if __name__ == '__main__':
    # Remember whether the vendor was explicitly passed before argv is consumed.
    _vendor_on_cli = any(a == '--vendor' or a.startswith('--vendor=') for a in sys.argv[1:])

    config = parse_task_args(
        description="Extract invoice data from PDFs in a folder.",
        default_task="extract_invoices",
        config_keys=['pdf_dir', 'auto_detect_vendor', 'vendor'],
    )

    # If the vendor was explicitly passed as a CLI arg, force it.
    if _vendor_on_cli:
        config['auto_detect_vendor'] = False

    extract_invoices(config)
