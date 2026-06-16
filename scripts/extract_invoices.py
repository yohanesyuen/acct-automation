"""
Extract invoice data from PDFs for all configured companies.

Two modes:
  1. Folder scan (recommended): set 'pdf_dir' to a folder. PDFs are discovered
     recursively and the company is guessed automatically from content.
  2. Per-company: set 'uie_pdf', 'lp_pdf', and/or 'mjm_pdf' to specific PDF paths.

Usage:
    python scripts/extract_invoices.py
    python scripts/extract_invoices.py --no-gui --pdf-dir /path/to/invoices
    python scripts/extract_invoices.py --no-gui --uie-pdf path/to/uie.pdf
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

COMPANIES = [
    ('uie_pdf',  'UIE Industrial',  UIEIndustrialExtractor),
    ('lp_pdf',   'LP Construction', LPConstructionExtractor),
    ('mjm_pdf',  'MJM Services',    MJMServicesExtractor),
]

_EXTRACTOR_LABELS = {
    UIEIndustrialExtractor:  'UIE Industrial',
    LPConstructionExtractor: 'LP Construction',
    MJMServicesExtractor:    'MJM Services',
}


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

    if pdf_dir and os.path.isdir(pdf_dir):
        # --- Folder scan mode ---
        pdfs = list(find_pdfs(pdf_dir))
        if not pdfs:
            print(f"No PDF files found in: {pdf_dir}")
            return

        print(f"Scanning {len(pdfs)} PDF(s) in {pdf_dir}\n")

        for pdf_path in pdfs:
            print(f"\n{'='*60}")
            print(f"  {pdf_path.name}")

            extractor_cls = guess_extractor(str(pdf_path))
            if extractor_cls is None:
                print("  No matching extractor found — skipping.")
                continue

            label = _EXTRACTOR_LABELS.get(extractor_cls, extractor_cls.__name__)
            print(f"  Detected company: {label}")
            print(f"{'='*60}")

            try:
                result = extractor_cls().extract(str(pdf_path))
                _print_result(result)
            except Exception as e:
                print(f"  Error: {e}")

    else:
        # --- Per-company mode ---
        for pdf_key, label, extractor_cls in COMPANIES:
            pdf_path = config.get(pdf_key, '')
            if not pdf_path or not os.path.isfile(pdf_path):
                print(f"\n  Skipping {label} — no PDF path configured.")
                continue

            print(f"\n{'='*60}")
            print(f"  {label}")
            print(f"{'='*60}")

            try:
                result = extractor_cls().extract(pdf_path)
                _print_result(result)
            except Exception as e:
                print(f"  Error: {e}")


if __name__ == '__main__':
    config = parse_task_args(
        description="Extract invoice data from PDFs for all companies.",
        default_task="extract_invoices",
        config_keys=['pdf_dir', 'uie_pdf', 'lp_pdf', 'mjm_pdf'],
    )
    extract_invoices(config)
