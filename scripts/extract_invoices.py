"""
Extract invoice data from PDFs for all configured companies.

Runs each company's extractor against its configured PDF path and prints
the extracted fields and line items. PDF paths are stored in tasks/extract_invoices.yml.

Usage:
    python scripts/extract_invoices.py
    python scripts/extract_invoices.py --no-gui --uie-pdf path/to/uie.pdf
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.extraction import UIEIndustrialExtractor, LPConstructionExtractor, MJMServicesExtractor
from lib.task_config import parse_task_args

COMPANIES = [
    ('uie_pdf',  'UIE Industrial',  UIEIndustrialExtractor),
    ('lp_pdf',   'LP Construction', LPConstructionExtractor),
    ('mjm_pdf',  'MJM Services',    MJMServicesExtractor),
]


def extract_invoices(config):
    for pdf_key, label, extractor_cls in COMPANIES:
        pdf_path = config.get(pdf_key, '')
        if not pdf_path or pdf_path == 'PDF_PATH':
            print(f"\n  Skipping {label} — no PDF path configured.")
            continue

        print(f"\n{'='*60}")
        print(f"  {label}")
        print(f"{'='*60}")

        result = extractor_cls().extract(pdf_path)
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


if __name__ == '__main__':
    config = parse_task_args(
        description="Extract invoice data from PDFs for all companies.",
        default_task="extract_invoices",
        config_keys=['uie_pdf', 'lp_pdf', 'mjm_pdf'],
    )
    extract_invoices(config)
