"""
Extract Excel attachments from Outlook emails that match GRN criteria.

This is functionally identical to extract_grn.py but uses the
'extract_grn_excel' task config (different report filename).

Usage:
    python scripts/extract_grn_excel.py
    python scripts/extract_grn_excel.py --sender-email someone@example.com
"""

import sys
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.task_config import parse_task_args
from extract_grn import extract_grn

if __name__ == "__main__":
    config = parse_task_args(
        description="Extract Excel attachments from Outlook emails (GRN criteria).",
        default_task="extract_grn_excel",
    )
    extract_grn(config)
