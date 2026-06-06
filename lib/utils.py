"""
Shared utility functions for the acct-automation project.

Contains helpers for filename handling, GR number extraction from text,
and attachment file extension checks.
"""

import os
import re
from typing import Callable, List


# Common Excel file extensions
EXCEL_EXTENSIONS = [".xlsx", ".xls", ".xlsm", ".xlsb"]

# Pre-compiled GR number patterns
GR_PATTERN = re.compile(r"GR[\s\-_]?\d+")
GR_NUMBER_8DIGIT = re.compile(r"\b(2\d{7})\b")


def create_extension_filter(extensions: List[str]) -> Callable[[str], bool]:
    """
    Create a filter function that checks if a filename has one of the given extensions.

    This is the canonical extension-filtering factory. Use it to build
    reusable predicates for any file-type check.

    Args:
        extensions: List of extensions with leading dot (e.g. ['.xlsx', '.pdf']).

    Returns:
        A callable that takes a filename and returns True if it matches.
    """
    normalized = [ext.lower() if ext.startswith('.') else f'.{ext.lower()}' for ext in extensions]

    def _filter(filename: str) -> bool:
        return os.path.splitext(filename)[1].lower() in normalized

    return _filter


# Pre-built filter for Excel files
is_excel_file: Callable[[str], bool] = create_extension_filter(EXCEL_EXTENSIONS)
is_excel_file.__doc__ = (
    "Check if a filename has a recognized Excel extension "
    "(.xlsx, .xls, .xlsm, .xlsb — case-insensitive)."
)


def sanitize_filename(name: str, max_length: int = 50) -> str:
    """
    Remove invalid filesystem characters and truncate to a max length.

    Replaces characters that are illegal in Windows filenames
    (``\\ / : * ? " < > |``) with underscores.

    Args:
        name: Original filename or subject string.
        max_length: Maximum number of characters to keep. Defaults to 50.

    Returns:
        Sanitized string safe for use as a filename component.
    """
    sanitized = re.sub(r'[\\/:*?"<>|]', '_', name)
    return sanitized[:max_length]


def extract_gr_numbers_from_text(text: str) -> List[str]:
    """
    Extract GR number references from a body of text.

    Matches patterns like "GR 12345", "GR-12345", "GR_12345", "GR12345".

    Args:
        text: The text to search (e.g. email body or subject).

    Returns:
        Deduplicated list of matched GR reference strings.
    """
    return list(set(GR_PATTERN.findall(text)))


def make_date_prefixed_filename(received_time, original_filename: str) -> str:
    """
    Create a filename prefixed with a datetime stamp.

    Format: ``YYYYMMDD_HHMMSS_originalfilename.ext``

    Args:
        received_time: A datetime-like object with a strftime method
                       (e.g. from Outlook's ReceivedTime property).
        original_filename: The original attachment filename.

    Returns:
        New filename string with date prefix.
    """
    date_prefix = received_time.strftime("%Y%m%d_%H%M%S")
    return f"{date_prefix}_{original_filename}"
