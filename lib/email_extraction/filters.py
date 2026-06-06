"""
Filename filter factories for .msg attachment extraction.

Provides convenience functions that build reusable predicates for filtering
attachment filenames by prefix or extension.
"""

from typing import Callable, List

from lib.utils import create_extension_filter


def create_filename_prefix_filter(prefix: str) -> Callable[[str], bool]:
    """
    Create a filter function that checks if filename starts with given prefix.

    Args:
        prefix: Prefix to match.

    Returns:
        Filter function that returns True if filename starts with prefix.
    """
    def filter_func(filename: str) -> bool:
        return filename.startswith(prefix)

    return filter_func


def create_filename_extension_filter(extensions: List[str]) -> Callable[[str], bool]:
    """
    Create a filter function that checks if filename has one of the given extensions.

    This is a thin wrapper around :func:`lib.utils.create_extension_filter`
    kept for backward compatibility.

    Args:
        extensions: List of extensions (with or without dot, e.g., ['.xlsx', 'pdf']).

    Returns:
        Filter function that returns True if filename has matching extension.
    """
    return create_extension_filter(extensions)
