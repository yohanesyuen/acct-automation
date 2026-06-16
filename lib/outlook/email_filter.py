"""
Email filtering utilities for Outlook inbox items.

Provides generator-based filters that yield matching email items from
an Outlook folder based on sender address, keyword, and attachment criteria.
"""

from typing import Generator, List, Optional, Union


def filter_emails(
    folder,
    sender_email: Optional[Union[str, List[str]]] = None,
    keyword: Optional[Union[str, List[str]]] = None,
    require_attachments: bool = True,
    verbose: bool = False,
    include_subfolders: bool = False,
) -> Generator:
    """
    Yield emails matching optional sender address and keyword filters.

    Iterates over all items in the given Outlook folder (and optionally all
    of its subfolders), yielding those that satisfy all provided criteria.
    All filters are optional.

    Args:
        folder: Outlook folder COM object (e.g. Inbox) to iterate.
        sender_email: Optional sender filter. Can be:
                      - A single string (substring match, case-insensitive)
                      - A list of strings (match if ANY sender matches)
                      - None (no sender filtering)
        keyword: Optional keyword filter. Can be:
                 - A single string (match in subject or body, case-insensitive)
                 - A list of strings (match if ANY keyword found in subject or body)
                 - None (no keyword filtering)
        require_attachments: If True, only yield emails that have at
                            least one attachment.
        verbose: If True, print diagnostic info about filter decisions.
        include_subfolders: If True, also search all subfolders recursively.

    Yields:
        Outlook MailItem COM objects that match all criteria.
    """
    # Normalize sender_email to a list for uniform handling
    if sender_email is None:
        sender_list = None
    elif isinstance(sender_email, str):
        sender_list = [sender_email] if sender_email else None
    else:
        sender_list = list(sender_email) if sender_email else None

    # Normalize keyword to a list for uniform handling
    if keyword is None:
        keyword_list = None
    elif isinstance(keyword, str):
        keyword_list = [keyword] if keyword else None
    else:
        keyword_list = list(keyword) if keyword else None

    # Collect folders to search
    if include_subfolders:
        from .connection import iter_subfolders
        folders = list(iter_subfolders(folder))
        if verbose:
            print(f"  [filter] Searching {len(folders)} folder(s) (including subfolders)")
    else:
        folders = [folder]

    total_scanned = 0
    sender_matched = 0
    keyword_matched = 0
    attachment_filtered = 0
    yielded = 0

    for current_folder in folders:
        if verbose and include_subfolders:
            try:
                print(f"  [filter] Scanning: {current_folder.Name}")
            except Exception:
                pass

        for email in current_folder.Items:
            try:
                sender = email.SenderEmailAddress or ""
                subject = email.Subject or ""
                body = email.Body or ""
            except Exception:
                continue

            total_scanned += 1

            # Sender match (if specified) — match if ANY entry is a substring
            if sender_list:
                sender_lower = sender.lower()
                if not any(s.lower() in sender_lower for s in sender_list):
                    if verbose and not include_subfolders and total_scanned <= 5:
                        print(f"  [filter] SKIP sender mismatch: {sender[:60]}")
                    continue
                sender_matched += 1

            # Keyword match (if specified) — match if ANY keyword in subject or body
            if keyword_list:
                subject_lower = subject.lower()
                body_lower = body.lower()
                if not any(k.lower() in subject_lower or k.lower() in body_lower for k in keyword_list):
                    if verbose:
                        print(f"  [filter] SKIP keyword mismatch: \"{subject[:50]}\" (sender: {sender[:40]})")
                    continue
                keyword_matched += 1

            # Attachment requirement
            if require_attachments and email.Attachments.Count == 0:
                attachment_filtered += 1
                if verbose:
                    print(f"  [filter] SKIP no attachments: \"{subject[:50]}\"")
                continue

            yielded += 1
            if verbose and yielded <= 10:
                print(f"  [filter] MATCH: \"{subject[:50]}\" (sender: {sender[:40]}, "
                      f"attachments: {email.Attachments.Count})")
            yield email

    if verbose:
        print(f"\n  [filter] Summary: scanned={total_scanned}, sender_match={sender_matched}, "
              f"keyword_match={keyword_matched}, no_attachments_skipped={attachment_filtered}, "
              f"yielded={yielded}")


def filter_emails_by_sender_and_keyword(
    inbox,
    sender_email: Union[str, List[str]],
    keyword: Optional[Union[str, List[str]]] = None,
    require_attachments: bool = True,
) -> Generator:
    """
    Yield emails matching a sender address (or list) and optional keyword filter.

    This is a convenience wrapper around :func:`filter_emails`.

    Args:
        inbox: Outlook folder COM object (e.g. Inbox) to iterate.
        sender_email: Substring (or list of substrings) to match against
                      SenderEmailAddress (case-insensitive).
        keyword: Optional keyword (or list of keywords) to search for in
                 Subject and Body. If None, no keyword filtering is applied.
        require_attachments: If True, only yield emails that have at
                            least one attachment.

    Yields:
        Outlook MailItem COM objects that match all criteria.
    """
    return filter_emails(inbox, sender_email=sender_email, keyword=keyword, require_attachments=require_attachments)


def filter_emails_by_keyword(
    inbox,
    keyword: Union[str, List[str]],
    require_attachments: bool = True,
) -> Generator:
    """
    Yield emails containing a keyword (or any of a list) in subject or body.

    This is a convenience wrapper around :func:`filter_emails`.

    Args:
        inbox: Outlook folder COM object (e.g. Inbox) to iterate.
        keyword: Keyword or list of keywords to search for in Subject and Body.
        require_attachments: If True, only yield emails that have at
                            least one attachment.

    Yields:
        Outlook MailItem COM objects that match the criteria.
    """
    return filter_emails(inbox, keyword=keyword, require_attachments=require_attachments)
