"""
Email filtering utilities for Outlook inbox items.

Provides generator-based filters that yield matching email items from
an Outlook folder based on sender address, keyword, and attachment criteria.
"""

from typing import Generator, Optional


def filter_emails(
    inbox,
    sender_email: Optional[str] = None,
    keyword: Optional[str] = None,
    require_attachments: bool = True,
) -> Generator:
    """
    Yield emails matching optional sender address and keyword filters.

    Iterates over all items in the given Outlook folder, yielding those
    that satisfy all provided criteria. All filters are optional — if none
    are specified, all items (optionally with attachments) are yielded.

    Args:
        inbox: Outlook folder COM object (e.g. Inbox) to iterate.
        sender_email: Optional substring to match against SenderEmailAddress
                      (case-insensitive). If None, sender is not filtered.
        keyword: Optional keyword to search for in Subject and Body
                 (case-insensitive). If None, no keyword filtering is applied.
        require_attachments: If True, only yield emails that have at
                            least one attachment.

    Yields:
        Outlook MailItem COM objects that match all criteria.
    """
    for email in inbox.Items:
        try:
            sender = email.SenderEmailAddress or ""
            subject = email.Subject or ""
            body = email.Body or ""
        except Exception:
            continue

        # Sender match (if specified)
        if sender_email and sender_email.lower() not in sender.lower():
            continue

        # Keyword match (if specified)
        if keyword:
            if keyword.lower() not in subject.lower() and keyword.lower() not in body.lower():
                continue

        # Attachment requirement
        if require_attachments and email.Attachments.Count == 0:
            continue

        yield email


def filter_emails_by_sender_and_keyword(
    inbox,
    sender_email: str,
    keyword: Optional[str] = None,
    require_attachments: bool = True,
) -> Generator:
    """
    Yield emails matching a sender address and optional keyword filter.

    This is a convenience wrapper around :func:`filter_emails`.

    Args:
        inbox: Outlook folder COM object (e.g. Inbox) to iterate.
        sender_email: Substring to match against SenderEmailAddress
                      (case-insensitive).
        keyword: Optional keyword to search for in Subject and Body.
                 If None, no keyword filtering is applied.
        require_attachments: If True, only yield emails that have at
                            least one attachment.

    Yields:
        Outlook MailItem COM objects that match all criteria.
    """
    return filter_emails(inbox, sender_email=sender_email, keyword=keyword, require_attachments=require_attachments)


def filter_emails_by_keyword(
    inbox,
    keyword: str,
    require_attachments: bool = True,
) -> Generator:
    """
    Yield emails containing a keyword in subject or body.

    This is a convenience wrapper around :func:`filter_emails`.

    Args:
        inbox: Outlook folder COM object (e.g. Inbox) to iterate.
        keyword: Keyword to search for in Subject and Body.
        require_attachments: If True, only yield emails that have at
                            least one attachment.

    Yields:
        Outlook MailItem COM objects that match the criteria.
    """
    return filter_emails(inbox, keyword=keyword, require_attachments=require_attachments)
