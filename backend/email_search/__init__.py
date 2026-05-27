#!/usr/bin/env python3
"""Email search functionality for Outlook Skill.

This package provides various email search capabilities including:
|- Subject-based searches
|- Sender-based searches  
|- Recipient-based searches
|- Body content searches
|- Shared search utilities
"""

# Import from the modular search components
from .subject_search import search_email_by_subject
from .sender_search import search_email_by_sender
from .recipient_search import search_email_by_recipient
from .body_search import search_email_by_body

# Import from email listing module
from .email_listing import list_recent_emails, list_recent_emails_multi, get_emails_from_folder

# Import from outlook_session modules (moved from search_utils)
from ..outlook_session.folder_operations import get_folder_emails, list_folders
# Email operations simplified - most operations now in tools/email_operations.py

# Direct imports for search functions (previously wrapped in search_utils)
from .sender_search import search_email_by_sender as search_email_by_from
from .recipient_search import search_email_by_recipient as search_email_by_to

# Import shared utilities
from .search_common import get_folder_path_safe, get_date_limit, is_server_search_supported, extract_email_info, extract_email_info_minimal, clear_com_attribute_cache

# Import parallel extraction
from .parallel_extractor import extract_emails_optimized

# Import unified search
from .unified_search import (
    find_related_emails,
    find_thread_by_email_id,
    unified_search,
)

# Import server search
from .server_search import (
    multi_folder_search,
    search_by_conversation_id,
    server_side_search,
)

__all__ = [
    # Search functions from modular components
    "search_email_by_subject",
    "search_email_by_sender",
    "search_email_by_recipient",
    "search_email_by_body",

    # Email listing functions
    "list_folders",
    "list_recent_emails",
    "list_recent_emails_multi",
    "get_emails_from_folder",
    "get_folder_emails",

    # Search functions from modular components
    "search_email_by_from",
    "search_email_by_to",

    # Shared utilities
    "get_folder_path_safe",
    "get_date_limit",
    "is_server_search_supported",
    "extract_email_info",
    "extract_email_info_minimal",
    "clear_com_attribute_cache",
    "extract_emails_optimized",

    # Search implementations
    "unified_search",
    "server_side_search",
    "multi_folder_search",
    "search_by_conversation_id",

    # Thread and related email search
    "find_thread_by_email_id",
    "find_related_emails",
]