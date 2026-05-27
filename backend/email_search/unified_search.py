"""
Unified search functionality for email operations.

This module provides the main unified search function that coordinates
between server-side search implementations, including multi-folder search.
"""

# Standard library imports
import sys

# Type imports
from typing import Any, Dict, List, Optional, Tuple

# Local application imports
from ..config import search_config
from ..logging_config import get_logger
from ..outlook_session.session_manager import OutlookSessionManager
from ..validation import BatchProcessing
from ..validators import EmailSearchParams
from .search_common import (
    extract_email_info,
    get_folder_path_safe,
    is_server_search_supported,
)
from .server_search import (
    multi_folder_search,
    search_by_conversation_id,
    search_related_emails,
    server_side_search,
)

logger = get_logger(__name__)


def unified_search(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    folder_names: Optional[List[str]] = None,
    match_all: bool = True,
    search_type: str = "subject",
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Unified search function that prioritizes fast server-side search.

    Supports both single-folder (folder_name) and multi-folder (folder_names) search.
    When folder_names is provided, it takes precedence over folder_name.

    Args:
        search_term: The term to search for
        days: Number of days to look back
        folder_name: Single folder name (default: Inbox)
        folder_names: List of folder names for cross-folder search
        match_all: Whether to match all terms (AND logic)
        search_type: Type of search (subject, sender, recipient, body)

    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    if not search_term or not isinstance(search_term, str):
        return [], "Search term must be a non-empty string"

    if days < 1:
        return [], "Days must be at least 1"

    if days > search_config.MAX_SEARCH_DAYS:
        error_msg = f"Days must be between 1 and {search_config.MAX_SEARCH_DAYS}"
        print(f"Error: {error_msg}", file=sys.stderr)
        return [], error_msg

    try:
        # Resolve folders: folder_names takes precedence
        if folder_names and len(folder_names) > 0:
            # Multi-folder search — already returns email dicts
            if not is_server_search_supported(search_type):
                return [], (
                    f"Search type '{search_type}' is not supported "
                    f"for multi-folder search"
                )

            email_list = multi_folder_search(
                search_term, days, search_type, folder_names, match_all
            )
            folder_label = ", ".join(folder_names)

            if not email_list:
                return [], f"No emails found matching '{search_term}'"

            logger.info(
                f"Search completed, returning {len(email_list)} results"
            )
            message = f"Found {len(email_list)} emails in '{folder_label}'"
            return email_list, message
        else:
            # Single folder search — extract to dicts inside session
            # Default folder depends on search type:
            #   subject → Inbox + Sent Items (multi-folder)
            #   sender/body → Inbox, recipient → Sent Items
            if folder_name is None and search_type == "subject":
                from .server_search import DEFAULT_SEARCH_FOLDERS
                email_list = multi_folder_search(
                    search_term, days, search_type, DEFAULT_SEARCH_FOLDERS, match_all
                )
                folder_label = " + ".join(DEFAULT_SEARCH_FOLDERS)
                if not email_list:
                    return [], f"No emails found matching '{search_term}'"
                message = f"Found {len(email_list)} emails in '{folder_label}'"
                return email_list, message

            if folder_name is None and search_type == "recipient":
                folder_path = "Sent Items"
            else:
                folder_path = get_folder_path_safe(folder_name)

            with OutlookSessionManager() as session:
                folder = session.get_folder(folder_path)
                if not folder:
                    return [], f"Folder '{folder_path}' not found"

                if is_server_search_supported(search_type):
                    try:
                        results = server_side_search(
                            folder,
                            search_term,
                            days,
                            search_type,
                            match_all,
                            session.outlook_namespace,
                        )
                        if not results:
                            return [], (
                                f"No emails found in '{folder_path}' "
                                f"matching '{search_term}'"
                            )
                    except Exception as e:
                        logger.error(f"Server-side search failed: {e}")
                        return [], (
                            f"Search failed for '{search_term}' "
                            f"in '{folder_path}'"
                        )
                else:
                    logger.warning(
                        f"Search type '{search_type}' not supported "
                        f"by server-side search"
                    )
                    return [], (
                        f"Search type '{search_type}' is not supported "
                        f"for performance reasons"
                    )

                # Extract to dicts while COM session is alive
                email_list = []
                from .search_common import clear_com_attribute_cache, extract_email_info
                clear_com_attribute_cache()

                batch_size = BatchProcessing.DEFAULT_BATCH_SIZE
                total_results = len(results)

                for batch_start in range(0, total_results, batch_size):
                    batch_end = min(batch_start + batch_size, total_results)
                    batch_items = results[batch_start:batch_end]

                    for item in batch_items:
                        try:
                            email_data = extract_email_info(item)
                            if email_data:
                                email_list.append(email_data)
                        except Exception as e:
                            logger.warning(f"Failed to extract email info: {e}")
                            continue

                    if batch_start % 200 == 0:
                        clear_com_attribute_cache()

            folder_label = folder_path

        if not email_list:
            return [], f"No valid emails found"

        # Sort by received time (newest first)
        email_list.sort(
            key=lambda x: x.get("received_time", ""), reverse=True
        )

        logger.info(
            f"Search completed, returning {len(email_list)} results"
        )

        message = f"Found {len(email_list)} emails in '{folder_label}'"
        return email_list, message

    except Exception as e:
        error_msg = f"Error searching emails: {e}"
        logger.error(error_msg)
        return [], error_msg


def find_thread_by_email_id(
    email_id: str, folder_names: Optional[List[str]] = None
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Find all emails in the same conversation thread as the given email.

    Works across folders (Inbox + Sent Items by default) and finds
    all related emails even when subjects have changed.

    Args:
        email_id: EntryID of the reference email
        folder_names: Folders to search (default: Inbox + Sent Items)

    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    if not email_id:
        return [], "Email ID is required"

    if folder_names is None:
        folder_names = ["Inbox", "Sent Items"]

    try:
        with OutlookSessionManager() as session:
            # Get the reference email to find its conversation ID
            try:
                ref_item = session.outlook_namespace.GetItemFromID(email_id)
            except Exception as e:
                return [], f"Email not found: {email_id}"

            conv_id = getattr(ref_item, 'ConversationID', '')
            if not conv_id:
                return [], (
                    "This email has no conversation ID "
                    "(may be a standalone message)"
                )

            conv_id = str(conv_id)
            logger.info(
                f"Finding thread for conversation: {conv_id[:40]}..."
            )

        # Search for all emails with this conversation ID
        # search_by_conversation_id now returns dicts (extracted inside session)
        email_list = search_by_conversation_id(conv_id, folder_names)

        if not email_list:
            return [], "No other emails found in this thread"

        # Sort by received time (oldest first for thread view)
        email_list.sort(key=lambda x: x.get("received_time", ""))

        return email_list, (
            f"Found {len(email_list)} emails in conversation thread"
        )

    except Exception as e:
        error_msg = f"Error finding thread: {e}"
        logger.error(error_msg)
        return [], error_msg


def find_related_emails(
    email_id: str,
    days: int = 90,
    strategies: Optional[List[str]] = None,
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Find emails related to the given email using multiple strategies.

    Strategies (in order of confidence):
    - thread: Same conversation ID (confidence=1.0)
    - sender: Same sender within time window (confidence=0.7)
    - keyword: Shared keywords in subject (confidence=0.4)

    Args:
        email_id: EntryID of the reference email
        days: Lookback days for sender and keyword strategies
        strategies: Which strategies to use (default: all)

    Returns:
        Tuple of (list of email dicts sorted by relevance, status message)
    """
    if not email_id:
        return [], "Email ID is required"

    if days > search_config.MAX_SEARCH_DAYS:
        days = search_config.MAX_SEARCH_DAYS

    try:
        result = search_related_emails(email_id, days, strategies)
        combined = result.get("combined", [])

        if not combined:
            ref = result.get("reference_email", {})
            ref_subject = ref.get("subject", "Unknown")
            return [], (
                f"No related emails found for '{ref_subject}'"
            )

        ref_info = result.get("reference_email", {})
        thread_count = len(result.get("thread_results", []))
        sender_count = len(result.get("sender_results", []))
        keyword_count = len(result.get("keyword_results", []))

        strategy_parts = []
        if thread_count:
            strategy_parts.append(f"{thread_count} by thread")
        if sender_count:
            strategy_parts.append(f"{sender_count} by sender")
        if keyword_count:
            strategy_parts.append(f"{keyword_count} by keyword")

        ref_subject = ref_info.get("subject", "Unknown")
        return combined, (
            f"Found {len(combined)} related emails for '{ref_subject}' "
            f"({', '.join(strategy_parts)})"
        )

    except Exception as e:
        error_msg = f"Error finding related emails: {e}"
        logger.error(error_msg)
        return [], error_msg
