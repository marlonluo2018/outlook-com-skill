"""
Server-side search functionality for email operations.

This module provides functions for performing server-side searches using
Outlook's AdvancedSearch functionality, which is more efficient for large folders.
"""

# Standard library imports
import time
from datetime import datetime, timedelta, timezone
from typing import Any, List, Optional, Tuple

# Local application imports
from ..logging_config import get_logger
from ..outlook_session.session_manager import OutlookSessionManager
from .search_common import get_date_limit

logger = get_logger(__name__)

# Default folders to search when cross-folder is requested
DEFAULT_SEARCH_FOLDERS = ["Inbox", "Sent Items"]


def _build_search_criteria(search_term: str, days: int, search_type: str, match_all: bool) -> str:
    """Build SQL search criteria string for Outlook Restrict/AdvancedSearch."""
    date_limit = get_date_limit(days)
    escaped_search_term = search_term.replace("'", "''")
    sql_conditions = []

    sql_conditions.append(
        f"urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
    )

    if search_type == "subject":
        sql_conditions.append(
            f"urn:schemas:httpmail:subject LIKE '%{escaped_search_term}%'"
        )
    elif search_type == "sender":
        sql_conditions.append(
            f"(urn:schemas:httpmail:fromname LIKE '%{escaped_search_term}%' OR "
            f"urn:schemas:httpmail:senderemail LIKE '%{escaped_search_term}%' OR "
            f"urn:schemas:httpmail:fromemail LIKE '%{escaped_search_term}%')"
        )
    elif search_type == "recipient":
        sql_conditions.append(
            f"urn:schemas:httpmail:to LIKE '%{escaped_search_term}%'"
        )
    elif search_type == "body":
        sql_conditions.append(
            f"urn:schemas:httpmail:textdescription LIKE '%{escaped_search_term}%'"
        )
    elif search_type == "conversation":
        # Search by conversation ID - no date filter needed for thread
        sql_conditions = [
            f"http://schemas.microsoft.com/mapi/proptag/0x3013001F = '{escaped_search_term}'"
        ]

    return "@SQL=" + " AND ".join(sql_conditions)


def _search_single_folder(
    folder,
    search_criteria: str,
    max_results: int = 2000,
    namespace=None,
    search_type: Optional[str] = None,
) -> List[Any]:
    """Search a single folder using Items.Restrict, with manual fallback.

    When Restrict returns 0 results (often due to unreliable DAV property
    support), fall back to a date-filtered manual scan that checks the
    relevant Outlook fields for the requested search type.
    """
    folder_path = getattr(folder, 'FolderPath', str(folder))
    need_fallback = False
    use_restrict = search_type != "conversation"

    if use_restrict:
        try:
            items = folder.Items
            restricted_items = items.Restrict(search_criteria)
            results = list(restricted_items)
            logger.info(f"Restrict found {len(results)} results in '{folder.Name}'")
            if results:
                return results[:max_results]
            # 0 results — don't trust it; DAV properties may be unsupported
            need_fallback = True
        except Exception as e:
            logger.warning(f"Restrict failed for '{folder.Name}': {e}")
            need_fallback = True
    else:
        logger.info(
            f"Skipping Restrict for conversation search in '{folder.Name}' "
            f"and using manual scan directly"
        )
        need_fallback = True

    if not need_fallback:
        return []

    # Fallback: parse date + search terms from criteria and filter manually.
    # AdvancedSearch is unreliable (disconnected COM objects), so we iterate
    # items by index — the approach proven by find-recent.
    try:
        import re

        date_match = re.search(
            r"datereceived\s*>=\s*'(\d{4}-\d{2}-\d{2})'", search_criteria
        )
        if date_match:
            date_limit = datetime.strptime(
                date_match.group(1), '%Y-%m-%d'
            ).replace(tzinfo=timezone.utc)
        else:
            date_limit = datetime.now(timezone.utc) - timedelta(days=30)

        term_matches = re.findall(r"LIKE\s+'%([^']+)%'", search_criteria)
        search_terms = [t.lower() for t in term_matches]

        conversation_match = re.search(
            r"0x3013001F\s*=\s*'([^']+)'", search_criteria
        )
        conversation_term = (
            conversation_match.group(1).lower() if conversation_match else ""
        )

        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        total = getattr(items, 'Count', 0)
        results = []

        for i in range(min(total, 2000)):
            try:
                item = items.Item(i + 1)
                if not item:
                    continue
                if (
                    hasattr(item, 'ReceivedTime')
                    and item.ReceivedTime
                ):
                    item_time = item.ReceivedTime
                    if getattr(item_time, 'tzinfo', None) is None:
                        item_time = item_time.replace(tzinfo=timezone.utc)
                    if item_time < date_limit:
                        continue

                if hasattr(item, 'Class') and item.Class != 43:
                    continue

                if conversation_term:
                    item_conversation_id = str(
                        getattr(item, 'ConversationID', '') or ''
                    ).lower()
                    if item_conversation_id != conversation_term:
                        continue
                elif search_terms:
                    searchable_parts = []

                    if search_type == "subject":
                        searchable_parts.append(getattr(item, 'Subject', '') or '')
                    elif search_type == "sender":
                        searchable_parts.append(getattr(item, 'SenderName', '') or '')
                        searchable_parts.append(
                            getattr(item, 'SenderEmailAddress', '') or ''
                        )
                    elif search_type == "recipient":
                        searchable_parts.append(getattr(item, 'To', '') or '')
                        searchable_parts.append(getattr(item, 'CC', '') or '')
                        try:
                            recipients = getattr(item, 'Recipients', None)
                            if recipients:
                                for r_idx in range(1, recipients.Count + 1):
                                    recipient = recipients.Item(r_idx)
                                    searchable_parts.append(
                                        getattr(recipient, 'Name', '') or ''
                                    )
                                    searchable_parts.append(
                                        getattr(recipient, 'Address', '') or ''
                                    )
                        except Exception:
                            pass
                    elif search_type == "body":
                        searchable_parts.append(getattr(item, 'Body', '') or '')
                    else:
                        searchable_parts.append(getattr(item, 'Subject', '') or '')
                        searchable_parts.append(getattr(item, 'SenderName', '') or '')
                        searchable_parts.append(getattr(item, 'To', '') or '')
                        searchable_parts.append(getattr(item, 'CC', '') or '')

                    combined = ' '.join(searchable_parts).lower()
                    if not any(term in combined for term in search_terms):
                        continue

                results.append(item)
            except Exception:
                continue

        logger.info(
            f"Manual fallback found {len(results)} results in '{folder.Name}'"
        )
        return results[:max_results]
    except Exception as e2:
        logger.error(
            f"Manual fallback also failed for '{folder.Name}': {e2}"
        )

    return []


def _resolve_folders(session, folder_names: Optional[List[str]] = None) -> List[Any]:
    """Resolve folder names to Outlook folder objects."""
    if not folder_names:
        folder_names = ["Inbox"]

    folders = []
    for name in folder_names:
        try:
            folder = session.get_folder(name)
            if folder:
                folders.append(folder)
            else:
                logger.warning(f"Folder '{name}' not found, skipping")
        except Exception as e:
            logger.warning(f"Error resolving folder '{name}': {e}")

    return folders


def server_side_search(
    folder, search_term: str, days: int, search_type: str, match_all: bool, namespace=None
) -> List[Any]:
    """
    Perform server-side search using Outlook's Restrict method.

    Falls back to manual scanning when Outlook's indexed property search
    returns no results or the relevant DAV property is unreliable.
    """
    try:
        search_criteria = _build_search_criteria(
            search_term, days, search_type, match_all
        )

        logger.info(f"Server-side search criteria: {search_criteria}")

        folder_path = (
            folder.FolderPath if hasattr(folder, 'FolderPath') else str(folder)
        )
        logger.info(f"Folder path: {folder_path}")

        return _search_single_folder(
            folder,
            search_criteria,
            namespace=namespace,
            search_type=search_type,
        )

    except Exception as e:
        logger.error(f"Server-side search failed: {e}")
        logger.error(f"Error type: {type(e)}")
        logger.error(f"Error details: {str(e)}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return []


def multi_folder_search(
    search_term: str,
    days: int,
    search_type: str,
    folder_names: List[str],
    match_all: bool = True,
) -> List[dict]:
    """
    Search across multiple folders and return combined results as email dicts.

    COM items are extracted to dicts INSIDE the session context so they
    remain usable after the session closes. Each result is tagged with
    its source folder name. Sorted by received time (newest first).

    Args:
        search_term: The term to search for
        days: Number of days to look back
        search_type: Type of search (subject, sender, recipient, body)
        folder_names: List of folder names to search (e.g. ["Inbox", "Sent Items"])
        match_all: Whether to match all terms (AND logic)

    Returns:
        Combined list of email dicts from all folders
    """
    search_criteria = _build_search_criteria(search_term, days, search_type, match_all)
    logger.info(f"Multi-folder search: type={search_type}, folders={folder_names}")

    all_results = []
    seen_ids = set()

    with OutlookSessionManager() as session:
        from .search_common import extract_email_info
        namespace = getattr(session, 'outlook_namespace', None)
        for folder_name in folder_names:
            try:
                folder = session.get_folder(folder_name)
                if not folder:
                    logger.warning(f"Skipping missing folder: {folder_name}")
                    continue

                results = _search_single_folder(
                    folder,
                    search_criteria,
                    namespace=namespace,
                    search_type=search_type,
                )
                for item in results:
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in seen_ids:
                        seen_ids.add(entry_id)
                        # Extract to dict while COM is alive
                        email_data = extract_email_info(item)
                        if email_data:
                            all_results.append(email_data)

                logger.info(f"'{folder_name}': {len(results)} results")
            except Exception as e:
                logger.warning(f"Error searching '{folder_name}': {e}")
                continue

        # Sort by received time, newest first (inside session for COM safety)
        try:
            all_results.sort(
                key=lambda x: x.get("received_time", ""),
                reverse=True,
            )
        except Exception as e:
            logger.warning(f"Error sorting results: {e}")

    logger.info(f"Multi-folder search total: {len(all_results)} results")
    return all_results


def _search_by_conversation_id_in_session(
    session: OutlookSessionManager,
    conversation_id: str,
    folder_names: Optional[List[str]] = None,
) -> List[dict]:
    """Find all emails with the same conversation ID using an existing session."""
    if not conversation_id:
        return []

    if not folder_names:
        folder_names = ["Inbox", "Sent Items"]

    search_criteria = _build_search_criteria(
        conversation_id, days=365, search_type="conversation", match_all=True
    )
    logger.info(f"Thread search: conversation_id={conversation_id[:30]}...")

    all_results = []
    seen_ids = set()

    from .search_common import extract_email_info

    namespace = getattr(session, 'outlook_namespace', None)
    for folder_name in folder_names:
        try:
            folder = session.get_folder(folder_name)
            if not folder:
                continue

            results = _search_single_folder(
                folder,
                search_criteria,
                namespace=namespace,
                search_type="conversation",
            )
            for item in results:
                entry_id = getattr(item, 'EntryID', '')
                if entry_id and entry_id not in seen_ids:
                    seen_ids.add(entry_id)
                    email_data = extract_email_info(item)
                    if email_data:
                        all_results.append(email_data)
        except Exception as e:
            logger.warning(f"Error in thread search '{folder_name}': {e}")
            continue

    try:
        all_results.sort(
            key=lambda x: x.get("received_time", "")
        )
    except Exception:
        pass

    logger.info(f"Thread search: found {len(all_results)} emails in conversation")
    return all_results


def search_by_conversation_id(conversation_id: str, folder_names: Optional[List[str]] = None) -> List[dict]:
    """
    Find all emails sharing the same conversation ID across specified folders.

    Uses Outlook's ConversationID property to track email threads,
    even when subjects change (RE:/Fwd: prefixes, topic changes).
    Returns email dicts extracted inside the session context.

    Args:
        conversation_id: The conversation ID to search for
        folder_names: Folders to search (defaults to Inbox + Sent Items)

    Returns:
        List of email dicts in the same thread, sorted by time
    """
    with OutlookSessionManager() as session:
        return _search_by_conversation_id_in_session(
            session, conversation_id, folder_names
        )


def _extract_recipient_set(email_info: dict, exclude_address: str = "") -> set:
    """Extract all recipient addresses (lowercase) from to + cc, excluding the given address."""
    addresses = set()
    exclude_lower = exclude_address.lower().strip()
    for recip in email_info.get("to_recipients", []) + email_info.get("cc_recipients", []):
        addr = (recip.get("address") or "").lower().strip()
        if addr and addr != exclude_lower:
            addresses.add(addr)
    return addresses


def _count_recipient_overlap(candidate_info: dict, ref_recipients: set) -> int:
    """Count how many of the reference recipients appear in the candidate's to/cc/sender."""
    candidate_addrs = set()
    for recip in candidate_info.get("to_recipients", []) + candidate_info.get("cc_recipients", []):
        addr = (recip.get("address") or "").lower().strip()
        if addr:
            candidate_addrs.add(addr)
    return len(ref_recipients & candidate_addrs)


def search_related_emails(
    email_id: str,
    days: int = 90,
    strategies: Optional[List[str]] = None,
    exclude_thread: bool = False,
) -> dict:
    """
    Find emails related to a given email using multiple strategies.

    Strategies:
    - thread: Same conversation ID (highest confidence)
    - sender: Same sender within expanded time window
    - recipient: Shared recipients (people overlap)
    - keyword: Extract key terms and search subject/body

    Args:
        email_id: EntryID of the reference email
        days: Lookback window for sender, recipient, and keyword strategies
        strategies: List of strategies to use (default: all four)
        exclude_thread: If True, skip the thread strategy

    Returns:
        Dict with strategy results and combined ranked list
    """
    if strategies is None:
        strategies = ["thread", "sender", "recipient", "keyword"]

    if exclude_thread and "thread" in strategies:
        strategies = [s for s in strategies if s != "thread"]

    output = {
        "reference_email": None,
        "thread_results": [],
        "sender_results": [],
        "recipient_results": [],
        "keyword_results": [],
        "combined": [],
    }

    with OutlookSessionManager() as session:
        # Get the reference email
        try:
            ref_item = session.outlook_namespace.GetItemFromID(email_id)
        except Exception as e:
            logger.error(f"Reference email not found: {e}")
            return output

        from .search_common import extract_email_info

        ref_info = extract_email_info(ref_item)
        output["reference_email"] = ref_info

        seen_ids = {email_id}
        thread_results = []
        sender_results = []
        recipient_results = []
        keyword_results = []

        # Strategy 1: Thread (ConversationID)
        if "thread" in strategies:
            conv_id = ref_info.get("conversation_id", "")
            if conv_id:
                thread_items = _search_by_conversation_id_in_session(
                    session,
                    conv_id,
                    ["Inbox", "Sent Items"],
                )
                for item in thread_items:
                    entry_id = item.get('entry_id', '')
                    if entry_id and entry_id not in seen_ids:
                        seen_ids.add(entry_id)
                        # Already a dict from thread search
                        info = dict(item)
                        info["_confidence"] = 1.0
                        info["_strategy"] = "thread"
                        thread_results.append(info)
                logger.info(f"Strategy 'thread': found {len(thread_results)}")
            output["thread_results"] = thread_results

        # Strategy 2: Same sender + time window
        if "sender" in strategies:
            sender_name = ref_info.get("sender", "")
            if sender_name and sender_name != "Unknown":
                sender_keywords = _extract_search_keywords(ref_info)
                sender_strong_keywords, _ = _split_keywords_by_strength(
                    sender_keywords
                )
                sender_topic_keywords = sender_strong_keywords or sender_keywords[:1]

                sender_criteria = _build_search_criteria(
                    sender_name, days, "sender", match_all=True
                )
                namespace = getattr(session, 'outlook_namespace', None)
                for folder_name in ["Inbox", "Sent Items"]:
                    try:
                        folder = session.get_folder(folder_name)
                        if not folder:
                            continue
                        items = _search_single_folder(
                            folder,
                            sender_criteria,
                            max_results=100,
                            namespace=namespace,
                            search_type="sender",
                        )
                        for item in items:
                            eid = getattr(item, 'EntryID', '')
                            if eid and eid not in seen_ids:
                                info = extract_email_info(item)
                                if not info:
                                    continue

                                sender_overlap = _count_keyword_overlap(
                                    info, sender_topic_keywords
                                )
                                if sender_topic_keywords and sender_overlap < 1:
                                    continue

                                seen_ids.add(eid)
                                info["_confidence"] = min(
                                    0.60 + (0.05 * max(sender_overlap - 1, 0)),
                                    0.75,
                                )
                                info["_strategy"] = "sender"
                                info["_sender_keyword_overlap"] = sender_overlap
                                sender_results.append(info)
                    except Exception as e:
                        logger.warning(f"Sender strategy error in '{folder_name}': {e}")
                logger.info(f"Strategy 'sender': found {len(sender_results)}")
            output["sender_results"] = sender_results

        # Strategy 3: Recipient overlap (shared people)
        if "recipient" in strategies:
            try:
                current_user = session.outlook.Session.CurrentUser
                eu = current_user.AddressEntry.GetExchangeUser()
                my_email = (eu.PrimarySmtpAddress if eu else "").lower()
            except Exception:
                my_email = ""

            ref_recipients = _extract_recipient_set(ref_info, exclude_address=my_email)
            try:
                sender_addr = (getattr(ref_item, 'SenderEmailAddress', '') or '').lower().strip()
                if sender_addr and sender_addr != my_email and '@' in sender_addr:
                    ref_recipients.add(sender_addr)
            except Exception:
                pass

            if len(ref_recipients) >= 1:
                namespace = getattr(session, 'outlook_namespace', None)
                # Use display names for sender search (more reliable than addresses)
                ref_names = set()
                for recip in ref_info.get("to_recipients", []) + ref_info.get("cc_recipients", []):
                    name = (recip.get("name") or "").strip()
                    if name and name.lower() != my_email:
                        ref_names.add(name)
                search_names = list(ref_names)[:3]
                for name in search_names:
                    recip_criteria = _build_search_criteria(
                        name, days, "sender", match_all=True
                    )
                    for folder_name in ["Inbox", "Sent Items"]:
                        try:
                            folder = session.get_folder(folder_name)
                            if not folder:
                                continue
                            items = _search_single_folder(
                                folder,
                                recip_criteria,
                                max_results=50,
                                namespace=namespace,
                                search_type="sender",
                            )
                            for item in items:
                                eid = getattr(item, 'EntryID', '')
                                if eid and eid not in seen_ids:
                                    info = extract_email_info(item)
                                    if not info:
                                        continue
                                    overlap = _count_recipient_overlap(
                                        info, ref_recipients
                                    )
                                    if overlap < 2:
                                        continue
                                    seen_ids.add(eid)
                                    info["_confidence"] = min(
                                        0.55 + (0.05 * max(overlap - 2, 0)),
                                        0.70,
                                    )
                                    info["_strategy"] = "recipient"
                                    info["_recipient_overlap"] = overlap
                                    recipient_results.append(info)
                        except Exception as e:
                            logger.warning(f"Recipient strategy error in '{folder_name}': {e}")
                logger.info(f"Strategy 'recipient': found {len(recipient_results)}")
            output["recipient_results"] = recipient_results

        # Strategy 4: Keyword extraction + subject search
        if "keyword" in strategies:
            keywords = _extract_search_keywords(ref_info)
            if keywords:
                namespace = getattr(session, 'outlook_namespace', None)
                strong_keywords, weak_keywords = _split_keywords_by_strength(keywords)
                if not strong_keywords:
                    strong_keywords = keywords[:1]
                all_keywords = strong_keywords + [
                    kw for kw in weak_keywords if kw not in strong_keywords
                ]
                min_overlap = 2 if len(all_keywords) >= 2 else 1

                for kw in strong_keywords[:3]:  # Search only with strong keywords
                    kw_criteria = _build_search_criteria(
                        kw, days, "subject", match_all=True
                    )
                    for folder_name in ["Inbox", "Sent Items"]:
                        try:
                            folder = session.get_folder(folder_name)
                            if not folder:
                                continue
                            items = _search_single_folder(
                                folder,
                                kw_criteria,
                                max_results=50,
                                namespace=namespace,
                                search_type="subject",
                            )
                            for item in items:
                                eid = getattr(item, 'EntryID', '')
                                if eid and eid not in seen_ids:
                                    info = extract_email_info(item)
                                    if not info:
                                        continue

                                    strong_overlap = _count_keyword_overlap(
                                        info, strong_keywords
                                    )
                                    if strong_overlap < 1:
                                        continue

                                    overlap = _count_keyword_overlap(
                                        info, all_keywords
                                    )
                                    if overlap < min_overlap:
                                        continue

                                    seen_ids.add(eid)
                                    info["_confidence"] = min(
                                        0.45
                                        + (0.10 * max(strong_overlap - 1, 0))
                                        + (0.05 * max(overlap - strong_overlap, 0)),
                                        0.70,
                                    )
                                    info["_strategy"] = f"keyword:{kw}"
                                    info["_keyword_overlap"] = overlap
                                    info["_strong_keyword_overlap"] = strong_overlap
                                    keyword_results.append(info)
                        except Exception as e:
                            logger.warning(f"Keyword strategy error: {e}")
                logger.info(f"Strategy 'keyword': found {len(keyword_results)}")
            output["keyword_results"] = keyword_results

        # Combine and sort by confidence (desc), then by time (desc)
        all_combined = thread_results + sender_results + recipient_results + keyword_results
        try:
            all_combined.sort(
                key=lambda x: (
                    -x.get("_confidence", 0),
                    x.get("received_time", ""),
                ),
                reverse=False,
            )
            # For equal confidence, newest first
            # Actually let's sort: confidence desc, then time desc
            all_combined.sort(
                key=lambda x: (
                    -x.get("_confidence", 0),
                    x.get("received_time", ""),
                ),
                reverse=False,
            )
            # Re-sort: primary = confidence (high first), secondary = time (new first)
            from operator import itemgetter
            all_combined.sort(
                key=lambda x: x.get("received_time", ""),
                reverse=True,
            )
            all_combined.sort(
                key=lambda x: x.get("_confidence", 0),
                reverse=True,
            )
        except Exception:
            pass

        output["combined"] = all_combined

    return output


def _split_keywords_by_strength(keywords: List[str]) -> Tuple[List[str], List[str]]:
    """Separate strong topic keywords from weaker context keywords."""
    weak_keywords = {
        "philippines", "china", "india", "singapore", "malaysia",
        "indonesia", "thailand", "vietnam", "global", "regional",
        "apac", "asean", "team", "project", "program", "learning",
        "services", "foundation", "requester", "requesters",
    }

    strong = []
    weak = []

    for kw in keywords:
        normalized = (kw or "").strip().lower()
        if not normalized:
            continue

        if normalized in weak_keywords:
            weak.append(normalized)
            continue

        strong.append(normalized)

    return list(dict.fromkeys(strong)), list(dict.fromkeys(weak))


def _count_keyword_overlap(email_info: dict, keywords: List[str]) -> int:
    """Count how many reference keywords appear in the candidate email."""
    searchable_parts = [
        email_info.get("subject", "") or "",
        email_info.get("sender", "") or "",
        email_info.get("preview", "") or "",
    ]
    haystack = " ".join(searchable_parts).lower()
    return sum(1 for kw in keywords if kw.lower() in haystack)


def _extract_search_keywords(email_info: dict) -> List[str]:
    """Extract high-signal keywords from email subject for related search."""
    keywords = []
    subject = email_info.get("subject", "")

    # Remove common prefixes and bracketed tags like [EXTERNAL]
    cleaned = subject
    for prefix in ["RE:", "Re:", "FW:", "Fwd:", "FWD:"]:
        cleaned = cleaned.replace(prefix, "")

    import re

    cleaned = re.sub(r"\[[^\]]+\]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    # Drop generic words that create noisy global matches.
    common_words = {
        "the", "and", "for", "from", "this", "that", "with", "have",
        "your", "you", "are", "not", "was", "all", "can", "has",
        "about", "which", "will", "would", "been", "they", "their",
        "please", "thanks", "regards", "dear", "hello", "attached",
        "external", "training", "request", "update", "follow", "followup",
        "reminder", "quote", "quotation", "email", "mail", "team",
        "reply", "response", "responses", "urgent", "proposal", "invoice",
        "foundation", "course", "session", "private", "cohort", "virtual",
        "instructor", "instructed", "learnquest", "learn", "quest", "ibm",
    }

    words = re.findall(r"[A-Za-z0-9][A-Za-z0-9_\-\.#]*", cleaned)
    for word in words:
        normalized = word.strip("._-#").lower()
        if not normalized:
            continue
        if normalized in common_words:
            continue
        if normalized.isdigit() and len(normalized) <= 4:
            # Bare years like 2026 are usually too broad as keyword seeds.
            continue
        if len(normalized) <= 2 and not normalized.isupper():
            continue
        keywords.append(normalized)

    # Prefer more specific keywords first.
    keywords = list(dict.fromkeys(keywords))
    keywords.sort(
        key=lambda kw: (
            any(ch.isdigit() for ch in kw),
            len(kw),
        ),
        reverse=True,
    )

    return keywords[:10]
