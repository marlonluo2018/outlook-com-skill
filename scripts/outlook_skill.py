"""
Outlook Skill CLI for BrainClaw
Command-line interface for Outlook email operations using email IDs
"""

import sys
import os
import re
import json
import argparse
from typing import Optional

# Force UTF-8 for stdin, stdout and stderr on Windows
if sys.stdin and hasattr(sys.stdin, 'reconfigure') and sys.stdin.encoding != 'utf-8':
    sys.stdin.reconfigure(encoding='utf-8')
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

# Add parent directory to path to allow imports from backend and tools
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

# Import from local backend and tools directories
from backend.outlook_session.session_manager import OutlookSessionManager
from backend.email_search import (
    list_recent_emails,
    list_recent_emails_multi,
    search_email_by_subject,
    search_email_by_from,
    search_email_by_to,
    search_email_by_body,
    list_folders,
    find_thread_by_email_id,
    find_related_emails,
    unified_search,
)
from backend.email_composition import compose_email
from backend.outlook_session.contact_operations import get_contact_by_email, get_contact_by_name, get_display_name_from_email
from backend.config import search_config, display_config


def cmd_list_folders(args):
    """List all Outlook folders"""
    try:
        folders = list_folders(hide_system_folders=args.hide_system)
        print("\nAvailable folders:")
        for folder in folders:
            print(folder)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_list_recent(args):
    """List recent emails with their IDs"""
    try:
        if args.folder:
            emails, message = list_recent_emails(args.folder, args.days)
        else:
            emails, message = list_recent_emails_multi(days=args.days)
        email_count = len(emails)

        if getattr(args, 'json', False):
            serializable = []
            for e in emails:
                entry = {
                    'entry_id': e.get('entry_id') or e.get('id', ''),
                    'subject': e.get('subject', ''),
                    'sender': e.get('sender', ''),
                    'to_recipients': e.get('to_recipients', []),
                    'cc_recipients': e.get('cc_recipients', []),
                    'received_time': e.get('received_time', ''),
                    'folder': e.get('folder_name', e.get('folder', '')),
                    'has_attachments': e.get('has_attachments', False),
                    'meeting_status': e.get('meeting_status', ''),
                    'attachments_count': e.get('attachments_count', 0),
                    'body_preview': e.get('body_preview', ''),
                }
                serializable.append(entry)
            print(json.dumps(serializable, ensure_ascii=False))
            return 0

        print(f"\n✅ Found {email_count} recent emails\n")

        if emails:
            _display_email_list(emails, show_folder=True)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def _folder_emoji(folder_name: str) -> str:
    """Return emoji indicator for folder type."""
    name_lower = folder_name.lower()
    if "sent" in name_lower:
        return "\U0001F4E4"  # 📤
    elif "inbox" in name_lower:
        return "\U0001F4E5"  # 📥
    elif "draft" in name_lower:
        return "\U0001F4DD"  # 📝
    elif "deleted" in name_lower or "trash" in name_lower:
        return "\U0001F5D1"  # 🗑
    return "\U0001F4C1"  # 📁


def _normalize_search_days(days: Optional[int]) -> int:
    """Normalize direct find search days while preserving broad-search capability."""
    if days is None:
        return search_config.DIRECT_FIND_DEFAULT_DAYS
    if days < 1:
        return 1
    return min(days, search_config.MAX_SEARCH_DAYS)


def _build_body_preview(body_text: str) -> str:
    """Build a compact terminal-safe preview from an email body."""
    if not body_text:
        return ""

    stop_markers = (
        'from:',
        'sent:',
        'subject:',
        'to:',
        'cc:',
        'original message',
        '-----original message-----',
        'zjqcmqryfpfptbannerstart',
        'zjqcmqryfpfptbannerend',
        'notice:this is an external sender',
        'this message is from an external sender',
    )

    preview_lines = []
    preview_budget = max(getattr(display_config, 'PREVIEW_LENGTH', 200) * 2, 200)

    for raw_line in body_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        lower = line.lower()
        if any(marker in lower for marker in stop_markers):
            break
        if re.fullmatch(r'[_=\-\s]{5,}', line):
            break
        if line.startswith('<http') or line.startswith('http://') or line.startswith('https://'):
            continue
        if 'proofpoint.com' in lower or line == 'Report Suspicious':
            continue

        line = re.sub(r'\s+', ' ', line)
        if not line:
            continue

        preview_lines.append(line)

        if len(' '.join(preview_lines)) >= preview_budget or len(preview_lines) >= 4:
            break

    preview = re.sub(r'\s+', ' ', ' '.join(preview_lines)).strip()
    max_preview_length = getattr(display_config, 'PREVIEW_LENGTH', 200)
    if len(preview) > max_preview_length:
        preview = preview[:max_preview_length].rstrip() + "..."
    return preview


def _display_email_list(emails, show_folder=True):
    """Display formatted email list with folder markers."""
    def extract_display_name(recipient_string):
        if not recipient_string:
            return ""
        if '</o=' in recipient_string.lower() or '/cn=' in recipient_string.lower():
            if '<' in recipient_string:
                return recipient_string.split('<')[0].strip()
        if '<' in recipient_string and '@' in recipient_string:
            return recipient_string.split('<')[0].strip()
        return recipient_string.strip()

    for idx, email_data in enumerate(emails, 1):
        email_id = email_data.get('id') or email_data.get('entry_id', '')
        print(f"{'='*80}")
        folder_name = email_data.get('folder', '')
        folder_indicator = f" {_folder_emoji(folder_name)} {folder_name}" if show_folder and folder_name else ""

        meeting = email_data.get('meeting_status', '')
        meeting_icon = ""
        if meeting == "meeting_request":
            meeting_icon = " 📅 Meeting Invite"
        elif meeting == "meeting_canceled":
            meeting_icon = " ❌ Canceled"
        elif meeting == "meeting":
            meeting_icon = " 📅 Meeting"
        elif not meeting:
            # Fast check: subject + sender heuristics (no body access)
            subj_lower = (email_data.get('subject') or '').lower()
            sender_lower = (email_data.get('sender') or '').lower()
            event_sender_kw = ('events', 'calendar', 'webinar', 'noreply')
            event_subj_kw = ('webinar', 'join us', 'register now', 'you are invited',
                             "you're invited", 'invitation', 'save the date',
                             'live event', 'virtual event')
            if any(kw in sender_lower for kw in event_sender_kw) or \
               any(kw in subj_lower for kw in event_subj_kw):
                meeting_icon = " 📅 Event"

        print(f"Email #{idx}{folder_indicator}{meeting_icon}")
        print(f"{'='*80}")

        subject = email_data.get('subject', 'No Subject')
        sender = email_data.get('sender', 'Unknown')
        received = email_data.get('received_time', 'Unknown')

        print(f"ID: {email_id}")
        print(f"Subject: {subject}")
        print(f"From: {sender}")

        to_recipients = email_data.get('to_recipients', [])
        if to_recipients:
            to_list = []
            for recipient in to_recipients:
                name = recipient.get('name', '')
                address = recipient.get('address', '')
                display_name = extract_display_name(name) if name else extract_display_name(address)
                if display_name:
                    to_list.append(display_name)
            if to_list:
                print(f"To: {'; '.join(to_list)}")

        cc_recipients = email_data.get('cc_recipients', [])
        if cc_recipients:
            cc_list = []
            for recipient in cc_recipients:
                name = recipient.get('name', '')
                address = recipient.get('address', '')
                display_name = extract_display_name(name) if name else extract_display_name(address)
                if display_name:
                    cc_list.append(display_name)
            if cc_list:
                print(f"CC: {'; '.join(cc_list)}")

        print(f"Received: {received}")

        # Show confidence/strategy for related search
        confidence = email_data.get('_confidence')
        strategy = email_data.get('_strategy')
        if confidence is not None and strategy:
            stars = "★" * int(confidence * 5) + "☆" * (5 - int(confidence * 5))
            print(f"Relevance: {stars} ({strategy})")

        has_attachments = email_data.get('has_attachments', False)
        attachments = email_data.get('attachments', [])
        attachments_count = email_data.get('attachments_count', 0)
        embedded_images_count = email_data.get('embedded_images_count', 0)

        if has_attachments and attachments:
            print(f"\n📎 Attachments ({attachments_count}):")
            for attachment in attachments:
                name = attachment.get('name', 'Unknown')
                size = attachment.get('size', 0)
                size_kb = size / 1024 if size > 0 else 0
                print(f"  - {name} ({size_kb:.1f} KB)")

        if embedded_images_count > 0:
            embedded_images = email_data.get('embedded_images', [])
            if embedded_images:
                names = ", ".join(img.get('name', 'unknown') for img in embedded_images)
                print(f"\U0001F5BC  Embedded images ({embedded_images_count}): {names}")
            else:
                print(f"\U0001F5BC  Embedded images: {embedded_images_count}")

        try:
            with OutlookSessionManager() as session:
                email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(email_id)
                body_text = email_item.Body if hasattr(email_item, 'Body') else ""
                preview = _build_body_preview(body_text)
                if preview:
                    print(f"\nPreview: {preview}")
        except Exception:
            pass

        print()


def _print_thread_summary(emails):
    """Print a one-line thread summary: count, participants, date span."""
    if not emails:
        return
    senders = list(dict.fromkeys(e.get('sender', 'Unknown') for e in emails))
    times = [e.get('received_time', '') for e in emails if e.get('received_time')]
    times.sort()

    participant_str = ", ".join(senders[:3])
    if len(senders) > 3:
        participant_str += f" +{len(senders) - 3} more"

    span_str = ""
    if len(times) >= 2:
        start = times[0][:10]
        end = times[-1][:10]
        if start != end:
            try:
                from datetime import datetime
                d1 = datetime.strptime(start, "%Y-%m-%d")
                d2 = datetime.strptime(end, "%Y-%m-%d")
                span_days = (d2 - d1).days
                span_str = f" | Span: {start} → {end} ({span_days} days)"
            except Exception:
                span_str = f" | Span: {start} → {end}"
        else:
            span_str = f" | Date: {start}"
    elif times:
        span_str = f" | Date: {times[0][:10]}"

    print(f"\U0001F4CA Thread: {len(emails)} messages | Participants: {participant_str}{span_str}")
    print()


def _display_email_list_brief(emails, show_folder=True):
    """Compact single-line display with email ID on second line."""
    for idx, email_data in enumerate(emails, 1):
        email_id = email_data.get('id') or email_data.get('entry_id', '')
        folder_name = email_data.get('folder', '')
        folder_tag = f" {_folder_emoji(folder_name)}" if show_folder and folder_name else ""

        received = (email_data.get('received_time', '') or '')[:10]
        sender = (email_data.get('sender', 'Unknown') or 'Unknown')[:20]
        subject = (email_data.get('subject', 'No Subject') or 'No Subject')[:50]

        confidence = email_data.get('_confidence')
        strategy = email_data.get('_strategy', '')
        stars_str = ""
        if confidence is not None:
            stars = "★" * int(confidence * 5) + "☆" * (5 - int(confidence * 5))
            stars_str = f" {stars}"

        print(f"# {idx}{folder_tag} {received} | {sender:<20} | {subject}{stars_str}")
        print(f"  ID: {email_id}")


def cmd_search(args):
    """Search emails and display with IDs - supports multi-folder."""
    try:
        effective_days = _normalize_search_days(args.days)
        if args.days != effective_days:
            print(
                f"\nℹ️ Direct find search window adjusted to {effective_days} days "
                f"(allowed range: 1-{search_config.MAX_SEARCH_DAYS})."
            )

        # Resolve folders for multi-folder search
        if args.folders:
            folder_names = [f.strip() for f in args.folders.split(',')]
        else:
            folder_names = None

        emails, note = unified_search(
            search_term=args.query,
            days=effective_days,
            folder_name=args.folder,
            folder_names=folder_names,
            match_all=args.match_all,
            search_type=args.type,
        )
        print(f"\n✅ {note}\n")

        if emails:
            show_folder = (folder_names is not None and len(folder_names) > 1) or \
                          (folder_names is None and args.folder is None and args.type == "subject")
            _display_email_list(emails, show_folder=show_folder)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_get_email(args):
    """Get full email details by ID"""
    try:
        email_id = args.email_id or getattr(args, 'email_id_flag', None)
        if not email_id:
            print("Error: email_id is required (positional or --id)", file=sys.stderr)
            return 1
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, email_id)

            fields = None
            if getattr(args, 'fields', None):
                fields = set(f.strip().lower() for f in args.fields.split(','))

            if not fields:
                print("\nFull email details:")
                print(f"ID: {email_id}")
            if not fields or 'subject' in fields:
                print(f"Subject: {email_item.Subject}")
            if not fields or 'from' in fields:
                print(f"From: {email_item.SenderName} <{email_item.SenderEmailAddress}>")
            if not fields or 'to' in fields:
                if email_item.To:
                    print(f"To: {email_item.To}")
            if not fields or 'cc' in fields:
                if email_item.CC:
                    print(f"CC: {email_item.CC}")
            if not fields or 'date' in fields:
                rt = getattr(email_item, 'ReceivedTime', None)
                print(f"Date: {rt.replace(tzinfo=None) if rt else 'Unknown'}")
            if not fields or 'body' in fields:
                print(f"\nBody:\n{email_item.Body}")

            if (not fields or 'body' in fields) and email_item.Attachments.Count > 0:
                regular = []
                embedded = []
                for i in range(1, email_item.Attachments.Count + 1):
                    attach = email_item.Attachments.Item(i)
                    if _is_embedded_image(attach):
                        embedded.append(attach)
                    else:
                        regular.append(attach)

                if regular:
                    print(f"\n\U0001F4CE Attachments ({len(regular)}):")
                    for attach in regular:
                        size_kb = attach.Size / 1024
                        print(f"  - {attach.FileName} ({size_kb:.1f} KB)")

                if embedded:
                    import tempfile
                    temp_dir = os.path.join(tempfile.gettempdir(), "outlook_inline", email_id[:16])
                    os.makedirs(temp_dir, exist_ok=True)
                    print(f"\n\U0001F5BC  Embedded images (auto-saved):")
                    used_names = set()
                    for attach in embedded:
                        filename = attach.FileName
                        if filename in used_names:
                            stem, ext = os.path.splitext(filename)
                            counter = 2
                            while f"{stem}_{counter}{ext}" in used_names:
                                counter += 1
                            filename = f"{stem}_{counter}{ext}"
                        used_names.add(filename)
                        save_path = os.path.join(temp_dir, filename)
                        attach.SaveAsFile(save_path)
                        size_kb = attach.Size / 1024
                        print(f"  - {save_path} ({size_kb:.1f} KB)")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def _get_email_item(session, email_id):
    """Get email by ID, retrying across all Outlook stores when needed."""
    namespace = session.namespace or session.outlook.GetNamespace("MAPI")

    try:
        return namespace.GetItemFromID(email_id)
    except Exception as first_error:
        last_error = first_error

        try:
            stores = getattr(namespace, "Folders", None)
            if stores:
                for i in range(1, stores.Count + 1):
                    try:
                        store_root = stores.Item(i)
                        store_id = getattr(store_root, "StoreID", "")
                        if not store_id:
                            continue
                        return namespace.GetItemFromID(email_id, store_id)
                    except Exception as store_error:
                        last_error = store_error
                        continue
        except Exception:
            pass

        if "moved or deleted" in str(last_error).lower():
            print("Error: Outlook could not resolve the email by its current item handle.")
            print("This can happen after recall, move, or mailbox-store changes.")
            print("Please search for the email again to get a fresh current email ID.")

        raise last_error


def _add_attachments(mail_item, attach_str):
    """Add file attachments to a mail item."""
    if not attach_str:
        return
    import os
    for filepath in attach_str.split(","):
        filepath = filepath.strip().strip('"')
        if not os.path.exists(filepath):
            print(f"WARNING: Attachment not found: {filepath}")
            continue
        mail_item.Attachments.Add(filepath)


def _attach_emails(mail_item, attach_email_str, session):
    """Attach other emails as .msg files to a mail item."""
    if not attach_email_str:
        return
    import os, tempfile
    temp_files = []
    for eid in attach_email_str.split(","):
        eid = eid.strip().strip('"')
        if not eid:
            continue
        try:
            source = _get_email_item(session, eid)
            subject = str(getattr(source, "Subject", "email"))[:50]
            safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in subject)
            temp_path = os.path.join(tempfile.gettempdir(), f"{safe_name}.msg")
            source.SaveAs(temp_path, 3)  # olMSG
            mail_item.Attachments.Add(temp_path)
            temp_files.append(temp_path)
        except Exception as e:
            print(f"WARNING: Could not attach email {eid[:20]}...: {e}")
    for f in temp_files:
        try:
            os.remove(f)
        except OSError:
            pass


def _set_importance(mail_item, importance_str):
    """Set email importance/priority flag. Values: high, low, normal (default)."""
    if not importance_str:
        return
    level = importance_str.strip().lower()
    if level == "high":
        mail_item.Importance = 2  # olImportanceHigh
    elif level == "low":
        mail_item.Importance = 0  # olImportanceLow


def _add_inline_images(mail_item, inline_str):
    """Embed images into the HTML body using CID.

    Format: "filepath:cid_name" (comma separated for multiple).
    Windows paths (C:\...) are handled — only the last colon past position 1
    is treated as the filepath:cid separator.

    If the body already contains cid: references for the images, only sets the
    CID property (user controls placement). Otherwise auto-prepends <img> tags.
    """
    if not inline_str:
        return
    import os
    cids = []
    for entry in inline_str.split(","):
        entry = entry.strip().strip('"')
        last_colon = entry.rfind(":")
        if last_colon > 1:
            filepath = entry[:last_colon]
            cid = entry[last_colon + 1:]
        else:
            filepath = entry
            cid = None
        if not os.path.exists(filepath):
            print(f"WARNING: Inline image not found: {filepath}")
            continue
        if not cid:
            cid = os.path.basename(filepath).replace(" ", "_")
        # Remove existing attachments with the same CID (e.g. from forwarded emails)
        for i in range(mail_item.Attachments.Count, 0, -1):
            try:
                existing_cid = mail_item.Attachments.Item(i).PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                )
                if existing_cid == cid:
                    mail_item.Attachments.Remove(i)
            except Exception:
                pass
        attachment = mail_item.Attachments.Add(filepath)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
        )
        cids.append(cid)
    if cids and not any(f"cid:{c}" in mail_item.HTMLBody for c in cids):
        img_html = "".join(f'<img src="cid:{c}" style="max-width:100%;"><br>' for c in cids)
        mail_item.HTMLBody = img_html + mail_item.HTMLBody
    elif cids:
        mail_item.HTMLBody = mail_item.HTMLBody


IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico', '.webp'}
DOCUMENT_EXTENSIONS = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.zip', '.rar'}


def _is_embedded_image(attachment):
    """Detect if an attachment is an embedded/inline image (4-method detection)."""
    file_name = getattr(attachment, 'FileName', '') or getattr(attachment, 'DisplayName', 'Unknown')
    ext = os.path.splitext(file_name)[1].lower()

    if ext in DOCUMENT_EXTENSIONS:
        return False
    if ext not in IMAGE_EXTENSIONS:
        return False

    # Method 1: Content-ID / Content-Location
    try:
        if hasattr(attachment, 'PropertyAccessor'):
            content_id = attachment.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F")
            if content_id and str(content_id).strip():
                return True
            content_loc = attachment.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3713001F")
            if content_loc and str(content_loc).strip():
                return True
    except Exception:
        pass

    # Method 2: Attachment type (3=Embedded, 4=OLE)
    att_type = getattr(attachment, 'Type', 1)
    if att_type in (3, 4):
        return True

    # Method 3: Filename heuristics
    lower_name = file_name.lower()
    if any(p in lower_name for p in ('image', 'img', 'cid:', 'embedded')):
        return True
    stem = os.path.splitext(lower_name)[0]
    if stem.isdigit() or (len(stem) <= 2 and stem.isalnum()):
        return True

    # Method 4: Small image size (< 10 KB)
    size = getattr(attachment, 'Size', 0)
    if 0 < size < 10240:
        return True

    return False


_FONT_STYLE = 'font-family:Calibri,sans-serif;font-size:11pt;'


def _wrap_body_font(html_body: str) -> str:
    """Wrap HTML body content in a div with consistent font styling."""
    if not html_body:
        return ""
    return f'<div style="{_FONT_STYLE}">{html_body}</div>'


def _format_forward_message_html(message_text: str) -> str:
    """Convert plain text or HTML-ish input into the simple prepended block used for forwards.
    If input already contains HTML tags, pass through as-is (no <p> wrapping)."""
    if not message_text:
        return ""
    if re.search(r'<[a-zA-Z][^>]*>', message_text):
        return _wrap_body_font(message_text)
    inner = '<p>' + message_text.replace('\n\n', '</p><p>').replace('\n', '<br>') + '</p>'
    return _wrap_body_font(inner)


def _ensure_utf8_charset(html: str) -> str:
    """Ensure HTML declares UTF-8 charset. Replace any existing charset or add one."""
    charset_pattern = re.compile(
        r'<meta[^>]*charset=["\']?[^"\'>]+["\']?[^>]*/?>',
        re.IGNORECASE
    )
    content_type_pattern = re.compile(
        r'<meta[^>]*http-equiv=["\']?Content-Type["\']?[^>]*/?>',
        re.IGNORECASE
    )
    utf8_meta = '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'

    if charset_pattern.search(html):
        html = charset_pattern.sub(utf8_meta, html, count=1)
    elif content_type_pattern.search(html):
        html = content_type_pattern.sub(utf8_meta, html, count=1)
    else:
        head_match = re.search(r'(<head[^>]*>)', html, re.IGNORECASE)
        if head_match:
            insert_pos = head_match.end()
            html = html[:insert_pos] + utf8_meta + html[insert_pos:]
    return html


def _inject_before_body(prepend_html: str, full_html: str) -> str:
    """Insert prepend_html inside the <body> tag and ensure UTF-8 charset."""
    full_html = _ensure_utf8_charset(full_html)
    body_match = re.search(r'(<body[^>]*>)', full_html, re.IGNORECASE)
    if body_match:
        insert_pos = body_match.end()
        return full_html[:insert_pos] + prepend_html + full_html[insert_pos:]
    return prepend_html + full_html


def _get_smtp_address(recipient):
    """Extract SMTP email address from an Outlook Recipient object."""
    try:
        if recipient.AddressEntry.Type == "EX":
            exchange_user = recipient.AddressEntry.GetExchangeUser()
            if exchange_user:
                return exchange_user.PrimarySmtpAddress
    except:
        pass
    return recipient.Address or recipient.Name


def _remove_self_from_recipients(reply, current_user_email):
    """Remove current user from reply recipients."""
    to_remove = []
    for i in range(1, reply.Recipients.Count + 1):
        recipient = reply.Recipients.Item(i)
        recipient_email = _get_smtp_address(recipient)
        if recipient_email and recipient_email.lower() == current_user_email.lower():
            to_remove.append(i)
    for i in reversed(to_remove):
        reply.Recipients.Remove(i)


def _normalize_cc(cc_arg):
    """Normalize args.cc (list from append or single string) into one comma-separated string."""
    if not cc_arg:
        return ""
    if isinstance(cc_arg, list):
        return ",".join(cc_arg)
    return cc_arg


def _add_recipients(reply, to_str, cc_arg):
    """Append --to and --cc to existing recipients."""
    if to_str:
        for r in to_str.replace(";", ",").split(","):
            r = r.strip()
            if r:
                reply.Recipients.Add(r)
    cc_str = _normalize_cc(cc_arg)
    if cc_str:
        for r in cc_str.replace(";", ",").split(","):
            r = r.strip()
            if r:
                cc_recip = reply.Recipients.Add(r)
                cc_recip.Type = 2


def _build_reply_header(email_item, current_user_email):
    """Build standard reply attribution header (From/Sent/To/Cc/Subject)."""
    sent_on = email_item.SentOn.strftime("%A, %B %d, %Y %I:%M %p") if email_item.SentOn else ""
    original_to = getattr(email_item, "To", "") or ""
    original_cc = getattr(email_item, "CC", "") or ""
    original_subject = getattr(email_item, "Subject", "") or ""
    header = (
        '<div style="border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0 0 0">'
        f'<p><b>From:</b> {email_item.SenderName} &lt;{current_user_email}&gt;<br>'
        f'<b>Sent:</b> {sent_on}<br>'
        f'<b>To:</b> {original_to}<br>'
    )
    if original_cc:
        header += f'<b>Cc:</b> {original_cc}<br>'
    header += f'<b>Subject:</b> {original_subject}</p></div>'
    return header


def _print_sent_entry_id(session, subject):
    """Retrieve and print EntryID of the most recently sent email."""
    import time
    from datetime import datetime, timedelta, timezone
    threshold = datetime.now(timezone.utc) - timedelta(seconds=90)
    for attempt in range(15):
        time.sleep(2)
        try:
            ns = session.outlook.GetNamespace("MAPI")
            sent_folder = ns.GetDefaultFolder(5)  # 5 = olFolderSentMail
            items = sent_folder.Items
            items.Sort("[SentOn]", True)
            item = items.GetFirst()
            if not item:
                continue
            sent_on = item.SentOn
            if hasattr(sent_on, 'astimezone'):
                sent_utc = sent_on.astimezone(timezone.utc)
            else:
                sent_utc = sent_on.replace(tzinfo=timezone.utc)
            import re
            bare_subject = re.sub(r'^(RE:\s*|FW:\s*)+', '', subject or '', flags=re.IGNORECASE)
            bare_item_subject = re.sub(r'^(RE:\s*|FW:\s*)+', '', item.Subject or '', flags=re.IGNORECASE)
            if sent_utc >= threshold and (not subject or bare_subject in bare_item_subject):
                print(f"EntryID: {item.EntryID}")
                return
        except Exception:
            pass


def _resolve_body(args, required=True):
    """Resolve body from --body-stdin or positional/flag body arg.
    Priority: stdin > positional arg.
    Returns True on success (args.body is set), False on error."""
    if getattr(args, 'body_stdin', False):
        if not sys.stdin.isatty():
            args.body = sys.stdin.read()
        else:
            print("Error: --body-stdin specified but no piped input detected.", file=sys.stderr)
            return False
    if required and not args.body:
        print("Error: body or --body-stdin is required.", file=sys.stderr)
        return False
    return True


def cmd_reply(args):
    """Reply to an email. Default: reply-all. --only: sender (From) only."""
    if not _resolve_body(args):
        return 1
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)
            current_user = session.outlook.Session.CurrentUser
            current_user_email = (
                current_user.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                if current_user.AddressEntry.GetExchangeUser() else ""
            )

            parent_folder = email_item.Parent
            is_sent_items = "Sent Items" in parent_folder.Name or "已发送邮件" in parent_folder.Name

            if args.only:
                # --only: reply to From (sender) only
                if is_sent_items:
                    reply = session.outlook.CreateItem(0)
                    _add_recipients(reply, args.to, args.cc)
                    reply.Display(False)
                    signature_html = reply.HTMLBody
                    original_body = email_item.HTMLBody if email_item.HTMLBody else f"<p>{email_item.Body}</p>"
                    reply.HTMLBody = _wrap_body_font(args.body) + signature_html + _build_reply_header(email_item, current_user_email) + original_body
                else:
                    reply = email_item.Reply()
                    _add_recipients(reply, args.to, args.cc)
                    reply.HTMLBody = _wrap_body_font(args.body) + reply.HTMLBody
            elif is_sent_items:
                # Sent Items: create new email with original recipients
                reply = session.outlook.CreateItem(0)
                for i in range(1, email_item.Recipients.Count + 1):
                    recip = email_item.Recipients.Item(i)
                    smtp = _get_smtp_address(recip)
                    if smtp and smtp.lower() != current_user_email.lower():
                        new_recip = reply.Recipients.Add(smtp)
                        new_recip.Type = recip.Type  # 1=To, 2=CC
                _add_recipients(reply, args.to, args.cc)
                reply.Display(False)
                signature_html = reply.HTMLBody
                original_body = email_item.HTMLBody if email_item.HTMLBody else f"<p>{email_item.Body}</p>"
                reply.HTMLBody = _wrap_body_font(args.body) + signature_html + _build_reply_header(email_item, current_user_email) + original_body
            else:
                # Default: ReplyAll
                reply = email_item.ReplyAll()
                _remove_self_from_recipients(reply, current_user_email)
                _add_recipients(reply, args.to, args.cc)
                reply.HTMLBody = _wrap_body_font(args.body) + reply.HTMLBody

            reply.HTMLBody = _ensure_utf8_charset(reply.HTMLBody)
            reply.Subject = f"RE: {email_item.Subject}" if not email_item.Subject.startswith("RE:") else email_item.Subject

            if reply.Recipients.Count == 0:
                print("Error: No recipients found.")
                return 1

            count = reply.Recipients.Count
            _add_inline_images(reply, args.inline_image)
            _add_attachments(reply, args.attach)
            _attach_emails(reply, args.attach_email, session)
            _set_importance(reply, args.importance)
            subject = reply.Subject
            reply.Send()
            mode = "Reply (From only)" if args.only else "Reply-all"
            print(f"{mode} sent to {count} recipient(s)")
            _print_sent_entry_id(session, subject)
            return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_compose(args):
    """Compose and send new email (always HTML format)"""
    if not _resolve_body(args):
        return 1
    try:
        to_list = [x.strip() for x in args.to.replace(";", ",").split(",")] if args.to else []
        cc_normalized = _normalize_cc(args.cc)
        cc_list = [x.strip() for x in cc_normalized.replace(";", ",").split(",")] if cc_normalized else []

        with OutlookSessionManager() as session:
            mail = session.outlook.CreateItem(0)  # 0 = olMailItem

            # Set recipients
            for recipient in to_list:
                mail.Recipients.Add(recipient)
            if cc_list:
                for cc_recipient in cc_list:
                    cc_recip = mail.Recipients.Add(cc_recipient)
                    cc_recip.Type = 2  # 2 = olCC

            mail.Subject = args.subject

            # Display briefly to trigger Outlook signature insertion
            mail.Display(False)

            # Prepend body to signature HTML (same pattern as reply)
            mail.HTMLBody = _ensure_utf8_charset(_wrap_body_font(args.body) + mail.HTMLBody)

            _add_inline_images(mail, args.inline_image)
            _add_attachments(mail, args.attach)
            _attach_emails(mail, args.attach_email, session)
            _set_importance(mail, args.importance)
            mail.Send()
            total_recipients = len(to_list) + len(cc_list)
            print(f"HTML email sent successfully to {total_recipients} recipient(s)")
            _print_sent_entry_id(session, args.subject)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_batch_forward(args):
    """Batch forward email to multiple recipients by email ID"""
    try:
        # Import batch configuration from backend/config.py
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
        from backend.config import batch_config
        
        # Use configured batch size from backend/config.py
        batch_size = batch_config.OUTLOOK_BCC_LIMIT
        
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)
            email_subject = str(getattr(email_item, "Subject", "No Subject"))

            # Read CSV file (handle BOM if present)
            import csv
            recipients = []
            with open(args.csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    row_lower = {k.lower(): v for k, v in row.items()}
                    if 'email' in row_lower:
                        recipients.append(row_lower['email'].strip())
            
            if not recipients:
                print("Error: No email addresses found in CSV", file=sys.stderr)
                return 1
            
            # Forward to recipients in batches (batch size from config file)
            import time
            from datetime import datetime, timezone
            total_sent = 0
            num_batches = (len(recipients) + batch_size - 1) // batch_size
            send_start_time = datetime.now(timezone.utc)

            for i in range(0, len(recipients), batch_size):
                batch = recipients[i:i + batch_size]
                batch_num = i // batch_size + 1

                # Create forward
                forward = email_item.Forward()

                # Add custom message if provided (insert into body tag like reply does)
                if args.message:
                    forward.HTMLBody = _format_forward_message_html(args.message) + forward.HTMLBody

                # Add recipients as BCC (to protect privacy)
                for recipient in batch:
                    bcc_recip = forward.Recipients.Add(recipient)
                    bcc_recip.Type = 3  # 3 = olBCC

                # Resolve all recipients before sending
                forward.Recipients.ResolveAll()

                # Send
                forward.Send()
                total_sent += len(batch)
                print(f"Sent batch {batch_num}/{num_batches}: {len(batch)} recipients")

            print(f"\nTotal recipients: {total_sent}")
            print(f"Successfully forwarded email: {email_subject}")

            # Retrieve EntryIDs for all sent batches
            fw_subject = f"FW: {email_subject}" if not email_subject.startswith("FW:") else email_subject
            import re
            bare_subject = re.sub(r'^(RE:\s*|FW:\s*)+', '', fw_subject or '', flags=re.IGNORECASE)
            threshold = send_start_time

            for attempt in range(15):
                time.sleep(2)
                try:
                    ns = session.outlook.GetNamespace("MAPI")
                    sent_folder = ns.GetDefaultFolder(5)
                    items = sent_folder.Items
                    items.Sort("[SentOn]", True)
                    found_ids = []
                    item = items.GetFirst()
                    while item and len(found_ids) < num_batches:
                        sent_on = item.SentOn
                        if hasattr(sent_on, 'astimezone'):
                            sent_utc = sent_on.astimezone(timezone.utc)
                        else:
                            sent_utc = sent_on.replace(tzinfo=timezone.utc)
                        if sent_utc < threshold:
                            break
                        bare_item_subject = re.sub(r'^(RE:\s*|FW:\s*)+', '', item.Subject or '', flags=re.IGNORECASE)
                        if bare_subject in bare_item_subject:
                            found_ids.append(item.EntryID)
                        item = items.GetNext()
                    if len(found_ids) >= num_batches:
                        for idx, eid in enumerate(reversed(found_ids), 1):
                            print(f"EntryID (batch {idx}): {eid}")
                        break
                except Exception:
                    pass

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_forward(args):
    """Forward an email to specified recipients with optional CC and custom message."""
    if not _resolve_body(args, required=False):
        return 1
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            forward = email_item.Forward()

            # Meeting items preserve original attendees as recipients — clear them
            while forward.Recipients.Count > 0:
                forward.Recipients.Remove(1)

            # Set subject: use --subject override or default FW: prefix
            original_subject = str(getattr(email_item, "Subject", "No Subject"))
            if getattr(args, 'subject', None):
                forward.Subject = args.subject
            else:
                forward.Subject = f"FW: {original_subject}" if not original_subject.startswith("FW:") else original_subject

            # Add To recipients
            if args.to:
                for r in args.to.replace(";", ",").split(","):
                    r = r.strip()
                    if r:
                        forward.Recipients.Add(r)

            # Add CC recipients (args.cc is a list from action='append', each entry may be comma-separated)
            if args.cc:
                cc_joined = ",".join(args.cc)
                for r in cc_joined.replace(";", ",").split(","):
                    r = r.strip()
                    if r:
                        cc_recip = forward.Recipients.Add(r)
                        cc_recip.Type = 2  # 2 = olCC

            # Prepend custom message if provided
            if args.body:
                prepend_html = _format_forward_message_html(args.body)
                try:
                    original_html = forward.HTMLBody
                    forward.HTMLBody = _inject_before_body(prepend_html, original_html)
                except Exception:
                    # Meeting items / special types need inspector initialization
                    _ = forward.GetInspector
                    forward.HTMLBody = _inject_before_body(prepend_html, forward.HTMLBody)

            if forward.Recipients.Count == 0:
                print("Error: No recipients specified. Use --to.")
                return 1

            resolved = forward.Recipients.ResolveAll()
            if resolved is False:
                print("Error: One or more recipients could not be resolved.")
                return 1

            recipient_count = forward.Recipients.Count
            final_subject = str(getattr(forward, "Subject", original_subject))

            _add_inline_images(forward, args.inline_image)
            _add_attachments(forward, args.attach)
            _attach_emails(forward, args.attach_email, session)
            _set_importance(forward, args.importance)
            forward.Send()
            print(f"Forward sent to {recipient_count} recipient(s)")
            print(f"Subject: {final_subject}")
            _print_sent_entry_id(session, final_subject)
            return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_redirect(args):
    """Redirect an email: clear all recipients, add new TO/CC, preserve original body."""
    if not _resolve_body(args):
        return 1
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)
            forward = email_item.Forward()

            # Remove all pre-populated recipients
            while forward.Recipients.Count > 0:
                forward.Recipients.Remove(1)

            # Add new recipients
            _add_recipients(forward, args.to, args.cc)

            if forward.Recipients.Count == 0:
                print("Error: No recipients specified. Use --to or --cc.")
                return 1

            # Subject
            original_subject = str(getattr(email_item, "Subject", "No Subject"))
            forward.Subject = f"FW: {original_subject}" if not original_subject.startswith("FW:") else original_subject

            # Signature + body
            forward.Display(False)
            signature_html = forward.HTMLBody
            forward.HTMLBody = _ensure_utf8_charset(_wrap_body_font(args.body) + signature_html)

            recipient_count = forward.Recipients.Count
            final_subject = forward.Subject

            _add_inline_images(forward, args.inline_image)
            _add_attachments(forward, args.attach)
            forward.Send()
            print(f"Redirected to {recipient_count} recipient(s)")
            print(f"Subject: {final_subject}")
            _print_sent_entry_id(session, final_subject)
            return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_create_folder(args):
    """Create a new folder"""
    try:
        with OutlookSessionManager() as session:
            result = session.create_folder(args.name, args.parent)
            print(result)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_remove_folder(args):
    """Remove a folder"""
    try:
        from backend.outlook_session.folder_operations import remove_folder
        result = remove_folder(args.name)
        print(result)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_move_email(args):
    """Move an email to a folder by email ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(args.email_id)
            
            # Get target folder
            target_folder = session.get_folder(args.folder)
            if not target_folder:
                print(f"Error: Folder '{args.folder}' not found", file=sys.stderr)
                return 1
            
            # Move email
            email_item.Move(target_folder)
            print(f"Successfully moved email to '{args.folder}'")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_delete_email(args):
    """Delete an email by ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(args.email_id)
            subject = email_item.Subject
            email_item.Delete()
            print(f"Successfully deleted email: {subject}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_recall(args):
    """Recall a sent email via Exchange server"""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            if not email_item.Sent:
                print("Error: Can only recall sent emails.", file=sys.stderr)
                return 1

            subject = email_item.Subject
            recipients = email_item.To
            print(f"Recalling: {subject}")
            print(f"Recipients: {recipients}")

            email_item.Display()
            inspector = email_item.GetInspector
            inspector.CommandBars.ExecuteMso("RecallThisMessage")
            print(f"\n✅ Recall dialog opened for: {subject}")
            print("Please confirm in the Outlook dialog, then close the email window.")
            print("Note: Recall only works for internal Exchange recipients who haven't read the message.")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_lookup_contact(args):
    """Look up contact information by email address or display name"""
    try:
        query = args.query
        if '@' in query:
            contact_info = get_contact_by_email(query)
            results = [contact_info] if contact_info else []
        else:
            results = get_contact_by_name(query)

        if not results:
            print(f"No contact found for: {query}")
            return 0

        print(f"\n{'='*50}")
        print(f" Contact Results ({len(results)} match{'es' if len(results) > 1 else ''})")
        print(f"{'='*50}")

        for i, contact in enumerate(results, 1):
            if len(results) > 1:
                print(f"\n--- [{i}] ---")
            else:
                print()
            print(f"Display Name: {contact['display_name']}")
            print(f"Email: {contact['email']}")
            if contact.get('alias'):
                print(f"Alias: {contact['alias']}")
            if contact.get('first_name'):
                print(f"First Name: {contact['first_name']}")
            if contact.get('last_name'):
                print(f"Last Name: {contact['last_name']}")
            if contact.get('company'):
                print(f"Company: {contact['company']}")
            if contact.get('department'):
                print(f"Department: {contact['department']}")
            if contact.get('job_title'):
                print(f"Job Title: {contact['job_title']}")
            if contact.get('office'):
                print(f"Office: {contact['office']}")
            if contact.get('business_phone'):
                print(f"Business Phone: {contact['business_phone']}")
            if contact.get('mobile_phone'):
                print(f"Mobile: {contact['mobile_phone']}")
            if contact.get('city') or contact.get('state'):
                location_parts = [p for p in [contact.get('city'), contact.get('state')] if p]
                print(f"Location: {', '.join(location_parts)}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_find_thread(args):
    """Find all emails in the same conversation thread."""
    try:
        emails, message = find_thread_by_email_id(
            args.email_id,
            folder_names=(
                [f.strip() for f in args.folders.split(',')]
                if args.folders else None
            ),
            fuzzy=getattr(args, 'fuzzy', False),
        )
        print(f"\n\U0001F9F5 {message}\n")

        if emails:
            _print_thread_summary(emails)
            if getattr(args, 'brief', False):
                _display_email_list_brief(emails, show_folder=True)
            else:
                _display_email_list(emails, show_folder=True)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_download_attachment(args):
    """Download attachments from an email to a local directory."""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            if email_item.Attachments.Count == 0:
                print("No attachments found on this email.")
                return 0

            save_dir = args.output_dir or os.path.join(os.path.expanduser("~"), "Downloads")
            os.makedirs(save_dir, exist_ok=True)

            saved = []
            for i in range(1, email_item.Attachments.Count + 1):
                attach = email_item.Attachments.Item(i)
                # Skip embedded images (cid: references) unless --all flag
                # Only skip if it looks like an image (not PDF/doc/etc)
                if not args.all and hasattr(attach, 'PropertyAccessor'):
                    try:
                        content_id = attach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                        if content_id:
                            ext = os.path.splitext(attach.FileName)[1].lower()
                            image_exts = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.ico', '.svg'}
                            if ext in image_exts:
                                continue
                    except Exception:
                        pass

                if args.filename and attach.FileName.lower() != args.filename.lower():
                    continue

                save_path = os.path.join(save_dir, attach.FileName)
                attach.SaveAsFile(save_path)
                size_kb = os.path.getsize(save_path) / 1024
                saved.append((attach.FileName, save_path, size_kb))
                print(f"✅ Saved: {attach.FileName} ({size_kb:.1f} KB) → {save_path}")

            if not saved:
                if args.filename:
                    print(f"Attachment '{args.filename}' not found.")
                else:
                    print("No file attachments found (only embedded images).")
            else:
                print(f"\n{len(saved)} attachment(s) saved to: {save_dir}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_find_related(args):
    """Find emails related to a given email using multiple strategies."""
    try:
        strategies = None
        if args.strategies:
            strategies = [s.strip() for s in args.strategies.split(',')]

        emails, message = find_related_emails(
            args.email_id,
            days=args.days,
            strategies=strategies,
            exclude_thread=args.exclude_thread,
            max_results=getattr(args, 'max_results', None),
        )
        print(f"\n\U0001F517 {message}\n")

        if emails:
            if getattr(args, 'brief', False):
                _display_email_list_brief(emails, show_folder=True)
            else:
                _display_email_list(emails, show_folder=True)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_get_html(args):
    """Get the raw HTMLBody of an email for template editing."""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            subject = getattr(email_item, 'Subject', 'No Subject')
            sender = getattr(email_item, 'SenderName', 'Unknown')
            html_body = getattr(email_item, 'HTMLBody', '')

            print(f"Subject: {subject}")
            print(f"From: {sender}")
            print(f"To: {getattr(email_item, 'To', '')}")
            if getattr(email_item, 'CC', ''):
                print(f"CC: {email_item.CC}")
            print(f"HTML length: {len(html_body)} chars")
            print(f"\n{'='*60}")
            print(f"HTML_START")
            print(html_body)
            print(f"HTML_END")
            print(f"{'='*60}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_edit_html(args):
    """Edit an email's HTML and save.

    If the email is a draft (unsent): modifies it IN PLACE and saves.
    If the email is sent/received: creates a new draft with modified content.
    """
    try:
        with OutlookSessionManager() as session:
            source_item = _get_email_item(session, args.email_id)

            is_draft = not getattr(source_item, 'Sent', True)

            # Start with source HTML
            html_body = getattr(source_item, 'HTMLBody', '')
            if not html_body:
                print("Error: Source email has no HTML body.", file=sys.stderr)
                return 1

            # Apply --body-file FIRST (full body replacement)
            if args.body_file:
                with open(args.body_file, 'r', encoding='utf-8') as f:
                    html_body = f.read()
                print(f"Body replaced from file: {args.body_file}")

            # Apply --replace pairs on top
            replacements_made = 0
            if args.replace:
                for pair in args.replace:
                    if '::' not in pair:
                        print(f"Error: Invalid replace format '{pair}'. Use 'old::new'.", file=sys.stderr)
                        return 1
                    old, new = pair.split('::', 1)
                    if old in html_body:
                        html_body = html_body.replace(old, new)
                        replacements_made += 1
                    else:
                        print(f"WARNING: '{old}' not found in HTML body.")

            if is_draft:
                # --- Modify existing draft in place ---
                draft = source_item

                if args.subject:
                    draft.Subject = args.subject
                if args.to:
                    draft.To = args.to.replace(',', '; ')
                cc_normalized = _normalize_cc(args.cc)
                if cc_normalized:
                    draft.CC = cc_normalized.replace(',', '; ')

                draft.HTMLBody = html_body
                draft.Save()

                print(f"✅ Draft updated in place")
            else:
                # --- Create new draft from sent/received email ---
                draft = session.outlook.CreateItem(0)  # olMailItem

                if args.subject:
                    draft.Subject = args.subject
                else:
                    draft.Subject = getattr(source_item, 'Subject', '')

                if args.to:
                    draft.To = args.to.replace(',', '; ')
                else:
                    to_addrs = []
                    for i in range(1, source_item.Recipients.Count + 1):
                        recip = source_item.Recipients.Item(i)
                        if recip.Type == 1:  # olTo
                            to_addrs.append(_get_smtp_address(recip))
                    draft.To = '; '.join(to_addrs)

                cc_normalized = _normalize_cc(args.cc)
                if cc_normalized:
                    draft.CC = cc_normalized.replace(',', '; ')
                else:
                    cc_addrs = []
                    for i in range(1, source_item.Recipients.Count + 1):
                        recip = source_item.Recipients.Item(i)
                        if recip.Type == 2:  # olCC
                            cc_addrs.append(_get_smtp_address(recip))
                    if cc_addrs:
                        draft.CC = '; '.join(cc_addrs)

                # Set modified HTML body first
                draft.HTMLBody = html_body

                # Copy embedded images to preserve cid: references in HTML
                import tempfile
                for i in range(1, source_item.Attachments.Count + 1):
                    attach = source_item.Attachments.Item(i)
                    try:
                        cid = attach.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                        )
                    except Exception:
                        cid = None
                    if cid:
                        temp_path = os.path.join(tempfile.gettempdir(), attach.FileName)
                        attach.SaveAsFile(temp_path)
                        new_attach = draft.Attachments.Add(temp_path)
                        new_attach.PropertyAccessor.SetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
                        )
                        os.remove(temp_path)

                # Force Outlook to re-resolve CID references
                draft.HTMLBody = draft.HTMLBody

                # Copy file attachments (non-embedded) from source if requested
                if args.copy_attachments:
                    for i in range(1, source_item.Attachments.Count + 1):
                        attach = source_item.Attachments.Item(i)
                        try:
                            cid = attach.PropertyAccessor.GetProperty(
                                "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                            )
                            if cid:
                                continue
                        except Exception:
                            pass
                        import tempfile
                        temp_path = os.path.join(tempfile.gettempdir(), attach.FileName)
                        attach.SaveAsFile(temp_path)
                        draft.Attachments.Add(temp_path)

                draft.Save()
                print(f"✅ New draft created in Drafts folder")

            print(f"   Subject: {draft.Subject}")
            print(f"   To: {draft.To}")
            if draft.CC:
                print(f"   CC: {draft.CC}")
            if replacements_made:
                print(f"   Replacements applied: {replacements_made}")
            print(f"   HTML length: {len(html_body)} chars")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_send_draft(args):
    """Send an existing draft email."""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            if getattr(email_item, 'Sent', False):
                print("Error: This email has already been sent.", file=sys.stderr)
                return 1

            subject = getattr(email_item, 'Subject', 'No Subject')
            to = getattr(email_item, 'To', '')
            cc = getattr(email_item, 'CC', '')

            if not to:
                print("Error: Draft has no recipients.", file=sys.stderr)
                return 1

            email_item.Send()
            print(f"✅ Email sent")
            print(f"   Subject: {subject}")
            print(f"   To: {to}")
            if cc:
                print(f"   CC: {cc}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


PR_OOF_STATE = 'http://schemas.microsoft.com/mapi/proptag/0x661D000B'


def cmd_get_ooo(args):
    """Get current Out of Office status."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        store = namespace.DefaultStore
        pa = store.PropertyAccessor

        oof_state = pa.GetProperty(PR_OOF_STATE)
        print(f"{'=' * 50}")
        print(f" Out of Office — {store.DisplayName}")
        print(f"{'=' * 50}")
        print(f"\n  Status: {'🟢 ENABLED' if oof_state else '⚪ DISABLED'}")
        if oof_state:
            print("\n  To disable: py -3 scripts/outlook_skill.py disable-ooo")
        else:
            print("\n  To enable: py -3 scripts/outlook_skill.py set-ooo")
        print(f"\n  💡 To view/edit auto-reply message: Outlook > File > Automatic Replies")

        pythoncom.CoUninitialize()
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_set_ooo(args):
    """Enable Out of Office auto-reply."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        store = namespace.DefaultStore
        pa = store.PropertyAccessor

        current = pa.GetProperty(PR_OOF_STATE)
        if current:
            print(f"⚠️ Out of Office is already enabled for {store.DisplayName}.")
            pythoncom.CoUninitialize()
            return 0

        pa.SetProperty(PR_OOF_STATE, True)
        verified = pa.GetProperty(PR_OOF_STATE)
        if verified:
            print(f"✅ Out of Office ENABLED for {store.DisplayName}.")
            print(f"   💡 Make sure your auto-reply message is set in Outlook > File > Automatic Replies")
        else:
            print(f"⚠️ Failed to enable OOF — state did not change.")

        pythoncom.CoUninitialize()
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_disable_ooo(args):
    """Disable Out of Office auto-reply."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        store = namespace.DefaultStore
        pa = store.PropertyAccessor

        current = pa.GetProperty(PR_OOF_STATE)
        if not current:
            print(f"⚪ Out of Office is already disabled for {store.DisplayName}.")
            pythoncom.CoUninitialize()
            return 0

        pa.SetProperty(PR_OOF_STATE, False)
        verified = pa.GetProperty(PR_OOF_STATE)
        if not verified:
            print(f"✅ Out of Office DISABLED for {store.DisplayName}.")
        else:
            print(f"⚠️ Failed to disable OOF — state did not change.")

        pythoncom.CoUninitialize()
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def _get_signatures_path():
    """Return the Outlook signatures folder path."""
    return os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures')


def _read_sig_file(filepath):
    """Read a signature file, handling UTF-16, UTF-8, and legacy encodings."""
    with open(filepath, 'rb') as f:
        raw = f.read()
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
        return raw.decode('utf-16')
    try:
        return raw.decode('utf-8')
    except UnicodeDecodeError:
        return raw.decode('latin-1')


def cmd_get_signature(args):
    """List all signatures or view a specific one."""
    try:
        sig_path = _get_signatures_path()
        if not os.path.exists(sig_path):
            print("No signatures folder found at:", sig_path)
            return 1

        htm_files = [f for f in os.listdir(sig_path) if f.endswith('.htm')]
        if not htm_files:
            print("No signatures found.")
            return 1

        name = getattr(args, 'name', None)
        fmt = getattr(args, 'format', 'text')

        if name:
            htm_path = os.path.join(sig_path, f"{name}.htm")
            txt_path = os.path.join(sig_path, f"{name}.txt")
            if not os.path.exists(htm_path):
                print(f"Signature '{name}' not found.")
                print(f"Available: {', '.join(f[:-4] for f in htm_files)}")
                return 1
            if fmt == 'html':
                print(_read_sig_file(htm_path))
            else:
                if os.path.exists(txt_path):
                    print(_read_sig_file(txt_path))
                else:
                    print(_read_sig_file(htm_path))
        else:
            print(f"{'=' * 60}")
            print(f" Signatures ({len(htm_files)} found)")
            print(f"{'=' * 60}")
            for htm_file in htm_files:
                sig_name = htm_file[:-4]
                txt_path = os.path.join(sig_path, f"{sig_name}.txt")
                print(f"\n--- {sig_name} ---")
                if os.path.exists(txt_path):
                    print(_read_sig_file(txt_path).strip())
                else:
                    print("(no .txt version available)")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def _update_rtf_signature(rtf_path, find_text, replace_text):
    """Update text in an RTF signature file. Returns True if updated."""
    if not os.path.exists(rtf_path):
        return False
    with open(rtf_path, 'rb') as f:
        content = f.read().decode('latin-1')
    escaped_find = _rtf_escape(find_text)
    escaped_replace = _rtf_escape(replace_text)
    if escaped_find in content:
        content = content.replace(escaped_find, escaped_replace)
        with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
            f.write(content)
        return True
    if find_text in content:
        content = content.replace(find_text, escaped_replace)
        with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
            f.write(content)
        return True
    return False


def _rtf_escape(text):
    """Escape non-ASCII characters for RTF using \\uN notation."""
    result = []
    for ch in text:
        if ord(ch) > 127:
            result.append(f'\\u{ord(ch)}?')
        else:
            result.append(ch)
    return ''.join(result)


def _rebuild_rtf_content(rtf_path, lines):
    """Replace all content paragraphs in RTF with new lines. Generates new RTF if needed."""
    import re as _re

    content = ''
    if os.path.exists(rtf_path):
        with open(rtf_path, 'rb') as f:
            content = f.read().decode('latin-1')

    # Try to find and replace existing content block
    if content:
        pattern = _re.compile(
            r'(\\ltrch\\fcs0 \\insrsid\d+ \\hich\\af31506\\dbch\\af31505\\loch\\f31506 )'
            r'(.+?)'
            r'(\r\n\\par \}\\pard)',
            _re.DOTALL
        )
        match = pattern.search(content)
        if match:
            rtf_prefix = r'\hich\af31506\dbch\af31505\loch\f31506 '
            rtf_lines = []
            for line in lines:
                if line.strip():
                    rtf_lines.append(rtf_prefix + _rtf_escape(line))
                else:
                    rtf_lines.append('')
            new_content = '\r\n\\par '.join(rtf_lines)
            content = content[:match.start(2)] + new_content + content[match.end(2):]
            with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
                f.write(content)
            return True

    # Fallback: generate minimal valid RTF from scratch
    escaped_lines = []
    for line in lines:
        escaped_lines.append(_rtf_escape(line) if line.strip() else '')
    par_content = '\\par\r\n'.join(escaped_lines)

    rtf = (
        r'{\rtf1\ansi\ansicpg1252\deff0\nouicompat'
        r'{\fonttbl{\f0\fswiss\fcharset0 Aptos;}}'
        '\r\n'
        r'{\*\generator Microsoft Word 15}'
        r'\viewkind4\uc1'
        '\r\n'
        r'\pard\f0\fs22 '
        + par_content +
        '\r\n\\par }'
    )
    with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
        f.write(rtf)
    return True


def cmd_update_signature(args):
    """Update a signature's content."""
    try:
        sig_path = _get_signatures_path()
        if not os.path.exists(sig_path):
            print("No signatures folder found at:", sig_path)
            return 1

        name = args.name
        htm_path = os.path.join(sig_path, f"{name}.htm")
        txt_path = os.path.join(sig_path, f"{name}.txt")
        rtf_path = os.path.join(sig_path, f"{name}.rtf")

        if not os.path.exists(htm_path):
            htm_files = [f for f in os.listdir(sig_path) if f.endswith('.htm')]
            print(f"Signature '{name}' not found.")
            if htm_files:
                print(f"Available: {', '.join(f[:-4] for f in htm_files)}")
            return 1

        import shutil
        shutil.copy2(htm_path, htm_path + '.bak')
        if os.path.exists(txt_path):
            shutil.copy2(txt_path, txt_path + '.bak')
        if os.path.exists(rtf_path):
            shutil.copy2(rtf_path, rtf_path + '.bak')

        if args.text:
            lines = args.text.split('\\n')
            line_end = '\r\n'
            paragraphs = []
            for line in lines:
                if line.strip():
                    paragraphs.append(f'<p class=MsoNormal>{line}<o:p></o:p></p>')
                else:
                    paragraphs.append(f'<p class=MsoNormal><o:p>&nbsp;</o:p></p>')
            body_html = line_end.join(paragraphs)

            # Preserve original HTML template (head/styles) if backup exists
            bak_path = htm_path + '.bak'
            if os.path.exists(bak_path):
                template = _read_sig_file(bak_path)
            else:
                template = _read_sig_file(htm_path)

            # Replace everything between <div> and </div> with new paragraphs
            import re as _re
            div_pattern = _re.compile(r'(<div[^>]*>).*?(</div>)', _re.DOTALL)
            match = div_pattern.search(template)
            if match:
                htm_content = template[:match.start(1)] + match.group(1) + line_end + body_html + line_end + match.group(2) + template[match.end(2):]
            else:
                htm_content = (
                    '<html xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
                    'xmlns:w="urn:schemas-microsoft-com:office:word"\r\n'
                    'xmlns="http://www.w3.org/TR/REC-html40">\r\n'
                    '<head><meta http-equiv=Content-Type content="text/html; charset=utf-8">\r\n'
                    '<meta name=Generator content="Microsoft Word 15"></head>\r\n'
                    f'<body>\r\n<div>\r\n\r\n{body_html}\r\n\r\n</div>\r\n</body>\r\n</html>\r\n'
                )

            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(htm_content)
            txt_content = '\r\n'.join(line if line.strip() else '' for line in lines)
            with open(txt_path, 'w', encoding='utf-16') as f:
                f.write(txt_content)

            # Update RTF
            rtf_ok = _rebuild_rtf_content(rtf_path, lines)

            print(f"✅ Signature '{name}' replaced from text.")
            print(f"   HTM: ✓ | TXT: ✓ | RTF: {'✓' if rtf_ok else '✗ (manual check needed)'}")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")
            for line in lines:
                print(f"   {line}")

        elif args.body:
            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(args.body)
            from html.parser import HTMLParser
            class _TextExtractor(HTMLParser):
                def __init__(self):
                    super().__init__()
                    self.result = []
                def handle_data(self, data):
                    self.result.append(data)
            extractor = _TextExtractor()
            extractor.feed(args.body)
            with open(txt_path, 'w', encoding='utf-16') as f:
                f.write(''.join(extractor.result))
            print(f"✅ Signature '{name}' replaced (full HTML body).")
            print(f"   ⚠️ RTF not updated — use --text mode for full update.")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")

        elif args.after and args.insert:
            htm_content = _read_sig_file(htm_path)

            if args.after not in htm_content:
                print(f"⚠️ Text '{args.after}' not found in signature '{name}'.")
                return 1

            anchor_idx = htm_content.find(args.after)
            end_p = htm_content.find('</p>', anchor_idx)
            if end_p == -1:
                insert_point = anchor_idx + len(args.after)
            else:
                insert_point = end_p + len('</p>')

            line_end = '\r\n' if '\r\n' in htm_content else '\n'
            new_paragraphs = ''
            for line in args.insert.split('\\n'):
                if line.strip():
                    new_paragraphs += f'{line_end}<p class=MsoNormal>{line}<o:p></o:p></p>'
                else:
                    new_paragraphs += f'{line_end}<p class=MsoNormal><o:p>&nbsp;</o:p></p>'

            htm_content = htm_content[:insert_point] + new_paragraphs + htm_content[insert_point:]
            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(htm_content)

            # Update .txt
            if os.path.exists(txt_path):
                txt_content = _read_sig_file(txt_path)
                txt_line_end = '\r\n' if '\r\n' in txt_content else '\n'
                txt_insert = txt_line_end + args.insert.replace('\\n', txt_line_end)
                anchor_idx_txt = txt_content.find(args.after)
                if anchor_idx_txt != -1:
                    eol = txt_content.find(txt_line_end, anchor_idx_txt)
                    if eol == -1:
                        eol = len(txt_content)
                    txt_content = txt_content[:eol] + txt_insert + txt_content[eol:]
                    with open(txt_path, 'w', encoding='utf-16') as f:
                        f.write(txt_content)

            # Update RTF — insert after anchor text
            if os.path.exists(rtf_path):
                rtf_content = open(rtf_path, 'rb').read().decode('latin-1')
                rtf_prefix = r'\hich\af31506\dbch\af31505\loch\f31506 '
                if args.after in rtf_content:
                    rtf_anchor_idx = rtf_content.find(args.after)
                    rtf_eol = rtf_content.find('\r\n', rtf_anchor_idx)
                    if rtf_eol == -1:
                        rtf_eol = len(rtf_content)
                    rtf_new_lines = ''
                    for line in args.insert.split('\\n'):
                        if line.strip():
                            rtf_new_lines += f'\r\n\\par {rtf_prefix}{_rtf_escape(line)}'
                        else:
                            rtf_new_lines += '\r\n\\par '
                    rtf_content = rtf_content[:rtf_eol] + rtf_new_lines + rtf_content[rtf_eol:]
                    with open(rtf_path, 'w', encoding='latin-1', newline='') as f:
                        f.write(rtf_content)

            insert_lines = args.insert.replace('\\n', '\n')
            print(f"✅ Signature '{name}' updated (inserted after '{args.after}').")
            print(f"   Added: {insert_lines}")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")

        elif args.find and args.replace is not None:
            htm_content = _read_sig_file(htm_path)

            if args.find not in htm_content:
                if os.path.exists(txt_path):
                    txt_content = _read_sig_file(txt_path)
                    if args.find not in txt_content:
                        print(f"⚠️ Text '{args.find}' not found in signature '{name}'.")
                        return 1

            count = htm_content.count(args.find)
            htm_content = htm_content.replace(args.find, args.replace)
            with open(htm_path, 'w', encoding='utf-8') as f:
                f.write(htm_content)

            if os.path.exists(txt_path):
                txt_content = _read_sig_file(txt_path)
                txt_content = txt_content.replace(args.find, args.replace)
                with open(txt_path, 'w', encoding='utf-16') as f:
                    f.write(txt_content)

            # Update RTF
            _update_rtf_signature(rtf_path, args.find, args.replace)

            print(f"✅ Signature '{name}' updated ({count} replacement(s)).")
            print(f"   '{args.find}' → '{args.replace}'")
            print(f"   Backup: {htm_path}.bak")
            print(f"   ⚠️ Restart Outlook for changes to take effect.")
        else:
            print("Error: provide --text, --body, --find/--replace, or --after/--insert.")
            return 1

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def main():
    parser = argparse.ArgumentParser(description='Outlook Skill for BrainClaw - Email Management CLI')
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # List folders command
    parser_list_folders = subparsers.add_parser('list-folders', help='List all Outlook folders')
    parser_list_folders.add_argument('--hide-system', action='store_true', help='Hide system folders')
    parser_list_folders.set_defaults(func=cmd_list_folders)
    
    # List recent emails command
    parser_list_recent = subparsers.add_parser('find-recent', help='Find recent emails with IDs')
    parser_list_recent.add_argument('--days', type=int, default=7, help='Days back to search (1-30)')
    parser_list_recent.add_argument('--folder', type=str, default=None, help='Folder name (default: Inbox)')
    parser_list_recent.add_argument('--json', action='store_true', help='Output JSON for piping to email_sync.py')
    parser_list_recent.set_defaults(func=cmd_list_recent)
    
    # Search emails command
    parser_search = subparsers.add_parser('find', help='Find emails by subject, sender, recipient, or body')
    parser_search.add_argument('--type', required=True, choices=['subject', 'sender', 'recipient', 'body'], help='Search type')
    parser_search.add_argument('--query', required=True, help='Search query')
    parser_search.add_argument(
        '--days',
        type=int,
        default=search_config.DIRECT_FIND_DEFAULT_DAYS,
        help=(
            f"Days back to search "
            f"(default: {search_config.DIRECT_FIND_DEFAULT_DAYS}, "
            f"allowed range: 1-{search_config.MAX_SEARCH_DAYS})"
        ),
    )
    parser_search.add_argument('--folder', type=str, default=None, help='Folder name (default: Inbox)')
    parser_search.add_argument('--folders', type=str, default=None, help='Comma-separated folder names for cross-folder search')
    parser_search.add_argument('--match-all', action='store_true', default=True, help='Match all terms (AND logic)')
    parser_search.set_defaults(func=cmd_search)
    
    # Get email command
    parser_get_email = subparsers.add_parser('get-email', aliases=['read'], help='Get full email details by ID')
    parser_get_email.add_argument('email_id', nargs='?', default=None, help='Email ID from search results')
    parser_get_email.add_argument('--id', dest='email_id_flag', help='Email ID (alternative to positional)')
    parser_get_email.add_argument('--fields', help='Comma-separated fields to show (from,to,cc,subject,date,body)')
    parser_get_email.set_defaults(func=cmd_get_email)
    
    # Reply command (default: reply-all; --only: sender only)
    parser_reply = subparsers.add_parser('reply', help='Reply to email (default: reply-all; --only: From only)')
    parser_reply.add_argument('email_id', help='Email ID from search results')
    parser_reply.add_argument('body', nargs='?', default=None, help='Reply text in HTML format')
    parser_reply.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_reply.add_argument('--to', help='Additional To recipients (comma separated)')
    parser_reply.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_reply.add_argument('--attach', help='File path(s) to attach (comma separated)')
    parser_reply.add_argument('--attach-email', help='Email ID(s) to attach as .msg (comma separated)')
    parser_reply.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_reply.add_argument('--importance', choices=['high', 'low'], help='Set email importance (high or low)')
    parser_reply.add_argument('--only', action='store_true', help='Reply to From (sender) only instead of reply-all')
    parser_reply.set_defaults(func=cmd_reply)
    
    # Compose command
    parser_compose = subparsers.add_parser('compose', help='Compose and send new email')
    parser_compose.add_argument('--to', required=True, help='To recipients (comma separated)')
    parser_compose.add_argument('--subject', required=True, help='Email subject')
    parser_compose.add_argument('--body', help='Email body')
    parser_compose.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_compose.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_compose.add_argument('--attach', help='File path(s) to attach (comma separated)')
    parser_compose.add_argument('--attach-email', help='Email ID(s) to attach as .msg (comma separated)')
    parser_compose.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_compose.add_argument('--importance', choices=['high', 'low'], help='Set email importance (high or low)')
    parser_compose.set_defaults(func=cmd_compose)
    
    # Forward command
    parser_forward = subparsers.add_parser('forward', help='Forward an email to specified recipients')
    parser_forward.add_argument('email_id', help='Email ID from search results')
    parser_forward.add_argument('--to', required=True, help='To recipients (comma separated)')
    parser_forward.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_forward.add_argument('--body', help='Custom message to prepend')
    parser_forward.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_forward.add_argument('--attach', help='File path(s) to attach (comma separated)')
    parser_forward.add_argument('--attach-email', help='Email ID(s) to attach as .msg (comma separated)')
    parser_forward.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_forward.add_argument('--importance', choices=['high', 'low'], help='Set email importance (high or low)')
    parser_forward.add_argument('--subject', help='Override the forwarded subject (default: FW: <original>)')
    parser_forward.set_defaults(func=cmd_forward)

    # Redirect command (clear all recipients, add new ones)
    parser_redirect = subparsers.add_parser('redirect', help='Redirect email: clear all recipients, set new TO/CC')
    parser_redirect.add_argument('email_id', help='Email ID from search results')
    parser_redirect.add_argument('body', nargs='?', default=None, help='Message body in HTML format')
    parser_redirect.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_redirect.add_argument('--to', required=True, help='To recipients (comma separated)')
    parser_redirect.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_redirect.add_argument('--attach', help='File path(s) to attach (comma separated)')
    parser_redirect.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_redirect.set_defaults(func=cmd_redirect)

    # Batch forward command
    parser_batch = subparsers.add_parser('batch-forward', help='Batch forward email by ID to multiple recipients')
    parser_batch.add_argument('email_id', help='Email ID from search results')
    parser_batch.add_argument('csv_path', help='Path to CSV file with email addresses')
    parser_batch.add_argument('--message', help='Custom message to prepend (HTML format)')
    parser_batch.set_defaults(func=cmd_batch_forward)
    
    # Create folder command
    parser_create_folder = subparsers.add_parser('create-folder', help='Create a new folder')
    parser_create_folder.add_argument('name', help='Folder name')
    parser_create_folder.add_argument('--parent', help='Parent folder name')
    parser_create_folder.set_defaults(func=cmd_create_folder)
    
    # Remove folder command
    parser_remove_folder = subparsers.add_parser('remove-folder', help='Remove a folder')
    parser_remove_folder.add_argument('name', help='Folder name or path')
    parser_remove_folder.set_defaults(func=cmd_remove_folder)
    
    # Move email command
    parser_move = subparsers.add_parser('move-email', help='Move an email to a folder by ID')
    parser_move.add_argument('email_id', help='Email ID from search results')
    parser_move.add_argument('folder', help='Target folder name')
    parser_move.set_defaults(func=cmd_move_email)
    
    # Delete email command
    parser_delete = subparsers.add_parser('delete-email', help='Delete an email by ID')
    parser_delete.add_argument('email_id', help='Email ID from search results')
    parser_delete.set_defaults(func=cmd_delete_email)

    # Recall email command
    parser_recall = subparsers.add_parser('recall', help='Recall a sent email via Exchange')
    parser_recall.add_argument('email_id', help='Email ID from search results (must be in Sent Items)')
    parser_recall.set_defaults(func=cmd_recall)

    # Lookup contact command
    parser_lookup = subparsers.add_parser('lookup-contact', help='Look up contact information by email or display name')
    parser_lookup.add_argument('query', help='Email address or display name to look up')
    parser_lookup.set_defaults(func=cmd_lookup_contact)

    # Find thread command
    parser_thread = subparsers.add_parser('find-thread', help='Find all emails in same conversation thread')
    parser_thread.add_argument('email_id', help='Email ID from search results')
    parser_thread.add_argument('--folders', type=str, default=None, help='Folders to search (default: Inbox,Sent Items)')
    parser_thread.add_argument('--fuzzy', action='store_true', help='Also find emails with similar subjects (token overlap)')
    parser_thread.add_argument('--brief', action='store_true', help='Compact single-line output')
    parser_thread.set_defaults(func=cmd_find_thread)

    # Find related command
    parser_related = subparsers.add_parser('find-related', help='Find emails related to a given email')
    parser_related.add_argument('email_id', help='Email ID from search results')
    parser_related.add_argument('--days', type=int, default=90, help='Days back for sender/keyword strategies')
    parser_related.add_argument('--strategies', type=str, default=None, help='Strategies: thread,sender,recipient,keyword (default: all)')
    parser_related.add_argument('--exclude-thread', action='store_true', help='Skip thread strategy (useful after find-thread)')
    parser_related.add_argument('--max', type=int, dest='max_results', default=None, help='Maximum results to return (default: from config)')
    parser_related.add_argument('--brief', action='store_true', help='Compact single-line output')
    parser_related.set_defaults(func=cmd_find_related)

    # Download attachment command
    parser_download = subparsers.add_parser('download-attachment', help='Download attachments from an email to local directory')
    parser_download.add_argument('email_id', help='Email ID from search results')
    parser_download.add_argument('--output-dir', help='Directory to save attachments (default: ~/Downloads)')
    parser_download.add_argument('--filename', help='Download only this specific attachment filename')
    parser_download.add_argument('--all', action='store_true', help='Include embedded images (default: skip them)')
    parser_download.set_defaults(func=cmd_download_attachment)

    # Get HTML command
    parser_get_html = subparsers.add_parser('get-html', help='Get raw HTMLBody of an email for template editing')
    parser_get_html.add_argument('email_id', help='Email ID from search results')
    parser_get_html.set_defaults(func=cmd_get_html)

    # Send draft command
    parser_send_draft = subparsers.add_parser('send-draft', help='Send an existing draft email')
    parser_send_draft.add_argument('email_id', help='Draft email ID')
    parser_send_draft.set_defaults(func=cmd_send_draft)

    # Edit HTML command (template editing → save to Drafts)
    parser_edit_html = subparsers.add_parser('edit-html', help='Edit email HTML and save as new draft')
    parser_edit_html.add_argument('email_id', help='Source email ID (used as template)')
    parser_edit_html.add_argument('--replace', action='append', help='Text replacement: "old::new" (repeatable)')
    parser_edit_html.add_argument('--subject', help='Override subject line')
    parser_edit_html.add_argument('--to', help='Override To recipients (comma separated)')
    parser_edit_html.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_edit_html.add_argument('--body-file', help='Replace entire HTML body from file path')
    parser_edit_html.add_argument('--copy-attachments', action='store_true', help='Copy non-embedded attachments from source')
    parser_edit_html.set_defaults(func=cmd_edit_html)

    # Get OOO status command
    parser_get_ooo = subparsers.add_parser('get-ooo', help='Get Out of Office status')
    parser_get_ooo.set_defaults(func=cmd_get_ooo)

    # Set OOO command
    parser_set_ooo = subparsers.add_parser('set-ooo', help='Enable Out of Office auto-reply')
    parser_set_ooo.set_defaults(func=cmd_set_ooo)

    # Disable OOO command
    parser_disable_ooo = subparsers.add_parser('disable-ooo', help='Disable Out of Office auto-reply')
    parser_disable_ooo.set_defaults(func=cmd_disable_ooo)

    # Get signature command
    parser_get_sig = subparsers.add_parser('get-signature', help='List all signatures or view a specific one')
    parser_get_sig.add_argument('name', nargs='?', default=None, help='Signature name (omit to list all)')
    parser_get_sig.add_argument('--format', choices=['text', 'html'], default='text', help='Output format (default: text)')
    parser_get_sig.set_defaults(func=cmd_get_signature)

    # Update signature command
    parser_update_sig = subparsers.add_parser('update-signature', help='Modify a signature')
    parser_update_sig.add_argument('name', help='Signature name')
    parser_update_sig.add_argument('--text', help='Full plain text replacement (use \\n for line breaks)')
    parser_update_sig.add_argument('--body', help='Full HTML replacement for the signature')
    parser_update_sig.add_argument('--find', help='Text to find (use with --replace)')
    parser_update_sig.add_argument('--replace', help='Replacement text (use with --find)')
    parser_update_sig.add_argument('--after', help='Insert new content after this text (use with --insert)')
    parser_update_sig.add_argument('--insert', help='Text to insert (use \\n for new lines, use with --after)')
    parser_update_sig.set_defaults(func=cmd_update_signature)

    # Parse arguments
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    # Execute command
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())