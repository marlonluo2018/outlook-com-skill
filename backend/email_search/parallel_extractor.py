"""
Parallel email extraction for performance optimization.
"""

# Standard library imports
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any, Dict, List

# Local application imports
from ..logging_config import get_logger

logger = get_logger(__name__)

# Thread-local storage for COM objects
_thread_local = threading.local()

_MAPI_START_PROP = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/820D0040'
_MAPI_END_PROP = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/820E0040'


def _get_sender_smtp(item) -> str:
    """Helper to safely extract the sender's SMTP address."""
    try:
        email_type = getattr(item, 'SenderEmailType', '')
        email_address = getattr(item, 'SenderEmailAddress', '')
        if email_type == "EX":
            sender = getattr(item, 'Sender', None)
            if sender:
                user = sender.GetExchangeUser()
                if user and user.PrimarySmtpAddress:
                    return user.PrimarySmtpAddress
        return email_address or ""
    except Exception:
        try:
            return getattr(item, 'SenderEmailAddress', '')
        except Exception:
            return ""


def _get_recipients_parallel(item):
    """Helper to safely extract resolved To and CC recipients with SMTP."""
    to_list = []
    cc_list = []
    try:
        recipients = getattr(item, 'Recipients', None)
        if recipients:
            for recipient in recipients:
                try:
                    # Get type: 1 = To, 2 = CC
                    rec_type = getattr(recipient, 'Type', 1)
                    name = getattr(recipient, 'Name', '')
                    address = getattr(recipient, 'Address', '')
                    # Resolve SMTP if EX address
                    if address and address.startswith("/o="):
                        ae = getattr(recipient, 'AddressEntry', None)
                        if ae:
                            user = ae.GetExchangeUser()
                            if user and user.PrimarySmtpAddress:
                                address = user.PrimarySmtpAddress
                    recipient_info = {"name": name, "address": address}
                    if rec_type == 1:
                        to_list.append(recipient_info)
                    elif rec_type == 2:
                        cc_list.append(recipient_info)
                except Exception:
                    continue
    except Exception:
        pass
    return to_list, cc_list


def _meeting_status_label(raw: int) -> str:
    """Convert raw MeetingStatus int to string label matching search_common convention."""
    if raw == 3:
        return "meeting_request"
    elif raw in (5, 7):
        return "meeting_canceled"
    elif raw == 1:
        return "meeting"
    return ""


def _extract_email_info_parallel(item_data: Dict[str, Any]) -> Dict[str, Any]:
    """Extract email info from item data in a thread-safe manner."""
    try:
        # Extract basic attributes
        entry_id = item_data.get('EntryID', '')
        subject = item_data.get('Subject', 'No Subject')
        sender = item_data.get('SenderName', 'Unknown')
        sender_email = item_data.get('SenderEmail', '')
        received_time = item_data.get('ReceivedTime', None)
        
        # Extract recipients - handle both formats
        to_recipients = item_data.get('to_recipients', [])
        cc_recipients = item_data.get('cc_recipients', [])
        
        # If recipients are not already extracted, try to extract from To/CC fields
        if not to_recipients and item_data.get('To'):
            to_field = str(item_data.get('To', ''))
            if to_field:
                to_list = to_field.split(';')
                to_recipients = [{"address": addr.strip(), "name": addr.strip()} for addr in to_list if addr.strip()]
        
        if not cc_recipients and item_data.get('CC'):
            cc_field = str(item_data.get('CC', ''))
            if cc_field:
                cc_list = cc_field.split(';')
                cc_recipients = [{"address": addr.strip(), "name": addr.strip()} for addr in cc_list if addr.strip()]
        
        # Extract attachment info
        has_attachments = item_data.get('has_attachments', False)
        attachments = item_data.get('attachments', [])
        embedded_images_count = item_data.get('embedded_images_count', 0)
        embedded_images = item_data.get('embedded_images', [])

        start_time = item_data.get('Start', None)
        start_str = ""
        if start_time:
            try:
                start_str = str(start_time.replace(tzinfo=None))
            except Exception:
                start_str = str(start_time) if start_time else ""

        end_time = item_data.get('End', None)
        end_str = ""
        if end_time:
            try:
                end_str = str(end_time.replace(tzinfo=None))
            except Exception:
                end_str = str(end_time) if end_time else ""

        return {
            "entry_id": entry_id,
            "subject": subject,
            "sender": sender,
            "sender_email": sender_email,
            "received_time": str(received_time.replace(tzinfo=None)) if received_time else "Unknown",
            "start_time": start_str,
            "end_time": end_str,
            "to_recipients": to_recipients,
            "cc_recipients": cc_recipients,
            "has_attachments": has_attachments,
            "attachments": attachments,
            "attachments_count": len(attachments),
            "embedded_images_count": embedded_images_count,
            "embedded_images": embedded_images,
            "body_preview": item_data.get('body_preview', ''),
            "unread": item_data.get('UnRead', False),
            "message_class": item_data.get('MessageClass', ''),
            "meeting_status": _meeting_status_label(item_data.get('MeetingStatus', 0)),
        }
    except Exception as e:
        logger.debug(f"Error in parallel extraction: {e}")
        return {
            "entry_id": item_data.get('EntryID', ''),
            "subject": "No Subject",
            "sender": "Unknown",
            "received_time": "Unknown",
            "to_recipients": [],
            "cc_recipients": [],
            "has_attachments": False,
            "attachments": [],
            "attachments_count": 0,
            "body_preview": "",
            "unread": False
        }

def extract_emails_parallel(items: List[Any], max_workers: int = 4) -> List[Dict[str, Any]]:
    """
    Extract email information from a list of Outlook items using parallel processing.
    
    Args:
        items: List of Outlook MailItem objects
        max_workers: Maximum number of worker threads
        
    Returns:
        List of email dictionaries
    """
    if not items:
        return []
    
    try:
        # Convert items to dictionaries first to avoid COM threading issues
        logger.info(f"Converting {len(items)} items to dictionaries for parallel processing")
        
        item_dicts = []
        for item in items:
            try:
                raw_body = ''
                try:
                    raw_body = getattr(item, 'Body', '') or ''
                except Exception:
                    pass
                try:
                    to_val = getattr(item, 'To', '') or ''
                except Exception:
                    to_val = ''
                try:
                    cc_val = getattr(item, 'CC', '') or ''
                except Exception:
                    cc_val = ''

                start_val = None
                end_val = None
                try:
                    msg_class = getattr(item, 'MessageClass', '') or ''
                    if 'Schedule' in msg_class or getattr(item, 'MeetingStatus', 0):
                        pa = item.PropertyAccessor
                        start_val = pa.GetProperty(_MAPI_START_PROP)
                        end_val = pa.GetProperty(_MAPI_END_PROP)
                except Exception:
                    pass

                to_rec_list, cc_rec_list = _get_recipients_parallel(item)
                item_dict = {
                    'EntryID': getattr(item, 'EntryID', ''),
                    'Subject': getattr(item, 'Subject', 'No Subject'),
                    'SenderName': getattr(item, 'SenderName', 'Unknown'),
                    'SenderEmail': _get_sender_smtp(item),
                    'ReceivedTime': getattr(item, 'ReceivedTime', None),
                    'Start': start_val,
                    'End': end_val,
                    'To': to_val,
                    'CC': cc_val,
                    'to_recipients': to_rec_list,
                    'cc_recipients': cc_rec_list,
                    'UnRead': getattr(item, 'UnRead', False),
                    'MessageClass': getattr(item, 'MessageClass', ''),
                    'MeetingStatus': getattr(item, 'MeetingStatus', 0),
                    'body_preview': raw_body[:200].strip()
                }
                
                # Extract attachment info with embedded image detection
                try:
                    attachments = getattr(item, 'Attachments', None)
                    if attachments:
                        attachments_list = []
                        embedded_images_count = 0
                        embedded_images_list = []

                        for att in attachments:
                            try:
                                file_name = getattr(att, 'FileName', '') or getattr(att, 'DisplayName', 'Unknown')
                                is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico'))

                                # Check if it's an embedded image using multiple methods
                                is_embedded = False

                                # Method 1: Check Content-ID property
                                try:
                                    content_id = getattr(att, 'PropertyAccessor', None)
                                    if content_id:
                                        cid = content_id.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                        if cid and cid.strip():
                                            is_embedded = True
                                except Exception:
                                    pass

                                # Method 2: Check if filename contains CID-like patterns
                                if not is_embedded and is_image:
                                    if 'cid:' in file_name.lower() or file_name.startswith('image'):
                                        is_embedded = True

                                # Method 3: Check attachment type
                                try:
                                    att_type = getattr(att, 'Type', 1)
                                    if att_type == 6:  # Embedded message
                                        is_embedded = True
                                except Exception:
                                    pass

                                # Count embedded images and collect filenames
                                if is_embedded and is_image:
                                    embedded_images_count += 1
                                    embedded_images_list.append({
                                        'name': file_name,
                                        'size': getattr(att, 'Size', 0)
                                    })
                                else:
                                    # Only add non-embedded attachments to the list
                                    attachments_list.append({
                                        'filename': file_name,
                                        'size': getattr(att, 'Size', 0)
                                    })

                            except Exception:
                                continue

                        item_dict['has_attachments'] = len(attachments_list) > 0
                        item_dict['attachments'] = attachments_list
                        item_dict['embedded_images_count'] = embedded_images_count
                        item_dict['embedded_images'] = embedded_images_list
                    else:
                        item_dict['has_attachments'] = False
                        item_dict['attachments'] = []
                        item_dict['embedded_images_count'] = 0
                        item_dict['embedded_images'] = []
                except Exception:
                    item_dict['has_attachments'] = False
                    item_dict['attachments'] = []
                    item_dict['embedded_images_count'] = 0
                    item_dict['embedded_images'] = []
                
                item_dicts.append(item_dict)
            except Exception as e:
                logger.debug(f"Error converting item to dict: {e}")
                continue
        
        logger.info(f"Processing {len(item_dicts)} items in parallel with {max_workers} workers")
        
        # Process items in parallel
        email_list = []
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_item = {executor.submit(_extract_email_info_parallel, item_dict): item_dict 
                             for item_dict in item_dicts}
            
            # Collect results as they complete
            for future in as_completed(future_to_item):
                try:
                    email_data = future.result()
                    if email_data and email_data.get("entry_id"):
                        email_list.append(email_data)
                except Exception as e:
                    logger.debug(f"Error processing item in parallel: {e}")
                    continue
        
        logger.info(f"Parallel extraction completed: {len(email_list)} emails extracted")
        return email_list
        
    except Exception as e:
        logger.error(f"Error in parallel extraction: {e}")
        # Fallback to sequential processing
        return extract_emails_sequential_fallback(items)

def extract_emails_sequential_fallback(items: List[Any]) -> List[Dict[str, Any]]:
    """Optimized sequential extraction for small datasets with minimal overhead."""
    email_list = []
    
    # Pre-allocate list for better performance if size is known
    if hasattr(items, '__len__'):
        email_list = [None] * len(items)
        index = 0
    
    for item in items:
        try:
            # Minimal attribute access with error handling
            entry_id = getattr(item, 'EntryID', '')
            if not entry_id:
                continue
                
            subject = getattr(item, 'Subject', 'No Subject') or 'No Subject'
            sender = getattr(item, 'SenderName', 'Unknown') or 'Unknown'
            sender_email = _get_sender_smtp(item)
            
            received_time = getattr(item, 'ReceivedTime', None)
            received_str = str(received_time.replace(tzinfo=None)) if received_time else "Unknown"
            
            # Extract recipient information via Recipients collection
            to_recipients, cc_recipients = _get_recipients_parallel(item)
            
            # Extract attachment info with embedded image detection
            has_attachments = False
            attachments = []
            embedded_images_count = 0
            embedded_images_list = []
            try:
                attachments_obj = getattr(item, 'Attachments', None)
                if attachments_obj:
                    has_attachments = attachments_obj.Count > 0
                    attachments_list = []

                    for att in attachments_obj:
                        try:
                            file_name = getattr(att, 'FileName', '') or getattr(att, 'DisplayName', 'Unknown')
                            is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico'))

                            # Check if it's an embedded image using multiple methods
                            is_embedded = False

                            # Method 1: Check Content-ID property
                            try:
                                content_id = getattr(att, 'PropertyAccessor', None)
                                if content_id:
                                    cid = content_id.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                    if cid and cid.strip():
                                        is_embedded = True
                            except Exception:
                                pass

                            # Method 2: Check if filename contains CID-like patterns
                            if not is_embedded and is_image:
                                if 'cid:' in file_name.lower() or file_name.startswith('image'):
                                    is_embedded = True

                            # Method 3: Check attachment type
                            try:
                                att_type = getattr(att, 'Type', 1)
                                if att_type == 6:  # Embedded message
                                    is_embedded = True
                            except Exception:
                                pass

                            # Count embedded images and collect filenames
                            if is_embedded and is_image:
                                embedded_images_count += 1
                                embedded_images_list.append({
                                    'name': file_name,
                                    'size': getattr(att, 'Size', 0)
                                })
                            else:
                                # Only add non-embedded attachments to the list
                                attachments_list.append({
                                    'filename': file_name,
                                    'size': getattr(att, 'Size', 0)
                                })

                        except Exception:
                            continue

                    attachments = attachments_list
            except Exception:
                has_attachments = False
                attachments = []
                embedded_images_count = 0
                embedded_images_list = []
            
            # Extract unread status
            unread = getattr(item, 'UnRead', False)

            # Body preview for quick scope judgment
            body_preview = ""
            try:
                raw_body = getattr(item, 'Body', '') or ''
                body_preview = raw_body[:200].strip()
            except Exception:
                pass

            message_class = getattr(item, 'MessageClass', '') or ''
            meeting_status_raw = getattr(item, 'MeetingStatus', 0) or 0

            start_str = ""
            end_str = ""
            try:
                if 'Schedule' in message_class or meeting_status_raw:
                    pa = item.PropertyAccessor
                    start_val = pa.GetProperty(_MAPI_START_PROP)
                    if start_val:
                        start_str = str(start_val.replace(tzinfo=None))
                    end_val = pa.GetProperty(_MAPI_END_PROP)
                    if end_val:
                        end_str = str(end_val.replace(tzinfo=None))
            except Exception:
                pass

            email_data = {
                "entry_id": entry_id,
                "subject": subject,
                "sender": sender,
                "sender_email": sender_email,
                "received_time": received_str,
                "start_time": start_str,
                "end_time": end_str,
                "to_recipients": to_recipients,
                "cc_recipients": cc_recipients,
                "has_attachments": has_attachments,
                "attachments": attachments,
                "attachments_count": len(attachments),
                "embedded_images_count": embedded_images_count,
                "embedded_images": embedded_images_list,
                "body_preview": body_preview,
                "unread": unread,
                "message_class": message_class,
                "meeting_status": _meeting_status_label(meeting_status_raw),
            }
            
            if hasattr(items, '__len__'):
                email_list[index] = email_data
                index += 1
            else:
                email_list.append(email_data)
                
        except Exception:
            # Silent fail for performance - skip problematic items
            continue
    
    # Remove None values if pre-allocation was used
    if hasattr(items, '__len__') and index < len(email_list):
        email_list = email_list[:index]
    
    return email_list

def extract_emails_optimized(items: List[Any], use_parallel: bool = True, max_workers: int = 4) -> List[Dict[str, Any]]:
    """
    Optimized email extraction with automatic fallback and improved small dataset handling.
    
    Args:
        items: List of Outlook MailItem objects
        use_parallel: Whether to use parallel processing
        max_workers: Maximum number of worker threads (if parallel)
        
    Returns:
        List of email dictionaries
    """
    if not items:
        return []
    
    item_count = len(items)
    
    # Optimized thresholds for better performance
    if item_count < 20:  # Very small datasets: sequential is definitely faster
        return extract_emails_sequential_fallback(items)
    elif item_count < 50:  # Small datasets: use sequential with minimal overhead
        return extract_emails_sequential_fallback(items)
    elif item_count < 100:  # Medium datasets: use sequential or light parallel
        return extract_emails_sequential_fallback(items)
    else:  # Large datasets: use parallel processing
        if use_parallel:
            return extract_emails_parallel(items, max_workers)
        else:
            return extract_emails_sequential_fallback(items)