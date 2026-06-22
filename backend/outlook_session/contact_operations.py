"""
Contact operations for Outlook.

This module provides functions to interact with Outlook contacts,
including looking up display names from email addresses.
"""

from typing import Optional, Dict, Any, List
from ..logging_config import get_logger
from .session_manager import OutlookSessionManager

logger = get_logger(__name__)


def _extract_exchange_user_info(exchange_user) -> Dict[str, Any]:
    """Extract all available fields from an ExchangeUser COM object."""
    def _safe_get(attr):
        try:
            val = getattr(exchange_user, attr, None)
            return val if val else None
        except Exception:
            return None

    return {
        'display_name': _safe_get('Name'),
        'email': _safe_get('PrimarySmtpAddress'),
        'first_name': _safe_get('FirstName'),
        'last_name': _safe_get('LastName'),
        'company': _safe_get('CompanyName'),
        'job_title': _safe_get('JobTitle'),
        'department': _safe_get('Department'),
        'office': _safe_get('OfficeLocation'),
        'business_phone': _safe_get('BusinessTelephoneNumber'),
        'mobile_phone': _safe_get('MobileTelephoneNumber'),
        'city': _safe_get('City'),
        'state': _safe_get('StateOrProvince'),
        'alias': _safe_get('Alias'),
    }


def get_contact_by_email(email_address: str) -> Optional[Dict[str, Any]]:
    """
    Look up a contact by email address and return their information.
    
    Args:
        email_address: The email address to search for
        
    Returns:
        Dictionary with contact info (display_name, email, etc.) or None if not found
    """
    try:
        with OutlookSessionManager() as session:
            namespace = session.outlook_namespace
            contacts_folder = namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
            
            # Search for contact with this email address
            # Try multiple email fields (Email1Address, Email2Address, Email3Address)
            for i in range(1, 4):
                try:
                    filter_str = f"[Email{i}Address] = '{email_address}'"
                    contact = contacts_folder.Items.Find(filter_str)
                    
                    if contact:
                        return {
                            'display_name': contact.FullName or contact.FileAs,
                            'email': email_address,
                            'first_name': getattr(contact, 'FirstName', None),
                            'last_name': getattr(contact, 'LastName', None),
                            'company': getattr(contact, 'CompanyName', None),
                            'job_title': getattr(contact, 'JobTitle', None),
                            'department': getattr(contact, 'Department', None),
                            'office': getattr(contact, 'OfficeLocation', None),
                            'business_phone': getattr(contact, 'BusinessTelephoneNumber', None),
                            'mobile_phone': getattr(contact, 'MobileTelephoneNumber', None),
                            'city': getattr(contact, 'BusinessAddressCity', None),
                            'state': getattr(contact, 'BusinessAddressState', None),
                            'alias': None,
                        }
                except Exception as e:
                    logger.debug(f"Error searching Email{i}Address: {e}")
                    continue
            
            # If not found in contacts, try resolving via Exchange
            try:
                recipient = session.outlook.Session.CreateRecipient(email_address)
                recipient.Resolve()

                if recipient.Resolved:
                    address_entry = recipient.AddressEntry
                    exchange_user = address_entry.GetExchangeUser()

                    if exchange_user:
                        return _extract_exchange_user_info(exchange_user)
                    else:
                        return {
                            'display_name': address_entry.Name,
                            'email': email_address,
                            'first_name': None,
                            'last_name': None,
                            'company': None,
                            'job_title': None,
                            'department': None,
                            'office': None,
                            'business_phone': None,
                            'mobile_phone': None,
                            'city': None,
                            'state': None,
                            'alias': None,
                        }
            except Exception as e:
                logger.debug(f"Error resolving via Exchange: {e}")
            
            return None
            
    except Exception as e:
        logger.error(f"Error looking up contact: {e}")
        return None


def get_contact_by_name(display_name: str) -> List[Dict[str, Any]]:
    """
    Look up contacts by display name via Exchange GAL.

    Always searches the GAL to return ALL matching contacts,
    not just the first resolved match.

    Args:
        display_name: The display name to search for (e.g., "HONG YANG")

    Returns:
        List of matching contacts. Empty list if none found.
    """
    try:
        with OutlookSessionManager() as session:
            namespace = session.outlook_namespace
            gal = None
            for al in namespace.AddressLists:
                if "Global Address List" in al.Name:
                    gal = al
                    break

            if not gal:
                logger.debug("GAL not found, falling back to single resolve")
                return _resolve_single_contact(session, display_name)

            results = []
            seen_emails = set()
            search_upper = display_name.upper()
            entries = gal.AddressEntries

            def _collect_from_position(start_name):
                """Jump to a GAL position and collect consecutive matches."""
                try:
                    entry = entries.Item(start_name)
                except Exception:
                    return
                while entry:
                    if search_upper in entry.Name.upper():
                        try:
                            exchange_user = entry.GetExchangeUser()
                            if exchange_user:
                                info = _extract_exchange_user_info(exchange_user)
                                email_key = (info.get('email') or '').lower()
                                if email_key and email_key not in seen_emails:
                                    seen_emails.add(email_key)
                                    results.append(info)
                        except Exception:
                            pass
                    else:
                        break
                    entry = entries.GetNext()

            _collect_from_position(display_name)
            prev_count = len(results)
            max_iterations = 50
            iteration = 0
            while iteration < max_iterations:
                iteration += 1
                if not results and prev_count == 0:
                    break
                last_name = results[-1]['display_name'] if results else display_name
                next_prefix = last_name + "a"
                _collect_from_position(next_prefix)
                if len(results) == prev_count:
                    break
                prev_count = len(results)

            if results:
                return results

            return _resolve_single_contact(session, display_name)

    except Exception as e:
        logger.error(f"Error looking up contact by name: {e}")
        return []


def _resolve_single_contact(session, display_name: str) -> List[Dict[str, Any]]:
    """Fallback: resolve a single contact via CreateRecipient."""
    try:
        recipient = session.outlook.Session.CreateRecipient(display_name)
        recipient.Resolve()

        if recipient.Resolved:
            address_entry = recipient.AddressEntry
            exchange_user = address_entry.GetExchangeUser()

            if exchange_user:
                return [_extract_exchange_user_info(exchange_user)]
            else:
                return [{
                    'display_name': address_entry.Name,
                    'email': address_entry.Address,
                    'first_name': None,
                    'last_name': None,
                    'company': None,
                    'job_title': None,
                    'department': None,
                    'office': None,
                    'business_phone': None,
                    'mobile_phone': None,
                    'city': None,
                    'state': None,
                    'alias': None,
                }]
    except Exception as e:
        logger.debug(f"Single resolve fallback error: {e}")

    return []


def get_display_name_from_email(email_address: str) -> Optional[str]:
    """
    Get the display name for an email address.

    This function tries multiple methods:
    1. Look up in Outlook Contacts
    2. Resolve via Exchange Global Address List

    Args:
        email_address: The email address to look up

    Returns:
        Display name if found, None otherwise
    """
    contact_info = get_contact_by_email(email_address)
    return contact_info['display_name'] if contact_info else None