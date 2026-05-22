"""
Contact operations for Outlook.

This module provides functions to interact with Outlook contacts,
including looking up display names from email addresses.
"""

from typing import Optional, Dict, Any, List
from ..logging_config import get_logger
from .session_manager import OutlookSessionManager

logger = get_logger(__name__)


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
                            'first_name': contact.FirstName,
                            'last_name': contact.LastName,
                            'company': contact.CompanyName,
                            'job_title': contact.JobTitle
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
                        return {
                            'display_name': exchange_user.Name,
                            'email': exchange_user.PrimarySmtpAddress,
                            'first_name': exchange_user.FirstName,
                            'last_name': exchange_user.LastName,
                            'company': exchange_user.CompanyName,
                            'job_title': exchange_user.JobTitle
                        }
                    else:
                        return {
                            'display_name': address_entry.Name,
                            'email': email_address,
                            'first_name': None,
                            'last_name': None,
                            'company': None,
                            'job_title': None
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

    Args:
        display_name: The display name to search for (e.g., "HONG YANG")

    Returns:
        List of matching contacts. Empty list if none found.
    """
    try:
        with OutlookSessionManager() as session:
            recipient = session.outlook.Session.CreateRecipient(display_name)
            recipient.Resolve()

            if recipient.Resolved:
                address_entry = recipient.AddressEntry
                exchange_user = address_entry.GetExchangeUser()

                if exchange_user:
                    return [{
                        'display_name': exchange_user.Name,
                        'email': exchange_user.PrimarySmtpAddress,
                        'first_name': exchange_user.FirstName,
                        'last_name': exchange_user.LastName,
                        'company': exchange_user.CompanyName,
                        'job_title': exchange_user.JobTitle
                    }]
                else:
                    return [{
                        'display_name': address_entry.Name,
                        'email': address_entry.Address,
                        'first_name': None,
                        'last_name': None,
                        'company': None,
                        'job_title': None
                    }]

            # Ambiguous or not found — search GAL directly
            namespace = session.outlook_namespace
            gal = None
            for al in namespace.AddressLists:
                if "Global Address List" in al.Name:
                    gal = al
                    break

            if not gal:
                return []

            results = []
            search_upper = display_name.upper()
            try:
                entries = gal.AddressEntries
                entry = entries.Item(display_name)
                while entry:
                    if search_upper in entry.Name.upper():
                        try:
                            exchange_user = entry.GetExchangeUser()
                            if exchange_user:
                                results.append({
                                    'display_name': exchange_user.Name,
                                    'email': exchange_user.PrimarySmtpAddress,
                                    'first_name': exchange_user.FirstName,
                                    'last_name': exchange_user.LastName,
                                    'company': exchange_user.CompanyName,
                                    'job_title': exchange_user.JobTitle
                                })
                        except Exception:
                            pass
                    else:
                        break
                    entry = entries.GetNext()
            except Exception as e:
                logger.debug(f"GAL search error: {e}")

            return results

    except Exception as e:
        logger.error(f"Error looking up contact by name: {e}")
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