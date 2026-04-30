"""
Email operations for Outlook session management.

This module provides email-related operations such as moving emails between folders,
managing email policies, and retrieving email details.

Note: This module has been simplified. Most operations are now handled directly
in tools/email_operations.py using message_id instead of cache-based email_number.
"""

# Type imports
from typing import Any, Dict

# Local application imports
from ..logging_config import get_logger
from ..outlook_session.session_manager import OutlookSessionManager
from .exceptions import InvalidParameterError, OperationFailedError

logger = get_logger(__name__)


class EmailOperations:
    """Handles all email-related operations for Outlook."""
    
    def __init__(self, session_manager):
        """Initialize with a session manager instance."""
        self.session_manager = session_manager

    def get_email_by_id(self, message_id: str) -> Dict[str, Any]:
        """
        Get email details by message_id.
        
        Args:
            message_id: The Outlook entry_id of the email
        
        Returns:
            Email dictionary with full details
        """
        try:
            if not message_id or not isinstance(message_id, str):
                raise ValueError("message_id must be a non-empty string")
            
            item = self.session_manager.namespace.GetItemFromID(message_id)
            if not item:
                raise ValueError(f"Could not find email with message_id: {message_id}")
            
            # Extract basic email data
            email_data = {
                "entry_id": message_id,
                "subject": getattr(item, 'Subject', 'No Subject'),
                "sender": getattr(item, 'SenderName', 'Unknown'),
                "received_time": str(getattr(item, 'ReceivedTime', 'Unknown')),
                "body": getattr(item, 'Body', ''),
            }
            
            logger.info(f"Retrieved email: {email_data.get('subject', 'No Subject')}")
            return email_data
            
        except Exception as e:
            error_msg = f"Error getting email by message_id: {e}"
            logger.error(error_msg)
            raise ValueError(error_msg)


def get_email_by_id(message_id: str) -> Dict[str, Any]:
    """Get email details by message_id."""
    with OutlookSessionManager() as session_manager:
        email_ops = EmailOperations(session_manager)
        return email_ops.get_email_by_id(message_id)
