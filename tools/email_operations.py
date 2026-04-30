"""Email operations tools for Outlook MCP Server."""

from typing import Dict, Any, Union, List, Optional
from backend.email_composition import reply_to_email_by_message_id, compose_email
from backend.outlook_session import OutlookSessionManager
from backend.validation import ValidationError


def reply_to_email_tool(
    message_id: str,
    reply_text: str,
    to_recipients: Union[str, List[str], None] = None,
    cc_recipients: Union[str, List[str], None] = None
) -> Dict[str, Any]:
    """Reply to an email using its message_id with custom recipients if provided

    Args:
        message_id: The Outlook entry_id of the email to reply to
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]

    Behavior:
        - When both to_recipients and cc_recipients are None:
          * Uses ReplyAll() to maintain original recipients
        - When either parameter is provided:
          * Uses Reply() with specified recipients
          * Any None parameters will result in empty recipient fields
        - Single email strings and lists of email strings are both accepted

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not message_id or not isinstance(message_id, str):
        raise ValidationError("message_id must be a non-empty string")
    if not reply_text or not isinstance(reply_text, str):
        raise ValidationError("Reply text must be a non-empty string")
    
    try:
        result = reply_to_email_by_message_id(message_id, reply_text, to_recipients, cc_recipients)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error replying to email: {str(e)}"}


def compose_email_tool(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> Dict[str, Any]:
    """Compose and send a new email

    Args:
        recipient_email: Email address(es) of the recipient(s) - can be single email or semicolon-separated list
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Optional CC email address(es) - can be single email or semicolon-separated list

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not recipient_email or not isinstance(recipient_email, str):
        raise ValidationError("Recipient email must be a non-empty string")
    if not subject or not isinstance(subject, str):
        raise ValidationError("Subject must be a non-empty string")
    if not body or not isinstance(body, str):
        raise ValidationError("Body must be a non-empty string")
    
    try:
        # Parse semicolon-separated email addresses into lists
        to_recipients = [email.strip() for email in recipient_email.split(';') if email.strip()]
        cc_recipients = None
        if cc_email:
            cc_recipients = [email.strip() for email in cc_email.split(';') if email.strip()]
        
        result = compose_email(to_recipients, subject, body, cc_recipients)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error composing email: {str(e)}"}


def move_email_tool(message_id: str, target_folder_name: str) -> Dict[str, Any]:
    """Move an email to the specified folder using its message_id.

    Args:
        message_id: The Outlook entry_id of the email to move
        target_folder_name: Name or path of the target folder (supports nested paths like "user@company.com/Inbox/SubFolder1/SubFolder2")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved successfully to target_folder"
        }

    Note:
        IMPORTANT: Target folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
    """
    if not message_id or not isinstance(message_id, str):
        raise ValidationError("message_id must be a non-empty string")
    if not target_folder_name or not isinstance(target_folder_name, str):
        raise ValidationError("Target folder name must be a non-empty string")

    try:
        with OutlookSessionManager() as session:
            # Get the email item
            item = session.namespace.GetItemFromID(message_id)
            if not item:
                return {"type": "text", "text": f"Error: Could not find email with message_id {message_id}"}
            
            # Get target folder
            target_folder = session.get_folder(target_folder_name)
            if not target_folder:
                return {"type": "text", "text": f"Error: Target folder '{target_folder_name}' not found"}
            
            # Move the email
            item.Move(target_folder)
            return {"type": "text", "text": f"Email moved successfully to '{target_folder_name}'"}
    except Exception as e:
        return {"type": "text", "text": f"Error moving email: {str(e)}"}


def delete_email_tool(message_id: str) -> Dict[str, Any]:
    """Move an email to the Deleted Items folder using its message_id.

    Args:
        message_id: The Outlook entry_id of the email to delete

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved to Deleted Items successfully"
        }

    Note:
        This tool moves the email to the Deleted Items folder instead of permanently deleting it.
    """
    if not message_id or not isinstance(message_id, str):
        raise ValidationError("message_id must be a non-empty string")

    try:
        # Move to Deleted Items folder
        return move_email_tool(message_id, "Deleted Items")
    except Exception as e:
        return {"type": "text", "text": f"Error deleting email: {str(e)}"}