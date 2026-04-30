"""Email viewing tools for Outlook MCP Server - Simplified without cache."""

# Type imports
from typing import Any, Dict, Optional

# Local application imports
from backend.outlook_session import OutlookSessionManager
from backend.validation import (
    ValidationError,
    validate_days_parameter,
    validate_folder_name
)


def load_emails_by_folder_tool(folder_path: str, days: int = None, max_emails: int = None) -> Dict[str, Any]:
    """Load emails from a specific folder and return them directly with message_id.

    **LLM Note**: This function enforces strict mutual exclusion between 'days' and 'max_emails' parameters.
    You CANNOT use both parameters together. Choose either time-based loading (days) or number-based loading (max_emails).

    Args:
        folder_path: Path to the folder (supports nested paths like "user@company.com/Inbox/SubFolder1")
        days: Number of days to look back (max: 30) - mutually exclusive with max_emails
        max_emails: Maximum number of emails to load (mutually exclusive with days)

    Returns:
        dict: Response containing list of emails with message_id:
        {
            "type": "json",
            "data": {
                "count": 5,
                "message": "Found 5 emails...",
                "emails": [
                    {
                        "entry_id": "...",
                        "subject": "...",
                        "sender": "...",
                        ...
                    }
                ]
            }
        }

    Note:
        IMPORTANT: Folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
        
        Usage examples:
        - Time-based: load_emails_by_folder_tool("Inbox", days=7)
        - Number-based: load_emails_by_folder_tool("Inbox", max_emails=50)
        - Cannot use both: load_emails_by_folder_tool("Inbox", days=7, max_emails=50) - raises error
    """
    try:
        validate_folder_name(folder_path)
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    
    # Enforce mutual exclusion
    if days is not None and max_emails is not None:
        return {"type": "text", "text": "Cannot specify both 'days' and 'max_emails' parameters. Use either time-based (days) or number-based (max_emails) loading, not both."}
    
    # Set default behavior
    if days is None and max_emails is None:
        days = 7
    
    # Validate parameters
    try:
        if days is not None:
            validate_days_parameter(days)
        
        if max_emails is not None and (not isinstance(max_emails, int) or max_emails < 1):
            raise ValidationError("max_emails must be a positive integer when specified")
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    
    try:
        # Determine max_emails based on parameters
        if max_emails is not None:
            actual_max_emails = min(max_emails, 1000)  # Cap at 1000
        else:
            actual_max_emails = 10000  # High limit for time-based search
        
        with OutlookSessionManager() as outlook_session:
            email_list, message = outlook_session.get_folder_emails(
                folder_path, 
                actual_max_emails, 
                fast_mode=True, 
                days_filter=days if max_emails is None else None
            )
            
            # Return emails directly in JSON format
            return {
                "type": "json",
                "data": {
                    "count": len(email_list),
                    "message": message,
                    "emails": email_list
                }
            }
    except Exception as e:
        return {"type": "text", "text": f"Error loading emails from folder: {str(e)}"}