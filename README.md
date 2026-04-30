# Outlook Skill for BrainClaw

A Python-based skill for managing Microsoft Outlook emails through command-line interface. This skill provides comprehensive email operations including search, compose, reply, forward, and folder management.

## Overview

This is a **skill module** for the BrainClaw assistant system that interfaces directly with Microsoft Outlook via COM automation. It uses **message_id** for email identification and operates without caching - all operations work directly with Outlook.

## Key Features

- Email Search: Search by subject, sender, recipient, or body content
- Email Operations: Compose, reply, forward, move, and delete emails
- Batch Operations: Forward emails to multiple recipients via CSV
- Folder Management: Create, list, and manage Outlook folders
- Contact Lookup: Retrieve contact information by email address
- HTML Email Support: All emails use HTML format for rich content

## Requirements

- Operating System: Windows 10 or later
- Microsoft Outlook: 2016 or later (must be running)
- Python: 3.8 or later
- Dependencies: pywin32 and pythoncom for Windows COM automation

## Installation

Install Python dependencies:
```bash
pip install -r config/requirements.txt
```

Ensure Outlook is running - the skill requires Outlook to be open and logged in. Works with both desktop and Microsoft 365 Outlook.

## Usage

All commands are executed through the CLI script located at scripts/outlook_skill.py

Basic command structure:
```bash
py -3 scripts/outlook_skill.py <command> [options]
```

### List Recent Emails

```bash
py -3 scripts/outlook_skill.py list-recent --days 7 --folder "Inbox"
```

Options:
- --days: Number of days to look back (1-30, default: 7)
- --folder: Folder name (default: Inbox)

Output includes message ID, subject, sender, recipients, attachments, and body preview.

### Search Emails

```bash
py -3 scripts/outlook_skill.py search --type subject --query "Meeting" --days 30
py -3 scripts/outlook_skill.py search --type sender --query "John Smith" --days 30
py -3 scripts/outlook_skill.py search --type recipient --query "Jane Doe" --days 30
py -3 scripts/outlook_skill.py search --type body --query "project update" --days 30
```

Options:
- --type: Search type (subject, sender, recipient, body)
- --query: Search term (required)
- --days: Days to look back (1-30, default: 30)
- --folder: Folder to search (default: Inbox)
- --match-all: Match all terms with AND logic (default: true)

### Contact Lookup

```bash
py -3 scripts/outlook_skill.py lookup-contact "user@example.com"
```

Why use this? Outlook MAPI doesn't reliably search by email address. Use this to get the display name, then search by name.

Workflow:
1. Lookup contact to get display name
2. Search by name instead of email address

### Reply to Email

```bash
py -3 scripts/outlook_skill.py reply "<message_id>" "<p>Your HTML reply</p>"
py -3 scripts/outlook_skill.py reply "<message_id>" "<p>Your HTML reply</p>" --send
py -3 scripts/outlook_skill.py reply "<message_id>" "<p>Your HTML reply</p>" --to "user1@example.com" --send
```

Options:
- --send: Actually send the email (default: preview only)
- --to: Override To recipients (comma-separated)
- --cc: Override CC recipients (comma-separated)

### Compose New Email

```bash
py -3 scripts/outlook_skill.py compose --to "user@example.com" --subject "Meeting" --body "<p>Hello,</p><p>Let's meet tomorrow.</p>"
```

Options:
- --to: To recipients (comma-separated, required)
- --subject: Email subject (required)
- --body: Email body in HTML format (required)
- --cc: CC recipients (comma-separated, optional)

Note: Compose always sends immediately (no preview mode).

### Batch Forward

```bash
py -3 scripts/outlook_skill.py batch-forward "<message_id>" "recipients.csv" --message "FYI"
```

CSV Format:
```csv
email
user1@example.com
user2@example.com
user3@example.com
```

Features:
- Uses BCC for privacy (recipients don't see each other)
- Automatically batches in groups of 500 (Outlook limit)
- Optional custom message prepended to forwarded content

### Folder Management

```bash
py -3 scripts/outlook_skill.py list-folders
py -3 scripts/outlook_skill.py create-folder "ProjectX" --parent "Inbox"
py -3 scripts/outlook_skill.py remove-folder "ProjectX"
```

### Email Management

```bash
py -3 scripts/outlook_skill.py move-email "<message_id>" "Archive"
py -3 scripts/outlook_skill.py delete-email "<message_id>"
py -3 scripts/outlook_skill.py get-email "<message_id>"
```

## Architecture

### Project Structure

```
outlook-skill/
├── backend/                      
│   ├── email_search/            
│   │   ├── unified_search.py    
│   │   ├── server_search.py     
│   │   ├── subject_search.py    
│   │   ├── sender_search.py     
│   │   ├── recipient_search.py  
│   │   └── body_search.py       
│   ├── outlook_session/         
│   │   ├── session_manager.py   
│   │   ├── email_operations.py  
│   │   ├── folder_operations.py 
│   │   └── contact_operations.py
│   ├── config.py                
│   ├── validation.py            
│   └── email_composition.py     
├── tools/                       
│   ├── search_tools.py          
│   ├── email_operations.py      
│   ├── folder_tools.py          
│   └── registration.py          
├── scripts/
│   └── outlook_skill.py         
└── config/
    └── requirements.txt         
```

### Key Design Principles

1. No Caching: All operations work directly with Outlook COM objects
2. Message ID Based: Uses Outlook's EntryID for email identification
3. Server-Side Search: Prioritizes fast server-side search over client-side
4. Batch Processing: Handles large result sets efficiently
5. HTML Format: All emails use HTML for rich formatting

### Technical Details

- COM Automation: Uses win32com.client for Outlook integration
- Session Management: Context manager pattern for proper resource cleanup
- Error Handling: Comprehensive error handling with retry logic
- Threading: Proper COM threading initialization with pythoncom
- Memory Management: Periodic COM cache clearing to prevent memory growth

## HTML Email Format

All email bodies must be in HTML format.

Simple Email:
```html
<p>Hello,</p>
<p>This is a simple message.</p>
<p>Best regards,<br>Your Name</p>
```

Formatted Email:
```html
<p>Dear Team,</p>
<p>Please review the following:</p>
<ul>
  <li><strong>Item 1</strong>: Description</li>
  <li><strong>Item 2</strong>: Description</li>
</ul>
<p>Thank you,<br>Manager</p>
```

## Search Tips

### Searching by Email Address

Outlook MAPI doesn't reliably search by email address. Use this workflow:

1. Lookup contact first to get display name
2. Use display name in search

### Search Performance

- Subject/Sender/Recipient: Fast (server-side search)
- Body content: Slower (requires loading full email content)
- Limit days: Use smaller day ranges for faster results

### Match Logic

- --match-all true (default): Requires ALL terms to match (AND logic)
- --match-all false: Matches ANY term (OR logic)

## Configuration

Configuration is centralized in backend/config.py:

- Search limits: Max 30 days lookback
- Batch sizes: Optimized for performance
- Display settings: Text truncation, date formats
- Outlook constants: COM object types and folder IDs

## Troubleshooting

### Common Issues

1. "Failed to connect to Outlook"
   - Ensure Outlook is running and logged in
   - Check Windows COM permissions

2. "Email has been moved or deleted"
   - The message_id is stale
   - Search for the email again to get current ID

3. "No recipients found"
   - Verify the original email has valid recipients
   - Check if you're replying to a sent item correctly

4. Search returns no results
   - Try broader search terms
   - Increase the --days parameter
   - Verify the folder name is correct

### Debug Mode

Enable detailed logging by modifying backend/logging_config.py

## API Reference

### Email Data Structure

Each email returned contains:

```python
{
    "id": "message_id",
    "subject": "Email subject",
    "sender": "Sender Name",
    "sender_email": "sender@example.com",
    "received_time": "2024-01-01 12:00:00",
    "to_recipients": [
        {"name": "Recipient Name", "address": "recipient@example.com"}
    ],
    "cc_recipients": [],
    "has_attachments": true,
    "attachments": [
        {"name": "file.pdf", "size": 102400}
    ],
    "attachments_count": 1,
    "embedded_images_count": 0
}
```

## Privacy & Security

- 100% Local Processing: All email operations happen on your computer
- No Cloud Services: Works entirely offline with local Outlook installation
- Secure by Design: Uses your existing Outlook installation and credentials
- No Data Leaves Machine: Email content never sent to external servers
- Windows COM API: Direct integration with Outlook via win32com

## Performance

- Fast Search: 100 emails loaded in approximately 2 seconds
- Real-time: Searches complete instantly
- Batch Operations: Forward to 100+ recipients in 2 minutes
- Local Processing: No network latency

## Related Files

- SKILL.md - Quick reference guide
- scripts/outlook_skill.py - Main CLI implementation
- backend/config.py - Configuration settings
- tools/registration.py - MCP tool registration

## Contributing

This skill is part of the BrainClaw assistant system. For contributions:

1. Follow the existing code structure
2. Add proper error handling
3. Update documentation for new features
4. Test with various Outlook configurations

## License

Part of the BrainClaw assistant system.

---

Note: This skill operates directly with Outlook COM objects without caching. All email operations use message_id (Outlook EntryID) for identification.