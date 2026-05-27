# Outlook Skill

A Python-based skill for managing Microsoft Outlook emails through command-line interface. This skill provides comprehensive email operations including search, compose, reply, forward, and folder management.

## Overview

This is a **universal skill module** that can be integrated with any AI assistant system. It interfaces directly with Microsoft Outlook via COM automation, uses **message_id** for email identification, and operates without caching - all operations work directly with Outlook.

## Key Features

- Email Search: Search by subject, sender, recipient, or body content
- Thread Tracking: Find all emails in the same conversation (even when subject changes)
- Related Email Discovery: Multi-strategy search by sender, recipients, keywords
- Email Operations: Compose, reply, forward, redirect, move, and delete emails
- Batch Operations: Forward emails to multiple recipients via CSV
- Folder Management: Create, list, and manage Outlook folders
- Contact Lookup: Retrieve contact information by email or display name
- Attachment Download: Save attachments to local directory
- HTML Email Support: All emails use HTML format for rich content

## Requirements

- Operating System: Windows 10 or later
- Microsoft Outlook: 2016 or later (must be running)
- Python: 3.8 or later
- Dependencies: pywin32 and pythoncom for Windows COM automation

## Installation

Install Python dependencies:

```bash
pip install -r requirements.txt
```

Ensure Outlook is running - the skill requires Outlook to be open and logged in. Works with both desktop and Microsoft 365 Outlook.

## Integration with AI Systems

This skill can be integrated with various AI assistant systems:

- **BrainClaw**: Native integration as a skill module
- **Custom AI Systems**: Direct CLI integration or Python API
- **Other Assistants**: Adaptable to any system that can execute Python scripts

The skill provides CLI interface and programmatic API for flexible integration.

## Usage

All commands are executed through the CLI script located at scripts/outlook_skill.py

Basic command structure:

```bash
py -3 scripts/outlook_skill.py <command> [options]
```

### Find Recent Emails

```bash
py -3 scripts/outlook_skill.py find-recent --days 7
py -3 scripts/outlook_skill.py find-recent --days 7 --folder "Inbox"
```

Options:

- --days: Number of days to look back (1-365, default: 7)
- --folder: Single folder override (default: Inbox + Sent Items)

Output includes email ID, subject, sender, recipients, attachments, folder indicator, and body preview.

### Find Emails (Search)

```bash
py -3 scripts/outlook_skill.py find --type subject --query "Meeting" --days 14
py -3 scripts/outlook_skill.py find --type sender --query "John Smith" --days 14
py -3 scripts/outlook_skill.py find --type recipient --query "Jane Doe" --days 14
py -3 scripts/outlook_skill.py find --type body --query "project update" --days 14
```

Options:

- --type: Search type (subject, sender, recipient, body) — required
- --query: Search term — required
- --days: Days to look back (1-365, default: 14)
- --folder: Single folder override
- --folders: Comma-separated folder names for cross-folder search
- --match-all: Match all terms with AND logic (default: true)

Default folder behavior by search type:

- `subject` / `body` → Inbox + Sent Items (both folders)
- `sender` → Inbox only
- `recipient` → Sent Items only

### Find Thread (Conversation Tracking)

Given an email, find all emails in the same conversation — even if subjects changed (RE:/FW: prefixes, topic drift).

```bash
py -3 scripts/outlook_skill.py find-thread "<email_id>"
py -3 scripts/outlook_skill.py find-thread "<email_id>" --fuzzy
py -3 scripts/outlook_skill.py find-thread "<email_id>" --brief
```

Options:

- --folders: Override search folders (default: Inbox + Sent Items)
- --fuzzy: Also find emails with similar subjects when conversation ID breaks (e.g., cross-tenant forwards)
- --brief: Compact single-line output (still shows email ID)

Output includes a thread summary: message count, participants, and date span.

**How it works:**

1. Reads the conversation ID from the given email (Outlook assigns the same ID to all replies in a thread)
2. Searches Inbox + Sent Items for all emails with that same conversation ID
3. With `--fuzzy`: also finds emails where subject keywords overlap ≥ 60% within ±7 days (catches broken threads)
4. Results sorted chronologically (oldest first)

### Find Related Emails (Multi-Strategy Discovery)

Given an email, find all related emails across multiple dimensions — not just the same thread, but same sender, same group of people, or same topic.

```bash
py -3 scripts/outlook_skill.py find-related "<email_id>"
py -3 scripts/outlook_skill.py find-related "<email_id>" --exclude-thread
py -3 scripts/outlook_skill.py find-related "<email_id>" --strategies sender,keyword
py -3 scripts/outlook_skill.py find-related "<email_id>" --max 10 --brief
```

Options:

- --days: Lookback window for sender/recipient/keyword strategies (default: 90)
- --strategies: Comma-separated list (default: all four — thread, sender, recipient, keyword)
- --exclude-thread: Skip thread strategy (useful after find-thread to avoid duplicates)
- --max: Limit results returned (default: 20, configurable in config.py)
- --brief: Compact single-line output (still shows email ID)

**How it works — 4 strategies ranked by relevance:**

| Strategy  | What it finds                                           | Relevance |
| --------- | ------------------------------------------------------- | --------- |
| thread    | Same conversation ID                                    | ★★★★★     |
| sender    | Same sender + overlapping topic keywords                | ★★★★      |
| recipient | ≥ 2 shared recipients in To/CC (same group discussing)  | ★★★       |
| keyword   | Shared meaningful keywords from subject + body          | ★★★       |

Results are merged, deduplicated, and sorted by relevance then time.

**Performance:** Uses a single merged scan per folder (not separate scans per strategy), so it's fast even with all 4 strategies enabled.

### Contact Lookup

```bash
py -3 scripts/outlook_skill.py lookup-contact "user@example.com"
py -3 scripts/outlook_skill.py lookup-contact "HONG YANG"
```

Accepts email address or display name. Auto-detects format by presence of `@`.

Returns: Display name, email, alias, company, department, job title, office, phone, mobile, location.

Why use this? Outlook MAPI doesn't reliably search by email address. Use this to get the display name, then search by name.

### Get Full Email Details

```bash
py -3 scripts/outlook_skill.py get-email "<email_id>"
```

Returns complete email content: full HTML body, all attachments, metadata. Embedded images are auto-extracted to a temp directory with paths printed in output.

### ReplyAll (Default Reply)

```bash
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>HTML body</p>"
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>HTML body</p>" --cc "extra@example.com"
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>HTML body</p>" --attach "C:\path\file.pdf"
```

Keeps ALL original To + CC recipients. `--to`/`--cc` APPEND to existing.

Options:

- --to: Additional To recipients (comma separated)
- --cc: Additional CC recipients (comma separated)
- --attach: File path(s) to attach (comma separated for multiple)
- --inline-image: Embed image inline (format: `filepath:cid_name`, comma separated)

### Reply (Sender Only)

```bash
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>"
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>" --to "specific@example.com"
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>" --attach "C:\path\file.pdf"
```

Replies to sender only. `--to`/`--cc` specify EXACT extra recipients (original To/CC NOT included).

Options:

- --to: Extra To recipients (comma separated)
- --cc: Extra CC recipients (comma separated)
- --attach: File path(s) to attach (comma separated)
- --inline-image: Embed image inline (format: `filepath:cid_name`, comma separated)

### Compose New Email

```bash
py -3 scripts/outlook_skill.py compose --to "user@example.com" --subject "Meeting" --body "<p>Hello</p>"
py -3 scripts/outlook_skill.py compose --to "user@example.com" --subject "Report" --body "<p>See attached</p>" --attach "C:\path\file.pdf"
py -3 scripts/outlook_skill.py compose --to "user@example.com" --subject "Photo" --body "<p><img src='cid:pic1'></p>" --inline-image "C:\path\img.png:pic1"
```

Options:

- --to: To recipients (comma-separated, required)
- --subject: Email subject (required)
- --body: Email body in HTML format (required)
- --cc: CC recipients (comma-separated)
- --attach: File path(s) to attach (comma separated)
- --inline-image: Embed image inline (format: `filepath:cid_name`, comma separated)

Sends immediately when called.

### Forward (Single)

```bash
py -3 scripts/outlook_skill.py forward "<email_id>" --to "user@example.com"
py -3 scripts/outlook_skill.py forward "<email_id>" --to "user1@example.com,user2@example.com" --cc "manager@example.com" --body "<p>FYI</p>"
py -3 scripts/outlook_skill.py forward "<email_id>" --to "user@example.com" --attach "C:\path\extra.pdf"
```

Options:

- --to: To recipients (comma-separated, required)
- --cc: CC recipients (comma-separated)
- --body: Custom HTML message to prepend
- --attach: Additional file path(s) to attach (comma separated)
- --inline-image: Embed image inline (format: `filepath:cid_name`, comma separated)

Preserves original email formatting. Subject auto-prefixed with `FW:`.

### Redirect (Clear Recipients + New TO/CC)

```bash
py -3 scripts/outlook_skill.py redirect "<email_id>" "<p>New message</p>" --to "a@example.com,b@example.com"
py -3 scripts/outlook_skill.py redirect "<email_id>" "<p>FYI</p>" --to "a@example.com" --cc "b@example.com"
```

Clears all existing TO and CC recipients, then adds new ones. Preserves original email body as quoted content.

Options:

- body: HTML message prepended above original content (required)
- --to: New TO recipients (comma separated, required)
- --cc: New CC recipients (comma separated)
- --attach: File path(s) to attach (comma separated)
- --inline-image: Embed image inline (format: `filepath:cid_name`, comma separated)

### Batch Forward

```bash
py -3 scripts/outlook_skill.py batch-forward "<email_id>" "recipients.csv" --message "<p>FYI</p>"
```

CSV Format (single column named "email", supports BOM encoding):

```csv
email
user1@example.com
user2@example.com
user3@example.com
```

Features:

- Uses BCC for privacy (recipients don't see each other)
- Automatically splits into batches of 500 (configurable in config.py)
- Optional HTML message prepended to forwarded content

### Download Attachment

```bash
py -3 scripts/outlook_skill.py download-attachment "<email_id>"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --output-dir "C:\temp"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --filename "report.pdf"
```

Options:

- --output-dir: Directory to save attachments (default: ~/Downloads)
- --filename: Download only a specific attachment by name
- --all: Include embedded images (default: skips inline images, keeps PDFs/docs)

### Inline Images

Embed images directly in email body using Content-ID (CID):

```bash
py -3 scripts/outlook_skill.py compose --to "user@example.com" --subject "Report" \
  --body "<p>See chart: <img src='cid:chart1'></p>" \
  --inline-image "C:\path\chart.png:chart1"
```

- Format: `filepath:cid_name` (comma separated for multiple)
- Reference in HTML: `<img src="cid:cid_name">`
- Works with: compose, reply, replyall, forward, redirect

### Folder Management

```bash
py -3 scripts/outlook_skill.py create-folder "ProjectX" --parent "Inbox"
py -3 scripts/outlook_skill.py remove-folder "ProjectX"
```

### Email Management

```bash
py -3 scripts/outlook_skill.py move-email "<email_id>" "Archive"
py -3 scripts/outlook_skill.py delete-email "<email_id>"
```

## Architecture

### Project Structure

```text
outlook-skill/
├── backend/
│   ├── email_search/
│   │   ├── unified_search.py     # Public search API (find, find-thread, find-related)
│   │   ├── server_search.py      # Core search engine (Restrict, merged scan, strategies)
│   │   ├── search_common.py      # Shared utilities (extract_email_info, folder helpers)
│   │   ├── email_listing.py      # Recent email listing
│   │   ├── parallel_extractor.py # Parallel email data extraction
│   │   ├── subject_search.py
│   │   ├── sender_search.py
│   │   ├── recipient_search.py
│   │   └── body_search.py
│   ├── outlook_session/
│   │   ├── session_manager.py    # COM session lifecycle (context manager)
│   │   ├── email_operations.py   # Send, reply, forward operations
│   │   ├── folder_operations.py  # Folder CRUD
│   │   └── contact_operations.py # Contact/GAL lookup
│   ├── config.py                 # All configuration constants
│   ├── validation.py             # Input validation
│   ├── email_composition.py      # Email building helpers
│   └── logging_config.py         # Logging setup
├── scripts/
│   └── outlook_skill.py          # CLI entry point
├── SKILL.md                      # Quick reference (AI-oriented)
├── README.md                     # Full documentation
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
```

Formatted Email:

```html
<p>Dear Team,</p>
<p>Please review the following:</p>
<ul>
  <li><strong>Item 1</strong>: Description</li>
  <li><strong>Item 2</strong>: Description</li>
</ul>
```

**Special characters:** Replace `$` with `&#36;` in HTML body to avoid shell variable interpolation.

```html
<!-- Wrong: $80,000 becomes ,000 -->
<p>Cost: $80,000 USD</p>

<!-- Correct: use HTML entity -->
<p>Cost: &#36;80,000 USD</p>
```

Common entities: `$` = `&#36;` | `&` = `&amp;` | `<` = `&lt;` | `>` = `&gt;`

## Search Tips

### Searching by Email Address

Outlook MAPI doesn't reliably search by email address. Use this workflow:

1. `lookup-contact "user@example.com"` → get display name
2. `find --type sender --query "Display Name"` → search by name

### Search Performance

- Subject/Sender/Recipient: Fast (server-side Restrict filter)
- Body content: Uses subject pre-filter first, then validates body match (configurable read limit)
- find-related: Single merged scan per folder for all strategies
- find-thread: Tries Restrict first, falls back to manual scan

### Match Logic

- --match-all true (default): Requires ALL terms to match (AND logic)
- --match-all false: Matches ANY term (OR logic)

### Recommended Search Workflow

```bash
# 1. Start narrow and recent
py -3 scripts/outlook_skill.py find --type subject --query "topic" --days 14

# 2. If not found, widen time window
py -3 scripts/outlook_skill.py find --type subject --query "topic" --days 45

# 3. From any result, find the full conversation thread
py -3 scripts/outlook_skill.py find-thread "<email_id>"

# 4. For broader context, find related emails across threads
py -3 scripts/outlook_skill.py find-related "<email_id>"
```

## Configuration

Configuration is centralized in `backend/config.py`:

- Search limits: Max 365 days lookback, default 14 days for direct find
- Related search: Max results (default: 20), lookback (default: 90 days)
- Fuzzy matching: Subject similarity threshold (0.60), time proximity (±7 days)
- Body keyword extraction: First 500 chars for keyword strategy
- Batch sizes: BCC limit (500), extraction batch sizes
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

- Server-side search: Uses Outlook Restrict filter (near-instant for subject/sender/recipient)
- Merged scan: find-related scans each folder once for all strategies (not per-strategy)
- find-thread: Tries Restrict first, falls back to manual scan only when needed
- Body search: Subject pre-filter reduces emails needing full body read
- Batch operations: Forward to 500+ recipients per batch via BCC
- Local processing: All operations via COM, no network latency

## Related Files

- SKILL.md — Quick reference for AI integration (triggers, operations, format rules)
- scripts/outlook_skill.py — CLI entry point (all commands)
- backend/config.py — Centralized configuration (all tunable parameters)

## Contributing

Contributions are welcome! For contributions:

1. Follow the existing code structure
2. Add proper error handling
3. Update documentation for new features
4. Test with various Outlook configurations
5. Submit pull requests with clear descriptions

## License

MIT License - See LICENSE file for details.

---

Note: This skill operates directly with Outlook COM objects without caching. All email operations use message_id (Outlook EntryID) for identification.