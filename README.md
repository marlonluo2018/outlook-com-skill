# Outlook Skill

A Python-based skill for managing Microsoft Outlook emails through command-line interface. Provides 22 commands covering search, compose, reply, forward, template editing, folder management, and more.

## Overview

This is a **universal skill module** that can be integrated with any AI assistant system. It interfaces directly with Microsoft Outlook via COM automation, uses **message_id** (EntryID) for email identification, and operates without caching — all operations work directly with Outlook.

## Key Features

- **Email Search**: Search by subject, sender, recipient, or body content
- **Thread Tracking**: Find all emails in the same conversation (even when subject changes)
- **Related Email Discovery**: Multi-strategy search by sender, recipients, keywords
- **Email Operations**: Compose, reply, reply-all, forward, redirect, recall, move, and delete
- **Template Editing**: Use any email as a template — read HTML, make targeted replacements, preview in Drafts, send when ready
- **Batch Operations**: Forward emails to multiple recipients via CSV (BCC, auto-batching)
- **Folder Management**: Create, list, and manage Outlook folders
- **Contact Lookup**: Retrieve contact information by email or display name (Exchange GAL)
- **Attachment Handling**: Download attachments, embed inline images via CID
- **HTML Email Support**: All emails use HTML format for rich content

## Requirements

- Operating System: Windows 10 or later
- Microsoft Outlook: 2016 or later (must be running)
- Python: 3.8 or later
- Dependencies: pywin32 (for Windows COM automation)

## Installation

```bash
pip install -r requirements.txt
```

Ensure Outlook is running — the skill requires Outlook to be open and logged in.

## Integration

- **BrainClaw**: Native integration as a skill module
- **Custom AI Systems**: Direct CLI integration or Python API
- **Other Assistants**: Adaptable to any system that can execute Python scripts

## Usage

All commands are executed through the CLI:

```bash
py -3 scripts/outlook_skill.py <command> [options]
```

---

### Search & Discovery

#### Find Recent Emails

```bash
py -3 scripts/outlook_skill.py find-recent --days 7
py -3 scripts/outlook_skill.py find-recent --days 7 --folder "Inbox"
```

- Default: Inbox + Sent Items (both folders)
- `--days`: 1-365 (default: 7)
- `--folder`: Override to single folder
- Output: ID, subject, sender, To/CC, attachments, folder indicator, body preview

#### Find Emails (Search)

```bash
py -3 scripts/outlook_skill.py find --type subject --query "Meeting" --days 14
py -3 scripts/outlook_skill.py find --type sender --query "John Smith"
py -3 scripts/outlook_skill.py find --type recipient --query "Jane Doe"
py -3 scripts/outlook_skill.py find --type body --query "project update"
```

- `--type`: subject, sender, recipient, body (required)
- `--query`: Search term (required)
- `--days`: 1-365 (default: 14)
- `--folder` / `--folders`: Override search location
- `--match-all`: AND logic (default: true)

Default folder by type: `subject`/`body` → Inbox + Sent Items | `sender` → Inbox | `recipient` → Sent Items

#### Find Thread (Conversation)

```bash
py -3 scripts/outlook_skill.py find-thread "<email_id>"
py -3 scripts/outlook_skill.py find-thread "<email_id>" --fuzzy
py -3 scripts/outlook_skill.py find-thread "<email_id>" --brief
```

Finds ALL emails sharing the same ConversationID across Inbox + Sent Items.

- `--fuzzy`: Also find emails with similar subjects (token overlap ≥ 0.6)
- `--brief`: Compact single-line output
- `--folders`: Override search folders
- Shows thread summary: message count, participants, date span

#### Find Related Emails (Multi-Strategy)

```bash
py -3 scripts/outlook_skill.py find-related "<email_id>"
py -3 scripts/outlook_skill.py find-related "<email_id>" --exclude-thread
py -3 scripts/outlook_skill.py find-related "<email_id>" --strategies sender,keyword --max 10
```

4 strategies ranked by relevance:

| Strategy  | What it finds                                          | Relevance |
|-----------|--------------------------------------------------------|-----------|
| thread    | Same conversation ID                                   | ★★★★★     |
| sender    | Same sender + overlapping topic keywords               | ★★★★      |
| recipient | ≥2 shared recipients in To/CC                          | ★★★       |
| keyword   | Shared meaningful keywords from subject + body         | ★★★       |

- `--days`: Lookback window (default: 90)
- `--strategies`: Comma-separated list (default: all four)
- `--exclude-thread`: Skip thread strategy
- `--max`: Limit results (default: 20)
- `--brief`: Compact output

#### Contact Lookup

```bash
py -3 scripts/outlook_skill.py lookup-contact "user@example.com"
py -3 scripts/outlook_skill.py lookup-contact "HONG YANG"
```

Auto-detects format (`@` → email lookup, no `@` → Exchange GAL name lookup).
Returns: display name, email, alias, company, department, job title, office, phone, mobile, location.

#### Get Full Email Details

```bash
py -3 scripts/outlook_skill.py get-email "<email_id>"
```

Returns complete email: full body (plain text), all attachments, metadata. Embedded images auto-extracted to temp directory.

---

### Email Composition & Sending

#### Compose New Email

```bash
py -3 scripts/outlook_skill.py compose --to "user@example.com" --subject "Subject" --body "<p>HTML body</p>"
py -3 scripts/outlook_skill.py compose --to "a@b.com" --subject "Report" --body "<p>See attached</p>" --attach "C:\file.pdf"
py -3 scripts/outlook_skill.py compose --to "a@b.com" --subject "Photo" --body "<p><img src='cid:pic1'></p>" --inline-image "C:\img.png:pic1"
```

- `--to`: Recipients (comma-separated, required)
- `--subject`: Subject line (required)
- `--body`: HTML body (required)
- `--cc`: CC recipients
- `--attach`: File paths (comma-separated)
- `--inline-image`: Embed images (`filepath:cid_name`, comma-separated)
- Sends immediately when called

#### ReplyAll (Default Reply)

```bash
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>HTML body</p>"
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>body</p>" --cc "extra@ibm.com"
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>body</p>" --attach "C:\file.pdf"
```

Keeps ALL original To + CC recipients. `--to`/`--cc` APPEND to existing.

#### Reply (Sender Only)

```bash
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>"
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>body</p>" --to "specific@ibm.com"
```

Replies to sender only. `--to`/`--cc` specify EXACT extra recipients (original To/CC NOT included).

#### Forward

```bash
py -3 scripts/outlook_skill.py forward "<email_id>" --to "user@example.com"
py -3 scripts/outlook_skill.py forward "<email_id>" --to "a@b.com,c@d.com" --cc "mgr@b.com" --body "<p>FYI</p>"
```

- `--to`: Recipients (required)
- `--cc`: CC recipients
- `--body`: Custom HTML message to prepend
- `--attach`: Additional attachments
- Subject auto-prefixed with `FW:`

#### Redirect (Clear Recipients + New TO/CC)

```bash
py -3 scripts/outlook_skill.py redirect "<email_id>" "<p>New message</p>" --to "a@b.com"
py -3 scripts/outlook_skill.py redirect "<email_id>" "<p>FYI</p>" --to "a@b.com" --cc "b@b.com"
```

Clears ALL existing TO/CC, adds new ones. Preserves original email body as quoted content.

#### Batch Forward

```bash
py -3 scripts/outlook_skill.py batch-forward "<email_id>" "recipients.csv" --message "<p>FYI</p>"
```

- CSV: single column named `email` (supports BOM encoding)
- Uses BCC for privacy
- Auto-splits into batches of 500 (configurable in config.py)

#### Recall Sent Email

```bash
py -3 scripts/outlook_skill.py recall "<email_id>"
```

Recalls a sent email via Exchange server. Only works for internal Exchange recipients who haven't read the message.

---

### Template Editing (Email as Template)

Use any email as a starting point — modify HTML content, swap photos, change text — and preview in Outlook before sending.

#### Get Email HTML

```bash
py -3 scripts/outlook_skill.py get-html "<email_id>"
```

Returns raw HTMLBody for inspection. Output wrapped in `HTML_START` / `HTML_END` markers.

#### Edit Email HTML

```bash
py -3 scripts/outlook_skill.py edit-html "<email_id>" --replace "old text::new text"
py -3 scripts/outlook_skill.py edit-html "<email_id>" --replace "Q1::Q2" --replace "Jan::Apr"
py -3 scripts/outlook_skill.py edit-html "<email_id>" --replace "old::new" --subject "New Subject" --to "new@ibm.com"
py -3 scripts/outlook_skill.py edit-html "<email_id>" --body-file "C:\temp\modified.html"
```

Behavior:

- **Draft (unsent)**: Modifies in place — same email updated, no duplicates
- **Sent/Received**: Creates a new draft — original untouched

Options:

- `--replace "old::new"`: Text replacement (repeatable)
- `--subject`: Override subject
- `--to` / `--cc`: Override recipients
- `--body-file`: Replace entire HTML from file
- `--copy-attachments`: Copy file attachments from source

#### Send Draft

```bash
py -3 scripts/outlook_skill.py send-draft "<draft_email_id>"
```

Sends an existing draft. Validates: must be unsent, must have recipients.

#### Typical Template Workflow

```bash
# 1. Read the template email's HTML
py -3 scripts/outlook_skill.py get-html "<template_id>"

# 2. Apply targeted replacements (draft auto-saved)
py -3 scripts/outlook_skill.py edit-html "<template_id>" --replace "old name::new name" --replace "old title::new title" --to "recipient@example.com"

# 3. Preview in Outlook Drafts folder

# 4. Iterate if needed (same draft updated in place)
py -3 scripts/outlook_skill.py edit-html "<draft_id>" --replace "fix this::to that"

# 5. Send when ready
py -3 scripts/outlook_skill.py send-draft "<draft_id>"
```

---

### Attachments & Images

#### Download Attachment

```bash
py -3 scripts/outlook_skill.py download-attachment "<email_id>"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --output-dir "C:\temp"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --filename "report.pdf"
```

- `--output-dir`: Save directory (default: ~/Downloads)
- `--filename`: Download specific attachment only
- `--all`: Include embedded images (default: skips inline images)

#### Inline Images (CID)

Embed images directly in email body:

```bash
py -3 scripts/outlook_skill.py compose --to "user@example.com" --subject "Report" \
  --body "<p>Chart: <img src='cid:chart1'></p>" \
  --inline-image "C:\path\chart.png:chart1"
```

- Format: `filepath:cid_name` (comma-separated for multiple)
- Reference in HTML: `<img src="cid:cid_name">`
- Works with: compose, reply, replyall, forward, redirect

---

### Folder & Email Management

```bash
py -3 scripts/outlook_skill.py list-folders
py -3 scripts/outlook_skill.py list-folders --hide-system
py -3 scripts/outlook_skill.py create-folder "ProjectX" --parent "Inbox"
py -3 scripts/outlook_skill.py remove-folder "ProjectX"
py -3 scripts/outlook_skill.py move-email "<email_id>" "Archive"
py -3 scripts/outlook_skill.py delete-email "<email_id>"
```

---

## Architecture

### Project Structure

```text
outlook-skill/
├── scripts/
│   └── outlook_skill.py            # CLI entry point (22 commands)
├── backend/
│   ├── config.py                   # Centralized configuration
│   ├── email_composition.py        # Email building helpers
│   ├── batch_operations.py         # Batch forward logic
│   ├── validation.py               # Input validation rules
│   ├── validators.py               # Pydantic parameter validators
│   ├── logging_config.py           # Logging setup
│   ├── shared.py                   # Shared constants
│   ├── utils.py                    # General utilities
│   ├── email_metadata.py           # Metadata structures
│   ├── email_data_extractor.py     # Email data extraction
│   ├── email_utils.py              # Email utility functions
│   ├── email_search/
│   │   ├── __init__.py             # Public search API exports
│   │   ├── unified_search.py       # Orchestrates search strategies
│   │   ├── server_search.py        # Core engine (Restrict, merged scan)
│   │   ├── search_common.py        # Shared search utilities
│   │   ├── parallel_extractor.py   # Parallel email extraction
│   │   ├── email_listing.py        # Recent email listing
│   │   ├── subject_search.py       # Subject search
│   │   ├── sender_search.py        # Sender search
│   │   ├── recipient_search.py     # Recipient search (To+CC)
│   │   └── body_search.py          # Body content search
│   └── outlook_session/
│       ├── __init__.py             # Session exports
│       ├── session_manager.py      # COM session lifecycle
│       ├── email_operations.py     # Send, reply, forward
│       ├── folder_operations.py    # Folder CRUD
│       ├── contact_operations.py   # Contact/GAL lookup
│       ├── decorators.py           # COM operation decorators
│       ├── utils.py                # Session utilities
│       └── exceptions.py           # Custom exceptions
├── tools/                          # Legacy (deprecated, unused)
├── SKILL.md                        # Quick reference (AI-oriented)
├── README.md                       # Full documentation
├── requirements.txt                # pywin32==311
└── LICENSE                         # MIT
```

### Design Principles

1. **No Caching**: All operations work directly with Outlook COM objects
2. **Message ID Based**: Uses Outlook's EntryID for email identification
3. **Server-Side Search**: Prioritizes Restrict filter over client-side scan
4. **In-Place Editing**: Draft templates modify the same email, no duplicates
5. **HTML Format**: All emails use HTML for rich formatting
6. **Single Merged Scan**: find-related scans each folder once for all strategies

### Technical Details

- COM Automation: win32com.client for Outlook integration
- Session Management: Context manager pattern for resource cleanup
- Error Handling: Retry across multiple stores for EntryID resolution
- Threading: Proper COM threading initialization with pythoncom
- Encoding: UTF-8 throughout, safe_encode_text for COM boundary

---

## HTML Email Format

All email bodies use HTML format:

```html
<p>Hello,</p>
<p>Please review the following:</p>
<ul>
  <li><strong>Item 1</strong>: Description</li>
</ul>
```

**Special characters:** Replace `$` with `&#36;` to avoid shell interpolation.

Common entities: `$` = `&#36;` | `&` = `&amp;` | `<` = `&lt;` | `>` = `&gt;`

---

## Search Tips

### By Email Address

Outlook MAPI doesn't reliably search by email address:

1. `lookup-contact "user@example.com"` → get display name
2. `find --type sender --query "Display Name"` → search by name

### Performance

- Subject/Sender/Recipient: Fast (server-side Restrict)
- Body: Subject pre-filter + body validation
- find-related: Single merged scan per folder
- find-thread: Restrict first, falls back to manual scan

### Recommended Workflow

```bash
# Start narrow → widen → thread → related
find --type subject --query "topic" --days 14
find --type subject --query "topic" --days 45
find-thread "<email_id>"
find-related "<email_id>"
```

---

## Configuration

Centralized in `backend/config.py`:

- Search: Max 365 days, default 14 days for direct find
- Related: Max 20 results, 90-day lookback
- Fuzzy: Subject similarity 0.60, time proximity ±7 days
- Batch: BCC limit 500 (configurable)
- Body keywords: First 500 chars extracted

---

## Troubleshooting

| Issue                              | Solution                                |
| ---------------------------------- | --------------------------------------- |
| "Failed to connect to Outlook"     | Ensure Outlook is running and logged in |
| "Email has been moved or deleted"  | Search again for fresh ID               |
| "No recipients found"              | Check reply target has valid recipients |
| Search returns nothing             | Broaden terms, increase --days          |

---

## Privacy & Security

- 100% local processing — no cloud services
- Uses existing Outlook installation and credentials
- Email content never sent to external servers
- Windows COM API via win32com

---

## License

MIT License — See LICENSE file for details.
