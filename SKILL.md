---
name: outlook
description: Microsoft Outlook email management - search, list, compose, reply, forward, download attachments, thread tracking
triggers: [
  "check email", "check inbox", "any new emails", "what's new",
  "show recent emails", "show emails", "list emails",
  "find emails about", "find all emails from", "search for emails",
  "find thread", "find conversation",
  "find related",
  "draft email", "compose", "write email", "new email",
  "reply", "forward", "send to",
  "batch forward", "mass forward", "forward to multiple",
  "get email", "view email", "show email details",
  "download attachment", "save attachment", "get attachment",
  "lookup contact", "who is"
]
operations: ["find-recent", "find", "compose", "reply", "forward", "batch-forward", "download-attachment", "contact-lookup", "find-thread", "find-related", "get-email"]
---

# Outlook Skill

> **⚠️ ALL emails use HTML format:** `<p>text</p>`, `<br>`, `<strong>bold</strong>`

## Commands

### Find Recent Emails
```bash
py -3 scripts/outlook_skill.py find-recent --days 7
```
- Default: **Inbox only** (your sent emails are tracked in task files)
- Shows: To/CC, attachments, body preview, folder indicator (📥/📤)
- `--days`: 1-365 (default: 7)
- `--folder`: override folder (rarely needed)

### Find Emails
```bash
py -3 scripts/outlook_skill.py find --type subject --query "Name" --days 14
```
- Default folder depends on `--type`:
  - `sender`, `subject`, `body` → **Inbox** only
  - `recipient` → **Sent Items** only
- `--type`: subject, sender, recipient, body
- `recipient` search matches recipients in sent mail using **To + CC** fields and resolved Outlook recipient names/addresses
- `--query`: search text (required)
- `--days`: 1-365 for direct `find` searches (default: 14)
- `--folders`: use only when explicitly searching across folders (searches Inbox + Sent Items)
- **AI guidance:** start with a small recent window first (usually 7-14 days)
- If the first search does not find the email, widen the date range gradually and make the query more specific before broadening further
- Use [`find-thread`](assistant_brain/skills/outlook-skill/SKILL.md:49) or [`find-related`](assistant_brain/skills/outlook-skill/SKILL.md:59) when older or broader history is needed

### Find Thread
```bash
py -3 scripts/outlook_skill.py find-thread "<email_id>"
```
- **Auto-searches Inbox + Sent Items** — thread completeness requires both
- Finds ALL emails sharing the same ConversationID
- Subjects can differ (RE:/Fwd: prefixes, topic changes don't matter)
- Uses a reliable full-folder scan path for conversation matching when Outlook filtering is inconsistent
- Results sorted chronologically (oldest first)

### Find Related Emails
```bash
py -3 scripts/outlook_skill.py find-related "<email_id>"
```
- **Auto-searches Inbox + Sent Items** — multi-strategy needs full data
- Strategies: thread (★5) + sender (★3) + keyword (★2)
- Results sorted by relevance
- Multi-strategy search for emails related to a given email:
  - **thread** (★5): Same ConversationID
  - **sender** (★3): Same sender within time window, but only when the email also overlaps with the reference topic
  - **keyword** (★2): Shared meaningful topic keywords from the subject/content
- Generic noise terms such as external/training/request are intentionally ignored during keyword extraction
- Sender and keyword matching are intentionally tighter to reduce unrelated same-sender and boilerplate matches
- Results sorted by relevance (confidence score)
- `--strategies`: comma-separated (default: all three)

### Contact Lookup (Use Before Search by Email)
```bash
py -3 scripts/outlook_skill.py lookup-contact "user@domain.com"
py -3 scripts/outlook_skill.py lookup-contact "HONG YANG"
```
- Accepts: Email address OR display name
- Auto-detects: `@` present → email lookup; no `@` → name lookup via Exchange GAL
- Returns: Display name, email, company, job title
- **Why:** Outlook search by email address unreliable; use display name instead
- **Name lookup:** Resolves display names against Exchange Global Address List

### ReplyAll (default)
```bash
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>HTML body</p>"
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>HTML body</p>" --cc "extra@ibm.com"
py -3 scripts/outlook_skill.py replyall "<email_id>" "<p>HTML body</p>" --attach "C:\path\file.pdf"
```
- Keeps ALL original To + CC recipients. `--to`/`--cc` APPEND to existing.
- `--attach`: File path(s) to attach (comma separated for multiple)
- **This is the default reply command.** Use unless you need to narrow recipients.
- **⚠️ ALWAYS show draft to user first — NEVER send before user approval**

### Reply (specify mode)
```bash
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>"
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>" --to "specific@ibm.com"
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>" --attach "C:\path\file.pdf"
```
- Replies to sender only. `--to`/`--cc` specify EXACT extra recipients (original To/CC NOT included).
- `--attach`: File path(s) to attach (comma separated for multiple)
- Use when you want to narrow the recipient list.

### Compose Email
```bash
py -3 scripts/outlook_skill.py compose --to "email" --subject "text" --body "<p>HTML</p>"
py -3 scripts/outlook_skill.py compose --to "email" --subject "text" --body "<p>HTML</p>" --attach "C:\path\file.pdf"
```
- `--attach`: File path(s) to attach (comma separated for multiple)
- **⚠️ ALWAYS show draft to user in chat window first — NEVER send before user approval**
- AI presents the email as readable plain text in chat
- Only call this command after user explicitly confirms "send" or "approve"
- Sends immediately when called

### Forward (single)
```bash
py -3 scripts/outlook_skill.py forward "<email_id>" --to "user@domain.com"
py -3 scripts/outlook_skill.py forward "<email_id>" --to "user1@ibm.com,user2@ibm.com" --cc "manager@ibm.com" --body "<p>FYI</p>"
py -3 scripts/outlook_skill.py forward "<email_id>" --to "user@domain.com" --attach "C:\path\file.pdf"
```
- Forwards an email to specified recipients
- `--to` (required): Comma-separated list of To recipients
- `--cc` (optional): Comma-separated list of CC recipients
- `--body` (optional): Custom HTML message to prepend
- `--attach` (optional): File path(s) to attach (comma separated for multiple)
- Subject auto-prefixed with `FW:`
- Preserves original email formatting
- **⚠️ ALWAYS show draft to user first — NEVER send before user approval**

### Batch Forward
```bash
py -3 scripts/outlook_skill.py batch-forward "<email_id>" "recipients.csv" --message "<p>HTML body</p>"
```
- CSV: single column named "email" (supports BOM encoding)
- `--message`: Optional HTML message to prepend (same format as reply)
- Uses BCC for privacy
- Preserves original email formatting
- Automatically splits large recipient lists into batches
- **Batch size:** Configured in [`backend/config.py`](backend/config.py) (default: 500)

### Download Attachment
```bash
py -3 scripts/outlook_skill.py download-attachment "<email_id>"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --output-dir "C:\temp"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --filename "report.pdf"
```
- Saves email attachments to local directory (default: `~/Downloads`)
- `--output-dir`: Override save location
- `--filename`: Download only a specific attachment by name
- `--all`: Include embedded images (default: skips inline images, keeps PDFs/docs)
- Use with `reply --attach` or `forward --attach` to relay attachments

### Get Full Email Details
```bash
py -3 scripts/outlook_skill.py get-email "<email_id>"
```
- Returns complete email: full body, all attachments, metadata
- Use after search/thread/related to read the actual content

## Configuration

All configuration is centralized in [`backend/config.py`](backend/config.py).

**To change batch size:**
Edit `backend/config.py` and modify:
```python
class BatchConfig:
    OUTLOOK_BCC_LIMIT = 500  # Change this value
```

**Batch size recommendations:**
- **500** (default): Recommended for production use
- **100**: For testing with smaller batches
- **1000**: Maximum (may hit Exchange server limits)

## HTML Format Examples

```html
<p>Dear John,</p>
<p>Message text here.</p>
<p>Best regards,<br>Marlon</p>
```

## ⚠️ Special Characters in Email Body

**CRITICAL:** Replace `$` with `&#36;` in HTML body to avoid shell variable issues.

```html
<!-- ❌ WRONG: $80,000 displays as ,000 -->
<p>Cost: $80,000 USD</p>

<!-- ✅ CORRECT: Use HTML entity -->
<p>Cost: &#36;80,000 USD</p>
```

**Common HTML entities:** `$` = `&#36;` | `&` = `&amp;` | `<` = `&lt;` | `>` = `&gt;`

## Find Workflow for Email Addresses

1. Lookup display name: `lookup-contact "user@domain.com"`
2. Find by display name: `find --type sender --query "Display Name"`

**Why:** Outlook MAPI doesn't reliably search by email address

## Recommended AI Usage Flow

### Finding All Emails About a Topic
```bash
# Step 1: Start narrow and recent with a specific query
py -3 scripts/outlook_skill.py find --type subject --query "voucher approval" --folders "Inbox,Sent Items" --days 14

# Step 2: If not found, widen the time window but make the query more specific
py -3 scripts/outlook_skill.py find --type subject --query "voucher approval philippines" --folders "Inbox,Sent Items" --days 45

# Step 3: From any result, find the full thread
py -3 scripts/outlook_skill.py find-thread "<entry_id>"

# Step 4: For even more context, find related across threads
py -3 scripts/outlook_skill.py find-related "<entry_id>"
```

## Requirements

- Microsoft Outlook 2016+ (running)
- Windows 10+
- Python 3.8+ with pywin32
- SQLite 3.35+ (included with Python 3.8+)
