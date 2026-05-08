---
name: outlook
description: Microsoft Outlook email management - search, list, compose, reply, forward
triggers: ["outlook", "email", "list emails", "search email"]
operations: ["list-recent", "search", "compose", "reply", "batch-forward", "contact-lookup"]
---

# Outlook Skill

> **⚠️ ALL emails use HTML format:** `<p>text</p>`, `<br>`, `<strong>bold</strong>`

## Commands

### List Recent Emails
```bash
py -3 scripts/outlook_skill.py list-recent --days 7
```
- Shows: To/CC, attachments (with filenames), body preview
- `--days`: 1-30 (default: 7)
- `--folder`: optional (default: Inbox)

### Search Emails
```bash
py -3 scripts/outlook_skill.py search --type sender --query "Name" --days 30
```
- `--type`: subject, sender, recipient, body
- `--query`: search text (required)
- `--days`: 1-30 (default: 30)
- Shows same details as list-recent

### Contact Lookup (Use Before Search by Email)
```bash
py -3 scripts/outlook_skill.py lookup-contact "user@domain.com"
```
- Returns: Display name, company, job title
- **Why:** Outlook search by email address unreliable; use display name instead

### Reply to Email
```bash
py -3 scripts/outlook_skill.py reply "<email_id>" "<p>HTML body</p>" [--send]
```
- Default: Preview only
- `--send`: Actually send email

### Compose Email
```bash
py -3 scripts/outlook_skill.py compose --to "email" --subject "text" --body "<p>HTML</p>"
```
- Always sends immediately

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

## Search Workflow for Email Addresses

1. Lookup display name: `lookup-contact "user@domain.com"`
2. Search by display name: `search --type sender --query "Display Name"`

**Why:** Outlook MAPI doesn't reliably search by email address

## Requirements

- Microsoft Outlook 2016+ (running)
- Windows 10+
- Python 3.8+ with pywin32