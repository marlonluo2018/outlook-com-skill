"""
Outlook Skill CLI for BrainClaw
Command-line interface for Outlook email operations using email IDs
"""

import sys
import os
import argparse
from typing import Optional

# Set UTF-8 encoding for stdout to handle emojis and special characters
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# Add parent directory to path to allow imports from backend and tools
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

# Import from local backend and tools directories
from backend.outlook_session.session_manager import OutlookSessionManager
from backend.email_search import (
    list_recent_emails,
    search_email_by_subject,
    search_email_by_from,
    search_email_by_to,
    search_email_by_body,
    list_folders
)
from backend.email_composition import compose_email
# Caching removed - working directly with email results
from backend.outlook_session.contact_operations import get_contact_by_email, get_display_name_from_email


def cmd_list_folders(args):
    """List all Outlook folders"""
    try:
        folders = list_folders(hide_system_folders=args.hide_system)
        print("\nAvailable folders:")
        for folder in folders:
            print(folder)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_list_recent(args):
    """List recent emails with their IDs"""
    
    def extract_display_name(recipient_string):
        """Extract display name from Exchange format or email string"""
        if not recipient_string:
            return ""
        
        # Check if it's Exchange format: Name </o=ExchangeLabs/...>
        if '</o=' in recipient_string.lower() or '/cn=' in recipient_string.lower():
            # Extract the name before the < symbol
            if '<' in recipient_string:
                return recipient_string.split('<')[0].strip()
        
        # Check if it's standard email format: Name <email@domain.com>
        if '<' in recipient_string and '@' in recipient_string:
            return recipient_string.split('<')[0].strip()
        
        # Otherwise return as is
        return recipient_string.strip()
    
    try:
        emails, message = list_recent_emails(args.folder, args.days)
        
        # Count emails
        email_count = len(emails)
        print(f"\n✅ Found {email_count} recent emails\n")
        
        # Display emails with complete details
        if emails:
            for idx, email_data in enumerate(emails, 1):
                email_id = email_data.get('id') or email_data.get('entry_id', '')
                print(f"{'='*80}")
                print(f"Email #{idx}")
                print(f"{'='*80}")
                
                # Basic info
                subject = email_data.get('subject', 'No Subject')
                sender = email_data.get('sender', 'Unknown')
                received = email_data.get('received_time', 'Unknown')
                
                print(f"ID: {email_id}")
                print(f"Subject: {subject}")
                print(f"From: {sender}")
                
                # To recipients - extract display names only
                to_recipients = email_data.get('to_recipients', [])
                if to_recipients:
                    to_list = []
                    for recipient in to_recipients:
                        name = recipient.get('name', '')
                        address = recipient.get('address', '')
                        
                        # Extract display name from Exchange format
                        display_name = extract_display_name(name) if name else extract_display_name(address)
                        
                        if display_name:
                            to_list.append(display_name)
                    
                    if to_list:
                        print(f"To: {'; '.join(to_list)}")
                
                # CC recipients - extract display names only
                cc_recipients = email_data.get('cc_recipients', [])
                if cc_recipients:
                    cc_list = []
                    for recipient in cc_recipients:
                        name = recipient.get('name', '')
                        address = recipient.get('address', '')
                        
                        # Extract display name from Exchange format
                        display_name = extract_display_name(name) if name else extract_display_name(address)
                        
                        if display_name:
                            cc_list.append(display_name)
                    
                    if cc_list:
                        print(f"CC: {'; '.join(cc_list)}")
                
                print(f"Received: {received}")
                
                # Attachments
                has_attachments = email_data.get('has_attachments', False)
                attachments = email_data.get('attachments', [])
                attachments_count = email_data.get('attachments_count', 0)
                embedded_images_count = email_data.get('embedded_images_count', 0)
                
                if has_attachments and attachments:
                    print(f"\n📎 Attachments ({attachments_count}):")
                    for attachment in attachments:
                        # Handle both dict and object attachment formats
                        if isinstance(attachment, dict):
                            name = attachment.get('name') or attachment.get('filename') or 'Unknown'
                            size = attachment.get('size', 0)
                        else:
                            name = getattr(attachment, 'name', None) or getattr(attachment, 'filename', 'Unknown')
                            size = getattr(attachment, 'size', 0)
                        
                        size_kb = size / 1024 if size > 0 else 0
                        print(f"  - {name} ({size_kb:.1f} KB)")
                
                if embedded_images_count > 0:
                    print(f"🖼️  Embedded images: {embedded_images_count}")
                
                # Get body preview - extract the latest message content only
                try:
                    with OutlookSessionManager() as session:
                        email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(email_id)
                        body_text = email_item.Body if hasattr(email_item, 'Body') else ""
                        if body_text:
                            # Clean up the text: remove excessive whitespace and extract meaningful content
                            lines = body_text.split('\n')
                            cleaned_lines = []
                            
                            for line in lines:
                                stripped = line.strip()
                                # Stop at common email thread separators
                                if any(sep in stripped for sep in ['From:', 'Sent:', '________________________________',
                                                                   '________________________________________________________________________________',
                                                                   'Original Message', '-----Original Message-----']):
                                    break
                                # Skip empty lines at the start
                                if not cleaned_lines and not stripped:
                                    continue
                                # Add non-empty lines or single empty line for paragraph breaks
                                if stripped or (cleaned_lines and cleaned_lines[-1]):
                                    cleaned_lines.append(stripped if stripped else '')
                            
                            # Join lines and limit to reasonable length
                            preview = '\n'.join(cleaned_lines).strip()
                            
                            # Limit to 800 characters for readability
                            if len(preview) > 800:
                                preview = preview[:800] + "..."
                            
                            if preview:
                                print(f"\nPreview:\n{preview}")
                except Exception as e:
                    # If preview fails, just skip it
                    pass
                
                print()
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_search(args):
    """Search emails and display with IDs"""
    
    def extract_display_name(recipient_string):
        """Extract display name from Exchange format or email string"""
        if not recipient_string:
            return ""
        
        # Check if it's Exchange format: Name </o=ExchangeLabs/...>
        if '</o=' in recipient_string.lower() or '/cn=' in recipient_string.lower():
            # Extract the name before the < symbol
            if '<' in recipient_string:
                return recipient_string.split('<')[0].strip()
        
        # Check if it's standard email format: Name <email@domain.com>
        if '<' in recipient_string and '@' in recipient_string:
            return recipient_string.split('<')[0].strip()
        
        # Otherwise return as is
        return recipient_string.strip()
    
    try:
        if args.type == 'subject':
            emails, note = search_email_by_subject(args.query, args.days, args.folder, args.match_all)
        elif args.type == 'sender':
            emails, note = search_email_by_from(args.query, args.days, args.folder, args.match_all)
        elif args.type == 'recipient':
            emails, note = search_email_by_to(args.query, args.days, args.folder, args.match_all)
        elif args.type == 'body':
            emails, note = search_email_by_body(args.query, args.days, args.folder, args.match_all)
        else:
            print(f"Unknown search type: {args.type}", file=sys.stderr)
            return 1
        
        print(f"\n✅ Found {len(emails)} matching emails\n")
        
        # Display emails with complete details
        if emails:
            for idx, email_data in enumerate(emails, 1):
                email_id = email_data.get('id') or email_data.get('entry_id', '')
                print(f"{'='*80}")
                print(f"Email #{idx}")
                print(f"{'='*80}")
                
                # Basic info
                subject = email_data.get('subject', 'No Subject')
                sender = email_data.get('sender', 'Unknown')
                received = email_data.get('received_time', 'Unknown')
                
                print(f"ID: {email_id}")
                print(f"Subject: {subject}")
                print(f"From: {sender}")
                
                # To recipients - extract display names only
                to_recipients = email_data.get('to_recipients', [])
                if to_recipients:
                    to_list = []
                    for recipient in to_recipients:
                        name = recipient.get('name', '')
                        address = recipient.get('address', '')
                        
                        # Extract display name from Exchange format
                        display_name = extract_display_name(name) if name else extract_display_name(address)
                        
                        if display_name:
                            to_list.append(display_name)
                    
                    if to_list:
                        print(f"To: {'; '.join(to_list)}")
                
                # CC recipients - extract display names only
                cc_recipients = email_data.get('cc_recipients', [])
                if cc_recipients:
                    cc_list = []
                    for recipient in cc_recipients:
                        name = recipient.get('name', '')
                        address = recipient.get('address', '')
                        
                        # Extract display name from Exchange format
                        display_name = extract_display_name(name) if name else extract_display_name(address)
                        
                        if display_name:
                            cc_list.append(display_name)
                    
                    if cc_list:
                        print(f"CC: {'; '.join(cc_list)}")
                
                print(f"Received: {received}")
                
                # Attachments
                has_attachments = email_data.get('has_attachments', False)
                attachments = email_data.get('attachments', [])
                attachments_count = email_data.get('attachments_count', 0)
                embedded_images_count = email_data.get('embedded_images_count', 0)
                
                if has_attachments and attachments:
                    print(f"\n📎 Attachments ({attachments_count}):")
                    for attachment in attachments:
                        name = attachment.get('name', 'Unknown')
                        size = attachment.get('size', 0)
                        size_kb = size / 1024 if size > 0 else 0
                        print(f"  - {name} ({size_kb:.1f} KB)")
                
                if embedded_images_count > 0:
                    print(f"🖼️  Embedded images: {embedded_images_count}")
                
                # Get body preview - extract the latest message content only
                try:
                    with OutlookSessionManager() as session:
                        email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(email_id)
                        body_text = email_item.Body if hasattr(email_item, 'Body') else ""
                        if body_text:
                            # Clean up the text: remove excessive whitespace and extract meaningful content
                            lines = body_text.split('\n')
                            cleaned_lines = []
                            
                            for line in lines:
                                stripped = line.strip()
                                # Stop at common email thread separators
                                if any(sep in stripped for sep in ['From:', 'Sent:', '________________________________',
                                                                   '________________________________________________________________________________',
                                                                   'Original Message', '-----Original Message-----']):
                                    break
                                # Skip empty lines at the start
                                if not cleaned_lines and not stripped:
                                    continue
                                # Add non-empty lines or single empty line for paragraph breaks
                                if stripped or (cleaned_lines and cleaned_lines[-1]):
                                    cleaned_lines.append(stripped if stripped else '')
                            
                            # Join lines and limit to reasonable length
                            preview = '\n'.join(cleaned_lines).strip()
                            
                            # Limit to 800 characters for readability
                            if len(preview) > 800:
                                preview = preview[:800] + "..."
                            
                            if preview:
                                print(f"\nPreview:\n{preview}")
                except Exception as e:
                    # If preview fails, just skip it
                    pass
                
                print()
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_get_email(args):
    """Get full email details by ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(args.email_id)
            
            print("\nFull email details:")
            print(f"ID: {args.email_id}")
            print(f"Subject: {email_item.Subject}")
            print(f"From: {email_item.SenderName} <{email_item.SenderEmailAddress}>")
            if email_item.To:
                print(f"To: {email_item.To}")
            if email_item.CC:
                print(f"CC: {email_item.CC}")
            print(f"Date: {email_item.ReceivedTime}")
            print(f"\nBody:\n{email_item.Body}")
            
            if email_item.Attachments.Count > 0:
                print("\nAttachments:")
                for i in range(1, email_item.Attachments.Count + 1):
                    attach = email_item.Attachments.Item(i)
                    print(f"- {attach.FileName} ({attach.Size} bytes)")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_reply(args):
    """Reply to an email by ID (HTML format) - Display or Send"""
    try:
        with OutlookSessionManager() as session:
            email_subject = None
            
            # Try to get email metadata
            email_subject = 'Unknown'
            
            # Try namespace.GetItemFromID (reliable method)
            try:
                email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(args.email_id)
            except Exception as e:
                error_msg = str(e)
                if "moved or deleted" in error_msg.lower():
                    print("Error: The email has been moved or deleted from Outlook.")
                    print("Please search for the email again to get a current email ID.")
                    return 1
                else:
                    raise
            
            # Get current user's email address
            current_user = session.outlook.Session.CurrentUser
            current_user_email = (
                current_user.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                if current_user.AddressEntry.GetExchangeUser()
                else ""
            )
            
            # Check if this is from Sent Items
            parent_folder = email_item.Parent
            is_sent_items = "Sent Items" in parent_folder.Name or "已发送邮件" in parent_folder.Name
            
            if args.send:
                # Check if custom recipients are provided
                if args.to or args.cc:
                    # Use custom recipients
                    reply = session.outlook.CreateItem(0)  # 0 = olMailItem
                    
                    # Add custom To recipients
                    if args.to:
                        for recipient in args.to.split(","):
                            recipient = recipient.strip()
                            if recipient:
                                reply.Recipients.Add(recipient)
                    
                    # Add custom CC recipients
                    if args.cc:
                        for recipient in args.cc.split(","):
                            recipient = recipient.strip()
                            if recipient:
                                cc_recip = reply.Recipients.Add(recipient)
                                cc_recip.Type = 2  # 2 = olCC
                    
                    # Set subject
                    reply.Subject = f"RE: {email_item.Subject}" if not email_item.Subject.startswith("RE:") else email_item.Subject
                    
                    # Set HTML body with original email
                    separator = '<hr style="border: 1px solid #ccc; margin: 20px 0;">'
                    original_body = email_item.HTMLBody if email_item.HTMLBody else f"<p>{email_item.Body}</p>"
                    reply.HTMLBody = args.body + separator + original_body
                elif is_sent_items:
                    # For Sent Items, create a new email with original recipients
                    reply = session.outlook.CreateItem(0)  # 0 = olMailItem
                    
                    # Add original To recipients
                    if email_item.To:
                        for recipient in email_item.To.split(";"):
                            recipient = recipient.strip()
                            if recipient and recipient.lower() != current_user_email.lower():
                                reply.Recipients.Add(recipient)
                    
                    # Add original CC recipients
                    if email_item.CC:
                        for recipient in email_item.CC.split(";"):
                            recipient = recipient.strip()
                            if recipient and recipient.lower() != current_user_email.lower():
                                cc_recip = reply.Recipients.Add(recipient)
                                cc_recip.Type = 2  # 2 = olCC
                    
                    # Set subject
                    reply.Subject = f"RE: {email_item.Subject}" if not email_item.Subject.startswith("RE:") else email_item.Subject
                    
                    # Set HTML body with original email
                    separator = '<hr style="border: 1px solid #ccc; margin: 20px 0;">'
                    original_body = email_item.HTMLBody if email_item.HTMLBody else f"<p>{email_item.Body}</p>"
                    reply.HTMLBody = args.body + separator + original_body
                else:
                    # For Inbox/other folders, use ReplyAll
                    reply = email_item.ReplyAll()
                    
                    # Remove current user from recipients
                    recipients_to_remove = []
                    for i in range(1, reply.Recipients.Count + 1):
                        recipient = reply.Recipients.Item(i)
                        recipient_email = recipient.Address
                        try:
                            if recipient.AddressEntry.Type == "EX":
                                recipient_email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                        except:
                            pass
                        
                        if recipient_email.lower() == current_user_email.lower():
                            recipients_to_remove.append(i)
                    
                    for i in reversed(recipients_to_remove):
                        reply.Recipients.Remove(i)
                    
                    # Set HTML body (Outlook adds its own separator automatically)
                    reply.HTMLBody = args.body + reply.HTMLBody
                
                # Check if there are any recipients left
                if reply.Recipients.Count == 0:
                    print("Error: No recipients found.")
                    print(f"Original To: {email_item.To}")
                    print(f"Original CC: {email_item.CC}")
                    print("Please verify the email has valid recipients.")
                    return 1
                
                # Save recipient count before Send (reply object becomes stale after Send)
                recipient_count = reply.Recipients.Count
                
                # Send
                reply.Send()
                print(f"Email sent successfully to {recipient_count} recipient(s)")
            else:
                # Just display the email content (no saving)
                print("\n" + "="*60)
                print("EMAIL PREVIEW (Not saved, not sent)")
                print("="*60)
                
                # Check if custom recipients provided
                if args.to or args.cc:
                    print(f"To: {args.to if args.to else '(none)'}")
                    if args.cc:
                        print(f"CC: {args.cc}")
                    print(f"\nNote: Using custom recipients")
                elif is_sent_items:
                    # Show original recipients from Sent Items
                    print(f"To: {email_item.To}")
                    if email_item.CC:
                        print(f"CC: {email_item.CC}")
                    print(f"\nNote: Replying to original recipients from Sent Items (excluding {current_user_email})")
                else:
                    # Show ReplyAll recipients
                    print(f"To: {email_item.SenderName} <{email_item.SenderEmailAddress}>")
                    if email_item.CC:
                        print(f"CC: {email_item.CC}")
                    print(f"\nNote: Will reply to all original recipients (excluding {current_user_email})")
                
                print(f"Subject: RE: {email_item.Subject}")
                print("\n" + "-"*60)
                print("HTML Body:")
                print("-"*60)
                print(args.body)
                print("\n" + "="*60)
                print("Use --send flag to send this email")
                print("="*60)
        
        return 0
    except Exception as e:
        error_msg = str(e)
        if "moved or deleted" in error_msg.lower():
            print("Error: The email has been moved or deleted from Outlook.")
            print("Please search for the email again to get a current email ID.")
        else:
            print(f"Error: {error_msg}", file=sys.stderr)
        return 1


def cmd_compose(args):
    """Compose and send new email (always HTML format)"""
    try:
        to_list = [x.strip() for x in args.to.split(",")] if args.to else []
        cc_list = [x.strip() for x in args.cc.split(",")] if args.cc else []
        
        # Replace literal \n with actual newlines
        body = args.body.replace('\\n', '\n')
        
        # Convert plain text to HTML
        # Split by double newlines to get paragraphs
        paragraphs = body.split('\n\n')
        html_body = '<html><body>\n'
        for para in paragraphs:
            # Replace single newlines with <br> within paragraphs
            para = para.replace('\n', '<br>')
            html_body += f'<p>{para}</p>\n'
        html_body += '</body></html>'
        
        # Create email using Outlook COM API with HTML body
        with OutlookSessionManager() as session:
            mail = session.outlook.CreateItem(0)  # 0 = olMailItem
            
            # Set recipients
            for recipient in to_list:
                mail.Recipients.Add(recipient)
            if cc_list:
                for cc_recipient in cc_list:
                    cc_recip = mail.Recipients.Add(cc_recipient)
                    cc_recip.Type = 2  # 2 = olCC
            
            # Set subject and HTML body
            mail.Subject = args.subject
            mail.HTMLBody = html_body
            
            # Send
            mail.Send()
            total_recipients = len(to_list) + len(cc_list)
            print(f"HTML email sent successfully to {total_recipients} recipient(s)")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_batch_forward(args):
    """Batch forward email to multiple recipients by email ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(args.email_id)
            
            # Read CSV file
            import csv
            recipients = []
            with open(args.csv_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if 'email' in row:
                        recipients.append(row['email'].strip())
            
            if not recipients:
                print("Error: No email addresses found in CSV", file=sys.stderr)
                return 1
            
            # Forward to recipients in batches (max 500 per batch)
            batch_size = 500
            total_sent = 0
            
            for i in range(0, len(recipients), batch_size):
                batch = recipients[i:i + batch_size]
                
                # Create forward
                forward = email_item.Forward()
                
                # Add custom message if provided
                if args.message:
                    forward.Body = args.message + "\n\n" + forward.Body
                
                # Add recipients as BCC (to protect privacy)
                for recipient in batch:
                    forward.BCC += recipient + ";"
                
                # Send
                forward.Send()
                total_sent += len(batch)
                print(f"Sent batch {i//batch_size + 1}: {len(batch)} recipients")
            
            print(f"\nTotal recipients: {total_sent}")
            print(f"Successfully forwarded email: {email_item.Subject}")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_create_folder(args):
    """Create a new folder"""
    try:
        with OutlookSessionManager() as session:
            result = session.create_folder(args.name, args.parent)
            print(result)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_remove_folder(args):
    """Remove a folder"""
    try:
        from backend.outlook_session.folder_operations import remove_folder
        result = remove_folder(args.name)
        print(result)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_move_email(args):
    """Move an email to a folder by email ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(args.email_id)
            
            # Get target folder
            target_folder = session.get_folder(args.folder)
            if not target_folder:
                print(f"Error: Folder '{args.folder}' not found", file=sys.stderr)
                return 1
            
            # Move email
            email_item.Move(target_folder)
            print(f"Successfully moved email to '{args.folder}'")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_delete_email(args):
    """Delete an email by ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(args.email_id)
            subject = email_item.Subject
            email_item.Delete()
            print(f"Successfully deleted email: {subject}")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_lookup_contact(args):
    """Look up contact information by email address"""
    try:
        contact_info = get_contact_by_email(args.email)
        
        if contact_info:
            print("\nContact Information:")
            print(f"Display Name: {contact_info['display_name']}")
            print(f"Email: {contact_info['email']}")
            if contact_info.get('first_name'):
                print(f"First Name: {contact_info['first_name']}")
            if contact_info.get('last_name'):
                print(f"Last Name: {contact_info['last_name']}")
            if contact_info.get('company'):
                print(f"Company: {contact_info['company']}")
            if contact_info.get('job_title'):
                print(f"Job Title: {contact_info['job_title']}")
        else:
            print(f"No contact found for: {args.email}")
            print("\nTip: Try searching by partial email username instead:")
            username = args.email.split('@')[0] if '@' in args.email else args.email
            print(f"  py outlook_skill.py search --type sender --query \"{username}\"")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def main():
    parser = argparse.ArgumentParser(description='Outlook Skill for BrainClaw - Email Management CLI')
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # List folders command
    parser_list_folders = subparsers.add_parser('list-folders', help='List all Outlook folders')
    parser_list_folders.add_argument('--hide-system', action='store_true', help='Hide system folders')
    parser_list_folders.set_defaults(func=cmd_list_folders)
    
    # List recent emails command
    parser_list_recent = subparsers.add_parser('list-recent', help='List recent emails with IDs')
    parser_list_recent.add_argument('--days', type=int, default=7, help='Days back to search (1-30)')
    parser_list_recent.add_argument('--folder', type=str, default=None, help='Folder name (default: Inbox)')
    parser_list_recent.set_defaults(func=cmd_list_recent)
    
    # Search emails command
    parser_search = subparsers.add_parser('search', help='Search emails and display with IDs')
    parser_search.add_argument('--type', required=True, choices=['subject', 'sender', 'recipient', 'body'], help='Search type')
    parser_search.add_argument('--query', required=True, help='Search query')
    parser_search.add_argument('--days', type=int, default=30, help='Days back to search')
    parser_search.add_argument('--folder', type=str, default=None, help='Folder name (default: Inbox)')
    parser_search.add_argument('--match-all', action='store_true', default=True, help='Match all terms (AND logic)')
    parser_search.set_defaults(func=cmd_search)
    
    # Get email command
    parser_get_email = subparsers.add_parser('get-email', help='Get full email details by ID')
    parser_get_email.add_argument('email_id', help='Email ID from search results')
    parser_get_email.set_defaults(func=cmd_get_email)
    
    # Reply command
    parser_reply = subparsers.add_parser('reply', help='Reply to an email by ID (preview by default, use --send to send)')
    parser_reply.add_argument('email_id', help='Email ID from search results')
    parser_reply.add_argument('body', help='Reply text in HTML format')
    parser_reply.add_argument('--send', action='store_true', help='Send the email (default: preview only)')
    parser_reply.add_argument('--to', help='Override To recipients (comma separated)')
    parser_reply.add_argument('--cc', help='Override CC recipients (comma separated)')
    parser_reply.set_defaults(func=cmd_reply)
    
    # Compose command
    parser_compose = subparsers.add_parser('compose', help='Compose and send new email')
    parser_compose.add_argument('--to', required=True, help='To recipients (comma separated)')
    parser_compose.add_argument('--subject', required=True, help='Email subject')
    parser_compose.add_argument('--body', required=True, help='Email body')
    parser_compose.add_argument('--cc', help='CC recipients (comma separated)')
    parser_compose.set_defaults(func=cmd_compose)
    
    # Batch forward command
    parser_batch = subparsers.add_parser('batch-forward', help='Batch forward email by ID to multiple recipients')
    parser_batch.add_argument('email_id', help='Email ID from search results')
    parser_batch.add_argument('csv_path', help='Path to CSV file with email addresses')
    parser_batch.add_argument('--message', help='Custom message to prepend')
    parser_batch.set_defaults(func=cmd_batch_forward)
    
    # Create folder command
    parser_create_folder = subparsers.add_parser('create-folder', help='Create a new folder')
    parser_create_folder.add_argument('name', help='Folder name')
    parser_create_folder.add_argument('--parent', help='Parent folder name')
    parser_create_folder.set_defaults(func=cmd_create_folder)
    
    # Remove folder command
    parser_remove_folder = subparsers.add_parser('remove-folder', help='Remove a folder')
    parser_remove_folder.add_argument('name', help='Folder name or path')
    parser_remove_folder.set_defaults(func=cmd_remove_folder)
    
    # Move email command
    parser_move = subparsers.add_parser('move-email', help='Move an email to a folder by ID')
    parser_move.add_argument('email_id', help='Email ID from search results')
    parser_move.add_argument('folder', help='Target folder name')
    parser_move.set_defaults(func=cmd_move_email)
    
    # Delete email command
    parser_delete = subparsers.add_parser('delete-email', help='Delete an email by ID')
    parser_delete.add_argument('email_id', help='Email ID from search results')
    parser_delete.set_defaults(func=cmd_delete_email)
    
    # Lookup contact command
    parser_lookup = subparsers.add_parser('lookup-contact', help='Look up contact information by email address')
    parser_lookup.add_argument('email', help='Email address to look up')
    parser_lookup.set_defaults(func=cmd_lookup_contact)
    
    # Parse arguments
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    # Execute command
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())