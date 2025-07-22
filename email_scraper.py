#!/usr/bin/env python3
"""
AI Email Scraping Tool
Scrape email details without storing data - analyze on the go!
"""

import imaplib
import email
import ssl
import sys
import getpass
import signal
import time
from collections import defaultdict, Counter
from datetime import datetime, date
import re
from typing import List, Dict, Tuple
import argparse

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.progress import Progress, TaskID
from rich.prompt import Prompt, Confirm
from rich import print as rprint

console = Console()

class EmailScraper:
    def __init__(self):
        self.imap_server = None
        self.email_address = None
        self.connected = False
        
    def connect_to_email(self, email_address: str, password: str, server: str = None, port: int = 993):
        """Connect to email server using IMAP with multiple authentication methods"""
        try:
            self.email_address = email_address
            
            # Auto-detect server if not provided
            if not server:
                server = self._detect_imap_server(email_address)
            
            console.print(f"[yellow]Connecting to {server}:{port}...[/yellow]")
            
            # Try different authentication methods
            self.imap_server = self._try_different_auth_methods(email_address, password, server, port)
            
            if self.imap_server:
                self.connected = True
                console.print("[green]✓ Successfully connected to email server![/green]")
                return True
            else:
                console.print("[red]❌ All authentication methods failed![/red]")
                console.print("[yellow]💡 Suggestions:[/yellow]")
                console.print("  • For Gmail/Google Workspace: Use App Password")
                console.print("  • For corporate email: Contact your IT admin")
                console.print("  • Check if IMAP is enabled in your email settings")
                return False
            
        except Exception as e:
            console.print(f"[red]❌ Failed to connect: {str(e)}[/red]")
            return False
    
    def _try_different_auth_methods(self, email_addr: str, password: str, server: str, port: int = 993):
        """Try different authentication methods"""
        
        methods_to_try = [
            ("Standard SSL", lambda: imaplib.IMAP4_SSL(server, port)),
            ("SSL with custom context", lambda: imaplib.IMAP4_SSL(server, port, ssl_context=ssl.create_default_context())),
            ("Port 143 with STARTTLS", lambda: self._try_starttls(server)),
        ]
        
        for method_name, connection_func in methods_to_try:
            console.print(f"[blue]Trying {method_name}...[/blue]")
            try:
                imap = connection_func()
                console.print("[yellow]Authenticating...[/yellow]")
                imap.login(email_addr, password)
                console.print(f"[green]✅ Success with {method_name}![/green]")
                return imap
            except Exception as e:
                error_msg = str(e)[:100] + "..." if len(str(e)) > 100 else str(e)
                console.print(f"[red]❌ {method_name} failed: {error_msg}[/red]")
                continue
        
        return None
    
    def _try_starttls(self, server: str):
        """Try IMAP with STARTTLS on port 143"""
        imap = imaplib.IMAP4(server, 143)
        imap.starttls()
        return imap
    
    def _detect_imap_server(self, email_address: str) -> str:
        """Auto-detect IMAP server based on email domain"""
        domain = email_address.split('@')[1].lower()
        
        server_map = {
            'gmail.com': 'imap.gmail.com',
            'sigsi.com': 'imap.gmail.com',  # Google Workspace
            'outlook.com': 'outlook.office365.com',
            'hotmail.com': 'outlook.office365.com',
            'live.com': 'outlook.office365.com',
            'yahoo.com': 'imap.mail.yahoo.com',
            'icloud.com': 'imap.mail.me.com',
            'me.com': 'imap.mail.me.com',
            'aol.com': 'imap.aol.com',
        }
        
        return server_map.get(domain, f'imap.{domain}')
    
    def disconnect(self):
        """Disconnect from email server"""
        if self.imap_server and self.connected:
            try:
                self.imap_server.logout()
                console.print("[yellow]Disconnected from email server[/yellow]")
            except:
                pass
            finally:
                self.connected = False
    
    def search_emails_by_sender(self, sender_email: str, mailbox: str = 'INBOX') -> List[Dict]:
        """Search for emails from a specific sender"""
        if not self.connected:
            console.print("[red]Not connected to email server![/red]")
            return []
        
        try:
            # Select mailbox
            self.imap_server.select(mailbox)
            
            # Search for emails from sender
            console.print(f"[yellow]Searching for emails from {sender_email}...[/yellow]")
            result, message_ids = self.imap_server.search(None, f'FROM "{sender_email}"')
            
            if result != 'OK':
                console.print("[red]Search failed![/red]")
                return []
            
            email_list = []
            message_ids = message_ids[0].split()
            
            if not message_ids:
                console.print(f"[yellow]No emails found from {sender_email}[/yellow]")
                return []
            
            # Use the robust fetch method
            return self._fetch_emails_details(message_ids)
            
        except Exception as e:
            console.print(f"[red]Error searching emails: {str(e)}[/red]")
            return []
    
    def search_by_subject_keyword(self, keyword: str, mailbox: str = 'INBOX', limit: int = 100) -> List[Dict]:
        """Search for emails containing keyword in subject"""
        if not self.connected:
            console.print("[red]Not connected to email server![/red]")
            return []
        
        try:
            # Select mailbox
            self.imap_server.select(mailbox)
            
            # Search for emails with keyword in subject
            console.print(f"[yellow]Searching for emails with '{keyword}' in subject...[/yellow]")
            result, message_ids = self.imap_server.search(None, f'SUBJECT "{keyword}"')
            
            if result != 'OK':
                console.print("[red]Search failed![/red]")
                return []
            
            message_ids = message_ids[0].split()
            
            if not message_ids:
                console.print(f"[yellow]No emails found with '{keyword}' in subject[/yellow]")
                return []
            
            # Limit results if specified
            if limit and len(message_ids) > limit:
                message_ids = message_ids[-limit:]  # Get most recent emails
            
            # Use the robust fetch method
            return self._fetch_emails_details(message_ids, limit=limit)
            
        except Exception as e:
            console.print(f"[red]Error searching emails: {str(e)}[/red]")
            return []
    
    def search_emails_by_to(self, to_email: str, mailbox: str = 'INBOX') -> List[Dict]:
        """Search for emails sent to a specific recipient"""
        if not self.connected:
            console.print("[red]Not connected to email server![/red]")
            return []
        
        try:
            self.imap_server.select(mailbox)
            console.print(f"[yellow]Searching for emails sent to {to_email}...[/yellow]")
            result, message_ids = self.imap_server.search(None, f'TO "{to_email}"')
            
            if result != 'OK':
                console.print("[red]Search failed![/red]")
                return []
            
            message_ids = message_ids[0].split()
            
            if not message_ids:
                console.print(f"[yellow]No emails found sent to {to_email}[/yellow]")
                return []
            
            return self._fetch_emails_details(message_ids)
            
        except Exception as e:
            console.print(f"[red]Error searching emails: {str(e)}[/red]")
            return []
    
    def search_emails_by_cc(self, cc_email: str, mailbox: str = 'INBOX') -> List[Dict]:
        """Search for emails where someone was CC'd"""
        if not self.connected:
            console.print("[red]Not connected to email server![/red]")
            return []
        
        try:
            self.imap_server.select(mailbox)
            console.print(f"[yellow]Searching for emails with {cc_email} in CC...[/yellow]")
            result, message_ids = self.imap_server.search(None, f'CC "{cc_email}"')
            
            if result != 'OK':
                console.print("[red]Search failed![/red]")
                return []
            
            message_ids = message_ids[0].split()
            
            if not message_ids:
                console.print(f"[yellow]No emails found with {cc_email} in CC[/yellow]")
                return []
            
            return self._fetch_emails_details(message_ids)
            
        except Exception as e:
            console.print(f"[red]Error searching emails: {str(e)}[/red]")
            return []
    
    def search_emails_by_date_range(self, start_date: str, end_date: str, mailbox: str = 'INBOX') -> List[Dict]:
        """Search for emails within a date range (format: DD-MMM-YYYY)"""
        if not self.connected:
            console.print("[red]Not connected to email server![/red]")
            return []
        
        try:
            self.imap_server.select(mailbox)
            console.print(f"[yellow]Searching for emails from {start_date} to {end_date}...[/yellow]")
            
            # IMAP date format: DD-Mon-YYYY (e.g., 01-Jan-2024)
            search_criteria = f'SINCE "{start_date}" BEFORE "{end_date}"'
            result, message_ids = self.imap_server.search(None, search_criteria)
            
            if result != 'OK':
                console.print("[red]Search failed![/red]")
                return []
            
            message_ids = message_ids[0].split()
            
            if not message_ids:
                console.print(f"[yellow]No emails found in date range {start_date} to {end_date}[/yellow]")
                return []
            
            return self._fetch_emails_details(message_ids)
            
        except Exception as e:
            console.print(f"[red]Error searching emails: {str(e)}[/red]")
            return []
    
    def advanced_email_filter(self, filters: Dict, mailbox: str = 'INBOX') -> List[Dict]:
        """Advanced email filtering with multiple criteria and mailbox selection"""
        if not self.connected:
            console.print("[red]Not connected to email server![/red]")
            return []
        
        try:
            # Auto-detect best mailbox based on search criteria
            suggested_mailbox = self._suggest_mailbox(filters)
            if suggested_mailbox != mailbox:
                console.print(f"[yellow]💡 Suggestion: You might want to search in '{suggested_mailbox}' folder[/yellow]")
                
            # Get available mailboxes first
            available_mailboxes = self._get_available_mailboxes()
            
            # Search in multiple mailboxes if FROM is specified (likely looking for sent emails)
            mailboxes_to_search = []
            if filters.get('from'):
                # Use the same successful method as option 12 to find sent folder
                try:
                    result, folders = self.imap_server.list()
                    if result == 'OK':
                        sent_folder_candidates = []
                        for folder in folders:
                            folder_str = folder.decode('utf-8', errors='ignore').lower()
                            if 'sent' in folder_str:
                                # Extract the actual folder name
                                parts = folder.decode('utf-8', errors='ignore').split('"')
                                if len(parts) >= 3:
                                    folder_name = parts[-2]
                                    sent_folder_candidates.append(folder_name)
                        
                        # Add found sent folders + INBOX as backup
                        mailboxes_to_search.extend(sent_folder_candidates[:2])
                        mailboxes_to_search.append('INBOX')
                    else:
                        mailboxes_to_search = ['INBOX']
                except Exception:
                    mailboxes_to_search = ['INBOX']
            else:
                mailboxes_to_search = [mailbox]
            
            all_emails = []
            
            for search_mailbox in mailboxes_to_search:
                try:
                    console.print(f"[blue]🔍 Searching in mailbox: {search_mailbox}[/blue]")
                    # Use quotes for folder names with special characters (like Gmail folders)
                    if '[Gmail]' in search_mailbox or ' ' in search_mailbox:
                        self.imap_server.select(f'"{search_mailbox}"')
                    else:
                        self.imap_server.select(search_mailbox)
                    
                    # Build search criteria
                    search_criteria = []
                    
                    if filters.get('from'):
                        search_criteria.append(f'FROM "{filters["from"]}"')
                    
                    if filters.get('to'):
                        search_criteria.append(f'TO "{filters["to"]}"')
                    
                    if filters.get('cc'):
                        search_criteria.append(f'CC "{filters["cc"]}"')
                    
                    if filters.get('subject'):
                        search_criteria.append(f'SUBJECT "{filters["subject"]}"')
                    
                    if filters.get('start_date') and filters.get('end_date'):
                        search_criteria.append(f'SINCE "{filters["start_date"]}"')
                        search_criteria.append(f'BEFORE "{filters["end_date"]}"')
                    
                    if not search_criteria:
                        console.print("[yellow]No search criteria specified![/yellow]")
                        continue
                    
                    # Combine criteria with AND
                    final_criteria = ' '.join(search_criteria)
                    console.print(f"[blue]Search criteria: {final_criteria}[/blue]")
                    
                    result, message_ids = self.imap_server.search(None, final_criteria)
                    
                    if result == 'OK' and message_ids and message_ids[0]:
                        message_ids = message_ids[0].split()
                        if message_ids:
                            console.print(f"[green]Found {len(message_ids)} emails in {search_mailbox}[/green]")
                            mailbox_emails = self._fetch_emails_details(message_ids, limit=filters.get('limit', 100))
                            # Add mailbox info to each email
                            for email in mailbox_emails:
                                email['mailbox'] = search_mailbox
                            all_emails.extend(mailbox_emails)
                            
                            # If we found emails and only searching for FROM, we probably found what we need
                            if filters.get('from') and not filters.get('to'):
                                break
                
                except Exception as e:
                    console.print(f"[yellow]⚠️ Could not search in {search_mailbox}: {str(e)[:50]}[/yellow]")
                    continue
            
            if not all_emails:
                console.print("[yellow]No emails found matching the criteria in any mailbox[/yellow]")
                console.print("[cyan]💡 Try these troubleshooting steps:[/cyan]")
                console.print("  • Check if the email addresses are correct")
                console.print("  • Try searching with just one criteria (FROM or TO)")
                console.print("  • Check if the email was sent recently (might need time to sync)")
                console.print("  • Try expanding the date range")
                return []
            
            # Remove duplicates and limit results
            unique_emails = []
            seen_ids = set()
            for email in all_emails:
                if email['id'] not in seen_ids:
                    unique_emails.append(email)
                    seen_ids.add(email['id'])
            
            # Limit final results
            limit = filters.get('limit', 100)
            if len(unique_emails) > limit:
                unique_emails = unique_emails[-limit:]
            
            console.print(f"[green]Total unique emails found: {len(unique_emails)}[/green]")
            return unique_emails
            
        except Exception as e:
            console.print(f"[red]Error in advanced filtering: {str(e)}[/red]")
            return []
    
    def _suggest_mailbox(self, filters: Dict) -> str:
        """Suggest the best mailbox to search based on criteria"""
        if filters.get('from'):
            # If searching for emails FROM a specific address, likely want sent emails
            return '[Gmail]/Sent Mail'
        else:
            # Default to inbox for emails TO someone
            return 'INBOX'
    
    def _get_available_mailboxes(self) -> List[str]:
        """Get list of available mailboxes/folders"""
        try:
            result, folders = self.imap_server.list()
            if result == 'OK':
                mailbox_list = []
                for folder in folders:
                    # Parse folder name from IMAP LIST response
                    folder_str = folder.decode('utf-8', errors='ignore')
                    # Extract folder name (usually the last part after quotes)
                    if '"' in folder_str:
                        parts = folder_str.split('"')
                        if len(parts) >= 4:
                            folder_name = parts[-2]
                            mailbox_list.append(folder_name)
                    else:
                        # Fallback parsing
                        parts = folder_str.split()
                        if len(parts) >= 3:
                            folder_name = ' '.join(parts[2:])
                            mailbox_list.append(folder_name)
                return mailbox_list
            else:
                return ['INBOX']  # Fallback to INBOX
        except Exception as e:
            console.print(f"[yellow]⚠️ Could not get folder list: {str(e)[:50]}[/yellow]")
            return ['INBOX']
    
    def _fetch_emails_details(self, message_ids: List, limit: int = None) -> List[Dict]:
        """Helper method to fetch detailed email information with robust error handling"""
        if limit and len(message_ids) > limit:
            message_ids = message_ids[-limit:]  # Get most recent emails
        
        email_list = []
        errors = 0
        
        console.print(f"[blue]🔄 Processing {len(message_ids)} emails... (Press Ctrl+C to cancel)[/blue]")
        
        with Progress() as progress:
            task = progress.add_task(f"Processing {len(message_ids)} emails...", total=len(message_ids))
            
            for i, msg_id in enumerate(message_ids):
                try:
                    # First try to get just headers for speed
                    result, msg_data = self.imap_server.fetch(msg_id, '(BODY[HEADER.FIELDS (FROM TO CC SUBJECT DATE)] BODY[TEXT])')
                    
                    if result == 'OK' and msg_data and msg_data[0]:
                        try:
                            # Try to get full message if headers work
                            result2, full_msg_data = self.imap_server.fetch(msg_id, '(RFC822)')
                            
                            if result2 == 'OK' and full_msg_data and full_msg_data[0]:
                                email_body = full_msg_data[0][1]
                                email_message = email.message_from_bytes(email_body)
                                
                                email_info = {
                                    'id': msg_id.decode() if isinstance(msg_id, bytes) else str(msg_id),
                                    'from': email_message.get('From', 'Unknown'),
                                    'subject': email_message.get('Subject', 'No Subject'),
                                    'date': email_message.get('Date', 'Unknown'),
                                    'to': email_message.get('To', 'Unknown'),
                                    'cc': email_message.get('Cc', ''),
                                    'body': self._get_email_body(email_message)
                                }
                                email_list.append(email_info)
                            else:
                                # Fallback to header-only data
                                header_data = msg_data[0][1].decode('utf-8', errors='ignore')
                                email_info = self._parse_header_data(header_data, msg_id)
                                email_list.append(email_info)
                                
                        except Exception as e:
                            # Fallback to header-only data
                            try:
                                header_data = msg_data[0][1].decode('utf-8', errors='ignore')
                                email_info = self._parse_header_data(header_data, msg_id)
                                email_list.append(email_info)
                            except Exception:
                                errors += 1
                                console.print(f"[yellow]⚠️ Skipping problematic email {i+1}/{len(message_ids)}[/yellow]")
                    else:
                        errors += 1
                        
                except Exception as e:
                    errors += 1
                    if errors > 5:  # Stop if too many errors
                        console.print(f"[red]Too many errors encountered, stopping...[/red]")
                        break
                
                progress.update(task, advance=1)
                
                # Add small delay to prevent overwhelming the server
                if i % 5 == 0 and i > 0:
                    time.sleep(0.1)
        
        if errors > 0:
            console.print(f"[yellow]⚠️ Skipped {errors} problematic emails[/yellow]")
        
        return email_list
    
    def _parse_header_data(self, header_data: str, msg_id) -> Dict:
        """Parse email header data as fallback"""
        import re
        
        email_info = {
            'id': msg_id.decode() if isinstance(msg_id, bytes) else str(msg_id),
            'from': 'Unknown',
            'subject': 'No Subject',
            'date': 'Unknown',
            'to': 'Unknown',
            'cc': '',
            'body': 'Body not available'
        }
        
        try:
            lines = header_data.split('\n')
            for line in lines:
                if line.startswith('From:'):
                    email_info['from'] = line.replace('From:', '').strip()[:100]
                elif line.startswith('To:'):
                    email_info['to'] = line.replace('To:', '').strip()[:100]
                elif line.startswith('Cc:'):
                    email_info['cc'] = line.replace('Cc:', '').strip()[:100]
                elif line.startswith('Subject:'):
                    email_info['subject'] = line.replace('Subject:', '').strip()[:100]
                elif line.startswith('Date:'):
                    email_info['date'] = line.replace('Date:', '').strip()[:50]
        except Exception:
            pass
        
        return email_info
    
    def _get_email_body(self, email_message) -> str:
        """Extract email body content with timeout protection"""
        try:
            # Quick extraction with limits to prevent hanging
            if email_message.is_multipart():
                body = ""
                part_count = 0
                for part in email_message.walk():
                    part_count += 1
                    if part_count > 10:  # Limit parts to process
                        break
                    
                    content_type = part.get_content_type()
                    if content_type in ["text/plain", "text/html"]:
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                if isinstance(payload, bytes):
                                    body = payload.decode('utf-8', errors='ignore')
                                else:
                                    body = str(payload)
                                break
                        except Exception:
                            continue
                return body if body else "No readable text content"
            else:
                try:
                    payload = email_message.get_payload(decode=True)
                    if payload:
                        if isinstance(payload, bytes):
                            return payload.decode('utf-8', errors='ignore')
                        else:
                            return str(payload)
                    else:
                        return "No email content"
                except Exception:
                    return "Unable to decode email body"
        except Exception as e:
            return f"Error extracting body: {str(e)[:100]}"
    
    def get_email_statistics(self, mailbox: str = 'INBOX') -> Dict:
        """Get general email statistics"""
        if not self.connected:
            console.print("[red]Not connected to email server![/red]")
            return {}
        
        try:
            self.imap_server.select(mailbox)
            
            # Get all emails (just headers for efficiency)
            result, message_ids = self.imap_server.search(None, 'ALL')
            
            if result != 'OK':
                return {}
            
            message_ids = message_ids[0].split()
            total_emails = len(message_ids)
            
            console.print(f"[yellow]Analyzing {total_emails} emails...[/yellow]")
            
            sender_count = Counter()
            subject_keywords = Counter()
            
            # Sample a subset if too many emails (for performance)
            sample_size = min(1000, total_emails)
            if total_emails > sample_size:
                message_ids = message_ids[-sample_size:]  # Get most recent
                console.print(f"[yellow]Sampling {sample_size} most recent emails for analysis...[/yellow]")
            
            with Progress() as progress:
                task = progress.add_task("Analyzing emails...", total=len(message_ids))
                
                for msg_id in message_ids:
                    result, msg_data = self.imap_server.fetch(msg_id, '(BODY[HEADER.FIELDS (FROM SUBJECT DATE)])')
                    
                    if result == 'OK':
                        header_data = msg_data[0][1].decode('utf-8', errors='ignore')
                        
                        # Extract sender
                        from_match = re.search(r'From:.*?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', header_data)
                        if from_match:
                            sender_count[from_match.group(1).lower()] += 1
                        
                        # Extract subject keywords
                        subject_match = re.search(r'Subject:\s*(.+)', header_data)
                        if subject_match:
                            subject = subject_match.group(1).strip()
                            # Extract meaningful words (exclude common words)
                            words = re.findall(r'\b[a-zA-Z]{3,}\b', subject.lower())
                            for word in words:
                                if word not in ['the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'can', 'had', 'was', 'one', 'our', 'out', 'day', 'get', 'has', 'him', 'his', 'how', 'its', 'new', 'now', 'old', 'see', 'two', 'who', 'boy', 'did', 'may', 'say', 'she', 'use', 'her', 'way', 'will', 'your']:
                                    subject_keywords[word] += 1
                    
                    progress.update(task, advance=1)
            
            return {
                'total_emails': total_emails,
                'analyzed_emails': len(message_ids),
                'top_senders': sender_count.most_common(10),
                'top_subject_keywords': subject_keywords.most_common(15)
            }
            
        except Exception as e:
            console.print(f"[red]Error getting statistics: {str(e)}[/red]")
            return {}
    
    def display_email_list(self, emails: List[Dict], title: str = "Email Results"):
        """Display email list in a formatted table"""
        if not emails:
            console.print("[yellow]No emails to display[/yellow]")
            return
        
        table = Table(title=title, show_header=True, header_style="bold magenta")
        table.add_column("From", style="cyan")
        table.add_column("Subject", style="white")
        table.add_column("Date", style="green")
        table.add_column("Mailbox", style="yellow", max_width=15)
        
        for email_info in emails[:50]:  # Limit display to 50 emails
            # Clean up the from field
            from_field = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', email_info['from'])
            from_display = from_field.group(1) if from_field else email_info['from']
            
            # Truncate long subjects
            subject = email_info['subject'][:60] + "..." if len(email_info['subject']) > 60 else email_info['subject']
            
            # Parse and format date
            try:
                parsed_date = email.utils.parsedate_to_datetime(email_info['date'])
                date_display = parsed_date.strftime("%Y-%m-%d %H:%M")
            except:
                date_display = email_info['date'][:20]
            
            mailbox = email_info.get('mailbox', 'INBOX')
            table.add_row(from_display, subject, date_display, mailbox)
        
        console.print(table)
        
        if len(emails) > 50:
            console.print(f"[yellow]... and {len(emails) - 50} more emails[/yellow]")
    
    def display_statistics(self, stats: Dict):
        """Display email statistics in formatted panels"""
        if not stats:
            return
        
        # General stats panel
        general_panel = Panel(
            f"Total Emails: {stats['total_emails']:,}\nAnalyzed: {stats['analyzed_emails']:,}",
            title="📧 General Statistics",
            style="blue"
        )
        console.print(general_panel)
        
        # Top senders table
        if stats['top_senders']:
            sender_table = Table(title="👤 Top Email Senders", show_header=True, header_style="bold cyan")
            sender_table.add_column("Sender", style="white")
            sender_table.add_column("Count", style="green", justify="right")
            
            for sender, count in stats['top_senders']:
                sender_table.add_row(sender, str(count))
            
            console.print(sender_table)
        
        # Top subject keywords table
        if stats['top_subject_keywords']:
            keyword_table = Table(title="🔤 Top Subject Keywords", show_header=True, header_style="bold yellow")
            keyword_table.add_column("Keyword", style="white")
            keyword_table.add_column("Frequency", style="green", justify="right")
            
            for keyword, count in stats['top_subject_keywords']:
                keyword_table.add_row(keyword, str(count))
            
            console.print(keyword_table)
    
    def generate_ai_analysis_prompt(self, emails: List[Dict], user_request: str) -> str:
        """Generate a ChatGPT prompt for email analysis"""
        
        if not emails:
            return "No emails to analyze."
        
        # Prepare email data for analysis
        email_data = []
        for i, email in enumerate(emails[:20], 1):  # Limit to 20 emails for token management
            email_summary = {
                'email_number': i,
                'from': email.get('from', 'Unknown'),
                'to': email.get('to', 'Unknown'),
                'cc': email.get('cc', ''),
                'subject': email.get('subject', 'No Subject'),
                'date': email.get('date', 'Unknown'),
                'body_full': email.get('body', '')
            }
            email_data.append(email_summary)
        
        # Generate the ChatGPT prompt
        prompt = f"""
ROLE: You are an expert email analyst with deep experience in business communication, project management, and professional correspondence analysis.

TASK: Analyze the following {len(email_data)} emails and provide insights based on this specific request:

USER REQUEST: "{user_request}"

EMAIL DATA:
"""
        
        for email in email_data:
            prompt += f"""
EMAIL #{email['email_number']}:
- From: {email['from']}
- To: {email['to']}
- CC: {email['cc']}
- Subject: {email['subject']}
- Date: {email['date']}
- Body: {email['body_full']}
---
"""
        
        prompt += f"""

ANALYSIS INSTRUCTIONS:
1. Focus specifically on what the user requested: "{user_request}"
2. Provide actionable insights and specific findings
3. Include relevant statistics and patterns you observe
4. Quote specific email content when relevant
5. Organize your response clearly with headers and bullet points
6. If looking for specific information, highlight the exact details found
7. Identify any trends, recurring themes, or important patterns
8. Suggest any follow-up actions or recommendations based on the analysis

RESPONSE FORMAT:
- Start with a brief executive summary
- Provide detailed analysis addressing the user's specific request
- Include key findings and statistics
- End with actionable recommendations

Please provide a comprehensive analysis focusing on the user's specific request.
"""
        
        return prompt
    
    def display_ai_prompt(self, emails: List[Dict], user_request: str):
        """Display the generated AI prompt and save it to clipboard-ready format"""
        
        prompt = self.generate_ai_analysis_prompt(emails, user_request)
        
        # Create a panel to display the prompt
        console.print(Panel.fit(
            "[bold cyan]🤖 ChatGPT Analysis Prompt Generated![/bold cyan]\n"
            f"[yellow]For {len(emails)} filtered emails[/yellow]",
            style="green"
        ))
        
        console.print("\n[bold yellow]📋 COPY THIS PROMPT TO CHATGPT:[/bold yellow]")
        console.print("="*80)
        console.print(prompt)
        console.print("="*80)
        
        # Save to file for easy access
        try:
            filename = f"chatgpt_prompt_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(prompt)
            console.print(f"\n[green]✅ Prompt saved to file: {filename}[/green]")
        except Exception as e:
            console.print(f"[yellow]⚠️ Could not save to file: {str(e)}[/yellow]")
        
        console.print(f"\n[bold blue]📊 Analysis Details:[/bold blue]")
        console.print(f"  • Total emails analyzed: {len(emails)}")
        console.print(f"  • User request: '{user_request}'")
        console.print(f"  • Prompt length: ~{len(prompt)} characters")
        console.print(f"\n[bold green]🚀 Next Steps:[/bold green]")
        console.print("  1. Copy the prompt above")
        console.print("  2. Go to ChatGPT (https://chat.openai.com)")
        console.print("  3. Paste the prompt and get your analysis!")
    
    def interactive_advanced_filter(self) -> List[Dict]:
        """Interactive advanced email filtering"""
        console.print("\n[bold cyan]🔍 Advanced Email Filter[/bold cyan]")
        console.print("[yellow]Leave fields empty to skip that filter[/yellow]")
        
        filters = {}
        
        # Get filter criteria
        from_email = Prompt.ask("Filter by FROM email", default="").strip()
        if from_email:
            filters['from'] = from_email
        
        to_email = Prompt.ask("Filter by TO email", default="").strip()
        if to_email:
            filters['to'] = to_email
        
        cc_email = Prompt.ask("Filter by CC email", default="").strip()
        if cc_email:
            filters['cc'] = cc_email
        
        subject_keyword = Prompt.ask("Filter by SUBJECT keyword", default="").strip()
        if subject_keyword:
            filters['subject'] = subject_keyword
        
        # Date range filter
        use_date_filter = Confirm.ask("Filter by date range?", default=False)
        if use_date_filter:
            console.print("[yellow]Date format: DD-Mon-YYYY (e.g., 01-Jan-2024)[/yellow]")
            start_date = Prompt.ask("Start date (DD-Mon-YYYY)").strip()
            end_date = Prompt.ask("End date (DD-Mon-YYYY)").strip()
            
            if start_date and end_date:
                filters['start_date'] = start_date
                filters['end_date'] = end_date
        
        # Result limit
        limit = Prompt.ask("Maximum number of results", default="50")
        try:
            filters['limit'] = int(limit)
        except:
            filters['limit'] = 50
        
        if not any(filters.values()):
            console.print("[red]No filters specified![/red]")
            return []
        
        # Show what filters will be applied
        console.print(f"\n[bold blue]📋 Applied Filters:[/bold blue]")
        for key, value in filters.items():
            if value and key != 'limit':
                console.print(f"  • {key.upper()}: {value}")
        
        console.print(f"  • LIMIT: {filters['limit']}")
        
        if not Confirm.ask("\nProceed with these filters?", default=True):
            return []
        
        # Apply filters
        return self.advanced_email_filter(filters)
    
    def interactive_ai_analysis(self, emails: List[Dict]):
        """Interactive AI analysis prompt generation"""
        if not emails:
            console.print("[red]No emails to analyze![/red]")
            return
        
        console.print(f"\n[bold cyan]🤖 AI Analysis for {len(emails)} emails[/bold cyan]")
        
        # Predefined analysis templates
        console.print("\n[yellow]📝 Common Analysis Types:[/yellow]")
        console.print("1. Project status and progress summary")
        console.print("2. Action items and task assignments")
        console.print("3. Communication patterns and frequency")
        console.print("4. Meeting schedules and important dates")  
        console.print("5. Financial discussions and budgets")
        console.print("6. Technical issues and solutions")
        console.print("7. Client feedback and requirements")
        console.print("8. Custom analysis request")
        
        choice = Prompt.ask("Choose analysis type or 8 for custom", choices=["1", "2", "3", "4", "5", "6", "7", "8"])
        
        templates = {
            "1": "Analyze these emails for project status updates, milestones, deliverables, and overall progress. Identify any blockers, completed tasks, and upcoming deadlines.",
            "2": "Extract all action items, task assignments, and responsibilities mentioned in these emails. List who is assigned what, deadlines, and priority levels.",
            "3": "Analyze communication patterns: who emails whom most frequently, response times, communication style, and professional relationships.",
            "4": "Find all meeting schedules, calendar invitations, important dates, and time-sensitive information in these emails.",
            "5": "Identify any financial discussions, budget mentions, cost estimates, pricing, payments, or monetary transactions referenced in these emails.",
            "6": "Analyze technical discussions, bug reports, system issues, solutions provided, and technical decisions made in these emails.",
            "7": "Extract client feedback, requirements, change requests, satisfaction levels, and customer communication insights from these emails."
        }
        
        if choice == "8":
            user_request = Prompt.ask("Enter your custom analysis request")
        else:
            user_request = templates[choice]
            console.print(f"\n[green]Selected analysis: {user_request}[/green]")
            
            if not Confirm.ask("Use this analysis template?", default=True):
                user_request = Prompt.ask("Enter your custom analysis request")
        
        # Generate and display the prompt
        self.display_ai_prompt(emails, user_request)

def main():
    parser = argparse.ArgumentParser(description="AI Email Scraping Tool")
    parser.add_argument('--email', '-e', help="Email address")
    parser.add_argument('--server', '-s', help="IMAP server (auto-detected if not provided)")
    parser.add_argument('--port', '-p', type=int, default=993, help="IMAP port (default: 993)")
    
    args = parser.parse_args()
    
    console.print(Panel.fit(
        "[bold blue]🤖 AI Email Scraping Tool[/bold blue]\n"
        "[yellow]Analyze your emails on-the-go without storing data![/yellow]",
        style="blue"
    ))
    
    scraper = EmailScraper()
    
    try:
        # Get credentials
        if not args.email:
            email_address = Prompt.ask("📧 Enter your email address")
        else:
            email_address = args.email
        
        password = getpass.getpass("🔒 Enter your password: ")
        
        # Connect to email
        if not scraper.connect_to_email(email_address, password, args.server, args.port):
            sys.exit(1)
        
        # Main menu loop
        while True:
            console.print("\n" + "="*60)
            console.print("[bold cyan]🤖 AI Email Scraping Tool - Main Menu[/bold cyan]")
            console.print("="*60)
            
            console.print("\n[bold yellow]📊 BASIC ANALYSIS[/bold yellow]")
            console.print("1. 📈 Get general email statistics")
            console.print("2. 👤 Count emails from specific sender")
            console.print("3. 🔍 Search emails by subject keyword")
            console.print("4. 📋 List emails from specific sender")
            
            console.print("\n[bold yellow]🔍 ADVANCED FILTERING[/bold yellow]")
            console.print("5. 📧 Filter by TO recipient")
            console.print("6. 📬 Filter by CC recipient") 
            console.print("7. 📅 Filter by date range")
            console.print("8. 🎯 Advanced multi-filter search")
            console.print("11. 📤 Search SENT emails (emails you sent)")
            console.print("12. ⚡ Quick find TODAY's sent emails")
            
            console.print("\n[bold yellow]🤖 AI ANALYSIS[/bold yellow]")
            console.print("9. 🧠 Generate AI analysis prompt (filter first)")
            
            console.print("\n[bold yellow]⚙️ SYSTEM[/bold yellow]")
            console.print("10. 🚪 Exit")
            
            choice = Prompt.ask("Choose an option", choices=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"])
            
            if choice == "1":
                console.print("\n[bold]Getting email statistics...[/bold]")
                stats = scraper.get_email_statistics()
                scraper.display_statistics(stats)
            
            elif choice == "2":
                sender = Prompt.ask("\n👤 Enter sender email address")
                emails = scraper.search_emails_by_sender(sender)
                console.print(f"\n[bold green]Found {len(emails)} emails from {sender}[/bold green]")
                
                if emails and Confirm.ask("Would you like to see the email list?"):
                    scraper.display_email_list(emails, f"Emails from {sender}")
                
                # Offer AI analysis
                if emails and Confirm.ask("Generate AI analysis for these emails?", default=False):
                    scraper.interactive_ai_analysis(emails)
            
            elif choice == "3":
                keyword = Prompt.ask("\n🔍 Enter keyword to search in subjects")
                limit = Prompt.ask("Maximum number of results", default="50")
                try:
                    limit = int(limit)
                except:
                    limit = 50
                
                emails = scraper.search_by_subject_keyword(keyword, limit=limit)
                console.print(f"\n[bold green]Found {len(emails)} emails with '{keyword}' in subject[/bold green]")
                
                if emails:
                    scraper.display_email_list(emails, f"Emails with '{keyword}' in subject")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(emails)
            
            elif choice == "4":
                sender = Prompt.ask("\n👤 Enter sender email address")
                emails = scraper.search_emails_by_sender(sender)
                
                if emails:
                    scraper.display_email_list(emails, f"Emails from {sender}")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(emails)
            
            elif choice == "5":
                to_email = Prompt.ask("\n📧 Enter recipient email address (TO field)")
                emails = scraper.search_emails_by_to(to_email)
                
                if emails:
                    console.print(f"\n[bold green]Found {len(emails)} emails sent to {to_email}[/bold green]")
                    scraper.display_email_list(emails, f"Emails to {to_email}")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(emails)
            
            elif choice == "6":
                cc_email = Prompt.ask("\n📬 Enter email address to search in CC field")
                emails = scraper.search_emails_by_cc(cc_email)
                
                if emails:
                    console.print(f"\n[bold green]Found {len(emails)} emails with {cc_email} in CC[/bold green]")
                    scraper.display_email_list(emails, f"Emails CC'd to {cc_email}")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(emails)
            
            elif choice == "7":
                console.print("\n📅 [yellow]Date format: DD-Mon-YYYY (e.g., 01-Jan-2024, 15-Dec-2023)[/yellow]")
                start_date = Prompt.ask("Enter start date (DD-Mon-YYYY)")
                end_date = Prompt.ask("Enter end date (DD-Mon-YYYY)")
                
                emails = scraper.search_emails_by_date_range(start_date, end_date)
                
                if emails:
                    console.print(f"\n[bold green]Found {len(emails)} emails from {start_date} to {end_date}[/bold green]")
                    scraper.display_email_list(emails, f"Emails from {start_date} to {end_date}")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(emails)
            
            elif choice == "8":
                emails = scraper.interactive_advanced_filter()
                
                if emails:
                    console.print(f"\n[bold green]Found {len(emails)} emails matching your filters[/bold green]")
                    scraper.display_email_list(emails, "Filtered Email Results")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(emails)
            
            elif choice == "9":
                console.print("\n[bold cyan]🤖 AI Analysis Prompt Generator[/bold cyan]")
                console.print("[yellow]First, you need to filter some emails to analyze.[/yellow]")
                console.print("[yellow]Choose from the filtering options above (1-8) first![/yellow]")
                
                # Quick filter option
                quick_choice = Prompt.ask(
                    "Quick filter options",
                    choices=["subject", "sender", "advanced", "cancel"],
                    default="cancel"
                )
                
                emails = []
                if quick_choice == "subject":
                    keyword = Prompt.ask("Enter keyword to search in subjects")
                    limit = Prompt.ask("Maximum number of results", default="20")
                    try:
                        limit = int(limit)
                    except:
                        limit = 20
                    emails = scraper.search_by_subject_keyword(keyword, limit=limit)
                
                elif quick_choice == "sender":
                    sender = Prompt.ask("Enter sender email address")
                    emails = scraper.search_emails_by_sender(sender)
                
                elif quick_choice == "advanced":
                    emails = scraper.interactive_advanced_filter()
                
                if emails:
                    scraper.interactive_ai_analysis(emails)
            
            elif choice == "10":
                break
            
            elif choice == "11":
                console.print("\n📤 [bold cyan]Search SENT Emails (Emails You Sent)[/bold cyan]")
                recipient = Prompt.ask("Enter recipient email address (who you sent to)", default="").strip()
                subject_keyword = Prompt.ask("Enter subject keyword (optional)", default="").strip()
                
                # Find the sent folder
                available_mailboxes = scraper._get_available_mailboxes()
                sent_folders = [box for box in available_mailboxes if 'sent' in box.lower() or 'enviado' in box.lower()]
                
                if not sent_folders:
                    console.print("[yellow]⚠️ Could not find a sent mail folder. Trying common names...[/yellow]")
                    sent_folders = ['[Gmail]/Sent Mail', 'Sent', 'Sent Items']
                
                console.print(f"[blue]📁 Available sent folders: {', '.join(sent_folders[:3])}[/blue]")
                
                emails_found = []
                for sent_folder in sent_folders[:2]:  # Try max 2 sent folders
                    try:
                        console.print(f"[blue]🔍 Searching in: {sent_folder}[/blue]")
                        scraper.imap_server.select(sent_folder)
                        search_criteria = []
                        
                        if recipient:
                            search_criteria.append(f'TO "{recipient}"')
                        if subject_keyword:
                            search_criteria.append(f'SUBJECT "{subject_keyword}"')
                        
                        if not search_criteria:
                            # Get recent sent emails (last 7 days)
                            from datetime import timedelta
                            week_ago = (datetime.now() - timedelta(days=7)).strftime('%d-%b-%Y')
                            search_criteria.append(f'SINCE "{week_ago}"')
                        
                        final_criteria = ' '.join(search_criteria) if search_criteria else 'ALL'
                        console.print(f"[blue]Search criteria: {final_criteria}[/blue]")
                        
                        result, message_ids = scraper.imap_server.search(None, final_criteria)
                        
                        if result == 'OK' and message_ids and message_ids[0]:
                            message_ids = message_ids[0].split()
                            if message_ids:
                                console.print(f"[green]Found {len(message_ids)} emails in {sent_folder}[/green]")
                                folder_emails = scraper._fetch_emails_details(message_ids, limit=25)
                                # Mark which folder they came from
                                for email in folder_emails:
                                    email['mailbox'] = sent_folder
                                emails_found.extend(folder_emails)
                                break  # Found emails, no need to search other folders
                        else:
                            console.print(f"[yellow]No emails found in {sent_folder}[/yellow]")
                            
                    except Exception as e:
                        console.print(f"[yellow]⚠️ Could not search in {sent_folder}: {str(e)[:50]}[/yellow]")
                        continue
                
                if emails_found:
                    scraper.display_email_list(emails_found, f"📤 Sent Emails ({len(emails_found)} found)")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(emails_found)
                else:
                    console.print("[yellow]No sent emails found with those criteria[/yellow]")
                    console.print("[cyan]💡 Troubleshooting tips:[/cyan]")
                    console.print("  • Try leaving both fields blank to see recent sent emails")
                    console.print("  • Check if the recipient email address is correct")
                    console.print("  • Your sent emails might be in a different folder")
                    
                    # Show available folders for debugging
                    console.print(f"[blue]📁 Your available folders: {', '.join(available_mailboxes[:10])}[/blue]")
            
            elif choice == "12":
                console.print("\n⚡ [bold cyan]Quick Find TODAY's Sent Emails[/bold cyan]")
                console.print("[yellow]This will try multiple methods to find emails you sent today[/yellow]")
                
                today = datetime.now().strftime('%d-%b-%Y')
                console.print(f"[blue]🗓️ Searching for emails sent on: {today}[/blue]")
                
                # Try different approaches to find sent emails
                found_emails = []
                
                # Method 1: Try to list all folders and find sent-like folders
                try:
                    result, folders = scraper.imap_server.list()
                    if result == 'OK':
                        console.print(f"[blue]📁 Found {len(folders)} total folders[/blue]")
                        
                        sent_folder_candidates = []
                        for folder in folders:
                            folder_str = folder.decode('utf-8', errors='ignore').lower()
                            if 'sent' in folder_str or 'enviado' in folder_str or 'wyslane' in folder_str:
                                # Extract the actual folder name
                                parts = folder.decode('utf-8', errors='ignore').split('"')
                                if len(parts) >= 3:
                                    folder_name = parts[-2]
                                    sent_folder_candidates.append(folder_name)
                                    console.print(f"[green]📤 Found potential sent folder: {folder_name}[/green]")
                        
                        # Try each sent folder candidate
                        for folder_name in sent_folder_candidates[:3]:  # Try max 3
                            try:
                                console.print(f"[blue]🔍 Trying folder: {folder_name}[/blue]")
                                scraper.imap_server.select(f'"{folder_name}"')
                                
                                # Search for today's emails
                                result, message_ids = scraper.imap_server.search(None, f'SINCE "{today}"')
                                
                                if result == 'OK' and message_ids and message_ids[0]:
                                    message_ids = message_ids[0].split()
                                    if message_ids:
                                        console.print(f"[green]✅ Found {len(message_ids)} emails in {folder_name}![/green]")
                                        
                                        # Get email details
                                        emails = scraper._fetch_emails_details(message_ids[:10], limit=10)  # Max 10
                                        for email in emails:
                                            email['mailbox'] = folder_name
                                        found_emails.extend(emails)
                                        
                                        if found_emails:
                                            break  # Found some, stop searching
                                        
                            except Exception as e:
                                console.print(f"[yellow]⚠️ Could not access {folder_name}: {str(e)[:50]}[/yellow]")
                                continue
                                
                except Exception as e:
                    console.print(f"[yellow]⚠️ Could not list folders: {str(e)[:50]}[/yellow]")
                
                # If no sent emails found, search INBOX for emails FROM your address (might catch some)
                if not found_emails:
                    console.print("[blue]🔄 No luck with sent folders, checking INBOX for any emails from you...[/blue]")
                    try:
                        scraper.imap_server.select('INBOX')
                        your_email = scraper.email_address
                        result, message_ids = scraper.imap_server.search(None, f'FROM "{your_email}" SINCE "{today}"')
                        
                        if result == 'OK' and message_ids and message_ids[0]:
                            message_ids = message_ids[0].split()
                            if message_ids:
                                console.print(f"[green]Found {len(message_ids)} emails from your address in INBOX[/green]")
                                emails = scraper._fetch_emails_details(message_ids[:5], limit=5)
                                for email in emails:
                                    email['mailbox'] = 'INBOX'
                                found_emails.extend(emails)
                    except Exception as e:
                        console.print(f"[yellow]⚠️ INBOX search failed: {str(e)[:50]}[/yellow]")
                
                # Display results
                if found_emails:
                    scraper.display_email_list(found_emails, f"⚡ Today's Sent Emails ({len(found_emails)} found)")
                    
                    # Look for the specific email to Abhishek
                    abhishek_emails = [e for e in found_emails if 'abhishek.chand@sigsi.com' in e.get('to', '').lower()]
                    if abhishek_emails:
                        console.print(f"[bold green]🎯 Found {len(abhishek_emails)} emails to Abhishek![/bold green]")
                    
                    # Offer AI analysis
                    if Confirm.ask("Generate AI analysis for these emails?", default=False):
                        scraper.interactive_ai_analysis(found_emails)
                else:
                    console.print("[yellow]😔 No sent emails found for today[/yellow]")
                    console.print("[cyan]💡 Possible reasons:[/cyan]")
                    console.print("  • The email might not have been sent yet")
                    console.print("  • It might be in a different folder")
                    console.print("  • The email server might need time to sync")
                    console.print("  • Try using option 11 for more detailed sent email search")
            
            if not Confirm.ask("\nContinue with another operation?", default=True):
                break
    
    except KeyboardInterrupt:
        console.print("\n[yellow]Operation cancelled by user[/yellow]")
    except Exception as e:
        console.print(f"\n[red]An error occurred: {str(e)}[/red]")
    finally:
        scraper.disconnect()
        console.print("\n[green]Thanks for using AI Email Scraping Tool! 👋[/green]")

if __name__ == "__main__":
    main() 