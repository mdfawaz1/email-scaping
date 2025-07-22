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
from collections import defaultdict, Counter
from datetime import datetime
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
        """Connect to email server using IMAP"""
        try:
            self.email_address = email_address
            
            # Auto-detect server if not provided
            if not server:
                server = self._detect_imap_server(email_address)
            
            # Create SSL context
            context = ssl.create_default_context()
            
            # Connect to server
            console.print(f"[yellow]Connecting to {server}:{port}...[/yellow]")
            self.imap_server = imaplib.IMAP4_SSL(server, port, ssl_context=context)
            
            # Login
            console.print("[yellow]Authenticating...[/yellow]")
            self.imap_server.login(email_address, password)
            
            self.connected = True
            console.print("[green]✓ Successfully connected to email server![/green]")
            return True
            
        except Exception as e:
            console.print(f"[red]❌ Failed to connect: {str(e)}[/red]")
            return False
    
    def _detect_imap_server(self, email_address: str) -> str:
        """Auto-detect IMAP server based on email domain"""
        domain = email_address.split('@')[1].lower()
        
        server_map = {
            'gmail.com': 'imap.gmail.com',
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
            
            with Progress() as progress:
                task = progress.add_task(f"Processing {len(message_ids)} emails...", total=len(message_ids))
                
                for msg_id in message_ids:
                    result, msg_data = self.imap_server.fetch(msg_id, '(RFC822)')
                    
                    if result == 'OK':
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)
                        
                        email_info = {
                            'id': msg_id.decode(),
                            'from': email_message.get('From', 'Unknown'),
                            'subject': email_message.get('Subject', 'No Subject'),
                            'date': email_message.get('Date', 'Unknown'),
                            'to': email_message.get('To', 'Unknown')
                        }
                        email_list.append(email_info)
                    
                    progress.update(task, advance=1)
            
            return email_list
            
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
            
            email_list = []
            
            with Progress() as progress:
                task = progress.add_task(f"Processing {len(message_ids)} emails...", total=len(message_ids))
                
                for msg_id in message_ids:
                    result, msg_data = self.imap_server.fetch(msg_id, '(RFC822)')
                    
                    if result == 'OK':
                        email_body = msg_data[0][1]
                        email_message = email.message_from_bytes(email_body)
                        
                        email_info = {
                            'id': msg_id.decode(),
                            'from': email_message.get('From', 'Unknown'),
                            'subject': email_message.get('Subject', 'No Subject'),
                            'date': email_message.get('Date', 'Unknown'),
                            'to': email_message.get('To', 'Unknown')
                        }
                        email_list.append(email_info)
                    
                    progress.update(task, advance=1)
            
            return email_list
            
        except Exception as e:
            console.print(f"[red]Error searching emails: {str(e)}[/red]")
            return []
    
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
            
            table.add_row(from_display, subject, date_display)
        
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
            console.print("\n" + "="*50)
            console.print("[bold cyan]What would you like to do?[/bold cyan]")
            console.print("1. 📊 Get general email statistics")
            console.print("2. 👤 Count emails from specific sender")
            console.print("3. 🔍 Search emails by subject keyword")
            console.print("4. 📋 List emails from specific sender")
            console.print("5. 🚪 Exit")
            
            choice = Prompt.ask("Choose an option", choices=["1", "2", "3", "4", "5"])
            
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
            
            elif choice == "4":
                sender = Prompt.ask("\n👤 Enter sender email address")
                emails = scraper.search_emails_by_sender(sender)
                
                if emails:
                    scraper.display_email_list(emails, f"Emails from {sender}")
            
            elif choice == "5":
                break
            
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