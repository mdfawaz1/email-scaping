# 🤖 AI Email Scraping Tool

A powerful, privacy-focused email analysis tool that helps you scrape and analyze your email data **without storing anything locally**. Perfect for understanding your email patterns, finding specific communications, and getting insights from your inbox on-the-go!

## ✨ Features

- **📊 Email Statistics**: Get comprehensive analytics about your email patterns
- **👤 Sender Analysis**: Count and analyze emails from specific senders
- **🔍 Subject Search**: Search for keywords in email subjects
- **📋 Email Listing**: View detailed email lists with sender, subject, and date
- **🔒 Privacy-First**: No data storage - all analysis is done in real-time
- **🎨 Beautiful UI**: Rich terminal interface with progress bars and colored output
- **🔧 Auto-Detection**: Automatically detects IMAP settings for major email providers

## 🚀 Quick Start

### Prerequisites

- Python 3.7 or higher
- An email account with IMAP access enabled

### Installation

1. **Clone or download this repository**
```bash
git clone <repository-url>
cd email-scraping
```

2. **Install dependencies**
```bash
pip install -r requirements.txt
```

3. **Run the tool**
```bash
python email_scraper.py
```

## 📧 Email Provider Setup

### Gmail
1. Enable 2-Factor Authentication
2. Go to Google Account → Security → App passwords
3. Generate an app password for "Mail"
4. Use your email and the app password (not your regular password)

### Outlook/Hotmail/Live
- Regular password usually works
- If you have 2FA enabled, you might need an app password

### Yahoo Mail
1. Go to Yahoo Account Security
2. Generate an app password
3. Use the app password instead of your regular password

### iCloud
1. Go to Apple ID → Sign-In and Security
2. Generate an App-Specific Password
3. Use the app-specific password

## 🛠️ Usage

### Command Line Options

```bash
python email_scraper.py [options]

Options:
  -e, --email EMAIL     Your email address
  -s, --server SERVER   IMAP server (auto-detected if not provided)
  -p, --port PORT       IMAP port (default: 993)
```

### Interactive Menu

Once connected, you can:

1. **📊 Get general email statistics**
   - Total email count
   - Top email senders
   - Most common subject keywords

2. **👤 Count emails from specific sender**
   - Enter any email address
   - Get exact count of emails received
   - Option to view detailed list

3. **🔍 Search emails by subject keyword**
   - Search for any keyword in email subjects
   - Set maximum number of results
   - View matching emails with details

4. **📋 List emails from specific sender**
   - Get detailed list of all emails from a sender
   - Shows sender, subject, and date
   - Limited to 50 emails for readability

## 🔒 Privacy & Security

- **No Data Storage**: This tool doesn't save any email data to your computer
- **Real-time Analysis**: All processing happens in memory and is discarded after use
- **Secure Connection**: Uses SSL/TLS encryption for email server connections
- **Local Processing**: All analysis is done locally on your machine
- **No Tracking**: No analytics, tracking, or data collection

## 📊 Sample Output

### Email Statistics
```
📧 General Statistics
─────────────────────
Total Emails: 15,847
Analyzed: 1,000

👤 Top Email Senders
┌─────────────────────────────────┬───────┐
│ Sender                          │ Count │
├─────────────────────────────────┼───────┤
│ notifications@github.com        │   234 │
│ noreply@medium.com             │   156 │
│ team@slack.com                 │   89  │
└─────────────────────────────────┴───────┘

🔤 Top Subject Keywords
┌──────────────┬───────────┐
│ Keyword      │ Frequency │
├──────────────┼───────────┤
│ update       │    45     │
│ notification │    38     │
│ weekly       │    32     │
└──────────────┴───────────┘
```

### Email Search Results
```
Emails with 'invoice' in subject
┌─────────────────────────┬──────────────────────────┬─────────────────┐
│ From                    │ Subject                  │ Date            │
├─────────────────────────┼──────────────────────────┼─────────────────┤
│ billing@company.com     │ Invoice #12345 - Dec...  │ 2023-12-15 10:30│
│ payments@service.com    │ Monthly Invoice Ready    │ 2023-12-01 09:15│
└─────────────────────────┴──────────────────────────┴─────────────────┘
```

## 🔧 Troubleshooting

### Common Issues

**"Authentication failed"**
- For Gmail: Make sure you're using an app password, not your regular password
- For Yahoo: Generate and use an app-specific password
- For Outlook: Try your regular password first, then app password if needed

**"Connection timeout"**
- Check your internet connection
- Verify IMAP is enabled in your email settings
- Some corporate networks block IMAP ports

**"No emails found"**
- Check the sender email address for typos
- Try searching with partial email addresses
- Verify the sender has actually sent you emails

### Getting Help

If you encounter issues:
1. Run the configuration helper: `python config_template.py`
2. Check your email provider's IMAP settings
3. Ensure IMAP access is enabled in your email account settings

## 🤝 Contributing

Feel free to contribute by:
- Adding support for more email providers
- Improving the analysis algorithms
- Adding new search and analysis features
- Enhancing the user interface

## ⚠️ Disclaimers

- This tool requires your email credentials to function
- Always use app-specific passwords when available
- Be cautious when using on shared computers
- Some email providers may have rate limits for IMAP access
- Large mailboxes may take time to analyze

## 📄 License

This project is open source. Use responsibly and respect email provider terms of service.

---

**Happy Email Scraping! 🚀** 