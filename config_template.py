#!/usr/bin/env python3
"""
Email Configuration Template
Copy this file to config.py and update with your settings
"""

# Email server configurations for different providers
EMAIL_CONFIGS = {
    # Gmail
    'gmail.com': {
        'imap_server': 'imap.gmail.com',
        'imap_port': 993,
        'notes': [
            'You need to use an App Password instead of your regular password',
            'Enable 2-factor authentication first',
            'Go to Google Account > Security > App passwords',
            'Generate a new app password for "Mail"',
        ]
    },
    
    # Outlook/Hotmail/Live
    'outlook.com': {
        'imap_server': 'outlook.office365.com',
        'imap_port': 993,
        'notes': [
            'Regular password should work',
            'If you have 2FA enabled, you might need an app password',
        ]
    },
    
    'hotmail.com': {
        'imap_server': 'outlook.office365.com',
        'imap_port': 993,
        'notes': [
            'Regular password should work',
            'If you have 2FA enabled, you might need an app password',
        ]
    },
    
    # Yahoo Mail
    'yahoo.com': {
        'imap_server': 'imap.mail.yahoo.com',
        'imap_port': 993,
        'notes': [
            'You need to generate an App Password',
            'Go to Yahoo Account Security > Generate app password',
            'Use the app password instead of your regular password',
        ]
    },
    
    # iCloud
    'icloud.com': {
        'imap_server': 'imap.mail.me.com',
        'imap_port': 993,
        'notes': [
            'You need to use an App-Specific Password',
            'Go to Apple ID > Sign-In and Security > App-Specific Passwords',
            'Generate a password for this application',
        ]
    },
    
    # AOL
    'aol.com': {
        'imap_server': 'imap.aol.com',
        'imap_port': 993,
        'notes': [
            'You might need to enable IMAP in your AOL mail settings',
            'Use your regular AOL password',
        ]
    }
}

def get_email_config(email_address):
    """Get configuration for an email provider"""
    domain = email_address.split('@')[1].lower()
    return EMAIL_CONFIGS.get(domain, {
        'imap_server': f'imap.{domain}',
        'imap_port': 993,
        'notes': [
            f'Auto-detected IMAP server: imap.{domain}',
            'If this doesn\'t work, check your email provider\'s IMAP settings',
        ]
    })

def print_email_setup_instructions(email_address):
    """Print setup instructions for the given email provider"""
    config = get_email_config(email_address)
    domain = email_address.split('@')[1].lower()
    
    print(f"\n{'='*60}")
    print(f"📧 SETUP INSTRUCTIONS FOR {domain.upper()}")
    print(f"{'='*60}")
    print(f"IMAP Server: {config['imap_server']}")
    print(f"IMAP Port: {config['imap_port']}")
    print(f"Security: SSL/TLS")
    
    print(f"\n📝 IMPORTANT NOTES:")
    for i, note in enumerate(config['notes'], 1):
        print(f"   {i}. {note}")
    
    print(f"\n⚠️  SECURITY REMINDERS:")
    print(f"   • Never share your credentials with anyone")
    print(f"   • Use app-specific passwords when possible")
    print(f"   • This tool doesn't store your credentials")
    print(f"   • All email analysis is done in real-time")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    # Example usage
    test_emails = ['user@gmail.com', 'user@outlook.com', 'user@yahoo.com']
    for email in test_emails:
        print_email_setup_instructions(email) 