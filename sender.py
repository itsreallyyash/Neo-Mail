"""
Simple Email Testing Setup - Easiest Options for Testing
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# OPTION 1: Gmail (Recommended for Testing)
# Steps to setup:
# 1. Go to Gmail settings -> Security -> 2-Step Verification (enable if not already)
# 2. Go to App Passwords -> Generate app password for "Mail"
# 3. Use the 16-character app password (not your regular Gmail password)

GMAIL_CONFIG = {
    'server': 'smtp.gmail.com',
    'port': 587,
    'username': 'yash12628@gmail.com',  # Your Gmail address
    'password': 'rmvj seus ezmj vobu',  # Gmail App Password
    'from_email': 'yash12628@gmail.com'
}

# OPTION 2: Outlook/Hotmail (Alternative)
OUTLOOK_CONFIG = {
    'server': 'smtp-mail.outlook.com',
    'port': 587,
    'username': 'your_email@outlook.com',  # or @hotmail.com
    'password': 'your_outlook_password',
    'from_email': 'your_email@outlook.com'
}

# OPTION 3: For Local Testing Only (Doesn't actually send emails)
# This will print emails to console instead of sending them
class MockEmailSender:
    def __init__(self):
        self.sent_emails = []
    
    def send_email(self, to_email, subject, body):
        print(f"\n{'='*50}")
        print(f"MOCK EMAIL SENT")
        print(f"To: {to_email}")
        print(f"Subject: {subject}")
        print(f"{'='*50}")
        print(body)
        print(f"{'='*50}\n")
        
        self.sent_emails.append({
            'to': to_email,
            'subject': subject,
            'body': body
        })

# Quick Test Function
def test_email_connection(config):
    """Test if email configuration works"""
    try:
        server = smtplib.SMTP(config['server'], config['port'])
        server.starttls()
        server.login(config['username'], config['password'])
        server.quit()
        print("✅ Email configuration successful!")
        return True
    except Exception as e:
        print(f"❌ Email configuration failed: {str(e)}")
        return False

# Simple Send Function
def send_test_email(config, to_email, subject="Test Email", body="This is a test email."):
    """Send a simple test email"""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = config['from_email']
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Add body
        msg.attach(MIMEText(body, 'html' if '<' in body else 'plain'))
        
        # Connect and send
        server = smtplib.SMTP(config['server'], config['port'])
        server.starttls()
        server.login(config['username'], config['password'])
        
        text = msg.as_string()
        server.sendmail(config['from_email'], to_email, text)
        server.quit()
        
        print(f"✅ Test email sent successfully to {to_email}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send email: {str(e)}")
        return False

# Modified Investment Email Generator for Easy Testing
from investment_email_generator import InvestmentEmailGenerator

class TestableEmailGenerator(InvestmentEmailGenerator):
    def __init__(self, use_mock=False):
        super().__init__()
        self.use_mock = use_mock
        if use_mock:
            self.mock_sender = MockEmailSender()
    
    def send_emails_easy(self, emails, config=None, test_email=None):
        """Simplified email sending with better error handling"""
        
        if self.use_mock:
            # Mock sending - just print to console
            for investor_name, email_data in emails.items():
                self.mock_sender.send_email(
                    test_email or email_data['email'],
                    email_data['subject'],
                    email_data['content']
                )
            return True
        
        if not config:
            print("❌ Email configuration required for actual sending")
            return False
        
        # Test connection first
        if not test_email_connection(config):
            return False
        
        # Send emails
        success_count = 0
        for investor_name, email_data in emails.items():
            recipient = test_email or email_data['email']  # Use test email if provided
            
            if send_test_email(config, recipient, email_data['subject'], email_data['content']):
                success_count += 1
                print(f"✅ Email sent to {investor_name}")
            else:
                print(f"❌ Failed to send email to {investor_name}")
        
        print(f"\n📊 Results: {success_count}/{len(emails)} emails sent successfully")
        return success_count == len(emails)

# Example Usage
if __name__ == "__main__":
    print("🚀 Investment Email Generator - Testing Setup")
    
    # Ask for CSV file path
    csv_file = input("\nEnter CSV file path (or press Enter for sample data): ").strip()
    
    print("\nChoose testing option:")
    print("1. Mock Mode (No actual emails sent - Console output only)")
    print("2. Gmail Test (Requires Gmail App Password)")
    print("3. Outlook Test")
    
    choice = input("\nEnter choice (1-3): ").strip()
    
    if choice == "1":
        # Mock mode testing
        print("\n📝 MOCK MODE - Emails will be printed to console")
        generator = TestableEmailGenerator(use_mock=True)
        
        # Load data from CSV or use sample data
        if csv_file and csv_file.lower().endswith('.csv'):
            try:
                generator.load_data(file_path=csv_file)
                print(f"✅ Loaded data from {csv_file}")
            except Exception as e:
                print(f"❌ Error loading CSV: {e}")
                print("Using sample data instead...")
                generator.load_data()
        else:
            print("Using sample data...")
            generator.load_data()
        
        company_config = {
            'company_name': 'Test Company Pvt Ltd',
            'period_month': 'April',
            'period_year': '2025',
            'processing_date': 'April 25, 2025',
            'company_signature': 'Test Company',
            'company_address': 'Test Address\nTest City - 123456',
            'company_email': 'test@testcompany.com'
        }
        
        emails = generator.generate_emails(company_config)
        generator.send_emails_easy(emails)
        
    elif choice == "2":
        # Gmail testing
        print("\n📧 GMAIL MODE")
        
        # Get CSV file
        csv_file = input("Enter CSV file path (or press Enter for sample data): ").strip()
        
        gmail_user = input("Enter your Gmail address: ").strip()
        gmail_app_password = input("Enter your Gmail App Password (16 characters): ").strip()
        test_recipient = input("Enter test recipient email: ").strip()
        
        gmail_config = {
            'server': 'smtp.gmail.com',
            'port': 587,
            'username': gmail_user,
            'password': gmail_app_password,
            'from_email': gmail_user
        }
        
        # Test connection
        if test_email_connection(gmail_config):
            generator = TestableEmailGenerator(use_mock=False)
            
            # Load data from CSV or use sample data
            if csv_file and csv_file.lower().endswith('.csv'):
                try:
                    generator.load_data(file_path=csv_file)
                    print(f"✅ Loaded data from {csv_file}")
                except Exception as e:
                    print(f"❌ Error loading CSV: {e}")
                    print("Using sample data instead...")
                    generator.load_data()
            else:
                print("Using sample data...")
                generator.load_data()
            
            company_config = {
                'company_name': 'Test Company Pvt Ltd',
                'period_month': 'April',
                'period_year': '2025',
                'processing_date': 'April 25, 2025',
                'company_signature': 'Test Company',
                'company_address': 'Test Address\nTest City - 123456',
                'company_email': 'test@testcompany.com'
            }
            
            emails = generator.generate_emails(company_config)
            generator.send_emails_easy(emails, gmail_config, test_recipient)
    
    elif choice == "3":
        # Outlook testing
        print("\n📧 OUTLOOK MODE")
        
        # Get CSV file
        csv_file = input("Enter CSV file path (or press Enter for sample data): ").strip()
        
        outlook_user = input("Enter your Outlook email: ").strip()
        outlook_password = input("Enter your Outlook password: ").strip()
        test_recipient = input("Enter test recipient email: ").strip()
        
        outlook_config = {
            'server': 'smtp-mail.outlook.com',
            'port': 587,
            'username': outlook_user,
            'password': outlook_password,
            'from_email': outlook_user
        }
        
        # Test connection
        if test_email_connection(outlook_config):
            generator = TestableEmailGenerator(use_mock=False)
            
            # Load data from CSV or use sample data
            if csv_file and csv_file.lower().endswith('.csv'):
                try:
                    generator.load_data(file_path=csv_file)
                    print(f"✅ Loaded data from {csv_file}")
                except Exception as e:
                    print(f"❌ Error loading CSV: {e}")
                    print("Using sample data instead...")
                    generator.load_data()
            else:
                print("Using sample data...")
                generator.load_data()
            
            company_config = {
                'company_name': 'Test Company Pvt Ltd',
                'period_month': 'April',
                'period_year': '2025',
                'processing_date': 'April 25, 2025',
                'company_signature': 'Test Company',
                'company_address': 'Test Address\nTest City - 123456',
                'company_email': 'test@testcompany.com'
            }
            
            emails = generator.generate_emails(company_config)
            generator.send_emails_easy(emails, outlook_config, test_recipient)
    
    else:
        print("Invalid choice!")