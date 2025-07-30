#!/usr/bin/env python3
"""
Outlook Email Search Script
Searches for emails containing "beat" in title or content from start of month to today.
Displays results in HTML format and saves to text file.
"""

import os
import sys
from datetime import datetime, date
from exchangelib import Credentials, Account, DELEGATE, Configuration
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import html
import re

# Disable SSL verification warnings (only if needed for corporate environments)
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

def get_start_of_month():
    """Get the first day of the current month."""
    today = date.today()
    return date(today.year, today.month, 1)

def connect_to_outlook(email, password, server=None):
    """
    Connect to Outlook account.
    
    Args:
        email (str): Email address
        password (str): Password
        server (str, optional): Exchange server URL
    
    Returns:
        Account: Connected account object
    """
    try:
        credentials = Credentials(email, password)
        
        if server:
            config = Configuration(server=server, credentials=credentials)
            account = Account(primary_smtp_address=email, config=config, 
                           autodiscover=False, access_type=DELEGATE)
        else:
            account = Account(primary_smtp_address=email, credentials=credentials, 
                           autodiscover=True, access_type=DELEGATE)
        
        print(f"Successfully connected to {email}")
        return account
    
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return None

def search_emails(account, start_date, end_date, search_term="beat"):
    """
    Search for emails containing the search term in title or content.
    
    Args:
        account: Outlook account object
        start_date: Start date for search
        end_date: End date for search
        search_term: Term to search for
    
    Returns:
        list: List of matching emails
    """
    try:
        # Search in inbox
        inbox = account.inbox
        
        # Filter emails by date range and search term
        emails = inbox.filter(
            received__gte=start_date,
            received__lte=end_date
        )
        
        matching_emails = []
        
        for email in emails:
            # Check if search term is in subject or body
            subject = email.subject.lower() if email.subject else ""
            body = email.body.lower() if email.body else ""
            
            if search_term.lower() in subject or search_term.lower() in body:
                matching_emails.append(email)
        
        print(f"Found {len(matching_emails)} emails containing '{search_term}'")
        return matching_emails
    
    except Exception as e:
        print(f"Error searching emails: {e}")
        return []

def create_html_report(emails, search_term):
    """
    Create HTML report of matching emails.
    
    Args:
        emails: List of email objects
        search_term: Search term used
    
    Returns:
        str: HTML content
    """
    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Outlook Email Search Results - "{search_term}"</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
        }}
        .email {{
            border: 1px solid #ddd;
            margin: 15px 0;
            padding: 15px;
            border-radius: 5px;
            background-color: #fafafa;
        }}
        .email-header {{
            background-color: #0078d4;
            color: white;
            padding: 10px;
            margin: -15px -15px 15px -15px;
            border-radius: 5px 5px 0 0;
        }}
        .email-subject {{
            font-weight: bold;
            font-size: 16px;
            margin-bottom: 5px;
        }}
        .email-meta {{
            font-size: 12px;
            opacity: 0.9;
        }}
        .email-body {{
            line-height: 1.6;
            white-space: pre-wrap;
            max-height: 300px;
            overflow-y: auto;
            border: 1px solid #eee;
            padding: 10px;
            background-color: white;
        }}
        .highlight {{
            background-color: #ffff99;
            padding: 2px 4px;
            border-radius: 3px;
        }}
        .summary {{
            background-color: #e8f4fd;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Outlook Email Search Results</h1>
        
        <div class="summary">
            <h3>Search Summary</h3>
            <p><strong>Search Term:</strong> "{search_term}"</p>
            <p><strong>Date Range:</strong> {get_start_of_month().strftime('%B %d, %Y')} to {date.today().strftime('%B %d, %Y')}</p>
            <p><strong>Total Emails Found:</strong> {len(emails)}</p>
            <p><strong>Report Generated:</strong> {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>
"""
    
    if not emails:
        html_content += """
        <div class="email">
            <p><em>No emails found matching the search criteria.</em></p>
        </div>
"""
    else:
        for i, email in enumerate(emails, 1):
            # Clean and escape HTML content
            subject = html.escape(email.subject) if email.subject else "No Subject"
            sender = html.escape(str(email.sender)) if email.sender else "Unknown Sender"
            received = email.datetime_received.strftime('%B %d, %Y at %I:%M %p') if email.datetime_received else "Unknown Date"
            
            # Clean body content
            body = email.body if email.body else "No body content"
            body = html.escape(body)
            
            # Highlight search term in body
            body = re.sub(
                f'({re.escape(search_term)})', 
                r'<span class="highlight">\1</span>', 
                body, 
                flags=re.IGNORECASE
            )
            
            html_content += f"""
        <div class="email">
            <div class="email-header">
                <div class="email-subject">Email #{i}: {subject}</div>
                <div class="email-meta">
                    <strong>From:</strong> {sender} | 
                    <strong>Received:</strong> {received}
                </div>
            </div>
            <div class="email-body">{body}</div>
        </div>
"""
    
    html_content += """
    </div>
</body>
</html>
"""
    
    return html_content

def save_to_text_file(emails, search_term, filename="outlook_search_results.txt"):
    """
    Save email results to a text file.
    
    Args:
        emails: List of email objects
        search_term: Search term used
        filename: Output filename
    """
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"OUTLOOK EMAIL SEARCH RESULTS\n")
            f.write(f"=" * 50 + "\n\n")
            f.write(f"Search Term: {search_term}\n")
            f.write(f"Date Range: {get_start_of_month().strftime('%B %d, %Y')} to {date.today().strftime('%B %d, %Y')}\n")
            f.write(f"Total Emails Found: {len(emails)}\n")
            f.write(f"Report Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}\n\n")
            
            if not emails:
                f.write("No emails found matching the search criteria.\n")
            else:
                for i, email in enumerate(emails, 1):
                    f.write(f"EMAIL #{i}\n")
                    f.write(f"-" * 30 + "\n")
                    f.write(f"Subject: {email.subject or 'No Subject'}\n")
                    f.write(f"From: {email.sender or 'Unknown Sender'}\n")
                    f.write(f"Received: {email.datetime_received.strftime('%B %d, %Y at %I:%M %p') if email.datetime_received else 'Unknown Date'}\n")
                    f.write(f"Body:\n{email.body or 'No body content'}\n")
                    f.write("\n" + "=" * 50 + "\n\n")
        
        print(f"Results saved to {filename}")
        return True
    
    except Exception as e:
        print(f"Error saving to file: {e}")
        return False

def main():
    """Main function to run the email search."""
    print("Outlook Email Search Tool")
    print("=" * 40)
    
    # Get user credentials
    email = input("Enter your email address: ").strip()
    password = input("Enter your password: ").strip()
    
    # Optional: Exchange server URL (leave empty for autodiscover)
    server = input("Enter Exchange server URL (or press Enter for autodiscover): ").strip()
    if not server:
        server = None
    
    # Connect to Outlook
    account = connect_to_outlook(email, password, server)
    if not account:
        print("Failed to connect to Outlook. Please check your credentials.")
        return
    
    # Set search parameters
    start_date = get_start_of_month()
    end_date = date.today()
    search_term = "beat"
    
    print(f"\nSearching for emails containing '{search_term}' from {start_date} to {end_date}...")
    
    # Search for emails
    matching_emails = search_emails(account, start_date, end_date, search_term)
    
    if matching_emails:
        # Create HTML report
        html_content = create_html_report(matching_emails, search_term)
        
        # Save HTML to file
        html_filename = "outlook_search_results.html"
        try:
            with open(html_filename, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"HTML report saved to {html_filename}")
        except Exception as e:
            print(f"Error saving HTML file: {e}")
        
        # Save text version
        save_to_text_file(matching_emails, search_term)
        
        print(f"\nSearch completed successfully!")
        print(f"Found {len(matching_emails)} emails containing '{search_term}'")
        print(f"Results saved to:")
        print(f"  - {html_filename} (HTML format)")
        print(f"  - outlook_search_results.txt (Text format)")
        
    else:
        print(f"No emails found containing '{search_term}' in the specified date range.")

if __name__ == "__main__":
    main()