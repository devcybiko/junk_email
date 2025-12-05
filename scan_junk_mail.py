#!/usr/bin/env python3
"""
Scan junk mail folder and extract unique email addresses
Works with Amazon WorkMail (Exchange-compatible server)
"""

import json
import os
import time
from datetime import datetime
from dotenv import load_dotenv
from exchangelib import Credentials, Account, DELEGATE, Configuration
from exchangelib.folders import JunkEmail
from exchangelib.errors import ErrorServerBusy
import re
from collections import defaultdict


def extract_email_addresses(text):
    """Extract all email addresses from text using regex"""
    if not text:
        return []
    
    # Pattern to match email addresses
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.findall(email_pattern, str(text))


def scan_junk_mail(email_address, password, email_count=defaultdict(int), server='outlook.office365.com', batch_size=100):
    print(f"Connecting to {server} as {email_address}...")
    
    # Set up credentials with retry policy
    credentials = Credentials(username=email_address, password=password)
    
    # Configure the account with a more lenient retry policy
    config = Configuration(
        server=server, 
        credentials=credentials,
        retry_policy=None  # Disable automatic retries, we'll handle them ourselves
    )
    account = Account(
        primary_smtp_address=email_address,
        config=config,
        autodiscover=False,
        access_type=DELEGATE
    )
    
    print("Connected! Scanning junk mail folder (fetching headers only for speed)...")
    print(f"Using batch size of {batch_size} messages with delays to avoid throttling")
    
    # Access junk email folder
    junk_folder = account.junk
    
    new_emails = set()
    processed = 0
    retries = 0
    max_retries = 5
    
    # Get total count first
    try:
        total = junk_folder.total_count
        print(f"Found {total} messages in junk folder")
    except:
        total = "unknown"
    
    # Process in smaller batches with delays
    while True:
        try:
            # Only fetch sender field - this is MUCH faster than fetching full items
            # .only() tells exchangelib to only retrieve specific fields
            items = list(junk_folder.all()
                        .only('sender')
                        .order_by('-datetime_received')[processed:processed+batch_size])
            
            if not items:
                print("No more messages to process")
                break
            
            for item in items:
                try:
                    # Extract from sender (only field we fetched)
                    if hasattr(item, 'sender') and item.sender:
                        sender_email = item.sender.email_address
                        if sender_email:
                            email_count[sender_email.lower()] += 1
                            if email_count[sender_email.lower()] == 1:
                                new_emails.add(sender_email.lower())
                    
                    processed += 1
                    
                except Exception as e:
                    print(f"Error processing individual message: {e}")
                    continue
            
            print(f"Processed {processed} messages...")
            
            # Save progress after each batch
            if processed % (batch_size * 2) == 0:
                save_progress(email_count, processed)
            
            # Longer delay between batches to avoid throttling
            print("Waiting 1 second before next batch...")
            time.sleep(1)
            retries = 0  # Reset retries on success
                
        except ErrorServerBusy as e:
            retries += 1
            if retries > max_retries:
                print(f"Server too busy after {max_retries} retries. Saving progress and exiting.")
                save_progress(email_count, processed)
                break
            
            wait_time = 10 * retries  # Exponential backoff
            print(f"Server busy (attempt {retries}/{max_retries}), waiting {wait_time} seconds... (processed {processed} so far)")
            time.sleep(wait_time)
            continue
            
        except Exception as e:
            print(f"Unexpected error: {e}")
            save_progress(email_count, processed)
            break
    
    print(f"Total messages processed: {processed}")
    return email_count, new_emails


def save_progress(email_count, processed):
    """Save progress to file"""
    filename = "junk_email_addresses_progress.json"
    with open(filename, 'w') as f:
        json.dump({
            'email_count': dict(email_count),
            'processed': processed,
            'timestamp': time.time()
        }, f, indent=4)
    print(f"Progress saved to {filename}")


def main():
    # Load environment variables from ~/.secrets
    secrets_path = os.path.expanduser('~/.secrets')
    load_dotenv(secrets_path)
    
    # Configuration - Update these values
    # EMAIL = input("Enter your email address: ")
    EMAIL = 'gregory@greg-smith.com'
    PASSWORD = os.getenv('PASSWORD')
    if not PASSWORD:
        PASSWORD = input("Enter your password: ")
    # For Amazon WorkMail, the server format is usually:
    # ews.mail.<region>.awsapps.com
    # Example: ews.mail.us-east-1.awsapps.com
    # SERVER = input("Enter Exchange/WorkMail server (e.g., ews.mail.us-east-1.awsapps.com): ")
    SERVER = 'ews.mail.us-east-1.awsapps.com'
    
    # Load previous progress if it exists
    progress_file = 'junk_email_addresses_progress.json'
    if os.path.exists(progress_file):
        print(f"Found previous progress file, loading...")
        with open(progress_file) as f:
            progress = json.load(f)
            email_count = defaultdict(int, progress.get('email_count', {}))
            print(f"Resuming from {progress.get('processed', 0)} messages")
    else:
        with open("junk_email_addresses.json") as f:
            email_count = defaultdict(int, json.load(f))
            print(f"Read {len(email_count)} messages")
    
    # You can adjust batch_size: smaller = slower but less likely to throttle
    email_count, new_emails = scan_junk_mail(EMAIL, PASSWORD, email_count, SERVER, batch_size=100)
    
    print(f"\n{'='*60}")
    print(f"Found {len(email_count)} unique email addresses in junk mail")
    print(f"{'='*60}\n")
    
    # Sort by frequency (most common first)
    sorted_emails = sorted(email_count.items(), key=lambda x: x[1], reverse=True)
    
    print("Email Address".ljust(50) + "Count")
    print("-" * 60)

    for email, count in sorted_emails:
        print(f"{email.ljust(50)} {count}")
    
    if new_emails:
        print(f"\n{len(new_emails)} new email(s) found this run:")
        for email in new_emails:
            print(f"  - {email}")
        
        # Save new emails to timestamped file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_emails_filename = f"new_emails_{timestamp}.json"
        with open(new_emails_filename, 'w') as f:
            json.dump(sorted(list(new_emails)), f, indent=4)
        print(f"\nNew emails saved to {new_emails_filename}")

    filename = "junk_email_addresses.json"
    with open(filename, 'w') as f:
        json.dump(dict(email_count), f, indent=4)
    print(f"\nResults saved to {filename}")
    
    # Clean up progress file on successful completion
    if os.path.exists(progress_file):
        os.remove(progress_file)
        print(f"Progress file removed")
    
if __name__ == "__main__":
    main()
