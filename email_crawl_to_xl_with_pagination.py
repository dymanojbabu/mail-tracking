import re
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Define patterns to ignore (any email containing these words will be excluded)
IGNORE_PATTERNS = ["noreply", "no-reply", "internal", "Avinash", "thota", "avi"]

def get_service():
    """Authenticate and return a Gmail API service."""
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    return build('gmail', 'v1', credentials=creds)

def parse_email_address(email_string):
    """Extract full name and email from a string like '"John Doe" <john.doe@example.com>'."""
    match = re.match(r'(?:"?([^"]+)"?\s)?<?([\w\.-]+@[\w\.-]+)>?', email_string)
    if match:
        full_name, email = match.groups()
        if full_name:
            name_parts = full_name.split()
            first_name = name_parts[0] if name_parts else ""
            last_name = name_parts[1] if len(name_parts) > 1 else ""
        else:
            first_name, last_name = "", ""  # No name found
        return first_name, last_name, email
    return "", "", email_string  # If no match, assume it's just an email

def extract_emails(label):
    """Extract first name, last name, and email from all messages in Sent or Inbox, with pagination."""
    service = get_service()
    extracted_data = []
    next_page_token = None

    while True:
        results = service.users().messages().list(
            userId='me', labelIds=[label], maxResults=500, pageToken=next_page_token
        ).execute()

        messages = results.get('messages', [])
        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id'], format='metadata').execute()
            headers = msg_data.get('payload', {}).get('headers', [])

            for header in headers:
                if header['name'] in ['To', 'From', 'Cc', 'Bcc']:
                    emails = re.findall(r'[^,]+', header['value'])  # Split multiple recipients
                    for email_string in emails:
                        first_name, last_name, email = parse_email_address(email_string.strip())
                        if not any(pattern.lower() in email.lower() for pattern in IGNORE_PATTERNS):
                            extracted_data.append([first_name, last_name, email])

        next_page_token = results.get('nextPageToken')  # Move to the next page
        if not next_page_token:
            break  # Stop when there are no more pages

    return extracted_data

def main():
    """Extract all sent and received emails with names, then save to an Excel file."""
    print("\nExtracting Sent Emails (Paginated)...")
    sent_emails = extract_emails('SENT')

    print("\nExtracting Received Emails (Paginated)...")
    received_emails = extract_emails('INBOX')

    # Convert to DataFrames
    sent_df = pd.DataFrame(sent_emails, columns=["First Name", "Last Name", "Email"]) if sent_emails else pd.DataFrame(columns=["First Name", "Last Name", "Email"])
    received_df = pd.DataFrame(received_emails, columns=["First Name", "Last Name", "Email"]) if received_emails else pd.DataFrame(columns=["First Name", "Last Name", "Email"])

    # Save to a single Excel file with two sheets
    with pd.ExcelWriter("emails.xlsx") as writer:
        sent_df.to_excel(writer, sheet_name="Sent Emails", index=False)
        received_df.to_excel(writer, sheet_name="Received Emails", index=False)

    print("\n✅ Emails saved to 'emails.xlsx'")

main()
