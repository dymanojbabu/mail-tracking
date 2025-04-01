import re
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

def extract_emails(label):
    """Extract recipient emails from specified label (Sent or Inbox), ignoring patterns."""
    service = get_service()
    results = service.users().messages().list(userId='me', labelIds=[label], maxResults=100).execute()
    messages = results.get('messages', [])

    extracted_emails = set()
    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id'], format='metadata').execute()
        headers = msg_data.get('payload', {}).get('headers', [])

        for header in headers:
            if header['name'] in ['To', 'From', 'Cc', 'Bcc']:
                emails = re.findall(r'[\w\.-]+@[\w\.-]+', header['value'])
                # Filter out emails containing any ignored patterns
                filtered_emails = [email for email in emails if not any(pattern.lower() in email.lower() for pattern in IGNORE_PATTERNS)]
                extracted_emails.update(filtered_emails)

    return extracted_emails

def main():
    """Extract and print both sent and received emails."""
    print("\nExtracting Sent Emails...")
    sent_emails = extract_emails('SENT')
    print("\nExtracted Sent Emails (Filtered):")
    for email in sent_emails:
        print(email)

    print("\nExtracting Received Emails...")
    received_emails = extract_emails('INBOX')
    print("\nExtracted Received Emails (Filtered):")
    for email in received_emails:
        print(email)

main()
