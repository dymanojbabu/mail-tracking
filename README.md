# Mail Tracking

A Python-based email tracking system that extracts and organizes email data from your Gmail account using the Gmail API. This tool helps you maintain a structured record of your email communications by extracting sender and recipient information into an organized Excel spreadsheet.

## Features

- Extracts emails from both Sent and Inbox folders
- Supports pagination for handling large email volumes
- Parses email addresses and names into structured data
- Filters out unwanted emails based on predefined patterns
- Exports data to Excel with separate sheets for sent and received emails
- Handles multiple recipients (To, Cc, Bcc)

## Prerequisites

- Python 3.x
- Google Cloud Project with Gmail API enabled
- OAuth 2.0 credentials (`credentials.json`)

## Required Python Packages

```
google-auth-oauthlib
google-api-python-client
pandas
openpyxl
```

## Setup

1. Set up a Google Cloud Project and enable the Gmail API
2. Download your OAuth 2.0 credentials and save them as `credentials.json` in the project root
3. Install the required packages:
   ```bash
   pip install google-auth-oauthlib google-api-python-client pandas openpyxl
   ```

## Usage

1. Run the script with pagination (recommended for large mailboxes):
   ```bash
   python email_crawl_to_xl_with_pagination.py
   ```
   
2. Or run the simple version without pagination:
   ```bash
   python email_crawl_to_xl.py
   ```

The script will:
- Authenticate with your Gmail account
- Extract email data from your Sent and Inbox folders
- Save the data to `emails.xlsx` with two sheets: "Sent Emails" and "Received Emails"

## Output Format

The generated Excel file (`emails.xlsx`) contains two sheets:
- **Sent Emails**: Emails you've sent
- **Received Emails**: Emails you've received

Each sheet contains the following columns:
- First Name
- Last Name
- Email

## Customization

You can modify the `IGNORE_PATTERNS` list in the script to exclude specific email patterns from being extracted:

```python
IGNORE_PATTERNS = ["noreply", "no-reply", "internal", ...]
```

## Security Note

- Keep your `credentials.json` file secure and never commit it to version control
- The script requires read-only access to your Gmail account
- All data is processed locally on your machine

## Contributing

Feel free to submit issues and enhancement requests! 