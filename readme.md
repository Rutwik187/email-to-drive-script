# Gmail Attachment Downloader - Setup Guide

This guide will help you set up and use the Gmail Attachment Downloader script to download all attachments from your Gmail within a specified date range.

## Prerequisites

- Python 3.7 or higher
- A Google account with Gmail
- Basic familiarity with command line/terminal

## Step 1: Google Cloud Console Setup

### 1.1 Create a New Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click "Select a project" → "New Project"
3. Enter project name (e.g., "Gmail Attachment Downloader")
4. Click "Create"

### 1.2 Enable Gmail API
1. In the Google Cloud Console, go to "APIs & Services" → "Library"
2. Search for "Gmail API"
3. Click on "Gmail API" and click "Enable"

### 1.3 Create Credentials
1. Go to "APIs & Services" → "Credentials"
2. Click "Create Credentials" → "OAuth client ID"
3. If prompted, configure the OAuth consent screen:
   - Choose "External" (unless you have a Google Workspace account)
   - Fill in required fields:
     - App name: "Gmail Attachment Downloader"
     - User support email: Your email
     - Developer contact information: Your email
   - Click "Save and Continue"
   - Skip "Scopes" and "Test users" for now
4. Back to creating OAuth client ID:
   - Choose "Desktop application"
   - Name: "Gmail Attachment Downloader Client"
   - Click "Create"
5. Download the credentials JSON file
6. Rename it to `credentials.json` and place it in your project folder

## Step 2: Python Environment Setup

### 2.1 Create Virtual Environment (Recommended)
```bash
# Create virtual environment
python -m venv gmail_downloader_env

# Activate virtual environment
# On Windows:
gmail_downloader_env\Scripts\activate
# On macOS/Linux:
source gmail_downloader_env/bin/activate
```

### 2.2 Install Dependencies
```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

## Step 3: File Structure

Your project folder should look like this:
```
gmail_attachment_downloader/
├── gmail_attachment_downloader.py
├── requirements.txt
├── credentials.json (downloaded from Google Cloud Console)
├── token.json (will be created automatically)
└── gmail_attachments/ (will be created automatically)
```

## Step 4: Running the Script

### 4.1 Basic Usage
```bash
python gmail_attachment_downloader.py
```

The script will:
1. Prompt for authentication (first time only)
2. Ask for date range (defaults to last 30 days)
3. Ask for additional search filters (optional)
4. Ask for download folder (defaults to 'gmail_attachments')
5. Start downloading attachments

### 4.2 Programmatic Usage
You can also use the script programmatically:

```python
from gmail_attachment_downloader import GmailAttachmentDownloader

# Initialize downloader
downloader = GmailAttachmentDownloader()

# Download attachments from specific date range
summary = downloader.download_attachments_in_date_range(
    start_date='2024-01-01',
    end_date='2024-01-31',
    download_folder='my_downloads',
    additional_query='from:important@company.com'
)

print(summary)
```

## Step 5: Advanced Usage

### 5.1 Date Range Options
- Start and end dates should be in `YYYY-MM-DD` format
- Examples:
  - `2024-01-01` to `2024-01-31` (January 2024)
  - `2023-12-01` to `2024-02-29` (3 months)

### 5.2 Additional Search Filters
You can add Gmail search operators:
- `from:sender@example.com` - From specific sender
- `to:recipient@example.com` - To specific recipient
- `subject:"Important Document"` - Specific subject
- `filename:pdf` - Only PDF attachments
- `larger:10M` - Files larger than 10MB
- `label:important` - Emails with specific label

Examples:
- `from:hr@company.com filename:pdf` - PDF attachments from HR
- `subject:"Invoice" larger:1M` - Large invoice files
- `from:client@company.com after:2024/01/01 before:2024/02/01` - Client emails in January

### 5.3 Output Structure
```
gmail_attachments/
├── download_log.json
├── 2024-01-15_Important_Document_a1b2c3d4/
│   ├── document.pdf
│   └── spreadsheet.xlsx
├── 2024-01-16_Meeting_Notes_e5f6g7h8/
│   └── presentation.pptx
└── ...
```

Each email's attachments are stored in separate folders with:
- Date prefix
- Truncated subject
- Message ID suffix

## Step 6: Troubleshooting

### 6.1 Authentication Issues
- Make sure `credentials.json` is in the correct location
- Delete `token.json` and re-authenticate if needed
- Ensure your Google account has access to the project

### 6.2 API Quota Limits
- Gmail API has daily quota limits
- For large downloads, consider running in batches
- The script handles rate limiting automatically

### 6.3 File permission Errors
- Ensure write permissions in download folder
- Run as administrator if needed (Windows)

### 6.4 Common Errors
1. **"credentials.json not found"**: Download and rename the credentials file
2. **"Access denied"**: Check OAuth consent screen configuration
3. **"Quota exceeded"**: Wait 24 hours or request quota increase
4. **"Invalid date format"**: Use YYYY-MM-DD format

## Step 7: Security Considerations

### 7.1 Protect Your Credentials
- Keep `credentials.json` and `token.json` secure
- Don't commit these files to version control
- Use environment variables for production deployments

### 7.2 Minimal Permissions
- The script uses `gmail.readonly` scope only
- It cannot send, delete, or modify emails
- Only downloads attachment data

## Step 8: Customization Options

### 8.1 Modify Search Parameters
Edit the `search_messages_with_attachments` method to change:
- Default search query
- Maximum results per page
- Additional filters

### 8.2 Change File Organization
Edit the `process_message_parts` method to change:
- Folder naming convention
- File naming convention
- Duplicate handling

### 8.3 Add Progress Tracking
The script includes basic progress indicators. You can enhance:
- Add progress bars (using tqdm)
- Implement logging
- Add email notifications on completion

## Step 9: Batch Processing Large Archives

For downloading from large date ranges:

```python
from datetime import datetime, timedelta

downloader = GmailAttachmentDownloader()

# Process 1 month at a time
start = datetime(2020, 1, 1)
end = datetime(2024, 1, 1)

current = start
while current < end:
    month_end = current + timedelta(days=30)
    if month_end > end:
        month_end = end

    print(f"Processing: {current.strftime('%Y-%m-%d')} to {month_end.strftime('%Y-%m-%d')}")

    downloader.download_attachments_in_date_range(
        start_date=current.strftime('%Y-%m-%d'),
        end_date=month_end.strftime('%Y-%m-%d'),
        download_folder=f'gmail_archive_{current.year}_{current.month:02d}'
    )

    current = month_end
```

## Support and Contributing

For issues or improvements:
1. Check the troubleshooting section
2. Review Gmail API documentation
3. Consider rate limits and quota restrictions
4. Test with small date ranges first

Happy downloading! 📧📎
