
"""
Gmail Attachment Downloader Script
Downloads all attachments from Gmail within a specified date range.
"""

import os
import base64
import pickle
import json
import re
from datetime import datetime, timedelta
from pathlib import Path
import mimetypes
from typing import List, Dict, Any, Optional

# Google API imports
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

class GmailAttachmentDownloader:
    """Download attachments from Gmail within a specified date range."""

    # Gmail API scopes needed for reading messages and downloading attachments
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

    def __init__(self, credentials_file: str = 'credentials.json', token_file: str = 'token.json'):
        """
        Initialize the Gmail API client.

        Args:
            credentials_file: Path to Gmail API credentials JSON file
            token_file: Path to store authentication tokens
        """
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.service = None
        self._authenticate()

    def _authenticate(self):
        """Authenticate and build the Gmail API service."""
        creds = None

        # Load existing token
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, self.SCOPES)

        # If there are no valid credentials, get new ones
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(self.credentials_file):
                    raise FileNotFoundError(
                        f"Credentials file '{self.credentials_file}' not found. "
                        "Please download it from Google Cloud Console."
                    )

                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, self.SCOPES
                )
                creds = flow.run_local_server(port=0)

            # Save the credentials for next run
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())

        self.service = build('gmail', 'v1', credentials=creds)

    def format_date_for_query(self, date_obj: datetime) -> str:
        """Format datetime object for Gmail search query."""
        return date_obj.strftime('%Y/%m/%d')

    def search_messages_with_attachments(
        self, 
        start_date: datetime, 
        end_date: datetime,
        additional_query: str = ""
    ) -> List[Dict[str, Any]]:
        """
        Search for messages with attachments within date range.

        Args:
            start_date: Start date for search
            end_date: End date for search
            additional_query: Additional search parameters

        Returns:
            List of message metadata
        """
        try:
            # Build search query
            query_parts = [
                "has:attachment",
                f"after:{self.format_date_for_query(start_date)}",
                f"before:{self.format_date_for_query(end_date)}"
            ]

            if additional_query:
                query_parts.append(additional_query)

            query = " ".join(query_parts)
            print(f"Search query: {query}")

            # Search messages
            result = self.service.users().messages().list(
                userId='me', 
                q=query,
                maxResults=500  # Adjust as needed
            ).execute()

            messages = []
            if 'messages' in result:
                messages.extend(result['messages'])

            # Handle pagination
            while 'nextPageToken' in result:
                page_token = result['nextPageToken']
                result = self.service.users().messages().list(
                    userId='me',
                    q=query,
                    pageToken=page_token,
                    maxResults=500
                ).execute()
                if 'messages' in result:
                    messages.extend(result['messages'])

            print(f"Found {len(messages)} messages with attachments")
            return messages

        except HttpError as error:
            print(f"An error occurred during search: {error}")
            return []

    def get_message_details(self, message_id: str) -> Optional[Dict[str, Any]]:
        """Get full message details including attachments."""
        try:
            message = self.service.users().messages().get(
                userId='me', 
                id=message_id,
                format='full'
            ).execute()
            return message
        except HttpError as error:
            print(f"Error getting message {message_id}: {error}")
            return None

    def extract_message_info(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Extract useful information from message."""
        headers = message['payload'].get('headers', [])
        info = {
            'id': message['id'],
            'thread_id': message['threadId'],
            'subject': '',
            'from': '',
            'date': '',
            'internal_date': message.get('internalDate', '')
        }

        for header in headers:
            name = header['name'].lower()
            if name == 'subject':
                info['subject'] = header['value']
            elif name == 'from':
                info['from'] = header['value']
            elif name == 'date':
                info['date'] = header['value']

        return info

    def download_attachment(
        self, 
        message_id: str, 
        attachment_id: str, 
        filename: str, 
        download_folder: str
    ) -> bool:
        """Download a single attachment."""
        try:
            # Get attachment data
            attachment = self.service.users().messages().attachments().get(
                userId='me',
                messageId=message_id,
                id=attachment_id
            ).execute()

            # Decode attachment data
            data = attachment['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))

            # Ensure download folder exists
            Path(download_folder).mkdir(parents=True, exist_ok=True)

            # Create safe filename
            safe_filename = re.sub(r'[<>:"/\|?*]', '_', filename)
            file_path = os.path.join(download_folder, safe_filename)

            # Handle duplicate filenames
            counter = 1
            original_path = file_path
            while os.path.exists(file_path):
                name, ext = os.path.splitext(original_path)
                file_path = f"{name}_{counter}{ext}"
                counter += 1

            # Write file
            with open(file_path, 'wb') as f:
                f.write(file_data)

            print(f"Downloaded: {safe_filename} ({len(file_data)} bytes)")
            return True

        except HttpError as error:
            print(f"Error downloading attachment {filename}: {error}")
            return False
        except Exception as error:
            print(f"Unexpected error downloading {filename}: {error}")
            return False

    def process_message_parts(
        self, 
        parts: List[Dict[str, Any]], 
        message_id: str, 
        download_folder: str,
        message_info: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """Process message parts to find and download attachments."""
        attachments_info = []

        for part in parts:
            filename = part.get('filename', '')
            mime_type = part.get('mimeType', '')

            if (filename and filename.lower().endswith('.pdf')) or mime_type == 'application/pdf':
                # Get attachment info
                body = part.get('body', {})
                attachment_id = body.get('attachmentId')

                if attachment_id:
                    success = self.download_attachment(
                    message_id,
                    attachment_id,
                    filename,
                    download_folder  # Store all PDFs here
                    )


                    attachment_info = {
                        'filename': filename,
                        'attachment_id': attachment_id,
                        'size': body.get('size', 0),
                        'mime_type': part.get('mimeType', ''),
                        'downloaded': success,
                        'folder': download_folder if success else None
                    }
                    attachments_info.append(attachment_info)

            # Recursively process nested parts (for multipart messages)
            if 'parts' in part:
                nested_attachments = self.process_message_parts(
                    part['parts'], 
                    message_id, 
                    download_folder, 
                    message_info
                )
                attachments_info.extend(nested_attachments)

        return attachments_info

    def download_attachments_in_date_range(
        self,
        start_date: str,
        end_date: str,
        download_folder: str = "gmail_attachments",
        additional_query: str = "",
        save_log: bool = True
    ) -> Dict[str, Any]:
        """
        Download all attachments from Gmail within specified date range.

        Args:
            start_date: Start date in YYYY-MM-DD format
            end_date: End date in YYYY-MM-DD format  
            download_folder: Folder to save attachments
            additional_query: Additional Gmail search parameters
            save_log: Whether to save download log

        Returns:
            Summary of download operation
        """
        try:
            # Parse dates
            start_dt = datetime.strptime(start_date, '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')

            if start_dt > end_dt:
                raise ValueError("Start date must be before end date")

            print(f"Searching for attachments from {start_date} to {end_date}")
            print(f"Download folder: {download_folder}")

            # Search for messages
            messages = self.search_messages_with_attachments(
                start_dt, 
                end_dt, 
                additional_query
            )

            if not messages:
                print("No messages with attachments found in the specified date range.")
                return {
                    'total_messages': 0,
                    'total_attachments': 0,
                    'successful_downloads': 0,
                    'failed_downloads': 0,
                    'download_folder': download_folder
                }

            # Process each message
            download_log = []
            total_attachments = 0
            successful_downloads = 0
            failed_downloads = 0

            for i, message in enumerate(messages, 1):
                print(f"\nProcessing message {i}/{len(messages)}: {message['id']}")

                # Get message details
                message_details = self.get_message_details(message['id'])
                if not message_details:
                    continue

                # Extract message info
                message_info = self.extract_message_info(message_details)
                print(f"Subject: {message_info['subject']}")
                print(f"From: {message_info['from']}")
                print(f"Date: {message_info['date']}")

                # Process attachments
                payload = message_details.get('payload', {})
                parts = payload.get('parts', [])

                if parts:
                    attachments = self.process_message_parts(
                        parts, 
                        message['id'], 
                        download_folder,
                        message_info
                    )

                    total_attachments += len(attachments)
                    successful_downloads += sum(1 for a in attachments if a['downloaded'])
                    failed_downloads += sum(1 for a in attachments if not a['downloaded'])

                    # Add to log
                    log_entry = {
                        'message_info': message_info,
                        'attachments': attachments
                    }
                    download_log.append(log_entry)

            # Save download log
            if save_log and download_log:
                log_file = os.path.join(download_folder, 'download_log.json')
                Path(download_folder).mkdir(parents=True, exist_ok=True)

                with open(log_file, 'w', encoding='utf-8') as f:
                    json.dump(download_log, f, indent=2, ensure_ascii=False)
                print(f"\nDownload log saved to: {log_file}")

            # Print summary
            summary = {
                'total_messages': len(messages),
                'total_attachments': total_attachments,
                'successful_downloads': successful_downloads,
                'failed_downloads': failed_downloads,
                'download_folder': download_folder
            }

            print(f"\n=== DOWNLOAD SUMMARY ===")
            print(f"Messages processed: {summary['total_messages']}")
            print(f"Total attachments: {summary['total_attachments']}")
            print(f"Successfully downloaded: {summary['successful_downloads']}")
            print(f"Failed downloads: {summary['failed_downloads']}")
            print(f"Download folder: {summary['download_folder']}")

            return summary

        except ValueError as e:
            print(f"Date parsing error: {e}")
            return {}
        except Exception as e:
            print(f"Unexpected error: {e}")
            return {}


def main():
    """Example usage of Gmail Attachment Downloader."""
    try:
        # Initialize downloader
        downloader = GmailAttachmentDownloader()

        # Example: Download attachments from last 30 days
        end_date = datetime.now()
        start_date = end_date - timedelta(days=30)

        start_date_str = start_date.strftime('%Y-%m-%d')
        end_date_str = end_date.strftime('%Y-%m-%d')

        print("Gmail Attachment Downloader")
        print("=" * 40)
        print(f"Default date range: {start_date_str} to {end_date_str}")
        print()

        # You can customize these parameters
        custom_start = input(f"Enter start date (YYYY-MM-DD) or press Enter for {start_date_str}: ").strip()
        custom_end = input(f"Enter end date (YYYY-MM-DD) or press Enter for {end_date_str}: ").strip()

        if custom_start:
            start_date_str = custom_start
        if custom_end:
            end_date_str = custom_end

        # Optional: Add additional search filters
        additional_query = input("Enter additional search filters (optional, e.g., 'from:example@gmail.com'): ").strip()

        # Download folder
        download_folder = input("Enter download folder (or press Enter for 'gmail_attachments'): ").strip()
        if not download_folder:
            download_folder = "gmail_attachments"

        # Start download
        print(f"\nStarting download from {start_date_str} to {end_date_str}...")

        summary = downloader.download_attachments_in_date_range(
            start_date=start_date_str,
            end_date=end_date_str,
            download_folder=download_folder,
            additional_query=additional_query
        )

        print("\nDownload completed!")

    except KeyboardInterrupt:
        print("\nDownload cancelled by user.")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
