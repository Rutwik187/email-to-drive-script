"""
Gmail Attachment Downloader Script
Downloads all .pdf and .docx attachments from Gmail within a specified date range,
and uploads them to your Google Drive (flat—no subfolders).
"""

import os
import base64
import json
import re
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional

# Google API imports
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import io
from googleapiclient.http import MediaIoBaseUpload

class GmailAttachmentDownloader:
    """Download PDF/DOCX attachments from Gmail and upload to Google Drive."""

    SCOPES = [
        'https://www.googleapis.com/auth/gmail.readonly', 
        "https://www.googleapis.com/auth/drive.file"
    ]

    def __init__(self, credentials_file: str = 'credentials.json', token_file: str = 'token.json'):
        """
        Initialize the Gmail & Drive API clients.
        Args:
            credentials_file: Path to Gmail API credentials JSON file
            token_file: Path to store authentication tokens
        """
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.gmail_service = None
        self.drive_service = None
        self._authenticate()

    def _authenticate(self):
        """Authenticate and build Gmail & Drive API services."""
        creds = None
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, self.SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(self.credentials_file):
                    raise FileNotFoundError(
                        f"Credentials file '{self.credentials_file}' not found. "
                        "Please download it from Google Cloud Console."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
        self.gmail_service = build('gmail', 'v1', credentials=creds)
        self.drive_service = build('drive', 'v3', credentials=creds)

    def format_date_for_query(self, date_obj: datetime) -> str:
        return date_obj.strftime('%Y/%m/%d')

    def search_messages_with_attachments(
        self, 
        start_date: datetime, 
        end_date: datetime,
        additional_query: str = ""
    ) -> List[Dict[str, Any]]:
        """Search for messages with attachments within date range."""
        try:
            query_parts = [
                "has:attachment",
                f"after:{self.format_date_for_query(start_date)}",
                f"before:{self.format_date_for_query(end_date)}"
            ]
            if additional_query:
                query_parts.append(additional_query)
            query = " ".join(query_parts)
            print(f"Search query: {query}")
            result = self.gmail_service.users().messages().list(
                userId='me', q=query, maxResults=500
            ).execute()
            messages = []
            if 'messages' in result:
                messages.extend(result['messages'])
            while 'nextPageToken' in result:
                page_token = result['nextPageToken']
                result = self.gmail_service.users().messages().list(
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
        try:
            message = self.gmail_service.users().messages().get(
                userId='me', id=message_id, format='full'
            ).execute()
            return message
        except HttpError as error:
            print(f"Error getting message {message_id}: {error}")
            return None

    def extract_message_info(self, message: Dict[str, Any]) -> Dict[str, Any]:
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

    def upload_file_to_drive(self, file_data: bytes, filename: str, mimetype: str) -> Optional[str]:
        """Upload a file to Google Drive root, return file ID if successful."""
        try:
            file_metadata = {'name': filename, 'parents': ["1vrqio1vr_fhsdT7rSMYLcyk_Gv2tplvo"]}
            media = MediaIoBaseUpload(io.BytesIO(file_data), mimetype=mimetype, resumable=True)
            file = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            print(f"Uploaded to Drive: {filename} (file id: {file['id']})")
            return file['id']
        except HttpError as error:
            print(f"Drive upload failed for {filename}: {error}")
            return None

    def download_attachment_and_upload_to_drive(
        self, 
        message_id: str, 
        attachment_id: str,
        filename: str, 
        mime_type: str
    ) -> bool:
        """Download attachment from Gmail and upload to Google Drive."""
        try:
            attachment = self.gmail_service.users().messages().attachments().get(
                userId='me', messageId=message_id, id=attachment_id
            ).execute()
            data = attachment['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))

            file_id = self.upload_file_to_drive(file_data, filename, mime_type)
            return file_id is not None
        except HttpError as error:
            print(f"Error downloading/uploading {filename}: {error}")
            return False
        except Exception as error:
            print(f"Unexpected error downloading/uploading {filename}: {error}")
            return False

    def process_message_parts(
        self, 
        parts: List[Dict[str, Any]], 
        message_id: str, 
        message_info: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """Find and upload .pdf/.docx attachments to Drive root."""
        attachments_info = []
        for part in parts:
            filename = part.get('filename', '')
            mime_type = part.get('mimeType', '')

            is_pdf = (filename and filename.lower().endswith('.pdf')) or mime_type == 'application/pdf'
            is_docx = (filename and filename.lower().endswith('.docx')) or mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

            if is_pdf or is_docx:
                body = part.get('body', {})
                attachment_id = body.get('attachmentId')
                if attachment_id:
                    success = self.download_attachment_and_upload_to_drive(
                        message_id,
                        attachment_id,
                        filename,
                        mime_type or ("application/pdf" if is_pdf else "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    )
                    attachment_info = {
                        'filename': filename,
                        'attachment_id': attachment_id,
                        'size': body.get('size', 0),
                        'mime_type': mime_type,
                        'uploaded': success,
                        'destination': "Google Drive" if success else None
                    }
                    attachments_info.append(attachment_info)

            if 'parts' in part:
                attachments_info.extend(self.process_message_parts(
                    part['parts'], message_id, message_info
                ))
        return attachments_info

    def download_attachments_in_date_range(
        self,
        start_date: str,
        end_date: str,
        additional_query: str = "",
        save_log: bool = True
    ) -> Dict[str, Any]:
        """
        Download .pdf/.docx from Gmail within specified date range and upload to Google Drive.
        """
        try:
            start_dt = datetime.strptime(start_date, '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')
            if start_dt > end_dt:
                raise ValueError("Start date must be before end date")

            print(f"Searching for attachments from {start_date} to {end_date}")
            messages = self.search_messages_with_attachments(
                start_dt, end_dt, additional_query
            )
            if not messages:
                print("No messages with attachments found in the specified date range.")
                return {
                    'total_messages': 0,
                    'total_attachments': 0,
                    'successful_uploads': 0,
                    'failed_uploads': 0,
                }

            download_log = []
            total_attachments = 0
            successful_uploads = 0
            failed_uploads = 0

            for i, message in enumerate(messages, 1):
                print(f"\nProcessing message {i}/{len(messages)}: {message['id']}")
                message_details = self.get_message_details(message['id'])
                if not message_details:
                    continue
                message_info = self.extract_message_info(message_details)
                print(f"Subject: {message_info['subject']}")
                print(f"From: {message_info['from']}")
                print(f"Date: {message_info['date']}")

                payload = message_details.get('payload', {})
                parts = payload.get('parts', [])
                if parts:
                    attachments = self.process_message_parts(
                        parts, message['id'], message_info
                    )
                    total_attachments += len(attachments)
                    successful_uploads += sum(1 for a in attachments if a['uploaded'])
                    failed_uploads += sum(1 for a in attachments if not a['uploaded'])
                    download_log.append({
                        'message_info': message_info,
                        'attachments': attachments
                    })

            if save_log and download_log:
                log_file = f'drive_upload_log_{start_date}_to_{end_date}.json'
                with open(log_file, 'w', encoding='utf-8') as f:
                    json.dump(download_log, f, indent=2, ensure_ascii=False)
                print(f"\nUpload log saved to: {log_file}")

            summary = {
                'total_messages': len(messages),
                'total_attachments': total_attachments,
                'successful_uploads': successful_uploads,
                'failed_uploads': failed_uploads,
                'location': "Google Drive",
            }

            print(f"\n=== SUMMARY ===")
            print(f"Messages processed: {summary['total_messages']}")
            print(f"Total attachments: {summary['total_attachments']}")
            print(f"Successfully uploaded: {summary['successful_uploads']}")
            print(f"Failed uploads: {summary['failed_uploads']}")
            print(f"Upload location: {summary['location']}")
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
        downloader = GmailAttachmentDownloader()
        end_date = datetime.now()
        start_date = end_date - timedelta(days=30)
        start_date_str = start_date.strftime('%Y-%m-%d')
        end_date_str = end_date.strftime('%Y-%m-%d')
        print("Gmail Attachment Downloader (uploads to Drive)")
        print("=" * 40)
        print(f"Default date range: {start_date_str} to {end_date_str}\n")

        custom_start = input(f"Enter start date (YYYY-MM-DD) or press Enter for {start_date_str}: ").strip()
        custom_end = input(f"Enter end date (YYYY-MM-DD) or press Enter for {end_date_str}: ").strip()
        if custom_start:
            start_date_str = custom_start
        if custom_end:
            end_date_str = custom_end

        additional_query = input("Enter additional Gmail search filters (optional): ").strip()
        print(f"\nStarting upload from {start_date_str} to {end_date_str}...\n")

        summary = downloader.download_attachments_in_date_range(
            start_date=start_date_str,
            end_date=end_date_str,
            additional_query=additional_query
        )
        print("\nUpload completed!")

    except KeyboardInterrupt:
        print("\nUpload cancelled by user.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
