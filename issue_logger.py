"""
Issue Logger - Integrasi dengan Google Sheets untuk pencatatan incident.
Handles append & fetch operations untuk worksheet IssueLogs.
"""

import logging
import gspread
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from datetime import datetime
import pytz
from typing import List, Tuple, Optional
import os
import json

logger = logging.getLogger(__name__)

# Google Sheets scope
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


class IssueLogger:
    """
    Handles Google Sheets operations untuk incident logging.
    """

    def __init__(self):
        """
        Initialize IssueLogger dengan Google Sheets credentials.
        Credentials harus di-provide via GOOGLE_SHEETS_CREDENTIALS_JSON env var.
        """
        self.client = None
        self.spreadsheet = None
        self.worksheet = None
        self.timezone = pytz.timezone('Asia/Jakarta')
        self._initialize_sheets()

    def _initialize_sheets(self) -> bool:
        """
        Initialize Google Sheets client dan buka spreadsheet itops-ticket-log.
        
        Return:
            bool: True jika berhasil, False jika error
        """
        try:
            # Get credentials dari env var (base64 encoded JSON)
            creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
            if not creds_json:
                logger.error("GOOGLE_SHEETS_CREDENTIALS_JSON tidak ditemukan di .env")
                return False
            
            # Parse JSON credentials
            import base64
            try:
                # Try decoding as base64 first
                creds_dict = json.loads(base64.b64decode(creds_json).decode('utf-8'))
            except:
                # If not base64, treat as raw JSON
                creds_dict = json.loads(creds_json)
            
            # Create credentials
            credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
            
            # Initialize gspread client
            self.client = gspread.authorize(credentials)
            
            # Open spreadsheet by name
            self.spreadsheet = self.client.open('itops-ticket-log')
            logger.info("✓ Google Sheets client initialized")
            
            # Try to open IssueLogs worksheet, if not exist create it
            try:
                self.worksheet = self.spreadsheet.worksheet('IssueLogs')
                logger.info("✓ IssueLogs worksheet opened")
            except gspread.exceptions.WorksheetNotFound:
                logger.info("IssueLogs worksheet not found, creating...")
                self.worksheet = self.spreadsheet.add_worksheet('IssueLogs', rows=1000, cols=5)
                
                # Add header row
                headers = ['Tanggal Issue', 'Chat ID', 'Source', 'Detail Issue', 'Action Resolved']
                self.worksheet.append_row(headers)
                logger.info("✓ IssueLogs worksheet created with headers")
            
            return True
            
        except Exception as e:
            logger.error(f"Error initializing Google Sheets: {e}")
            return False

    def get_current_timestamp(self) -> str:
        """
        Get current timestamp dalam format YYYY-MM-DD HH:mm:ss WIB.
        
        Return:
            str: Formatted timestamp
        """
        now = datetime.now(self.timezone)
        return now.strftime('%Y-%m-%d %H:%M:%S WIB')

    def append_issue(self, chat_id: str, source: str, detail_issue: str, action_resolved: str) -> Tuple[bool, str]:
        """
        Append row baru ke Google Sheets dengan issue data.

        Args:
            chat_id: Chat ID user
            source: Source of issue (Web App, Mobile, POS, etc)
            detail_issue: Detail/description issue
            action_resolved: Action yang diambil untuk resolve

        Return:
            (success: bool, message: str)
        """
        try:
            if not self.worksheet:
                return False, "❌ Google Sheets tidak terhubung"
            
            # Get current timestamp
            timestamp = self.get_current_timestamp()
            
            # Prepare row data
            row_data = [timestamp, str(chat_id), source, detail_issue, action_resolved]
            
            # Append to worksheet
            self.worksheet.append_row(row_data)
            
            logger.info(f"✓ Issue appended for chat_id={chat_id}: {source}")
            return True, timestamp
            
        except Exception as e:
            error_msg = f"Error appending issue: {str(e)}"
            logger.error(error_msg)
            return False, error_msg

    def fetch_all_issues(self) -> Tuple[bool, List[List], Optional[str]]:
        """
        Fetch semua rows dari IssueLogs worksheet.

        Return:
            (success: bool, rows: List[List], error_msg: Optional[str])
        """
        try:
            if not self.worksheet:
                return False, [], "Google Sheets tidak terhubung"
            
            # Get all values
            all_rows = self.worksheet.get_all_values()
            
            if len(all_rows) <= 1:
                # Only header or empty
                return True, [], None
            
            # Skip header row (index 0)
            data_rows = all_rows[1:]
            
            logger.info(f"✓ Fetched {len(data_rows)} issue rows from Google Sheets")
            return True, data_rows, None
            
        except Exception as e:
            error_msg = f"Error fetching issues: {str(e)}"
            logger.error(error_msg)
            return False, [], error_msg

    def parse_noteissue_input(self, text: str) -> Tuple[bool, dict]:
        """
        Parse input dari /noteissue command.
        Format yang diharapkan:
        Source: Web App
        Kendala: Detail issue di sini
        Action: Action resolved di sini

        Args:
            text: Raw text input dari user

        Return:
            (success: bool, data: dict dengan keys: source, detail_issue, action_resolved)
        """
        try:
            # Split by newlines
            lines = text.strip().split('\n')
            
            parsed = {}
            for line in lines:
                if ':' not in line:
                    continue
                    
                key, value = line.split(':', 1)
                key = key.strip().lower()
                value = value.strip()
                
                if key == 'source':
                    parsed['source'] = value
                elif key in ['kendala', 'detail', 'issue']:
                    parsed['detail_issue'] = value
                elif key in ['action', 'aksi']:
                    parsed['action_resolved'] = value
            
            # Validate required fields
            if not all(k in parsed for k in ['source', 'detail_issue', 'action_resolved']):
                missing = [k for k in ['source', 'detail_issue', 'action_resolved'] if k not in parsed]
                return False, {'error': f"Kolom wajib diisi: {', '.join(missing)}"}
            
            return True, parsed
            
        except Exception as e:
            logger.error(f"Error parsing noteissue input: {e}")
            return False, {'error': str(e)}
