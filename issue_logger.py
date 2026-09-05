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
        Initialize Google Sheets client menggunakan existing Railway variables.
        Gunakan GOOGLE_SHEETS_CREDENTIALS dan GOOGLE_SHEETS_SPREADSHEET_ID.
        
        Return:
            bool: True jika berhasil, False jika error
        """
        try:
            # Get credentials dari env var (sudah terintegrasi di Railway)
            creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
            spreadsheet_id = os.getenv('GOOGLE_SHEETS_SPREADSHEET_ID')
            
            logger.info(f"DEBUG: creds_json exists: {bool(creds_json)}")
            logger.info(f"DEBUG: spreadsheet_id exists: {bool(spreadsheet_id)}")
            logger.info(f"DEBUG: spreadsheet_id value: {spreadsheet_id[:50] if spreadsheet_id else 'None'}...")
            
            if not creds_json:
                logger.error("❌ GOOGLE_SHEETS_CREDENTIALS tidak ditemukan di environment")
                return False
            
            if not spreadsheet_id:
                logger.error("❌ GOOGLE_SHEETS_SPREADSHEET_ID tidak ditemukan di environment")
                return False
            
            # Parse JSON credentials
            import base64
            try:
                # Try decoding as base64 first
                logger.info("Trying to decode credentials as base64...")
                creds_dict = json.loads(base64.b64decode(creds_json).decode('utf-8'))
                logger.info("✓ Credentials decoded from base64")
            except Exception as e:
                logger.info(f"Base64 decode failed ({e}), trying raw JSON...")
                try:
                    creds_dict = json.loads(creds_json)
                    logger.info("✓ Credentials parsed as raw JSON")
                except Exception as json_error:
                    logger.error(f"❌ Failed to parse credentials: {json_error}")
                    return False
            
            logger.info(f"Credentials type: {creds_dict.get('type', 'unknown')}")
            logger.info(f"Project ID: {creds_dict.get('project_id', 'unknown')}")
            
            # Create credentials
            logger.info("Creating service account credentials...")
            credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
            logger.info("✓ Service account credentials created")
            
            # Initialize gspread client
            logger.info("Initializing gspread client...")
            self.client = gspread.authorize(credentials)
            logger.info("✓ gspread client authorized")
            
            # Open spreadsheet by ID (more reliable than by name)
            logger.info(f"Opening spreadsheet by ID: {spreadsheet_id}")
            self.spreadsheet = self.client.open_by_key(spreadsheet_id)
            logger.info(f"✓ Spreadsheet opened: {self.spreadsheet.title}")
            
            # Try to open IssueLogs worksheet, if not exist create it
            try:
                logger.info("Looking for IssueLogs worksheet...")
                self.worksheet = self.spreadsheet.worksheet('IssueLogs')
                logger.info(f"✓ IssueLogs worksheet opened ({self.worksheet.row_count} rows)")
                
                # Check if worksheet has correct structure (7 columns with Username and Ticket Number)
                existing_headers = self.worksheet.row_values(1)
                logger.info(f"Existing headers: {existing_headers}")
                
                expected_headers = ['Tanggal Issue', 'Chat ID', 'Username', 'Ticket Number', 'Source', 'Detail Issue', 'Action Resolved']
                
                if len(existing_headers) < 7 or existing_headers[2:4] != expected_headers[2:4]:
                    logger.warning("⚠️ IssueLogs worksheet has old structure, needs migration")
                    logger.info("Attempting to update headers to include Username and Ticket Number...")
                    
                    try:
                        # Update first row with new headers
                        self.worksheet.update([expected_headers], range_name='A1:G1')
                        logger.info("✓ Headers updated to new structure (7 columns)")
                    except Exception as e:
                        logger.error(f"❌ Failed to update headers: {e}")
                        logger.info("Please manually update worksheet headers in Google Sheets:")
                        logger.info("Expected: " + str(expected_headers))
                else:
                    logger.info("✓ IssueLogs worksheet has correct structure")
                    
            except gspread.exceptions.WorksheetNotFound:
                logger.info("IssueLogs worksheet not found, creating...")
                import time
                
                # Create new worksheet with proper parameters
                self.worksheet = self.spreadsheet.add_worksheet('IssueLogs', rows=1000, cols=7)
                logger.info(f"✓ Worksheet created: {self.worksheet.title}")
                
                # Small delay to ensure worksheet is ready
                time.sleep(1)
                
                # Add header row (now with Username and Ticket Number)
                headers = ['Tanggal Issue', 'Chat ID', 'Username', 'Ticket Number', 'Source', 'Detail Issue', 'Action Resolved']
                
                try:
                    # Use update_cells with direct cell references (most reliable)
                    cells_to_update = []
                    for col_idx, header_value in enumerate(headers, start=1):
                        cell = gspread.Cell(row=1, col=col_idx, value=header_value)
                        cells_to_update.append(cell)
                    
                    self.worksheet.update_cells(cells_to_update)
                    logger.info(f"✓ Headers added via update_cells")
                except Exception as update_err:
                    logger.warning(f"update_cells failed: {update_err}, trying append_row...")
                    try:
                        # Fallback: append row
                        self.worksheet.append_row(headers)
                        logger.info(f"✓ Headers added via append_row")
                    except Exception as append_err:
                        logger.error(f"append_row failed too: {append_err}")
                        # Continue anyway, maybe headers weren't critical
                
                # Re-fetch worksheet to ensure proper initialization
                self.worksheet = self.spreadsheet.worksheet('IssueLogs')
                logger.info("✓ Worksheet re-fetched and ready")
            
            logger.info("✓✓✓ Google Sheets initialization SUCCESS ✓✓✓")
            return True
            
        except Exception as e:
            logger.error(f"❌ Error initializing Google Sheets: {e}", exc_info=True)
            return False

    def get_current_timestamp(self) -> str:
        """
        Get current timestamp dalam format YYYY-MM-DD HH:mm:ss WIB.
        
        Return:
            str: Formatted timestamp
        """
        now = datetime.now(self.timezone)
        return now.strftime('%Y-%m-%d %H:%M:%S WIB')

    def append_issue(self, chat_id: str, username: str, ticket_number: str, source: str, detail_issue: str, action_resolved: str) -> Tuple[bool, str]:
        """
        Append row baru ke Google Sheets dengan issue data.

        Args:
            chat_id: Chat ID user
            username: Username Telegram user
            ticket_number: Ticket number (optional)
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
            
            # Check worksheet structure to determine column count
            try:
                headers = self.worksheet.row_values(1)
                col_count = len([h for h in headers if h])  # Count non-empty headers
                logger.info(f"DEBUG: Worksheet headers count: {col_count}, headers: {headers[:7]}")
            except Exception as e:
                logger.warning(f"Could not read headers: {e}, assuming new structure")
                col_count = 7  # Default to new structure
            
            # Prepare row data based on structure
            if col_count >= 7:
                # New structure: [timestamp, chat_id, username, ticket_number, source, detail, action]
                row_data = [timestamp, str(chat_id), username or "N/A", ticket_number or "-", source, detail_issue, action_resolved]
                logger.info(f"Using new 7-column structure")
            else:
                # Old structure fallback: [timestamp, chat_id, source, detail, action]
                # This is for backward compatibility
                logger.warning("⚠️ Old worksheet structure detected (5 cols), appending without username/ticket")
                row_data = [timestamp, str(chat_id), source, detail_issue, action_resolved]
            
            # Append to worksheet using append_row
            self.worksheet.append_row(row_data)
            
            logger.info(f"✓ Issue appended for {username} (chat_id={chat_id}, ticket={ticket_number}): {source}")
            return True, timestamp
            
        except gspread.exceptions.APIError as api_err:
            error_msg = f"Google Sheets API Error: {str(api_err)}"
            logger.error(error_msg)
            return False, error_msg
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
        Ticket Number: TIK-123456 (opsional)
        Kendala: Detail issue di sini
        Action: Action resolved di sini

        Args:
            text: Raw text input dari user

        Return:
            (success: bool, data: dict dengan keys: source, ticket_number, detail_issue, action_resolved)
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
                elif key in ['ticket', 'ticket number', 'nomor ticket', 'no ticket']:
                    parsed['ticket_number'] = value
                elif key in ['kendala', 'detail', 'issue']:
                    parsed['detail_issue'] = value
                elif key in ['action', 'aksi']:
                    parsed['action_resolved'] = value
            
            # Validate required fields
            required_fields = ['source', 'detail_issue', 'action_resolved']
            if not all(k in parsed for k in required_fields):
                missing = [k for k in required_fields if k not in parsed]
                return False, {'error': f"Kolom wajib diisi: {', '.join(missing)}"}
            
            # Set ticket_number to empty string if not provided
            if 'ticket_number' not in parsed:
                parsed['ticket_number'] = ""
            
            return True, parsed
            
        except Exception as e:
            logger.error(f"Error parsing noteissue input: {e}")
            return False, {'error': str(e)}
