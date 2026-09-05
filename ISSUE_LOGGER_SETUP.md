# Issue Logger Setup Guide

Panduan setup untuk fitur `/noteissue` dan `/listissue` dengan Google Sheets integration.

## Prerequisites

- Google Cloud Project dengan enabled Google Sheets API
- Service Account JSON credentials dari Google Cloud Console
- Spreadsheet bernama "itops-ticket-log" (atau akan auto-created)

## Setup Steps

### 1. Create Google Cloud Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create new project atau select existing project
3. Enable Google Sheets API:
   - Search "Google Sheets API"
   - Click Enable
4. Create Service Account:
   - Go to "Credentials" → "Create Credentials" → "Service Account"
   - Fill service account name and click "Create and Continue"
   - Grant Editor role to the service account
   - Click "Continue"
   - Click "Done"

### 2. Generate JSON Key

1. Go to Service Account → Click on created service account
2. Go to "Keys" tab
3. Click "Add Key" → "Create new key"
4. Select JSON format
5. Click "Create" - file will auto-download

### 3. Configure .env

#### Option A: Base64 Encoding (Recommended for Railway)

```bash
# On Linux/Mac/Git Bash:
base64 -i /path/to/service-account.json

# On Windows PowerShell:
[Convert]::ToBase64String([System.IO.File]::ReadAllBytes("C:\path\to\service-account.json"))
```

Then copy output and paste into .env:

```env
GOOGLE_SHEETS_CREDENTIALS_JSON=<base64-encoded-json>
```

#### Option B: Raw JSON

```env
GOOGLE_SHEETS_CREDENTIALS_JSON={"type":"service_account","project_id":"...","..."}
```

### 4. Share Google Sheets (If using existing sheet)

1. Open spreadsheet "itops-ticket-log"
2. Click "Share" → Paste service account email
3. Grant Editor access

## Usage

### /noteissue - Record New Issue

**Format:**
```
/noteissue
Source: Web App / Mobile / POS
Kendala: Detail issue yang terjadi
Action: Action untuk resolve issue
```

**Example:**
```
/noteissue
Source: Web App
Kendala: POS Digital tidak bisa payment setelah update sistem
Action: Clear browser cache dan restart service POS
```

**Response:**
```
✅ Issue Berhasil Dicatat!

• Tanggal: 2026-09-04 15:30:45 WIB
• Source: Web App
• Detail Issue: POS Digital tidak bisa payment setelah update sistem
• Action Resolved: Clear browser cache dan restart service POS
```

### /listissue - View All Issues

**Usage:**
```
/listissue
```

**Response:**
- If data ≤ 3500 chars: Sends as formatted text
- If data > 3500 chars: Sends as Excel (.xlsx) file

## Data Format in Google Sheets

| Tanggal Issue | Chat ID | Source | Detail Issue | Action Resolved |
|---|---|---|---|---|
| 2026-09-04 15:30:45 WIB | -1001234567 | Web App | Issue detail... | Action taken... |
| 2026-09-04 16:45:20 WIB | -1001234567 | Mobile | Issue detail... | Action taken... |

## Troubleshooting

### Issue Logger not connecting

- Check `GOOGLE_SHEETS_CREDENTIALS_JSON` is set correctly
- Verify service account has access to "itops-ticket-log" sheet
- Check Railway logs: `⚠️  Issue Logger initialization failed`

### Base64 encoding issues

- Ensure no line breaks in base64 string
- Copy entire output as one line

### Worksheet creation failed

- Make sure service account has Editor access to spreadsheet
- Or sheet "IssueLogs" already exists

## Files

- `issue_logger.py` - Core Google Sheets operations
- `issue_commands.py` - Telegram command handlers
- `bot.py` - Integration point

## API Error Handling

All operations include try-except blocks:
- Connection errors: Return user-friendly error message
- API timeouts: Graceful fallback with message
- Invalid input: Format validation with help text
