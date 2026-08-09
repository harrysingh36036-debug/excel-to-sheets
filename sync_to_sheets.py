import os
import sys
import json
import glob
import openpyxl
import gspread
from oauth2client.service_account import ServiceAccountCredentials

excel_files = (
    glob.glob("**/*.xlsx", recursive=True)
    + glob.glob("**/*.xls", recursive=True)
)
excel_files = [f for f in excel_files if not f.startswith(".git/")]

if not excel_files:
    print("ERROR: No Excel file (.xlsx or .xls) found in the repository.")
    print("Upload a .xlsx or .xls file to trigger the sync.")
    sys.exit(1)

EXCEL_FILE = excel_files[0]
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")

if not GOOGLE_SHEET_ID:
    print("ERROR: GOOGLE_SHEET_ID secret is not set.")
    sys.exit(1)

if not SERVICE_ACCOUNT_JSON:
    print("ERROR: GOOGLE_SERVICE_ACCOUNT_JSON secret is not set.")
    sys.exit(1)

if not GOOGLE_SHEET_ID:
    print("ERROR: GOOGLE_SHEET_ID secret is not set.")
    sys.exit(1)

if not SERVICE_ACCOUNT_JSON:
    print("ERROR: GOOGLE_SERVICE_ACCOUNT_JSON secret is not set.")
    sys.exit(1)

try:
    service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
except json.JSONDecodeError:
    print("ERROR: GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON.")
    sys.exit(1)

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

credentials = ServiceAccountCredentials.from_json_keyfile_dict(
    service_account_info, scope
)
client = gspread.authorize(credentials)

try:
    sheet = client.open_by_key(GOOGLE_SHEET_ID)
except gspread.SpreadsheetNotFound:
    print(f"ERROR: Could not find spreadsheet with ID '{GOOGLE_SHEET_ID}'.")
    print("Make sure the Sheet ID is correct and the service account has edit access.")
    sys.exit(1)

worksheet = sheet.sheet1

workbook = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
worksheet_excel = workbook.active

rows = list(worksheet_excel.iter_rows(values_only=True))

if not rows:
    print("ERROR: The Excel file appears to be empty.")
    sys.exit(1)

print(f"Read {len(rows)} rows from '{EXCEL_FILE}' (sheet: '{worksheet_excel.title}').")

worksheet.clear()

worksheet.update("A1", rows)

sheet_url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/edit"
print(f"Successfully synced to Google Sheet: {sheet_url}")
