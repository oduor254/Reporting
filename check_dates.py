import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import certifi
import urllib3

JSON_FILE_PATH = r'C:\Users\Oduor\Downloads\JSON Files\reporting-488611-05bb9568ffb8.json'
SHEET_NAME = 'February 2026'
WORKSHEET_NAME = 'Shops'

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

scopes = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]

import os as os_module
os_module.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
os_module.environ['CURL_CA_BUNDLE'] = certifi.where()

creds = Credentials.from_service_account_file(JSON_FILE_PATH, scopes=scopes)
client = gspread.authorize(creds)

sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)
records = sheet.get_all_records()

# Get unique dates
dates = set()
for r in records:
    date_str = str(r.get("Date", "")).strip()
    if date_str:
        dates.add(date_str)

print(f"Total records: {len(records)}")
print(f"Unique dates found: {len(dates)}")
print("\nLast 10 dates (chronological):")
sorted_dates = sorted(dates, key=lambda x: datetime.strptime(x, "%m/%d/%Y"))
for date in sorted_dates[-10:]:
    print(f"  {date}")

print(f"\nLatest date in sheet: {sorted_dates[-1]}")
