import gspread
from google.oauth2.service_account import Credentials
import os
import certifi
import json

# Set SSL certs
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
os.environ['SSL_CERT_FILE'] = certifi.where()

JSON_FILE_PATH = r"C:\Users\Oduor\Downloads\JSON Files\reporting-488611-05bb9568ffb8.json"
SHEET_NAME = "Main Reporting Sheet 2026"
FEEDBACK_WORKSHEET_NAME = "Feedback"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def dump_participation_area():
    try:
        gc = gspread.service_account(filename=JSON_FILE_PATH, scopes=SCOPES)
        spreadsheet = gc.open(SHEET_NAME)
        worksheet = spreadsheet.worksheet(FEEDBACK_WORKSHEET_NAME)
        all_values = worksheet.get_all_values()
        
        print(f"Total rows: {len(all_values)}")
        for i in range(25, min(len(all_values), 60)):
            print(f"Row {i+1}: {all_values[i]}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    dump_participation_area()
