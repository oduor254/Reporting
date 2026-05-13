import gspread
from google.oauth2.service_account import Credentials
import os
import certifi

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

def debug_feedback_sheet():
    try:
        gc = gspread.service_account(filename=JSON_FILE_PATH, scopes=SCOPES)
        spreadsheet = gc.open(SHEET_NAME)
        worksheet = spreadsheet.worksheet(FEEDBACK_WORKSHEET_NAME)
        all_values = worksheet.get_all_values()
        
        print(f"Total rows in feedback sheet: {len(all_values)}")
        
        # Search for "Feedback Targets"
        target_found = False
        for idx, row in enumerate(all_values):
            row_str = " | ".join([str(c) for c in row])
            if "Feedback Targets" in row_str:
                print(f"\n[FOUND] 'Feedback Targets' at Row {idx+1}")
                target_found = True
                # Show surrounding rows
                for i in range(max(0, idx-1), min(len(all_values), idx+25)):
                    print(f"Row {i+1}: {all_values[i][:6]}")
                break
        
        if not target_found:
            print("\nUpdating search... scanning all rows for any participation-like table")
            for idx, row in enumerate(all_values):
                if row and len(row) > 0 and str(row[0]).strip().upper() == 'SHOP':
                    print(f"\n[POSSIBLE TABLE] 'SHOP' header found at Row {idx+1}")
                    for i in range(idx, min(len(all_values), idx+25)):
                         print(f"Row {i+1}: {all_values[i][:6]}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    debug_feedback_sheet()
