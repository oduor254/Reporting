import gspread
from google.oauth2.service_account import Credentials
import os
import certifi
import pandas as pd

# Set SSL certs
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
os.environ['SSL_CERT_FILE'] = certifi.where()

JSON_FILE_PATH = r"C:\Users\Oduor\Downloads\JSON Files\reporting-488611-05bb9568ffb8.json"
SHEET_NAME = "Reporting"
FEEDBACK_WORKSHEET_NAME = "Feedback"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def save_feedback_raw():
    try:
        gc = gspread.service_account(filename=JSON_FILE_PATH, scopes=SCOPES)
        spreadsheet = gc.open(SHEET_NAME)
        worksheet = spreadsheet.worksheet(FEEDBACK_WORKSHEET_NAME)
        all_values = worksheet.get_all_values()
        
        df = pd.DataFrame(all_values)
        df.to_csv("feedback_raw.csv", index=False, header=False)
        print(f"Saved {len(all_values)} rows to feedback_raw.csv")
        
        # Also print rows around 30 helpful to debug
        for i in range(25, min(len(all_values), 55)):
            print(f"Row {i+1}: {all_values[i]}")
            
    except Exception as e:
        import traceback
        print(f"Error: {repr(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    save_feedback_raw()
