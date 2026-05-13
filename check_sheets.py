import gspread
import pandas as pd
import os
import certifi
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
os.environ['SSL_CERT_FILE'] = certifi.where()
import ssl

JSON_FILE_PATH = r'C:\Users\Oduor\Downloads\JSON Files\reporting-488611-05bb9568ffb8.json'
SHEET_NAME = 'Reporting'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def find_targets():
    try:
        gc = gspread.service_account(filename=JSON_FILE_PATH, scopes=SCOPES)
        spreadsheet = gc.open(SHEET_NAME)
        for ws in spreadsheet.worksheets():
            print(f"Searching in sheet: {ws.title}")
            vals = ws.get_all_values()
            for r_idx, row in enumerate(vals):
                for c_idx, cell in enumerate(row):
                    if "Feedback Targets" in str(cell):
                        print(f"FOUND 'Feedback Targets' at Row {r_idx+1}, Col {c_idx+1} in '{ws.title}'")
                        # Print the table data around it
                        for r in vals[r_idx:r_idx+25]:
                            print(r[c_idx:c_idx+10])
                        return
        print("Could not find 'Feedback Targets' table title in any sheet.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == '__main__':
    find_targets()
