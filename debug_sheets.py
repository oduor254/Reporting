"""
Debug script to list all Google Sheets the service account can access
"""

import gspread
from google.oauth2.service_account import Credentials
import certifi
import os
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

JSON_FILE_PATH = r"C:\Users\Oduor\Downloads\JSON Files\reporting-488611-05bb9568ffb8.json"

# Set SSL certificates
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
os.environ['CURL_CA_BUNDLE'] = certifi.where()

scopes = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

try:
    creds = Credentials.from_service_account_file(JSON_FILE_PATH, scopes=scopes)
    client = gspread.authorize(creds)
    
    print("Service account authenticated successfully!\n")
    print("Accessible spreadsheets:\n")
    
    spreadsheets = client.list_spreadsheet_files()
    
    if spreadsheets:
        for sheet in spreadsheets:
            print(f"  • {sheet['name']}")
    else:
        print("  (No spreadsheets found - service account may not be shared with any sheets)")
    
    # Also extract service account email from credentials
    with open(JSON_FILE_PATH) as f:
        import json
        cred_data = json.load(f)
        print(f"\nService account email: {cred_data.get('client_email')}")
        print(f"\nShare your spreadsheet with this email address.")
        
except Exception as e:
    print(f"Error: {e}")
