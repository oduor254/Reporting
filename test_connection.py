import gspread
from google.oauth2.service_account import Credentials
import ssl
import certifi
import urllib3

try:
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
    
    # Try to open the sheet
    sheet = client.open(SHEET_NAME)
    print(f'Successfully opened sheet: {sheet.title}')
    
    # List all worksheets
    worksheets = sheet.worksheets()
    print(f'Available worksheets: {[ws.title for ws in worksheets]}')
    
    # Try to access the specific worksheet
    if WORKSHEET_NAME in [ws.title for ws in worksheets]:
        worksheet = sheet.worksheet(WORKSHEET_NAME)
        print(f'Successfully accessed worksheet: {worksheet.title}')
        print(f'Number of rows: {worksheet.row_count}')
    else:
        print(f'Worksheet "{WORKSHEET_NAME}" not found!')
        
except Exception as e:
    print(f'Error: {e}')
    import traceback
    traceback.print_exc()
