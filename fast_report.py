from flask import Flask, render_template, jsonify, request
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import time

app = Flask(__name__)

# Configuration
JSON_FILE_PATH = r'C:\Users\Oduor\Downloads\JSON Files\reporting-488611-05bb9568ffb8.json'
SHEET_NAME = 'Reporting'
WORKSHEET_NAME = 'Shops'
CACHE_DURATION = 600

# Cache variables
cached_data = None
last_fetch_time = None

def get_simple_data():
    """Fast data fetch without SSL complications"""
    global cached_data, last_fetch_time
    
    current_time = time.time()
    if cached_data is not None and last_fetch_time is not None:
        if current_time - last_fetch_time < CACHE_DURATION:
            print(f"[DEBUG] Returning cached data (age: {current_time - last_fetch_time:.0f}s)")
            return cached_data
    
    print("[INFO] Fetching data...")
    start_time = time.time()
    
    try:
        # Simple authentication
        gc = gspread.service_account(JSON_FILE_PATH)
        spreadsheet = gc.open(SHEET_NAME)
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
        
        # Fast fetch
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            return {"error": "No data available"}
        
        headers = all_values[0]
        data_rows = all_values[1:]
        df = pd.DataFrame(data_rows, columns=headers)
        
        fetch_time = time.time() - start_time
        print(f"[DEBUG] Data fetch took {fetch_time:.2f}s")
        
        # Cache the data
        cached_data = df
        last_fetch_time = current_time
        return df
        
    except Exception as e:
        print(f"[ERROR] {str(e)}")
        return {"error": str(e)}

@app.route('/')
def dashboard():
    return render_template('reporting_dashboard.html')

@app.route('/api/reporting-data')
def get_reporting_api_data():
    """Fast API endpoint"""
    try:
        data = get_simple_data()
        if "error" in data:
            return jsonify({'error': data['error']}), 500
            
        return jsonify({
            'status': 'success',
            'data': {
                'message': f'Fast server working - {len(data)} records loaded',
                'record_count': len(data),
                'cache_age': time.time() - last_fetch_time if last_fetch_time else 0
            }
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/clear-cache', methods=['POST'])
def clear_cache():
    """Clear cache"""
    global cached_data, last_fetch_time
    cached_data = None
    last_fetch_time = None
    return {'status': 'success', 'message': 'Cache cleared'}

if __name__ == '__main__':
    print("Starting FAST server on http://localhost:5003")
    print("Optimized for speed - no SSL complications")
    app.run(debug=False, port=5003, use_reloader=False)
