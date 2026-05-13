from flask import Flask, render_template, jsonify, request
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import numpy as np
from datetime import datetime
import json
import time

app = Flask(__name__)

# Configuration
JSON_FILE_PATH = 'service_account.json'
SHEET_NAME = 'Shops'
WORKSHEET_NAME = 'Shops'
FEEDBACK_WORKSHEET_NAME = 'Feedback'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
DATE_FORMAT = "%m/%d/%Y"
CACHE_DURATION = 600  # 10 minutes

# Cache variables
cached_data = None
feedback_cached_data = None
last_fetch_time = None
last_feedback_fetch_time = None

@app.route('/')
def dashboard():
    return render_template('reporting_dashboard.html')

@app.route('/api/reporting-data')
def get_reporting_api_data():
    """Simple API endpoint for reporting data"""
    try:
        # Return cached data if available
        if cached_data is not None:
            return jsonify({
                'status': 'success',
                'data': {'message': 'Server running in simple mode'}
            })
        
        return jsonify({
            'status': 'success',
            'data': {'message': 'Simple server working'}
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/clear-cache', methods=['POST'])
def clear_cache():
    """Clear all cached data and force fresh fetch."""
    global cached_data, feedback_cached_data, last_fetch_time, last_feedback_fetch_time
    
    try:
        cached_data = None
        feedback_cached_data = None
        last_fetch_time = None
        last_feedback_fetch_time = None
        
        return {
            'status': 'success',
            'message': 'All caches cleared successfully'
        }
        
    except Exception as e:
        return {
            'status': 'error',
            'message': f'Failed to clear cache: {str(e)}'
        }

if __name__ == '__main__':
    print("Starting simple server on http://localhost:5003")
    print("This is a minimal version for testing")
    app.run(debug=True, port=5003, use_reloader=False)
