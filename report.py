import certifi
import ssl
import os
import sys

# Force UTF-8 output on Windows so print() never fails on special characters
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# Set environment variables for SSL certificates
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
os.environ['CURL_CA_BUNDLE'] = certifi.where()

# Suppress SSL warnings if needed
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from flask import Flask, render_template, jsonify, request
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import numpy as np
from datetime import datetime
import json
import time
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter
import requests
from datetime import datetime, timedelta
from collections import Counter, defaultdict

app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False  # Allow UTF-8 characters in JSON responses

# ── Terminal request logging ──────────────────────────────────────────────────
_req_start = {}

@app.before_request
def _log_request_start():
    import flask
    _req_start[flask.g] = time.time()
    print(f"[REQ]  {request.method} {request.path}{('?' + request.query_string.decode()) if request.query_string else ''}")

@app.after_request
def _log_request_end(response):
    import flask
    elapsed = (time.time() - _req_start.pop(flask.g, time.time())) * 1000
    status = response.status_code
    icon = '✓' if status < 400 else ('✗' if status < 500 else '!!')
    print(f"[{icon}] {status}  {request.method} {request.path}  ({elapsed:.0f}ms)")
    return response
# ─────────────────────────────────────────────────────────────────────────────

# Configuration
# On Render: set GOOGLE_CREDENTIALS_PATH env var to the path of your service account JSON
# On local: falls back to the hardcoded path
JSON_FILE_PATH = os.environ.get('GOOGLE_CREDENTIALS_PATH', r'C:\Users\Oduor\Downloads\JSON Files\reporting-488611-05bb9568ffb8.json')
SHEET_NAME = 'Reporting'
WORKSHEET_NAME = 'Shops'  # Main data sheet
FEEDBACK_WORKSHEET_NAME = 'Feedback'  # Customer feedback sheet

# Google Sheets API scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Cache variables
cached_data = None
feedback_cached_data = None
last_fetch_time = None
last_feedback_fetch_time = None
CACHE_DURATION = 600  # Cache for 10 minutes instead of 5

# Date format used in your sheet's Date column
DATE_FORMAT = "%m/%d/%Y"

# Initialize cache variables at module level
feedback_cached_data = None
last_feedback_fetch_time = None
cached_data = None
last_fetch_time = None


def build_feedback_summary(feedback_rows, period=None):
    """Analyze customer feedback with dynamic period support and frontend-compatible keys."""
    if not feedback_rows:
        return {
            'total_feedback': 0,
            'daily_comparison': {},
            'weekly_comparison': {},
            'metrics_summary': {
                'total_metrics': 0,
                'numeric_metrics': 0,
                'positive_changes': 0,
                'negative_changes': 0,
                'no_changes': 0
            },
            'data_type': 'daily',
            'error': 'No feedback data available'
        }
    
    import re
    from datetime import datetime
    
    # 1. Map columns
    first_row = feedback_rows[0]
    cols = list(first_row.keys())
    
    # Identify column types
    date_cols = [c for c in cols if re.match(r'\d{1,2}/\d{1,2}/\d{4}', str(c))]
    week_cols = [c for c in cols if 'Week' in str(c) and re.search(r'\d', str(c))]
    month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    month_cols = [c for c in cols if c.strip() in month_names]
    
    # 2. Determine data type based on period or columns
    is_weekly_request = period and "Week" in str(period)
    if not is_weekly_request and not date_cols and week_cols:
        is_weekly_request = True
        
    is_monthly_request = False
    if not is_weekly_request:
        if period and any(m in str(period) for m in month_names):
            is_monthly_request = True
        elif not date_cols and not week_cols and month_cols:
            is_monthly_request = True

    result = {
        'total_feedback': len(feedback_rows),
        'metrics_summary': {
            'total_metrics': len(feedback_rows),
            'numeric_metrics': 0,
            'positive_changes': 0,
            'negative_changes': 0,
            'no_changes': 0,
            'date_columns': date_cols,
            'week_columns': week_cols,
            'month_columns': month_cols
        }
    }

    if is_weekly_request:
        result['data_type'] = 'weekly'
        weekly_comp = {}
        daily_comp = {} 
        
        target_week = ""
        prev_week = ""
        match = re.search(r'Week\s*(\d)', str(period))
        if match and week_cols:
            target_week = match.group(0)
            if target_week in week_cols:
                idx = week_cols.index(target_week)
                if idx > 0: prev_week = week_cols[idx-1]
        
        for row in feedback_rows:
            metric = row.get('Metric_Weekly', row.get('Metric', 'Unknown')).strip()
            if not metric: continue
            
            w_data = {
                'weekly_data': {c: row.get(c, '-') for c in week_cols},
                'change': row.get('% Change_Weekly', row.get('% Change', '-')),
                'status': row.get('Status_Weekly', row.get('Status', '-')),
                'target': row.get('Target_Weekly', row.get('Target', '100%'))
            }
            
            # Add automated comment for weekly vs target if applicable
            target_str = str(w_data['target']).strip()
            if not target_str or target_str == '-':
                target_str = "100%"
                w_data['target'] = "100%"

            target_val = target_str.replace('%', '').strip()
            current_val = str(row.get(target_week, '')).strip() if target_week else ''
            
            comment = ""
            try:
                if target_val and current_val and current_val != '-':
                    t_num = float(target_val)
                    
                    # 1. Extract percentage if present in parens, else raw number
                    pct_in_paren = re.search(r'\((\d+\.?\d*)\)', current_val)
                    if pct_in_paren:
                        c_num = float(pct_in_paren.group(1))
                    else:
                        c_match = re.search(r'(\d+\.?\d*)', current_val.replace(',', ''))
                        c_num = float(c_match.group(1)) if c_match else 0
                    
                    # 2. Heuristic: If it's a Rating (Metric contains specific words or value <= 5)
                    # Scale to 100% by multiplying by 20
                    is_rating = any(word in metric.lower() for word in ['rating', 'assistance', 'knowledge', 'experience', 'professionalism', 'overall', 'satisfaction'])
                    if (is_rating or c_num <= 5.0) and "responses" not in metric.lower() and "feedback" not in metric.lower():
                        if c_num <= 5.0: # Only scale if it looks like a 1-5 rating
                            c_num = c_num * 20
                            detail_str = f"Rating {c_num/20:.2f}/5 → {c_num:.1f}%"
                        else:
                            detail_str = f"{c_num:.1f}%"
                    else:
                        detail_str = f"{c_num}" if "responses" in metric.lower() else f"{c_num}%"

                    # 3. Handle Responses specially - comparison to 100% percentage doesn't make sense
                    if "responses" in metric.lower() or "received" in metric.lower():
                        comment = f"📊 Total: {c_num} responses"
                    elif c_num >= t_num:
                        comment = f"✅ Goal Met ({detail_str} vs {t_num}%)"
                    else:
                        diff = t_num - c_num
                        comment = f"⚠️ Target Gap: {diff:,.1f}% below goal ({detail_str})"
            except:
                pass
            w_data['comment'] = comment
            
            weekly_comp[metric] = w_data
            
            if target_week:
                val1 = row.get(prev_week, '-') if prev_week else '-'
                val2 = row.get(target_week, '-')
                chg = row.get('% Change_Weekly', row.get('% Change', '-'))
                
                daily_comp[metric] = {
                    'date1_name': prev_week or "Previous Week",
                    'date2_name': target_week,
                    'date1_value': val1,
                    'date2_value': val2,
                    'change': 0, 
                    'change_str': chg,
                    'analysis': row.get('Status_Weekly', row.get('Status', '-'))
                }
            
            pct_change = row.get('% Change_Weekly', row.get('% Change', '0'))

            try:
                clean_str = str(pct_change).replace('%', '').replace(',', '').strip()
                if not clean_str or clean_str == '-':
                    result['metrics_summary']['no_changes'] += 1
                else:
                    clean_val = float(clean_str)
                    result['metrics_summary']['numeric_metrics'] += 1
                    if clean_val > 0: result['metrics_summary']['positive_changes'] += 1
                    elif clean_val < 0: result['metrics_summary']['negative_changes'] += 1
                    else: result['metrics_summary']['no_changes'] += 1
            except:
                result['metrics_summary']['no_changes'] += 1
                
        result['weekly_comparison'] = weekly_comp
        result['daily_comparison'] = daily_comp # Populate this for Weekly too
    elif is_monthly_request:
        result['data_type'] = 'monthly'
        monthly_comp = {}
        
        for row in feedback_rows:
            metric = row.get('Metric_Monthly', row.get('Metric', 'Unknown')).strip()
            if not metric: continue
            
            m_data = {c: row.get(c, '-') for c in month_cols}
            chg = row.get('Change_Monthly', row.get('Change', '-'))
            analysis = row.get('Analysis_Monthly', row.get('Analysis', '-'))
            
            monthly_comp[metric] = {
                'month_data': m_data,
                'change': chg,
                'analysis': analysis
            }
            
            try:
                clean_str = str(chg).replace('%', '').replace(',', '').strip()

                if not clean_str or clean_str == '-':
                    result['metrics_summary']['no_changes'] += 1
                else:
                    clean_val = float(clean_str)
                    result['metrics_summary']['numeric_metrics'] += 1
                    if clean_val > 0: result['metrics_summary']['positive_changes'] += 1
                    elif clean_val < 0: result['metrics_summary']['negative_changes'] += 1
                    else: result['metrics_summary']['no_changes'] += 1
            except:
                result['metrics_summary']['no_changes'] += 1
                
        result['monthly_comparison'] = monthly_comp
    else:
        result['data_type'] = 'daily'
        daily_comp = {}
        
        # Identify target and previous dates
        target_date_str = ""
        prev_date_str = ""
        
        # Try to parse period if provided (YYYY-MM-DD)
        if period and re.match(r'\d{4}-\d{2}-\d{2}', str(period)):
            try:
                target_date_str = datetime.strptime(str(period), '%Y-%m-%d').strftime('%d/%m/%Y')
                # Google Sheets sometimes omits leading zeros in day/month
                if target_date_str not in date_cols:
                    # Python 3.9+ %-d isn't portable, but safe on most systems this tool uses. 
                    # Use a safer way:
                    d_obj = datetime.strptime(str(period), '%Y-%m-%d')
                    clean_target = f"{d_obj.day}/{d_obj.month}/{d_obj.year}"
                    if clean_target in date_cols: target_date_str = clean_target
            except:
                pass
        
        # Fallback or find previous
        if target_date_str in date_cols:
            idx = date_cols.index(target_date_str)
            if idx > 0: prev_date_str = date_cols[idx-1]
        elif date_cols:
            target_date_str = date_cols[-1]
            if len(date_cols) > 1: prev_date_str = date_cols[-2]

        for row in feedback_rows:
            metric = row.get('Metric_Daily', row.get('Metric', 'Unknown')).strip()
            if not metric: continue
            
            val1 = row.get(prev_date_str, '-')
            val2 = row.get(target_date_str, '-')
            change_str = row.get('Change_Daily', row.get('Change', '-'))
            analysis = row.get('Analysis_Daily', row.get('Analysis', '-'))

            
            # Change stats
            try:
                clean_str = str(change_str).replace('%', '').replace(',', '').strip()
                if not clean_str or clean_str == '-':
                    change_val = 0
                    result['metrics_summary']['no_changes'] += 1
                else:
                    change_val = float(clean_str)
                    result['metrics_summary']['numeric_metrics'] += 1
                    if change_val > 0: result['metrics_summary']['positive_changes'] += 1
                    elif change_val < 0: result['metrics_summary']['negative_changes'] += 1
                    else: result['metrics_summary']['no_changes'] += 1
            except:
                change_val = 0
                result['metrics_summary']['no_changes'] += 1

            daily_comp[metric] = {
                'date1_name': prev_date_str or "N/A",
                'date2_name': target_date_str or "N/A",
                'date1_value': val1,
                'date2_value': val2,
                'change': change_val,
                'change_str': change_str,
                'analysis': analysis
            }
        result['daily_comparison'] = daily_comp

    return result


def sanitize_text(value):
    """Replace characters that can't be encoded by the system codec with a safe substitute."""
    if not isinstance(value, str):
        return value
    # Encode to UTF-8 bytes then decode back — drops nothing.
    # Then additionally strip chars that cause charmap issues on Windows.
    try:
        return value.encode('utf-8', errors='replace').decode('utf-8')
    except Exception:
        return value.encode('ascii', errors='replace').decode('ascii')


def get_feedback_data():
    """Fetch and process customer feedback data with optimized loading."""
    global feedback_cached_data, last_feedback_fetch_time

    current_time = time.time()
    if (
        feedback_cached_data is not None
        and last_feedback_fetch_time is not None
        and current_time - last_feedback_fetch_time < CACHE_DURATION
    ):
        print(f"[DEBUG] Returning fresh feedback cache (age: {current_time - last_feedback_fetch_time:.0f}s)")
        return feedback_cached_data

    print("[INFO] Fetching feedback data from Google Sheets...")

    try:
        # Create a session with retry logic
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session = requests.Session()
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        
        creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
        if creds_json:
            creds_dict = json.loads(creds_json)
            gc = gspread.service_account_from_dict(creds_dict, scopes=SCOPES)
        else:
            gc = gspread.service_account(filename=JSON_FILE_PATH, scopes=SCOPES)
            
        # Increase timeout if possible (not all gspread versions support this directly on client)
        try:
            gc.http_client.timeout = 30
        except:
            pass
            
        spreadsheet = gc.open(SHEET_NAME)
        feedback_worksheet = spreadsheet.worksheet(FEEDBACK_WORKSHEET_NAME)

        all_values = feedback_worksheet.get_all_values()
        if not all_values or len(all_values) < 2:
            print("[WARNING] Feedback worksheet is empty")
            feedback_cached_data = [] 
            last_feedback_fetch_time = current_time
            return feedback_cached_data

        if len(all_values) < 4:
            feedback_cached_data = []
            last_feedback_fetch_time = current_time
            return feedback_cached_data

        # Custom header processing to handle multiple blocks
        raw_headers = [sanitize_text(h) for h in all_values[3]] 
        
        # Map headers to unique names based on position
        unique_headers = []
        for i, h in enumerate(raw_headers):
            if i < 6: # Daily Block (0-5)
                unique_headers.append(f"{h}_Daily" if h in ['Metric', 'Change', 'Analysis'] else h)
            elif i < 16: # Weekly Block (increased potentially for Target)
                unique_headers.append(f"{h}_Weekly" if h in ['Metric', '% Change', 'Status', 'Target'] else h)
            else: # Monthly Block (16+)
                unique_headers.append(f"{h}_Monthly" if h in ['Metric', 'Change', 'Analysis'] else h)
        
        headers = unique_headers
        data_rows = all_values[4:] 

        feedback_rows = []
        for row in data_rows:
            if not row or all(str(x).strip() == "" for x in row):
                continue
            if len(row) < len(headers):
                row = row + [""] * (len(headers) - len(row))
            sanitized_row = [sanitize_text(str(cell)) for cell in row]
            feedback_rows.append(dict(zip(headers, sanitized_row)))

        # Custom Parser for Participation Targets (Shop Table) - Highly Robust Search
        participation_rows = []
        try:
            p_start_row = -1
            col_offset = -1
            
            # Scan for "SHOP" or "LINKS SENT" header to find the table precisely
            for idx, row in enumerate(all_values):
                if not row: continue
                cleaned_row = [str(c).strip().upper() for c in row]
                
                # Look for header cells anywhere in the row
                if "SHOP" in cleaned_row and "LINKS SENT" in cleaned_row:
                    p_start_row = idx
                    col_offset = cleaned_row.index("SHOP")
                    break
                elif "LINKS SENT" in cleaned_row:
                    p_start_row = idx
                    # Find offset of first non-empty cell that might be SHOP
                    for c_idx, cell in enumerate(row):
                        if cell.strip():
                            col_offset = c_idx
                            break
                    break
            
            if p_start_row != -1 and col_offset != -1:
                print(f"[DEBUG] Found Participation table at Row {p_start_row+1}, Col {col_offset}")
                # Use standardized headers based on new specifications
                final_headers = ["SHOP", "Links Sent", "Online", "Walk-Ins", "Links %", "Total", "Target", "Performance"]
                
                for r_idx in range(p_start_row + 1, len(all_values)):
                    row = all_values[r_idx]
                    # Check if the SHOP cell (at offset) is empty
                    if len(row) <= col_offset or not str(row[col_offset]).strip():
                        continue
                    
                    shop_name = str(row[col_offset]).strip()
                    # Extract 6 columns starting from offset
                    p_vals = [str(c).strip() for c in row[col_offset:col_offset+6]]
                    if len(p_vals) < 6: p_vals += [""] * (6 - len(p_vals))
                    
                    # Skip if it's the header row again
                    if shop_name.upper() == 'SHOP': continue
                    
                    # Calculate performance vs 75% target for Online vs Links Sent
                    links_sent_str = p_vals[1].replace(',', '').strip()
                    online_str = p_vals[2].replace(',', '').strip()
                    
                    try:
                        links_sent = float(links_sent_str) if links_sent_str and links_sent_str != '-' else 0
                        online = float(online_str) if online_str and online_str != '-' else 0
                        
                        if links_sent > 0:
                            actual_pct = (online / links_sent) * 100
                            target_pct = 75.0
                            if actual_pct >= target_pct:
                                perf = "✅ Goal Met"
                            else:
                                perf = f"⚠️ {target_pct - actual_pct:.1f}% Gap"
                            target_str = "75%"
                        else:
                            target_str = "-"
                            perf = "-"
                    except: 
                        target_str = "-"
                        perf = "-"
                    
                    row_dict = dict(zip(final_headers, p_vals + [target_str, perf]))
                    participation_rows.append(row_dict)
                    
                    if shop_name.title() == 'Totals':
                        break
        except Exception as pe:
            print(f"[WARNING] Failed to parse Participation Targets: {pe}")
            import traceback
            traceback.print_exc()

        print(f"[INFO] Loaded {len(feedback_rows)} feedback metrics and {len(participation_rows)} participation records")
        
        result_data = {
            'metrics': feedback_rows,
            'participation': participation_rows
        }

        feedback_cached_data = result_data
        last_feedback_fetch_time = time.time()
        
        # Save local backup
        try:
            with open("feedback_data_cache.json", "w", encoding="utf-8") as f:
                json.dump(result_data, f, ensure_ascii=False)
            print("[INFO] Saved feedback data to local cache file")
        except Exception as e:
            print(f"[WARNING] Failed to save feedback cache file: {e}")
        
        return result_data

    except Exception as e:
        print(f"[ERROR] Failed to fetch feedback data: {str(e)}")
        
        if feedback_cached_data is not None:
            print("[INFO] Using stale memory cache due to error")
            return feedback_cached_data
            
        try:
            if os.path.exists("feedback_data_cache.json"):
                with open("feedback_data_cache.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                    print("[INFO] Loaded feedback data from local backup file")
                    feedback_cached_data = data
                    last_feedback_fetch_time = time.time() - CACHE_DURATION + 60
                    return data
        except Exception as backup_e:
            print(f"[ERROR] Failed to load local feedback backup: {backup_e}")
            
        raise Exception(f"Error fetching feedback data: {str(e)}")



@app.route('/api/refresh-now', methods=['POST'])
def refresh_now():
    """Force refresh data from Google Sheets"""
    global cached_data, last_fetch_time, feedback_cached_data, last_feedback_fetch_time
    cached_data = None
    last_fetch_time = None
    feedback_cached_data = None
    last_feedback_fetch_time = None
    return get_reporting_api_data()


@app.route('/api/feedback-data')
def get_feedback_api_data():
    """API endpoint to get customer feedback data with period support"""
    try:
        print("[API] Hitting Feedback API v2 (with period support)")
        period = request.args.get('period', 'Overall')
        raw_feedback_data = get_feedback_data()
        
        metrics_rows = raw_feedback_data.get('metrics', [])
        participation_rows = raw_feedback_data.get('participation', [])
        
        analysis = build_feedback_summary(metrics_rows, period=period)
        analysis['participation_targets'] = participation_rows
        
        return jsonify({
            "total_feedback": len(metrics_rows),
            "analysis": analysis
        })
    except Exception as e:
        print(f"Error in /api/feedback-data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def get_reporting_data():
    """Fetch and cache reporting data from Google Sheets.

    Returns a DataFrame with:
      - Date parsed to datetime
      - Shop standardized
      - Customer_ID created (Phone-based)
      - 'Organization' rows removed (if Gender column exists)
      - Year >= 2025 filter (configurable by editing below)
    """
    global cached_data, last_fetch_time

    current_time = time.time()
    if (
        cached_data is not None
        and last_fetch_time is not None
        and current_time - last_fetch_time < CACHE_DURATION
    ):
        print(f"[DEBUG] Returning cached data (age: {current_time - last_fetch_time:.0f}s)")
        return cached_data

    print("[INFO] Fetching data from Google Sheets...")
    start_time = time.time()

    try:
        creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
        if creds_json:
            creds_dict = json.loads(creds_json)
            gc = gspread.service_account_from_dict(creds_dict, scopes=SCOPES)
        else:
            gc = gspread.service_account(filename=JSON_FILE_PATH, scopes=SCOPES)
        spreadsheet = gc.open(SHEET_NAME)
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)

        print("[DEBUG] Fetching data from worksheet...")
        all_values = worksheet.get_all_values()

        if not all_values or len(all_values) < 2:
            raise ValueError("Worksheet is empty or has no headers")

        headers = [h.strip() for h in all_values[0]]
        data_rows = all_values[1:]

        df = pd.DataFrame(data_rows, columns=headers)

        if df.empty:
            raise ValueError("No data returned from Google Sheets")

        # Standardize column names (trim)
        df.columns = [str(c).strip() for c in df.columns]

        # Standardize Shop column name
        if "Location" in df.columns and "Shop" not in df.columns:
            df.rename(columns={"Location": "Shop"}, inplace=True)

        if "Shop" in df.columns:
            df["Shop"] = df["Shop"].astype(str).str.strip().str.title()
            # Filter out summary rows from shop data to avoid double counting
            df = df[~df['Shop'].str.lower().isin(['total', 'grand total', 'all shops', 'nan', 'none'])]
            # Filter out summary rows from shop data
            df = df[~df['Shop'].str.lower().isin(['total', 'grand total', 'all shops', 'nan', 'none'])]

        # Parse Date
        if "Date" not in df.columns:
            raise ValueError("Missing 'Date' column in Shops sheet")

        # Try strict format first (mm/dd/yyyy), then fall back
        df["Date"] = pd.to_datetime(df["Date"], format=DATE_FORMAT, errors="coerce")
        if df["Date"].isna().all():
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

        df = df.dropna(subset=["Date"])
        if df.empty:
            raise ValueError("All rows have invalid dates after parsing 'Date' column")

        # Filter to 2025+ (adjust if you want to include older years)
        df = df[df["Date"].dt.year >= 2025].copy()

        if df.empty:
            raise ValueError("No data matches the filter criteria (after removing Organization / year filter)")

        # Create customer identifier - Phone only (matches Excel pivot logic)
        if "Phone" in df.columns:
            df["Customer_ID"] = df["Phone"].astype(str).str.strip()
        else:
            # Fallback: use First Name + Date if Phone is missing
            fn = df["First Name"].astype(str).str.strip() if "First Name" in df.columns else "unknown"
            df["Customer_ID"] = (fn + "_" + df["Date"].dt.strftime("%Y%m%d")).astype(str)

        fetch_time = time.time() - start_time
        print(f"[DEBUG] Data load+clean took {fetch_time:.2f}s")
        print(f"[DEBUG] Final data shape: {df.shape}")
        print(f"[DEBUG] Date range: {df['Date'].min()} to {df['Date'].max()}")
        print(f"[DEBUG] Unique shops: {df['Shop'].nunique() if 'Shop' in df.columns else 'N/A'}")
        print(f"[DEBUG] Unique customers: {df['Customer_ID'].nunique()}")

        # Cache the cleaned data
        cached_data = df
        last_fetch_time = current_time

        # Best-effort local backup
        try:
            df.to_pickle("reporting_data_cache.pkl")
            print("[INFO] Saved data to local cache file")
        except Exception:
            pass

        return df

    except requests.exceptions.Timeout:
        print("[ERROR] Connection timed out to Google Sheets.")
        if cached_data is not None:
            print("[INFO] Using stale cache due to timeout")
            return cached_data
        try:
            df = pd.read_pickle("reporting_data_cache.pkl")
            print("[INFO] Loaded data from local backup file")
            cached_data = df
            last_fetch_time = time.time() - CACHE_DURATION + 60
            return df
        except Exception:
            raise Exception("Timeout connecting to Google Sheets and no local cache available. Please try again.")

    except Exception as e:
        print(f"[ERROR] Connection failed: {str(e)}")
        if cached_data is not None:
            print("[INFO] Using stale cache due to error")
            return cached_data
        try:
            df = pd.read_pickle("reporting_data_cache.pkl")
            print("[INFO] Loaded data from local backup file")
            cached_data = df
            last_fetch_time = time.time() - CACHE_DURATION + 60
            return df
        except Exception:
            raise Exception(f"Error fetching reporting data: {str(e)}")


def get_period_type(period_str):
    """Determine the period type from the period string"""
    if not period_str or period_str == "Overall":
        return "yearly"  # Default to yearly for overall
    elif period_str.startswith("Week"):
        return "weekly"
    elif len(period_str) == 10 and period_str.count('-') == 2:  # YYYY-MM-DD
        return "daily"
    elif period_str.startswith("Q"):  # Q1 2025, Q2 2025, etc.
        return "quarterly"
    elif period_str.startswith("H"):  # H1 2025, H2 2025
        return "semi_annual"
    elif len(period_str) == 4 and period_str.isdigit():  # 2025, 2026
        return "yearly"
    else:  # January 2025, February 2025, etc.
        return "monthly"


def get_analysis_date(period, raw_df):
    """Calculate analysis date based on the selected period"""
    if period == "Overall":
        return raw_df['Date'].max()
    elif period.startswith("Week"):
        # Parse "Week X of Month YYYY"
        parts = period.split()
        week_num = int(parts[1])
        month = parts[3]
        year = int(parts[4])
        
        # Convert month name to number
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
        month_num = month_map.get(month, 1)
        
        # Calculate Sunday-Saturday week
        start_of_month = pd.Timestamp(year=year, month=month_num, day=1)
        # Find Sunday of the week containing the 1st
        # pd.weekday: 0=Mon, ... 5=Sat, 6=Sun
        offset = (start_of_month.weekday() + 1) % 7
        first_sunday = start_of_month - pd.Timedelta(days=offset)
        
        start_date = first_sunday + pd.Timedelta(days=(week_num - 1) * 7)
        end_date = start_date + pd.Timedelta(days=6)
        
        # Get last day of the specified month
        if month_num == 12:
            last_day_of_month = pd.Timestamp(year=year + 1, month=1, day=1) - pd.Timedelta(days=1)
        else:
            last_day_of_month = pd.Timestamp(year=year, month=month_num + 1, day=1) - pd.Timedelta(days=1)
        
        # Constrain the end date to only include dates within the specified month
        end_date = min(end_date, last_day_of_month)
        
        # Return the end date as analysis date
        return min(end_date, raw_df['Date'].max())
    else:
        try:
            return pd.to_datetime(period)
        except:
            return raw_df['Date'].max()


def get_filtered_df(raw_df, shop, period):
    """Filter raw DataFrame by shop and period."""
    df = raw_df.copy()
    
    # Filter by shop
    if shop and shop != "All Shops":
        df = df[df['Shop'] == shop]
    
    if df.empty:
        return df
    
    # Filter by period
    if not period or period == "Overall":
        return df
    
    if period.startswith("Week"):
        parts = period.split()
        week_num = int(parts[1])
        month = parts[3]
        year = int(parts[4])
        
        # Month name to number
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
        month_num = month_map.get(month, 1)
        
        start_of_month = pd.Timestamp(year=year, month=month_num, day=1)
        # Find Sunday of the week containing the 1st
        offset = (start_of_month.weekday() + 1) % 7
        first_sunday = start_of_month - pd.Timedelta(days=offset)
        
        start_date = first_sunday + pd.Timedelta(days=(week_num - 1) * 7)
        end_date = start_date + pd.Timedelta(days=6)
        
        # Get last day of the specified month
        if month_num == 12:
            last_day_of_month = pd.Timestamp(year=year + 1, month=1, day=1) - pd.Timedelta(days=1)
        else:
            last_day_of_month = pd.Timestamp(year=year, month=month_num + 1, day=1) - pd.Timedelta(days=1)
        
        # Constrain the date range to only include dates within the specified month
        start_date = max(start_date, start_of_month)
        end_date = min(end_date, last_day_of_month)
        
        return df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
    
    if len(period) == 10 and period.count('-') == 2:
        # Daily YYYY-MM-DD
        target = pd.to_datetime(period)
        return df[df['Date'].dt.date == target.date()]
    
    # Monthly/Quarterly/Yearly/Semi-annual
    ptype = get_period_type(period)
    if ptype == "monthly":
        target = pd.to_datetime(period)
        return df[(df['Date'].dt.year == target.year) & (df['Date'].dt.month == target.month)]
    if ptype == "quarterly":
        q, y = period.split()
        q_num = int(q.replace("Q", ""))
        year = int(y)
        start_month = (q_num - 1) * 3 + 1
        end_month = start_month + 2
        return df[(df['Date'].dt.year == year) & (df['Date'].dt.month.between(start_month, end_month))]
    if ptype == "semi_annual":
        half, y = period.split()
        year = int(y)
        if half == "H1":
            return df[(df['Date'].dt.year == year) & (df['Date'].dt.month.between(1, 6))]
        return df[(df['Date'].dt.year == year) & (df['Date'].dt.month.between(7, 12))]
    if ptype == "yearly":
        year = int(period)
        return df[df['Date'].dt.year == year]
    
    return df


def get_last_week_of_month(month_num, year):
    """Return the last week number for a given month using Sunday-Saturday week buckets."""
    start_of_month = pd.Timestamp(year=year, month=month_num, day=1)
    offset = (start_of_month.weekday() + 1) % 7
    first_sunday = start_of_month - pd.Timedelta(days=offset)

    if month_num == 12:
        last_day_of_month = pd.Timestamp(year=year + 1, month=1, day=1) - pd.Timedelta(days=1)
    else:
        last_day_of_month = pd.Timestamp(year=year, month=month_num + 1, day=1) - pd.Timedelta(days=1)

    week_num = 1
    while True:
        week_start = first_sunday + pd.Timedelta(days=(week_num - 1) * 7)
        if week_start > last_day_of_month:
            return max(1, week_num - 1)
        week_num += 1


# ----------------------------- METRICS BUILDERS -----------------------------

def build_customer_metrics(today_rows, yesterday_rows):
    """Calculate customer overview metrics"""
    today_df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    yesterday_df = pd.DataFrame(yesterday_rows) if yesterday_rows else pd.DataFrame()
    
    if today_df.empty:
        return {
            'total_customers': {'value': 0, 'change': 0, 'change_pct': 0},
            'new_customers': {'value': 0, 'change': 0, 'change_pct': 0},
            'repeat_customers': {'value': 0, 'change': 0, 'change_pct': 0},
            'avg_orders_per_customer': {'value': 0, 'change': 0, 'change_pct': 0},
        }
    
    # Total customers = unique Customer_ID
    total_customers = today_df['Customer_ID'].nunique() if 'Customer_ID' in today_df.columns else 0
    yesterday_total = yesterday_df['Customer_ID'].nunique() if not yesterday_df.empty and 'Customer_ID' in yesterday_df.columns else 0
    
    # New vs repeat (by duplicates in period)
    cust_counts = today_df['Customer_ID'].value_counts() if 'Customer_ID' in today_df.columns else pd.Series(dtype=int)
    new_customers = int((cust_counts == 1).sum())
    repeat_customers = int((cust_counts > 1).sum())
    
    # Yesterday new vs repeat
    if not yesterday_df.empty and 'Customer_ID' in yesterday_df.columns:
        y_counts = yesterday_df['Customer_ID'].value_counts()
        y_new = int((y_counts == 1).sum())
        y_repeat = int((y_counts > 1).sum())
    else:
        y_new, y_repeat = 0, 0
    
    # Avg orders per customer
    avg_orders = float(len(today_df) / total_customers) if total_customers else 0
    y_avg_orders = float(len(yesterday_df) / yesterday_total) if yesterday_total else 0
    
    def calc_change(curr, prev):
        change = curr - prev
        pct = (change / prev * 100) if prev else 0
        return {'value': curr, 'change': change, 'change_pct': pct}
    
    return {
        'total_customers': calc_change(total_customers, yesterday_total),
        'new_customers': calc_change(new_customers, y_new),
        'repeat_customers': calc_change(repeat_customers, y_repeat),
        'avg_orders_per_customer': calc_change(round(avg_orders, 2), round(y_avg_orders, 2)),
    }


def build_store_performance(today_rows, yesterday_rows):
    """Build store performance metrics with period comparisons"""
    today_df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    yesterday_df = pd.DataFrame(yesterday_rows) if yesterday_rows else pd.DataFrame()
    
    if today_df.empty or 'Shop' not in today_df.columns:
        return []
    
    # 1. Current metrics
    store_summary = today_df.groupby('Shop').agg(
        unique_customers=('Customer_ID', 'nunique'),
        transactions=('Customer_ID', 'count')
    ).reset_index()
    
    # 2. Previous metrics
    if not yesterday_df.empty and 'Shop' in yesterday_df.columns:
        y_summary = yesterday_df.groupby('Shop').agg(
            y_unique_customers=('Customer_ID', 'nunique')
        ).reset_index()
        
        # Merge current with previous
        store_summary = pd.merge(store_summary, y_summary, on='Shop', how='left').fillna(0)
        store_summary['change'] = store_summary['unique_customers'] - store_summary['y_unique_customers']
    else:
        store_summary['y_unique_customers'] = 0
        store_summary['change'] = store_summary['unique_customers']
    
    store_summary['avg_orders_per_customer'] = (store_summary['transactions'] / store_summary['unique_customers']).round(2)
    
    # Rename for frontend compatibility while maintaining unique customer logic
    store_summary['display_count'] = store_summary['unique_customers']
    store_summary['display_prev'] = store_summary['y_unique_customers']
    
    # Sort by unique customers
    store_summary = store_summary.sort_values('unique_customers', ascending=False)
    
    return store_summary.to_dict('records')


def build_product_breakdown(today_rows):
    """Build product breakdown including products, categories, and colors"""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    
    result = {
        'products': [],
        'categories': [],
        'colors': []
    }
    
    # Products
    if not df.empty and 'Product' in df.columns:
        prod = df['Product'].astype(str).str.strip()
        product_summary = prod.value_counts().reset_index()
        product_summary.columns = ['product', 'count']
        result['products'] = product_summary.to_dict('records')
    
    # Categories
    if not df.empty and 'Category' in df.columns:
        cat = df['Category'].astype(str).str.strip()
        category_summary = cat.value_counts().reset_index()
        category_summary.columns = ['category', 'count']
        result['categories'] = category_summary.to_dict('records')
    
    # Colors
    if not df.empty and 'Color' in df.columns:
        color = df['Color'].astype(str).str.strip()
        color_summary = color.value_counts().reset_index()
        color_summary.columns = ['color', 'count']
        result['colors'] = color_summary.to_dict('records')
    
    return result


def build_repeat_analysis(today_rows, yesterday_rows=None):
    """Repeat analysis with comparisons"""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    ydf = pd.DataFrame(yesterday_rows) if yesterday_rows else pd.DataFrame()
    
    if df.empty or 'Customer_ID' not in df.columns:
        return {'repeat_customers': {'value': 0, 'prev': 0, 'change': 0}, 
                'new_customers': {'value': 0, 'prev': 0, 'change': 0}, 
                'repeat_pct': 0, 'new_pct': 0}
    
    # Current
    counts = df['Customer_ID'].value_counts()
    new_cust = int((counts == 1).sum())
    repeat = int((counts > 1).sum())
    total = new_cust + repeat
    
    # Previous
    if not ydf.empty and 'Customer_ID' in ydf.columns:
        y_counts = ydf['Customer_ID'].value_counts()
        y_new = int((y_counts == 1).sum())
        y_repeat = int((y_counts > 1).sum())
    else:
        y_new, y_repeat = 0, 0
    
    return {
        'repeat_customers': {'value': repeat, 'prev': y_repeat, 'change': repeat - y_repeat},
        'new_customers': {'value': new_cust, 'prev': y_new, 'change': new_cust - y_new},
        'repeat_pct': round(repeat / total * 100, 1) if total else 0,
        'new_pct': round(new_cust / total * 100, 1) if total else 0,
        'total': total
    }


def build_revenue(today_rows, yesterday_rows):
    """Revenue summary (handles currency symbols and commas, fallbacks to Price)."""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    ydf = pd.DataFrame(yesterday_rows) if yesterday_rows else pd.DataFrame()
    
    def to_num(s):
        if s.dtype == object:
            # Remove KES, commas, and other non-numeric chars except decimal
            s = s.str.replace(r'[^0-9.]', '', regex=True)
        return pd.to_numeric(s, errors='coerce').fillna(0)
    
    rev_col = 'Total' if 'Total' in df.columns else 'Price' if 'Price' in df.columns else None
    
    if df.empty or not rev_col:
        return {'revenue': {'value': 0, 'change': 0, 'change_pct': 0}}
    
    today_rev = float(to_num(df[rev_col]).sum())
    y_rev = float(to_num(ydf[rev_col]).sum()) if not ydf.empty and rev_col in ydf.columns else 0
    
    change = today_rev - y_rev
    pct = (change / y_rev * 100) if y_rev else 0
    
    num_trans = len(df)
    avg_per_trans = (today_rev / num_trans) if num_trans > 0 else 0
    
    # Calculate Average per Customer
    unique_cust = df['Customer_ID'].nunique() if 'Customer_ID' in df.columns else 0
    avg_per_cust = (today_rev / unique_cust) if unique_cust > 0 else 0
    
    # Previous average per customer for comparison
    prev_cust = ydf['Customer_ID'].nunique() if not ydf.empty and 'Customer_ID' in ydf.columns else 0
    prev_avg_per_cust = (y_rev / prev_cust) if prev_cust > 0 else 0
    
    cust_spend_change = avg_per_cust - prev_avg_per_cust
    
    return {
        'revenue': {
            'value': round(today_rev, 2), 
            'change': round(change, 2), 
            'change_pct': round(pct, 1),
            'avg_per_transaction': round(avg_per_trans, 2),
            'avg_per_customer': round(avg_per_cust, 2),
            'avg_per_customer_prev': round(prev_avg_per_cust, 2),
            'avg_per_customer_change': round(cust_spend_change, 2)
        }
    }


def build_comparison_summary(today_rows):
    """Quick comparison summary helper"""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    if df.empty:
        return {}
    return {
        'transactions': len(df),
        'unique_customers': df['Customer_ID'].nunique() if 'Customer_ID' in df.columns else 0
    }


def build_phone_kyc(today_rows, yesterday_rows=None):
    """Phone completeness summary with period comparisons"""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    ydf = pd.DataFrame(yesterday_rows) if yesterday_rows else pd.DataFrame()
    
    if df.empty or 'Phone' not in df.columns:
        return {'with_phone': {'value': 0, 'prev': 0, 'change': 0}, 
                'missing_phone': {'value': 0, 'prev': 0, 'change': 0}, 
                'invalid_phone': {'value': 0, 'prev': 0, 'change': 0}, 
                'with_phone_pct': 0}
    
    def get_kyc_data(data_df):
        if data_df.empty or 'Phone' not in data_df.columns:
            return 0, 0, 0
            
        invalid_vals = ['N/A', 'NA', 'null', 'None', '-']
        
        temp_df = data_df.copy()
        temp_df['Phone_Clean'] = temp_df['Phone'].astype(str).str.strip()
        
        if 'First Name' in temp_df.columns:
            name_series = temp_df['First Name'].astype(str).str.strip()
        elif 'Customer Name' in temp_df.columns:
            name_series = temp_df['Customer Name'].astype(str).str.strip()
        elif 'Name' in temp_df.columns:
            name_series = temp_df['Name'].astype(str).str.strip()
        else:
            name_series = pd.Series(temp_df.index.astype(str), index=temp_df.index)
            
        # Dedup based on Customer_ID typically, or Phone_Clean
        if 'Customer_ID' in temp_df.columns:
            temp_df['KYC_Dedup_Key'] = temp_df['Customer_ID']
        else:
            temp_df['KYC_Dedup_Key'] = temp_df['Phone_Clean']
            
        # Override deduplication key for invalid/missing phones to uniquely count them by name
        is_invalid = (temp_df['Phone_Clean'] == '') | temp_df['Phone_Clean'].isin(invalid_vals)
        temp_df.loc[is_invalid, 'KYC_Dedup_Key'] = temp_df.loc[is_invalid, 'Phone_Clean'] + "_" + name_series.loc[is_invalid]
        
        unique_df = temp_df.drop_duplicates(subset=['KYC_Dedup_Key'])
            
        phone_series = unique_df['Phone_Clean']
        w_phone = int(((phone_series != '') & (~phone_series.isin(invalid_vals))).sum())
        missing_count = int((phone_series == '').sum())
        invalid_count = int(phone_series.isin(invalid_vals).sum())
        return w_phone, missing_count, invalid_count

    # Current
    with_phone, missing, invalid = get_kyc_data(df)
    total = with_phone + missing + invalid
    valid_pct = round(with_phone / total * 100, 1) if total else 0
    
    # Previous
    y_with_phone, y_missing, y_invalid = get_kyc_data(ydf)
    
    return {
        'with_phone': {'value': with_phone, 'prev': y_with_phone, 'change': with_phone - y_with_phone}, 
        'missing_phone': {'value': missing, 'prev': y_missing, 'change': missing - y_missing}, 
        'invalid_phone': {'value': invalid, 'prev': y_invalid, 'change': invalid - y_invalid},
        'with_phone_pct': valid_pct
    }


def build_purchase_behavior(today_rows):
    """Purchase behavior by frequency"""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    if df.empty or 'Customer_ID' not in df.columns:
        return {}
    counts = df['Customer_ID'].value_counts()
    buckets = {
        '1_purchase': int((counts == 1).sum()),
        '2_purchases': int((counts == 2).sum()),
        '3_5_purchases': int(((counts >= 3) & (counts <= 5)).sum()),
        '6_plus_purchases': int((counts >= 6).sum()),
    }
    return buckets


def build_gender_performance(today_rows):
    """Gender performance counts based on unique customers"""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    if df.empty or 'Gender' not in df.columns:
        return []
    
    # Use unique customers for gender breakdown
    if 'Customer_ID' in df.columns:
        df_unique = df.drop_duplicates(subset=['Customer_ID'])
    else:
        df_unique = df

    g = df_unique['Gender'].astype(str).str.strip().str.title()
    s = g.value_counts().reset_index()
    s.columns = ['gender', 'count']
    
    total = s['count'].sum()
    if total > 0:
        s['percentage'] = (s['count'] / total * 100).round(1).astype(str)
    else:
        s['percentage'] = '0'
        
    return s.to_dict('records')


def build_gender_by_location(today_rows):
    """Gender by location segmentation based on unique customers per shop"""
    df = pd.DataFrame(today_rows) if today_rows else pd.DataFrame()
    if df.empty or 'Gender' not in df.columns or 'Shop' not in df.columns:
        return []
    
    # Drop duplicates so each customer counts once per shop
    if 'Customer_ID' in df.columns:
        df = df.drop_duplicates(subset=['Customer_ID', 'Shop'])
        
    df['Gender'] = df['Gender'].astype(str).str.strip().str.title()
    df['Shop'] = df['Shop'].astype(str).str.strip().str.title()
    seg = df.groupby(['Shop', 'Gender']).size().reset_index(name='count')
    return seg.to_dict('records')


def build_monthly_summary(all_rows, current_month):
    """Monthly summary helper based on raw rows"""
    print(f"[DEBUG] build_monthly_summary called with {len(all_rows)} rows for month {current_month}")
    df = pd.DataFrame(all_rows) if all_rows else pd.DataFrame()
    if df.empty or 'Date' not in df.columns:
        print("[DEBUG] Returning empty - no Date column")
        return {}
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])
    df['month'] = df['Date'].dt.strftime('%b')
    current = df[df['month'] == current_month].copy()
    if current.empty:
        print(f"[DEBUG] No data for month {current_month}")
        return {}

    total = len(current)
    print(f"[DEBUG] Total calculated: {total} (type: {type(total)})")
    
    female = int((current['Gender'].astype(str).str.lower().str.strip() == 'female').sum()) if 'Gender' in current.columns else 0
    male = int((current['Gender'].astype(str).str.lower().str.strip() == 'male').sum()) if 'Gender' in current.columns else 0
    
    # Revenue logic with fallback
    rev_col = 'Total' if 'Total' in current.columns else 'Price' if 'Price' in current.columns else None
    if rev_col:
        def clean_rev(s):
            if s.dtype == object:
                s = s.str.replace(r'[^0-9.]', '', regex=True)
            return pd.to_numeric(s, errors='coerce').fillna(0)
        revenue = float(clean_rev(current[rev_col]).sum())
    else:
        revenue = 0
    
    print(f"[DEBUG] Female: {female} (type: {type(female)}), Male: {male} (type: {type(male)})")
    
    # repeat vs new (phone-based)
    if 'Phone' in current.columns:
        current['Customer_ID'] = current['Phone'].astype(str).str.strip()
        counts = current['Customer_ID'].value_counts()
        repeat = int((counts > 1).sum())
        new_cust = int((counts == 1).sum())
    else:
        repeat, new_cust = 0, 0
    
    # daily trend (count by date)
    daily_trend = current.groupby(current['Date'].dt.date).size().reset_index(name='transactions')
    daily_trend.columns = ['date', 'transactions']
    daily_trend = daily_trend.to_dict('records')
    
    avg_daily = round(total / max(len(daily_trend), 1), 2)
    days_recorded = int(len(daily_trend))
    
    top_location = None
    if 'Shop' in current.columns:
        top_location = current['Shop'].value_counts().idxmax()
    
    top_product = None
    if 'Product' in current.columns:
        top_product = current['Product'].value_counts().idxmax()
    
    result = {
        "month": current_month,
        "total": total,
        "female": female,
        "male": male,
        "revenue": revenue,
        "repeat": repeat,
        "new": new_cust,
        "repeat_pct": round(repeat / total * 100, 1) if total else 0,
        "new_pct": round(new_cust / total * 100, 1) if total else 0,
        "avg_daily": avg_daily,
        "days_recorded": days_recorded,
        "daily_trend": daily_trend,
        "top_location": top_location,
        "top_product": top_product,
    }
    
    print(f"[DEBUG] build_monthly_summary returning: {result}")
    return result


def generate_insights(data):
    """Generate automated insights based on the report data"""
    insights = []
    
    # 1. Revenue Insight
    rev = data.get('revenue', {}).get('revenue', {})
    val = rev.get('value', 0)
    chg = rev.get('change', 0)
    pct = rev.get('change_pct', 0)
    
    if chg > 0:
        insights.append(f"🟢 Revenue is up by KES {chg:,.2f} ({pct:.1f}%) compared to the previous period.")
    elif chg < 0:
        insights.append(f"🔴 Revenue decreased by KES {abs(chg):,.2f} ({abs(pct):.1f}%) compared to previous.")
    else:
        insights.append(f"⚪ Revenue remained steady at KES {val:,.2f}.")

    # 2. Customer Insight
    cust = data.get('metrics', {}).get('total_customers', {})
    c_chg = cust.get('change', 0)
    if c_chg > 0:
        insights.append(f"📈 Customer base grew by {c_chg} unique customers.")
    
    # Customer Spend Insight
    avg_spend = data.get('revenue', {}).get('revenue', {}).get('avg_per_customer', 0)
    avg_chg = data.get('revenue', {}).get('revenue', {}).get('avg_per_customer_change', 0)
    if avg_chg > 0:
        insights.append(f"💰 High Value: Average spend per customer increased by KES {avg_chg:,.2f} to KES {avg_spend:,.2f}.")
    elif avg_chg < 0:
        insights.append(f"📉 Average spend per customer dropped by KES {abs(avg_chg):,.2f}.")
    
    repeat_pct = data.get('repeat', {}).get('repeat_pct', 0)
    if repeat_pct > 30:
        insights.append(f"🔄 Strong retention: {repeat_pct}% of customers are multi-buyers.")
    
    # 3. Store Insight
    stores = data.get('stores', [])
    if stores:
        # Get store with highest growth
        growth_stores = [s for s in stores if s.get('change', 0) > 0]
        if growth_stores:
            best_growth = max(growth_stores, key=lambda x: x.get('change', 0))
            insights.append(f"🏪 {best_growth.get('Shop')} showed the highest growth (+{best_growth.get('change')} customers).")
        
        top_store = stores[0]
        insights.append(f"🏆 {top_store.get('Shop')} remains the lead location with {top_store.get('display_count')} customers.")

    # 4. Product Insights
    prods = data.get('products', {}).get('products', [])
    if prods:
        top_prod = prods[0]
        insights.append(f"🏷️ Top Product: {top_prod.get('product')} (Sold {top_prod.get('count')} times).")
    
    cats = data.get('products', {}).get('categories', [])
    if cats:
        top_cat = cats[0]
        insights.append(f"📂 Top Category: {top_cat.get('category')}.")

    # 5. KYC Insight
    kyc_pct = data.get('kyc', {}).get('with_phone_pct', 0)
    if kyc_pct < 85:
        insights.append(f"⚠️ Data quality: Only {kyc_pct}% phone completeness. Missing data for {data.get('kyc', {}).get('missing_phone', {}).get('value')} records.")
    else:
        insights.append(f"✅ Data quality is excellent with {kyc_pct}% phone completeness.")

    return insights


@app.route('/')
@app.route('/dashboard')
def index():
    """Serve the dashboard page"""
    return render_template('reporting_dashboard.html')


@app.route('/api/reporting-data', methods=['GET'])
def get_reporting_api_data():
    """API endpoint to get reporting data"""
    try:
        start_time = time.time()
        
        # 1. Get query parameters from the URL
        shop = request.args.get('shop', 'All Shops')
        period = request.args.get('period', 'Overall')
        
        # 2. Load the full dataset
        raw_df = get_reporting_data()
        
        # 3. Generate filter options for the dropdowns based on raw data
        available_shops = sorted(raw_df['Shop'].unique().tolist())
        months = sorted(raw_df['Date'].dt.strftime('%B %Y').unique().tolist(), 
                        key=lambda x: pd.to_datetime(x))
        years = sorted(raw_df['Date'].dt.year.unique().astype(str).tolist())
        
        # Generate quarters
        quarters = []
        for year in years:
            for q in ['Q1', 'Q2', 'Q3', 'Q4']:
                quarters.append(f"{q} {year}")
        
        # Determine period type and analysis date
        period_type = get_period_type(period)
        analysis_date = get_analysis_date(period, raw_df)
        print(f"[INFO] Period: {period} -> Type: {period_type}")
        print(f"[INFO] Analysis Date: {analysis_date}")
        
        # 4. Filter the raw data based on shop and period FIRST
        filtered_transactions = get_filtered_df(raw_df, shop, period)
        
        if filtered_transactions.empty:
            return jsonify({'error': 'No data found for the selected filters'}), 404
        
        print(f"[DEBUG] Filtered transactions: {len(filtered_transactions)} records")
        print(f"[DEBUG] Filtered transactions unique customers: {filtered_transactions['Customer_ID'].nunique()}")
        
        # 5. For comparison, get previous period data
        prev_period = ""
        if period == "Overall":
            yesterday_rows = pd.DataFrame()
        elif len(period) == 10 and period.count('-') == 2:  # Daily: YYYY-MM-DD
            target_date = pd.to_datetime(period)
            prev_date = target_date - pd.Timedelta(days=1)
            prev_period = prev_date.strftime('%Y-%m-%d')
            yesterday_rows = get_filtered_df(raw_df, shop, prev_period)
        elif period.startswith("Week"):  # Weekly: "Week X of Month YYYY"
            parts = period.split()
            week_num = int(parts[1])
            month = parts[3]
            year = parts[4]
            
            if week_num > 1:
                prev_period = f"Week {week_num - 1} of {month} {year}"
            else:
                # When week_num == 1, use the final week number of the previous month
                month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                             "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
                month_num = month_map[month]
                if month_num == 1:
                    prev_month_num = 12
                    prev_year = int(year) - 1
                else:
                    prev_month_num = month_num - 1
                    prev_year = int(year)

                prev_month_name = [name for name, num in month_map.items() if num == prev_month_num][0]
                prev_week_num = get_last_week_of_month(prev_month_num, prev_year)
                prev_period = f"Week {prev_week_num} of {prev_month_name} {prev_year}"
            
            yesterday_rows = get_filtered_df(raw_df, shop, prev_period)
        else:
            try:
                target_date = pd.to_datetime(period)
                prev_date = target_date - pd.offsets.MonthBegin(1)
                prev_period = prev_date.strftime('%B %Y')
                yesterday_rows = get_filtered_df(raw_df, shop, prev_period)
            except:
                yesterday_rows = pd.DataFrame()
        
        # 6. Convert to dict format
        today_rows_dict = filtered_transactions.to_dict('records')
        yesterday_rows_dict = yesterday_rows.to_dict('records') if not yesterday_rows.empty else []
        
        # 7. Calculate all metrics
        current_month = analysis_date.strftime("%b")
        
        # Calculate date range info
        min_date = filtered_transactions['Date'].min()
        max_date = filtered_transactions['Date'].max()
        num_days = (max_date - min_date).days + 1
        
        # Determine comparison label
        if period == "Overall":
            comparison_label = "No Comparison"
        elif len(period) == 10 and period.count('-') == 2:
            target_date = pd.to_datetime(period)
            prev_date = target_date - pd.Timedelta(days=1)
            comparison_label = prev_date.strftime('%m/%d/%Y')
        elif period.startswith("Week"):
            comparison_label = prev_period
        else:
            try:
                target_date = pd.to_datetime(period)
                prev_date = target_date - pd.offsets.MonthBegin(1)
                comparison_label = prev_date.strftime('%B %Y')
            except:
                comparison_label = "Previous Period"

        display_period = period
        if period.startswith("Week"):
            # Keep the UI text concise for weekly selections
            parts = period.split()
            if len(parts) >= 2:
                display_period = f"Week {parts[1]}"

        # 8. Build the complete report data
        data = {
            'analysis_date': analysis_date.strftime('%Y-%m-%d'),
            'period_type': period_type,
            'date_range': {
                'start_date': min_date.strftime('%B %d, %Y'),
                'end_date': max_date.strftime('%B %d, %Y'),
                'num_days': int(num_days),
                'num_transactions': int(len(filtered_transactions))
            },
            'report_date': analysis_date.strftime("%b %d, %Y"),
            'period': period,
            'display_period': display_period,
            'current_label': display_period,
            'comparison_label': comparison_label,
            'analyst': 'Brayan Oduor',
            'generated_at': datetime.now().strftime("%H:%M, %d %b %Y"),
            'today_count': len(today_rows_dict),
            'metrics': build_customer_metrics(today_rows_dict, yesterday_rows_dict),
            'stores': build_store_performance(today_rows_dict, yesterday_rows_dict),
            'products': build_product_breakdown(today_rows_dict),
            'repeat': build_repeat_analysis(today_rows_dict, yesterday_rows_dict),
            'revenue': build_revenue(today_rows_dict, yesterday_rows_dict),
            'comparison': build_comparison_summary(today_rows_dict),
            'kyc': build_phone_kyc(today_rows_dict, yesterday_rows_dict),
            'purchase_behavior': build_purchase_behavior(today_rows_dict),
            'gender_performance': build_gender_performance(today_rows_dict),
            'segments': build_gender_by_location(today_rows_dict),
            'monthly': build_monthly_summary(
                (raw_df[raw_df['Shop'] == shop] if shop != 'All Shops' else raw_df).to_dict('records'), 
                current_month
            ),
            'filters': {
                'shops': available_shops,
                'periods': {
                    'Monthly': months,
                    'Quarterly': quarters,
                    'Yearly': years,
                    'Semi_Annually': [f"H1 {y}" for y in years] + [f"H2 {y}" for y in years]
                }
            }
        }
        
        # Add automated insights
        data['summary'] = generate_insights(data)
        
        print(f"[TIME] Reporting calculation ({shop}/{period}) in {time.time() - start_time:.2f}s")
        return jsonify(data)
    
    except Exception as e:
        print(f"Error in /api/reporting-data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500



if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    debug_mode = os.environ.get('FLASK_ENV', 'production') != 'production'
    print("\n" + "="*60)
    print("Starting Report Dashboard")
    print("="*60)
    print(f"[LINK] Open in browser: http://localhost:{port}")
    print(f"[INFO] Report Analysis using February 2026 Shops Data")
    print("="*60 + "\n")
    
    app.run(debug=debug_mode, port=port, host='0.0.0.0', use_reloader=False)