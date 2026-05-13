import requests
import json

try:
    # 1. Refresh cache
    print("Triggering refresh...")
    r_refresh = requests.post("http://localhost:5003/api/refresh-now")
    print(f"Refresh status: {r_refresh.status_code}")
    
    # 2. Fetch feedback data for a week
    print("Fetching feedback data for Week 1 of February 2026...")
    r_data = requests.get("http://localhost:5003/api/feedback-data?period=Week%201%20of%20February%202026")
    data = r_data.json()
    
    participation = data.get('analysis', {}).get('participation_targets', [])
    print(f"Found {len(participation)} participation records.")
    
    if participation:
        print("\nFirst 3 records:")
        for p in participation[:3]:
            print(p)
        print("\nTotals record:")
        for p in participation:
            if p.get('SHOP') == 'Totals':
                print(p)
    else:
        print("\n[ERROR] Still no participation data found in API response.")
        
except Exception as e:
    print(f"Error checking API: {e}")
