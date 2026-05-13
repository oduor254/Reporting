#!/usr/bin/env python3

import pandas as pd
from report import build_monthly_summary

# Test the specific function that's causing issues
def test_build_monthly_summary():
    # Sample data similar to what the function expects
    all_rows = [
        {'Date': '2026-01-15', 'Gender': 'female', 'Total': 100},
        {'Date': '2026-01-16', 'Gender': 'male', 'Total': 200}
    ]
    
    try:
        result = build_monthly_summary(all_rows, 'January')
        print("SUCCESS:", result)
        return True
    except Exception as e:
        print("ERROR:", str(e))
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_build_monthly_summary()
    if success:
        print("Function works correctly")
    else:
        print("Function has issues")
