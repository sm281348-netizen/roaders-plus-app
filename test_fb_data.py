import sys
import os

# Ensure app directory is in path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app import fetch_month_summary, _get_cached_sheet
import pandas as pd

print("1. Fetching df_all from cache...")
df_all = _get_cached_sheet("daily_data", hotel_type="站前館")

print(f"df_all empty? {df_all.empty if df_all is not None else True}")
if df_all is not None and not df_all.empty:
    print(f"Columns: {df_all.columns.tolist()}")
    
    # Check 2026-01
    df_2026 = df_all[df_all['date'].str.startswith('2026-01', na=False)]
    print(f"Rows for 2026-01: {len(df_2026)}")
    
    if not df_2026.empty:
        print("Data for 2026-01:")
        if 'rest_breakfast' in df_2026.columns:
            print(f"Total rest_breakfast: {df_2026['rest_breakfast'].sum()}")
        if 'rest_day_guests' in df_2026.columns:
            print(f"Total rest_day_guests: {df_2026['rest_day_guests'].sum()}")
        if 'rest_month_rev' in df_2026.columns:
            print(f"Total rest_month_rev: {df_2026[df_2026['rest_month_rev'] > 0]['rest_month_rev'].tail(1)}")
            
    # Print the dates in the dataframe
    print("Unique dates:")
    print(df_all['date'].unique()[:10])
    
print("2. Fetching month summary for 2026-01...")
# Note: fetch_month_summary returns a dictionary in the new version
res = fetch_month_summary(2026, 1)
print("fetch_month_summary keys:", res.keys())

# If the UI uses df_mtd directly, let's see df_mtd
if 'df' in res:
    df_mtd = res['df']
    print("df_mtd shape:", df_mtd.shape)
    if not df_mtd.empty:
        print("df_mtd dates:", df_mtd['date'].unique())
        if 'rest_breakfast' in df_mtd.columns:
            print("rest_breakfast sum:", df_mtd['rest_breakfast'].sum())
