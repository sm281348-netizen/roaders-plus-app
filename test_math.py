import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from app import _get_cached_sheet, standardize_df_dates

df = _get_cached_sheet('daily_data')
df = standardize_df_dates(df)
df['revenue'] = pd.to_numeric(df['total room revenue'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
df['sold_rooms'] = pd.to_numeric(df['total rooms sold'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
df['total_rooms'] = pd.to_numeric(df['rooms available to sell'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
df['occ_rate'] = pd.to_numeric(df['occ rate'].astype(str).str.replace('%', ''), errors='coerce').fillna(0)
df['adr_col'] = pd.to_numeric(df['adr (rooms sold)'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

# April 2026
df_apr = df[(df['date'] >= '2026-04-01') & (df['date'] <= '2026-04-30')]
# Wait, user's screenshot was 2025 data, let's filter for 2025-04!
df_apr = df[(df['date'] >= '2025-04-01') & (df['date'] <= '2025-04-30')]
if df_apr.empty:
    df_apr = df[(df['date'] >= '2026-04-01') & (df['date'] <= '2026-04-30')]

if not df_apr.empty:
    sum_rev = df_apr['revenue'].sum()
    sum_sold = df_apr['sold_rooms'].sum()
    sum_avail = df_apr['total_rooms'].sum()
    
    # Method 1: standard sum
    occ1 = (sum_sold / sum_avail) * 100
    adr1 = sum_rev / sum_sold
    revpar1 = adr1 * (occ1 / 100)
    
    # Method 2: average of dailies
    occ2 = df_apr['occ_rate'].mean()
    adr2 = df_apr['adr_col'].mean()
    revpar2 = (df_apr['revenue'] / df_apr['total_rooms']).mean()
    
    print(f"Method 1 (Sums): Occ={occ1:.2f}%, ADR={adr1:.2f}, RevPAR={revpar1:.2f}")
    print(f"Method 2 (Avgs): Occ={occ2:.2f}%, ADR={adr2:.2f}, RevPAR={revpar2:.2f}")
else:
    print("No data for April")
