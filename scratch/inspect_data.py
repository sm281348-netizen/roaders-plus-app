import pandas as pd
from streamlit_gsheets import GSheetsConnection
import streamlit as st

# Setup connection
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    
    print("Worksheets inspection...")
    possible_names = ["purchase data", "Purchase Data", "purchase_data", "Purchase_Data"]
    for name in possible_names:
        try:
            df = conn.read(worksheet=name, ttl="0")
            print(f"\n--- Worksheet: {name} ---")
            print(f"Columns: {df.columns.tolist()}")
            if '部門' in df.columns:
                print(f"Unique Departments: {df['部門'].unique().tolist()}")
            elif 'Dept' in df.columns:
                print(f"Unique Departments: {df['Dept'].unique().tolist()}")
            print("Sample data:")
            print(df.head(5))
        except Exception as e:
            print(f"Worksheet {name} not found or error: {e}")

    # Inspect daily_data columns too
    df_daily = conn.read(worksheet="daily_data", ttl="0")
    print(f"\n--- Worksheet: daily_data ---")
    print(f"Columns: {df_daily.columns.tolist()}")
    
except Exception as e:
    print(f"Connection error: {e}")
