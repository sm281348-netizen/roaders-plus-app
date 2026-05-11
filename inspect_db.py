import sys
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

conn = st.connection("gsheets", type=GSheetsConnection)
df = conn.read(worksheet="daily_data", ttl="0")
print("Total rows:", len(df))
print("\nLast 15 rows of date column:")
print(df['date'].tail(15).tolist())

try:
    print("\nCheck if today/yesterday exists:")
    recent = df[df['date'].isin(['2026-05-02', '2026-05-03', '2026-05-01'])]
    print(recent[['date', 'occ_rate', 'revenue', 'total_rooms']])
except Exception as e:
    print("Error querying specific dates:", e)
    
df_logs = conn.read(worksheet="daily_logs", ttl="0")
if df_logs is not None:
    print("\nLogs - Last 5 rows:")
    print(df_logs['date'].tail(5).tolist())
