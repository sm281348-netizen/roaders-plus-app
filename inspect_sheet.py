import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection
import sys

conn = st.connection("gsheets", type=GSheetsConnection)
df = conn.read(worksheet="daily_data", ttl="0")
print("Total rows:", len(df))
print("Last 10 rows date column:")
print(df['date'].tail(10).tolist())

df_logs = conn.read(worksheet="daily_logs", ttl="0")
if df_logs is not None:
    print("\nLogs - Last 5 dates:")
    print(df_logs['date'].tail(5).tolist())
