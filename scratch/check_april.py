import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

conn = st.connection("gsheets", type=GSheetsConnection)
df = conn.read(worksheet="daily_data", ttl="0")
print("Total rows:", len(df))

# Filter for April 2026
df_april = df[df['date'].str.startswith('2026-04', na=False)]
print("\nApril 2026 data:")
columns_to_show = [c for c in ['date', 'rest_hh_guests', 'rest_day_guests', 'bf_total_act', 'af_total_act', 'revenue'] if c in df.columns]
print(df_april[columns_to_show].to_string())
