import pandas as pd
from streamlit_gsheets import GSheetsConnection
import streamlit as st

conn = st.connection("gsheets", type=GSheetsConnection)
df = conn.read(worksheet="daily_data", ttl="0")
df_may = df[df['date'].str.startswith('2026-05', na=False)]
print("May data:")
print(df_may[['date', 'rest_month_rev', 'bf_total_act']])
