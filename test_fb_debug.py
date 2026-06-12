import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection

try:
    url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
    raw_st = st.connection("gsheets_station", type=GSheetsConnection)
    df_st = raw_st.read(worksheet="f&b_data", spreadsheet=url_st, ttl=0)
    print("Station f&b_data rows:", len(df_st) if df_st is not None else "None")
except Exception as e:
    print("Station f&b_data Error:", e)

try:
    url_th = st.secrets["connections"]["gsheets_theme"]["spreadsheet"]
    raw_th = st.connection("gsheets_theme", type=GSheetsConnection)
    df_th = raw_th.read(worksheet="f&b_data", spreadsheet=url_th, ttl=0)
    print("Theme f&b_data rows:", len(df_th) if df_th is not None else "None")
except Exception as e:
    print("Theme f&b_data Error:", e)

try:
    df_rep = raw_st.read(worksheet="f&b_report", spreadsheet=url_st, ttl=0)
    print("Station f&b_report rows:", len(df_rep) if df_rep is not None else "None")
except Exception as e:
    print("Station f&b_report Error:", e)
