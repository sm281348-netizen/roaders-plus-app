import pandas as pd
import streamlit as st
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from app import _get_cached_sheet

df = _get_cached_sheet('daily_data')
if df is not None and not df.empty:
    cols = ['date', 'total_rooms', 'sold_rooms', 'total inventory', 'occ rate', 'revenue', 'adr']
    valid_cols = [c for c in cols if c in df.columns]
    print(df[valid_cols].head(10).to_string())
else:
    print("Empty")
