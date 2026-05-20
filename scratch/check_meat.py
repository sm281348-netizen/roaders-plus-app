import pandas as pd
from streamlit_gsheets import GSheetsConnection
import streamlit as st

conn = st.connection("gsheets", type=GSheetsConnection)
df = conn.read(worksheet="purchase data", ttl="0")
df.columns = df.columns.astype(str).str.strip()

print("All columns:", df.columns.tolist())
date_col = next((c for c in df.columns if '日期' in c or 'Date' in c), None)
item_col = next((c for c in df.columns if any(k in c for k in ['品名', '項目', 'Item'])), None)
total_col = next((c for c in df.columns if any(k in c for k in ['小計', '金額', 'Total'])), None)
qty_col = next((c for c in df.columns if any(k in c for k in ['數量', 'Qty', 'Quantity'])), None)

df['日期'] = pd.to_datetime(df[date_col], errors='coerce')
df_april = df[(df['日期'] >= '2026-04-01') & (df['日期'] <= '2026-04-30')].copy()

meat_df = df_april[df_april[item_col].astype(str).str.contains('肉', na=False)].copy()
print(f"Total meat rows found: {len(meat_df)}")
if not meat_df.empty:
    print(meat_df[[date_col, item_col, qty_col, total_col]].head(20))
else:
    print("No meat rows found in April!")
