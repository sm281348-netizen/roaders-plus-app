import sqlite3
import pandas as pd

conn = sqlite3.connect("roaders_plus.db")
df = pd.read_sql("SELECT * FROM daily_data WHERE date LIKE '2026-04%'", conn)
print("April 2026 data in SQLite:")
columns_to_show = [c for c in ['date', 'rest_hh_guests', 'rest_day_guests', 'bf_total_act', 'af_total_act', 'revenue'] if c in df.columns]
if not df.empty:
    print(df[columns_to_show].to_string())
else:
    print("No April data in local SQLite database.")
