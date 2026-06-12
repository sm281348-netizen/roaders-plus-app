import pandas as pd
from app import conn

try:
    df = conn.read(worksheet="occ data", ttl=0)
    print("Columns:", df.columns.tolist())
    print("First 5 rows:")
    print(df.head(5))
except Exception as e:
    print("Error:", e)
