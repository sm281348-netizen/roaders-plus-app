import pandas as pd
df = pd.DataFrame({'date': ['2026-05-01', None, pd.NA]})
mask = df['date'].str.startswith('2026-05', na=False)
print(mask)
try:
    df.loc[mask, 'val'] = 1
    print("Success")
except Exception as e:
    print(f"Error: {e}")
