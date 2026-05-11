import pandas as pd

df = pd.DataFrame({'date': ['2026-04-30', '115/04/30', '04/30', pd.Timestamp('2026-05-01')]})
df['date_str'] = df['date'].astype(str).str.split(' ').str[0]
print("date_str:")
print(df['date_str'])

# what does to_datetime do?
df['parsed'] = pd.to_datetime(df['date_str'], errors='coerce')
print("\nparsed:")
print(df['parsed'])
