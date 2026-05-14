import pandas as pd
df = pd.DataFrame({'date': ['A', 'B', 'C']})
mask = df['date'] == 'A'
try:
    df.loc[mask, 'new_col'] = 100
    print("Success")
except Exception as e:
    print(f"Error: {e}")
