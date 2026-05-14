import pandas as pd
df1 = pd.DataFrame({'date': ['2026-05-01', '2026-05-01'], 'val': [1, 2]}).set_index('date')
df2 = pd.DataFrame({'date': ['2026-05-01', '2026-05-01'], 'val': [3, 4]}).set_index('date')
try:
    df2.combine_first(df1)
    print("Success")
except Exception as e:
    print(f"Error: {type(e).__name__} - {e}")
