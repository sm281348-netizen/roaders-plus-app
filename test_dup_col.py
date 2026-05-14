import pandas as pd
df = pd.DataFrame([[1, 2], [3, 4]], columns=['A', 'A'])
mask = pd.Series([True, False])
try:
    df.loc[mask, 'A'] = 100
    print("Success")
except Exception as e:
    print(f"Error: {type(e).__name__} - {e}")
