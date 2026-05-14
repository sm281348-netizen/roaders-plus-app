import pandas as pd
df_existing = pd.DataFrame({'date': ['2026-05-01', '2026-05-02', '']})
# This mimics standardize_df_dates
df_existing['date'] = df_existing['date'].astype(str)

months = {'2026-05'}
for m in months:
    mask = df_existing['date'].str.startswith(m, na=False)
    if mask.any():
        df_existing.loc[mask, 'rest_month_rev'] = 200000
        
print(df_existing)
