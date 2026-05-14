import pandas as pd

parsed_days = [{'date': '2026-05-01', 'rest_month_rev': 200000, 'rest_avg_spent': 500}]

df_existing = pd.DataFrame({'date': ['2026-05-01', '2026-05-02'], 'rest_month_rev': [100, 100]})

months = set("-".join(str(d['date']).split('-')[:2]) for d in parsed_days)
if not df_existing.empty and 'date' in df_existing.columns:
    for m in months:
        mask = df_existing['date'].str.startswith(m, na=False)
        if mask.any():
            df_existing.loc[mask, 'rest_month_rev'] = 200000
            df_existing.loc[mask, 'rest_avg_spent'] = 500

df_new = pd.DataFrame(parsed_days)

df_new = df_new.set_index('date')
if not df_existing.empty:
    df_existing = df_existing.set_index('date')
    # 以新上傳的資料優先蓋掉舊的，但如果是新資料缺少的欄位則保留舊的
    df_final = df_new.combine_first(df_existing).reset_index()
else:
    df_final = df_new.reset_index()
    
print("Success")
