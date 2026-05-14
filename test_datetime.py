import pandas as pd
df = pd.DataFrame({'date': pd.to_datetime(['2026-05-01', '2026-05-02'])})
def fix_d(val):
    if pd.isna(val): return ""
    v_str = str(val).split(' ')[0].strip()
    return v_str
df['date'] = df['date'].apply(fix_d)
print(df['date'].dtype)
mask = df['date'].str.startswith('2026-05', na=False)
print(mask.any())
