import altair as alt
import pandas as pd
df = pd.DataFrame({'date': ['2023-01-01', '2023-01-02', '2023-01-03'], 'occ_rate': [95.0, 75.0, 85.0]})
dt = pd.to_datetime(df['date'])
df['day'] = dt.dt.day
weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
df['weekday'] = dt.dt.weekday.map(weekday_map)
df['label'] = df['day'].astype(str) + " (" + df['weekday'] + ")"
df['color_category'] = df['occ_rate'].apply(lambda x: '>=90' if x >= 90.0 else ('>=80' if x >= 80.0 else '<80'))

base = alt.Chart(df).encode(
    x=alt.X('label:O', 
            title='日期', 
            sort=df['label'].tolist(),
            axis=alt.Axis(
                labelAngle=0,
                labelColor=alt.expr("datum.value.indexOf('五') > -1 || datum.value.indexOf('六') > -1 ? '#e74c3c' : '#2c3e50'")
            )),
    tooltip=['date', 'occ_rate']
)

bars = base.mark_bar().encode(
    y=alt.Y('occ_rate:Q', title='住房率 (%)', scale=alt.Scale(domain=[0, 100])),
    color=alt.Color(
        'color_category:N', 
        scale=alt.Scale(
            domain=['>=90', '>=80', '<80'], 
            range=['#e74c3c', '#3498db', '#2ecc71']
        ),
        legend=None
    )
)

text = base.mark_text(
    align='center',
    baseline='bottom',
    dy=-5,
    fontSize=14,
    fontWeight='bold'
).encode(
    y=alt.Y('occ_rate:Q'),
    text=alt.Text('occ_rate:Q', format='.1f')
)

chart = (bars + text).properties(title="Test", height=300)
chart.to_json()
print("SUCCESS")
