import altair as alt
import pandas as pd

df = pd.DataFrame({'day': [1, 2, 3], 'occ_rate': [70, 85, 95], 'weekday': ['(一)', '(五)', '(六)']})
df['label'] = df['day'].astype(str) + " " + df['weekday']

chart = alt.Chart(df).mark_bar().encode(
    x=alt.X('label:O', axis=alt.Axis(
        labelAngle=0,
        labelColor=alt.expr("datum.value.indexOf('五') > -1 || datum.value.indexOf('六') > -1 ? 'red' : 'black'")
    )),
    y='occ_rate:Q'
)
print(chart.to_json())
