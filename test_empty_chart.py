import altair as alt
import pandas as pd
df = pd.DataFrame({'a': ['1 (一)', '2 (五)', '3 (六)'], 'b': [10, 20, 30], 'color': ['#e74c3c', '#3498db', '#2ecc71']})
chart = alt.Chart(df).mark_bar().encode(
    x=alt.X('a:O'),
    y='b:Q',
    color=alt.Color('color:N', scale=None)
)
chart.save('chart.json')
