import altair as alt
import pandas as pd
df = pd.DataFrame({'a': ['1 (一)', '2 (五)', '3 (六)'], 'b': [1, 2, 3]})
chart = alt.Chart(df).mark_bar().encode(
    x=alt.X('a:O', axis=alt.Axis(
        labelColor=alt.expr("datum.value.indexOf('五') > -1 || datum.value.indexOf('六') > -1 ? 'red' : 'black'")
    )),
    y='b:Q'
)
chart.to_json()
