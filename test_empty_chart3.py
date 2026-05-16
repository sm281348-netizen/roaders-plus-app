import altair as alt
import pandas as pd
df = pd.DataFrame({'a': ['1 (一)', '2 (五)', '3 (六)'], 'b': [10, 20, 30]})

chart = alt.Chart(df).mark_bar().encode(
    x=alt.X('a:O', axis=alt.Axis(
        labelColor={"condition": {"test": "datum.value.indexOf('五') > -1 || datum.value.indexOf('六') > -1", "value": "#e74c3c"}, "value": "#2c3e50"}
    )),
    y='b:Q'
)

print(chart.to_json())
