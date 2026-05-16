import altair as alt
import pandas as pd
df = pd.DataFrame({'a': [1, 2, 3], 'b': [1, 2, 3]})
color=alt.condition(
    alt.datum.a >= 2, 
    alt.value('red'), 
    alt.condition(
        alt.datum.a >= 1, 
        alt.value('blue'), 
        alt.value('green')
    )
)
print(color)
