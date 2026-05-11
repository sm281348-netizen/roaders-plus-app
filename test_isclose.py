import math
import numpy as np
import pandas as pd

import datetime

df = pd.DataFrame({"counter_complaints": [np.nan]})
db_val = df.iloc[0]["counter_complaints"]

# What does the check do?
default_val = ""
curr_val = ""

if pd.isna(db_val) or db_val is None:
    norm_db = default_val
else:
    try:
        if isinstance(default_val, int): norm_db = int(float(db_val))
        elif isinstance(default_val, float): norm_db = float(db_val)
        else: norm_db = str(db_val)
    except:
        norm_db = default_val

print("curr_val == norm_db:", curr_val == norm_db)
print("type dict:", type(curr_val), type(norm_db))
