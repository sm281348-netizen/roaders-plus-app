import streamlit as st
import json

secrets_dict = {}
for k, v in st.secrets.items():
    try:
        secrets_dict[k] = dict(v)
    except:
        secrets_dict[k] = str(v)

with open("secrets_dump.json", "w") as f:
    json.dump(secrets_dict, f, indent=4)

st.write("Secrets dumped successfully!")
st.write(secrets_dict)
