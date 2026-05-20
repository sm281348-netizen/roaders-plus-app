import streamlit as st
import json

try:
    secrets_dict = {}
    for k, v in st.secrets.items():
        if isinstance(v, st.runtime.secrets.SecretsSection):
            secrets_dict[k] = dict(v)
        else:
            secrets_dict[k] = str(v)
    
    with open("secrets_dump.json", "w") as f:
        json.dump(secrets_dict, f, indent=4)
    print("Successfully dumped secrets to secrets_dump.json!")
except Exception as e:
    print("Dump error:", e)
