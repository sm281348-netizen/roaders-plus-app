import streamlit as st

if "log" not in st.session_state:
    st.session_state.log = []

def cb():
    st.session_state.log.append("Callback fired!")

if st.button("Change Date"):
    st.session_state["val"] = st.session_state.get("val", 0) + 1

st.number_input("Val", key="val", on_change=cb)

st.write("Log:")
st.write(st.session_state.log)
