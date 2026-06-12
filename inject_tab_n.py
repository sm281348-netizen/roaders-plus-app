with open('nationality_tab_snippet.py', 'r', encoding='utf-8') as f:
    snippet_code = f.read()

# Remove the import statements at the beginning
snippet_code = snippet_code.replace("import streamlit as st\nimport pandas as pd\nimport altair as alt\nimport io\n", "")

with open('app.py', 'a', encoding='utf-8') as f:
    f.write("\n\n" + snippet_code + "\n")
    
    f.write("""
if current_hotel != "採購":
    with tab_n:
        render_nationality_tab()
""")
