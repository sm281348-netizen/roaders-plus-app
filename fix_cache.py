
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Add hotel_type param to function definition
old1 = '@st.cache_data(ttl=60)\ndef _get_cached_sheet(worksheet):\n    """\u96c6\u4e2d\u5feb\u53d6\u5c64\uff1a\u6240\u6709\u552f\u8b80 Sheet \u8acb\u6c42\u8d70\u9019\u88e1\uff0c60s TTL\uff0c\u907f\u514d API 429"""'
new1 = '@st.cache_data(ttl=60)\ndef _get_cached_sheet(worksheet, hotel_type=""):\n    """\u96c6\u4e2d\u5feb\u53d6\u5c64\uff1a\u6240\u6709\u552f\u8b80 Sheet \u8acb\u6c42\u8d70\u9019\u88e1\uff0c60s TTL\uff0c\u907f\u514d API 429\n    hotel_type \u53c3\u6578\u7528\u65bc\u5340\u5206\u4e0d\u540c\u9928\u7684\u5feb\u53d6\uff0c\u907f\u514d\u8de8\u9928\u8cc7\u6599\u6c61\u67d3\u3002"""'

# Fix 2: read_google_sheet call site
old2 = '        return _get_cached_sheet(worksheet)'
new2 = '        return _get_cached_sheet(worksheet, hotel_type=current_hotel)'

# Fix 3: All direct _get_cached_sheet("daily_data") calls
old3 = '_get_cached_sheet("daily_data")'
new3 = '_get_cached_sheet("daily_data", hotel_type=current_hotel)'

count1 = content.count(old1)
count2 = content.count(old2)
count3 = content.count(old3)

content = content.replace(old1, new1)
content = content.replace(old2, new2)
content = content.replace(old3, new3)

with open('app.py', 'w', encoding='utf-8', newline='\r\n') as f:
    f.write(content)

print(f'Fix 1 (function def): {count1} replaced')
print(f'Fix 2 (read_google_sheet call): {count2} replaced')
print(f'Fix 3 (direct calls): {count3} replaced')
print('Done!')
