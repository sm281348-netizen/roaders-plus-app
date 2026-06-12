import xlrd

# 讀取最新的 EIS 檔案 (0601)
filepath = r"Y:\櫃台\金旭\每日EIS\2026\202606\EIS0320260601.XLS"
wb = xlrd.open_workbook(filepath, encoding_override='cp950')

print("=" * 60)
print("所有工作表名稱：")
for i, name in enumerate(wb.sheet_names()):
    ws = wb.sheet_by_index(i)
    print(f"  [{i}] {name} ({ws.nrows}列 x {ws.ncols}欄)")

print("\n" + "=" * 60)
print("【工作表 0】完整內容（逐列）")
ws0 = wb.sheet_by_index(0)
for r in range(ws0.nrows):
    row_vals = []
    for c in range(ws0.ncols):
        val = ws0.cell_value(r, c)
        if val != '' and val != 0.0:
            row_vals.append(f"[{c}]{repr(val)}")
    if row_vals:
        print(f"  Row{r:02d}: {' | '.join(row_vals)}")

print("\n" + "=" * 60)
print("【工作表 1】住客來源前 35 列")
ws1 = wb.sheet_by_index(1)
for r in range(min(35, ws1.nrows)):
    row_vals = []
    for c in range(ws1.ncols):
        val = ws1.cell_value(r, c)
        if val != '' and val != 0.0:
            row_vals.append(f"[{c}]{repr(val)}")
    if row_vals:
        print(f"  Row{r:02d}: {' | '.join(row_vals)}")
