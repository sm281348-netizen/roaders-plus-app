
# 讀取 EIS 函數
import xlrd
import os

date_str_test = "2026-06-01"
year = date_str_test[:4]
month = date_str_test[5:7]
day = date_str_test[8:10]
folder = f"Y:\\櫃台\\金旭\\每日EIS\\{year}\\{year}{month}"
filename = f"EIS03{year}{month}{day}.XLS"
filepath = os.path.join(folder, filename)

print(f"路徑: {filepath}")
print(f"檔案存在: {os.path.exists(filepath)}")

wb = xlrd.open_workbook(filepath, encoding_override='cp950')
ws = wb.sheet_by_index(1)  # 工作表 1 = 住客來源

# 找合計行（Row08）
print(f"\n工作表1 共 {ws.nrows} 列")
print("\n找合計行...")
for r in range(ws.nrows):
    row_vals = [ws.cell_value(r, c) for c in range(ws.ncols)]
    # 找有實際數字的合計行
    if r in [8, 9, 27]:
        print(f"Row{r:02d}: col2={row_vals[2]}, col3={row_vals[3]}, col5={row_vals[5]}, col6={row_vals[6]}")

# 驗證 Row08 是合計行
r = 8
revenue = ws.cell_value(r, 2)
total_rooms = ws.cell_value(r, 3)
adr = ws.cell_value(r, 5)
occ_pct = ws.cell_value(r, 6) * 100

print(f"\n=== 解析結果 ===")
print(f"revenue    = {revenue:,.0f}")
print(f"total_rooms= {total_rooms:.0f}")
print(f"adr        = {adr:,.0f}")
print(f"occ_rate   = {occ_pct:.1f}%")
