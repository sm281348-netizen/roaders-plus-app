r"""
金旭 PMS (Kingxu / HT2000) 每日夜審檔案解析引擎
===================================================
功能：
1. 自動適配【站前館】與【主題館】之內部伺服器目錄路徑 (Y:\ 與 W:\)
2. 解析每日營業日報表 (EIS03{YYYYMMDD}.XLS):
   - 提取客房營運核心數據 (occ_data: 可售房數、售出房數、住房率、營收、ADR)
   - 提取營業收入會計科目累計 (EIS_data: 各科目本日與本月累計金額)
3. 解析訂房日報表 (RVR12 / {YYYYMMDD}.XLSX):
   - 提取各大 OTA 與合約渠道訂房間數 (marketing_channel_data)
4. 提供單日提取與日期區間批次回補介面
"""

import os
import sys
import datetime
import pandas as pd
import xlrd
import openpyxl

# 雙館路徑與標準基礎設定
HOTEL_CONFIGS = {
    "站前館": {
        "eis_base": r"Y:\櫃台\金旭\每日EIS",
        "rvr_base": r"Y:\櫃台\金旭\訂房率統計表",
        "total_rooms_default": 128,
    },
    "主題館": {
        "eis_base": r"W:\主題-櫃台\金旭(帳務)\EIS報表",
        "rvr_base": r"W:\主題-櫃台\夜櫃訂房率",
        "total_rooms_default": 112,
    }
}

def clean_channel_name(name: str) -> str:
    """標準化渠道公司名稱"""
    if not name or pd.isna(name):
        return "其他"
    name = str(name).strip().upper()
    if 'AGODA' in name:
        return 'AGODA'
    if 'CTRIP' in name or 'TRIP.COM' in name or '攜程' in name:
        return 'CTRIP'
    if 'BOOKING' in name and 'BUTTON' not in name:
        return 'BOOKING'
    if 'EXPEDIA' in name:
        return 'EXPEDIA'
    if 'RAKUTEN' in name:
        return 'RAKUTEN'
    if '易遊' in name:
        return '易遊網'
    if '雄獅' in name or 'FG' in name:
        return '雄獅'
    if 'BOOKINGBUTTON' in name:
        return '官網(BookingButton)'
    if '【' in name and '】' in name:
        return name.split('】')[0].replace('【', '').strip()
    return name

def to_date_obj(date_val) -> datetime.date:
    """將各類日期輸入 (str, date, datetime) 轉為 datetime.date"""
    if isinstance(date_val, datetime.date):
        return date_val
    if isinstance(date_val, datetime.datetime):
        return date_val.date()
    s = str(date_val).strip().split(' ')[0].replace('/', '-').replace('.0', '')
    if len(s) == 8 and s.isdigit():
        return datetime.datetime.strptime(s, "%Y%m%d").date()
    return pd.to_datetime(s).date()

def resolve_file_paths(hotel_name: str, d_obj: datetime.date):
    """
    根據館別與日期解析出 EIS03 與 RVR12 的檔案路徑
    """
    cfg = HOTEL_CONFIGS.get(hotel_name)
    if not cfg:
        raise ValueError(f"未知的館別: {hotel_name}，僅支援 '站前館' 或 '主題館'")

    y_str = d_obj.strftime("%Y")
    ym_str = d_obj.strftime("%Y%m")
    ymd_str = d_obj.strftime("%Y%m%d")

    # 1. EIS03 路徑
    if hotel_name == "站前館":
        # 站前館目錄分層: {YYYY}\{YYYYMM}\EIS03{YYYYMMDD}.XLS
        eis_path = os.path.join(cfg["eis_base"], y_str, ym_str, f"EIS03{ymd_str}.XLS")
        if not os.path.exists(eis_path):
            eis_path_alt = os.path.join(cfg["eis_base"], f"EIS03{ymd_str}.XLS")
            if os.path.exists(eis_path_alt):
                eis_path = eis_path_alt
    else:
        # 主題館為扁平目錄或月份目錄
        eis_path = os.path.join(cfg["eis_base"], f"EIS03{ymd_str}.XLS")
        if not os.path.exists(eis_path):
            eis_path_alt = os.path.join(cfg["eis_base"], ym_str, f"EIS03{ymd_str}.XLS")
            if os.path.exists(eis_path_alt):
                eis_path = eis_path_alt

    # 2. RVR12 路徑
    if hotel_name == "站前館":
        # 站前館 RVR12: {YYYY}\{YYYYMM}\{YYYYMMDD}.XLSX
        rvr_path = os.path.join(cfg["rvr_base"], y_str, ym_str, f"{ymd_str}.XLSX")
    else:
        # 主題館 RVR12: {YYYYMM}\{YYYYMMDD}.XLSX
        rvr_path = os.path.join(cfg["rvr_base"], ym_str, f"{ymd_str}.XLSX")

    return eis_path, rvr_path

def parse_eis03_occ(file_path: str, hotel_name: str, d_obj: datetime.date) -> dict:
    """
    從 EIS03 提取住房營運指標 (對應 occ_data)
    回傳字典格式:
    { 'date', 'sold_rooms', 'total_rooms', 'occ_rate', 'revenue', 'adr' }
    """
    if not os.path.exists(file_path):
        return None

    try:
        wb = xlrd.open_workbook(file_path, encoding_override='cp950')
        ws0 = wb.sheet_by_name('營業總日報表')
        ws1 = wb.sheet_by_name('客房分析表')
    except Exception as e:
        print(f"[Error] 開啟 EIS03 失敗 ({file_path}): {e}")
        return None

    # 1. Sheet 0: 提取可售房數 (動態扣除 OOO) 與 出租率
    total_rooms = None
    occ_rate_pct = None
    for r in range(ws0.nrows):
        c1 = str(ws0.cell_value(r, 1)).strip()
        if c1 == '可售房數' and total_rooms is None:
            try:
                total_rooms = int(float(ws0.cell_value(r, 2)))
            except:
                pass
        elif c1 in ('實際銷售房數', '實際出租率') and total_rooms is not None:
            for c in range(2, min(5, ws0.ncols)):
                val = ws0.cell_value(r, c)
                if isinstance(val, (int, float)) and 0 < val <= 1.0:
                    occ_rate_pct = round(val * 100, 2)
                    break

    # 2. Sheet 1: 提取客房合計列 (售出房數、營收、ADR)
    sold_rooms = 0
    revenue = 0
    adr = 0
    for r in range(ws1.nrows):
        c1 = str(ws1.cell_value(r, 1)).strip()
        if c1 == '合計':
            try:
                revenue = int(float(ws1.cell_value(r, 2)))
                sold_rooms = int(float(ws1.cell_value(r, 3)))
                adr = int(float(ws1.cell_value(r, 5)))
            except:
                pass
            break

    # 若報表未解析出可售房數，使用館別預設房數
    if not total_rooms or total_rooms <= 0:
        total_rooms = HOTEL_CONFIGS[hotel_name]["total_rooms_default"]

    # 若出租率未讀出，以 售出 / 可售 計算
    if occ_rate_pct is None and total_rooms > 0:
        occ_rate_pct = round(sold_rooms / total_rooms * 100, 2)

    # 計算 ADR (若報表無)
    if (not adr or adr <= 0) and sold_rooms > 0:
        adr = int(round(revenue / sold_rooms))

    return {
        "date": d_obj.strftime("%Y-%m-%d"),
        "sold_rooms": sold_rooms,
        "total_rooms": total_rooms,
        "occ_rate": occ_rate_pct,
        "revenue": revenue,
        "adr": adr
    }

def parse_eis03_financial(file_path: str, hotel_name: str, d_obj: datetime.date) -> pd.DataFrame:
    """
    從 EIS03 提取營業收入會計科目明細 (對應 EIS_data)
    回傳 DataFrame 包含: ['年/月', '項目', '本月實際', '本日金額']
    """
    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:
        wb = xlrd.open_workbook(file_path, encoding_override='cp950')
        ws0 = wb.sheet_by_name('營業總日報表')
    except Exception as e:
        print(f"[Error] 開啟 EIS03 失敗 ({file_path}): {e}")
        return pd.DataFrame()

    ym_str = d_obj.strftime("%Y/%m")
    records = []

    # 遍歷營業總日報表第 3 列至第 35 列之會計項目
    for r in range(3, min(40, ws0.nrows)):
        item = str(ws0.cell_value(r, 1)).strip()
        if not item or item in ('小計', '餐廳合計', '總收入', '客房餐飲', '客房統計'):
            continue

        try:
            today_val = float(ws0.cell_value(r, 2))
        except:
            today_val = 0.0

        try:
            mtd_val = float(ws0.cell_value(r, 4))
        except:
            mtd_val = 0.0

        records.append({
            "年/月": ym_str,
            "項目": item,
            "本日金額": today_val,
            "本月實際": mtd_val
        })

    return pd.DataFrame(records)

def parse_rvr12_channels(file_path: str, hotel_name: str, d_obj: datetime.date) -> pd.DataFrame:
    """
    從 RVR12 訂房日報表提取 OTA 渠道通路數據 (對應 marketing_channel_data)
    回傳 DataFrame 包含: ['date', 'company name', 'rooms']
    """
    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
    except Exception as e:
        print(f"[Error] 開啟 RVR12 失敗 ({file_path}): {e}")
        return pd.DataFrame()

    header_row = None
    col_status = None
    col_company = None
    col_rooms = None

    # 定位標題欄位
    for r in range(1, min(10, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or '').strip() for c in range(1, ws.max_column + 1)]
        for idx, val in enumerate(row_vals, start=1):
            if '合約公司' in val or '公司名稱' in val:
                header_row = r
                col_company = idx
            elif '狀態' in val:
                col_status = idx
            elif '總房間數' in val or '房間數' in val:
                col_rooms = idx
        if header_row and col_company and col_rooms:
            break

    if not header_row:
        return pd.DataFrame()

    d_str = d_obj.strftime("%Y-%m-%d")
    channel_counts = {}

    for r in range(header_row + 1, ws.max_row + 1):
        status = str(ws.cell(r, col_status).value or '').strip().upper() if col_status else 'NEW'
        if status in ('CANC', 'VOID', '刪除'):
            continue

        raw_company = str(ws.cell(r, col_company).value or '').strip()
        if not raw_company or raw_company.startswith('===') or '合計' in raw_company:
            continue

        rooms_val = ws.cell(r, col_rooms).value
        try:
            rooms = int(float(str(rooms_val).replace(',', '')))
        except:
            rooms = 0

        if rooms <= 0:
            continue

        clean_name = clean_channel_name(raw_company)
        channel_counts[clean_name] = channel_counts.get(clean_name, 0) + rooms

    records = []
    for ch_name, count in sorted(channel_counts.items(), key=lambda x: x[1], reverse=True):
        records.append({
            "date": d_str,
            "company name": ch_name,
            "rooms": count
        })

    return pd.DataFrame(records)

def extract_pms_day(hotel_name: str, date_val):
    """
    綜合提取指定館別與日期的全量 PMS 數據
    回傳字典:
    {
        'date': str,
        'occ': dict,
        'eis': DataFrame,
        'channels': DataFrame,
        'files_found': bool
    }
    """
    d_obj = to_date_obj(date_val)
    eis_path, rvr_path = resolve_file_paths(hotel_name, d_obj)

    occ_data = parse_eis03_occ(eis_path, hotel_name, d_obj)
    eis_df = parse_eis03_financial(eis_path, hotel_name, d_obj)
    chan_df = parse_rvr12_channels(rvr_path, hotel_name, d_obj)

    files_found = os.path.exists(eis_path) or os.path.exists(rvr_path)

    return {
        "date": d_obj.strftime("%Y-%m-%d"),
        "hotel": hotel_name,
        "occ": occ_data,
        "eis": eis_df,
        "channels": chan_df,
        "files_found": files_found,
        "eis_path": eis_path,
        "rvr_path": rvr_path
    }

def extract_pms_range(hotel_name: str, start_date, end_date):
    """
    批次提取指定區間內的所有 PMS 數據
    """
    s_obj = to_date_obj(start_date)
    e_obj = to_date_obj(end_date)
    if s_obj > e_obj:
        s_obj, e_obj = e_obj, s_obj

    current = s_obj
    occ_rows = []
    chan_rows = []
    eis_dfs = []

    while current <= e_obj:
        day_res = extract_pms_day(hotel_name, current)
        if day_res["occ"]:
            occ_rows.append(day_res["occ"])
        if not day_res["channels"].empty:
            chan_rows.append(day_res["channels"])
        if not day_res["eis"].empty:
            eis_dfs.append(day_res["eis"])
        current += datetime.timedelta(days=1)

    df_occ = pd.DataFrame(occ_rows) if occ_rows else pd.DataFrame()
    df_chan = pd.concat(chan_rows, ignore_index=True) if chan_rows else pd.DataFrame()
    df_eis = eis_dfs[-1] if eis_dfs else pd.DataFrame()  # 財務取區間最後一天累計

    return {
        "hotel": hotel_name,
        "df_occ": df_occ,
        "df_chan": df_chan,
        "df_eis": df_eis
    }

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="金旭 PMS 數據解析引擎")
    parser.add_argument("--hotel", default="站前館", choices=["站前館", "主題館"], help="館別")
    parser.add_argument("--date", default=None, help="指定日期 (YYYY-MM-DD 或 YYYYMMDD)，預設為昨天")
    parser.add_argument("--days", type=int, default=1, help="回推解析天數")

    args = parser.parse_args()

    target_date = to_date_obj(args.date) if args.date else datetime.date.today() - datetime.timedelta(days=1)
    print(f"執行館別: {args.hotel}, 目標日期: {target_date}, 範圍: {args.days} 天")

    if args.days > 1:
        start_d = target_date - datetime.timedelta(days=args.days - 1)
        res = extract_pms_range(args.hotel, start_d, target_date)
        print("\n=== 批次 occ_data ===")
        print(res["df_occ"].to_string(index=False) if not res["df_occ"].empty else "無資料")
        print("\n=== 批次 channels (前 10 筆) ===")
        print(res["df_chan"].head(10).to_string(index=False) if not res["df_chan"].empty else "無資料")
    else:
        day_res = extract_pms_day(args.hotel, target_date)
        print("\n=== 單日 occ_data ===")
        print(day_res["occ"])
        print("\n=== 單日 channels ===")
        print(day_res["channels"].to_string(index=False) if not day_res["channels"].empty else "無資料")
        print(f"\n=== 單日 EIS 科目數: {len(day_res['eis'])} 項 ===")
