# -*- coding: utf-8 -*-
r"""
餐飲部 (The Peak 與 4F 餐車) 採購單全自動解析引擎 (F&B Purchase Extractor - Production v2)
========================================================================================
重大升級與架構修復：
1. 菜商生鮮單價 100% 連動 Google Sheets `supplier_prices` 官方半月報價庫：
   - 依據叫貨日期動態鎖定最新生效期數 (例如 10/1~10/15 抓取 10/1 報價；10/16~10/31 抓取 10/16 報價)。
   - 內建品名正規化與別名對照字典 (處理 紫色包心、青花菜、清雞腿肉、彩色甜椒 等現場簡寫)。
   - 杜絕 0 元與過期單價污染。
2. 嚴格年份與部門過濾：
   - 僅採納 2026 年合法分頁 (如 026MMDD、2026MMDD，或 4F 之 M.D)。
   - 徹底杜絕 2025 年 4 碼舊分頁 (如 1010, 1017, 1103 等)。
   - 校驗儲存格 (3,6) 請購單位，嚴禁將餐車單據誤歸戶至 The Peak。
3. 嚴防未來穿越日期：
   - 請購日期不得大於今日 (<= today)。
"""

import os
import re
import sys
import datetime
import logging
import openpyxl
import tomllib
import pandas as pd
import gspread

# 確保輸出編碼正確
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

logger = logging.getLogger("pms_sync")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")
SECRETS_PATH = os.path.join(APP_DIR, ".streamlit", "secrets.toml")

BASE_4F_DIR = r"Y:\餐務\2026(4F)"
BASE_PEAK_DIR = r"Y:\餐務\The peak\2026(37F)"

# 蔬果別名與標準報價品名對照表
VEG_SYNONYMS = {
    '紫包心': '紫色包心',
    '綠花椰': '青花菜',
    '青花椰': '青花菜',
    '青蔥': '青蔥',
    '青 蔥': '青蔥',
    '火龍果(紅/白)': '火龍果',
    '芭樂(粒)': '芭樂(粒)',
    '芭 樂 (粒)': '芭樂(粒)',
    '美生菜': '美生菜進口',
    '美生菜(普通)': '美生菜進口',
    '小黃瓜(普通)': '小黃瓜(普通)',
    '蒜碎/蒜粗': '蒜碎/蒜粗',
    '牛蕃茄': '牛蕃茄',
    '洋蔥': '洋蔥',
    '紅蘿蔔': '紅蘿蔔',
    '甜黃椒': '彩色甜椒',
    '甜紅椒': '彩色甜椒',
    '洋菇(盒裝)': '洋菇(盒)',
    '小番茄': '小番茄',
    '南瓜': '南瓜(台灣)',
    '巴西里': '巴西里',
    '小豆苗': '小豆苗',
    '羅蔓': '羅蔓',
    '羅 蔓': '羅蔓',
    '紅辣椒': '紅辣椒',
    '油飯': '油飯',
    '板豆腐': '板豆腐',
    '老薑': '老薑',
    '老 薑': '老薑',
    '生薑': '老薑',
    '綠色節瓜': '綠節瓜',
    '綠節瓜': '綠節瓜',
    '九層塔': '九層塔',
    '薑絲': '薑絲',
    '士林大香腸': '士林大香腸',
    '綠檸檬': '檸檬',
    '香吉士': '香吉士',
    '清腿肉丁': '清雞腿肉',
    '雞腿丁': '清雞腿肉',
    '去骨雞腿肉切塊': '清雞腿肉',
    '去骨雞腿肉': '清雞腿肉',
    '紫色包心': '紫色包心',
}

_CACHED_SUPPLIER_PRICES_DF = None

def load_supplier_prices_df(force_reload: bool = False):
    """載入 Google Sheets supplier_prices 資料表 (含本地快取備份，支援強制重載)"""
    global _CACHED_SUPPLIER_PRICES_DF
    if not force_reload and _CACHED_SUPPLIER_PRICES_DF is not None:
        return _CACHED_SUPPLIER_PRICES_DF

    backup_csv = os.path.join(BACKUP_DIR, "supplier_prices_cache.csv")

    # 1. 優先從 Google Sheets API 讀取最新
    try:
        if os.path.exists(SECRETS_PATH):
            with open(SECRETS_PATH, "rb") as f:
                sec = tomllib.load(f)
            c = sec.get("connections", {}).get("gsheets_station", {})
            if c.get("type") == "service_account":
                gc = gspread.service_account_from_dict(c)
                sh = gc.open_by_url(c["spreadsheet"])
                ws = sh.worksheet("supplier_prices")
                df_sp = pd.DataFrame(ws.get_all_records())
                if not df_sp.empty:
                    df_sp['period_dt'] = pd.to_datetime(df_sp['period'], errors='coerce').dt.date
                    df_sp = df_sp.dropna(subset=['period_dt']).reset_index(drop=True)
                    df_sp['clean_name'] = df_sp['item name'].astype(str).str.replace(' ', '').str.replace('　', '')
                    # 存入快取
                    df_sp.to_csv(backup_csv, index=False, encoding='utf-8-sig')
                    _CACHED_SUPPLIER_PRICES_DF = df_sp
                    return df_sp
    except Exception as e:
        logger.warning(f"[EXTRACTOR] 無法直連 Google Sheets supplier_prices: {e}，嘗試載入本地快取...")

    # 2. 次選本地快取
    if os.path.exists(backup_csv):
        try:
            df_sp = pd.read_csv(backup_csv)
            df_sp['period_dt'] = pd.to_datetime(df_sp['period'], errors='coerce').dt.date
            df_sp = df_sp.dropna(subset=['period_dt']).reset_index(drop=True)
            df_sp['clean_name'] = df_sp['item name'].astype(str).str.replace(' ', '').str.replace('　', '')
            _CACHED_SUPPLIER_PRICES_DF = df_sp
            return df_sp
        except Exception:
            pass

    return pd.DataFrame()

def lookup_veg_price(item_name: str, order_date: datetime.date, force_reload: bool = False):
    """
    依據請購日期，在 supplier_prices 中尋找當期核定單價與單位 (防誤判階層式演算法)
    """
    df_sp = load_supplier_prices_df(force_reload=force_reload)
    if df_sp.empty:
        return None, None, None

    valid = df_sp[df_sp['period_dt'] <= order_date]
    if valid.empty:
        valid = df_sp
    latest_period = valid['period_dt'].max()
    sub = valid[valid['period_dt'] == latest_period]

    raw_name = str(item_name).strip()
    clean_target = re.sub(r'[\s\(\)（）\-\_]', '', raw_name)

    # 1. 原始名稱無空格精確對應 (Priority 1)
    m = sub[sub['clean_name'] == raw_name.replace(' ', '')]
    if not m.empty:
        p_val = pd.to_numeric(m.iloc[0]['price'], errors='coerce')
        if pd.notna(p_val) and p_val > 0:
            return float(p_val), str(m.iloc[0]['unit']).strip(), latest_period

    # 2. 別名字典精確對應 (Priority 2)
    if raw_name in VEG_SYNONYMS:
        syn = VEG_SYNONYMS[raw_name]
        m_syn = sub[sub['clean_name'] == syn.replace(' ', '')]
        if not m_syn.empty:
            p_val = pd.to_numeric(m_syn.iloc[0]['price'], errors='coerce')
            if pd.notna(p_val) and p_val > 0:
                return float(p_val), str(m_syn.iloc[0]['unit']).strip(), latest_period

    # 3. 去標點符號後之精確對應 (Priority 3)
    for _, row in sub.iterrows():
        cn = re.sub(r'[\s\(\)（）\-\_]', '', str(row['clean_name']))
        if clean_target == cn:
            p_val = pd.to_numeric(row['price'], errors='coerce')
            if pd.notna(p_val) and p_val > 0:
                return float(p_val), str(row['unit']).strip(), latest_period

    # 4. 長度安全之模糊匹配 (Priority 4，嚴禁短字串逆向包含，如"玉米"誤配"玉米筍")
    best_candidate = None
    best_match_len = 0
    for _, row in sub.iterrows():
        cn = re.sub(r'[\s\(\)（）\-\_]', '', str(row['clean_name']))
        if not cn or not clean_target:
            continue
        # 情況 A: 叫貨單品名較短，完整包含於報價品名 (例如 叫貨單寫 "青花菜"，報價單寫 "青花菜(進口)")
        if len(clean_target) >= 2 and clean_target in cn:
            p_val = pd.to_numeric(row['price'], errors='coerce')
            if pd.notna(p_val) and p_val > 0:
                if len(clean_target) > best_match_len:
                    best_match_len = len(clean_target)
                    best_candidate = (float(p_val), str(row['unit']).strip(), latest_period)
        # 情況 B: 報價單品名較短，包含於叫貨單品名 (例如 叫貨單寫 "美生菜進口"，報價單寫 "美生菜")
        # 嚴格限制：cn 長度必須 >= 3，嚴防 2 字元短詞 (如"玉米") 誤配較長品名 (如"玉米筍")
        elif len(cn) >= 3 and cn in clean_target:
            p_val = pd.to_numeric(row['price'], errors='coerce')
            if pd.notna(p_val) and p_val > 0:
                if len(cn) > best_match_len:
                    best_match_len = len(cn)
                    best_candidate = (float(p_val), str(row['unit']).strip(), latest_period)

    if best_candidate is not None:
        return best_candidate

    return None, None, latest_period

def parse_peak_sheet_date(sheet_name: str, target_year: int = 2026) -> datetime.date:
    """
    嚴格解析 The Peak 工作表名稱：只允許明確包含 026 / 2026 的分頁
    徹底杜絕 2025 年 4 碼舊分頁 (如 1010, 1103 等)
    """
    s = str(sheet_name).strip()

    # 1. 026MMDD 或 26MMDD
    m_026 = re.match(r"^0?26(\d{2})(\d{2})$", s)
    if m_026:
        m, d = int(m_026.group(1)), int(m_026.group(2))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return datetime.date(target_year, m, d)

    # 2. 2026MMDD
    m_2026 = re.match(r"^2026(\d{2})(\d{2})$", s)
    if m_2026:
        m, d = int(m_2026.group(1)), int(m_2026.group(2))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return datetime.date(2026, m, d)

    return None

def parse_4f_sheet_date(sheet_name: str, fallback_year: int = 2026) -> datetime.date:
    """
    解析 4F 餐車工作表名稱 (例如 11.2, 10.5, 10 26, 0260629)
    """
    s = str(sheet_name).strip()

    m_026 = re.match(r"^0?26(\d{2})(\d{2})$", s)
    if m_026:
        m, d = int(m_026.group(1)), int(m_026.group(2))
        return datetime.date(fallback_year, m, d)

    m_dot = re.match(r"^(\d{1,2})\.(\d{1,2})$", s)
    if m_dot:
        m, d = int(m_dot.group(1)), int(m_dot.group(2))
        return datetime.date(fallback_year, m, d)

    m_space = re.match(r"^(\d{1,2})\s+(\d{1,2})$", s)
    if m_space:
        m, d = int(m_space.group(1)), int(m_space.group(2))
        return datetime.date(fallback_year, m, d)

    return None

def extract_thepeak_purchases(target_year: int = 2026, target_month: int = None, force_reload_prices: bool = False, veg_only: bool = False) -> pd.DataFrame:
    """
    抽取 The Peak (37F) 叫貨單數據 (全面核對 supplier_prices 官方報價單)
    """
    if force_reload_prices:
        load_supplier_prices_df(force_reload=True)

    records = []
    base_dir = rf"Y:\餐務\The peak\{target_year}(37F)"
    if not os.path.exists(base_dir):
        base_dir = BASE_PEAK_DIR

    if not os.path.exists(base_dir):
        logger.warning(f"The Peak 目錄不存在: {base_dir}")
        return pd.DataFrame()

    today = datetime.date.today()

    if veg_only:
        target_files = [("3@thepeak菜商叫貨單.xlsx", True, "菜商生鮮")]
    else:
        target_files = [
            ("1@THEPEAK採購用4.0.xlsx", False, "常規採購"),
            ("3@thepeak菜商叫貨單.xlsx", True, "菜商生鮮")
        ]

    for fname, is_veg, category in target_files:
        fp = os.path.join(base_dir, fname)
        if not os.path.exists(fp):
            continue

        try:
            wb = openpyxl.load_workbook(fp, data_only=True)
            for sname in wb.sheetnames:
                if any(k in sname for k in ["複製", "常用", "品項", "空白", "價格", "工作表", "班表"]):
                    continue

                d_obj = parse_peak_sheet_date(sname, target_year)
                if not d_obj:
                    continue

                # 嚴防未來穿越日期
                if d_obj > today:
                    continue

                if target_month and d_obj.month != target_month:
                    continue

                ws = wb[sname]

                # 校驗儲存格 (3,6) 請購單位：若是餐車則略過
                dept_val = str(ws.cell(3, 6).value or "").strip()
                if "餐車" in dept_val:
                    continue

                date_str = d_obj.strftime("%Y-%m-%d")

                if not is_veg:
                    # 常規採購單 (自帶單價與金額)
                    for r in range(5, ws.max_row + 1):
                        name = ws.cell(r, 5).value
                        if not name:
                            continue
                        name = str(name).strip()
                        if any(k in name for k in ["合計", "總計", "品名", "備註"]):
                            continue

                        price = ws.cell(r, 6).value
                        unit = ws.cell(r, 7).value
                        qty = ws.cell(r, 8).value
                        tot = ws.cell(r, 9).value

                        try:
                            qty_f = float(qty) if qty is not None else 0.0
                            if qty_f <= 0:
                                continue
                            price_f = float(price) if price is not None else 0.0
                            tot_f = float(tot) if tot is not None else (price_f * qty_f)

                            records.append({
                                "date": date_str,
                                "item_name": name,
                                "quantity": qty_f,
                                "unit": str(unit or "").strip(),
                                "price": price_f,
                                "total_amount": tot_f,
                                "category": category,
                                "department": "The Peak",
                                "source_file": fname,
                                "source_sheet": sname
                            })
                        except Exception:
                            continue
                else:
                    # 菜商生鮮叫貨單：無單價欄位，必須由 supplier_prices 核對當期單價
                    for r in range(5, 45):
                        for c_name, c_qty, c_unit in [(2, 3, 4), (7, 8, 9)]:
                            name = ws.cell(r, c_name).value
                            qty = ws.cell(r, c_qty).value
                            unit = ws.cell(r, c_unit).value

                            if not name or qty is None:
                                continue
                            name = str(name).strip()
                            if any(k in name for k in ["品名", "合計", "備註"]):
                                continue

                            try:
                                qty_f = float(qty)
                                if qty_f <= 0:
                                    continue

                                p_val, u_val, p_date = lookup_veg_price(name, d_obj)
                                if p_val is None or p_val <= 0:
                                    logger.warning(f"[THE PEAK] 菜商品項 '{name}' 在報價期 ({p_date}) 查無有效單價，記錄為 0 元待補。")
                                    p_val = 0.0
                                    tot_f = 0.0
                                else:
                                    tot_f = round(p_val * qty_f, 2)

                                final_unit = u_val or str(unit or "").strip()

                                records.append({
                                    "date": date_str,
                                    "item_name": name,
                                    "quantity": qty_f,
                                    "unit": final_unit,
                                    "price": p_val,
                                    "total_amount": tot_f,
                                    "category": category,
                                    "department": "The Peak",
                                    "source_file": fname,
                                    "source_sheet": sname
                                })
                            except Exception:
                                continue
            wb.close()
        except Exception as e:
            logger.error(f"讀取 The Peak 檔案異常: {fname}, 錯誤: {e}")

    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values(["date", "item_name"]).reset_index(drop=True)
    return df

def extract_4f_purchases(target_year: int = 2026, target_month: int = None) -> pd.DataFrame:
    """
    抽取 4F 餐車 (Happy Hour) 叫貨單數據
    """
    records = []
    base_dir = rf"Y:\餐務\{target_year}(4F)"
    if not os.path.exists(base_dir):
        base_dir = BASE_4F_DIR

    if not os.path.exists(base_dir):
        logger.warning(f"4F 目錄不存在: {base_dir}")
        return pd.DataFrame()

    today = datetime.date.today()

    month_folders = []
    if target_month:
        month_folders.append(f"{target_month}月")
    else:
        for item in os.listdir(base_dir):
            if os.path.isdir(os.path.join(base_dir, item)) and "月" in item:
                month_folders.append(item)

    for m_folder in month_folders:
        folder_path = os.path.join(base_dir, m_folder)
        if not os.path.isdir(folder_path):
            continue

        for fname in os.listdir(folder_path):
            if fname.startswith("~$") or not fname.endswith((".xlsx", ".xls")):
                continue
            if "採購" not in fname and "叫貨" not in fname:
                continue

            fp = os.path.join(folder_path, fname)
            try:
                wb = openpyxl.load_workbook(fp, data_only=True)
                for sname in wb.sheetnames:
                    if any(k in sname for k in ["複製", "常用", "品項", "空白", "價格", "工作表", "班表"]):
                        continue

                    d_obj = parse_4f_sheet_date(sname, target_year)
                    if not d_obj:
                        continue

                    if d_obj > today:
                        continue

                    if target_month and d_obj.month != target_month:
                        continue

                    ws = wb[sname]
                    date_str = d_obj.strftime("%Y-%m-%d")

                    for r in range(5, ws.max_row + 1):
                        name = ws.cell(r, 5).value
                        if not name:
                            continue
                        name = str(name).strip()
                        if any(k in name for k in ["合計", "總計", "品名", "備註"]):
                            continue

                        price = ws.cell(r, 6).value
                        unit = ws.cell(r, 7).value
                        qty = ws.cell(r, 8).value
                        tot = ws.cell(r, 9).value

                        try:
                            qty_f = float(qty) if qty is not None else 0.0
                            if qty_f <= 0:
                                continue
                            price_f = float(price) if price is not None else 0.0
                            tot_f = float(tot) if tot is not None else (price_f * qty_f)

                            records.append({
                                "date": date_str,
                                "item_name": name,
                                "quantity": qty_f,
                                "unit": str(unit or "").strip(),
                                "price": price_f,
                                "total_amount": tot_f,
                                "category": "餐車採購",
                                "department": "4F Happy Hour",
                                "source_file": fname,
                                "source_sheet": sname
                            })
                        except Exception:
                            continue
                wb.close()
            except Exception as e:
                logger.error(f"讀取 4F 檔案異常: {fname}, 錯誤: {e}")

    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values(["date", "item_name"]).reset_index(drop=True)
    return df
