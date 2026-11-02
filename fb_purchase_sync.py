# -*- coding: utf-8 -*-
r"""
餐飲請購單 (The Peak 與 4F 餐車) 每日自動比對與增量同步引擎 (F&B Purchase Auto-Sync)
================================================================================
功能說明：
1. 自動監控網路磁碟 Y:\餐務\ 最新 Excel 請購單 (4F 餐車 與 THE PEAK)。
2. 自動判斷當前月份，若月初 (1~7 日) 則自動涵蓋上月結算，支援跨月與跨週自動分流。
3. 直連 Google Sheets (4FHH_daily_purchase_report 與 thepeak_daily_purchase_report)。
4. 進行雙向集合差集比對 (Diff)，精準識別尚未登錄的新叫貨批次。
5. 具備嚴格防呆防重疊機制，將遺漏批次自動追加 (append_rows) 至 Google Sheets。
6. 本地同步儲存全量備份 CSV 至 data/pms_backup/，支援 Git 雲端同步。
"""

import os
import sys
import tomllib
import datetime
import logging
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
os.makedirs(BACKUP_DIR, exist_ok=True)

from fb_purchase_extractor import extract_4f_purchases, extract_thepeak_purchases

def get_gsheet_client():
    """透過 secrets.toml 取得 gspread 試算表物件"""
    if not os.path.exists(SECRETS_PATH):
        logger.error(f"[FB_SYNC] 找不到 secrets.toml: {SECRETS_PATH}")
        return None
    try:
        with open(SECRETS_PATH, "rb") as f:
            sec = tomllib.load(f)
        c = sec.get("connections", {}).get("gsheets_station", {})
        if c.get("type") != "service_account":
            logger.error("[FB_SYNC] gsheets_station 未設定為 service_account")
            return None
        gc = gspread.service_account_from_dict(c)
        sh = gc.open_by_url(c["spreadsheet"])
        return sh
    except Exception as e:
        logger.error(f"[FB_SYNC] 連接 Google Sheets 失敗: {e}", exc_info=True)
        return None

def sync_department_purchases(sh, ws_name: str, df_excel: pd.DataFrame, dept_name: str):
    """
    對單一餐飲部門執行比對與增量同步
    """
    if df_excel.empty:
        logger.info(f"[FB_SYNC] [{dept_name}] 網路磁碟無相關叫貨單資料，略過同步。")
        return 0

    try:
        ws = sh.worksheet(ws_name)
        existing_records = ws.get_all_records()
        df_existing = pd.DataFrame(existing_records)
        
        if not df_existing.empty:
            date_col = next((c for c in df_existing.columns if any(k in str(c) for k in ['日期', '請購日期', '叫貨日'])), df_existing.columns[0])
            existing_dates = set(pd.to_datetime(df_existing[date_col], errors='coerce').dt.strftime('%Y-%m-%d').dropna())
        else:
            existing_dates = set()

        excel_dates = sorted(df_excel['date'].unique())
        missing_batches = [d for d in excel_dates if d not in existing_dates]

        if not missing_batches:
            logger.info(f"[FB_SYNC] ✅ [{dept_name}] 線上已是最新狀態 (共 {len(excel_dates)} 個批次均已登記)，無漏更項目。")
            return 0

        logger.info(f"[FB_SYNC] 🚨 [{dept_name}] 偵測到 {len(missing_batches)} 個新批次待補入: {missing_batches}")
        
        df_missing = df_excel[df_excel['date'].isin(missing_batches)].copy()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        df_to_append = pd.DataFrame({
            "請購日期": df_missing['date'],
            "品項名稱": df_missing['item_name'],
            "當期單價": df_missing['price'],
            "單位": df_missing['unit'],
            "請購數量": df_missing['quantity'],
            "總價": df_missing['total_amount'],
            "建檔時間": now_str
        })

        cols = ['請購日期', '品項名稱', '當期單價', '單位', '請購數量', '總價', '建檔時間']
        values = df_to_append[cols].astype(str).values.tolist()

        ws.append_rows(values, value_input_option='USER_ENTERED')
        total_amt = df_missing['total_amount'].sum()
        logger.info(f"[FB_SYNC] 🎉 [{dept_name}] 成功自動同步追加 {len(values)} 列！(總金額: NT$ {total_amt:,.1f})")

        # 同步儲存至本地備份庫
        backup_csv = os.path.join(BACKUP_DIR, f"{ws_name}_latest.csv")
        df_all_updated = pd.concat([df_existing, df_to_append], ignore_index=True) if not df_existing.empty else df_to_append
        df_all_updated.to_csv(backup_csv, index=False, encoding='utf-8-sig')

        return len(values)

    except Exception as e:
        logger.error(f"[FB_SYNC] [{dept_name}] 同步寫入過程發生異常: {e}", exc_info=True)
        return -1

def run_fb_purchase_sync(target_date: datetime.date = None):
    """
    執行 F&B 採購自動同步主流程
    """
    if target_date is None:
        target_date = datetime.date.today()

    logger.info(f"========== 開始執行餐飲請購單 (F&B) 全自動同步 (基準日: {target_date}) ==========")
    sh = get_gsheet_client()
    if sh is None:
        logger.error("[FB_SYNC] 無法連接 Google Sheets，終止餐飲採購同步。")
        return False

    # 決定要抽取的月份清單：當月 + (若為月初前7天則涵蓋上月)
    months_to_scan = [target_date.month]
    if target_date.day <= 7:
        prev_month = (target_date.replace(day=1) - datetime.timedelta(days=1)).month
        if prev_month not in months_to_scan:
            months_to_scan.append(prev_month)

    year = target_date.year

    all_4f = []
    all_peak = []
    for m in months_to_scan:
        logger.info(f"[FB_SYNC] 正在解析網路磁碟 {year} 年 {m} 月叫貨單...")
        df_4f_m = extract_4f_purchases(target_year=year, target_month=m)
        if not df_4f_m.empty:
            all_4f.append(df_4f_m)

        df_peak_m = extract_thepeak_purchases(target_year=year, target_month=m)
        if not df_peak_m.empty:
            all_peak.append(df_peak_m)

    df_4f_total = pd.concat(all_4f, ignore_index=True) if all_4f else pd.DataFrame()
    df_peak_total = pd.concat(all_peak, ignore_index=True) if all_peak else pd.DataFrame()

    # 執行同步
    count_4f = sync_department_purchases(sh, "4FHH_daily_purchase_report", df_4f_total, "4F 餐車 Happy Hour")
    count_peak = sync_department_purchases(sh, "thepeak_daily_purchase_report", df_peak_total, "THE PEAK 37F")

    logger.info(f"========== 餐飲請購單同步完成 (4FHH 補入: {max(count_4f, 0)} 筆, THE PEAK 補入: {max(count_peak, 0)} 筆) ==========\n")
    return True

if __name__ == "__main__":
    # 配置本機控制台日誌輸出
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    run_fb_purchase_sync()
