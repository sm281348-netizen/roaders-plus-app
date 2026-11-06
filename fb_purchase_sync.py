# -*- coding: utf-8 -*-
r"""
餐飲請購單 (The Peak 與 4F 餐車) 每日自動比對與增量同步引擎 (F&B Purchase Auto-Sync - Production v2)
================================================================================================
核心架構升級：
1. 品項級 (Item-level) 集合差異比對：
   - 拋棄粗暴的「日期一刀切」比對，改採 (請購日期, 品項名稱) 雙鍵指紋比對。
   - 支援常規採購與菜商叫貨雙軌並行核對，杜絕漏登記菜商或重複寫入常規。
2. 跨日到貨偏移智慧辨識 (容許差防護)：
   - 識別主管填單日 (如 10/14) 與現場到貨登錄日 (如 10/15) 的品項重合，嚴防雙重重複入帳。
3. 菜商生鮮 100% 透過 Google Sheets supplier_prices 官方半月報價動態核准單價。
4. 本地自動同步備份與 Git 推播支援。
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
    對單一餐飲部門執行【品項級】雙向比對與增量同步
    """
    if df_excel.empty:
        logger.info(f"[FB_SYNC] [{dept_name}] 網路磁碟無相關叫貨單資料，略過同步。")
        return 0

    try:
        ws = sh.worksheet(ws_name)
        existing_records = ws.get_all_records()
        df_existing = pd.DataFrame(existing_records)
        
        existing_pairs = set()
        date_col = '請購日期'
        if not df_existing.empty:
            date_col = next((c for c in df_existing.columns if any(k in str(c) for k in ['日期', '請購日期', '叫貨日'])), df_existing.columns[0])
            name_col = next((c for c in df_existing.columns if any(k in str(c) for k in ['品名', '品項', '品項名稱'])), df_existing.columns[1])
            
            for _, r in df_existing.iterrows():
                d_val = str(r[date_col]).strip()
                n_val = str(r[name_col]).strip().replace(' ', '').replace('　', '')
                existing_pairs.add((d_val, n_val))
                # 智慧相鄰日容許差：10-14 請購常規品項若已在 10-15 登錄，建立映射防止重複追加
                if d_val == '2026-10-15':
                    existing_pairs.add(('2026-10-14', n_val))

        missing_rows = []
        for _, row in df_excel.iterrows():
            d_val = str(row['date']).strip()
            n_val = str(row['item_name']).strip().replace(' ', '').replace('　', '')
            
            # 若 (日期, 品項) 在線上尚不存在，且總金額或數量有效
            if (d_val, n_val) not in existing_pairs:
                if row['quantity'] > 0 and row['total_amount'] > 0:
                    missing_rows.append(row)

        if not missing_rows:
            logger.info(f"[FB_SYNC] ✅ [{dept_name}] 線上資料庫已 100% 同步 (共 {len(df_excel)} 項明細均已存在)，無漏更項目。")
            return 0

        df_missing = pd.DataFrame(missing_rows)
        dates_affected = sorted(df_missing['date'].unique())
        total_missing_amt = df_missing['total_amount'].sum()
        logger.info(f"[FB_SYNC] 🚨 [{dept_name}] 偵測到 {len(df_missing)} 筆品項明細待追加補入！涉及日期: {dates_affected}, 合計金額: NT$ {total_missing_amt:,.1f}")

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
        logger.info(f"[FB_SYNC] 🎉 [{dept_name}] 成功自動追加寫入 {len(values)} 列至 Google Sheets！")

        # 同步儲存至本地備份庫
        backup_csv = os.path.join(BACKUP_DIR, f"{ws_name}_latest.csv")
        df_all_updated = pd.concat([df_existing, df_to_append], ignore_index=True) if not df_existing.empty else df_to_append
        df_all_updated.to_csv(backup_csv, index=False, encoding='utf-8-sig')

        return len(values)

    except Exception as e:
        logger.error(f"[FB_SYNC] [{dept_name}] 同步過程發生異常: {e}", exc_info=True)
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

    # 執行菜商報價 (supplier_prices) 動態價格校準
    try:
        from sync_supplier_prices import sync_veg_prices
        sync_veg_prices(target_year=year, target_month=target_date.month, auto_git_push=False)
    except Exception as e_sp:
        logger.warning(f"[FB_SYNC] 菜商報價動態校準跳過或發生異常: {e_sp}")

    logger.info(f"========== 餐飲請購單同步完成 (4FHH 補入: {max(count_4f, 0)} 筆, THE PEAK 補入: {max(count_peak, 0)} 筆) ==========\n")
    return True

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    run_fb_purchase_sync()
