# -*- coding: utf-8 -*-
r"""
房務請購單每日自動同步與全量校準引擎 (HK Purchase Auto-Sync - Production)
========================================================================
功能：
1. 自動讀取 Y:\房務\例行性表格\2026年 之最新採購累計活頁簿。
2. 提取 2026 年度所有核定請購明細（區分 OPEX 消耗備品與 CAPEX 布巾資產）。
3. 與 Google Sheets 'hk_daily_purchase_report' 進行 (請購日期, 品項名稱) 指紋比對。
4. 批次補齊缺漏期數與明細（如 0728, 0922, 1006, 1020, 1103, 布巾專案）。
5. 同步更新本地備份 CSV: data/pms_backup/hk_purchases_2026.csv。
"""

import os
import sys
import toml
import logging
import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

logger = logging.getLogger("pms_sync")
if not logger.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")
SECRETS_PATH = os.path.join(APP_DIR, ".streamlit", "secrets.toml")
os.makedirs(BACKUP_DIR, exist_ok=True)

from hk_purchase_extractor import extract_hk_purchases_2026, find_latest_hk_purchase_file

EXPECTED_HEADERS = ['請購日期', '品項名稱', '當期單價', '單位', '請購數量', '總價', '建檔時間', '類別', '庫存量']

def get_gsheet_client():
    """取得 gspread client 與試算表"""
    if not os.path.exists(SECRETS_PATH):
        logger.error(f"[HK_SYNC] 找不到 secrets.toml: {SECRETS_PATH}")
        return None, None
    try:
        sec = toml.load(SECRETS_PATH)
        c = sec.get("connections", {}).get("gsheets_station", {})
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(c, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_url(c["spreadsheet"])
        return gc, sh
    except Exception as e:
        logger.error(f"[HK_SYNC] 連接 Google Sheets 失敗: {e}", exc_info=True)
        return None, None

def sync_hk_purchases(dry_run=False):
    """執行房務部 2026 請購資料補齊與同步"""
    latest_file = find_latest_hk_purchase_file()
    if not latest_file:
        logger.error("[HK_SYNC] 找不到 2026 房務採購活頁簿")
        return False, "找不到房務採購活頁簿"

    df_excel = extract_hk_purchases_2026(latest_file)
    if df_excel.empty:
        logger.warning("[HK_SYNC] 抽取結果為空")
        return False, "抽取結果為空"

    gc, sh = get_gsheet_client()
    if not sh:
        return False, "無法連線至 Google Sheets"

    try:
        ws = sh.worksheet("hk_daily_purchase_report")
    except Exception:
        logger.warning("[HK_SYNC] 找不到工作表 hk_daily_purchase_report，正在建立...")
        ws = sh.add_worksheet(title="hk_daily_purchase_report", rows=1000, cols=15)
        ws.append_row(EXPECTED_HEADERS)

    # 讀取現存記錄
    online_rows = ws.get_all_records()
    df_online = pd.DataFrame(online_rows)
    logger.info(f"[HK_SYNC] 線上目前筆數: {len(df_online)}")

    # 確保表頭包含 類別 與 庫存量
    header_vals = ws.row_values(1)
    if len(header_vals) < len(EXPECTED_HEADERS) or '類別' not in header_vals:
        logger.info(f"[HK_SYNC] 更新工作表表頭為 9 欄標準結構...")
        if not dry_run:
            ws.update(range_name='A1:I1', values=[EXPECTED_HEADERS])

    # 建立線上指紋庫 (請購日期, 品項名稱)
    online_fingerprints = set()
    if not df_online.empty:
        for _, row in df_online.iterrows():
            d = str(row.get('請購日期', '')).strip()
            item = str(row.get('品項名稱', '')).strip()
            if d and item:
                online_fingerprints.add((d, item))

    # 找出缺漏明細
    missing_records = []
    for _, row in df_excel.iterrows():
        d = str(row.get('請購日期', '')).strip()
        item = str(row.get('品項名稱', '')).strip()
        if (d, item) not in online_fingerprints:
            missing_records.append(row)

    logger.info(f"[HK_SYNC] 經比對，發現待補齊請購明細共 {len(missing_records)} 筆")

    if not missing_records:
        logger.info("[HK_SYNC] 線上資料已是最新，無須補齊。")
        # 仍更新本地 CSV 備份
        backup_file = os.path.join(BACKUP_DIR, "hk_purchases_2026.csv")
        df_excel.to_csv(backup_file, index=False, encoding="utf-8-sig")
        return True, "線上資料已是最新"

    df_missing = pd.DataFrame(missing_records)
    logger.info("\n待補齊各期統計：\n" + str(df_missing.groupby('工作表')['總價'].agg(['count', 'sum'])))

    # 組裝寫入清單
    rows_to_append = []
    for _, r in df_missing.iterrows():
        rows_to_append.append([
            str(r.get('請購日期', '')),
            str(r.get('品項名稱', '')),
            float(r.get('當期單價', 0)),
            str(r.get('單位', '')),
            float(r.get('請購數量', 0)),
            float(r.get('總價', 0)),
            str(r.get('建檔時間', '')),
            str(r.get('類別', '日常備品(OPEX)')),
            str(r.get('庫存量', ''))
        ])

    if dry_run:
        logger.info(f"[HK_SYNC] DRY-RUN 完成，預計寫入 {len(rows_to_append)} 筆。")
        return True, f"DRY-RUN 預計寫入 {len(rows_to_append)} 筆"

    # 批次追加寫入
    logger.info(f"[HK_SYNC] 正在批次寫入 {len(rows_to_append)} 筆至 Google Sheets...")
    ws.append_rows(rows_to_append, value_input_option='USER_ENTERED')
    logger.info("[HK_SYNC] Google Sheets 補齊完成！")

    # 更新本地備份
    backup_file = os.path.join(BACKUP_DIR, "hk_purchases_2026.csv")
    df_excel.to_csv(backup_file, index=False, encoding="utf-8-sig")
    logger.info(f"[HK_SYNC] 本地全量備份已更新: {backup_file}")

    return True, f"成功補齊 {len(rows_to_append)} 筆房務請購明細"

if __name__ == '__main__':
    dry = '--dry-run' in sys.argv
    success, msg = sync_hk_purchases(dry_run=dry)
    print(f"Result: success={success}, msg={msg}")
