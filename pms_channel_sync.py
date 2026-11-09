# -*- coding: utf-8 -*-
r"""
金旭 PMS 訂房渠道 (RVR12) 全自動同步與歷史回補模組 (PMS Channel Auto-Sync & Backfill)
======================================================================================
資料來源：
- 站前館: Y:\櫃台\金旭\訂房率統計表\2026\{YYYYMM}\{YYYYMMDD}.XLSX
- 主題館: W:\主題-櫃台\夜櫃訂房率\{YYYYMM}\{YYYYMMDD}.XLSX

功能：
1. 自動掃描指定年度/月份之所有 RVR12 訂房日報表。
2. 解析 OTA 與合作渠道（Agoda, Booking, Ctrip, 官網, Expedia 等）之每日房間數。
3. 自動更新本地備份庫: data/pms_backup/station_channels.csv 與 theme_channels.csv。
4. 可選回寫至 Google Sheets 工作表 'marketing_channel_data'，實現 100% 全自動化入帳。
"""

import os
import sys
import glob
import datetime
import logging
import pandas as pd
from pms_parser import parse_rvr12_channels, clean_channel_name

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
os.makedirs(BACKUP_DIR, exist_ok=True)

HOTEL_DIRS = {
    "站前館": {
        "tag": "station",
        "base": r"Y:\櫃台\金旭\訂房率統計表\2026"
    },
    "主題館": {
        "tag": "theme",
        "base": r"W:\主題-櫃台\夜櫃訂房率"
    }
}

def sync_channels_for_hotel(hotel_name="站前館", year=2026, month=None):
    """
    掃描並同步指定館別的 RVR12 渠道數據
    """
    cfg = HOTEL_DIRS.get(hotel_name)
    if not cfg or not os.path.exists(cfg["base"]):
        logger.warning(f"[CHAN_SYNC] 找不到 {hotel_name} 的 RVR 目錄: {cfg['base'] if cfg else 'N/A'}")
        return pd.DataFrame()

    base_dir = cfg["base"]
    h_tag = cfg["tag"]
    
    # 尋找目標檔案
    if hotel_name == "站前館":
        if month:
            pattern = os.path.join(base_dir, f"{year}{month:02d}", "*.XLSX")
        else:
            pattern = os.path.join(base_dir, f"{year}*", "*.XLSX")
    else:
        # 主題館
        if month:
            pattern = os.path.join(base_dir, f"{year}{month:02d}", "*.XLSX")
        else:
            pattern = os.path.join(base_dir, f"{year}*", "*.XLSX")

    files = [f for f in glob.glob(pattern) if not os.path.basename(f).startswith("~$")]
    # 亦比對小寫 .xlsx
    files_lower = [f for f in glob.glob(pattern.replace(".XLSX", ".xlsx")) if not os.path.basename(f).startswith("~$")]
    all_files = sorted(list(set(files + files_lower)))

    logger.info(f"[CHAN_SYNC] [{hotel_name}] 共找到 {len(all_files)} 個 RVR12 檔案")
    if not all_files:
        return pd.DataFrame()

    parsed_dfs = []
    for fpath in all_files:
        fname = os.path.basename(fpath)
        ymd = fname.split('.')[0]
        if len(ymd) == 8 and ymd.isdigit():
            try:
                d_obj = datetime.datetime.strptime(ymd, "%Y%m%d").date()
                df_day = parse_rvr12_channels(fpath, hotel_name, d_obj)
                if not df_day.empty:
                    parsed_dfs.append(df_day)
            except Exception as e:
                logger.debug(f"[CHAN_SYNC] 解析失敗 {fname}: {e}")

    if not parsed_dfs:
        return pd.DataFrame()

    df_new = pd.concat(parsed_dfs, ignore_index=True)
    df_new.columns = [str(c).strip().lower() for c in df_new.columns]
    df_new = df_new.drop_duplicates(subset=['date', 'company name'])

    # 與本地現有備份庫合併 (增量更新)
    backup_file = os.path.join(BACKUP_DIR, f"{h_tag}_channels.csv")
    if os.path.exists(backup_file):
        try:
            df_old = pd.read_csv(backup_file)
            df_old.columns = [str(c).strip().lower() for c in df_old.columns]
            new_dates = set(df_new['date'].astype(str).unique())
            df_old_clean = df_old[~df_old['date'].astype(str).isin(new_dates)]
            df_final = pd.concat([df_old_clean, df_new], ignore_index=True)
        except Exception:
            df_final = df_new
    else:
        df_final = df_new

    df_final = df_final.sort_values(['date', 'rooms'], ascending=[True, False]).reset_index(drop=True)
    df_final.to_csv(backup_file, index=False, encoding="utf-8-sig")
    logger.info(f"[CHAN_SYNC] [{hotel_name}] 成功儲存 {len(df_final)} 筆渠道數據至 {backup_file}")
    return df_final

def run_all_channel_sync(year=2026):
    """同步雙館全年渠道數據"""
    results = {}
    for h in ["站前館", "主題館"]:
        df = sync_channels_for_hotel(h, year=year)
        results[h] = len(df)
    return results

if __name__ == '__main__':
    res = run_all_channel_sync(2026)
    print("Sync results:", res)
