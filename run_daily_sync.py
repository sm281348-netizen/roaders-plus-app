# -*- coding: utf-8 -*-
"""
金旭 PMS 雙館每日定時同步腳本 (Daily Scheduled Sync Runner)
==============================================================
執行方式：
  python run_daily_sync.py [--date YYYY-MM-DD] [--days N]

說明：
1. 每日清晨 (例如 05:00) 於 Windows 工作排程器中自動無聲執行。
2. 自動解析【站前館】與【主題館】之夜審結算報表 (EIS03 + RVR12)。
3. 將當日及歷史最新數據儲存於本地備份庫 (data/pms_backup/)。
4. 輸出執行日誌至 logs/pms_sync.log，方便維運追蹤。
"""

import os
import sys
import datetime
import logging
import argparse
import pandas as pd

# 解決 Windows 主控台編碼問題
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# 導入解析引擎
from pms_parser import extract_pms_day, extract_pms_range, to_date_obj

# 設定工作目錄
APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(APP_DIR, "logs")
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

# 設定日誌
log_file = os.path.join(LOG_DIR, "pms_sync.log")
logger = logging.getLogger("pms_sync")
logger.setLevel(logging.INFO)
logger.handlers.clear()

file_h = logging.FileHandler(log_file, encoding="utf-8")
file_h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(file_h)

stream_h = logging.StreamHandler(sys.stdout)
stream_h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(stream_h)

def run_sync_for_hotel(hotel_name: str, target_date: datetime.date, days: int = 1):
    logger.info(f"========== 開始同步【{hotel_name}】數據 (目標日期: {target_date}, 範圍: {days} 天) ==========")
    start_date = target_date - datetime.timedelta(days=days - 1)

    try:
        res = extract_pms_range(hotel_name, start_date, target_date)
        df_occ = res["df_occ"]
        df_chan = res["df_chan"]
        df_eis = res["df_eis"]

        # 1. 儲存至本地備份庫 (確保零資料遺失)
        h_tag = "station" if hotel_name == "站前館" else "theme"
        
        if not df_occ.empty:
            occ_backup_path = os.path.join(BACKUP_DIR, f"{h_tag}_occ_data.csv")
            if os.path.exists(occ_backup_path):
                df_old_occ = pd.read_csv(occ_backup_path)
                df_occ['date'] = df_occ['date'].astype(str)
                df_old_occ['date'] = df_old_occ['date'].astype(str)
                df_final_occ = df_occ.set_index('date').combine_first(df_old_occ.set_index('date')).reset_index()
            else:
                df_final_occ = df_occ
            df_final_occ.sort_values('date').to_csv(occ_backup_path, index=False, encoding='utf-8-sig')
            logger.info(f"[SUCCESS] [{hotel_name}] occ_data 解析成功: 共 {len(df_occ)} 天，已更新至本地備份庫")
        else:
            logger.warning(f"[WARN] [{hotel_name}] 未能解析到 occ_data (檢查報表是否存在)")

        if not df_chan.empty:
            chan_backup_path = os.path.join(BACKUP_DIR, f"{h_tag}_channels.csv")
            if os.path.exists(chan_backup_path):
                df_old_chan = pd.read_csv(chan_backup_path)
                new_dates = set(df_chan['date'].astype(str).unique())
                df_old_clean = df_old_chan[~df_old_chan['date'].astype(str).isin(new_dates)]
                df_final_chan = pd.concat([df_old_clean, df_chan], ignore_index=True)
            else:
                df_final_chan = df_chan
            df_final_chan.sort_values('date').to_csv(chan_backup_path, index=False, encoding='utf-8-sig')
            logger.info(f"[SUCCESS] [{hotel_name}] marketing_channel_data 解析成功: 共 {len(df_chan)} 筆，已更新至本地備份庫")
        else:
            logger.warning(f"[WARN] [{hotel_name}] 未能解析到渠道明細 (檢查 RVR12 是否存在)")

        if not df_eis.empty:
            eis_backup_path = os.path.join(BACKUP_DIR, f"{h_tag}_eis_{target_date.strftime('%Y%m')}.csv")
            df_eis.to_csv(eis_backup_path, index=False, encoding='utf-8-sig')
            logger.info(f"[SUCCESS] [{hotel_name}] EIS_data 會計科目解析成功: 共 {len(df_eis)} 項，已儲存至備份庫")

        return True
    except Exception as e:
        logger.error(f"[FAIL] [{hotel_name}] 同步過程發生異常: {e}", exc_info=True)
        return False

def main():
    parser = argparse.ArgumentParser(description="金旭 PMS 雙館每日定時同步腳本")
    parser.add_argument("--date", default=None, help="目標日期 (YYYY-MM-DD)，預設為昨日")
    parser.add_argument("--days", type=int, default=1, help="回推天數，預設 1 天")
    parser.add_argument("--hotel", default="all", choices=["站前館", "主題館", "all"], help="指定館別")

    args = parser.parse_args()

    target_date = to_date_obj(args.date) if args.date else datetime.date.today() - datetime.timedelta(days=1)
    hotels = ["站前館", "主題館"] if args.hotel == "all" else [args.hotel]

    logger.info(f"啟動金旭 PMS 數據同步排程作業...")
    for h in hotels:
        run_sync_for_hotel(h, target_date, days=args.days)
    logger.info(f"雙館金旭數據同步作業完成。\n")

if __name__ == "__main__":
    main()
