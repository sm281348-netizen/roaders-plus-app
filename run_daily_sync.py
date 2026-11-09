# -*- coding: utf-8 -*-
"""
金旭 PMS 雙館每日定時同步腳本 (Daily Scheduled Sync Runner)
==============================================================
執行方式：
  python run_daily_sync.py [--date YYYY-MM-DD] [--days N] [--git-push]

說明：
1. 每日清晨 (例如 05:00) 於 Windows 工作排程器中自動無聲執行。
2. 自動解析【站前館】與【主題館】之夜審結算報表 (EIS03 + RVR12)。
3. 將當日及歷史最新數據儲存於本地備份庫 (data/pms_backup/)。
4. 支援 --git-push 自動提交並推播至 GitHub 倉庫，讓 Streamlit Cloud 零時差獲取最新數據。
5. 輸出執行日誌至 logs/pms_sync.log，方便維運追蹤。
"""

import os
import sys
import datetime
import logging
import argparse
import subprocess
import pandas as pd

# 解決 Windows 主控台編碼問題
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# 導入解析引擎
from pms_parser import extract_pms_day, extract_pms_range, to_date_obj
from pms_future_extractor import extract_future_pms_occ
from fb_purchase_sync import run_fb_purchase_sync
from pms_fnb_extractor import run_fnb_scheduled_sync

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

try:
    file_h = logging.FileHandler(log_file, encoding="utf-8")
    file_h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(file_h)
except Exception:
    pass

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

        # 1. 儲存歷史數據至本地備份庫 (確保零資料遺失)
        h_tag = "station" if hotel_name == "站前館" else "theme"
        occ_backup_path = os.path.join(BACKUP_DIR, f"{h_tag}_occ_data.csv")
        
        if not df_occ.empty:
            if os.path.exists(occ_backup_path):
                df_old_occ = pd.read_csv(occ_backup_path)
                df_occ['date'] = df_occ['date'].astype(str)
                df_old_occ['date'] = df_old_occ['date'].astype(str)
                df_final_occ = df_occ.set_index('date').combine_first(df_old_occ.set_index('date')).reset_index()
            else:
                df_final_occ = df_occ
            logger.info(f"[SUCCESS] [{hotel_name}] 歷史 occ_data (EIS03) 解析成功: 共 {len(df_occ)} 天")
        else:
            if os.path.exists(occ_backup_path):
                df_final_occ = pd.read_csv(occ_backup_path)
                df_final_occ['date'] = df_final_occ['date'].astype(str)
            else:
                df_final_occ = pd.DataFrame(columns=['date', 'sold_rooms', 'total_rooms', 'occ_rate', 'revenue', 'adr'])
            logger.warning(f"[WARN] [{hotel_name}] 本次未解析到新 EIS03，沿用現有備份庫")

        # 2. 抽取未來預約數據 (今天 ~ 2027-12-31 全量鏡像覆蓋)
        future_start_date = target_date + datetime.timedelta(days=1)
        future_start_str = future_start_date.strftime("%Y-%m-%d")
        try:
            df_future = extract_future_pms_occ(
                hotel_name=hotel_name,
                start_date=future_start_date,
                end_date=datetime.date(2027, 12, 31)
            )
            if not df_future.empty:
                df_future['date'] = df_future['date'].astype(str)
                # 歷史部分 (<= target_date 結算封帳): 來自 EIS03 審定日報
                # 未來部分 (> target_date): 100% 以今晨最新預約快照全量覆蓋 (鏡像覆蓋，避免退訂殘留)
                df_hist_part = df_final_occ[df_final_occ['date'].astype(str) < future_start_str]
                df_final_occ = pd.concat([df_hist_part, df_future], ignore_index=True).sort_values('date').reset_index(drop=True)
                logger.info(f"[SUCCESS] [{hotel_name}] 未來在手預約全量刷新成功 (共 {len(df_future)} 天，覆蓋至 2027-12-31)")
        except Exception as e_future:
            logger.error(f"[ERROR] [{hotel_name}] 未來預約抽取發生異常: {e_future}", exc_info=True)

        if not df_final_occ.empty:
            df_final_occ.sort_values('date').to_csv(occ_backup_path, index=False, encoding='utf-8-sig')
            logger.info(f"[SUCCESS] [{hotel_name}] occ_data 已儲存至本地備份庫: 共 {len(df_final_occ)} 天 (時間軸: {df_final_occ['date'].min()} ~ {df_final_occ['date'].max()})")

        if not df_chan.empty:
            df_chan.columns = [str(c).strip().lower() for c in df_chan.columns]
            chan_backup_path = os.path.join(BACKUP_DIR, f"{h_tag}_channels.csv")
            if os.path.exists(chan_backup_path):
                df_old_chan = pd.read_csv(chan_backup_path)
                df_old_chan.columns = [str(c).strip().lower() for c in df_old_chan.columns]
                new_dates = set(df_chan['date'].astype(str).unique())
                df_old_clean = df_old_chan[~df_old_chan['date'].astype(str).isin(new_dates)]
                df_final_chan = pd.concat([df_old_clean, df_chan], ignore_index=True)
            else:
                df_final_chan = df_chan
            df_final_chan.columns = [str(c).strip().lower() for c in df_final_chan.columns]
            df_final_chan = df_final_chan.loc[:, ~df_final_chan.columns.duplicated()]
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

def push_to_cloud_git():
    """
    將 data/pms_backup/ 的最新數據自動提交並推播至 GitHub main 分支
    """
    logger.info("檢查是否需要推播數據至 GitHub 雲端倉庫...")
    try:
        # 1. 暫存 data/pms_backup/ 與同步引擎腳本
        subprocess.run([
            "git", "add", 
            "data/pms_backup/", 
            "pms_future_extractor.py", 
            "run_daily_sync.py",
            "pms_channel_sync.py",
            "fb_purchase_extractor.py",
            "fb_purchase_sync.py",
            "pms_fnb_extractor.py",
            "sync_supplier_prices.py",
            "hk_purchase_extractor.py",
            "hk_purchase_sync.py"
        ], cwd=APP_DIR, check=True)
        
        # 2. 檢查有無異動
        diff_res = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=APP_DIR)
        if diff_res.returncode != 0:
            today_str = datetime.date.today().strftime('%Y-%m-%d')
            commit_msg = f"chore(sync): auto-sync PMS & FB purchase data {today_str} [skip ci]"
            logger.info(f"偵測到新數據，正在提交變更: {commit_msg}")
            subprocess.run(["git", "commit", "-m", commit_msg], cwd=APP_DIR, check=True)
            
            logger.info("正在推播至遠端 GitHub (main 分支)...")
            push_res = subprocess.run(
                ["git", "push", "origin", "main"],
                cwd=APP_DIR,
                capture_output=True,
                text=True,
                encoding="utf-8"
            )
            if push_res.returncode == 0:
                logger.info("[SUCCESS] 數據已成功推播至 GitHub！Streamlit Cloud 將自動同步呈現最新數據。")
            else:
                logger.error(f"[FAIL] Git 推播失敗: {push_res.stderr}")
        else:
            logger.info("[INFO] PMS 與餐飲數據庫無新異動，無需推播。")
    except Exception as e:
        logger.error(f"[FAIL] 執行 Git 雲端推播時發生例外: {e}")

def main():
    parser = argparse.ArgumentParser(description="金旭 PMS 雙館每日定時同步腳本")
    parser.add_argument("--date", default=None, help="目標日期 (YYYY-MM-DD)，預設為昨日")
    parser.add_argument("--days", type=int, default=1, help="回推天數，預設 1 天")
    parser.add_argument("--hotel", default="all", choices=["站前館", "主題館", "all"], help="指定館別")
    parser.add_argument("--git-push", action="store_true", help="是否自動推播最新數據至 GitHub 雲端倉庫")

    args = parser.parse_args()

    target_date = to_date_obj(args.date) if args.date else datetime.date.today() - datetime.timedelta(days=1)
    hotels = ["站前館", "主題館"] if args.hotel == "all" else [args.hotel]

    logger.info(f"啟動金旭 PMS 數據同步排程作業 (目標: {target_date}, 範圍: {args.days} 天, 雲端推播: {args.git_push})...")
    for h in hotels:
        run_sync_for_hotel(h, target_date, days=args.days)
    logger.info(f"雙館金旭數據解析與本地庫存檔作業完成。")

    # 執行雙館 F&B 附加服務 (f&b_data 及 f&b_report 預估) 全自動同步
    try:
        run_fnb_scheduled_sync()
    except Exception as e_fnb:
        logger.error(f"[FAIL] F&B 附加服務同步發生例外: {e_fnb}", exc_info=True)

    # 執行餐飲請購單 (The Peak 與 4F 餐車) 自動比對與雲端同步
    try:
        run_fb_purchase_sync(target_date=datetime.date.today())
    except Exception as e_fb:
        logger.error(f"[FAIL] 餐飲請購單同步發生例外: {e_fb}", exc_info=True)

    # 執行房務部請購單全量與增量自動比對與雲端同步
    try:
        from hk_purchase_sync import sync_hk_purchases
        sync_hk_purchases()
    except Exception as e_hk:
        logger.error(f"[FAIL] 房務請購單同步發生例外: {e_hk}", exc_info=True)

    if args.git_push:
        push_to_cloud_git()
        
    logger.info(f"所有排程任務執行完畢。\n")

if __name__ == "__main__":
    main()
