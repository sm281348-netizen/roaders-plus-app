# -*- coding: utf-8 -*-
"""
金旭 PMS 與 Google Sheets 資料庫同步引擎 (PMS Sync Engine)
============================================================
負責將 pms_parser 解析出的數據安全、冪等地寫入 Google Sheets：
- occ_data: 以 'date' 為主鍵進行 combine_first 覆蓋更新
- marketing_channel_data: 以 'date' 區間進行覆蓋更新，避免重複堆疊
- EIS_data: 以 '年/月' 與 '項目' 進行會計科目更新
"""

import datetime
import pandas as pd
from pms_parser import extract_pms_day, extract_pms_range, to_date_obj

def sync_occ_data(conn, hotel_name: str, new_occ_rows: list) -> int:
    """
    同步住房營運數據至 Google Sheets occ_data
    """
    if not new_occ_rows:
        return 0

    df_new = pd.DataFrame(new_occ_rows)
    if 'date' not in df_new.columns:
        return 0

    try:
        df_existing = conn.read(worksheet="occ_data", ttl=0)
    except Exception as e:
        print(f"讀取現有 occ_data 失敗: {e}")
        df_existing = pd.DataFrame()

    if df_existing is not None and not df_existing.empty:
        df_existing.columns = [str(c).strip().lower() for c in df_existing.columns]
        # 移除非法或空欄位
        df_existing = df_existing.loc[:, ~df_existing.columns.duplicated()]
        if 'date' in df_existing.columns:
            df_existing['date'] = df_existing['date'].astype(str).str.strip()
            df_new['date'] = df_new['date'].astype(str).str.strip()
            
            # 以 date 為 index 進行 combine_first 冪等合併 (新資料優先覆蓋舊資料)
            df_new_idx = df_new.set_index('date')
            df_exist_idx = df_existing.set_index('date')
            df_merged = df_new_idx.combine_first(df_exist_idx).reset_index()
        else:
            df_merged = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_merged = df_new

    # 排序與清理
    if 'date' in df_merged.columns:
        df_merged = df_merged.sort_values('date').reset_index(drop=True)

    conn.update(worksheet="occ_data", data=df_merged.fillna(""))
    return len(new_occ_rows)

def sync_channel_data(conn, hotel_name: str, df_new_channels: pd.DataFrame) -> int:
    """
    同步 OTA 訂房通路數據至 Google Sheets marketing_channel_data
    """
    if df_new_channels is None or df_new_channels.empty:
        return 0

    try:
        df_existing = conn.read(worksheet="marketing_channel_data", ttl=0)
    except Exception as e:
        print(f"讀取現有 marketing_channel_data 失敗: {e}")
        df_existing = pd.DataFrame()

    new_dates = set(df_new_channels['date'].astype(str).str.strip().unique())

    if df_existing is not None and not df_existing.empty:
        df_existing.columns = [str(c).strip().lower() for c in df_existing.columns]
        if 'date' in df_existing.columns:
            df_existing['date'] = df_existing['date'].astype(str).str.strip()
            # 移除新日期已存在的舊紀錄，避免重複累加
            df_clean_exist = df_existing[~df_existing['date'].isin(new_dates)]
            df_merged = pd.concat([df_clean_exist, df_new_channels], ignore_index=True)
        else:
            df_merged = pd.concat([df_existing, df_new_channels], ignore_index=True)
    else:
        df_merged = df_new_channels

    if 'date' in df_merged.columns:
        df_merged = df_merged.sort_values('date').reset_index(drop=True)

    conn.update(worksheet="marketing_channel_data", data=df_merged.fillna(""))
    return len(df_new_channels)

def sync_eis_data(conn, hotel_name: str, df_new_eis: pd.DataFrame) -> int:
    """
    同步會計收入科目數據至 Google Sheets EIS_data
    """
    if df_new_eis is None or df_new_eis.empty:
        return 0

    try:
        df_existing = conn.read(worksheet="EIS_data", ttl=0)
    except Exception as e:
        print(f"讀取現有 EIS_data 失敗: {e}")
        df_existing = pd.DataFrame()

    if df_existing is not None and not df_existing.empty:
        # 清理既有欄位
        col_ym = next((c for c in df_existing.columns if '年' in c and '月' in c), '年/月')
        col_item = next((c for c in df_existing.columns if '項目' in c), '項目')
        col_actual = next((c for c in df_existing.columns if '本月' in c and '實際' in c), '本月實際')

        df_existing = df_existing.copy()
        df_existing[col_ym] = df_existing[col_ym].astype(str).str.strip()
        df_existing[col_item] = df_existing[col_item].astype(str).str.strip()

        # 針對當月更新或附加
        target_ym = df_new_eis['年/月'].iloc[0]
        # 移除當月的舊科目紀錄，全量換成當月最新審定值
        df_clean = df_existing[df_existing[col_ym] != target_ym]
        
        # 確保新資料欄位名稱對齊
        df_to_add = df_new_eis[['年/月', '項目', '本月實際']].copy()
        df_merged = pd.concat([df_clean, df_to_add], ignore_index=True)
    else:
        df_merged = df_new_eis[['年/月', '項目', '本月實際']]

    conn.update(worksheet="EIS_data", data=df_merged.fillna(""))
    return len(df_new_eis)

def sync_pms_to_gsheets(conn, hotel_name: str, start_date=None, end_date=None) -> dict:
    """
    高階整合同步函數：
    抓取指定日期範圍之金旭數據並完整同步至雙館 Google Sheets。
    """
    if not end_date:
        end_date = datetime.date.today() - datetime.timedelta(days=1)
    if not start_date:
        start_date = end_date

    res = extract_pms_range(hotel_name, start_date, end_date)

    occ_count = 0
    chan_count = 0
    eis_count = 0

    if not res["df_occ"].empty:
        occ_count = sync_occ_data(conn, hotel_name, res["df_occ"].to_dict(orient="records"))

    if not res["df_chan"].empty:
        chan_count = sync_channel_data(conn, hotel_name, res["df_chan"])

    if not res["df_eis"].empty:
        eis_count = sync_eis_data(conn, hotel_name, res["df_eis"])

    return {
        "hotel": hotel_name,
        "date_range": f"{start_date} ~ {end_date}",
        "occ_synced": occ_count,
        "channels_synced": chan_count,
        "eis_synced": eis_count
    }
