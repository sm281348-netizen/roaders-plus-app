# -*- coding: utf-8 -*-
"""
金旭 PMS 雙館 F&B 附加服務 (早餐/下午茶) 全自動抽取與同步引擎 (PMS F&B Extractor & Sync)
================================================================================
功能：
1. 直連金旭 PMS 雙館 SQL Server 資料庫 (站前館與主題館)。
2. 從 HTHP (附加服務檔) 與 HTRV (預約主檔) 抽取有效在手訂單之早餐與下午茶。
3. 自動依據入住日與退房日展開為「每日餐期日曆明細」：
   - 早餐 (BF): 入住隔日清晨至退房當日 (ci_d + 1 ~ co_d)
   - 下午茶 (AT): 入住日至退房前一日 (ci_d ~ co_d - 1)
4. 生成符合 Google Sheets f&b_data 規範的 12 欄標準資料集：
   ['訂房狀況', '訂房碼', '房號', '訂房名稱/旅客姓名', '公司名稱', '服務事項', '服務日期', '服務時間(起)', '服務時間(迄)', '數量', '說明事項', '金額']
5. 支援「滑動窗口鏡像同步 (Rolling Window Sync)」：保護 30 天前歷史紀錄不變，動態覆蓋更新近 14 天至未來的在手預約。
6. 自動計算未來 14 天每日各餐期預估人次，精準單格回填至 f&b_report (不觸碰任何實際來客數欄位)。
"""

import os
import sys
import datetime
import logging
import pyodbc
import pandas as pd
import toml
import gspread

logger = logging.getLogger("pms_sync")

# 金旭 SQL Server 連線設定 (雙館)
HOTEL_SQL_CONFIGS = {
    "站前館": {
        "driver": "{SQL Server}",
        "server": r"192.168.10.160\SQL2019,1999",
        "database": "HT2000",
        "uid": "guisa",
        "pwd": "7669588",
        "timeout": 10,
        "gsheet_conn": "gsheets_station"
    },
    "主題館": {
        "driver": "{SQL Server}",
        "server": "192.168.10.160",
        "database": "HT2000",
        "uid": "guisa",
        "pwd": "7669588",
        "timeout": 10,
        "gsheet_conn": "gsheets_theme"
    }
}

# 渠道代碼美化對照表
CHANNEL_NAME_MAP = {
    'TBB': 'TBB 【BookingButton】BOOKINGBUTTOM',
    'AGODA': 'AGODA 【AGODA預付】Agoda Company Pte. Ltd.',
    'Expedia02': 'Expedia02 【EXPEDIA到付】EXPEDIA',
    'Expedia': 'Expedia 【EXPEDIA到付】EXPEDIA',
    'Booking': 'Booking 【Booking.com】Booking.com',
    'Ctrip': 'Ctrip 【攜程網預付】Ctrip',
    'Trip': 'Trip.com 【Trip.com】Trip.com',
    'ezTravel': '易遊網 【易遊網】ezTravel',
    'FIT': 'FIT 【散客/臨櫃】Walk-in'
}


def get_sql_connection(hotel_name: str):
    """建立指定館別之金旭 SQL Server 連線"""
    cfg = HOTEL_SQL_CONFIGS.get(hotel_name)
    if not cfg:
        raise ValueError(f"未知的館別: {hotel_name}")
    conn_str = (
        f"DRIVER={cfg['driver']};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['uid']};"
        f"PWD={cfg['pwd']};"
    )
    return pyodbc.connect(conn_str, timeout=cfg["timeout"])


def extract_fnb_daily_details(hotel_name: str, start_co_date: str = "20260101") -> pd.DataFrame:
    """
    從金旭 SQL Server 抽取附加服務並展開為每日餐期明細。
    
    :param hotel_name: '站前館' 或 '主題館'
    :param start_co_date: 最早退房日期 (預設 20260101)
    :return: 包含 12 欄之 DataFrame
    """
    logger.info(f"[{hotel_name}] 開始從金旭資料庫抽取附加服務 (退房日 >= {start_co_date})...")
    conn = get_sql_connection(hotel_name)
    cur = conn.cursor()

    sql = """
    SELECT 
        ISNULL(v.RV07, '') as status,
        ISNULL(v.RV01, '') as r_code,
        ISNULL(v.RV08, '') as room_no,
        ISNULL(v.RV11, '') as guest_name,
        ISNULL(v.RV10, '') as company_code,
        p.HP01 as service_code,
        ISNULL(p.HP11, '') as service_name,
        v.RV03 as ci_date,
        v.RV04 as co_date,
        p.HP06 as qty,
        ISNULL(p.HP12, 0) as amt
    FROM HTHP p
    JOIN HTRV v ON p.HP05 = v.RV01
    WHERE p.HP01 IN ('BF', 'AT')
      AND v.RV06 != 'CXNL'
      AND v.RV07 != 'CX'
      AND v.RV04 >= ?
    ORDER BY v.RV03, v.RV01
    """
    cur.execute(sql, (start_co_date,))
    rows = cur.fetchall()
    conn.close()

    logger.info(f"[{hotel_name}] SQL 查詢成功: 取得 {len(rows)} 筆附加服務主檔記錄")

    expanded_records = []
    for r in rows:
        status, r_code, room_no, guest_name, comp_code, s_code, s_name, ci_str, co_str, qty, amt = r
        try:
            ci_d = datetime.datetime.strptime(ci_str.strip(), "%Y%m%d").date()
            co_d = datetime.datetime.strptime(co_str.strip(), "%Y%m%d").date()
        except Exception:
            continue

        # 房號清洗：若非純數字 (如 NRM) 則轉為空字串
        r_room = room_no.strip() if room_no.strip().isdigit() else ''

        # 渠道名稱格式化
        c_code_clean = comp_code.strip()
        comp_display = CHANNEL_NAME_MAP.get(c_code_clean, c_code_clean)

        # 數量安全轉換
        try:
            qty_int = int(float(qty or 0))
        except Exception:
            qty_int = 0

        if qty_int <= 0:
            continue

        # 決定服務展開區間
        if s_code == 'BF':  # 早餐
            s_label = '早餐'
            curr_d = ci_d + datetime.timedelta(days=1)
            end_d = co_d
            if ci_d == co_d:  # 同日進退
                curr_d = ci_d
                end_d = ci_d
        else:  # 下午茶
            s_label = '下午茶'
            curr_d = ci_d
            end_d = co_d - datetime.timedelta(days=1)
            if ci_d == co_d:
                curr_d = ci_d
                end_d = ci_d

        while curr_d <= end_d:
            d_str = curr_d.strftime("%Y%m%d")
            expanded_records.append({
                '訂房狀況': status.strip(),
                '訂房碼': r_code.strip(),
                '房號': r_room,
                '訂房名稱/旅客姓名': guest_name.strip(),
                '公司名稱': comp_display,
                '服務事項': s_label,
                '服務日期': d_str,
                '服務時間(起)': '',
                '服務時間(迄)': '',
                '數量': qty_int,
                '說明事項': s_label,
                '金額': '0'
            })
            curr_d += datetime.timedelta(days=1)

    df_out = pd.DataFrame(expanded_records)
    if not df_out.empty:
        df_out = df_out.sort_values(by=['服務日期', '房號', '訂房碼', '服務事項']).reset_index(drop=True)
    logger.info(f"[{hotel_name}] 餐期日曆展開完成: 共 {len(df_out)} 筆每日服務明細")
    return df_out


def get_gspread_client():
    """獲取 gspread 服務帳戶客戶端"""
    app_dir = os.path.dirname(os.path.abspath(__file__))
    sec_path = os.path.join(app_dir, ".streamlit", "secrets.toml")
    with open(sec_path, encoding='utf-8') as f:
        sec = toml.load(f)
    cfg = sec['connections']['gsheets_station']
    sa_info = {
        k: cfg[k] for k in [
            'type', 'project_id', 'private_key_id', 'private_key',
            'client_email', 'client_id', 'auth_uri', 'token_uri',
            'auth_provider_x509_cert_url', 'client_x509_cert_url'
        ]
    }
    return gspread.service_account_from_dict(sa_info), sec


def sync_hotel_fnb_data(hotel_name: str, cutoff_days: int = 14) -> pd.DataFrame:
    """
    同步指定館別的 f&b_data 至 Google Sheets (採用滑動窗口鏡像同步)。
    
    :param hotel_name: '站前館' 或 '主題館'
    :param cutoff_days: 滑動窗口天數 (預設今日前 14 天)
    :return: 最終合併後的完整 DataFrame
    """
    cfg = HOTEL_SQL_CONFIGS.get(hotel_name)
    conn_key = cfg['gsheet_conn']

    today = datetime.date.today()
    cutoff_date = today - datetime.timedelta(days=cutoff_days)
    cutoff_str = cutoff_date.strftime("%Y%m%d")

    gc, sec = get_gspread_client()
    sh_url = sec['connections'][conn_key]['spreadsheet']
    sh = gc.open_by_url(sh_url)
    ws = sh.worksheet('f&b_data')

    # 1. 讀取現有資料
    existing_rows = ws.get_all_values()
    headers = ['訂房狀況', '訂房碼', '房號', '訂房名稱/旅客姓名', '公司名稱', '服務事項', '服務日期', '服務時間(起)', '服務時間(迄)', '數量', '說明事項', '金額']

    preserved_rows = []
    if len(existing_rows) > 1:
        for r in existing_rows[1:]:
            if len(r) > 6:
                d = r[6].strip().replace('.0', '')
                if len(d) == 8 and d.isdigit() and d < cutoff_str:
                    # 保留截止日之前的歷史記錄，補齊 12 欄
                    row_full = r[:12] + [''] * max(0, 12 - len(r))
                    preserved_rows.append(row_full)

    logger.info(f"[{hotel_name}] [f&b_data] 保留歷史存檔 (< {cutoff_str}): 共 {len(preserved_rows)} 筆")

    # 2. 從 SQL 抽取最新預約 (涵蓋 cutoff_date 至今及未來所有在手)
    df_new = extract_fnb_daily_details(hotel_name, start_co_date=cutoff_str)
    # 只取服務日期 >= cutoff_str 的記錄
    df_new_active = df_new[df_new['服務日期'] >= cutoff_str].copy()
    logger.info(f"[{hotel_name}] [f&b_data] 抽取最新在手預約 (>= {cutoff_str}): 共 {len(df_new_active)} 筆")

    # 3. 合併資料並進行嚴格防呆排序與去重
    df_pres = pd.DataFrame(preserved_rows, columns=headers) if preserved_rows else pd.DataFrame(columns=headers)
    df_combined = pd.concat([df_pres, df_new_active[headers]], ignore_index=True)
    
    if not df_combined.empty:
        # 清理服務日期
        df_combined['服務日期'] = df_combined['服務日期'].astype(str).str.strip().str.replace('.0', '', regex=False)
        # 去重：若同訂單在同日期同事項重複出現，優先保留有實際房號與 CO 已結帳狀態的記錄
        df_combined['has_room'] = df_combined['房號'].str.strip() != ''
        df_combined['is_co'] = df_combined['訂房狀況'].str.strip().str.startswith('CO')
        
        df_sorted = df_combined.sort_values(
            by=['服務日期', 'is_co', 'has_room', '訂房碼'], 
            ascending=[True, False, False, True]
        )
        df_dedup = df_sorted.drop_duplicates(
            subset=['訂房碼', '服務日期', '服務事項', '訂房名稱/旅客姓名'], 
            keep='first'
        ).copy()
        df_dedup = df_dedup.drop(columns=['has_room', 'is_co'])
        
        # 嚴格按「服務日期 -> 房號 -> 訂房碼 -> 服務事項」由小到大嚴格遞增排序
        df_final = df_dedup.sort_values(by=['服務日期', '房號', '訂房碼', '服務事項']).reset_index(drop=True)
        final_rows = [headers] + df_final[headers].astype(str).values.tolist()
    else:
        final_rows = [headers]

    # 4. 寫回 Google Sheet
    ws.clear()
    ws.update(values=final_rows, range_name="A1")
    logger.info(f"[SUCCESS] [{hotel_name}] [f&b_data] Google Sheets 同步成功！總筆數: {len(final_rows) - 1} 筆 (嚴格按服務日期遞增排序，已消除錯亂與重複)")

    return df_new


def sync_fb_forecast_to_report(df_st_fnb: pd.DataFrame, df_th_fnb: pd.DataFrame, forecast_days: int = 14):
    """
    將雙館未來的預估來客數自動回填至 f&b_report 的預估欄位 (絕不覆寫實際欄位)。
    """
    logger.info(f"開始核算雙館前 7 天與未來 {forecast_days} 天之餐期預估人數並回填 f&b_report...")
    today = datetime.date.today()
    target_dates = [(today + datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(-7, forecast_days + 1)]

    # 聚合雙館預估數
    def get_forecast_map(df):
        f_map = {}
        if df is not None and not df.empty and '服務日期' in df.columns:
            for _, r in df.iterrows():
                raw_d = str(r['服務日期']).replace('.0', '').strip()
                if len(raw_d) == 8 and raw_d.isdigit():
                    d_iso = f"{raw_d[:4]}-{raw_d[4:6]}-{raw_d[6:]}"
                else:
                    d_iso = raw_d
                s_item = str(r['服務事項']).strip()
                try:
                    q = int(float(r['數量'] or 0))
                except Exception:
                    q = 0
                if d_iso not in f_map:
                    f_map[d_iso] = {'早餐': 0, '下午茶': 0}
                if s_item in f_map[d_iso]:
                    f_map[d_iso][s_item] += q
        return f_map

    st_map = get_forecast_map(df_st_fnb)
    th_map = get_forecast_map(df_th_fnb)

    gc, sec = get_gspread_client()
    sh_url = sec['connections']['gsheets_station']['spreadsheet']
    sh = gc.open_by_url(sh_url)
    ws = sh.worksheet('f&b_report')

    all_rows = ws.get_all_values()
    date_to_row_idx = {}
    for i, r in enumerate(all_rows):
        if r and len(r) > 0:
            d_val = r[0].strip().replace('/', '-')
            if len(d_val) == 10 and d_val[4] == '-' and d_val[7] == '-':
                date_to_row_idx[d_val] = i + 1  # 1-indexed

    updates = []
    for d_iso in target_dates:
        row_num = date_to_row_idx.get(d_iso)
        if not row_num:
            # 日期不在現有表格中，追加新列
            ws.append_row([d_iso, '', '', '', '', '', '', '', '', ''])
            all_rows.append([d_iso, '', '', '', '', '', '', '', '', ''])
            row_num = len(all_rows)
            date_to_row_idx[d_iso] = row_num

        th_bf = th_map.get(d_iso, {}).get('早餐', 0)
        th_af = th_map.get(d_iso, {}).get('下午茶', 0)
        st_bf = st_map.get(d_iso, {}).get('早餐', 0)
        st_af = st_map.get(d_iso, {}).get('下午茶', 0)

        # 批次更新儲存格: B (主題早), D (站前早), F (主題午), H (站前午)
        curr_row = all_rows[row_num - 1]
        orig_th_bf = curr_row[1] if len(curr_row) > 1 else ''
        orig_st_bf = curr_row[3] if len(curr_row) > 3 else ''
        orig_th_af = curr_row[5] if len(curr_row) > 5 else ''
        orig_st_af = curr_row[7] if len(curr_row) > 7 else ''

        val_th_bf = str(th_bf) if th_bf > 0 else (orig_th_bf if orig_th_bf else '0')
        val_st_bf = str(st_bf) if st_bf > 0 else (orig_st_bf if orig_st_bf else '0')
        val_th_af = str(th_af) if th_af > 0 else (orig_th_af if orig_th_af else '0')
        val_st_af = str(st_af) if st_af > 0 else (orig_st_af if orig_st_af else '0')

        updates.append({'range': f'B{row_num}', 'values': [[val_th_bf]]})
        updates.append({'range': f'D{row_num}', 'values': [[val_st_bf]]})
        updates.append({'range': f'F{row_num}', 'values': [[val_th_af]]})
        updates.append({'range': f'H{row_num}', 'values': [[val_st_af]]})

    if updates:
        ws.batch_update(updates)
        logger.info(f"[SUCCESS] [f&b_report] 未來 {len(target_dates)} 天之預估來客數已回填至 Google Sheets！(共更新 {len(updates)} 個儲存格)")


def run_fnb_scheduled_sync():
    """每日 05:00 排程主執行入口"""
    logger.info("========== 開始執行雙館 F&B 附加服務全自動同步作業 ==========")
    try:
        # 1. 同步站前館 f&b_data
        df_st = sync_hotel_fnb_data("站前館", cutoff_days=14)
        # 2. 同步主題館 f&b_data
        df_th = sync_hotel_fnb_data("主題館", cutoff_days=14)
        # 3. 自動回填未來的預估來客數至 f&b_report
        sync_fb_forecast_to_report(df_st, df_th, forecast_days=14)
        logger.info("========== 雙館 F&B 數據與預估排程同步全數完成 ==========")
        return True
    except Exception as e:
        logger.error(f"[ERROR] F&B 附加服務同步失敗: {e}", exc_info=True)
        return False


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    run_fnb_scheduled_sync()
