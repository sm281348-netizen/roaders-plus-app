# -*- coding: utf-8 -*-
"""
金旭 PMS 未來在手預訂 (OTB) 全量抽取引擎 (Future OTB Extractor)
=============================================================
功能：
1. 直連金旭 PMS 核心 SQL Server 資料庫 (HT2000)。
2. 從 HTAC (每日入帳排程) 與 HTRV (預約主檔) 抽取指定日期區間 (今日 ~ 2027-12-31) 之每日在手訂單。
3. 關聯 HTRB (房況檔) 動態扣除每日維修/鎖房 (OOO)，精準計算可售房數 (Available Rooms)。
4. 建立「全量日曆脊柱 (Full Calendar Spine)」，確保所有未來日期完整鋪滿，被取消或無訂單之日自動歸零。
5. 輸出符合 occ_data 規範的 DataFrame:
   ['date', 'sold_rooms', 'total_rooms', 'occ_rate', 'revenue', 'adr']
"""

import datetime
import logging
import pyodbc
import pandas as pd

logger = logging.getLogger("pms_sync")

# 金旭 SQL Server 連線設定 (站前館)
SQL_SERVER_CONFIG = {
    "DRIVER": "{SQL Server}",
    "SERVER": r"192.168.10.160\SQL2019,1999",
    "DATABASE": "HT2000",
    "UID": "guisa",
    "PWD": "7669588",
    "TIMEOUT": 10
}

def get_sql_connection():
    """建立金旭 PMS SQL Server 資料庫連線"""
    conn_str = (
        f"DRIVER={SQL_SERVER_CONFIG['DRIVER']};"
        f"SERVER={SQL_SERVER_CONFIG['SERVER']};"
        f"DATABASE={SQL_SERVER_CONFIG['DATABASE']};"
        f"UID={SQL_SERVER_CONFIG['UID']};"
        f"PWD={SQL_SERVER_CONFIG['PWD']};"
    )
    return pyodbc.connect(conn_str, timeout=SQL_SERVER_CONFIG["TIMEOUT"])

def extract_future_pms_occ(
    hotel_name: str = "站前館",
    start_date: datetime.date = None,
    end_date: datetime.date = datetime.date(2027, 12, 31),
    default_total_rooms: int = 128
) -> pd.DataFrame:
    """
    從金旭 PMS 資料庫抽取自 start_date 至 end_date 的每日預約售房與房租數據。
    
    :param hotel_name: 館別名稱 ('站前館' 或 '主題館')
    :param start_date: 起始日期 (預設為今日)
    :param end_date: 結束日期 (預設為 2027-12-31)
    :param default_total_rooms: 預設總房間數 (站前館 128 間)
    :return: 包含完整未來日曆的 DataFrame
    """
    if hotel_name != "站前館":
        logger.info(f"[{hotel_name}] 目前未來預約抽取僅支援站前館 (HT2000)")
        return pd.DataFrame(columns=['date', 'sold_rooms', 'total_rooms', 'occ_rate', 'revenue', 'adr'])

    if start_date is None:
        start_date = datetime.date.today()

    start_str = start_date.strftime("%Y%m%d")
    end_str = end_date.strftime("%Y%m%d")

    logger.info(f"開始抽取【{hotel_name}】未來預約數據: {start_date} 至 {end_date} (共 {(end_date - start_date).days + 1} 天)...")

    try:
        conn = get_sql_connection()
        cur = conn.cursor()

        # 1. 查詢有效預約的每日房數與房租 (排除 CXNL 取消訂單)
        sql_bookings = """
        SELECT 
            a.AC02 as date_str,
            SUM(a.AC06) as sold_rooms,
            SUM(a.AC08) as revenue
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        WHERE a.AC02 >= ? AND a.AC02 <= ?
          AND v.RV06 = 'ACTV'
        GROUP BY a.AC02
        """
        cur.execute(sql_bookings, (start_str, end_str))
        bookings_map = {r[0]: (float(r[1]), float(r[2])) for r in cur.fetchall()}

        # 2. 查詢各日維修鎖房數 (OOO)
        sql_ooo = """
        SELECT 
            RB03 as date_str,
            COUNT(*) as ooo_count
        FROM HTRB
        WHERE RB03 >= ? AND RB03 <= ?
          AND RB05 = 'OOO'
        GROUP BY RB03
        """
        cur.execute(sql_ooo, (start_str, end_str))
        ooo_map = {r[0]: int(r[1]) for r in cur.fetchall()}

        conn.close()
        logger.info(f"SQL 查詢完成: 取得 {len(bookings_map)} 天在手訂單，{len(ooo_map)} 天維修房設定")

    except Exception as e:
        logger.error(f"連線或查詢金旭資料庫失敗: {e}", exc_info=True)
        return pd.DataFrame(columns=['date', 'sold_rooms', 'total_rooms', 'occ_rate', 'revenue', 'adr'])

    # 3. 構建全量日曆脊柱 (Full Calendar Spine)
    records = []
    curr = start_date
    while curr <= end_date:
        d_ymd = curr.strftime("%Y%m%d")
        d_iso = curr.strftime("%Y-%m-%d")

        # 動態計算可售房數 (128 - OOO，若無維修紀錄預設為 2 間維持 126 間標準可售)
        ooo_cnt = ooo_map.get(d_ymd, 2)
        total_rooms = max(1, default_total_rooms - ooo_cnt)

        # 售房數與營收
        sold_raw, rev_raw = bookings_map.get(d_ymd, (0.0, 0.0))
        sold = int(round(sold_raw))
        rev = int(round(rev_raw))

        adr = int(round(rev / sold)) if sold > 0 else 0
        occ = round(sold / total_rooms * 100, 2) if total_rooms > 0 else 0.0

        records.append({
            "date": d_iso,
            "sold_rooms": sold,
            "total_rooms": total_rooms,
            "occ_rate": occ,
            "revenue": rev,
            "adr": adr
        })
        curr += datetime.timedelta(days=1)

    df_res = pd.DataFrame(records)
    logger.info(f"【{hotel_name}】未來預約抽取完成，共產出 {len(df_res)} 筆日曆記錄")
    return df_res

if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    df = extract_future_pms_occ("站前館", datetime.date(2026, 11, 2), datetime.date(2027, 12, 31))
    print(df.head(10))
    print(df.tail(5))
