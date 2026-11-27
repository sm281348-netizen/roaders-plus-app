# -*- coding: utf-8 -*-
"""
金旭 PMS (HT2000) 收益情報與通路淨利全自動唯讀抽取引擎 (PMS Revenue Intelligence Sync)
====================================================================================
100% 唯讀 SELECT 查詢（零寫入），供每日 05:00 run_daily_sync.py 全自動執行並推送雲端。

輸出快照（存於 data/pms_backup/）：
1. pms_channel_intel.json
   - 結合 HTRV + HTAC + HTGR，精準區分各 OTA「預付 (Net)」與「到付 (Gross)」口徑。
   - 提供 MTD / 近30天 / 近90天 / 逐月（訂房日 Booked Date 與 入住日 Stay Date 雙視角）之：
     淨訂單數、取消筆數、No-Show、取消率、淨房晚、營收、原始 ADR、ALOS、提前預訂天數 (Lead Time)。
   - 未來 60 天「每日高取消風險通路佔比（Wash Risk）」供房務、餐飲與收益部預警。
2. pms_stly_pace.json
   - 未來 90 天每日在手 OTB vs 去年同期 STLY Pace（同時提供「同星期 -364 天 [SiteMinder 基準]」與「同日期 -365 天」）。
   - 運用 HTAC.AC07（取消房數保留欄）與 HTRV.RV001/RV005 精確還原去年今日在手訂單，並同步計算 1 日與 7 日淨增減（Net Pickup）。
3. 同步觸發 wallpaper_planner.refresh_pms_roomtype_inventory() 刷新 12 房型每日剩餘庫存。
"""

import os
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("OMP_NUM_THREADS", "1")
import json
import datetime
import logging
from collections import defaultdict

from pms_future_extractor import get_sql_connection

logger = logging.getLogger("pms_sync")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")
CHANNEL_INTEL_PATH = os.path.join(BACKUP_DIR, "pms_channel_intel.json")
STLY_PACE_PATH = os.path.join(BACKUP_DIR, "pms_stly_pace.json")
PROMO_BF_PATH = os.path.join(BACKUP_DIR, "pms_promo_bf_audit.json")
DYNAMIC_YIELD_PATH = os.path.join(BACKUP_DIR, "pms_dynamic_yield.json")
SM_SNAPSHOT_PATH = os.path.join(BACKUP_DIR, "siteminder_inventory_snapshot.json")
SM_HISTORY_DIR = os.path.join(BACKUP_DIR, "sm_history")
RT_INV_PATH = os.path.join(BACKUP_DIR, "wallpaper_roomtype_inventory.json")

# SiteMinder 標準含早加價（依實測 Distribution 衍生規則：每人淨加價 $400 ÷ 0.9）
SM_BF_ADDON_BY_RT = {
    "DF": (1778, 4), "F": (1778, 4),
    "DTP": (1333, 3),
    "LDT": (889, 2), "DT": (889, 2), "E": (889, 2), "ER": (889, 2), "PD": (889, 2), "PT": (889, 2),
    "SP": (0, 2), "S": (0, 2), "SD": (0, 2),
}


def classify_channel(rv10_code: str, gr03_name: str, gr07_pay: str):
    """
    將金旭 HTRV.RV10 與 HTGR.GR03 標準化為：
    (channel_group, sub_label, rate_basis, default_comm_pct, is_ota)
    rate_basis:
      - 'net_prepaid' : OTA 預付（HTAC.AC08 已為扣抽成後淨房價，預設再扣抽成 0%）
      - 'gross_pay_at_hotel' : OTA 到付（HTAC.AC08 為旅客刷卡毛價，需扣合約抽成）
      - 'direct' : 官網直訂 / 散客（免 OTA 抽成，僅計約 2% 金流手續費）
      - 'corporate' : 企業合約 / 票券 / 集團專案
    """
    code = (rv10_code or "").strip()
    cu = code.upper()
    name = (gr03_name or "").strip()
    nu = name.upper()
    pay = (gr07_pay or "").strip().upper()

    if cu in ("CTRIP01", "CTRIP02") or "CTRIP" in nu or "TRIP.COM" in nu:
        is_prepaid = (cu == "CTRIP01") or ("預付" in name) or (pay == "CARD")
        return (
            "Trip.com",
            "Trip.com 預付 (Net)" if is_prepaid else "Trip.com 到付 (Gross)",
            "net_prepaid" if is_prepaid else "gross_pay_at_hotel",
            0.0 if is_prepaid else 15.0,
            True,
        )
    if cu == "BOOKING" or "BOOKING.COM" in nu:
        return ("Booking.com", "Booking.com 到付 (Gross)", "gross_pay_at_hotel", 15.0, True)
    if cu in ("4251987901", "4251987902") or "AGODA" in nu:
        is_prepaid = (cu == "4251987901") or ("預付" in name) or (pay == "CARD")
        return (
            "Agoda",
            "Agoda 預付 (Net)" if is_prepaid else "Agoda 到付 (Gross)",
            "net_prepaid" if is_prepaid else "gross_pay_at_hotel",
            0.0 if is_prepaid else 15.0,
            True,
        )
    if cu == "TBB" or "BOOKINGBUTTON" in nu or "BOOKINGBUTTOM" in nu:
        return ("官網 Direct (TBB)", "官網直訂 (Direct)", "direct", 2.0, False)
    if cu in ("EXPEDIA01", "EXPEDIA02") or "EXPEDIA" in nu:
        is_prepaid = (cu == "EXPEDIA01") or ("預付" in name) or (pay == "CARD")
        return (
            "Expedia",
            "Expedia 預付 (Net)" if is_prepaid else "Expedia 到付 (Gross)",
            "net_prepaid" if is_prepaid else "gross_pay_at_hotel",
            0.0 if is_prepaid else 15.0,
            True,
        )
    if cu in ("9030352601", "2898392501") or "RAKUTEN" in nu or "樂天" in name:
        return ("Rakuten 樂天", "Rakuten 到付 (Gross)", "gross_pay_at_hotel", 12.0, True)
    if cu == "82883867" or "ASIAYO" in nu:
        return ("AsiaYo", "AsiaYo 預付 (Net)", "net_prepaid", 0.0, True)
    if cu == "24763920" or "FUNNOW" in nu:
        return ("FunNow", "FunNow 預付 (Net)", "net_prepaid", 0.0, True)
    if cu == "":
        return ("現場/電話散客", "現場/電話散客", "direct", 2.0, False)
    if cu in ("FG", "85099708", "TAIWANSTAY") or "小鹿" in name or "國旅" in name:
        return ("小鹿/票券/國旅", name or code, "corporate", 0.0, False)
    return ("其他企業/旅行社", name or code, "corporate", 0.0, False)


def _fetch_channel_window(cur, date_mode: str, start_ymd: str, end_ymd: str):
    """
    唯讀查詢指定日期區間之各通路績效（精確區分 HTRV 筆數與 HTAC 房晚，避免 JOIN 膨脹 ALOS/Lead）。
    date_mode: 'booked' (以 RV001 下訂日) 或 'stay' (以 AC02/RV03 入住日)
    """
    if date_mode == "booked":
        sql_rv = """
        SELECT
            LTRIM(RTRIM(ISNULL(v.RV10, ''))) AS code,
            ISNULL(LTRIM(RTRIM(g.GR03)), '') AS gr_name,
            ISNULL(LTRIM(RTRIM(g.GR07)), '') AS pay_type,
            COUNT(CASE WHEN v.RV06 = 'ACTV' THEN 1 END) AS net_bk,
            COUNT(CASE WHEN v.RV06 = 'CXNL' THEN 1 END) AS cx_bk,
            COUNT(CASE WHEN v.RV06 = 'NOSH' THEN 1 END) AS nosh_bk,
            AVG(CASE WHEN v.RV06 = 'ACTV' AND LEN(LTRIM(RTRIM(v.RV03)))=8 AND LEN(LTRIM(RTRIM(v.RV04)))=8
                     AND ISDATE(v.RV03)=1 AND ISDATE(v.RV04)=1
                THEN CAST(DATEDIFF(day, CONVERT(date, v.RV03, 112), CONVERT(date, v.RV04, 112)) AS FLOAT) END) AS alos,
            AVG(CASE WHEN v.RV06 = 'ACTV' AND LEN(LTRIM(RTRIM(v.RV001)))=8 AND LEN(LTRIM(RTRIM(v.RV03)))=8
                     AND ISDATE(v.RV001)=1 AND ISDATE(v.RV03)=1
                THEN CAST(DATEDIFF(day, CONVERT(date, v.RV001, 112), CONVERT(date, v.RV03, 112)) AS FLOAT) END) AS lead_d
        FROM HTRV v
        LEFT JOIN HTGR g ON LTRIM(RTRIM(v.RV10)) = LTRIM(RTRIM(g.GR01))
        WHERE v.RV001 >= ? AND v.RV001 <= ?
        GROUP BY LTRIM(RTRIM(ISNULL(v.RV10, ''))), ISNULL(LTRIM(RTRIM(g.GR03)), ''), ISNULL(LTRIM(RTRIM(g.GR07)), '')
        """
        sql_ac = """
        SELECT
            LTRIM(RTRIM(ISNULL(v.RV10, ''))) AS code,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC06 ELSE 0 END) AS net_rn,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC08 ELSE 0 END) AS net_rev,
            SUM(CASE WHEN v.RV06 IN ('CXNL','NOSH') THEN a.AC07 ELSE 0 END) AS cx_rn,
            SUM(CASE WHEN v.RV06 IN ('CXNL','NOSH') THEN a.AC08 ELSE 0 END) AS cx_rev
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        WHERE v.RV001 >= ? AND v.RV001 <= ?
        GROUP BY LTRIM(RTRIM(ISNULL(v.RV10, '')))
        """
    else:
        sql_rv = """
        SELECT
            LTRIM(RTRIM(ISNULL(v.RV10, ''))) AS code,
            ISNULL(LTRIM(RTRIM(g.GR03)), '') AS gr_name,
            ISNULL(LTRIM(RTRIM(g.GR07)), '') AS pay_type,
            COUNT(CASE WHEN v.RV06 = 'ACTV' THEN 1 END) AS net_bk,
            COUNT(CASE WHEN v.RV06 = 'CXNL' THEN 1 END) AS cx_bk,
            COUNT(CASE WHEN v.RV06 = 'NOSH' THEN 1 END) AS nosh_bk,
            AVG(CASE WHEN v.RV06 = 'ACTV' AND LEN(LTRIM(RTRIM(v.RV03)))=8 AND LEN(LTRIM(RTRIM(v.RV04)))=8
                     AND ISDATE(v.RV03)=1 AND ISDATE(v.RV04)=1
                THEN CAST(DATEDIFF(day, CONVERT(date, v.RV03, 112), CONVERT(date, v.RV04, 112)) AS FLOAT) END) AS alos,
            AVG(CASE WHEN v.RV06 = 'ACTV' AND LEN(LTRIM(RTRIM(v.RV001)))=8 AND LEN(LTRIM(RTRIM(v.RV03)))=8
                     AND ISDATE(v.RV001)=1 AND ISDATE(v.RV03)=1
                THEN CAST(DATEDIFF(day, CONVERT(date, v.RV001, 112), CONVERT(date, v.RV03, 112)) AS FLOAT) END) AS lead_d
        FROM HTRV v
        LEFT JOIN HTGR g ON LTRIM(RTRIM(v.RV10)) = LTRIM(RTRIM(g.GR01))
        WHERE v.RV03 >= ? AND v.RV03 <= ?
        GROUP BY LTRIM(RTRIM(ISNULL(v.RV10, ''))), ISNULL(LTRIM(RTRIM(g.GR03)), ''), ISNULL(LTRIM(RTRIM(g.GR07)), '')
        """
        sql_ac = """
        SELECT
            LTRIM(RTRIM(ISNULL(v.RV10, ''))) AS code,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC06 ELSE 0 END) AS net_rn,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC08 ELSE 0 END) AS net_rev,
            SUM(CASE WHEN v.RV06 IN ('CXNL','NOSH') THEN a.AC07 ELSE 0 END) AS cx_rn,
            SUM(CASE WHEN v.RV06 IN ('CXNL','NOSH') THEN a.AC08 ELSE 0 END) AS cx_rev
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        WHERE a.AC02 >= ? AND a.AC02 <= ?
        GROUP BY LTRIM(RTRIM(ISNULL(v.RV10, '')))
        """

    cur.execute(sql_rv, (start_ymd, end_ymd))
    rv_map = {}
    for code, gr_name, pay_type, nb, cx, nosh, alos, lead in cur.fetchall():
        rv_map[code] = {
            "code": code,
            "gr_name": gr_name,
            "pay_type": pay_type,
            "net_bk": int(nb or 0),
            "cx_bk": int(cx or 0),
            "nosh_bk": int(nosh or 0),
            "alos": round(float(alos or 0.0), 2),
            "lead_d": round(float(lead or 0.0), 1),
        }

    cur.execute(sql_ac, (start_ymd, end_ymd))
    ac_map = {}
    for code, net_rn, net_rev, cx_rn, cx_rev in cur.fetchall():
        ac_map[code] = {
            "net_rn": int(round(float(net_rn or 0))),
            "net_rev": int(round(float(net_rev or 0))),
            "cx_rn": int(round(float(cx_rn or 0))),
            "cx_rev": int(round(float(cx_rev or 0))),
        }

    all_codes = sorted(set(rv_map.keys()) | set(ac_map.keys()))
    rows = []
    for code in all_codes:
        r = rv_map.get(code, {"gr_name": "", "pay_type": "", "net_bk": 0, "cx_bk": 0, "nosh_bk": 0, "alos": 0.0, "lead_d": 0.0})
        a = ac_map.get(code, {"net_rn": 0, "net_rev": 0, "cx_rn": 0, "cx_rev": 0})
        if r["net_bk"] == 0 and r["cx_bk"] == 0 and a["net_rn"] == 0:
            continue
        grp, sub_lbl, basis, def_comm, is_ota = classify_channel(code, r["gr_name"], r["pay_type"])
        tot_bk = r["net_bk"] + r["cx_bk"] + r["nosh_bk"]
        cx_rate = round((r["cx_bk"] + r["nosh_bk"]) / tot_bk * 100.0, 1) if tot_bk > 0 else 0.0
        raw_adr = int(round(a["net_rev"] / a["net_rn"])) if a["net_rn"] > 0 else 0
        rows.append({
            "code": code or "(空白/現場)",
            "gr_name": r["gr_name"],
            "channel": grp,
            "sub_label": sub_lbl,
            "rate_basis": basis,
            "default_comm_pct": def_comm,
            "is_ota": is_ota,
            "net_bk": r["net_bk"],
            "cx_bk": r["cx_bk"],
            "nosh_bk": r["nosh_bk"],
            "cancel_rate_pct": cx_rate,
            "net_rn": a["net_rn"],
            "net_rev": a["net_rev"],
            "raw_adr": raw_adr,
            "cx_rn": a["cx_rn"],
            "cx_rev": a["cx_rev"],
            "alos": r["alos"],
            "lead_days": r["lead_d"],
        })
    rows.sort(key=lambda x: x["net_rn"], reverse=True)
    return rows


def extract_pms_channel_intel(today: datetime.date = None):
    """抽取站前館各視窗通路明細與未來 60 天每日洗盤風險（100% 唯讀）。"""
    if today is None:
        today = datetime.date.today()

    conn = get_sql_connection()
    cur = conn.cursor()

    mtd_start = today.replace(day=1).strftime("%Y%m%d")
    today_ymd = today.strftime("%Y%m%d")
    d30_start = (today - datetime.timedelta(days=29)).strftime("%Y%m%d")
    d90_start = (today - datetime.timedelta(days=89)).strftime("%Y%m%d")
    ytd_start = today.replace(month=1, day=1).strftime("%Y%m%d")

    # 計算當月與次月月底
    import calendar
    m_end_day = calendar.monthrange(today.year, today.month)[1]
    cur_m_end = today.replace(day=m_end_day).strftime("%Y%m%d")
    nm_y, nm_m = (today.year + 1, 1) if today.month == 12 else (today.year, today.month + 1)
    nm_start = f"{nm_y}{nm_m:02d}01"
    nm_end = f"{nm_y}{nm_m:02d}{calendar.monthrange(nm_y, nm_m)[1]:02d}"

    windows = {
        "booked_mtd": {
            "label": f"本月累計 MTD ({today.strftime('%Y-%m-01')} ~ {today.strftime('%Y-%m-%d')}，訂房日)",
            "mode": "booked",
            "rows": _fetch_channel_window(cur, "booked", mtd_start, today_ymd),
        },
        "booked_30d": {
            "label": f"近 30 天 ({(today - datetime.timedelta(days=29)).strftime('%Y-%m-%d')} ~ {today.strftime('%Y-%m-%d')}，訂房日)",
            "mode": "booked",
            "rows": _fetch_channel_window(cur, "booked", d30_start, today_ymd),
        },
        "booked_90d": {
            "label": f"近 90 天 ({(today - datetime.timedelta(days=89)).strftime('%Y-%m-%d')} ~ {today.strftime('%Y-%m-%d')}，訂房日)",
            "mode": "booked",
            "rows": _fetch_channel_window(cur, "booked", d90_start, today_ymd),
        },
        "stay_cur_month": {
            "label": f"本月入住 ({today.strftime('%Y-%m')} 全月，入住日)",
            "mode": "stay",
            "rows": _fetch_channel_window(cur, "stay", mtd_start, cur_m_end),
        },
        "stay_next_month": {
            "label": f"次月在手 ({nm_y}-{nm_m:02d} 全月，入住日)",
            "mode": "stay",
            "rows": _fetch_channel_window(cur, "stay", nm_start, nm_end),
        },
        "stay_ytd": {
            "label": f"{today.year} 年度累計 YTD ({today.year}-01-01 ~ {today.strftime('%Y-%m-%d')}，入住日)",
            "mode": "stay",
            "rows": _fetch_channel_window(cur, "stay", ytd_start, today_ymd),
        },
    }

    # 未來 60 天每日通路佔比與「高取消率通路洗盤風險 (Wash Risk)」
    fut_end = (today + datetime.timedelta(days=60)).strftime("%Y%m%d")
    cur.execute("""
        SELECT
            LTRIM(RTRIM(a.AC02)) AS stay_date,
            LTRIM(RTRIM(ISNULL(v.RV10, ''))) AS code,
            ISNULL(LTRIM(RTRIM(g.GR03)), '') AS gr_name,
            ISNULL(LTRIM(RTRIM(g.GR07)), '') AS pay_type,
            SUM(a.AC06) AS rn,
            SUM(a.AC08) AS rev
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        LEFT JOIN HTGR g ON LTRIM(RTRIM(v.RV10)) = LTRIM(RTRIM(g.GR01))
        WHERE a.AC02 >= ? AND a.AC02 <= ?
          AND v.RV06 = 'ACTV'
        GROUP BY LTRIM(RTRIM(a.AC02)), LTRIM(RTRIM(ISNULL(v.RV10, ''))),
                 ISNULL(LTRIM(RTRIM(g.GR03)), ''), ISNULL(LTRIM(RTRIM(g.GR07)), '')
    """, (today_ymd, fut_end))

    day_chan = defaultdict(lambda: defaultdict(lambda: {"rn": 0, "rev": 0}))
    for dt, code, gr_name, pay_type, rn, rev in cur.fetchall():
        if len(dt) != 8:
            continue
        iso = f"{dt[:4]}-{dt[4:6]}-{dt[6:8]}"
        grp, _, _, _, _ = classify_channel(code, gr_name, pay_type)
        day_chan[iso][grp]["rn"] += int(round(float(rn or 0)))
        day_chan[iso][grp]["rev"] += int(round(float(rev or 0)))

    conn.close()

    # 取得近 90 天各通路實際取消率作為洗盤加權基準
    cx_rate_by_chan = {}
    grp_bk = defaultdict(lambda: [0, 0])
    for r in windows["booked_90d"]["rows"]:
        grp_bk[r["channel"]][0] += r["net_bk"]
        grp_bk[r["channel"]][1] += r["cx_bk"] + r["nosh_bk"]
    for ch, (nb, cx) in grp_bk.items():
        cx_rate_by_chan[ch] = round(cx / (nb + cx) * 100.0, 1) if (nb + cx) > 0 else 0.0

    daily_wash = []
    for i in range(61):
        d = today + datetime.timedelta(days=i)
        iso = d.strftime("%Y-%m-%d")
        cmap = day_chan.get(iso, {})
        tot_rn = sum(v["rn"] for v in cmap.values())
        high_cx_rn = sum(v["rn"] for ch, v in cmap.items() if ch in ("Agoda", "Expedia"))
        direct_rn = cmap.get("官網 Direct (TBB)", {}).get("rn", 0)
        trip_rn = cmap.get("Trip.com", {}).get("rn", 0)
        booking_rn = cmap.get("Booking.com", {}).get("rn", 0)
        agoda_rn = cmap.get("Agoda", {}).get("rn", 0)
        expedia_rn = cmap.get("Expedia", {}).get("rn", 0)
        # 依領先期衰減估算潛在洗盤房數（距離入住越近，已熬過取消期越不會退）
        lead_factor = min(1.0, max(0.15, i / 30.0))
        est_wash_rn = round(
            sum(v["rn"] * (cx_rate_by_chan.get(ch, 15.0) / 100.0) * 0.55 * lead_factor for ch, v in cmap.items()),
            1,
        )
        daily_wash.append({
            "date": iso,
            "tot_rn": tot_rn,
            "trip_rn": trip_rn,
            "booking_rn": booking_rn,
            "agoda_rn": agoda_rn,
            "expedia_rn": expedia_rn,
            "direct_rn": direct_rn,
            "high_cx_rn": high_cx_rn,
            "high_cx_share_pct": round(high_cx_rn / tot_rn * 100.0, 1) if tot_rn > 0 else 0.0,
            "est_wash_rn": est_wash_rn,
        })

    payload = {
        "as_of": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "base_date": today.strftime("%Y-%m-%d"),
        "source": "JinXu HT2000 (HTRV + HTAC + HTGR) Read-Only SQL",
        "channel_90d_cancel_rates": cx_rate_by_chan,
        "windows": windows,
        "daily_wash_60d": daily_wash,
    }
    os.makedirs(BACKUP_DIR, exist_ok=True)
    with open(CHANNEL_INTEL_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    logger.info(f"[SUCCESS] 金旭通路淨利與取消情報已更新 -> {CHANNEL_INTEL_PATH}")
    return payload


def extract_pms_stly_pace(today: datetime.date = None, horizon_days: int = 90):
    """
    從金旭 HT2000 抽取未來 90 天每日 OTB、1日/7日淨 Pickup，以及去年同期 STLY Pace：
      - 同星期比較 (-364 天，與 SiteMinder Pace 預設口徑一致)
      - 同日期比較 (-365 天，方便跨年/節慶固定日期比對)
      - 去年最終結算實績 (Final)
    """
    if today is None:
        today = datetime.date.today()

    conn = get_sql_connection()
    cur = conn.cursor()

    end_d = today + datetime.timedelta(days=horizon_days)
    s_cur = today.strftime("%Y%m%d")
    e_cur = end_d.strftime("%Y%m%d")
    d_1ago = (today - datetime.timedelta(days=1)).strftime("%Y%m%d")
    d_7ago = (today - datetime.timedelta(days=7)).strftime("%Y%m%d")

    # 1. 今年未來每日 OTB (含 1日/7日新增)
    cur.execute("""
        SELECT
            LTRIM(RTRIM(a.AC02)) AS stay_date,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC06 ELSE 0 END) AS otb_rn,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC08 ELSE 0 END) AS otb_rev,
            SUM(CASE WHEN v.RV06 = 'ACTV' AND v.RV001 > ? THEN a.AC06 ELSE 0 END) AS new_1d_rn,
            SUM(CASE WHEN v.RV06 IN ('CXNL','NOSH') AND v.RV005 > ? AND v.RV001 <= ? THEN a.AC07 ELSE 0 END) AS cx_1d_rn,
            SUM(CASE WHEN v.RV06 = 'ACTV' AND v.RV001 > ? THEN a.AC06 ELSE 0 END) AS new_7d_rn,
            SUM(CASE WHEN v.RV06 IN ('CXNL','NOSH') AND v.RV005 > ? AND v.RV001 <= ? THEN a.AC07 ELSE 0 END) AS cx_7d_rn
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        WHERE a.AC02 >= ? AND a.AC02 <= ?
        GROUP BY LTRIM(RTRIM(a.AC02))
    """, (d_1ago, d_1ago, d_1ago, d_7ago, d_7ago, d_7ago, s_cur, e_cur))
    cur_map = {}
    for row in cur.fetchall():
        dt = row[0]
        cur_map[dt] = {
            "rn": int(round(float(row[1] or 0))),
            "rev": int(round(float(row[2] or 0))),
            "p1d": int(round(float(row[3] or 0) - float(row[4] or 0))),
            "p7d": int(round(float(row[5] or 0) - float(row[6] or 0))),
        }

    # 2. 今年每日 OOO 鎖房數
    cur.execute("""
        SELECT LTRIM(RTRIM(RB03)), COUNT(*)
        FROM HTRB
        WHERE RB03 >= ? AND RB03 <= ? AND RB05 = 'OOO'
        GROUP BY LTRIM(RTRIM(RB03))
    """, (s_cur, e_cur))
    ooo_map = {r[0]: int(r[1]) for r in cur.fetchall()}

    # 3. 去年同期 STLY（範圍涵蓋 -365 天至 end_d - 364 天）
    ly_start = (today - datetime.timedelta(days=366)).strftime("%Y%m%d")
    ly_end = (end_d - datetime.timedelta(days=363)).strftime("%Y%m%d")
    asof_ly_dow = (today - datetime.timedelta(days=364)).strftime("%Y%m%d")
    asof_ly_date = (today - datetime.timedelta(days=365)).strftime("%Y%m%d")

    cur.execute("""
        SELECT
            LTRIM(RTRIM(a.AC02)) AS stay_date,
            SUM(CASE WHEN v.RV001 <= ? AND (v.RV06 = 'ACTV' OR (v.RV06 IN ('CXNL','NOSH') AND v.RV005 > ?))
                     THEN (CASE WHEN v.RV06 = 'ACTV' THEN a.AC06 ELSE a.AC07 END) ELSE 0 END) AS stly_dow_rn,
            SUM(CASE WHEN v.RV001 <= ? AND (v.RV06 = 'ACTV' OR (v.RV06 IN ('CXNL','NOSH') AND v.RV005 > ?))
                     THEN a.AC08 ELSE 0 END) AS stly_dow_rev,
            SUM(CASE WHEN v.RV001 <= ? AND (v.RV06 = 'ACTV' OR (v.RV06 IN ('CXNL','NOSH') AND v.RV005 > ?))
                     THEN (CASE WHEN v.RV06 = 'ACTV' THEN a.AC06 ELSE a.AC07 END) ELSE 0 END) AS stly_date_rn,
            SUM(CASE WHEN v.RV001 <= ? AND (v.RV06 = 'ACTV' OR (v.RV06 IN ('CXNL','NOSH') AND v.RV005 > ?))
                     THEN a.AC08 ELSE 0 END) AS stly_date_rev,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC06 ELSE 0 END) AS final_rn,
            SUM(CASE WHEN v.RV06 = 'ACTV' THEN a.AC08 ELSE 0 END) AS final_rev
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        WHERE a.AC02 >= ? AND a.AC02 <= ?
        GROUP BY LTRIM(RTRIM(a.AC02))
    """, (
        asof_ly_dow, asof_ly_dow, asof_ly_dow, asof_ly_dow,
        asof_ly_date, asof_ly_date, asof_ly_date, asof_ly_date,
        ly_start, ly_end,
    ))
    ly_map = {}
    for row in cur.fetchall():
        ly_map[row[0]] = {
            "stly_dow_rn": int(round(float(row[1] or 0))),
            "stly_dow_rev": int(round(float(row[2] or 0))),
            "stly_date_rn": int(round(float(row[3] or 0))),
            "stly_date_rev": int(round(float(row[4] or 0))),
            "final_rn": int(round(float(row[5] or 0))),
            "final_rev": int(round(float(row[6] or 0))),
        }

    conn.close()

    dow_zh = ["一", "二", "三", "四", "五", "六", "日"]
    daily_rows = []
    for i in range(horizon_days + 1):
        d = today + datetime.timedelta(days=i)
        ds = d.strftime("%Y%m%d")
        d_ly_dow = (d - datetime.timedelta(days=364)).strftime("%Y%m%d")
        d_ly_date = (d - datetime.timedelta(days=365)).strftime("%Y%m%d")

        c = cur_map.get(ds, {"rn": 0, "rev": 0, "p1d": 0, "p7d": 0})
        ooo = ooo_map.get(ds, 2)
        cap = max(1, 128 - ooo)
        adr = int(round(c["rev"] / c["rn"])) if c["rn"] > 0 else 0
        occ = round(c["rn"] / cap * 100.0, 1)

        ld = ly_map.get(d_ly_dow, {})
        stly_dow_rn = ld.get("stly_dow_rn", 0)
        stly_dow_rev = ld.get("stly_dow_rev", 0)
        stly_dow_adr = int(round(stly_dow_rev / stly_dow_rn)) if stly_dow_rn > 0 else 0
        ly_dow_final_rn = ld.get("final_rn", 0)
        ly_dow_final_rev = ld.get("final_rev", 0)
        ly_dow_final_adr = int(round(ly_dow_final_rev / ly_dow_final_rn)) if ly_dow_final_rn > 0 else 0

        ldt = ly_map.get(d_ly_date, {})
        stly_date_rn = ldt.get("stly_date_rn", 0)
        stly_date_rev = ldt.get("stly_date_rev", 0)
        stly_date_adr = int(round(stly_date_rev / stly_date_rn)) if stly_date_rn > 0 else 0
        ly_date_final_rn = ldt.get("final_rn", 0)
        ly_date_final_rev = ldt.get("final_rev", 0)
        ly_date_final_adr = int(round(ly_date_final_rev / ly_date_final_rn)) if ly_date_final_rn > 0 else 0

        daily_rows.append({
            "date": d.strftime("%Y-%m-%d"),
            "dow": dow_zh[d.weekday()],
            "is_weekend": d.weekday() in (4, 5),
            "lead_days": i,
            "cap": cap,
            "ooo": ooo,
            "sold_rn": c["rn"],
            "rem_rn": cap - c["rn"],
            "occ_pct": occ,
            "otb_rev": c["rev"],
            "otb_adr": adr,
            "pickup_1d_rn": c["p1d"],
            "pickup_7d_rn": c["p7d"],
            # 同星期 (-364天) STLY
            "stly_dow_date": f"{d_ly_dow[:4]}-{d_ly_dow[4:6]}-{d_ly_dow[6:8]}",
            "stly_dow_rn": stly_dow_rn,
            "stly_dow_rev": stly_dow_rev,
            "stly_dow_adr": stly_dow_adr,
            "yoy_dow_rn": c["rn"] - stly_dow_rn,
            "yoy_dow_adr": adr - stly_dow_adr if (adr > 0 and stly_dow_adr > 0) else 0,
            "ly_dow_final_rn": ly_dow_final_rn,
            "ly_dow_final_adr": ly_dow_final_adr,
            # 同日期 (-365天) STLY
            "stly_date_ref": f"{d_ly_date[:4]}-{d_ly_date[4:6]}-{d_ly_date[6:8]}",
            "stly_date_rn": stly_date_rn,
            "stly_date_rev": stly_date_rev,
            "stly_date_adr": stly_date_adr,
            "yoy_date_rn": c["rn"] - stly_date_rn,
            "yoy_date_adr": adr - stly_date_adr if (adr > 0 and stly_date_adr > 0) else 0,
            "ly_date_final_rn": ly_date_final_rn,
            "ly_date_final_adr": ly_date_final_adr,
        })

    payload = {
        "as_of": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "base_date": today.strftime("%Y-%m-%d"),
        "horizon_days": horizon_days,
        "source": "JinXu HT2000 (HTAC + HTRV + HTRB) Read-Only SQL",
        "daily": daily_rows,
    }
    os.makedirs(BACKUP_DIR, exist_ok=True)
    with open(STLY_PACE_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    logger.info(f"[SUCCESS] 金旭未來 {horizon_days} 天 STLY Pace 與 Pickup 已更新 -> {STLY_PACE_PATH}")
    return payload


def _diagnose_order_promo(code: str, channel: str, rate_basis: str, ratio: float, rt: str):
    """依據金旭實收單價 ÷ SiteMinder 牌告價，白話判讀該筆訂單套用了哪種促銷或折扣。"""
    cu = (code or "").upper().strip()
    if cu == "TAIWANSTAY":
        return "✅ 國旅補助原價 (100%)", False
    if channel == "官網 Direct (TBB)":
        if ratio >= 0.83:
            return "✅ 官網優惠價 (-15%)", False
        if ratio >= 0.78:
            return "✅ 官網連住3晚 (-20%)", False
        if ratio >= 0.73:
            return "✅ 官網連住5晚 (-25%)", False
        return "🔴 官網異常深折 (<73%，請核對優惠碼)", True
    if channel == "Booking.com":
        if ratio >= 0.88:
            return "✅ Booking 原價/9折 (淺折)", False
        if ratio >= 0.83:
            return "✅ Booking 常態85折 (Genius/促銷)", False
        if ratio >= 0.76:
            return "🟠 Booking 深度促銷 (約76~82折)", True
        return "🔴 Booking 超深促銷 (<76折，疑重疊活動)", True
    if channel == "Trip.com":
        if rate_basis == "gross_pay_at_hotel":
            if ratio >= 0.83:
                return "✅ Trip 到付常態85折", False
            return "🟠 Trip 到付額外促銷 (<83折)", True
        # 預付 Net (常態 0.85 × 0.85 = 0.7225；機酒/景酒 0.80 × 0.85 = 0.68)
        if ratio >= 0.71:
            return "✅ Trip 預付常態 (85折×扣15%佣)", False
        if ratio >= 0.665:
            if rt in ("DF", "F", "DT", "E"):
                return "🟠 Trip <機酒/景酒> (-20%×扣佣)", True
            return "🟠 Trip 預付額外促銷 (約78~83折實收)", True
        return "🔴 Trip 預付異常深折 (疑疊加促銷 <66.5%)", True
    if channel == "Agoda":
        if rate_basis == "gross_pay_at_hotel":
            return "✅ Agoda 到付常態" if ratio >= 0.83 else "🟠 Agoda 到付深折", ratio < 0.83
        if ratio >= 0.71:
            return "✅ Agoda 預付常態 (85折×扣15%佣)", False
        if ratio >= 0.66:
            return "🟠 Agoda 預付額外促銷 (約8折×扣佣)", True
        return "🔴 Agoda 預付異常深折 (<66%，疑疊加促銷)", True
    if channel == "Expedia":
        if rate_basis == "net_prepaid":
            if ratio >= 0.78:
                return "✅ Expedia 預付常態 (+5%倍率×9折×扣佣)", False
            if ratio >= 0.71:
                return "🟠 Expedia 促銷活動 (+5%×85折×扣佣)", False
            return "🔴 Expedia 異常深折 (<71%)", True
        return "✅ Expedia 到付" if ratio >= 0.83 else "🟠 Expedia 到付深折", ratio < 0.83
    if "Rakuten" in channel:
        return "✅ 樂天方案 (-10%~-20%)" if ratio >= 0.78 else "🟠 樂天深度促銷", ratio < 0.78
    return f"其他專案 ({ratio*100:.1f}%)", ratio < 0.70


def _norm_orig_rt(rt_code: str) -> str:
    c = (rt_code or "").strip()
    return "DT" if c == "DR" else c


def _hhmmss_to_sec(t8: str):
    s = str(t8 or "").strip().zfill(8)
    if len(s) >= 6 and s[:6].isdigit():
        hh, mm, ss = int(s[:2]), int(s[2:4]), int(s[4:6])
        cc = int(s[6:8]) if len(s) >= 8 and s[6:8].isdigit() else 0
        return hh * 3600 + mm * 60 + ss + cc / 100.0
    return None


def _los_tier(nights: int) -> str:
    n = max(1, int(nights or 1))
    if n == 1:
        return "1N"
    if n == 2:
        return "2N"
    if n <= 4:
        return "3-4N"
    return "5N+"


def extract_pms_promo_and_bf_audit(today: datetime.date = None):
    """
    從金旭 HT2000 + HT2000_Log.HTACL 抽取定價稽核真相（100% 唯讀 SELECT，已剔除現場免費升等汙染）：
      1. 同日期 × 同原始房型「含早 vs 不含早」真實成交差價。
      2. 嚴格區分「同日 4 小時內＋同連住級距」vs「跨日浮動調價觀察」之含早差價明細。
      3. 近 14 天新進訂單實收價 vs 優先對照「下訂當日歸檔 BAR」（若無則對照最新快照）之促銷反查。
    """
    if today is None:
        today = datetime.date.today()

    conn = get_sql_connection()
    cur = conn.cursor()

    s_bf = (today - datetime.timedelta(days=30)).strftime("%Y%m%d")
    e_bf = (today + datetime.timedelta(days=60)).strftime("%Y%m%d")

    # 1. 同日期 × 同原始房型「含早 vs 不含早」真實成交差價（JOIN HTACL 首次 INSERT 排除升等汙染）
    cur.execute("""
        WITH orig AS (
            SELECT
                LTRIM(RTRIM(AC01)) AS rv01,
                LTRIM(RTRIM(AC02)) AS stay_d,
                LTRIM(RTRIM(AC03)) AS orig_rt,
                ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(AC01)), LTRIM(RTRIM(AC02)) ORDER BY DateTime ASC) AS rn
            FROM HT2000_Log.dbo.HTACL
            WHERE Action = 'INSERT' AND AC02 >= ? AND AC02 <= ?
        ),
        per_day_rt AS (
            SELECT
                a.AC02 AS stay_date,
                CASE WHEN ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) = 'DR' THEN 'DT'
                     ELSE ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) END AS rt,
                CASE WHEN a.AC11 LIKE '%不含早%' THEN 'RO'
                     WHEN a.AC11 LIKE '%含早%' THEN 'BF' END AS meal,
                SUM(a.AC06) AS rn,
                SUM(a.AC08) AS rev
            FROM HTAC a
            JOIN HTRV v ON a.AC01 = v.RV01
            LEFT JOIN orig o ON LTRIM(RTRIM(a.AC01)) = o.rv01 AND LTRIM(RTRIM(a.AC02)) = o.stay_d AND o.rn = 1
            WHERE a.AC02 >= ? AND a.AC02 <= ?
              AND v.RV06 = 'ACTV' AND a.AC06 > 0
              AND (a.AC11 LIKE '%不含早%' OR a.AC11 LIKE '%含早%')
            GROUP BY a.AC02,
                     CASE WHEN ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) = 'DR' THEN 'DT'
                          ELSE ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) END,
                     CASE WHEN a.AC11 LIKE '%不含早%' THEN 'RO' WHEN a.AC11 LIKE '%含早%' THEN 'BF' END
        )
        SELECT
            ro.rt,
            COUNT(*) AS matched_days,
            SUM(ro.rn) AS ro_rn,
            SUM(bf.rn) AS bf_rn,
            AVG(ro.rev / ro.rn) AS avg_ro_adr,
            AVG(bf.rev / bf.rn) AS avg_bf_adr,
            AVG((bf.rev / bf.rn) - (ro.rev / ro.rn)) AS avg_diff
        FROM per_day_rt ro
        JOIN per_day_rt bf ON ro.stay_date = bf.stay_date AND ro.rt = bf.rt
        WHERE ro.meal = 'RO' AND bf.meal = 'BF'
        GROUP BY ro.rt
        ORDER BY ro.rt
    """, (s_bf, e_bf, s_bf, e_bf))

    bf_by_rt = []
    for rt, mdays, ro_rn, bf_rn, avg_ro, avg_bf, avg_diff in cur.fetchall():
        sm_addon, pax = SM_BF_ADDON_BY_RT.get(rt, (889, 2))
        sm_90_addon = int(round(sm_addon * 0.9))
        act_diff = int(round(float(avg_diff or 0)))
        per_pax_act = int(round(act_diff / pax)) if pax > 0 else 0
        bf_by_rt.append({
            "rt": rt,
            "pax": pax,
            "matched_days": int(mdays),
            "ro_rn": int(round(float(ro_rn or 0))),
            "bf_rn": int(round(float(bf_rn or 0))),
            "avg_ro_adr": int(round(float(avg_ro or 0))),
            "avg_bf_adr": int(round(float(avg_bf or 0))),
            "sm_shelf_diff": sm_addon,
            "sm_90_diff": sm_90_addon,
            "actual_avg_diff": act_diff,
            "actual_per_pax": per_pax_act,
        })

    # 2. 同日期 × 同原始房型 × 同通路「含早 vs 不含早」明細
    cur.execute("""
        WITH orig AS (
            SELECT
                LTRIM(RTRIM(AC01)) AS rv01,
                LTRIM(RTRIM(AC02)) AS stay_d,
                LTRIM(RTRIM(AC03)) AS orig_rt,
                ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(AC01)), LTRIM(RTRIM(AC02)) ORDER BY DateTime ASC) AS rn
            FROM HT2000_Log.dbo.HTACL
            WHERE Action = 'INSERT' AND AC02 >= ? AND AC02 <= ?
        ),
        per_day_ch AS (
            SELECT
                a.AC02 AS stay_date,
                CASE WHEN ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) = 'DR' THEN 'DT'
                     ELSE ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) END AS rt,
                LTRIM(RTRIM(ISNULL(v.RV10,''))) AS ch_code,
                ISNULL(LTRIM(RTRIM(g.GR03)),'') AS gr_name,
                ISNULL(LTRIM(RTRIM(g.GR07)),'') AS pay_type,
                CASE WHEN a.AC11 LIKE '%不含早%' THEN 'RO' WHEN a.AC11 LIKE '%含早%' THEN 'BF' END AS meal,
                SUM(a.AC06) AS rn,
                SUM(a.AC08) AS rev
            FROM HTAC a
            JOIN HTRV v ON a.AC01 = v.RV01
            LEFT JOIN HTGR g ON LTRIM(RTRIM(v.RV10)) = LTRIM(RTRIM(g.GR01))
            LEFT JOIN orig o ON LTRIM(RTRIM(a.AC01)) = o.rv01 AND LTRIM(RTRIM(a.AC02)) = o.stay_d AND o.rn = 1
            WHERE a.AC02 >= ? AND a.AC02 <= ?
              AND v.RV06 = 'ACTV' AND a.AC06 > 0
              AND (a.AC11 LIKE '%不含早%' OR a.AC11 LIKE '%含早%')
            GROUP BY a.AC02,
                     CASE WHEN ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) = 'DR' THEN 'DT'
                          ELSE ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) END,
                     LTRIM(RTRIM(ISNULL(v.RV10,''))),
                     ISNULL(LTRIM(RTRIM(g.GR03)),''), ISNULL(LTRIM(RTRIM(g.GR07)),''),
                     CASE WHEN a.AC11 LIKE '%不含早%' THEN 'RO' WHEN a.AC11 LIKE '%含早%' THEN 'BF' END
        )
        SELECT
            ro.rt, ro.ch_code, ro.gr_name, ro.pay_type,
            COUNT(*) AS pairs,
            AVG((bf.rev / bf.rn) - (ro.rev / ro.rn)) AS avg_diff,
            MIN((bf.rev / bf.rn) - (ro.rev / ro.rn)) AS min_diff,
            MAX((bf.rev / bf.rn) - (ro.rev / ro.rn)) AS max_diff
        FROM per_day_ch ro
        JOIN per_day_ch bf ON ro.stay_date = bf.stay_date AND ro.rt = bf.rt AND ro.ch_code = bf.ch_code
        WHERE ro.meal = 'RO' AND bf.meal = 'BF'
        GROUP BY ro.rt, ro.ch_code, ro.gr_name, ro.pay_type
        HAVING COUNT(*) >= 2
        ORDER BY ro.rt, pairs DESC
    """, (s_bf, e_bf, s_bf, e_bf))

    bf_by_ch = []
    for rt, code, gname, ptype, pairs, avg_d, min_d, max_d in cur.fetchall():
        grp, sub_lbl, basis, _, _ = classify_channel(code, gname, ptype)
        sm_addon, pax = SM_BF_ADDON_BY_RT.get(rt, (889, 2))
        act_d = int(round(float(avg_d or 0)))
        per_pax = int(round(act_d / pax)) if pax > 0 else 0
        bf_by_ch.append({
            "rt": rt,
            "pax": pax,
            "channel": grp,
            "sub_label": sub_lbl,
            "pairs": int(pairs),
            "sm_shelf_diff": sm_addon,
            "actual_avg_diff": act_d,
            "min_diff": int(round(float(min_d or 0))),
            "max_diff": int(round(float(max_d or 0))),
            "per_pax_bf": per_pax,
            "is_low_bf": per_pax < 300,
        })

    # 3. 含早 vs 不含早實單配對（加入原始下訂房型、精確進單秒數 RV002 與連住晚數 LOS，嚴格區分同日4h內 vs 跨日浮動價差！）
    bf_inv_start = (today - datetime.timedelta(days=30)).strftime("%Y%m%d")
    bf_inv_end = (today + datetime.timedelta(days=60)).strftime("%Y%m%d")
    cur.execute("""
        WITH orig AS (
            SELECT
                LTRIM(RTRIM(AC01)) AS rv01,
                LTRIM(RTRIM(AC02)) AS stay_d,
                LTRIM(RTRIM(AC03)) AS orig_rt,
                ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(AC01)), LTRIM(RTRIM(AC02)) ORDER BY DateTime ASC) AS rn
            FROM HT2000_Log.dbo.HTACL
            WHERE Action = 'INSERT' AND AC02 >= ? AND AC02 <= ?
        ),
        nights AS (
            SELECT
                v.RV01 AS rv_no,
                LTRIM(RTRIM(ISNULL(v.RV29,''))) AS ota_ref,
                v.RV001 AS bk_date,
                RIGHT('00000000' + LTRIM(RTRIM(ISNULL(v.RV002,'0'))), 8) AS bk_time,
                CASE WHEN LEN(LTRIM(RTRIM(v.RV03)))=8 AND LEN(LTRIM(RTRIM(v.RV04)))=8
                          AND ISDATE(v.RV03)=1 AND ISDATE(v.RV04)=1
                     THEN DATEDIFF(day, CONVERT(date, v.RV03, 112), CONVERT(date, v.RV04, 112))
                     ELSE 1 END AS los_nights,
                a.AC02 AS stay_date,
                CASE WHEN ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) = 'DR' THEN 'DT'
                     ELSE ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) END AS rt,
                LTRIM(RTRIM(ISNULL(v.RV10,''))) AS ch_code,
                ISNULL(LTRIM(RTRIM(g.GR03)),'') AS gr_name,
                ISNULL(LTRIM(RTRIM(g.GR07)),'') AS pay_type,
                CASE WHEN a.AC11 LIKE '%不含早%' THEN 'RO' WHEN a.AC11 LIKE '%含早%' THEN 'BF' END AS meal,
                a.AC08 / a.AC06 AS unit_rate
            FROM HTAC a
            JOIN HTRV v ON a.AC01 = v.RV01
            LEFT JOIN HTGR g ON LTRIM(RTRIM(v.RV10)) = LTRIM(RTRIM(g.GR01))
            LEFT JOIN orig o ON LTRIM(RTRIM(a.AC01)) = o.rv01 AND LTRIM(RTRIM(a.AC02)) = o.stay_d AND o.rn = 1
            WHERE a.AC02 >= ? AND a.AC02 <= ?
              AND v.RV06 = 'ACTV' AND a.AC06 > 0 AND a.AC08 > 0
              AND (a.AC11 LIKE '%不含早%' OR a.AC11 LIKE '%含早%')
              AND LEN(LTRIM(RTRIM(v.RV001))) = 8 AND ISDATE(v.RV001) = 1
        )
        SELECT
            bf.stay_date, bf.rt, bf.ch_code, bf.gr_name, bf.pay_type,
            bf.ota_ref AS bf_ota, bf.bk_date AS bf_bk, bf.bk_time AS bf_t, bf.los_nights AS bf_los, bf.unit_rate AS bf_rate,
            ro.ota_ref AS ro_ota, ro.bk_date AS ro_bk, ro.bk_time AS ro_t, ro.los_nights AS ro_los, ro.unit_rate AS ro_rate,
            DATEDIFF(day, CONVERT(date, ro.bk_date, 112), CONVERT(date, bf.bk_date, 112)) AS bk_gap_days,
            (bf.unit_rate - ro.unit_rate) AS diff_twd
        FROM nights bf
        JOIN nights ro
          ON bf.stay_date = ro.stay_date
         AND bf.rt = ro.rt
         AND bf.ch_code = ro.ch_code
         AND bf.meal = 'BF' AND ro.meal = 'RO'
         AND ABS(DATEDIFF(day, CONVERT(date, ro.bk_date, 112), CONVERT(date, bf.bk_date, 112))) <= 7
        WHERE (bf.unit_rate - ro.unit_rate) < 600
        ORDER BY diff_twd ASC, bf.stay_date ASC
    """, (bf_inv_start, bf_inv_end, bf_inv_start, bf_inv_end))
    bf_inv_raw = cur.fetchall()

    # 載入最新 SiteMinder 快照與歷史歸檔快取（供精準比對下訂當日 BAR）
    sm_inv = {}
    if os.path.exists(SM_SNAPSHOT_PATH):
        try:
            with open(SM_SNAPSHOT_PATH, "r", encoding="utf-8") as f:
                sm_inv = json.load(f).get("inventory", {})
        except Exception:
            sm_inv = {}

    sm_hist_cache = {}
    def _get_bar_for_bk_and_stay(bk_iso_str: str, stay_iso_str: str, rt_code: str):
        if bk_iso_str not in sm_hist_cache:
            hpath = os.path.join(SM_HISTORY_DIR, f"{bk_iso_str}.json")
            if os.path.exists(hpath):
                try:
                    with open(hpath, "r", encoding="utf-8") as hf:
                        sm_hist_cache[bk_iso_str] = json.load(hf).get("inventory", {})
                except Exception:
                    sm_hist_cache[bk_iso_str] = None
            else:
                sm_hist_cache[bk_iso_str] = None
        h_inv = sm_hist_cache[bk_iso_str]
        if h_inv:
            r_val = int(h_inv.get(stay_iso_str, {}).get("rooms", {}).get(rt_code, {}).get("r", 0) or 0)
            if r_val > 0:
                return r_val, "下訂當日BAR"
        r_latest = int(sm_inv.get(stay_iso_str, {}).get("rooms", {}).get(rt_code, {}).get("r", 0) or 0)
        return r_latest, "最新快照BAR"

    bf_inversion_alerts = []
    seen_bf_pair = set()
    for (
        st, rt, code, gname, ptype,
        b_ota, b_bk, b_t, b_los, b_rate,
        r_ota, r_bk, r_t, r_los, r_rate,
        gap_d, diff,
    ) in bf_inv_raw:
        sm_addon, pax = SM_BF_ADDON_BY_RT.get(rt, (889, 2))
        b_rate_i = int(round(float(b_rate)))
        r_rate_i = int(round(float(r_rate)))
        diff_i = int(round(float(diff)))
        per_pax = int(round(diff_i / pax)) if pax > 0 else 0
        if diff_i > 0 and per_pax >= 280:
            continue
        pair_key = (st, rt, code, b_ota, r_ota)
        if pair_key in seen_bf_pair:
            continue
        seen_bf_pair.add(pair_key)

        st_iso = f"{st[:4]}-{st[4:6]}-{st[6:8]}"
        b_bk_iso = f"{b_bk[:4]}-{b_bk[4:6]}-{b_bk[6:8]}"
        r_bk_iso = f"{r_bk[:4]}-{r_bk[4:6]}-{r_bk[6:8]}"
        grp, sub_lbl, basis, _, _ = classify_channel(code, gname, ptype)

        gap_d_i = int(gap_d)
        gap_hrs = None
        if gap_d_i == 0:
            sb, sr = _hhmmss_to_sec(b_t), _hhmmss_to_sec(r_t)
            if sb is not None and sr is not None:
                gap_hrs = round(abs(sb - sr) / 3600.0, 2)

        b_los_i, r_los_i = max(1, int(b_los or 1)), max(1, int(r_los or 1))
        same_los = (_los_tier(b_los_i) == _los_tier(r_los_i))
        is_strict = (gap_d_i == 0) and (gap_hrs is not None and gap_hrs <= 4.0) and same_los

        exp_bf_addon = int(round(400 * pax * (0.85 if basis == "net_prepaid" else 1.0)))
        expected_bf_rate = r_rate_i + exp_bf_addon
        bf_vs_expected_pct = round(b_rate_i / expected_bf_rate * 100.0, 1) if expected_bf_rate > 0 else 0.0
        extra_disc_pct = round(100.0 - bf_vs_expected_pct, 1)
        shortfall_vs_400 = (400 * pax) - diff_i

        sm_ro, bar_mode = _get_bar_for_bk_and_stay(b_bk_iso, st_iso, rt)
        sm_bf_pct_str = "—"
        if sm_ro > 0 and sm_addon > 0:
            sm_bf = sm_ro + sm_addon
            sm_bf_pct = round(b_rate_i / sm_bf * 100.0, 1)
            sm_ro_pct = round(r_rate_i / sm_ro * 100.0, 1)
            sm_bf_pct_str = f"含早{sm_bf_pct}% vs 不含早{sm_ro_pct}% ({bar_mode})"

        if is_strict:
            if diff_i <= 0:
                severity = "🔴 同日4h內含早≤不含早"
                reason = f"同日 {gap_hrs:.1f}h 內下訂且同連住級距({b_los_i}晚 vs {r_los_i}晚)，含早反少收 ${abs(diff_i):,}（需確認是否疊加促銷或VIP會員差）"
            else:
                severity = "🟠 同日4h內早餐偏低"
                reason = f"同日 {gap_hrs:.1f}h 內下訂({b_los_i}晚 vs {r_los_i}晚)，每人早餐實收 ${per_pax}/人（可能含會員等級折扣差異）"
        else:
            why_noise = []
            if gap_d_i != 0:
                why_noise.append(f"跨日下訂差{abs(gap_d_i)}天(受浮動調價影響)")
            elif gap_hrs is not None and gap_hrs > 4.0:
                why_noise.append(f"同日間隔{gap_hrs:.1f}h(可能盤中調價)")
            if not same_los:
                why_noise.append(f"連住不同({b_los_i}晚 vs {r_los_i}晚)")
            noise_tag = "、".join(why_noise) if why_noise else "動態價格變動"
            severity = "📊 浮動價差觀察(非鐵證)"
            reason = f"{noise_tag}；兩單帳面差 {diff_i:+,} 元（屬浮動房價或連住折扣軌跡，非設定錯誤鐵證）"

        bf_inversion_alerts.append({
            "severity": severity,
            "is_strict_intraday": is_strict,
            "stay_date": st_iso,
            "rt": rt,
            "pax": pax,
            "channel": grp,
            "sub_label": sub_lbl,
            "bf_rate": b_rate_i,
            "bf_bk_date": b_bk_iso,
            "bf_los": b_los_i,
            "bf_ota_ref": b_ota,
            "ro_rate": r_rate_i,
            "ro_bk_date": r_bk_iso,
            "ro_los": r_los_i,
            "ro_ota_ref": r_ota,
            "bk_gap_days": gap_d_i,
            "bk_gap_hours": gap_hrs,
            "diff_twd": diff_i,
            "per_pax_bf": per_pax,
            "shortfall_vs_400": shortfall_vs_400,
            "extra_disc_pct": extra_disc_pct,
            "sm_pct_compare": sm_bf_pct_str,
            "reason": reason,
        })

    # 4. 近 14 天新進訂單實收價 vs 原始下訂房型 BAR（剔除免費升等造成之假深折！）
    bk_start = (today - datetime.timedelta(days=14)).strftime("%Y%m%d")
    stay_start = today.strftime("%Y%m%d")
    stay_end = (today + datetime.timedelta(days=45)).strftime("%Y%m%d")

    cur.execute("""
        WITH orig AS (
            SELECT
                LTRIM(RTRIM(AC01)) AS rv01,
                LTRIM(RTRIM(AC02)) AS stay_d,
                LTRIM(RTRIM(AC03)) AS orig_rt,
                ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(AC01)), LTRIM(RTRIM(AC02)) ORDER BY DateTime ASC) AS rn
            FROM HT2000_Log.dbo.HTACL
            WHERE Action = 'INSERT' AND AC02 >= ? AND AC02 <= ?
        )
        SELECT
            v.RV01 AS rv_no,
            LTRIM(RTRIM(ISNULL(v.RV29,''))) AS ota_ref,
            v.RV001 AS bk_date,
            a.AC02 AS stay_date,
            CASE WHEN ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) = 'DR' THEN 'DT'
                 ELSE ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) END AS orig_rt,
            LTRIM(RTRIM(a.AC03)) AS cur_rt,
            LTRIM(RTRIM(ISNULL(v.RV10,''))) AS ch_code,
            ISNULL(LTRIM(RTRIM(g.GR03)),'') AS gr_name,
            ISNULL(LTRIM(RTRIM(g.GR07)),'') AS pay_type,
            LTRIM(RTRIM(ISNULL(a.AC11,''))) AS plan_name,
            a.AC06 AS rooms,
            a.AC08 AS rev
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        LEFT JOIN HTGR g ON LTRIM(RTRIM(v.RV10)) = LTRIM(RTRIM(g.GR01))
        LEFT JOIN orig o ON LTRIM(RTRIM(a.AC01)) = o.rv01 AND LTRIM(RTRIM(a.AC02)) = o.stay_d AND o.rn = 1
        WHERE v.RV001 >= ? AND v.RV06 = 'ACTV'
          AND a.AC02 >= ? AND a.AC02 <= ?
          AND a.AC06 > 0 AND a.AC08 > 0
        ORDER BY v.RV001 DESC, a.AC02 ASC
    """, (stay_start, stay_end, bk_start, stay_start, stay_end))

    promo_orders = []
    ratios_all = []
    by_chan_agg = defaultdict(lambda: {"n": 0, "ratios": [], "deep_cnt": 0, "pms_sum": 0, "bar_sum": 0})
    for rv_no, ota_ref, bk_d, st_d, orig_rt, cur_rt, code, gname, ptype, pname, rms, rev in cur.fetchall():
        if len(st_d) != 8 or len(bk_d) != 8:
            continue
        st_iso = f"{st_d[:4]}-{st_d[4:6]}-{st_d[6:8]}"
        bk_iso = f"{bk_d[:4]}-{bk_d[4:6]}-{bk_d[6:8]}"
        sm_ro, bar_mode = _get_bar_for_bk_and_stay(bk_iso, st_iso, orig_rt)
        if sm_ro <= 0:
            continue
        is_bf = ("含早" in pname) and ("不含早" not in pname)
        bf_addon = SM_BF_ADDON_BY_RT.get(orig_rt, (889, 2))[0] if is_bf else 0
        sm_ref = sm_ro + bf_addon
        pms_unit = int(round(float(rev) / float(rms)))
        ratio = round(pms_unit / sm_ref, 4) if sm_ref > 0 else 0.0
        if ratio <= 0.25 or ratio >= 1.8:
            continue

        grp, sub_lbl, basis, _, _ = classify_channel(code, gname, ptype)
        diag, is_deep = _diagnose_order_promo(code, grp, basis, ratio, orig_rt)
        if orig_rt != cur_rt:
            diag = f"{diag} [已升等→{cur_rt}]"

        ratios_all.append(ratio)
        agg = by_chan_agg[sub_lbl]
        agg["n"] += 1
        agg["ratios"].append(ratio)
        agg["pms_sum"] += pms_unit
        agg["bar_sum"] += sm_ref
        if is_deep:
            agg["deep_cnt"] += 1

        promo_orders.append({
            "bk_date": bk_iso,
            "stay_date": st_iso,
            "rt": orig_rt,
            "cur_rt": cur_rt,
            "meal": "含早" if is_bf else "不含早",
            "channel": grp,
            "sub_label": sub_lbl,
            "ota_ref": ota_ref,
            "sm_ref_rate": sm_ref,
            "bar_match_mode": bar_mode,
            "pms_actual_rate": pms_unit,
            "ratio_pct": round(ratio * 100.0, 1),
            "diff_twd": pms_unit - sm_ref,
            "promo_diag": diag,
            "is_deep_promo": is_deep,
        })

    conn.close()

    import statistics
    chan_summary = []
    for sub_lbl, d in sorted(by_chan_agg.items(), key=lambda x: -x[1]["n"]):
        med_r = statistics.median(d["ratios"]) if d["ratios"] else 0.0
        min_r = min(d["ratios"]) if d["ratios"] else 0.0
        max_r = max(d["ratios"]) if d["ratios"] else 0.0
        chan_summary.append({
            "sub_label": sub_lbl,
            "nights": d["n"],
            "avg_sm_ref": int(round(d["bar_sum"] / d["n"])) if d["n"] else 0,
            "avg_pms_actual": int(round(d["pms_sum"] / d["n"])) if d["n"] else 0,
            "med_ratio_pct": round(med_r * 100.0, 1),
            "min_ratio_pct": round(min_r * 100.0, 1),
            "max_ratio_pct": round(max_r * 100.0, 1),
            "deep_promo_nights": d["deep_cnt"],
        })

    emp_ratio = round(sum(ratios_all) / len(ratios_all), 4) if ratios_all else 0.77
    bf_inversion_alerts.sort(key=lambda x: (not x["is_strict_intraday"], x["diff_twd"], x["stay_date"]))
    promo_orders.sort(key=lambda x: (not x["is_deep_promo"], x["ratio_pct"], x["stay_date"]))

    payload = {
        "as_of": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "base_date": today.strftime("%Y-%m-%d"),
        "empirical_actual_to_bar_ratio": emp_ratio,
        "matched_order_nights": len(ratios_all),
        "bf_vs_ro_by_rt": bf_by_rt,
        "bf_vs_ro_by_channel": bf_by_ch,
        "bf_inversion_alerts": bf_inversion_alerts[:150],
        "promo_summary_by_channel": chan_summary,
        "recent_promo_orders": promo_orders[:250],
    }
    os.makedirs(BACKUP_DIR, exist_ok=True)
    with open(PROMO_BF_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    logger.info(f"[SUCCESS] 金旭促銷折數反查與含早差價健檢已更新 ({len(ratios_all)} 筆對照) -> {PROMO_BF_PATH}")
    return payload


def extract_pms_dynamic_yield(today: datetime.date = None, horizon_days: int = 45):
    """
    串接 HT2000 + HT2000_Log.HTACL（100% 唯讀），還原零升等汙染的「動態收益參謀」底層指標：
      1. 近 60 天（前30天+後30天）現場真實免費升等流量矩陣（量化 ER 專用升等池貢獻與零增本 vs 跨床型升等）。
      2. 未來 30 天 12 房型「原始真實買氣 (Original Net Pickup) vs 物理排房移轉」與近期調價方向。
      3. 未來 45 天逐日「低階升等導流 vs 臨期高階清倉」動態診斷（納入床型相容＋前後連貫空房檢核）。
      4. 未來 45 天單日單房型「成交價軌跡 (Booking Curve)」，一眼看懂浮動調價歷程。
    """
    if today is None:
        today = datetime.date.today()

    conn = get_sql_connection()
    cur = conn.cursor()

    s_upg = (today - datetime.timedelta(days=30)).strftime("%Y%m%d")
    e_upg = (today + datetime.timedelta(days=30)).strftime("%Y%m%d")

    # 1. 近 60 天真實現場免費升等／改房型矩陣（排除 2026-03 DR->DT 房型整併）
    cur.execute("""
        WITH orig AS (
            SELECT
                LTRIM(RTRIM(AC01)) AS rv01,
                LTRIM(RTRIM(AC02)) AS stay_d,
                LTRIM(RTRIM(AC03)) AS orig_rt,
                ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(AC01)), LTRIM(RTRIM(AC02)) ORDER BY DateTime ASC) AS rn
            FROM HT2000_Log.dbo.HTACL
            WHERE Action = 'INSERT' AND AC02 >= ? AND AC02 <= ?
        )
        SELECT
            o.orig_rt,
            LTRIM(RTRIM(a.AC03)) AS cur_rt,
            SUM(a.AC06) AS rn_sum,
            COUNT(DISTINCT a.AC01) AS res_cnt,
            AVG(a.AC08 / NULLIF(a.AC06, 0)) AS avg_rate
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        JOIN orig o ON LTRIM(RTRIM(a.AC01)) = o.rv01 AND LTRIM(RTRIM(a.AC02)) = o.stay_d AND o.rn = 1
        WHERE a.AC02 >= ? AND a.AC02 <= ?
          AND v.RV06 = 'ACTV' AND a.AC06 > 0
          AND o.orig_rt <> LTRIM(RTRIM(a.AC03))
          AND NOT (o.orig_rt = 'DR' AND LTRIM(RTRIM(a.AC03)) = 'DT')
        GROUP BY o.orig_rt, LTRIM(RTRIM(a.AC03))
        ORDER BY rn_sum DESC
    """, (s_upg, e_upg, s_upg, e_upg))

    queen_low = {"S", "SP", "PD", "SD"}
    queen_zero_cost_up = {"SP", "PD", "ER", "E", "LDT"}
    multi_bed_high = {"DT", "DTP", "F", "DF"}

    upg_matrix = []
    tot_upg_rn = 0
    zero_cost_rn = 0
    er_abs_rn = 0
    cross_bed_rn = 0

    for ort_raw, crt, rn_s, r_cnt, avg_r in cur.fetchall():
        ort = _norm_orig_rt(ort_raw)
        if ort == crt:
            continue
        rn_i = int(round(float(rn_s or 0)))
        if rn_i <= 0:
            continue
        tot_upg_rn += rn_i
        if crt == "ER":
            er_abs_rn += rn_i
        if ort in queen_low and crt in queen_zero_cost_up:
            cat = "🟢 同床型零增本升等(大床→大床)"
            zero_cost_rn += rn_i
        elif crt in multi_bed_high and ort not in multi_bed_high:
            cat = "🟡 跨床型大容量升等(增清潔/排擠)"
            cross_bed_rn += rn_i
        else:
            cat = "🔵 雙床/特殊調度或平移"

        upg_matrix.append({
            "orig_rt": ort_raw,
            "cur_rt": crt,
            "room_nights": rn_i,
            "orders": int(r_cnt or 0),
            "avg_rate": int(round(float(avg_r or 0))),
            "category": cat,
        })

    # 2. 未來 45 天有效與近期退訂明細（一次拉齊供 Velocity + Diagnostics + Booking Curve 共用）
    s_fut = today.strftime("%Y%m%d")
    e_fut = (today + datetime.timedelta(days=horizon_days)).strftime("%Y%m%d")
    d_7ago = (today - datetime.timedelta(days=7)).strftime("%Y%m%d")
    d_14ago = (today - datetime.timedelta(days=14)).strftime("%Y%m%d")

    cur.execute("""
        WITH orig AS (
            SELECT
                LTRIM(RTRIM(AC01)) AS rv01,
                LTRIM(RTRIM(AC02)) AS stay_d,
                LTRIM(RTRIM(AC03)) AS orig_rt,
                ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(AC01)), LTRIM(RTRIM(AC02)) ORDER BY DateTime ASC) AS rn
            FROM HT2000_Log.dbo.HTACL
            WHERE Action = 'INSERT' AND AC02 >= ? AND AC02 <= ?
        )
        SELECT
            a.AC02 AS stay_d,
            v.RV001 AS bk_d,
            v.RV005 AS mod_d,
            v.RV06 AS status,
            CASE WHEN ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) = 'DR' THEN 'DT'
                 ELSE ISNULL(o.orig_rt, LTRIM(RTRIM(a.AC03))) END AS orig_rt,
            LTRIM(RTRIM(a.AC03)) AS cur_rt,
            CASE WHEN a.AC11 LIKE '%不含早%' THEN 'RO'
                 WHEN a.AC11 LIKE '%含早%' THEN 'BF' ELSE 'RO' END AS meal,
            a.AC06 AS actv_rn,
            a.AC07 AS cx_rn,
            a.AC08 AS rev
        FROM HTAC a
        JOIN HTRV v ON a.AC01 = v.RV01
        LEFT JOIN orig o ON LTRIM(RTRIM(a.AC01)) = o.rv01 AND LTRIM(RTRIM(a.AC02)) = o.stay_d AND o.rn = 1
        WHERE a.AC02 >= ? AND a.AC02 <= ?
          AND (v.RV06 = 'ACTV' OR (v.RV06 IN ('CXNL', 'NOSH') AND v.RV005 > ?))
    """, (s_fut, e_fut, s_fut, e_fut, d_14ago))

    fut_rows = cur.fetchall()
    conn.close()

    # 讀取 12 房型每日剩餘庫存與 STLY Pace 快照
    rt_inv_data = {}
    if os.path.exists(RT_INV_PATH):
        try:
            with open(RT_INV_PATH, "r", encoding="utf-8") as f:
                rt_inv_data = json.load(f)
        except Exception:
            rt_inv_data = {}
    rt_master = rt_inv_data.get("rt_master", {})
    daily_avail = rt_inv_data.get("daily_avail", {})

    pace_map = {}
    if os.path.exists(STLY_PACE_PATH):
        try:
            with open(STLY_PACE_PATH, "r", encoding="utf-8") as f:
                for pr in json.load(f).get("daily", []):
                    pace_map[pr["date"]] = pr
        except Exception:
            pass

    d30_end_ymd = (today + datetime.timedelta(days=29)).strftime("%Y%m%d")
    rt_order_12 = ["DF", "F", "DTP", "LDT", "DT", "E", "ER", "PD", "PT", "SP", "S", "SD"]

    rt_agg = {
        k: {
            "orig_rn": 0, "cur_rn": 0,
            "new7": 0, "cx7": 0, "new14": 0, "cx14": 0,
            "rev_all": 0, "rn_rev_all": 0,
            "rev_7d": 0, "rn_rev_7d": 0,
            "rev_prior": 0, "rn_rev_prior": 0,
        }
        for k in rt_order_12
    }
    day_orig_pk7 = defaultdict(lambda: defaultdict(int))
    curve_grp = defaultdict(lambda: {"rn": 0, "rev": 0, "upg_rn": 0})

    for stay_d, bk_d, mod_d, st_code, ort, crt, meal, actv_rn, cx_rn, rev in fut_rows:
        stay_d = str(stay_d or "").strip()
        bk_d = str(bk_d or "").strip()
        mod_d = str(mod_d or "").strip()
        if len(stay_d) != 8:
            continue
        stay_iso = f"{stay_d[:4]}-{stay_d[4:6]}-{stay_d[6:8]}"

        if st_code == "ACTV":
            rn_i = int(round(float(actv_rn or 0)))
            rev_i = int(round(float(rev or 0)))
            if rn_i <= 0:
                continue

            if stay_d <= d30_end_ymd:
                if ort in rt_agg:
                    rt_agg[ort]["orig_rn"] += rn_i
                    if rev_i > 0:
                        rt_agg[ort]["rev_all"] += rev_i
                        rt_agg[ort]["rn_rev_all"] += rn_i
                    if bk_d > d_7ago:
                        rt_agg[ort]["new7"] += rn_i
                        if rev_i > 0:
                            rt_agg[ort]["rev_7d"] += rev_i
                            rt_agg[ort]["rn_rev_7d"] += rn_i
                    else:
                        if rev_i > 0:
                            rt_agg[ort]["rev_prior"] += rev_i
                            rt_agg[ort]["rn_rev_prior"] += rn_i
                    if bk_d > d_14ago:
                        rt_agg[ort]["new14"] += rn_i
                if crt in rt_agg:
                    rt_agg[crt]["cur_rn"] += rn_i

            if bk_d > d_7ago:
                day_orig_pk7[stay_iso][ort] += rn_i

            if len(bk_d) == 8 and rev_i > 0:
                bk_iso = f"{bk_d[:4]}-{bk_d[4:6]}-{bk_d[6:8]}"
                ck = (stay_iso, ort, bk_iso, meal)
                curve_grp[ck]["rn"] += rn_i
                curve_grp[ck]["rev"] += rev_i
                if ort != crt:
                    curve_grp[ck]["upg_rn"] += rn_i
        else:
            # CXNL / NOSH cancelled within 14d -> subtract from net original pickup if originally booked before window!
            crn_i = int(round(float(cx_rn or 0)))
            if crn_i <= 0:
                continue
            if mod_d > d_7ago and bk_d <= d_7ago:
                day_orig_pk7[stay_iso][ort] -= crn_i
                if stay_d <= d30_end_ymd and ort in rt_agg:
                    rt_agg[ort]["cx7"] += crn_i
            if stay_d <= d30_end_ymd and mod_d > d_14ago and bk_d <= d_14ago and ort in rt_agg:
                rt_agg[ort]["cx14"] += crn_i

    rt_velocity_30d = []
    for k in rt_order_12:
        a = rt_agg[k]
        adr_all = int(round(a["rev_all"] / a["rn_rev_all"])) if a["rn_rev_all"] > 0 else 0
        adr_7d = int(round(a["rev_7d"] / a["rn_rev_7d"])) if a["rn_rev_7d"] > 0 else 0
        adr_prior = int(round(a["rev_prior"] / a["rn_rev_prior"])) if a["rn_rev_prior"] > 0 else 0
        delta_7d = (adr_7d - adr_prior) if (adr_7d > 0 and adr_prior > 0) else 0
        rt_velocity_30d.append({
            "rt": k,
            "tot_inv": int(rt_master.get(k, {}).get("total_inv", 0)),
            "orig_demand_rn": a["orig_rn"],
            "cur_assigned_rn": a["cur_rn"],
            "net_upgrade_shift": a["cur_rn"] - a["orig_rn"],
            "orig_pickup_7d": a["new7"] - a["cx7"],
            "orig_pickup_14d": a["new14"] - a["cx14"],
            "orig_adr_all": adr_all,
            "orig_adr_7d": adr_7d,
            "orig_adr_prior": adr_prior,
            "adr_delta_7d": delta_7d,
        })

    # 3. 整理未來 45 天逐日「低階升等導流 vs 臨期高階清倉」診斷（含前後連貫空房 Continuity 檢核！）
    dow_zh = ["一", "二", "三", "四", "五", "六", "日"]
    iso_list = [(today + datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(horizon_days + 1)]

    def _pool_rem(d_iso: str, keys):
        rts = daily_avail.get(d_iso, {})
        return sum(max(0, int(rts.get(k, {}).get("rem", 0))) for k in keys)

    daily_diag = []
    for idx, d_iso in enumerate(iso_list):
        d_obj = today + datetime.timedelta(days=idx)
        pc = pace_map.get(d_iso, {})
        rts = daily_avail.get(d_iso, {})
        pk_map = day_orig_pk7.get(d_iso, {})

        entry_rem = _pool_rem(d_iso, ("S", "SP", "PD"))
        entry_pk7 = sum(pk_map.get(k, 0) for k in ("S", "SP", "PD"))

        er_rem = max(0, int(rts.get("ER", {}).get("rem", 0)))
        q_upg_rem = _pool_rem(d_iso, ("ER", "E", "LDT"))
        q_upg_pk7 = sum(pk_map.get(k, 0) for k in ("ER", "E", "LDT"))

        # 連住連貫緩衝：檢查 d 與 d+1（若有 d-1 且 idx>0 亦參考 d+1 為主）之最低連貫可升等空房，避免連住客中途換房
        adj_indices = [idx] + ([idx + 1] if idx + 1 < len(iso_list) else [])
        q_upg_cont_rem = min(_pool_rem(iso_list[j], ("ER", "E", "LDT")) for j in adj_indices)
        dt_rem = max(0, int(rts.get("DT", {}).get("rem", 0)))
        dt_cont_rem = min(_pool_rem(iso_list[j], ("DT",)) for j in adj_indices)
        dt_pk7 = pk_map.get("DT", 0)

        pt_rem = max(0, int(rts.get("PT", {}).get("rem", 0)))
        pt_pk7 = pk_map.get("PT", 0)

        fam_rem = _pool_rem(d_iso, ("DF", "F", "DTP"))
        fam_pk7 = sum(pk_map.get(k, 0) for k in ("DF", "F", "DTP"))

        occ_pct = float(pc.get("occ_pct", 0.0))
        rem_rn = int(pc.get("rem_rn", sum(max(0, int(v.get("rem", 0))) for v in rts.values())))

        if 0 <= idx <= 2 and rem_rn >= 14 and (q_upg_rem + dt_rem + fam_rem) >= 10:
            scode = "lastmin_high_clear"
            slabel = "⚡ 臨期高階清倉窗口"
            snote = (
                f"距入住僅 {idx} 天、全館尚餘 {rem_rn} 間，其中高階/大容量剩 {q_upg_rem + dt_rem + fam_rem} 間"
                f"（ER+E+LDT:{q_upg_rem}、DT:{dt_rem}、家庭:{fam_rem}），單靠低階升等已來不及去化，宜確認高階臨期直售水位。"
            )
        elif idx >= 2 and entry_rem <= 3 and entry_pk7 >= 3 and q_upg_cont_rem >= 3 and q_upg_pk7 <= 1:
            scode = "upgrade_capture"
            slabel = "🔄 預排升等導流（免砍高階價）"
            snote = (
                f"主力大床(S+SP+PD)僅剩 {entry_rem} 間且近7日原始淨進單 +{entry_pk7} 晚；"
                f"同床型高階(ER+E+LDT)當日空 {q_upg_rem} 間、連貫2晚空 {q_upg_cont_rem} 間（其中 ER 專用池空 {er_rem} 間）且高階7日僅 {q_upg_pk7:+d} 晚。"
                "建議現場預排 SP/PD 免費升等至 ER/E，騰出低階續賣以守住高階價格錨點！"
            )
        elif occ_pct >= 88.0 and (entry_pk7 + dt_pk7 + fam_pk7) >= 6:
            scode = "hold_and_push"
            slabel = "📈 買氣強勁・穩價鎖利"
            snote = f"住房率已達 {occ_pct:.1f}% 且近期原始買氣強，無須額外降價促銷。"
        else:
            scode = "balanced"
            slabel = "✅ 結構平衡"
            snote = f"主力大床剩 {entry_rem} 間(7日{entry_pk7:+d})｜大床升等池剩 {q_upg_rem} 間｜DT剩 {dt_rem} 間｜家庭剩 {fam_rem} 間"

        daily_diag.append({
            "date": d_iso,
            "dow": dow_zh[d_obj.weekday()],
            "lead_days": idx,
            "occ_pct": occ_pct,
            "rem_rn": rem_rn,
            "entry_rem": entry_rem,
            "entry_pk7": entry_pk7,
            "er_rem": er_rem,
            "q_upg_rem": q_upg_rem,
            "q_upg_cont_rem": q_upg_cont_rem,
            "q_upg_pk7": q_upg_pk7,
            "pt_rem": pt_rem,
            "pt_pk7": pt_pk7,
            "dt_rem": dt_rem,
            "dt_cont_rem": dt_cont_rem,
            "dt_pk7": pk_map.get("DT", 0),
            "fam_rem": fam_rem,
            "fam_pk7": fam_pk7,
            "strategy_code": scode,
            "strategy_label": slabel,
            "evidence_note": snote,
        })

    # 4. 整理 Booking Curves 點位（按 stay_date, bk_date 排序）
    booking_curves = []
    for (st_iso, ort, bk_iso, meal), v in sorted(curve_grp.items()):
        try:
            lead_bk = (datetime.date.fromisoformat(st_iso) - datetime.date.fromisoformat(bk_iso)).days
        except Exception:
            lead_bk = 0
        booking_curves.append({
            "stay_date": st_iso,
            "orig_rt": ort,
            "bk_date": bk_iso,
            "lead_at_bk": lead_bk,
            "meal": meal,
            "rn": v["rn"],
            "avg_rate": int(round(v["rev"] / v["rn"])) if v["rn"] > 0 else 0,
            "upgraded_rn": v["upg_rn"],
        })

    er_share = round(er_abs_rn / tot_upg_rn * 100.0, 1) if tot_upg_rn > 0 else 0.0
    out_payload = {
        "as_of": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "base_date": today.strftime("%Y-%m-%d"),
        "source": "JinXu HT2000 + HT2000_Log.HTACL Read-Only SQL (Zero Upgrade Bias)",
        "summary": {
            "upg_window_label": "近30天已入住＋未來30天在手 (共60天)",
            "total_upgraded_rn": tot_upg_rn,
            "zero_cost_queen_upg_rn": zero_cost_rn,
            "er_absorbed_rn": er_abs_rn,
            "er_share_pct": er_share,
            "cross_bed_upg_rn": cross_bed_rn,
        },
        "upgrade_matrix_60d": upg_matrix[:35],
        "rt_velocity_30d": rt_velocity_30d,
        "daily_yield_diagnostics": daily_diag,
        "booking_curves": booking_curves,
    }
    os.makedirs(BACKUP_DIR, exist_ok=True)
    with open(DYNAMIC_YIELD_PATH, "w", encoding="utf-8") as f:
        json.dump(out_payload, f, ensure_ascii=False, indent=2)
    logger.info(
        f"[SUCCESS] 動態收益與原始買氣軌跡已更新 (60d升等={tot_upg_rn}晚, ER池={er_abs_rn}晚, 曲線點={len(booking_curves)}) -> {DYNAMIC_YIELD_PATH}"
    )
    return out_payload


def run_pms_revenue_intel_sync():
    """每日 05:00 全自動同步入口：通路淨利 + 90天Pace + 12房型庫存 + 動態收益與升等軌跡(HTACL) + 促銷與含早差價健檢。"""
    results = {
        "channel_intel": False,
        "stly_pace": False,
        "roomtype_inventory": False,
        "dynamic_yield": False,
        "promo_bf_audit": False,
    }
    # 若現有 SiteMinder 快照尚未歸檔至 sm_history/，順手確保歷史歸檔初始化
    try:
        if os.path.exists(SM_SNAPSHOT_PATH):
            from siteminder_cdp_sync import archive_siteminder_snapshot
            with open(SM_SNAPSHOT_PATH, "r", encoding="utf-8") as sf:
                archive_siteminder_snapshot(json.load(sf))
    except Exception:
        pass

    try:
        extract_pms_channel_intel()
        results["channel_intel"] = True
    except Exception as e:
        logger.error(f"[FAIL] 金旭通路淨利情報抽取失敗: {e}", exc_info=True)

    try:
        extract_pms_stly_pace()
        results["stly_pace"] = True
    except Exception as e:
        logger.error(f"[FAIL] 金旭 STLY Pace 抽取失敗: {e}", exc_info=True)

    try:
        import wallpaper_planner
        wallpaper_planner.refresh_pms_roomtype_inventory()
        results["roomtype_inventory"] = True
        logger.info("[SUCCESS] 金旭 12 房型每日剩餘庫存快照已同步更新。")
    except Exception as e:
        logger.error(f"[FAIL] 金旭 12 房型庫存快照刷新失敗: {e}", exc_info=True)

    try:
        extract_pms_dynamic_yield()
        results["dynamic_yield"] = True
    except Exception as e:
        logger.error(f"[FAIL] 金旭動態收益與升等軌跡抽取失敗: {e}", exc_info=True)

    try:
        extract_pms_promo_and_bf_audit()
        results["promo_bf_audit"] = True
    except Exception as e:
        logger.error(f"[FAIL] 金旭促銷與含早差價健檢抽取失敗: {e}", exc_info=True)

    try:
        import decision_engine
        decision_engine.append_daily_ledger_snapshot()
        logger.info("[SUCCESS] 每日預警覆盤日誌 (decision_ledger.json) 已自動更新。")
    except Exception as e:
        logger.warning(f"[WARN] 更新每日預警覆盤日誌失敗: {e}")

    return results


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    res = run_pms_revenue_intel_sync()
    print("Sync results:", res)
