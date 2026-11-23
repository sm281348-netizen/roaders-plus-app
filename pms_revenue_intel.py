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


def run_pms_revenue_intel_sync():
    """每日 05:00 全自動同步入口：通路淨利情報 + 90 天 STLY Pace + 12 房型剩餘庫存快照。"""
    results = {"channel_intel": False, "stly_pace": False, "roomtype_inventory": False}
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

    return results


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    res = run_pms_revenue_intel_sync()
    print("Sync results:", res)
