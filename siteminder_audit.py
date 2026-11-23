# -*- coding: utf-8 -*-
"""
路徒 PLUS 行旅（站前館）SiteMinder 定價稽核雷達 × 金旭 PMS 收益情報中樞
==========================================================================
100% 唯讀分析模組（嚴禁修改任何房價、庫存或設定）。
整合兩大自動化數據源（均於每日 05:00 run_daily_sync.py 自動更新並推播雲端）：
  1. 金旭 PMS (HT2000) 全自動唯讀 SQL：
     - data/pms_backup/pms_channel_intel.json（通路預付 Net / 到付 Gross、取消率、ALOS、Lead Time、洗盤風險）
     - data/pms_backup/pms_stly_pace.json（未來 90 天 OTB vs 去年同期 STLY -364d/-365d 及 1日/7日淨 Pickup）
     - data/pms_backup/wallpaper_roomtype_inventory.json（12 房型每日真實剩餘空房與 OOO）
  2. SiteMinder 100% 唯讀 CDP 快照：
     - data/pms_backup/siteminder_inventory_snapshot.json（未來 42 天標準房價、OTA 配額、Stop Sell、MinLOS、漏單警示）
     - data/pms_backup/siteminder_sync_status.json（每日同步狀態紀錄）
  3. 亦支援拖拉上傳 SiteMinder 匯出之 Pace / Booking performance CSV 即時比對。
"""

import os
import io
import json
import calendar
import datetime
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")

SM_SNAPSHOT_PATH = os.path.join(BACKUP_DIR, "siteminder_inventory_snapshot.json")
SM_STATUS_PATH = os.path.join(BACKUP_DIR, "siteminder_sync_status.json")
CHAN_INTEL_PATH = os.path.join(BACKUP_DIR, "pms_channel_intel.json")
STLY_PACE_PATH = os.path.join(BACKUP_DIR, "pms_stly_pace.json")
RT_INV_PATH = os.path.join(BACKUP_DIR, "wallpaper_roomtype_inventory.json")
OCC_CSV_PATH = os.path.join(BACKUP_DIR, "station_occ_data.csv")

# 依金旭 HTRT 定義之房型階梯順序（高 → 低）
RT_TIER_ORDER = ["DF", "F", "DTP", "LDT", "DT", "E", "PD", "PT", "SP", "S"]
RT_NAMES = {
    "DF": "豪華家庭房 DF",
    "F": "家庭房 F",
    "DTP": "豪華三人房 DTP",
    "LDT": "奢華雙人房 LDT",
    "DT": "豪華雙人房 DT",
    "E": "101行政 E(05)",
    "ER": "101行政 ER (未掛OTA)",
    "PD": "尊爵雙人房 PD",
    "PT": "尊爵雙床房 PT",
    "SP": "高級雙人房 SP",
    "S": "標準雙人房 S",
    "SD": "無障礙 SD (未掛OTA)",
}

# 必須嚴格維持「高階 >= 低階」的房型階梯檢查對 (高階, 低階, 是否為同級雙人/雙床)
TIER_PAIRS = [
    ("DF", "F", False),
    ("F", "DTP", False),
    ("DTP", "LDT", False),
    ("LDT", "DT", False),
    ("DT", "E", False),
    ("E", "PD", False),
    ("E", "PT", False),
    ("PD", "SP", False),
    ("PT", "SP", False),
    ("SP", "S", False),
]


def _load_json_safe(path: str, default=None):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {} if default is None else default


def _inject_module_css():
    st.markdown(
        """
<style>
.sm-hero {
    background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
    color: #f8fafc;
    border-radius: 14px;
    padding: 20px 24px;
    border-left: 6px solid #f59e0b;
    margin-bottom: 16px;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.16);
}
.sm-hero h2 {
    margin: 0 0 6px 0;
    font-size: 1.45rem;
    color: #ffffff;
}
.sm-hero p {
    margin: 0;
    color: #cbd5e1;
    font-size: 0.92rem;
    line-height: 1.5;
}
.sm-badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 700;
    margin-right: 6px;
    margin-top: 8px;
}
.sm-badge-ok { background: rgba(16,185,129,0.18); color: #6ee7b7; border: 1px solid rgba(16,185,129,0.5); }
.sm-badge-warn { background: rgba(245,158,11,0.20); color: #fde68a; border: 1px solid rgba(245,158,11,0.55); }
.sm-badge-err { background: rgba(239,68,68,0.22); color: #fca5a5; border: 1px solid rgba(239,68,68,0.6); }

.sm-kpi {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 14px 16px;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
    height: 100%;
}
.sm-kpi-title { font-size: 0.82rem; color: #64748b; font-weight: 600; margin-bottom: 4px; }
.sm-kpi-val { font-size: 1.55rem; color: #0f172a; font-weight: 800; line-height: 1.2; font-variant-numeric: tabular-nums; }
.sm-kpi-sub { font-size: 0.78rem; color: #475569; margin-top: 5px; line-height: 1.35; }
.sm-alert-box {
    background: #fef2f2;
    border: 1px solid #fecaca;
    border-left: 5px solid #ef4444;
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 14px;
    color: #7f1d1d;
    font-size: 0.9rem;
    line-height: 1.5;
}
</style>
""",
        unsafe_allow_html=True,
    )


def run_pricing_audit_rules(
    sm_snap: dict,
    pace_data: dict,
    rt_inv: dict,
    events_map: dict,
    stly_mode: str = "dow",
    last_min_days: int = 3,
    underprice_pace_rn: int = 5,
    fat_finger_drop_pct: float = 40.0,
    allow_nye_101_exception: bool = True,
):
    """
    針對 SiteMinder 掛牌價快照 × 金旭 PMS STLY Pace × 12 房型庫存 × 台北活動，
    自動執行 6 大定價疏漏稽核，回傳結構化異常清單 DataFrame。
    """
    inv_map = sm_snap.get("inventory", {})
    if not inv_map:
        return pd.DataFrame()

    pace_by_date = {r["date"]: r for r in pace_data.get("daily", [])}
    pms_daily_rt = rt_inv.get("daily_avail", {})
    base_date_str = pace_data.get("base_date") or datetime.date.today().strftime("%Y-%m-%d")
    try:
        base_d = datetime.date.fromisoformat(base_date_str)
    except Exception:
        base_d = datetime.date.today()

    dow_zh = ["一", "二", "三", "四", "五", "六", "日"]
    dates_sorted = sorted(inv_map.keys())
    alerts = []

    # 預先計算同星期平均 S 基準價（供臨期滯銷與 Fat-Finger 比較）
    dow_s_rates = {i: [] for i in range(7)}
    for d_str in dates_sorted:
        try:
            wd = datetime.date.fromisoformat(d_str).weekday()
        except Exception:
            continue
        sr = inv_map[d_str].get("rooms", {}).get("S", {}).get("r", 0)
        if sr > 0:
            dow_s_rates[wd].append(sr)
    dow_s_med = {wd: float(np.median(vals)) if vals else 0.0 for wd, vals in dow_s_rates.items()}

    for d_str in dates_sorted:
        day_obj = inv_map[d_str]
        rooms = day_obj.get("rooms", {})
        ss_flag = str(day_obj.get("ss", "N")).upper()
        mlos = int(day_obj.get("mlos", 1) or 1)

        try:
            d_obj = datetime.date.fromisoformat(d_str)
            wd = d_obj.weekday()
            dow_lbl = dow_zh[wd]
            lead_d = (d_obj - base_d).days
        except Exception:
            wd, dow_lbl, lead_d = 0, "", 99

        pace = pace_by_date.get(d_str, {})
        yoy_rn = pace.get("yoy_dow_rn" if stly_mode == "dow" else "yoy_date_rn", 0)
        yoy_adr = pace.get("yoy_dow_adr" if stly_mode == "dow" else "yoy_date_adr", 0)
        stly_adr = pace.get("stly_dow_adr" if stly_mode == "dow" else "stly_date_adr", 0)
        otb_adr = pace.get("otb_adr", 0)
        occ_pct = pace.get("occ_pct", 0.0)
        pms_rem_tot = pace.get("rem_rn", None)
        ev_name = events_map.get(d_str, "")

        sm_tot_avail = sum(int(v.get("a", 0)) for v in rooms.values())

        # 規則 1：房型階梯倒掛（高階價 < 低階價）
        for hi, lo, _ in TIER_PAIRS:
            if hi not in rooms or lo not in rooms:
                continue
            r_hi, a_hi = int(rooms[hi].get("r", 0)), int(rooms[hi].get("a", 0))
            r_lo, a_lo = int(rooms[lo].get("r", 0)), int(rooms[lo].get("a", 0))
            if r_hi <= 0 or r_lo <= 0:
                continue
            if r_lo > r_hi:
                # 跨年夜 101 景觀行政房 E(05) > DT 屬煙火景觀溢價，可選擇標註為提醒而非嚴重倒掛
                is_nye_101 = allow_nye_101_exception and d_str == "2026-12-31" and hi == "DT" and lo == "E"
                diff = r_lo - r_hi
                if is_nye_101:
                    sev = "🟡 提醒"
                    note = f"跨年夜 101 景觀房 E(05) ({r_lo:,}) 高於 DT ({r_hi:,}) +{diff:,}（屬煙火景觀溢價，確認為刻意操作即可）"
                else:
                    sev = "🔴 嚴重" if (a_lo > 0 or diff >= 500) else "🟠 警告"
                    avail_note = "雙方或低階仍有庫存，客人下單即倒掛！" if a_lo > 0 else "低階目前庫存為 0，但一旦有人退訂釋房將立刻呈現倒掛。"
                    note = f"低階 {lo} ({r_lo:,}, 剩{a_lo}) 高於高階 {hi} ({r_hi:,}, 剩{a_hi}) +{diff:,} 元。{avail_note}"
                alerts.append({
                    "日期": f"{d_str} ({dow_lbl})",
                    "date": d_str,
                    "嚴重度": sev,
                    "異常類型": "📉 房價階梯倒掛",
                    "涉及房型": f"{lo} > {hi}",
                    "掛牌價/數據摘要": f"{lo} ${r_lo:,} vs {hi} ${r_hi:,}",
                    "說明與收益部提問建議": note,
                    "台北活動": ev_name,
                })
            elif r_lo == r_hi and (a_hi > 0 and a_lo > 0):
                # 規則 2：跨階同價零價差（且雙方皆有可售庫存！）
                alerts.append({
                    "日期": f"{d_str} ({dow_lbl})",
                    "date": d_str,
                    "嚴重度": "🟠 警告",
                    "異常類型": "⚖️ 跨階同價零價差",
                    "涉及房型": f"{lo} ＝ {hi}",
                    "掛牌價/數據摘要": f"皆掛 ${r_hi:,}（庫存 {lo}:{a_lo} / {hi}:{a_hi}）",
                    "說明與收益部提問建議": f"{hi} 與低一階 {lo} 同價且雙雙有房可售，等於免費升級壓縮原本可收價差，建議拉開合理階梯。",
                    "台北活動": ev_name,
                })

        # 規則 3：臨期滯銷卻硬撐高價 (Overpricing)
        s_rate = int(rooms.get("S", {}).get("r", 0))
        med_s = dow_s_med.get(wd, 0.0)
        if 0 <= lead_d <= last_min_days and sm_tot_avail >= 10 and med_s > 0 and s_rate >= med_s * 1.20:
            if yoy_rn <= -2 or occ_pct < 85.0:
                pct_above = round((s_rate / med_s - 1.0) * 100.0, 1)
                alerts.append({
                    "日期": f"{d_str} ({dow_lbl})",
                    "date": d_str,
                    "嚴重度": "🟠 警告",
                    "異常類型": "❄️ 臨期高價滯銷",
                    "涉及房型": "全館主力房",
                    "掛牌價/數據摘要": f"S ${s_rate:,}（高於同週中位 +{pct_above}%）｜Pace {yoy_rn:+d}晚",
                    "說明與收益部提問建議": f"距入住僅 {lead_d} 天、SiteMinder 仍釋出 {sm_tot_avail} 間（金旭住房率 {occ_pct:.1f}%），標準房價卻高於同星期中位數 ${int(med_s):,} 達 +{pct_above}%，建議評估臨期促銷去化。",
                    "台北活動": ev_name,
                })

        # 規則 4：買氣強勁／大活動但 ADR 低於去年同期 (Underpricing)
        if pace and yoy_rn >= underprice_pace_rn and yoy_adr <= -150 and stly_adr > 0:
            adr_pct = round(yoy_adr / stly_adr * 100.0, 1)
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})",
                "date": d_str,
                "嚴重度": "🟠 警告",
                "異常類型": "🔥 買氣強卻賣便宜",
                "涉及房型": "全館均價",
                "掛牌價/數據摘要": f"Pace {yoy_rn:+d}晚｜ADR ${otb_adr:,} (YoY {adr_pct:+.1f}%)",
                "說明與收益部提問建議": f"在手房晚已大幅領先去年同期 +{yoy_rn} 晚（住房率 {occ_pct:.1f}%），但成交 ADR 反較去年低 ${abs(yoy_adr):,}，建議檢視是否早鳥促銷放量過多並提早墊高剩餘房價。",
                "台北活動": ev_name,
            })

        # 規則 5：幽靈關房 / 高壓日 MinLOS 檢核
        if ss_flag == "Y" and (pms_rem_tot is not None and pms_rem_tot >= 4):
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})",
                "date": d_str,
                "嚴重度": "🔴 嚴重",
                "異常類型": "🚫 疑似幽靈關房",
                "涉及房型": "全館",
                "掛牌價/數據摘要": f"Stop Sell = Y｜金旭剩餘 {pms_rem_tot} 間",
                "說明與收益部提問建議": "金旭 PMS 仍有空房，但 SiteMinder 顯示 Stop Sell，請立即確認是否誤鎖！",
                "台北活動": ev_name,
            })
        if mlos >= 3 and sm_tot_avail >= 4 and lead_d <= 45:
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})",
                "date": d_str,
                "嚴重度": "🟡 提醒",
                "異常類型": "🔒 高連住門檻尾房檢核",
                "涉及房型": "全館",
                "掛牌價/數據摘要": f"MinLOS = {mlos} 晚｜OTA 剩 {sm_tot_avail} 間",
                "說明與收益部提問建議": f"當日設定最低連住 {mlos} 晚且尚有 {sm_tot_avail} 間可售尾房，建議追蹤去化速度，若臨近兩週仍有空房應評估降為 2 晚或 1 晚。",
                "台北活動": ev_name,
            })

        # 規則 6：Fat-Finger 暴跌防呆（單日 S 房價較同星期中位數暴跌 > 門檻）
        if med_s > 0 and s_rate > 0 and s_rate < med_s * (1.0 - fat_finger_drop_pct / 100.0):
            drop_pct = round((1.0 - s_rate / med_s) * 100.0, 1)
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})",
                "date": d_str,
                "嚴重度": "🔴 嚴重",
                "異常類型": "⚡ 疑誤輸入暴跌 (Fat-Finger)",
                "涉及房型": "S",
                "掛牌價/數據摘要": f"S ${s_rate:,}（較同週中位暴跌 -{drop_pct}%）",
                "說明與收益部提問建議": "掛牌價異常暴跌，請立刻核對是否手動輸入少打位數！",
                "台北活動": ev_name,
            })

        # 補充：金旭 ER (8間) 庫存未上架提醒（僅彙總於首日避免重複 42 筆）
        if d_str == dates_sorted[0]:
            er_rem = pms_daily_rt.get(d_str, {}).get("ER", {}).get("rem", None)
            if er_rem is not None:
                alerts.append({
                    "日期": "全區間結構",
                    "date": "0000-00-00",
                    "嚴重度": "🟡 提醒",
                    "異常類型": "🧩 金旭 vs SiteMinder 房型缺口",
                    "涉及房型": "ER (8間) / SD (3間)",
                    "掛牌價/數據摘要": f"SiteMinder 僅掛 10 房型｜今日 ER 剩 {er_rem} 間",
                    "說明與收益部提問建議": "金旭 HTRT 現行 12 房型中之 ER (8間) 與 SD (3間可售) 未見於 SiteMinder，請確認 ER 是否併入 E 房型池銷售或刻意留作現場升等池。",
                    "台北活動": "",
                })

    df_alerts = pd.DataFrame(alerts)
    if not df_alerts.empty:
        sev_rank = {"🔴 嚴重": 1, "🟠 警告": 2, "🟡 提醒": 3}
        df_alerts["_rank"] = df_alerts["嚴重度"].map(sev_rank).fillna(9)
        df_alerts = df_alerts.sort_values(["_rank", "date"]).drop(columns=["_rank"]).reset_index(drop=True)
    return df_alerts


def build_excel_audit_report(df_alerts, df_rates, df_pace, df_chan):
    """產出 4 分頁專業定價稽核 Excel (.xlsx) 供經理會議對質與存查。"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (df_alerts if not df_alerts.empty else pd.DataFrame([{"訊息": "目前無觸發異常"}])).to_excel(
            writer, sheet_name="1_定價疏漏清單", index=False
        )
        (df_rates if not df_rates.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="2_未來42天掛牌價矩陣", index=False
        )
        (df_pace if not df_pace.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="3_STLY進度與Pickup", index=False
        )
        (df_chan if not df_chan.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="4_通路真實淨利排行", index=False
        )
        for ws in writer.book.worksheets:
            for col in ws.columns:
                max_len = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 4, 12), 46)
    return buf.getvalue()


def _parse_uploaded_sm_csv(uploaded_file):
    """彈性解析使用者手動上傳之 SiteMinder Pace 或 Booking performance CSV。"""
    try:
        df = pd.read_csv(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception:
            return pd.DataFrame()


def render_siteminder_audit(taipei_events_df=None, targets_df=None):
    """渲染『📡 專案：SiteMinder 定價稽核雷達』獨立互動模組（站前館專屬）。"""
    _inject_module_css()

    # 1. 載入所有本地雲端快照（零 SQL 直連依賴，確保 Streamlit Cloud 100% 穩健）
    sm_snap = _load_json_safe(SM_SNAPSHOT_PATH)
    sm_status = _load_json_safe(SM_STATUS_PATH)
    chan_intel = _load_json_safe(CHAN_INTEL_PATH)
    pace_data = _load_json_safe(STLY_PACE_PATH)
    rt_inv = _load_json_safe(RT_INV_PATH)

    health = sm_snap.get("health", {})
    sm_asof = sm_snap.get("as_of", "尚未擷取")
    pms_asof = pace_data.get("as_of") or chan_intel.get("as_of") or "尚未同步"
    rt_asof = rt_inv.get("as_of", "尚未同步")
    n_inv_days = len(sm_snap.get("inventory", {}))

    fail_cnt = int(health.get("delivery_failures_today", 0) or 0)
    mapped_n = health.get("rates_mapped", 0)
    mapped_tot = health.get("rates_total", 0)
    comp_cnt = int(health.get("competitors_configured", 0) or 0)

    # 整理台北活動對照表
    events_map = {}
    if taipei_events_df is not None and not taipei_events_df.empty and "date" in taipei_events_df.columns:
        for _, er in taipei_events_df.iterrows():
            ds = str(er.get("date", "")).strip()[:10]
            en = str(er.get("event_name", "")).strip()
            if ds and en:
                events_map[ds] = f"{events_map[ds]} / {en}" if ds in events_map else en

    # 2. 頂部 Hero Banner
    b_fail_cls = "sm-badge-err" if fail_cnt > 0 else "sm-badge-ok"
    b_map_cls = "sm-badge-warn" if (mapped_tot > 0 and mapped_n < mapped_tot) else "sm-badge-ok"
    b_comp_cls = "sm-badge-warn" if comp_cnt == 0 else "sm-badge-ok"
    st.markdown(
        f"""
<div class="sm-hero">
  <h2>📡 站前館 SiteMinder 定價稽核雷達 × 金旭 PMS 收益情報中樞</h2>
  <p>
    100% 唯讀營運稽核｜結合 <b>每日 05:00 金旭 HT2000 自動同步</b>（通路預付/到付淨利、取消洗盤、STLY Pace、12 房型剩餘庫存）
    與 <b>SiteMinder 未來 {n_inv_days} 天掛牌價網格</b>，一鍵抓出收益策略部定價盲點與通路獲利真相。
  </p>
  <div>
    <span class="sm-badge sm-badge-ok">🕒 金旭 05:00 快照：{pms_asof}</span>
    <span class="sm-badge sm-badge-ok">🌐 SiteMinder 掛牌快照：{sm_asof}</span>
    <span class="sm-badge {b_fail_cls}">🚨 今日 PMS 派單失敗：{fail_cnt} 筆</span>
    <span class="sm-badge {b_map_cls}">🔗 GoldenUp 對應：{mapped_n}/{mapped_tot}</span>
    <span class="sm-badge {b_comp_cls}">👀 競品監控設定：{comp_cnt} 家</span>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    # 3. 緊急營運止血警示（若有派單失敗）
    if fail_cnt > 0:
        fd = health.get("failed_delivery_detail", {})
        if fd:
            st.markdown(
                f"""
<div class="sm-alert-box">
  <b>🚨 今日營運止血警示：SiteMinder 有 {fail_cnt} 筆訂單派送至金旭 PMS 失敗！</b><br>
  • <b>通路／訂單編號</b>：{fd.get('channel','-')} <code>{fd.get('reservation_id','-')}</code>
  ｜<b>房型／方案</b>：{fd.get('room_type','-')} ({fd.get('rate_plan','-')})
  ｜<b>入退日</b>：{fd.get('checkin','-')} → {fd.get('checkout','-')}（{fd.get('nights','-')} 晚，TWD {fd.get('total_twd',0):,.2f}）<br>
  • <b>異動時間與研判</b>：{fd.get('modified_at','-')} — {fd.get('likely_reason','請櫃檯人工核對金旭訂單狀態。')}
</div>
""",
                unsafe_allow_html=True,
            )

    # 4. 全域參數模擬與稽核設定控制台
    with st.expander("⚙️ 全域稽核門檻、各通路合約抽成試算參數 ＆ 手動補傳 SiteMinder CSV（點擊展開）", expanded=False):
        c_p1, c_p2, c_p3 = st.columns(3)
        with c_p1:
            st.markdown("**1️⃣ 稽核規則與同期比較口徑**")
            stly_mode_lbl = st.radio(
                "STLY 去年同期比對基準",
                ["同星期比對 (-364天｜SiteMinder 預設)", "同日期比對 (-365天｜跨年節慶對照)"],
                horizontal=False,
                key="sm_stly_mode",
            )
            stly_mode = "dow" if "-364" in stly_mode_lbl else "date"
            allow_nye_101 = st.checkbox("跨年夜 12/31 允許 101 景觀房 E(05) > DT（降級為提醒）", value=True, key="sm_nye_101")
            last_min_days = st.slider("臨期滯銷檢核天數 (≤ N 天)", 1, 14, 3, key="sm_lastmin_d")
        with c_p2:
            st.markdown("**2️⃣ 各通路到付毛價抽成試算 (%)**")
            st.caption("金旭 HTGR 已自動標記「預付=淨價(已扣抽成)」、「到付=毛價(未扣抽成)」，避免重複扣兩次。")
            comm_booking = st.slider("Booking.com 到付抽成 (%)", 10.0, 22.0, 15.0, 0.5, key="sm_c_bk")
            comm_ota_gross = st.slider("Trip / Agoda / Expedia 到付單抽成 (%)", 10.0, 22.0, 15.0, 0.5, key="sm_c_otag")
            comm_direct = st.slider("官網 Direct 金流手續費 (%)", 0.0, 5.0, 2.0, 0.5, key="sm_c_dir")
            comm_rakuten = st.slider("Rakuten 樂天到付抽成 (%)", 8.0, 18.0, 12.0, 0.5, key="sm_c_rak")
        with c_p3:
            st.markdown("**3️⃣ 稽核敏感度 & 手動加掛 CSV**")
            underprice_pace_rn = st.slider("買氣強領先門檻 (Pace ≥ +N 晚且 ADR 下降)", 2, 20, 5, key="sm_up_rn")
            fat_finger_drop_pct = st.slider("Fat-Finger 暴跌防呆門檻 (%)", 25.0, 60.0, 40.0, 5.0, key="sm_ff_pct")
            up_csv = st.file_uploader("選用：拖拉上傳 SiteMinder Pace / Booking Performance CSV", type=["csv"], key="sm_up_csv")

    if up_csv is not None:
        df_up = _parse_uploaded_sm_csv(up_csv)
        if not df_up.empty:
            with st.expander(f"📎 已載入您手動上傳的 SiteMinder CSV（{up_csv.name}，共 {len(df_up)} 列）", expanded=True):
                st.dataframe(df_up, use_container_width=True, height=260)

    # 執行自動 6 大規則掃描
    df_alerts = run_pricing_audit_rules(
        sm_snap=sm_snap,
        pace_data=pace_data,
        rt_inv=rt_inv,
        events_map=events_map,
        stly_mode=stly_mode,
        last_min_days=last_min_days,
        underprice_pace_rn=underprice_pace_rn,
        fat_finger_drop_pct=fat_finger_drop_pct,
        allow_nye_101_exception=allow_nye_101,
    )

    # 預先計算未來 31 天 Pace 摘要與本月通路淨房價差（供頂部 4 張 KPI 卡）
    pace_daily = pace_data.get("daily", [])
    pace_31 = pace_daily[:31]
    rev_31 = sum(r["otb_rev"] for r in pace_31)
    rn_31 = sum(r["sold_rn"] for r in pace_31)
    stly_rev_31 = sum(r["stly_dow_rev" if stly_mode == "dow" else "stly_date_rev"] for r in pace_31)
    stly_rn_31 = sum(r["stly_dow_rn" if stly_mode == "dow" else "stly_date_rn"] for r in pace_31)
    p7d_31 = sum(r["pickup_7d_rn"] for r in pace_31)
    yoy_rev_pct = ((rev_31 / stly_rev_31 - 1.0) * 100.0) if stly_rev_31 > 0 else 0.0
    yoy_rn_diff = rn_31 - stly_rn_31

    # 計算通路真實 Net ADR（依 MTD 訂房日）
    def _calc_net_adr_df(rows):
        if not rows:
            return pd.DataFrame()
        recs = []
        for r in rows:
            basis = r["rate_basis"]
            ch = r["channel"]
            if basis == "net_prepaid":
                comm = 0.0
                basis_zh = "預付(已淨價)"
            elif basis == "gross_pay_at_hotel":
                if ch == "Booking.com":
                    comm = comm_booking
                elif "Rakuten" in ch:
                    comm = comm_rakuten
                else:
                    comm = comm_ota_gross
                basis_zh = f"到付毛價(扣{comm:.0f}%)"
            elif basis == "direct":
                comm = comm_direct
                basis_zh = f"直客(金流{comm:.1f}%)"
            else:
                comm = 0.0
                basis_zh = "合約/專案"

            net_real_rev = int(round(r["net_rev"] * (1.0 - comm / 100.0)))
            net_real_adr = int(round(net_real_rev / r["net_rn"])) if r["net_rn"] > 0 else 0
            recs.append({
                **r,
                "口徑": basis_zh,
                "適用抽成%": comm,
                "飯店實拿淨營收": net_real_rev,
                "飯店實拿 Net ADR": net_real_adr,
            })
        return pd.DataFrame(recs)

    mtd_rows = chan_intel.get("windows", {}).get("booked_mtd", {}).get("rows", [])
    df_mtd_detail = _calc_net_adr_df(mtd_rows)
    direct_net_adr, trip_net_adr, top_ota_name = 0, 0, "Trip.com"
    if not df_mtd_detail.empty:
        g_dir = df_mtd_detail[df_mtd_detail["channel"] == "官網 Direct (TBB)"]
        if not g_dir.empty and g_dir["net_rn"].sum() > 0:
            direct_net_adr = int(round(g_dir["飯店實拿淨營收"].sum() / g_dir["net_rn"].sum()))
        g_ota = df_mtd_detail[df_mtd_detail["is_ota"]].groupby("channel", as_index=False).agg(
            rn=("net_rn", "sum"), net_rev=("飯店實拿淨營收", "sum")
        )
        if not g_ota.empty:
            top_row = g_ota.sort_values("rn", ascending=False).iloc[0]
            top_ota_name = top_row["channel"]
            trip_net_adr = int(round(top_row["net_rev"] / top_row["rn"])) if top_row["rn"] > 0 else 0

    sev_red = int((df_alerts["嚴重度"] == "🔴 嚴重").sum()) if not df_alerts.empty else 0
    sev_org = int((df_alerts["嚴重度"] == "🟠 警告").sum()) if not df_alerts.empty else 0
    sev_yel = int((df_alerts["嚴重度"] == "🟡 提醒").sum()) if not df_alerts.empty else 0

    # 5. 4 欄 Executive KPI Cards
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            f"""
<div class="sm-kpi" style="border-top: 4px solid #2563eb;">
  <div class="sm-kpi-title">📈 未來 31 天在手營收 (金旭 OTB)</div>
  <div class="sm-kpi-val">NT$ {rev_31/10000:,.1f} 萬</div>
  <div class="sm-kpi-sub">
    STLY 進度：<b>{yoy_rev_pct:+.1f}%</b>（<b>{yoy_rn_diff:+d}</b> 房晚）<br>
    近 7 日淨增訂 (Pickup)：<b>{p7d_31:+d}</b> 房晚
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f"""
<div class="sm-kpi" style="border-top: 4px solid #f97316;">
  <div class="sm-kpi-title">🚨 未來 {n_inv_days} 天定價異常檢出</div>
  <div class="sm-kpi-val">{len(df_alerts)} 項</div>
  <div class="sm-kpi-sub">
    🔴 嚴重倒掛/暴跌：<b>{sev_red}</b> 項｜🟠 警告：<b>{sev_org}</b> 項<br>
    🟡 提醒/結構檢核：<b>{sev_yel}</b> 項
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
    with k3:
        gap_net = direct_net_adr - trip_net_adr
        ratio_eq = (direct_net_adr / trip_net_adr) if trip_net_adr > 0 else 0
        st.markdown(
            f"""
<div class="sm-kpi" style="border-top: 4px solid #10b981;">
  <div class="sm-kpi-title">💎 官網 vs {top_ota_name} 真實淨房價差</div>
  <div class="sm-kpi-val">+NT$ {gap_net:,}</div>
  <div class="sm-kpi-sub">
    官網實拿 <b>${direct_net_adr:,}</b> vs {top_ota_name} <b>${trip_net_adr:,}</b><br>
    賣 1 間官網 ≈ 賣 <b>{ratio_eq:.2f}</b> 間 {top_ota_name}
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
    with k4:
        unmapped = max(0, mapped_tot - mapped_n)
        st.markdown(
            f"""
<div class="sm-kpi" style="border-top: 4px solid #8b5cf6;">
  <div class="sm-kpi-title">🔌 SiteMinder × 金旭連線健檢</div>
  <div class="sm-kpi-val">{mapped_n}/{mapped_tot} 方案</div>
  <div class="sm-kpi-sub">
    未對應：<b>{unmapped}</b> 方案｜今日漏單：<b>{fail_cnt}</b> 筆<br>
    競品設定：<b>{comp_cnt}</b> 家（未開啟比價雷達）
  </div>
</div>
""",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # 預先整理 42 天掛牌價 DataFrame 供 Tab 2 與 Excel 匯出共用
    inv_map = sm_snap.get("inventory", {})
    pms_daily_rt = rt_inv.get("daily_avail", {})
    pace_by_date = {r["date"]: r for r in pace_daily}
    dow_zh = ["一", "二", "三", "四", "五", "六", "日"]
    rate_rows = []
    for d_str in sorted(inv_map.keys()):
        try:
            wd = dow_zh[datetime.date.fromisoformat(d_str).weekday()]
        except Exception:
            wd = ""
        d_obj = inv_map[d_str]
        rooms = d_obj.get("rooms", {})
        pc = pace_by_date.get(d_str, {})
        row_dict = {
            "日期": f"{d_str} ({wd})",
            "金旭住房率%": pc.get("occ_pct", 0.0),
            "金旭真實剩房": pc.get("rem_rn", 0),
            "OTA釋出合計": sum(int(v.get("a", 0)) for v in rooms.values()),
            "MinLOS": int(d_obj.get("mlos", 1) or 1),
            "台北活動": events_map.get(d_str, ""),
        }
        for rt in RT_TIER_ORDER:
            cell = rooms.get(rt, {})
            r_v, a_v = int(cell.get("r", 0)), int(cell.get("a", 0))
            row_dict[rt] = f"${r_v:,} ({a_v})" if r_v > 0 else "—"
        er_rem = pms_daily_rt.get(d_str, {}).get("ER", {}).get("rem", None)
        row_dict["金旭ER剩房"] = er_rem if er_rem is not None else "—"
        rate_rows.append(row_dict)
    df_rates_export = pd.DataFrame(rate_rows)
    df_pace_export = pd.DataFrame(pace_daily)

    # 6. 五大互動功能分頁
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🚨 定價疏漏雷達 & 對質清單",
        "📅 42天掛牌價 × 金旭庫存矩陣",
        "📈 STLY Pace × Pickup × 達標試算",
        "💰 通路真實淨利 (Net ADR) & 洗盤預警",
        "⚙️ 05:00 全自動同步狀態",
    ])

    # ------------------------------------------------------------------
    # TAB 1: 定價疏漏雷達 & 收益部對質清單
    # ------------------------------------------------------------------
    with tab1:
        st.subheader("🚨 收益策略部定價疏漏自動掃描結果（100% 實測數據）")
        if df_alerts.empty:
            st.success("✅ 目前掃描區間內未發現定價異常！")
        else:
            f1, f2, f3 = st.columns([1.2, 1.8, 1.2])
            with f1:
                sev_opts = sorted(df_alerts["嚴重度"].unique().tolist())
                sel_sevs = st.multiselect("篩選嚴重度", sev_opts, default=sev_opts, key="sm_f_sev")
            with f2:
                cat_opts = sorted(df_alerts["異常類型"].unique().tolist())
                sel_cats = st.multiselect("篩選異常類型", cat_opts, default=cat_opts, key="sm_f_cat")
            with f3:
                excel_bytes = build_excel_audit_report(df_alerts, df_rates_export, df_pace_export, df_mtd_detail)
                st.download_button(
                    "📥 匯出 4 分頁稽核 Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"站前館_SiteMinder定價稽核_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            df_show = df_alerts[df_alerts["嚴重度"].isin(sel_sevs) & df_alerts["異常類型"].isin(sel_cats)].copy()
            st.dataframe(
                df_show[["日期", "嚴重度", "異常類型", "涉及房型", "掛牌價/數據摘要", "說明與收益部提問建議", "台北活動"]],
                use_container_width=True,
                hide_index=True,
                height=440,
            )

            with st.expander("💬 一鍵複製：給收益策略部的會議對質／LINE 溫和核對清單", expanded=False):
                top_red = df_alerts[df_alerts["嚴重度"].isin(["🔴 嚴重", "🟠 警告"])].head(10)
                lines = [
                    f"【站前館 定價與通路系統例行核對｜基準日 {datetime.date.today().strftime('%Y-%m-%d')}】",
                    "各位收益部夥伴好，我們用系統交叉核對未來 6 週掛牌價與金旭 PMS 庫存時，整理出幾項想請協助確認的項目：",
                    "",
                ]
                if fail_cnt > 0:
                    fd = health.get("failed_delivery_detail", {})
                    lines.append(f"0. 【緊急】今日有 {fail_cnt} 筆派單至金旭失敗（{fd.get('channel','')} #{fd.get('reservation_id','')}），另 GoldenUp 目前對應 {mapped_n}/{mapped_tot}（尚有 {max(0,mapped_tot-mapped_n)} 方案未對應）。")
                for idx, r in enumerate(top_red.to_dict("records"), 1):
                    lines.append(f"{idx}. {r['日期']}｜{r['異常類型']}（{r['涉及房型']}）：{r['掛牌價/數據摘要']} — {r['說明與收益部提問建議']}")
                if comp_cnt == 0:
                    lines.append(f"\n📌 補充建議：SiteMinder『Competitor rates』目前設定為 0 家競品，建議加入台北車站周邊 5–8 家核心同業以利即時對標。")
                lines.append("再麻煩協助確認上述日期是否為刻意控價或規則連動疏漏，感謝協助！")
                st.code("\n".join(lines), language="text")

    # ------------------------------------------------------------------
    # TAB 2: 未來 42 天掛牌價 × 金旭剩餘空房 × 台北活動矩陣
    # ------------------------------------------------------------------
    with tab2:
        st.subheader(f"📅 未來 {n_inv_days} 天主控價『標準房價(不含早)』× 金旭真實空房矩陣")
        st.caption(
            "格內格式：`$掛牌價 (SiteMinder 開放庫存)`。紅底＝該房型價低於其低一階房型（倒掛）；黃底＝與相鄰階同價且雙雙有房。"
            "最左側同步對照金旭 PMS 真實剩餘空房（128 扣 OOO 扣已售）與 ER 未上架空房數。"
        )

        if not inv_map:
            st.info("尚未載入 SiteMinder 掛牌價快照。")
        else:
            # 構建樣式標色函式
            def _style_rate_matrix(df_in):
                styles = pd.DataFrame("", index=df_in.index, columns=df_in.columns)
                dates_list = sorted(inv_map.keys())
                for idx, d_str in enumerate(dates_list):
                    if idx >= len(df_in):
                        break
                    rooms = inv_map[d_str].get("rooms", {})
                    for hi, lo, _ in TIER_PAIRS:
                        if hi in rooms and lo in rooms:
                            rh, ah = int(rooms[hi].get("r", 0)), int(rooms[hi].get("a", 0))
                            rl, al = int(rooms[lo].get("r", 0)), int(rooms[lo].get("a", 0))
                            if rl > rh > 0:
                                if not (allow_nye_101 and d_str == "2026-12-31" and hi == "DT" and lo == "E"):
                                    styles.at[idx, lo] = "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
                                    styles.at[idx, hi] = "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
                            elif rl == rh > 0 and ah > 0 and al > 0:
                                styles.at[idx, lo] = "background-color: #fef3c7; color: #92400e; font-weight: 700;"
                                styles.at[idx, hi] = "background-color: #fef3c7; color: #92400e; font-weight: 700;"
                return styles

            st.dataframe(
                df_rates_export.style.apply(_style_rate_matrix, axis=None).format({"金旭住房率%": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True,
                height=520,
            )

            # 圖表：金旭真實剩房 vs SiteMinder OTA 開放配額對比
            st.markdown("#### 🔍 金旭 PMS 真實空房 vs SiteMinder OTA 開放配額落差（含 `ER` 8 間未掛 OTA 檢視）")
            fig_cap = go.Figure()
            fig_cap.add_trace(go.Bar(
                x=df_rates_export["日期"], y=df_rates_export["金旭真實剩房"],
                name="金旭 PMS 真實剩餘空房 (12房型)", marker_color="#3b82f6", opacity=0.75,
            ))
            fig_cap.add_trace(go.Scatter(
                x=df_rates_export["日期"], y=df_rates_export["OTA釋出合計"],
                name="SiteMinder 開放配額合計 (10房型)", mode="lines+markers",
                line=dict(color="#f97316", width=3),
            ))
            fig_cap.update_layout(
                height=340, margin=dict(l=20, r=20, t=20, b=40),
                legend=dict(orientation="h", y=1.12), hovermode="x unified",
                yaxis_title="間數",
            )
            st.plotly_chart(fig_cap, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 3: STLY Pace × 7日 Pickup × 月底達標價量試算器
    # ------------------------------------------------------------------
    with tab3:
        st.subheader("📈 未來 90 天 STLY 去年同期進度 (Pace) × 近 7 日淨增訂 (Pickup)")
        if not pace_daily:
            st.warning("尚未同步金旭 STLY Pace 快照，請確認每日 05:00 排程或於『05:00 同步狀態』分頁手動刷新。")
        else:
            df_p = pd.DataFrame(pace_daily)
            col_h1, col_h2 = st.columns([1.2, 2.8])
            with col_h1:
                show_days = st.slider("檢視未來天數", 14, min(90, len(df_p)), 42, key="sm_pace_days")
            df_pv = df_p.head(show_days).copy()
            stly_rn_col = "stly_dow_rn" if stly_mode == "dow" else "stly_date_rn"
            stly_adr_col = "stly_dow_adr" if stly_mode == "dow" else "stly_date_adr"
            yoy_rn_col = "yoy_dow_rn" if stly_mode == "dow" else "yoy_date_rn"
            yoy_adr_col = "yoy_dow_adr" if stly_mode == "dow" else "yoy_date_adr"
            final_rn_col = "ly_dow_final_rn" if stly_mode == "dow" else "ly_date_final_rn"

            df_pv["日期標籤"] = df_pv["date"].str[5:] + "(" + df_pv["dow"] + ")"

            fig_pace = go.Figure()
            fig_pace.add_trace(go.Bar(
                x=df_pv["日期標籤"], y=df_pv["sold_rn"],
                name="今年在手 OTB (房晚)", marker_color="#2563eb", opacity=0.82,
            ))
            fig_pace.add_trace(go.Scatter(
                x=df_pv["日期標籤"], y=df_pv[stly_rn_col],
                name="去年今日 STLY (房晚)", mode="lines+markers",
                line=dict(color="#f59e0b", width=2.5, dash="dot"),
            ))
            fig_pace.add_trace(go.Scatter(
                x=df_pv["日期標籤"], y=df_pv[final_rn_col],
                name="去年最終結算 (房晚)", mode="lines",
                line=dict(color="#94a3b8", width=1.5, dash="dash"),
            ))
            fig_pace.update_layout(
                height=360, margin=dict(l=20, r=20, t=25, b=40),
                legend=dict(orientation="h", y=1.12), hovermode="x unified",
                yaxis_title="房晚數 (RN)",
            )
            st.plotly_chart(fig_pace, use_container_width=True)

            # ADR 剪刀差 & 7 日淨 Pickup 雙圖
            c_ch1, c_ch2 = st.columns(2)
            with c_ch1:
                fig_adr = go.Figure()
                fig_adr.add_trace(go.Scatter(x=df_pv["日期標籤"], y=df_pv["otb_adr"], name="今年 OTB ADR", line=dict(color="#10b981", width=3)))
                fig_adr.add_trace(go.Scatter(x=df_pv["日期標籤"], y=df_pv[stly_adr_col], name="去年今日 STLY ADR", line=dict(color="#ef4444", width=2, dash="dot")))
                fig_adr.update_layout(title="今年在手 ADR vs 去年今日 STLY ADR", height=300, margin=dict(l=15, r=15, t=35, b=30), legend=dict(orientation="h", y=1.12), hovermode="x unified")
                st.plotly_chart(fig_adr, use_container_width=True)
            with c_ch2:
                colors_pk = ["#10b981" if v >= 0 else "#ef4444" for v in df_pv["pickup_7d_rn"]]
                fig_pk = go.Figure(go.Bar(x=df_pv["日期標籤"], y=df_pv["pickup_7d_rn"], marker_color=colors_pk, name="7日淨Pickup"))
                fig_pk.update_layout(title="近 7 日淨增訂 (Net Pickup 房晚)", height=300, margin=dict(l=15, r=15, t=35, b=30), hovermode="x unified")
                st.plotly_chart(fig_pk, use_container_width=True)

            st.divider()
            # 月底達標價量試算器
            st.markdown("#### 🎯 月度營收達標「價量試算器」（結合已結算實績 ＋ 金旭 OTB ＋ SiteMinder 掛牌價）")
            df_occ = pd.read_csv(OCC_CSV_PATH) if os.path.exists(OCC_CSV_PATH) else pd.DataFrame()
            if not df_occ.empty:
                df_occ["date"] = df_occ["date"].astype(str).str[:10]
                df_occ["ym"] = df_occ["date"].str[:7]
                cur_ym = datetime.date.today().strftime("%Y-%m")
                avail_yms = sorted([ym for ym in df_occ["ym"].unique() if ym >= cur_ym])[:4]
                if avail_yms:
                    sc1, sc2, sc3 = st.columns([1, 1.2, 1.2])
                    with sc1:
                        sel_ym = st.selectbox("選擇試算月份", avail_yms, index=min(1, len(avail_yms) - 1), key="sm_sim_ym")
                    # 嘗試從 Google Sheets targets 帶入預設月目標
                    def_target = 16000000
                    if targets_df is not None and not targets_df.empty:
                        try:
                            t_copy = targets_df.copy()
                            t_copy.columns = [str(c).strip().lower() for c in t_copy.columns]
                            m_col = next((c for c in t_copy.columns if "month" in c or "date" in c or "年" in c), None)
                            r_col = next((c for c in t_copy.columns if "target" in c or "rev" in c or "目標" in c), None)
                            if m_col and r_col:
                                hit = t_copy[t_copy[m_col].astype(str).str.replace("/", "-").str.startswith(sel_ym)]
                                if not hit.empty:
                                    v_t = pd.to_numeric(str(hit.iloc[0][r_col]).replace(",", ""), errors="coerce")
                                    if pd.notna(v_t) and v_t > 0:
                                        def_target = int(round(v_t))
                        except Exception:
                            pass
                    with sc2:
                        target_rev = st.number_input("該月客房營收目標 (TWD)", min_value=1000000, max_value=50000000, value=def_target, step=500000, key="sm_sim_tgt")
                    with sc3:
                        sell_through = st.slider("預估剩餘空房去化率 (%)", 10, 100, 65, 5, key="sm_sim_st")

                    m_df = df_occ[df_occ["ym"] == sel_ym]
                    cur_rev = int(round(pd.to_numeric(m_df["revenue"], errors="coerce").fillna(0).sum()))
                    cur_sold = int(round(pd.to_numeric(m_df["sold_rooms"], errors="coerce").fillna(0).sum()))
                    tot_cap = int(round(pd.to_numeric(m_df["total_rooms"], errors="coerce").fillna(126).sum()))
                    today_iso = datetime.date.today().strftime("%Y-%m-%d")
                    fut_m = m_df[m_df["date"] >= today_iso]
                    rem_fut_rooms = int(round((pd.to_numeric(fut_m["total_rooms"], errors="coerce").fillna(126) - pd.to_numeric(fut_m["sold_rooms"], errors="coerce").fillna(0)).clip(lower=0).sum()))

                    # 取得該月未來日期之 SiteMinder 主力雙人房 (SP/PD) 均價作為參考掛牌價
                    sm_ref_prices = []
                    for d_s in fut_m["date"].tolist():
                        rms = inv_map.get(d_s, {}).get("rooms", {})
                        vals = [rms[k]["r"] for k in ("SP", "PD", "DT") if k in rms and rms[k].get("r", 0) > 0]
                        if vals:
                            sm_ref_prices.append(np.mean(vals))
                    sm_bar_avg = int(round(np.mean(sm_ref_prices))) if sm_ref_prices else (int(round(cur_rev / cur_sold)) if cur_sold > 0 else 5500)
                    est_add_rn = int(round(rem_fut_rooms * (sell_through / 100.0)))
                    est_final_rev = cur_rev + est_add_rn * sm_bar_avg
                    gap_rev = target_rev - cur_rev
                    need_adr = int(round(gap_rev / est_add_rn)) if (gap_rev > 0 and est_add_rn > 0) else 0

                    mc1, mc2, mc3, mc4 = st.columns(4)
                    mc1.metric(f"{sel_ym} 目前已入帳＋OTB", f"NT$ {cur_rev:,}", f"已售 {cur_sold:,} / {tot_cap:,} 間")
                    mc2.metric("距離月目標尚缺", f"NT$ {max(0, gap_rev):,}", "已達標 🎉" if gap_rev <= 0 else f"未來尚有 {rem_fut_rooms:,} 空房")
                    mc3.metric("現行 SiteMinder 主力掛牌均價", f"NT$ {sm_bar_avg:,}", f"去化 {sell_through}% = +{est_add_rn:,} 房晚")
                    mc4.metric("預估月底總營收", f"NT$ {est_final_rev:,}", f"{'超標' if est_final_rev >= target_rev else '缺口'} NT$ {est_final_rev - target_rev:+,}")
                    if gap_rev > 0 and need_adr > 0:
                        st.info(
                            f"💡 **參謀建議**：若剩餘空房去化率為 **{sell_through}%（{est_add_rn:,} 房晚）**，"
                            f"剩餘房源實收均價須達 **NT$ {need_adr:,}** 才能達標（目前 SiteMinder 主力掛牌均價約 **NT$ {sm_bar_avg:,}**，扣抽成約 NT$ {int(sm_bar_avg*0.85):,}）。"
                        )

    # ------------------------------------------------------------------
    # TAB 4: 通路真實淨利排行 (Net ADR) & 洗盤取消率雷達
    # ------------------------------------------------------------------
    with tab4:
        st.subheader("💰 通路真實淨利排行（拆穿「預付淨價 vs 到付毛價」盲區）")
        st.caption(
            "資料來源：每日 05:00 金旭 HT2000 (HTRV＋HTAC＋HTGR) 100% 唯讀全自動同步。"
            "SiteMinder 預設把「Booking.com 到付毛價」與「Agoda/Trip 預付淨價」混在同一張表比 ADR，本表依金旭 HTGR 付款別還原「飯店實拿 Net ADR」。"
        )
        windows = chan_intel.get("windows", {})
        if not windows:
            st.warning("尚未產生金旭通路淨利情報快照。")
        else:
            win_keys = list(windows.keys())
            win_labels = {k: windows[k]["label"] for k in win_keys}
            sel_win = st.selectbox("選擇統計區間與日期視角", win_keys, format_func=lambda k: win_labels[k], key="sm_chan_win")
            view_level = st.radio("檢視層級", ["按通路彙總 (合併預付與到付)", "展開明細 (區分預付 Net vs 到付 Gross 代碼)"], horizontal=True, key="sm_chan_lvl")

            df_ch = _calc_net_adr_df(windows[sel_win]["rows"])
            if not df_ch.empty:
                tot_rn_w = df_ch["net_rn"].sum()
                if "彙總" in view_level:
                    g = df_ch.groupby("channel", as_index=False).agg(
                        淨訂單=("net_bk", "sum"),
                        取消筆數=("cx_bk", "sum"),
                        NoShow=("nosh_bk", "sum"),
                        淨房晚=("net_rn", "sum"),
                        金旭帳面營收=("net_rev", "sum"),
                        飯店實拿淨營收=("飯店實拿淨營收", "sum"),
                        ALOS=("alos", "mean"),
                        提前預訂天=("lead_days", "mean"),
                    )
                    g["房晚佔比%"] = np.where(tot_rn_w > 0, (g["淨房晚"] / tot_rn_w * 100.0).round(1), 0.0)
                    g["帳面 ADR"] = np.where(g["淨房晚"] > 0, (g["金旭帳面營收"] / g["淨房晚"]).round(0), 0).astype(int)
                    g["飯店實拿 Net ADR"] = np.where(g["淨房晚"] > 0, (g["飯店實拿淨營收"] / g["淨房晚"]).round(0), 0).astype(int)
                    g["取消率%"] = np.where((g["淨訂單"] + g["取消筆數"] + g["NoShow"]) > 0, ((g["取消筆數"] + g["NoShow"]) / (g["淨訂單"] + g["取消筆數"] + g["NoShow"]) * 100.0).round(1), 0.0)
                    g["ALOS"] = g["ALOS"].round(2)
                    g["提前預訂天"] = g["提前預訂天"].round(1)
                    g = g.sort_values("淨房晚", ascending=False).reset_index(drop=True)
                    g = g.rename(columns={"channel": "通路"})

                    st.dataframe(
                        g[["通路", "淨房晚", "房晚佔比%", "帳面 ADR", "飯店實拿 Net ADR", "取消率%", "淨訂單", "取消筆數", "ALOS", "提前預訂天", "飯店實拿淨營收"]].style.format({
                            "淨房晚": "{:,}", "房晚佔比%": "{:.1f}%", "帳面 ADR": "${:,}", "飯店實拿 Net ADR": "${:,}",
                            "取消率%": "{:.1f}%", "淨訂單": "{:,}", "取消筆數": "{:,}", "ALOS": "{:.2f}", "提前預訂天": "{:.1f}天", "飯店實拿淨營收": "${:,}",
                        }),
                        use_container_width=True,
                        hide_index=True,
                    )

                    # 氣泡圖：實拿 Net ADR vs 取消率（氣泡大小 = 房晚）
                    g_plot = g[g["淨房晚"] >= 5]
                    if not g_plot.empty:
                        fig_bub = px.scatter(
                            g_plot, x="取消率%", y="飯店實拿 Net ADR", size="淨房晚", color="通路",
                            hover_name="通路", text="通路", size_max=55,
                            title="各通路獲利品質象限：實拿 Net ADR（越高越好）vs 取消率（越低越好）",
                        )
                        fig_bub.update_traces(textposition="top center")
                        fig_bub.update_layout(height=390, margin=dict(l=20, r=20, t=40, b=20))
                        st.plotly_chart(fig_bub, use_container_width=True)
                else:
                    df_d = df_ch.rename(columns={
                        "code": "金旭代碼", "sub_label": "細分通路", "net_rn": "淨房晚",
                        "raw_adr": "金旭帳面ADR", "cancel_rate_pct": "取消率%", "alos": "ALOS", "lead_days": "提前預訂天",
                    })
                    st.dataframe(
                        df_d[["金旭代碼", "細分通路", "口徑", "淨房晚", "金旭帳面ADR", "飯店實拿 Net ADR", "取消率%", "ALOS", "提前預訂天", "飯店實拿淨營收"]].style.format({
                            "淨房晚": "{:,}", "金旭帳面ADR": "${:,}", "飯店實拿 Net ADR": "${:,}",
                            "取消率%": "{:.1f}%", "ALOS": "{:.2f}", "提前預訂天": "{:.1f}天", "飯店實拿淨營收": "${:,}",
                        }),
                        use_container_width=True,
                        hide_index=True,
                    )

            st.divider()
            # 未來 60 天「退訂洗盤預警 (Wash Risk)」
            st.markdown("#### 🌊 未來 60 天「高取消通路洗盤風險預警」（連動房務排班／備貨／壁紙壓房）")
            wash_rows = chan_intel.get("daily_wash_60d", [])
            if wash_rows:
                df_w = pd.DataFrame(wash_rows)
                fig_w = go.Figure()
                fig_w.add_trace(go.Bar(x=df_w["date"], y=df_w["high_cx_rn"], name="高取消通路在手 (Agoda+Expedia)", marker_color="#f97316"))
                fig_w.add_trace(go.Bar(x=df_w["date"], y=df_w["trip_rn"], name="Trip.com", marker_color="#3b82f6"))
                fig_w.add_trace(go.Bar(x=df_w["date"], y=df_w["booking_rn"], name="Booking.com", marker_color="#1d4ed8"))
                fig_w.add_trace(go.Bar(x=df_w["date"], y=df_w["direct_rn"], name="官網 Direct", marker_color="#10b981"))
                fig_w.add_trace(go.Scatter(x=df_w["date"], y=df_w["est_wash_rn"], name="預估潛在洗盤退訂間數", mode="lines", line=dict(color="#ef4444", width=2.5)))
                fig_w.update_layout(barmode="stack", height=350, margin=dict(l=20, r=20, t=20, b=30), legend=dict(orientation="h", y=1.12), hovermode="x unified", yaxis_title="房晚數")
                st.plotly_chart(fig_w, use_container_width=True)

                high_risk = df_w[df_w["high_cx_share_pct"] >= 35.0]
                if not high_risk.empty:
                    st.warning(
                        f"⚠️ 未來 60 天共有 **{len(high_risk)} 天** 的在手訂單中，高取消通路（Agoda＋Expedia）佔比 ≥ 35%，"
                        f"平均預估潛在洗盤退訂約 **{high_risk['est_wash_rn'].mean():.1f} 間/日**，排班與備貨建議保留彈性緩衝。"
                    )

    # ------------------------------------------------------------------
    # TAB 5: 05:00 全自動同步狀態與本機即戰工具
    # ------------------------------------------------------------------
    with tab5:
        st.subheader("⚙️ 每日 05:00 全自動同步狀態（黃金混合架構）")
        c_s1, c_s2 = st.columns(2)
        with c_s1:
            st.markdown("##### 1️⃣ 金旭 HT2000 唯讀 SQL（05:00 100% 無人值守）")
            st.success(f"✅ 通路淨利與取消情報：`{chan_intel.get('as_of', '未同步')}`")
            st.success(f"✅ 未來 90 天 STLY Pace & Pickup：`{pace_data.get('as_of', '未同步')}`")
            st.success(f"✅ 12 房型每日剩餘庫存：`{rt_asof}`")
            st.caption("已整合進 `run_daily_sync.py` 每日清晨 05:00 自動執行並推播 GitHub / Streamlit Cloud，0 登入過期問題。")
        with c_s2:
            st.markdown("##### 2️⃣ SiteMinder 掛牌價快照（05:00 順路同步 ＋ 日間一鍵更新）")
            st.info(f"🌐 最新有效掛牌快照：`{sm_asof}`（涵蓋 {n_inv_days} 天）")
            if sm_status:
                st.caption(f"最近一次嘗試：`{sm_status.get('last_attempt_at','-')}`｜狀態：`{sm_status.get('status','-')}`｜說明：{sm_status.get('detail','-')}")
            st.markdown(
                """
- **05:00 無彈窗全自動用法**：平常雙擊專案內 `launch_siteminder_chrome.bat` 開啟專用 Chrome 視窗（Port `9223`）並登入 SiteMinder 保持開啟，05:00 排程即可靜默唯讀抓完最新 42 天掛牌價推播雲端。
- **日間隨時手動更新**：雙擊 `sync_siteminder_now.bat`，15 秒內唯讀抓取最新掛牌價並自動推上雲端版。
- **安全保證**：若半夜 SiteMinder 自動登出或電腦未開 Chrome，系統會自動保留最後一份有效快照，絕不報錯、絕不中斷 05:00 其他報表同步。
"""
            )

        st.divider()
        b1, b2 = st.columns(2)
        with b1:
            if st.button("🔄 立即直連金旭 PMS 重算最新通路淨利與 Pace（僅飯店區網本機可用）", use_container_width=True):
                with st.spinner("正在唯讀查詢金旭 HT2000 ..."):
                    try:
                        import pms_revenue_intel
                        res = pms_revenue_intel.run_pms_revenue_intel_sync()
                        st.success(f"✅ 同步完成：{res}")
                        st.rerun()
                    except Exception as e:
                        st.warning(f"雲端或非飯店區網環境無法直連 SQL Server（已自動使用最新快照）：{e}")
        with b2:
            if st.button("🌐 立即唯讀刷新本機 Chrome SiteMinder 掛牌價（僅本機開啟 Chrome 可用）", use_container_width=True):
                with st.spinner("正在透過本機 CDP 唯讀擷取 SiteMinder 42 天掛牌價..."):
                    try:
                        import siteminder_cdp_sync
                        res = siteminder_cdp_sync.run_siteminder_sync(unattended=False, windows_count=3)
                        if res.get("status") == "success":
                            st.success(f"✅ {res.get('detail')}")
                            st.rerun()
                        else:
                            st.warning(f"⚠️ {res.get('detail')}")
                    except Exception as e:
                        st.warning(f"目前非本機 Chrome 環境（已保留最新快照）：{e}")
