# -*- coding: utf-8 -*-
"""
路徒 PLUS 行旅（站前館）SiteMinder 定價稽核雷達 × 金旭 PMS 收益情報中樞
==========================================================================
100% 唯讀分析模組（嚴禁修改任何房價、庫存或設定）。
直指飯店經理最核心的 4 大管理提問（不搞花俏複雜切換）：
  1. 收益部放的掛牌價合不合理？為什麼這天高、那天低？（連動住房率、剩房、去年同期 Pace、台北活動與倒掛掃描）
  2. 這時候到底做了什麼促銷？（比對 SiteMinder 牌告價 vs 金旭實際進單成交價，抓出深折與疊加促銷）
  3. 同產品（同房型）的「含早 vs 不含早」差價差多少錢？（SiteMinder 設定加價 vs 金旭同日同房型實收差價）
  4. 各通路扣除抽成後的真實 Net ADR 與退訂洗盤風險？
"""

import os
import io
import json
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
SM_HISTORY_DIR = os.path.join(BACKUP_DIR, "sm_history")
CHAN_INTEL_PATH = os.path.join(BACKUP_DIR, "pms_channel_intel.json")
STLY_PACE_PATH = os.path.join(BACKUP_DIR, "pms_stly_pace.json")
PROMO_BF_PATH = os.path.join(BACKUP_DIR, "pms_promo_bf_audit.json")
DYN_YIELD_PATH = os.path.join(BACKUP_DIR, "pms_dynamic_yield.json")
RT_INV_PATH = os.path.join(BACKUP_DIR, "wallpaper_roomtype_inventory.json")
OCC_CSV_PATH = os.path.join(BACKUP_DIR, "station_occ_data.csv")

# 依金旭 HTRT 定義之房型階梯順序（高 → 低）
RT_TIER_ORDER = ["DF", "F", "DTP", "LDT", "DT", "E", "PD", "PT", "SP", "S"]
RT_NAMES = {
    "DF": "豪華家庭房 DF (4人)",
    "F": "家庭房 F (4人)",
    "DTP": "豪華三人房 DTP (3人)",
    "LDT": "奢華雙人房 LDT (2人)",
    "DT": "豪華雙人房 DT (2人)",
    "E": "101行政 E(05) (2人)",
    "ER": "101行政 ER (未掛OTA)",
    "PD": "尊爵雙人房 PD (2人)",
    "PT": "尊爵雙床房 PT (2人)",
    "SP": "高級雙人房 SP (無含早)",
    "S": "標準雙人房 S (無含早)",
}

# 分床型產品線階梯（區分「一大床線」與「雙床/多人線」，避免將雙小床 PT 與單大床 E 跨床型誤判為倒掛）
TIER_PAIRS = [
    # 多人／雙床線 (Multi-pax / 2 Beds)
    ("DF", "F"),
    ("F", "DTP"),
    ("DTP", "DT"),
    ("DT", "PT"),
    # 一大床線 (1 Queen/King Bed)
    ("LDT", "E"),
    ("E", "PD"),
    ("PD", "SP"),
    ("SP", "S"),
    # 頂規三人 vs 奢華雙人
    ("DTP", "LDT"),
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
.sm-hero h2 { margin: 0 0 6px 0; font-size: 1.45rem; color: #ffffff; }
.sm-hero p { margin: 0; color: #cbd5e1; font-size: 0.92rem; line-height: 1.5; }
.sm-badge {
    display: inline-block; padding: 3px 10px; border-radius: 999px;
    font-size: 0.78rem; font-weight: 700; margin-right: 6px; margin-top: 8px;
}
.sm-badge-ok { background: rgba(16,185,129,0.18); color: #6ee7b7; border: 1px solid rgba(16,185,129,0.5); }
.sm-badge-warn { background: rgba(245,158,11,0.20); color: #fde68a; border: 1px solid rgba(245,158,11,0.55); }
.sm-badge-err { background: rgba(239,68,68,0.22); color: #fca5a5; border: 1px solid rgba(239,68,68,0.6); }
.sm-kpi {
    background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px;
    padding: 14px 16px; box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05); height: 100%;
}
.sm-kpi-title { font-size: 0.82rem; color: #64748b; font-weight: 600; margin-bottom: 4px; }
.sm-kpi-val { font-size: 1.55rem; color: #0f172a; font-weight: 800; line-height: 1.2; font-variant-numeric: tabular-nums; }
.sm-kpi-sub { font-size: 0.78rem; color: #475569; margin-top: 5px; line-height: 1.35; }
.sm-alert-box {
    background: #fef2f2; border: 1px solid #fecaca; border-left: 5px solid #ef4444;
    border-radius: 10px; padding: 12px 16px; margin-bottom: 14px; color: #7f1d1d; font-size: 0.9rem; line-height: 1.5;
}
.sm-dcard {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 13px 16px 11px 16px;
    margin-bottom: 6px;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
}
.sm-dcard-locked {
    background: #f8fafc;
    border: 1px dashed #cbd5e1;
    opacity: 0.92;
}
.sm-dcard-top {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 4px;
    flex-wrap: wrap;
}
.sm-dpill {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 0.74rem;
    font-weight: 700;
    white-space: nowrap;
}
.sm-dtitle {
    font-size: 0.98rem;
    font-weight: 800;
    color: #0f172a;
}
.sm-dsum {
    font-size: 0.85rem;
    color: #475569;
    line-height: 1.4;
}
.sm-lane-hdr {
    font-size: 0.96rem;
    font-weight: 800;
    color: #1e293b;
    padding: 6px 12px;
    background: #f1f5f9;
    border-radius: 8px;
    margin-bottom: 10px;
}
</style>
""",
        unsafe_allow_html=True,
    )


def _render_card_detail_body(card: dict, ack_state: dict):
    """渲染單一決策卡的完整 Drill-Down 詳情內容（共用予 Dialog 彈窗或降級展開區）。"""
    import decision_engine

    if card.get("detail_intro"):
        st.markdown(card["detail_intro"])

    metrics = card.get("metrics") or []
    if metrics:
        m_cols = st.columns(len(metrics))
        for col, (lbl, val) in zip(m_cols, metrics):
            col.metric(lbl, val)

    df_d = card.get("detail_df")
    if df_d is not None and isinstance(df_d, pd.DataFrame) and not df_d.empty:
        st.markdown("##### 📋 相關異常明細")
        st.dataframe(df_d, use_container_width=True, hide_index=True, height=min(320, 44 + 35 * len(df_d)))

    actions = card.get("actions") or []
    if actions:
        st.markdown("##### 🎯 建議行動方案")
        for idx, act in enumerate(actions, 1):
            st.markdown(f"**{idx}.** {act}")

    copy_txt = card.get("copy_text") or ""
    if copy_txt:
        st.markdown("##### 💬 一鍵複製核對話術（可直接貼 LINE / 會議紀錄）")
        st.code(copy_txt, language="text")

    if not card.get("is_locked"):
        st.divider()
        b1, b2 = st.columns(2)
        cid = card["id"]
        with b1:
            if st.button("✅ 標記已處理 / 已交辦", key=f"dlg_done_{cid}", type="primary", use_container_width=True):
                ack_state[cid] = {"status": "done", "at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "title": card["title"]}
                st.session_state["sm_decision_ack"] = ack_state
                decision_engine.save_ack_state(ack_state)
                st.rerun()
        with b2:
            if st.button("🔕 刻意為之（暫時隱藏）", key=f"dlg_ign_{cid}", use_container_width=True):
                ack_state[cid] = {"status": "ignored", "at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "title": card["title"]}
                st.session_state["sm_decision_ack"] = ack_state
                decision_engine.save_ack_state(ack_state)
                st.rerun()


if hasattr(st, "dialog"):
    @st.dialog("🧭 異常細節與行動方案", width="large")
    def _open_decision_dialog(card: dict, ack_state: dict):
        st.markdown(f"### {card['badge']}　{card['title']}")
        _render_card_detail_body(card, ack_state)
else:
    def _open_decision_dialog(card: dict, ack_state: dict):
        st.session_state["sm_inline_open_card"] = card["id"]


def _render_compact_card(card: dict, ack_state: dict):
    """最前端極簡卡片：僅呈現「燈號 + 短標題 + 1 行量化結論 + 查看詳情按鈕」。"""
    sev_style = {
        "critical": ("#ef4444", "#fef2f2", "#b91c1c"),
        "warning":  ("#f97316", "#fff7ed", "#c2410c"),
        "info":     ("#3b82f6", "#eff6ff", "#1d4ed8"),
        "verified": ("#10b981", "#ecfdf5", "#047857"),
        "locked":   ("#94a3b8", "#f1f5f9", "#475569"),
    }
    border_c, pill_bg, pill_fg = sev_style.get(card.get("severity", "info"), ("#3b82f6", "#eff6ff", "#1d4ed8"))
    lock_cls = " sm-dcard-locked" if card.get("is_locked") else ""

    c_txt, c_btn = st.columns([4.4, 1.35], vertical_alignment="center")
    with c_txt:
        st.markdown(
            f"""
<div class="sm-dcard{lock_cls}" style="border-left: 5px solid {border_c};">
  <div class="sm-dcard-top">
    <span class="sm-dpill" style="background:{pill_bg}; color:{pill_fg};">{card['badge']}</span>
    <span class="sm-dtitle">{card['title']}</span>
  </div>
  <div class="sm-dsum">{card['summary']}</div>
</div>
""",
            unsafe_allow_html=True,
        )
    with c_btn:
        btn_lbl = "🔓 解鎖方式" if card.get("is_locked") else "🔎 查看詳情"
        if st.button(btn_lbl, key=f"open_card_{card['id']}", use_container_width=True):
            _open_decision_dialog(card, ack_state)

    if not hasattr(st, "dialog") and st.session_state.get("sm_inline_open_card") == card["id"]:
        with st.container(border=True):
            _render_card_detail_body(card, ack_state)


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
    """針對未來 42 天 SiteMinder 掛牌價 × 金旭 STLY Pace × 12 房型剩房 × 台北活動執行 6 大定價稽核。"""
    inv_map = sm_snap.get("inventory", {})
    if not inv_map:
        return pd.DataFrame(), {}

    pace_by_date = {r["date"]: r for r in pace_data.get("daily", [])}
    pms_daily_rt = rt_inv.get("daily_avail", {})
    base_date_str = pace_data.get("base_date") or datetime.date.today().strftime("%Y-%m-%d")
    try:
        base_d = datetime.date.fromisoformat(base_date_str)
    except Exception:
        base_d = datetime.date.today()

    dow_zh = ["一", "二", "三", "四", "五", "六", "日"]
    fixed_holidays = {"12-24", "12-25", "12-29", "12-30", "12-31", "01-01", "01-02"}
    # 嚴格過濾掉已過去日期 (< base_date_str)，徹底杜絕歷史日查無未來 Pace 而補 0% 誤判！
    dates_sorted = [d for d in sorted(inv_map.keys()) if d >= base_date_str]
    alerts = []
    day_verdict = {}  # date -> (vs_dow_str, plain_chinese_reason)

    # 計算「一般常態同星期」主力雙人房 PD 的中位掛牌價（排除耶誕/跨年 12/23~01/02 避免基準線被旺季拉歪）
    dow_pd_rates = {i: [] for i in range(7)}
    dow_pd_all = {i: [] for i in range(7)}
    for d_str in dates_sorted:
        try:
            wd = datetime.date.fromisoformat(d_str).weekday()
        except Exception:
            continue
        pdr = inv_map[d_str].get("rooms", {}).get("PD", {}).get("r", 0)
        if pdr > 0:
            dow_pd_all[wd].append(pdr)
            mmdd = d_str[5:]
            if not ("12-23" <= mmdd <= "12-31" or "01-01" <= mmdd <= "01-02"):
                dow_pd_rates[wd].append(pdr)
    dow_pd_med = {
        wd: float(np.median(dow_pd_rates[wd])) if dow_pd_rates[wd] else (float(np.median(dow_pd_all[wd])) if dow_pd_all[wd] else 0.0)
        for wd in range(7)
    }

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
        # 固定日曆節慶 (耶誕/跨年) 強制對齊同月同日 (-365d)，避免把今年 12/30 錯比去年跨年夜 12/31！
        eff_mode = "date" if d_str[5:] in fixed_holidays else stly_mode
        yoy_rn = pace.get("yoy_dow_rn" if eff_mode == "dow" else "yoy_date_rn", 0)
        yoy_adr = pace.get("yoy_dow_adr" if eff_mode == "dow" else "yoy_date_adr", 0)
        stly_adr = pace.get("stly_dow_adr" if eff_mode == "dow" else "stly_date_adr", 0)
        otb_adr = pace.get("otb_adr", 0)
        occ_pct = pace.get("occ_pct", 0.0)
        pms_rem_tot = pace.get("rem_rn", None)
        ev_name = events_map.get(d_str, "")
        sm_tot_avail = sum(int(v.get("a", 0)) for v in rooms.values())

        day_inv_cnt = 0
        day_latent_cnt = 0
        day_clear_cnt = 0
        day_eq_cnt = 0
        day_overprice = False
        day_underprice = False

        # 規則 1 & 2：動態收益視角之房型階梯檢核（區分「雙有庫存真倒掛」vs「低階完售・高階合理清倉」vs「D+3↑潛伏退訂風險」）
        for hi, lo in TIER_PAIRS:
            if hi not in rooms or lo not in rooms:
                continue
            r_hi, a_hi = int(rooms[hi].get("r", 0)), int(rooms[hi].get("a", 0))
            r_lo, a_lo = int(rooms[lo].get("r", 0)), int(rooms[lo].get("a", 0))
            if r_hi <= 0 or r_lo <= 0:
                continue
            if r_lo > r_hi:
                is_nye_101 = allow_nye_101_exception and d_str == "2026-12-31" and hi == "DT" and lo == "E"
                diff = r_lo - r_hi
                if is_nye_101:
                    alerts.append({
                        "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🟡 提醒",
                        "異常類型": "🎆 景觀特殊溢價", "涉及房型": f"{lo} > {hi}",
                        "掛牌價/數據摘要": f"{lo} ${r_lo:,} vs {hi} ${r_hi:,}",
                        "說明與收益部提問建議": f"跨年夜 101 景觀房 E(05) 高於 DT +{diff:,}（煙火景觀溢價，屬合理操作）",
                        "台北活動": ev_name,
                    })
                elif a_lo == 0 and lead_d <= 2:
                    # 臨期 D-0~D-2 低階已完售(0間)，收益部調降高階清庫存 → 100% 合理動態去化，不列異常！
                    day_clear_cnt += 1
                elif a_lo == 0 and lead_d >= 3:
                    # 遠中程低階暫滿(0間)、高階降價低於低階原價 → 僅列提醒（防免費取消回庫變真倒掛），不報嚴重/警告！
                    day_latent_cnt += 1
                    alerts.append({
                        "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🟡 提醒",
                        "異常類型": "💤 潛伏價差(低階已滿)", "涉及房型": f"{lo}(0間) > {hi}",
                        "掛牌價/數據摘要": f"{lo} ${r_lo:,}(剩0) vs {hi} ${r_hi:,}(剩{a_hi})",
                        "說明與收益部提問建議": f"低階 {lo} 已滿(0間)、高階 {hi} 降價去化屬合理浮動操作；僅提醒距入住尚有 {lead_d} 天，若 {lo} 發生免費取消回庫將觸發倒掛。",
                        "台北活動": ev_name,
                    })
                else:
                    # a_lo > 0：低階仍有庫存卻賣得比高階貴
                    day_inv_cnt += 1
                    sev = "🔴 嚴重" if a_hi > 0 else "🟠 警告"
                    cat_lbl = "📉 雙有庫存掛牌倒掛" if a_hi > 0 else "📉 低階高於已滿高階"
                    avail_note = f"雙方皆有庫存（{lo}剩{a_lo}、{hi}剩{a_hi}），旅客搜尋直接看到小房比大房貴，流失高階溢價與升等主導權！" if a_hi > 0 else f"低階 {lo} 仍有 {a_lo} 間在售但高於已滿之 {hi}。"
                    alerts.append({
                        "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": sev,
                        "異常類型": cat_lbl, "涉及房型": f"{lo} > {hi}",
                        "掛牌價/數據摘要": f"{lo} ${r_lo:,}(剩{a_lo}) vs {hi} ${r_hi:,}(剩{a_hi})",
                        "說明與收益部提問建議": f"低階 {lo} 高於高階 {hi} +{diff:,} 元。{avail_note}",
                        "台北活動": ev_name,
                    })
            elif r_lo == r_hi and (a_hi > 0 and a_lo > 0):
                day_eq_cnt += 1
                alerts.append({
                    "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🟠 警告",
                    "異常類型": "⚖️ 跨階同價零價差", "涉及房型": f"{lo} ＝ {hi}",
                    "掛牌價/數據摘要": f"皆掛 ${r_hi:,}（庫存 {lo}:{a_lo} / {hi}:{a_hi}）",
                    "說明與收益部提問建議": f"{hi} 與低一階 {lo} 同價且雙雙有房，若低階跑得快建議拉開價差改由預排升等消化。",
                    "台北活動": ev_name,
                })

        # 規則 3：臨期滯銷卻偏高 (Overpricing)
        pd_rate = int(rooms.get("PD", {}).get("r", 0))
        med_pd = dow_pd_med.get(wd, 0.0)
        vs_dow_pct = round((pd_rate / med_pd - 1.0) * 100.0, 1) if (pd_rate > 0 and med_pd > 0) else 0.0

        if 0 <= lead_d <= last_min_days and sm_tot_avail >= 10 and vs_dow_pct >= 20.0:
            if yoy_rn <= -2 or occ_pct < 85.0:
                day_overprice = True
                alerts.append({
                    "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🟠 警告",
                    "異常類型": "❄️ 臨期高價滯銷", "涉及房型": "全館主力房",
                    "掛牌價/數據摘要": f"PD ${pd_rate:,}（高於平常週{dow_lbl} +{vs_dow_pct}%）｜剩{pms_rem_tot}間",
                    "說明與收益部提問建議": f"距入住僅 {lead_d} 天、金旭尚餘 {pms_rem_tot} 間（住房率 {occ_pct:.1f}%），掛牌價卻高於平常週{dow_lbl} +{vs_dow_pct}%，請確認為何未降價去化。",
                    "台北活動": ev_name,
                })

        # 規則 4：買氣強勁卻賣比去年便宜 (Underpricing)
        if pace and yoy_rn >= underprice_pace_rn and yoy_adr <= -150 and stly_adr > 0:
            day_underprice = True
            adr_pct = round(yoy_adr / stly_adr * 100.0, 1)
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🟠 警告",
                "異常類型": "🔥 買氣強卻賣便宜", "涉及房型": "全館均價",
                "掛牌價/數據摘要": f"Pace {yoy_rn:+d}晚｜成交ADR ${otb_adr:,} (YoY {adr_pct:+.1f}%)",
                "說明與收益部提問建議": f"在手房晚已領先去年同期 +{yoy_rn} 晚（住房率 {occ_pct:.1f}%），成交均價卻低於去年 ${abs(yoy_adr):,}，建議檢視是否提早放了過多低價促銷。",
                "台北活動": ev_name,
            })

        # 規則 5：幽靈關房 / 跨年高 MinLOS 尾房檢核
        if ss_flag == "Y" and (pms_rem_tot is not None and pms_rem_tot >= 4):
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🔴 嚴重",
                "異常類型": "🚫 疑似幽靈關房", "涉及房型": "全館",
                "掛牌價/數據摘要": f"Stop Sell = Y｜金旭剩餘 {pms_rem_tot} 間",
                "說明與收益部提問建議": "金旭 PMS 仍有空房但 SiteMinder 設為 Stop Sell，請立即確認是否誤鎖！",
                "台北活動": ev_name,
            })
        if mlos >= 3 and sm_tot_avail >= 4 and lead_d <= 45:
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🟡 提醒",
                "異常類型": "🔒 高連住門檻尾房檢核", "涉及房型": "全館",
                "掛牌價/數據摘要": f"MinLOS = {mlos} 晚｜OTA 剩 {sm_tot_avail} 間",
                "說明與收益部提問建議": f"設最低連住 {mlos} 晚且尚有 {sm_tot_avail} 間尾房，建議追蹤去化速度避免臨近賣不掉。",
                "台北活動": ev_name,
            })

        # 規則 6：Fat-Finger 暴跌防呆（須同時低於絕對地板價 <$3,200，避免把連假收尾 $7,422 誤判成少打一個零）
        if vs_dow_pct <= -fat_finger_drop_pct and 0 < pd_rate < 3200:
            alerts.append({
                "日期": f"{d_str} ({dow_lbl})", "date": d_str, "嚴重度": "🔴 嚴重",
                "異常類型": "⚡ 疑誤輸入暴跌", "涉及房型": "PD",
                "掛牌價/數據摘要": f"PD ${pd_rate:,}（較平常週{dow_lbl} {vs_dow_pct:.1f}%）",
                "說明與收益部提問建議": "掛牌價異常暴跌，請立即核對是否誤輸入！",
                "台北活動": ev_name,
            })

        if d_str == dates_sorted[0]:
            er_rem = pms_daily_rt.get(d_str, {}).get("ER", {}).get("rem", None)
            if er_rem is not None:
                alerts.append({
                    "日期": "全區間結構", "date": "0000-00-00", "嚴重度": "🟡 提醒",
                    "異常類型": "🧩 金旭 vs SiteMinder 房型缺口", "涉及房型": "ER (8間) / SD (3間)",
                    "掛牌價/數據摘要": f"SiteMinder 僅掛 10 房型｜今日 ER 剩 {er_rem} 間",
                    "說明與收益部提問建議": "確認 ER (8間) 是併入 E 銷售、保留升等還是漏上架。",
                    "台北活動": "",
                })

        # 產生每日白話「為何高／為何低、合不合理」診斷結論（直接放 Tab 2！）
        vs_str = f"{vs_dow_pct:+.0f}%"
        reasons = []
        if day_inv_cnt > 0:
            reasons.append(f"🔴 雙有房倒掛({day_inv_cnt}處)")
        if day_clear_cnt > 0:
            reasons.append(f"✅ 低階完售・高階清倉({day_clear_cnt}處)")
        elif day_latent_cnt > 0:
            reasons.append(f"💤 低階滿・高階降價({day_latent_cnt}處)")
        if day_eq_cnt > 0:
            reasons.append(f"🟠 同價({day_eq_cnt}處)")
        if day_overprice:
            reasons.append(f"❄️ 偏高(剩{pms_rem_tot}間卻高於平日)")
        elif day_underprice:
            reasons.append(f"🔥 偏低(領先+{yoy_rn}晚但ADR降)")
        elif d_str in ("2026-12-31", "2027-01-01"):
            reasons.append(f"✅ 跨年高價(MinLOS={mlos})")
        elif d_str in ("2026-12-24", "2026-12-25", "2026-12-26"):
            reasons.append("✅ 耶誕檔期拉價")
        elif ev_name:
            reasons.append(f"✅ 活動支撐({ev_name[:10]})")
        elif occ_pct >= 88.0 and vs_dow_pct >= 10:
            reasons.append(f"✅ 合理拉價(住房率{occ_pct:.0f}%)")
        elif vs_dow_pct >= 25 and occ_pct < 70:
            reasons.append(f"🟡 掛價偏高({vs_str}｜住房率{occ_pct:.0f}%)")
        elif vs_dow_pct <= -15 and occ_pct >= 75:
            reasons.append(f"🟡 掛價偏保守({vs_str}｜住房率{occ_pct:.0f}%)")
        elif not reasons:
            reasons.append("✅ 區間合理")
        day_verdict[d_str] = (vs_str, " ｜ ".join(reasons))

    df_alerts = pd.DataFrame(alerts)
    if not df_alerts.empty:
        sev_rank = {"🔴 嚴重": 1, "🟠 警告": 2, "🟡 提醒": 3}
        df_alerts["_rank"] = df_alerts["嚴重度"].map(sev_rank).fillna(9)
        df_alerts = df_alerts.sort_values(["_rank", "date"]).drop(columns=["_rank"]).reset_index(drop=True)
    return df_alerts, day_verdict


def build_excel_audit_report(df_alerts, df_rates, df_bf, df_promo, df_bf_inv=None):
    """產出 5 分頁專業定價稽核 Excel (.xlsx) 供經理會議對質與存查。"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (df_alerts if not df_alerts.empty else pd.DataFrame([{"訊息": "目前無觸發異常"}])).to_excel(
            writer, sheet_name="1_定價異常對質清單", index=False
        )
        (df_rates if not df_rates.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="2_42天定價合理性矩陣", index=False
        )
        (df_bf if not df_bf.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="3_含早vs不含早差價", index=False
        )
        (df_bf_inv if df_bf_inv is not None and not df_bf_inv.empty else pd.DataFrame([{"訊息": "無含早倒掛實單"}])).to_excel(
            writer, sheet_name="4_含早倒掛抓包實單", index=False
        )
        (df_promo if not df_promo.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="5_促銷折數反查明細", index=False
        )
        for ws in writer.book.worksheets:
            for col in ws.columns:
                max_len = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 4, 12), 46)
    return buf.getvalue()


def render_siteminder_audit(taipei_events_df=None, targets_df=None):
    """渲染『📡 專案：SiteMinder 定價稽核雷達』獨立互動模組（站前館專屬）。"""
    _inject_module_css()

    sm_snap = _load_json_safe(SM_SNAPSHOT_PATH)
    sm_status = _load_json_safe(SM_STATUS_PATH)
    chan_intel = _load_json_safe(CHAN_INTEL_PATH)
    pace_data = _load_json_safe(STLY_PACE_PATH)
    promo_bf = _load_json_safe(PROMO_BF_PATH)
    dyn_yield = _load_json_safe(DYN_YIELD_PATH)
    rt_inv = _load_json_safe(RT_INV_PATH)

    health = sm_snap.get("health", {})
    sm_asof = sm_snap.get("as_of", "尚未擷取")
    pms_asof = pace_data.get("as_of") or chan_intel.get("as_of") or "尚未同步"
    rt_asof = rt_inv.get("as_of", "尚未同步")
    n_inv_days = len(sm_snap.get("inventory", {}))

    fail_cnt = int(health.get("delivery_failures_today", 0) or 0)
    mapped_n = health.get("rates_mapped", 94)
    mapped_tot = health.get("rates_total", 105)
    comp_cnt = int(health.get("competitors_configured", 0) or 0)

    events_map = {}
    if taipei_events_df is not None and not taipei_events_df.empty and "date" in taipei_events_df.columns:
        for _, er in taipei_events_df.iterrows():
            ds = str(er.get("date", "")).strip()[:10]
            en = str(er.get("event_name", "")).strip()
            if ds and en:
                events_map[ds] = f"{events_map[ds]} / {en}" if ds in events_map else en

    import decision_engine

    # 預先執行 6 大規則（使用預設門檻，下方進階設定可覆蓋）
    stly_mode_lbl = st.session_state.get("sm_stly_mode", "同星期比對 (-364天｜SiteMinder 預設)")
    stly_mode = "dow" if "-364" in stly_mode_lbl else "date"
    allow_nye_101 = st.session_state.get("sm_nye_101", True)
    last_min_days = st.session_state.get("sm_lastmin_d", 3)
    underprice_pace_rn = st.session_state.get("sm_up_rn", 5)

    df_alerts, day_verdict = run_pricing_audit_rules(
        sm_snap=sm_snap, pace_data=pace_data, rt_inv=rt_inv, events_map=events_map,
        stly_mode=stly_mode, last_min_days=last_min_days,
        underprice_pace_rn=underprice_pace_rn, allow_nye_101_exception=allow_nye_101,
    )

    bundle = decision_engine.build_daily_decision_bundle(
        sm_alerts_df=df_alerts,
        taipei_events_df=taipei_events_df,
    )
    dh = bundle["health"]

    # 1. 頂部精簡 Hero Banner（含即時時效閘門＋負庫存防爆房哨兵燈號）
    pms_badge_cls = "sm-badge-ok" if dh["pms_fresh"] else "sm-badge-warn"
    sm_badge_cls = "sm-badge-warn" if dh["sm_stale"] else "sm-badge-ok"
    sm_badge_txt = (
        f"🔒 SiteMinder 快照過期 {dh['sm_age_h']/24:.1f} 天（已上鎖防誤報）"
        if dh["sm_stale"]
        else f"🌐 SiteMinder 掛牌新鮮：{sm_asof}"
    )
    inv_badge_txt = f"✅ 房型校驗 {dh['rt_total_inv']} 間" if dh["rt_inv_ok"] else f"⚠️ 房型異常 ({dh['rt_total_inv']}間)"
    ov_cnt = int(dh.get("overbook_count", 0) or 0)
    l1_cnt = int(dh.get("last1_14d_count", 0) or 0)
    l1_noesc = int(dh.get("last1_14d_no_escape", 0) or 0)
    ov_badge_cls = "sm-badge-err" if ov_cnt > 0 else "sm-badge-ok"
    ov_badge_txt = f"🚨 負庫存爆房 {ov_cnt} 筆！" if ov_cnt > 0 else "✅ 全館負庫存 0 筆"
    l1_badge_cls = "sm-badge-warn" if l1_noesc > 0 else "sm-badge-ok"
    l1_badge_txt = (
        f"⚡ 頂規剩1間 {l1_cnt} 筆 ({l1_noesc}零退路)"
        if l1_cnt > 0
        else "✅ 近14天無頂規剩1間"
    )
    upg_60d = int(dh.get("upg_60d_rn", 0) or 0)
    er_pct = float(dh.get("er_share_pct", 0.0) or 0.0)
    upg_badge_txt = f"🔄 60天升等消化 {upg_60d} 晚 (ER池 {er_pct:.0f}%)" if upg_60d > 0 else "🔄 動態升等引擎就緒"
    st.markdown(
        f"""
<div class="sm-hero">
  <h2>📡 站前館營運決策與定價稽核雷達 <span style="font-size:0.75rem;background:#f59e0b;color:#0f172a;padding:2px 8px;border-radius:999px;vertical-align:middle;">v4.0 動態收益參謀</span></h2>
  <p>
    以「<b>浮動房價 × 低階升等導流</b>」全局視角自動剔除跨日調價與免費升等汙染；最前端收束為精簡結論，點選 <b>「🔎 查看詳情」</b> 即可展開明細、軌跡與一鍵 LINE 話術。
  </p>
  <div>
    <span class="sm-badge {pms_badge_cls}">🕒 金旭最新：{dh['pms_as_of']}</span>
    <span class="sm-badge {ov_badge_cls}">{ov_badge_txt}</span>
    <span class="sm-badge {l1_badge_cls}">{l1_badge_txt}</span>
    <span class="sm-badge sm-badge-ok">{upg_badge_txt}</span>
    <span class="sm-badge {sm_badge_cls}">{sm_badge_txt}</span>
    <span class="sm-badge sm-badge-ok">{inv_badge_txt}</span>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    # 2. 最前端極簡決策收束看板（Progressive Disclosure：只看結論，點擊展開詳情）
    if "sm_decision_ack" not in st.session_state:
        st.session_state["sm_decision_ack"] = decision_engine.load_ack_state()
    ack_state = st.session_state["sm_decision_ack"]

    active_pm = [c for c in bundle["postmortem_cards"] if c["id"] not in ack_state]
    active_int = [c for c in bundle["internal_cards"] if c["id"] not in ack_state]
    active_rev = [c for c in bundle["revenue_cards"] if c["id"] not in ack_state]
    acked_cards = [
        c for c in (bundle["postmortem_cards"] + bundle["internal_cards"] + bundle["revenue_cards"])
        if c["id"] in ack_state
    ]

    # 2a. 若有昨日/近期預警事後覆盤，以極簡單列置頂呈現
    if active_pm:
        for pm_c in active_pm:
            _render_compact_card(pm_c, ack_state)

    # 2b. 雙欄權責分流（左：館內今日拍板 ｜ 右：總部收益部鐵證對質）
    col_l, col_r = st.columns(2, gap="medium")
    with col_l:
        st.markdown(f'<div class="sm-lane-hdr">🏢 館內今日拍板（{len(active_int)} 項）</div>', unsafe_allow_html=True)
        if not active_int:
            st.success("✅ 今日館內無待處理異常！")
        else:
            for card in active_int:
                _render_compact_card(card, ack_state)

    with col_r:
        st.markdown(f'<div class="sm-lane-hdr">📡 總部收益部對質・僅留鐵證（{len(active_rev)} 項）</div>', unsafe_allow_html=True)
        if not active_rev:
            st.success("✅ 目前無待對質定價異常！")
        else:
            for card in active_rev:
                _render_compact_card(card, ack_state)

    # 2c. 已處理／已隱藏閉環收納列（不占主畫面）
    if acked_cards:
        with st.expander(f"✅ 已處理 / 已暫時隱藏項目（{len(acked_cards)} 項，點此檢視或還原）", expanded=False):
            for ac in acked_cards:
                meta = ack_state.get(ac["id"], {})
                tag = "已處理" if meta.get("status") == "done" else "刻意為之隱藏"
                a1, a2, a3 = st.columns([3.8, 1.2, 1.0], vertical_alignment="center")
                a1.markdown(f"~~{ac['badge']} {ac['title']}~~ — *{tag} ({meta.get('at', '')})*")
                if a2.button("🔎 詳情", key=f"ack_view_{ac['id']}", use_container_width=True):
                    _open_decision_dialog(ac, ack_state)
                if a3.button("↩️ 還原", key=f"ack_undo_{ac['id']}", use_container_width=True):
                    ack_state.pop(ac["id"], None)
                    st.session_state["sm_decision_ack"] = ack_state
                    decision_engine.save_ack_state(ack_state)
                    st.rerun()

    st.divider()
    st.markdown("### 📂 底層原始數據查證專區（需要深查全表時再展開檢視）")

    # 3. 進階參數設定（預設收合）
    with st.expander("⚙️ 進階稽核門檻與各通路合約抽成微調（預設已套用站前館實測值，平時無需調整）", expanded=False):
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            stly_mode_lbl = st.radio(
                "STLY 去年同期比對基準",
                ["同星期比對 (-364天｜SiteMinder 預設)", "同日期比對 (-365天｜跨年對照)"],
                key="sm_stly_mode",
            )
            stly_mode = "dow" if "-364" in stly_mode_lbl else "date"
            allow_nye_101 = st.checkbox("跨年夜 12/31 允許 101 景觀房 E(05) > DT（降級為提醒）", value=True, key="sm_nye_101")
            last_min_days = st.slider("臨期滯銷檢核天數 (≤ N 天)", 1, 14, 3, key="sm_lastmin_d")
            underprice_pace_rn = st.slider("買氣強領先門檻 (Pace ≥ +N 晚且 ADR 下降)", 2, 20, 5, key="sm_up_rn")
        with c_p2:
            st.caption("金旭 HTGR 已自動標記「預付=淨價(已扣佣)」、「到付=毛價(未扣佣)」，僅對到付單扣除下方抽成。")
            comm_booking = st.slider("Booking.com 到付抽成 (%)", 10.0, 22.0, 15.0, 0.5, key="sm_c_bk")
            comm_ota_gross = st.slider("Trip / Agoda / Expedia 到付單抽成 (%)", 10.0, 22.0, 15.0, 0.5, key="sm_c_otag")
            comm_direct = st.slider("官網 Direct 金流手續費 (%)", 0.0, 5.0, 2.0, 0.5, key="sm_c_dir")
            comm_rakuten = st.slider("Rakuten 樂天到付抽成 (%)", 8.0, 18.0, 12.0, 0.5, key="sm_c_rak")

    # 若上方參數有調整，同步重跑一次 6 大規則供下方 5 分頁查證
    df_alerts, day_verdict = run_pricing_audit_rules(
        sm_snap=sm_snap, pace_data=pace_data, rt_inv=rt_inv, events_map=events_map,
        stly_mode=stly_mode, last_min_days=last_min_days,
        underprice_pace_rn=underprice_pace_rn, allow_nye_101_exception=allow_nye_101,
    )

    # 計算 KPI 數字
    pace_daily = pace_data.get("daily", [])
    pace_31 = pace_daily[:31]
    rev_31 = sum(r["otb_rev"] for r in pace_31)
    rn_31 = sum(r["sold_rn"] for r in pace_31)
    stly_rev_31 = sum(r["stly_dow_rev" if stly_mode == "dow" else "stly_date_rev"] for r in pace_31)
    stly_rn_31 = sum(r["stly_dow_rn" if stly_mode == "dow" else "stly_date_rn"] for r in pace_31)
    p7d_31 = sum(r["pickup_7d_rn"] for r in pace_31)
    yoy_rev_pct = ((rev_31 / stly_rev_31 - 1.0) * 100.0) if stly_rev_31 > 0 else 0.0
    yoy_rn_diff = rn_31 - stly_rn_31

    sev_red = int((df_alerts["嚴重度"] == "🔴 嚴重").sum()) if not df_alerts.empty else 0
    sev_org = int((df_alerts["嚴重度"] == "🟠 警告").sum()) if not df_alerts.empty else 0

    emp_ratio = float(promo_bf.get("empirical_actual_to_bar_ratio", 0.77) or 0.77)
    deep_cnt = sum(1 for r in promo_bf.get("recent_promo_orders", []) if r.get("is_deep_promo"))
    matched_n = int(promo_bf.get("matched_order_nights", 0) or 0)

    # 計算含早平均每人實收（雙人房為主）
    bf_rt_list = promo_bf.get("bf_vs_ro_by_rt", [])
    avg_bf_pax = int(round(np.mean([r["actual_per_pax"] for r in bf_rt_list if r.get("actual_per_pax", 0) > 0]))) if bf_rt_list else 0

    # 預先過濾有效未來掛牌日期（>= base_date），徹底排除過去歷史日
    inv_map = sm_snap.get("inventory", {})
    fut_inv_dates = [d for d in sorted(inv_map.keys()) if d >= dh["base_date"]]
    n_fut_inv_days = len(fut_inv_dates)

    # 4 欄 Executive KPI Cards（與頂部時效閘門完全同步，零精神分裂！）
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        if dh["sm_stale"]:
            st.markdown(
                f"""
<div class="sm-kpi" style="border-top: 4px solid #94a3b8;">
  <div class="sm-kpi-title">🔒 掛牌價稽核（快照已過期 {dh['sm_age_h']/24:.1f} 天）</div>
  <div class="sm-kpi-val" style="color:#64748b;">待同步解鎖</div>
  <div class="sm-kpi-sub">
    快照時間：<b>{sm_asof}</b><br>
    金旭庫存與買氣判讀仍 100% 即時運作中
  </div>
</div>
""",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f"""
<div class="sm-kpi" style="border-top: 4px solid #f97316;">
  <div class="sm-kpi-title">🚨 未來 {n_fut_inv_days} 天定價檢核結果</div>
  <div class="sm-kpi-val">{sev_red + sev_org} 項待確認</div>
  <div class="sm-kpi-sub">
    🔴 嚴重倒掛：<b>{sev_red}</b> 項｜🟠 警告：<b>{sev_org}</b> 項<br>
    （另含 {max(0, len(df_alerts) - sev_red - sev_org)} 項潛伏/連住提醒）
  </div>
</div>
""",
                unsafe_allow_html=True,
            )
    with k2:
        st.markdown(
            f"""
<div class="sm-kpi" style="border-top: 4px solid #8b5cf6;">
  <div class="sm-kpi-title">🏷️ 近14天實單 vs 牌告折數</div>
  <div class="sm-kpi-val">{emp_ratio*100:.1f}%</div>
  <div class="sm-kpi-sub">
    對照 <b>{matched_n}</b> 房晚實單｜異常深折：<b>{deep_cnt}</b> 筆<br>
    （常態 Booking/官網85折、預付淨價約72.3%）
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f"""
<div class="sm-kpi" style="border-top: 4px solid #10b981;">
  <div class="sm-kpi-title">🍳 同日同房含早實收加價</div>
  <div class="sm-kpi-val">NT$ {avg_bf_pax:,} /人</div>
  <div class="sm-kpi-sub">
    SiteMinder 牌告雙人 <b>+$889</b> (9折$800)<br>
    註：SP / S 兩房型全通路皆未開含早
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f"""
<div class="sm-kpi" style="border-top: 4px solid #2563eb;">
  <div class="sm-kpi-title">📈 未來 31 天金旭 OTB 進度</div>
  <div class="sm-kpi-val">NT$ {rev_31/10000:,.1f} 萬</div>
  <div class="sm-kpi-sub">
    STLY 同期：<b>{yoy_rev_pct:+.1f}%</b>（<b>{yoy_rn_diff:+d}</b> 晚）<br>
    近 7 日淨增訂 Pickup：<b>{p7d_31:+d}</b> 晚
  </div>
</div>
""",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # 預先整理未來有效天數定價合理性矩陣（僅納 >= base_date，對齊 PD 自有住房率與同星期常態基準）
    pms_daily_rt = rt_inv.get("daily_avail", {})
    pace_by_date = {r["date"]: r for r in pace_daily}
    dow_zh = ["一", "二", "三", "四", "五", "六", "日"]

    # 常態同星期 PD 掛牌中位數基準線（排除耶誕跨年 12/23~01/02）
    dow_pd_bucket = {i: [] for i in range(7)}
    for d_str in fut_inv_dates:
        try:
            wd_i = datetime.date.fromisoformat(d_str).weekday()
        except Exception:
            continue
        pr_v = int(inv_map[d_str].get("rooms", {}).get("PD", {}).get("r", 0))
        if pr_v > 0 and not ("12-23" <= d_str[5:] <= "12-31" or "01-01" <= d_str[5:] <= "01-02"):
            dow_pd_bucket[wd_i].append(pr_v)
    dow_pd_baseline = {w: int(round(np.median(v))) if v else 5200 for w, v in dow_pd_bucket.items()}

    rate_rows = []
    pd_occ_series = []
    pd_rate_series = []
    pd_base_series = []
    for d_str in fut_inv_dates:
        try:
            wd_idx = datetime.date.fromisoformat(d_str).weekday()
            wd = dow_zh[wd_idx]
        except Exception:
            wd_idx, wd = 0, ""
        d_obj = inv_map[d_str]
        rooms = d_obj.get("rooms", {})
        pc = pace_by_date.get(d_str, {})
        vs_str, verdict_str = day_verdict.get(d_str, ("—", ""))

        # 計算 PD 房型自有金旭即時住房率（與 PD 掛牌價 100% 同口徑對齊！）
        pd_rt_info = pms_daily_rt.get(d_str, {}).get("PD", {})
        pd_cap = max(1, int(pd_rt_info.get("tot", 22)) - int(pd_rt_info.get("ooo", 0)))
        pd_sold = int(pd_rt_info.get("sold", 0))
        pd_occ = round(min(105.0, pd_sold / pd_cap * 100.0), 1)
        pd_occ_series.append(pd_occ)
        pd_rate_series.append(int(rooms.get("PD", {}).get("r", 0)))
        pd_base_series.append(dow_pd_baseline.get(wd_idx, 5200))

        row_dict = {
            "日期": f"{d_str} ({wd})",
            "定價合理性判讀（為何高/低）": verdict_str,
            "較平常同週": vs_str,
            "住房率%": pc.get("occ_pct", 0.0),
            "PD自有住房率%": pd_occ,
            "金旭剩房": pc.get("rem_rn", 0),
            "台北活動": events_map.get(d_str, ""),
        }
        for rt in RT_TIER_ORDER:
            cell = rooms.get(rt, {})
            r_v, a_v = int(cell.get("r", 0)), int(cell.get("a", 0))
            row_dict[rt] = f"${r_v:,} ({a_v})" if r_v > 0 else "—"
        rate_rows.append(row_dict)
    df_rates_export = pd.DataFrame(rate_rows)
    df_bf_export = pd.DataFrame(bf_rt_list)
    df_bf_inv_export = pd.DataFrame(promo_bf.get("bf_inversion_alerts", []))
    df_promo_export = pd.DataFrame(promo_bf.get("recent_promo_orders", []))

    # 5 大精簡焦點分頁
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🚨 定價異常 & 收益部對質清單",
        "📅 42天定價合理性（為何高/低）",
        "🏷️ 促銷折數反查 & 含早/不含早差價",
        "📈 浮動軌跡 (Booking Curve) × 升等導流 & 淨利",
        "⚙️ 05:00 全自動同步狀態",
    ])

    # ------------------------------------------------------------------
    # TAB 1: 定價異常 & 收益部對質清單
    # ------------------------------------------------------------------
    with tab1:
        st.subheader("🚨 收益策略部定價疏漏自動掃描清單")
        if dh["sm_stale"]:
            st.warning(f"⚠️ 注意：目前 SiteMinder 掛牌快照停在 `{sm_asof}`（已逾 {dh['sm_age_h']:.0f} 小時），下方掛牌類項目僅供歷史參考，請同步最新快照後再用於對質。")
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
                excel_bytes = build_excel_audit_report(df_alerts, df_rates_export, df_bf_export, df_promo_export, df_bf_inv_export)
                st.download_button(
                    "📥 匯出 5 分頁稽核 Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"站前館_SiteMinder定價稽核_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            df_show = df_alerts[df_alerts["嚴重度"].isin(sel_sevs) & df_alerts["異常類型"].isin(sel_cats)].copy()
            st.dataframe(
                df_show[["日期", "嚴重度", "異常類型", "涉及房型", "掛牌價/數據摘要", "說明與收益部提問建議", "台北活動"]],
                use_container_width=True, hide_index=True, height=430,
            )

            with st.expander("💬 一鍵複製：給收益策略部的會議對質／LINE 核對清單", expanded=False):
                top_red = df_alerts[df_alerts["嚴重度"].isin(["🔴 嚴重", "🟠 警告"])].head(10)
                lines = [
                    f"【站前館 定價與促銷例行核對｜基準日 {datetime.date.today().strftime('%Y-%m-%d')}】",
                    "各位收益部夥伴好，我們比對未來 6 週掛牌價與金旭實單後，整理幾項想請協助確認：",
                    "",
                ]
                if fail_cnt > 0:
                    fd = health.get("failed_delivery_detail", {})
                    lines.append(f"0. 【緊急】今日有 {fail_cnt} 筆派單至金旭失敗（{fd.get('channel','')} #{fd.get('reservation_id','')}），已請櫃檯人工核對。")
                for idx, r in enumerate(top_red.to_dict("records"), 1):
                    lines.append(f"{idx}. {r['日期']}｜{r['異常類型']}（{r['涉及房型']}）：{r['掛牌價/數據摘要']} — {r['說明與收益部提問建議']}")
                if comp_cnt == 0:
                    lines.append("\n📌 補充建議：SiteMinder『Competitor rates』目前為 0 家競品，建議加入台北車站周邊 5–8 家核心同業作為調價依據。")
                st.code("\n".join(lines), language="text")

    # ------------------------------------------------------------------
    # TAB 2: 42天定價合理性（為何高／為何低）— 乾淨直白，零過去日汙染
    # ------------------------------------------------------------------
    with tab2:
        st.subheader(f"📅 未來 {n_fut_inv_days} 天掛牌價合理性總表（標準房價不含早）")
        if dh["sm_stale"]:
            st.warning(f"⚠️ SiteMinder 掛牌價停在 `{sm_asof}`（過期 {dh['sm_age_h']/24:.1f} 天）；下方已自動剔除過去日期，住房率與剩房為金旭最新實況。")
        st.caption(
            "直接呈現收益部於 SiteMinder 放出的標準牌告價與庫存 `$掛牌價 (OTA配額)`，並在最左側自動比對「同星期常態基準、金旭住房率、PD自有住房率與台北活動」。"
            "（紅底＝雙有庫存同床型倒掛；淺藍底＝低階已滿・高階降價清倉；黃底＝同床型跨階同價）"
        )
        if df_rates_export.empty:
            st.info("目前無未來有效 SiteMinder 掛牌價快照，下方自動切換呈現「金旭未來 35 天在手成交 ADR vs 全館住房率」永不斷線檢視。")
        else:
            def _style_rate_matrix(df_in):
                styles = pd.DataFrame("", index=df_in.index, columns=df_in.columns)
                for idx, d_str in enumerate(fut_inv_dates):
                    if idx >= len(df_in):
                        break
                    rooms = inv_map[d_str].get("rooms", {})
                    for hi, lo in TIER_PAIRS:
                        if hi in rooms and lo in rooms:
                            rh, ah = int(rooms[hi].get("r", 0)), int(rooms[hi].get("a", 0))
                            rl, al = int(rooms[lo].get("r", 0)), int(rooms[lo].get("a", 0))
                            if rl > rh > 0:
                                if al > 0:
                                    styles.at[idx, lo] = "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
                                    styles.at[idx, hi] = "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
                                else:
                                    styles.at[idx, hi] = "background-color: #e0f2fe; color: #0369a1; font-weight: 600;"
                            elif rl == rh > 0 and ah > 0 and al > 0:
                                styles.at[idx, lo] = "background-color: #fef3c7; color: #92400e; font-weight: 700;"
                                styles.at[idx, hi] = "background-color: #fef3c7; color: #92400e; font-weight: 700;"
                return styles

            st.dataframe(
                df_rates_export.style.apply(_style_rate_matrix, axis=None).format({
                    "住房率%": "{:.1f}%",
                    "PD自有住房率%": "{:.1f}%",
                }),
                use_container_width=True, hide_index=True, height=540,
            )

        # 修復版圖表：過濾過去日、對齊 PD 自有住房率、加上同星期常態基準虛線、修正圖層順序與整數刻度（若 SM 過期則自動切換金旭即時 ADR 永不斷線！）
        if not df_rates_export.empty and not dh["sm_stale"]:
            st.markdown("#### 📊 尊爵雙人 PD 掛牌價 vs PD 自有住房率 ＆ 同星期常態基準（同口徑對齊）")
            st.caption("💡 橘點高於『灰色同星期基準虛線』代表拉價、低於虛線代表讓價；背景藍柱為金旭 PD 自有住房率（已排除跨房型干擾與歷史過去日）。")
            fig_sync = go.Figure()
            # 背景層 (yaxis)：PD 自有住房率藍柱（置於底層，絕不遮蓋橘線變灰！）
            fig_sync.add_trace(go.Bar(
                x=df_rates_export["日期"],
                y=pd_occ_series,
                name="金旭 PD 自有住房率 (%)",
                marker_color="rgba(147, 197, 253, 0.55)",
                yaxis="y",
                customdata=df_rates_export["住房率%"],
                hovertemplate="PD 自有住房率: %{y:.1f}% (全館 %{customdata:.1f}%)<extra></extra>",
            ))
            # 前景層 1 (yaxis2)：同星期常態基準虛線（解決平假日 2 倍價差視覺誤導！）
            fig_sync.add_trace(go.Scatter(
                x=df_rates_export["日期"],
                y=pd_base_series,
                name="PD 同星期常態基準價",
                mode="lines",
                line=dict(color="#64748b", width=1.8, dash="dot"),
                yaxis="y2",
            ))
            # 前景層 2 (yaxis2)：PD 最新掛牌價鮮橘實線（最上層清晰呈現！）
            fig_sync.add_trace(go.Scatter(
                x=df_rates_export["日期"],
                y=pd_rate_series,
                name="尊爵雙人 PD 掛牌價 (TWD)",
                mode="lines+markers",
                line=dict(color="#f97316", width=3),
                marker=dict(size=7, color="#ea580c"),
                yaxis="y2",
            ))
            fig_sync.update_layout(
                height=360,
                margin=dict(l=20, r=20, t=20, b=40),
                legend=dict(orientation="h", y=1.13),
                hovermode="x unified",
                yaxis=dict(title="PD 自有住房率 (%)", side="right", range=[0, 108], dtick=25, showgrid=False),
                yaxis2=dict(title="PD 掛牌價 (TWD)", overlaying="y", side="left", rangemode="tozero"),
            )
            st.plotly_chart(fig_sync, use_container_width=True)
        elif pace_daily:
            st.markdown("#### 📊 金旭未來 35 天即時在手 ADR vs 全館住房率（永不斷線自動視角）")
            df_fb_chart = pd.DataFrame(pace_daily[:35])
            df_fb_chart["日期標籤"] = df_fb_chart["date"] + " (" + df_fb_chart["dow"] + ")"
            fig_fb = go.Figure()
            fig_fb.add_trace(go.Bar(
                x=df_fb_chart["日期標籤"], y=df_fb_chart["occ_pct"],
                name="全館在手住房率 (%)", marker_color="rgba(147, 197, 253, 0.55)", yaxis="y",
            ))
            fig_fb.add_trace(go.Scatter(
                x=df_fb_chart["日期標籤"], y=df_fb_chart["stly_dow_adr"],
                name="去年同期成交 ADR", mode="lines", line=dict(color="#64748b", width=1.8, dash="dot"), yaxis="y2",
            ))
            fig_fb.add_trace(go.Scatter(
                x=df_fb_chart["日期標籤"], y=df_fb_chart["otb_adr"],
                name="今年在手成交 ADR (TWD)", mode="lines+markers", line=dict(color="#2563eb", width=3), yaxis="y2",
            ))
            fig_fb.update_layout(
                height=360, margin=dict(l=20, r=20, t=20, b=40),
                legend=dict(orientation="h", y=1.13), hovermode="x unified",
                yaxis=dict(title="住房率 (%)", side="right", range=[0, 108], dtick=25, showgrid=False),
                yaxis2=dict(title="成交 ADR (TWD)", overlaying="y", side="left", rangemode="tozero"),
            )
            st.plotly_chart(fig_fb, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 3: 促銷折數反查 & 同產品「含早 vs 不含早」真實差價（經理核心關注！）
    # ------------------------------------------------------------------
    with tab3:
        st.subheader("🍳 一、同產品（同房型）『含早 vs 不含早』差價健檢")
        st.caption(
            "直接比較收益部在 SiteMinder 設定的「牌告含早加價」與金旭真實訂單「同入住日 × 同房型」的實際成交差價，"
            "確認含早方案是否合理獲利，或在哪個通路打折過頭（註：SP 與 S 全館未開放含早方案）。"
        )
        if not bf_rt_list:
            st.info("尚未產生含早差價對照快照，請執行一次 05:00 同步或點 Tab 5 立即更新。")
        else:
            df_bf_rt = pd.DataFrame(bf_rt_list)
            df_bf_rt["房型"] = df_bf_rt["rt"].map(lambda x: RT_NAMES.get(x, x))
            df_bf_rt["牌告每人"] = (df_bf_rt["sm_90_diff"] / df_bf_rt["pax"]).round(0).astype(int)
            st.dataframe(
                df_bf_rt[["房型", "matched_days", "avg_ro_adr", "avg_bf_adr", "sm_shelf_diff", "sm_90_diff", "actual_avg_diff", "actual_per_pax"]].rename(columns={
                    "matched_days": "同日比對天數", "avg_ro_adr": "不含早均價", "avg_bf_adr": "含早均價",
                    "sm_shelf_diff": "SiteMinder牌告加價", "sm_90_diff": "9折目標加價",
                    "actual_avg_diff": "金旭同日實收差價", "actual_per_pax": "實收每人早餐費",
                }).style.format({
                    "不含早均價": "${:,}", "含早均價": "${:,}",
                    "SiteMinder牌告加價": "+${:,}", "9折目標加價": "+${:,}",
                    "金旭同日實收差價": "+${:,}", "實收每人早餐費": "${:,}/人",
                }),
                use_container_width=True, hide_index=True,
            )

            bf_ch_list = promo_bf.get("bf_vs_ro_by_channel", [])
            if bf_ch_list:
                with st.expander("🔍 展開各通路「同日 × 同房型 × 同通路」含早差價明細（抓出哪家 OTA 含早賣太便宜）", expanded=False):
                    df_bfc = pd.DataFrame(bf_ch_list)
                    only_low = st.checkbox("僅顯示異常偏低（每人實收早餐費 < NT$ 350 或曾出現負差價）", value=False, key="sm_bf_low")
                    if only_low:
                        df_bfc = df_bfc[(df_bfc["per_pax_bf"] < 350) | (df_bfc["min_diff"] < 0)]
                    st.dataframe(
                        df_bfc[["rt", "sub_label", "pairs", "sm_shelf_diff", "actual_avg_diff", "per_pax_bf", "min_diff", "max_diff"]].rename(columns={
                            "rt": "房型", "sub_label": "通路", "pairs": "同日組數",
                            "sm_shelf_diff": "牌告加價", "actual_avg_diff": "實收平均差價",
                            "per_pax_bf": "每人實收早餐", "min_diff": "最低差價", "max_diff": "最高差價",
                        }).style.format({
                            "牌告加價": "+${:,}", "實收平均差價": "+${:,}",
                            "每人實收早餐": "${:,}/人", "最低差價": "{:+,}", "最高差價": "{:+,}",
                        }),
                        use_container_width=True, hide_index=True, height=300,
                    )

            # 含早 vs 不含早實單比對（預設嚴格過濾跨日浮動調價干擾！）
            bf_inv_list = promo_bf.get("bf_inversion_alerts", [])
            st.markdown("#### 🔍 同產品『含早 vs 不含早』實單差價檢核（已回溯 HTACL 原始下訂房型）")
            st.caption(
                "在浮動房價架構下，跨日下訂之價差多源於市場動態調價而非設定錯誤。預設僅顯示 **「同一天 4 小時內下訂＋同連住晚數級距」** 的嚴格比對樣本；"
                "若想觀察跨日浮動調價造成的帳面價差，可取消下方勾選。"
            )
            if not bf_inv_list:
                st.success("✅ 目前近 90 天區間內無異常含早差價樣本！")
            else:
                df_bi = pd.DataFrame(bf_inv_list)
                if "is_strict_intraday" not in df_bi.columns:
                    df_bi["is_strict_intraday"] = df_bi["bk_gap_days"] == 0
                bi0, bi1, bi2 = st.columns([1.6, 1.3, 1.8])
                with bi0:
                    only_strict = st.checkbox("🎯 僅看「同日 4h 內＋同連住級距」嚴格比對", value=True, key="sm_bi_strict")
                with bi1:
                    only_inv_neg = st.checkbox("🔴 僅看「含早 ≤ 不含早」", value=False, key="sm_bi_neg")
                with bi2:
                    bi_chs = sorted(df_bi["channel"].unique().tolist())
                    sel_bi_chs = st.multiselect("篩選通路", bi_chs, default=bi_chs, key="sm_bi_chs")
                df_bi_v = df_bi[df_bi["channel"].isin(sel_bi_chs)]
                if only_strict:
                    df_bi_v = df_bi_v[df_bi_v["is_strict_intraday"] == True]
                if only_inv_neg:
                    df_bi_v = df_bi_v[df_bi_v["diff_twd"] <= 0]
                if df_bi_v.empty:
                    st.success("✅ 在嚴格同日 4 小時內且同連住級距條件下，目前 **0 筆** 含早倒掛或異常重折！（其餘皆為正常跨日浮動調價或連住差異）")
                else:
                    st.dataframe(
                        df_bi_v[[
                            "severity", "stay_date", "rt", "sub_label",
                            "bf_rate", "bf_bk_date", "bf_ota_ref",
                            "ro_rate", "ro_bk_date", "ro_ota_ref",
                            "diff_twd", "per_pax_bf", "shortfall_vs_400", "sm_pct_compare", "reason"
                        ]].rename(columns={
                            "severity": "判讀分類", "stay_date": "入住日", "rt": "原始房型", "sub_label": "通路",
                            "bf_rate": "含早實收", "bf_bk_date": "含早訂房日", "bf_ota_ref": "含早OTA單號",
                            "ro_rate": "不含早實收", "ro_bk_date": "不含早訂房日", "ro_ota_ref": "不含早OTA單號",
                            "diff_twd": "兩單差價", "per_pax_bf": "每人實收", "shortfall_vs_400": "較$400/人差額",
                            "sm_pct_compare": "vs牌告BAR比例", "reason": "動態歸因說明",
                        }).style.format({
                            "含早實收": "${:,}", "不含早實收": "${:,}",
                            "兩單差價": "{:+,}", "每人實收": "${:,}/人",
                            "較$400/人差額": "{:+,}",
                        }),
                        use_container_width=True, hide_index=True, height=340,
                    )

        st.divider()
        st.subheader("🏷️ 二、這時候到底做了什麼促銷？（SiteMinder 牌告價 vs 金旭實際成交價反查）")
        st.caption(
            "OTA 訂單經由 SiteMinder 拋入金旭後金額 100% 吻合；將金旭近 14 天真實成交單價 ÷ 同日同房型今日 SiteMinder 掛牌價，"
            "即可反查各通路實際折數與促銷狀況。（💡 判讀要領：訂房日在 **近 1～3 天** 的單最能精準反映當下促銷折數；若訂房日在 1～2 週前卻遠低於今日掛牌價，則代表當時早鳥開價過低、後來才急拉價格！）"
        )
        p_sum = promo_bf.get("promo_summary_by_channel", [])
        if p_sum:
            df_ps = pd.DataFrame(p_sum)
            st.dataframe(
                df_ps.rename(columns={
                    "sub_label": "通路（含預付/到付別）", "nights": "對照房晚", "avg_sm_ref": "平均SiteMinder牌告",
                    "avg_pms_actual": "平均金旭實收", "med_ratio_pct": "實收中位折數%",
                    "min_ratio_pct": "最低砍到%", "max_ratio_pct": "最高%", "deep_promo_nights": "異常深折房晚數",
                }).style.format({
                    "平均SiteMinder牌告": "${:,}", "平均金旭實收": "${:,}",
                    "實收中位折數%": "{:.1f}%", "最低砍到%": "{:.1f}%", "最高%": "{:.1f}%",
                }),
                use_container_width=True, hide_index=True,
            )

        p_orders = promo_bf.get("recent_promo_orders", [])
        if p_orders:
            df_po = pd.DataFrame(p_orders)
            c_pf1, c_pf2 = st.columns([1.4, 2.0])
            with c_pf1:
                only_deep = st.checkbox("🔴 僅看「異常深折 / 機酒促銷」訂單", value=True, key="sm_only_deep")
            with c_pf2:
                ch_opts = sorted(df_po["channel"].unique().tolist())
                sel_chs = st.multiselect("篩選通路", ch_opts, default=ch_opts, key="sm_promo_chs")
            df_po_v = df_po[df_po["channel"].isin(sel_chs)]
            if only_deep:
                df_po_v = df_po_v[df_po_v["is_deep_promo"]]
            st.dataframe(
                df_po_v[["bk_date", "stay_date", "rt", "meal", "sub_label", "sm_ref_rate", "pms_actual_rate", "ratio_pct", "promo_diag", "ota_ref"]].rename(columns={
                    "bk_date": "訂房日", "stay_date": "入住日", "rt": "房型", "meal": "餐食",
                    "sub_label": "通路", "sm_ref_rate": "SiteMinder牌告", "pms_actual_rate": "金旭成交價",
                    "ratio_pct": "實收比例%", "promo_diag": "系統反查促銷判讀", "ota_ref": "OTA單號",
                }).style.format({
                    "SiteMinder牌告": "${:,}", "金旭成交價": "${:,}", "實收比例%": "{:.1f}%",
                }),
                use_container_width=True, hide_index=True, height=360,
            )

    # ------------------------------------------------------------------
    # TAB 4: STLY 訂房進度 (Pace) & 通路真實淨利 (Net ADR) + 洗盤預警
    # ------------------------------------------------------------------
    with tab4:
        st.subheader("🎢 一、單日浮動調價與進單軌跡（Booking Curve｜零升等汙染原始房型）")
        st.caption(
            "回溯金旭 `HTACL` 首次寫入之「原始下訂房型」，100% 剔除現場免費升等造成的假低價與假買氣，"
            "直接還原特定入住日從遠期（D-60）到近期（D-1）的真實成交價走勢與收益部調價軌跡。"
        )
        bc_list = dyn_yield.get("booking_curves", [])
        diag_by_d = {d["date"]: d for d in dyn_yield.get("daily_yield_diagnostics", [])}
        if bc_list:
            df_bc_all = pd.DataFrame(bc_list)
            bc_dates = sorted(df_bc_all["stay_date"].unique().tolist())
            # 預設挑選未來 14 天內點位最豐富的一天（如週末滿房日）方便一眼看懂曲線
            near_dates = [d for d in bc_dates[:16]]
            def_bc_d = max(near_dates, key=lambda x: len(df_bc_all[df_bc_all["stay_date"] == x])) if near_dates else bc_dates[0]
            def_idx = bc_dates.index(def_bc_d) if def_bc_d in bc_dates else 0

            def _fmt_bc_date(d_iso: str):
                dg = diag_by_d.get(d_iso, {})
                return f"{d_iso} ({dg.get('dow','')})｜住房率 {dg.get('occ_pct',0):.0f}%｜{dg.get('strategy_label','')}"

            c_bc1, c_bc2 = st.columns([1.8, 2.2])
            with c_bc1:
                sel_bc_date = st.selectbox("選擇目標入住日", bc_dates, index=def_idx, format_func=_fmt_bc_date, key="sm_bc_date")
            df_bc_day = df_bc_all[df_bc_all["stay_date"] == sel_bc_date].copy()
            avail_rts = sorted(df_bc_day["orig_rt"].unique().tolist(), key=lambda x: RT_TIER_ORDER.index(x) if x in RT_TIER_ORDER else 99)
            with c_bc2:
                sel_bc_rts = st.multiselect("篩選原始下訂房型", avail_rts, default=avail_rts[:6], key="sm_bc_rts")

            dg_sel = diag_by_d.get(sel_bc_date, {})
            if dg_sel:
                st.info(f"🧭 **{sel_bc_date} ({dg_sel.get('dow','')}) 動態收益診斷**：**{dg_sel.get('strategy_label','')}** — {dg_sel.get('evidence_note','')}")

            df_bc_plot = df_bc_day[df_bc_day["orig_rt"].isin(sel_bc_rts)].sort_values("bk_date")
            if not df_bc_plot.empty:
                df_bc_plot["系列"] = df_bc_plot["orig_rt"] + " (" + df_bc_plot["meal"].map({"RO": "不含早", "BF": "含早"}).fillna("不含早") + ")"
                df_bc_plot["氣泡大小"] = df_bc_plot["rn"].clip(lower=1, upper=8) * 5 + 6
                fig_bc = px.line(
                    df_bc_plot,
                    x="bk_date",
                    y="avg_rate",
                    color="系列",
                    markers=True,
                    hover_data={
                        "bk_date": True,
                        "lead_at_bk": True,
                        "rn": True,
                        "avg_rate": ":,",
                        "upgraded_rn": True,
                        "系列": False,
                    },
                    labels={
                        "bk_date": "旅客下訂日期",
                        "avg_rate": "原始房型成交單價 (NT$)",
                        "lead_at_bk": "提前天數(D-N)",
                        "rn": "成交房晚",
                        "upgraded_rn": "後續現場升等晚數",
                    },
                )
                fig_bc.update_traces(marker=dict(size=9))
                fig_bc.update_layout(
                    height=350,
                    margin=dict(l=20, r=20, t=20, b=35),
                    legend=dict(orientation="h", y=1.14),
                    hovermode="x unified",
                )
                st.plotly_chart(fig_bc, use_container_width=True)

        vel_list = dyn_yield.get("rt_velocity_30d", [])
        upg_mat = dyn_yield.get("upgrade_matrix_60d", [])
        if vel_list or upg_mat:
            with st.expander("🔄 展開「12 房型真實原始買氣 vs 現場免費升等消化矩陣（揭密 ER 專用升等池）」", expanded=False):
                dy_sum = dyn_yield.get("summary", {})
                st.caption(
                    f"統計視窗：{dy_sum.get('upg_window_label','近60天')}｜全館共發生 **{dy_sum.get('total_upgraded_rn',0)} 晚** 現場改房型／免費升等，"
                    f"其中 **同床型零增本大床升等 {dy_sum.get('zero_cost_queen_upg_rn',0)} 晚**、**ER(8間未掛OTA專用池) 吸收 {dy_sum.get('er_absorbed_rn',0)} 晚 ({dy_sum.get('er_share_pct',0):.1f}%)**。"
                )
                vcol1, vcol2 = st.columns([1.45, 1.15])
                with vcol1:
                    st.markdown("##### 📊 未來 30 天各房型「真實買氣 vs 排房移轉」")
                    if vel_list:
                        df_vel = pd.DataFrame(vel_list)
                        st.dataframe(
                            df_vel[[
                                "rt", "tot_inv", "orig_demand_rn", "cur_assigned_rn",
                                "net_upgrade_shift", "orig_pickup_7d", "orig_adr_all", "adr_delta_7d"
                            ]].rename(columns={
                                "rt": "房型", "tot_inv": "總房數", "orig_demand_rn": "原始下訂(晚)",
                                "cur_assigned_rn": "現排房(晚)", "net_upgrade_shift": "升等淨移入",
                                "orig_pickup_7d": "7日原始淨買氣", "orig_adr_all": "原始成交ADR", "adr_delta_7d": "近7日均價變化",
                            }).style.format({
                                "升等淨移入": "{:+d}", "7日原始淨買氣": "{:+d}",
                                "原始成交ADR": "${:,}", "近7日均價變化": "{:+,}",
                            }),
                            use_container_width=True, hide_index=True, height=430,
                        )
                with vcol2:
                    st.markdown("##### 🔀 近 60 天現場免費升等路徑排行")
                    if upg_mat:
                        df_um = pd.DataFrame(upg_mat)
                        df_um["路徑"] = df_um["orig_rt"] + " ➔ " + df_um["cur_rt"]
                        st.dataframe(
                            df_um[["路徑", "room_nights", "orders", "avg_rate", "category"]].rename(columns={
                                "room_nights": "房晚數", "orders": "訂單數", "avg_rate": "原單均價", "category": "成本與床型分類",
                            }).style.format({"原單均價": "${:,}"}),
                            use_container_width=True, hide_index=True, height=430,
                        )

        st.divider()
        st.subheader("📈 二、未來 STLY 去年同期進度 (Pace) × 7 日淨增訂 (Pickup)")
        if pace_daily:
            df_p = pd.DataFrame(pace_daily)
            show_days = st.slider("檢視未來天數", 14, min(90, len(df_p)), 42, key="sm_pace_days")
            df_pv = df_p.head(show_days).copy()
            stly_rn_col = "stly_dow_rn" if stly_mode == "dow" else "stly_date_rn"
            stly_adr_col = "stly_dow_adr" if stly_mode == "dow" else "stly_date_adr"
            final_rn_col = "ly_dow_final_rn" if stly_mode == "dow" else "ly_date_final_rn"
            df_pv["日期標籤"] = df_pv["date"].str[5:] + "(" + df_pv["dow"] + ")"

            fig_pace = go.Figure()
            fig_pace.add_trace(go.Bar(x=df_pv["日期標籤"], y=df_pv["sold_rn"], name="今年在手 OTB (晚)", marker_color="#2563eb", opacity=0.82))
            fig_pace.add_trace(go.Scatter(x=df_pv["日期標籤"], y=df_pv[stly_rn_col], name="去年今日 STLY (晚)", mode="lines+markers", line=dict(color="#f59e0b", width=2.5, dash="dot")))
            fig_pace.add_trace(go.Scatter(x=df_pv["日期標籤"], y=df_pv[final_rn_col], name="去年最終結算 (晚)", mode="lines", line=dict(color="#94a3b8", width=1.5, dash="dash")))
            fig_pace.update_layout(height=340, margin=dict(l=20, r=20, t=20, b=35), legend=dict(orientation="h", y=1.12), hovermode="x unified", yaxis_title="房晚數")
            st.plotly_chart(fig_pace, use_container_width=True)

            # 月底達標試算器（改採真實實測折數 emp_ratio，零高估！）
            df_occ = pd.read_csv(OCC_CSV_PATH) if os.path.exists(OCC_CSV_PATH) else pd.DataFrame()
            if not df_occ.empty:
                df_occ["date"] = df_occ["date"].astype(str).str[:10]
                df_occ["ym"] = df_occ["date"].str[:7]
                cur_ym = datetime.date.today().strftime("%Y-%m")
                avail_yms = sorted([ym for ym in df_occ["ym"].unique() if ym >= cur_ym])[:4]
                if avail_yms:
                    with st.expander("🎯 月底營收達標試算器（依近14天實測成交折數還原預估進帳，不高估）", expanded=False):
                        sc1, sc2, sc3 = st.columns([1, 1.2, 1.2])
                        with sc1:
                            sel_ym = st.selectbox("選擇月份", avail_yms, index=min(1, len(avail_yms) - 1), key="sm_sim_ym")
                        def_target = 16000000
                        if targets_df is not None and not targets_df.empty:
                            try:
                                tc = targets_df.copy()
                                tc.columns = [str(c).strip().lower() for c in tc.columns]
                                m_col = next((c for c in tc.columns if "month" in c or "date" in c or "年" in c), None)
                                r_col = next((c for c in tc.columns if "target" in c or "rev" in c or "目標" in c), None)
                                if m_col and r_col:
                                    hit = tc[tc[m_col].astype(str).str.replace("/", "-").str.startswith(sel_ym)]
                                    if not hit.empty:
                                        v_t = pd.to_numeric(str(hit.iloc[0][r_col]).replace(",", ""), errors="coerce")
                                        if pd.notna(v_t) and v_t > 0:
                                            def_target = int(round(v_t))
                            except Exception:
                                pass
                        with sc2:
                            target_rev = st.number_input("月營收目標 (TWD)", 1000000, 50000000, def_target, 500000, key="sm_sim_tgt")
                        with sc3:
                            sell_through = st.slider("預估剩餘空房去化率 (%)", 10, 100, 65, 5, key="sm_sim_st")

                        m_df = df_occ[df_occ["ym"] == sel_ym]
                        cur_rev = int(round(pd.to_numeric(m_df["revenue"], errors="coerce").fillna(0).sum()))
                        cur_sold = int(round(pd.to_numeric(m_df["sold_rooms"], errors="coerce").fillna(0).sum()))
                        today_iso = datetime.date.today().strftime("%Y-%m-%d")
                        fut_m = m_df[m_df["date"] >= today_iso]
                        rem_fut = int(round((pd.to_numeric(fut_m["total_rooms"], errors="coerce").fillna(126) - pd.to_numeric(fut_m["sold_rooms"], errors="coerce").fillna(0)).clip(lower=0).sum()))
                        sm_ref_p = [
                            np.mean([inv_map[d]["rooms"][k]["r"] for k in ("SP", "PD", "DT") if k in inv_map[d].get("rooms", {}) and inv_map[d]["rooms"][k].get("r", 0) > 0])
                            for d in fut_m["date"].tolist() if d in inv_map and any(k in inv_map[d].get("rooms", {}) for k in ("SP", "PD", "DT"))
                        ]
                        sm_bar_avg = int(round(np.mean(sm_ref_p))) if sm_ref_p else (int(round(cur_rev / cur_sold)) if cur_sold > 0 else 5500)
                        est_real_adr = int(round(sm_bar_avg * emp_ratio))
                        est_add_rn = int(round(rem_fut * (sell_through / 100.0)))
                        est_final = cur_rev + est_add_rn * est_real_adr
                        mc1, mc2, mc3, mc4 = st.columns(4)
                        mc1.metric(f"{sel_ym} 已入帳＋OTB", f"NT$ {cur_rev:,}", f"已售 {cur_sold:,} 晚")
                        mc2.metric("距離目標缺口", f"NT$ {max(0, target_rev - cur_rev):,}", f"尚餘 {rem_fut:,} 空房")
                        mc3.metric("預估實收進帳 ADR", f"NT$ {est_real_adr:,}", f"掛牌 ${sm_bar_avg:,} × 實測 {emp_ratio*100:.1f}%")
                        mc4.metric("預估月底總營收", f"NT$ {est_final:,}", f"{'超標' if est_final >= target_rev else '缺口'} NT$ {est_final - target_rev:+,}")

        st.divider()
        st.subheader("💰 通路真實淨利排行（還原預付 Net vs 到付 Gross）＆ 洗盤取消率")
        windows = chan_intel.get("windows", {})
        if windows:
            win_keys = list(windows.keys())
            sel_win = st.selectbox("選擇統計區間", win_keys, format_func=lambda k: windows[k]["label"], key="sm_chan_win")
            rows = windows[sel_win]["rows"]
            recs = []
            for r in rows:
                b, ch = r["rate_basis"], r["channel"]
                comm = 0.0 if b in ("net_prepaid", "corporate") else (comm_booking if ch == "Booking.com" else (comm_rakuten if "Rakuten" in ch else (comm_direct if b == "direct" else comm_ota_gross)))
                net_real_rev = int(round(r["net_rev"] * (1.0 - comm / 100.0)))
                recs.append({**r, "飯店實拿淨營收": net_real_rev})
            df_ch = pd.DataFrame(recs)
            if not df_ch.empty:
                tot_rn_w = df_ch["net_rn"].sum()
                g = df_ch.groupby("channel", as_index=False).agg(
                    淨訂單=("net_bk", "sum"), 取消=("cx_bk", "sum"), NoShow=("nosh_bk", "sum"),
                    淨房晚=("net_rn", "sum"), 金旭帳面營收=("net_rev", "sum"),
                    飯店實拿淨營收=("飯店實拿淨營收", "sum"), ALOS=("alos", "mean"), 提前預訂天=("lead_days", "mean"),
                )
                g["佔比%"] = np.where(tot_rn_w > 0, (g["淨房晚"] / tot_rn_w * 100.0).round(1), 0.0)
                g["帳面ADR"] = np.where(g["淨房晚"] > 0, (g["金旭帳面營收"] / g["淨房晚"]).round(0), 0).astype(int)
                g["實拿Net ADR"] = np.where(g["淨房晚"] > 0, (g["飯店實拿淨營收"] / g["淨房晚"]).round(0), 0).astype(int)
                g["取消率%"] = np.where((g["淨訂單"] + g["取消"] + g["NoShow"]) > 0, ((g["取消"] + g["NoShow"]) / (g["淨訂單"] + g["取消"] + g["NoShow"]) * 100.0).round(1), 0.0)
                g["ALOS"] = g["ALOS"].round(2)
                g["提前預訂天"] = g["提前預訂天"].round(1)
                g = g.sort_values("淨房晚", ascending=False).rename(columns={"channel": "通路"})
                st.dataframe(
                    g[["通路", "淨房晚", "佔比%", "帳面ADR", "實拿Net ADR", "取消率%", "ALOS", "提前預訂天", "飯店實拿淨營收"]].style.format({
                        "淨房晚": "{:,}", "佔比%": "{:.1f}%", "帳面ADR": "${:,}", "實拿Net ADR": "${:,}",
                        "取消率%": "{:.1f}%", "ALOS": "{:.2f}", "提前預訂天": "{:.1f}天", "飯店實拿淨營收": "${:,}",
                    }),
                    use_container_width=True, hide_index=True,
                )

                wash_rows = chan_intel.get("daily_wash_60d", [])
                if wash_rows:
                    df_w = pd.DataFrame(wash_rows)
                    high_risk = df_w[df_w["high_cx_share_pct"] >= 35.0]
                    if not high_risk.empty:
                        st.warning(
                            f"🌊 **退訂洗盤預警**：未來 60 天有 **{len(high_risk)} 天** 的在手訂單中高取消通路（Agoda＋Expedia）佔比 ≥ 35%，"
                            f"預估日均潛在退訂洗盤約 **{high_risk['est_wash_rn'].mean():.1f} 間**，房務排班與備貨建議留彈性。"
                        )

    # ------------------------------------------------------------------
    # TAB 5: 05:00 全自動同步狀態
    # ------------------------------------------------------------------
    with tab5:
        st.subheader("⚙️ 每日 05:00 全自動同步狀態")
        c_s1, c_s2 = st.columns(2)
        with c_s1:
            st.markdown("##### 1️⃣ 金旭 HT2000 唯讀 SQL（05:00 無需開 Chrome 全自動）")
            st.success(f"✅ 通路真實淨利與取消情報：`{chan_intel.get('as_of', '未同步')}`")
            st.success(f"✅ 促銷折數反查與含早差價：`{promo_bf.get('as_of', '未同步')}`")
            st.success(f"✅ 未來 90 天 STLY Pace & Pickup：`{pace_data.get('as_of', '未同步')}`")
            st.success(f"✅ 動態收益與原始買氣軌跡 (HTACL)：`{dyn_yield.get('as_of', '未同步')}`")
            st.success(f"✅ 12 房型每日剩餘庫存：`{rt_asof}`")
        with c_s2:
            n_sm_hist = len([fn for fn in os.listdir(SM_HISTORY_DIR) if fn.endswith(".json")]) if os.path.exists(SM_HISTORY_DIR) else 0
            st.markdown("##### 2️⃣ SiteMinder 42 天掛牌價快照（唯讀＋歷史歸檔）")
            st.info(f"🌐 最新有效掛牌快照：`{sm_asof}`（涵蓋 {n_inv_days} 天｜已累積歷史歸檔 `{n_sm_hist}` 天）")
            if sm_status:
                st.caption(f"最近狀態：`{sm_status.get('last_attempt_at','-')}`｜{sm_status.get('detail','-')}")
            st.markdown(
                """
- **最省心用法**：平常不用刻意開著 SiteMinder；每週開會前或想查掛牌價時，打開 Chrome 登入 SiteMinder 並雙擊 `sync_siteminder_now.bat`，15 秒自動更新推上雲端。
- **想讓 05:00 也自動抓掛牌價**：雙擊 `launch_siteminder_chrome.bat` 開啟專用視窗（Port `9223`）登入 SiteMinder 縮小放著即可。
"""
            )

        st.divider()
        b1, b2 = st.columns(2)
        with b1:
            if st.button("🔄 立即直連金旭重算促銷、含早差價與 Pace（飯店區網本機用）", use_container_width=True):
                with st.spinner("正在唯讀查詢金旭 HT2000 ..."):
                    try:
                        import pms_revenue_intel
                        res = pms_revenue_intel.run_pms_revenue_intel_sync()
                        st.success(f"✅ 同步完成：{res}")
                        st.rerun()
                    except Exception as e:
                        st.warning(f"雲端環境無法直連 SQL Server（已自動使用最新快照）：{e}")
        with b2:
            if st.button("🌐 立即唯讀刷新本機 Chrome SiteMinder 掛牌價", use_container_width=True):
                with st.spinner("正在唯讀擷取 SiteMinder 42 天掛牌價..."):
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
