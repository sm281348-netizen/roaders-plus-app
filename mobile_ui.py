# -*- coding: utf-8 -*-
"""
mobile_ui.py
Hotel Master 專屬手機版介面模組（修正版方案 A）
==============================================
核心設計原則：
1. 100% 共用 app.py 的既有 Google Sheets 快取、權限與核心計算邏輯，零雙邊漂移。
2. 使用 Streamlit 原生水平分段導覽（不採會觸發整頁重載而遺失登入的 HTML 假底欄）。
3. 現場餐期登記採 st.form + 原生數字鍵盤 + 分餐期局部欄位安全寫入（防卡頓、防多人覆蓋）。
4. 「分析月份」與「現場填報日期」完全解耦，解決採購角色月底鎖定衝突。
5. 行動版圖表採純 HTML/CSS 響應式條狀圖與進度條，徹底杜絕 Altair 觸控鎖死滑動。
6. 嚴格遵守館別權限：僅站前館登入者可跨館切換主題館；主題館與採購角色鎖定自身權限。
"""

import os
import re
import html
import calendar
import datetime
import pandas as pd
import streamlit as st


# ─────────────────────────────────────────────────────────────
# 1. 行動版專屬 CSS 與共用微元件
# ─────────────────────────────────────────────────────────────
def _inject_mobile_css():
    st.markdown(
        """
<style>
/* 行動版緊湊版面：隱藏桌面側邊欄並收緊邊距 */
[data-testid="stSidebar"],
[data-testid="collapsedControl"] {
    display: none !important;
}
.main .block-container {
    max-width: 540px !important;
    padding-top: 0.65rem !important;
    padding-bottom: 3.5rem !important;
    padding-left: 0.85rem !important;
    padding-right: 0.85rem !important;
    margin: 0 auto !important;
}
/* 觸控目標最小 44px 規範 */
.stButton > button,
.stDownloadButton > button,
.stFormSubmitButton > button {
    min-height: 44px !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    font-size: 0.9rem !important;
}
.stFormSubmitButton > button {
    min-height: 48px !important;
}
/* 強制數字等寬對齊 */
.m-num {
    font-variant-numeric: tabular-nums;
}
/* 頂部識別列 */
.m-topbar {
    background: linear-gradient(135deg, #1f2c56 0%, #2b3f78 100%);
    color: #ffffff;
    border-radius: 14px;
    padding: 12px 16px;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 4px 12px rgba(31, 44, 86, 0.18);
}
.m-topbar-title {
    font-size: 16px;
    font-weight: 800;
    letter-spacing: 0.3px;
}
.m-topbar-sub {
    font-size: 12px;
    color: #cbd5e1;
    margin-top: 2px;
}
.m-pill-badge {
    background: rgba(241, 158, 56, 0.22);
    border: 1px solid rgba(241, 158, 56, 0.7);
    color: #fde68a;
    font-size: 12px;
    font-weight: 700;
    padding: 4px 10px;
    border-radius: 999px;
    white-space: nowrap;
}
/* 2x2 Hero KPI 固定雙欄網格（在 375px 手機上絕不折行成單欄） */
.m-kpi-grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 10px;
    margin: 8px 0 12px 0;
}
.m-kpi-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 13px 14px;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
}
.m-kpi-label {
    font-size: 12px;
    font-weight: 600;
    color: #64748b;
    margin-bottom: 4px;
}
.m-kpi-val {
    font-size: 22px;
    font-weight: 800;
    color: #0f172a;
    line-height: 1.2;
    font-variant-numeric: tabular-nums;
}
.m-kpi-sub {
    font-size: 12px;
    color: #475569;
    margin-top: 4px;
    font-variant-numeric: tabular-nums;
}
/* 3 欄精簡利潤/摘要卡 */
.m-tri-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 8px;
    margin-top: 8px;
}
.m-tri-cell {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 10px 6px;
    text-align: center;
}
.m-tri-label {
    font-size: 11.5px;
    color: #64748b;
    font-weight: 600;
}
.m-tri-val {
    font-size: 15px;
    font-weight: 800;
    margin-top: 3px;
    font-variant-numeric: tabular-nums;
}
/* 通用卡片容器 */
.m-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 14px 16px;
    margin-bottom: 12px;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
}
.m-card-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 8px;
}
.m-card-title {
    font-size: 15px;
    font-weight: 800;
    color: #0f172a;
}
.m-badge-ok {
    background: #dcfce7; color: #166534;
    font-size: 12px; font-weight: 700; padding: 3px 9px; border-radius: 999px;
}
.m-badge-warn {
    background: #fef3c7; color: #92400e;
    font-size: 12px; font-weight: 700; padding: 3px 9px; border-radius: 999px;
}
.m-badge-err {
    background: #fee2e2; color: #991b1b;
    font-size: 12px; font-weight: 700; padding: 3px 9px; border-radius: 999px;
}
.m-badge-muted {
    background: #f1f5f9; color: #475569;
    font-size: 12px; font-weight: 700; padding: 3px 9px; border-radius: 999px;
}
/* 進度條 */
.m-bar-bg {
    width: 100%;
    height: 9px;
    background: #e2e8f0;
    border-radius: 999px;
    overflow: hidden;
    margin: 6px 0;
}
.m-bar-fill {
    height: 100%;
    border-radius: 999px;
}
/* 明細列表列 */
.m-list-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 9px 0;
    border-top: 1px solid #f1f5f9;
    gap: 8px;
}
.m-list-main {
    font-size: 14px;
    font-weight: 700;
    color: #0f172a;
}
.m-list-sub {
    font-size: 12px;
    color: #64748b;
    margin-top: 2px;
}
.m-list-right {
    text-align: right;
    white-space: nowrap;
    font-variant-numeric: tabular-nums;
}
/* 主題館暫無數據精簡提示卡 */
.m-notice-amber {
    background: linear-gradient(135deg, #fffaf0 0%, #fef3c7 100%);
    border-left: 5px solid #f59e0b;
    border-radius: 12px;
    padding: 14px 16px;
    margin: 8px 0 12px 0;
}
.m-notice-title {
    font-size: 14px;
    font-weight: 800;
    color: #92400e;
}
.m-notice-body {
    font-size: 13px;
    color: #78350f;
    line-height: 1.55;
    margin-top: 4px;
}
/* 純 CSS 行動柱狀圖（零觸控捲動衝突） */
.m-chart-wrap {
    display: flex;
    align-items: flex-end;
    justify-content: space-between;
    gap: 6px;
    height: 155px;
    padding-top: 22px;
    margin-top: 6px;
}
.m-chart-col {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    min-width: 0;
}
.m-chart-val {
    font-size: 11px;
    font-weight: 700;
    color: #334155;
    margin-bottom: 4px;
    font-variant-numeric: tabular-nums;
}
.m-chart-bar {
    width: 100%;
    max-width: 34px;
    border-radius: 6px 6px 2px 2px;
}
.m-chart-lbl {
    font-size: 11px;
    color: #64748b;
    margin-top: 5px;
    white-space: nowrap;
}
/* 讓水平 radio 像分頁膠囊 */
div[role="radiogroup"] {
    gap: 6px !important;
    flex-wrap: nowrap !important;
    overflow-x: auto !important;
    padding-bottom: 4px !important;
}
/* 強制手機版精簡動作按鈕列維持水平並排，不折疊成多行 */
@media (max-width: 640px) {
    [data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
        gap: 6px !important;
    }
    [data-testid="stHorizontalBlock"] > [data-testid="column"],
    [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
        min-width: 0 !important;
        flex: 1 1 0 !important;
    }
}
</style>
        """,
        unsafe_allow_html=True,
    )


def _safe_int(val, default=0) -> int:
    try:
        if val is None or pd.isna(val):
            return default
        s = str(val).strip().replace(",", "").replace("%", "")
        if not s:
            return default
        return int(round(float(s)))
    except Exception:
        return default


def _safe_float(val, default=0.0) -> float:
    try:
        if val is None or pd.isna(val):
            return default
        s = str(val).strip().replace(",", "").replace("%", "")
        if not s:
            return default
        return float(s)
    except Exception:
        return default


def _weekday_zh(d: datetime.date) -> str:
    return ["一", "二", "三", "四", "五", "六", "日"][d.weekday()]


def _render_nav_pills(options, key, default_val):
    """優先使用 Streamlit 1.40+ segmented_control，舊版自動降級為水平 radio。"""
    if key not in st.session_state or st.session_state[key] not in options:
        st.session_state[key] = default_val
    if hasattr(st, "segmented_control"):
        chosen = st.segmented_control(
            "功能分頁",
            options=options,
            default=st.session_state[key],
            key=f"{key}_seg",
            label_visibility="collapsed",
        )
        if chosen and chosen in options:
            st.session_state[key] = chosen
        return st.session_state[key]
    chosen = st.radio(
        "功能分頁",
        options=options,
        index=options.index(st.session_state[key]),
        horizontal=True,
        key=f"{key}_rad",
        label_visibility="collapsed",
    )
    st.session_state[key] = chosen
    return chosen


def _recent_months_list(n: int = 12):
    today = datetime.date.today()
    y, m = today.year, today.month
    out = []
    for _ in range(n):
        out.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return out


def _render_theme_compact_notice(title: str, summary: str, points: list):
    st.markdown(
        f"""
<div class="m-notice-amber">
<div class="m-notice-title">⚠️ 暫無數據可提供此服務（{html.escape(title)}）</div>
<div class="m-notice-body">主題館目前尚未接入此站前館內部數據，暫不顯示空白或誤差統計。</div>
</div>
        """,
        unsafe_allow_html=True,
    )
    with st.expander("💡 展開了解此模組可呈現之戰略資訊"):
        st.write(summary)
        for p in points:
            st.markdown(f"- {p}")
        st.caption("若主題館主管對此戰略分析有興趣，後續可再討論資料串接導入。")


# ─────────────────────────────────────────────────────────────
# 2. 安全分餐期局部寫入 f&b_report（防多人互蓋、絕不觸碰預估欄）
# ─────────────────────────────────────────────────────────────
def _update_fb_report_partial(date_str: str, meal_scope: str, field_values: dict):
    """
    精準只更新所選餐期的實際來客欄位，保留其餘餐期與 PMS 預估欄位不被覆寫：
      - 'breakfast' : 只更新 C (主題早實際) 與 E (站前早實際)
      - 'afternoon' : 只更新 G (主題午實際) 與 I (站前午實際)
      - 'hh'        : 只更新 J (4F HH 實際)
    """
    import gspread

    try:
        cfg = st.secrets["connections"]["gsheets_station"]
        sa_info = {
            k: cfg[k]
            for k in [
                "type",
                "project_id",
                "private_key_id",
                "private_key",
                "client_email",
                "client_id",
                "auth_uri",
                "token_uri",
                "auth_provider_x509_cert_url",
                "client_x509_cert_url",
            ]
        }
        gc = gspread.service_account_from_dict(sa_info)
        sh = gc.open_by_url(cfg["spreadsheet"])
        ws = sh.worksheet("f&b_report")
        all_rows = ws.get_all_values()

        m_t = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(date_str).strip())
        norm_target = (
            f"{m_t.group(1)}-{int(m_t.group(2)):02d}-{int(m_t.group(3)):02d}"
            if m_t
            else str(date_str).strip().replace("/", "-")
        )

        target_row_idx = None
        for i, r in enumerate(all_rows):
            if r and len(r) > 0:
                m_r = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(r[0]).strip())
                r_norm = (
                    f"{m_r.group(1)}-{int(m_r.group(2)):02d}-{int(m_r.group(3)):02d}"
                    if m_r
                    else str(r[0]).strip().replace("/", "-")
                )
                if r_norm == norm_target:
                    target_row_idx = i + 1
                    break

        if not target_row_idx:
            # 該日期尚未建列，先補一列空白日期再局部更新
            ws.append_row([norm_target, "", "", "", "", "", "", "", "", ""])
            target_row_idx = len(all_rows) + 1

        updates = []
        if meal_scope == "breakfast":
            updates.append({"range": f"C{target_row_idx}", "values": [[str(field_values["th_bf_act"])]]})
            updates.append({"range": f"E{target_row_idx}", "values": [[str(field_values["zq_bf_act"])]]})
        elif meal_scope == "afternoon":
            updates.append({"range": f"G{target_row_idx}", "values": [[str(field_values["th_af_act"])]]})
            updates.append({"range": f"I{target_row_idx}", "values": [[str(field_values["zq_af_act"])]]})
        elif meal_scope == "hh":
            updates.append({"range": f"J{target_row_idx}", "values": [[str(field_values["hh_act"])]]})

        if updates:
            ws.batch_update(updates, value_input_option="USER_ENTERED")
        st.cache_data.clear()
        return True, "儲存成功"
    except Exception as e:
        return False, str(e)


# ─────────────────────────────────────────────────────────────
# 3. 各分頁渲染器
# ─────────────────────────────────────────────────────────────
def _render_mobile_home(ctx: dict):
    current_hotel = ctx["current_hotel"]
    today = datetime.date.today()

    if "mobile_home_date" not in st.session_state:
        st.session_state["mobile_home_date"] = today

    # 日期選擇獨立滿寬列 + 下方 3 鍵水平快切（避免窄螢幕擠壓日期框）
    picked = st.date_input(
        "📅 檢視日期",
        value=st.session_state["mobile_home_date"],
        key="m_home_date_input",
        label_visibility="collapsed",
    )
    if picked != st.session_state["mobile_home_date"]:
        st.session_state["mobile_home_date"] = picked
        st.rerun()

    d_c1, d_c2, d_c3 = st.columns(3)
    with d_c1:
        if st.button("‹ 前一日", key="m_prev_day", use_container_width=True):
            st.session_state["mobile_home_date"] -= datetime.timedelta(days=1)
            st.rerun()
    with d_c2:
        if st.button("📍 今日", key="m_today_btn", use_container_width=True):
            st.session_state["mobile_home_date"] = today
            st.rerun()
    with d_c3:
        if st.button("次一日 ›", key="m_next_day", use_container_width=True):
            st.session_state["mobile_home_date"] += datetime.timedelta(days=1)
            st.rerun()

    sel_date = st.session_state["mobile_home_date"]
    d_str = sel_date.strftime("%Y-%m-%d")
    st.caption(f"📅 目前檢視：**{sel_date.strftime('%Y/%m/%d')}（週{_weekday_zh(sel_date)}）**")

    # 讀取當日與當月真實 PMS 營運數據
    day_data = ctx["get_daily_data"](d_str) or {}
    m_sum = ctx["fetch_month_summary_by_hotel"](sel_date.year, sel_date.month, current_hotel)

    occ_val = _safe_float(day_data.get("occ_rate", 0.0))
    adr_val = _safe_int(day_data.get("adr", 0))
    rev_val = _safe_int(day_data.get("revenue", 0))
    sold_rooms = _safe_int(day_data.get("sold_rooms", 0))
    total_rooms = _safe_int(day_data.get("total_rooms", 112 if current_hotel == "主題館" else 128))
    if sold_rooms <= 0 and total_rooms > 0 and occ_val > 0:
        sold_rooms = int(round(total_rooms * occ_val / 100.0))
    revpar_day = int(round(adr_val * occ_val / 100.0)) if adr_val > 0 else 0

    # 當日雙館餐飲來客數
    fb_today = ctx["compute_fb_mtd"](d_str, d_str, _dummy_hotel=current_hotel) or {}
    bf_tot = _safe_int(fb_today.get("total_act_bf", 0))
    af_tot = _safe_int(fb_today.get("total_act_af", 0))
    bf_zq = _safe_int(fb_today.get("bf_zq_act", 0))
    bf_th = _safe_int(fb_today.get("bf_theme_act", 0))

    if occ_val >= 90.0:
        st.success(f"🎉 滿房達標！{sel_date.strftime('%m/%d')} 住房率達 **{occ_val:.1f}%**")

    # 2x2 Hero KPI Grid（100% 真實自動資料）
    st.markdown(
        f"""
<div class="m-kpi-grid">
<div class="m-kpi-card">
<div class="m-kpi-label">🏨 當日住房率 (OCC)</div>
<div class="m-kpi-val">{occ_val:.1f}%</div>
<div class="m-kpi-sub">已售 {sold_rooms} / 可售 {total_rooms} 間</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">💵 平均房價 (ADR)</div>
<div class="m-kpi-val">NT$ {adr_val:,}</div>
<div class="m-kpi-sub">RevPAR NT$ {revpar_day:,}</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">💰 當日客房營收</div>
<div class="m-kpi-val">NT$ {rev_val:,}</div>
<div class="m-kpi-sub">本月累計 NT$ {int(m_sum.get('rev', 0)):,}</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">🥐 雙館當日餐飲</div>
<div class="m-kpi-val">{bf_tot + af_tot} 人</div>
<div class="m-kpi-sub">早 {bf_tot} (站{bf_zq}/主{bf_th})｜午 {af_tot}</div>
</div>
</div>
        """,
        unsafe_allow_html=True,
    )

    # 本月進度摘要卡
    _, last_day = calendar.monthrange(sel_date.year, sel_date.month)
    days_passed = sel_date.day if (sel_date.year, sel_date.month) == (today.year, today.month) else last_day
    pct_month = min(100, int(round(days_passed / last_day * 100)))
    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">📊 {sel_date.month} 月營運累計摘要</div>
<span class="m-badge-muted">月進度 {pct_month}%</span>
</div>
<div class="m-bar-bg"><div class="m-bar-fill" style="width:{pct_month}%; background:#1f2c56;"></div></div>
<div class="m-tri-grid">
<div class="m-tri-cell">
<div class="m-tri-label">月均 OCC</div>
<div class="m-tri-val" style="color:#1f2c56;">{float(m_sum.get('avg_occ', 0.0)):.1f}%</div>
</div>
<div class="m-tri-cell">
<div class="m-tri-label">月均 ADR</div>
<div class="m-tri-val" style="color:#0f172a;">${int(round(m_sum.get('avg_adr', 0))):,}</div>
</div>
<div class="m-tri-cell">
<div class="m-tri-label">滿房 ≥90%</div>
<div class="m-tri-val" style="color:#16a34a;">{int(m_sum.get('occ90_days', 0))} 天</div>
</div>
</div>
</div>
        """,
        unsafe_allow_html=True,
    )

    # Rolling CPOR 精簡利潤卡（站前館真實計算；主題館呈現友善提示）
    if current_hotel == "主題館":
        _render_theme_compact_notice(
            "💎 即時利潤戰情看板 (Rolling CPOR)",
            "整合全館各部門（房務、櫃台、餐飲）當月預估採購成本與在手入住房數，即時衡量單房變動成本 (CPOR) 與單房淨利。",
            [
                "即時追蹤單房變動成本 (Rolling CPOR)",
                "結合 ADR 與餐飲攤提營收計算單房即時淨利 (Net TRevPOR)",
                "監控今日各部門即時請購總額 (Live Spend)",
            ],
        )
    else:
        try:
            df_purch = ctx["_get_all_purchase_clean"]()
            occ_df = ctx["_get_occ_data_cached_v2"](hotel_type=current_hotel)
            carts = [
                (ctx["fetch_hk_daily_purchase_report"](), "房務"),
                (ctx["fetch_fd_daily_purchase_report"](), "櫃台"),
                (ctx["fetch_cs_daily_purchase_report"](), "工務"),
                (ctx["fetch_thepeak_daily_purchase_report"](), "The Peak"),
                (ctx["fetch_4fhh_daily_purchase_report"](), "Happy Hour"),
            ]
            all_live = []
            live_spend = 0.0
            for df_daily, default_dept in carts:
                if df_daily is not None and not df_daily.empty:
                    d_col = next((c for c in df_daily.columns if "日期" in str(c)), None)
                    i_col = next((c for c in df_daily.columns if "品項" in str(c) or "品名" in str(c)), None)
                    p_col = next((c for c in df_daily.columns if any(k in str(c) for k in ["總價", "金額", "小計"])), None)
                    if d_col and i_col and p_col:
                        df_conv = pd.DataFrame()
                        df_conv["_date_parsed"] = pd.to_datetime(df_daily[d_col], errors="coerce")
                        df_conv = df_conv[df_conv["_date_parsed"].notna()].copy()
                        df_conv["_date_str"] = df_conv["_date_parsed"].dt.strftime("%Y-%m-%d")
                        df_conv["_item"] = df_daily.loc[df_conv.index, i_col].astype(str).str.strip()
                        df_conv["_price"] = pd.to_numeric(
                            df_daily.loc[df_conv.index, p_col].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                            errors="coerce",
                        ).fillna(0.0)
                        df_conv["_dept"] = default_dept
                        all_live.append(df_conv)
                        live_spend += float(df_conv.loc[df_conv["_date_str"] == d_str, "_price"].sum())

            if all_live:
                df_live_comb = pd.concat(all_live, ignore_index=True)
                if not df_purch.empty:
                    off_pairs = set(zip(df_purch["_date_str"], df_purch["_item"]))
                    mask_new = df_live_comb.apply(lambda r: (r["_date_str"], r["_item"]) not in off_pairs, axis=1)
                    df_purch = pd.concat([df_purch, df_live_comb[mask_new]], ignore_index=True)
                else:
                    df_purch = df_live_comb

            all_depts = df_purch["_dept"].dropna().astype(str).unique().tolist() if not df_purch.empty else []
            target_depts = ["房務", "櫃台"]
            for d_name in all_depts:
                du = d_name.upper()
                if any(k in du for k in ["PEAK", "餐廳", "THEPEAK", "餐飲", "HH", "HAPPY", "4FHH", "歡樂時光"]):
                    if d_name not in target_depts:
                        target_depts.append(d_name)

            m_start_str = sel_date.replace(day=1).strftime("%Y-%m-%d")
            fb_mtd = ctx["compute_fb_mtd"](m_start_str, d_str, _dummy_hotel=current_hotel) or {}
            zq_g = _safe_int(fb_mtd.get("bf_zq_act", 0)) + _safe_int(fb_mtd.get("af_zq_act", 0))
            th_g = _safe_int(fb_mtd.get("bf_theme_act", 0)) + _safe_int(fb_mtd.get("af_theme_act", 0))
            fb_share = (zq_g / (zq_g + th_g)) if (zq_g + th_g) > 0 else 0.55

            tot_cost = 0.0
            eom_rooms = 0
            mtd_rooms = 0
            for dept in target_depts:
                df_d = df_purch[df_purch["_dept"].str.strip() == dept] if not df_purch.empty else pd.DataFrame()
                met = ctx["compute_dept_cpr_metrics"](df_d, occ_df, sel_date, dept_label=dept)
                d_cost = float(met.get("eom_forecast_cost", 0.0))
                if dept not in ["房務", "櫃台", "工務"]:
                    d_cost *= fb_share
                tot_cost += d_cost
                if dept == "櫃台":
                    eom_rooms = met.get("eom_rooms", 0)
                    mtd_rooms = met.get("mtd_rooms", 0)

            cpor_val = (tot_cost / eom_rooms) if eom_rooms > 0 else 0.0
            rest_mrev = float(fb_mtd.get("revenue", 0.0)) * fb_share
            fb_revpor = (rest_mrev / mtd_rooms) if mtd_rooms > 0 else 0.0
            net_profit = (adr_val + fb_revpor - cpor_val) if (adr_val + fb_revpor) > 0 else 0.0

            st.markdown(
                f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">💎 即時利潤戰情 (Rolling CPOR)</div>
<span class="m-badge-ok">自動核算</span>
</div>
<div class="m-tri-grid">
<div class="m-tri-cell">
<div class="m-tri-label">動態單房成本</div>
<div class="m-tri-val" style="color:#dc2626;">${int(round(cpor_val)):,}</div>
</div>
<div class="m-tri-cell">
<div class="m-tri-label">單房即時淨利</div>
<div class="m-tri-val" style="color:#16a34a;">${int(round(net_profit)):,}</div>
</div>
<div class="m-tri-cell">
<div class="m-tri-label">當日各部採購</div>
<div class="m-tri-val" style="color:#d97706;">${int(round(live_spend)):,}</div>
</div>
</div>
</div>
                """,
                unsafe_allow_html=True,
            )
        except Exception:
            pass

    # 當日營運日誌快速預覽
    cur_log = (ctx["get_daily_log"](d_str) or "").strip()
    preview_txt = html.escape(cur_log[:160] + ("..." if len(cur_log) > 160 else "")) if cur_log else "尚無當日營運日誌，可點選上方「✍️ 填報」快速紀錄。"
    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">📝 {sel_date.strftime('%m/%d')} 營運日誌摘要</div>
<span class="m-badge-muted">{'已紀錄' if cur_log else '未填寫'}</span>
</div>
<div style="font-size:13px; color:#334155; line-height:1.55; white-space:pre-line;">{preview_txt}</div>
</div>
        """,
        unsafe_allow_html=True,
    )


def _render_mobile_input(ctx: dict):
    current_hotel = ctx["current_hotel"]
    today = datetime.date.today()

    if "mobile_input_date" not in st.session_state:
        st.session_state["mobile_input_date"] = today

    # 獨立的現場填報日期（滿寬呈現，絕不受採購月底鎖定影響！）
    inp_date = st.date_input(
        "📅 填報日期（預設今日）",
        value=st.session_state["mobile_input_date"],
        key="m_input_date_widget",
    )
    if inp_date != st.session_state["mobile_input_date"]:
        st.session_state["mobile_input_date"] = inp_date
        st.rerun()
    if st.session_state["mobile_input_date"] != today:
        if st.button("📍 跳回今日", key="m_input_today", use_container_width=True):
            st.session_state["mobile_input_date"] = today
            st.rerun()

    d_str = st.session_state["mobile_input_date"].strftime("%Y-%m-%d")

    if current_hotel == "採購":
        sub_mode = "🍽️ 現場餐期登記"
    else:
        sub_mode = _render_nav_pills(["🍽️ 現場餐期登記", "📝 每日營運日誌"], "m_input_sub", "🍽️ 現場餐期登記")

    if sub_mode == "🍽️ 現場餐期登記":
        raw_map, err = ctx["fetch_fb_report_raw_dict"]()
        if err:
            st.warning(f"⚠️ 讀取 f&b_report 發生提示：{err}")
        row_vals = raw_map.get(d_str, [""] * 9)
        th_bf_est, th_bf_act = _safe_int(row_vals[0]), _safe_int(row_vals[1])
        zq_bf_est, zq_bf_act = _safe_int(row_vals[2]), _safe_int(row_vals[3])
        th_af_est, th_af_act = _safe_int(row_vals[4]), _safe_int(row_vals[5])
        zq_af_est, zq_af_act = _safe_int(row_vals[6]), _safe_int(row_vals[7])
        hh_act = _safe_int(row_vals[8])

        st.markdown(
            f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">🍽️ {d_str} 餐期現況總覽</div>
<span class="m-badge-ok">局部安全寫入</span>
</div>
<div class="m-tri-grid">
<div class="m-tri-cell">
<div class="m-tri-label">🥐 雙館早餐</div>
<div class="m-tri-val">{th_bf_act + zq_bf_act} 人</div>
</div>
<div class="m-tri-cell">
<div class="m-tri-label">🍰 雙館午茶</div>
<div class="m-tri-val">{th_af_act + zq_af_act} 人</div>
</div>
<div class="m-tri-cell">
<div class="m-tri-label">🥂 4F HH</div>
<div class="m-tri-val">{hh_act} 人</div>
</div>
</div>
</div>
            """,
            unsafe_allow_html=True,
        )

        meal_tab = _render_nav_pills(["🥐 早餐", "🍰 下午茶", "🥂 4F HH"], "m_meal_tab", "🥐 早餐")
        st.caption("💡 只會更新您目前選擇的餐期欄位，絕不會覆蓋其他同仁填寫的餐期或 PMS 預估人數。")

        if meal_tab == "🥐 早餐":
            with st.form(key=f"m_fb_bf_form_{d_str}"):
                st.markdown(f"**站前館早餐** *(PMS 預估 `{zq_bf_est}` 人)*")
                in_zq_bf = st.number_input("站前館早餐實際人數", min_value=0, max_value=999, value=zq_bf_act, step=1)
                st.markdown(f"**主題館早餐** *(PMS 預估 `{th_bf_est}` 人)*")
                in_th_bf = st.number_input("主題館早餐實際人數", min_value=0, max_value=999, value=th_bf_act, step=1)
                if st.form_submit_button("💾 儲存早餐實際人數", type="primary", use_container_width=True):
                    ok, msg = _update_fb_report_partial(
                        d_str, "breakfast", {"th_bf_act": int(in_th_bf), "zq_bf_act": int(in_zq_bf)}
                    )
                    if ok:
                        st.toast(f"✅ {d_str} 早餐人數已安全存檔！")
                        st.rerun()
                    else:
                        st.error(f"儲存失敗：{msg}")

        elif meal_tab == "🍰 下午茶":
            with st.form(key=f"m_fb_af_form_{d_str}"):
                st.markdown(f"**站前館下午茶** *(PMS 預估 `{zq_af_est}` 人)*")
                in_zq_af = st.number_input("站前館下午茶實際人數", min_value=0, max_value=999, value=zq_af_act, step=1)
                st.markdown(f"**主題館下午茶** *(PMS 預估 `{th_af_est}` 人)*")
                in_th_af = st.number_input("主題館下午茶實際人數", min_value=0, max_value=999, value=th_af_act, step=1)
                if st.form_submit_button("💾 儲存下午茶實際人數", type="primary", use_container_width=True):
                    ok, msg = _update_fb_report_partial(
                        d_str, "afternoon", {"th_af_act": int(in_th_af), "zq_af_act": int(in_zq_af)}
                    )
                    if ok:
                        st.toast(f"✅ {d_str} 下午茶人數已安全存檔！")
                        st.rerun()
                    else:
                        st.error(f"儲存失敗：{msg}")

        else:
            with st.form(key=f"m_fb_hh_form_{d_str}"):
                st.markdown("**4F 餐車 Happy Hour 實際來客數**")
                in_hh = st.number_input("4F HH 實際人數", min_value=0, max_value=999, value=hh_act, step=1)
                if st.form_submit_button("💾 儲存 Happy Hour 實際人數", type="primary", use_container_width=True):
                    ok, msg = _update_fb_report_partial(d_str, "hh", {"hh_act": int(in_hh)})
                    if ok:
                        st.toast(f"✅ {d_str} 4F HH 人數已安全存檔！")
                        st.rerun()
                    else:
                        st.error(f"儲存失敗：{msg}")

    else:
        # 每日營運日誌
        log_ss_key = f"m_log_draft_{current_hotel}_{d_str}"
        if log_ss_key not in st.session_state:
            st.session_state[log_ss_key] = ctx["get_daily_log"](d_str) or ""

        st.caption("🏷️ 點選快速插入分類標籤：")
        t1, t2, t3, t4 = st.columns(4)
        for col_w, tag in zip([t1, t2, t3, t4], ["交班重點", "客訴處理", "VIP接待", "設備修繕"]):
            if col_w.button(f"#{tag[:2]}", key=f"m_tag_{d_str}_{tag}", use_container_width=True, help=f"插入 [{tag}]"):
                cur_t = st.session_state[log_ss_key].strip()
                st.session_state[log_ss_key] = (cur_t + "\n" if cur_t else "") + f"[{tag}] "
                st.rerun()

        with st.form(key=f"m_daily_log_form_{current_hotel}_{d_str}"):
            log_text = st.text_area(
                f"📝 {current_hotel} — {d_str} 營運日誌",
                value=st.session_state[log_ss_key],
                height=220,
                placeholder="輸入交班重點、客訴處理、VIP 接待或設備修繕紀錄...",
            )
            if st.form_submit_button("💾 儲存今日營運日誌", type="primary", use_container_width=True):
                ctx["save_daily_log"](d_str, log_text.strip())
                st.session_state[log_ss_key] = log_text.strip()
                st.toast(f"✅ {d_str} 營運日誌已存檔！")
                st.rerun()


def _load_month_fb_purchases(ctx: dict, year: int, month: int):
    """與 app.py 採購分析完全一致的 (dept, YYYY-MM) 官方+日報無縫縫合邏輯。"""
    df_p = ctx["get_purchase_data_cached"]()
    if df_p is None:
        df_p = pd.DataFrame()
    else:
        df_p = df_p.copy()

    def _parse_d(val):
        if pd.isna(val):
            return None
        s = re.sub(r"\.0$", "", str(val).strip())
        if not s or s in ("nan", "None", "NaT"):
            return None
        if "/" in s:
            parts = s.split("/")
            if len(parts) == 3 and parts[0].isdigit():
                y_v = int(parts[0])
                if y_v < 1900:
                    y_v += 1911
                try:
                    return datetime.date(y_v, int(parts[1]), int(parts[2]))
                except Exception:
                    pass
        try:
            return pd.to_datetime(val).date()
        except Exception:
            return None

    def _bucket(d):
        du = str(d).upper()
        if "4" in du or any(k in du for k in ["HH", "HAPPY", "歡樂時光"]):
            return "Happy Hour"
        if any(k in du for k in ["PEAK", "餐廳", "THEPEAK", "餐飲"]):
            return "The Peak"
        return str(d).strip()

    rows = []
    official_dept_ym = set()
    if not df_p.empty:
        df_p.columns = df_p.columns.astype(str).str.strip()
        d_col = next((c for c in df_p.columns if "日期" in c or "Date" in c), None)
        dept_col = next((c for c in df_p.columns if "部門" in c or "Dept" in c or "工地" in c), None)
        tot_col = next((c for c in df_p.columns if "小計" in c or "金額" in c or "Total" in c), None)
        item_col = next((c for c in df_p.columns if any(k in c for k in ["品名", "品項", "項次說明", "項目", "Item"])), None)
        if d_col and dept_col and tot_col:
            df_p["_d"] = df_p[d_col].apply(_parse_d)
            df_p = df_p[df_p["_d"].notna()].copy()
            df_p["_bucket"] = df_p[dept_col].apply(_bucket)
            df_p["_ym"] = df_p["_d"].apply(lambda x: f"{x.year:04d}-{x.month:02d}")
            df_p["_amt"] = pd.to_numeric(
                df_p[tot_col].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce"
            ).fillna(0.0)
            official_dept_ym = set(zip(df_p["_bucket"], df_p["_ym"]))
            ym_str = f"{year:04d}-{month:02d}"
            sub = df_p[(df_p["_ym"] == ym_str) & (df_p["_bucket"].isin(["The Peak", "Happy Hour"]))]
            for _, r in sub.iterrows():
                rows.append(
                    {
                        "date": r["_d"].strftime("%Y-%m-%d"),
                        "bucket": r["_bucket"],
                        "item": str(r[item_col]).strip() if item_col else "未分類",
                        "amount": float(r["_amt"]),
                    }
                )

    ym_str = f"{year:04d}-{month:02d}"
    for df_live, b_name in [
        (ctx["fetch_thepeak_daily_purchase_report"](), "The Peak"),
        (ctx["fetch_4fhh_daily_purchase_report"](), "Happy Hour"),
    ]:
        if (b_name, ym_str) in official_dept_ym:
            continue
        if df_live is not None and not df_live.empty:
            d_c = next((c for c in df_live.columns if "日期" in str(c) or "叫貨" in str(c)), None)
            i_c = next((c for c in df_live.columns if "品項" in str(c) or "品名" in str(c)), None)
            p_c = next((c for c in df_live.columns if any(k in str(c) for k in ["總價", "金額", "小計"])), None)
            if d_c and p_c:
                df_l = df_live.copy()
                df_l["_d"] = pd.to_datetime(df_l[d_c], errors="coerce").dt.date
                df_l = df_l[df_l["_d"].notna()].copy()
                df_l["_ym"] = df_l["_d"].apply(lambda x: f"{x.year:04d}-{x.month:02d}")
                sub_l = df_l[df_l["_ym"] == ym_str]
                for _, r in sub_l.iterrows():
                    amt = _safe_float(r[p_c])
                    rows.append(
                        {
                            "date": r["_d"].strftime("%Y-%m-%d"),
                            "bucket": b_name,
                            "item": str(r[i_c]).strip() if i_c else "未分類",
                            "amount": amt,
                        }
                    )

    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["date", "bucket", "item", "amount"])


def _render_mobile_purchase(ctx: dict):
    current_hotel = ctx["current_hotel"]
    months = _recent_months_list(12)
    sel_ym = st.selectbox("📅 選擇分析月份", months, index=0, key="m_purch_ym")
    y, m = int(sel_ym[:4]), int(sel_ym[5:7])
    _, last_day = calendar.monthrange(y, m)
    m_start = f"{y:04d}-{m:02d}-01"
    m_end = f"{y:04d}-{m:02d}-{last_day:02d}"

    fb_mtd = ctx["compute_fb_mtd"](m_start, m_end, _dummy_hotel=current_hotel) or {}

    if current_hotel == "主題館":
        th_g = _safe_int(fb_mtd.get("bf_theme_act", 0)) + _safe_int(fb_mtd.get("af_theme_act", 0))
        zq_g = _safe_int(fb_mtd.get("bf_zq_act", 0)) + _safe_int(fb_mtd.get("af_zq_act", 0))
        tot_g = th_g + zq_g
        th_pct = (th_g / tot_g * 100.0) if tot_g > 0 else 0.0
        st.markdown(
            f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">🍽️ {sel_ym} 雙館餐飲營運 (MTD)</div>
<span class="m-badge-ok">雙館共用</span>
</div>
<div class="m-kpi-grid" style="margin-bottom:0;">
<div class="m-kpi-card">
<div class="m-kpi-label">主題館實際來客</div>
<div class="m-kpi-val">{th_g:,} 人</div>
<div class="m-kpi-sub">佔雙館 {th_pct:.1f}%</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">兩館餐飲結算營收</div>
<div class="m-kpi-val">NT$ {int(round(fb_mtd.get('revenue', 0))):,}</div>
<div class="m-kpi-sub">站前 {zq_g:,} 人</div>
</div>
</div>
</div>
            """,
            unsafe_allow_html=True,
        )
        _render_theme_compact_notice(
            "💰 全館部門採購分析",
            "跨部門整合會計採購大表與現場請購日報，協助控管預算執行率、單房/單客耗用成本與單價異常。",
            [
                "全館各部門採購總額追蹤與剩餘預算預警",
                "品項級單價異常偵測與供應商報價漲跌幅追蹤",
                "雙引擎智慧叫貨預測與各週採購配額建議",
            ],
        )
        return

    df_fb_p = _load_month_fb_purchases(ctx, y, m)
    peak_df = df_fb_p[df_fb_p["bucket"] == "The Peak"] if not df_fb_p.empty else pd.DataFrame()
    hh_df = df_fb_p[df_fb_p["bucket"] == "Happy Hour"] if not df_fb_p.empty else pd.DataFrame()

    peak_spent = float(peak_df["amount"].sum()) if not peak_df.empty else 0.0
    hh_spent = float(hh_df["amount"].sum()) if not hh_df.empty else 0.0

    df_fb_daily = ctx["get_combined_fb_daily_df"](y, m, current_hotel)
    peak_guests = int(df_fb_daily["peak_guests"].sum()) if (df_fb_daily is not None and not df_fb_daily.empty and "peak_guests" in df_fb_daily.columns) else 0
    hh_guests = int(df_fb_daily["hh_act"].sum()) if (df_fb_daily is not None and not df_fb_daily.empty and "hh_act" in df_fb_daily.columns) else 0

    peak_cpg = (peak_spent / peak_guests) if peak_guests > 0 else 0.0
    hh_cpg = (hh_spent / hh_guests) if hh_guests > 0 else 0.0

    target_cpg = 150
    cpg_badge_cls = "m-badge-ok" if peak_cpg <= target_cpg else ("m-badge-warn" if peak_cpg <= target_cpg * 1.08 else "m-badge-err")
    cpg_label = f"CPG ${int(round(peak_cpg))} {'安全' if peak_cpg <= target_cpg else '偏高'}"
    cpg_pct = min(100, int(round((peak_cpg / target_cpg) * 100))) if target_cpg > 0 else 0
    bar_color = "#16a34a" if peak_cpg <= target_cpg else ("#d97706" if peak_cpg <= target_cpg * 1.08 else "#dc2626")

    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">🏰 The Peak (37F) — {sel_ym}</div>
<span class="{cpg_badge_cls}">{cpg_label}</span>
</div>
<div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:4px;">
<div>
<div class="m-kpi-label">本月總採購額</div>
<div class="m-kpi-val">NT$ {int(round(peak_spent)):,}</div>
</div>
<div style="text-align:right;">
<div class="m-kpi-label">已發生來客</div>
<div style="font-size:18px; font-weight:800; color:#1f2c56;" class="m-num">{peak_guests:,} 人</div>
</div>
</div>
<div class="m-bar-bg"><div class="m-bar-fill" style="width:{cpg_pct}%; background:{bar_color};"></div></div>
<div style="display:flex; justify-content:space-between; font-size:12px; color:#64748b;" class="m-num">
<span>目前 CPG：NT$ {peak_cpg:.1f} / 客</span>
<span>目標上限：NT$ {target_cpg}</span>
</div>
</div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">🥂 Happy Hour (4F) — {sel_ym}</div>
<span class="m-badge-muted">每客 NT$ {int(round(hh_cpg)):,}</span>
</div>
<div class="m-kpi-grid" style="margin-bottom:0;">
<div class="m-kpi-card">
<div class="m-kpi-label">本月總採購額</div>
<div class="m-kpi-val">NT$ {int(round(hh_spent)):,}</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">本月總來客數</div>
<div class="m-kpi-val">{hh_guests:,} 人</div>
</div>
</div>
</div>
        """,
        unsafe_allow_html=True,
    )

    # 行動卡片流：本月 Top 10 採購品項（取代寬表格）
    if not df_fb_p.empty:
        grp = (
            df_fb_p.groupby(["item", "bucket"], as_index=False)
            .agg(total=("amount", "sum"), cnt=("amount", "count"))
            .sort_values("total", ascending=False)
            .head(10)
        )
        tot_all = max(peak_spent + hh_spent, 1.0)
        rows_html = ""
        for _, r in grp.iterrows():
            share = r["total"] / tot_all * 100.0
            rows_html += (
                f'<div class="m-list-row">'
                f'<div><div class="m-list-main">{html.escape(str(r["item"]))}</div>'
                f'<div class="m-list-sub">{html.escape(str(r["bucket"]))} · 叫貨 {int(r["cnt"])} 次</div></div>'
                f'<div class="m-list-right"><div class="m-list-main">NT$ {int(round(r["total"])):,}</div>'
                f'<div class="m-list-sub">{share:.1f}%</div></div>'
                f"</div>"
            )
        st.markdown(
            f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">📋 本月前 10 大採購品項</div>
<span class="m-badge-muted">共 {len(df_fb_p)} 筆</span>
</div>
{rows_html}
</div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.info(f"💡 {sel_ym} 尚無 The Peak / 4FHH 採購紀錄。")


def _render_veg_price_cards(ctx: dict):
    sp_df = ctx["fetch_supplier_prices"]()
    if sp_df is None or sp_df.empty:
        st.info("💡 尚未載入菜商報價資料 (`supplier_prices`)。")
        return
    periods = sorted(sp_df["period_dt"].unique())
    if len(periods) < 2:
        st.info("💡 菜商報價期數不足兩期，無法計算漲跌幅。")
        return

    latest_p, prev_p = periods[-1], periods[-2]
    df_cur = sp_df[sp_df["period_dt"] == latest_p].groupby("item_name", as_index=False).agg({"price": "mean", "unit": "first"})
    df_prv = sp_df[sp_df["period_dt"] == prev_p].groupby("item_name", as_index=False).agg({"price": "mean"})
    merged = pd.merge(df_cur, df_prv, on="item_name", how="inner", suffixes=("_cur", "_prv"))
    merged = merged[merged["price_prv"] > 0].copy()
    merged["chg_pct"] = (merged["price_cur"] - merged["price_prv"]) / merged["price_prv"] * 100.0

    risers = merged.sort_values("chg_pct", ascending=False).head(8)
    drops = merged.sort_values("chg_pct", ascending=True).head(5)

    def _build_rows(sub_df):
        out = ""
        for _, r in sub_df.iterrows():
            chg = float(r["chg_pct"])
            color = "#dc2626" if chg > 10 else ("#d97706" if chg > 0 else ("#16a34a" if chg < 0 else "#64748b"))
            sign = "+" if chg > 0 else ""
            out += (
                f'<div class="m-list-row">'
                f'<div><div class="m-list-main">{html.escape(str(r["item_name"]))}</div>'
                f'<div class="m-list-sub">前次 ${r["price_prv"]:.0f} → 最新 ${r["price_cur"]:.0f} / {html.escape(str(r["unit"]))}</div></div>'
                f'<div class="m-list-right" style="color:{color}; font-weight:800; font-size:14px;">{sign}{chg:.1f}%</div>'
                f"</div>"
            )
        return out

    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">🚨 菜價漲幅警示 Top 8</div>
<span class="m-badge-muted">{latest_p} vs {prev_p}</span>
</div>
{_build_rows(risers)}
</div>
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">📉 菜價跌幅優惠 Top 5</div>
<span class="m-badge-ok">議價參考</span>
</div>
{_build_rows(drops)}
</div>
        """,
        unsafe_allow_html=True,
    )


def _render_mobile_trend(ctx: dict):
    current_hotel = ctx["current_hotel"]
    if current_hotel == "採購":
        _render_veg_price_cards(ctx)
        return

    months = _recent_months_list(12)
    sel_ym = st.selectbox("📅 選擇趨勢月份", months, index=0, key="m_trend_ym")
    y, m = int(sel_ym[:4]), int(sel_ym[5:7])

    sub_view = _render_nav_pills(["📈 住房走勢", "🌍 國籍與渠道", "🛒 菜價紅綠燈"], "m_trend_sub", "📈 住房走勢")

    if sub_view == "🛒 菜價紅綠燈":
        _render_veg_price_cards(ctx)
        return

    if sub_view == "📈 住房走勢":
        m_cur = ctx["fetch_month_summary_by_hotel"](y, m, current_hotel)
        m_ly = ctx["fetch_month_summary_by_hotel"](y - 1, m, current_hotel)

        occ_diff = float(m_cur.get("avg_occ", 0)) - float(m_ly.get("avg_occ", 0))
        adr_diff = float(m_cur.get("avg_adr", 0)) - float(m_ly.get("avg_adr", 0))
        rev_diff = float(m_cur.get("rev", 0)) - float(m_ly.get("rev", 0))
        has_ly = not m_ly.get("df", pd.DataFrame()).empty

        def _delta_html(val, unit="", is_money=False):
            if not has_ly:
                return '<span style="color:#64748b;">去年同期無資料</span>'
            clr = "#16a34a" if val >= 0 else "#dc2626"
            arrow = "▲" if val >= 0 else "▼"
            txt = f"${abs(int(round(val))):,}" if is_money else f"{abs(val):.1f}{unit}"
            return f'<span style="color:{clr}; font-weight:700;">{arrow} {txt} YoY</span>'

        st.markdown(
            f"""
<div class="m-kpi-grid">
<div class="m-kpi-card">
<div class="m-kpi-label">{sel_ym} 平均 OCC</div>
<div class="m-kpi-val">{float(m_cur.get('avg_occ', 0)):.1f}%</div>
<div class="m-kpi-sub">{_delta_html(occ_diff, '%')}</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">{sel_ym} 平均 ADR</div>
<div class="m-kpi-val">NT$ {int(round(m_cur.get('avg_adr', 0))):,}</div>
<div class="m-kpi-sub">{_delta_html(adr_diff, is_money=True)}</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">月總客房營收</div>
<div class="m-kpi-val">${int(round(m_cur.get('rev', 0)))/10000:.1f}萬</div>
<div class="m-kpi-sub">{_delta_html(rev_diff, is_money=True)}</div>
</div>
<div class="m-kpi-card">
<div class="m-kpi-label">月均 RevPAR</div>
<div class="m-kpi-val">NT$ {int(round(m_cur.get('revpar', 0))):,}</div>
<div class="m-kpi-sub">滿房 ≥90% 共 {int(m_cur.get('occ90_days', 0))} 天</div>
</div>
</div>
            """,
            unsafe_allow_html=True,
        )

        # 純 CSS 行動柱狀圖（最近 7 天有售房之紀錄，零捲動干擾）
        df_m = m_cur.get("df", pd.DataFrame())
        if df_m is not None and not df_m.empty and "date" in df_m.columns and "occ_rate" in df_m.columns:
            df_valid = df_m[df_m["occ_rate"] > 0].sort_values("date").tail(7)
            if not df_valid.empty:
                bars_html = ""
                for _, r in df_valid.iterrows():
                    occ_r = float(r["occ_rate"])
                    bar_h = max(14, min(115, int(round(occ_r * 1.1))))
                    bar_clr = "#16a34a" if occ_r >= 90 else ("#1f2c56" if occ_r >= 80 else "#f19e38")
                    d_lbl = str(r["date"])[5:]
                    bars_html += (
                        f'<div class="m-chart-col">'
                        f'<div class="m-chart-val">{occ_r:.0f}%</div>'
                        f'<div class="m-chart-bar" style="height:{bar_h}px; background:{bar_clr};"></div>'
                        f'<div class="m-chart-lbl">{html.escape(d_lbl)}</div>'
                        f"</div>"
                    )
                st.markdown(
                    f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">📊 近期 7 日住房率走勢</div>
<span class="m-badge-muted">綠 ≥90%｜藍 ≥80%</span>
</div>
<div class="m-chart-wrap">{bars_html}</div>
</div>
                    """,
                    unsafe_allow_html=True,
                )
        return

    # 🌍 國籍 Top 5 + 📉 渠道 Top 5（純 CSS 水平條狀排行）
    ym_compact = f"{y:04d}{m:02d}"
    df_nat = ctx["_get_cached_sheet_v3"]("nationality_report", hotel_type=current_hotel)
    nat_rows_html = ""
    nat_total = 0
    if df_nat is not None and not df_nat.empty:
        df_n = df_nat.copy()
        df_n.columns = [str(c).strip().lower() for c in df_n.columns]

        def _ym_norm(v):
            s = str(v).strip().replace(".0", "")
            if re.match(r"^\d{6}$", s):
                return s
            m2 = re.match(r"^(\d{4})[-/](\d{1,2})", s)
            if m2:
                return f"{m2.group(1)}{int(m2.group(2)):02d}"
            return s[:6] if re.match(r"^\d{8}$", s) else ""

        if "date" in df_n.columns and "month" not in df_n.columns:
            df_n["month"] = df_n["date"].apply(_ym_norm).replace("", pd.NA).ffill().fillna("")
        elif "month" in df_n.columns:
            df_n["month"] = df_n["month"].apply(_ym_norm).replace("", pd.NA).ffill().fillna("")

        if {"month", "nation", "person"}.issubset(set(df_n.columns)):
            df_t = df_n[df_n["month"].astype(str).str.strip() == ym_compact].copy()
            df_t["person"] = pd.to_numeric(df_t["person"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
            df_t = df_t[(df_t["person"] > 0) & (~df_t["nation"].astype(str).str.contains("合計|總計|小計|Total", na=False, case=False))]
            if not df_t.empty:
                df_g = df_t.groupby("nation", as_index=False)["person"].sum().sort_values("person", ascending=False)
                nat_total = int(round(df_g["person"].sum()))
                for _, r in df_g.head(5).iterrows():
                    pct = (r["person"] / nat_total * 100.0) if nat_total > 0 else 0.0
                    w_pct = min(100, int(round(pct * 2.2)))
                    nat_rows_html += (
                        f'<div style="margin-bottom:9px;">'
                        f'<div style="display:flex; justify-content:space-between; font-size:13px;" class="m-num">'
                        f'<strong>{html.escape(str(r["nation"]))}</strong>'
                        f'<span style="color:#475569;">{int(round(r["person"])):,} 人 ({pct:.1f}%)</span></div>'
                        f'<div class="m-bar-bg"><div class="m-bar-fill" style="width:{w_pct}%; background:#1f2c56;"></div></div>'
                        f"</div>"
                    )

    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">🌍 {sel_ym} 國籍 Top 5</div>
<span class="m-badge-muted">總計 {nat_total:,} 人</span>
</div>
{nat_rows_html or '<div style="font-size:13px; color:#64748b;">尚無該月國籍數據</div>'}
</div>
        """,
        unsafe_allow_html=True,
    )

    # OTA 渠道 Top 5
    df_ch = ctx["_get_cached_sheet_v3"]("marketing_channel_data", hotel_type=current_hotel)
    ch_rows_html = ""
    ch_total = 0
    if df_ch is not None and not df_ch.empty:
        df_c = df_ch.copy()
        df_c.columns = [str(c).strip().lower() for c in df_c.columns]
        df_c = df_c.loc[:, ~df_c.columns.duplicated()]
        if "date" in df_c.columns:
            df_c["date"] = df_c["date"].replace("", pd.NA).bfill().astype(str).str.strip()
            df_c = df_c[df_c["date"].str.startswith(sel_ym)]
            comp_c = next((c for c in df_c.columns if "company" in c or "channel" in c), None)
            rm_c = next((c for c in df_c.columns if "room" in c), None)
            if comp_c and rm_c and not df_c.empty:
                df_c[rm_c] = pd.to_numeric(df_c[rm_c].astype(str).str.replace(",", ""), errors="coerce").fillna(0)

                def _norm_ch(v):
                    u = str(v).strip().upper()
                    if "AGODA" in u:
                        return "AGODA"
                    if "CTRIP" in u or "TRIP.COM" in u:
                        return "CTRIP / Trip.com"
                    if "BOOKING" in u and "BUTTON" not in u:
                        return "BOOKING"
                    if "EXPEDIA" in u:
                        return "EXPEDIA"
                    if "官網" in u or "BOOKINGBUTTON" in u:
                        return "官網"
                    return str(v).strip()

                df_c["_ch"] = df_c[comp_c].apply(_norm_ch)
                grp_c = df_c.groupby("_ch", as_index=False)[rm_c].sum().sort_values(rm_c, ascending=False)
                ch_total = int(round(grp_c[rm_c].sum()))
                for _, r in grp_c.head(5).iterrows():
                    pct = (r[rm_c] / ch_total * 100.0) if ch_total > 0 else 0.0
                    w_pct = min(100, int(round(pct * 1.6)))
                    ch_rows_html += (
                        f'<div style="margin-bottom:9px;">'
                        f'<div style="display:flex; justify-content:space-between; font-size:13px;" class="m-num">'
                        f'<strong>{html.escape(str(r["_ch"]))}</strong>'
                        f'<span style="color:#475569;">{int(round(r[rm_c])):,} 間 ({pct:.1f}%)</span></div>'
                        f'<div class="m-bar-bg"><div class="m-bar-fill" style="width:{w_pct}%; background:#f19e38;"></div></div>'
                        f"</div>"
                    )

    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">📉 {sel_ym} OTA 渠道 Top 5</div>
<span class="m-badge-muted">總計 {ch_total:,} 間</span>
</div>
{ch_rows_html or '<div style="font-size:13px; color:#64748b;">尚無該月渠道數據</div>'}
</div>
        """,
        unsafe_allow_html=True,
    )


def _render_mobile_more(ctx: dict):
    current_hotel = ctx["current_hotel"]
    login_role = ctx["login_role"]
    today = datetime.date.today()

    st.markdown(
        f"""
<div class="m-card">
<div class="m-card-head">
<div class="m-card-title">⚙️ 系統工具與進階功能</div>
<span class="m-badge-muted">權限：{html.escape(login_role)}</span>
</div>
<div style="font-size:13px; color:#475569; line-height:1.6;">
💡 <strong>加入手機主畫面小技巧</strong>：在手機瀏覽器網址列加上 <code>?view=mobile</code> 後加入主畫面，之後點開就會直接進入專屬手機版。
</div>
</div>
        """,
        unsafe_allow_html=True,
    )

    if current_hotel == "主題館":
        _render_theme_compact_notice(
            "房務 / 工務 / 櫃台 / 人事 / 免費服務優化",
            "以上 5 個進階管理模組目前僅串接站前館內部採購與人事資料，主題館暫無對應數據來源。",
            [
                "🧹 房務數據：清潔人效、布巾洗滌成本與備品庫存預警",
                "🔧 工務數據：修繕耗材紀律與突發維修支出追蹤",
                "🏢 櫃台數據：客用迎賓備品 CPOR 與雙引擎叫貨預測",
                "👥 人事概況：部門人力結構、薪資與獎金動態試算",
                "💡 免費服務優化：39 項贈品價值 vs. 成本四象限節流分析",
            ],
        )

    # 匯出本月完整 Markdown 報告
    st.markdown("##### 📤 匯出本月數據報告 (Markdown)")
    exp_ym = st.selectbox("選擇匯出月份", _recent_months_list(6), index=0, key="m_exp_ym")
    ey, em = int(exp_ym[:4]), int(exp_ym[5:7])
    if st.button("📥 產生本月 Markdown 報告", use_container_width=True, key="m_gen_md"):
        with st.spinner(f"正在彙整 {exp_ym} 數據報告..."):
            try:
                md_content = ctx["generate_monthly_export"](ey, em)
                st.download_button(
                    label=f"✅ 下載 Hotel_月份報告_{exp_ym}.md",
                    data=md_content,
                    file_name=f"Hotel_月份報告_{exp_ym}.md",
                    mime="text/markdown",
                    use_container_width=True,
                    type="primary",
                )
            except Exception as e:
                st.error(f"報告產生失敗：{e}")

    st.divider()
    st.caption("如需使用 13 頁店務會議 PPTX 產製、完整多表比對或人事獎金試算，請切換回電腦完整版：")
    c_m1, c_m2 = st.columns(2)
    with c_m1:
        if st.button("💻 切換電腦完整版", use_container_width=True, key="m_more_desktop"):
            st.session_state["ui_mode"] = "desktop"
            try:
                if "view" in st.query_params:
                    del st.query_params["view"]
            except Exception:
                pass
            st.rerun()
    with c_m2:
        if st.button("🔄 清除快取重同步", use_container_width=True, key="m_more_refresh"):
            st.cache_data.clear()
            st.toast("✅ 已清除快取並重新同步最新資料")
            st.rerun()

    st.write("")
    if st.button("🔒 安全登出（切換帳號）", use_container_width=True, key="m_logout_btn"):
        for k in ["authenticated", "hotel_type", "login_role"]:
            st.session_state.pop(k, None)
        st.rerun()


# ─────────────────────────────────────────────────────────────
# 4. 手機版總入口
# ─────────────────────────────────────────────────────────────
def render_mobile_app(ctx: dict):
    """
    由 app.py 在登入驗證後、桌面側邊欄建立前呼叫。
    ctx 為相依注入字典，確保零循環匯入並 100% 共用快取與商業邏輯。
    """
    _inject_mobile_css()

    current_hotel = ctx["current_hotel"]
    login_role = ctx["login_role"]
    now_str = datetime.datetime.now().strftime("%m/%d %H:%M")

    # 1. 頂部緊湊品牌列
    st.markdown(
        f"""
<div class="m-topbar">
<div>
<div class="m-topbar-title">🦌 Hotel Master 行動版</div>
<div class="m-topbar-sub">路徒 PLUS 行旅 · {now_str}</div>
</div>
<div class="m-pill-badge">{'🏨 ' + current_hotel if current_hotel != '採購' else '🛒 採購中心'}</div>
</div>
        """,
        unsafe_allow_html=True,
    )

    # 2. 頂部快捷操作列（僅站前館登入者顯示跨館切換按鈕，嚴守權限）
    if login_role == "站前館":
        b1, b2, b3 = st.columns([1.4, 1.1, 1.3])
        with b1:
            target_h = "主題館" if current_hotel == "站前館" else "站前館"
            icon_h = "🎨 切主題館" if target_h == "主題館" else "🏨 切站前館"
            if st.button(icon_h, key="m_switch_hotel", use_container_width=True):
                st.session_state["hotel_type"] = target_h
                st.rerun()
        with b2:
            if st.button("🔄 同步", key="m_top_sync", use_container_width=True):
                st.cache_data.clear()
                st.toast("✅ 已重新同步最新資料")
                st.rerun()
        with b3:
            if st.button("💻 電腦版", key="m_top_desktop", use_container_width=True):
                st.session_state["ui_mode"] = "desktop"
                try:
                    if "view" in st.query_params:
                        del st.query_params["view"]
                except Exception:
                    pass
                st.rerun()
    else:
        b1, b2 = st.columns(2)
        with b1:
            if st.button("🔄 重新同步", key="m_top_sync_nr", use_container_width=True):
                st.cache_data.clear()
                st.toast("✅ 已重新同步最新資料")
                st.rerun()
        with b2:
            if st.button("💻 切換電腦版", key="m_top_desktop_nr", use_container_width=True):
                st.session_state["ui_mode"] = "desktop"
                try:
                    if "view" in st.query_params:
                        del st.query_params["view"]
                except Exception:
                    pass
                st.rerun()

    # 3. 角色動態主分頁導覽（原生 Streamlit 元件，零重載不掉登入）
    if current_hotel == "採購":
        nav_opts = ["💰 預算", "✍️ 餐期", "🛒 菜價", "⋯ 更多"]
        active_tab = _render_nav_pills(nav_opts, "m_active_tab_purch", "💰 預算")
    else:
        nav_opts = ["🏠 總覽", "✍️ 填報", "🍽️ 採購", "📊 趨勢", "⋯ 更多"]
        active_tab = _render_nav_pills(nav_opts, "m_active_tab_hotel", "🏠 總覽")

    # 4. 分頁內容路由
    if active_tab == "🏠 總覽":
        _render_mobile_home(ctx)
    elif active_tab in ("✍️ 填報", "✍️ 餐期"):
        _render_mobile_input(ctx)
    elif active_tab in ("🍽️ 採購", "💰 預算"):
        _render_mobile_purchase(ctx)
    elif active_tab in ("📊 趨勢", "🛒 菜價"):
        _render_mobile_trend(ctx)
    else:
        _render_mobile_more(ctx)
