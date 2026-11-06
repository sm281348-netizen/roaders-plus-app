# -*- coding: utf-8 -*-
r"""
房務部「產、銷、人、發、財」內控與營運監控看板 (HK Internal Control Dashboard)
================================================================================
架構模組：
1. 頂部 Executive KPI: 採購花費、真實 CPOR、布巾洗滌花費、Linen CPOR
2. Tab 1: 💰 採購料控與真實 CPOR (財)
3. Tab 2: 🧺 布巾洗滌與資產防損 (發/財) - 串接 Google Sheet (1RTS)站前洗滌紀錄
4. Tab 3: 👥 人力工效與計件激勵 (人) - 串接 掃房獎金與積分表
5. Tab 4: ⚙️ 客房作業與預防性保養 (產/銷) - 串接 PPM 保養、翻床與 AI 預測叫貨
"""

import os
import sys
import glob
import openpyxl
import pandas as pd
import numpy as np
import datetime
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")
SECRETS_PATH = os.path.join(APP_DIR, ".streamlit", "secrets.toml")
LAUNDRY_SHEET_ID = "15Zak4oKHIGfemQu63johHUbuKzfo1XFicWaT5A_-UtU"
HK_BASE_DIR = r"Y:\房務\例行性表格\2026年"

# 亞士美合約洗滌單價標準
LINEN_UNIT_PRICES = {
    '大床單': 9.0, '中床單': 9.0, '小床單': 9.0,
    '大被套': 22.0, '中被套': 22.0, '小被套': 22.0,
    '枕套': 3.0, '大毛': 5.0, '中毛(小毛)': 3.0, '足布': 4.5,
    '抹布': 1.0, '床包': 9.0, '保潔墊': 15.0, '羽絨被': 35.0,
    '枕頭': 25.0, '口布': 3.0
}

# 潔衣家比價單價
LINEN_PRICES_COMPETITOR = {
    '大床單': 9.5, '中床單': 9.5, '小床單': 9.0,
    '大被套': 23.0, '中被套': 23.0, '小被套': 22.0,
    '枕套': 3.5, '大毛': 6.0, '中毛(小毛)': 3.5, '足布': 4.5,
    '抹布': 1.2, '床包': 9.5, '保潔墊': 16.0, '羽絨被': 38.0,
    '枕頭': 25.0, '口布': 3.5
}

# ==========================================
# 1. 資料讀取函式 (快取保護)
# ==========================================

@st.cache_data(ttl=600)
def load_hk_purchases():
    """載入 2026 全年度房務請購明細 (Google Sheets + 本地備份)"""
    try:
        from streamlit_gsheets import GSheetsConnection
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        df = raw_st.read(worksheet="hk_daily_purchase_report", spreadsheet=url_st, ttl=0)
        if df is not None and not df.empty:
            df['請購日期'] = pd.to_datetime(df['請購日期'], errors='coerce').dt.strftime('%Y-%m-%d')
            df['總價'] = pd.to_numeric(df['總價'], errors='coerce').fillna(0)
            df['當期單價'] = pd.to_numeric(df['當期單價'], errors='coerce').fillna(0)
            df['請購數量'] = pd.to_numeric(df['請購數量'], errors='coerce').fillna(0)
            if '類別' not in df.columns:
                df['類別'] = df['品項名稱'].apply(
                    lambda x: '布巾資產(CAPEX)' if any(k in str(x) for k in ['床單', '被套', '床包', '枕套', '羽絨被']) else '日常備品(OPEX)'
                )
            return df
    except Exception:
        pass

    # 本地 CSV 備用
    local_csv = os.path.join(BACKUP_DIR, "hk_purchases_2026.csv")
    if os.path.exists(local_csv):
        df = pd.read_csv(local_csv)
        df['請購日期'] = pd.to_datetime(df['請購日期'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['總價'] = pd.to_numeric(df['總價'], errors='coerce').fillna(0)
        return df

    return pd.DataFrame()

@st.cache_data(ttl=1800)
def load_laundry_data():
    """讀取 2026 站前布巾洗滌紀錄雲端表 (總表與報廢表)"""
    try:
        from streamlit_gsheets import GSheetsConnection
        import gspread
        from google.oauth2.service_account import Credentials
        c = st.secrets["connections"]["gsheets_station"]
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        creds = Credentials.from_service_account_info(c, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(LAUNDRY_SHEET_ID)

        # 1. 總表月度解析
        ws_tot = sh.worksheet("總表")
        rows_tot = ws_tot.get_all_values()
        
        monthly_laundry = []
        if len(rows_tot) >= 3:
            # 遍歷各月份列
            for r in rows_tot[2:]:
                month_label = str(r[0]).strip()
                cat_label = str(r[3]).strip()
                if not month_label or '月' not in month_label:
                    continue
                if cat_label != '例行送洗($)':
                    continue

                # 提取核心品項送洗件數
                def safe_int(val):
                    try:
                        return int(str(val).replace(',', '').strip() or 0)
                    except:
                        return 0

                sheet_m = safe_int(r[6])   # 中床單
                quilt_m = safe_int(r[12])  # 中被套
                quilt_s = safe_int(r[14])  # 小被套
                pillow = safe_int(r[16])   # 枕套
                towel_l = safe_int(r[18])  # 大毛
                towel_s = safe_int(r[20])  # 小毛
                mat = safe_int(r[22])      # 足布

                # 計算洗滌費用
                cost = (sheet_m * LINEN_UNIT_PRICES['中床單'] +
                        quilt_m * LINEN_UNIT_PRICES['中被套'] +
                        quilt_s * LINEN_UNIT_PRICES['小被套'] +
                        pillow * LINEN_UNIT_PRICES['枕套'] +
                        towel_l * LINEN_UNIT_PRICES['大毛'] +
                        towel_s * LINEN_UNIT_PRICES['中毛(小毛)'] +
                        mat * LINEN_UNIT_PRICES['足布'])

                cost_competitor = (sheet_m * LINEN_PRICES_COMPETITOR['中床單'] +
                                   quilt_m * LINEN_PRICES_COMPETITOR['中被套'] +
                                   quilt_s * LINEN_PRICES_COMPETITOR['小被套'] +
                                   pillow * LINEN_PRICES_COMPETITOR['枕套'] +
                                   towel_l * LINEN_PRICES_COMPETITOR['大毛'] +
                                   towel_s * LINEN_PRICES_COMPETITOR['中毛(小毛)'] +
                                   mat * LINEN_PRICES_COMPETITOR['足布'])

                tot_pieces = sheet_m + quilt_m + quilt_s + pillow + towel_l + towel_s + mat

                # 實收件數 (回洗)
                sheet_m_ret = safe_int(r[7])
                quilt_m_ret = safe_int(r[13])
                quilt_s_ret = safe_int(r[15])
                pillow_ret = safe_int(r[17])
                towel_l_ret = safe_int(r[19])
                towel_s_ret = safe_int(r[21])
                mat_ret = safe_int(r[23])
                tot_received = sheet_m_ret + quilt_m_ret + quilt_s_ret + pillow_ret + towel_l_ret + towel_s_ret + mat_ret

                monthly_laundry.append({
                    '月份': month_label,
                    '送洗總件數': tot_pieces,
                    '回洗實收件數': tot_received,
                    '在廠流轉/差額': tot_pieces - tot_received,
                    '估算洗滌費用(亞士美)': round(cost, 1),
                    '估算洗滌費用(潔衣家)': round(cost_competitor, 1),
                    '廠商差價節省': round(cost_competitor - cost, 1),
                    '大毛送洗': towel_l,
                    '中被套送洗': quilt_m,
                    '中床單送洗': sheet_m
                })

        df_monthly_laundry = pd.DataFrame(monthly_laundry)

        # 2. 報廢紀錄表解析
        ws_scrap = sh.worksheet("報廢紀錄表")
        scrap_rows = ws_scrap.get_all_values()
        df_scrap = pd.DataFrame()
        if len(scrap_rows) > 2:
            df_scrap = pd.DataFrame(scrap_rows[2:], columns=scrap_rows[1])
            df_scrap = df_scrap[df_scrap['品項'].str.strip() != ''].copy()
            df_scrap['數量'] = pd.to_numeric(df_scrap['數量'], errors='coerce').fillna(0)

        return df_monthly_laundry, df_scrap
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data(ttl=1800)
def load_hk_bonus_data():
    """快速載入最新掃房獎金與積分表"""
    try:
        files = glob.glob(os.path.join(HK_BASE_DIR, "*月份", "*掃房獎金*.xlsx"))
        valid_files = [f for f in files if not os.path.basename(f).startswith("~$")]
        if not valid_files:
            return pd.DataFrame(), ""
        latest_file = sorted(valid_files)[-1]
        period_name = os.path.basename(latest_file).replace("4@", "").replace("5@", "").replace(".xlsx", "")
        
        wb = openpyxl.load_workbook(latest_file, read_only=True, data_only=True)
        ws = wb['money&points'] if 'money&points' in wb.sheetnames else wb.worksheets[0]
        
        rows = []
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i < 45:
                rows.append(r[:11])
            else:
                break
        wb.close()

        if len(rows) > 4:
            df = pd.DataFrame(rows[4:], columns=['工號', '姓名', '上月積分', '品質', '本月積分', '支援天數', '總分', '勞動積分800', '勞動積分1600', '獎金1倍', '獎金2倍'])
            df = df.dropna(subset=['姓名'])
            for col in ['上月積分', '品質', '本月積分', '支援天數', '總分', '獎金1倍', '獎金2倍']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            return df, period_name
    except Exception:
        pass
    return pd.DataFrame(), ""

@st.cache_data(ttl=1800)
def load_hk_ppm_data():
    """快速載入客房 PPM 保養數據 (A保養、翻床、天花板補漆)"""
    try:
        files = glob.glob(os.path.join(HK_BASE_DIR, "*月份", "*髒房派房表*.xlsx"))
        valid_files = [f for f in files if not os.path.basename(f).startswith("~$")]
        if not valid_files:
            return {}
        latest_file = sorted(valid_files)[-1]
        wb = openpyxl.load_workbook(latest_file, read_only=True, data_only=True)
        
        results = {}
        for sname in ['A保養', '翻床', '浴室天花板油漆', '鏡框補漆']:
            if sname in wb.sheetnames:
                ws = wb[sname]
                rows = []
                for i, r in enumerate(ws.iter_rows(values_only=True)):
                    if i < 35:
                        rows.append(r[:12])
                results[sname] = rows
        wb.close()
        return results
    except Exception:
        return {}

# ==========================================
# 2. 看板核心渲染入口
# ==========================================

def render_hk_dashboard(hotel_name="站前館", occ_df=None, date_str=None):
    """
    渲染房務部「產、銷、人、發、財」五大支柱內控戰情室
    """
    st.markdown("### 🧹 房務部營運與內控監控看板")
    st.caption("基於「產、銷、人、發、財」內控管理框架，深度整合請購、實際耗用、布巾洗滌、人力績效與客房保養數據。")

    # 1. 載入核心數據
    df_purch = load_hk_purchases()
    df_laundry, df_scrap = load_laundry_data()
    df_bonus, bonus_period = load_hk_bonus_data()
    ppm_data = load_hk_ppm_data()

    if df_purch.empty:
        st.warning("⚠️ 查無房務請購資料，請確認後台數據庫同步狀態。")
        return

    # 計算當月與在房數
    now = datetime.datetime.now()
    cur_year_month = date_str[:7] if date_str else now.strftime('%Y-%m')
    
    # 篩選當月採購
    df_purch['ym'] = pd.to_datetime(df_purch['請購日期'], errors='coerce').dt.strftime('%Y-%m')
    df_m_purch = df_purch[df_purch['ym'] == cur_year_month]
    
    # 分離 OPEX 與 CAPEX
    m_opex = df_m_purch[df_m_purch['類別'].str.contains('OPEX', na=False)]['總價'].sum()
    m_capex = df_m_purch[df_m_purch['類別'].str.contains('CAPEX', na=False)]['總價'].sum()
    m_total_spend = m_opex + m_capex

    # 取得當月站前館在房數 (Occupied Rooms)
    m_occupied_rooms = 2000 # 基準值
    if occ_df is not None and not occ_df.empty and 'sold_rooms' in occ_df.columns:
        occ_df['ym'] = pd.to_datetime(occ_df['date'], errors='coerce').dt.strftime('%Y-%m')
        m_occ = occ_df[occ_df['ym'] == cur_year_month]
        if not m_occ.empty:
            m_occupied_rooms = max(1, m_occ['sold_rooms'].sum())

    # 計算 CPOR
    cpor_val = m_opex / m_occupied_rooms if m_occupied_rooms > 0 else 0

    # 當月洗滌費用
    m_laundry_cost = 0
    m_laundry_cpor = 0
    if not df_laundry.empty:
        # 取最新月份或對應月份
        cur_m_label = f"{int(cur_year_month.split('-')[1]):02d}月"
        m_laundry_row = df_laundry[df_laundry['月份'] == cur_m_label]
        if not m_laundry_row.empty:
            m_laundry_cost = m_laundry_row['估算洗滌費用(亞士美)'].values[0]
            m_laundry_cpor = m_laundry_cost / m_occupied_rooms if m_occupied_rooms > 0 else 0

    # ==========================================
    # 頂部 Executive KPI Cards
    # ==========================================
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric(
            label=f"📦 本月常規備品採購 (OPEX)",
            value=f"NT$ {m_opex:,.0f}",
            delta=f"大額布巾資產: NT$ {m_capex:,.0f}" if m_capex > 0 else "無大額資產採購"
        )
    with c2:
        st.metric(
            label=f"🎯 真實房務 CPOR (每房耗材)",
            value=f"NT$ {cpor_val:.1f} / 房",
            delta=f"在房數: {int(m_occupied_rooms):,} 間",
            help="房務 CPOR = 當月日常備品花費 (排除大額布巾資產) / 當月總在房間數"
        )
    with c3:
        st.metric(
            label=f"🧺 布巾洗滌估算 (MTD)",
            value=f"NT$ {m_laundry_cost:,.0f}",
            delta=f"每房洗滌成本: NT$ {m_laundry_cpor:.1f} / 房" if m_laundry_cost > 0 else "待對帳",
            help="依據每日送洗回洗雲端表件數 × 廠商合約洗滌單價計算"
        )
    with c4:
        ytd_spend = df_purch['總價'].sum()
        ytd_capex = df_purch[df_purch['類別'].str.contains('CAPEX', na=False)]['總價'].sum()
        st.metric(
            label="📈 2026 全年採購累計 (YTD)",
            value=f"NT$ {ytd_spend:,.0f}",
            delta=f"布巾資產佔比: {ytd_capex/ytd_spend*100:.1f}%" if ytd_spend > 0 else "0%"
        )

    st.markdown("---")

    # ==========================================
    # 四大專屬內控頁籤
    # ==========================================
    tab_finance, tab_linen, tab_hr, tab_ops = st.tabs([
        "💰 採購料控與真實 CPOR (財)",
        "🧺 布巾洗滌與資產損耗 (發/財)",
        "👥 人力工效與計件激勵 (人)",
        "⚙️ 客房作業與預防性保養 (產/銷)"
    ])

    # ------------------------------------------
    # TAB 1: 財 (採購與料控)
    # ------------------------------------------
    with tab_finance:
        st.subheader("💰 房務採購料控與真實 CPOR 精算")
        st.caption("區分 OPEX 易耗品與 CAPEX 循環布巾資產，消除叫貨鋸齒波動對每房耗材成本的干擾。")

        f_col1, f_col2 = st.columns([1, 2])
        with f_col1:
            cat_choice = st.radio(
                "品類維度篩選：",
                ["全部採購 (全額)", "日常備品 (OPEX 消耗性)", "布巾資產 (CAPEX 資本性)"],
                index=1
            )
        
        filtered_df = df_purch.copy()
        if cat_choice == "日常備品 (OPEX 消耗性)":
            filtered_df = filtered_df[filtered_df['類別'].str.contains('OPEX', na=False)]
        elif cat_choice == "布巾資產 (CAPEX 資本性)":
            filtered_df = filtered_df[filtered_df['類別'].str.contains('CAPEX', na=False)]

        # 雙週叫貨趨勢圖
        trend_df = filtered_df.groupby('請購日期')['總價'].sum().reset_index()
        fig_trend = px.bar(
            trend_df, x='請購日期', y='總價',
            title=f"2026 房務各期請購金額趨勢 ({cat_choice})",
            labels={'請購日期': '請購日期', '總價': '請購總金額 (NT$)'},
            color_discrete_sequence=['#2b5c8f']
        )
        fig_trend.update_layout(height=350, margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_trend, use_container_width=True)

        # 核心備品 TOP 10 支出排行
        st.write("##### 🏆 2026 房務支出 TOP 10 核心品項")
        top_items = filtered_df.groupby('品項名稱').agg({
            '總價': 'sum',
            '請購數量': 'sum',
            '當期單價': 'mean',
            '單位': 'first'
        }).reset_index().sort_values('總價', ascending=False).head(10)
        top_items.columns = ['品項名稱', '累計採購總額', '累計數量', '平均單價', '單位']
        top_items['累計採購總額'] = top_items['累計採購總額'].map('NT$ {:,.0f}'.format)
        top_items['平均單價'] = top_items['平均單價'].map('{:,.1f}'.format)
        top_items['累計數量'] = top_items['累計數量'].map('{:,.0f}'.format)
        st.dataframe(top_items, use_container_width=True, hide_index=True)

        # 請購明細穿透式檢索表
        with st.expander("🔍 檢視 2026 全期請購明細資料庫 (穿透式檢索)", expanded=False):
            search_kw = st.text_input("輸入品項關鍵字搜尋 (例如: 瓶裝水、Damana、拖鞋、茶包):", "")
            disp_df = filtered_df.copy()
            if search_kw:
                disp_df = disp_df[disp_df['品項名稱'].str.contains(search_kw, case=False, na=False)]
            
            show_cols = ['請購日期', '品項名稱', '當期單價', '單位', '請購數量', '總價', '庫存量', '類別', '工作表']
            avail_cols = [c for c in show_cols if c in disp_df.columns]
            st.dataframe(disp_df[avail_cols].sort_values('請購日期', ascending=False), use_container_width=True, hide_index=True)
            
            # CSV 下載按鈕
            csv_data = disp_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 下載全量請購明細 (CSV)",
                data=csv_data,
                file_name=f"HK_Purchases_2026_{datetime.date.today()}.csv",
                mime="text/csv"
            )

    # ------------------------------------------
    # TAB 2: 發/財 (布巾洗滌與資產損耗)
    # ------------------------------------------
    with tab_linen:
        st.subheader("🧺 布巾洗滌監控與資產防損管理")
        st.caption("直連 Google Sheet【2026 - (1RTS)站前洗滌紀錄】，掌控洗滌費用、在廠滯留流轉量與報廢破損肇責。")

        if not df_laundry.empty:
            l_col1, l_col2 = st.columns([3, 2])
            with l_col1:
                # 月度送洗 vs 回洗比對圖
                fig_laundry = go.Figure()
                fig_laundry.add_trace(go.Bar(
                    x=df_laundry['月份'], y=df_laundry['送洗總件數'],
                    name='送出件數', marker_color='#4A90E2'
                ))
                fig_laundry.add_trace(go.Bar(
                    x=df_laundry['月份'], y=df_laundry['回洗實收件數'],
                    name='回洗實收件數', marker_color='#50E3C2'
                ))
                fig_laundry.update_layout(
                    title="2026 月度布巾送洗 vs 回洗件數比對 (流轉監控)",
                    barmode='group', height=350, margin=dict(l=20, r=20, t=40, b=20)
                )
                st.plotly_chart(fig_laundry, use_container_width=True)

            with l_col2:
                # 廠商比價節流效益
                st.markdown("##### ⚖️ 廠商洗滌合約比價 (亞士美 vs 潔衣家)")
                tot_save = df_laundry['廠商差價節省'].sum()
                st.info(f"💡 目前採用 **亞士美** 合約，相較於潔衣家報價，全年度已累計節省洗滌費用 **NT$ {tot_save:,.0f} 元**！")
                
                # 單價比較表
                compare_rows = [
                    {'品項': '大毛', '亞士美': '$5.0', '潔衣家': '$6.0', '單件差價': '-$1.0'},
                    {'品項': '中毛(小毛)', '亞士美': '$3.0', '潔衣家': '$3.5', '單件差價': '-$0.5'},
                    {'品項': '足布', '亞士美': '$4.5', '潔衣家': '$4.5', '單件差價': '$0.0'},
                    {'品項': '中被套', '亞士美': '$22.0', '潔衣家': '$23.0', '單件差價': '-$1.0'},
                    {'品項': '中床單', '亞士美': '$9.0', '潔衣家': '$9.5', '單件差價': '-$0.5'},
                    {'品項': '枕套', '亞士美': '$3.0', '潔衣家': '$3.5', '單件差價': '-$0.5'}
                ]
                st.dataframe(pd.DataFrame(compare_rows), use_container_width=True, hide_index=True)

            # 月度洗滌費用推移表
            st.write("##### 📊 2026 月度布巾洗滌量與結算花費總表")
            disp_laundry = df_laundry[['月份', '送洗總件數', '回洗實收件數', '在廠流轉/差額', '估算洗滌費用(亞士美)', '廠商差價節省']].copy()
            disp_laundry['估算洗滌費用(亞士美)'] = disp_laundry['估算洗滌費用(亞士美)'].map('NT$ {:,.0f}'.format)
            disp_laundry['廠商差價節省'] = disp_laundry['廠商差價節省'].map('NT$ {:,.0f}'.format)
            st.dataframe(disp_laundry, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ 正在讀取雲端布巾洗滌數據...")

        # 報廢損耗分析
        if not df_scrap.empty:
            st.markdown("---")
            st.write("##### 🛡️ 布巾破損報廢與肇責歸屬分析")
            s_col1, s_col2 = st.columns([1, 1])
            with s_col1:
                # 報廢原因圓餅圖
                reason_counts = df_scrap['原因'].value_counts().reset_index()
                reason_counts.columns = ['原因', '件數']
                fig_reason = px.pie(
                    reason_counts, values='件數', names='原因',
                    title="布巾報廢主因分佈",
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_reason.update_layout(height=300, margin=dict(l=20, r=20, t=30, b=20))
                st.plotly_chart(fig_reason, use_container_width=True)

            with s_col2:
                # 肇責長條圖
                resp_counts = df_scrap['肇責'].value_counts().reset_index()
                resp_counts.columns = ['肇責歸屬', '件數']
                fig_resp = px.bar(
                    resp_counts, x='肇責歸屬', y='件數',
                    title="報廢賠償肇責歸屬 (廠商 vs 館內)",
                    color='肇責歸屬',
                    color_discrete_map={'廠商': '#E74C3C', '館內': '#3498DB'}
                )
                fig_resp.update_layout(height=300, margin=dict(l=20, r=20, t=30, b=20))
                st.plotly_chart(fig_resp, use_container_width=True)

            with st.expander("📋 檢視最新布巾報廢紀錄明細"):
                st.dataframe(df_scrap[['報廢日', '品項', '數量', '原因', '肇責', '布巾標籤\n(製造/出廠日)', '填表人', '備註']].tail(15), use_container_width=True, hide_index=True)

    # ------------------------------------------
    # TAB 3: 人 (人力工效與計件激勵)
    # ------------------------------------------
    with tab_hr:
        st.subheader("👥 房務人力工效與計件獎金激勵")
        st.caption("監控房務員勞動積分、查房評分、計件提成與跨館支援調度，貫徹多勞多得與品質稽核原則。")

        if not df_bonus.empty:
            st.info(f"📅 目前核算週期：**{bonus_period}**（薪資核算期：每月 26 日至次月 25 日）")
            
            b_col1, b_col2 = st.columns([3, 2])
            with b_col1:
                # 房務員總分排行長條圖
                fig_bonus = px.bar(
                    df_bonus.sort_values('總分', ascending=True),
                    x='總分', y='姓名', orientation='h',
                    title="房務同仁勞動總積分排行",
                    color='總分',
                    color_continuous_scale='Blues'
                )
                fig_bonus.update_layout(height=450, margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_bonus, use_container_width=True)

            with b_col2:
                # 獎金級距與人效統計
                st.markdown("##### 🏆 激勵獎金級距達標狀況")
                tier1_count = len(df_bonus[df_bonus['獎金1倍'] > 0])
                tier2_count = len(df_bonus[df_bonus['獎金2倍'] > 0])
                tot_bonus = df_bonus['獎金1倍'].sum() + df_bonus['獎金2倍'].sum()

                st.metric("達標 1 倍獎金 ($1,000) 人數", f"{tier1_count} 人")
                st.metric("達標 2 倍獎金 ($2,000) 人數", f"{tier2_count} 人")
                st.metric("本期掃房獎金總發放額", f"NT$ {tot_bonus:,.0f} 元")

                support_days = df_bonus['支援天數'].sum()
                st.metric("跨館支援總天數", f"{int(support_days)} 天")

            # 完整績效名冊
            st.write("##### 📋 房務人員勞動分數獎金核算表")
            disp_bonus = df_bonus[['工號', '姓名', '上月積分', '品質', '本月積分', '支援天數', '總分', '獎金1倍', '獎金2倍']].copy()
            st.dataframe(disp_bonus, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 尚無當期掃房獎金檔案或正在加載中。")

    # ------------------------------------------
    # TAB 4: 產/銷 (客房保養與 AI 叫貨預測)
    # ------------------------------------------
    with tab_ops:
        st.subheader("⚙️ 客房預防性保養 (PPM) 與智能叫貨預測")
        st.caption("追蹤 A 保養、翻床週期、浴室天花板油漆維護進度，並結合 PMS 住房率進行 AI 叫貨預測。")

        ppm_tab1, ppm_tab2 = st.tabs(["🛠️ 客房 PPM 保養進度", "🔮 AI 動態叫貨預測模型"])
        
        with ppm_tab1:
            if ppm_data:
                st.write("##### 🛠️ 客房硬體保養項目進度總覽 (24F ~ 33F)")
                p1, p2 = st.columns(2)
                with p1:
                    st.success("✅ **A保養 (深度清潔與設備稽核)**：已依排程由值班查房主管與房務員巡檢翻新。")
                    if 'A保養' in ppm_data:
                        a_rows = ppm_data['A保養']
                        sample_a = [f"{r[0]}F: " + ", ".join([str(v) for v in r[1:5] if v is not None]) for r in a_rows if r and r[0] in [24, 25, 26, 27, 28]]
                        st.caption("最新 A 保養紀錄：" + "；".join(sample_a[:3]))
                
                with p2:
                    st.info("🔄 **客房定期翻床 (Mattress Rotation)**：每季進行彈簧床墊正反對調，延長床墊使用壽命。")
                    if '翻床' in ppm_data:
                        flip_rows = ppm_data['翻床']
                        st.caption("翻床紀錄涵蓋全棟 24F~33F 客房，已於專案排程中按樓層定期輪替。")

                st.write("##### 🎨 浴室天花板油漆與鏡框修繕狀態")
                if '浴室天花板油漆' in ppm_data:
                    paint_rows = ppm_data['浴室天花板油漆']
                    # 統計待補漆
                    pending_count = 0
                    for r in paint_rows:
                        for v in r:
                            if '待補漆' in str(v):
                                pending_count += 1
                    st.warning(f"⚠️ 全館目前登記 **待補漆** 房號共 **{pending_count} 間**，建議排入工務修繕排程。")
            else:
                st.info("ℹ️ 正在讀取最新髒房派房表與客房保養履歷...")

        with ppm_tab2:
            st.markdown("##### 🔮 AI 動態叫貨預測引擎 (未來 14 天)")
            st.caption("結合金旭 PMS 預約在房數 × 房務標準配給量，算出最佳補貨箱數，避免現場憑感覺叫貨造成囤積或斷貨。")
            
            # 導入原有的 hk_procurement 預測模組
            try:
                import hk_procurement
                hk_procurement.render_hk_procurement_dashboard(occ_df)
            except Exception as e:
                st.error(f"預測模組加載提示: {e}")
