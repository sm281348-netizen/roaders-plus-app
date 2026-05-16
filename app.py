import streamlit as st
import datetime
import pandas as pd
import time
import random
import os
import re
import altair as alt
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from streamlit_gsheets import GSheetsConnection
import holidays

# APP_VERSION: 1.3.0 - RECONSTRUCTED
TARGET_HOLIDAY_COUNTRIES = {
    'KR': '[韓]',
    'SG': '[星]',
    'HK': '[港]',
    'JP': '[日]',
    'US': '[美]',
    'TW': '[台]'
}
OTHER_HOLIDAY_COUNTRIES = ['PH', 'MY', 'TH', 'VN']

EVENT_TYPE_LABELS = {
    '演唱會': '[演]',
    '展覽': '[展]',
    '賽事': '[賽]',
    '其他': '[活]'
}

@st.cache_data(ttl=600)
def fetch_taipei_events():
    """讀取試算表中的台北重大活動分頁"""
    try:
        df = conn.read(worksheet="taipei_events", ttl="10m")
        if df is not None and not df.empty:
            df = standardize_df_dates(df)
            return df
    except Exception:
        return pd.DataFrame(columns=['date', 'event_name', 'event_type', 'venue'])
    return pd.DataFrame(columns=['date', 'event_name', 'event_type', 'venue'])

@st.cache_data(ttl=86400 * 30)
def translate_to_zh(text):
    if not text: return text
    try:
        from deep_translator import GoogleTranslator
        translator = GoogleTranslator(source='auto', target='zh-TW')
        return translator.translate(text)
    except:
        return text

@st.cache_data(ttl=86400)
def fetch_holidays_for_month(year, month):
    h_objs = {}
    for code in list(TARGET_HOLIDAY_COUNTRIES.keys()) + OTHER_HOLIDAY_COUNTRIES:
        try:
            h_objs[code] = holidays.country_holidays(code, years=[year])
        except Exception:
            continue
            
    import calendar
    _, last_day = calendar.monthrange(year, month)
    
    result = {}
    for day in range(1, last_day + 1):
        dt_obj = datetime.date(year, month, day)
        dt_str = dt_obj.strftime('%Y-%m-%d')
        day_flags = set()
        day_details = []
        has_other = False
        
        for code, h_obj in h_objs.items():
            if dt_obj in h_obj:
                raw_name = h_obj.get(dt_obj)
                h_name = translate_to_zh(raw_name)
                if code in TARGET_HOLIDAY_COUNTRIES:
                    day_flags.add(TARGET_HOLIDAY_COUNTRIES[code])
                    day_details.append(f"{TARGET_HOLIDAY_COUNTRIES[code]} {code}: {h_name}")
                else:
                    has_other = True
                    day_details.append(f"🌍 {code}: {h_name}")
        
        flags_str = "".join(sorted(list(day_flags)))
        if has_other: flags_str += "🌍"
        if flags_str:
            result[dt_str] = {'flags': flags_str, 'details': day_details}
    return result

@st.cache_data(ttl=86400)
def fetch_upcoming_holidays(start_date, days=30):
    years = {start_date.year, (start_date + datetime.timedelta(days=days)).year}
    h_objs = {}
    for code in list(TARGET_HOLIDAY_COUNTRIES.keys()) + OTHER_HOLIDAY_COUNTRIES:
        try:
            h_objs[code] = holidays.country_holidays(code, years=list(years))
        except: pass
    result = []
    for i in range(days + 1):
        dt_obj = start_date + datetime.timedelta(days=i)
        day_details = []
        day_flags = set()
        has_other = False
        for code, h_obj in h_objs.items():
            if dt_obj in h_obj:
                raw_name = h_obj.get(dt_obj)
                h_name = translate_to_zh(raw_name)
                if code in TARGET_HOLIDAY_COUNTRIES:
                    day_flags.add(TARGET_HOLIDAY_COUNTRIES[code])
                    day_details.append(f"{TARGET_HOLIDAY_COUNTRIES[code]} {code}: {h_name}")
                else:
                    has_other = True
                    day_details.append(f"🌍 {code}: {h_name}")
        flags_str = "".join(sorted(list(day_flags)))
        if has_other: flags_str += "🌍"
        if flags_str:
            result.append({
                'date': dt_obj.strftime('%Y-%m-%d'),
                'flags': flags_str,
                'details': ", ".join(day_details)
            })
    return result

st.set_page_config(page_title="路徒Plus行旅站前館營運日誌", layout="wide")

if "authenticated" not in st.session_state:
    st.markdown("<h2 style='text-align: center;'>🔒 歡迎登入 路徒Plus行旅 營運日誌</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>為了保護營業機密，請輸入管理員通行碼進入系統。</p>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd = st.text_input("管理員通行碼", type="password")
        if pwd:
            correct_password = st.secrets.get("admin_password", "roaders123")
            if pwd == correct_password:
                st.session_state["authenticated"] = True
                st.rerun()
            else: st.error("❌ 密碼錯誤，請重新輸入。")
    st.stop()

conn = st.connection("gsheets", type=GSheetsConnection)

def standardize_df_dates(df):
    if df is None or df.empty or 'date' not in df.columns: return df
    def fix_d(val):
        if pd.isna(val) or str(val).strip() == '' or str(val).strip() == 'NaT': return ""
        v_str = str(val).split(' ')[0].strip()
        m_tw = re.match(r'^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})$', v_str)
        if m_tw:
            y, m, d = int(m_tw.group(1)), int(m_tw.group(2)), int(m_tw.group(3))
            if y < 1000: y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
        m_md = re.match(r'^(\d{1,2})[/-](\d{1,2})$', v_str)
        if m_md:
            import datetime
            curr_y = datetime.datetime.now().year
            m, d = int(m_md.group(1)), int(m_md.group(2))
            return f"{curr_y:04d}-{m:02d}-{d:02d}"
        try:
            p = pd.to_datetime(v_str)
            if pd.notna(p): return p.strftime('%Y-%m-%d')
        except: pass
        return v_str
    df['date'] = df['date'].apply(fix_d)
    return df

def get_daily_data(d_str):
    try:
        df = conn.read(worksheet="daily_data", ttl="1m")
        if df is not None and not df.empty:
            df = standardize_df_dates(df)
            df = df.drop_duplicates(subset='date', keep='last')
            res = df[df['date'] == d_str]
            if not res.empty:
                data_dict = res.iloc[0].to_dict()
                numeric_cols = [
                    'occ_rate', 'adr', 'revenue', 'total_rooms', 'counter_expense', 
                    'cleaned_rooms', 'hk_checkout_extend', 'hk_avg_clean', 'hk_expense',
                    'rest_breakfast', 'rest_month_guests', 'rest_day_guests', 'rest_avg_guests',
                    'rest_month_rev', 'rest_avg_spent', 'rest_peak_expense', 'rest_hh_guests',
                    'maint_repair_rooms', 'maint_expense',
                    'bf_theme_est', 'bf_theme_act', 'bf_zq_est', 'bf_zq_act', 'bf_total_est', 'bf_total_act',
                    'af_theme_est', 'af_theme_act', 'af_zq_est', 'af_zq_act', 'af_total_est', 'af_total_act'
                ]
                for col in numeric_cols:
                    if col in data_dict and (pd.isna(data_dict[col]) or data_dict[col] is None):
                        data_dict[col] = 0
                return data_dict
    except Exception: pass
    return {}

def save_daily_data(d_str, data_dict):
    try:
        df = conn.read(worksheet="daily_data", ttl="0")
        if df is None: df = pd.DataFrame()
        df = standardize_df_dates(df)
        data_dict['date'] = d_str
        new_row = pd.DataFrame([data_dict])
        if 'date' in df.columns and d_str in df['date'].values:
            df = df[df['date'] != d_str]
        df = pd.concat([df, new_row], ignore_index=True)
        if 'date' in df.columns:
            df = df.sort_values('date').reset_index(drop=True)
        conn.update(worksheet="daily_data", data=df.fillna(""))
        st.cache_data.clear()
    except Exception as e: st.error(f"儲存失敗: {e}")

def get_monthly_target(month_str):
    try:
        df = conn.read(worksheet="targets", ttl="1m")
        if df is not None and not df.empty:
            res = df[df['month'] == month_str]
            if not res.empty: return int(res.iloc[0]['target_revenue'])
    except Exception: pass
    return 0

def save_monthly_target(month_str, target):
    try:
        df = conn.read(worksheet="targets", ttl="0")
        if df is None or df.empty: df = pd.DataFrame(columns=["month", "target_revenue"])
        if month_str in df['month'].values: df.loc[df['month'] == month_str, 'target_revenue'] = target
        else:
            new_row = pd.DataFrame([{"month": month_str, "target_revenue": target}])
            df = pd.concat([df, new_row], ignore_index=True)
        conn.update(worksheet="targets", data=df.fillna(""))
        st.cache_data.clear()
        return True
    except: return False

def get_daily_log(d_str):
    try:
        df = conn.read(worksheet="daily_logs", ttl="1m")
        if df is not None and not df.empty:
            res = df[df['date'] == d_str]
            if not res.empty: return str(res.iloc[0]['log']).strip()
    except: pass
    return ""

def save_daily_log(d_str, log_text):
    try:
        df = conn.read(worksheet="daily_logs", ttl="0")
        if df is None or df.empty: df = pd.DataFrame(columns=["date", "log"])
        df = standardize_df_dates(df)
        new_row = pd.DataFrame([{'date': d_str, 'log': log_text}])
        if d_str in df['date'].values: df = df[df['date'] != d_str]
        df = pd.concat([df, new_row], ignore_index=True)
        if 'date' in df.columns: df = df.sort_values('date').reset_index(drop=True)
        conn.update(worksheet="daily_logs", data=df.fillna(""))
        st.cache_data.clear()
        return True
    except Exception: return False

def get_month_delta(d, delta):
    year = d.year
    month = d.month + delta
    while month > 12: month -= 12; year += 1
    while month < 1: month += 12; year -= 1
    return datetime.date(year, month, 1)

def fetch_month_summary(year, month):
    import calendar
    m_start = f"{year}-{month:02d}-01"
    _, last_day = calendar.monthrange(year, month)
    m_end = f"{year}-{month:02d}-{last_day:02d}"
    try:
        df_all = conn.read(worksheet="daily_data", ttl="10m")
        if df_all is not None and not df_all.empty:
            df_all = standardize_df_dates(df_all)
            df_all = df_all.drop_duplicates(subset='date', keep='last')
            df = df_all[(df_all['date'] >= m_start) & (df_all['date'] <= m_end)].copy()
        else: df = pd.DataFrame()
    except Exception: df = pd.DataFrame()
    res = {'rev': 0.0, 'rooms': 0.0, 'sellable': 0.0, 'occ90_days': 0, 'avg_occ': 0.0, 'avg_adr': 0.0, 'revpar': 0.0, 'df': df, 'month_label': f"{year}-{month:02d}"}
    if not df.empty:
        num_cols = ['revenue', 'total_rooms', 'occ_rate', 'adr']
        for c in num_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(',', '').str.replace('%', '')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
        for _, r in df.iterrows():
            rev = float(r['revenue']); rm = float(r['total_rooms']); occ = float(r['occ_rate']); adr = float(r['adr'])
            if rev == 0 and adr > 0 and rm > 0: rev = adr * rm
            if rm == 0 and rev > 0 and adr > 0: rm = rev / adr
            if rm > 0 or rev > 0:
                res['rev'] += rev; res['rooms'] += rm
                if occ > 0: res['sellable'] += (rm / (occ / 100.0))
                if occ >= 90.0: res['occ90_days'] += 1
        res['avg_occ'] = (res['rooms'] / res['sellable'] * 100.0) if res['sellable'] > 0 else 0.0
        res['avg_adr'] = (res['rev'] / res['rooms']) if res['rooms'] > 0 else 0.0
        res['revpar'] = (res['avg_occ'] / 100.0) * res['avg_adr']
    return res

st.sidebar.header("📅 日期選擇")
if 'sidebar_date' not in st.session_state: st.session_state['sidebar_date'] = datetime.date.today()

field_mapping = {
    'input_occ': ('occ_rate', 0.0), 'input_adr': ('adr', 0), 'input_rev': ('revenue', 0), 'input_rooms': ('total_rooms', 0),
    'input_complaints': ('counter_complaints', ""), 'input_counter_exp': ('counter_expense', 0),
    'input_cleaned': ('cleaned_rooms', 0), 'input_hk_co': ('hk_checkout_extend', 0), 'input_hk_avg': ('hk_avg_clean', 0.0), 'input_hk_exp': ('hk_expense', 0),
    'input_bf_theme_est': ('bf_theme_est', 0), 'input_bf_theme_act': ('bf_theme_act', 0), 'input_bf_zq_est': ('bf_zq_est', 0), 'input_bf_zq_act': ('bf_zq_act', 0), 'input_bf_total_est': ('bf_total_est', 0), 'input_bf_total_act': ('bf_total_act', 0),
    'input_af_theme_est': ('af_theme_est', 0), 'input_af_theme_act': ('af_theme_act', 0), 'input_af_zq_est': ('af_zq_est', 0), 'input_af_zq_act': ('af_zq_act', 0), 'input_af_total_est': ('af_total_est', 0), 'input_af_total_act': ('af_total_act', 0),
    'input_rest_mrev': ('rest_month_rev', 0), 'input_rest_aspent': ('rest_avg_spent', 0), 'input_rest_exp': ('rest_peak_expense', 0), 'input_hh_act': ('rest_hh_guests', 0), 'input_peak_act': ('rest_day_guests', 0),
    'input_repair': ('maint_repair_rooms', 0), 'input_maint_rec': ('maint_records', ""), 'input_maint_exp': ('maint_expense', 0)
}

def sync_st_to_db(target_d_str):
    db_data = get_daily_data(target_d_str)
    update_dict = {}; has_changes = False
    for ss_key, (db_col, default_val) in field_mapping.items():
        if ss_key in st.session_state:
            curr_val = st.session_state[ss_key]
            update_dict[db_col] = curr_val
            db_val = db_data.get(db_col)
            norm_db = default_val if pd.isna(db_val) or db_val is None else (int(float(db_val)) if isinstance(default_val, int) else (float(db_val) if isinstance(default_val, float) else str(db_val)))
            if curr_val != norm_db: has_changes = True
    if has_changes: save_daily_data(target_d_str, update_dict)
    if 'input_daily_log' in st.session_state:
        curr_log = st.session_state['input_daily_log'].strip(); db_log = str(get_daily_log(target_d_str) or "").strip()
        if curr_log != db_log: save_daily_log(target_d_str, curr_log)

def prev_day(): st.session_state['sidebar_date'] -= datetime.timedelta(days=1)
def next_day(): st.session_state['sidebar_date'] += datetime.timedelta(days=1)

col1, col2 = st.sidebar.columns(2)
col1.button("⬅️ 前一天", on_click=prev_day); col2.button("後一天 ➡️", on_click=next_day)
selected_date = st.sidebar.date_input("選擇日期", value=st.session_state['sidebar_date'], key='sidebar_date')
date_str = str(selected_date)

day_data = get_daily_data(date_str)
if st.session_state.get('_last_loaded_date') != date_str:
    for ss_key, (db_col, default_val) in field_mapping.items():
        val = day_data.get(db_col)
        st.session_state[ss_key] = default_val if pd.isna(val) or val is None else (int(val) if isinstance(default_val, int) else (float(val) if isinstance(default_val, float) else str(val)))
    st.session_state['input_daily_log'] = get_daily_log(date_str)
    st.session_state['_last_loaded_date'] = date_str

def on_input_change(): sync_st_to_db(date_str)

tab1, tab2, tab3, tab4, tab5, tab6, tab_p, tab7 = st.tabs(["📊 營運總覽", "📒 每日日誌", "🏨 櫃台房務", "🍽️ 餐廳數據", "🔧 工務數據", "📅 全月分析", "💰 採購分析", "👥 人事概況"])

with tab4:
    st.header("🍽️ 餐廳營運數據")
    st.markdown("#### 🌞 早餐數據")
    c1, c2, c3 = st.columns(3); c1.number_input("【主題】 實際來客", key="input_bf_theme_act", on_change=on_input_change); c2.number_input("【站前】 實際來客", key="input_bf_zq_act", on_change=on_input_change); c3.number_input("【兩館總和】 實際", key="input_bf_total_act", on_change=on_input_change)
    st.markdown("#### 🍰 下午茶數據")
    c1, c2, c3 = st.columns(3); c1.number_input("【主題】 實際來客", key="input_af_theme_act", on_change=on_input_change); c2.number_input("【站前】 實際來客", key="input_af_zq_act", on_change=on_input_change); c3.number_input("【兩館總和】 實際", key="input_af_total_act", on_change=on_input_change)
    st.markdown("#### 📊 月報結算總數與雜項")
    c1, c2, c3 = st.columns(3); c1.number_input("已結算營收 (全月)", key="input_rest_mrev", on_change=on_input_change); c2.number_input("平均客單價", key="input_rest_aspent", on_change=on_input_change); c3.number_input("THE PEAK 請購費用", key="input_rest_exp", on_change=on_input_change)
    col_rest1, col_rest2 = st.columns(2); col_rest1.number_input("The Peak 當日來客數 (手動)", key="input_peak_act", on_change=on_input_change); col_rest2.number_input("Happy Hour 當日來客數", key="input_hh_act", on_change=on_input_change)

with tab_p:
    st.header("💰 採購花費分析統計")
    try:
        possible_names = ["purchase data", "Purchase Data", "purchase_data", "Purchase_Data"]
        df_purchase = None; used_name = ""
        for name in possible_names:
            try:
                df_purchase = conn.read(worksheet=name, ttl="1m")
                if df_purchase is not None and not df_purchase.empty: used_name = name; break
            except: continue
        if df_purchase is not None and not df_purchase.empty:
            df_purchase.columns = df_purchase.columns.astype(str).str.strip()
            date_col = next((c for c in df_purchase.columns if '日期' in c or 'Date' in c), None)
            dept_col = next((c for c in df_purchase.columns if '部門' in c or 'Dept' in c or '工地' in c), None)
            total_col = next((c for c in df_purchase.columns if '小計' in c or '金額' in c or 'Total' in c), None)
            
            def robust_date_parse(val):
                if pd.isna(val): return None
                try: return pd.to_datetime(val).date()
                except: return None
            df_purchase['日期'] = df_purchase[date_col].apply(robust_date_parse)
            df_purchase[dept_col] = df_purchase[dept_col].fillna("未分類").astype(str).str.strip()
            df_purchase = df_purchase[df_purchase['日期'].notna()]
            m_start = selected_date.replace(day=1)
            import calendar; _, last_day = calendar.monthrange(selected_date.year, selected_date.month); m_end = selected_date.replace(day=last_day)
            df_month = df_purchase[(df_purchase['日期'] >= m_start) & (df_purchase['日期'] <= m_end)].copy()
            
            if not df_month.empty:
                df_month['小計'] = pd.to_numeric(df_month[total_col], errors='coerce').fillna(0)
                total_month_expense = df_month['小計'].sum()
                st.metric(f"📅 {selected_date.strftime('%Y-%m')} 總開銷", f"NT$ {int(total_month_expense):,}")
                
                dept_summary = df_month.groupby(dept_col)['小計'].sum().reset_index(); dept_summary.columns = ['部門', '小計']
                
                # --- 餐飲績效分析 (模糊匹配) ---
                st.divider(); st.subheader("🍽️ 餐飲績效與成本深度分析 (Cash-basis)")
                m_data = fetch_month_summary(selected_date.year, selected_date.month); df_daily_rest = m_data['df']
                
                if not df_daily_rest.empty:
                    target_cols = ['rest_day_guests', 'rest_hh_guests', 'bf_total_act', 'af_total_act']
                    for c in target_cols:
                        if c in df_daily_rest.columns: df_daily_rest[c] = pd.to_numeric(df_daily_rest[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                        else: df_daily_rest[c] = 0
                    
                    df_daily_rest['eff_peak_guests'] = df_daily_rest.apply(lambda r: r['rest_day_guests'] if r['rest_day_guests']>0 else (r['bf_total_act']+r['af_total_act']), axis=1)
                    
                    all_depts = dept_summary['部門'].tolist()
                    hh_matched = [d for d in all_depts if '4' in d or any(k in d.upper() for k in ['HH', 'HAPPY'])]
                    peak_matched = [d for d in all_depts if (any(k in d.upper() for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲'])) and (d not in hh_matched)]
                    
                    st.info(f"📍 偵測到部門: HH(`{hh_matched}`), Peak(`{peak_matched}`)")
                    
                    df_peak_p = df_month[df_month[dept_col].isin(peak_matched)].copy()
                    df_hh_p = df_month[df_month[dept_col].isin(hh_matched)].copy()
                    
                    total_peak_cost = df_peak_p['小計'].sum(); total_hh_cost = df_hh_p['小計'].sum()
                    total_peak_guests = df_daily_rest['eff_peak_guests'].sum(); total_hh_guests = df_daily_rest['rest_hh_guests'].sum()
                    peak_cpg = total_peak_cost / total_peak_guests if total_peak_guests > 0 else 0
                    hh_cpg = total_hh_cost / total_hh_guests if total_hh_guests > 0 else 0
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**🏰 The Peak (餐廳)**")
                        st.metric("本月總採購額", f"NT$ {int(total_peak_cost):,}")
                        is_auto = "(自動)" if (df_daily_rest['rest_day_guests'].sum() == 0 and total_peak_guests > 0) else ""
                        st.metric(f"本月總來客數 {is_auto}", f"{int(total_peak_guests):,} 人")
                        st.metric("平均每客成本 (CPG)", f"NT$ {int(peak_cpg):,}")
                    with c2:
                        st.markdown("**🥂 Happy Hour (HH)**")
                        st.metric("本月總採購額", f"NT$ {int(total_hh_cost):,}")
                        st.metric("本月總來客數", f"{int(total_hh_guests):,} 人")
                        st.metric("平均每客服務成本", f"NT$ {int(hh_cpg):,}")
    except Exception as e: st.error(f"分析出錯: {e}")
