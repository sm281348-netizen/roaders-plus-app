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
        # 讀取試算表中的 taipei_events 分頁
        df = conn.read(worksheet="taipei_events", ttl="10m")
        if df is not None and not df.empty:
            df = standardize_df_dates(df)
            return df
    except Exception:
        # 如果分頁不存在或讀取失敗，回傳空 DataFrame
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
    """回傳當月所有目標國家的國定假日字典: { 'YYYY-MM-DD': {'flags': '...', 'details': [...] } }"""
    h_objs = {}
    for code in list(TARGET_HOLIDAY_COUNTRIES.keys()) + OTHER_HOLIDAY_COUNTRIES:
        try:
            h_objs[code] = holidays.country_holidays(code, years=[year])
        except Exception as e:
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
        
        # Sort flags to maintain consistent order
        flags_str = "".join(sorted(list(day_flags)))
        if has_other:
            flags_str += "🌍"
            
        if flags_str:
            result[dt_str] = {
                'flags': flags_str,
                'details': day_details
            }
            
    return result

@st.cache_data(ttl=86400)
def fetch_upcoming_holidays(start_date, days=30):
    """回傳未來 N 天內的假日"""
    years = {start_date.year, (start_date + datetime.timedelta(days=days)).year}
    h_objs = {}
    for code in list(TARGET_HOLIDAY_COUNTRIES.keys()) + OTHER_HOLIDAY_COUNTRIES:
        try:
            h_objs[code] = holidays.country_holidays(code, years=list(years))
        except:
            pass
            
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

# 設定頁面資訊
st.set_page_config(page_title="路徒Plus行旅站前館營運日誌", layout="wide")

# --- 安全防護：全站密碼攔截 ---
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
            else:
                st.error("❌ 密碼錯誤，請重新輸入。")
    st.stop()
# -----------------------------

# -- 資料庫連線初始化 (Google Sheets 版) --
conn = st.connection("gsheets", type=GSheetsConnection)

def init_db():
    """
    確保 Google Sheets 中有正確的分頁與標題行。
    由於 st-gsheets-connection 的運作機制，初次使用時需確保 Sheets 存在。
    """
    # 這裡我們不撰寫複雜的初始化代碼，因為使用者需先建立 Sheet。
    # 但我們可以預定義欄位給後續寫入使用。
    pass
init_db()

# -- 基本資料庫讀寫函數 (需優先定義以供導航邏輯使用) --
def standardize_df_dates(df):
    if df is None or df.empty or 'date' not in df.columns:
        return df
    def fix_d(val):
        if pd.isna(val) or str(val).strip() == '' or str(val).strip() == 'NaT': 
            return ""
        v_str = str(val).split(' ')[0].strip()
        
        import re
        # 處理民國年或簡寫 (例如 115/4/30 或 115-04-30)
        m_tw = re.match(r'^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})$', v_str)
        if m_tw:
            y, m, d = int(m_tw.group(1)), int(m_tw.group(2)), int(m_tw.group(3))
            if y < 1000: y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
            
        # 處理只有月跟日的狀況 (例如 4/30 或 04-30)
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
        # 讀取完整表單 (快取設定為 1 分鐘)
        df = conn.read(worksheet="daily_data", ttl="1m")
        if df is not None and not df.empty:
            # 確保日期欄位為字串格式 (YYYY-MM-DD) 以供比對
            df = standardize_df_dates(df)
            # 確保唯一
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
    except Exception:
        pass
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
    except Exception as e:
        st.error(f"儲存失敗: {e}")

def get_monthly_target(month_str):
    try:
        df = conn.read(worksheet="targets", ttl="1m")
        if df is not None and not df.empty:
            res = df[df['month'] == month_str]
            if not res.empty:
                return int(res.iloc[0]['target_revenue'])
    except Exception:
        pass
    return 0

def save_monthly_target(month_str, target):
    try:
        df = conn.read(worksheet="targets", ttl="0")
        if df is None or df.empty:
            df = pd.DataFrame(columns=["month", "target_revenue"])
        
        if month_str in df['month'].values:
            df.loc[df['month'] == month_str, 'target_revenue'] = target
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
            if not res.empty:
                return str(res.iloc[0]['log']).strip()
    except:
        pass
    # Fallback to daily_data if not found in daily_logs (backward compatibility)
    try:
        df_old = conn.read(worksheet="daily_data", ttl="1m")
        if df_old is not None and not df_old.empty:
            res = df_old[df_old['date'] == d_str]
            if not res.empty and 'daily_work_log' in res.columns:
                return str(res.iloc[0]['daily_work_log']).strip()
    except:
        pass
    return ""

def save_daily_log(d_str, log_text):
    try:
        df = conn.read(worksheet="daily_logs", ttl="0")
        if df is None or df.empty:
            df = pd.DataFrame(columns=["date", "log"])
            
        df = standardize_df_dates(df)
        
        # 確保欄位存在
        if 'date' not in df.columns or 'log' not in df.columns:
            df = pd.DataFrame(columns=["date", "log"])

        new_row = pd.DataFrame([{'date': d_str, 'log': log_text}])
        
        if d_str in df['date'].values:
            df = df[df['date'] != d_str]
        
        df = pd.concat([df, new_row], ignore_index=True)
        if 'date' in df.columns:
            df = df.sort_values('date').reset_index(drop=True)
        conn.update(worksheet="daily_logs", data=df.fillna(""))
        st.cache_data.clear()
        st.toast(f"✅ {d_str} 日誌已自動對齊 Google Sheet！")
        return True
    except Exception as e:
        # 這裡不使用 st.error 以免干擾輸入，但可以列印到日誌或使用 toast
        print(f"DEBUG: 日誌儲存失敗: {e}")
        return False

def get_month_delta(d, delta):
    year = d.year
    month = d.month + delta
    while month > 12:
        month -= 12
        year += 1
    while month < 1:
        month += 12
        year -= 1
    return datetime.date(year, month, 1)

def prepare_monthly_report(year, month):
    try:
        df_all = conn.read(worksheet="daily_data", ttl="10m")
        month_str = f"{year}-{month:02d}"
        df = df_all[df_all['date'].str.startswith(month_str, na=False)].sort_values('date')
    except:
        return "--- 讀取失敗 ---"
    
    if df.empty: return "--- 當月無紀錄 ---"
    
    full_report = ""
    for d in sorted(df['date'].unique()):
        full_report += generate_report_text(d) + "\n\n"
    return full_report

def minguo_to_western(d_str):
    """
    將 民國/月/日 (如 115/03/02 或 0115/03/02) 轉換為 Python date 對象。
    """
    if pd.isna(d_str) or not isinstance(d_str, str): return None
    try:
        # 移除前導零並拆分
        parts = d_str.strip().split('/')
        if len(parts) == 3:
            year = int(parts[0])
            # 如果是 115 或 0115，這應是民國年
            if year < 1000: # 民國年編號通常不大於 1000
                year += 1911
            return datetime.date(year, int(parts[1]), int(parts[2]))
    except:
        pass
    return None

def fetch_month_summary(year, month):
    import calendar
    m_start = f"{year}-{month:02d}-01"
    _, last_day = calendar.monthrange(year, month)
    m_end = f"{year}-{month:02d}-{last_day:02d}"
    
    try:
        df_all = conn.read(worksheet="daily_data", ttl="10m")
        if df_all is not None and not df_all.empty:
            df_all = standardize_df_dates(df_all)
            # 確保日期唯一，避免重複加總
            df_all = df_all.drop_duplicates(subset='date', keep='last')
            df = df_all[(df_all['date'] >= m_start) & (df_all['date'] <= m_end)].copy()
        else:
            df = pd.DataFrame()
    except Exception:
        df = pd.DataFrame()
    
    res = {
        'rev': 0.0, 'rooms': 0.0, 'sellable': 0.0, 'occ90_days': 0,
        'avg_occ': 0.0, 'avg_adr': 0.0, 'revpar': 0.0, 'df': df,
        'month_label': f"{year}-{month:02d}"
    }
    
    if not df.empty:
        # 確保數值欄位為 float
        num_cols = ['revenue', 'total_rooms', 'occ_rate', 'adr']
        for c in num_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(',', '').str.replace('%', '')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

        for _, r in df.iterrows():
            rev = float(r['revenue'])
            rm = float(r['total_rooms'])
            occ = float(r['occ_rate'])
            adr = float(r['adr'])
            
            if rev == 0 and adr > 0 and rm > 0: rev = adr * rm
            if rm == 0 and rev > 0 and adr > 0: rm = rev / adr
            
            if rm > 0 or rev > 0:
                res['rev'] += rev
                res['rooms'] += rm
                if occ > 0:
                    res['sellable'] += (rm / (occ / 100.0))
                if occ >= 90.0:
                    res['occ90_days'] += 1
        
        res['avg_occ'] = (res['rooms'] / res['sellable'] * 100.0) if res['sellable'] > 0 else 0.0
        res['avg_adr'] = (res['rev'] / res['rooms']) if res['rooms'] > 0 else 0.0
        res['revpar'] = (res['avg_occ'] / 100.0) * res['avg_adr']
        
    return res

@st.cache_data(ttl=3600)
def fetch_yearly_metrics(year):
    try:
        df_all = conn.read(worksheet="daily_data", ttl="10m")
        if df_all is None or df_all.empty: return 0.0, 0.0
        df_all = standardize_df_dates(df_all)
        df_all = df_all.drop_duplicates(subset='date', keep='last')
        
        y_start = f"{year}-01-01"
        y_end = f"{year}-12-31"
        df = df_all[(df_all['date'] >= y_start) & (df_all['date'] <= y_end)].copy()
        if df.empty: return 0.0, 0.0
        
        num_cols = ['revenue', 'total_rooms']
        for c in num_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(',', '').str.replace('%', '')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
                
        tot_rev = df['revenue'].sum()
        tot_rms = df['total_rooms'].sum()
        yearly_adr = (tot_rev / tot_rms) if tot_rms > 0 else 0.0
        
        taipei_events_df = fetch_taipei_events()
        e_dates = set(taipei_events_df['date'].unique()) if not taipei_events_df.empty else set()
        
        h_dates = set()
        for m in range(1, 13):
            h_dict = fetch_holidays_for_month(year, m)
            for d_str, info in h_dict.items():
                if info['flags']: h_dates.add(d_str)
                
        df['is_e'] = df['date'].isin(e_dates)
        df['is_h'] = df['date'].isin(h_dates)
        df_pure = df[~df['is_e'] & ~df['is_h']]
        
        p_rev = df_pure['revenue'].sum()
        p_rms = df_pure['total_rooms'].sum()
        yearly_pure_adr = (p_rev / p_rms) if p_rms > 0 else 0.0
        
        return yearly_adr, yearly_pure_adr
    except Exception:
        return 0.0, 0.0

# -- 側邊欄：進階日期選擇器 --
st.sidebar.caption(f"🚀 最後更新時間: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
st.sidebar.header("📅 日期選擇")
if 'sidebar_date' not in st.session_state:
    st.session_state['sidebar_date'] = datetime.date.today()

# 定義欄位映射 (必須在儲存與載入函數之前)
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
    # 先獲取目前的 DB 資料作為比對基準
    db_data = get_daily_data(target_d_str)
    
    # 同步數值數據
    update_dict = {}
    has_changes = False
    
    for ss_key, (db_col, default_val) in field_mapping.items():
        if ss_key in st.session_state:
            curr_val = st.session_state[ss_key]
            update_dict[db_col] = curr_val
            
            # 從 DB 解析原本應該長怎樣
            db_val = db_data.get(db_col)
            if pd.isna(db_val) or db_val is None:
                norm_db = default_val
            else:
                try:
                    if isinstance(default_val, int): norm_db = int(float(db_val))
                    elif isinstance(default_val, float): norm_db = float(db_val)
                    else: norm_db = str(db_val)
                except:
                    norm_db = default_val
            
            # 判斷是否真的改變
            if isinstance(curr_val, float) and isinstance(norm_db, float):
                import math
                if not math.isclose(curr_val, norm_db, abs_tol=1e-5):
                    has_changes = True
                    with open("debug_save.log", "a") as f: f.write(f"[{target_d_str}] DIFF: {db_col} ({type(curr_val)})={curr_val} vs DB {norm_db} ({type(norm_db)})\n")
            elif curr_val != norm_db:
                has_changes = True
                with open("debug_save.log", "a") as f: f.write(f"[{target_d_str}] DIFF: {db_col} ({type(curr_val)})={repr(curr_val)} vs DB {repr(norm_db)} ({type(norm_db)})\n")

    if has_changes:
        save_daily_data(target_d_str, update_dict)
    
    # 單獨同步日誌
    if 'input_daily_log' in st.session_state:
        curr_log = st.session_state['input_daily_log'].strip()
        db_log = str(get_daily_log(target_d_str) or "").strip()
        if curr_log != db_log:
            with open("debug_save.log", "a") as f: f.write(f"[{target_d_str}] LOG DIFF: curr={repr(curr_log)} vs db={repr(db_log)}\n")
            save_daily_log(target_d_str, curr_log)

def prev_day():
    st.session_state['sidebar_date'] -= datetime.timedelta(days=1)

def next_day():
    st.session_state['sidebar_date'] += datetime.timedelta(days=1)

col1, col2 = st.sidebar.columns(2)
col1.button("⬅️ 前一天", on_click=prev_day)
col2.button("後一天 ➡️", on_click=next_day)

selected_date = st.sidebar.date_input("選擇日期", value=st.session_state['sidebar_date'], key='sidebar_date')
date_str = str(selected_date)

# 追蹤當前正在編輯的日期
if '_actual_current_date' not in st.session_state:
    st.session_state['_actual_current_date'] = date_str
if '_data_is_loaded' not in st.session_state:
    st.session_state['_data_is_loaded'] = False

if st.session_state['_actual_current_date'] != date_str:
    # 移除原本無條件在切換日期時自動存檔的邏輯，避免單純查看舊日期造成不必要的寫入或位置跳動
    st.session_state['_actual_current_date'] = date_str
    st.session_state['_data_is_loaded'] = False

# --- 新增：週次預覽選擇器 ---
weekly_options = ["--- 關閉週預覽 ---", "第1週 (1-7號)", "第2週 (8-14號)", "第3週 (15-21號)", "第4週 (22-28號)", "第5週 (29號起)"]
selected_week = st.sidebar.selectbox("快速查閱區間：", weekly_options, index=0, key="weekly_view_select")
# --------------------------------------------------




day_data = get_daily_data(date_str)
if st.session_state.get('_last_loaded_date') != date_str or st.session_state.get('_last_week_view') != selected_week:
    for ss_key, (db_col, default_val) in field_mapping.items():
        val = day_data.get(db_col)
        # Handle nan/null from Pandas/SQLite gracefully
        if pd.isna(val) or val is None:
            st.session_state[ss_key] = default_val
        else:
            if isinstance(default_val, int): st.session_state[ss_key] = int(val)
            elif isinstance(default_val, float): st.session_state[ss_key] = float(val)
            else: st.session_state[ss_key] = str(val)
    
    # 獲取日誌
    st.session_state['input_daily_log'] = get_daily_log(date_str)
    
    st.session_state['_last_loaded_date'] = date_str
    st.session_state['_last_week_view'] = selected_week
    st.session_state['_data_is_loaded'] = True # 標記為已載入，此後任何變動或換日才允許存檔

def on_input_change():
    # 使用 session_state 中的當前日期，確保 callback 觸發時日期正確
    target_d = st.session_state.get('_actual_current_date')
    if target_d:
        sync_st_to_db(target_d)

st.sidebar.divider()
st.sidebar.subheader("📤 數據匯出與備份")

def generate_report_text(d_str):
    data = get_daily_data(d_str)
    if not data: return f"--- {d_str} 無紀錄 ---"
    
    report = []
    report.append(f"========================================")
    report.append(f"🏨 路徒行旅 Plus 站前館 - 營運日誌 ({d_str})")
    report.append(f"========================================\n")
    
    def safe_int_val(v):
        try:
            if pd.isna(v) or v is None: return 0
            return int(float(v))
        except: return 0

    report.append(f"【📊 營運指標】")
    report.append(f"- 住房率: {data.get('occ_rate', 0)}%")
    report.append(f"- ADR: NT$ {safe_int_val(data.get('adr', 0)):,}")
    report.append(f"- 總營收: NT$ {safe_int_val(data.get('revenue', 0)):,}")
    report.append(f"- 總住房數: {safe_int_val(data.get('total_rooms', 0))} 間\n")
    
    report.append(f"【💼 櫃台與房務】")
    report.append(f"- 負評客訴: {data.get('counter_complaints', '無')}")
    report.append(f"- 櫃台請購: {safe_int_val(data.get('counter_expense', 0))} 元")
    report.append(f"- 總清消房數: {safe_int_val(data.get('cleaned_rooms', 0))} 間")
    report.append(f"- 房務請購: {safe_int_val(data.get('hk_expense', 0))} 元\n")
    
    report.append(f"【🍽️ 餐廳數據 (兩館實際來客)】")
    report.append(f"- 早餐總計: {safe_int_val(data.get('bf_total_act', 0))} 人")
    report.append(f"- 下午茶總計: {safe_int_val(data.get('af_total_act', 0))} 人")
    report.append(f"- Happy Hour: {safe_int_val(data.get('rest_hh_guests', 0))} 人")
    report.append(f"- 餐廳營收(全月): {safe_int_val(data.get('rest_month_rev', 0))} 元\n")
    
    report.append(f"【🔧 工務紀錄】")
    report.append(f"- 待修房數: {data.get('maint_repair_rooms', 0)} 間")
    report.append(f"- 修繕細節: {data.get('maint_records', '無')}\n")
    
    report.append(f"【📝 每日營運紀錄細節】")
    report.append(f"{get_daily_log(d_str) or '無紀錄内容'}")
    report.append(f"\n" + "-"*40 + "\n")
    
    return "\n".join(report)

# 1. 單日匯出
single_report = generate_report_text(date_str)
st.sidebar.download_button(
    label="📄 當日營運紀錄匯出",
    data=single_report,
    file_name=f"Roaders_Plus_Daily_{date_str}.txt",
    mime="text/plain",
    use_container_width=True
)

# 2. 全月匯出
month_str = selected_date.strftime('%Y-%m')
if f"monthly_report_{month_str}" not in st.session_state:
    if st.sidebar.button(f"📅 當月 {month_str} 營運紀錄匯出", use_container_width=True):
        df_all = conn.read(worksheet="daily_data", ttl="0")
        df_month = df_all[df_all['date'].str.startswith(month_str, na=False)].sort_values('date')
        
        if df_month.empty:
            st.sidebar.warning(f"⚠️ {month_str} 尚無任何資料。")
        else:
            with st.sidebar.status("正在產生報表...", expanded=False):
                full_month_text = f"【路徒行旅 Plus 站前館 {month_str} 全月營運紀錄匯總】\n\n"
                for d in df_all['date']:
                    full_month_text += generate_report_text(d) + "\n\n"
                st.session_state[f"monthly_report_{month_str}"] = full_month_text
            st.rerun()
else:
    st.sidebar.download_button(
        label=f"⬇️ 下載 {month_str} 紀錄 (.txt)",
        data=st.session_state[f"monthly_report_{month_str}"],
        file_name=f"Roaders_Plus_Monthly_{month_str}.txt",
        mime="text/plain",
        use_container_width=True,
    )
    if st.sidebar.button("🔄 重新產生", key="clear_monthly"):
        del st.session_state[f"monthly_report_{month_str}"]
        st.rerun()

# 側邊欄底部移除多餘區塊

# -- 報表解析與寫入資料庫 --
def parse_and_save_jinxu(file):
    try:
        if file.name.endswith('.csv'):
            df_test = pd.read_csv(file, nrows=20, header=None)
            is_csv = True
        else:
            df_test = pd.read_excel(file, nrows=20, header=None)
            is_csv = False

        header_idx = 0
        for i in range(len(df_test)):
            row_str = "".join(str(val) for val in df_test.iloc[i].values)
            if '日期' in row_str:
                header_idx = i
                break
                
        file.seek(0)
        try:
            df = pd.read_csv(file, skiprows=header_idx) if is_csv else pd.read_excel(file, skiprows=header_idx)
        except Exception as e:
            # 嘗試不同的 engine
            file.seek(0)
            df = pd.read_excel(file, skiprows=header_idx, engine='openpyxl')
            
        df.columns = df.columns.astype(str).str.replace(r'[\s\n\r]', '', regex=True)
        
        date_col = next((c for c in df.columns if '日期' in c), None)
        occ_col = next((c for c in df.columns if '住房率' in c or '訂房率' in c or '出租率' in c or 'OCC' in c.upper()), None)
        adr_col = next((c for c in df.columns if '平均房價' in c or 'ADR' in c.upper()), None)
        
        rev_col = next((c for c in df.columns if '客房收入' in c or '客房營收' in c or '總營收' in c or '營業額' in c or '實際營收' in c), None)
        rooms_col = next((c for c in df.columns if ('住房數' in c or '出租' in c or '售出' in c or '實住' in c) and '可售' not in c), None)
        if not rooms_col:
            rooms_col = next((c for c in df.columns if ('房間數' in c or '客房數' in c) and '可售' not in c), None)

        if not date_col:
            st.error("⚠️ 解析失敗：找不到『日期』欄位，請檢查報表格式。")
            return False

        # --- 強化日期解析邏輯 ---
        def robust_parse_date(val):
            if pd.isna(val) or str(val).strip() == '': return None
            s = str(val).strip().split('.')[0] # 移除 .0
            # 嘗試 YYYYMMDD
            try:
                if len(s) == 8 and s.isdigit():
                    return pd.to_datetime(s, format='%Y%m%d').date()
            except: pass
            # 嘗試一般解析 (YYYY-MM-DD, YYYY/MM/DD 等)
            try:
                return pd.to_datetime(s).date()
            except: pass
            return None

        df['標準日期'] = df[date_col].apply(robust_parse_date)
        
        df_new_records = pd.DataFrame()
        updates = []
        for index, row in df.iterrows():
            d_obj = row['標準日期']
            if not d_obj: continue
            
            d_str = str(d_obj)
            occ = 0.0
            if occ_col and pd.notna(row.get(occ_col)):
                raw_occ_str = str(row.get(occ_col)).strip()
                has_percent = '%' in raw_occ_str
                clean_occ = raw_occ_str.replace('%', '').replace(',', '')
                try:
                    occ = float(clean_occ)
                    if 0 < occ <= 1.0 and not has_percent:
                        occ = occ * 100.0
                except: occ = 0.0

            adr = int(float(str(row.get(adr_col, '0')).replace(',', ''))) if adr_col and pd.notna(row.get(adr_col)) else 0
            rev = int(float(str(row.get(rev_col, '0')).replace(',', ''))) if rev_col and pd.notna(row.get(rev_col)) else 0
            rooms = int(float(str(row.get(rooms_col, '0')).replace(',', ''))) if rooms_col and pd.notna(row.get(rooms_col)) else 0

            updates.append({'date': d_str, 'occ_rate': occ, 'adr': adr, 'revenue': rev, 'total_rooms': rooms})
            
        if updates:
            df_existing = conn.read(worksheet="daily_data", ttl="0")
            if df_existing is None: df_existing = pd.DataFrame()
            df_existing = standardize_df_dates(df_existing)
            df_new = pd.DataFrame(updates)
            
            # 合併數據 (以日期為 key，部分更新)
            df_new = df_new.set_index('date')
            if not df_existing.empty:
                df_existing = df_existing.set_index('date')
                df_final = df_new.combine_first(df_existing).reset_index()
            else:
                df_final = df_new.reset_index()
                
            if 'date' in df_final.columns:
                df_final = df_final.sort_values('date').reset_index(drop=True)
                
            conn.update(worksheet="daily_data", data=df_final.fillna(""))
            st.cache_data.clear()
            return len(updates)
            
        return 0
    except Exception as e:
        import traceback
        st.error(f"解析櫃台報表失敗: {e}\n{traceback.format_exc()}")
        return False

# -- 餐廳報表解析與寫入資料庫 --
def parse_and_save_restaurant(file, current_year):
    try:
        # 讀取 Excel 檔案的所有內容
        df = pd.read_excel(file, header=None)
        
        month_rev = 0
        avg_spent = 0
        
        # 1. 改良版：搜尋全表尋找月結算關鍵字 (不再侷限於第 0 欄)
        for i, row in df.iterrows():
            row_str = " ".join([str(v) for v in row if pd.notna(v)])
            # 尋找營收
            if ('已結算營收' in row_str or '月營收' in row_str) and '早餐' not in row_str and '下午茶' not in row_str:
                for val in row:
                    s_val = str(val).strip()
                    if any(c.isdigit() for c in s_val) and not any(k in s_val for k in ['已結算營收', '月營收']):
                        try:
                            clean_val = s_val.replace('NT$', '').replace('$', '').replace(',', '').strip()
                            month_rev = int(float(clean_val))
                            break
                        except: continue
            # 尋找客單價
            if '平均客單價' in row_str or '客單價' in row_str:
                for val in row:
                    s_val = str(val).strip()
                    if any(c.isdigit() for c in s_val) and '客單價' not in s_val:
                        try:
                            clean_val = s_val.replace('NT$', '').replace('$', '').replace(',', '').strip()
                            avg_spent = int(float(clean_val))
                            break
                        except: continue

        parsed_days = []
        # 2. 尋找每日明細 (修正 Regex 讓其更具包容度)
        for i, row in df.iterrows():
            col0 = str(row[0]).strip()
            m = re.search(r'(\d{1,2})/(\d{1,2})', col0)
            if m:
                month_val, day_val = m.groups()
                d_str = f"{current_year}-{int(month_val):02d}-{int(day_val):02d}"
                
                def safe_int(val):
                    if pd.isna(val) or str(val).strip() == '': return 0
                    try: 
                        # 處理 Excel 讀入時可能的科學符號或逗號
                        return int(float(str(val).replace(',', '').strip()))
                    except: return 0
                
                # 假設欄位順序不變 (根據 Roaders Plus 常用報表格式)
                # 早餐相關 (1-6)
                row_vals = row.values.tolist()
                bf_theme_est = safe_int(row_vals[1]) if len(row_vals) > 1 else 0
                bf_theme_act = safe_int(row_vals[2]) if len(row_vals) > 2 else 0
                bf_zq_est = safe_int(row_vals[3]) if len(row_vals) > 3 else 0
                bf_zq_act = safe_int(row_vals[4]) if len(row_vals) > 4 else 0
                bf_total_est = safe_int(row_vals[5]) if len(row_vals) > 5 else 0
                bf_total_act = safe_int(row_vals[6]) if len(row_vals) > 6 else 0
                
                # 下午茶相關 (7-12)
                af_theme_est = safe_int(row_vals[7]) if len(row_vals) > 7 else 0
                af_theme_act = safe_int(row_vals[8]) if len(row_vals) > 8 else 0
                af_zq_est = safe_int(row_vals[9]) if len(row_vals) > 9 else 0
                af_zq_act = safe_int(row_vals[10]) if len(row_vals) > 10 else 0
                af_total_est = safe_int(row_vals[11]) if len(row_vals) > 11 else 0
                af_total_act = safe_int(row_vals[12]) if len(row_vals) > 12 else 0
                
                parsed_days.append({
                    'date': d_str,
                    'rest_month_rev': month_rev, 'rest_avg_spent': avg_spent,
                    'bf_theme_est': bf_theme_est, 'bf_theme_act': bf_theme_act,
                    'bf_zq_est': bf_zq_est, 'bf_zq_act': bf_zq_act,
                    'bf_total_est': bf_total_est, 'bf_total_act': bf_total_act,
                    'af_theme_est': af_theme_est, 'af_theme_act': af_theme_act,
                    'af_zq_est': af_zq_est, 'af_zq_act': af_zq_act,
                    'af_total_est': af_total_est, 'af_total_act': af_total_act
                })

        if parsed_days:
            # 讀取現有庫內資料
            df_existing = conn.read(worksheet="daily_data", ttl="0")
            if df_existing is None: df_existing = pd.DataFrame()
            
            # 重要：確保現有資料的 date 也是字串，否則 combine_first 的 join 會失效
            df_existing = standardize_df_dates(df_existing)
            
            # --- 修復：如果成功解析出月結算營收或客單價，強制更新現有資料庫中該月份的所有紀錄 ---
            # 避免使用者先點擊了未來的日期產生了帶有舊營收的紀錄，導致 MTD 永遠抓到最後一天的舊數據
            months = set("-".join(str(d['date']).split('-')[:2]) for d in parsed_days)
            if not df_existing.empty and 'date' in df_existing.columns:
                for m in months:
                    mask = df_existing['date'].str.startswith(m, na=False)
                    if mask.any():
                        if month_rev > 0:
                            df_existing.loc[mask, 'rest_month_rev'] = month_rev
                        if avg_spent > 0:
                            df_existing.loc[mask, 'rest_avg_spent'] = avg_spent
            
            df_new = pd.DataFrame(parsed_days)
            
            # 合併數據 (以日期為 key，部分更新)
            df_new = df_new.set_index('date')
            if not df_existing.empty:
                df_existing = df_existing.set_index('date')
                # 以新上傳的資料優先蓋掉舊的，但如果是新資料缺少的欄位則保留舊的
                df_final = df_new.combine_first(df_existing).reset_index()
            else:
                df_final = df_new.reset_index()
                
            if 'date' in df_final.columns:
                df_final = df_final.sort_values('date').reset_index(drop=True)
                
            # 寫回資料庫
            conn.update(worksheet="daily_data", data=df_final.fillna(""))
            st.cache_data.clear()
            
        # 清除快取以確保重整後能看到新數據
        st.session_state['_last_loaded_date'] = None
        return len(parsed_days)
    except Exception as e:
        import traceback
        st.error(f"解析餐廳報表失敗: {str(e)}")
        with st.expander("🔍 查看錯誤細節"):
            st.code(traceback.format_exc())
        with open("debug_error.log", "w") as f:
            f.write(traceback.format_exc())
        return False

# 頁面標題
st.title("路徒Plus行旅站前館營運日誌")

# 主畫面
tab1, tab_m, tab6, tab_p, tab3, tab4, tab5, tab7 = st.tabs(["📊 營運總覽", "📈 月分析專區", "📝 每日營運紀錄", "💰 採購分析", "🧹 房務數據", "🍽️ 餐廳數據", "🔧 工務數據", "👥 人事概況"])


with tab1:
    st.header("📊 營運總覽")
    
    # 注入專屬 CSS 與 Card 產生器
    st.markdown("""
    <style>
    .metric-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 18px 20px;
        margin: 8px 0 16px 0;
        box-shadow: 0 4px 10px rgba(0,0,0,0.06);
        border-left: 6px solid #4CAF50;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.12);
    }
    .metric-title {
        color: #7f8c8d;
        font-size: 0.95rem;
        font-weight: 600;
        margin-bottom: 8px;
    }
    .metric-value {
        color: #2c3e50;
        font-size: 1.8rem;
        font-weight: 800;
        letter-spacing: 0.5px;
    }
    .card-theme-blue { border-left-color: #3498db; }
    .card-theme-orange { border-left-color: #f39c12; }
    .card-theme-purple { border-left-color: #9b59b6; }
    .card-theme-red { border-left-color: #e74c3c; }
    .card-theme-green { border-left-color: #2ecc71; }
    .card-bg-dark {
        background: linear-gradient(135deg, #1f2c56 0%, #2e437c 100%);
    }
    .card-bg-dark .metric-title { color: #d8e2fb; }
    .card-bg-dark .metric-value { color: #ffffff; }
    </style>
    """, unsafe_allow_html=True)

    def make_card(title, value, color_class="card-theme-blue", bg_class="", icon=""):
        return f'''
        <div class="metric-card {color_class} {bg_class}">
            <div class="metric-title">{icon} {title}</div>
            <div class="metric-value">{value}</div>
        </div>
        '''
    
    occ_val = st.session_state.get('input_occ', 0.0)
    if occ_val >= 90.0:
        st.success("🎉 **滿房慶祝！今日住房率達到 90% 以上，全館辛苦了！** 🎉")

    # -- 今日看板 --
    st.subheader(f"今日全館營運大看板 ({date_str})")
    adr_val = st.session_state.get('input_adr', 0)
    rev_val = st.session_state.get('input_rev', 0)
    
    def safe_format_int(v):
        try:
            if pd.isna(v) or v is None: return 0
            return int(float(v))
        except: return 0

    kpi_html = f"""
    <style>
    .kpi-container {{ display: flex; justify-content: space-around; flex-wrap: wrap; margin-bottom: 30px; }}
    .kpi-circle {{ width: 170px; height: 170px; border-radius: 50%; background: linear-gradient(135deg, #1f2c56 0%, #2e437c 100%); color: white; display: flex; flex-direction: column; justify-content: center; align-items: center; box-shadow: 0 8px 15px rgba(0,0,0,0.15); border: 4px solid #4CAF50; margin: 15px; }}
    .kpi-title {{ font-size: 16px; margin-bottom: 8px; color: #d8e2fb; }}
    .kpi-value {{ font-size: 26px; font-weight: bold; }}
    </style>
    <div class="kpi-container">
        <div class="kpi-circle"><div class="kpi-title">今日住房率</div><div class="kpi-value">{occ_val}%</div></div>
        <div class="kpi-circle"><div class="kpi-title">ADR</div><div class="kpi-value">NT$ {safe_format_int(adr_val):,}</div></div>
        <div class="kpi-circle"><div class="kpi-title">總營收</div><div class="kpi-value">NT$ {safe_format_int(rev_val):,}</div></div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.success("🧹 **房務狀況**")
        total_occ = st.session_state.get('input_rooms', 0)
        cleaned = st.session_state.get('input_cleaned', 0)
        st.metric("目標清消總數 (來自金旭)", f"{total_occ} 間")
        st.caption(f"手動紀錄清消: {cleaned} 間 (差額: {cleaned - total_occ})")
    with col2:
        st.warning("🔧 **工務狀況**")
        repairs = st.session_state.get('input_repair', 0)
        st.metric("今日待修房數", f"{repairs} 間", delta="🔴 需處理" if repairs>0 else "🟢 正常", delta_color="off")
    with col3:
        st.error("🍽️ **餐廳狀況**")
        bf_total_act = st.session_state.get('input_bf_total_act', 0)
        st.metric("今日雙館早餐總來客", f"{safe_format_int(bf_total_act)} 人")
        
    st.divider()

    # -- 月度累計模式 (MTD Analysis) --
    st.subheader(f"📅 本月累計分析 (MTD: {selected_date.strftime('%Y-%m')})")
    start_of_month = selected_date.replace(day=1).strftime('%Y-%m-%d')
    
    try:
        df_all = conn.read(worksheet="daily_data", ttl="10m")
        if df_all is not None and not df_all.empty:
            df_all = standardize_df_dates(df_all)
            # 防止重複資料毀掉加總
            df_all = df_all.drop_duplicates(subset='date', keep='last')
            df_mtd = df_all[(df_all['date'] >= start_of_month) & (df_all['date'] <= date_str)].copy()
        else:
            df_mtd = pd.DataFrame()
    except Exception as e:
        st.sidebar.error(f"⚠️ 讀取數據時發生錯誤: {e}")
        df_mtd = pd.DataFrame()

    if not df_mtd.empty:
        # 先將所有可能計算的欄位轉為數值，避免 Google Sheets 帶來的字串問題
        for col in ['bf_theme_act', 'bf_zq_act', 'af_theme_act', 'af_zq_act',
                    'bf_total_act', 'af_total_act', 'bf_total_est', 'af_total_est',
                    'rest_month_rev', 'rest_avg_spent']:
            if col in df_mtd.columns:
                df_mtd[col] = pd.to_numeric(df_mtd[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

        mtd_rooms = 0.0
        mtd_rev = 0.0
        total_sellable = 0.0
        
        for _, r in df_mtd.iterrows():
            # 強化字串清理防護
            def clean_num(val):
                if pd.isna(val): return 0.0
                try: return float(str(val).replace(',', '').replace('%', ''))
                except: return 0.0
                
            o = clean_num(r.get('occ_rate'))
            adr = clean_num(r.get('adr'))
            rev = clean_num(r.get('revenue'))
            rm = clean_num(r.get('total_rooms'))
            
            # 容錯處理：若 Excel 某天缺營收但有 ADR 和房數，或缺房數但有營收，做數學回推
            if rev == 0 and adr > 0 and rm > 0:
                rev = adr * rm
            if rm == 0 and rev > 0 and adr > 0:
                rm = rev / adr
                
            # 只加總有實際營業數據的日期（排除未來的 0）
            if rm > 0 or rev > 0:
                mtd_rooms += rm
                mtd_rev += rev
                if o > 0:
                    total_sellable += (rm / (o / 100.0))
        
        mtd_occ = (mtd_rooms / total_sellable * 100.0) if total_sellable > 0 else 0.0
        mtd_adr = (mtd_rev / mtd_rooms) if mtd_rooms > 0 else 0.0
        
        # 獲取餐廳資料 (正確結算，不重複加總)
        rest_mrev = 0
        if not df_mtd.empty and 'rest_month_rev' in df_mtd.columns:
            valid_rest = df_mtd[df_mtd['rest_month_rev'] > 0]
            if not valid_rest.empty:
                rest_mrev = valid_rest.iloc[-1]['rest_month_rev']
        
        grand_total_rev = mtd_rev + rest_mrev
        
        # 顯示四大指標
        st.write("##### 🏨 房務營運 MTD")
        c1, c2, c3 = st.columns(3)
        c1.markdown(make_card("MTD 累計住房率", f"{mtd_occ:.1f}%", "card-theme-blue", "card-bg-dark", "🏨"), unsafe_allow_html=True)
        c2.markdown(make_card("MTD 累計 ADR", f"NT$ {int(mtd_adr):,}", "card-theme-green", "card-bg-dark", "💳"), unsafe_allow_html=True)
        c3.markdown(make_card("MTD 房務累計營收", f"NT$ {int(mtd_rev):,}", "card-theme-orange", "card-bg-dark", "💰"), unsafe_allow_html=True)
        
        st.write("##### 🏁 全館合併營收 (MTD)")
        g1, g2 = st.columns([1, 2])
        g1.markdown(make_card("餐廳結算營收", f"NT$ {int(rest_mrev):,}", "card-theme-purple", "card-bg-dark", "🍽️"), unsafe_allow_html=True)
        g2.markdown(make_card("✨ 全館 MTD 總營收", f"NT$ {int(grand_total_rev):,}", "card-theme-red", "card-bg-dark", "🚀"), unsafe_allow_html=True)
        
        st.markdown("<br><hr style='margin: 5px 0; border: 1px dashed #ddd;'>", unsafe_allow_html=True)
        st.write("##### 🍽️ 餐廳營運累計 (MTD)")
        
        # MTD 餐廳計算
        mtd_bf_theme = df_mtd['bf_theme_act'].sum() if 'bf_theme_act' in df_mtd.columns else 0
        mtd_bf_zq = df_mtd['bf_zq_act'].sum() if 'bf_zq_act' in df_mtd.columns else 0
        mtd_af_theme = df_mtd['af_theme_act'].sum() if 'af_theme_act' in df_mtd.columns else 0
        mtd_af_zq = df_mtd['af_zq_act'].sum() if 'af_zq_act' in df_mtd.columns else 0
        
        # 本月整體總和
        mtd_total_bf_act = df_mtd['bf_total_act'].sum() if 'bf_total_act' in df_mtd.columns else 0
        mtd_total_af_act = df_mtd['af_total_act'].sum() if 'af_total_act' in df_mtd.columns else 0
        
        # 為了更精確，僅採計「有預估客數」或「有實際客數」的日子為工作日（這會完美略過月底那些全是 0 的未來天數）
        if 'bf_total_act' in df_mtd.columns:
            active_bf_days = len(df_mtd[(df_mtd['bf_total_est'] > 0) | (df_mtd['bf_total_act'] > 0)])
        else:
            active_bf_days = 0
        
        if 'af_total_act' in df_mtd.columns:
            active_af_days = len(df_mtd[(df_mtd['af_total_est'] > 0) | (df_mtd['af_total_act'] > 0)])
        else:
            active_af_days = 0
            
        total_bf_days = active_bf_days if active_bf_days > 0 else 1
        total_af_days = active_af_days if active_af_days > 0 else 1
        
        mtd_avg_bf = mtd_total_bf_act / total_bf_days
        mtd_avg_af = mtd_total_af_act / total_af_days
        mtd_avg_total = mtd_avg_bf + mtd_avg_af
        
        # 獲取餐廳月度總結
        # 改用最後一筆有值的記錄作為結算值，通常比較準確 (假設報表是累計生成的)
        rest_month_rev = rest_mrev # 前面已計算過
        rest_avg_spent = 0
        if not df_mtd.empty and 'rest_avg_spent' in df_mtd.columns:
            valid_aspent = df_mtd[df_mtd['rest_avg_spent'] > 0]
            if not valid_aspent.empty:
                rest_avg_spent = valid_aspent.iloc[-1]['rest_avg_spent']
        
        st.markdown("<h6 style='color:#555; margin-top:15px;'>📌【站前館】MTD 累計</h6>", unsafe_allow_html=True)
        sz1, sz2, sz3 = st.columns(3)
        sz1.markdown(make_card("早餐 (實際)", f"{int(mtd_bf_zq)} 人", "card-theme-orange", "", "🥐"), unsafe_allow_html=True)
        sz2.markdown(make_card("下午茶 (實際)", f"{int(mtd_af_zq)} 人", "card-theme-purple", "", "🍰"), unsafe_allow_html=True)
        sz3.markdown(make_card("站前合計 (實際)", f"{int(mtd_bf_zq + mtd_af_zq)} 人", "card-theme-blue", "", "👥"), unsafe_allow_html=True)

        st.markdown("<h6 style='color:#555; margin-top:20px;'>📌【主題館】MTD 累計</h6>", unsafe_allow_html=True)
        st1, st2, st3 = st.columns(3)
        st1.markdown(make_card("早餐 (實際)", f"{int(mtd_bf_theme)} 人", "card-theme-orange", "", "🥐"), unsafe_allow_html=True)
        st2.markdown(make_card("下午茶 (實際)", f"{int(mtd_af_theme)} 人", "card-theme-purple", "", "🍰"), unsafe_allow_html=True)
        st3.markdown(make_card("主題合計 (實際)", f"{int(mtd_bf_theme + mtd_af_theme)} 人", "card-theme-blue", "", "👥"), unsafe_allow_html=True)
        
        st.markdown("<h6 style='color:#555; margin-top:20px;'>👑【兩館合併總覽】</h6>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.markdown(make_card("兩館早餐 (實際)", f"{int(mtd_total_bf_act)} 人", "card-theme-orange", "card-bg-dark", "🥐"), unsafe_allow_html=True)
        m2.markdown(make_card("兩館下午茶 (實際)", f"{int(mtd_total_af_act)} 人", "card-theme-purple", "card-bg-dark", "🍰"), unsafe_allow_html=True)
        m3.markdown(make_card("全月結算營收", f"NT$ {int(rest_month_rev):,}", "card-theme-green", "card-bg-dark", "💰"), unsafe_allow_html=True)
        m4.markdown(make_card("平均客單價", f"NT$ {int(rest_avg_spent):,}", "card-theme-red", "card-bg-dark", "🧾"), unsafe_allow_html=True)

        st.markdown("<h6 style='color:#555; margin-top:20px;'>📉【兩館日平均來客】</h6>", unsafe_allow_html=True)
        a1, a2, a3 = st.columns(3)
        a1.markdown(make_card("兩館早餐平均", f"{mtd_avg_bf:.1f} 人/日", "card-theme-orange", "", "✨"), unsafe_allow_html=True)
        a2.markdown(make_card("兩館下午茶平均", f"{mtd_avg_af:.1f} 人/日", "card-theme-purple", "", "✨"), unsafe_allow_html=True)
        a3.markdown(make_card("兩館整體總平均", f"{mtd_avg_total:.1f} 人/日", "card-theme-blue", "", "📈"), unsafe_allow_html=True)
        
    else:
        st.info("💡 資料庫中目前尚未有這個月的記錄。")

with tab_m:
    st.header("📈 月分析專區")
    
    # 1. 取得四個月數據 (M-2, M-1, M, M+1)
    prev_prev_m_date = get_month_delta(selected_date, -2)
    prev_m_date = get_month_delta(selected_date, -1)
    next_m_date = get_month_delta(selected_date, 1)
    
    m_prev_prev = fetch_month_summary(prev_prev_m_date.year, prev_prev_m_date.month)
    m_prev = fetch_month_summary(prev_m_date.year, prev_m_date.month)
    m_curr = fetch_month_summary(selected_date.year, selected_date.month)
    m_next = fetch_month_summary(next_m_date.year, next_m_date.month)
    
    # 取得去年同月數據 (YoY)
    m_curr_ly = fetch_month_summary(selected_date.year - 1, selected_date.month)
    
    st.markdown("#### 🏆 本月總覽與去年同期比較 (YoY)")
    if not m_curr['df'].empty and not m_curr_ly['df'].empty:
        col1, col2, col3 = st.columns(3)
        
        adr_diff = m_curr['avg_adr'] - m_curr_ly['avg_adr']
        adr_pct = (adr_diff / m_curr_ly['avg_adr'] * 100) if m_curr_ly['avg_adr'] > 0 else 0
        adr_color = "#2ecc71" if adr_diff >= 0 else "#e74c3c"
        adr_sign = "+" if adr_diff >= 0 else ""
        col1.markdown(f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {adr_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>當月平均 ADR</p><strong style='font-size:22px;'>NT$ {int(m_curr['avg_adr']):,}</strong><p style='margin:5px 0 0 0; font-size:13px; color:{adr_color}; font-weight:bold;'>較去年同期 {adr_sign}NT$ {int(adr_diff):,} ({adr_sign}{adr_pct:.1f}%)</p></div>", unsafe_allow_html=True)
        
        occ_diff = m_curr['avg_occ'] - m_curr_ly['avg_occ']
        occ_color = "#2ecc71" if occ_diff >= 0 else "#e74c3c"
        occ_sign = "+" if occ_diff >= 0 else ""
        col2.markdown(f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {occ_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>當月平均 OCC</p><strong style='font-size:22px;'>{m_curr['avg_occ']:.1f}%</strong><p style='margin:5px 0 0 0; font-size:13px; color:{occ_color}; font-weight:bold;'>較去年同期 {occ_sign}{occ_diff:.1f}%</p></div>", unsafe_allow_html=True)
        
        rev_diff = m_curr['rev'] - m_curr_ly['rev']
        rev_pct = (rev_diff / m_curr_ly['rev'] * 100) if m_curr_ly['rev'] > 0 else 0
        rev_color = "#2ecc71" if rev_diff >= 0 else "#e74c3c"
        rev_sign = "+" if rev_diff >= 0 else ""
        col3.markdown(f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {rev_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>當月總客房營收</p><strong style='font-size:22px;'>NT$ {int(m_curr['rev']):,}</strong><p style='margin:5px 0 0 0; font-size:13px; color:{rev_color}; font-weight:bold;'>較去年同期 {rev_sign}NT$ {int(rev_diff):,} ({rev_sign}{rev_pct:.1f}%)</p></div>", unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom:20px;'></div>", unsafe_allow_html=True)
    else:
        if m_curr['df'].empty:
            st.info("💡 本月尚無數據，無法與去年同期比較。")
        elif m_curr_ly['df'].empty:
            st.info("💡 去年同月尚無歷史對比資料。")

    # 取得台北重大活動資料
    taipei_events_df = fetch_taipei_events()
    
    # --- A. 每日住房率概況 (四個月對比) ---
    st.subheader("📊 每日住房率概況比較 (四個月)")
    col_chart1, col_chart2, col_chart3, col_chart4 = st.columns(4)
    
    def render_occ_chart(month_data, title_suffix):
        df = month_data['df'].copy()
        if df.empty:
            st.info(f"💡 {month_data['month_label']} 尚無數據。")
            return
            
        # 預先新增全月平均 ADR 基準線欄位，與主資料集完美共用同一資料來源以解決尺度分裂問題
        avg_adr = month_data.get('avg_adr', 0)
        df['adr_baseline'] = avg_adr
        df['adr_baseline_text'] = ''
        
        y_adr, y_pure_adr = fetch_yearly_metrics(int(month_data['month_label'].split('-')[0]))
        df['y_adr_baseline'] = y_adr
        df['y_adr_text'] = ''
        df['y_pure_adr_baseline'] = y_pure_adr
        df['y_pure_adr_text'] = ''

        if not df.empty:
            if avg_adr > 0: df.loc[df.index[-1], 'adr_baseline_text'] = f"${int(avg_adr):,} (月)"
            if y_adr > 0: df.loc[df.index[-1], 'y_adr_text'] = f"${int(y_adr):,} (年)"
            if y_pure_adr > 0: df.loc[df.index[-1], 'y_pure_adr_text'] = f"${int(y_pure_adr):,} (純平)"
            
        dt = pd.to_datetime(df['date'])
        df['day'] = dt.dt.day
        weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
        df['weekday'] = dt.dt.weekday.map(weekday_map)
        df['label'] = df['day'].astype(str) + " (" + df['weekday'] + ")"
        
        df['color_category'] = df['occ_rate'].apply(lambda x: '>=90' if x >= 90.0 else ('>=80' if x >= 80.0 else '<80'))
        
        if not df.empty:
            y_str, m_str = df['date'].iloc[0].split('-')[:2]
            holidays_dict = fetch_holidays_for_month(int(y_str), int(m_str))
            
            # 合併假日與台北活動標籤
            def get_combined_flags_list(d_str):
                import re
                h_f_str = holidays_dict.get(d_str, {}).get('flags', '')
                h_flags = re.findall(r'\[.*?\]|🌍', h_f_str)
                
                e_flags = []
                if not taipei_events_df.empty:
                    day_events = taipei_events_df[taipei_events_df['date'] == d_str]
                    for _, row in day_events.iterrows():
                        e_label = EVENT_TYPE_LABELS.get(row['event_type'], '[活]')
                        if e_label not in e_flags:
                            e_flags.append(e_label)
                return h_flags + e_flags

            # 建立多層標籤資料 (最多支援 5 層垂直堆疊，避免過度擁擠)
            for i in range(5):
                df[f'flag_{i}'] = df['date'].apply(lambda d: get_combined_flags_list(d)[i] if len(get_combined_flags_list(d)) > i else '')
        else:
            for i in range(5):
                df[f'flag_{i}'] = ''
            
        # ==========================================
        # 1. 建立 OCC 子圖 (長條圖 + 住房百分比文字標籤 + 活動/節慶)
        # ==========================================
        base_occ = alt.Chart(df).encode(
            x=alt.X('label:O', 
                    title='日期', 
                    sort=df['label'].tolist(),
                    axis=alt.Axis(labelAngle=0)),
            tooltip=['date', 'occ_rate', 'adr']
        )
        
        bars = base_occ.mark_bar().encode(
            y=alt.Y('occ_rate:Q', title='住房率 (%)', scale=alt.Scale(domain=[0, 100])),
            color=alt.Color(
                'color_category:N', 
                scale=alt.Scale(
                    domain=['>=90', '>=80', '<80'], 
                    range=['#e74c3c', '#3498db', '#2ecc71']
                ),
                legend=None
            )
        )
        
        # 住房率文字標籤 (自然繼承 OCC 軸，不需畫新軸)
        text = base_occ.mark_text(
            align='center',
            baseline='bottom',
            dy=-5,
            fontSize=14,
            fontWeight='bold'
        ).encode(
            y='occ_rate:Q',
            text=alt.Text('occ_rate:Q', format='.1f')
        )

        # 建立多層垂直標籤
        occ_layers = [bars, text]
        for i in range(5):
            offset = -22 - (i * 13)
            f_layer = base_occ.mark_text(
                align='center',
                baseline='bottom',
                dy=offset,
                fontSize=11,
                fontWeight='normal'
            ).encode(
                y='occ_rate:Q',
                text=f'flag_{i}:N'
            )
            occ_layers.append(f_layer)
            
        occ_chart = alt.layer(*occ_layers)
        
        # 計算當月 ADR 的邊界，鎖定統一的 Y 軸比例尺 domain，消除 Altair 多資料來源尺度獨立導致的錯位 Bug
        valid_adrs = df[df['adr'] > 0]['adr'] if not df.empty else pd.Series([])
        if not valid_adrs.empty:
            adr_min = max(0, int(valid_adrs.min() * 0.9))
            adr_max = int(valid_adrs.max() * 1.1)
        else:
            adr_min = 2000
            adr_max = 8000
            
        avg_adr = month_data.get('avg_adr', 0)
        y_adr, y_pure_adr = fetch_yearly_metrics(int(month_data['month_label'].split('-')[0]))
        
        if avg_adr > 0:
            adr_min = min(adr_min, int(avg_adr * 0.9))
            adr_max = max(adr_max, int(avg_adr * 1.1))
        if y_adr > 0:
            adr_min = min(adr_min, int(y_adr * 0.9))
            adr_max = max(adr_max, int(y_adr * 1.1))
        if y_pure_adr > 0:
            adr_min = min(adr_min, int(y_pure_adr * 0.9))
            adr_max = max(adr_max, int(y_pure_adr * 1.1))
            
        adr_scale = alt.Scale(domain=[adr_min, adr_max], zero=False)
        
        # ==========================================
        # 2. 建立 ADR 子圖 (折線圖 + 資料點 + 紅色平均房價基準線 + 紅色金額數值標記)
        # ==========================================
        base_adr = alt.Chart(df).encode(
            x=alt.X('label:O', sort=df['label'].tolist()), # 自然與 OCC X 軸合併
            tooltip=['date', 'occ_rate', 'adr']
        )
        
        adr_line = base_adr.mark_line(color='#ff9f43', strokeWidth=3, interpolate='monotone').encode(
            y=alt.Y('adr:Q', title='平均房價 (NT$)', axis=alt.Axis(titleColor='#ff9f43', format='$,.0f'), scale=adr_scale)
        )
        adr_points = base_adr.mark_circle(color='black', size=100, stroke='white', strokeWidth=1.5).encode(
            y=alt.Y('adr:Q', scale=adr_scale)
        )
        
        adr_layers = [adr_line, adr_points]
        
        # 繪製全月平均 ADR 紅色基準線與右側數值標記
        if avg_adr > 0:
            # 建立水平紅色虛線 (共用相同 df 解決尺度獨立 bug，且因不含 X 編碼故保證水平)
            baseline_rule = alt.Chart(df).mark_rule(
                color='#e74c3c', 
                strokeWidth=1.5, 
                strokeDash=[5, 5]
            ).encode(
                y=alt.Y('adr_baseline:Q', scale=adr_scale)
            )
            
            # 建立紅色文字標籤 (共用相同 df，只在最後一天繪製文字，完美對齊)
            baseline_text = alt.Chart(df).mark_text(
                align='right',     # 改為靠右對齊，讓文字往圖表內部 (左側) 延伸
                baseline='middle', 
                dx=-8,             # 向左偏移 8 像素，避免與外側的年 ADR 數值重疊
                color='#000000',
                fontSize=12,
                fontWeight='bold'
            ).encode(
                x=alt.X('label:O', sort=df['label'].tolist()),
                y=alt.Y('adr_baseline:Q', scale=adr_scale),
                text='text:N' if 'text' in df.columns else 'adr_baseline_text:N'
            )
            adr_layers.extend([baseline_rule, baseline_text])
            
        # 繪製年 ADR 黃色基準線
        if df.get('y_adr_baseline', pd.Series()).max() > 0:
            y_adr_rule = alt.Chart(df).mark_rule(color='#f1c40f', strokeWidth=1.5, strokeDash=[5, 5]).encode(y=alt.Y('y_adr_baseline:Q', scale=adr_scale))
            y_adr_text = alt.Chart(df).mark_text(
                align='left', baseline='middle', dx=8, dy=-14, color='#000000', fontSize=11, fontWeight='bold'
            ).encode(
                x=alt.X('label:O', sort=df['label'].tolist()), y=alt.Y('y_adr_baseline:Q', scale=adr_scale), text='y_adr_text:N'
            )
            adr_layers.extend([y_adr_rule, y_adr_text])
            
        # 繪製年純平日 ADR 黑色基準線
        if df.get('y_pure_adr_baseline', pd.Series()).max() > 0:
            yp_adr_rule = alt.Chart(df).mark_rule(color='#000000', strokeWidth=1.5, strokeDash=[5, 5]).encode(y=alt.Y('y_pure_adr_baseline:Q', scale=adr_scale))
            yp_adr_text = alt.Chart(df).mark_text(align='left', baseline='middle', dx=8, dy=14, color='#000000', fontSize=11, fontWeight='bold').encode(
                x=alt.X('label:O', sort=df['label'].tolist()), y=alt.Y('y_pure_adr_baseline:Q', scale=adr_scale), text='y_pure_adr_text:N'
            )
            adr_layers.extend([yp_adr_rule, yp_adr_text])
            
        adr_chart = alt.layer(*adr_layers)
        
        # ==========================================
        # 3. 結合兩個子圖，宣告 Y 軸為獨立雙軸，實現完美對齊
        # ==========================================
        chart = alt.layer(occ_chart, adr_chart).resolve_scale(
            y='independent'
        ).properties(title=f"{month_data['month_label']} {title_suffix}", height=400)
        
        st.altair_chart(chart, use_container_width=True)

    with col_chart1: render_occ_chart(m_prev_prev, "(前前月)")
    with col_chart2: render_occ_chart(m_prev, "(上月)")
    with col_chart3: render_occ_chart(m_curr, "(本月)")
    with col_chart4: render_occ_chart(m_next, "(下月)")
    
    # --- A2. 去年同期軌跡對比 (YoY Daily Comparison) ---
    st.markdown("#### 📅 去年同期軌跡對比 (YoY Daily Comparison)")
    if not m_curr['df'].empty and not m_curr_ly['df'].empty:
        df_ty = m_curr['df'].copy()
        df_ly = m_curr_ly['df'].copy()
        
        if 'day' not in df_ty.columns: df_ty['day'] = pd.to_datetime(df_ty['date']).dt.day
        if 'day' not in df_ly.columns: df_ly['day'] = pd.to_datetime(df_ly['date']).dt.day
            
        df_ty['year'] = '今年'
        df_ly['year'] = '去年'
        
        df_yoy = pd.concat([df_ty[['day', 'adr', 'year']], df_ly[['day', 'adr', 'year']]], ignore_index=True)
        df_yoy['adr'] = pd.to_numeric(df_yoy['adr'], errors='coerce').fillna(0)
        
        # 設定 Y 軸比例尺
        yoy_adr_min = max(0, int(df_yoy['adr'].min() * 0.9))
        yoy_adr_max = int(df_yoy['adr'].max() * 1.1)
        if yoy_adr_min == yoy_adr_max: yoy_adr_max += 1000
        
        yoy_chart = alt.Chart(df_yoy).mark_line(point=True, strokeWidth=3).encode(
            x=alt.X('day:O', title='日期 (Day of Month)'),
            y=alt.Y('adr:Q', title='平均房價 (NT$)', scale=alt.Scale(domain=[yoy_adr_min, yoy_adr_max], zero=False)),
            color=alt.Color('year:N', 
                scale=alt.Scale(domain=['今年', '去年'], range=['#ff9f43', '#bdc3c7']),
                legend=alt.Legend(title="年份", orient="top-left")
            ),
            strokeDash=alt.condition(alt.datum.year == '去年', alt.value([5, 5]), alt.value([0])),
            tooltip=['day', 'year', 'adr']
        ).properties(height=350)
        
        st.altair_chart(yoy_chart, use_container_width=True)
    
    st.markdown("<div style='margin-bottom:30px;'></div>", unsafe_allow_html=True)
    
    # --- B. 關鍵表現數據分析 ---
    st.markdown("#### 🌟 關鍵表現數據分析")
    
    def calc_key_metrics(m_data):
        df = m_data.get('df', pd.DataFrame())
        res = {'high_adr_days': 0, 'top20_rev_avg': 0, 'bot20_rev_avg': 0, 'dual_match_days': 0, 'month_label': m_data.get('month_label', '')}
        if df is None or df.empty: return res
        
        avg_adr = m_data.get('avg_adr', 0)
        
        # 確保數值正確
        df['adr_val'] = pd.to_numeric(df['adr'], errors='coerce').fillna(0)
        df['rev_val'] = pd.to_numeric(df['revenue'], errors='coerce').fillna(0)
        
        # 高於當月平均 ADR 天數
        df['is_high_adr'] = df['adr_val'] > avg_adr
        res['high_adr_days'] = int(df['is_high_adr'].sum())
        
        # 八二法則 (前 20% 與後 20%)
        n_days = len(df)
        n_top = max(1, int(round(n_days * 0.2)))
        
        df_sorted = df.sort_values('rev_val', ascending=False)
        top20_df = df_sorted.head(n_top)
        bot20_df = df_sorted.tail(n_top)
        
        res['top20_rev_avg'] = top20_df['rev_val'].mean() if not top20_df.empty else 0
        res['bot20_rev_avg'] = bot20_df['rev_val'].mean() if not bot20_df.empty else 0
        
        # 雙冠天數：前 20% 營收日中，ADR 也大於當月平均 ADR 的天數
        dual_match_df = top20_df[top20_df['is_high_adr']]
        res['dual_match_days'] = int(len(dual_match_df))
        res['dual_match_dates'] = dual_match_df['date'].sort_values().tolist() if not dual_match_df.empty else []
        
        return res

    curr_metrics = calc_key_metrics(m_curr)
    prev_metrics = calc_key_metrics(m_prev)
    pprev_metrics = calc_key_metrics(m_prev_prev)
    next_metrics = calc_key_metrics(m_next)
    
    def metric_diff_card(label, diff, target_label, unit="天"):
        color = '#2ecc71' if diff >= 0 else '#e74c3c'
        status = '本月多' if diff > 0 else '較少' if diff < 0 else '持平'
        return f'<div style="flex: 1; min-width: 150px; background: #fff; padding: 12px; border-radius: 8px; border: 1px solid #eee; margin-bottom: 10px;"><p style="margin:0; font-size:12px; color:#999;">與 {target_label} 相比</p><div style="display: flex; align-items: baseline; gap: 8px; margin-top: 5px;"><strong style="font-size:18px; color:{color};">{abs(diff)} {unit}</strong><span style="font-size:11px; color:#666;">({status})</span></div></div>'

    # 天數差異
    diff_adr_pp = curr_metrics['high_adr_days'] - pprev_metrics['high_adr_days']
    diff_adr_p = curr_metrics['high_adr_days'] - prev_metrics['high_adr_days']
    diff_adr_n = curr_metrics['high_adr_days'] - next_metrics['high_adr_days']
    
    diff_dual_pp = curr_metrics['dual_match_days'] - pprev_metrics['dual_match_days']
    diff_dual_p = curr_metrics['dual_match_days'] - prev_metrics['dual_match_days']
    diff_dual_n = curr_metrics['dual_match_days'] - next_metrics['dual_match_days']

    kp_col1, kp_col2 = st.columns([1.5, 1])
    
    with kp_col1:
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #3498db; margin-bottom: 20px;">
            <p style="margin:0; font-size:14px; color:#666;">📈 <strong>高於當月平均 ADR 天數 (本月: {curr_metrics['high_adr_days']} 天)</strong></p>
            <div style="display: flex; gap: 15px; margin-top: 10px; flex-wrap: wrap;">
                {metric_diff_card("前前月", diff_adr_pp, pprev_metrics['month_label'])}
                {metric_diff_card("上月", diff_adr_p, prev_metrics['month_label'])}
                {metric_diff_card("下月預期", diff_adr_n, next_metrics['month_label'])}
            </div>
            <p style="margin:15px 0 0 0; font-size:14px; color:#666;">🏆 <strong>雙冠天數：前 20% 營收且高 ADR (本月: {curr_metrics['dual_match_days']} 天)</strong></p>
            <div style="display: flex; gap: 15px; margin-top: 10px; flex-wrap: wrap;">
                {metric_diff_card("前前月", diff_dual_pp, pprev_metrics['month_label'])}
                {metric_diff_card("上月", diff_dual_p, prev_metrics['month_label'])}
                {metric_diff_card("下月預期", diff_dual_n, next_metrics['month_label'])}
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with kp_col2:
        st.markdown(f"""
        <div style="background: #fffcf5; padding: 15px; border-radius: 10px; border-left: 5px solid #f39c12; margin-bottom: 20px; height: 100%;">
            <p style="margin:0; font-size:14px; color:#666;">📊 <strong>八二法則：極端值分析 (本月)</strong></p>
            <div style="margin-top: 20px;">
                <p style="margin:0; font-size:13px; color:#999;">🔥 前 20% 營收日 (Top 20%) 平均營收</p>
                <h3 style="margin: 5px 0 15px 0; color: #d35400;">NT$ {int(curr_metrics['top20_rev_avg']):,}</h3>
                <p style="margin:0; font-size:13px; color:#999;">❄️ 後 20% 營收日 (Bottom 20%) 平均營收</p>
                <h3 style="margin: 5px 0 15px 0; color: #7f8c8d;">NT$ {int(curr_metrics['bot20_rev_avg']):,}</h3>
                <hr style="border: 0; border-top: 1px dashed #eee; margin: 15px 0;">
                <p style="margin:0; font-size:12px; color:#888;">💡 <strong>解讀</strong>：當前後 20% 的平均營收差距擴大時，代表淡旺日的業績差距大，可針對淡日加強促銷。</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # --- B3. OCC vs ADR 四象限定價診斷圖 ---
    st.markdown("#### 🎯 定價水位診斷：住房率 vs 平均房價 四象限分析（以年純平 ADR 為底線基準）")
    scatter_df = m_curr['df'].copy()
    if not scatter_df.empty:
        scatter_df['occ_val'] = pd.to_numeric(scatter_df['occ_rate'], errors='coerce').fillna(0)
        scatter_df['adr_val'] = pd.to_numeric(scatter_df['adr'], errors='coerce').fillna(0)
        scatter_df['day'] = pd.to_datetime(scatter_df['date']).dt.day

        # 以「年純平 ADR」作為 Y 軸分界（最客觀的裸實力底線，不受淡日拉低）
        y_adr_s, y_pure_adr_s = fetch_yearly_metrics(selected_date.year)
        adr_anchor = y_pure_adr_s if y_pure_adr_s > 0 else m_curr.get('avg_adr', scatter_df['adr_val'].mean())
        anchor_label = f'年純平 ADR ${int(adr_anchor):,}'
        anchor_color = '#000000'
        occ_threshold = 75.0  # 高住房率門檻

        def classify_quadrant(row):
            hi_occ = row['occ_val'] >= occ_threshold
            hi_adr = row['adr_val'] >= adr_anchor
            if hi_occ and hi_adr: return '🟠 理想（高OCC+高ADR）'
            if hi_occ and not hi_adr: return '🔴 賤賣（高OCC+低ADR）'
            if not hi_occ and hi_adr: return '🟡 定價偏高（低OCC+高ADR）'
            return '🔵 淡季死水（低OCC+低ADR）'

        scatter_df['象限'] = scatter_df.apply(classify_quadrant, axis=1)

        color_map = {
            '🟠 理想（高OCC+高ADR）': '#ff9f43',
            '🔴 賤賣（高OCC+低ADR）': '#e74c3c',
            '🟡 定價偏高（低OCC+高ADR）': '#f1c40f',
            '🔵 淡季死水（低OCC+低ADR）': '#3498db',
        }

        scatter_chart = alt.Chart(scatter_df).mark_circle(size=100, opacity=0.8).encode(
            x=alt.X('occ_val:Q', title='住房率 (%)', scale=alt.Scale(domain=[0, 105])),
            y=alt.Y('adr_val:Q', title='平均房價 ADR (NT$)', scale=alt.Scale(zero=False)),
            color=alt.Color('象限:N',
                scale=alt.Scale(
                    domain=list(color_map.keys()),
                    range=list(color_map.values())
                ),
                legend=alt.Legend(title="象限分類", orient="bottom", columns=2)
            ),
            tooltip=[
                alt.Tooltip('date:N', title='日期'),
                alt.Tooltip('occ_val:Q', title='住房率 (%)', format='.1f'),
                alt.Tooltip('adr_val:Q', title='ADR (NT$)', format=',.0f'),
                alt.Tooltip('象限:N', title='象限'),
            ]
        )

        # 年純平 ADR 水平輔助線
        adr_rule = alt.Chart(pd.DataFrame({'y': [adr_anchor]})).mark_rule(
            color=anchor_color, strokeDash=[6, 3], strokeWidth=2
        ).encode(y='y:Q')
        adr_label = alt.Chart(pd.DataFrame({'y': [adr_anchor], 'x': [105], 'text': [anchor_label]})).mark_text(
            align='right', dx=-4, dy=-8, color=anchor_color, fontSize=11, fontWeight='bold'
        ).encode(x='x:Q', y='y:Q', text='text:N')

        # 75% OCC 垂直輔助線
        occ_rule = alt.Chart(pd.DataFrame({'x': [occ_threshold]})).mark_rule(
            color='#7f8c8d', strokeDash=[6, 3], strokeWidth=1.5
        ).encode(x='x:Q')
        occ_label = alt.Chart(pd.DataFrame({'x': [occ_threshold], 'y': [scatter_df['adr_val'].max() * 1.05], 'text': ['75% OCC 門檻']})).mark_text(
            align='left', dx=4, color='#7f8c8d', fontSize=11, fontWeight='bold'
        ).encode(x='x:Q', y='y:Q', text='text:N')

        final_chart = alt.layer(scatter_chart, adr_rule, adr_label, occ_rule, occ_label).properties(
            height=380,
            title=f"{m_curr['month_label']} 每日定價水位診斷（每個點代表一天，以年純平 ADR 為底線）"
        )
        st.altair_chart(final_chart, use_container_width=True)

        # 各象限天數摘要
        q_counts = scatter_df['象限'].value_counts()
        q_cols = st.columns(4)
        for i, (q_name, color) in enumerate(color_map.items()):
            cnt = q_counts.get(q_name, 0)
            q_cols[i].markdown(
                f"<div style='background:{color}22; border-left:4px solid {color}; padding:10px; border-radius:6px; text-align:center;'>"
                f"<p style='margin:0; font-size:12px; color:#555;'>{q_name}</p>"
                f"<strong style='font-size:22px;'>{cnt} 天</strong></div>",
                unsafe_allow_html=True
            )
        st.write("")

        # --- 定價成功率 (Pricing Success Rate) ---
        ideal_cnt = q_counts.get('🟠 理想（高OCC+高ADR）', 0)
        cheap_cnt = q_counts.get('🔴 賤賣（高OCC+低ADR）', 0)
        high_occ_total = ideal_cnt + cheap_cnt
        success_rate = (ideal_cnt / high_occ_total * 100) if high_occ_total > 0 else 0
        
        # 計算上個月的定價成功率作為對比
        prev_scatter_df = m_prev['df'].copy()
        prev_success_rate = 0
        if not prev_scatter_df.empty:
            prev_scatter_df['occ_val'] = pd.to_numeric(prev_scatter_df['occ_rate'], errors='coerce').fillna(0)
            prev_scatter_df['adr_val'] = pd.to_numeric(prev_scatter_df['adr'], errors='coerce').fillna(0)
            prev_scatter_df['hi_occ'] = prev_scatter_df['occ_val'] >= occ_threshold
            prev_scatter_df['hi_adr'] = prev_scatter_df['adr_val'] >= adr_anchor
            prev_ideal = int((prev_scatter_df['hi_occ'] & prev_scatter_df['hi_adr']).sum())
            prev_cheap = int((prev_scatter_df['hi_occ'] & ~prev_scatter_df['hi_adr']).sum())
            prev_total = prev_ideal + prev_cheap
            prev_success_rate = (prev_ideal / prev_total * 100) if prev_total > 0 else 0
        
        rate_diff = success_rate - prev_success_rate
        rate_color = '#2ecc71' if rate_diff >= 0 else '#e74c3c'
        rate_sign = '+' if rate_diff >= 0 else ''
        
        if success_rate >= 80:
            bar_color = '#2ecc71'
            verdict = '🟢 定價能力優秀'
        elif success_rate >= 60:
            bar_color = '#f39c12'
            verdict = '🟡 定價能力尚可'
        else:
            bar_color = '#e74c3c'
            verdict = '🔴 定價能力待改善'
            
        st.markdown(f"""
        <div style="background:#f8f9fa; border-radius:10px; padding:20px; margin-top:10px; border-left: 5px solid {bar_color};">
            <p style="margin:0 0 8px 0; font-size:14px; color:#555;">
                📐 <strong>高住房日定價成功率</strong>
                <span style="font-size:12px; color:#aaa; margin-left:8px;">高OCC 天數共 {high_occ_total} 天，其中 {int(ideal_cnt)} 天 ADR 超過年純平基準</span>
            </p>
            <div style="display:flex; align-items:baseline; gap:15px; flex-wrap:wrap;">
                <strong style="font-size:40px; color:{bar_color};">{success_rate:.1f}%</strong>
                <span style="font-size:14px;">{verdict}</span>
                <span style="font-size:14px; color:{rate_color}; font-weight:bold;">vs 上月 {prev_success_rate:.1f}% ({rate_sign}{rate_diff:.1f}%)</span>
            </div>
            <div style="background:#e0e0e0; border-radius:999px; height:10px; margin-top:10px;">
                <div style="background:{bar_color}; width:{min(success_rate, 100):.1f}%; height:10px; border-radius:999px; transition: width 0.5s;"></div>
            </div>
            <p style="margin:8px 0 0 0; font-size:12px; color:#888;">💡 目標：讓「賤賣天數」每月減少 1-2 天，持續將成功率推向 80%</p>
        </div>
        """, unsafe_allow_html=True)
        st.write("")

    st.divider()

    # --- 新增：雙冠備戰行事曆 (Peak Demand Radar) ---

    if curr_metrics.get('dual_match_dates'):
        st.markdown("#### 🎯 雙冠備戰行事曆 (Peak Demand Radar)")
        st.info("💡 系統自動揪出本月符合「高營收」且「高均價」的雙冠日。請針對以下日期提早準備生鮮食材，並可適度放寬單客成本 (CPG) 以滿足高端客群期待。")
        
        radar_cols = st.columns(min(max(len(curr_metrics['dual_match_dates']), 1), 5))
        for i, d_date in enumerate(curr_metrics['dual_match_dates']):
            day_row = m_curr['df'][m_curr['df']['date'] == d_date]
            bf_count = 0
            if not day_row.empty:
                bf_col = 'bf_total_act' if 'bf_total_act' in day_row.columns and pd.to_numeric(day_row['bf_total_act'].iloc[0], errors='coerce') > 0 else 'bf_total_est'
                if bf_col in day_row.columns:
                    bf_count = pd.to_numeric(day_row[bf_col], errors='coerce').fillna(0).iloc[0]
                    
            c = radar_cols[i % 5]
            c.markdown(f"""
            <div style="background: #fff; border: 2px solid #e74c3c; border-radius: 8px; padding: 15px; text-align: center; margin-bottom: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <h4 style="margin:0; color:#e74c3c;">{d_date[5:]}</h4>
                <p style="margin:5px 0 0 0; font-size:12px; color:#666;">早餐預估: <strong>{int(bf_count)}</strong> 人</p>
            </div>
            """, unsafe_allow_html=True)
            
        st.divider()

    # --- B2. 即將到來的重大活動與假日警報 ---
    st.markdown("#### 🚨 即將到來的重大活動與假日警報 (未來 30 天)")
    upcoming_holidays = fetch_upcoming_holidays(selected_date, 30)
    
    # 合併台北重大活動至警報列表 (分開呈現假日與活動)
    combined_alerts = []
    h_map = {h['date']: h for h in upcoming_holidays}
    
    for i in range(31):
        d_obj = selected_date + datetime.timedelta(days=i)
        d_str = d_obj.strftime('%Y-%m-%d')
        
        h_info = h_map.get(d_str)
        e_list = []
        e_labels = []
        if not taipei_events_df.empty:
            day_events = taipei_events_df[taipei_events_df['date'] == d_str]
            for _, row in day_events.iterrows():
                v_suffix = f" <span style='color:#777;'>@{row['venue']}</span>" if pd.notna(row['venue']) and str(row['venue']).strip() != "" else ""
                e_list.append(f"🏟️ {row['event_name']}{v_suffix}")
                e_labels.append(EVENT_TYPE_LABELS.get(row['event_type'], '[活]'))
        
        if h_info or e_list:
            all_flags = (h_info['flags'] if h_info else "") + "".join(sorted(list(set(e_labels))))
            details_html = ""
            if h_info:
                details_html += f"<div style='margin-bottom:4px; color:#856404;'>🌍 {h_info['details']}</div>"
            if e_list:
                details_html += "<div style='color:#2c3e50;'>" + "<br>".join(e_list) + "</div>"
            
            combined_alerts.append({
                'date': d_str,
                'flags': all_flags,
                'details_html': details_html
            })

    if combined_alerts:
        alert_html = "<div style='display: flex; gap: 10px; overflow-x: auto; padding-bottom: 10px;'>"
        for h in combined_alerts:
            alert_html += f"<div style='min-width: 250px; background: #fff3cd; border-left: 4px solid #ffc107; padding: 10px; border-radius: 5px;'><strong style='color: #856404;'>{h['date']}</strong> <span style='font-size: 1.1em;'>{h['flags']}</span><br><div style='font-size: 0.85em; margin-top:5px;'>{h['details_html']}</div></div>"
        alert_html += "</div>"
        st.write(alert_html, unsafe_allow_html=True)
    else:
        st.info("未來 30 天內無重大假日或台北活動。")

    st.divider()

    # --- C. 假日與活動績效分析 ---
    st.markdown("#### 🌍 績效貢獻度交叉分析")
    curr_df = m_curr['df'].copy()
    if not curr_df.empty:
        y_str, m_str = curr_df['date'].iloc[0].split('-')[:2]
        h_dict = fetch_holidays_for_month(int(y_str), int(m_str))
        h_dates = {d for d, info in h_dict.items() if info['flags']}
        e_dates = set(taipei_events_df['date'].unique()) if not taipei_events_df.empty else set()
        
        curr_df['is_h'] = curr_df['date'].isin(h_dates)
        curr_df['is_e'] = curr_df['date'].isin(e_dates)
        curr_df['is_any'] = curr_df['is_h'] | curr_df['is_e']

        def render_impact_row(df, condition_col, title, icon):
            holiday_days = df[df[condition_col]]
            non_holiday_days = df[~df[condition_col]]
            
            h_occ = holiday_days['occ_rate'].mean() if len(holiday_days) > 0 else 0
            h_adr = holiday_days['revenue'].sum() / holiday_days['total_rooms'].sum() if len(holiday_days) > 0 and holiday_days['total_rooms'].sum() > 0 else 0
            nh_occ = non_holiday_days['occ_rate'].mean() if len(non_holiday_days) > 0 else 0
            nh_adr = non_holiday_days['revenue'].sum() / non_holiday_days['total_rooms'].sum() if len(non_holiday_days) > 0 and non_holiday_days['total_rooms'].sum() > 0 else 0
            
            diff_occ = h_occ - nh_occ
            diff_adr = h_adr - nh_adr
            
            st.markdown(f"**{icon} {title}**")
            c1, c2, c3 = st.columns(3)
            c1.markdown(f"<div style='background:#f1f8ff; padding:10px; border-radius:5px; border-left:3px solid #3498db; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>有標籤 ({len(holiday_days)}天)</p><strong style='font-size:16px;'>{h_occ:.1f}% / NT$ {int(h_adr):,}</strong></div>", unsafe_allow_html=True)
            c2.markdown(f"<div style='background:#f8f9fa; padding:10px; border-radius:5px; border-left:3px solid #ccc; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>無標籤 ({len(non_holiday_days)}天)</p><strong style='font-size:16px;'>{nh_occ:.1f}% / NT$ {int(nh_adr):,}</strong></div>", unsafe_allow_html=True)
            color = "#2ecc71" if diff_occ >= 0 else "#e74c3c"
            c3.markdown(f"<div style='background:#f0fff4; padding:10px; border-radius:5px; border-left:3px solid #2ecc71; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>帶動效益</p><strong style='font-size:16px; color:{color};'>{diff_occ:+.1f}% / NT$ {int(diff_adr):+,}</strong></div>", unsafe_allow_html=True)
            st.markdown("<div style='margin-bottom:15px;'></div>", unsafe_allow_html=True)

        def render_exclusive_matrix(df, title_suffix=""):
            st.markdown(f"**📐 四象限排他性交叉分析 {title_suffix}**")
            
            is_e = df['is_e']
            is_h = df['is_h']
            
            df_pure_weekday = df[~is_e & ~is_h]
            df_pure_event = df[is_e & ~is_h]
            df_pure_holiday = df[~is_e & is_h]
            df_double_impact = df[is_e & is_h]
            
            def get_metrics(sub_df):
                days = len(sub_df)
                if days == 0:
                    return 0, 0, days
                occ = sub_df['occ_rate'].mean()
                adr = sub_df['revenue'].sum() / sub_df['total_rooms'].sum() if sub_df['total_rooms'].sum() > 0 else 0
                return occ, adr, days
            
            occ_pw, adr_pw, days_pw = get_metrics(df_pure_weekday)
            occ_pe, adr_pe, days_pe = get_metrics(df_pure_event)
            occ_ph, adr_ph, days_ph = get_metrics(df_pure_holiday)
            occ_di, adr_di, days_di = get_metrics(df_double_impact)
            
            col1, col2, col3, col4 = st.columns(4)
            
            def format_diff(val, is_percent=False):
                if val == 0:
                    return "0.0%" if is_percent else "NT$ 0"
                sign = "+" if val > 0 else ""
                if is_percent:
                    return f"{sign}{val:.1f}%"
                else:
                    return f"{sign}NT$ {int(val):,}"

            # 象限 4: 純平日
            with col1:
                st.markdown(
                    f"<div style='background:#1e293b; padding:15px; border-radius:8px; border-left:4px solid #94a3b8; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#94a3b8; font-weight:bold;'>【象限 4】純平日 ({days_pw}天)</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>基準對照組</p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_pw:.1f}% / NT$ {int(adr_pw):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:11px; color:#64748b;'>無活動、無節慶的基準線</p>"
                    f"</div>", 
                    unsafe_allow_html=True
                )
                
            # 象限 1: 純活動日
            with col2:
                diff_occ = occ_pe - occ_pw if days_pe > 0 and days_pw > 0 else 0
                diff_adr = adr_pe - adr_pw if days_pe > 0 and days_pw > 0 else 0
                color = "#10b981" if diff_adr >= 0 else "#ef4444"
                bg_style = "background:#0f172a; border-left:4px solid #3b82f6;"
                desc = f"淨效益: <span style='color:{color}; font-weight:bold;'>{format_diff(diff_occ, True)} / {format_diff(diff_adr)}</span>" if days_pe > 0 else "無數據"
                st.markdown(
                    f"<div style='{bg_style} padding:15px; border-radius:8px; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#3b82f6; font-weight:bold;'>【象限 1】純活動日 ({days_pe}天)</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>僅台北重大活動</p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_pe:.1f}% / NT$ {int(adr_pe):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:12px; color:#cbd5e1;'>{desc}</p>"
                    f"</div>", 
                    unsafe_allow_html=True
                )
                
            # 象限 2: 純節慶日
            with col3:
                diff_occ = occ_ph - occ_pw if days_ph > 0 and days_pw > 0 else 0
                diff_adr = adr_ph - adr_pw if days_ph > 0 and days_pw > 0 else 0
                color = "#10b981" if diff_adr >= 0 else "#ef4444"
                bg_style = "background:#0f172a; border-left:4px solid #eab308;"
                desc = f"淨效益: <span style='color:{color}; font-weight:bold;'>{format_diff(diff_occ, True)} / {format_diff(diff_adr)}</span>" if days_ph > 0 else "無數據"
                st.markdown(
                    f"<div style='{bg_style} padding:15px; border-radius:8px; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#eab308; font-weight:bold;'>【象限 2】純節慶日 ({days_ph}天)</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>僅外國節慶</p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_ph:.1f}% / NT$ {int(adr_ph):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:12px; color:#cbd5e1;'>{desc}</p>"
                    f"</div>", 
                    unsafe_allow_html=True
                )
                
            # 象限 3: 黃金雙重日
            with col4:
                diff_occ = occ_di - occ_pw if days_di > 0 and days_pw > 0 else 0
                diff_adr = adr_di - adr_pw if days_di > 0 and days_pw > 0 else 0
                color = "#10b981" if diff_adr >= 0 else "#ef4444"
                bg_style = "background:#1e1b4b; border-left:4px solid #a855f7;"
                desc = f"淨效益: <span style='color:{color}; font-weight:bold;'>{format_diff(diff_occ, True)} / {format_diff(diff_adr)}</span>" if days_di > 0 else "無數據"
                st.markdown(
                    f"<div style='{bg_style} padding:15px; border-radius:8px; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#a855f7; font-weight:bold;'>【象限 3】黃金雙重日 ({days_di}天)</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>活動 ＋ 節慶疊加</p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_di:.1f}% / NT$ {int(adr_di):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:12px; color:#cbd5e1;'>{desc}</p>"
                    f"</div>", 
                    unsafe_allow_html=True
                )
            st.markdown("<div style='margin-bottom:20px;'></div>", unsafe_allow_html=True)

        render_impact_row(curr_df, 'is_any', "綜合分析 (假日 + 台北活動)", "📊")
        render_impact_row(curr_df, 'is_h', "僅外國節慶分析", "🌍")
        render_impact_row(curr_df, 'is_e', "僅台北重大活動分析", "🏟️")
        
        st.markdown("<div style='margin-bottom:20px;'></div>", unsafe_allow_html=True)
        render_exclusive_matrix(curr_df, "(當月)")

        st.divider()

        # --- C2. 過去三個月合計績效分析 (長期趨勢) ---
        st.markdown("#### ⏳ 過去三個月合計績效分析 (長期趨勢)")
        # 取得前三個月日期
        m1_date = get_month_delta(selected_date, -1)
        m2_date = get_month_delta(selected_date, -2)
        m3_date = get_month_delta(selected_date, -3)
        
        m1_sum = fetch_month_summary(m1_date.year, m1_date.month)
        m2_sum = fetch_month_summary(m2_date.year, m2_date.month)
        m3_sum = fetch_month_summary(m3_date.year, m3_date.month)
        
        hist_df = pd.concat([m1_sum['df'], m2_sum['df'], m3_sum['df']], ignore_index=True)
        
        if not hist_df.empty:
            # 準備歷史資料的標籤
            def get_hist_flags(row):
                d = row['date']
                y, m = int(d.split('-')[0]), int(d.split('-')[1])
                h_f = fetch_holidays_for_month(y, m).get(d, {}).get('flags', '')
                e_f = ""
                if not taipei_events_df.empty:
                    de = taipei_events_df[taipei_events_df['date'] == d]
                    for _, r in de.iterrows(): e_f += EVENT_TYPE_LABELS.get(r['event_type'], '[活]')
                return (h_f != ''), (e_f != '')

            # 為了效能，預先抓取這幾個月的假日資料
            hist_h_dates = set()
            for md in [m1_date, m2_date, m3_date]:
                hd = fetch_holidays_for_month(md.year, md.month)
                for d, info in hd.items():
                    if info['flags']: hist_h_dates.add(d)
            
            hist_df['is_h'] = hist_df['date'].isin(hist_h_dates)
            hist_df['is_e'] = hist_df['date'].isin(set(taipei_events_df['date'].unique())) if not taipei_events_df.empty else False
            hist_df['is_any'] = hist_df['is_h'] | hist_df['is_e']
            
            render_impact_row(hist_df, 'is_any', "綜合分析 (過去三個月合計)", "📊")
            render_impact_row(hist_df, 'is_h', "僅外國節慶分析 (過去三個月合計)", "🌍")
            render_impact_row(hist_df, 'is_e', "僅台北重大活動分析 (過去三個月合計)", "🏟️")
            
            st.markdown("<div style='margin-bottom:20px;'></div>", unsafe_allow_html=True)
            render_exclusive_matrix(hist_df, "(過去三個月合計)")
        else:
            st.info("尚無足夠的歷史數據進行長期趨勢分析。")
            
        with st.expander("📅 查看本月所有假日與台北活動詳細清單"):
            # Combine details for expander
            all_dates = sorted(set(list(h_dict.keys()) + (taipei_events_df['date'].tolist() if not taipei_events_df.empty else [])))
            has_any = False
            for d in all_dates:
                if d.startswith(f"{y_str}-{m_str}"):
                    h_info = h_dict.get(d, {'flags': '', 'details': []})
                    e_info = ""
                    if not taipei_events_df.empty:
                        de = taipei_events_df[taipei_events_df['date'] == d]
                        for _, r in de.iterrows():
                            v_suffix = f" @{r['venue']}" if pd.notna(r['venue']) and str(r['venue']).strip() != "" else ""
                            e_info += f", 🏟️ {r['event_name']}{v_suffix} ({r['event_type']})"
                    
                    if h_info['flags'] or e_info:
                        st.markdown(f"- **{d}** {h_info['flags']}{e_info}: {', '.join(h_info['details'])}")
                        has_any = True
            if not has_any:
                st.write("本月無任何重大活動或假日。")
    else:
        st.info("本月尚無營運數據可供分析。")
    
    st.divider()
    
    # --- D. 月度營運指標 (四個月對比) ---
    st.subheader("📌 月度營運指標對比")
    
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    
    def render_metric_col(month_data, label):
        st.markdown(f"<p style='text-align:center; color:#777; margin-bottom:10px;'>{label} ({month_data['month_label']})</p>", unsafe_allow_html=True)
        if not month_data['df'].empty:
            st.markdown(make_card("當月總營收", f"NT$ {int(month_data['rev']):,}", "card-theme-orange", "card-bg-dark"), unsafe_allow_html=True)
            st.markdown(make_card("當月平均房價", f"NT$ {int(month_data['avg_adr']):,}", "card-theme-green", "card-bg-dark"), unsafe_allow_html=True)
            st.markdown(make_card("當月住房率", f"{month_data['avg_occ']:.1f}%", "card-theme-blue", "card-bg-dark"), unsafe_allow_html=True)
            st.markdown(make_card("當月 RevPAR", f"NT$ {int(month_data['revpar']):,}", "card-theme-purple", "card-bg-dark"), unsafe_allow_html=True)
        else:
            st.info("暫無數據")

    with col_m1: render_metric_col(m_prev_prev, "⏪ 前前月")
    with col_m2: render_metric_col(m_prev, "◀️ 上月")
    with col_m3: render_metric_col(m_curr, "✨ 本月")
    with col_m4: render_metric_col(m_next, "▶️ 下月")
    
    # --- D. 月度營運指標 - 關鍵差異 ---
    st.markdown("#### 🔍 月度營運指標：關鍵差異對比 (本月 vs 其他月份)")
    
    def calculate_diff_row(current_val, compare_val, is_currency=True, is_percent=False):
        if compare_val == 0: return "<span style='color:#777;'>-</span>"
        diff = current_val - compare_val
        if is_currency:
            diff_str = f"{'▲' if diff >= 0 else '▼'} NT$ {abs(int(diff)):,}"
        elif is_percent:
            diff_str = f"{'▲' if diff >= 0 else '▼'} {abs(diff):.1f}%"
        else:
            diff_str = f"{'▲' if diff >= 0 else '▼'} {abs(diff):.1f}"
        
        color = "#2ecc71" if diff >= 0 else "#e74c3c" # 增加為綠色，減少為紅色
        return f"<span style='color:{color}; font-weight:bold;'>{diff_str}</span>"

    diff_table_html = f"""
    <table style="width:100%; border-collapse: collapse; margin-top: 10px; font-size: 15px;">
        <tr style="background-color: #f1f3f6; text-align: left;">
            <th style="padding: 12px; border: 1px solid #ddd;">指標項目</th>
            <th style="padding: 12px; border: 1px solid #ddd;">與前前月 ({m_prev_prev['month_label']}) 相比</th>
            <th style="padding: 12px; border: 1px solid #ddd;">與上月 ({m_prev['month_label']}) 相比</th>
            <th style="padding: 12px; border: 1px solid #ddd;">與下月 ({m_next['month_label']}) 相比</th>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">當月總營收</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['rev'], m_prev_prev['rev'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['rev'], m_prev['rev'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['rev'], m_next['rev'])}</td>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">當月平均房價 (ADR)</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_adr'], m_prev_prev['avg_adr'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_adr'], m_prev['avg_adr'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_adr'], m_next['avg_adr'])}</td>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">當月住房率 (%)</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_occ'], m_prev_prev['avg_occ'], False, True)}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_occ'], m_prev['avg_occ'], False, True)}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_occ'], m_next['avg_occ'], False, True)}</td>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">當月 RevPAR</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['revpar'], m_prev_prev['revpar'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['revpar'], m_prev['revpar'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['revpar'], m_next['revpar'])}</td>
        </tr>
    </table>
    """
    st.write(diff_table_html, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("註：RevPAR 計算方式為「當月平均住房率 × 當月平均房價」；差異對比中 ▲ 代表本月較高，▼ 代表本月較低。")

    st.divider()

    # --- 3. 達標分析指數 ---
    st.subheader("🎯 達標分析指數")
    
    # 獲取與保存目標 (針對所選月份)
    month_key = selected_date.strftime('%Y-%m')
    current_target = get_monthly_target(month_key)
    m_rev = m_curr['rev'] # 使用剛剛計算好的本月營收
    
    t_col1, t_col2 = st.columns([1, 2])
    with t_col1:
        new_target = st.number_input(f"設定 {month_key} 目標業績 (NT$)", min_value=0, step=10000, value=current_target, key=f"target_input_{month_key}")
        if new_target != current_target:
            save_monthly_target(month_key, new_target)
            st.toast(f"已更新 {month_key} 目標業績！")
            time.sleep(0.5)
            st.rerun()
    
    if new_target > 0:
        gap = new_target - m_rev
        stretch_goal = new_target * 1.1
        stretch_gap = stretch_goal - m_rev
        progress = min(1.0, m_rev / new_target)
        st.progress(progress, text=f"目標達成率: {progress*100:.1f}%")
        
        # 營收進度外推預估 (Run-Rate Forecast)
        active_days = m_curr['df'][m_curr['df']['revenue'] > 0]
        elapsed_days = len(active_days)
        
        if elapsed_days > 0:
            import calendar
            total_days = calendar.monthrange(selected_date.year, selected_date.month)[1]
            daily_avg = m_rev / elapsed_days
            projected_rev = daily_avg * total_days
            projected_progress = projected_rev / new_target
            
            # 預警顏色與文字
            status_color = "#2ecc71" if projected_rev >= new_target else "#ef4444"
            status_icon = "📈" if projected_rev >= new_target else "⚠️"
            status_text = "依目前進度，預計**可順利達標**！" if projected_rev >= new_target else "依目前進度，**達標可能有難度**，建議調整動態定價或加強促銷！"
            
            st.markdown(
                f"<div style='background: #1e293b; padding: 15px; border-radius: 8px; border-left: 5px solid {status_color}; margin-top: 10px; margin-bottom: 15px; color: #f8fafc;'>"
                f"<p style='margin:0; font-size:13px; color:#94a3b8;'>🔮 <strong>當月營收進度外推預估 (Pacing Forecast)</strong></p>"
                f"<div style='display: flex; gap: 20px; align-items: center; margin-top: 5px; flex-wrap: wrap; font-size: 13px;'>"
                f"<div>已統計天數: <strong style='color:#f1f5f9;'>{elapsed_days} / {total_days} 天</strong></div>"
                f"<div>當前日均營收: <strong style='color:#f1f5f9;'>NT$ {int(daily_avg):,}</strong></div>"
                f"<div>預估月底總營收: <strong style='color:{status_color}; font-size: 15px;'>NT$ {int(projected_rev):,}</strong></div>"
                f"<div>預估最終達成率: <strong style='color:{status_color}; font-size: 15px;'>{projected_progress*100:.1f}%</strong></div>"
                f"</div>"
                f"<p style='margin: 8px 0 0 0; font-size: 12px; color: #cbd5e1;'>{status_icon} {status_text}</p>"
                f"</div>",
                unsafe_allow_html=True
            )
        
        a_col1, a_col2, a_col3 = st.columns(3)
        if gap <= 0:
            t_card = make_card("目標達成狀況", "🎉 已達標！", "card-theme-green", "", "✅")
        else:
            t_card = make_card("距離目標還差", f"NT$ {int(gap):,}", "card-theme-red", "", "🎯")
        a_col1.markdown(t_card, unsafe_allow_html=True)
        a_col2.markdown(make_card("超標目標 (+10%)", f"NT$ {int(stretch_goal):,}", "card-theme-orange", "", "🚀"), unsafe_allow_html=True)
        if stretch_gap <= 0:
            s_card = make_card("超標達成狀況", "🔥 已超標達成！", "card-theme-green", "card-bg-dark", "🏆")
        else:
            s_card = make_card("距離超標還差", f"NT$ {int(stretch_gap):,}", "card-theme-purple", "", "⚡")
        a_col3.markdown(s_card, unsafe_allow_html=True)
    else:
        st.info("💡 請在上方輸入本月目標業績，系統將自動為您計算達標差距。")

with tab3:
    st.header("🧹 房務數據")
    st.number_input("今日總清消房數", min_value=0, step=1, key="input_cleaned", on_change=on_input_change)
    st.number_input("退/續數量", min_value=0, step=1, key="input_hk_co", on_change=on_input_change)
    st.number_input("每人平均掃房數", min_value=0.0, step=0.1, key="input_hk_avg", on_change=on_input_change)
    st.number_input("房務請購費用", min_value=0, step=100, key="input_hk_exp", on_change=on_input_change)

with tab4:
    st.header("🍽️ 餐廳數據")
    st.subheader("📁 數據報表上傳")
    rest_file = st.file_uploader("上傳餐廳報表 (Excel)，會自動把整份報表寫入資料庫！", type=["xls", "xlsx"], key="rest_uploader")
    
    if rest_file:
        # 在寫入前增加預覽區
        try:
            # 暫時執行解析 (不存入資料庫)
            # 為了效率與介面，我們在這裡做個簡化的預覽
            df_preview = pd.read_excel(rest_file, header=None)
            st.info("🔍 **報表內容初步掃描：**")
            
            p_month_rev = 0
            p_avg_spent = 0
            found_days = 0
            
            for i, row in df_preview.iterrows():
                row_str = " ".join([str(v) for v in row if pd.notna(v)])
                if ('已結算營收' in row_str or '月營收' in row_str) and '早餐' not in row_str and '下午茶' not in row_str:
                    for v in row:
                        if any(c.isdigit() for c in str(v)) and not any(k in str(v) for k in ['已結算營收', '月營收']):
                            try: p_month_rev = int(float(str(v).replace('NT$', '').replace('$', '').replace(',', '').strip())); break
                            except: pass
                if '客單價' in row_str:
                    for v in row:
                        if any(c.isdigit() for c in str(v)) and '客單價' not in str(v):
                            try: p_avg_spent = int(float(str(v).replace('NT$', '').replace('$', '').replace(',', '').strip())); break
                            except: pass
                if re.search(r'\d{1,2}/\d{1,2}', str(row[0])): found_days += 1
            
            pv_col1, pv_col2, pv_col3 = st.columns(3)
            pv_col1.metric("辨識出月結算營收", f"NT$ {p_month_rev:,}")
            pv_col2.metric("辨識出平均客單價", f"NT$ {p_avg_spent:,}")
            pv_col3.metric("辨識出每日明細", f"{found_days} 筆")
            
            if p_month_rev == 0:
                st.warning("⚠️ 系統未能從報表中自動找到「月結算營收」，請確認報表格式或手動檢查。")

            if st.button("📥 確認無誤，寫入系統資料庫", key="rest_btn"):
                saved_count = parse_and_save_restaurant(rest_file, selected_date.year)
                if saved_count:
                    st.success(f"✅ 成功更新 {saved_count} 筆每日餐廳資料！")
                    time.sleep(1)
                    st.rerun()
        except Exception as ex:
            st.error(f"預覽失敗: {ex}")

    st.divider()
    st.subheader(f"餐廳手動確認區 ({date_str})")
    
    st.markdown("#### 🌞 早餐數據")
    b1, b2, b3 = st.columns(3)
    b1.number_input("【主題】預估來客", min_value=0, step=1, key="input_bf_theme_est", on_change=on_input_change)
    b1.number_input("【主題】實際來客", min_value=0, step=1, key="input_bf_theme_act", on_change=on_input_change)
    
    b2.number_input("【站前】預估來客", min_value=0, step=1, key="input_bf_zq_est", on_change=on_input_change)
    b2.number_input("【站前】實際來客", min_value=0, step=1, key="input_bf_zq_act", on_change=on_input_change)
    
    b3.number_input("【兩館總和】預估", min_value=0, step=1, key="input_bf_total_est", on_change=on_input_change)
    b3.number_input("【兩館總和】實際", min_value=0, step=1, key="input_bf_total_act", on_change=on_input_change)

    st.markdown("#### 🍰 下午茶數據")
    a1, a2, a3 = st.columns(3)
    a1.number_input("【主題】預估來客", min_value=0, step=1, key="input_af_theme_est", on_change=on_input_change)
    a1.number_input("【主題】實際來客", min_value=0, step=1, key="input_af_theme_act", on_change=on_input_change)
    
    a2.number_input("【站前】預估來客", min_value=0, step=1, key="input_af_zq_est", on_change=on_input_change)
    a2.number_input("【站前】實際來客", min_value=0, step=1, key="input_af_zq_act", on_change=on_input_change)
    
    a3.number_input("【兩館總和】預估", min_value=0, step=1, key="input_af_total_est", on_change=on_input_change)
    a3.number_input("【兩館總和】實際", min_value=0, step=1, key="input_af_total_act", on_change=on_input_change)

    st.markdown("#### 📊 月報結算總數與雜項")
    c1, c2, c3 = st.columns(3)
    c1.number_input("已結算營收 (全月)", min_value=0, step=100, key="input_rest_mrev", on_change=on_input_change)
    c2.number_input("平均客單價", min_value=0, step=10, key="input_rest_aspent", on_change=on_input_change)
    c3.number_input("THE PEAK 請購費用", min_value=0, step=100, key="input_rest_exp", on_change=on_input_change)
    
    col_rest1, col_rest2 = st.columns(2)
    col_rest1.number_input("The Peak 當日來客數", min_value=0, step=1, key="input_peak_act", on_change=on_input_change)
    col_rest2.number_input("Happy Hour 當日來客數", min_value=0, step=1, key="input_hh_act", on_change=on_input_change)

with tab5:
    st.header("🔧 工務數據")
    st.number_input("今日待修房數", min_value=0, step=1, key="input_repair", on_change=on_input_change)
    st.text_area("修繕紀錄", key="input_maint_rec", on_change=on_input_change)
    st.number_input("工務請購費用", min_value=0, step=100, key="input_maint_exp", on_change=on_input_change)

with tab6:
    st.header("📝 每日營運紀錄")

    # --- 金旭報表上傳 + 手動輸入 (從原「櫃台數據」移入) ---
    with st.expander("📁 金旭報表上傳 & 當日數字手動確認", expanded=False):
        jinxu_file = st.file_uploader("上傳金旭報表 (Excel/CSV)，會自動把整份報表寫入資料庫！", type=["csv", "xls", "xlsx"], key="jinxu_uploader")
        if jinxu_file:
            if st.button("📥 寫入系統資料庫"):
                saved_count = parse_and_save_jinxu(jinxu_file)
                if saved_count:
                    st.success(f"✅ 成功將 {saved_count} 筆每日資料存入系統資料庫！切換日期即可自動調出。")
                    time.sleep(1)
                    st.rerun()
        st.divider()
        st.subheader(f"📋 當日數字手動確認 ({date_str})")
        rc1, rc2, rc3 = st.columns(3)
        rc1.number_input("訂房率 (%)", min_value=0.0, max_value=100.0, step=0.1, key="input_occ", on_change=on_input_change)
        rc2.number_input("ADR (平均房價)", min_value=0, step=10, key="input_adr", on_change=on_input_change)
        rc3.number_input("總營收", min_value=0, step=100, key="input_rev", on_change=on_input_change)
        rc4, rc5 = st.columns(2)
        rc4.number_input("總住房數", min_value=0, step=1, key="input_rooms", on_change=on_input_change)
        rc5.number_input("櫃台請購費用", min_value=0, step=100, key="input_counter_exp", on_change=on_input_change)
        st.text_area("負評客訴", key="input_complaints", on_change=on_input_change)
    

    if selected_week != "--- 關閉週預覽 ---":
        import calendar
        _, last_day_of_month = calendar.monthrange(selected_date.year, selected_date.month)
        
        # 解析選擇的區間
        week_idx = weekly_options.index(selected_week)
        start_d = (week_idx - 1) * 7 + 1
        if week_idx == 5:
            end_d = last_day_of_month
        else:
            end_d = min(start_d + 6, last_day_of_month)
            
        st.subheader(f"📋 {selected_week} 快速審視模式")
        st.info(f"正在查看 {selected_date.year}年度 {selected_date.month}月份 ({start_d}號 至 {end_d}號) 的完整紀錄。")
        
        # 獲取該區間所有資料
        c_month_str = selected_date.strftime('%Y-%m')
        
        for day in range(start_d, end_d + 1):
            target_date = f"{c_month_str}-{day:02d}"
            # 這裡我們呼叫 get_daily_data
            d_data = get_daily_data(target_date)
            
            with st.expander(f"📅 {target_date} 營運紀錄", expanded=True):
                day_log = get_daily_log(target_date)
                if day_log:
                    st.markdown(f"**【當日日誌細節】**\n\n{day_log}")
                    st.divider()
                    col_a, colb, colc = st.columns(3)
                    col_a.metric("住房率", f"{d_data.get('occ_rate', 0)}%")
                    colb.metric("ADR", f"NT$ {int(d_data.get('adr', 0)):,}")
                    colc.metric("營收", f"NT$ {int(d_data.get('revenue', 0)):,}")
                else:
                    st.write("🌑 此日期尚無任何日誌紀錄。")
        
        if st.button("⬅️ 返回今日編輯模式"):
            st.rerun()

    else:
        st.info(f"💡 請在下方詳細填寫 **{date_str}** 的各項營運日誌與重點工作回報。這裡的紀錄會自動儲存，切換日期或關閉網頁也不用擔心遺失。")
        st.text_area("✍️ 今日工作與營運細節報告：", height=500, key="input_daily_log", placeholder="可以在這裡記錄交班重點、客訴特殊處理、VIP 接待細節、設備大修紀錄...等", on_change=on_input_change)

with tab_p:
    st.header("💰 採購花費分析統計")
    
    current_month_str = selected_date.strftime('%Y-%m')
    
    try:
        # 讀取採購數據 (降低 TTL 以確保更新及時)
        possible_names = ["purchase data", "Purchase Data", "purchase_data", "Purchase_Data"]
        df_purchase = None
        used_name = ""
        
        for name in possible_names:
            try:
                df_purchase = conn.read(worksheet=name, ttl="1m")
                if df_purchase is not None and not df_purchase.empty:
                    used_name = name
                    break
            except:
                continue
        
        if df_purchase is not None and not df_purchase.empty:
            # 清理欄位名稱 (移除空格)
            df_purchase.columns = df_purchase.columns.astype(str).str.strip()
            
            # 尋找關鍵欄位 (自動識別可能的名稱變體)
            date_col = next((c for c in df_purchase.columns if '日期' in c or 'Date' in c), None)
            dept_col = next((c for c in df_purchase.columns if '部門' in c or 'Dept' in c or '工地' in c), None)
            total_col = next((c for c in df_purchase.columns if '小計' in c or '金額' in c or 'Total' in c), None)
            
            if not date_col or not dept_col or not total_col:
                missing = [c for c, found in [('日期', date_col), ('部門', dept_col), ('小計', total_col)] if not found]
                st.error(f"❌ 採購分頁缺少必要欄位：{', '.join(missing)}")
                st.write("目前偵測到的欄位有：", list(df_purchase.columns))
                st.stop()

            # 確保日期欄位為日期型態 (支援民國年與一般西元年)
            def robust_date_parse(val):
                if pd.isna(val): return None
                s = str(val).strip()
                # 判斷是否為民國年格式 (含 / 且部分較小)
                if '/' in s:
                    res = minguo_to_western(s)
                    if res: return res
                # 嘗試標準解析
                try: return pd.to_datetime(val).date()
                except: return None

            df_purchase['日期'] = df_purchase[date_col].apply(robust_date_parse)
            
            # 處理部門欄位空值 (歸類到「未分類」)
            df_purchase[dept_col] = df_purchase[dept_col].fillna("未分類").astype(str).str.strip()
            df_purchase.loc[df_purchase[dept_col] == "", dept_col] = "未分類"
            
            # 過濾 NaT/None
            df_purchase = df_purchase[df_purchase['日期'].notna()]
            
            # 過濾當月數據
            m_start = selected_date.replace(day=1)
            import calendar
            _, last_day = calendar.monthrange(selected_date.year, selected_date.month)
            m_end = selected_date.replace(day=last_day)
            
            df_month = df_purchase[(df_purchase['日期'] >= m_start) & (df_purchase['日期'] <= m_end)].copy()
            
            # --- 新增：取得上個月數據用於 MoM 分析 ---
            prev_m_date = get_month_delta(selected_date, -1)
            pm_start = prev_m_date.replace(day=1)
            _, pm_last_day = calendar.monthrange(prev_m_date.year, prev_m_date.month)
            pm_end = prev_m_date.replace(day=pm_last_day)
            df_prev_month = df_purchase[(df_purchase['日期'] >= pm_start) & (df_purchase['日期'] <= pm_end)].copy()
            
            if not df_month.empty:
                # 數值清理
                df_month['小計'] = pd.to_numeric(df_month[total_col], errors='coerce').fillna(0)
                if not df_prev_month.empty:
                    df_prev_month['小計'] = pd.to_numeric(df_prev_month[total_col], errors='coerce').fillna(0)
                
                total_month_expense = df_month['小計'].sum()
                total_prev_expense = df_prev_month['小計'].sum() if not df_prev_month.empty else 0
                
                # 計算增長率
                mom_delta = total_month_expense - total_prev_expense
                mom_pcnt = (mom_delta / total_prev_expense * 100) if total_prev_expense > 0 else 0
                
                # 1. 本月總開銷與 MoM
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #1f2c56 0%, #2e437c 100%); padding: 25px; border-radius: 15px; text-align: center; color: white; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
                    <p style="margin: 0; font-size: 1.1rem; opacity: 0.8;">📅 {current_month_str} 本月總開銷金額</p>
                    <h1 style="margin: 10px 0 0 0; font-size: 3rem; font-weight: 800; letter-spacing: 1px;">NT$ {int(total_month_expense):,}</h1>
                </div>
                """, unsafe_allow_html=True)
                
                # 顯示 MoM 指標
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.metric("上月同期總額", f"NT$ {int(total_prev_expense):,}")
                with col_m2:
                    st.metric("月增長金額 (MoM)", f"NT$ {int(mom_delta):,}", delta=int(mom_delta), delta_color="inverse")
                with col_m3:
                    st.metric("月增長百分比", f"{mom_pcnt:.1f}%", delta=f"{mom_pcnt:.1f}%", delta_color="inverse")
                
                st.divider()

                # --- 異常值監控：找出增長過快的部門 ---
                st.subheader("⚠️ 採購異常監控 (MoM Spikes)")
                # 計算各部門本月 vs 上月
                curr_depts = df_month.groupby(dept_col)['小計'].sum().reset_index()
                curr_depts.columns = ['部門', '小計']
                
                if not df_prev_month.empty:
                    prev_depts = df_prev_month.groupby(dept_col)['小計'].sum().reset_index()
                    prev_depts.columns = ['部門', '小計']
                else:
                    prev_depts = pd.DataFrame(columns=['部門', '小計'])
                
                comparison = pd.merge(curr_depts, prev_depts, on='部門', how='left', suffixes=('_今', '_昨')).fillna(0)
                
                # 安全計算變動率 (避免 ZeroDivisionError 與 Indexing 類型報錯)
                def calc_mom_ratio(row):
                    if row['小計_昨'] > 0:
                        return (row['小計_今'] - row['小計_昨']) / row['小計_昨'] * 100
                    return 100.0 if row['小計_今'] > 0 else 0.0
                
                comparison['變動率'] = comparison.apply(calc_mom_ratio, axis=1)
                
                # 找出變動率大於 20% 且金額大於一定門檻的 (例如 > 2000)
                spikes = comparison[(comparison['變動率'] > 20) & (comparison['小計_今'] > 2000)].sort_values('變動率', ascending=False)
                
                if not spikes.empty:
                    for _, row in spikes.iterrows():
                        st.warning(f"🚩 **{row['部門']}** 本月開銷異常！較上月增長 **{row['變動率']:.1f}%** (NT$ {int(row['小計_今']):,})")
                else:
                    st.success("✅ 目前各部門採購金額平穩，未偵測到異常大幅波動。")

                st.divider()
                
                # 2. 部門佔比圓餅圖
                st.subheader("📊 各部門請購佔比分析")
                dept_summary = df_month.groupby(dept_col)['小計'].sum().reset_index()
                dept_summary.columns = ['部門', '小計']
                
                # 繪製圓餅圖 (依照金額排序)
                base = alt.Chart(dept_summary).encode(
                    theta=alt.Theta(field="小計", type="quantitative", stack=True),
                    color=alt.Color(
                        field="部門", 
                        type="nominal", 
                        scale=alt.Scale(scheme='category10'), 
                        legend=alt.Legend(title="部門", orient="right"),
                        sort=alt.SortField("小計", order="descending")
                    ),
                    order=alt.Order("小計", sort="descending"),
                    tooltip=["部門", alt.Tooltip("小計", format=",.0f", title="總金額 (NT$)")]
                ).properties(height=450)
                
                # 圓餅主體
                chart_arc = base.mark_arc(innerRadius=60, outerRadius=120, stroke="#fff")
                
                # 在圓餅切片上顯示金額
                chart_text = base.mark_text(radius=90, size=14, fontWeight="bold", color="white").encode(
                    text=alt.Text("小計:Q", format=",.0f")
                )
                
                st.altair_chart(chart_arc + chart_text, use_container_width=True)
                
                # --- 新增：餐飲績效分析 (The Peak & Happy Hour) ---
                st.divider()
                st.subheader("🍽️ 餐飲績效與成本深度分析 (Cash-basis)")
                
                # 獲取當月每日數據 (含餐廳來客數)
                m_data = fetch_month_summary(selected_date.year, selected_date.month)
                df_daily_rest = m_data['df']
                
                if not df_daily_rest.empty:
                    # 確保必要欄位存在並轉為數值
                    target_cols = ['rest_day_guests', 'rest_hh_guests', 'revenue', 'bf_total_act', 'af_total_act']
                    for c in target_cols:
                        if c in df_daily_rest.columns:
                            df_daily_rest[c] = pd.to_numeric(df_daily_rest[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                        else:
                            df_daily_rest[c] = 0
                    
                    # --- 自動化邏輯：如果 The Peak 來客數為 0，則自動加總 早餐 + 下午茶 ---
                    def calculate_peak_guests(row):
                        if row['rest_day_guests'] > 0:
                            return row['rest_day_guests']
                        return row['bf_total_act'] + row['af_total_act']
                    
                    df_daily_rest['effective_peak_guests'] = df_daily_rest.apply(calculate_peak_guests, axis=1)

                    # 篩選 The Peak 與 Happy Hour 採購 (強力模糊匹配)
                    all_depts_list = dept_summary['部門'].astype(str).tolist()
                    
                    # HH 匹配：包含 '4'、'HH' 或 'HAPPY'
                    hh_matched = [d for d in all_depts_list if '4' in d or any(k in d.upper() for k in ['HH', 'HAPPY', '歡樂時光'])]
                    # Peak 匹配：包含 'PEAK' 或 '餐廳'，且排除 HH 部門
                    peak_matched = [d for d in all_depts_list if (any(k in d.upper() for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲'])) and (d not in hh_matched)]
                    
                    with st.expander("🛠️ 數據匹配校準器 (若數據不正確請點開)"):
                        st.info(f"📍 偵測到之所有部門: `{all_depts_list}`")
                        st.success(f"🍷 歸類為 Happy Hour (HH) 之部門: `{hh_matched}`")
                        st.success(f"🏰 歸類為 The Peak (餐廳) 之部門: `{peak_matched}`")
                    
                    df_peak_purchase = df_month[df_month[dept_col].isin(peak_matched)].copy()
                    df_hh_purchase = df_month[df_month[dept_col].isin(hh_matched)].copy()
                    
                    # --- 進階匹配：若部門抓不到 HH，嘗試從品名抓取 ---
                    if df_hh_purchase.empty:
                        item_col = next((c for c in df_month.columns if any(k in c for k in ['品名', '項目', 'Item'])), None)
                        if item_col:
                            df_hh_purchase = df_month[df_month[item_col].astype(str).str.upper().str.contains('HH|HAPPY|歡樂時光', na=False)].copy()
                    
                    # 計算每日採購總額（改用『以週為單位均攤』修正採購日 vs 消耗日失真）
                    df_daily_rest['日期_obj'] = pd.to_datetime(df_daily_rest['date']).dt.date
                    df_daily_rest['日期_dt'] = pd.to_datetime(df_daily_rest['date'])
                    
                    def spread_weekly_cost(df_purchase, df_daily_base):
                        """將採購費用以週為單位，均攤到當週有來客的每一天"""
                        if df_purchase.empty or df_daily_base.empty:
                            return pd.Series(0, index=df_daily_base['日期_obj'])
                        
                        # 加上 ISO 週別
                        df_purchase = df_purchase.copy()
                        df_purchase['week'] = pd.to_datetime(df_purchase['日期']).dt.isocalendar().week.astype(int)
                        df_purchase['year'] = pd.to_datetime(df_purchase['日期']).dt.isocalendar().year.astype(int)
                        weekly_cost = df_purchase.groupby(['year', 'week'])['小計'].sum().reset_index()
                        
                        df_base = df_daily_base.copy()
                        df_base['week'] = df_base['日期_dt'].dt.isocalendar().week.astype(int)
                        df_base['year'] = df_base['日期_dt'].dt.isocalendar().year.astype(int)
                        df_base['has_guest'] = df_base['effective_peak_guests'] > 0
                        
                        # 每週有來客的天數
                        days_per_week = df_base.groupby(['year', 'week'])['has_guest'].sum().reset_index()
                        days_per_week.columns = ['year', 'week', 'active_days']
                        days_per_week['active_days'] = days_per_week['active_days'].replace(0, 1)  # 防零除
                        
                        # 合併週成本
                        df_base = pd.merge(df_base, weekly_cost, on=['year', 'week'], how='left').fillna(0)
                        df_base = pd.merge(df_base, days_per_week, on=['year', 'week'], how='left')
                        df_base['spread_cost'] = df_base['小計'] / df_base['active_days']
                        
                        return df_base.set_index('日期_obj')['spread_cost']
                    
                    # 用週均攤計算每日成本
                    peak_spread = spread_weekly_cost(df_peak_purchase, df_daily_rest)
                    hh_spread = spread_weekly_cost(df_hh_purchase, df_daily_rest)
                    
                    # 合併來客數與週均攤成本
                    analysis_df = df_daily_rest[['日期_obj', 'effective_peak_guests', 'rest_hh_guests', 'revenue']].copy()
                    analysis_df['peak_cost'] = analysis_df['日期_obj'].map(peak_spread).fillna(0)
                    analysis_df['hh_cost'] = analysis_df['日期_obj'].map(hh_spread).fillna(0)
                    
                    # --- 累計分析邏輯：計算本月至今的累積數據 ---
                    analysis_df = analysis_df.sort_values('日期_obj')
                    analysis_df['cum_peak_cost'] = analysis_df['peak_cost'].cumsum()
                    analysis_df['cum_peak_guests'] = analysis_df['effective_peak_guests'].cumsum()
                    analysis_df['cum_hh_cost'] = analysis_df['hh_cost'].cumsum()
                    analysis_df['cum_hh_guests'] = analysis_df['rest_hh_guests'].cumsum()
                    
                    # 計算累積 CPG (這才是真實的平均成本走勢)
                    analysis_df['cum_peak_cpg'] = analysis_df.apply(lambda r: r['cum_peak_cost']/r['cum_peak_guests'] if r['cum_peak_guests']>0 else 0, axis=1)
                    analysis_df['cum_hh_cpg'] = analysis_df.apply(lambda r: r['cum_hh_cost']/r['cum_hh_guests'] if r['cum_hh_guests']>0 else 0, axis=1)

                    # UI 呈現 (本月總結)
                    total_peak_cost = analysis_df['cum_peak_cost'].iloc[-1] if not analysis_df.empty else 0
                    total_peak_guests = analysis_df['cum_peak_guests'].iloc[-1] if not analysis_df.empty else 0
                    final_peak_cpg = total_peak_cost / total_peak_guests if total_peak_guests > 0 else 0
                    
                    total_hh_cost = analysis_df['cum_hh_cost'].iloc[-1] if not analysis_df.empty else 0
                    total_hh_guests = analysis_df['cum_hh_guests'].iloc[-1] if not analysis_df.empty else 0
                    final_hh_cpg = total_hh_cost / total_hh_guests if total_hh_guests > 0 else 0

                    # --- 📈 The Peak CPG 月趨勢圖（過去 6 個月）---
                    st.markdown("##### 📈 The Peak CPG 月趨勢（過去 6 個月）")
                    
                    trend_rows = []
                    for n_back in range(5, -1, -1):  # 從 5 個月前到本月
                        t_date = get_month_delta(selected_date, -n_back)
                        t_label = t_date.strftime('%Y-%m')
                        
                        # 抓該月採購數據
                        t_start = t_date.replace(day=1)
                        import calendar as _cal
                        _, t_last = _cal.monthrange(t_date.year, t_date.month)
                        t_end = t_date.replace(day=t_last)
                        df_t_purchase = df_purchase[(df_purchase['日期'] >= t_start) & (df_purchase['日期'] <= t_end)].copy()
                        
                        if not df_t_purchase.empty:
                            df_t_purchase['小計'] = pd.to_numeric(df_t_purchase[total_col], errors='coerce').fillna(0)
                        
                        # 抓該月來客數
                        t_m_data = fetch_month_summary(t_date.year, t_date.month)
                        t_df = t_m_data.get('df', pd.DataFrame())
                        t_guests = 0
                        if not t_df.empty:
                            for _c in ['rest_day_guests', 'bf_total_act', 'af_total_act']:
                                if _c in t_df.columns:
                                    t_df[_c] = pd.to_numeric(t_df[_c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                            if 'rest_day_guests' in t_df.columns and t_df['rest_day_guests'].sum() > 0:
                                t_guests = t_df['rest_day_guests'].sum()
                            elif 'bf_total_act' in t_df.columns:
                                t_guests = (t_df['bf_total_act'] + t_df.get('af_total_act', 0)).sum()
                        
                        # 篩選 The Peak 採購
                        t_peak_cost = 0
                        if not df_t_purchase.empty and dept_col in df_t_purchase.columns:
                            t_all_depts = df_t_purchase[dept_col].astype(str).unique().tolist()
                            t_hh = [d for d in t_all_depts if '4' in d or any(k in d.upper() for k in ['HH', 'HAPPY'])]
                            t_peak_depts = [d for d in t_all_depts if any(k in d.upper() for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲']) and d not in t_hh]
                            t_peak_cost = df_t_purchase[df_t_purchase[dept_col].isin(t_peak_depts)]['小計'].sum()
                        
                        t_cpg = t_peak_cost / t_guests if t_guests > 0 else None
                        trend_rows.append({'月份': t_label, 'CPG': t_cpg, '目標': 150})
                    
                    trend_df = pd.DataFrame(trend_rows).dropna(subset=['CPG'])
                    
                    if not trend_df.empty and len(trend_df) >= 2:
                        base = alt.Chart(trend_df)
                        cpg_line = base.mark_line(point=True, strokeWidth=2.5, color='#1f2c56').encode(
                            x=alt.X('月份:N', title='月份', sort=None),
                            y=alt.Y('CPG:Q', title='每客成本 CPG (NT$)', scale=alt.Scale(zero=False)),
                            tooltip=[alt.Tooltip('月份:N', title='月份'), alt.Tooltip('CPG:Q', title='CPG (NT$)', format=',.0f')]
                        )
                        target_line = alt.Chart(pd.DataFrame({'y': [150]})).mark_rule(
                            color='#e74c3c', strokeDash=[6, 3], strokeWidth=1.5
                        ).encode(y='y:Q')
                        target_label = alt.Chart(pd.DataFrame({'y': [150], 'x': [trend_df['月份'].iloc[-1]], 'text': ['目標 $150']})).mark_text(
                            align='right', dx=-4, dy=-8, color='#e74c3c', fontSize=11, fontWeight='bold'
                        ).encode(x='x:N', y='y:Q', text='text:N')
                        st.altair_chart(
                            alt.layer(cpg_line, target_line, target_label).properties(height=220),
                            use_container_width=True
                        )
                    else:
                        st.info("💡 需要至少 2 個月的數據才能顯示 CPG 趨勢圖。")
                    
                    st.divider()

                    # --- 📊 採購花費 vs 早餐來客數 相關性驗證（以週為單位）---
                    st.markdown("##### 📊 採購花費 vs 早餐來客數 相關性驗證（週）")
                    st.caption("💡 兩條線的形狀應趨近一致。若某週「採購↑ 來客↓」或「採購↓ 來客↑」，代表食材控管可能有問題。")
                    
                    corr_df = analysis_df[['日期_obj', 'peak_cost', 'effective_peak_guests']].copy()
                    corr_df['日期_dt'] = pd.to_datetime(corr_df['日期_obj'])
                    corr_df['week'] = corr_df['日期_dt'].dt.isocalendar().week.astype(int)
                    corr_df['year'] = corr_df['日期_dt'].dt.isocalendar().year.astype(int)
                    corr_df['week_start'] = corr_df['日期_dt'].apply(lambda x: x - pd.Timedelta(days=x.dayofweek))
                    
                    weekly_corr = corr_df.groupby('week_start').agg(
                        採購金額=('peak_cost', 'sum'),
                        來客人數=('effective_peak_guests', 'sum')
                    ).reset_index()
                    weekly_corr['週次'] = weekly_corr['week_start'].dt.strftime('W%V\n%m/%d')
                    
                    # 標準化成 0–100%（對各自最大值）
                    max_cost = weekly_corr['採購金額'].max()
                    max_guest = weekly_corr['來客人數'].max()
                    weekly_corr['採購(%)'] = (weekly_corr['採購金額'] / max_cost * 100).round(1) if max_cost > 0 else 0
                    weekly_corr['來客(%)'] = (weekly_corr['來客人數'] / max_guest * 100).round(1) if max_guest > 0 else 0
                    weekly_corr['背道而馳'] = (abs(weekly_corr['採購(%)'] - weekly_corr['來客(%)']) > 25).map({True: '⚠️ 異常', False: '✅ 正常'})
                    
                    if not weekly_corr.empty and max_cost > 0 and max_guest > 0:
                        # 轉成長格式給 Altair
                        melt_df = weekly_corr.melt(
                            id_vars=['週次', '背道而馳', '採購金額', '來客人數'],
                            value_vars=['採購(%)', '來客(%)'],
                            var_name='指標', value_name='標準化數值'
                        )
                        color_map = {'採購(%)': '#e67e22', '來客(%)': '#2980b9'}
                        
                        corr_chart = alt.Chart(melt_df).mark_line(point=True, strokeWidth=2.5).encode(
                            x=alt.X('週次:N', title='週次', sort=None),
                            y=alt.Y('標準化數值:Q', title='相對比例 (% of max)', scale=alt.Scale(domain=[0, 110])),
                            color=alt.Color('指標:N',
                                scale=alt.Scale(domain=list(color_map.keys()), range=list(color_map.values())),
                                legend=alt.Legend(title='指標', orient='bottom')
                            ),
                            tooltip=[
                                alt.Tooltip('週次:N', title='週次'),
                                alt.Tooltip('指標:N', title='指標'),
                                alt.Tooltip('採購金額:Q', title='採購金額 (NT$)', format=',.0f'),
                                alt.Tooltip('來客人數:Q', title='來客人數 (人)', format=',.0f'),
                                alt.Tooltip('背道而馳:N', title='健康狀態'),
                            ]
                        ).properties(height=220)
                        
                        st.altair_chart(corr_chart, use_container_width=True)
                        
                        # 標出背道而馳的週次
                        bad_weeks = weekly_corr[weekly_corr['背道而馳'] == '⚠️ 異常']
                        if not bad_weeks.empty:
                            for _, bw in bad_weeks.iterrows():
                                diff = bw['採購(%)'] - bw['來客(%)']
                                direction = "採購偏高（來客少但食材買太多）" if diff > 0 else "來客偏高（來客多但食材買太少）"
                                st.warning(f"⚠️ **{bw['週次'].replace(chr(10), ' ')}** 出現背道而馳！{direction}　採購 NT$ {int(bw['採購金額']):,} | 來客 {int(bw['來客人數'])} 人")
                        else:
                            st.success("✅ 本月各週採購花費與來客人數走勢一致，食材控管健康。")
                    else:
                        st.info("💡 本月資料不足，無法進行相關性分析。")

                    st.divider()
                    c_ana1, c_ana2 = st.columns(2)


                    with c_ana1:
                        st.markdown(f"<div style='background:#f8f9fa; padding:15px; border-radius:10px; border-top:4px solid #1f2c56;'>", unsafe_allow_html=True)
                        st.markdown(f"**🏰 The Peak (餐廳)**")
                        st.metric("本月總採購額", f"NT$ {int(total_peak_cost):,}")
                        is_auto = "(自動加總)" if (df_daily_rest['rest_day_guests'].sum() == 0 and total_peak_guests > 0) else ""
                        st.metric(f"本月總來客數 {is_auto}", f"{int(total_peak_guests):,} 人")
                        
                        # CPG 顏色警示 (目標 $150)
                        peak_target = 150
                        delta_val = peak_target - final_peak_cpg
                        st.metric("平均每客成本 (CPG)", f"NT$ {int(final_peak_cpg):,}", delta=f"{int(delta_val)} (距離目標)" if delta_val >=0 else f"{int(delta_val)} (已超標)", delta_color="normal" if delta_val >=0 else "inverse")
                        st.markdown("</div>", unsafe_allow_html=True)
                        
                        # --- 新增：財務預測與目標控管 ---
                        st.write("")
                        st.markdown("##### 🎯 財務目標控管")
                        # 1. 成本佔比 (餐飲成本 / 總營收)
                        total_hotel_rev = m_data['rev']
                        cost_ratio = (total_peak_cost / total_hotel_rev * 100) if total_hotel_rev > 0 else 0
                        st.write(f"📊 目前成本佔總營收比例: **{cost_ratio:.1f}%**")
                        
                        # 2. 月底支出預測
                        import calendar
                        _, last_day_num = calendar.monthrange(selected_date.year, selected_date.month)
                        current_day_num = len(analysis_df)
                        if current_day_num > 0:
                            daily_avg_cost = total_peak_cost / current_day_num
                            forecast_total = total_peak_cost + (daily_avg_cost * (last_day_num - current_day_num))
                            
                            forecast_color = "red" if final_peak_cpg > peak_target else "green"
                            st.markdown(f"🔮 月底預估總支出: <span style='color:{forecast_color}; font-weight:bold;'>NT$ {int(forecast_total):,}</span>", unsafe_allow_html=True)
                            if final_peak_cpg > peak_target:
                                st.warning(f"⚠️ 警告：目前每客成本 ({int(final_peak_cpg)}) 已高於目標 {peak_target} 元，請檢視進貨項目或份量控管。")
                        # ----------------------------
                    with c_ana2:
                        st.markdown(f"<div style='background:#fff9f0; padding:15px; border-radius:10px; border-top:4px solid #ff9f43;'>", unsafe_allow_html=True)
                        st.markdown(f"**🥂 Happy Hour (HH)**")
                        st.metric("本月總採購額", f"NT$ {int(total_hh_cost):,}")
                        st.metric("本月總來客數", f"{int(total_hh_guests):,} 人")
                        st.metric("平均每客服務成本", f"NT$ {int(final_hh_cpg):,}")
                        if total_hh_cost > 0 and total_hh_guests == 0:
                            st.warning("⚠️ 有產生 HH 採購費用但總來客數為 0！請至「🍽️ 餐廳數據」補登以計算每客服務成本 (CPG)。")
                        st.markdown("</div>", unsafe_allow_html=True)

                    st.write("")
                    # 趨勢圖表
                    st.markdown("#### 📈 本月累計每客成本趨勢 (Monthly Cumulative CPG)")
                    analysis_df['日期_str'] = analysis_df['日期_obj'].astype(str)
                    
                    # 整合圖表
                    base_chart = alt.Chart(analysis_df).encode(x=alt.X('日期_str:T', title='日期'))
                    
                    peak_line = base_chart.mark_line(point=True, color='#1f2c56', strokeWidth=3).encode(
                        y=alt.Y('cum_peak_cpg:Q', title='累計平均成本 (NT$)'),
                        tooltip=['日期_str', alt.Tooltip('cum_peak_guests', title='累計來客'), alt.Tooltip('cum_peak_cost', title='累計採購'), alt.Tooltip('cum_peak_cpg', format='.0f', title='累計 CPG')]
                    )
                    
                    st.altair_chart(peak_line.properties(title="The Peak 累計平均成本趨勢", height=300), use_container_width=True)
                    
                    if total_hh_guests > 0:
                        st.write("")
                        st.markdown("#### 🥂 Happy Hour 累計成本分析")
                        
                        # 顯示累計人數 vs 累計成本
                        hh_chart_base = alt.Chart(analysis_df).encode(x=alt.X('日期_str:T', title='日期'))
                        
                        # 長條圖顯示累計 CPG
                        hh_bar = hh_chart_base.mark_bar(color='#ff9f43', opacity=0.7).encode(
                            y=alt.Y('cum_hh_cpg:Q', title='累計平均成本 (NT$)'),
                            tooltip=[
                                '日期_str', 
                                alt.Tooltip('cum_hh_guests', title='累計來客 (分母)'), 
                                alt.Tooltip('cum_hh_cost', title='累計採購 (分子)'), 
                                alt.Tooltip('cum_hh_cpg', format='.1f', title='累計 CPG')
                            ]
                        )
                        
                        # 疊加一條線顯示累計人數的成長 (確保分母正確)
                        hh_guest_line = hh_chart_base.mark_line(color='#e67e22', strokeDash=[5,5]).encode(
                            y=alt.Y('cum_hh_guests:Q', title='累計人數'),
                            tooltip=['日期_str', alt.Tooltip('cum_hh_guests', title='累計人數')]
                        )
                        
                        st.altair_chart(alt.layer(hh_bar, hh_guest_line).resolve_scale(y='independent').properties(title="Happy Hour 累計趨勢 (長條:成本, 虛線:人數)", height=300), use_container_width=True)
                        
                    # --- 新增：雙冠日食材消耗對比分析 (Dynamic CPG Analysis) ---
                    st.divider()
                    st.markdown("#### 🎯 雙冠日 vs 一般日：食材消耗對比分析")
                    
                    # 獲取雙冠日清單
                    curr_metrics = calc_key_metrics(m_data)
                    dual_match_dates = curr_metrics.get('dual_match_dates', [])
                    
                    if dual_match_dates:
                        # 將日期標記為雙冠日
                        analysis_df['is_dual_match'] = analysis_df['日期_str'].isin(dual_match_dates)
                        
                        df_dual = analysis_df[analysis_df['is_dual_match']]
                        df_normal = analysis_df[~analysis_df['is_dual_match']]
                        
                        # 計算雙冠日 CPG
                        dual_peak_cost = df_dual['peak_cost'].sum()
                        dual_peak_guests = df_dual['effective_peak_guests'].sum()
                        dual_cpg = dual_peak_cost / dual_peak_guests if dual_peak_guests > 0 else 0
                        
                        # 計算一般日 CPG
                        normal_peak_cost = df_normal['peak_cost'].sum()
                        normal_peak_guests = df_normal['effective_peak_guests'].sum()
                        normal_cpg = normal_peak_cost / normal_peak_guests if normal_peak_guests > 0 else 0
                        
                        cpg_col1, cpg_col2 = st.columns(2)
                        
                        with cpg_col1:
                            st.markdown(f"""
                            <div style="background:#fff5e6; border-left:4px solid #e67e22; padding:15px; border-radius:8px;">
                                <p style="margin:0; font-size:13px; color:#e67e22; font-weight:bold;">🏆 雙冠日 (共 {len(df_dual)} 天)</p>
                                <h3 style="margin:5px 0;">NT$ {int(dual_cpg):,} / 客</h3>
                                <p style="margin:0; font-size:12px; color:#666;">總食材花費: NT$ {int(dual_peak_cost):,} | 服務客數: {int(dual_peak_guests):,} 人</p>
                            </div>
                            """, unsafe_allow_html=True)
                            
                        with cpg_col2:
                            st.markdown(f"""
                            <div style="background:#f8f9fa; border-left:4px solid #95a5a6; padding:15px; border-radius:8px;">
                                <p style="margin:0; font-size:13px; color:#7f8c8d; font-weight:bold;">📉 一般日 (共 {len(df_normal)} 天)</p>
                                <h3 style="margin:5px 0;">NT$ {int(normal_cpg):,} / 客</h3>
                                <p style="margin:0; font-size:12px; color:#666;">總食材花費: NT$ {int(normal_peak_cost):,} | 服務客數: {int(normal_peak_guests):,} 人</p>
                            </div>
                            """, unsafe_allow_html=True)
                            
                        # 顯示策略建議（基於比例，而非絕對差值）
                        st.write("")
                        target_ratio = 1.10  # 雙冠日 CPG 應達到一般日的 110%
                        actual_ratio = (dual_cpg / normal_cpg) if normal_cpg > 0 else 0
                        ratio_pct = actual_ratio * 100
                        
                        if actual_ratio >= target_ratio:
                            st.success(f"💡 **主動備戰策略成功！** 雙冠日的單客成本（NT$ {int(dual_cpg):,}）達到一般日（NT$ {int(normal_cpg):,}）的 **{ratio_pct:.0f}%**，超過 110% 目標。代表你在大日子前有主動備了更好的食材，與高房價形成正向配對。")
                        elif actual_ratio >= 0.90:
                            diff_to_target = int(normal_cpg * target_ratio - dual_cpg)
                            st.info(f"⚖️ **採購尚未主動分級。** 雙冠日 CPG 為一般日的 {ratio_pct:.0f}%（目標 ≥ 110%）。由於週均攤讓旺日（人多）天然壓低 CPG，這個差距屬於合理的規模效應。建議在雙冠日當週多編列 NT$ {diff_to_target:,} / 人左右的品質預算，讓高端客感受得到差異。")
                        else:
                            # 計算行動指引數據
                            avg_normal_guests = (normal_peak_guests / len(df_normal)) if len(df_normal) > 0 else 0
                            peak_target_cpg = 150  # 目標 CPG 上限（與前面財務目標一致）
                            # 建議週採購上限：目標 CPG × 一般日平均每日來客數 × 7 天
                            recommended_weekly_budget = int(peak_target_cpg * avg_normal_guests * 7)
                            # 本月實際週均採購
                            total_weeks = max(1, round(len(df_normal) / 7))
                            actual_weekly_avg = int(normal_peak_cost / total_weeks) if total_weeks > 0 else 0
                            overrun = actual_weekly_avg - recommended_weekly_budget
                            
                            st.error(
                                f"⚠️ **平日食材成本明顯偏高（雙冠日 CPG 僅為一般日的 {ratio_pct:.0f}%）**\n\n"
                                f"📊 **一般日數據**\n"
                                f"- 一般日平均每日來客數：**{avg_normal_guests:.1f} 人**\n"
                                f"- 一般日單客食材成本 (CPG)：**NT$ {int(normal_cpg):,}**\n\n"
                                f"💰 **週採購建議**\n"
                                f"- 以目標 CPG $150 計算，建議每週 The Peak 採購上限：**NT$ {recommended_weekly_budget:,}**\n"
                                f"- 本月實際週均採購：**NT$ {actual_weekly_avg:,}**\n"
                                f"- {'🔴 超出建議上限：NT$ ' + f'{overrun:,}' if overrun > 0 else '🟢 在目標範圍內'}\n\n"
                                f"📋 **可能原因（請擇一追查）**\n"
                                f"1. 平日來客數也偏高，被迫追加採購（合理，可對照 OCC 確認）\n"
                                f"2. 平日備料過多，有生鮮報廢（須檢視）\n"
                                f"3. 領用未確實盤點（須追查）"
                            )
                    else:
                        st.info("💡 本月目前無符合條件的雙冠日，無法進行對比分析。")
                        st.caption("💡 虛線代表累積來客數。如果長條圖在月初是空的，代表該時段尚未產生 HH 相關的採購支出。")
                    
                    st.info("💡 **分析小撇步**：當「每客成本」異常偏高時，請檢查該日期是否有大宗採購進入庫存，或來客數輸入是否正確。")

                else:
                    st.info("尚未偵測到本月的餐廳來客數據，無法進行成本效益分析。")
                
                st.divider()
                
                # 3. 各部門詳細統計
                st.subheader("🏢 各部門經費分析")
                
                # 取得所有部門
                departments = dept_summary.sort_values('小計', ascending=False)['部門'].tolist()
                
                for dept in departments:
                    dept_df = df_month[df_month[dept_col] == dept].copy()
                    dept_total = dept_df['小計'].sum()
                    
                    with st.expander(f"📌 {dept} (總計: NT$ {int(dept_total):,})", expanded=False):
                        # --- 新增：Top 5 高額品項排行榜 ---
                        item_name_col = next((c for c in dept_df.columns if any(k in c for k in ['品名', '項目', 'Item'])), None)
                        if item_name_col:
                            st.markdown("##### 🏆 前五名高額採購品項")
                            top_items = dept_df.groupby(item_name_col)['小計'].sum().sort_values(ascending=False).head(5).reset_index()
                            t_cols = st.columns(5)
                            for idx, row in top_items.iterrows():
                                with t_cols[idx]:
                                    st.metric(f"No.{idx+1} {row[item_name_col][:8]}", f"NT$ {int(row['小計']):,}")
                        st.divider()

                        # --- 新增：排序控制 ---
                        sort_by = st.selectbox(f"排序方式 ({dept})", ["日期 (新→舊)", "金額 (高→低)", "金額 (低→高)", "品項名稱"], key=f"sort_{dept}")
                        
                        if sort_by == "金額 (高→低)":
                            dept_df = dept_df.sort_values('小計', ascending=False)
                        elif sort_by == "金額 (低→高)":
                            dept_df = dept_df.sort_values('小計', ascending=True)
                        elif sort_by == "日期 (新→舊)":
                            dept_df = dept_df.sort_values('日期', ascending=False)
                        elif sort_by == "品項名稱" and item_name_col:
                            dept_df = dept_df.sort_values(item_name_col)

                        # 顯示該部門表格
                        cols_to_show = [c for c in ['日期', '供應商', '品名', '規格', '數量', '單位', '單價', '小計'] if c in dept_df.columns]
                        if not cols_to_show:
                             cols_to_show = dept_df.columns.tolist()
                             
                        st.dataframe(
                            dept_df[cols_to_show],
                            use_container_width=True,
                            hide_index=True
                        )
                        
                # --- 🎯 4. 單品食材消耗率與精準採購方案分析 ---
                if 'analysis_df' in locals() and not analysis_df.empty:
                    st.divider()
                    st.subheader("🎯 單品食材消耗率與精準採購方案分析")
                    st.caption("分析特定關鍵食材品項（如：蛋、高麗菜、海鮮等）的每客平均消耗量，並自動產出精準叫貨配比建議。")

                    item_col = next((c for c in df_month.columns if any(k in c for k in ['品名', '項目', 'Item'])), None)
                    qty_col = next((c for c in df_month.columns if any(k in c for k in ['數量', 'Qty', 'Quantity'])), None)
                    unit_col = next((c for c in df_month.columns if any(k in c for k in ['單位', 'Unit'])), None)
                    price_col = next((c for c in df_month.columns if any(k in c for k in ['單價', 'Price', 'Rate'])), None)
                    
                    if item_col and qty_col:
                        # 擷取常用關鍵字選項
                        all_items = df_month[item_col].dropna().astype(str).str.strip()
                        all_items = all_items[all_items != ""]
                        
                        common_keywords = ["蛋", "菜", "肉", "奶", "米", "麵", "油", "海鮮", "雞", "豬", "牛", "魚"]
                        found_keywords = [k for k in common_keywords if any(k in x for x in all_items)]
                        if not found_keywords:
                            found_keywords = ["蛋"]
                        
                        c_sel1, c_sel2 = st.columns([1, 1])
                        with c_sel1:
                            selected_keyword = st.selectbox(
                                "🔍 選擇分析品項關鍵字",
                                options=found_keywords + ["(自訂輸入)"],
                                index=0,
                                key="item_analysis_keyword_select"
                            )
                        with c_sel2:
                            if selected_keyword == "(自訂輸入)":
                                search_term = st.text_input("✍️ 輸入自訂食材名稱 (例如: 高麗菜)", "蛋", key="item_analysis_custom_input")
                            else:
                                search_term = selected_keyword
                        
                        # 篩選匹配的採購項目
                        item_mask = df_month[item_col].astype(str).str.contains(search_term, na=False, case=False)
                        item_df = df_month[item_mask].copy()
                        
                        if not item_df.empty:
                            # 數值清理
                            item_df['cleaned_qty'] = pd.to_numeric(item_df[qty_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                            item_df['cleaned_total'] = pd.to_numeric(item_df['小計'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                            
                            # 單位判斷
                            most_common_unit = "單位"
                            if unit_col in item_df.columns:
                                most_common_unit = item_df[unit_col].mode().iloc[0] if not item_df[unit_col].empty else "單位"
                            
                            # 每日採購整合
                            item_df['日期_obj'] = pd.to_datetime(item_df['日期']).dt.date
                            daily_item_qty = item_df.groupby('日期_obj')['cleaned_qty'].sum().reset_index()
                            daily_item_cost = item_df.groupby('日期_obj')['cleaned_total'].sum().reset_index()
                            
                            # 合併每日來客
                            item_analysis_df = analysis_df[['日期_obj', 'effective_peak_guests']].copy()
                            item_analysis_df = pd.merge(item_analysis_df, daily_item_qty, on='日期_obj', how='left').fillna(0)
                            item_analysis_df = pd.merge(item_analysis_df, daily_item_cost, on='日期_obj', how='left').fillna(0)
                            
                            # 週彙總計算
                            item_analysis_df['日期_dt'] = pd.to_datetime(item_analysis_df['日期_obj'])
                            item_analysis_df['week_start'] = item_analysis_df['日期_dt'].apply(lambda x: x - pd.Timedelta(days=x.dayofweek))
                            
                            weekly_item = item_analysis_df.groupby('week_start').agg(
                                總採購量=('cleaned_qty', 'sum'),
                                總費用=('cleaned_total', 'sum'),
                                來客人數=('effective_peak_guests', 'sum')
                            ).reset_index()
                            
                            weekly_item['週次'] = pd.to_datetime(weekly_item['week_start']).dt.strftime('W%V\n%m/%d')
                            weekly_item['每客平均消耗量'] = weekly_item.apply(
                                lambda r: r['總採購量'] / r['來客人數'] if r['來客人數'] > 0 else 0, axis=1
                            )
                            
                            # 計算月平均與平均單價
                            total_qty_month = weekly_item['總採購量'].sum()
                            total_guests_month = weekly_item['來客人數'].sum()
                            avg_rate_month = total_qty_month / total_guests_month if total_guests_month > 0 else 0
                            avg_unit_price = item_df['cleaned_total'].sum() / item_df['cleaned_qty'].sum() if item_df['cleaned_qty'].sum() > 0 else 0
                            
                            st.write("")
                            st.markdown(f"##### 📊 **「{search_term}」消耗數據指標**")
                            
                            c_m1, c_m2, c_m3 = st.columns(3)
                            c_m1.metric("本月總採購量", f"{total_qty_month:,.1f} {most_common_unit}")
                            c_m2.metric("每客平均消耗量 (使用率)", f"{avg_rate_month:.2f} {most_common_unit}/人", help="總採購量 / 總來客數")
                            c_m3.metric("平均採購單價", f"NT$ {avg_unit_price:,.1f} /{most_common_unit}")
                            
                            # 圖表呈現
                            st.write("")
                            st.markdown(f"###### 📈 週來客數 vs 「{search_term}」採購量相對走勢")
                            
                            max_w_qty = weekly_item['總採購量'].max()
                            max_w_guests = weekly_item['來客人數'].max()
                            weekly_item['採購量(%)'] = (weekly_item['總採購量'] / max_w_qty * 100).round(1) if max_w_qty > 0 else 0
                            weekly_item['來客(%)'] = (weekly_item['來客人數'] / max_w_guests * 100).round(1) if max_w_guests > 0 else 0
                            
                            melt_item_df = weekly_item.melt(
                                id_vars=['週次', '總採購量', '來客人數', '總費用'],
                                value_vars=['採購量(%)', '來客(%)'],
                                var_name='指標', value_name='標準化數值'
                            )
                            
                            item_color_map = {'採購量(%)': '#e67e22', '來客(%)': '#2980b9'}
                            
                            item_chart = alt.Chart(melt_item_df).mark_line(point=True, strokeWidth=2.5).encode(
                                x=alt.X('週次:N', title='週次', sort=None),
                                y=alt.Y('標準化數值:Q', title='相對比例 (% of max)', scale=alt.Scale(domain=[0, 110])),
                                color=alt.Color('指標:N',
                                    scale=alt.Scale(domain=list(item_color_map.keys()), range=list(item_color_map.values())),
                                    legend=alt.Legend(title='指標', orient='bottom')
                                ),
                                tooltip=[
                                    alt.Tooltip('週次:N', title='週次'),
                                    alt.Tooltip('指標:N', title='指標'),
                                    alt.Tooltip('總採購量:Q', title=f'總採購量 ({most_common_unit})', format=',.1f'),
                                    alt.Tooltip('來客人數:Q', title='來客人數 (人)', format=',.0f'),
                                    alt.Tooltip('總費用:Q', title='總費用 (NT$)', format=',.0f'),
                                ]
                            ).properties(height=200)
                            
                            st.altair_chart(item_chart, use_container_width=True)
                            
                            # 🔮 採購方案精算
                            st.write("")
                            st.markdown(f"##### 🔮 「{search_term}」精準採購方案預算機")
                            st.write("設定您未來的預計來客數，系統會自動幫您推算最合理的採購量與叫貨時程建議。")
                            
                            col_calc1, col_calc2 = st.columns([1, 1])
                            with col_calc1:
                                input_guests = st.number_input(
                                    "📅 未來一週預計總來客數",
                                    min_value=10,
                                    max_value=5000,
                                    value=int(total_guests_month / 4) if total_guests_month > 0 else 500,
                                    step=50,
                                    key="item_calc_guests_input_widget"
                                )
                                
                                # 分類叫貨週期提示
                                vendor_type = "菜商"
                                if any(x in search_term for x in ["蛋", "卵"]):
                                    vendor_type = "蛋商"
                                elif any(x in search_term for x in ["肉", "雞", "豬", "牛", "魚"]):
                                    vendor_type = "肉商"
                                elif any(x in search_term for x in ["雜", "油", "米", "麵"]):
                                    vendor_type = "雜貨"
                                
                                st.info(
                                    f"💡 **建議配比原則 ({vendor_type})**\n\n"
                                    f"- 自動配比可防止單次進貨量過大導致新鮮度下降或報廢損耗。\n"
                                    f"- 可依現行實際叫貨週期彈性調整叫貨。"
                                )
                                
                            with col_calc2:
                                # 包含 5% 安全庫存緩衝
                                recommended_qty = input_guests * avg_rate_month * 1.05
                                est_cost = recommended_qty * avg_unit_price
                                
                                st.markdown(
                                    f"<div style='background:#2e437c15; border-left:4px solid #2e437c; padding:15px; border-radius:8px;'>"
                                    f"<h4 style='margin:0; color:#2e437c;'>建議採購總量</h4>"
                                    f"<h2 style='margin:5px 0; color:#2e437c;'>{recommended_qty:,.1f} {most_common_unit}</h2>"
                                    f"<p style='margin:0; font-size:12px; color:#666;'>已包含 5% 安全庫存緩衝</p>"
                                    f"<hr style='margin:10px 0; border:none; border-top:1px solid #ddd;'>"
                                    f"<h5 style='margin:0; color:#333;'>預估採購費用: <strong style='font-size:18px;'>NT$ {int(est_cost):,}</strong></h5>"
                                    f"</div>",
                                    unsafe_allow_html=True
                                )
                                
                                # 分週配送週期比例建議
                                st.markdown("📋 **叫貨週期配送配比推薦**")
                                if vendor_type == "蛋商":
                                    st.markdown("- **週一 (40%)**：建議採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(recommended_qty * 0.4, most_common_unit, int(est_cost * 0.4)))
                                    st.markdown("- **週三 (30%)**：建議採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(recommended_qty * 0.3, most_common_unit, int(est_cost * 0.3)))
                                    st.markdown("- **週五 (30%)**：建議採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(recommended_qty * 0.3, most_common_unit, int(est_cost * 0.3)))
                                elif vendor_type == "菜商":
                                    st.markdown("- **平日每日均攤 (60%)**：每次到貨建議 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(recommended_qty * 0.12, most_common_unit, int(est_cost * 0.12)))
                                    st.markdown("- **週五加強 (40%)**：一次叫足 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(recommended_qty * 0.4, most_common_unit, int(est_cost * 0.4)))
                                else:
                                    st.markdown("- **單次足額採購 (100%)**：於週一或合約到貨日一次性採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(recommended_qty, most_common_unit, int(est_cost)))
                                    
                        else:
                            st.warning(f"⚠️ 在目前的採購資料中，找不到含有「{search_term}」的品項名稱。")
                            st.info("💡 請嘗試選擇其他常用關鍵字，或自訂輸入更精確的關鍵字（如：雞蛋、高麗菜）。")
                    else:
                        st.info("💡 採購分頁缺少「品名」或「數量」欄位，無法進行單品消耗率分析。")
                        
            else:
                st.info(f"💡 {current_month_str} 尚未有採購數據紀錄。")
                st.write(f"ℹ️ 在「**{used_name}**」分頁中總共發現 {len(df_purchase)} 筆資料，但沒有符合 {current_month_str} 的紀錄。")
                with st.expander("🛠️ 點此查看分頁中的前 5 筆原始資料 (除錯用)"):
                    st.write(df_purchase.head(5))
        else:
            st.warning(f"⚠️ 無法在 Google Sheet 中找到採購分頁 (嘗試過: {', '.join(possible_names)})。")
            st.info("💡 請確認分頁名稱是否正確，且分頁中至少已填入一行資料。")
            
    except Exception as e:
        if "WorksheetNotFound" in str(e):
             st.error(f"❌ 找不到採購相關分頁！請確認 Google Sheet 中的分頁名稱（如 purchase data）。")
        else:
            st.error(f"讀取採購數據出錯: {e}")
        import traceback
        st.expander("錯誤詳細資訊").code(traceback.format_exc())

# --- 📅 本月接下來各週採購金額建議（獨立區塊，不依賴採購數據）---
with tab_p:
    st.divider()
    st.markdown("#### 📅 本月接下來各週採購金額建議")
    
    from datetime import date as dt_date, timedelta as dt_timedelta
    import calendar as cal_lib
    today_dt2 = dt_date.today()
    _, last_day_num2 = cal_lib.monthrange(selected_date.year, selected_date.month)
    month_end_dt2 = dt_date(selected_date.year, selected_date.month, last_day_num2)
    
    is_cur_or_fut = (selected_date.year, selected_date.month) >= (today_dt2.year, today_dt2.month)
    
    if is_cur_or_fut:
        # 1. 優先用本月餐廳數據
        fw_m_data = fetch_month_summary(selected_date.year, selected_date.month)
        fw_df = fw_m_data.get('df', pd.DataFrame())
        avg_fw = 0
        fw_label = ''
        if not fw_df.empty and 'bf_total_act' in fw_df.columns:
            fw_df = fw_df.copy()
            fw_df['_bf'] = pd.to_numeric(fw_df['bf_total_act'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
            active_fw = fw_df[fw_df['_bf'] > 0]['_bf']
            if not active_fw.empty:
                avg_fw = active_fw.mean()
                fw_label = f"本月實際 ({len(active_fw)} 天記錄)"
        
        # 2. 備援：改用上個月早餐來客數
        if avg_fw == 0:
            fw_prev = fetch_month_summary(m_prev['year'], m_prev['month']) if 'year' in m_prev else {}
            fw_prev_df = fw_prev.get('df', pd.DataFrame()) if fw_prev else pd.DataFrame()
            if not fw_prev_df.empty and 'bf_total_act' in fw_prev_df.columns:
                fw_prev_df = fw_prev_df.copy()
                fw_prev_df['_bf'] = pd.to_numeric(fw_prev_df['bf_total_act'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                prev_active = fw_prev_df[fw_prev_df['_bf'] > 0]['_bf']
                if not prev_active.empty:
                    avg_fw = prev_active.mean()
                    fw_label = f"⚠️ 以上月平均推估（本月尚無餐廳資料）"
        
        # 3. 雙冠日清單（來自 tab_m 的 calc_key_metrics）
        fw_curr_metrics = calc_key_metrics(fw_m_data)
        fw_dual_dates = set(fw_curr_metrics.get('dual_match_dates', []))
        
        # 4. 逐週生成
        fw_week_plans = []
        fw_cursor = today_dt2
        fw_seen = set()
        while fw_cursor <= month_end_dt2:
            mon = fw_cursor - dt_timedelta(days=fw_cursor.weekday())
            sun = mon + dt_timedelta(days=6)
            ws = max(mon, dt_date(selected_date.year, selected_date.month, 1))
            we = min(sun, month_end_dt2)
            if ws not in fw_seen and we >= today_dt2:
                fw_seen.add(ws)
                wdates = [(ws + dt_timedelta(days=i)).strftime('%Y-%m-%d') for i in range((we - ws).days + 1)]
                has_d = any(d in fw_dual_dates for d in wdates)
                days_cnt = len(wdates)
                fw_week_plans.append({
                    'label': f"{ws.strftime('%m/%d')} ～ {we.strftime('%m/%d')}",
                    'has_dual': has_d,
                    'recommended': int((150 * 1.15 if has_d else 150) * avg_fw * days_cnt),
                    'dual_labels': [d[5:] for d in wdates if d in fw_dual_dates],
                    'days_cnt': days_cnt,
                })
            fw_cursor = sun + dt_timedelta(days=1)
        
        if avg_fw > 0 and fw_week_plans:
            st.caption(f"💡 預估基準：每日平均來客數 **{avg_fw:.1f} 人**（{fw_label}）。雙冠週採購上限自動提高 15%。")
            for wp in fw_week_plans:
                color = '#e67e22' if wp['has_dual'] else '#2980b9'
                dual_note = f"　🎯 含雙冠日：{', '.join(wp['dual_labels'])}" if wp['has_dual'] else ""
                c1, c2 = st.columns([2, 1])
                c1.markdown(f"**{wp['label']}**{dual_note}")
                c2.markdown(
                    f"<div style='background:{color}22; border-left:3px solid {color}; padding:8px 12px; border-radius:6px; text-align:center;'>"
                    f"<strong style='font-size:16px;'>NT$ {wp['recommended']:,}</strong>"
                    f"<br><span style='font-size:11px; color:#666;'>建議週採購上限 ({wp['days_cnt']}天)</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
        else:
            st.info("💡 尚無足夠的來客數據來估算週採購預算。請確認上個月的餐廳早餐來客數已填寫。")
    else:
        st.info("💡 「週採購建議」僅適用於當月或未來月份。")

with tab7:
    st.header("👥 人事概況")
    
    # -- 人事管理函數 (Google Sheets 版) --
    def get_all_employees():
        try:
            df = conn.read(worksheet="employees", ttl="1m")
            return df if df is not None else pd.DataFrame()
        except:
            return pd.DataFrame()

    def add_employee(e_id, name, dept, pos, salary):
        try:
            df = conn.read(worksheet="employees", ttl="0")
            required_cols = ["employee_id", "name", "dept", "position", "salary"]
            
            if df is None or df.empty or not all(c in df.columns for c in required_cols):
                if df is None or df.empty:
                    df = pd.DataFrame(columns=required_cols)
                else:
                    for c in required_cols:
                        if c not in df.columns:
                            df[c] = ""
                            
            if str(e_id) in df['employee_id'].astype(str).values:
                return "ID_EXISTS"
                
            new_emp = pd.DataFrame([{"employee_id": str(e_id), "name": name, "dept": dept, "position": pos, "salary": salary}])
            df = pd.concat([df, new_emp], ignore_index=True)
            conn.update(worksheet="employees", data=df.fillna(""))
            return True
        except Exception as e:
            return str(e)

    def delete_employee(e_id):
        try:
            df = conn.read(worksheet="employees", ttl="0")
            if df is not None and not df.empty and 'employee_id' in df.columns:
                df['employee_id'] = df['employee_id'].astype(str)
                df = df[df['employee_id'] != str(e_id)]
                conn.update(worksheet="employees", data=df.fillna(""))
        except:
            pass

    # -- UI: 新增員工區 --
    with st.expander("➕ 新增新進員工資訊", expanded=False):
        with st.form("add_employee_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            new_id = col1.text_input("員工編號 (必填)")
            new_name = col2.text_input("姓名 (必填)")
            
            new_dept = st.selectbox("所屬部門", ["路徒Plus行旅站前館", "櫃檯", "房務", "工務", "The Peak"])
            new_pos = st.text_input("職位")
            new_salary = st.number_input("薪資", min_value=0, step=1000)
            
            submit_btn = st.form_submit_button("✅ 確認新增")
            if submit_btn:
                if not new_id or not new_name:
                    st.error("❌ 請填寫員工編號與姓名！")
                else:
                    res = add_employee(new_id, new_name, new_dept, new_pos, new_salary)
                    if res == True:
                        st.success(f"✅ 成功新增員工：{new_name}")
                        st.rerun()
                    elif res == "ID_EXISTS":
                        st.error("❌ 員工編號已存在，請檢查是否重覆。")
                    else:
                        st.error(f"❌ 新增失敗：{res}")

    st.divider()

    # -- UI: 員工列表與排序 --
    df_emp = get_all_employees()
    
    # 過濾掉空白的資料列 (如 Google Sheets 常見的結尾空白行)
    if not df_emp.empty and 'employee_id' in df_emp.columns:
        df_emp['employee_id'] = df_emp['employee_id'].astype(str).str.strip()
        # 移除 pandas 自動將數字轉為 float 所產生的 .0 結尾
        df_emp['employee_id'] = df_emp['employee_id'].str.replace(r'\.0$', '', regex=True)
        df_emp = df_emp[df_emp['employee_id'] != '']
        df_emp = df_emp[df_emp['employee_id'].str.lower() != 'nan']
    
    if df_emp.empty:
        st.info("💡 目前資料庫中尚無員工資訊。")
    else:
        # 計算總薪資 (排除職位為 PT 的人)
        # 確保 position 欄位存在且處理大小寫
        if 'position' in df_emp.columns:
            non_pt_df = df_emp[df_emp['position'].fillna('').astype(str).str.upper() != 'PT']
            total_salary = pd.to_numeric(non_pt_df['salary'], errors='coerce').fillna(0).sum()
        else:
            total_salary = pd.to_numeric(df_emp['salary'], errors='coerce').fillna(0).sum()

        st.markdown(f"""
        <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; border-left: 5px solid #4CAF50; margin-bottom: 20px;">
            <p style="margin: 0; font-size: 14px; color: #666;">💰 正職員工薪資總計</p>
            <h2 style="margin: 0; color: #2e437c;">NT$ {int(total_salary):,}</h2>
            <p style="margin: 5px 0 0 0; font-size: 12px; color: #999;">* 已自動排除職位名稱為 "PT" 的人員數據</p>
        </div>
        """, unsafe_allow_html=True)

        col_sort, col_search = st.columns([1, 1])
        sort_opt = col_sort.selectbox("排序方式", ["員工編號順序", "薪資 (由高到低)", "薪資 (由低到高)", "按部門排序"])
        search_query = col_search.text_input("🔍 搜尋姓名或編號")

        # 搜尋過濾
        if search_query:
            df_emp = df_emp[df_emp['name'].astype(str).str.contains(search_query, case=False) | df_emp['employee_id'].str.contains(search_query, case=False)]

        # 確保 salary 為數值以便排序
        if 'salary' in df_emp.columns:
            df_emp['salary'] = pd.to_numeric(df_emp['salary'], errors='coerce').fillna(0)

        # 排序邏輯
        if sort_opt == "員工編號順序":
            df_emp = df_emp.sort_values("employee_id")
        elif sort_opt == "薪資 (由高到低)":
            if 'salary' in df_emp.columns: df_emp = df_emp.sort_values("salary", ascending=False)
        elif sort_opt == "薪資 (由低到高)":
            if 'salary' in df_emp.columns: df_emp = df_emp.sort_values("salary", ascending=True)
        elif sort_opt == "按部門排序":
            if 'dept' in df_emp.columns: df_emp = df_emp.sort_values(["dept", "employee_id"])

        # 自定義表格顯示
        st.write(f"📊 目前共有 {len(df_emp)} 位員工")
        
        # 標題列
        header_cols = st.columns([1.5, 1.5, 1.5, 1.5, 1.5, 1])
        header_cols[0].markdown("**員工編號**")
        header_cols[1].markdown("**姓名**")
        header_cols[2].markdown("**部門**")
        header_cols[3].markdown("**職位**")
        header_cols[4].markdown("**薪資**")
        header_cols[5].markdown("**操作**")
        
        st.divider()
        
        for idx, row in df_emp.iterrows():
            row_cols = st.columns([1.5, 1.5, 1.5, 1.5, 1.5, 1])
            row_cols[0].write(row.get('employee_id', ''))
            row_cols[1].write(row.get('name', ''))
            row_cols[2].write(row.get('dept', ''))
            row_cols[3].write(row.get('position', ''))
            
            salary_val = row.get('salary', 0)
            try: salary_int = int(float(salary_val))
            except: salary_int = 0
            row_cols[4].write(f"NT$ {salary_int:,}")
            
            # 使用 idx 來保證 key 絕對唯一，避免 StreamlitDuplicateElementKey
            if row_cols[5].button("🗑️", key=f"del_{idx}_{row.get('employee_id', '')}", help="刪除此員工"):
                delete_employee(row.get('employee_id', ''))
                st.toast(f"已刪除員工: {row['name']}")
                time.sleep(0.5)
                st.rerun()
