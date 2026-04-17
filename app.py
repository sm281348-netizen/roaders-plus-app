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
def get_daily_data(d_str):
    try:
        # 讀取完整表單 (快取設定為 1 分鐘)
        df = conn.read(worksheet="daily_data", ttl="1m")
        if df is not None and not df.empty:
            res = df[df['date'] == d_str]
            if not res.empty:
                data_dict = res.iloc[0].to_dict()
                numeric_cols = [
                    'occ_rate', 'adr', 'revenue', 'total_rooms', 'counter_expense', 
                    'cleaned_rooms', 'hk_checkout_extend', 'hk_avg_clean', 'hk_expense',
                    'rest_breakfast', 'rest_month_guests', 'rest_day_guests', 'rest_avg_guests',
                    'rest_month_rev', 'rest_avg_spent', 'rest_peak_expense',
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
        
        data_dict['date'] = d_str
        new_row = pd.DataFrame([data_dict])
        
        if 'date' in df.columns and d_str in df['date'].values:
            df = df[df['date'] != d_str]
        
        df = pd.concat([df, new_row], ignore_index=True)
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
        
        # 確保欄位存在
        if 'date' not in df.columns or 'log' not in df.columns:
            df = pd.DataFrame(columns=["date", "log"])

        new_row = pd.DataFrame([{'date': d_str, 'log': log_text}])
        
        if d_str in df['date'].values:
            df = df[df['date'] != d_str]
        
        df = pd.concat([df, new_row], ignore_index=True)
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
    'input_rest_mrev': ('rest_month_rev', 0), 'input_rest_aspent': ('rest_avg_spent', 0), 'input_rest_exp': ('rest_peak_expense', 0), 'input_rest_car': ('rest_car_data', ""),
    'input_repair': ('maint_repair_rooms', 0), 'input_maint_rec': ('maint_records', ""), 'input_maint_exp': ('maint_expense', 0)
}

def sync_st_to_db(target_d_str):
    # 同步數值數據
    update_dict = {db_col: st.session_state[ss_key] for ss_key, (db_col, _) in field_mapping.items() if ss_key in st.session_state}
    if update_dict:
        save_daily_data(target_d_str, update_dict)
    
    # 單獨同步日誌
    if 'input_daily_log' in st.session_state:
        save_daily_log(target_d_str, st.session_state['input_daily_log'])

def prev_day():
    st.session_state['sidebar_date'] -= datetime.timedelta(days=1)

def next_day():
    st.session_state['sidebar_date'] += datetime.timedelta(days=1)

col1, col2 = st.sidebar.columns(2)
col1.button("⬅️ 前一天", on_click=prev_day)
col2.button("後一天 ➡️", on_click=next_day)

selected_date = st.sidebar.date_input("選擇日期", value=st.session_state['sidebar_date'], key='sidebar_date')
date_str = str(selected_date)

# --- 核心修復：檢測日期切換並自動強制存檔舊日期 ---
# 追蹤當前正在編輯的日期
if '_actual_current_date' not in st.session_state:
    st.session_state['_actual_current_date'] = date_str
# 追蹤當前 session_state 內容是否已經從資料庫載入完成 (防止存入預設的 0)
if '_data_is_loaded' not in st.session_state:
    st.session_state['_data_is_loaded'] = False

if st.session_state['_actual_current_date'] != date_str:
    # 只有在「確定已經載入過舊日期資料」的情況下，才在切換時存檔
    if st.session_state.get('_data_is_loaded', False):
        sync_st_to_db(st.session_state['_actual_current_date'])
    
    # 切換日期標記，並重設載入狀態（因為新日期的資料還沒讀取）
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
            df_new = pd.DataFrame(updates)
            
            # 合併數據 (以日期為 key，部分更新)
            df_new = df_new.set_index('date')
            if not df_existing.empty:
                df_existing = df_existing.set_index('date')
                df_final = df_new.combine_first(df_existing).reset_index()
            else:
                df_final = df_new.reset_index()
                
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
        df = pd.read_excel(file, header=None)
        
        month_rev = 0
        avg_spent = 0
        
        # 尋找底部的月結算資料
        for i, row in df.iterrows():
            col0 = str(row[0]).strip()
            if '已結算營收' in col0 and '早餐' not in col0 and '下午茶' not in col0:
                for val in row[1:]:
                    if pd.notna(val) and str(val).strip() != '':
                        month_rev = int(float(str(val).replace('NT$', '').replace(',', '').strip()))
                        break
            if '平均客單價' == col0:
                for val in row[1:]:
                    if pd.notna(val) and str(val).strip() != '':
                        avg_spent = int(float(str(val).replace('NT$', '').replace(',', '').strip()))
                        break

        parsed_days = []
        for i, row in df.iterrows():
            col0 = str(row[0]).strip()
            m = re.match(r'^(\d{1,2})/(\d{1,2})', col0)
            if m:
                month_val, day_val = m.groups()
                d_str = f"{current_year}-{int(month_val):02d}-{int(day_val):02d}"
                
                def safe_int(val):
                    try: 
                        if pd.isna(val): return 0
                        return int(float(str(val).replace(',', '').strip()))
                    except: return 0
                        
                bf_theme_est = safe_int(row[1]) if len(row) > 1 else 0
                bf_theme_act = safe_int(row[2]) if len(row) > 2 else 0
                bf_zq_est = safe_int(row[3]) if len(row) > 3 else 0
                bf_zq_act = safe_int(row[4]) if len(row) > 4 else 0
                bf_total_est = safe_int(row[5]) if len(row) > 5 else 0
                bf_total_act = safe_int(row[6]) if len(row) > 6 else 0
                
                af_theme_est = safe_int(row[7]) if len(row) > 7 else 0
                af_theme_act = safe_int(row[8]) if len(row) > 8 else 0
                af_zq_est = safe_int(row[9]) if len(row) > 9 else 0
                af_zq_act = safe_int(row[10]) if len(row) > 10 else 0
                af_total_est = safe_int(row[11]) if len(row) > 11 else 0
                af_total_act = safe_int(row[12]) if len(row) > 12 else 0
                
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
            df_existing = conn.read(worksheet="daily_data", ttl="0")
            if df_existing is None: df_existing = pd.DataFrame()
            df_new = pd.DataFrame(parsed_days)
            
            # 合併數據 (以日期為 key，部分更新)
            df_new = df_new.set_index('date')
            if not df_existing.empty:
                df_existing = df_existing.set_index('date')
                df_final = df_new.combine_first(df_existing).reset_index()
            else:
                df_final = df_new.reset_index()
                
            conn.update(worksheet="daily_data", data=df_final.fillna(""))
            st.cache_data.clear()
            
        st.session_state['_last_loaded_date'] = None
        return len(parsed_days)
    except Exception as e:
        import traceback
        st.error(f"解析餐廳報表失敗: {e}\n{traceback.format_exc()}")
        return False

# 頁面標題
st.title("路徒Plus行旅站前館營運日誌")

# 主畫面
tab1, tab_m, tab6, tab_p, tab2, tab3, tab4, tab5, tab7 = st.tabs(["📊 營運總覽", "📈 月分析專區", "📝 每日營運紀錄", "💰 採購分析", "💼 櫃台數據", "🧹 房務數據", "🍽️ 餐廳數據", "🔧 工務數據", "👥 人事概況"])

with tab2:
    st.header("💼 櫃台數據")
    st.subheader("📁 數據報表上傳")
    jinxu_file = st.file_uploader("上傳金旭報表 (Excel/CSV)，會自動把整份報表寫入資料庫！", type=["csv", "xls", "xlsx"], key="jinxu_uploader")
    
    if jinxu_file:
        if st.button("📥 寫入系統資料庫"):
            saved_count = parse_and_save_jinxu(jinxu_file)
            if saved_count:
                st.success(f"✅ 成功將 {saved_count} 筆每日資料存入系統資料庫！切換日期即可自動調出。")
                time.sleep(1)
                st.rerun()

    st.divider()
    st.subheader(f"櫃台手動確認區 ({date_str})")
    st.number_input("訂房率 (%)", min_value=0.0, max_value=100.0, step=0.1, key="input_occ", on_change=on_input_change)
    st.number_input("總營收", min_value=0, step=100, key="input_rev", on_change=on_input_change)
    st.number_input("ADR (平均房價)", min_value=0, step=10, key="input_adr", on_change=on_input_change)
    st.number_input("總住房數", min_value=0, step=1, key="input_rooms", on_change=on_input_change)
    st.text_area("負評客訴", key="input_complaints", on_change=on_input_change)
    st.number_input("櫃台請購費用", min_value=0, step=100, key="input_counter_exp", on_change=on_input_change)

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

    # -- 月度累計模式 (MTD Analysis) --
    st.subheader(f"📅 本月累計分析 (MTD: {selected_date.strftime('%Y-%m')})")
    start_of_month = selected_date.replace(day=1).strftime('%Y-%m-%d')
    
    try:
        df_all = conn.read(worksheet="daily_data", ttl="10m")
        if df_all is not None and not df_all.empty:
            df_mtd = df_all[(df_all['date'] >= start_of_month) & (df_all['date'] <= date_str)].copy()
        else:
            df_mtd = pd.DataFrame()
    except:
        df_mtd = pd.DataFrame()

    if not df_mtd.empty:
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
        
        m1, m2, m3 = st.columns(3)
        m1.markdown(make_card("MTD 累計住房率", f"{mtd_occ:.1f}%", "card-theme-blue", "card-bg-dark", "🏨"), unsafe_allow_html=True)
        m2.markdown(make_card("MTD 累計 ADR", f"NT$ {int(mtd_adr):,}", "card-theme-green", "card-bg-dark", "💳"), unsafe_allow_html=True)
        m3.markdown(make_card("MTD 累計總營收", f"NT$ {int(mtd_rev):,}", "card-theme-orange", "card-bg-dark", "💰"), unsafe_allow_html=True)
        
        st.markdown("<br><hr style='margin: 5px 0; border: 1px dashed #ddd;'>", unsafe_allow_html=True)
        st.write("##### 🍽️ 餐廳營運累計 (MTD)")
        
        # MTD 餐廳計算
        mtd_bf_theme = df_mtd['bf_theme_act'].fillna(0).sum() if 'bf_theme_act' in df_mtd.columns else 0
        mtd_bf_zq = df_mtd['bf_zq_act'].fillna(0).sum() if 'bf_zq_act' in df_mtd.columns else 0
        mtd_af_theme = df_mtd['af_theme_act'].fillna(0).sum() if 'af_theme_act' in df_mtd.columns else 0
        mtd_af_zq = df_mtd['af_zq_act'].fillna(0).sum() if 'af_zq_act' in df_mtd.columns else 0
        
        # 本月整體總和
        mtd_total_bf_act = df_mtd['bf_total_act'].fillna(0).sum() if 'bf_total_act' in df_mtd.columns else 0
        mtd_total_af_act = df_mtd['af_total_act'].fillna(0).sum() if 'af_total_act' in df_mtd.columns else 0
        
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
        
        rest_month_rev = df_mtd['rest_month_rev'].fillna(0).max() if 'rest_month_rev' in df_mtd.columns else 0
        rest_avg_spent = df_mtd['rest_avg_spent'].fillna(0).max() if 'rest_avg_spent' in df_mtd.columns else 0
        
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
    
    # --- A. 每日住房率概況 (四個月對比) ---
    st.subheader("📊 每日住房率概況比較 (四個月)")
    col_chart1, col_chart2, col_chart3, col_chart4 = st.columns(4)
    
    def render_occ_chart(month_data, title_suffix):
        df = month_data['df'].copy()
        if df.empty:
            st.info(f"💡 {month_data['month_label']} 尚無數據。")
            return
        df['day'] = pd.to_datetime(df['date']).dt.day
        
        # 建立 Altair 圖表
        base = alt.Chart(df).encode(
            x=alt.X('day:O', title='日期', axis=alt.Axis(labelAngle=0)),
            tooltip=['date', 'occ_rate']
        )
        
        bars = base.mark_bar().encode(
            y=alt.Y('occ_rate:Q', title='住房率 (%)', scale=alt.Scale(domain=[0, 100])),
            color=alt.condition(alt.datum.occ_rate >= 90, alt.value('#e74c3c'), alt.value('#3498db'))
        )
        
        # 新增文字標籤 (固定顯示在長條上方)
        text = base.mark_text(
            align='center',
            baseline='bottom',
            dy=-5,
            fontSize=14,
            fontWeight='bold'
        ).encode(
            y=alt.Y('occ_rate:Q'),
            text=alt.Text('occ_rate:Q', format='.1f')
        )
        
        chart = (bars + text).properties(title=f"{month_data['month_label']} {title_suffix}", height=300)
        st.altair_chart(chart, use_container_width=True)

    with col_chart1: render_occ_chart(m_prev_prev, "(前前月)")
    with col_chart2: render_occ_chart(m_prev, "(上月)")
    with col_chart3: render_occ_chart(m_curr, "(本月)")
    with col_chart4: render_occ_chart(m_next, "(下月)")
    
    # --- B. 每日住房率 - 關鍵差異 ---
    st.markdown("#### 🔍 每日住房率：關鍵差異分析 (與各月比對)")
    diff_curr_prev_prev = m_curr['occ90_days'] - m_prev_prev['occ90_days']
    diff_curr_prev = m_curr['occ90_days'] - m_prev['occ90_days']
    diff_curr_next = m_curr['occ90_days'] - m_next['occ90_days']
    
    def occ_diff_card(label, diff, target_label):
        color = '#2ecc71' if diff >= 0 else '#e74c3c'
        status = '本月多' if diff > 0 else '較少' if diff < 0 else '持平'
        return f'<div style="flex: 1; min-width: 150px; background: #fff; padding: 12px; border-radius: 8px; border: 1px solid #eee; margin-bottom: 10px;"><p style="margin:0; font-size:12px; color:#999;">與 {target_label} 相比</p><div style="display: flex; align-items: baseline; gap: 8px; margin-top: 5px;"><strong style="font-size:18px; color:{color};">{abs(diff)} 天</strong><span style="font-size:11px; color:#666;">({status})</span></div></div>'

    html_content = f"""
    <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #3498db; margin-bottom: 20px;">
        <p style="margin:0; font-size:14px; color:#666;">📊 <strong>達 90% 住房率天數比對 (本月: {m_curr['occ90_days']} 天)</strong></p>
        <div style="display: flex; gap: 15px; margin-top: 10px; flex-wrap: wrap;">
            {occ_diff_card("前前月", diff_curr_prev_prev, m_prev_prev['month_label'])}
            {occ_diff_card("上月", diff_curr_prev, m_prev['month_label'])}
            {occ_diff_card("下月預期", diff_curr_next, m_next['month_label'])}
        </div>
    </div>
    """
    st.write(html_content, unsafe_allow_html=True)
    
    st.divider()
    
    # --- C. 月度營運指標 (四個月對比) ---
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

    st.divider()

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
        if st.button("📥 寫入系統資料庫", key="rest_btn"):
            saved_count = parse_and_save_restaurant(rest_file, selected_date.year)
            if saved_count:
                st.success(f"✅ 成功將 {saved_count} 筆每日餐廳資料存入系統資料庫！切換日期即可自動調出。")
                time.sleep(1)
                st.rerun()

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
    st.text_area("4樓餐車數據", key="input_rest_car", on_change=on_input_change)

with tab5:
    st.header("🔧 工務數據")
    st.number_input("今日待修房數", min_value=0, step=1, key="input_repair", on_change=on_input_change)
    st.text_area("修繕紀錄", key="input_maint_rec", on_change=on_input_change)
    st.number_input("工務請購費用", min_value=0, step=100, key="input_maint_exp", on_change=on_input_change)

with tab6:
    st.header("📝 每日營運紀錄")
    
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
            
            # 過濾 NaT/None
            df_purchase = df_purchase[df_purchase['日期'].notna()]
            
            # 過濾當月數據
            m_start = selected_date.replace(day=1)
            import calendar
            _, last_day = calendar.monthrange(selected_date.year, selected_date.month)
            m_end = selected_date.replace(day=last_day)
            
            df_month = df_purchase[(df_purchase['日期'] >= m_start) & (df_purchase['日期'] <= m_end)].copy()
            
            if not df_month.empty:
                # 數值清理
                df_month['小計'] = pd.to_numeric(df_month[total_col], errors='coerce').fillna(0)
                # 其他欄位清理 (如果存在的話)
                price_col = next((c for c in df_month.columns if '單價' in c or 'Price' in c), None)
                qty_col = next((c for c in df_month.columns if '數量' in c or 'Qty' in c), None)
                
                if price_col: df_month['單價'] = pd.to_numeric(df_month[price_col], errors='coerce').fillna(0)
                if qty_col: df_month['數量'] = pd.to_numeric(df_month[qty_col], errors='coerce').fillna(0)
                
                total_month_expense = df_month['小計'].sum()
                
                # 1. 本月總開銷
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #1f2c56 0%, #2e437c 100%); padding: 25px; border-radius: 15px; text-align: center; color: white; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
                    <p style="margin: 0; font-size: 1.1rem; opacity: 0.8;">📅 {current_month_str} 本月總開銷金額</p>
                    <h1 style="margin: 10px 0 0 0; font-size: 3rem; font-weight: 800; letter-spacing: 1px;">NT$ {int(total_month_expense):,}</h1>
                </div>
                """, unsafe_allow_html=True)
                
                # 2. 部門佔比圓餅圖
                st.subheader("📊 各部門請購佔比分析")
                dept_summary = df_month.groupby(dept_col)['小計'].sum().reset_index()
                # 統一重命名方便繪圖與顯示
                dept_summary.columns = ['部門', '小計']
                
                # 繪製圓餅圖
                pie_chart = alt.Chart(dept_summary).mark_arc(innerRadius=60, stroke="#fff").encode(
                    theta=alt.Theta(field="小計", type="quantitative"),
                    color=alt.Color(field="部門", type="nominal", scale=alt.Scale(scheme='category10'), legend=alt.Legend(title="部門")),
                    tooltip=["部門", alt.Tooltip("小計", format=",.0f", title="總金額 (NT$)")]
                ).properties(height=400)
                
                st.altair_chart(pie_chart, use_container_width=True)
                
                st.divider()
                
                # 3. 各部門詳細統計
                st.subheader("🏢 各部門經費分析")
                
                # 取得所有部門
                departments = dept_summary.sort_values('小計', ascending=False)['部門'].tolist()
                
                for dept in departments:
                    dept_df = df_month[df_month[dept_col] == dept]
                    dept_total = dept_df['小計'].sum()
                    
                    with st.expander(f"📌 {dept} (總計: NT$ {int(dept_total):,})", expanded=False):
                        # 顯示該部門表格 (動態選擇要顯示的欄位)
                        cols_to_show = [c for c in ['日期', '供應商', '品名', '規格', '數量', '單位', '單價', '小計'] if c in dept_df.columns]
                        if not cols_to_show:
                             cols_to_show = dept_df.columns.tolist()
                             
                        st.dataframe(
                            dept_df[cols_to_show].sort_values('日期'),
                            use_container_width=True,
                            hide_index=True
                        )
                        
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
            if df is None: df = pd.DataFrame(columns=["employee_id", "name", "dept", "position", "salary"])
            
            if e_id in df['employee_id'].values:
                return "ID_EXISTS"
                
            new_emp = pd.DataFrame([{"employee_id": e_id, "name": name, "dept": dept, "position": pos, "salary": salary}])
            df = pd.concat([df, new_emp], ignore_index=True)
            conn.update(worksheet="employees", data=df.fillna(""))
            return True
        except Exception as e:
            return str(e)

    def delete_employee(e_id):
        try:
            df = conn.read(worksheet="employees", ttl="0")
            if df is not None:
                df = df[df['employee_id'] != e_id]
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
    
    if df_emp.empty:
        st.info("💡 目前資料庫中尚無員工資訊。")
    else:
        # 計算總薪資 (排除職位為 PT 的人)
        # 確保 position 欄位存在且處理大小寫
        if 'position' in df_emp.columns:
            non_pt_df = df_emp[df_emp['position'].fillna('').str.upper() != 'PT']
            total_salary = non_pt_df['salary'].sum()
        else:
            total_salary = df_emp['salary'].sum()

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
            df_emp = df_emp[df_emp['name'].str.contains(search_query, case=False) | df_emp['employee_id'].str.contains(search_query, case=False)]

        # 排序邏輯
        if sort_opt == "員工編號順序":
            df_emp = df_emp.sort_values("employee_id")
        elif sort_opt == "薪資 (由高到低)":
            df_emp = df_emp.sort_values("salary", ascending=False)
        elif sort_opt == "薪資 (由低到高)":
            df_emp = df_emp.sort_values("salary", ascending=True)
        elif sort_opt == "按部門排序":
            df_emp = df_emp.sort_values(["dept", "employee_id"])

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
            row_cols[0].write(row['employee_id'])
            row_cols[1].write(row['name'])
            row_cols[2].write(row['dept'])
            row_cols[3].write(row['position'])
            row_cols[4].write(f"NT$ {int(row['salary']):,}")
            
            if row_cols[5].button("🗑️", key=f"del_{row['employee_id']}", help="刪除此員工"):
                delete_employee(row['employee_id'])
                st.toast(f"已刪除員工: {row['name']}")
                time.sleep(0.5)
                st.rerun()
