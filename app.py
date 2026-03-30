import streamlit as st
import datetime
import pandas as pd
import time
import random
import sqlite3
import os
import re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# 設定頁面資訊
st.set_page_config(page_title="路徒行旅 Plus 站前館營運日誌", layout="wide")

# --- 安全防護：全站密碼攔截 ---
if "authenticated" not in st.session_state:
    st.markdown("<h2 style='text-align: center;'>🔒 歡迎登入 路徒行旅 Plus 營運日誌</h2>", unsafe_allow_html=True)
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

# -- 資料庫初始化 --
def init_db():
    conn = sqlite3.connect('roaders_plus.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS daily_data (
            date TEXT PRIMARY KEY,
            occ_rate REAL,
            adr INTEGER,
            revenue INTEGER,
            total_rooms INTEGER,
            
            counter_complaints TEXT,
            counter_expense INTEGER,
            
            cleaned_rooms INTEGER,
            hk_checkout_extend INTEGER,
            hk_avg_clean REAL,
            hk_expense INTEGER,
            
            rest_breakfast INTEGER,
            rest_month_guests INTEGER,
            rest_day_guests INTEGER,
            rest_avg_guests REAL,
            rest_month_rev INTEGER,
            rest_avg_spent INTEGER,
            rest_peak_expense INTEGER,
            rest_car_data TEXT,
            
            maint_repair_rooms INTEGER,
            maint_records TEXT,
            maint_expense INTEGER
        )
    ''')
    
    # 動態新增餐廳分館欄位（為了相容舊有資料庫結構）
    new_columns = [
        "bf_theme_est INTEGER", "bf_theme_act INTEGER",
        "bf_zq_est INTEGER", "bf_zq_act INTEGER",
        "bf_total_est INTEGER", "bf_total_act INTEGER",
        "af_theme_est INTEGER", "af_theme_act INTEGER",
        "af_zq_est INTEGER", "af_zq_act INTEGER",
        "af_total_est INTEGER", "af_total_act INTEGER",
        "daily_work_log TEXT"
    ]
    for col_def in new_columns:
        col_name = col_def.split()[0]
        try:
            c.execute(f"ALTER TABLE daily_data ADD COLUMN {col_def}")
        except sqlite3.OperationalError:
            pass # column already exists
            
    # --- 新增人事管理資料表 ---
    c.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            employee_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            dept TEXT NOT NULL,
            position TEXT,
            salary INTEGER
        )
    ''')
    
    conn.commit()
    conn.close()

init_db()

# -- 側邊欄：進階日期選擇器 --
st.sidebar.header("📅 日期導覽")
if 'sidebar_date' not in st.session_state:
    st.session_state['sidebar_date'] = datetime.date.today()

def prev_day(): st.session_state['sidebar_date'] -= datetime.timedelta(days=1)
def next_day(): st.session_state['sidebar_date'] += datetime.timedelta(days=1)

col1, col2 = st.sidebar.columns(2)
col1.button("⬅️ 前一天", on_click=prev_day)
col2.button("後一天 ➡️", on_click=next_day)

selected_date = st.sidebar.date_input("選擇日期", value=st.session_state['sidebar_date'], key='sidebar_date')
date_str = str(selected_date)

# --- 新增：週次預覽選擇器 (提前定義以供載入邏輯使用) ---
import calendar
_, last_day_of_month = calendar.monthrange(selected_date.year, selected_date.month)
weekly_options = [
    "--- 關閉週預覽 ---",
    "第1週 (1號 - 7號)",
    "第2週 (8號 - 14號)",
    "第3週 (15-21號)",
    "第4週 (22-28號)",
    f"第5週 (29號 - {last_day_of_month}號)"
]
selected_week = st.sidebar.selectbox("快速查閱區間：", weekly_options, index=0, key="weekly_view_select")
# --------------------------------------------------

# -- 資料庫讀寫函數 --
def get_daily_data(d_str):
    conn = sqlite3.connect('roaders_plus.db')
    df = pd.read_sql_query("SELECT * FROM daily_data WHERE date=?", conn, params=(d_str,))
    conn.close()
    if not df.empty:
        return df.iloc[0].to_dict()
    return {}

def save_daily_data(d_str, data_dict):
    conn = sqlite3.connect('roaders_plus.db')
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO daily_data (date) VALUES (?)", (d_str,))
    
    set_clause = ", ".join([f"{k} = ?" for k in data_dict.keys()])
    values = list(data_dict.values()) + [d_str]
    
    if set_clause:
        c.execute(f"UPDATE daily_data SET {set_clause} WHERE date = ?", values)
    conn.commit()
    conn.close()

# 自動載入當日資料到 session_state
day_data = get_daily_data(date_str)

field_mapping = {
    'input_occ': ('occ_rate', 0.0),
    'input_adr': ('adr', 0),
    'input_rev': ('revenue', 0),
    'input_rooms': ('total_rooms', 0),
    
    'input_complaints': ('counter_complaints', ""),
    'input_counter_exp': ('counter_expense', 0),
    
    'input_cleaned': ('cleaned_rooms', 0),
    'input_hk_co': ('hk_checkout_extend', 0),
    'input_hk_avg': ('hk_avg_clean', 0.0),
    'input_hk_exp': ('hk_expense', 0),
    
    'input_bf_theme_est': ('bf_theme_est', 0),
    'input_bf_theme_act': ('bf_theme_act', 0),
    'input_bf_zq_est': ('bf_zq_est', 0),
    'input_bf_zq_act': ('bf_zq_act', 0),
    'input_bf_total_est': ('bf_total_est', 0),
    'input_bf_total_act': ('bf_total_act', 0),
    
    'input_af_theme_est': ('af_theme_est', 0),
    'input_af_theme_act': ('af_theme_act', 0),
    'input_af_zq_est': ('af_zq_est', 0),
    'input_af_zq_act': ('af_zq_act', 0),
    'input_af_total_est': ('af_total_est', 0),
    'input_af_total_act': ('af_total_act', 0),
    
    'input_rest_mrev': ('rest_month_rev', 0),
    'input_rest_aspent': ('rest_avg_spent', 0),
    'input_rest_exp': ('rest_peak_expense', 0),
    'input_rest_car': ('rest_car_data', ""),
    
    'input_repair': ('maint_repair_rooms', 0),
    'input_maint_rec': ('maint_records', ""),
    'input_maint_exp': ('maint_expense', 0),
    
    'input_daily_log': ('daily_work_log', "")
}

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
    st.session_state['_last_loaded_date'] = date_str
    st.session_state['_last_week_view'] = selected_week

# 新增：自動儲存函數，避免切換日期時資料遺失
def sync_st_to_db():
    update_dict = {db_col: st.session_state[ss_key] for ss_key, (db_col, _) in field_mapping.items() if ss_key in st.session_state}
    if update_dict:
        save_daily_data(date_str, update_dict)

def on_input_change():
    sync_st_to_db()

st.sidebar.divider()
st.sidebar.subheader("📤 數據匯出與備份")

def generate_report_text(d_str):
    data = get_daily_data(d_str)
    if not data: return f"--- {d_str} 無紀錄 ---"
    
    report = []
    report.append(f"========================================")
    report.append(f"🏨 路徒行旅 Plus 站前館 - 營運日誌 ({d_str})")
    report.append(f"========================================\n")
    
    report.append(f"【📊 營運指標】")
    report.append(f"- 住房率: {data.get('occ_rate', 0)}%")
    report.append(f"- ADR: NT$ {int(data.get('adr', 0)):,}")
    report.append(f"- 總營收: NT$ {int(data.get('revenue', 0)):,}")
    report.append(f"- 總住房數: {data.get('total_rooms', 0)} 間\n")
    
    report.append(f"【💼 櫃台與房務】")
    report.append(f"- 負評客訴: {data.get('counter_complaints', '無')}")
    report.append(f"- 櫃台請購: {data.get('counter_expense', 0)} 元")
    report.append(f"- 總清消房數: {data.get('cleaned_rooms', 0)} 間")
    report.append(f"- 房務請購: {data.get('hk_expense', 0)} 元\n")
    
    report.append(f"【🍽️ 餐廳數據 (兩館實際來客)】")
    report.append(f"- 早餐總計: {int(data.get('bf_total_act', 0))} 人")
    report.append(f"- 下午茶總計: {int(data.get('af_total_act', 0))} 人")
    report.append(f"- 餐廳營收(全月): {data.get('rest_month_rev', 0)} 元\n")
    
    report.append(f"【🔧 工務紀錄】")
    report.append(f"- 待修房數: {data.get('maint_repair_rooms', 0)} 間")
    report.append(f"- 修繕細節: {data.get('maint_records', '無')}\n")
    
    report.append(f"【📝 每日營運紀錄細節】")
    report.append(f"{data.get('daily_work_log', '無紀錄内容')}")
    report.append(f"\n" + "-"*40 + "\n")
    
    return "\n".join(report)

# 1. 單日匯出
single_report = generate_report_text(date_str)
st.sidebar.download_button(
    label="📄 匯出當日紀錄 (.txt)",
    data=single_report,
    file_name=f"Roaders_Plus_Daily_{date_str}.txt",
    mime="text/plain",
    use_container_width=True
)

# 2. 全月匯出
month_str = selected_date.strftime('%Y-%m')
if f"monthly_report_{month_str}" not in st.session_state:
    if st.sidebar.button(f"📅 準備 {month_str} 紀錄匯出", use_container_width=True):
        conn = sqlite3.connect('roaders_plus.db')
        df_all = pd.read_sql_query("SELECT date FROM daily_data WHERE date LIKE ? ORDER BY date ASC", conn, params=(f"{month_str}%",))
        conn.close()
        
        if df_all.empty:
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

st.sidebar.divider()
st.sidebar.subheader("📅 週次紀錄快速審視 (已於上方選擇)")
st.sidebar.info(f"當前模式：{selected_week}")

st.sidebar.divider()
if st.sidebar.button("💾 強制儲存今日所有變更", use_container_width=True):
    sync_st_to_db()
    st.sidebar.success("✅ 今日資料已安全寫入資料庫！")

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
        df = pd.read_csv(file, skiprows=header_idx) if is_csv else pd.read_excel(file, skiprows=header_idx)
        df.columns = df.columns.astype(str).str.replace(r'[\s\n\r]', '', regex=True)
        
        date_col = next((c for c in df.columns if '日期' in c), None)
        occ_col = next((c for c in df.columns if '住房率' in c or '訂房率' in c or '出租率' in c or 'OCC' in c.upper()), None)
        adr_col = next((c for c in df.columns if '平均房價' in c or 'ADR' in c.upper()), None)
        
        # 營收：優先抓客房收入或總營收，避開「可售總金額」
        rev_col = next((c for c in df.columns if '客房收入' in c or '客房營收' in c or '總營收' in c or '營業額' in c or '實際營收' in c), None)
        # 房數：優先找出租/住房，避開「可售」
        rooms_col = next((c for c in df.columns if ('住房數' in c or '出租' in c or '售出' in c or '實住' in c) and '可售' not in c), None)
        if not rooms_col:
            rooms_col = next((c for c in df.columns if ('房間數' in c or '客房數' in c) and '可售' not in c), None)

        if not date_col:
            st.error("⚠️ 解析失敗：找不到『日期』欄位")
            return False

        df[date_col] = df[date_col].astype(str).str.replace(r'\.0$', '', regex=True)
        df['標準日期'] = pd.to_datetime(df[date_col], format='%Y%m%d', errors='coerce').dt.date
        null_mask = df['標準日期'].isnull()
        if null_mask.any():
            df.loc[null_mask, '標準日期'] = pd.to_datetime(df.loc[null_mask, date_col], errors='coerce').dt.date
            
        conn = sqlite3.connect('roaders_plus.db')
        c = conn.cursor()
        
        records_saved = 0
        for _, row in df.iterrows():
            d_str = str(row['標準日期'])
            if d_str == 'NaT' or d_str == 'None': continue
            
            # 處理訂房率 (可能為 '92%', '92.5' 字串，或 0.925 浮點數)
            occ = 0.0
            if occ_col and pd.notna(row.get(occ_col)):
                raw_occ_str = str(row.get(occ_col)).strip()
                has_percent = '%' in raw_occ_str
                clean_occ = raw_occ_str.replace('%', '').replace(',', '')
                try:
                    occ = float(clean_occ)
                    if 0 < occ <= 1.0 and not has_percent:
                        occ = occ * 100.0  # 把 0.92 轉成 92.0
                except ValueError:
                    occ = 0.0

            adr = int(float(str(row.get(adr_col, '0')).replace(',', ''))) if adr_col and pd.notna(row.get(adr_col)) else 0
            rev = int(float(str(row.get(rev_col, '0')).replace(',', ''))) if rev_col and pd.notna(row.get(rev_col)) else 0
            rooms = int(float(str(row.get(rooms_col, '0')).replace(',', ''))) if rooms_col and pd.notna(row.get(rooms_col)) else 0

            c.execute("INSERT OR IGNORE INTO daily_data (date) VALUES (?)", (d_str,))
            c.execute("""
                UPDATE daily_data 
                SET occ_rate = ?, adr = ?, revenue = ?, total_rooms = ?
                WHERE date = ?
            """, (occ, adr, rev, rooms, d_str))
            records_saved += 1
            
        conn.commit()
        conn.close()
        
        # 強制立刻更新畫面的狀態
        st.session_state['_last_loaded_date'] = None
        return records_saved
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
            # 判斷是否為「3/1週日」格式
            m = re.match(r'^(\d{1,2})/(\d{1,2})', col0)
            if m:
                month, day = m.groups()
                d_str = f"{current_year}-{int(month):02d}-{int(day):02d}"
                
                def safe_int(val):
                    try: return int(float(val))
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
                    'bf_theme_est': bf_theme_est, 'bf_theme_act': bf_theme_act,
                    'bf_zq_est': bf_zq_est, 'bf_zq_act': bf_zq_act,
                    'bf_total_est': bf_total_est, 'bf_total_act': bf_total_act,
                    'af_theme_est': af_theme_est, 'af_theme_act': af_theme_act,
                    'af_zq_est': af_zq_est, 'af_zq_act': af_zq_act,
                    'af_total_est': af_total_est, 'af_total_act': af_total_act
                })

        conn = sqlite3.connect('roaders_plus.db')
        c = conn.cursor()
        
        for r in parsed_days:
            c.execute("INSERT OR IGNORE INTO daily_data (date) VALUES (?)", (r['date'],))
            c.execute("""
                UPDATE daily_data 
                SET rest_month_rev = ?, rest_avg_spent = ?,
                    bf_theme_est=?, bf_theme_act=?, bf_zq_est=?, bf_zq_act=?, bf_total_est=?, bf_total_act=?,
                    af_theme_est=?, af_theme_act=?, af_zq_est=?, af_zq_act=?, af_total_est=?, af_total_act=?
                WHERE date = ?
            """, (month_rev, avg_spent,
                  r['bf_theme_est'], r['bf_theme_act'], r['bf_zq_est'], r['bf_zq_act'], r['bf_total_est'], r['bf_total_act'],
                  r['af_theme_est'], r['af_theme_act'], r['af_zq_est'], r['af_zq_act'], r['af_total_est'], r['af_total_act'],
                  r['date']))
            
        conn.commit()
        conn.close()
        
        st.session_state['_last_loaded_date'] = None
        return len(parsed_days)
    except Exception as e:
        import traceback
        st.error(f"解析餐廳報表失敗: {e}\n{traceback.format_exc()}")
        return False

# 頁面標題
st.title("路徒行旅 Plus 站前館營運日誌")

# 主畫面
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📊 營運總覽", "💼 櫃台數據", "🧹 房務數據", "🍽️ 餐廳數據", "🔧 工務數據", "📝 每日營運紀錄", "👥 人事概況"])

with tab2:
    st.header("💼 櫃台數據")
    st.subheader("📁 金旭報表上傳區 (支援全月匯入)")
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
    conn = sqlite3.connect('roaders_plus.db')
    df_mtd = pd.read_sql_query("SELECT * FROM daily_data WHERE date >= ? AND date <= ?", conn, params=(start_of_month, date_str))
    conn.close()

    if not df_mtd.empty:
        mtd_rooms = 0.0
        mtd_rev = 0.0
        total_sellable = 0.0
        
        for _, r in df_mtd.iterrows():
            o = float(r['occ_rate']) if pd.notna(r['occ_rate']) else 0.0
            adr = float(r['adr']) if pd.notna(r['adr']) else 0.0
            rev = float(r['revenue']) if pd.notna(r['revenue']) else 0.0
            rm = float(r['total_rooms']) if pd.notna(r['total_rooms']) else 0.0
            
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

    st.divider()

    # -- 今日看板 --
    st.subheader(f"今日全館營運大看板 ({date_str})")
    adr_val = st.session_state.get('input_adr', 0)
    rev_val = st.session_state.get('input_rev', 0)
    
    kpi_html = f"""
    <style>
    .kpi-container {{ display: flex; justify-content: space-around; flex-wrap: wrap; margin-bottom: 30px; }}
    .kpi-circle {{ width: 170px; height: 170px; border-radius: 50%; background: linear-gradient(135deg, #1f2c56 0%, #2e437c 100%); color: white; display: flex; flex-direction: column; justify-content: center; align-items: center; box-shadow: 0 8px 15px rgba(0,0,0,0.15); border: 4px solid #4CAF50; margin: 15px; }}
    .kpi-title {{ font-size: 16px; margin-bottom: 8px; color: #d8e2fb; }}
    .kpi-value {{ font-size: 26px; font-weight: bold; }}
    </style>
    <div class="kpi-container">
        <div class="kpi-circle"><div class="kpi-title">今日住房率</div><div class="kpi-value">{occ_val}%</div></div>
        <div class="kpi-circle"><div class="kpi-title">ADR</div><div class="kpi-value">NT$ {adr_val:,}</div></div>
        <div class="kpi-circle"><div class="kpi-title">總營收</div><div class="kpi-value">NT$ {rev_val:,}</div></div>
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
        st.metric("今日雙館早餐總來客", f"{int(bf_total_act)} 人")

with tab3:
    st.header("🧹 房務數據")
    st.number_input("今日總清消房數", min_value=0, step=1, key="input_cleaned", on_change=on_input_change)
    st.number_input("退/續數量", min_value=0, step=1, key="input_hk_co", on_change=on_input_change)
    st.number_input("每人平均掃房數", min_value=0.0, step=0.1, key="input_hk_avg", on_change=on_input_change)
    st.number_input("房務請購費用", min_value=0, step=100, key="input_hk_exp", on_change=on_input_change)

with tab4:
    st.header("🍽️ 餐廳數據")
    st.subheader("📁 餐廳報表上傳區")
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
        # 解析選擇的區間
        week_idx = weekly_options.index(selected_week)
        start_d = (week_idx - 1) * 7 + 1
        if week_idx == 5:
            end_d = last_day_of_month
        else:
            end_d = start_d + 6
            
        st.subheader(f"📋 {selected_week} 快速審視模式")
        st.info(f"正在查看 {selected_date.year}年{selected_date.month}月 {start_d}號 至 {end_d}號 的完整紀錄。")
        
        # 獲取該區間所有資料
        conn = sqlite3.connect('roaders_plus.db')
        c_month_str = selected_date.strftime('%Y-%m')
        
        for day in range(start_d, end_d + 1):
            target_date = f"{c_month_str}-{day:02d}"
            # 這裡我們呼叫 get_daily_data
            d_data = get_daily_data(target_date)
            
            with st.expander(f"📅 {target_date} 營運紀錄", expanded=True):
                if d_data and d_data.get('daily_work_log'):
                    st.markdown(f"**【當日日誌細節】**\n\n{d_data['daily_work_log']}")
                    st.divider()
                    col_a, colb, colc = st.columns(3)
                    col_a.metric("住房率", f"{d_data.get('occ_rate', 0)}%")
                    colb.metric("ADR", f"NT$ {int(d_data.get('adr', 0)):,}")
                    colc.metric("營收", f"NT$ {int(d_data.get('revenue', 0)):,}")
                else:
                    st.write("🌑 此日期尚無任何日誌紀錄。")
        conn.close()
        
        if st.button("⬅️ 返回今日編輯模式"):
            st.rerun()

    else:
        st.info(f"💡 請在下方詳細填寫 **{date_str}** 的各項營運日誌與重點工作回報。這裡的紀錄會自動儲存，切換日期或關閉網頁也不用擔心遺失。")
        st.text_area("✍️ 今日工作與營運細節報告：", height=500, key="input_daily_log", placeholder="可以在這裡記錄交班重點、客訴特殊處理、VIP 接待細節、設備大修紀錄...等", on_change=on_input_change)

with tab7:
    st.header("👥 人事概況")
    
    # -- 人事管理函數 --
    def get_all_employees():
        conn = sqlite3.connect('roaders_plus.db')
        df = pd.read_sql_query("SELECT * FROM employees", conn)
        conn.close()
        return df

    def add_employee(e_id, name, dept, pos, salary):
        try:
            conn = sqlite3.connect('roaders_plus.db')
            c = conn.cursor()
            c.execute("INSERT INTO employees (employee_id, name, dept, position, salary) VALUES (?, ?, ?, ?, ?)",
                      (e_id, name, dept, pos, salary))
            conn.commit()
            conn.close()
            return True
        except sqlite3.IntegrityError:
            return "ID_EXISTS"
        except Exception as e:
            return str(e)

    def delete_employee(e_id):
        conn = sqlite3.connect('roaders_plus.db')
        c = conn.cursor()
        c.execute("DELETE FROM employees WHERE employee_id = ?", (e_id,))
        conn.commit()
        conn.close()

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
