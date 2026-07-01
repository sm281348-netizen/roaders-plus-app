import traceback
import streamlit as st

# --- Gspread Retry Patch ---
import gspread.client
import time
from gspread.exceptions import APIError

if not hasattr(gspread.client.Client, '_original_request_patched'):
    original_request = gspread.client.Client.request

    def patched_request(self, method, endpoint, params=None, data=None, json=None, files=None, headers=None):
        for attempt in range(5):
            try:
                return original_request(self, method, endpoint, params=params, data=data, json=json, files=files, headers=headers)
            except APIError as e:
                err_msg = str(e).lower()
                is_retryable = False
                
                if hasattr(e, 'response') and getattr(e.response, 'status_code', 0) == 429:
                    is_retryable = True
                elif "429" in err_msg or "resource_exhausted" in err_msg or "quota" in err_msg or "redacted" in err_msg:
                    is_retryable = True
                    
                if is_retryable:
                    if attempt < 4:
                        time.sleep(15)  # Sleep 15s to wait out the 60s quota window
                        continue
                raise e

    gspread.client.Client.request = patched_request
    gspread.client.Client._original_request_patched = True
# ---------------------------

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


@st.cache_data(ttl=3600)
def fetch_taipei_events():
    """讀取試算表中的台北重大活動分頁"""
    try:
        # 讀取試算表中的 taipei_events 分頁
        df = read_google_sheet("taipei_events", ttl="1h")
        if df is not None and not df.empty:
            df = standardize_df_dates(df)
            return df
    except Exception:
        pass
    return pd.DataFrame(columns=['date', 'event_name', 'event_type', 'venue'])


@st.cache_data(ttl=3600)
def fetch_supplier_prices():
    """讀取菜價表 supplier_prices 分頁，回傳標準化 DataFrame"""
    try:
        from streamlit_gsheets import GSheetsConnection
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        df = raw_st.read(worksheet="supplier_prices", spreadsheet=url_st, ttl="1h")
        if df is None or df.empty:
            return pd.DataFrame()
        # 欄位名稱標準化 (item name → item_name)
        df.columns = [str(c).strip().lower().replace(' ', '_') for c in df.columns]
        
        # 支援中文欄位自動對應
        col_mapping = {
            '日期': 'period',
            '品項': 'item_name',
            '品名': 'item_name',
            '品項名稱': 'item_name',
            '單位': 'unit',
            '單價': 'price',
            '價格': 'price',
            '金額': 'price'
        }
        df.rename(columns=col_mapping, inplace=True)
        # 必要欄位檢查
        required = {'period', 'item_name', 'unit', 'price'}
        if not required.issubset(set(df.columns)):
            return pd.DataFrame()
        # 清理資料
        df = df.dropna(subset=['item_name', 'price'])
        df['price'] = pd.to_numeric(df['price'], errors='coerce')
        df = df.dropna(subset=['price'])
        df['item_name'] = df['item_name'].astype(str).str.strip()
        df['unit'] = df['unit'].astype(str).str.strip()
        # period 日期解析 (支援 YYYY/M/D, YYYY-MM-DD, Timestamp, float 等各種格式)

        def parse_period(v):
            # 若已是 date/datetime，直接轉
            if isinstance(v, datetime.date):
                return v if not isinstance(v, datetime.datetime) else v.date()
            # 若是 pandas Timestamp
            if isinstance(v, pd.Timestamp):
                return v.date()
            # 嘗試 pd.to_datetime (最萬用，處理 float 序號、各種字串格式)
            try:
                return pd.to_datetime(v, dayfirst=False).date()
            except:
                pass
            return None
        df['period_dt'] = df['period'].apply(parse_period)
        df = df.dropna(subset=['period_dt'])
        df = df.sort_values('period_dt').reset_index(drop=True)
        return df
    except Exception as e:
        return pd.DataFrame()


@st.cache_data(ttl=3600)
def fetch_thepeak_daily_purchase_report():
    """讀取 thepeak_daily_purchase_report 分頁，使用 60s 快取以避免 API 限制"""
    try:
        from streamlit_gsheets import GSheetsConnection
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        # 直接使用確定的分頁名稱
        target_ws = "thepeak_daily_purchase_report"
        df = raw_st.read(worksheet=target_ws, spreadsheet=url_st, ttl=0)
        if df is None:
            return pd.DataFrame()
        return df
    except Exception:
        return pd.DataFrame()

def append_thepeak_daily_purchase_report(new_rows_df):
    """附加新資料到 thepeak_daily_purchase_report 分頁"""
    try:
        from streamlit_gsheets import GSheetsConnection
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        
        target_ws = "thepeak_daily_purchase_report"
        try:
            df_old = raw_st.read(worksheet=target_ws, spreadsheet=url_st, ttl=0)
        except Exception as e:
            if target_ws in str(e) or "WorksheetNotFound" in str(type(e)):
                st.error(f"無法在 RTS_backup 中找到 '{target_ws}' 分頁。\n目前該試算表有的分頁為: {', '.join(ws_titles)}")
                return False
            raise e
        
        if df_old is None or df_old.empty:
            df_new = new_rows_df
        else:
            # 確保欄位型態一致避免警告，並保留原本所有資料
            df_new = pd.concat([df_old, new_rows_df], ignore_index=True)
            
        raw_st.update(worksheet=target_ws, data=df_new, spreadsheet=url_st)
        return True
    except Exception as e:
        st.error(f"寫入失敗: {e}")
        return False


@st.cache_data(ttl=3600)
def fetch_4fhh_daily_purchase_report():
    """讀取 4FHH_daily_purchase_report 分頁，使用 60s 快取以避免 API 限制"""
    try:
        from streamlit_gsheets import GSheetsConnection
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        # 直接使用確定的分頁名稱
        target_ws = "4FHH_daily_purchase_report"
        df = raw_st.read(worksheet=target_ws, spreadsheet=url_st, ttl=0)
        if df is None:
            return pd.DataFrame()
        return df
    except Exception:
        return pd.DataFrame()

def append_4fhh_daily_purchase_report(new_rows_df):
    """附加新資料到 4FHH_daily_purchase_report 分頁"""
    try:
        from streamlit_gsheets import GSheetsConnection
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        
        target_ws = "4FHH_daily_purchase_report"
        try:
            df_old = raw_st.read(worksheet=target_ws, spreadsheet=url_st, ttl=0)
        except Exception as e:
            if target_ws in str(e) or "WorksheetNotFound" in str(type(e)):
                st.error(f"無法在 RTS_backup 中找到 '{target_ws}' 分頁。\n目前該試算表有的分頁為: {', '.join(ws_titles)}")
                return False
            raise e
        
        if df_old is None or df_old.empty:
            df_new = new_rows_df
        else:
            # 確保欄位型態一致避免警告，並保留原本所有資料
            df_new = pd.concat([df_old, new_rows_df], ignore_index=True)
            
        raw_st.update(worksheet=target_ws, data=df_new, spreadsheet=url_st)
        return True
    except Exception as e:
        st.error(f"寫入失敗: {e}")
        return False
def get_market_index_df(sp_df):
    """將 supplier_prices DataFrame 轉換為大盤物價指數 (Market Price Index) DataFrame
    使用連鎖指數法 (Chain-linking) 以支援中途新增或移除的食材。"""
    if sp_df is None or sp_df.empty:
        return pd.DataFrame()

    periods_available = sorted(sp_df['period_dt'].unique())
    if len(periods_available) < 1:
        return pd.DataFrame()

    index_data = []
    prev_dict = sp_df[sp_df['period_dt'] == periods_available[0]].set_index(['item_name', 'unit'])['price'].to_dict()
    current_index = 100.0
    
    index_data.append({
        'period_dt': periods_available[0],
        'period_str': str(periods_available[0]),
        'month_label': periods_available[0].strftime('%Y-%m'),
        'index': round(current_index, 1)
    })

    for i in range(1, len(periods_available)):
        p = periods_available[i]
        curr_df = sp_df[sp_df['period_dt'] == p]
        curr_dict = curr_df.set_index(['item_name', 'unit'])['price'].to_dict()
        
        ratios = []
        for key, price in curr_dict.items():
            if key in prev_dict and prev_dict[key] > 0 and pd.notna(price):
                ratios.append(price / prev_dict[key])
                
        if ratios:
            period_ratio = sum(ratios) / len(ratios)
            current_index = current_index * period_ratio
            
        index_data.append({
            'period_dt': p,
            'period_str': str(p),
            'month_label': p.strftime('%Y-%m'),
            'index': round(current_index, 1)
        })
        
        prev_dict = curr_dict

    return pd.DataFrame(index_data)


def calc_key_metrics(m_data):
    df = m_data.get('df', pd.DataFrame())
    res = {'high_adr_days': 0, 'top20_rev_avg': 0, 'bot20_rev_avg': 0,
           'dual_match_days': 0, 'month_label': m_data.get('month_label', '')}
    if df is None or df.empty:
        return res

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

    res['top20_rev_avg'] = top20_df['rev_val'].mean(
    ) if not top20_df.empty else 0
    res['bot20_rev_avg'] = bot20_df['rev_val'].mean(
    ) if not bot20_df.empty else 0

    # 雙冠天數：前 20% 營收日中，ADR 也大於當月平均 ADR 的天數
    dual_match_df = top20_df[top20_df['is_high_adr']]
    res['dual_match_days'] = int(len(dual_match_df))
    res['dual_match_dates'] = dual_match_df['date'].sort_values(
    ).tolist() if not dual_match_df.empty else []

    return res


@st.cache_data(ttl=86400 * 30)
def translate_to_zh(text):
    if not text:
        return text
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
                    day_details.append(
                        f"{TARGET_HOLIDAY_COUNTRIES[code]} {code}: {h_name}")
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
    years = {start_date.year,
             (start_date + datetime.timedelta(days=days)).year}
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
                    day_details.append(
                        f"{TARGET_HOLIDAY_COUNTRIES[code]} {code}: {h_name}")
                else:
                    has_other = True
                    day_details.append(f"🌍 {code}: {h_name}")

        flags_str = "".join(sorted(list(day_flags)))
        if has_other:
            flags_str += "🌍"

        if flags_str:
            result.append({
                'date': dt_obj.strftime('%Y-%m-%d'),
                'flags': flags_str,
                'details': ", ".join(day_details)
            })
    return result


# 設定頁面資訊
st.set_page_config(page_title="路徒Plus行旅站前館營運日誌", layout="wide")

password_station = st.secrets.get("admin_password", "roaders123")
password_theme = st.secrets.get("theme_password", "theme456")
password_purchase = st.secrets.get("purchase_password", "thepeak37")

if "authenticated" not in st.session_state:
    st.markdown("<h2 style='text-align: center;'>🔒 Welcome to Hotel Master</h2>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd = st.text_input("管理員通行碼", type="password")
        if pwd:
            if pwd == password_station:
                st.session_state["authenticated"] = True
                st.session_state["hotel_type"] = "站前館"
                st.rerun()
            elif pwd == password_theme:
                st.session_state["authenticated"] = True
                st.session_state["hotel_type"] = "主題館"
                st.rerun()
            elif pwd == password_purchase:
                st.session_state["authenticated"] = True
                st.session_state["hotel_type"] = "採購"
                st.rerun()
            else:
                st.error("❌ 密碼錯誤，請重新輸入")
                st.stop()
        else:
            st.stop()
# -----------------------------

# 取得目前館別（必須在連線前定義）
current_hotel = st.session_state.get("hotel_type", "站前館")


def get_google_sheet_error_hint(e):
    """根據 Google Sheets 連線錯誤類型回傳對應的中文建議"""
    msg = str(e)
    if "invalid_grant" in msg or "Token" in msg or "oauth" in msg.lower():
        return "🔑 憑證已過期或無效，請至 Streamlit Cloud Secrets 重新設定 Google Service Account。"
    if "quota" in msg.lower() or "rate limit" in msg.lower() or "429" in msg:
        return "⏳ API 配額已超限，請稍後再試。"
    if "403" in msg or "forbidden" in msg.lower() or "permission" in msg.lower():
        return "🚫 沒有權限，請確認 Google Sheet 已與 Service Account Email 共用編輯權限。"
    if "404" in msg or "not found" in msg.lower():
        return "❓ 找不到試算表，請確認 Secrets 中的 spreadsheet URL 是否正確。"
    if "Worksheet" in msg and "not found" in msg:
        return "📋 找不到該分頁名稱，請確認分頁名稱拼寫是否正確。"
    return None


# 試算表 URL（非敏感資訊，hardcode 作為 fallback 確保可靠）
_STATION_SPREADSHEET = st.secrets.get(
    "station_spreadsheet_url",
    "https://docs.google.com/spreadsheets/d/190DAPuSoorfuQzLb1f8E-jAVCnmm6gXC7YrahxCL-VQ/edit"
)
_THEME_SPREADSHEET = st.secrets.get(
    "theme_spreadsheet_url",
    "https://docs.google.com/spreadsheets/d/1zigbiXDK362v8pvkpFxEkLmBR6R4pCNy_qg7CCmcF0I/edit"
)
_ACTIVE_SPREADSHEET = _THEME_SPREADSHEET if current_hotel == "主題館" else _STATION_SPREADSHEET


class _ConnWrapper:
    """包住 GSheetsConnection，讓所有 read()/update() 自動帶 spreadsheet URL"""
    def __init__(self, raw_conn, spreadsheet_url):
        self._raw = raw_conn
        self._url = spreadsheet_url

    def read(self, worksheet=None, spreadsheet=None, **kwargs):
        return self._raw.read(worksheet=worksheet, spreadsheet=spreadsheet or self._url, **kwargs)

    def update(self, worksheet=None, data=None, spreadsheet=None, **kwargs):
        return self._raw.update(worksheet=worksheet, data=data, spreadsheet=spreadsheet or self._url, **kwargs)


try:
    if current_hotel == "主題館":
        _raw_conn = st.connection("gsheets_theme", type=GSheetsConnection)
    else:
        _raw_conn = st.connection("gsheets_station", type=GSheetsConnection)
    conn = _ConnWrapper(_raw_conn, _ACTIVE_SPREADSHEET)
except Exception as e:
    hint = get_google_sheet_error_hint(e)
    err_msg = f"無法建立 Google Sheets 連線: {e}"
    if hint:
        err_msg += f"\n建議: {hint}"
    st.error(err_msg)
    st.stop()


@st.cache_data(ttl=3600)
def _get_occ_data_cached_v2(hotel_type=""):
    try:
        df = conn.read(worksheet="occ_data", ttl=0)
    except Exception as e:
        raise e

    if df is not None and not df.empty:
        df.columns = [str(c).strip().lower() for c in df.columns]
        # 移除可能重複的欄位名稱 (例如使用者不小心在 Google Sheets 建了兩個 date 欄位)，避免 pandas 引發 KeyError: 0
        df = df.loc[:, ~df.columns.duplicated()]
        if 'date' not in df.columns and 'net occupancy' not in df.columns:
            for idx, r in df.head(10).iterrows():
                row_vals = [str(v).strip().lower() for v in r.values]
                if 'date' in row_vals or 'net occupancy' in row_vals:
                    df.columns = row_vals
                    df = df.iloc[idx+1:].reset_index(drop=True)
                    break
        if 'date' in df.columns or 'net occupancy' in df.columns:
            rename_map = {
                'date': 'date',
                'net occupancy': 'occ_rate',
                'adr (rooms sold)': 'adr',
                'total room revenue': 'revenue',
                'rooms available to sell': 'total_rooms',
                'total rooms sold': 'sold_rooms'
            }
            df = df.rename(columns=rename_map)
            df = df.loc[:, ~df.columns.duplicated()]
            df = df.loc[:, df.columns.notna() & (df.columns != 'nan') & (df.columns != '')]
            if 'date' in df.columns:
                def format_golden_date(d):
                    ds = str(d).strip()
                    if ds.endswith('.0'):
                        ds = ds[:-2]
                    if len(ds) == 8 and ds.isdigit():
                        return f"{ds[:4]}-{ds[4:6]}-{ds[6:]}"
                    if ds.isdigit() and 40000 < int(ds) < 60000:
                        import datetime
                        dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(ds))
                        return dt.strftime('%Y-%m-%d')
                    return ds
                df['date'] = df['date'].apply(format_golden_date)
            if 'date' in df.columns:
                df = df[df['date'].astype(str).str.strip() != '']
                df = df[~df['date'].astype(str).str.contains("合計|總計|total", case=False, na=False)]
        else:
            if 'date' not in df.columns:
                df['date'] = pd.Series(dtype='str')
            if 'revenue' not in df.columns:
                df['revenue'] = 0
    return df

@st.cache_data(ttl=3600)
def _get_cached_sheet_v3(worksheet, hotel_type=""):
    """集中快取層：所有唯讀 Sheet 請求走這裡，60s TTL，避免 API 429
    hotel_type 參數用於區分不同館的快取，避免跨館資料污染。
    注意：F&B 資料解析不在此函數內，請使用 compute_fb_mtd() 獨立讀取。"""
    if worksheet == "occ_data":
        return _get_occ_data_cached_v2(hotel_type)
    
    try:
        df = conn.read(worksheet=worksheet, ttl=0)
    except Exception as e:
        raise e
    return df


@st.cache_data(ttl=3600)
def compute_fb_mtd(start_date_str, end_date_str, _dummy_hotel=""):
    """
    獨立讀取 f&b_report 並計算指定日期區間的 MTD F&B 統計。
    快取 300 秒，不依賴 daily_data 快取函數，避免 API 429。

    f&b_report 欄位結構 (0-indexed，跳過前兩列標題後從日期行開始):
      [0] 日期  [1] 主題早預估  [2] 主題早實際  [3] 站前早預估  [4] 站前早實際
      [5] 主題下午預估  [6] 主題下午實際  [7] 站前下午預估  [8] 站前下午實際
      [9] HH 4F 實際
    """
    import re
    from streamlit_gsheets import GSheetsConnection

    result = {
        'bf_theme_act': 0, 'bf_zq_act': 0,
        'af_theme_act': 0, 'af_zq_act': 0,
        'bf_theme_est': 0, 'bf_zq_est': 0,
        'af_theme_est': 0, 'af_zq_est': 0,
        'hh_act': 0,
        'total_est': 0,
        'total_act_bf': 0,
        'total_act_af': 0,
        'revenue': 0,
        'avg_spent': 0,
        'matched_days': 0,   # 有早餐客數的天數（為日平均分母）
        'active_af_days': 0, # 有下午茶客數的天數
        'report_rows': 0,
        'error': ''
    }

    def safe_int(v):
        try:
            return int(float(str(v).replace(',', '').strip()))
        except:
            return 0

    def parse_date(raw):
        import pandas as pd
        if pd.isna(raw) or str(raw).strip() == '':
            return None
        s = str(raw).replace('.0', '').strip()
        m1 = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
        m2 = re.match(r'^(\d{8})$', s)
        if m1:
            return f"{m1.group(1)}-{int(m1.group(2)):02d}-{int(m1.group(3)):02d}"
        if m2 and s.isdigit():
            return f"{s[:4]}-{s[4:6]}-{s[6:]}"
        try:
            if s.isdigit() and len(s) == 5:
                d = pd.to_datetime('1899-12-30') + pd.to_timedelta(int(s), unit='D')
                return d.strftime('%Y-%m-%d')
            d = pd.to_datetime(raw)
            return d.strftime('%Y-%m-%d')
        except:
            return None

    # --- 讀取 f&b_report (RTS / 站前館，此分頁只放在 RTS) ---
    df_report = None
    try:
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        c_st = _ConnWrapper(raw_st, url_st)
        df_report = c_st.read(worksheet="f&b_report", ttl=0)
    except Exception as e:
        result['error'] += f'[f&b_report 讀取失敗: {e}] '

    if df_report is not None and not df_report.empty:
        result['report_rows'] = len(df_report)
        for _, row in df_report.iterrows():
            d = parse_date(row.iloc[0])
            if d and start_date_str <= d <= end_date_str:
                bf_theme_act = safe_int(row.iloc[2]) if len(row) > 2 else 0
                bf_zq_act    = safe_int(row.iloc[4]) if len(row) > 4 else 0
                af_theme_act = safe_int(row.iloc[6]) if len(row) > 6 else 0
                af_zq_act    = safe_int(row.iloc[8]) if len(row) > 8 else 0

                result['bf_theme_est'] += safe_int(row.iloc[1]) if len(row) > 1 else 0
                result['bf_theme_act'] += bf_theme_act
                result['bf_zq_est']    += safe_int(row.iloc[3]) if len(row) > 3 else 0
                result['bf_zq_act']    += bf_zq_act
                result['af_theme_est'] += safe_int(row.iloc[5]) if len(row) > 5 else 0
                result['af_theme_act'] += af_theme_act
                result['af_zq_est']    += safe_int(row.iloc[7]) if len(row) > 7 else 0
                result['af_zq_act']    += af_zq_act
                result['hh_act']       += safe_int(row.iloc[9]) if len(row) > 9 else 0

                # 只計算有早餐/下午茶實際客數的天數，避免把「未來空白日」也計入分母
                if (bf_theme_act + bf_zq_act) > 0:
                    result['matched_days'] += 1
                if (af_theme_act + af_zq_act) > 0:
                    result['active_af_days'] += 1

    # 彙總計算
    result['total_est'] = (result['bf_theme_est'] + result['bf_zq_est'] +
                           result['af_theme_est'] + result['af_zq_est'])
    result['total_act_bf'] = result['bf_theme_act'] + result['bf_zq_act']
    result['total_act_af'] = result['af_theme_act'] + result['af_zq_act']
    total_act = result['total_act_bf'] + result['total_act_af']

    # 營收 = 總預估人數 × 300（若預估為0則改用實際）
    est_for_rev = result['total_est'] if result['total_est'] > 0 else total_act
    result['revenue'] = est_for_rev * 300
    result['avg_spent'] = int(result['revenue'] / total_act) if total_act > 0 else 0

    return result


@st.cache_data(ttl=3600)
def fetch_fb_daily_df(year, month, _dummy_hotel=""):
    """
    從 f&b_report 讀取指定年月的每日 F&B 來客數。
    回傳 DataFrame，欄位為：date (YYYY-MM-DD), bf_act, af_act, hh_act, peak_guests
    """
    import re
    from streamlit_gsheets import GSheetsConnection

    def safe_int(v):
        try:
            return int(float(str(v).replace(',', '').strip()))
        except:
            return 0

    def parse_date(raw):
        import pandas as pd
        if pd.isna(raw) or str(raw).strip() == '':
            return None
        s = str(raw).replace('.0', '').strip()
        # 1) YYYY-MM-DD or YYYY/MM/DD
        m1 = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
        # 2) YYYYMMDD
        m2 = re.match(r'^(\d{8})$', s)
        
        if m1:
            return f"{m1.group(1)}-{int(m1.group(2)):02d}-{int(m1.group(3)):02d}"
        if m2 and s.isdigit():
            return f"{s[:4]}-{s[4:6]}-{s[6:]}"
            
        try:
            # 3) Excel serial date (e.g. 45413 for 2024-05-01 approx)
            if s.isdigit() and len(s) == 5:
                d = pd.to_datetime('1899-12-30') + pd.to_timedelta(int(s), unit='D')
                return d.strftime('%Y-%m-%d')
            # 4) Let pandas try to parse it (M/D/YYYY, etc)
            d = pd.to_datetime(raw)
            return d.strftime('%Y-%m-%d')
        except:
            return None

    m_start = f"{year}-{month:02d}-01"
    import calendar
    _, last_day = calendar.monthrange(year, month)
    m_end = f"{year}-{month:02d}-{last_day:02d}"

    rows = []
    try:
        conn_name = "gsheets_theme" if _dummy_hotel == "主題館" else "gsheets_station"
        raw_st = st.connection(conn_name, type=GSheetsConnection)
        url_st = st.secrets["connections"][conn_name]["spreadsheet"]
        c_st = _ConnWrapper(raw_st, url_st)
        df_report = c_st.read(worksheet="f&b_report", ttl=0)
        if df_report is not None and not df_report.empty:
            for _, row in df_report.iterrows():
                d = parse_date(row.iloc[0])
                if d and m_start <= d <= m_end:
                    bf_theme = safe_int(row.iloc[2]) if len(row) > 2 else 0
                    bf_zq    = safe_int(row.iloc[4]) if len(row) > 4 else 0
                    af_theme = safe_int(row.iloc[6]) if len(row) > 6 else 0
                    af_zq    = safe_int(row.iloc[8]) if len(row) > 8 else 0
                    hh_act   = safe_int(row.iloc[9]) if len(row) > 9 else 0
                    bf_act   = bf_theme + bf_zq
                    af_act   = af_theme + af_zq
                    rows.append({
                        'date': d,
                        'bf_act': bf_act,
                        'af_act': af_act,
                        'hh_act': hh_act,
                        'peak_guests': bf_act + af_act
                    })
    except Exception:
        pass

    if rows:
        return pd.DataFrame(rows)
    return pd.DataFrame(columns=['date', 'bf_act', 'af_act', 'hh_act', 'peak_guests'])

@st.cache_data(ttl=3600)
def fetch_fb_future_data(hotel_type="站前館"):
    """
    從 f&b_data 讀取包含未來的 F&B 預約客數明細。
    回傳清理後的 DataFrame。
    """
    import pandas as pd
    import re
    from streamlit_gsheets import GSheetsConnection
    try:
        conn_name = "gsheets_theme" if hotel_type == "主題館" else "gsheets_station"
        raw_st = st.connection(conn_name, type=GSheetsConnection)
        url_st = st.secrets["connections"][conn_name]["spreadsheet"]
        c_st = _ConnWrapper(raw_st, url_st)
        df = c_st.read(worksheet="f&b_data", ttl=0)
        
        if df is not None and not df.empty:
            if '數量' in df.columns:
                df['數量'] = pd.to_numeric(df['數量'], errors='coerce').fillna(0).astype(int)
            if '服務日期' in df.columns:
                df['服務日期'] = df['服務日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
                def format_dt(s):
                    import pandas as pd
                    if pd.isna(s) or s == '' or s.lower() == 'nan':
                        return s
                    if len(s) == 8 and s.isdigit():
                        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
                    m = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
                    if m:
                        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                    try:
                        if s.isdigit() and len(s) == 5:
                            d = pd.to_datetime('1899-12-30') + pd.to_timedelta(int(s), unit='D')
                            return d.strftime('%Y-%m-%d')
                        d = pd.to_datetime(s)
                        return d.strftime('%Y-%m-%d')
                    except:
                        return s
                df['date'] = df['服務日期'].apply(format_dt)
            return df
    except Exception as e:
        import logging
        logging.error(f"Error reading f&b_data: {e}")
        pass
    return pd.DataFrame()

def get_combined_fb_daily_df(year, month, current_hotel):
    import pandas as pd
    df = fetch_fb_daily_df(year, month)
    if df.empty:
        return pd.DataFrame(columns=['date', 'bf_act', 'af_act', 'hh_act', 'peak_guests'])
    return df

def get_combined_fb_future_data():
    import pandas as pd
    df_st = fetch_fb_future_data(hotel_type="站前館")
    df_th = fetch_fb_future_data(hotel_type="主題館")
    return pd.concat([df_st, df_th], ignore_index=True)

@st.cache_data(ttl=3600)


@st.cache_data(ttl=3600)
def get_all_employees_cached():
    try:
        from streamlit_gsheets import GSheetsConnection
        import streamlit as st
        import pandas as pd
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        df = raw_st.read(worksheet="employees", spreadsheet=url_st, ttl=0)
        return df if df is not None else pd.DataFrame()
    except:
        import pandas as pd
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def get_purchase_data_cached():
    import pandas as pd
    import streamlit as st
    from streamlit_gsheets import GSheetsConnection
    possible_names = ["purchase data", "Purchase Data", "purchase_data", "Purchase_Data"]
    try:
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        for name in possible_names:
            try:
                df = raw_st.read(worksheet=name, spreadsheet=url_st, ttl=0)
                if df is not None and not df.empty:
                    return df
            except:
                continue
    except:
        pass
    return pd.DataFrame()


def get_prediction_snapshots():
    import pandas as pd
    import streamlit as st
    from streamlit_gsheets import GSheetsConnection
    try:
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        df = raw_st.read(worksheet="prediction_snapshots", spreadsheet=url_st, ttl="0")
        if df is None or df.empty or 'snapshot_date' not in df.columns:
            return pd.DataFrame(columns=['snapshot_date', 'target_date_start', 'target_date_end', 'bf_conv_rate', 'af_conv_rate', 'future_bf', 'future_af', 'future_total'])
        return df
    except Exception:
        return pd.DataFrame(columns=['snapshot_date', 'target_date_start', 'target_date_end', 'bf_conv_rate', 'af_conv_rate', 'future_bf', 'future_af', 'future_total'])

def save_prediction_snapshot(new_row_dict):
    import pandas as pd
    import streamlit as st
    from streamlit_gsheets import GSheetsConnection
    df = get_prediction_snapshots()
    new_df = pd.DataFrame([new_row_dict])
    df = pd.concat([df, new_df], ignore_index=True)
    try:
        raw_st = st.connection("gsheets_station", type=GSheetsConnection)
        url_st = st.secrets["connections"]["gsheets_station"]["spreadsheet"]
        raw_st.update(worksheet="prediction_snapshots", data=df.fillna(""), spreadsheet=url_st)
        return True
    except Exception as e:
        st.error(f"儲存快照失敗: {e}")
        return False

def read_google_sheet(worksheet, ttl="1h"):
    try:
        return _get_cached_sheet_v3(worksheet, hotel_type=current_hotel)
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"Google Sheet 讀取失敗：{worksheet} ({e})"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
        return None


# -- 基本資料庫讀寫函數 (需優先定義以供導航邏輯使用) --


def standardize_df_dates(df):
    if df is None or df.empty or 'date' not in df.columns:
        return df

    def fix_d(val):
        if pd.isna(val) or str(val).strip() == '' or str(val).strip() == 'NaT':
            return ""
        v_str = str(val).split(' ')[0].strip()

        # 去掉 .0 尾綴 (Google Sheets 浮點數)
        if v_str.endswith('.0'):
            v_str = v_str[:-2]

        # 處理 Excel 數字日期 (例如 45809)
        if v_str.isdigit() and 40000 < int(v_str) < 60000:
            import datetime
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v_str))
            return dt.strftime('%Y-%m-%d')

        # 處理 YYYYMM 6位數 (例如 202605 → 2026-05-01)
        if v_str.isdigit() and len(v_str) == 6:
            y, m = int(v_str[:4]), int(v_str[4:])
            if 1900 < y < 2100 and 1 <= m <= 12:
                return f"{y:04d}-{m:02d}-01"

        # 處理 YYYYMMDD 8位數 (例如 20260611 → 2026-06-11)
        if v_str.isdigit() and len(v_str) == 8:
            return f"{v_str[:4]}-{v_str[4:6]}-{v_str[6:]}"

        import re
        # 處理民國年或簡寫 (例如 115/4/30 或 115-04-30)
        m_tw = re.match(r'^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})$', v_str)
        if m_tw:
            y, m, d = int(m_tw.group(1)), int(
                m_tw.group(2)), int(m_tw.group(3))
            if y < 1000:
                y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"

        # 處理只有月跟日的狀況 (例如 4/30 或 04-30)
        m_md = re.match(r'^(\d{1,2})[/-](\d{1,2})$', v_str)
        if m_md:
            import datetime
            # 使用 sidebar_date 的年份，若無則 fallback 到當前年份
            selected_d = st.session_state.get('sidebar_date', datetime.datetime.now())
            curr_y = selected_d.year
            m, d = int(m_md.group(1)), int(m_md.group(2))
            return f"{curr_y:04d}-{m:02d}-{d:02d}"

        try:
            p = pd.to_datetime(v_str)
            if pd.notna(p):
                return p.strftime('%Y-%m-%d')
        except:
            pass
        return v_str
    df['date'] = df['date'].apply(fix_d)
    return df


def get_daily_data(d_str):
    try:
        # 走集中快取層 (60s TTL)，避免大量 rerun 打爆 API
        df = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel)
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
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"讀取 occ_data 失敗: {repr(e)}\n{traceback.format_exc()}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
    return {}


def save_daily_data(d_str, data_dict):
    try:
        df = conn.read(worksheet="occ_data", ttl="0")
        if df is None:
            df = pd.DataFrame()

        df = standardize_df_dates(df)

        data_dict['date'] = d_str
        new_row = pd.DataFrame([data_dict])

        if 'date' in df.columns and d_str in df['date'].values:
            df = df[df['date'] != d_str]

        df = pd.concat([df, new_row], ignore_index=True)
        if 'date' in df.columns:
            df = df.sort_values('date').reset_index(drop=True)
        conn.update(worksheet="occ_data", data=df.fillna(""))
        _get_occ_data_cached_v2.clear()
        fetch_yearly_metrics.clear()
        return True
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"儲存失敗: {e}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
        return False


def get_monthly_target(month_str):
    try:
        df = read_google_sheet("targets", ttl="1h")
        if df is not None and not df.empty:
            res = df[df['month'] == month_str]
            if not res.empty:
                return int(res.iloc[0]['target_revenue'])
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"讀取 targets 失敗: {e}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
    return 0


def save_monthly_target(month_str, target):
    try:
        df = conn.read(worksheet="targets", ttl="0")
        if df is None or df.empty:
            df = pd.DataFrame(columns=["month", "target_revenue"])

        if month_str in df['month'].values:
            df.loc[df['month'] == month_str, 'target_revenue'] = target
        else:
            new_row = pd.DataFrame(
                [{"month": month_str, "target_revenue": target}])
            df = pd.concat([df, new_row], ignore_index=True)

        conn.update(worksheet="targets", data=df.fillna(""))
        _get_cached_sheet_v3.clear()
        return True
    except:
        return False


def get_daily_log(d_str):
    try:
        df = read_google_sheet("daily_logs", ttl="1h")
        if df is not None and not df.empty:
            res = df[df['date'] == d_str]
            if not res.empty:
                return str(res.iloc[0]['log']).strip()
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"讀取 daily_logs 失敗: {e}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
    # Fallback to daily_data if not found in daily_logs (backward compatibility)
    # 直接走快取，不額外打 API
    try:
        df_old = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel)
        if df_old is not None and not df_old.empty:
            df_old = standardize_df_dates(df_old)
            res = df_old[df_old['date'] == d_str]
            if not res.empty and 'daily_work_log' in res.columns:
                return str(res.iloc[0]['daily_work_log']).strip()
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"讀取 occ_data (fallback) 失敗: {repr(e)}\n{traceback.format_exc()}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
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
        _get_cached_sheet_v3.clear()
        st.toast(f"✅ {d_str} 日誌已自動對齊 Google Sheet！")
        return True
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"日誌儲存失敗: {e}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
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
        df_all = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel)
        month_str = f"{year}-{month:02d}"
        df = df_all[df_all['date'].str.startswith(
            month_str, na=False)].sort_values('date')
    except:
        return "--- 讀取失敗 ---"

    if df.empty:
        return "--- 當月無紀錄 ---"

    full_report = ""
    for d in sorted(df['date'].unique()):
        full_report += generate_report_text(d) + "\n\n"
    return full_report


def minguo_to_western(d_str):
    """
    將 民國/月/日 (如 115/03/02 或 0115/03/02) 轉換為 Python date 對象。
    """
    if pd.isna(d_str) or not isinstance(d_str, str):
        return None
    try:
        # 移除前導零並拆分
        parts = d_str.strip().split('/')
        if len(parts) == 3:
            year = int(parts[0])
            # 如果是 115 或 0115，這應是民國年
            if year < 1000:  # 民國年編號通常不大於 1000
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
        df_all = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel)
        if df_all is not None and not df_all.empty:
            df_all = standardize_df_dates(df_all)
            # 確保日期唯一，避免重複加總
            df_all = df_all.drop_duplicates(subset='date', keep='last')
            df = df_all[(df_all['date'] >= m_start) &
                        (df_all['date'] <= m_end)].copy()
        else:
            df = pd.DataFrame()
    except Exception:
        df = pd.DataFrame()

    res = {
        'rev': 0.0, 'rooms': 0.0, 'sold_rooms': 0.0, 'sellable': 0.0, 'occ90_days': 0,
        'avg_occ': 0.0, 'avg_adr': 0.0, 'revpar': 0.0, 'df': df,
        'month_label': f"{year}-{month:02d}"
    }

    if not df.empty:
        # 確保數值欄位為 float
        num_cols = ['revenue', 'total_rooms', 'sold_rooms', 'occ_rate', 'adr']
        for c in num_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(
                    ',', '').str.replace('%', '')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

        for _, r in df.iterrows():
            rev = float(r.get('revenue', 0.0))
            rm = float(r.get('total_rooms', 0.0))
            sold = float(r.get('sold_rooms', 0.0))
            occ = float(r.get('occ_rate', 0.0))

            if rm > 0 or rev > 0 or sold > 0:
                res['rev'] += rev
                res['rooms'] += rm
                res['sold_rooms'] += sold
                if occ >= 90.0:
                    res['occ90_days'] += 1

        res['avg_occ'] = (res['sold_rooms'] / res['rooms'] * 100.0) if res['rooms'] > 0 else 0.0
        res['avg_adr'] = (res['rev'] / res['sold_rooms']) if res['sold_rooms'] > 0 else 0.0
        res['revpar'] = res['avg_adr'] * (res['avg_occ'] / 100.0)

    return res


def get_other_revenue(year_month_str):
    """
    從 EIS_data 分頁取得特定月份（例如 '202601'）的「其他收入」。
    白名單項目：['電話費', '影印費', '傳真費', '冰箱飲料', '洗衣費', '付費電視', '清潔費', '場租', '設施使用費', '雜項', '客房其他收入', '餐飲費', '商品收入']
    """
    try:
        df = _get_cached_sheet_v3("EIS_data", hotel_type=current_hotel)
    except Exception:
        return 0.0
        
    if df is None or df.empty:
        return 0.0
    
    # 清理欄位名稱
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    if '年/月' not in df.columns or '項目' not in df.columns or '本月實際' not in df.columns:
        col_ym = next((c for c in df.columns if '年' in c and '月' in c), None)
        col_item = next((c for c in df.columns if '項目' in c), None)
        col_actual = next((c for c in df.columns if '本月' in c and '實際' in c), None)
        if not col_ym or not col_item or not col_actual:
            return 0.0
    else:
        col_ym = '年/月'
        col_item = '項目'
        col_actual = '本月實際'
        
    # 向下填滿「年月」欄位
    df[col_ym] = df[col_ym].replace('', pd.NA).ffill()
    df[col_ym] = df[col_ym].astype(str).str.strip().str.replace('.0', '', regex=False)
    
    # 篩選月份
    df_month = df[df[col_ym] == str(year_month_str)]
    if df_month.empty:
        return 0.0
        
    # 白名單
    whitelist = ['電話費', '影印費', '傳真費', '冰箱飲料', '洗衣費', '付費電視', '清潔費', '場租', '設施使用費', '雜項', '客房其他收入', '餐飲費', '商品收入']
    
    # 過濾與加總
    df_month = df_month.copy()
    df_month[col_item] = df_month[col_item].astype(str).str.strip()
    mask = df_month[col_item].isin(whitelist)
    
    df_filtered = df_month[mask].copy()
    if df_filtered.empty:
        return 0.0
        
    # 將本月實際轉為數值
    df_filtered[col_actual] = df_filtered[col_actual].astype(str).str.strip().str.replace(',', '').str.replace('%', '')
    df_filtered[col_actual] = pd.to_numeric(df_filtered[col_actual], errors='coerce').fillna(0.0)
    
    return float(df_filtered[col_actual].sum())

@st.cache_data(ttl=3600)
def fetch_yearly_metrics(year):
    try:
        df_all = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel)
        if df_all is None or df_all.empty:
            return 0.0, 0.0
        df_all = standardize_df_dates(df_all)
        df_all = df_all.drop_duplicates(subset='date', keep='last')

        y_start = f"{year}-01-01"
        y_end = f"{year}-12-31"
        df = df_all[(df_all['date'] >= y_start) &
                    (df_all['date'] <= y_end)].copy()
        if df.empty:
            return 0.0, 0.0

        num_cols = ['revenue', 'total_rooms']
        for c in num_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(
                    ',', '').str.replace('%', '')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

        tot_rev = df['revenue'].sum()
        tot_rms = df['total_rooms'].sum()
        yearly_adr = (tot_rev / tot_rms) if tot_rms > 0 else 0.0

        taipei_events_df = fetch_taipei_events()
        e_dates = set(taipei_events_df['date'].unique(
        )) if not taipei_events_df.empty else set()

        h_dates = set()
        for m in range(1, 13):
            h_dict = fetch_holidays_for_month(year, m)
            for d_str, info in h_dict.items():
                if info['flags']:
                    h_dates.add(d_str)

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
st.sidebar.caption(
    f"🚀 最後更新時間: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
st.sidebar.header("📅 日期選擇")
_today = datetime.date.today()
if 'sidebar_date' not in st.session_state:
    # 首次建立 session：設定為今日，並記錄初始化日期
    st.session_state['sidebar_date'] = _today
    st.session_state['_session_init_date'] = str(_today)
else:
    # Session 已存在：若初始化日期不是今天，代表 session 是舊的，重設日期到今日
    _init_date = st.session_state.get('_session_init_date', '')
    if _init_date != str(_today):
        st.session_state['sidebar_date'] = _today
        st.session_state['_session_init_date'] = str(_today)

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
                    if isinstance(default_val, int):
                        norm_db = int(float(db_val))
                    elif isinstance(default_val, float):
                        norm_db = float(db_val)
                    else:
                        norm_db = str(db_val)
                except:
                    norm_db = default_val

            # 判斷是否真的改變
            if isinstance(curr_val, float) and isinstance(norm_db, float):
                import math
                if not math.isclose(curr_val, norm_db, abs_tol=1e-5):
                    has_changes = True
            elif curr_val != norm_db:
                has_changes = True

    if has_changes:
        save_daily_data(target_d_str, update_dict)

    # 單獨同步日誌
    if 'input_daily_log' in st.session_state:
        curr_log = st.session_state['input_daily_log'].strip()
        db_log = str(get_daily_log(target_d_str) or "").strip()
        if curr_log != db_log:
            save_daily_log(target_d_str, curr_log)


def prev_day():
    st.session_state['sidebar_date'] -= datetime.timedelta(days=1)

def next_day():
    st.session_state['sidebar_date'] += datetime.timedelta(days=1)

def prev_month():
    import calendar
    curr = st.session_state['sidebar_date']
    first_day = curr.replace(day=1)
    prev_m_last = first_day - datetime.timedelta(days=1)
    st.session_state['sidebar_date'] = prev_m_last

def next_month():
    import calendar
    curr = st.session_state['sidebar_date']
    if curr.month == 12:
        next_m_first = curr.replace(year=curr.year+1, month=1, day=1)
    else:
        next_m_first = curr.replace(month=curr.month+1, day=1)
    _, last_day = calendar.monthrange(next_m_first.year, next_m_first.month)
    st.session_state['sidebar_date'] = next_m_first.replace(day=last_day)

col1, col2 = st.sidebar.columns(2)
col1.button("⬅️ 前月", on_click=prev_month)
col2.button("後月 ➡️", on_click=next_month)

if current_hotel == "採購":
    def on_purchasing_date_change():
        # 當使用者手動選取日期時，強制將其轉換為該月最後一天
        import calendar
        d = st.session_state['sidebar_date']
        _, last_day = calendar.monthrange(d.year, d.month)
        st.session_state['sidebar_date'] = d.replace(day=last_day)
        
    selected_raw = st.sidebar.date_input(
        "選擇月份 (點選該月任一天，將自動跳轉至月底)", 
        key='sidebar_date',
        on_change=on_purchasing_date_change
    )
    selected_date = st.session_state['sidebar_date']
else:
    selected_raw = st.sidebar.date_input(
        "選擇日期", 
        key='sidebar_date'
    )
    selected_date = st.session_state['sidebar_date']

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
if current_hotel != "採購":
    weekly_options = ["--- 關閉週預覽 ---",
                      "第1週 (1-7號)", "第2週 (8-14號)", "第3週 (15-21號)", "第4週 (22-28號)", "第5週 (29號起)"]
    selected_week = st.sidebar.selectbox(
        "快速查閱區間：", weekly_options, index=0, key="weekly_view_select")
else:
    selected_week = "--- 關閉週預覽 ---"

st.sidebar.divider()
st.sidebar.subheader("🔄 系統快取管理")
if st.sidebar.button("強制重新同步資料", help="從 Google Sheets 重新拉取最新資料 (若資料沒有更新時可使用)"):
    st.cache_data.clear()
    st.rerun()

# --------------------------------------------------


day_data = get_daily_data(date_str)
if st.session_state.get('_last_loaded_date') != date_str or st.session_state.get('_last_week_view') != selected_week:
    for ss_key, (db_col, default_val) in field_mapping.items():
        val = day_data.get(db_col)
        # Handle nan/null from Pandas/SQLite gracefully
        if pd.isna(val) or val is None:
            st.session_state[ss_key] = default_val
        else:
            if isinstance(default_val, int):
                st.session_state[ss_key] = int(val)
            elif isinstance(default_val, float):
                st.session_state[ss_key] = float(val)
            else:
                st.session_state[ss_key] = str(val)

    # 獲取日誌
    st.session_state['input_daily_log'] = get_daily_log(date_str)

    st.session_state['_last_loaded_date'] = date_str
    st.session_state['_last_week_view'] = selected_week
    st.session_state['_data_is_loaded'] = True  # 標記為已載入，此後任何變動或換日才允許存檔


def on_input_change():
    # 使用 session_state 中的當前日期，確保 callback 觸發時日期正確
    target_d = st.session_state.get('_actual_current_date')
    if target_d:
        sync_st_to_db(target_d)


if current_hotel != "採購":
    st.sidebar.divider()
    st.sidebar.subheader("📤 數據匯出與備份")


    def generate_report_text(d_str):
        data = get_daily_data(d_str)
        if not data:
            return f"--- {d_str} 無紀錄 ---"

        report = []
        report.append(f"========================================")
        report.append(f"🏨 路徒行旅 Plus {current_hotel} - 營運日誌 ({d_str})")
        report.append(f"========================================\n")

        def safe_int_val(v):
            try:
                if pd.isna(v) or v is None:
                    return 0
                return int(float(v))
            except:
                return 0

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
        report.append(
            f"- Happy Hour: {safe_int_val(data.get('rest_hh_guests', 0))} 人")
        report.append(
            f"- 餐廳營收(全月): {safe_int_val(data.get('rest_month_rev', 0))} 元\n")

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
            # 匯出按鈕走快取即可，不需要強制最新
            df_all = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel)
            df_month = df_all[df_all['date'].str.startswith(
                month_str, na=False)].sort_values('date')

            if df_month.empty:
                st.sidebar.warning(f"⚠️ {month_str} 尚無任何資料。")
            else:
                with st.sidebar.status("正在產生報表...", expanded=False):
                    full_month_text = f"【路徒行旅 Plus {current_hotel} {month_str} 全月營運紀錄匯總】\n\n"
                    for d in df_month['date']:
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


def sync_from_eis_local(d_str):
    """
    從 Y:\\ 網路磁碟自動讀取站前館 EIS XLS，解析當日核心營運數據。
    僅適用於站前館（current_hotel == '站前館'）。
    回傳 dict {'revenue', 'total_rooms', 'adr', 'occ_rate'} 或 None（失敗時）。
    """
    import xlrd
    import os

    try:
        year = d_str[:4]
        month = d_str[5:7]
        day = d_str[8:10]
        folder = os.path.join("Y:\\", "櫃台", "金旭", "每日EIS", year, f"{year}{month}")
        filename = f"EIS03{year}{month}{day}.XLS"
        filepath = os.path.join(folder, filename)

        if not os.path.exists(filepath):
            import datetime as _dt
            _req_date = _dt.date(int(year), int(month), int(day))
            _today_date = _dt.date.today()
            if _req_date >= _today_date:
                _hint = f"\n⚠️ 提示：{_req_date.strftime('%Y/%m/%d')} 的 EIS 報表通常在當日深夜或隔日凌晨 3 點後才會產生，請稍後再試。"
            else:
                _hint = f"\n⚠️ 提示：{_req_date.strftime('%Y/%m/%d')} 的 EIS 檔案可能未產生（例如休館日），請確認。"
            return None, f"找不到 EIS 檔案：{filepath}\n請確認 Y:\\ 磁碟已連線，且當日報表已產生。{_hint}"

        wb = xlrd.open_workbook(filepath, encoding_override='cp950')
        ws = wb.sheet_by_index(1)  # 工作表 1 = 住客來源

        # 合計行在 Row08（欄位：col2=收入, col3=間數, col5=ADR, col6=OCC小數）
        r = 8
        revenue = int(ws.cell_value(r, 2))
        total_rooms = int(ws.cell_value(r, 3))
        adr = int(ws.cell_value(r, 5))
        occ_raw = ws.cell_value(r, 6)
        occ_rate = round(occ_raw * 100, 1) if occ_raw <= 1.0 else round(occ_raw, 1)

        return {
            'revenue': revenue,
            'total_rooms': total_rooms,
            'adr': adr,
            'occ_rate': occ_rate,
        }, None

    except Exception as e:
        return None, f"讀取 EIS 失敗：{e}"


def batch_sync_from_eis_local(start_date, end_date):
    """
    批次掃描 start_date ~ end_date 範圍內所有存在的 EIS XLS 並解析。
    回傳 (results, errors)：
      results = [{'date': str, 'revenue': int, 'total_rooms': int, 'adr': int, 'occ_rate': float}, ...]
      errors  = [{'date': str, 'reason': str}, ...]
    """
    import xlrd
    import os
    import datetime

    results = []
    errors = []

    cur = start_date
    while cur <= end_date:
        d_str = str(cur)
        year = d_str[:4]
        month = d_str[5:7]
        day = d_str[8:10]
        folder = os.path.join("Y:\\", "櫃台", "金旭", "每日EIS", year, f"{year}{month}")
        filename = f"EIS03{year}{month}{day}.XLS"
        filepath = os.path.join(folder, filename)

        if not os.path.exists(filepath):
            errors.append({'date': d_str, 'reason': '檔案不存在（可能為休館日或尚未產生）'})
        else:
            try:
                wb = xlrd.open_workbook(filepath, encoding_override='cp950')
                ws = wb.sheet_by_index(1)
                r = 8
                revenue = int(ws.cell_value(r, 2))
                total_rooms = int(ws.cell_value(r, 3))
                adr = int(ws.cell_value(r, 5))
                occ_raw = ws.cell_value(r, 6)
                occ_rate = round(occ_raw * 100, 1) if occ_raw <= 1.0 else round(occ_raw, 1)
                results.append({
                    'date': d_str,
                    'revenue': revenue,
                    'total_rooms': total_rooms,
                    'adr': adr,
                    'occ_rate': occ_rate,
                })
            except Exception as e:
                errors.append({'date': d_str, 'reason': f'解析失敗：{e}'})

        cur += datetime.timedelta(days=1)

    return results, errors


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
            df = pd.read_csv(file, skiprows=header_idx) if is_csv else pd.read_excel(
                file, skiprows=header_idx)
        except Exception as e:
            # 嘗試不同的 engine
            file.seek(0)
            df = pd.read_excel(file, skiprows=header_idx, engine='openpyxl')

        df.columns = df.columns.astype(str).str.replace(
            r'[\s\n\r]', '', regex=True)

        date_col = next((c for c in df.columns if '日期' in c), None)
        occ_col = next(
            (c for c in df.columns if '住房率' in c or '訂房率' in c or '出租率' in c or 'OCC' in c.upper()), None)
        adr_col = next(
            (c for c in df.columns if '平均房價' in c or 'ADR' in c.upper()), None)

        rev_col = next(
            (c for c in df.columns if '客房收入' in c or '客房營收' in c or '總營收' in c or '營業額' in c or '實際營收' in c), None)
        rooms_col = next((c for c in df.columns if (
            '住房數' in c or '出租' in c or '售出' in c or '實住' in c) and '可售' not in c), None)
        if not rooms_col:
            rooms_col = next((c for c in df.columns if (
                '房間數' in c or '客房數' in c) and '可售' not in c), None)

        if not date_col:
            st.error("⚠️ 解析失敗：找不到『日期』欄位，請檢查報表格式。")
            return False

        # --- 強化日期解析邏輯 ---
        def robust_parse_date(val):
            if pd.isna(val) or str(val).strip() == '':
                return None
            s = str(val).strip().split('.')[0]  # 移除 .0
            # 嘗試 YYYYMMDD
            try:
                if len(s) == 8 and s.isdigit():
                    return pd.to_datetime(s, format='%Y%m%d').date()
            except:
                pass
            # 嘗試一般解析 (YYYY-MM-DD, YYYY/MM/DD 等)
            try:
                return pd.to_datetime(s).date()
            except:
                pass
            return None

        df['標準日期'] = df[date_col].apply(robust_parse_date)

        df_new_records = pd.DataFrame()
        updates = []
        for index, row in df.iterrows():
            d_obj = row['標準日期']
            if not d_obj:
                continue

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
                except:
                    occ = 0.0

            adr = int(float(str(row.get(adr_col, '0')).replace(',', ''))
                      ) if adr_col and pd.notna(row.get(adr_col)) else 0
            rev = int(float(str(row.get(rev_col, '0')).replace(',', ''))
                      ) if rev_col and pd.notna(row.get(rev_col)) else 0
            rooms = int(float(str(row.get(rooms_col, '0')).replace(
                ',', ''))) if rooms_col and pd.notna(row.get(rooms_col)) else 0

            updates.append({'date': d_str, 'occ_rate': occ,
                           'adr': adr, 'revenue': rev, 'total_rooms': rooms})

        if updates:
            df_existing = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel).copy()
            if df_existing is None:
                df_existing = pd.DataFrame()
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

            conn.update(worksheet="occ_data", data=df_final.fillna(""))
            _get_occ_data_cached_v2.clear()
            fetch_yearly_metrics.clear()
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
                            clean_val = s_val.replace('NT$', '').replace(
                                '$', '').replace(',', '').strip()
                            month_rev = int(float(clean_val))
                            break
                        except:
                            continue
            # 尋找客單價
            if '平均客單價' in row_str or '客單價' in row_str:
                for val in row:
                    s_val = str(val).strip()
                    if any(c.isdigit() for c in s_val) and '客單價' not in s_val:
                        try:
                            clean_val = s_val.replace('NT$', '').replace(
                                '$', '').replace(',', '').strip()
                            avg_spent = int(float(clean_val))
                            break
                        except:
                            continue

        parsed_days = []
        # 2. 尋找每日明細 (修正 Regex 讓其更具包容度)
        for i, row in df.iterrows():
            col0 = str(row[0]).strip()
            m = re.search(r'(\d{1,2})/(\d{1,2})', col0)
            if m:
                month_val, day_val = m.groups()
                d_str = f"{current_year}-{int(month_val):02d}-{int(day_val):02d}"

                def safe_int(val):
                    if pd.isna(val) or str(val).strip() == '':
                        return 0
                    try:
                        # 處理 Excel 讀入時可能的科學符號或逗號
                        return int(float(str(val).replace(',', '').strip()))
                    except:
                        return 0

                # 假設欄位順序不變 (根據 Roaders Plus 常用報表格式)
                # 早餐相關 (1-6)
                row_vals = row.values.tolist()
                bf_theme_est = safe_int(
                    row_vals[1]) if len(row_vals) > 1 else 0
                bf_theme_act = safe_int(
                    row_vals[2]) if len(row_vals) > 2 else 0
                bf_zq_est = safe_int(row_vals[3]) if len(row_vals) > 3 else 0
                bf_zq_act = safe_int(row_vals[4]) if len(row_vals) > 4 else 0
                bf_total_est = safe_int(
                    row_vals[5]) if len(row_vals) > 5 else 0
                bf_total_act = safe_int(
                    row_vals[6]) if len(row_vals) > 6 else 0

                # 下午茶相關 (7-12)
                af_theme_est = safe_int(
                    row_vals[7]) if len(row_vals) > 7 else 0
                af_theme_act = safe_int(
                    row_vals[8]) if len(row_vals) > 8 else 0
                af_zq_est = safe_int(row_vals[9]) if len(row_vals) > 9 else 0
                af_zq_act = safe_int(row_vals[10]) if len(row_vals) > 10 else 0
                af_total_est = safe_int(row_vals[11]) if len(
                    row_vals) > 11 else 0
                af_total_act = safe_int(row_vals[12]) if len(
                    row_vals) > 12 else 0

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
            df_existing = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel).copy()
            if df_existing is None:
                df_existing = pd.DataFrame()

            # 重要：確保現有資料的 date 也是字串，否則 combine_first 的 join 會失效
            df_existing = standardize_df_dates(df_existing)

            # --- 修復：如果成功解析出月結算營收或客單價，強制更新現有資料庫中該月份的所有紀錄 ---
            # 避免使用者先點擊了未來的日期產生了帶有舊營收的紀錄，導致 MTD 永遠抓到最後一天的舊數據
            months = set("-".join(str(d['date']).split('-')[:2])
                         for d in parsed_days)
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
            conn.update(worksheet="occ_data", data=df_final.fillna(""))
            _get_occ_data_cached_v2.clear()
            fetch_yearly_metrics.clear()

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
current_hotel = st.session_state.get("hotel_type", "站前館")

st.title(f"Hotel Master - {current_hotel}")

# 主畫面：依角色決定顯示哪些分頁
st.sidebar.divider()
st.sidebar.subheader("📌 系統功能導覽")
if current_hotel == "採購":
    menu_options = ["💰 採購分析", "🛒 菜價分析"]
else:
    menu_options = [
        "📊 營運總覽", "📈 月分析專區", "📝 每日營運紀錄", 
        "💰 採購分析", "🛒 菜價分析", "🧹 房務數據", 
        "🍽️ 餐廳數據", "🔧 工務數據", "👥 人事概況", 
        "🌍 國籍分析", "📉 渠道分析", "📋 營運檢討報告"
    ]
selected_page = st.sidebar.radio("請選擇功能：", menu_options, label_visibility="collapsed")


if current_hotel != "採購":
    if selected_page == "📊 營運總覽":
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
                if pd.isna(v) or v is None:
                    return 0
                return int(float(v))
            except:
                return 0

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
            st.metric("今日待修房數", f"{repairs} 間", delta="🔴 需處理" if repairs >
                      0 else "🟢 正常", delta_color="off")
        with col3:
            st.error("🍽️ **餐廳狀況**")
            bf_total_act = st.session_state.get('input_bf_total_act', 0)
            st.metric("今日雙館早餐總來客", f"{safe_format_int(bf_total_act)} 人")

        st.divider()

        # -- 月度累計模式 (MTD Analysis) --
        st.subheader(f"📅 本月累計分析 (MTD: {selected_date.strftime('%Y-%m')})")
        start_of_month = selected_date.replace(day=1).strftime('%Y-%m-%d')

        try:
            df_all = _get_cached_sheet_v3("occ_data", hotel_type=current_hotel)
            if df_all is not None and not df_all.empty:
                df_all = standardize_df_dates(df_all)
                df_all = df_all.drop_duplicates(subset='date', keep='last')
                df_mtd = df_all[(df_all['date'] >= start_of_month)
                                & (df_all['date'] <= date_str)].copy()
            else:
                df_mtd = pd.DataFrame()
        except Exception as e:
            st.sidebar.error(f"⚠️ 讀取數據時發生錯誤: {e}")
            df_mtd = pd.DataFrame()

        if not df_mtd.empty:
            for col in ['occ_rate', 'adr', 'revenue', 'total_rooms']:
                if col in df_mtd.columns:
                    df_mtd[col] = pd.to_numeric(df_mtd[col].astype(
                        str).str.replace(',', '').str.replace('%', ''), errors='coerce').fillna(0)

            mtd_rooms = 0.0
            mtd_rev = 0.0
            total_sellable = 0.0

            for _, r in df_mtd.iterrows():
                def clean_num(val):
                    if pd.isna(val):
                        return 0.0
                    try:
                        return float(str(val).replace(',', '').replace('%', ''))
                    except:
                        return 0.0

                o = clean_num(r.get('occ_rate'))
                adr = clean_num(r.get('adr'))
                rev = clean_num(r.get('revenue'))
                rm = clean_num(r.get('total_rooms'))

                if rev == 0 and adr > 0 and rm > 0:
                    rev = adr * rm
                if rm == 0 and rev > 0 and adr > 0:
                    rm = rev / adr

                if rm > 0 or rev > 0:
                    mtd_rooms += rm
                    mtd_rev += rev
                    if o > 0:
                        total_sellable += (rm / (o / 100.0))

            mtd_occ = (mtd_rooms / total_sellable * 100.0) if total_sellable > 0 else 0.0
            mtd_adr = (mtd_rev / mtd_rooms) if mtd_rooms > 0 else 0.0

            # --- 讀取 F&B 資料（直接呼叫獨立函數，不走 daily_data 快取）---
            fb = compute_fb_mtd(start_of_month, date_str, _dummy_hotel=current_hotel)
            rest_mrev = fb['revenue']
            grand_total_rev = mtd_rev + rest_mrev

            # 顯示四大指標
            st.write("##### 🏨 房務營運 MTD")
            c1, c2, c3 = st.columns(3)
            c1.markdown(make_card(
                "MTD 累計住房率", f"{mtd_occ:.1f}%", "card-theme-blue", "card-bg-dark", "🏨"), unsafe_allow_html=True)
            c2.markdown(make_card("MTD 累計 ADR", f"NT$ {int(mtd_adr):,}",
                        "card-theme-green", "card-bg-dark", "💳"), unsafe_allow_html=True)
            c3.markdown(make_card("MTD 房務累計營收", f"NT$ {int(mtd_rev):,}",
                        "card-theme-orange", "card-bg-dark", "💰"), unsafe_allow_html=True)

            st.write("##### 🏁 全館合併營收 (MTD)")
            g1, g2 = st.columns([1, 2])
            g1.markdown(make_card("餐廳結算營收", f"NT$ {int(rest_mrev):,}",
                        "card-theme-purple", "card-bg-dark", "🍽️"), unsafe_allow_html=True)
            g2.markdown(make_card("✨ 全館 MTD 總營收", f"NT$ {int(grand_total_rev):,}",
                        "card-theme-red", "card-bg-dark", "🚀"), unsafe_allow_html=True)

            st.markdown(
                "<br><hr style='margin: 5px 0; border: 1px dashed #ddd;'>", unsafe_allow_html=True)
            st.write("##### 🍽️ 餐廳營運累計 (MTD)")

            if fb['error']:
                st.warning(f"⚠️ F&B 資料載入提示：{fb['error']}")

            mtd_bf_theme = fb['bf_theme_act']
            mtd_bf_zq    = fb['bf_zq_act']
            mtd_af_theme = fb['af_theme_act']
            mtd_af_zq    = fb['af_zq_act']
            mtd_total_bf_act = fb['total_act_bf']
            mtd_total_af_act = fb['total_act_af']

            active_bf_days = fb['matched_days']
            active_af_days = fb.get('active_af_days', active_bf_days)
            avg_total_bf = (mtd_total_bf_act / active_bf_days) if active_bf_days > 0 else 0
            avg_total_af = (mtd_total_af_act / active_af_days) if active_af_days > 0 else 0
            mtd_avg_total = avg_total_bf + avg_total_af

            rest_avg_spent = fb['avg_spent']

            st.markdown(
                "<h6 style='color:#555; margin-top:15px;'>📌【站前館】MTD 累計</h6>", unsafe_allow_html=True)
            sz1, sz2, sz3 = st.columns(3)
            sz1.markdown(make_card(
                "早餐 (實際)", f"{int(mtd_bf_zq)} 人", "card-theme-orange", "", "🥐"), unsafe_allow_html=True)
            sz2.markdown(make_card(
                "下午茶 (實際)", f"{int(mtd_af_zq)} 人", "card-theme-purple", "", "🍰"), unsafe_allow_html=True)
            sz3.markdown(make_card(
                "站前合計 (實際)", f"{int(mtd_bf_zq + mtd_af_zq)} 人", "card-theme-blue", "", "👥"), unsafe_allow_html=True)

            st.markdown(
                "<h6 style='color:#555; margin-top:20px;'>📌【主題館】MTD 累計</h6>", unsafe_allow_html=True)
            st1, st2, st3 = st.columns(3)
            st1.markdown(make_card(
                "早餐 (實際)", f"{int(mtd_bf_theme)} 人", "card-theme-orange", "", "🥐"), unsafe_allow_html=True)
            st2.markdown(make_card(
                "下午茶 (實際)", f"{int(mtd_af_theme)} 人", "card-theme-purple", "", "🍰"), unsafe_allow_html=True)
            st3.markdown(make_card(
                "主題合計 (實際)", f"{int(mtd_bf_theme + mtd_af_theme)} 人", "card-theme-blue", "", "👥"), unsafe_allow_html=True)

            st.markdown(
                "<h6 style='color:#555; margin-top:20px;'>👑【兩館合併總覽】</h6>", unsafe_allow_html=True)
            m1, m2, m3, m4 = st.columns(4)
            m1.markdown(make_card("兩館早餐 (實際)", f"{int(mtd_total_bf_act)} 人",
                        "card-theme-orange", "card-bg-dark", "🥐"), unsafe_allow_html=True)
            m2.markdown(make_card("兩館下午茶 (實際)", f"{int(mtd_total_af_act)} 人",
                        "card-theme-purple", "card-bg-dark", "🍰"), unsafe_allow_html=True)
            m3.markdown(make_card("全月結算營收", f"NT$ {int(rest_mrev):,}",
                        "card-theme-green", "card-bg-dark", "💰"), unsafe_allow_html=True)
            m4.markdown(make_card("平均客單價", f"NT$ {int(rest_avg_spent):,}",
                        "card-theme-red", "card-bg-dark", "🧾"), unsafe_allow_html=True)

            st.markdown(
                "<h6 style='color:#555; margin-top:20px;'>📉【兩館日平均來客】</h6>", unsafe_allow_html=True)
            a1, a2, a3 = st.columns(3)
            a1.markdown(make_card("兩館早餐平均", f"{avg_total_bf:.1f} 人/日",
                        "card-theme-orange", "", "✨"), unsafe_allow_html=True)
            a2.markdown(make_card(
                "兩館下午茶平均", f"{avg_total_af:.1f} 人/日", "card-theme-purple", "", "✨"), unsafe_allow_html=True)
            a3.markdown(make_card(
                "兩館整體總平均", f"{mtd_avg_total:.1f} 人/日", "card-theme-blue", "", "📈"), unsafe_allow_html=True)

        else:
            st.info("💡 資料庫中目前尚未有這個月的記錄。")


if current_hotel != "採購":
    if selected_page == "📈 月分析專區":
        st.header("📈 月分析專區")

        # 1. 取得四個月數據 (M-2, M-1, M, M+1)
        prev_prev_m_date = get_month_delta(selected_date, -2)
        prev_m_date = get_month_delta(selected_date, -1)
        next_m_date = get_month_delta(selected_date, 1)

        m_prev_prev = fetch_month_summary(
            prev_prev_m_date.year, prev_prev_m_date.month)
        m_prev = fetch_month_summary(prev_m_date.year, prev_m_date.month)
        m_curr = fetch_month_summary(selected_date.year, selected_date.month)
        m_next = fetch_month_summary(next_m_date.year, next_m_date.month)

        # 取得去年同月數據 (YoY)
        m_curr_ly = fetch_month_summary(
            selected_date.year - 1, selected_date.month)

        st.markdown("#### 🏆 本月總覽與去年同期比較 (YoY)")
        if not m_curr['df'].empty and not m_curr_ly['df'].empty:
            col1, col2, col3 = st.columns(3)

            adr_diff = m_curr['avg_adr'] - m_curr_ly['avg_adr']
            adr_pct = (adr_diff / m_curr_ly['avg_adr']
                       * 100) if m_curr_ly['avg_adr'] > 0 else 0
            adr_color = "#2ecc71" if adr_diff >= 0 else "#e74c3c"
            adr_sign = "+" if adr_diff >= 0 else ""
            col1.markdown(
                f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {adr_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>當月平均 ADR</p><strong style='font-size:22px;'>NT$ {int(m_curr['avg_adr']):,}</strong><p style='margin:5px 0 0 0; font-size:13px; color:{adr_color}; font-weight:bold;'>較去年同期 {adr_sign}NT$ {int(adr_diff):,} ({adr_sign}{adr_pct:.1f}%)</p></div>", unsafe_allow_html=True)

            occ_diff = m_curr['avg_occ'] - m_curr_ly['avg_occ']
            occ_color = "#2ecc71" if occ_diff >= 0 else "#e74c3c"
            occ_sign = "+" if occ_diff >= 0 else ""
            col2.markdown(
                f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {occ_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>當月平均 OCC</p><strong style='font-size:22px;'>{m_curr['avg_occ']:.1f}%</strong><p style='margin:5px 0 0 0; font-size:13px; color:{occ_color}; font-weight:bold;'>較去年同期 {occ_sign}{occ_diff:.1f}%</p></div>", unsafe_allow_html=True)

            rev_diff = m_curr['rev'] - m_curr_ly['rev']
            rev_pct = (rev_diff / m_curr_ly['rev']
                       * 100) if m_curr_ly['rev'] > 0 else 0
            rev_color = "#2ecc71" if rev_diff >= 0 else "#e74c3c"
            rev_sign = "+" if rev_diff >= 0 else ""
            col3.markdown(
                f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {rev_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>當月總客房營收</p><strong style='font-size:22px;'>NT$ {int(m_curr['rev']):,}</strong><p style='margin:5px 0 0 0; font-size:13px; color:{rev_color}; font-weight:bold;'>較去年同期 {rev_sign}NT$ {int(rev_diff):,} ({rev_sign}{rev_pct:.1f}%)</p></div>", unsafe_allow_html=True)
            st.markdown("<div style='margin-bottom:20px;'></div>",
                        unsafe_allow_html=True)
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

            y_adr, y_pure_adr = fetch_yearly_metrics(
                int(month_data['month_label'].split('-')[0]))
            df['y_adr_baseline'] = y_adr
            df['y_adr_text'] = ''
            df['y_pure_adr_baseline'] = y_pure_adr
            df['y_pure_adr_text'] = ''

            if not df.empty:
                if avg_adr > 0:
                    df.loc[df.index[-1],
                           'adr_baseline_text'] = f"${int(avg_adr):,} (月)"
                if y_adr > 0:
                    df.loc[df.index[-1], 'y_adr_text'] = f"${int(y_adr):,} (年)"
                if y_pure_adr > 0:
                    df.loc[df.index[-1],
                           'y_pure_adr_text'] = f"${int(y_pure_adr):,} (純平)"

            dt = pd.to_datetime(df['date'])
            df['day'] = dt.dt.day
            weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
            df['weekday'] = dt.dt.weekday.map(weekday_map)
            df['label'] = df['day'].astype(str) + " (" + df['weekday'] + ")"

            df['color_category'] = df['occ_rate'].apply(
                lambda x: '>=90' if x >= 90.0 else ('>=80' if x >= 80.0 else '<80'))

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
                            e_label = EVENT_TYPE_LABELS.get(
                                row['event_type'], '[活]')
                            if e_label not in e_flags:
                                e_flags.append(e_label)
                    return h_flags + e_flags

                # 建立多層標籤資料 (最多支援 5 層垂直堆疊，避免過度擁擠)
                for i in range(5):
                    df[f'flag_{i}'] = df['date'].apply(lambda d: get_combined_flags_list(
                        d)[i] if len(get_combined_flags_list(d)) > i else '')
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
                y=alt.Y('occ_rate:Q', title='住房率 (%)',
                        scale=alt.Scale(domain=[0, 100])),
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
            valid_adrs = df[df['adr'] >
                            0]['adr'] if not df.empty else pd.Series([])
            if not valid_adrs.empty:
                adr_min = max(0, int(valid_adrs.min() * 0.9))
                adr_max = int(valid_adrs.max() * 1.1)
            else:
                adr_min = 2000
                adr_max = 8000

            avg_adr = month_data.get('avg_adr', 0)


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
                x=alt.X('label:O', sort=df['label'].tolist()),  # 自然與 OCC X 軸合併
                tooltip=['date', 'occ_rate', 'adr']
            )

            adr_line = base_adr.mark_line(color='#ff9f43', strokeWidth=3, interpolate='monotone').encode(
                y=alt.Y('adr:Q', title='平均房價 (NT$)', axis=alt.Axis(
                    titleColor='#ff9f43', format='$,.0f'), scale=adr_scale)
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
                y_adr_rule = alt.Chart(df).mark_rule(color='#f1c40f', strokeWidth=1.5, strokeDash=[
                    5, 5]).encode(y=alt.Y('y_adr_baseline:Q', scale=adr_scale))
                y_adr_text = alt.Chart(df).mark_text(
                    align='left', baseline='middle', dx=8, dy=-14, color='#000000', fontSize=11, fontWeight='bold'
                ).encode(
                    x=alt.X('label:O', sort=df['label'].tolist()), y=alt.Y('y_adr_baseline:Q', scale=adr_scale), text='y_adr_text:N'
                )
                adr_layers.extend([y_adr_rule, y_adr_text])

            # 繪製年純平日 ADR 黑色基準線
            if df.get('y_pure_adr_baseline', pd.Series()).max() > 0:
                yp_adr_rule = alt.Chart(df).mark_rule(color='#000000', strokeWidth=1.5, strokeDash=[
                    5, 5]).encode(y=alt.Y('y_pure_adr_baseline:Q', scale=adr_scale))
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

        with col_chart1:
            render_occ_chart(m_prev_prev, "(前前月)")
        with col_chart2:
            render_occ_chart(m_prev, "(上月)")
        with col_chart3:
            render_occ_chart(m_curr, "(本月)")
        with col_chart4:
            render_occ_chart(m_next, "(下月)")

        # --- A2. 去年同期軌跡對比 (YoY Daily Comparison) ---
        st.markdown("#### 📅 去年同期軌跡對比 (YoY Daily Comparison)")
        if not m_curr['df'].empty and not m_curr_ly['df'].empty:
            df_ty = m_curr['df'].copy()
            df_ly = m_curr_ly['df'].copy()

            if 'day' not in df_ty.columns:
                df_ty['day'] = pd.to_datetime(df_ty['date']).dt.day
            if 'day' not in df_ly.columns:
                df_ly['day'] = pd.to_datetime(df_ly['date']).dt.day

            df_ty['year'] = '今年'
            df_ly['year'] = '去年'

            df_yoy = pd.concat([df_ty[['day', 'adr', 'year']],
                               df_ly[['day', 'adr', 'year']]], ignore_index=True)
            df_yoy['adr'] = pd.to_numeric(df_yoy['adr'], errors='coerce').fillna(0)

            # 設定 Y 軸比例尺
            yoy_adr_min = max(0, int(df_yoy['adr'].min() * 0.9))
            yoy_adr_max = int(df_yoy['adr'].max() * 1.1)
            if yoy_adr_min == yoy_adr_max:
                yoy_adr_max += 1000

            yoy_chart = alt.Chart(df_yoy).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X('day:O', title='日期 (Day of Month)'),
                y=alt.Y('adr:Q', title='平均房價 (NT$)', scale=alt.Scale(
                    domain=[yoy_adr_min, yoy_adr_max], zero=False)),
                color=alt.Color('year:N',
                                scale=alt.Scale(domain=['今年', '去年'], range=[
                                                '#ff9f43', '#bdc3c7']),
                                legend=alt.Legend(title="年份", orient="top-left")
                                ),
                strokeDash=alt.condition(
                    alt.datum.year == '去年', alt.value([5, 5]), alt.value([0])),
                tooltip=['day', 'year', 'adr']
            ).properties(height=350)

            st.altair_chart(yoy_chart, use_container_width=True)

        st.markdown("<div style='margin-bottom:30px;'></div>",
                    unsafe_allow_html=True)

        # --- B. 關鍵表現數據分析 ---
        st.markdown("#### 🌟 關鍵表現數據分析")



        curr_metrics = calc_key_metrics(m_curr)
        prev_metrics = calc_key_metrics(m_prev)
        pprev_metrics = calc_key_metrics(m_prev_prev)
        next_metrics = calc_key_metrics(m_next)

        def metric_diff_card(label, diff, target_label, unit="天"):
            color = '#2ecc71' if diff >= 0 else '#e74c3c'
            status = '本月多' if diff > 0 else '較少' if diff < 0 else '持平'
            return f'<div style="flex: 1; min-width: 150px; background: #fff; padding: 12px; border-radius: 8px; border: 1px solid #eee; margin-bottom: 10px;"><p style="margin:0; font-size:12px; color:#999;">與 {target_label} 相比</p><div style="display: flex; align-items: baseline; gap: 8px; margin-top: 5px;"><strong style="font-size:18px; color:{color};">{abs(diff)} {unit}</strong><span style="font-size:11px; color:#666;">({status})</span></div></div>'

        # 天數差異
        diff_adr_pp = curr_metrics['high_adr_days'] - \
            pprev_metrics['high_adr_days']
        diff_adr_p = curr_metrics['high_adr_days'] - prev_metrics['high_adr_days']
        diff_adr_n = curr_metrics['high_adr_days'] - next_metrics['high_adr_days']

        diff_dual_pp = curr_metrics['dual_match_days'] - \
            pprev_metrics['dual_match_days']
        diff_dual_p = curr_metrics['dual_match_days'] - \
            prev_metrics['dual_match_days']
        diff_dual_n = curr_metrics['dual_match_days'] - \
            next_metrics['dual_match_days']

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
                <p style="margin:0; font-size:14px; color:#666;">📊 <strong>高低營收分析 (本月)</strong></p>
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
            scatter_df['occ_val'] = pd.to_numeric(
                scatter_df['occ_rate'], errors='coerce').fillna(0)
            scatter_df['adr_val'] = pd.to_numeric(
                scatter_df['adr'], errors='coerce').fillna(0)
            scatter_df['day'] = pd.to_datetime(scatter_df['date']).dt.day

            # 以「年純平 ADR」作為 Y 軸分界（最客觀的裸實力底線，不受淡日拉低）
            y_adr_s, y_pure_adr_s = fetch_yearly_metrics(selected_date.year)
            adr_anchor = y_pure_adr_s if y_pure_adr_s > 0 else m_curr.get(
                'avg_adr', scatter_df['adr_val'].mean())
            anchor_label = f'年純平 ADR ${int(adr_anchor):,}'
            anchor_color = '#000000'
            occ_threshold = 85.0  # 高住房率門檻

            def classify_quadrant(row):
                hi_occ = row['occ_val'] >= occ_threshold
                hi_adr = row['adr_val'] >= adr_anchor
                if hi_occ and hi_adr:
                    return '🟠 理想（高OCC+高ADR）'
                if hi_occ and not hi_adr:
                    return '🔴 賤賣（高OCC+低ADR）'
                if not hi_occ and hi_adr:
                    return '🟡 定價偏高（低OCC+高ADR）'
                return '🔵 淡季死水（低OCC+低ADR）'

            scatter_df['象限'] = scatter_df.apply(classify_quadrant, axis=1)

            color_map = {
                '🟠 理想（高OCC+高ADR）': '#ff9f43',
                '🔴 賤賣（高OCC+低ADR）': '#e74c3c',
                '🟡 定價偏高（低OCC+高ADR）': '#f1c40f',
                '🔵 淡季死水（低OCC+低ADR）': '#3498db',
            }

            scatter_chart = alt.Chart(scatter_df).mark_circle(size=100, opacity=0.8).encode(
                x=alt.X('occ_val:Q', title='住房率 (%)',
                        scale=alt.Scale(domain=[0, 105])),
                y=alt.Y('adr_val:Q', title='平均房價 ADR (NT$)',
                        scale=alt.Scale(zero=False)),
                color=alt.Color('象限:N',
                                scale=alt.Scale(
                                    domain=list(color_map.keys()),
                                    range=list(color_map.values())
                                ),
                                legend=alt.Legend(
                                    title="象限分類", orient="bottom", columns=2)
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

            # 85% OCC 垂直輔助線
            occ_rule = alt.Chart(pd.DataFrame({'x': [occ_threshold]})).mark_rule(
                color='#7f8c8d', strokeDash=[6, 3], strokeWidth=1.5
            ).encode(x='x:Q')
            occ_label = alt.Chart(pd.DataFrame({'x': [occ_threshold], 'y': [scatter_df['adr_val'].max() * 1.05], 'text': ['85% OCC 門檻']})).mark_text(
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
            success_rate = (ideal_cnt / high_occ_total *
                            100) if high_occ_total > 0 else 0

            # 計算上個月的定價成功率作為對比
            prev_scatter_df = m_prev['df'].copy()
            prev_success_rate = 0
            if not prev_scatter_df.empty:
                prev_scatter_df['occ_val'] = pd.to_numeric(
                    prev_scatter_df['occ_rate'], errors='coerce').fillna(0)
                prev_scatter_df['adr_val'] = pd.to_numeric(
                    prev_scatter_df['adr'], errors='coerce').fillna(0)
                prev_scatter_df['hi_occ'] = prev_scatter_df['occ_val'] >= occ_threshold
                prev_scatter_df['hi_adr'] = prev_scatter_df['adr_val'] >= adr_anchor
                prev_ideal = int(
                    (prev_scatter_df['hi_occ'] & prev_scatter_df['hi_adr']).sum())
                prev_cheap = int(
                    (prev_scatter_df['hi_occ'] & ~prev_scatter_df['hi_adr']).sum())
                prev_total = prev_ideal + prev_cheap
                prev_success_rate = (prev_ideal / prev_total *
                                     100) if prev_total > 0 else 0

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
                    v_suffix = f" <span style='color:#777;'>@{row['venue']}</span>" if pd.notna(
                        row['venue']) and str(row['venue']).strip() != "" else ""
                    e_list.append(f"🏟️ {row['event_name']}{v_suffix}")
                    e_labels.append(EVENT_TYPE_LABELS.get(
                        row['event_type'], '[活]'))

            if h_info or e_list:
                all_flags = (h_info['flags'] if h_info else "") + \
                    "".join(sorted(list(set(e_labels))))
                details_html = ""
                if h_info:
                    details_html += f"<div style='margin-bottom:4px; color:#856404;'>🌍 {h_info['details']}</div>"
                if e_list:
                    details_html += "<div style='color:#2c3e50;'>" + \
                        "<br>".join(e_list) + "</div>"

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
            e_dates = set(taipei_events_df['date'].unique(
            )) if not taipei_events_df.empty else set()

            curr_df['is_h'] = curr_df['date'].isin(h_dates)
            curr_df['is_e'] = curr_df['date'].isin(e_dates)
            curr_df['is_any'] = curr_df['is_h'] | curr_df['is_e']

            def render_impact_row(df, condition_col, title, icon):
                holiday_days = df[df[condition_col]]
                non_holiday_days = df[~df[condition_col]]

                h_occ = holiday_days['occ_rate'].mean() if len(
                    holiday_days) > 0 else 0
                h_adr = holiday_days['revenue'].sum() / holiday_days['total_rooms'].sum() if len(
                    holiday_days) > 0 and holiday_days['total_rooms'].sum() > 0 else 0
                nh_occ = non_holiday_days['occ_rate'].mean() if len(
                    non_holiday_days) > 0 else 0
                nh_adr = non_holiday_days['revenue'].sum() / non_holiday_days['total_rooms'].sum(
                ) if len(non_holiday_days) > 0 and non_holiday_days['total_rooms'].sum() > 0 else 0

                diff_occ = h_occ - nh_occ
                diff_adr = h_adr - nh_adr

                st.markdown(f"**{icon} {title}**")
                c1, c2, c3 = st.columns(3)
                c1.markdown(
                    f"<div style='background:#f1f8ff; padding:10px; border-radius:5px; border-left:3px solid #3498db; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>有標籤 ({len(holiday_days)}天)</p><strong style='font-size:16px;'>{h_occ:.1f}% / NT$ {int(h_adr):,}</strong></div>", unsafe_allow_html=True)
                c2.markdown(
                    f"<div style='background:#f8f9fa; padding:10px; border-radius:5px; border-left:3px solid #ccc; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>無標籤 ({len(non_holiday_days)}天)</p><strong style='font-size:16px;'>{nh_occ:.1f}% / NT$ {int(nh_adr):,}</strong></div>", unsafe_allow_html=True)
                color = "#2ecc71" if diff_occ >= 0 else "#e74c3c"
                c3.markdown(
                    f"<div style='background:#f0fff4; padding:10px; border-radius:5px; border-left:3px solid #2ecc71; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>帶動效益</p><strong style='font-size:16px; color:{color};'>{diff_occ:+.1f}% / NT$ {int(diff_adr):+,}</strong></div>", unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:15px;'></div>",
                            unsafe_allow_html=True)

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
                    adr = sub_df['revenue'].sum(
                    ) / sub_df['total_rooms'].sum() if sub_df['total_rooms'].sum() > 0 else 0
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
                st.markdown("<div style='margin-bottom:20px;'></div>",
                            unsafe_allow_html=True)

            render_impact_row(curr_df, 'is_any', "綜合分析 (假日 + 台北活動)", "📊")
            render_impact_row(curr_df, 'is_h', "僅外國節慶分析", "🌍")
            render_impact_row(curr_df, 'is_e', "僅台北重大活動分析", "🏟️")

            st.markdown("<div style='margin-bottom:20px;'></div>",
                        unsafe_allow_html=True)
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

            hist_df = pd.concat([m1_sum['df'], m2_sum['df'],
                                m3_sum['df']], ignore_index=True)

            if not hist_df.empty:
                # 準備歷史資料的標籤
                def get_hist_flags(row):
                    d = row['date']
                    y, m = int(d.split('-')[0]), int(d.split('-')[1])
                    h_f = fetch_holidays_for_month(
                        y, m).get(d, {}).get('flags', '')
                    e_f = ""
                    if not taipei_events_df.empty:
                        de = taipei_events_df[taipei_events_df['date'] == d]
                        for _, r in de.iterrows():
                            e_f += EVENT_TYPE_LABELS.get(r['event_type'], '[活]')
                    return (h_f != ''), (e_f != '')

                # 為了效能，預先抓取這幾個月的假日資料
                hist_h_dates = set()
                for md in [m1_date, m2_date, m3_date]:
                    hd = fetch_holidays_for_month(md.year, md.month)
                    for d, info in hd.items():
                        if info['flags']:
                            hist_h_dates.add(d)

                hist_df['is_h'] = hist_df['date'].isin(hist_h_dates)
                hist_df['is_e'] = hist_df['date'].isin(
                    set(taipei_events_df['date'].unique())) if not taipei_events_df.empty else False
                hist_df['is_any'] = hist_df['is_h'] | hist_df['is_e']

                render_impact_row(hist_df, 'is_any', "綜合分析 (過去三個月合計)", "📊")
                render_impact_row(hist_df, 'is_h', "僅外國節慶分析 (過去三個月合計)", "🌍")
                render_impact_row(hist_df, 'is_e', "僅台北重大活動分析 (過去三個月合計)", "🏟️")

                st.markdown("<div style='margin-bottom:20px;'></div>",
                            unsafe_allow_html=True)
                render_exclusive_matrix(hist_df, "(過去三個月合計)")
            else:
                st.info("尚無足夠的歷史數據進行長期趨勢分析。")

            with st.expander("📅 查看本月所有假日與台北活動詳細清單"):
                # Combine details for expander
                all_dates = sorted(set(list(h_dict.keys(
                )) + (taipei_events_df['date'].tolist() if not taipei_events_df.empty else [])))
                has_any = False
                for d in all_dates:
                    if d.startswith(f"{y_str}-{m_str}"):
                        h_info = h_dict.get(d, {'flags': '', 'details': []})
                        e_info = ""
                        if not taipei_events_df.empty:
                            de = taipei_events_df[taipei_events_df['date'] == d]
                            for _, r in de.iterrows():
                                v_suffix = f" @{r['venue']}" if pd.notna(
                                    r['venue']) and str(r['venue']).strip() != "" else ""
                                e_info += f", 🏟️ {r['event_name']}{v_suffix} ({r['event_type']})"

                        if h_info['flags'] or e_info:
                            st.markdown(
                                f"- **{d}** {h_info['flags']}{e_info}: {', '.join(h_info['details'])}")
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
            st.markdown(
                f"<p style='text-align:center; color:#777; margin-bottom:10px;'>{label} ({month_data['month_label']})</p>", unsafe_allow_html=True)
            if not month_data['df'].empty:
                st.markdown(make_card(
                    "當月總營收", f"NT$ {int(month_data['rev']):,}", "card-theme-orange", "card-bg-dark"), unsafe_allow_html=True)
                st.markdown(make_card(
                    "當月平均房價", f"NT$ {int(month_data['avg_adr']):,}", "card-theme-green", "card-bg-dark"), unsafe_allow_html=True)
                st.markdown(make_card(
                    "當月住房率", f"{month_data['avg_occ']:.1f}%", "card-theme-blue", "card-bg-dark"), unsafe_allow_html=True)
                st.markdown(make_card(
                    "當月 RevPAR", f"NT$ {int(month_data['revpar']):,}", "card-theme-purple", "card-bg-dark"), unsafe_allow_html=True)
            else:
                st.info("暫無數據")

        with col_m1:
            render_metric_col(m_prev_prev, "⏪ 前前月")
        with col_m2:
            render_metric_col(m_prev, "◀️ 上月")
        with col_m3:
            render_metric_col(m_curr, "✨ 本月")
        with col_m4:
            render_metric_col(m_next, "▶️ 下月")

        # --- D. 月度營運指標 - 關鍵差異 ---
        st.markdown("#### 🔍 月度營運指標：關鍵差異對比 (本月 vs 其他月份)")

        def calculate_diff_row(current_val, compare_val, is_currency=True, is_percent=False):
            if compare_val == 0:
                return "<span style='color:#777;'>-</span>"
            diff = current_val - compare_val
            if is_currency:
                diff_str = f"{'▲' if diff >= 0 else '▼'} NT$ {abs(int(diff)):,}"
            elif is_percent:
                diff_str = f"{'▲' if diff >= 0 else '▼'} {abs(diff):.1f}%"
            else:
                diff_str = f"{'▲' if diff >= 0 else '▼'} {abs(diff):.1f}"

            color = "#2ecc71" if diff >= 0 else "#e74c3c"  # 增加為綠色，減少為紅色
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
        m_rev = m_curr['rev']  # 使用剛剛計算好的本月營收

        t_col1, t_col2 = st.columns([1, 2])
        with t_col1:
            new_target = st.number_input(f"設定 {month_key} 目標業績 (NT$)", min_value=0,
                                         step=10000, value=current_target, key=f"target_input_{month_key}")
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
            if not m_curr['df'].empty and 'revenue' in m_curr['df'].columns:
                active_days = m_curr['df'][pd.to_numeric(m_curr['df']['revenue'], errors='coerce').fillna(0) > 0]
                elapsed_days = len(active_days)
            else:
                elapsed_days = 0

            if elapsed_days > 0:
                import calendar
                total_days = calendar.monthrange(
                    selected_date.year, selected_date.month)[1]
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
                t_card = make_card(
                    "距離目標還差", f"NT$ {int(gap):,}", "card-theme-red", "", "🎯")
            a_col1.markdown(t_card, unsafe_allow_html=True)
            a_col2.markdown(make_card(
                "超標目標 (+10%)", f"NT$ {int(stretch_goal):,}", "card-theme-orange", "", "🚀"), unsafe_allow_html=True)
            if stretch_gap <= 0:
                s_card = make_card("超標達成狀況", "🔥 已超標達成！",
                                   "card-theme-green", "card-bg-dark", "🏆")
            else:
                s_card = make_card(
                    "距離超標還差", f"NT$ {int(stretch_gap):,}", "card-theme-purple", "", "⚡")
            a_col3.markdown(s_card, unsafe_allow_html=True)
        else:
            st.info("💡 請在上方輸入本月目標業績，系統將自動為您計算達標差距。")

if current_hotel != "採購":
    if selected_page == "🧹 房務數據":
        st.header("🧹 房務數據")
        with st.form(f"form_hk_{date_str}"):
            st.number_input("今日總清消房數", min_value=0, step=1, key="input_cleaned")
            st.number_input("退/續數量", min_value=0, step=1, key="input_hk_co")
            st.number_input("每人平均掃房數", min_value=0.0, step=0.1, key="input_hk_avg")
            st.number_input("房務請購費用", min_value=0, step=100, key="input_hk_exp")
            if st.form_submit_button("💾 儲存房務數據", type="primary", use_container_width=True):
                sync_st_to_db(date_str)
                st.success("✅ 房務數據已儲存！")

if current_hotel != "採購":
    if selected_page == "🍽️ 餐廳數據":
        st.header("🍽️ 餐廳數據")
        st.subheader("📁 數據報表上傳")
        rest_file = st.file_uploader("上傳餐廳報表 (Excel)，會自動把整份報表寫入資料庫！", type=[
                                     "xls", "xlsx"], key="rest_uploader")

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
                                try:
                                    p_month_rev = int(float(str(v).replace(
                                        'NT$', '').replace('$', '').replace(',', '').strip()))
                                    break
                                except:
                                    pass
                    if '客單價' in row_str:
                        for v in row:
                            if any(c.isdigit() for c in str(v)) and '客單價' not in str(v):
                                try:
                                    p_avg_spent = int(float(str(v).replace(
                                        'NT$', '').replace('$', '').replace(',', '').strip()))
                                    break
                                except:
                                    pass
                    if re.search(r'\d{1,2}/\d{1,2}', str(row[0])):
                        found_days += 1

                pv_col1, pv_col2, pv_col3 = st.columns(3)
                pv_col1.metric("辨識出月結算營收", f"NT$ {p_month_rev:,}")
                pv_col2.metric("辨識出平均客單價", f"NT$ {p_avg_spent:,}")
                pv_col3.metric("辨識出每日明細", f"{found_days} 筆")

                if p_month_rev == 0:
                    st.warning("⚠️ 系統未能從報表中自動找到「月結算營收」，請確認報表格式或手動檢查。")

                if st.button("📥 確認無誤，寫入系統資料庫", key="rest_btn"):
                    saved_count = parse_and_save_restaurant(
                        rest_file, selected_date.year)
                    if saved_count:
                        st.success(f"✅ 成功更新 {saved_count} 筆每日餐廳資料！")
                        time.sleep(1)
                        st.rerun()
            except Exception as ex:
                st.error(f"預覽失敗: {ex}")

        st.divider()
        st.subheader(f"餐廳手動確認區 ({date_str})")

        with st.form(f"form_rest_{date_str}"):
            st.markdown("#### 🌞 早餐數據")
            b1, b2, b3 = st.columns(3)
            b1.number_input("【主題】預估來客", min_value=0, step=1, key="input_bf_theme_est")
            b1.number_input("【主題】實際來客", min_value=0, step=1, key="input_bf_theme_act")

            b2.number_input("【站前】預估來客", min_value=0, step=1, key="input_bf_zq_est")
            b2.number_input("【站前】實際來客", min_value=0, step=1, key="input_bf_zq_act")

            b3.number_input("【兩館總和】預估", min_value=0, step=1, key="input_bf_total_est")
            b3.number_input("【兩館總和】實際", min_value=0, step=1, key="input_bf_total_act")

            st.markdown("#### 🍰 下午茶數據")
            a1, a2, a3 = st.columns(3)
            a1.number_input("【主題】預估來客", min_value=0, step=1, key="input_af_theme_est")
            a1.number_input("【主題】實際來客", min_value=0, step=1, key="input_af_theme_act")

            a2.number_input("【站前】預估來客", min_value=0, step=1, key="input_af_zq_est")
            a2.number_input("【站前】實際來客", min_value=0, step=1, key="input_af_zq_act")

            a3.number_input("【兩館總和】預估", min_value=0, step=1, key="input_af_total_est")
            a3.number_input("【兩館總和】實際", min_value=0, step=1, key="input_af_total_act")

            st.markdown("#### 📊 月報結算總數與雜項")
            c1, c2, c3 = st.columns(3)
            c1.number_input("已結算營收 (全月)", min_value=0, step=100, key="input_rest_mrev")
            c2.number_input("平均客單價", min_value=0, step=10, key="input_rest_aspent")
            c3.number_input("THE PEAK 請購費用", min_value=0, step=100, key="input_rest_exp")

            col_rest1, col_rest2 = st.columns(2)
            col_rest1.number_input("The Peak 當日來客數", min_value=0, step=1, key="input_peak_act")
            col_rest2.number_input("Happy Hour 當日來客數", min_value=0, step=1, key="input_hh_act")

            if st.form_submit_button("💾 儲存餐廳數據", type="primary", use_container_width=True):
                sync_st_to_db(date_str)
                st.success("✅ 餐廳數據已儲存！")


if current_hotel != "採購":
    if selected_page == "🔧 工務數據":
        st.header("🔧 工務數據")
        with st.form(f"form_maint_{date_str}"):
            st.number_input("今日待修房數", min_value=0, step=1, key="input_repair")
            st.text_area("修繕紀錄", key="input_maint_rec")
            st.number_input("工務請購費用", min_value=0, step=100, key="input_maint_exp")
            if st.form_submit_button("💾 儲存工務數據", type="primary", use_container_width=True):
                sync_st_to_db(date_str)
                st.success("✅ 工務數據已儲存！")

if current_hotel != "採購":
    if selected_page == "📝 每日營運紀錄":
        st.header("📝 每日營運紀錄")

        # --- 金旭報表上傳 + 手動輸入 (從原「櫃台數據」移入) ---
        with st.expander("📁 金旭報表上傳 & 當日數字手動確認", expanded=False):

            # === EIS 自動同步區塊（僅站前館） ===
            if current_hotel == "站前館":
                st.markdown("#### 🔄 從 EIS 自動同步當日數據")
                st.caption(f"自動讀取 Y:\\\\櫃台\\金旭\\每日EIS 中 **{date_str}** 的報表，解析住房率、ADR、營收、間數。")
                eis_col1, eis_col2 = st.columns([1, 3])
                with eis_col1:
                    eis_sync_btn = st.button("🔄 從 EIS 同步", key="eis_sync_btn", type="primary")
                with eis_col2:
                    if eis_sync_btn:
                        with st.spinner("正在讀取 EIS 報表..."):
                            eis_data, eis_err = sync_from_eis_local(date_str)
                        if eis_err:
                            st.error(f"❌ {eis_err}")
                        elif eis_data:
                            # 預覽解析結果
                            st.success("✅ EIS 解析成功！以下為讀取結果：")
                            p1, p2, p3, p4 = st.columns(4)
                            p1.metric("住房率", f"{eis_data['occ_rate']:.1f}%")
                            p2.metric("ADR", f"NT$ {eis_data['adr']:,}")
                            p3.metric("總營收", f"NT$ {eis_data['revenue']:,}")
                            p4.metric("出租間數", f"{eis_data['total_rooms']} 間")
                            st.session_state['_eis_pending'] = eis_data
                            st.session_state['_eis_pending_date'] = date_str

                # 確認寫入按鈕
                if st.session_state.get('_eis_pending') and st.session_state.get('_eis_pending_date') == date_str:
                    eis_data_pending = st.session_state['_eis_pending']
                    if st.button("💾 確認寫入 Google Sheet", key="eis_confirm_btn"):
                        existing = get_daily_data(date_str)
                        merged = {**existing, **eis_data_pending, 'date': date_str}
                        if save_daily_data(date_str, merged):
                            # 同步更新 session_state 讓欄位即時反映
                            st.session_state['input_occ'] = eis_data_pending['occ_rate']
                            st.session_state['input_adr'] = eis_data_pending['adr']
                            st.session_state['input_rev'] = eis_data_pending['revenue']
                            st.session_state['input_rooms'] = eis_data_pending['total_rooms']
                            del st.session_state['_eis_pending']
                            del st.session_state['_eis_pending_date']
                            st.success(f"✅ {date_str} 的 EIS 數據已成功寫入 Google Sheet！")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("❌ 寫入失敗，請重試。")

                st.divider()

                # === 批次同步區塊 ===
                st.markdown("#### 📅 批次同步（多日補登）")
                st.caption("若有幾天未即時同步，可在此選擇日期範圍，一次補齊所有 EIS 資料。")
                b_col1, b_col2 = st.columns(2)
                batch_start = b_col1.date_input("起始日期", value=selected_date - datetime.timedelta(days=3), key="eis_batch_start")
                batch_end = b_col2.date_input("結束日期", value=selected_date - datetime.timedelta(days=1), key="eis_batch_end")

                if st.button("🔍 掃描可用的 EIS 報表", key="eis_batch_scan"):
                    if batch_start > batch_end:
                        st.error("❌ 起始日期不能晚於結束日期！")
                    elif (batch_end - batch_start).days > 60:
                        st.error("❌ 一次最多掃描 60 天。")
                    else:
                        with st.spinner(f"正在掃描 {batch_start} ~ {batch_end} 的 EIS 報表..."):
                            b_results, b_errors = batch_sync_from_eis_local(batch_start, batch_end)
                        st.session_state['_eis_batch_results'] = b_results
                        st.session_state['_eis_batch_errors'] = b_errors

                # 顯示掃描結果
                if st.session_state.get('_eis_batch_results') is not None:
                    b_results = st.session_state['_eis_batch_results']
                    b_errors = st.session_state.get('_eis_batch_errors', [])

                    if b_results:
                        st.success(f"✅ 找到 **{len(b_results)}** 天的 EIS 報表可同步：")
                        preview_df = pd.DataFrame(b_results)[['date', 'occ_rate', 'adr', 'revenue', 'total_rooms']]
                        preview_df.columns = ['日期', '住房率(%)', 'ADR', '總營收', '出租間數']
                        preview_df['住房率(%)'] = preview_df['住房率(%)'].apply(lambda x: f"{x:.1f}%")
                        preview_df['ADR'] = preview_df['ADR'].apply(lambda x: f"NT$ {x:,}")
                        preview_df['總營收'] = preview_df['總營收'].apply(lambda x: f"NT$ {x:,}")
                        st.dataframe(preview_df, use_container_width=True, hide_index=True)

                    if b_errors:
                        with st.expander(f"⚠️ {len(b_errors)} 天無法同步（點擊查看）"):
                            for e in b_errors:
                                st.caption(f"- {e['date']}：{e['reason']}")

                    if b_results:
                        if st.button(f"💾 確認批次寫入 {len(b_results)} 天資料", key="eis_batch_confirm", type="primary"):
                            success_count = 0
                            fail_count = 0
                            with st.spinner("批次寫入中..."):
                                for row in b_results:
                                    existing = get_daily_data(row['date'])
                                    merged = {**existing, **{k: v for k, v in row.items() if k != 'date'}, 'date': row['date']}
                                    if save_daily_data(row['date'], merged):
                                        success_count += 1
                                    else:
                                        fail_count += 1
                            del st.session_state['_eis_batch_results']
                            if '_eis_batch_errors' in st.session_state:
                                del st.session_state['_eis_batch_errors']
                            if success_count:
                                st.success(f"✅ 成功寫入 {success_count} 天！{'⚠️ ' + str(fail_count) + ' 天失敗。' if fail_count else ''}")
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error("❌ 全部寫入失敗，請重試。")
                st.divider()


            # === 手動上傳金旭報表（多日批次） ===
            jinxu_file = st.file_uploader(
                "上傳金旭報表 (Excel/CSV)，會自動把整份報表寫入資料庫！", type=["csv", "xls", "xlsx"], key="jinxu_uploader")
            if jinxu_file:
                if st.button("📥 寫入系統資料庫"):
                    saved_count = parse_and_save_jinxu(jinxu_file)
                    if saved_count:
                        st.success(f"✅ 成功將 {saved_count} 筆每日資料存入系統資料庫！切換日期即可自動調出。")
                        time.sleep(1)
                        st.rerun()
            st.divider()
            st.subheader(f"📋 當日數字手動確認 ({date_str})")
            with st.form(f"form_daily_{date_str}"):
                rc1, rc2, rc3 = st.columns(3)
                rc1.number_input("訂房率 (%)", min_value=0.0, max_value=100.0,
                                 step=0.1, key="input_occ")
                rc2.number_input("ADR (平均房價)", min_value=0, step=10, key="input_adr")
                rc3.number_input("總營收", min_value=0, step=100, key="input_rev")
                rc4, rc5 = st.columns(2)
                rc4.number_input("總住房數", min_value=0, step=1, key="input_rooms")
                rc5.number_input("櫃台請購費用", min_value=0, step=100, key="input_counter_exp")
                st.text_area("負評客訴", key="input_complaints")
                if st.form_submit_button("💾 儲存當日數據", type="primary", use_container_width=True):
                    sync_st_to_db(date_str)
                    st.success("✅ 當日數據已儲存！")


        if selected_week != "--- 關閉週預覽 ---":
            import calendar
            _, last_day_of_month = calendar.monthrange(
                selected_date.year, selected_date.month)

            # 解析選擇的區間
            week_idx = weekly_options.index(selected_week)
            start_d = (week_idx - 1) * 7 + 1
            if week_idx == 5:
                end_d = last_day_of_month
            else:
                end_d = min(start_d + 6, last_day_of_month)

            st.subheader(f"📋 {selected_week} 快速審視模式")
            st.info(
                f"正在查看 {selected_date.year}年度 {selected_date.month}月份 ({start_d}號 至 {end_d}號) 的完整紀錄。")

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
            st.info(
                f"💡 請在下方詳細填寫 **{date_str}** 的各項營運日誌與重點工作回報。這裡的紀錄會自動儲存，切換日期或關閉網頁也不用擔心遺失。")
            st.text_area("✍️ 今日工作與營運細節報告：", height=500, key="input_daily_log",
                         placeholder="可以在這裡記錄交班重點、客訴特殊處理、VIP 接待細節、設備大修紀錄...等", on_change=on_input_change)

if selected_page == "💰 採購分析":
    st.header("💰 採購花費分析統計")

    current_month_str = selected_date.strftime('%Y-%m')

    try:
        # 讀取採購數據 (降低 TTL 以確保更新及時)
        df_purchase = get_purchase_data_cached()

        if df_purchase is not None and not df_purchase.empty:
            # 清理欄位名稱 (移除空格)
            df_purchase.columns = df_purchase.columns.astype(str).str.strip()

            # 尋找關鍵欄位 (自動識別可能的名稱變體)
            date_col = next(
                (c for c in df_purchase.columns if '日期' in c or 'Date' in c), None)
            dept_col = next(
                (c for c in df_purchase.columns if '部門' in c or 'Dept' in c or '工地' in c), None)
            total_col = next(
                (c for c in df_purchase.columns if '小計' in c or '金額' in c or 'Total' in c), None)

            if not date_col or not dept_col or not total_col:
                missing = [c for c, found in [
                    ('日期', date_col), ('部門', dept_col), ('小計', total_col)] if not found]
                st.error(f"❌ 採購分頁缺少必要欄位：{', '.join(missing)}")
                st.write("目前偵測到的欄位有：", list(df_purchase.columns))
                st.stop()

            # 展開可查看偵測結果的小將器（方便尋找問題）
            with st.expander("🔍 資料偵測資訊（診斷中，確認後可收起）", expanded=True):
                st.write(f"日期欄：`{date_col}` ｜ 部門欄：`{dept_col}` ｜ 金額欄：`{total_col}`")
                st.write("所有欄位：", list(df_purchase.columns))
                sample_dates = df_purchase[date_col].dropna().head(5).tolist()
                st.write("前 5 筆日期**原始內容**（解析前）：", sample_dates)
                sample_depts = df_purchase[dept_col].dropna().unique()[:10].tolist()
                st.write("部門欄樣本內容：", sample_depts)
                st.write(f"purchase_data 總筆數：{len(df_purchase)} 筆")
                
                # --- Debug dump ---
                try:
                    df_purchase[[date_col, dept_col, total_col]].dropna().head(50).to_csv("scratch/debug_purchase_dates.csv", index=False, encoding='utf-8-sig')
                except Exception:
                    pass


            # 確保日期欄位為日期型態 (支援民國年與一般西元年)
            def robust_date_parse(val):
                if pd.isna(val):
                    return None
                s = str(val).strip().replace('.0', '')  # 處理 202505.0 浮點格式
                if not s or s in ('nan', 'None', 'NaT'):
                    return None
                # 民國年格式 (含 /)
                if '/' in s:
                    res = minguo_to_western(s)
                    if res:
                        return res
                # YYYYMM 6位數字 (如 202505) → 當月1日
                import re as _re_dp
                if _re_dp.match(r'^\d{6}$', s):
                    try:
                        return pd.to_datetime(s, format='%Y%m').date()
                    except:
                        pass
                # YYYYMMDD 8位數字 (如 20250525)
                if _re_dp.match(r'^\d{8}$', s):
                    try:
                        return pd.to_datetime(s, format='%Y%m%d').date()
                    except:
                        pass
                # 其他標準格式
                try:
                    return pd.to_datetime(val).date()
                except:
                    return None

            df_purchase['日期'] = df_purchase[date_col].apply(robust_date_parse)

            # 處理部門欄位空值 (歸類到「未分類」)
            df_purchase[dept_col] = df_purchase[dept_col].fillna(
                "未分類").astype(str).str.strip()
            df_purchase.loc[df_purchase[dept_col] == "", dept_col] = "未分類"

            # 過濾 NaT/None
            df_purchase = df_purchase[df_purchase['日期'].notna()]
            
            # --- 新增：找出官方 purchase_data 中已經有紀錄的所有 Year-Month ---
            official_ym = set(pd.to_datetime(df_purchase['日期']).dt.strftime('%Y-%m').dropna().unique())

            # --- 新增：將 thepeak_daily_purchase_report 無縫匯入 df_purchase 以供全域呈現 ---
            df_daily_report = fetch_thepeak_daily_purchase_report()
            if not df_daily_report.empty and '總價' in df_daily_report.columns:
                append_df = pd.DataFrame()
                if '請購日期' in df_daily_report.columns:
                    append_df['日期'] = pd.to_datetime(df_daily_report['請購日期'], errors='coerce').dt.date
                append_df[dept_col] = "The Peak"
                append_df[total_col] = pd.to_numeric(df_daily_report['總價'], errors='coerce').fillna(0)
                
                # 自動排除「官方總表已經有該月資料」的每日請購紀錄，防止 Double Counting
                append_df['ym'] = pd.to_datetime(append_df['日期']).dt.strftime('%Y-%m')
                append_df = append_df[~append_df['ym'].isin(official_ym)].drop(columns=['ym'])

                if not append_df.empty:
                    # 嘗試對應品名欄位以顯示在明細中
                    item_desc_col = next((c for c in df_purchase.columns if '品名' in c or '項次說明' in c or '明細' in c or '項目' in c), None)
                    if '品項名稱' in df_daily_report.columns:
                        item_str = df_daily_report['品項名稱'].astype(str)
                        if '請購數量' in df_daily_report.columns and '單位' in df_daily_report.columns:
                            item_str += " (" + df_daily_report['請購數量'].astype(str) + " " + df_daily_report['單位'].astype(str) + ")"
                        
                        if item_desc_col:
                            append_df[item_desc_col] = item_str
                        else:
                            append_df['備註(系統生成)'] = item_str
                    
                    df_purchase = pd.concat([df_purchase, append_df], ignore_index=True)
                    df_purchase = df_purchase[df_purchase['日期'].notna()]

            # --- 新增：將 4FHH_daily_purchase_report 無縫匯入 df_purchase 以供全域呈現 ---
            df_hh_report = fetch_4fhh_daily_purchase_report()
            if not df_hh_report.empty and '總價' in df_hh_report.columns:
                append_hh_df = pd.DataFrame()
                if '請購日期' in df_hh_report.columns:
                    append_hh_df['日期'] = pd.to_datetime(df_hh_report['請購日期'], errors='coerce').dt.date
                append_hh_df[dept_col] = "Happy Hour"
                append_hh_df[total_col] = pd.to_numeric(df_hh_report['總價'], errors='coerce').fillna(0)
                
                # 自動排除「官方總表已經有該月資料」的每日請購紀錄，防止 Double Counting
                append_hh_df['ym'] = pd.to_datetime(append_hh_df['日期']).dt.strftime('%Y-%m')
                append_hh_df = append_hh_df[~append_hh_df['ym'].isin(official_ym)].drop(columns=['ym'])

                if not append_hh_df.empty:
                    # 嘗試對應品名欄位以顯示在明細中
                    item_desc_col = next((c for c in df_purchase.columns if '品名' in c or '項次說明' in c or '明細' in c or '項目' in c), None)
                    if '品項名稱' in df_hh_report.columns:
                        item_str = df_hh_report['品項名稱'].astype(str)
                        if '請購數量' in df_hh_report.columns and '單位' in df_hh_report.columns:
                            item_str += " (" + df_hh_report['請購數量'].astype(str) + " " + df_hh_report['單位'].astype(str) + ")"
                        
                        if item_desc_col:
                            append_hh_df[item_desc_col] = item_str
                        else:
                            append_hh_df['備註(系統生成)'] = item_str
                    
                    df_purchase = pd.concat([df_purchase, append_hh_df], ignore_index=True)
                    df_purchase = df_purchase[df_purchase['日期'].notna()]


            # 過濾當月數據
            m_start = selected_date.replace(day=1)
            import calendar
            _, last_day = calendar.monthrange(
                selected_date.year, selected_date.month)
            m_end = selected_date.replace(day=last_day)

            df_month = df_purchase[(df_purchase['日期'] >= m_start) & (
                df_purchase['日期'] <= m_end)].copy()

            # 診斷：日期解析後的狀況
            with st.expander("🔍 日期解析診斷（診斷中，確認後可收起）", expanded=True):
                parsed_sample = df_purchase['日期'].dropna().head(5).tolist()
                st.write(f"解析後前 5 筆日期：{parsed_sample}")
                st.write(f"篩選區間：{m_start} ～ {m_end}")
                st.write(f"purchase_data 解析後剩餘總筆數：{len(df_purchase)} 筆（NaT 已移除）")
                st.write(f"符合 {selected_date.strftime('%Y-%m')} 的筆數：**{len(df_month)} 筆**")
                if df_month.empty:
                    st.error("⚠️ 當月筆數為 0！日期可能為其他年份或格式仍無法解析。")
                    all_dates = df_purchase['日期'].dropna().sort_values().unique()
                    if len(all_dates) > 0:
                        st.write(f"實際資料最早日期：{all_dates[0]}，最晚日期：{all_dates[-1]}")

            # --- 新增：取得上個月數據用於 MoM 分析 ---
            prev_m_date = get_month_delta(selected_date, -1)
            pm_start = prev_m_date.replace(day=1)
            _, pm_last_day = calendar.monthrange(
                prev_m_date.year, prev_m_date.month)
            pm_end = prev_m_date.replace(day=pm_last_day)
            df_prev_month = df_purchase[(df_purchase['日期'] >= pm_start) & (
                df_purchase['日期'] <= pm_end)].copy()
            # --- 新增：💰 本月剩餘預算戰情室 (The Peak) ---
            st.subheader("🎯 The Peak 本月剩餘預算戰情室 (Dynamic Budget)")
            
            # --- 新增：Projected CPG 佔位符 ---
            projected_cpg_container = st.empty()
            st.markdown("<br>", unsafe_allow_html=True)
            
            # 1. 計算 The Peak 本月已花費
            peak_spent = 0
            hh_spent = 0
            if not df_month.empty:
                _df_m_tmp = df_month.copy()
                _df_m_tmp['小計'] = pd.to_numeric(_df_m_tmp[total_col], errors='coerce').fillna(0)
                curr_depts_tmp = _df_m_tmp.groupby(dept_col)['小計'].sum().reset_index()
                all_d_list = curr_depts_tmp[dept_col].astype(str).tolist()
                hh_m = [d for d in all_d_list if '4' in d or any(k in d.upper() for k in ['HH', 'HAPPY', '歡樂時光'])]
                peak_m = [d for d in all_d_list if any(k in d.upper() for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲']) and d not in hh_m]
                peak_spent = curr_depts_tmp[curr_depts_tmp[dept_col].isin(peak_m)]['小計'].sum()
                hh_spent = curr_depts_tmp[curr_depts_tmp[dept_col].isin(hh_m)]['小計'].sum()
            
            # 2. 計算來客數: 歷史 (1號到昨天) + 未來 (今天到月底)
            import datetime
            import calendar
            today_d = datetime.date.today()
            hist_guests = 0
            future_guests = 0
            hist_bf = 0
            hist_af = 0
            hist_hh = 0
            
            # 歷史客數 (拆分早餐/下午茶) - 使用兩館加總資料
            df_fb_hist = get_combined_fb_daily_df(selected_date.year, selected_date.month, current_hotel)
            if not df_fb_hist.empty:
                df_fb_hist['date_obj'] = pd.to_datetime(df_fb_hist['date']).dt.date
                if (selected_date.year, selected_date.month) < (today_d.year, today_d.month):
                    hist_guests = df_fb_hist['peak_guests'].sum()
                    hist_bf = df_fb_hist['bf_act'].sum()
                    hist_af = df_fb_hist['af_act'].sum()
                    if 'hh_act' in df_fb_hist.columns:
                        hist_hh = df_fb_hist['hh_act'].sum()
                elif (selected_date.year, selected_date.month) == (today_d.year, today_d.month):
                    past_df = df_fb_hist[df_fb_hist['date_obj'] < today_d]
                    hist_guests = past_df['peak_guests'].sum()
                    hist_bf = past_df['bf_act'].sum()
                    hist_af = past_df['af_act'].sum()
                    if 'hh_act' in past_df.columns:
                        hist_hh = past_df['hh_act'].sum()
            
            # 本月歷史早/午比例推算
            _total_hist_peak = hist_bf + hist_af
            if _total_hist_peak > 0:
                _bf_ratio = hist_bf / _total_hist_peak
                _af_ratio = hist_af / _total_hist_peak
            else:
                _bf_ratio = 0.93  # 預設
                _af_ratio = 0.07

            # 到客率計算：統計目前選擇日期的過去 30 天 (動態浮動)
            ref_start_d = selected_date - datetime.timedelta(days=30)
            ref_end_d = selected_date - datetime.timedelta(days=1)
            
            start_30d_str = ref_start_d.strftime('%Y-%m-%d')
            end_1d_str = ref_end_d.strftime('%Y-%m-%d')
            
            fb_ref = compute_fb_mtd(start_30d_str, end_1d_str)

            _bf_ref_est = fb_ref.get('bf_theme_est', 0) + fb_ref.get('bf_zq_est', 0)
            _bf_ref_act = fb_ref.get('total_act_bf', 0)
            _default_bf_rate = (_bf_ref_act / _bf_ref_est) if _bf_ref_est > 0 else 1.0

            _af_ref_est = fb_ref.get('af_theme_est', 0) + fb_ref.get('af_zq_est', 0)
            _af_ref_act = fb_ref.get('total_act_af', 0)
            _default_af_rate = (_af_ref_act / _af_ref_est) if _af_ref_est > 0 else 1.0

            
            # 未來客數原始數據
            raw_future_guests = 0
            raw_future_bf = 0
            raw_future_af = 0
            if (selected_date.year, selected_date.month) >= (today_d.year, today_d.month):
                df_fb_fut = get_combined_fb_future_data()
                if not df_fb_fut.empty and 'date' in df_fb_fut.columns:
                    df_fb_fut['date_obj'] = pd.to_datetime(df_fb_fut['date']).dt.date
                    _, last_d = calendar.monthrange(selected_date.year, selected_date.month)
                    m_start_dt = datetime.date(selected_date.year, selected_date.month, 1)
                    m_end_dt = datetime.date(selected_date.year, selected_date.month, last_d)
                    calc_start = max(m_start_dt, today_d)
                    fut_mask = (df_fb_fut['date_obj'] >= calc_start) & (df_fb_fut['date_obj'] <= m_end_dt)
                    if '數量' in df_fb_fut.columns:
                        _fut_df = df_fb_fut[fut_mask]
                        raw_future_guests = _fut_df['數量'].sum()
                        
                        _meal_col = next((c for c in _fut_df.columns if any(k in c for k in ['餐別', '類型', '服務', 'meal', 'type'])), None)
                        if _meal_col:
                            _bf_kw = ['早', 'breakfast', 'bf']
                            _af_kw = ['下午', 'afternoon', 'tea', 'af']
                            _bf_mask = _fut_df[_meal_col].astype(str).str.lower().str.contains('|'.join(_bf_kw), na=False)
                            _af_mask = _fut_df[_meal_col].astype(str).str.lower().str.contains('|'.join(_af_kw), na=False)
                            raw_future_bf = _fut_df.loc[_bf_mask, '數量'].sum()
                            raw_future_af = _fut_df.loc[_af_mask, '數量'].sum()
                        else:
                            raw_future_bf = int(raw_future_guests * _bf_ratio)
                            raw_future_af = int(raw_future_guests * _af_ratio)

            # 3. UI 呈現：到客率與 CPG
            col_w1, col_w2, col_w3 = st.columns([1, 1, 1.5])
            with col_w1:
                target_cpg = st.number_input("🎯 設定目標單客成本 (Target CPG)", value=150, step=5, min_value=1)
            
            with col_w2:
                _rate_source_label = f"基準：{start_30d_str} 至 {end_1d_str} 實際到客率"
                bf_conv_rate = st.slider(
                    "⚙️ 早餐到客率折算 (%)",
                    min_value=0, max_value=120,
                    value=int(min(120, max(0, _default_bf_rate * 100))),
                    step=1,
                    help=f"{_rate_source_label}：早餐 {_default_bf_rate*100:.1f}%"
                )
                af_conv_rate = st.slider(
                    "⚙️ 下午茶到客率折算 (%)",
                    min_value=0, max_value=120,
                    value=int(min(120, max(0, _default_af_rate * 100))),
                    step=1,
                    help=f"{_rate_source_label}：下午茶 {_default_af_rate*100:.1f}%"
                )
                st.caption(f"📊 數值預設以「選擇日期({selected_date.strftime('%Y-%m-%d')})前30天」歷史表現為基準")
            
            # 套用折算率（直覺公式：未來總預約人數 × 到客率）
            adj_future_bf = int(raw_future_guests * (bf_conv_rate / 100.0))
            adj_future_af = int(raw_future_guests * (af_conv_rate / 100.0))
            future_guests = adj_future_bf + adj_future_af
            
            total_est_guests = hist_guests + future_guests
            
            # 本月估計總人次 (用於 K值系統)
            est_bf_total = hist_bf + adj_future_bf
            est_af_total = max(hist_af + adj_future_af, 1)

            total_budget = total_est_guests * target_cpg
            remaining_budget = total_budget - peak_spent
            remaining_cpg = remaining_budget / future_guests if future_guests > 0 else 0
            
            st.session_state['_budget_est_guests'] = total_est_guests
            st.session_state['_budget_total'] = total_budget
            
            # --- 修正：計算 Projected EOM CPG (改用 Current MTD CPG 或前月 CPG 推算) ---
            # 如果本月尚未有客數 (例如月初)，則改採「前月實際 CPG」作為未來消耗率預測基準，而非理想的目標值。
            fallback_cpg = target_cpg
            fallback_source = f"目標目標消耗率 (NT${int(target_cpg)})"
            try:
                import dateutil.relativedelta
                prev_d = selected_date - dateutil.relativedelta.relativedelta(months=1)
                df_fb_hist_prev = get_combined_fb_daily_df(prev_d.year, prev_d.month, current_hotel)
                prev_hist_guests = df_fb_hist_prev['peak_guests'].sum() if not df_fb_hist_prev.empty and 'peak_guests' in df_fb_hist_prev.columns else 0
                
                prev_peak_spent = 0
                if 'df_prev_month' in locals() and not df_prev_month.empty:
                    _df_pm = df_prev_month.copy()
                    _df_pm['小計'] = pd.to_numeric(_df_pm[total_col], errors='coerce').fillna(0)
                    _pm_depts = _df_pm.groupby(dept_col)['小計'].sum().reset_index()
                    _all_pm_d = _pm_depts[dept_col].astype(str).tolist()
                    _pm_hh = [d for d in _all_pm_d if '4' in d or any(k in d.upper() for k in ['HH', 'HAPPY', '歡樂時光'])]
                    _pm_peak = [d for d in _all_pm_d if any(k in d.upper() for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲']) and d not in _pm_hh]
                    prev_peak_spent = _pm_depts[_pm_depts[dept_col].isin(_pm_peak)]['小計'].sum()
                
                if prev_hist_guests > 0 and prev_peak_spent > 0:
                    fallback_cpg = prev_peak_spent / prev_hist_guests
                    fallback_source = f"前月實際消耗率 (NT${int(fallback_cpg)})"
            except Exception:
                pass

            if hist_guests > 0:
                current_mtd_cpg = peak_spent / hist_guests
                used_rate_source = f"目前消耗率 (NT${int(current_mtd_cpg)})"
            else:
                current_mtd_cpg = fallback_cpg
                used_rate_source = fallback_source
                
            expected_future_spend = future_guests * current_mtd_cpg
            projected_eom_cost = peak_spent + expected_future_spend
            projected_cpg = projected_eom_cost / total_est_guests if total_est_guests > 0 else 0
            cpg_delta = projected_cpg - target_cpg
            
            # 根據狀況給予顏色
            if projected_cpg > target_cpg * 1.05:
                status_color = "#e74c3c" # Red
                status_icon = "🚨 嚴重超支預警"
            elif projected_cpg > target_cpg:
                status_color = "#f39c12" # Orange
                status_icon = "⚠️ 輕微超支預警"
            else:
                status_color = "#2ecc71" # Green
                status_icon = "✅ 預算落點安全"
                
            # 渲染將延後至 k 值計算完成後
            with col_w3:
                st.markdown(f"**本月預估總客數**：`{int(total_est_guests):,}` 人")
                st.caption(
                    f"└ 歷史已發生 `{int(hist_guests):,}` 人 "
                    f"（早餐 `{int(hist_bf):,}` / 下午茶 `{int(hist_af):,}`）"
                )
                st.caption(
                    f"└ 未來預約 `{int(raw_future_guests):,}` 人 × 到客率 "
                    f"→ 早餐預估 `{int(adj_future_bf):,}` ({bf_conv_rate}%) "
                    f"/ 下午茶預估 `{int(adj_future_af):,}` ({af_conv_rate}%) "
                    f"= 合計 `{int(future_guests):,}` 人"
                )
                st.markdown(f"**本月總預算額度**：`NT$ {int(total_budget):,}`")
                st.markdown(f"**本月餐廳已花費**：`NT$ {int(peak_spent):,}`")
            
            st.markdown("#### 🚨 接下來採購戰略指標")
            c1, c2, c3 = st.columns(3)
            c1.metric("剩餘可用預算", f"NT$ {int(remaining_budget):,}")
            c2.metric("未來預約人數", f"{int(future_guests):,} 人")
            
            if remaining_cpg < 0:
                c3.error(f"❌ 已經超支！無法計算剩餘 CPG")
            elif remaining_cpg < target_cpg * 0.8:
                c3.warning(f"⚠️ 剩餘 CPG: ${remaining_cpg:.1f} (預算吃緊，請調整高價食材比例)")
            else:
                c3.success(f"✅ 剩餘 CPG: ${remaining_cpg:.1f} (預算充裕，可正常採購)")

            # ─────────────────────────────────────────────────
            # 🔢 早餐 / 下午茶 CPG 拆分分析 (k 值系統)
            # ─────────────────────────────────────────────────
            st.divider()
            st.markdown("#### 🍳 早餐 vs 🍰 下午茶 單客成本拆分分析")

            # 說明早/午比例來源
            _has_hist = hist_bf + hist_af > 0
            _ratio_src = f"本月已發生數據 (早餐 {int(hist_bf):,} 人 / 下午茶 {int(hist_af):,} 人)" if _has_hist else "預設估算 (早餐 93% / 下午茶 7%)"
            st.caption(f"📊 本月客數比例來源：{_ratio_src} → 估計早餐 `{est_bf_total:,}` 人 / 下午茶 `{est_af_total:,}` 人")

            # 計算 k 值動態上限
            CB_MIN = 80   # 早餐每人食材成本的絕對底線
            B = est_bf_total
            A = est_af_total
            T = target_cpg
            if A > 0 and CB_MIN > 0:
                k_max_raw = ((B + A) * T - CB_MIN * B) / (CB_MIN * A)
                k_max = max(1.0, round(k_max_raw, 1))
            else:
                k_max = 5.0

            k_col1, k_col2 = st.columns([1, 2])
            with k_col1:
                k_val = st.slider(
                    "⚖️ 下午茶相對成本係數 k",
                    min_value=1.0,
                    max_value=float(k_max),
                    value=min(1.8, float(k_max)),
                    step=0.1,
                    help=f"k=1 代表早/午食材成本相同；k 越高代表下午茶比早餐每人貴越多倍。\n系統自動計算上限：當早餐 CPG 低於 NT${CB_MIN} 時視為不可接受，此時 k_max = {k_max:.1f}"
                )

            # 核心公式 (基於目標)
            if B + k_val * A > 0:
                cb = (B + A) * T / (B + k_val * A)
                ca = k_val * cb
            else:
                cb = T
                ca = T

            _check = (B * cb + A * ca) / (B + A) if (B + A) > 0 else 0
            
            # --- 新增：核心公式 (基於剩餘救援 Remaining CPG) ---
            R = remaining_cpg
            if B + k_val * A > 0:
                rem_cb = (B + A) * R / (B + k_val * A)
                rem_ca = k_val * rem_cb
            else:
                rem_cb = R
                rem_ca = R
            
            # 延後渲染的 Projected CPG 大面板
            if B + k_val * A > 0:
                proj_cb = (B + A) * projected_cpg / (B + k_val * A)
                proj_ca = k_val * proj_cb
            else:
                proj_cb = projected_cpg
                proj_ca = projected_cpg

            hh_projected_cpg = hh_spent / hist_hh if hist_hh > 0 else 0

            projected_cpg_container.markdown(f"""
            <div style="background: linear-gradient(135deg, #1f2c56 0%, #2e437c 100%); padding: 25px; border-radius: 15px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
                <div style="flex: 1;">
                    <h3 style="margin:0; color:#ecf0f1; font-size:1.3rem;">🔮 本月預估落點單客成本 (Projected EOM CPG)</h3>
                    <div style="font-size:0.9rem; color:#bdc3c7; margin-top:8px;">基於目前累積花費 <b>NT${int(peak_spent):,}</b>，並假設未來維持{used_rate_source} 推算</div>
                    <div style="display:flex; gap: 15px; margin-top: 15px;">
                        <div style="background: rgba(46, 204, 113, 0.15); border: 1px solid rgba(46, 204, 113, 0.3); padding: 8px 15px; border-radius: 8px;">
                            <span style="font-size:0.8rem; color:#ecf0f1;">🍳 預估早餐落點</span><br>
                            <span style="font-size:1.2rem; font-weight:bold; color:#2ecc71;">NT$ {proj_cb:.1f}</span>
                        </div>
                        <div style="background: rgba(243, 156, 18, 0.15); border: 1px solid rgba(243, 156, 18, 0.3); padding: 8px 15px; border-radius: 8px;">
                            <span style="font-size:0.8rem; color:#ecf0f1;">🍰 預估下午茶落點</span><br>
                            <span style="font-size:1.2rem; font-weight:bold; color:#f39c12;">NT$ {proj_ca:.1f}</span>
                        </div>
                    </div>
                </div>
                <div style="text-align: right; padding-left: 20px;">
                    <h1 style="margin:0; color:{status_color}; font-size:3.2rem; font-weight:900; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">NT$ {projected_cpg:.1f}</h1>
                    <div style="color:{status_color}; font-size:1.1rem; font-weight:bold; margin-top:5px;">{status_icon} (與目標差距: {cpg_delta:+.1f})</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # --- 獨立顯示 HH 預估落點 ---
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #2c1a40 0%, #4a2b66 100%); padding: 15px 25px; border-radius: 15px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 15px rgba(0,0,0,0.2); margin-top: 15px;">
                <div style="flex: 1;">
                    <h3 style="margin:0; color:#ecf0f1; font-size:1.1rem;">🥂 Happy Hour 預估落點單客成本</h3>
                    <div style="font-size:0.85rem; color:#bdc3c7; margin-top:5px;">基於目前累積 HH 花費 <b>NT${int(hh_spent):,}</b> 與 MTD 累積客數推算，與 The Peak 完全獨立計算。</div>
                </div>
                <div style="text-align: right; padding-left: 20px;">
                    <h1 style="margin:0; color:#d2b4de; font-size:2.4rem; font-weight:900; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">NT$ {hh_projected_cpg:.1f}</h1>
                </div>
            </div>
            """, unsafe_allow_html=True)

            with k_col2:
                cb_color = "#27ae60" if cb >= 100 else ("#f39c12" if cb >= CB_MIN else "#e74c3c")
                ca_color = "#27ae60" if ca <= 250 else ("#f39c12" if ca <= 350 else "#e74c3c")
                st.markdown(f"""
                <div style="display:flex; gap:12px; margin-top:8px;">
                    <div style="flex:1; background:#1e2d40; border-left:5px solid {cb_color}; border-radius:10px; padding:16px; text-align:center; position:relative;">
                        <div style="color:#aaa; font-size:0.8rem;">🍳 目標理想分配 (理想上限)</div>
                        <div style="color:{cb_color}; font-size:1.8rem; font-weight:800; margin:6px 0;">NT$ {cb:.0f}</div>
                        <div style="color:#aaa; font-size:0.75rem;">總均攤預估 {est_bf_total:,} 人次</div>
                    </div>
                    <div style="flex:1; background:#1e2d40; border-left:5px solid {ca_color}; border-radius:10px; padding:16px; text-align:center;">
                        <div style="color:#aaa; font-size:0.8rem;">🍰 目標理想分配 (理想上限)</div>
                        <div style="color:{ca_color}; font-size:1.8rem; font-weight:800; margin:6px 0;">NT$ {ca:.0f}</div>
                        <div style="color:#aaa; font-size:0.75rem;">總均攤預估 {est_af_total:,} 人次</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # --- 新增：真實剩餘救援指標 ---
                if rem_cb < 0:
                    rescue_msg = f"<div style='color:#e74c3c; margin-top:10px; font-weight:bold;'>🚨 預算已經透支！無法提供未來剩餘上限。</div>"
                else:
                    rescue_msg = f"""
                    <div style="background: rgba(231, 76, 60, 0.1); border: 1px solid #e74c3c; border-radius: 8px; padding: 12px; margin-top: 10px;">
                        <div style="color:#e74c3c; font-size:0.85rem; font-weight:bold; margin-bottom:5px;">⚠️ 真實採購防線 (扣除已花費後的嚴格上限)</div>
                        <div style="display:flex; justify-content:space-between; color:#ecf0f1; font-size:0.9rem;">
                            <span>🍳 早餐剩餘嚴格上限：<b>NT$ {rem_cb:.0f}</b> /人</span>
                            <span>🍰 下午茶剩餘嚴格上限：<b>NT$ {rem_ca:.0f}</b> /人</span>
                        </div>
                    </div>
                    """
                st.markdown(rescue_msg, unsafe_allow_html=True)

            # k 值動態說明文字
            st.markdown("")
            if k_val >= k_max * 0.95:
                st.error(f"🚫 **k 值已達上限 ({k_val:.1f} / 上限 {k_max:.1f})**\n\n早餐每人成本已壓至 NT${cb:.0f}（接近最低底線 ${CB_MIN}）。繼續調高 k 將導致早餐食材嚴重縮水，影響餐點品質與住客滿意度。建議立即調低 k 值。")
            elif k_val > 2.5:
                st.warning(f"⚠️ **k 值偏高 ({k_val:.1f})**\n\n目前下午茶每人食材預算達 NT${ca:.0f}，早餐僅剩 NT${cb:.0f}。\n- **再往高調的風險**：早餐供應品質下滑，住客投訴風險提升。\n- **往低調的效果**：早餐食材預算提升至 NT${((B+A)*T/(B+(k_val-0.3)*A)):.0f}，下午茶縮至 NT${((k_val-0.3)*(B+A)*T/(B+(k_val-0.3)*A)):.0f}。\n→ 建議範圍：k = 1.5 ～ 2.2")
            elif k_val < 1.2:
                st.info(f"ℹ️ **k 值偏低 ({k_val:.1f})**\n\n早餐與下午茶每人食材成本幾乎相同（NT${cb:.0f} vs NT${ca:.0f}）。\n- **再往低調的風險**：下午茶預算不足，甜點備料品質難以維持。\n- **往高調的效果**：下午茶預算提升至 NT${((k_val+0.3)*(B+A)*T/(B+(k_val+0.3)*A)):.0f}，更真實反映備料成本。\n→ 建議至少設 k = 1.5")
            else:
                st.success(f"✅ **k 值合理 ({k_val:.1f})** ─ 早餐 NT${cb:.0f} / 下午茶 NT${ca:.0f}\n\n目前分配符合業界常規，加權總 CPG = NT${_check:.0f}（等於目標 ${int(T)}）。\n- 若往上調 k 至 {k_val+0.2:.1f} → 下午茶提升至 NT${((k_val+0.2)*(B+A)*T/(B+(k_val+0.2)*A)):.0f}，早餐降至 NT${((B+A)*T/(B+(k_val+0.2)*A)):.0f}\n- 若往下調 k 至 {k_val-0.2:.1f} → 早餐提升至 NT${((B+A)*T/(B+(k_val-0.2)*A)):.0f}，下午茶降至 NT${((k_val-0.2)*(B+A)*T/(B+(k_val-0.2)*A)):.0f}")

            # 月份早/午預算總額概覽
            bf_budget_total = cb * est_bf_total
            af_budget_total = ca * est_af_total
            st.markdown(f"""
            <div style="background:linear-gradient(135deg,#1a1a2e,#16213e); border-radius:12px; padding:16px; margin-top:12px; display:flex; justify-content:space-around; text-align:center;">
                <div><div style="color:#aaa;font-size:0.8rem;">🍳 本月早餐食材總預算</div><div style="color:#2ecc71;font-size:1.4rem;font-weight:700;">NT$ {int(bf_budget_total):,}</div></div>
                <div style="color:#555;font-size:2rem; align-self:center;">+</div>
                <div><div style="color:#aaa;font-size:0.8rem;">🍰 本月下午茶食材總預算</div><div style="color:#e67e22;font-size:1.4rem;font-weight:700;">NT$ {int(af_budget_total):,}</div></div>
                <div style="color:#555;font-size:2rem; align-self:center;">=</div>
                <div><div style="color:#aaa;font-size:0.8rem;">🎯 本月總預算上限</div><div style="color:#3498db;font-size:1.4rem;font-weight:700;">NT$ {int(total_budget):,}</div></div>
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 🎯 預測準確度回顧與快照系統
            # ==========================================
            st.markdown("<br>", unsafe_allow_html=True)
            st.divider()
            st.subheader("📊 預測準確度回顧與快照 (Prediction Accuracy)")
            st.info("💡 此區塊讓您每週手動儲存「未來 7 天」的客數預測快照，並能在下週檢視預測準確度。")
            
            # --- 1. 儲存快照區塊 ---
            snap_col1, snap_col2 = st.columns([1, 1])
            
            import datetime
            today_date = datetime.date.today()
            target_start = today_date + datetime.timedelta(days=1)
            target_end = today_date + datetime.timedelta(days=7)
            
            # 計算未來 7 天的預估人數 (為了快照)
            snap_future_bf = 0
            snap_future_af = 0
            if 'df_fb_fut' in locals() and not df_fb_fut.empty and 'date_obj' in df_fb_fut.columns and '數量' in df_fb_fut.columns:
                snap_mask = (df_fb_fut['date_obj'] >= target_start) & (df_fb_fut['date_obj'] <= target_end)
                snap_raw = df_fb_fut[snap_mask]['數量'].sum()
                snap_future_bf = int(snap_raw * (bf_conv_rate / 100.0))
                snap_future_af = int(snap_raw * (af_conv_rate / 100.0))
            snap_future_total = snap_future_bf + snap_future_af
            
            with snap_col1:
                st.markdown(f"**準備存檔的預測區間：** `{target_start}` 至 `{target_end}` (未來 7 天)")
                st.markdown(f"**預測人數：** 🍳 早 `{snap_future_bf}` 人 ｜ 🍰 午 `{snap_future_af}` 人 ｜ 總計 `{snap_future_total}` 人")
                st.markdown(f"**套用折算率：** 早餐 `{bf_conv_rate}%` ｜ 下午茶 `{af_conv_rate}%`")
                
                if st.button("💾 儲存本週預測快照", use_container_width=True):
                    new_snap = {
                        'snapshot_date': today_date.strftime('%Y-%m-%d'),
                        'target_date_start': target_start.strftime('%Y-%m-%d'),
                        'target_date_end': target_end.strftime('%Y-%m-%d'),
                        'bf_conv_rate': bf_conv_rate,
                        'af_conv_rate': af_conv_rate,
                        'future_bf': snap_future_bf,
                        'future_af': snap_future_af,
                        'future_total': snap_future_total
                    }
                    with st.spinner("正在將預測資料寫入 Google Sheets..."):
                        if save_prediction_snapshot(new_snap):
                            st.success(f"✅ 快照儲存成功！下週即可回頭檢視此區間的準確率。")
                            st.rerun()

            # --- 2. 回顧過往準確度 ---
            with snap_col2:
                df_snapshots = get_prediction_snapshots()
                if df_snapshots.empty:
                    st.warning("目前尚無任何預測快照記錄。請先儲存快照，一週後即可在此檢視準確率。")
                else:
                    st.markdown("**🔍 歷史預測準確度覆盤**")
                    # 篩選出已經過期的快照 (target_date_end < today)
                    df_snapshots['target_date_end'] = pd.to_datetime(df_snapshots['target_date_end']).dt.date
                    df_past_snaps = df_snapshots[df_snapshots['target_date_end'] < today_date].copy()
                    
                    if df_past_snaps.empty:
                        st.info("快照區間尚未結束，等設定的「未來 7 天」過完後，即可進行對比。")
                        st.dataframe(df_snapshots[['snapshot_date', 'target_date_start', 'target_date_end', 'future_total']].tail(3), hide_index=True)
                    else:
                        # 取最新一筆已過期的快照來分析
                        latest_snap = df_past_snaps.sort_values('target_date_end', ascending=False).iloc[0]
                        p_start = latest_snap['target_date_start']
                        p_end = latest_snap['target_date_end']
                        p_total = int(latest_snap['future_total'])
                        
                        # 從歷史 FB 報表抓取實際到客數
                        actual_fb = compute_fb_mtd(str(p_start), str(p_end))
                        a_bf = actual_fb.get('total_act_bf', 0)
                        a_af = actual_fb.get('total_act_af', 0)
                        a_total = a_bf + a_af
                        
                        diff = a_total - p_total
                        accuracy = (min(a_total, p_total) / max(a_total, p_total) * 100) if max(a_total, p_total) > 0 else 100
                        
                        st.markdown(f"📅 **覆盤區間：** `{p_start}` 至 `{p_end}`")
                        st.markdown(f"🎯 **當時預測：** `{p_total}` 人次")
                        st.markdown(f"🏆 **最終實際：** `{a_total}` 人次")
                        
                        acc_color = "green" if accuracy >= 90 else ("orange" if accuracy >= 75 else "red")
                        diff_text = f"+{diff}" if diff > 0 else str(diff)
                        st.markdown(f"📊 **準確率：** <span style='color:{acc_color}; font-size:1.2rem; font-weight:bold;'>{accuracy:.1f}%</span> (誤差: {diff_text} 人)", unsafe_allow_html=True)
                        
                        if diff > 0:
                            st.caption("提示：實際人數大於預測，若誤差過大可能導致食材備料不足，建議微調到客率滑桿。")
                        elif diff < 0:
                            st.caption("提示：實際人數少於預測，若誤差過大可能導致食材報廢浪費，建議下調到客率。")


            st.divider()


            if not df_month.empty:
                # 數值清理
                df_month['小計'] = pd.to_numeric(
                    df_month[total_col], errors='coerce').fillna(0)
                if not df_prev_month.empty:
                    df_prev_month['小計'] = pd.to_numeric(
                        df_prev_month[total_col], errors='coerce').fillna(0)

                total_month_expense = df_month['小計'].sum()
                total_prev_expense = df_prev_month['小計'].sum(
                ) if not df_prev_month.empty else 0

                # 計算增長率
                mom_delta = total_month_expense - total_prev_expense
                mom_pcnt = (mom_delta / total_prev_expense *
                            100) if total_prev_expense > 0 else 0

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
                    st.metric("月增長金額 (MoM)", f"NT$ {int(mom_delta):,}", delta=int(
                        mom_delta), delta_color="inverse")
                with col_m3:
                    st.metric(
                        "月增長百分比", f"{mom_pcnt:.1f}%", delta=f"{mom_pcnt:.1f}%", delta_color="inverse")


                # --- 異常值監控：找出增長過快的部門 ---
                st.subheader("⚠️ 採購異常監控 (MoM Spikes)")
                # 計算各部門本月 vs 上月
                curr_depts = df_month.groupby(
                    dept_col)['小計'].sum().reset_index()
                curr_depts.columns = ['部門', '小計']

                if not df_prev_month.empty:
                    prev_depts = df_prev_month.groupby(
                        dept_col)['小計'].sum().reset_index()
                    prev_depts.columns = ['部門', '小計']
                else:
                    prev_depts = pd.DataFrame(columns=['部門', '小計'])

                comparison = pd.merge(
                    curr_depts, prev_depts, on='部門', how='left', suffixes=('_今', '_昨')).fillna(0)

                # 安全計算變動率 (避免 ZeroDivisionError 與 Indexing 類型報錯)
                def calc_mom_ratio(row):
                    if row['小計_昨'] > 0:
                        return (row['小計_今'] - row['小計_昨']) / row['小計_昨'] * 100
                    return 100.0 if row['小計_今'] > 0 else 0.0

                comparison['變動率'] = comparison.apply(calc_mom_ratio, axis=1)

                # 找出變動率大於 20% 且金額大於一定門檻的 (例如 > 2000)
                spikes = comparison[(comparison['變動率'] > 20) & (
                    comparison['小計_今'] > 2000)].sort_values('變動率', ascending=False)

                if not spikes.empty:
                    for _, row in spikes.iterrows():
                        st.warning(
                            f"🚩 **{row['部門']}** 本月開銷異常！較上月增長 **{row['變動率']:.1f}%** (NT$ {int(row['小計_今']):,})")
                else:
                    st.success("✅ 目前各部門採購金額平穩，未偵測到異常大幅波動。")

                st.divider()

                # 2. 部門佔比圓餅圖
                st.subheader("📊 各部門請購佔比分析")
                dept_summary = df_month.groupby(
                    dept_col)['小計'].sum().reset_index()
                dept_summary.columns = ['部門', '小計']

                # 繪製圓餅圖 (依照金額排序)
                base = alt.Chart(dept_summary).encode(
                    theta=alt.Theta(
                        field="小計", type="quantitative", stack=True),
                    color=alt.Color(
                        field="部門",
                        type="nominal",
                        scale=alt.Scale(scheme='category10'),
                        legend=alt.Legend(title="部門", orient="right"),
                        sort=alt.SortField("小計", order="descending")
                    ),
                    order=alt.Order("小計", sort="descending"),
                    tooltip=["部門", alt.Tooltip(
                        "小計", format=",.0f", title="總金額 (NT$)")]
                ).properties(height=450)

                # 圓餅主體
                chart_arc = base.mark_arc(
                    innerRadius=60, outerRadius=120, stroke="#fff")

                # 在圓餅切片上顯示金額
                chart_text = base.mark_text(radius=90, size=14, fontWeight="bold", color="white").encode(
                    text=alt.Text("小計:Q", format=",.0f")
                )

                st.altair_chart(chart_arc + chart_text,
                                use_container_width=True)

                # --- 新增：餐飲績效分析 (The Peak & Happy Hour) ---
                st.divider()
                st.subheader("🍽️ 餐飲績效與成本深度分析 (Cash-basis)")

                # 獲取當月每日數據 (含餐廳來客數)
                m_data = fetch_month_summary(
                    selected_date.year, selected_date.month)
                df_daily_rest = m_data['df']

                # 【關鍵修正】occ_data 沒有 F&B 欄位，直接從 f&b_report 注入每日來客數
                df_fb_daily = get_combined_fb_daily_df(
                    selected_date.year, selected_date.month, current_hotel)
                if not df_fb_daily.empty and not df_daily_rest.empty:
                    df_daily_rest = df_daily_rest.merge(df_fb_daily, on='date', how='left')
                    df_daily_rest['bf_act']      = df_daily_rest['bf_act'].fillna(0)
                    df_daily_rest['af_act']      = df_daily_rest['af_act'].fillna(0)
                    df_daily_rest['hh_act']      = df_daily_rest['hh_act'].fillna(0)
                    df_daily_rest['peak_guests'] = df_daily_rest['peak_guests'].fillna(0)
                elif not df_fb_daily.empty:
                    # df_daily_rest 是空的，用 fb_daily 代替
                    df_daily_rest = df_fb_daily.copy()

                if not df_daily_rest.empty:
                    # 確保必要欄位存在並轉為數值
                    target_cols = ['peak_guests', 'hh_act', 'revenue', 'bf_act', 'af_act']
                    for c in target_cols:
                        if c in df_daily_rest.columns:
                            df_daily_rest[c] = pd.to_numeric(df_daily_rest[c].astype(
                                str).str.replace(',', ''), errors='coerce').fillna(0)
                        else:
                            df_daily_rest[c] = 0

                    # 使用真實 F&B 來客數 (直接用 peak_guests 和 hh_act)
                    df_daily_rest['rest_day_guests'] = df_daily_rest['peak_guests']
                    df_daily_rest['rest_hh_guests']  = df_daily_rest['hh_act']
                    df_daily_rest['bf_total_act']    = df_daily_rest['bf_act']
                    df_daily_rest['af_total_act']    = df_daily_rest['af_act']


                    # --- 自動化邏輯：如果 The Peak 來客數為 0，則自動加總 早餐 + 下午茶 ---
                    def calculate_peak_guests(row):
                        if row['rest_day_guests'] > 0:
                            return row['rest_day_guests']
                        return row['bf_total_act'] + row['af_total_act']

                    df_daily_rest['effective_peak_guests'] = df_daily_rest.apply(
                        calculate_peak_guests, axis=1)

                    # 篩選 The Peak 與 Happy Hour 採購 (強力模糊匹配)
                    all_depts_list = dept_summary['部門'].astype(str).tolist()

                    # HH 匹配：包含 '4'、'HH' 或 'HAPPY'
                    hh_matched = [d for d in all_depts_list if '4' in d or any(
                        k in d.upper() for k in ['HH', 'HAPPY', '歡樂時光'])]
                    # Peak 匹配：包含 'PEAK' 或 '餐廳'，且排除 HH 部門
                    peak_matched = [d for d in all_depts_list if (any(k in d.upper(
                    ) for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲'])) and (d not in hh_matched)]

                    with st.expander("🛠️ 數據匹配校準器 (若數據不正確請點開)"):
                        st.info(f"📍 偵測到之所有部門: `{all_depts_list}`")
                        st.success(
                            f"🍷 歸類為 Happy Hour (HH) 之部門: `{hh_matched}`")
                        st.success(
                            f"🏰 歸類為 The Peak (餐廳) 之部門: `{peak_matched}`")

                        st.divider()
                        st.markdown("**🔍 本月採購原始明細 (DEBUG)**")
                        _debug_cols = [c for c in [date_col, dept_col, total_col] + [
                            c for c in df_month.columns if '品名' in c or 'Item' in c or '項目' in c
                        ] if c in df_month.columns]
                        st.caption(f"The Peak 原始列 (共 {len(df_month[df_month[dept_col].isin(peak_matched)])} 筆，合計 NT$ {int(df_month[df_month[dept_col].isin(peak_matched)][total_col].apply(pd.to_numeric, errors='coerce').sum()):,})")
                        st.dataframe(df_month[df_month[dept_col].isin(peak_matched)][_debug_cols].sort_values(date_col), use_container_width=True)
                        st.caption(f"Happy Hour 原始列 (共 {len(df_month[df_month[dept_col].isin(hh_matched)])} 筆，合計 NT$ {int(df_month[df_month[dept_col].isin(hh_matched)][total_col].apply(pd.to_numeric, errors='coerce').sum()):,})")
                        st.dataframe(df_month[df_month[dept_col].isin(hh_matched)][_debug_cols].sort_values(date_col), use_container_width=True)
                        st.caption(f"⚠️ 當月所有採購列共 {len(df_month)} 筆 | 整個 purchase data 表共 {len(df_purchase)} 筆")

                    df_peak_purchase = df_month[df_month[dept_col].isin(
                        peak_matched)].copy()
                    df_hh_purchase = df_month[df_month[dept_col].isin(
                        hh_matched)].copy()

                    # --- 進階匹配：若部門抓不到 HH，嘗試從品名抓取 ---
                    if df_hh_purchase.empty:
                        item_col = next((c for c in df_month.columns if any(
                            k in c for k in ['品名', '項目', 'Item'])), None)
                        if item_col:
                            df_hh_purchase = df_month[df_month[item_col].astype(
                                str).str.upper().str.contains('HH|HAPPY|歡樂時光', na=False)].copy()

                    # 計算每日採購總額（改用『以週為單位均攤』修正採購日 vs 消耗日失真）
                    df_daily_rest['日期_obj'] = pd.to_datetime(
                        df_daily_rest['date']).dt.date
                    df_daily_rest['日期_dt'] = pd.to_datetime(
                        df_daily_rest['date'])

                    def spread_monthly_cost(df_purchase_input, df_daily_base):
                        """將全月總費用平分到當月有來客的每一天 (保存規模經濟差異)"""
                        if df_purchase_input.empty or df_daily_base.empty:
                            return pd.Series(0, index=df_daily_base['日期_obj'])
                        
                        df_purchase_local = df_purchase_input.copy()
                        df_purchase_local[total_col] = pd.to_numeric(df_purchase_local[total_col], errors='coerce').fillna(0)
                        total_month_cost = df_purchase_local[total_col].sum()
                        
                        df_base = df_daily_base.copy()
                        
                        has_guests = df_base['effective_peak_guests'] > 0
                        days_with_guests = has_guests.sum()
                        
                        df_base['spread_cost'] = 0.0
                        if days_with_guests > 0:
                            df_base.loc[has_guests, 'spread_cost'] = total_month_cost / days_with_guests
                        else:
                            # 沒客人則平分給所有天數
                            df_base['spread_cost'] = total_month_cost / len(df_base)
                            
                        return df_base.set_index('日期_obj')['spread_cost']

                    # 用月均攤計算每日成本
                    peak_spread = spread_monthly_cost(
                        df_peak_purchase, df_daily_rest)
                    hh_spread = spread_monthly_cost(
                        df_hh_purchase, df_daily_rest)

                    # 合併來客數與週均攤成本
                    analysis_df = df_daily_rest[[
                        '日期_obj', 'effective_peak_guests', 'bf_act', 'af_act', 'rest_hh_guests', 'revenue']].copy()
                    analysis_df['peak_cost'] = analysis_df['日期_obj'].map(
                        peak_spread).fillna(0)
                    analysis_df['hh_cost'] = analysis_df['日期_obj'].map(
                        hh_spread).fillna(0)

                    # --- 累計分析邏輯：計算本月至今的累積數據 ---
                    analysis_df = analysis_df.sort_values('日期_obj')
                    analysis_df['cum_peak_cost'] = analysis_df['peak_cost'].cumsum()
                    analysis_df['cum_peak_guests'] = analysis_df['effective_peak_guests'].cumsum(
                    )
                    analysis_df['cum_hh_cost'] = analysis_df['hh_cost'].cumsum()
                    analysis_df['cum_hh_guests'] = analysis_df['rest_hh_guests'].cumsum(
                    )

                    # 計算累積 CPG (這才是真實的平均成本走勢)
                    analysis_df['cum_peak_cpg'] = analysis_df.apply(
                        lambda r: r['cum_peak_cost']/r['cum_peak_guests'] if r['cum_peak_guests'] > 0 else 0, axis=1)
                    analysis_df['cum_hh_cpg'] = analysis_df.apply(
                        lambda r: r['cum_hh_cost']/r['cum_hh_guests'] if r['cum_hh_guests'] > 0 else 0, axis=1)

                    # UI 呈現 (本月總結)
                    total_peak_cost = analysis_df['cum_peak_cost'].iloc[-1] if not analysis_df.empty else 0
                    total_peak_guests = analysis_df['cum_peak_guests'].iloc[-1] if not analysis_df.empty else 0
                    final_peak_cpg = total_peak_cost / \
                        total_peak_guests if total_peak_guests > 0 else 0

                    total_hh_cost = analysis_df['cum_hh_cost'].iloc[-1] if not analysis_df.empty else 0
                    total_hh_guests = analysis_df['cum_hh_guests'].iloc[-1] if not analysis_df.empty else 0
                    final_hh_cpg = total_hh_cost / total_hh_guests if total_hh_guests > 0 else 0

                    # --- 📈 The Peak CPG 防禦力績效圖 (CPG vs 菜商指數) ---
                    st.markdown("##### 📈 The Peak CPG 防禦力績效圖 (CPG vs 菜商指數)")
                    st.caption(
                        "💡 **防禦力判定**：若紅虛線(菜價)上升，但藍實線(CPG)持平或下降，代表採購防禦成功！")

                    trend_rows = []
                    for n_back in range(5, -1, -1):  # 從 5 個月前到本月
                        t_date = get_month_delta(selected_date, -n_back)
                        t_label = t_date.strftime('%Y-%m')

                        # 抓該月採購數據
                        t_start = t_date.replace(day=1)
                        import calendar as _cal
                        _, t_last = _cal.monthrange(t_date.year, t_date.month)
                        t_end = t_date.replace(day=t_last)
                        df_t_purchase = df_purchase[(df_purchase['日期'] >= t_start) & (
                            df_purchase['日期'] <= t_end)].copy()

                        if not df_t_purchase.empty:
                            df_t_purchase['小計'] = pd.to_numeric(
                                df_t_purchase[total_col], errors='coerce').fillna(0)

                        # 抓該月來客數
                        t_m_data = fetch_month_summary(
                            t_date.year, t_date.month)
                        t_df = t_m_data.get('df', pd.DataFrame())
                        t_guests = 0
                        if not t_df.empty:
                            for _c in ['rest_day_guests', 'bf_total_act', 'af_total_act']:
                                if _c in t_df.columns:
                                    t_df[_c] = pd.to_numeric(t_df[_c].astype(
                                        str).str.replace(',', ''), errors='coerce').fillna(0)
                            if 'rest_day_guests' in t_df.columns and t_df['rest_day_guests'].sum() > 0:
                                t_guests = t_df['rest_day_guests'].sum()
                            elif 'bf_total_act' in t_df.columns:
                                t_guests = (
                                    t_df['bf_total_act'] + t_df.get('af_total_act', 0)).sum()

                        # 篩選 The Peak 採購
                        t_peak_cost = 0
                        if not df_t_purchase.empty and dept_col in df_t_purchase.columns:
                            t_all_depts = df_t_purchase[dept_col].astype(
                                str).unique().tolist()
                            t_hh = [d for d in t_all_depts if '4' in d or any(
                                k in d.upper() for k in ['HH', 'HAPPY', '歡樂時光'])]
                            t_peak_depts = [d for d in t_all_depts if any(
                                k in d.upper() for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲']) and d not in t_hh]
                            t_peak_cost = df_t_purchase[df_t_purchase[dept_col].isin(
                                t_peak_depts)]['小計'].sum()

                        t_cpg = t_peak_cost / t_guests if t_guests > 0 else None
                        trend_rows.append(
                            {'月份': t_label, 'CPG': t_cpg, '目標': 150})

                    trend_df = pd.DataFrame(trend_rows).dropna(subset=['CPG'])

                    if not trend_df.empty and len(trend_df) >= 2:
                        # 取得大盤指數並與 trend_df 合併
                        sp_df = fetch_supplier_prices()
                        idx_df = get_market_index_df(sp_df)

                        if not idx_df.empty:
                            # 若同個月份有多期指數，取平均
                            idx_monthly = idx_df.groupby('month_label')[
                                'index'].mean().reset_index()
                            trend_df = trend_df.merge(
                                idx_monthly, left_on='月份', right_on='month_label', how='left')
                        else:
                            trend_df['index'] = None

                        base = alt.Chart(trend_df)

                        # 左軸：CPG 藍實線
                        cpg_line = base.mark_line(point=True, strokeWidth=3, color='#1f2c56').encode(
                            x=alt.X('月份:N', title='月份', sort=None),
                            y=alt.Y('CPG:Q', title='每客成本 CPG (NT$)', scale=alt.Scale(
                                zero=False), axis=alt.Axis(titleColor='#1f2c56')),
                            tooltip=[
                                alt.Tooltip('月份:N', title='月份'),
                                alt.Tooltip(
                                    'CPG:Q', title='CPG (NT$)', format=',.0f')
                            ]
                        )

                        target_line = alt.Chart(pd.DataFrame({'y': [150]})).mark_rule(
                            color='#1f2c56', strokeDash=[6, 3], strokeWidth=1.5, opacity=0.5
                        ).encode(y='y:Q')
                        target_label = alt.Chart(pd.DataFrame({'y': [150], 'x': [trend_df['月份'].iloc[-1]], 'text': ['目標 $150']})).mark_text(
                            align='right', dx=-4, dy=-8, color='#1f2c56', fontSize=11, fontWeight='bold', opacity=0.8
                        ).encode(x='x:N', y='y:Q', text='text:N')

                        cpg_layer = alt.layer(
                            cpg_line, target_line, target_label)

                        # 右軸：大盤指數 紅虛線
                        has_index_data = False
                        if 'index' in trend_df.columns and not trend_df['index'].isna().all():
                            valid_idx_count = trend_df['index'].notna().sum()
                            has_index_data = True

                            idx_line = base.mark_line(point={'color': '#e74c3c', 'size': 60}, strokeDash=[5, 5], strokeWidth=2, color='#e74c3c').encode(
                                x=alt.X('月份:N', sort=None),
                                y=alt.Y('index:Q', title='菜商大盤指數 (100=基準)', scale=alt.Scale(
                                    zero=False), axis=alt.Axis(titleColor='#e74c3c')),
                                tooltip=[
                                    alt.Tooltip('月份:N', title='月份'),
                                    alt.Tooltip(
                                        'index:Q', title='菜商指數', format=',.1f')
                                ]
                            )
                            chart = alt.layer(cpg_layer, idx_line).resolve_scale(
                                y='independent')

                            if valid_idx_count < 2:
                                st.info(
                                    "💡 提醒：大盤指數（紅虛線）的對應月份不足 2 個月，因此目前圖表上僅會顯示一個紅色點點。請在「菜單分析」分頁補齊過去月份的菜價資料以顯示完整折線。")
                        else:
                            chart = cpg_layer

                        st.altair_chart(chart.properties(
                            height=280), use_container_width=True)
                    else:
                        st.info("💡 需要至少 2 個月的數據才能顯示 CPG 趨勢圖。")

                    st.divider()

                    # --- 📊 採購花費 vs 早餐來客數 相關性驗證（以週為單位）---
                    st.markdown("##### 📊 採購花費 vs 早餐來客數 相關性驗證（週）")
                    st.caption(
                        "💡 兩條線的形狀應趨近一致。若某週「採購↑ 來客↓」或「採購↓ 來客↑」，代表食材控管可能有問題。")

                    corr_df = analysis_df[[
                        '日期_obj', 'peak_cost', 'effective_peak_guests']].copy()
                    corr_df['日期_dt'] = pd.to_datetime(corr_df['日期_obj'])
                    corr_df['week'] = corr_df['日期_dt'].dt.isocalendar(
                    ).week.astype(int)
                    corr_df['year'] = corr_df['日期_dt'].dt.isocalendar(
                    ).year.astype(int)
                    corr_df['week_start'] = corr_df['日期_dt'].apply(
                        lambda x: x - pd.Timedelta(days=x.dayofweek))

                    weekly_corr = corr_df.groupby('week_start').agg(
                        採購金額=('peak_cost', 'sum'),
                        來客人數=('effective_peak_guests', 'sum')
                    ).reset_index()
                    weekly_corr['週次'] = weekly_corr['week_start'].dt.strftime(
                        'W%V\n%m/%d')

                    # 標準化成 0–100%（對各自最大值）
                    max_cost = weekly_corr['採購金額'].max()
                    max_guest = weekly_corr['來客人數'].max()
                    weekly_corr['採購(%)'] = (
                        weekly_corr['採購金額'] / max_cost * 100).round(1) if max_cost > 0 else 0
                    weekly_corr['來客(%)'] = (
                        weekly_corr['來客人數'] / max_guest * 100).round(1) if max_guest > 0 else 0
                    weekly_corr['背道而馳'] = (abs(
                        weekly_corr['採購(%)'] - weekly_corr['來客(%)']) > 25).map({True: '⚠️ 異常', False: '✅ 正常'})

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
                            y=alt.Y('標準化數值:Q', title='相對比例 (% of max)',
                                    scale=alt.Scale(domain=[0, 110])),
                            color=alt.Color('指標:N',
                                            scale=alt.Scale(domain=list(
                                                color_map.keys()), range=list(color_map.values())),
                                            legend=alt.Legend(
                                                title='指標', orient='bottom')
                                            ),
                            tooltip=[
                                alt.Tooltip('週次:N', title='週次'),
                                alt.Tooltip('指標:N', title='指標'),
                                alt.Tooltip(
                                    '採購金額:Q', title='採購金額 (NT$)', format=',.0f'),
                                alt.Tooltip(
                                    '來客人數:Q', title='來客人數 (人)', format=',.0f'),
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
                                st.warning(
                                    f"⚠️ **{bw['週次'].replace(chr(10), ' ')}** 出現背道而馳！{direction}　採購 NT$ {int(bw['採購金額']):,} | 來客 {int(bw['來客人數'])} 人")
                        else:
                            st.success("✅ 本月各週採購花費與來客人數走勢一致，食材控管健康。")
                    else:
                        st.info("💡 本月資料不足，無法進行相關性分析。")

                    st.divider()
                    c_ana1, c_ana2 = st.columns(2)

                    with c_ana1:
                        st.markdown(
                            f"<div style='background:#f8f9fa; padding:15px; border-radius:10px; border-top:4px solid #1f2c56;'>", unsafe_allow_html=True)
                        st.markdown(f"**🏰 The Peak (餐廳)**")
                        st.metric("本月總採購額", f"NT$ {int(total_peak_cost):,}")
                        is_auto = "(自動加總)" if (
                            df_daily_rest['rest_day_guests'].sum() == 0 and total_peak_guests > 0) else ""
                        st.metric(f"本月總來客數 {is_auto}",
                                  f"{int(total_peak_guests):,} 人")

                        # CPG 顏色警示 (目標 $150)
                        peak_target = 150
                        delta_val = peak_target - final_peak_cpg
                        st.metric("平均每客成本 (CPG)", f"NT$ {int(final_peak_cpg):,}", delta=f"{int(delta_val)} (距離目標)" if delta_val >=
                                  0 else f"{int(delta_val)} (已超標)", delta_color="normal" if delta_val >= 0 else "inverse")
                        st.markdown("</div>", unsafe_allow_html=True)

                        # --- 新增：財務預測與目標控管 ---
                        st.write("")
                        st.markdown("##### 🎯 財務目標控管")
                        # 1. 成本佔比 (餐飲成本 / 總營收)
                        total_hotel_rev = m_data['rev']
                        cost_ratio = (
                            total_peak_cost / total_hotel_rev * 100) if total_hotel_rev > 0 else 0
                        st.write(f"📊 目前成本佔總營收比例: **{cost_ratio:.1f}%**")

                        # 2. 月底支出預測
                        import calendar
                        _, last_day_num = calendar.monthrange(
                            selected_date.year, selected_date.month)
                        current_day_num = len(analysis_df)
                        if current_day_num > 0:
                            daily_avg_cost = total_peak_cost / current_day_num
                            forecast_total = total_peak_cost + \
                                (daily_avg_cost * (last_day_num - current_day_num))

                            forecast_color = "red" if final_peak_cpg > peak_target else "green"
                            st.markdown(
                                f"🔮 月底預估總支出: <span style='color:{forecast_color}; font-weight:bold;'>NT$ {int(forecast_total):,}</span>", unsafe_allow_html=True)
                            if final_peak_cpg > peak_target:
                                st.warning(
                                    f"⚠️ 警告：目前每客成本 ({int(final_peak_cpg)}) 已高於目標 {peak_target} 元，請檢視進貨項目或份量控管。")
                        # ----------------------------
                    with c_ana2:
                        st.markdown(
                            f"<div style='background:#fff9f0; padding:15px; border-radius:10px; border-top:4px solid #ff9f43;'>", unsafe_allow_html=True)
                        st.markdown(f"**🥂 Happy Hour (HH)**")
                        st.metric("本月總採購額", f"NT$ {int(total_hh_cost):,}")
                        st.metric("本月總來客數", f"{int(total_hh_guests):,} 人")
                        st.metric("平均每客服務成本", f"NT$ {int(final_hh_cpg):,}")
                        if total_hh_cost > 0 and total_hh_guests == 0:
                            st.warning(
                                "⚠️ 有產生 HH 採購費用但總來客數為 0！請至「🍽️ 餐廳數據」補登以計算每客服務成本 (CPG)。")
                        st.markdown("</div>", unsafe_allow_html=True)

                    st.write("")
                    # 趨勢圖表
                    st.markdown("#### 📈 本月累計每客成本趨勢 (Monthly Cumulative CPG)")
                    analysis_df['日期_str'] = analysis_df['日期_obj'].astype(str)

                    # 整合圖表
                    base_chart = alt.Chart(analysis_df).encode(
                        x=alt.X('日期_str:O', title='日期'))

                    peak_line = base_chart.mark_line(point=True, color='#1f2c56', strokeWidth=3).encode(
                        y=alt.Y('cum_peak_cpg:Q', title='累計平均成本 (NT$)'),
                        tooltip=['日期_str', alt.Tooltip('cum_peak_guests', title='累計來客'), alt.Tooltip(
                            'cum_peak_cost', title='累計採購'), alt.Tooltip('cum_peak_cpg', format='.0f', title='累計 CPG')]
                    )

                    st.altair_chart(peak_line.properties(
                        title="The Peak 累計平均成本趨勢", height=300), use_container_width=True)

                    if total_hh_guests > 0:
                        st.write("")
                        st.markdown("#### 🥂 Happy Hour 累計成本分析")

                        # 顯示累計人數 vs 累計成本
                        hh_chart_base = alt.Chart(analysis_df).encode(
                            x=alt.X('日期_str:O', title='日期'))

                        # 長條圖顯示累計 CPG
                        hh_bar = hh_chart_base.mark_bar(color='#ff9f43', opacity=0.7).encode(
                            y=alt.Y('cum_hh_cpg:Q', title='累計平均成本 (NT$)'),
                            tooltip=[
                                '日期_str',
                                alt.Tooltip('cum_hh_guests',
                                            title='累計來客 (分母)'),
                                alt.Tooltip('cum_hh_cost', title='累計採購 (分子)'),
                                alt.Tooltip(
                                    'cum_hh_cpg', format='.1f', title='累計 CPG')
                            ]
                        )

                        # 疊加一條線顯示累計人數的成長 (確保分母正確)
                        hh_guest_line = hh_chart_base.mark_line(color='#e67e22', strokeDash=[5, 5]).encode(
                            y=alt.Y('cum_hh_guests:Q', title='累計人數'),
                            tooltip=['日期_str', alt.Tooltip(
                                'cum_hh_guests', title='累計人數')]
                        )

                        st.altair_chart(alt.layer(hh_bar, hh_guest_line).resolve_scale(y='independent').properties(
                            title="Happy Hour 累計趨勢 (長條:成本, 虛線:人數)", height=300), use_container_width=True)

                    # --- 新增：雙冠日食材消耗對比分析 (Dynamic CPG Analysis) ---
                    st.divider()
                    st.markdown("#### 🎯 雙冠日 vs 一般日：食材消耗對比分析")

                    # 獲取雙冠日清單
                    curr_metrics = calc_key_metrics(m_data)
                    dual_match_dates = curr_metrics.get('dual_match_dates', [])

                    if dual_match_dates:
                        # 將日期標記為雙冠日
                        analysis_df['is_dual_match'] = analysis_df['日期_str'].isin(
                            dual_match_dates)

                        df_dual = analysis_df[analysis_df['is_dual_match']]
                        df_normal = analysis_df[~analysis_df['is_dual_match']]

                        # 計算雙冠日 CPG
                        dual_peak_cost = df_dual['peak_cost'].sum()
                        dual_peak_guests = df_dual['effective_peak_guests'].sum()
                        dual_cpg = dual_peak_cost / dual_peak_guests if dual_peak_guests > 0 else 0
                        
                        dual_bf = df_dual['bf_act'].sum()
                        dual_af = df_dual['af_act'].sum()
                        if dual_bf + k_val * dual_af > 0:
                            dual_cb = (dual_bf + dual_af) * dual_cpg / (dual_bf + k_val * dual_af)
                            dual_ca = k_val * dual_cb
                        else:
                            dual_cb = dual_cpg
                            dual_ca = dual_cpg

                        # 計算一般日 CPG
                        normal_peak_cost = df_normal['peak_cost'].sum()
                        normal_peak_guests = df_normal['effective_peak_guests'].sum()
                        normal_cpg = normal_peak_cost / normal_peak_guests if normal_peak_guests > 0 else 0
                        
                        normal_bf = df_normal['bf_act'].sum()
                        normal_af = df_normal['af_act'].sum()
                        if normal_bf + k_val * normal_af > 0:
                            normal_cb = (normal_bf + normal_af) * normal_cpg / (normal_bf + k_val * normal_af)
                            normal_ca = k_val * normal_cb
                        else:
                            normal_cb = normal_cpg
                            normal_ca = normal_cpg

                        cpg_col1, cpg_col2 = st.columns(2)

                        with cpg_col1:
                            st.markdown(f"""
                            <div style="background:#fff5e6; border-left:4px solid #e67e22; padding:15px; border-radius:8px;">
                                <p style="margin:0; font-size:13px; color:#e67e22; font-weight:bold;">🏆 雙冠日 (共 {len(df_dual)} 天)</p>
                                <h3 style="margin:5px 0;">總平均 NT$ {int(dual_cpg):,} / 客</h3>
                                <div style="display:flex; gap:10px; margin: 10px 0;">
                                    <div style="background:rgba(230,126,34,0.1); padding:5px 10px; border-radius:5px;">🍳 早 <span style="font-weight:bold;">NT${dual_cb:.1f}</span></div>
                                    <div style="background:rgba(230,126,34,0.1); padding:5px 10px; border-radius:5px;">🍰 午 <span style="font-weight:bold;">NT${dual_ca:.1f}</span></div>
                                </div>
                                <p style="margin:0; font-size:12px; color:#666;">總食材花費: NT$ {int(dual_peak_cost):,} | 服務客數: {int(dual_peak_guests):,} 人</p>
                            </div>
                            """, unsafe_allow_html=True)

                        with cpg_col2:
                            st.markdown(f"""
                            <div style="background:#f8f9fa; border-left:4px solid #95a5a6; padding:15px; border-radius:8px;">
                                <p style="margin:0; font-size:13px; color:#7f8c8d; font-weight:bold;">📉 一般日 (共 {len(df_normal)} 天)</p>
                                <h3 style="margin:5px 0;">總平均 NT$ {int(normal_cpg):,} / 客</h3>
                                <div style="display:flex; gap:10px; margin: 10px 0;">
                                    <div style="background:rgba(149,165,166,0.1); padding:5px 10px; border-radius:5px;">🍳 早 <span style="font-weight:bold;">NT${normal_cb:.1f}</span></div>
                                    <div style="background:rgba(149,165,166,0.1); padding:5px 10px; border-radius:5px;">🍰 午 <span style="font-weight:bold;">NT${normal_ca:.1f}</span></div>
                                </div>
                                <p style="margin:0; font-size:12px; color:#666;">總食材花費: NT$ {int(normal_peak_cost):,} | 服務客數: {int(normal_peak_guests):,} 人</p>
                            </div>
                            """, unsafe_allow_html=True)

                        # 顯示策略建議（基於比例，而非絕對差值）
                        st.write("")
                        target_ratio = 1.10  # 雙冠日 CPG 應達到一般日的 110%
                        actual_ratio = (
                            dual_cpg / normal_cpg) if normal_cpg > 0 else 0
                        ratio_pct = actual_ratio * 100

                        if actual_ratio >= target_ratio:
                            st.success(
                                f"💡 **主動備戰策略成功！** 雙冠日的單客成本（NT$ {int(dual_cpg):,}）達到一般日（NT$ {int(normal_cpg):,}）的 **{ratio_pct:.0f}%**，超過 110% 目標。代表你在大日子前有主動備了更好的食材，與高房價形成正向配對。")
                        elif actual_ratio >= 0.90:
                            diff_to_target = int(
                                normal_cpg * target_ratio - dual_cpg)
                            st.info(
                                f"⚖️ **採購尚未主動分級。** 雙冠日 CPG 為一般日的 {ratio_pct:.0f}%（目標 ≥ 110%）。由於週均攤讓旺日（人多）天然壓低 CPG，這個差距屬於合理的規模效應。建議在雙冠日當週多編列 NT$ {diff_to_target:,} / 人左右的品質預算，讓高端客感受得到差異。")
                        else:
                            # 計算行動指引數據
                            avg_normal_guests = (
                                normal_peak_guests / len(df_normal)) if len(df_normal) > 0 else 0
                            peak_target_cpg = 150  # 目標 CPG 上限（與前面財務目標一致）
                            # 建議週採購上限：目標 CPG × 一般日平均每日來客數 × 7 天
                            recommended_weekly_budget = int(
                                peak_target_cpg * avg_normal_guests * 7)
                            # 本月實際週均採購
                            total_weeks = max(1, round(len(df_normal) / 7))
                            actual_weekly_avg = int(
                                normal_peak_cost / total_weeks) if total_weeks > 0 else 0
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
                        st.caption(
                            "💡 虛線代表累積來客數。如果長條圖在月初是空的，代表該時段尚未產生 HH 相關的採購支出。")

                    st.info(
                        "💡 **分析小撇步**：當「每客成本」異常偏高時，請檢查該日期是否有大宗採購進入庫存，或來客數輸入是否正確。")

                else:
                    st.info("尚未偵測到本月的餐廳來客數據，無法進行成本效益分析。")

                st.divider()

                # 3. 各部門詳細統計
                st.subheader("🏢 各部門經費分析")

                # 取得所有部門
                departments = dept_summary.sort_values(
                    '小計', ascending=False)['部門'].tolist()

                for dept in departments:
                    dept_df = df_month[df_month[dept_col] == dept].copy()
                    dept_total = dept_df['小計'].sum()

                    with st.expander(f"📌 {dept} (總計: NT$ {int(dept_total):,})", expanded=False):
                        # --- 新增：Top 5 高額品項排行榜 ---
                        item_name_col = next((c for c in dept_df.columns if any(
                            k in c for k in ['品名', '項目', 'Item'])), None)
                        if item_name_col:
                            st.markdown("##### 🏆 前五名高額採購品項")
                            top_items = dept_df.groupby(item_name_col)['小計'].sum(
                            ).sort_values(ascending=False).head(5).reset_index()
                            t_cols = st.columns(5)
                            for idx, row in top_items.iterrows():
                                with t_cols[idx]:
                                    st.metric(
                                        f"No.{idx+1} {row[item_name_col][:8]}", f"NT$ {int(row['小計']):,}")
                        st.divider()

                        # --- 新增：排序控制 ---
                        sort_by = st.selectbox(f"排序方式 ({dept})", [
                                               "日期 (新→舊)", "金額 (高→低)", "金額 (低→高)", "品項名稱"], key=f"sort_{dept}")

                        if sort_by == "金額 (高→低)":
                            dept_df = dept_df.sort_values(
                                '小計', ascending=False)
                        elif sort_by == "金額 (低→高)":
                            dept_df = dept_df.sort_values('小計', ascending=True)
                        elif sort_by == "日期 (新→舊)":
                            dept_df = dept_df.sort_values(
                                '日期', ascending=False)
                        elif sort_by == "品項名稱" and item_name_col:
                            dept_df = dept_df.sort_values(item_name_col)

                        # 顯示該部門表格
                        cols_to_show = [c for c in [
                            '日期', '供應商', '品名', '規格', '數量', '單位', '單價', '小計'] if c in dept_df.columns]
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
                    st.caption(
                        "分析特定關鍵食材品項（如：蛋、高麗菜、海鮮等）的每客平均消耗量，並自動產出精準叫貨配比建議。")

                    item_col = next((c for c in df_month.columns if any(
                        k in c for k in ['品名', '項目', 'Item'])), None)
                    qty_col = next((c for c in df_month.columns if any(
                        k in c for k in ['數量', 'Qty', 'Quantity'])), None)
                    unit_col = next((c for c in df_month.columns if any(
                        k in c for k in ['單位', 'Unit'])), None)
                    price_col = next((c for c in df_month.columns if any(
                        k in c for k in ['單價', 'Price', 'Rate'])), None)

                    if item_col and qty_col:
                        # 擷取常用關鍵字選項
                        all_items = df_month[item_col].dropna().astype(
                            str).str.strip()
                        all_items = all_items[all_items != ""]

                        common_keywords = ["蛋", "菜", "肉", "奶",
                                           "米", "麵", "油", "海鮮", "雞", "豬", "牛", "魚"]
                        found_keywords = [k for k in common_keywords if any(
                            k in x for x in all_items)]
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
                                search_term = st.text_input(
                                    "✍️ 輸入自訂食材名稱 (例如: 高麗菜)", "蛋", key="item_analysis_custom_input")
                            else:
                                search_term = selected_keyword

                        # 篩選匹配的採購項目
                        item_mask = df_month[item_col].astype(
                            str).str.contains(search_term, na=False, case=False)
                        item_df = df_month[item_mask].copy()

                        if not item_df.empty:
                            # 數值清理
                            item_df['cleaned_qty'] = pd.to_numeric(item_df[qty_col].astype(
                                str).str.replace(',', ''), errors='coerce').fillna(0)
                            item_df['cleaned_total'] = pd.to_numeric(item_df['小計'].astype(
                                str).str.replace(',', ''), errors='coerce').fillna(0)

                            # 單位判斷
                            most_common_unit = "單位"
                            if unit_col in item_df.columns:
                                most_common_unit = item_df[unit_col].mode(
                                ).iloc[0] if not item_df[unit_col].empty else "單位"

                            # 每日採購整合
                            item_df['日期_obj'] = pd.to_datetime(
                                item_df['日期']).dt.date
                            daily_item_qty = item_df.groupby(
                                '日期_obj')['cleaned_qty'].sum().reset_index()
                            daily_item_cost = item_df.groupby(
                                '日期_obj')['cleaned_total'].sum().reset_index()

                            # 合併每日來客
                            item_analysis_df = analysis_df[[
                                '日期_obj', 'effective_peak_guests']].copy()
                            item_analysis_df = pd.merge(
                                item_analysis_df, daily_item_qty, on='日期_obj', how='left').fillna(0)
                            item_analysis_df = pd.merge(
                                item_analysis_df, daily_item_cost, on='日期_obj', how='left').fillna(0)

                            # 週彙總計算
                            item_analysis_df['日期_dt'] = pd.to_datetime(
                                item_analysis_df['日期_obj'])
                            item_analysis_df['week_start'] = item_analysis_df['日期_dt'].apply(
                                lambda x: x - pd.Timedelta(days=x.dayofweek))

                            weekly_item = item_analysis_df.groupby('week_start').agg(
                                總採購量=('cleaned_qty', 'sum'),
                                總費用=('cleaned_total', 'sum'),
                                來客人數=('effective_peak_guests', 'sum')
                            ).reset_index()

                            weekly_item['週次'] = pd.to_datetime(
                                weekly_item['week_start']).dt.strftime('W%V\n%m/%d')
                            weekly_item['每客平均消耗量'] = weekly_item.apply(
                                lambda r: r['總採購量'] / r['來客人數'] if r['來客人數'] > 0 else 0, axis=1
                            )

                            # 計算月平均與平均單價
                            total_qty_month = weekly_item['總採購量'].sum()
                            total_guests_month = weekly_item['來客人數'].sum()
                            avg_rate_month = total_qty_month / \
                                total_guests_month if total_guests_month > 0 else 0
                            avg_unit_price = item_df['cleaned_total'].sum(
                            ) / item_df['cleaned_qty'].sum() if item_df['cleaned_qty'].sum() > 0 else 0

                            st.write("")
                            st.markdown(f"##### 📊 **「{search_term}」消耗數據指標**")

                            c_m1, c_m2, c_m3 = st.columns(3)
                            c_m1.metric(
                                "本月總採購量", f"{total_qty_month:,.1f} {most_common_unit}")
                            c_m2.metric(
                                "每客平均消耗量 (使用率)", f"{avg_rate_month:.2f} {most_common_unit}/人", help="總採購量 / 總來客數")
                            c_m3.metric(
                                "平均採購單價", f"NT$ {avg_unit_price:,.1f} /{most_common_unit}")

                            # 圖表呈現
                            st.write("")
                            st.markdown(
                                f"###### 📈 週來客數 vs 「{search_term}」採購量相對走勢")

                            max_w_qty = weekly_item['總採購量'].max()
                            max_w_guests = weekly_item['來客人數'].max()
                            weekly_item['採購量(%)'] = (
                                weekly_item['總採購量'] / max_w_qty * 100).round(1) if max_w_qty > 0 else 0
                            weekly_item['來客(%)'] = (
                                weekly_item['來客人數'] / max_w_guests * 100).round(1) if max_w_guests > 0 else 0

                            melt_item_df = weekly_item.melt(
                                id_vars=['週次', '總採購量', '來客人數', '總費用'],
                                value_vars=['採購量(%)', '來客(%)'],
                                var_name='指標', value_name='標準化數值'
                            )

                            item_color_map = {
                                '採購量(%)': '#e67e22', '來客(%)': '#2980b9'}

                            item_chart = alt.Chart(melt_item_df).mark_line(point=True, strokeWidth=2.5).encode(
                                x=alt.X('週次:N', title='週次', sort=None),
                                y=alt.Y('標準化數值:Q', title='相對比例 (% of max)',
                                        scale=alt.Scale(domain=[0, 110])),
                                color=alt.Color('指標:N',
                                                scale=alt.Scale(domain=list(item_color_map.keys()), range=list(
                                                    item_color_map.values())),
                                                legend=alt.Legend(
                                                    title='指標', orient='bottom')
                                                ),
                                tooltip=[
                                    alt.Tooltip('週次:N', title='週次'),
                                    alt.Tooltip('指標:N', title='指標'),
                                    alt.Tooltip(
                                        '總採購量:Q', title=f'總採購量 ({most_common_unit})', format=',.1f'),
                                    alt.Tooltip(
                                        '來客人數:Q', title='來客人數 (人)', format=',.0f'),
                                    alt.Tooltip(
                                        '總費用:Q', title='總費用 (NT$)', format=',.0f'),
                                ]
                            ).properties(height=200)

                            st.altair_chart(
                                item_chart, use_container_width=True)

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
                                    value=int(
                                        total_guests_month / 4) if total_guests_month > 0 else 500,
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
                                    st.markdown("- **週一 (40%)**：建議採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(
                                        recommended_qty * 0.4, most_common_unit, int(est_cost * 0.4)))
                                    st.markdown("- **週三 (30%)**：建議採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(
                                        recommended_qty * 0.3, most_common_unit, int(est_cost * 0.3)))
                                    st.markdown("- **週五 (30%)**：建議採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(
                                        recommended_qty * 0.3, most_common_unit, int(est_cost * 0.3)))
                                elif vendor_type == "菜商":
                                    st.markdown("- **平日每日均攤 (60%)**：每次到貨建議 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(
                                        recommended_qty * 0.12, most_common_unit, int(est_cost * 0.12)))
                                    st.markdown("- **週五加強 (40%)**：一次叫足 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(
                                        recommended_qty * 0.4, most_common_unit, int(est_cost * 0.4)))
                                else:
                                    st.markdown("- **單次足額採購 (100%)**：於週一或合約到貨日一次性採購 **`{:.1f}`** {}（預估單次費用：**NT$ {:,}**）".format(
                                        recommended_qty, most_common_unit, int(est_cost)))

                        else:
                            st.warning(
                                f"⚠️ 在目前的採購資料中，找不到含有「{search_term}」的品項名稱。")
                            st.info("💡 請嘗試選擇其他常用關鍵字，或自訂輸入更精確的關鍵字（如：雞蛋、高麗菜）。")
                    
        # ── C. 本期菜價總覽 ──────────────────────────────
        st.markdown("#### 📋 C. 本期菜價總覽")
        latest_period = periods_available[-1]
        latest_df = sp_df[sp_df['period_dt'] == latest_period].copy()

        # ── 計算歷史請購頻率 (從 thepeak_daily_purchase_report 及 4FHH) ────────
        df_peak_hist = fetch_thepeak_daily_purchase_report()
        df_hh_hist = fetch_4fhh_daily_purchase_report()
        
        freq_map = {}  # item_name -> purchase count
        for _h_df in [df_peak_hist, df_hh_hist]:
            if not _h_df.empty:
                _item_col = next((c for c in _h_df.columns if '品項' in c or 'item' in c.lower()), None)
                if _item_col:
                    for item_val in _h_df[_item_col].dropna().astype(str).values:
                        # 品項名稱有時候包含數量資訊如 "雞蛋 (10 個)", 嘗試取純品項名
                        item_clean = item_val.split('(')[0].strip()
                        freq_map[item_clean] = freq_map.get(item_clean, 0) + 1

        # 用品項名模糊對應到 sp_df 中的 item_name (帶寬鬆匹配)
        all_sp_items = latest_df['item_name'].astype(str).tolist()
        frequent_items_set = set()
        item_freq_count = {}  # item_name -> total frequency
        for sp_item in all_sp_items:
            matched_count = 0
            for freq_item, cnt in freq_map.items():
                if sp_item in freq_item or freq_item in sp_item:
                    matched_count += cnt
            if matched_count > 0:
                frequent_items_set.add(sp_item)
                item_freq_count[sp_item] = matched_count
        
        # 也從請購單歷史品項名稱中直接配對到 sp_df
        for freq_item, cnt in freq_map.items():
            if freq_item in all_sp_items:
                frequent_items_set.add(freq_item)
                item_freq_count[freq_item] = item_freq_count.get(freq_item, 0) + cnt
        
        n_frequent = len(frequent_items_set)
        
        if n_periods >= 2:
            prev_period = periods_available[-2]
            prev_df = sp_df[sp_df['period_dt'] == prev_period][['item_name', 'unit', 'price']].rename(columns={'price': 'prev_price'})
            latest_df = latest_df.merge(prev_df, on=['item_name', 'unit'], how='left')
            latest_df['change'] = latest_df['price'] - latest_df['prev_price']
            # 漏洞4修復：防止 prev_price 為 0 時產生除以零錯誤
            latest_df['change_pct'] = latest_df.apply(
                lambda r: round(r['change'] / r['prev_price'] * 100, 1)
                if pd.notna(r.get('prev_price')) and r.get('prev_price', 0) > 0
                else float('nan'),
                axis=1
            )

            # --- 計算今年至今(全歷史)的純平基準與極值，並統計每個品項的歷史期數 ---
            current_year = selected_date.year
            ytd_sp_df = sp_df[sp_df['period_dt'].apply(lambda x: getattr(x, 'year', None)) == current_year] if not sp_df.empty else sp_df
            ytd_stats = ytd_sp_df.groupby(['item_name', 'unit'])['price'].agg(['mean', 'max', 'min', 'count']).reset_index()
            ytd_stats.rename(columns={'mean': 'ytd_avg', 'max': 'ytd_max', 'min': 'ytd_min', 'count': 'ytd_periods'}, inplace=True)
            latest_df = latest_df.merge(ytd_stats, on=['item_name', 'unit'], how='left')

            def fmt_change(row):
                if pd.isna(row.get('change')):
                    return '—'
                sign = '+' if row['change'] > 0 else ''
                color = '#e74c3c' if row['change'] > 0 else ('#2ecc71' if row['change'] < 0 else '#888')
                arrow = '▲' if row['change'] > 0 else ('▼' if row['change'] < 0 else '─')
                return f"<span style='color:{color};font-weight:bold;'>{arrow} {sign}{row['change_pct']:.1f}%</span>"

            latest_df['漲跌'] = latest_df.apply(fmt_change, axis=1)

            def fmt_ytd(row):
                if pd.isna(row.get('ytd_avg')):
                    return '—'
                avg_p = row['ytd_avg']
                curr_p = row['price']
                if pd.isna(curr_p) or avg_p == 0:
                    return '—'

                diff_pct = ((curr_p - avg_p) / avg_p * 100)
                sign = '+' if diff_pct > 0 else ''
                color = '#e74c3c' if diff_pct > 0 else ('#2ecc71' if diff_pct < 0 else '#888')
                text = f"<span style='color:{color};'>{sign}{diff_pct:.1f}%</span>"

                badges = ""
                if row['ytd_max'] > row['ytd_min']:
                    if curr_p >= row['ytd_max']:
                        badges = " <span style='background:#e74c3c;color:white;font-size:10px;padding:2px 4px;border-radius:4px;margin-left:4px;'>歷史高點</span>"
                    elif curr_p <= row['ytd_min']:
                        badges = " <span style='background:#2ecc71;color:white;font-size:10px;padding:2px 4px;border-radius:4px;margin-left:4px;'>歷史低點</span>"

                return f"均 {avg_p:.1f} ({text}){badges}"

            latest_df['純平基準對照'] = latest_df.apply(fmt_ytd, axis=1)
            
            # 加入請購頻率欄位
            latest_df['請購次數'] = latest_df['item_name'].map(lambda x: item_freq_count.get(x, 0))
            latest_df['_is_frequent'] = latest_df['item_name'].isin(frequent_items_set)
            
            display_cols = ['item_name', 'unit', 'price', '漲跌', '純平基準對照', '請購次數']
            col_rename = {'item_name': '品項', 'unit': '單位', 'price': f'本期單價 ({latest_period})', '請購次數': '📦 請購次數'}
        else:
            latest_df['請購次數'] = latest_df['item_name'].map(lambda x: item_freq_count.get(x, 0))
            latest_df['_is_frequent'] = latest_df['item_name'].isin(frequent_items_set)
            display_cols = ['item_name', 'unit', 'price', '請購次數']
            col_rename = {'item_name': '品項', 'unit': '單位', 'price': f'本期單價 ({latest_period})', '請購次數': '📦 請購次數'}

        # ── 計算完全部資料後再儲存 full_latest_df，供 D/E/F 使用（包含 change/ytd 欄位）──
        full_latest_df = latest_df.copy()

        # ── 搜尋、篩選、排序 UI ─────────────────────────────
        # 主要篩選：以 radio 按鈕群組呈現，強調「經常性請購」為預設
        c_info, c_filter = st.columns([2, 1])
        with c_info:
            if n_frequent > 0:
                st.info(f"📦 已從歷史請購紀錄中識別出 **{n_frequent}** 個經常性採購品項（共從 `thepeak_daily_purchase_report` 及 `4FHH_daily_purchase_report` 統計得出）。")
            else:
                st.info("💡 尚無歷史請購紀錄可供分析。送出請購單後，「經常性請購」分類將自動建立。")
        
        with c_filter:
            pass  # placeholder

        search_kw = st.text_input("🔍 搜尋品項", placeholder="輸入關鍵字，如：高麗菜、雞蛋")
        
        # 篩選器一列三欄
        sort_col, filter_col, extra_col = st.columns(3)
        with sort_col:
            sort_option = st.selectbox("↕️ 排序方式", [
                "📦 請購頻率：由高至低",
                "預設 (不排序)",
                "💰 本期單價：由高至低",
                "💰 本期單價：由低至高",
                "📈 漲幅：由高至低 (漲最多)",
                "📉 跌幅：由低至高 (跌最多)"
            ])
            
        with filter_col:
            filter_option = st.selectbox("🏷️ 狀態篩選", [
                "⭐ 經常性請購",
                "📋 所有品項",
                "🚨 歷史高點",
                "✅ 歷史低點",
                "─ 正常區間"
            ])

        with extra_col:
            min_freq = st.number_input(
                "🔢 最低請購次數篩選",
                min_value=0,
                value=0,
                step=1,
                help="只顯示歷史請購次數 ≥ 此值的品項（設為 0 則不過濾）"
            )

        # 1. 套用搜尋
        if search_kw:
            latest_df = latest_df[latest_df['item_name'].str.contains(search_kw, na=False)]

        # 2. 套用狀態篩選
        if filter_option == "⭐ 經常性請購":
            latest_df = latest_df[latest_df['_is_frequent']].copy()
        elif filter_option != "📋 所有品項" and 'ytd_max' in latest_df.columns:
            is_high = (latest_df['price'] >= latest_df['ytd_max']) & (latest_df['ytd_max'] > latest_df['ytd_min'])
            is_low = (latest_df['price'] <= latest_df['ytd_min']) & (latest_df['ytd_max'] > latest_df['ytd_min'])
            if filter_option == "🚨 歷史高點":
                latest_df = latest_df[is_high]
            elif filter_option == "✅ 歷史低點":
                latest_df = latest_df[is_low]
            elif filter_option == "─ 正常區間":
                latest_df = latest_df[~(is_high | is_low)]

        # 3. 套用最低請購次數篩選
        if min_freq > 0:
            latest_df = latest_df[latest_df['請購次數'] >= min_freq]

        # 4. 套用排序
        if sort_option == "📦 請購頻率：由高至低":
            latest_df = latest_df.sort_values('請購次數', ascending=False)
        elif sort_option == "💰 本期單價：由高至低":
            latest_df = latest_df.sort_values('price', ascending=False)
        elif sort_option == "💰 本期單價：由低至高":
            latest_df = latest_df.sort_values('price', ascending=True)
        elif "漲幅：由高至低" in sort_option and 'change_pct' in latest_df.columns:
            latest_df = latest_df.sort_values('change_pct', ascending=False)
        elif "跌幅：由低至高" in sort_option and 'change_pct' in latest_df.columns:
            latest_df = latest_df.sort_values('change_pct', ascending=True)
        
        # 若是「經常性請購」模式預設按頻率排序（若使用者未選其他）
        if filter_option == "⭐ 經常性請購" and sort_option == "📦 請購頻率：由高至低":
            latest_df = latest_df.sort_values('請購次數', ascending=False)

        # 5. 新增「今日是否在菜價表」欄位 (供快速確認用)
        if '品項' not in latest_df.columns:
            show_df = latest_df[[c for c in display_cols if c in latest_df.columns]].rename(columns=col_rename)
        else:
            show_df = latest_df[[c for c in display_cols if c in latest_df.columns]].rename(columns=col_rename)

        # 顯示結果計數
        total_items_showing = len(show_df)
        mode_label = "（⭐ 經常性請購）" if filter_option == "⭐ 經常性請購" else f"（{filter_option}）"
        st.caption(f"🔎 目前顯示 **{total_items_showing}** 個品項 {mode_label}")

        if n_periods >= 2:
            st.write(show_df.to_html(escape=False, index=False), unsafe_allow_html=True)
        else:
            st.dataframe(show_df, use_container_width=True, hide_index=True)
        
        # ── 快速提示卡：「經常性請購」模式下的注意品項 ───────────────
        if filter_option == "⭐ 經常性請購" and n_periods >= 2 and not latest_df.empty and 'ytd_max' in latest_df.columns:
            _alert_highs = latest_df[(latest_df['price'] >= latest_df['ytd_max']) & (latest_df['ytd_max'] > latest_df['ytd_min']) & latest_df['_is_frequent']]
            if not _alert_highs.empty:
                _high_names = '、'.join(_alert_highs['item_name'].head(5).tolist())
                st.error(f"🚨 **常購品項中有歷史高點！** `{_high_names}` 目前為今年最貴，建議本週評估替代或減量採購。")
            _alert_lows = latest_df[(latest_df['price'] <= latest_df['ytd_min']) & (latest_df['ytd_max'] > latest_df['ytd_min']) & latest_df['_is_frequent']]
            if not _alert_lows.empty:
                _low_names = '、'.join(_alert_lows['item_name'].head(5).tolist())
                st.success(f"✅ **常購品項中有歷史低點！** `{_low_names}` 目前為今年最便宜，是囤貨的好時機！")


    except Exception as e:
        if "WorksheetNotFound" in str(e):
            st.error(f"❌ 找不到採購相關分頁！請確認 Google Sheet 中的分頁名稱（如 purchase data）。")
        else:
            st.error(f"讀取採購數據出錯: {e}")
        import traceback
        st.expander("錯誤詳細資訊").code(traceback.format_exc())

# --- 📅 本月接下來各週採購金額建議（獨立區塊，不依賴採購數據）---
if selected_page == "💰 採購分析":
    st.divider()
    st.markdown("#### 📅 本月接下來各週採購金額建議")

    from datetime import date as dt_date, timedelta as dt_timedelta
    import calendar as cal_lib
    today_dt2 = dt_date.today()
    _, last_day_num2 = cal_lib.monthrange(
        selected_date.year, selected_date.month)
    month_end_dt2 = dt_date(
        selected_date.year, selected_date.month, last_day_num2)

    is_cur_or_fut = (selected_date.year, selected_date.month) >= (
        today_dt2.year, today_dt2.month)

    if is_cur_or_fut:
        # 新邏輯：直接讀取未來的預約資料
        df_fb_fut = get_combined_fb_future_data()
        if not df_fb_fut.empty and 'date' in df_fb_fut.columns:
            df_fb_fut['date_obj'] = pd.to_datetime(df_fb_fut['date']).dt.date
        else:
            df_fb_fut = pd.DataFrame()
            
        # 取得雙冠日清單（來自 tab_m 的 calc_key_metrics）
        fw_m_data = fetch_month_summary(selected_date.year, selected_date.month)
        fw_curr_metrics = calc_key_metrics(fw_m_data)
        fw_dual_dates = set(fw_curr_metrics.get('dual_match_dates', []))

        # 取得剩餘可用總預算 (從上方戰情室傳遞)
        rem_budget_val = remaining_budget if 'remaining_budget' in locals() else 0
        rem_budget_val = max(0, rem_budget_val)  # 若為負數(超支)，強制配額為 0
        
        # 先計算這段時間的未來總客數，以計算分配比例
        temp_cursor = today_dt2
        temp_seen = set()
        overall_future_guests = 0
        while temp_cursor <= month_end_dt2:
            mon = temp_cursor - dt_timedelta(days=temp_cursor.weekday())
            sun = mon + dt_timedelta(days=6)
            ws = max(mon, dt_date(selected_date.year, selected_date.month, 1))
            we = min(sun, month_end_dt2)
            if ws not in temp_seen and we >= today_dt2:
                temp_seen.add(ws)
                if not df_fb_fut.empty and '數量' in df_fb_fut.columns:
                    w_mask = (df_fb_fut['date_obj'] >= ws) & (df_fb_fut['date_obj'] <= we)
                    overall_future_guests += df_fb_fut.loc[w_mask, '數量'].sum()
            temp_cursor = sun + dt_timedelta(days=1)
            
        # 確保 target_cpg 變數存在 (從上方戰情室繼承，若無則預設 150)
        cpg_val = target_cpg if 'target_cpg' in locals() else 150
        
        # 逐週生成
        fw_week_plans = []
        fw_cursor = today_dt2
        fw_seen = set()
        total_fut_guests_in_weeks = 0
        
        while fw_cursor <= month_end_dt2:
            mon = fw_cursor - dt_timedelta(days=fw_cursor.weekday())
            sun = mon + dt_timedelta(days=6)
            ws = max(mon, dt_date(selected_date.year, selected_date.month, 1))
            we = min(sun, month_end_dt2)
            if ws not in fw_seen and we >= today_dt2:
                fw_seen.add(ws)
                wdates = [(ws + dt_timedelta(days=i)).strftime('%Y-%m-%d')
                          for i in range((we - ws).days + 1)]
                has_d = any(d in fw_dual_dates for d in wdates)
                days_cnt = len(wdates)
                
                raw_week_guests = 0
                raw_week_bf = 0
                raw_week_af = 0
                week_guests = 0
                week_bf = 0
                week_af = 0
                if not df_fb_fut.empty and '數量' in df_fb_fut.columns:
                    w_mask = (df_fb_fut['date_obj'] >= ws) & (df_fb_fut['date_obj'] <= we)
                    raw_week_guests = df_fb_fut.loc[w_mask, '數量'].sum()
                    # 區分早餐 vs 下午茶（依 f&b_data 的「餐別」或「類型」欄位）
                    _meal_col = next((c for c in df_fb_fut.columns if any(k in c for k in ['餐別', '類型', '服務', 'meal', 'type'])), None)
                    if _meal_col:
                        _bf_kw = ['早', 'breakfast', 'bf']
                        _af_kw = ['下午', 'afternoon', 'tea', 'af']
                        _bf_mask = w_mask & df_fb_fut[_meal_col].astype(str).str.lower().str.contains('|'.join(_bf_kw), na=False)
                        _af_mask = w_mask & df_fb_fut[_meal_col].astype(str).str.lower().str.contains('|'.join(_af_kw), na=False)
                        raw_week_bf = df_fb_fut.loc[_bf_mask, '數量'].sum()
                        raw_week_af = df_fb_fut.loc[_af_mask, '數量'].sum()
                    else:
                        # 無餐別欄位：用歷史比例估算
                        _w_bf_r = _bf_ratio if '_bf_ratio' in locals() else 0.93
                        _w_af_r = _af_ratio if '_af_ratio' in locals() else 0.07
                        raw_week_bf = int(raw_week_guests * _w_bf_r)
                        raw_week_af = int(raw_week_guests * _w_af_r)
                    
                    # 讀取上方戰情室的折算係數 (如果有)
                    _w_bf_rate = bf_conv_rate / 100.0 if 'bf_conv_rate' in locals() else 1.0
                    _w_af_rate = af_conv_rate / 100.0 if 'af_conv_rate' in locals() else 1.0
                    
                    week_bf = int(raw_week_bf * _w_bf_rate)
                    week_af = int(raw_week_af * _w_af_rate)
                    week_guests = week_bf + week_af
                
                total_fut_guests_in_weeks += week_guests
                
                # 1. 計算「理論需求 (Standard Need)」
                standard_need = int((cpg_val * 1.15 if has_d else cpg_val) * week_guests)
                
                # 2. 計算「戰情室分配限額 (Adjusted Limit)」
                allocated_limit = 0
                # 使用從 tab_p 傳過來的已折算 future_guests 作為總未來分母
                _overall_fut = future_guests if 'future_guests' in locals() else overall_future_guests
                if _overall_fut > 0:
                    guest_ratio = week_guests / _overall_fut
                    allocated_limit = int(rem_budget_val * guest_ratio)
                
                fw_week_plans.append({
                    'label': f"{ws.strftime('%m/%d')} ～ {we.strftime('%m/%d')}",
                    'has_dual': has_d,
                    'standard_need': standard_need,
                    'allocated_limit': allocated_limit,
                    'week_guests': week_guests,
                    'raw_week_bf': raw_week_bf,
                    'raw_week_af': raw_week_af,
                    'week_bf': week_bf,
                    'week_af': week_af,
                    'dual_labels': [d[5:] for d in wdates if d in fw_dual_dates],
                    'days_cnt': days_cnt,
                })
            fw_cursor = sun + dt_timedelta(days=1)

        if total_fut_guests_in_weeks > 0 and fw_week_plans:
            st.caption(
                f"💡 **雙軌預算系統**：系統自動將「戰情室真實剩餘預算」按各週客數佔比分配給未來的每一週。")
            for wp in fw_week_plans:
                dual_note = f"　🎯 含雙冠日：{', '.join(wp['dual_labels'])}" if wp['has_dual'] else ""
                
                alloc = wp['allocated_limit']
                std = wp['standard_need']
                w_bf = wp.get('week_bf', 0)
                w_af = wp.get('week_af', 0)
                r_bf = wp.get('raw_week_bf', 0)
                r_af = wp.get('raw_week_af', 0)
                
                # 早/午人次說明文字 (加入折算資訊)
                if r_bf + r_af > 0:
                    bf_label = f"預估 {int(r_bf)} → <span style='color:#3498db;'>{int(w_bf)}</span> 人" if r_bf != w_bf else f"{int(w_bf)} 人"
                    af_label = f"預估 {int(r_af)} → <span style='color:#3498db;'>{int(w_af)}</span> 人" if r_af != w_af else f"{int(w_af)} 人"
                    bf_af_label = f"🍳 早餐 {bf_label} ／ 🍰 下午茶 {af_label}"
                else:
                    bf_af_label = f"合計 <strong>{int(wp['week_guests'])}</strong> 人"
                
                # 早/午成本拆分（如果 k 值系統的 cb / ca 在作用域中）
                _cb = cb if 'cb' in locals() and cb > 0 else None
                _ca = ca if 'ca' in locals() and ca > 0 else None
                if _cb and _ca and w_bf + w_af > 0:
                    _bf_cost = int(_cb * w_bf)
                    _af_cost = int(_ca * w_af)
                    cost_hint = f"<span style='font-size:11px; color:#aaa;'>🍳 早餐食材 NT${_bf_cost:,} ／ 🍰 下午茶食材 NT${_af_cost:,}</span>"
                else:
                    cost_hint = ""
                
                if alloc == 0 and std > 0:
                    status_color = '#c0392b'
                    status_text = "🚨 預算已徹底透支"
                elif alloc < std * 0.8:
                    status_color = '#e74c3c'
                    status_text = "🚨 預算嚴重吃緊 (月初超支)"
                elif alloc < std:
                    status_color = '#f39c12'
                    status_text = "⚠️ 預算稍微緊縮 (需採購平價替代品)"
                else:
                    status_color = '#27ae60'
                    status_text = "✅ 預算充裕安全 (有額外結餘可運用)"
                
                c1, c2, c3 = st.columns([1.5, 1, 1])
                c1.markdown(
                    f"**{wp['label']}**（{bf_af_label}）<br>"
                    f"{cost_hint}<br>"
                    f"<span style='font-size:13px;color:{status_color};font-weight:600;'>{status_text}</span>"
                    f"<span style='font-size:12px;color:#888;'>{dual_note}</span>",
                    unsafe_allow_html=True
                )
                
                c2.markdown(
                    f"<div style='background:#f1f2f6; border-left:3px solid #7f8fa6; padding:8px; border-radius:6px; text-align:center;'>"
                    f"<strong style='font-size:14px; color:#2f3640;'>NT$ {std:,}</strong>"
                    f"<br><span style='font-size:11px; color:#7f8fa6;'>單純客量理論需求</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                
                c3.markdown(
                    f"<div style='background:{status_color}15; border-left:3px solid {status_color}; padding:8px; border-radius:6px; text-align:center;'>"
                    f"<strong style='font-size:15px; color:{status_color};'>NT$ {alloc:,}</strong>"
                    f"<br><span style='font-size:11px; color:{status_color};'>戰情室真實配額</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                st.write("") # small spacing
        else:
            st.info("💡 目前 `f&b_data` 中沒有這段時間的預約人數資料，或是預約客數為 0，無法估算週採購預算。")
    else:
        st.info("💡 「週採購建議」僅適用於當月或未來月份。")

# =====================================================
# 🛒 菜價分析 tab_s
# =====================================================
if selected_page == "🛒 菜價分析":
    st.header("🛒 菜價分析")
    sp_df = fetch_supplier_prices()

    if sp_df.empty:
        st.warning(
            "⚠️ 尚未讀取到菜價資料。請確認 Google Sheets 中已建立 `supplier_prices` 分頁，且 `period`、`item_name`、`unit`、`price` 欄位已填寫。")
    else:
        periods_available = sorted(sp_df['period_dt'].unique())
        periods_str = [str(p) for p in periods_available]
        n_periods = len(periods_available)

        st.caption(
            f"📋 目前共有 **{n_periods}** 期菜價資料（{periods_str[0]} ～ {periods_str[-1]}）｜共 {len(sp_df['item_name'].unique())} 個品項")

        # ── 0. 📝 填寫本日請購單 ─────────────────────────
        st.markdown("### 📝 填寫本日請購單")
        st.caption("💡 於此處填寫當日請購數量，總金額將自動匯入戰情室「本月餐廳已花費」中進行預算連動。")
        
        # UI: 日期選擇器與部門選擇器
        import datetime
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            order_date = st.date_input("📅 請購日期 (預設為今日)", value=datetime.date.today())
        with col_d2:
            order_dept = st.selectbox("🏢 請購部門", ["The Peak (餐廳)", "Happy Hour (HH)"])
        
        # 根據請購日期，尋找對應的菜價期數 (尋找小於等於請購日期的最新一期菜價)
        order_date_pd = pd.to_datetime(order_date).date()
        valid_periods = [p for p in periods_available if p <= order_date_pd]
        
        if valid_periods:
            target_period = max(valid_periods)
        else:
            # 如果請購日期比所有菜價表都早，就拿最舊的一期
            target_period = min(periods_available)
            
        df_target = sp_df[sp_df['period_dt'] == target_period].copy()
        # 移除重複的品項名稱，避免 data_editor 迴圈時發生第二筆把第一筆的數量刪除的 Bug
        df_target = df_target.drop_duplicates(subset=['item_name'], keep='last')
        
        st.info(f"ℹ️ 系統已自動判斷並載入 **{target_period}** 的菜價表作為計價基礎。")
        
        # 初始化購物車 (優化 A：分部門儲存，切換部門時自動隔離購物車)
        cart_key = f"purchase_cart_{order_dept}"
        if cart_key not in st.session_state:
            st.session_state[cart_key] = {}
        # 相容舊的 purchase_cart 格式，若存在則遷移
        if 'purchase_cart' in st.session_state and st.session_state.purchase_cart:
            st.session_state[cart_key] = st.session_state.purchase_cart
            del st.session_state['purchase_cart']

        # 獨立的搜尋框，避免遮擋
        search_term = st.text_input("🔍 搜尋品項名稱", "", help="輸入關鍵字即可過濾下方表格")
        
        # 漏洞1修復：只建立一次 df_order_form，搜尋過濾後帶入購物車數量
        df_order_form = df_target[['item_name', 'price', 'unit']].copy()
        if search_term:
            df_order_form = df_order_form[df_order_form['item_name'].str.contains(search_term, na=False, case=False)]
            
        # 帶入購物車數量
        df_order_form['請購數量'] = df_order_form['item_name'].apply(lambda x: st.session_state[cart_key].get(x, 0.0))
        df_order_form = df_order_form.rename(columns={'item_name': '品項名稱', 'price': '當期單價', 'unit': '單位'})
        df_order_form = df_order_form[['品項名稱', '當期單價', '單位', '請購數量']]
        
        st.markdown("請在下方表格直接點擊 **請購數量** 進行填寫：")
        edited_df = st.data_editor(
            df_order_form,
            disabled=['品項名稱', '當期單價', '單位'],
            hide_index=True,
            use_container_width=True,
            key=f"editor_{search_term}", # 當搜尋改變時強制重新渲染，避免狀態衝突
            column_config={
                "請購數量": st.column_config.NumberColumn(
                    "請購數量",
                    min_value=0.0,
                    step=0.1,
                    format="%.1f",
                    help="請填寫大於 0 的數量 (支援小數)"
                )
            }
        )
        
        # 將編輯結果寫回購物車（依部門分開儲存）
        for _, row in edited_df.iterrows():
            item = row['品項名稱']
            qty = row['請購數量']
            if qty > 0:
                st.session_state[cart_key][item] = qty
            elif item in st.session_state[cart_key]:
                del st.session_state[cart_key][item]
                
        # 結算與已選清單（從分部門購物車讀取）
        final_order_list = []
        current_total = 0
        for item, qty in st.session_state[cart_key].items():
            price_series = df_target.loc[df_target['item_name'] == item, 'price']
            if not price_series.empty:
                price = price_series.values[0]
                unit = df_target.loc[df_target['item_name'] == item, 'unit'].values[0]
                cost = price * qty
                current_total += cost
                final_order_list.append({
                    '品項名稱': item,
                    '當期單價': price,
                    '單位': unit,
                    '請購數量': qty,
                    '總價': cost
                })
        
        if final_order_list:
            st.markdown("### 🛒 目前已選購清單")
            st.dataframe(pd.DataFrame(final_order_list)[['品項名稱', '當期單價', '單位', '請購數量', '總價']], hide_index=True, use_container_width=True)
            
        st.markdown(f"**本次請購預估總金額：** <span style='color:#e74c3c; font-size:1.2rem; font-weight:bold;'>NT$ {int(current_total):,}</span>", unsafe_allow_html=True)
        
        submitted = st.button("🚀 確認送出請購單", type="primary")
        if submitted:
            if current_total <= 0:
                st.warning("⚠️ 購物車是空的，請至少填寫一項大於 0 的數量！")
            else:
                final_order_df = pd.DataFrame(final_order_list)
                final_order_df.insert(0, '請購日期', order_date.strftime("%Y-%m-%d"))
                final_order_df['建檔時間'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # 漏洞2修復：用 try/except 包裹，防止 success 未定義導致 NameError
                success = False
                with st.spinner("正在將請購單寫入資料庫..."):
                    try:
                        if "Happy Hour" in order_dept:
                            success = append_4fhh_daily_purchase_report(final_order_df)
                            if success:
                                fetch_4fhh_daily_purchase_report.clear()
                        else:
                            success = append_thepeak_daily_purchase_report(final_order_df)
                            if success:
                                fetch_thepeak_daily_purchase_report.clear()
                    except Exception as e:
                        success = False
                        st.error(f"❌ 寫入資料庫時發生錯誤，請稍後再試。錯誤訊息：{e}")
                        
                    if success:
                        # 漏洞3修復：送出後清空對應部門購物車，並自動重整頁面 (無需手動 Reboot)
                        st.session_state[cart_key] = {}
                        st.success(f"✅ 請購單送出成功！本次共為 {order_dept} 採購 {len(final_order_df)} 項，總計 NT$ {int(current_total):,}。頁面將自動重整以更新預算...")
                        import time; time.sleep(1.5)
                        st.rerun()
        
        st.divider()

        # ── A. 菜商物價指數 ──────────────────────────────
        if n_periods >= 2:
            st.markdown("#### 📈 A. 菜商物價指數")
            st.info(
                "💡 **什麼是大盤指數？** 以第一期（基準期）的整體物價為 100 分。如果本期指數為 105，代表飯店整體的食材採購成本「通膨了 5%」。**這是與供應商談判及調整 CPG 預算的最強客觀依據！**\n\n⚠️ **注意**：本指數採用**等權重平均計算**，每個品項貢獻相同的權重，高波動但採購量小的品項（如香草、松露）可能略微放大指數，請搭配品項趨勢圖一併判斷。")

            index_df = get_market_index_df(sp_df)
            base_period = periods_available[0]

            if not index_df.empty:
                latest_idx = index_df.iloc[-1]['index']
                prev_idx = index_df.iloc[-2]['index']
                diff_idx = latest_idx - prev_idx

                ic1, ic2 = st.columns([1, 3])
                with ic1:
                    st.metric(label="本期大盤指數", value=f"{latest_idx:.1f}",
                              delta=f"{diff_idx:+.1f} 點 (vs上期)", delta_color="inverse")
                    st.caption(f"基準期：{base_period} (=100)")
                with ic2:
                    # 取得折線圖的上下限，並包含 100
                    all_idx_vals = index_df['index'].tolist() + [100]
                    idx_min = max(0, int(min(all_idx_vals) * 0.98))
                    idx_max = int(max(all_idx_vals) * 1.02)

                    line_chart = alt.Chart(index_df).mark_line(point=True, strokeWidth=3, color='#e74c3c').encode(
                        x=alt.X('period_str:O', title='期別',
                                axis=alt.Axis(labelAngle=-30)),
                        y=alt.Y('index:Q', title='指數', scale=alt.Scale(
                            domain=[idx_min, idx_max], zero=False)),
                        tooltip=[
                            alt.Tooltip('period_str:N', title='期別'),
                            alt.Tooltip('index:Q', title='大盤指數', format='.1f'),
                        ]
                    )

                    base_line = alt.Chart(pd.DataFrame({'y': [100]})).mark_rule(
                        strokeDash=[5, 5], color='gray').encode(y='y:Q')
                    st.altair_chart(
                        (base_line + line_chart).properties(height=250), use_container_width=True)
                st.divider()
            else:
                index_df = pd.DataFrame()

        # ── B. 本月食材安全預算範圍 ──────────────────────────────
        st.markdown("#### 💰 B. 本月食材安全預算範圍")
        st.caption("ℹ️ 此預算為本月餐廳**整體**食材預算，涵蓋 The Peak 與 Happy Hour 所有部門的合計總額度。")

        # 抓取本月總預算與客數 (直接沿用採購分析 tab_p 戰情室的計算結果)
        total_guests_for_budget = st.session_state.get('_budget_est_guests', 0)
        cur_budget = st.session_state.get('_budget_total', 0)
        current_month_str = f"{selected_date.year}-{selected_date.month:02d}"

        if total_guests_for_budget == 0 or cur_budget == 0:
            st.info("💡 預算與客數資料未載入。請先點擊上方「📊 採購分析」頁籤以完成資料讀取。")
        elif total_guests_for_budget > 0 and cur_budget > 0:
            bc1, bc2 = st.columns([1, 2])
            with bc1:
                st.metric(f"本月總備餐人次 ({current_month_str})", f"{int(total_guests_for_budget):,} 人", help="包含早餐與下午茶的總預估客數 (與戰情室連動)")
                st.metric("本月食材總預算額度", f"NT$ {int(cur_budget):,}", help="此為根據「採購分析」戰情室設定所計算出的總額度")
                
                # 將 cur_budget 賦值給 total_budget，供右側 bc2 的分配邏輯使用
                total_budget = cur_budget

            with bc2:
                # 根據 latest_idx 決定配額
                if n_periods >= 2:
                    if latest_idx > 105:
                        status = "🚨 市場通膨惡化 (大盤 > 105)"
                        def_pct, norm_pct, risk_pct = 0.70, 0.20, 0.10
                        advice = "強烈建議將 70% 預算鎖定在「高防禦」避風港食材，嚴格限縮高風險食材採購。"
                    elif latest_idx < 95:
                        status = "📉 市場低點 (大盤 < 95)"
                        def_pct, norm_pct, risk_pct = 0.40, 0.30, 0.30
                        advice = "目前處於買方市場，可釋放 30% 預算「抄底」囤貨原本偏貴的高風險食材。"
                    else:
                        status = "⚖️ 市場平穩"
                        def_pct, norm_pct, risk_pct = 0.50, 0.30, 0.20
                        advice = "市場波動正常，維持標準 5:3:2 採購比例。"
                else:
                    status = "—"
                    def_pct, norm_pct, risk_pct = 0.50, 0.30, 0.20
                    advice = "資料不足，維持標準比例。"

                st.markdown(f"**市場判定：{status}**")
                st.caption(f"💡 戰略建議：{advice}")

                # 畫進度條
                st.markdown(f"""
                <div style='display:flex; height:24px; border-radius:12px; overflow:hidden; margin-bottom:10px;'>
                    <div style='width:{def_pct*100}%; background-color:#2ecc71; display:flex; align-items:center; justify-content:center; color:white; font-size:12px; font-weight:bold;'>防禦 {def_pct*100:.0f}%</div>
                    <div style='width:{norm_pct*100}%; background-color:#f1c40f; display:flex; align-items:center; justify-content:center; color:#333; font-size:12px; font-weight:bold;'>一般 {norm_pct*100:.0f}%</div>
                    <div style='width:{risk_pct*100}%; background-color:#e74c3c; display:flex; align-items:center; justify-content:center; color:white; font-size:12px; font-weight:bold;'>風險 {risk_pct*100:.0f}%</div>
                </div>
                """, unsafe_allow_html=True)

                # 配額金額
                col_d, col_n, col_r = st.columns(3)
                col_d.metric("🛡️ 避風港配額", f"${int(total_budget * def_pct):,}")
                col_n.metric("🥬 一般配額", f"${int(total_budget * norm_pct):,}")
                col_r.metric("🚨 高風險配額", f"${int(total_budget * risk_pct):,}")

        else:
            st.info("💡 目前查無本月預估總早餐人數，無法試算總預算。請確認上方月分析已載入當月資料。")

        st.divider()

        # ── D. 本期 vs 上期漲跌排行 ──────────────────────
        if n_periods >= 2:
            st.markdown("#### 📊 D. 本期 vs 上期：漲跌排行")
            # 防呆：若 Section C 因異常未執行導致 full_latest_df 未定義，用基本資料代替
            if 'full_latest_df' not in dir() and 'full_latest_df' not in locals():
                full_latest_df = sp_df[sp_df['period_dt'] == periods_available[-1]].copy() if not sp_df.empty else pd.DataFrame()
            ranked = full_latest_df.dropna(subset=['change']).copy() if 'change' in full_latest_df.columns else pd.DataFrame()
            ranked = ranked.sort_values('change_pct', ascending=False)

            bc1, bc2 = st.columns(2)
            with bc1:
                st.markdown("**🔴 漲幅最大 Top 5**")
                top_up = ranked.head(5)
                for _, r in top_up.iterrows():
                    if r['change_pct'] > 0:
                        st.markdown(
                            f"<div style='background:#fdf2f2; border-left:4px solid #e74c3c; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#e74c3c; font-size:13px;'>▲ +{r['change_pct']:.1f}%</span>"
                            f"<br><span style='font-size:12px; color:#888;'>{r['prev_price']:.0f} → {r['price']:.0f} 元/{r['unit']}</span></div>",
                            unsafe_allow_html=True
                        )
            with bc2:
                st.markdown("**🟢 跌幅最大 Top 5**")
                top_down = ranked.tail(5).iloc[::-1]
                for _, r in top_down.iterrows():
                    if r['change_pct'] < 0:
                        st.markdown(
                            f"<div style='background:#f2fdf5; border-left:4px solid #2ecc71; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#2ecc71; font-size:13px;'>▼ {r['change_pct']:.1f}%</span>"
                            f"<br><span style='font-size:12px; color:#888;'>{r['prev_price']:.0f} → {r['price']:.0f} 元/{r['unit']}</span></div>",
                            unsafe_allow_html=True
                        )
            st.divider()

        # ── E. 品項趨勢圖 ────────────────────────────────
        if n_periods >= 2:
            st.markdown("#### 📈 E. 品項歷期趨勢")
            all_items = sorted(sp_df['item_name'].unique().tolist())
            selected_items = st.multiselect(
                "選擇品項（可多選）",
                options=all_items,
                default=all_items[:3] if len(all_items) >= 3 else all_items,
                placeholder="請選擇要比對的品項"
            )
            if selected_items:
                trend_df = sp_df[sp_df['item_name'].isin(
                    selected_items)].copy()
                trend_df['period_str'] = trend_df['period_dt'].astype(str)
                price_min = int(trend_df['price'].min() * 0.85)
                price_max = int(trend_df['price'].max() * 1.15)
                trend_chart = alt.Chart(trend_df).mark_line(point=True, strokeWidth=2).encode(
                    x=alt.X('period_str:O', title='期別',
                            axis=alt.Axis(labelAngle=-30)),
                    y=alt.Y('price:Q', title='單價', scale=alt.Scale(
                        domain=[price_min, price_max], zero=False)),
                    color=alt.Color(
                        'item_name:N', legend=alt.Legend(title='品項')),
                    tooltip=[
                        alt.Tooltip('period_str:N', title='期別'),
                        alt.Tooltip('item_name:N', title='品項'),
                        alt.Tooltip('price:Q', title='單價', format='.1f'),
                        alt.Tooltip('unit:N', title='單位'),
                    ]
                ).properties(height=380)
                st.altair_chart(trend_chart, use_container_width=True)
            st.divider()
        else:
            st.info("💡 目前只有一期資料，累積下一期菜單後即可查看趨勢圖與漲跌比對。")
            st.divider()

        # ── F. 叫貨戰略建議 ──────────────────────────────
        st.markdown("#### 🎯 F. 叫貨戰略建議")
        if n_periods >= 2:
            # 防呆：若 full_latest_df 未定義，補上基本資料
            if 'full_latest_df' not in dir() and 'full_latest_df' not in locals():
                full_latest_df = sp_df[sp_df['period_dt'] == periods_available[-1]].copy() if not sp_df.empty else pd.DataFrame()
            ranked_all = full_latest_df.dropna(subset=['change_pct']).copy() if 'change_pct' in full_latest_df.columns else pd.DataFrame()
            # 持續漲價：漲幅 > 5%
            alert_up = ranked_all[ranked_all['change_pct'] > 5].sort_values(
                'change_pct', ascending=False)
            # 明顯降價：跌幅 > 5%
            alert_down = ranked_all[ranked_all['change_pct']
                                    < -5].sort_values('change_pct')

            # 漏洞5修復：歷史天價/低點警報需至少 3 期資料，避免剛上線時的假警報
            MIN_PERIODS_FOR_ATH = 3
            alert_all_time_high = ranked_all[
                (ranked_all['price'] >= ranked_all['ytd_max']) &
                (ranked_all['ytd_max'] > ranked_all['ytd_min']) &
                (ranked_all.get('ytd_periods', 0) >= MIN_PERIODS_FOR_ATH)
            ]
            # 歷史低點警報 (同樣需至少 3 期)
            alert_all_time_low = ranked_all[
                (ranked_all['price'] <= ranked_all['ytd_min']) &
                (ranked_all['ytd_max'] > ranked_all['ytd_min']) &
                (ranked_all.get('ytd_periods', 0) >= MIN_PERIODS_FOR_ATH)
            ]

            if not alert_all_time_high.empty:
                high_items = '、'.join(
                    alert_all_time_high['item_name'].head(5).tolist())
                st.error(
                    f"🚨 **【歷史高點警報】**：{high_items} 目前為今年最高價！\n\n👉 強烈建議：全面停用或切換至高防禦(穩定)食材，直到價格回落。")

            if not alert_up.empty:
                up_items = '、'.join(alert_up['item_name'].head(5).tolist())
                st.warning(
                    f"📈 **短期漲幅警示（>{5}%）**：{up_items}\n\n👉 建議：評估替代食材，或提前確認這週用量是否能縮減。")

            if not alert_all_time_low.empty:
                low_items = '、'.join(
                    alert_all_time_low['item_name'].head(5).tolist())
                st.success(
                    f"✅ **【歷史低點進場】**：{low_items} 目前來到今年低價！\n\n👉 建議：可在不超庫存、確保新鮮的前提下多囤貨，鎖住 CPG。")
            elif not alert_down.empty:
                down_items = '、'.join(alert_down['item_name'].head(5).tolist())
                st.success(
                    f"📉 **短期降價機會（>{5}%↓）**：{down_items}\n\n👉 建議：這批相對便宜，可適量多叫。")

            if alert_up.empty and alert_down.empty and alert_all_time_high.empty and alert_all_time_low.empty:
                st.info("✅ 本期菜價整體穩定，無明顯異常波動，按原採購計畫執行即可。")

            # 彙整摘要表
            with st.expander("📋 完整戰略摘要表"):
                summary_rows = []
                for _, r in ranked_all.iterrows():
                    is_ath = (r['price'] >= r['ytd_max']) and (
                        r['ytd_max'] > r['ytd_min']) and (r.get('ytd_periods', 0) >= MIN_PERIODS_FOR_ATH)
                    is_atl = (r['price'] <= r['ytd_min']) and (
                        r['ytd_max'] > r['ytd_min']) and (r.get('ytd_periods', 0) >= MIN_PERIODS_FOR_ATH)

                    if is_ath:
                        strategy = "🚨 歷史高點！強烈建議停用"
                    elif is_atl:
                        strategy = "✅ 歷史低點！建議囤貨"
                    elif r['change_pct'] > 5:
                        strategy = "⚠️ 短期漲價，考慮替代"
                    elif r['change_pct'] < -5:
                        strategy = "📉 短期降價，可增量"
                    else:
                        strategy = "─ 穩定，照常叫貨"

                    ytd_avg = r.get('ytd_avg', 0)
                    if pd.isna(ytd_avg) or ytd_avg == 0:
                        ytd_str = '—'
                    else:
                        diff_pct = ((r['price'] - ytd_avg) / ytd_avg * 100)
                        ytd_str = f"均 {ytd_avg:.1f} ({'+' if diff_pct>0 else ''}{diff_pct:.1f}%)"
                        if is_ath:
                            ytd_str += " [高點]"
                        elif is_atl:
                            ytd_str += " [低點]"

                    summary_rows.append({
                        '品項': r['item_name'],
                        '本期單價': f"{r['price']:.0f} 元/{r['unit']}",
                        '短期漲跌': f"{'+' if r['change_pct']>0 else ''}{r['change_pct']:.1f}%",
                        '純平基準對照': ytd_str,
                        '戰略建議': strategy
                    })
                st.dataframe(pd.DataFrame(summary_rows),
                             use_container_width=True, hide_index=True)
        else:
            st.info("💡 叫貨戰略建議需要至少兩期資料才能比對。下一次菜單收到後，貼到 `supplier_prices` 分頁即可自動產生建議。")

        # ── G. 食材風險防禦分級 (價格波動度分析) ──────────────────
        st.divider()
        st.markdown("#### 🛡️ G. 食材風險防禦分級 (價格波動度分析)")
        if 'ytd_stats' in locals() and not ytd_stats.empty:
            st.info(
                "💡 **波動率 = (歷史最高價 - 歷史最低價) / 歷史最低價**。代表該食材在今年內可能暴漲的最大幅度。\n\n👉 **戰略建議**：主廚在雙冠日應盡量避開右側的高風險食材，多使用左側的高防禦食材來穩定 CPG。")

            vol_df = ytd_stats[ytd_stats['ytd_min'] > 0].copy()
            vol_df['volatility'] = (
                vol_df['ytd_max'] - vol_df['ytd_min']) / vol_df['ytd_min'] * 100

            high_risk = vol_df[vol_df['volatility'] > 50].sort_values(
                'volatility', ascending=False)
            low_risk = vol_df[vol_df['volatility'] <= 20].sort_values(
                'volatility', ascending=True)

            vc1, vc2 = st.columns(2)
            with vc1:
                st.markdown("##### 🛡️ 高防禦避風港 (波動極低 ≤ 20%)")
                if not low_risk.empty:
                    for _, r in low_risk.head(10).iterrows():
                        st.markdown(
                            f"<div style='background:#f2fdf5; border-left:4px solid #2ecc71; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#888; font-size:12px;'>最大漲幅: <span style='color:#2ecc71;'>{r['volatility']:.0f}%</span></span>"
                            f"<br><span style='font-size:12px; color:#666;'>區間: {r['ytd_min']:.0f} ~ {r['ytd_max']:.0f} 元/{r['unit']} (均 {r['ytd_avg']:.0f})</span></div>",
                            unsafe_allow_html=True
                        )
                else:
                    st.write("目前尚無低波動食材")

            with vc2:
                st.markdown("##### 🚨 高風險地雷區 (波動劇烈 > 50%)")
                if not high_risk.empty:
                    for _, r in high_risk.head(10).iterrows():
                        st.markdown(
                            f"<div style='background:#fdf2f2; border-left:4px solid #e74c3c; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#888; font-size:12px;'>最大漲幅: <span style='color:#e74c3c;'>{r['volatility']:.0f}%</span></span>"
                            f"<br><span style='font-size:12px; color:#666;'>區間: {r['ytd_min']:.0f} ~ {r['ytd_max']:.0f} 元/{r['unit']} (均 {r['ytd_avg']:.0f})</span></div>",
                            unsafe_allow_html=True
                        )
                else:
                    st.write("目前尚無高波動食材")
        else:
            st.info("💡 需要至少兩期資料才能分析食材的價格波動度。")

        # ── H. 雙冠備戰行事曆 (Peak Demand Radar) ──────────────────
        st.divider()
        if 'm_curr' in locals() or 'm_curr' in globals():
            s_curr_metrics = calc_key_metrics(m_curr)
            if s_curr_metrics.get('dual_match_dates'):
                st.markdown("#### 🎯 H. 雙冠備戰行事曆 (Peak Demand Radar)")
                st.info(
                    "💡 這是系統自動揪出本月符合「高營收」且「高均價」的雙冠日。**請現場採購人員特別注意這幾天的菜價與備料！** 隔天早餐高峰期人數預估如下，可考慮採用單價較低或正在降價的替代葉菜，來維持高品質又守住 CPG 目標。")

                import datetime
                # 載入 F&B 資料來預估備餐人數
                # 過去的實際人數從 f&b_report 取，未來的預估人數從 f&b_data 取
                today_d = datetime.date.today()
                df_hist = get_combined_fb_daily_df(selected_date.year, selected_date.month, current_hotel)
                df_fut = get_combined_fb_future_data()
                
                weekdays = ['一', '二', '三', '四', '五', '六', '日']

                s_radar_cols = st.columns(min(max(len(s_curr_metrics['dual_match_dates']), 1), 5))
                for i, d_date in enumerate(s_curr_metrics['dual_match_dates']):
                    # 隔天日期
                    curr_dt = datetime.datetime.strptime(d_date, '%Y-%m-%d')
                    next_dt = curr_dt + datetime.timedelta(days=1)
                    next_day = next_dt.strftime('%Y-%m-%d')
                    
                    wd1 = weekdays[curr_dt.weekday()]
                    wd2 = weekdays[next_dt.weekday()]

                    bf_count = 0
                    # 如果 next_dt 在過去，查 df_hist；如果 next_dt 在未來，查 df_fut
                    if next_dt.date() < today_d:
                        if not df_hist.empty and 'date' in df_hist.columns:
                            row = df_hist[df_hist['date'] == next_day]
                            if not row.empty and 'bf_act' in row.columns:
                                bf_count = pd.to_numeric(row['bf_act'], errors='coerce').fillna(0).iloc[0]
                    else:
                        # 優化 C：未來日期使用預估欄位 bf_est，若不存在才退回 bf_act
                        if not df_fut.empty and 'date' in df_fut.columns:
                            row = df_fut[df_fut['date'] == next_day]
                            if not row.empty:
                                est_col = 'bf_est' if 'bf_est' in row.columns else ('bf_act' if 'bf_act' in row.columns else None)
                                if est_col:
                                    bf_count = pd.to_numeric(row[est_col], errors='coerce').sum() # 可能有多列，全部加總

                    c = s_radar_cols[i % 5]
                    c.markdown(f"""
                    <div style="background: #fff; border: 2px solid #e74c3c; border-radius: 8px; padding: 15px; text-align: center; margin-bottom: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                        <p style="margin:0; font-size:12px; color:#aaa; letter-spacing:0.5px;">🏨 雙冠入住日</p>
                        <p style="margin:4px 0; font-size:14px; color:#666;">{d_date[5:]} (星期{wd1})</p>
                        <hr style="border:0; border-top:1px dashed #eee; margin:8px 0;">
                        <p style="margin:0; font-size:12px; color:#e74c3c; font-weight:bold; letter-spacing:0.5px;">🥐 備餐日（早餐高峰）</p>
                        <h3 style="margin:4px 0; color:#e74c3c; font-size:24px; font-weight:900;">{next_day[5:]} (星期{wd2})</h3>
                        <p style="margin:4px 0 0 0; font-size:14px; color:#333;">預估備餐: <strong style="color:#e74c3c; font-size:16px;">{int(bf_count)}</strong> 人</p>
                    </div>
                    """, unsafe_allow_html=True)

if current_hotel != "採購":
    if selected_page == "👥 人事概況":
        st.header("👥 人事概況")

        # -- 人事管理函數 (Google Sheets 版) --
        def get_all_employees():
            try:
                df = get_all_employees_cached()
                return df if df is not None else pd.DataFrame()
            except:
                return pd.DataFrame()

        def add_employee(e_id, name, dept, pos, salary):
            try:
                df = conn.read(worksheet="employees", ttl="0")
                required_cols = ["employee_id", "name",
                                 "dept", "position", "salary"]

                if df is None or df.empty or not all(c in df.columns for c in required_cols):
                    if df is None or df.empty:
                        df = pd.DataFrame(columns=required_cols)
                    else:
                        for c in required_cols:
                            if c not in df.columns:
                                df[c] = ""

                if str(e_id) in df['employee_id'].astype(str).values:
                    return "ID_EXISTS"

                new_emp = pd.DataFrame([{"employee_id": str(
                    e_id), "name": name, "dept": dept, "position": pos, "salary": salary}])
                df = pd.concat([df, new_emp], ignore_index=True)
                conn.update(worksheet="employees", data=df.fillna(""))
                get_all_employees_cached.clear()
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
                get_all_employees_cached.clear()
            except:
                pass

        # -- UI: 新增員工區 --
        with st.expander("➕ 新增新進員工資訊", expanded=False):
            with st.form("add_employee_form", clear_on_submit=True):
                col1, col2 = st.columns(2)
                new_id = col1.text_input("員工編號 (必填)")
                new_name = col2.text_input("姓名 (必填)")

                new_dept = st.selectbox(
                    "所屬部門", ["路徒Plus行旅站前館", "櫃檯", "房務", "工務", "The Peak"])
                new_pos = st.text_input("職位")
                new_salary = st.number_input("薪資", min_value=0, step=1000)

                submit_btn = st.form_submit_button("✅ 確認新增")
                if submit_btn:
                    if not new_id or not new_name:
                        st.error("❌ 請填寫員工編號與姓名！")
                    else:
                        res = add_employee(
                            new_id, new_name, new_dept, new_pos, new_salary)
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
            df_emp['employee_id'] = df_emp['employee_id'].str.replace(
                r'\.0$', '', regex=True)
            df_emp = df_emp[df_emp['employee_id'] != '']
            df_emp = df_emp[df_emp['employee_id'].str.lower() != 'nan']

        if df_emp.empty:
            st.info("💡 目前資料庫中尚無員工資訊。")
        else:
            # 計算總薪資 (排除職位為 PT 的人)
            # 確保 position 欄位存在且處理大小寫
            if 'position' in df_emp.columns:
                non_pt_df = df_emp[df_emp['position'].fillna(
                    '').astype(str).str.upper() != 'PT']
                total_salary = pd.to_numeric(
                    non_pt_df['salary'], errors='coerce').fillna(0).sum()
            else:
                total_salary = pd.to_numeric(
                    df_emp['salary'], errors='coerce').fillna(0).sum()

            st.markdown(f"""
            <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; border-left: 5px solid #4CAF50; margin-bottom: 20px;">
                <p style="margin: 0; font-size: 14px; color: #666;">💰 正職員工薪資總計</p>
                <h2 style="margin: 0; color: #2e437c;">NT$ {int(total_salary):,}</h2>
                <p style="margin: 5px 0 0 0; font-size: 12px; color: #999;">* 已自動排除職位名稱為 "PT" 的人員數據</p>
            </div>
            """, unsafe_allow_html=True)

            col_sort, col_search = st.columns([1, 1])
            sort_opt = col_sort.selectbox(
                "排序方式", ["員工編號順序", "薪資 (由高到低)", "薪資 (由低到高)", "按部門排序"])
            search_query = col_search.text_input("🔍 搜尋姓名或編號")

            # 搜尋過濾
            if search_query:
                df_emp = df_emp[df_emp['name'].astype(str).str.contains(
                    search_query, case=False) | df_emp['employee_id'].str.contains(search_query, case=False)]

            # 確保 salary 為數值以便排序
            if 'salary' in df_emp.columns:
                df_emp['salary'] = pd.to_numeric(
                    df_emp['salary'], errors='coerce').fillna(0)

            # 排序邏輯
            if sort_opt == "員工編號順序":
                df_emp = df_emp.sort_values("employee_id")
            elif sort_opt == "薪資 (由高到低)":
                if 'salary' in df_emp.columns:
                    df_emp = df_emp.sort_values("salary", ascending=False)
            elif sort_opt == "薪資 (由低到高)":
                if 'salary' in df_emp.columns:
                    df_emp = df_emp.sort_values("salary", ascending=True)
            elif sort_opt == "按部門排序":
                if 'dept' in df_emp.columns:
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
                row_cols[0].write(row.get('employee_id', ''))
                row_cols[1].write(row.get('name', ''))
                row_cols[2].write(row.get('dept', ''))
                row_cols[3].write(row.get('position', ''))

                salary_val = row.get('salary', 0)
                try:
                    salary_int = int(float(salary_val))
                except:
                    salary_int = 0
                row_cols[4].write(f"NT$ {salary_int:,}")

                # 使用 idx 來保證 key 絕對唯一，避免 StreamlitDuplicateElementKey
                if row_cols[5].button("🗑️", key=f"del_{idx}_{row.get('employee_id', '')}", help="刪除此員工"):
                    delete_employee(row.get('employee_id', ''))
                    st.toast(f"已刪除員工: {row['name']}")
                    time.sleep(0.5)
                    st.rerun()



def clean_nation_name(nation_str):
    if not isinstance(nation_str, str):
        return str(nation_str)
    
    import re
    # 只保留中文字元
    cleaned = re.sub(r'[^\u4e00-\u9fa5]', '', nation_str)
    
    # 若全部非中文，退回原字串並去空白
    if not cleaned:
        cleaned = str(nation_str).strip()
        
    # 常見金旭與觀光署的國籍名稱別名對應
    synonyms = {
        '澳大利亞': '澳洲',
        '中國大陸': '大陸',
        '中國': '大陸',
        '南韓': '韓國',
        '大韓民國': '韓國',
        '紐西蘭': '紐西蘭', # 預防有時寫新西蘭
        '新西蘭': '紐西蘭',
        '阿拉伯聯合大公國': '阿聯'
    }
    
    return synonyms.get(cleaned, cleaned)

def parse_tourism_bureau_excel(uploaded_file):
    try:
        # 改用無 header 讀取，避免合併儲存格造成欄位錯亂
        df = pd.read_excel(uploaded_file, header=None)
        
        # 找出國籍與數據開始的那一行
        start_row = 0
        for i, row in df.iterrows():
            row_str = ' '.join([str(x) for x in row.values]).lower()
            if 'nationality' in row_str or '國籍' in row_str or 'growth rate' in row_str or '成長率' in row_str:
                start_row = i
                break
                
        data_df = df.iloc[start_row+1:].copy()
        
        # 動態尋找真正的「國籍」欄位（包含日本、韓國等字眼的欄位）
        nat_col_idx = -1
        for col in data_df.columns:
            # 抓取該欄位前 15 筆非空值拼成字串來判斷
            col_str = ' '.join([str(x) for x in data_df[col].dropna().head(15).values])
            if '日本' in col_str or 'japan' in col_str.lower() or '韓國' in col_str or 'korea' in col_str.lower():
                nat_col_idx = col
                break
                
        if nat_col_idx != -1:
            # 國籍欄位右邊的欄位，依序過濾出包含數字的實體數據欄位
            numeric_cols = []
            for col in data_df.columns:
                if col > nat_col_idx:
                    # 檢查這欄是否有很多數字 (排除空白欄位)
                    num_count = pd.to_numeric(data_df[col], errors='coerce').notna().sum()
                    if num_count > 5:  # 超過5個數字就認為是有效數據欄
                        numeric_cols.append(col)
            
            # 通常至少需要 3 個數據欄位 (當期, 同期, 成長率)
            if len(numeric_cols) >= 3:
                curr_col = numeric_cols[0]
                prev_col = numeric_cols[1]
                growth_col = numeric_cols[2]
                
                res_df = data_df[[nat_col_idx, curr_col, prev_col, growth_col]].copy()
                res_df.columns = ['Nation_Raw', 'Curr_Arrivals', 'Prev_Arrivals', 'Growth_Rate_Pct']
                
                # 清理資料：移除空值與小計/總計列
                res_df = res_df.dropna(subset=['Nation_Raw'])
                res_df = res_df[~res_df['Nation_Raw'].astype(str).str.contains('Table|Total|計|小計|Growth|Nationality|Jan|洲地區', na=False, case=False)]
                
                # 整理國籍名稱以便比對
                res_df['Nation_Clean'] = res_df['Nation_Raw'].apply(clean_nation_name)
                
                # 確保數值正確
                res_df['Curr_Arrivals'] = pd.to_numeric(res_df['Curr_Arrivals'], errors='coerce').fillna(0)
                res_df['Growth_Rate_Pct'] = pd.to_numeric(res_df['Growth_Rate_Pct'], errors='coerce').fillna(0)
                
                # 過濾掉 Nation_Clean 為空
                res_df = res_df[res_df['Nation_Clean'].str.len() > 0]
                
                return res_df
            else:
                st.error("找不到足夠的數據欄位 (當期、同期、成長率)。")
                return None
        else:
            st.error("在檔案中找不到包含國名的欄位。")
            return None
    except Exception as e:
        st.error(f"解析觀光署報表失敗: {e}")
        return None

def format_pct(val):
    if pd.isna(val) or val == 0:
        return "0%"
    if val < 0.5:
        if val >= 0.01:
            return f"{val:.2f}%"
        else:
            return f"{val:.3f}%"
    return f"{int(val + 0.5)}%"

def clean_channel_name(name):
    name = str(name).strip().upper()
    if 'AGODA' in name: return 'AGODA'
    if 'CTRIP' in name or 'TRIP.COM' in name: return 'CTRIP'
    if 'BOOKING' in name: return 'BOOKING'
    if 'EXPEDIA' in name: return 'EXPEDIA'
    if 'RAKUTEN' in name: return 'RAKUTEN'
    if '官網' in name: return '官網'
    if '小鹿文娛' in name or 'FG' in name: return '小鹿文娛'
    if '【' in name and '】' in name:
        return name.split('】')[0].replace('【', '')
    return name

def render_channel_tab():
    st.header("📉 渠道分析專區")
    
    # -----------------------------------
    # 新增：營收統計區塊
    # -----------------------------------
    st.subheader("💰 營收統計")
    st.markdown("呈現近三個月的住宿、休息與其他營收綜合指標。")
    
    curr_date = st.session_state.get('sidebar_date')
    if curr_date:
        import datetime
        from dateutil.relativedelta import relativedelta
        
        dates = [
            curr_date,
            curr_date - relativedelta(months=1),
            curr_date - relativedelta(months=2)
        ]
        
        stats = []
        for d in dates:
            y, m = d.year, d.month
            msum = fetch_month_summary(y, m)
            other_rev = get_other_revenue(f"{y}{m:02d}")
            stats.append({
                'month_label': f"{m}月",
                'occ': msum['avg_occ'],
                'adr': msum['avg_adr'],
                'revpar': msum['revpar'],
                'room_rev': msum['rev'],
                'other_rev': other_rev,
                'year': y,
                'month': m
            })
            
        with st.expander("📝 填寫 / 編輯休息數據 (點擊展開)", expanded=False):
            st.caption("填寫完成後，點擊「🔄 更新表格」，下方表格才會更新（不會再每個欄位觸發一次）。")
            with st.form("form_rest_stats"):
                cols = st.columns(3)
                rest_inputs_temp = []
                for i, s in enumerate(stats):
                    with cols[i]:
                        st.markdown(f"**{s['month_label']} 休息數據**")
                        r_occ = st.number_input("使用率 (%)", value=0.00, step=0.01, format="%.2f", key=f"r_occ_{s['year']}_{s['month']}")
                        r_adr = st.number_input("ADR (元)", value=0, step=10, key=f"r_adr_{s['year']}_{s['month']}")
                        rest_inputs_temp.append({'occ': r_occ, 'adr': r_adr})
                if st.form_submit_button("🔄 更新表格", type="primary", use_container_width=True):
                    st.session_state['_rest_inputs_submitted'] = rest_inputs_temp

        # 從 session_state 讀取已提交的數值（只有按「更新表格」後才會 rerun）
        rest_inputs = st.session_state.get('_rest_inputs_submitted', [{'occ': 0, 'adr': 0}] * len(stats))
        if len(rest_inputs) < len(stats):
            rest_inputs = rest_inputs + [{'occ': 0, 'adr': 0}] * (len(stats) - len(rest_inputs))

        table_html = f"""
<div style="display: flex; justify-content: flex-start;">
<table style="width: auto; text-align: center; border-collapse: collapse; font-family: sans-serif; white-space: nowrap;">
<tr style="background-color: #f7a037; color: white; font-weight: bold;">
<td rowspan="2" style="border: 1px solid white; padding: 8px 15px;">{curr_date.year}年</td>
<td colspan="2" style="border: 1px solid white; padding: 8px 15px;">住宿</td>
<td colspan="2" style="border: 1px solid white; padding: 8px 15px;">休息</td>
<td rowspan="2" style="border: 1px solid white; padding: 8px 15px;">RevPAR<br>(元)</td>
<td rowspan="2" style="border: 1px solid white; padding: 8px 15px;">房費<br>總額</td>
<td rowspan="2" style="border: 1px solid white; padding: 8px 15px;">其他<br>收入</td>
<td rowspan="2" style="border: 1px solid white; padding: 8px 15px;">營業總額</td>
</tr>
<tr style="background-color: #fce2c4; color: black; font-weight: bold; font-size: 0.9em;">
<td style="border: 1px solid white; padding: 8px 15px;">訂房率(%)</td>
<td style="border: 1px solid white; padding: 8px 15px;">ADR<br>(元)</td>
<td style="border: 1px solid white; padding: 8px 15px;">使用率(%)</td>
<td style="border: 1px solid white; padding: 8px 15px;">ADR<br>(元)</td>
</tr>
"""
        
        for i, s in enumerate(stats):
            bg_color = "#fffaf0" if i % 2 != 0 else "#faece6"
            month_color = "red" if i == 0 else "black"
            font_weight = "bold" if i == 0 else "normal"
            
            r_in = rest_inputs[i]
            rest_occ_str = f"{r_in['occ']:g}" if r_in['occ'] > 0 else ""
            rest_adr_str = f"${int(r_in['adr']):,}" if r_in['adr'] > 0 else ""
            
            total_room_rev = s['room_rev']
            other_rev = s['other_rev']
            total_revenue = total_room_rev + other_rev
            
            table_html += f"""
<tr style="background-color: {bg_color}; font-size: 1.05em;">
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color}; font-weight: {font_weight};">{s['month_label']}</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">{s['occ']:.1f}%</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">${int(s['adr']):,}</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">{rest_occ_str}{'%' if rest_occ_str else ''}</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">{rest_adr_str}</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">${int(s['revpar']):,}</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">${int(total_room_rev):,}</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">${int(other_rev):,}</td>
<td style="border: 1px solid white; padding: 12px 15px; color: {month_color};">${int(total_revenue):,}</td>
</tr>
"""
            
        table_html += "</table></div><br>"
        st.markdown(table_html, unsafe_allow_html=True)
        st.divider()

    st.subheader("📊 訂房渠道分析")
    st.markdown("分析本飯店各訂房渠道之訂單分佈與佔比。")
    
    with st.spinner("載入渠道資料中..."):
        try:
            df_raw = read_google_sheet("marketing_channel_data")
            if df_raw is not None and not df_raw.empty:
                # 清理與統一欄位名稱
                df_raw.columns = [str(c).strip().lower() for c in df_raw.columns]

                if 'date' in df_raw.columns:
                    # 處理合併儲存格 (向前填補 date)
                    df_raw['date'] = df_raw['date'].replace('', pd.NA).ffill()
                
                if set(['date', 'company name', 'rooms']).issubset(set(df_raw.columns)):
                    # 將日期格式統一轉換為 YYYY-MM-DD
                    df_raw = standardize_df_dates(df_raw)
                    
                    curr_date = st.session_state.get('sidebar_date')
                    df_agg = pd.DataFrame()
                    
                    if curr_date:
                        curr_ym2 = curr_date.strftime("%Y-%m")
                        mask = df_raw['date'].astype(str).str.startswith(curr_ym2, na=False)
                        df_t = df_raw[mask].copy()
                        
                        if not df_t.empty:
                            df_t['rooms'] = df_t['rooms'].astype(str).str.replace(',', '', regex=False)
                            df_t['rooms'] = pd.to_numeric(df_t['rooms'], errors='coerce').fillna(0)
                            df_t = df_t[df_t['rooms'] > 0]
                            
                            if not df_t.empty:
                                df_t['channel_clean'] = df_t['company name'].apply(clean_channel_name)
                                df_agg = df_t.groupby('channel_clean', as_index=False).agg({'rooms': 'sum'})
                                
                                total_rooms = df_agg['rooms'].sum()
                                df_agg['rooms_pct'] = (df_agg['rooms'] / total_rooms * 100).round(3)
                                
                        st.info(f"📅 目前顯示飯店數據區間：**{curr_date.strftime('%Y 年 %m 月')}** (依據左側選單)")
        except Exception as e:
            st.error(f"讀取 RTS_backup(marketing_channel_data) 發生錯誤: {e}")
            
    if 'df_agg' not in locals() or df_agg.empty:
        st.warning("⚠️ 尚無資料")
        return
        
    st.divider()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏆 主力渠道分佈")
        
        curr_total = int(df_agg['rooms'].sum()) if not df_agg.empty else 0
        curr_m_str = curr_date.month if curr_date else ""
        
        def render_channel_table(df_a, bg_header, bg_row1, bg_row2):
            if df_a.empty:
                return "<p style='text-align:center; padding: 20px;'>無資料</p>"
            
            top10 = df_a.sort_values('rooms', ascending=False).head(10)
            rows_html = ""
            for i, row in top10.reset_index(drop=True).iterrows():
                bg_c = bg_row1 if (i % 2 == 0) else bg_row2
                pct_str = format_pct(row['rooms_pct'])
                rows_html += f"""<tr style="background-color: {bg_c};">
<td style="padding: 12px; border: 1px solid white;">{row['channel_clean']}</td>
<td style="padding: 12px; border: 1px solid white;">{int(row['rooms'])}</td>
<td style="padding: 12px; border: 1px solid white;">{pct_str}</td>
</tr>"""
            
            # 加入合計
            total_sum = int(df_a['rooms'].sum())
            rows_html += f"""<tr style="background-color: {bg_row1 if (len(top10) % 2 == 0) else bg_row2}; font-weight: bold;">
<td style="padding: 12px; border: 1px solid white;">合計</td>
<td style="padding: 12px; border: 1px solid white;">{total_sum}</td>
<td style="padding: 12px; border: 1px solid white;"></td>
</tr>"""

            return f"""<table style="width: 100%; text-align: center; border-collapse: collapse; font-size: 16px;">
<tr style="background-color: {bg_header}; color: black; font-weight: bold;">
<th style="padding: 12px; border: 1px solid white; text-align: center;">住宿</th>
<th style="padding: 12px; border: 1px solid white; text-align: center;">筆數</th>
<th style="padding: 12px; border: 1px solid white; text-align: center;">百分比</th>
</tr>
{rows_html}
</table>"""

        curr_html = render_channel_table(df_agg, "#f7a63b", "#fcebda", "#fef5ef")
        
        html_layout = f"""<div style="display: flex; gap: 15px; width: 100%;">
<div style="flex: 1;">
{curr_html}
</div>
</div>"""
        st.markdown(html_layout, unsafe_allow_html=True)

def render_nationality_tab():
    st.header("🌍 國籍客源分析專區")
    st.markdown("分析本飯店各國籍旅客分佈，並可結合交通部觀光署大盤數據進行交叉比對。")
    
    # 1. 讀取目前館別的 nationality_report（依右上角選單）
    df_hotel = None
    df_agg = None
    with st.spinner("載入飯店客源資料中..."):
        try:
            import re as _re

            def _extract_yyyymm(v):
                """從各種格式中解析出 YYYYMM 字串，否則回傳空字串"""
                s = str(v).strip().replace('.0', '')
                if not s or s in ('nan', 'None', 'NaT', ''):
                    return ''
                if _re.match(r'^\d{6}$', s):
                    return s
                m2 = _re.match(r'^(\d{4})[-/](\d{1,2})', s)
                if m2:
                    return f"{m2.group(1)}{int(m2.group(2)):02d}"
                if _re.match(r'^\d{8}$', s):
                    return s[:6]
                return ''

            df_hotel_raw = read_google_sheet("nationality_report")

            if df_hotel_raw is not None and not df_hotel_raw.empty:
                df_hotel_raw = df_hotel_raw.copy()
                df_hotel_raw.columns = [str(c).strip().lower() for c in df_hotel_raw.columns]
                if 'date' in df_hotel_raw.columns and 'month' not in df_hotel_raw.columns:
                    date_col = df_hotel_raw['date'].apply(_extract_yyyymm)
                    date_col = date_col.replace('', pd.NA).ffill().fillna('')
                    df_hotel_raw['month'] = date_col.astype(str).str.strip()
                df_hotel_raw = df_hotel_raw.dropna(subset=['nation'])
                df_hotel_raw = df_hotel_raw[df_hotel_raw['nation'].astype(str).str.strip() != '']

            if df_hotel_raw is not None and not df_hotel_raw.empty:
                df_hotel = df_hotel_raw.copy()
                
                # 確保必要欄位存在
                if set(['month', 'nation', 'person', 'rate', 'nights']).issubset(set(df_hotel.columns)):
                    curr_date = st.session_state.get('sidebar_date')
                    df_agg = pd.DataFrame()
                    df_prev_agg = pd.DataFrame()
                    prev_date = None
                    
                    if curr_date:
                        import datetime
                        prev_date = curr_date.replace(day=1) - datetime.timedelta(days=1)
                        
                        def get_month_agg(target_date):
                            ym1 = target_date.strftime("%Y%m")
                            ym2 = target_date.strftime("%Y-%m")
                            ym3 = target_date.strftime("%Y/%m")
                            
                            mask = df_hotel['month'].astype(str).str.contains(f"{ym1}|{ym2}|{ym3}", na=False, regex=True)
                            df_t = df_hotel[mask].copy()
                            
                            if df_t.empty:
                                return pd.DataFrame()
                                
                            for c in ['person', 'rate', 'nights']:
                                df_t[c] = df_t[c].astype(str).str.replace(',', '', regex=False)
                                df_t[c] = pd.to_numeric(df_t[c], errors='coerce').fillna(0)
                            
                            df_t = df_t[df_t['person'] > 0]
                            df_t = df_t[~df_t['nation'].astype(str).str.contains('合計|總計|Total', na=False, case=False)]
                            
                            if not df_t.empty:
                                df_t['nation_clean'] = df_t['nation'].apply(clean_nation_name)
                                # 先把國籍代碼（3碼英文）也提取出來，作為第二層合併依據
                                import re as _re2
                                df_t['nation_code'] = df_t['nation'].apply(
                                    lambda v: (_re2.match(r'^([A-Z]{2,4})', str(v)) or [None, str(v)])[1][:4].strip()
                                )
                                # 以 nation_code（如 HKG）為主鍵分組，避免同一國籍因撇號等格式差異被分成兩列
                                d_agg = df_t.groupby(['nation_code', 'nation_clean'], as_index=False).agg({
                                    'nights': 'sum', 'person': 'sum', 'rate': 'sum'
                                })
                                # 用 nation_code + nation_clean 組合成顯示名稱
                                d_agg['nation'] = d_agg['nation_code'] + d_agg['nation_clean']
                                d_agg['adr'] = d_agg.apply(lambda r: round(r['rate'] / r['nights']) if r['nights'] > 0 else 0, axis=1)
                                total_p = d_agg['person'].sum()
                                d_agg['nights_pct'] = (d_agg['person'] / total_p * 100).round(3) if total_p > 0 else 0
                                return d_agg
                            return pd.DataFrame()
                            
                        df_agg = get_month_agg(curr_date)
                        df_prev_agg = get_month_agg(prev_date)
                        
                        st.info(f"📅 目前顯示飯店數據區間：**{curr_date.strftime('%Y 年 %m 月')}** (依據左側選單)")
        except Exception as e:
            st.error(f"讀取 nationality_report（兩館）發生錯誤: {e}")
    
    if df_agg is None or df_agg.empty:
        st.warning("⚠️ 尚無資料")
        return
        
    st.divider()
    
    # 2. 上傳大盤資料
    st.subheader("⚔️ 大盤 vs 飯店實際業績交叉分析")
    st.markdown("您可以將從觀光署下載的 `月表1-3(來臺旅客_按國籍分析).xlsx` 直接拖曳至下方，系統將自動進行比對。")
    
    uploaded_file = st.file_uploader("上傳觀光署國籍分析報表 (Excel)", type=['xlsx', 'xls'])
    
    df_bureau = None
    if uploaded_file is not None:
        df_bureau = parse_tourism_bureau_excel(uploaded_file)
        if df_bureau is not None and not df_bureau.empty:
            st.success("✅ 觀光署資料解析成功！")
            
            # 合併資料 (Cross-Analysis)
            df_merged = pd.merge(df_agg, df_bureau, left_on='nation_clean', right_on='Nation_Clean', how='inner')
            
            if not df_merged.empty:
                st.markdown("#### 🚀 潛力與流失市場警報矩陣")
                st.markdown("X 軸為 **觀光署大盤成長率 (%)**，Y 軸為 **本飯店該國籍房晚佔比 (%)**。")
                
                # 建立散佈圖
                scatter = alt.Chart(df_merged).mark_circle(size=200).encode(
                    x=alt.X('Growth_Rate_Pct:Q', title='觀光署大盤成長率 (%)'),
                    y=alt.Y('nights_pct:Q', title='飯店房晚佔比 (%)'),
                    color=alt.Color('nation_clean:N', legend=None),
                    tooltip=['nation_clean', 'Growth_Rate_Pct', 'nights_pct', 'nights', 'rate']
                ).properties(height=400)
                
                # 加上文字標籤
                text = scatter.mark_text(
                    align='left', baseline='middle', dx=10, fontSize=12
                ).encode(text='nation_clean:N')
                
                # 加上平均十字線
                rule_x = alt.Chart(df_merged).mark_rule(color='red', strokeDash=[5,5]).encode(x='mean(Growth_Rate_Pct):Q')
                rule_y = alt.Chart(df_merged).mark_rule(color='blue', strokeDash=[5,5]).encode(y='mean(nights_pct):Q')
                
                st.altair_chart(scatter + text + rule_x + rule_y, use_container_width=True)
                
                # 警報區間提示
                high_growth = df_merged[df_merged['Growth_Rate_Pct'] > df_merged['Growth_Rate_Pct'].mean()]
                missed_market = high_growth[high_growth['nights_pct'] < df_merged['nights_pct'].mean()]
                
                if not missed_market.empty:
                    st.warning("⚠️ **流失商機警告 (大盤高成長，但飯店佔比低)：**\n" + 
                               ", ".join([f"{r['nation_clean']} (大盤+{r['Growth_Rate_Pct']}%)" for _, r in missed_market.iterrows()]))
            else:
                st.info("無法將飯店國籍與觀光署資料配對，請確認國籍名稱格式。")
    
    st.divider()
    
    # 3. 飯店自身數據視覺化
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏆 主力客源分佈 (前五名與上月比較)")
        
        curr_total = int(df_agg['person'].sum()) if not df_agg.empty else 0
        prev_total = int(df_prev_agg['person'].sum()) if not df_prev_agg.empty else 0
        
        curr_date = st.session_state.get('sidebar_date')
        curr_m_str = curr_date.month if curr_date else ""
        
        import datetime
        prev_date = curr_date.replace(day=1) - datetime.timedelta(days=1) if curr_date else None
        prev_m_str = prev_date.month if prev_date else ""
        
        def render_top5_table(df_a, bg_header, bg_row1, bg_row2):
            if df_a.empty:
                return "<p style='text-align:center; padding: 20px;'>無資料</p>"
            
            top5 = df_a.sort_values('person', ascending=False).head(5)
            rows_html = ""
            for i, row in top5.reset_index(drop=True).iterrows():
                bg_c = bg_row1 if (i % 2 == 0) else bg_row2
                pct_str = format_pct(row['nights_pct'])
                rows_html += f"""<tr style="background-color: {bg_c};">
<td style="padding: 12px; border: 1px solid white;">{row['nation']}</td>
<td style="padding: 12px; border: 1px solid white;">{int(row['person'])}</td>
<td style="padding: 12px; border: 1px solid white;">{pct_str}</td>
</tr>"""
                
            return f"""<table style="width: 100%; text-align: center; border-collapse: collapse; font-size: 16px;">
<tr style="background-color: {bg_header}; color: black; font-weight: bold;">
<th style="padding: 12px; border: 1px solid white; text-align: center;">國籍</th>
<th style="padding: 12px; border: 1px solid white; text-align: center;">人數</th>
<th style="padding: 12px; border: 1px solid white; text-align: center;">百分比</th>
</tr>
{rows_html}
</table>"""
            
        curr_html = render_top5_table(df_agg, "#f7a63b", "#fcebda", "#fef5ef")
        prev_html = render_top5_table(df_prev_agg, "#08969e", "#d2e2e6", "#eef2f5")
        
        html_layout = f"""<div style="display: flex; gap: 15px; width: 100%;">
<div style="flex: 1;">
<div style="font-size: 20px; font-weight: bold; color: red; margin-bottom: 10px;">{curr_m_str}月 總間數 {curr_total}</div>
{curr_html}
</div>
<div style="flex: 1;">
<div style="font-size: 20px; font-weight: bold; color: red; margin-bottom: 10px;">{prev_m_str}月 總間數 {prev_total}</div>
{prev_html}
</div>
</div>"""
        st.markdown(html_layout, unsafe_allow_html=True)
        
    with col2:
        st.subheader("💰 國籍別營收與 ADR")
        
        # 取 Top 15 依營收排序
        df_bar = df_agg.sort_values('rate', ascending=False).head(15)
        
        # Base chart
        base = alt.Chart(df_bar).encode(x=alt.X('nation_clean:N', sort='-y', title='國籍'))
        
        # Bar chart for revenue
        bar = base.mark_bar(color='#3498db').encode(
            y=alt.Y('rate:Q', title='總營收 (NTD)'),
            tooltip=['nation', 'rate', 'nights']
        )
        
        # Line chart for ADR
        line = base.mark_line(color='#e74c3c', point=True).encode(
            y=alt.Y('adr:Q', title='平均房價 ADR (NTD)'),
            tooltip=['nation', 'adr']
        )
        
        # Layer them
        combo = alt.layer(bar, line).resolve_scale(y='independent').properties(height=350)
        st.altair_chart(combo, use_container_width=True)
        
    st.divider()
    
    # 4. 資料表
    st.subheader("📊 詳細數據表")
    st.dataframe(
        df_agg[['nation', 'person', 'nights_pct', 'nights', 'rate', 'adr']].sort_values('person', ascending=False).rename(columns={
            'nation': '國籍 (代碼)',
            'person': '總人數',
            'nights_pct': '人數佔比 (%)',
            'nights': '總房晚數',
            'rate': '總營收',
            'adr': '平均房價 (ADR)'
        }),
        use_container_width=True,
        hide_index=True
    )


if current_hotel != "採購":
    if selected_page == "🌍 國籍分析":
        render_nationality_tab()
    
    if selected_page == "📉 渠道分析":
        render_channel_tab()



def render_report_tab():
    import datetime
    from dateutil.relativedelta import relativedelta
    import altair as alt
    
    st.header("📋 營運檢討總結報告")
    st.markdown("本頁面整合本月份全館核心數據，提供您快速匯報與重點回顧使用。")
    
    curr_date = st.session_state.get('sidebar_date')
    if not curr_date:
        st.warning("請先在側邊欄選擇日期！")
        return
        
    year, month = curr_date.year, curr_date.month
    month_str = f"{year}-{month:02d}"
    last_month_date = curr_date - relativedelta(months=1)
    last_year_date = curr_date - relativedelta(years=1)
    
    with st.spinner("🚀 正在為您彙整本月全館運營數據..."):
        # Fetching Data
        # 1. Summary KPIs
        curr_summary = fetch_month_summary(year, month)
        ly_summary = fetch_month_summary(last_year_date.year, last_year_date.month)
        lm_summary = fetch_month_summary(last_month_date.year, last_month_date.month)
        y_adr, y_pure_adr = fetch_yearly_metrics(year)
        
        # 2. Budget / CPG
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        start_of_month = f"{year}-{month:02d}-01"
        end_of_month = f"{year}-{month:02d}-{last_day:02d}"
        
        # Calculate actual peak_spent and hist_guests with robust merging
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        start_of_month = f"{year}-{month:02d}-01"
        end_of_month = f"{year}-{month:02d}-{last_day:02d}"
        
        # 1. Get hist_guests (dynamically merge with f&b_report for safety)
        hist_guests = 0
        df_occ = curr_summary.get('df')
        df_fb_daily = get_combined_fb_daily_df(year, month, current_hotel)
        
        if df_occ is not None and not df_occ.empty:
            df_merged = df_occ.copy()
            if not df_fb_daily.empty:
                df_merged = df_merged.merge(df_fb_daily, on='date', how='left')
        elif not df_fb_daily.empty:
            df_merged = df_fb_daily.copy()
        else:
            df_merged = pd.DataFrame()
            
        if not df_merged.empty:
            for c in ['peak_guests', 'bf_act', 'af_act', 'bf_total_act', 'af_total_act', 'rest_day_guests']:
                if c in df_merged.columns:
                    df_merged[c] = pd.to_numeric(df_merged[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                else:
                    df_merged[c] = 0
            
            for _, row in df_merged.iterrows():
                # Prefer peak_guests or rest_day_guests, otherwise bf + af
                if row['peak_guests'] > 0:
                    hist_guests += row['peak_guests']
                elif row['rest_day_guests'] > 0:
                    hist_guests += row['rest_day_guests']
                else:
                    bf = row['bf_total_act'] if row['bf_total_act'] > 0 else row['bf_act']
                    af = row['af_total_act'] if row['af_total_act'] > 0 else row['af_act']
                    hist_guests += (bf + af)

        # 2. Get peak_spent (from purchase_data + daily report)
        df_purchase = get_purchase_data_cached()
        peak_spent = 0
        
        # Debug diagnostics vars
        diag_date_col = None
        diag_dept_col = None
        diag_total_col = None
        diag_purchase_empty = True
        diag_df_t_empty = True
        diag_raw_dates_preview = []
        diag_parsed_dates_preview = []
        
        if df_purchase is not None and not df_purchase.empty:
            diag_purchase_empty = False
            df_purchase = df_purchase.copy()
            df_purchase.columns = df_purchase.columns.astype(str).str.strip()
            date_col = next((c for c in df_purchase.columns if '日期' in c or 'Date' in c), None)
            dept_col = next((c for c in df_purchase.columns if '部門' in c or 'Dept' in c or '單位' in c or '工地' in c), None)
            total_col = next((c for c in df_purchase.columns if '小計' in c or '總計' in c or '金額' in c or 'Total' in c), None)
            
            diag_date_col = date_col
            diag_dept_col = dept_col
            diag_total_col = total_col
            
            if date_col and dept_col and total_col:
                def robust_date_parse(val):
                    if pd.isna(val):
                        return None
                    s = str(val).strip().replace('.0', '')
                    if not s or s in ('nan', 'None', 'NaT'):
                        return None
                    if '/' in s:
                        res = minguo_to_western(s)
                        if res: return res
                    import re as _re_dp
                    if _re_dp.match(r'^\d{6}$', s):
                        try: return pd.to_datetime(s, format='%Y%m').date()
                        except: pass
                    if _re_dp.match(r'^\d{8}$', s):
                        try: return pd.to_datetime(s, format='%Y%m%d').date()
                        except: pass
                    try: return pd.to_datetime(val).date()
                    except: return None
                
                df_purchase['日期'] = df_purchase[date_col].apply(robust_date_parse)
                
                diag_raw_dates_preview = list(df_purchase[date_col].head(5))
                diag_parsed_dates_preview = list(df_purchase['日期'].head(5).astype(str))
                
                # Retrieve all year-month pairs in the official data to prevent double counting
                official_ym = set(pd.to_datetime(df_purchase['日期']).dt.strftime('%Y-%m').dropna().unique())
                
                # Combine daily report if applicable
                df_daily_report = fetch_thepeak_daily_purchase_report()
                if not df_daily_report.empty and '總價' in df_daily_report.columns:
                    append_df = pd.DataFrame()
                    if '請購日期' in df_daily_report.columns:
                        append_df['日期'] = pd.to_datetime(df_daily_report['請購日期'], errors='coerce').dt.date
                    append_df[dept_col] = "The Peak"
                    append_df[total_col] = pd.to_numeric(df_daily_report['總價'], errors='coerce').fillna(0)
                    
                    # Prevent Double Counting
                    append_df['ym'] = pd.to_datetime(append_df['日期']).dt.strftime('%Y-%m')
                    append_df = append_df[~append_df['ym'].isin(official_ym)].drop(columns=['ym'])
                    
                    if not append_df.empty:
                        df_purchase = pd.concat([df_purchase, append_df], ignore_index=True)
                
                # Filter for current month using robustly parsed date
                df_t = df_purchase[pd.to_datetime(df_purchase['日期']).dt.strftime('%Y-%m') == month_str].copy()
                if not df_t.empty:
                    diag_df_t_empty = False
                    df_t['小計'] = pd.to_numeric(df_t[total_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    t_all_depts = df_t[dept_col].astype(str).unique().tolist()
                    t_hh = [d for d in t_all_depts if '4' in d or any(k in d.upper() for k in ['HH', 'HAPPY', '歡樂時光'])]
                    t_peak_depts = [d for d in t_all_depts if any(k in d.upper() for k in ['PEAK', '餐廳', 'THEPEAK', '餐飲']) and d not in t_hh]
                    peak_spent = df_t[df_t[dept_col].isin(t_peak_depts)]['小計'].sum()

        cpg_actual = peak_spent / hist_guests if hist_guests > 0 else 0
        
        # Only show debug info if calculated CPG is 0
        if cpg_actual == 0:
            with st.expander("🔍 CPG 計算除錯診斷資訊 (CPG為0時自動顯示)", expanded=True):
                st.write(f"**計算年月**: {year}-{month:02d}")
                st.write(f"**來客數 (Denominator)**: {hist_guests} (df_occ 空: {df_occ is None or df_occ.empty}, df_fb_daily 空: {df_fb_daily.empty})")
                st.write(f"**採購花費 (Numerator)**: {peak_spent} (df_purchase 空: {diag_purchase_empty})")
                st.write(f"**偵測欄位**: 日期欄={diag_date_col} | 工地/部門欄={diag_dept_col} | 金額欄={diag_total_col}")
                st.write(f"**日期欄前5筆原始資料**: {diag_raw_dates_preview}")
                st.write(f"**日期欄前5筆轉換後資料**: {diag_parsed_dates_preview}")
                if df_purchase is not None and not df_purchase.empty:
                    st.write(f"**總表欄位清單**: {list(df_purchase.columns)}")
                    st.write(f"**當月採購列數**: {0 if diag_df_t_empty else len(df_t)}")
        
        # Only show debug info if calculated CPG is 0
        if cpg_actual == 0:
            with st.expander("🔍 CPG 計算除錯診斷資訊 (CPG為0時自動顯示)", expanded=True):
                st.write(f"**計算年月**: {year}-{month:02d}")
                st.write(f"**來客數 (Denominator)**: {hist_guests} (df_occ 空: {df_occ is None or df_occ.empty}, df_fb_daily 空: {df_fb_daily.empty})")
                st.write(f"**採購花費 (Numerator)**: {peak_spent} (df_purchase 空: {diag_purchase_empty})")
                st.write(f"**偵測欄位**: 日期欄={diag_date_col} | 工地/部門欄={diag_dept_col} | 金額欄={diag_total_col}")
                if df_purchase is not None and not df_purchase.empty:
                    st.write(f"**總表欄位清單**: {list(df_purchase.columns)}")
                    st.write(f"**當月採購列數**: {0 if diag_df_t_empty else len(df_t)}")
        
        # Only show debug info if calculated CPG is 0
        if cpg_actual == 0:
            with st.expander("🔍 CPG 計算除錯診斷資訊 (CPG為0時自動顯示)", expanded=True):
                st.write(f"**計算年月**: {year}-{month:02d}")
                st.write(f"**來客數 (Denominator)**: {hist_guests} (df_occ 空: {df_occ is None or df_occ.empty}, df_fb_daily 空: {df_fb_daily.empty})")
                st.write(f"**採購花費 (Numerator)**: {peak_spent} (df_purchase 空: {diag_purchase_empty})")
                st.write(f"**偵測欄位**: 日期欄={diag_date_col} | 工地/部門欄={diag_dept_col} | 金額欄={diag_total_col}")
                if df_purchase is not None and not df_purchase.empty:
                    st.write(f"**總表欄位清單**: {list(df_purchase.columns)}")
                    st.write(f"**當月採購列數**: {0 if diag_df_t_empty else len(df_t)}")
        cpg_target = st.session_state.get('tab_p_target_cpg', 150)
        
        # 3. Trends for the year
        year_summaries = []
        for m in range(1, month + 1):
            s = fetch_month_summary(year, m)
            if s['df'].empty and m < month:
                continue
            year_summaries.append({
                'month': f"{year}-{m:02d}",
                'month_dt': datetime.date(year, m, 1),
                'occ': s['avg_occ'],
                'adr': s['avg_adr'],
                'revpar': s['revpar']
            })
            
        df_trends = pd.DataFrame(year_summaries)
        
        # 4. Supplier Prices (Market Index)
        sp_df = fetch_supplier_prices()
        idx_df = get_market_index_df(sp_df)
        market_idx_curr = 100
        market_idx_prev = 100
        if not idx_df.empty:
            curr_row = idx_df[idx_df['month_label'] == f"{year}-{month:02d}"]
            if not curr_row.empty:
                market_idx_curr = curr_row.iloc[0]['index']
            prev_row = idx_df[idx_df['month_label'] == f"{last_month_date.year}-{last_month_date.month:02d}"]
            if not prev_row.empty:
                market_idx_prev = prev_row.iloc[0]['index']

        # 5. Channel and Nationality
        top_channels = []
        df_channel = read_google_sheet("marketing_channel_data")
        if df_channel is not None and not df_channel.empty:
            df_channel.columns = [str(c).strip().lower() for c in df_channel.columns]
            if 'date' in df_channel.columns:
                df_channel['date'] = df_channel['date'].replace('', pd.NA).ffill()
                if set(['date', 'company name', 'rooms']).issubset(set(df_channel.columns)):
                    df_channel = standardize_df_dates(df_channel)
                    df_c = df_channel[df_channel['date'].str.startswith(f"{year}-{month:02d}")]
                    if not df_c.empty:
                        df_c['rooms'] = pd.to_numeric(df_c['rooms'].astype(str).str.strip().str.replace(',', ''), errors='coerce').fillna(0)
                        top_channels = df_c.groupby('company name')['rooms'].sum().nlargest(3).index.tolist()

        top_nations = []
        df_nation = _get_cached_sheet_v3("nationality_report", hotel_type=current_hotel)
        if df_nation is not None and not df_nation.empty:
            df_nation.columns = [str(c).strip().lower().replace('\n', '') for c in df_nation.columns]
            date_col = next((c for c in df_nation.columns if '年/月' in c or 'date' in c), None)
            if date_col:
                df_nation[date_col] = df_nation[date_col].replace('', pd.NA).ffill().fillna('')
                mask = df_nation[date_col].astype(str).str.contains(f"{year}.*{month:02d}")
                df_n = df_nation[mask]
                if not df_n.empty:
                    exclude_cols = [date_col, 'total', 'grand total', 'subtotal', '總計', '小計']
                    nation_cols = [c for c in df_n.columns if c not in exclude_cols]
                    sums = {}
                    for c in nation_cols:
                        val = pd.to_numeric(df_n[c].astype(str).str.strip().str.replace(',', ''), errors='coerce').sum()
                        if val > 0: sums[c] = val
                    if sums:
                        top_nations = sorted(sums, key=sums.get, reverse=True)[:3]
                        top_nations = [n.title() for n in top_nations]

        # 6. Daily Logs
        df_logs = _get_cached_sheet_v3("daily_logs", hotel_type=current_hotel)
        log_events = []
        if df_logs is not None and not df_logs.empty and 'date' in df_logs.columns and 'events' in df_logs.columns:
            df_logs = standardize_df_dates(df_logs)
            df_m_logs = df_logs[df_logs['date'].str.startswith(f"{year}-{month:02d}")]
            for _, r in df_m_logs.iterrows():
                ev = str(r['events']).strip()
                if ev and ev.lower() not in ['nan', 'none', '']:
                    log_events.append(f"{str(r['date'])[8:10]}日: {ev}")

    # UI Rendering
    st.markdown("---")
    st.subheader(f"📊 1. 執行摘要 (Executive Summary) - {year}年{month}月")
    
    occ_val = curr_summary['avg_occ']
    occ_ly = ly_summary['avg_occ']
    occ_diff = occ_val - occ_ly
    occ_status = "🟢 優" if occ_val >= 85 else ("🟡 平" if occ_val >= 70 else "🔴 差")
    
    adr_val = curr_summary['avg_adr']
    adr_diff = adr_val - y_adr
    adr_status = "🟢 優" if adr_val >= y_adr else ("🟡 平" if adr_val >= y_adr*0.9 else "🔴 差")
    
    revpar_val = curr_summary['revpar']
    revpar_ly = ly_summary['revpar']
    revpar_diff = revpar_val - revpar_ly
    revpar_status = "🟢 優" if revpar_diff > 0 else "🔴 差"
    
    # removed cpg_actual
    cpg_status = "🟢 佳 (未超標)" if cpg_actual <= cpg_target else "🔴 警告 (超標)"
    
    fh_days = curr_summary['occ90_days']
    fh_status = "🟢 優" if fh_days >= 4 else ("🟡 平" if fh_days >= 1 else "🔴 差")

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("本月 OCC", f"{occ_val:.1f}%", f"{occ_diff:+.1f}% vs 去年", help=f"狀態: {occ_status}")
    with col2:
        st.metric("本月 ADR", f"NT$ {int(adr_val):,}", f"{int(adr_diff):+} vs 年度均價", help=f"狀態: {adr_status}")
    with col3:
        st.metric("本月 RevPAR", f"NT$ {int(revpar_val):,}", f"{int(revpar_diff):+} vs 去年", help=f"狀態: {revpar_status}")
    with col4:
        st.metric("食材 CPG", f"NT$ {int(cpg_actual):,}", f"{int(cpg_actual - cpg_target):+} vs 目標", help=f"狀態: {cpg_status}")
    with col5:
        st.metric("滿房天數 (≥90%)", f"{fh_days} 天", f"狀態: {fh_status}")

    st.markdown("---")
    st.subheader("📈 2. 趨勢分析 (Trend Analysis)")
    
    if not df_trends.empty:
        base = alt.Chart(df_trends).encode(x=alt.X('month_dt:T', title='月份', axis=alt.Axis(format='%m月')))
        line_occ = base.mark_line(color='#3498db', point=True).encode(y=alt.Y('occ:Q', title='OCC (%)', scale=alt.Scale(domain=[0, 100])))
        line_adr = base.mark_line(color='#e74c3c', point=True).encode(y=alt.Y('adr:Q', title='ADR (NT$)'))
        
        st.altair_chart(alt.layer(line_occ, line_adr).resolve_scale(y='independent'), use_container_width=True)
        
        current_quarter = (month - 1) // 3 + 1
        q_months = [current_quarter * 3 - 2, current_quarter * 3 - 1, current_quarter * 3]
        q_df = df_trends[df_trends['month_dt'].apply(lambda x: getattr(x, 'month', -1)).isin(q_months)]
        if not q_df.empty:
            q_df = q_df.sort_values('revpar', ascending=False).reset_index(drop=True)
            rank = q_df[q_df['month'] == month_str].index[0] + 1 if month_str in q_df['month'].values else "?"
            st.info(f"🏆 **季度排名：** 本月 RevPAR 在 Q{current_quarter} 中排名第 **{rank}** / {len(q_df)}。")
            
    st.markdown("---")
    st.subheader("🛒 3. 採購 & 食材成本分析")
    st.write(f"- **食材 CPG：** 實際 `NT$ {int(cpg_actual)}` / 目標 `NT$ {int(cpg_target)}`")
    idx_diff = market_idx_curr - market_idx_prev
    idx_arrow = "📈 增加" if idx_diff > 0 else ("📉 減少" if idx_diff < 0 else "持平")
    st.write(f"- **大盤物價指數：** 本月指數 `{market_idx_curr:.1f}`，較上月 {idx_arrow} `{abs(idx_diff):.1f}`")

    st.markdown("---")
    st.subheader("🌍 4. 來客結構分析")
    colA, colB = st.columns(2)
    with colA:
        st.markdown("**前三大國籍客源**")
        if top_nations:
            for i, n in enumerate(top_nations): st.write(f"{i+1}. {n}")
        else:
            st.write("無資料或尚未載入")
    with colB:
        st.markdown("**前三大訂房渠道**")
        if top_channels:
            for i, c in enumerate(top_channels): st.write(f"{i+1}. {c}")
        else:
            st.write("無資料或尚未載入")
            
    st.markdown("---")
    st.subheader("📝 5. 本月重點營運日誌")
    if log_events:
        with st.expander("展開查看每日日誌重點", expanded=True):
            for e in log_events:
                st.write(f"- {e}")
    else:
        st.write("本月暫無營運日誌紀錄。")

    st.markdown("---")
    st.subheader("🤖 6. 自動檢討評語 (Auto-Generated Summary)")
    
    good_pts = []
    bad_pts = []
    
    if occ_diff > 0: good_pts.append(f"住房率 (OCC) 達 {occ_val:.1f}%，超越去年同期 (+{occ_diff:.1f}%)。")
    else: bad_pts.append(f"住房率 (OCC) 為 {occ_val:.1f}%，較去年同期衰退 ({occ_diff:.1f}%)，需留意集客力。")
    
    if adr_diff >= 0: good_pts.append(f"平均房價 (ADR) 表現亮眼 (NT$ {int(adr_val):,})，高於年度平均基準。")
    else: bad_pts.append(f"平均房價 (ADR) (NT$ {int(adr_val):,}) 低於年度平均基準，可能受到淡季或促銷影響。")
    
    if cpg_actual > cpg_target: bad_pts.append(f"食材成本 (CPG) 超標，實際 NT$ {int(cpg_actual)} 高於目標 NT$ {int(cpg_target)}。")
    else: good_pts.append(f"食材成本 (CPG) 控制得宜，維持在目標 NT$ {int(cpg_target)} 內。")
    
    if idx_diff > 0: bad_pts.append(f"大盤物價指數較上月上漲 {idx_diff:.1f}，推升採購成本壓力。")
    
    summary_text = f"""**【{year}年{month}月 營運總結報告】**

本月營運結果整體呈現 {'成長' if revpar_diff >= 0 else '衰退'} 趨勢，RevPAR 較去年同期 {'增加' if revpar_diff >= 0 else '減少'} NT$ {abs(int(revpar_diff)):,}。

**✅ 本月亮點 (Strengths)**
{chr(10).join(['- ' + p for p in good_pts]) if good_pts else '- 無明顯亮點'}

**⚠️ 待改善與隱患 (Weaknesses & Threats)**
{chr(10).join(['- ' + p for p in bad_pts]) if bad_pts else '- 營運狀況平穩，無明顯隱患'}

**📌 下月建議方向 (Action Plans)**
- **營收策略**：{'加強平日促銷以提升 OCC' if occ_val < 70 else '在維持 OCC 的前提下，逐步拉升 ADR'}。
- **成本控制**：{'物價上漲壓力大，建議檢視前十大高單價食材並尋找替代供應商' if idx_diff > 0 or cpg_actual > cpg_target else '維持現有採購節奏'}。
- **客源開發**：本月主要客群為 {', '.join(top_nations) if top_nations else '未知'}，可針對這些市場推出專屬包裝。
"""
    
    st.markdown(summary_text)
    
    st.markdown("---")
    st.subheader("📑 7. 匯出報告 (Copy / Paste)")
    st.code(summary_text, language='markdown')

if current_hotel != "採購":
    if selected_page == "📋 營運檢討報告":
        render_report_tab()
