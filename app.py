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
    'KR': '[?',
    'SG': '[?',
    'HK': '[皜珠',
    'JP': '[?包',
    'US': '[蝢',
    'TW': '[?財'
}
OTHER_HOLIDAY_COUNTRIES = ['PH', 'MY', 'TH', 'VN']

EVENT_TYPE_LABELS = {
    '瞍??: '[瞍',
    '撅汗': '[撅',
    '鞈賭?': '[鞈稽',
    '?嗡?': '[瘣蒸'
}


@st.cache_data(ttl=600)
def fetch_taipei_events():
    """霈?岫蝞”銝剔??啣??之瘣餃???"""
    try:
        # 霈?岫蝞”銝剔? taipei_events ??
        df = read_google_sheet("taipei_events", ttl="10m")
        if df is not None and not df.empty:
            df = standardize_df_dates(df)
            return df
    except Exception:
        pass
    return pd.DataFrame(columns=['date', 'event_name', 'event_type', 'venue'])


@st.cache_data(ttl=600)
def fetch_supplier_prices():
    """霈???寡” supplier_prices ??嚗??單?皞? DataFrame"""
    try:
        df = read_google_sheet("supplier_prices", ttl="10m")
        if df is None or df.empty:
            return pd.DataFrame()
        # 甈??迂璅???(item name ??item_name)
        df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
        # 敹?甈?瑼Ｘ
        required = {'period', 'item_name', 'unit', 'price'}
        if not required.issubset(set(df.columns)):
            return pd.DataFrame()
        # 皜?鞈?
        df = df.dropna(subset=['item_name', 'price'])
        df['price'] = pd.to_numeric(df['price'], errors='coerce')
        df = df.dropna(subset=['price'])
        df['item_name'] = df['item_name'].astype(str).str.strip()
        df['unit'] = df['unit'].astype(str).str.strip()
        # period ?交?閫?? (?舀 YYYY/M/D, YYYY-MM-DD, Timestamp, float 蝑?蝔格撘?

        def parse_period(v):
            # ?亙歇??date/datetime嚗?亥?
            if isinstance(v, datetime.date):
                return v if not isinstance(v, datetime.datetime) else v.date()
            # ?交 pandas Timestamp
            if isinstance(v, pd.Timestamp):
                return v.date()
            # ?岫 pd.to_datetime (??祉嚗???float 摨???蝔桀?銝脫撘?
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


def get_market_index_df(sp_df):
    """撠?supplier_prices DataFrame 頧??箏之?斤?寞???(Market Price Index) DataFrame"""
    if sp_df is None or sp_df.empty:
        return pd.DataFrame()

    periods_available = sorted(sp_df['period_dt'].unique())
    if len(periods_available) < 2:
        return pd.DataFrame()

    base_period = periods_available[0]
    base_df = sp_df[sp_df['period_dt'] == base_period]
    base_dict = base_df.set_index(['item_name', 'unit'])['price'].to_dict()

    index_data = []
    for p in periods_available:
        curr_df = sp_df[sp_df['period_dt'] == p]
        ratios = []
        for _, r in curr_df.iterrows():
            key = (r['item_name'], r['unit'])
            if key in base_dict and base_dict[key] > 0 and pd.notna(r['price']):
                ratios.append(r['price'] / base_dict[key])

        idx_val = (sum(ratios) / len(ratios) * 100) if ratios else 100
        index_data.append({
            'period_dt': p,
            'period_str': str(p),
            'month_label': p.strftime('%Y-%m'),
            'index': round(idx_val, 1)
        })

    return pd.DataFrame(index_data)


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
    """??嗆???璅?摰嗥????摮: { 'YYYY-MM-DD': {'flags': '...', 'details': [...] } }"""
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
                    day_details.append(f"?? {code}: {h_name}")

        # Sort flags to maintain consistent order
        flags_str = "".join(sorted(list(day_flags)))
        if has_other:
            flags_str += "??"

        if flags_str:
            result[dt_str] = {
                'flags': flags_str,
                'details': day_details
            }

    return result


@st.cache_data(ttl=86400)
def fetch_upcoming_holidays(start_date, days=30):
    """??芯? N 憭拙????""
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
                    day_details.append(f"?? {code}: {h_name}")

        flags_str = "".join(sorted(list(day_flags)))
        if has_other:
            flags_str += "??"

        if flags_str:
            result.append({
                'date': dt_obj.strftime('%Y-%m-%d'),
                'flags': flags_str,
                'details': ", ".join(day_details)
            })
    return result


# 閮剖??鞈?
st.set_page_config(page_title="頝臬?Plus銵?蝡?擗函??隤?, layout="wide")

password_station = st.secrets.get("admin_password", "roaders123")
password_theme = st.secrets.get("theme_password", "theme456")

if "authenticated" not in st.session_state:
    st.markdown("<h2 style='text-align: center;'>?? Welcome to Hotel Master</h2>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd = st.text_input("蝞∠??⊿?蝣?, type="password")
        if pwd:
            if pwd == password_station:
                st.session_state["authenticated"] = True
                st.session_state["hotel_type"] = "蝡?擗?
                st.rerun()
            elif pwd == password_theme:
                st.session_state["authenticated"] = True
                st.session_state["hotel_type"] = "銝駁?擗?
                st.rerun()
            else:
                st.error("??撖Ⅳ?航炊嚗??頛詨")
                st.stop()
        else:
            st.stop()
# -----------------------------

# ???桀?擗典嚗???????蝢抬?
current_hotel = st.session_state.get("hotel_type", "蝡?擗?)


def get_google_sheet_error_hint(e):
    """?寞? Google Sheets ????航炊憿??撠??葉?遣霅?""
    msg = str(e)
    if "invalid_grant" in msg or "Token" in msg or "oauth" in msg.lower():
        return "?? ??撌脤????⊥?嚗???Streamlit Cloud Secrets ?閮剖? Google Service Account??
    if "quota" in msg.lower() or "rate limit" in msg.lower() or "429" in msg:
        return "??API ??撌脰???隢?敺?閰艾?
    if "403" in msg or "forbidden" in msg.lower() or "permission" in msg.lower():
        return "? 瘝?甈?嚗?蝣箄? Google Sheet 撌脰? Service Account Email ?梁蝺刻摩甈???
    if "404" in msg or "not found" in msg.lower():
        return "???曆??啗岫蝞”嚗?蝣箄? Secrets 銝剔? spreadsheet URL ?臬甇?Ⅱ??
    if "Worksheet" in msg and "not found" in msg:
        return "?? ?曆??啗府???迂嚗?蝣箄????迂?澆神?臬甇?Ⅱ??
    return None


# 閰衣?銵?URL嚗???鞈?嚗ardcode 雿 fallback 蝣箔??舫?嚗?
_STATION_SPREADSHEET = st.secrets.get(
    "station_spreadsheet_url",
    "https://docs.google.com/spreadsheets/d/190DAPuSoorfuQzLb1f8E-jAVCnmm6gXC7YrahxCL-VQ/edit"
)
_THEME_SPREADSHEET = st.secrets.get(
    "theme_spreadsheet_url",
    "https://docs.google.com/spreadsheets/d/1zigbiXDK362v8pvkpFxEkLmBR6R4pCNy_qg7CCmcF0I/edit"
)
_ACTIVE_SPREADSHEET = _THEME_SPREADSHEET if current_hotel == "銝駁?擗? else _STATION_SPREADSHEET


class _ConnWrapper:
    """?? GSheetsConnection嚗????read()/update() ?芸?撣?spreadsheet URL"""
    def __init__(self, raw_conn, spreadsheet_url):
        self._raw = raw_conn
        self._url = spreadsheet_url

    def read(self, worksheet=None, spreadsheet=None, **kwargs):
        return self._raw.read(worksheet=worksheet, spreadsheet=spreadsheet or self._url, **kwargs)

    def update(self, worksheet=None, data=None, spreadsheet=None, **kwargs):
        return self._raw.update(worksheet=worksheet, data=data, spreadsheet=spreadsheet or self._url, **kwargs)


try:
    if current_hotel == "銝駁?擗?:
        _raw_conn = st.connection("gsheets_theme", type=GSheetsConnection)
    else:
        _raw_conn = st.connection("gsheets_station", type=GSheetsConnection)
    conn = _ConnWrapper(_raw_conn, _ACTIVE_SPREADSHEET)
except Exception as e:
    hint = get_google_sheet_error_hint(e)
    err_msg = f"?⊥?撱箇? Google Sheets ???: {e}"
    if hint:
        err_msg += f"\n撱箄降: {hint}"
    st.error(err_msg)
    st.stop()


@st.cache_data(ttl=60)
def _get_cached_sheet(worksheet, hotel_type=""):
    """?葉敹怠?撅歹???霈 Sheet 隢?韏圈ㄐ嚗?0s TTL嚗??API 429
    hotel_type ??冽????尹?翰???踹?頝券尹鞈?瘙⊥???""
    return conn.read(worksheet=worksheet, ttl=0)


def read_google_sheet(worksheet, ttl="1m"):
    try:
        return _get_cached_sheet(worksheet, hotel_type=current_hotel)
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"Google Sheet 霈?仃??{worksheet} ({e})"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
        return None


# -- ?箸鞈?摨怨?撖怠??(??芸?摰儔隞乩?撠?摩雿輻) --


def standardize_df_dates(df):
    if df is None or df.empty or 'date' not in df.columns:
        return df

    def fix_d(val):
        if pd.isna(val) or str(val).strip() == '' or str(val).strip() == 'NaT':
            return ""
        v_str = str(val).split(' ')[0].strip()

        import re
        # ??瘞?撟湔?蝪∪神 (靘? 115/4/30 ??115-04-30)
        m_tw = re.match(r'^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})$', v_str)
        if m_tw:
            y, m, d = int(m_tw.group(1)), int(
                m_tw.group(2)), int(m_tw.group(3))
            if y < 1000:
                y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"

        # ???芣????亦??瘜?(靘? 4/30 ??04-30)
        m_md = re.match(r'^(\d{1,2})[/-](\d{1,2})$', v_str)
        if m_md:
            import datetime
            curr_y = datetime.datetime.now().year
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
        # 韏圈?銝剖翰?惜 (60s TTL)嚗?之??rerun ?? API
        df = _get_cached_sheet("daily_data", hotel_type=current_hotel)
        if df is not None and not df.empty:
            # 蝣箔??交?甈??箏?銝脫撘?(YYYY-MM-DD) 隞乩?瘥?
            df = standardize_df_dates(df)
            # 蝣箔??臭?
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
        msg = f"霈??daily_data 憭望?: {e}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
    return {}


def save_daily_data(d_str, data_dict):
    try:
        df = conn.read(worksheet="daily_data", ttl="0")
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
        conn.update(worksheet="daily_data", data=df.fillna(""))
        _get_cached_sheet.clear()
        return True
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"?脣?憭望?: {e}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
        return False


def get_monthly_target(month_str):
    try:
        df = read_google_sheet("targets", ttl="1m")
        if df is not None and not df.empty:
            res = df[df['month'] == month_str]
            if not res.empty:
                return int(res.iloc[0]['target_revenue'])
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"霈??targets 憭望?: {e}"
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
        _get_cached_sheet.clear()
        return True
    except:
        return False


def get_daily_log(d_str):
    try:
        df = read_google_sheet("daily_logs", ttl="1m")
        if df is not None and not df.empty:
            res = df[df['date'] == d_str]
            if not res.empty:
                return str(res.iloc[0]['log']).strip()
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"霈??daily_logs 憭望?: {e}"
        if hint:
            msg += f"\n{hint}"
        st.error(msg)
    # Fallback to daily_data if not found in daily_logs (backward compatibility)
    # ?湔韏啣翰??銝?憭? API
    try:
        df_old = _get_cached_sheet("daily_data", hotel_type=current_hotel)
        if df_old is not None and not df_old.empty:
            df_old = standardize_df_dates(df_old)
            res = df_old[df_old['date'] == d_str]
            if not res.empty and 'daily_work_log' in res.columns:
                return str(res.iloc[0]['daily_work_log']).strip()
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"霈??daily_data (fallback) 憭望?: {e}"
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

        # 蝣箔?甈?摮
        if 'date' not in df.columns or 'log' not in df.columns:
            df = pd.DataFrame(columns=["date", "log"])

        new_row = pd.DataFrame([{'date': d_str, 'log': log_text}])

        if d_str in df['date'].values:
            df = df[df['date'] != d_str]

        df = pd.concat([df, new_row], ignore_index=True)
        if 'date' in df.columns:
            df = df.sort_values('date').reset_index(drop=True)
        conn.update(worksheet="daily_logs", data=df.fillna(""))
        _get_cached_sheet.clear()
        st.toast(f"??{d_str} ?亥?撌脰??朣?Google Sheet嚗?)
        return True
    except Exception as e:
        hint = get_google_sheet_error_hint(e)
        msg = f"?亥??脣?憭望?: {e}"
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
        df_all = _get_cached_sheet("daily_data", hotel_type=current_hotel)
        month_str = f"{year}-{month:02d}"
        df = df_all[df_all['date'].str.startswith(
            month_str, na=False)].sort_values('date')
    except:
        return "--- 霈?仃??---"

    if df.empty:
        return "--- ?嗆??∠???---"

    full_report = ""
    for d in sorted(df['date'].unique()):
        full_report += generate_report_text(d) + "\n\n"
    return full_report


def minguo_to_western(d_str):
    """
    撠?瘞?/????(憒?115/03/02 ??0115/03/02) 頧???Python date 撠情??
    """
    if pd.isna(d_str) or not isinstance(d_str, str):
        return None
    try:
        # 蝘駁???嗡蒂??
        parts = d_str.strip().split('/')
        if len(parts) == 3:
            year = int(parts[0])
            # 憒???115 ??0115嚗??舀??僑
            if year < 1000:  # 瘞?撟渡楊?虜銝之??1000
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
        df_all = _get_cached_sheet("daily_data", hotel_type=current_hotel)
        if df_all is not None and not df_all.empty:
            df_all = standardize_df_dates(df_all)
            # 蝣箔??交??臭?嚗??銴?蝮?
            df_all = df_all.drop_duplicates(subset='date', keep='last')
            df = df_all[(df_all['date'] >= m_start) &
                        (df_all['date'] <= m_end)].copy()
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
        # 蝣箔??詨潭?雿 float
        num_cols = ['revenue', 'total_rooms', 'occ_rate', 'adr']
        for c in num_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(
                    ',', '').str.replace('%', '')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

        for _, r in df.iterrows():
            rev = float(r['revenue'])
            rm = float(r['total_rooms'])
            occ = float(r['occ_rate'])
            adr = float(r['adr'])

            if rev == 0 and adr > 0 and rm > 0:
                rev = adr * rm
            if rm == 0 and rev > 0 and adr > 0:
                rm = rev / adr

            if rm > 0 or rev > 0:
                res['rev'] += rev
                res['rooms'] += rm
                if occ > 0:
                    res['sellable'] += (rm / (occ / 100.0))
                if occ >= 90.0:
                    res['occ90_days'] += 1

        res['avg_occ'] = (res['rooms'] / res['sellable'] *
                          100.0) if res['sellable'] > 0 else 0.0
        res['avg_adr'] = (res['rev'] / res['rooms']
                          ) if res['rooms'] > 0 else 0.0
        res['revpar'] = (res['avg_occ'] / 100.0) * res['avg_adr']

    return res


@st.cache_data(ttl=3600)
def fetch_yearly_metrics(year):
    try:
        df_all = _get_cached_sheet("daily_data", hotel_type=current_hotel)
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


# -- ?湧?甈??脤??交??豢???--
st.sidebar.caption(
    f"?? ?敺?唳??? {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
st.sidebar.header("?? ?交??豢?")
if 'sidebar_date' not in st.session_state:
    st.session_state['sidebar_date'] = datetime.date.today()

# 摰儔甈??? (敹??典摮?頛?賣銋?)
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
    # ???? DB 鞈?雿瘥??箸?
    db_data = get_daily_data(target_d_str)

    # ?郊?詨潭??
    update_dict = {}
    has_changes = False

    for ss_key, (db_col, default_val) in field_mapping.items():
        if ss_key in st.session_state:
            curr_val = st.session_state[ss_key]
            update_dict[db_col] = curr_val

            # 敺?DB 閫????府?瑟見
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

            # ?斗?臬???寡?
            if isinstance(curr_val, float) and isinstance(norm_db, float):
                import math
                if not math.isclose(curr_val, norm_db, abs_tol=1e-5):
                    has_changes = True
                    with open("debug_save.log", "a") as f:
                        f.write(
                            f"[{target_d_str}] DIFF: {db_col} ({type(curr_val)})={curr_val} vs DB {norm_db} ({type(norm_db)})\n")
            elif curr_val != norm_db:
                has_changes = True
                with open("debug_save.log", "a") as f:
                    f.write(
                        f"[{target_d_str}] DIFF: {db_col} ({type(curr_val)})={repr(curr_val)} vs DB {repr(norm_db)} ({type(norm_db)})\n")

    if has_changes:
        save_daily_data(target_d_str, update_dict)

    # ?桃?郊?亥?
    if 'input_daily_log' in st.session_state:
        curr_log = st.session_state['input_daily_log'].strip()
        db_log = str(get_daily_log(target_d_str) or "").strip()
        if curr_log != db_log:
            with open("debug_save.log", "a") as f:
                f.write(
                    f"[{target_d_str}] LOG DIFF: curr={repr(curr_log)} vs db={repr(db_log)}\n")
            save_daily_log(target_d_str, curr_log)


def prev_day():
    st.session_state['sidebar_date'] -= datetime.timedelta(days=1)


def next_day():
    st.session_state['sidebar_date'] += datetime.timedelta(days=1)


col1, col2 = st.sidebar.columns(2)
col1.button("漎? ??憭?, on_click=prev_day)
col2.button("敺?憭??∴?", on_click=next_day)

selected_date = st.sidebar.date_input(
    "?豢??交?", value=st.session_state['sidebar_date'], key='sidebar_date')
date_str = str(selected_date)

# 餈質馱?嗅?甇?蝺刻摩???
if '_actual_current_date' not in st.session_state:
    st.session_state['_actual_current_date'] = date_str
if '_data_is_loaded' not in st.session_state:
    st.session_state['_data_is_loaded'] = False

if st.session_state['_actual_current_date'] != date_str:
    # 蝘駁??⊥?隞嗅???交????瑼??摩嚗?蝝???交???銝?閬?撖怠??蝵株歲??
    st.session_state['_actual_current_date'] = date_str
    st.session_state['_data_is_loaded'] = False

# --- ?啣?嚗望活?汗?豢???---
weekly_options = ["--- ???梢?閬?---",
                  "蝚???(1-7??", "蝚???(8-14??", "蝚???(15-21??", "蝚???(22-28??", "蝚???(29?絲)"]
selected_week = st.sidebar.selectbox(
    "敹恍?勗???", weekly_options, index=0, key="weekly_view_select")
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

    # ?脣??亥?
    st.session_state['input_daily_log'] = get_daily_log(date_str)

    st.session_state['_last_loaded_date'] = date_str
    st.session_state['_last_week_view'] = selected_week
    st.session_state['_data_is_loaded'] = True  # 璅??箏歇頛嚗迨敺遙雿??????閮勗?瑼?


def on_input_change():
    # 雿輻 session_state 銝剔??嗅??交?嚗Ⅱ靽?callback 閫貊??迤蝣?
    target_d = st.session_state.get('_actual_current_date')
    if target_d:
        sync_st_to_db(target_d)


st.sidebar.divider()
st.sidebar.subheader("? ?豢??臬??隞?)


def generate_report_text(d_str):
    data = get_daily_data(d_str)
    if not data:
        return f"--- {d_str} ?∠???---"

    report = []
    report.append(f"========================================")
    report.append(f"? 頝臬?銵? Plus 蝡?擗?- ???亥? ({d_str})")
    report.append(f"========================================\n")

    def safe_int_val(v):
        try:
            if pd.isna(v) or v is None:
                return 0
            return int(float(v))
        except:
            return 0

    report.append(f"?????????)
    report.append(f"- 雿?? {data.get('occ_rate', 0)}%")
    report.append(f"- ADR: NT$ {safe_int_val(data.get('adr', 0)):,}")
    report.append(f"- 蝮賜??? NT$ {safe_int_val(data.get('revenue', 0)):,}")
    report.append(f"- 蝮賭??踵: {safe_int_val(data.get('total_rooms', 0))} ?n")

    report.append(f"???瑹???)
    report.append(f"- 鞎?摰Ｚ迄: {data.get('counter_complaints', '??)}")
    report.append(f"- 瑹隢頃: {safe_int_val(data.get('counter_expense', 0))} ??)
    report.append(f"- 蝮賣?瘨?? {safe_int_val(data.get('cleaned_rooms', 0))} ??)
    report.append(f"- ?踹?隢頃: {safe_int_val(data.get('hk_expense', 0))} ?n")

    report.append(f"??踝? 擗輒?豢? (?拚尹撖阡?靘恥)??)
    report.append(f"- ?拚?蝮質?: {safe_int_val(data.get('bf_total_act', 0))} 鈭?)
    report.append(f"- 銝??嗥蜇閮? {safe_int_val(data.get('af_total_act', 0))} 鈭?)
    report.append(
        f"- Happy Hour: {safe_int_val(data.get('rest_hh_guests', 0))} 鈭?)
    report.append(
        f"- 擗輒?(?冽?): {safe_int_val(data.get('rest_month_rev', 0))} ?n")

    report.append(f"???撌亙?蝝??)
    report.append(f"- 敺耨?踵: {data.get('maint_repair_rooms', 0)} ??)
    report.append(f"- 靽桃?蝝啁?: {data.get('maint_records', '??)}\n")

    report.append(f"???瘥??蝝?敦蝭??)
    report.append(f"{get_daily_log(d_str) or '?∠???摰?}")
    report.append(f"\n" + "-"*40 + "\n")

    return "\n".join(report)


# 1. ?格?臬
single_report = generate_report_text(date_str)
st.sidebar.download_button(
    label="?? ?嗆??蝝???,
    data=single_report,
    file_name=f"Roaders_Plus_Daily_{date_str}.txt",
    mime="text/plain",
    use_container_width=True
)

# 2. ?冽??臬
month_str = selected_date.strftime('%Y-%m')
if f"monthly_report_{month_str}" not in st.session_state:
    if st.sidebar.button(f"?? ?嗆? {month_str} ??蝝???, use_container_width=True):
        # ?臬??韏啣翰??荔?銝?閬撥?嗆???
        df_all = _get_cached_sheet("daily_data", hotel_type=current_hotel)
        df_month = df_all[df_all['date'].str.startswith(
            month_str, na=False)].sort_values('date')

        if df_month.empty:
            st.sidebar.warning(f"?? {month_str} 撠隞颱?鞈???)
        else:
            with st.sidebar.status("甇??Ｙ??梯”...", expanded=False):
                full_month_text = f"?楝敺???Plus 蝡?擗?{month_str} ?冽???蝝?蝮賬n\n"
                for d in df_all['date']:
                    full_month_text += generate_report_text(d) + "\n\n"
                st.session_state[f"monthly_report_{month_str}"] = full_month_text
            st.rerun()
else:
    st.sidebar.download_button(
        label=f"漎? 銝? {month_str} 蝝??(.txt)",
        data=st.session_state[f"monthly_report_{month_str}"],
        file_name=f"Roaders_Plus_Monthly_{month_str}.txt",
        mime="text/plain",
        use_container_width=True,
    )
    if st.sidebar.button("?? ??Ｙ?", key="clear_monthly"):
        del st.session_state[f"monthly_report_{month_str}"]
        st.rerun()

# ?湧?甈??函宏?文?擗?憛?

# -- ?梯”閫???神?亥??澈 --


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
            if '?交?' in row_str:
                header_idx = i
                break

        file.seek(0)
        try:
            df = pd.read_csv(file, skiprows=header_idx) if is_csv else pd.read_excel(
                file, skiprows=header_idx)
        except Exception as e:
            # ?岫銝???engine
            file.seek(0)
            df = pd.read_excel(file, skiprows=header_idx, engine='openpyxl')

        df.columns = df.columns.astype(str).str.replace(
            r'[\s\n\r]', '', regex=True)

        date_col = next((c for c in df.columns if '?交?' in c), None)
        occ_col = next(
            (c for c in df.columns if '雿?? in c or '閮?? in c or '?箇??? in c or 'OCC' in c.upper()), None)
        adr_col = next(
            (c for c in df.columns if '撟喳??踹' in c or 'ADR' in c.upper()), None)

        rev_col = next(
            (c for c in df.columns if '摰Ｘ?嗅' in c or '摰Ｘ?' in c or '蝮賜??? in c or '?平憿? in c or '撖阡??' in c), None)
        rooms_col = next((c for c in df.columns if (
            '雿?? in c or '?箇?' in c or '?桀' in c or '撖虫?' in c) and '?臬' not in c), None)
        if not rooms_col:
            rooms_col = next((c for c in df.columns if (
                '?輸??? in c or '摰Ｘ?? in c) and '?臬' not in c), None)

        if not date_col:
            st.error("?? 閫??憭望?嚗銝???雿?隢炎?亙銵冽撘?)
            return False

        # --- 撘瑕??交?閫???摩 ---
        def robust_parse_date(val):
            if pd.isna(val) or str(val).strip() == '':
                return None
            s = str(val).strip().split('.')[0]  # 蝘駁 .0
            # ?岫 YYYYMMDD
            try:
                if len(s) == 8 and s.isdigit():
                    return pd.to_datetime(s, format='%Y%m%d').date()
            except:
                pass
            # ?岫銝?祈圾??(YYYY-MM-DD, YYYY/MM/DD 蝑?
            try:
                return pd.to_datetime(s).date()
            except:
                pass
            return None

        df['璅??交?'] = df[date_col].apply(robust_parse_date)

        df_new_records = pd.DataFrame()
        updates = []
        for index, row in df.iterrows():
            d_obj = row['璅??交?']
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
            df_existing = _get_cached_sheet("daily_data", hotel_type=current_hotel).copy()
            if df_existing is None:
                df_existing = pd.DataFrame()
            df_existing = standardize_df_dates(df_existing)
            df_new = pd.DataFrame(updates)

            # ?蔥?豢? (隞交? key嚗???
            df_new = df_new.set_index('date')
            if not df_existing.empty:
                df_existing = df_existing.set_index('date')
                df_final = df_new.combine_first(df_existing).reset_index()
            else:
                df_final = df_new.reset_index()

            if 'date' in df_final.columns:
                df_final = df_final.sort_values('date').reset_index(drop=True)

            conn.update(worksheet="daily_data", data=df_final.fillna(""))
            _get_cached_sheet.clear()
            return len(updates)

        return 0
    except Exception as e:
        import traceback
        st.error(f"閫??瑹?梯”憭望?: {e}\n{traceback.format_exc()}")
        return False

# -- 擗輒?梯”閫???神?亥??澈 --


def parse_and_save_restaurant(file, current_year):
    try:
        # 霈??Excel 瑼????摰?
        df = pd.read_excel(file, header=None)

        month_rev = 0
        avg_spent = 0

        # 1. ?寡?????刻”撠??蝞??萄? (銝?靘琿??潛洵 0 甈?
        for i, row in df.iterrows():
            row_str = " ".join([str(v) for v in row if pd.notna(v)])
            # 撠?
            if ('撌脩?蝞??? in row_str or '???? in row_str) and '?拚?' not in row_str and '銝??? not in row_str:
                for val in row:
                    s_val = str(val).strip()
                    if any(c.isdigit() for c in s_val) and not any(k in s_val for k in ['撌脩?蝞???, '????]):
                        try:
                            clean_val = s_val.replace('NT$', '').replace(
                                '$', '').replace(',', '').strip()
                            month_rev = int(float(clean_val))
                            break
                        except:
                            continue
            # 撠摰Ｗ??
            if '撟喳?摰Ｗ?? in row_str or '摰Ｗ?? in row_str:
                for val in row:
                    s_val = str(val).strip()
                    if any(c.isdigit() for c in s_val) and '摰Ｗ?? not in s_val:
                        try:
                            clean_val = s_val.replace('NT$', '').replace(
                                '$', '').replace(',', '').strip()
                            avg_spent = int(float(clean_val))
                            break
                        except:
                            continue

        parsed_days = []
        # 2. 撠瘥?敦 (靽格迤 Regex 霈?游?捆摨?
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
                        # ?? Excel 霈?交??航??摮貊泵????
                        return int(float(str(val).replace(',', '').strip()))
                    except:
                        return 0

                # ?身甈???銝? (?寞? Roaders Plus 撣貊?梯”?澆?)
                # ?拚??賊? (1-6)
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

                # 銝??嗥??(7-12)
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
            # 霈??澈?扯???
            df_existing = _get_cached_sheet("daily_data", hotel_type=current_hotel).copy()
            if df_existing is None:
                df_existing = pd.DataFrame()

            # ??嚗Ⅱ靽???? date 銋摮葡嚗??combine_first ??join ?仃??
            df_existing = standardize_df_dates(df_existing)

            # --- 靽桀儔嚗????圾???蝞??嗆?摰Ｗ?對?撘瑕?湔?暹?鞈?摨思葉閰脫?隞賜??????---
            # ?踹?雿輻??暺?鈭靘??交??Ｙ?鈭葆???????撠 MTD 瘞賊???敺?憭拍????
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

            # ?蔥?豢? (隞交? key嚗???
            df_new = df_new.set_index('date')
            if not df_existing.empty:
                df_existing = df_existing.set_index('date')
                # 隞交銝?????????雿???啗??撩撠?甈???????
                df_final = df_new.combine_first(df_existing).reset_index()
            else:
                df_final = df_new.reset_index()

            if 'date' in df_final.columns:
                df_final = df_final.sort_values('date').reset_index(drop=True)

            # 撖怠?鞈?摨?
            conn.update(worksheet="daily_data", data=df_final.fillna(""))
            _get_cached_sheet.clear()

        # 皜敹怠?隞亦Ⅱ靽??游??賜??唳?豢?
        st.session_state['_last_loaded_date'] = None
        return len(parsed_days)
    except Exception as e:
        import traceback
        st.error(f"閫??擗輒?梯”憭望?: {str(e)}")
        with st.expander("?? ?亦??航炊蝝啁?"):
            st.code(traceback.format_exc())
        with open("debug_error.log", "w") as f:
            f.write(traceback.format_exc())
        return False


# ?璅?
current_hotel = st.session_state.get("hotel_type", "蝡?擗?)

st.title(f"Hotel Master - {current_hotel}")
# 銝餌??
tab1, tab_m, tab6, tab_p, tab_s, tab3, tab4, tab5, tab7 = st.tabs(
    ["?? ??蝮質汗", "?? ?????", "?? 瘥??蝝??, "? ?∟頃??", "?? ???", "?完 ?踹??豢?", "?儭?擗輒?豢?", "? 撌亙??豢?", "? 鈭箔?璁?"])


with tab1:
    st.header("?? ??蝮質汗")

    # 瘜典撠惇 CSS ??Card ?Ｙ???
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
        st.success("?? **皛踵?嗥?嚗??乩??輻?? 90% 隞乩?嚗擗刻??虫?嚗?* ??")

    # -- 隞? --
    st.subheader(f"隞?券尹??憭抒???({date_str})")
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
        <div class="kpi-circle"><div class="kpi-title">隞雿??/div><div class="kpi-value">{occ_val}%</div></div>
        <div class="kpi-circle"><div class="kpi-title">ADR</div><div class="kpi-value">NT$ {safe_format_int(adr_val):,}</div></div>
        <div class="kpi-circle"><div class="kpi-title">蝮賜???/div><div class="kpi-value">NT$ {safe_format_int(rev_val):,}</div></div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.success("?完 **?踹??瘜?*")
        total_occ = st.session_state.get('input_rooms', 0)
        cleaned = st.session_state.get('input_cleaned', 0)
        st.metric("?格?皜?蝮賣 (靘?)", f"{total_occ} ??)
        st.caption(f"??蝝??瘨? {cleaned} ??(撌桅?: {cleaned - total_occ})")
    with col2:
        st.warning("? **撌亙??瘜?*")
        repairs = st.session_state.get('input_repair', 0)
        st.metric("隞敺耨?踵", f"{repairs} ??, delta="? ???" if repairs >
                  0 else "? 甇?虜", delta_color="off")
    with col3:
        st.error("?儭?**擗輒?瘜?*")
        bf_total_act = st.session_state.get('input_bf_total_act', 0)
        st.metric("隞?尹?拚?蝮賭?摰?, f"{safe_format_int(bf_total_act)} 鈭?)

    st.divider()

    # -- ?漲蝝航?璅∪? (MTD Analysis) --
    st.subheader(f"?? ?祆?蝝航??? (MTD: {selected_date.strftime('%Y-%m')})")
    start_of_month = selected_date.replace(day=1).strftime('%Y-%m-%d')

    try:
        df_all = _get_cached_sheet("daily_data", hotel_type=current_hotel)
        if df_all is not None and not df_all.empty:
            df_all = standardize_df_dates(df_all)
            # ?脫迫??鞈?瘥??蝮?
            df_all = df_all.drop_duplicates(subset='date', keep='last')
            df_mtd = df_all[(df_all['date'] >= start_of_month)
                            & (df_all['date'] <= date_str)].copy()
        else:
            df_mtd = pd.DataFrame()
    except Exception as e:
        st.sidebar.error(f"?? 霈????潛??航炊: {e}")
        df_mtd = pd.DataFrame()

    if not df_mtd.empty:
        # ?????質?蝞?甈?頧?詨潘??踹? Google Sheets 撣嗡???銝脣?憿?
        for col in ['bf_theme_act', 'bf_zq_act', 'af_theme_act', 'af_zq_act',
                    'bf_total_act', 'af_total_act', 'bf_total_est', 'af_total_est',
                    'rest_month_rev', 'rest_avg_spent']:
            if col in df_mtd.columns:
                df_mtd[col] = pd.to_numeric(df_mtd[col].astype(
                    str).str.replace(',', ''), errors='coerce').fillna(0)

        mtd_rooms = 0.0
        mtd_rev = 0.0
        total_sellable = 0.0

        for _, r in df_mtd.iterrows():
            # 撘瑕?摮葡皜??脰風
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

            # 摰寥??嚗 Excel ?予蝻箇??嗡???ADR ??賂??撩?踵雿??嚗??詨飛?
            if rev == 0 and adr > 0 and rm > 0:
                rev = adr * rm
            if rm == 0 and rev > 0 and adr > 0:
                rm = rev / adr

            # ?芸?蝮賣?撖阡??平?豢??????芯???0嚗?
            if rm > 0 or rev > 0:
                mtd_rooms += rm
                mtd_rev += rev
                if o > 0:
                    total_sellable += (rm / (o / 100.0))

        mtd_occ = (mtd_rooms / total_sellable *
                   100.0) if total_sellable > 0 else 0.0
        mtd_adr = (mtd_rev / mtd_rooms) if mtd_rooms > 0 else 0.0

        # ?脣?擗輒鞈? (甇?Ⅱ蝯?嚗????蜇)
        rest_mrev = 0
        if not df_mtd.empty and 'rest_month_rev' in df_mtd.columns:
            valid_rest = df_mtd[df_mtd['rest_month_rev'] > 0]
            if not valid_rest.empty:
                rest_mrev = valid_rest.iloc[-1]['rest_month_rev']

        grand_total_rev = mtd_rev + rest_mrev

        # 憿舐內?之??
        st.write("##### ? ?踹??? MTD")
        c1, c2, c3 = st.columns(3)
        c1.markdown(make_card(
            "MTD 蝝航?雿??, f"{mtd_occ:.1f}%", "card-theme-blue", "card-bg-dark", "?"), unsafe_allow_html=True)
        c2.markdown(make_card("MTD 蝝航? ADR", f"NT$ {int(mtd_adr):,}",
                    "card-theme-green", "card-bg-dark", "?"), unsafe_allow_html=True)
        c3.markdown(make_card("MTD ?踹?蝝航??", f"NT$ {int(mtd_rev):,}",
                    "card-theme-orange", "card-bg-dark", "?"), unsafe_allow_html=True)

        st.write("##### ?? ?券尹?蔥? (MTD)")
        g1, g2 = st.columns([1, 2])
        g1.markdown(make_card("擗輒蝯??", f"NT$ {int(rest_mrev):,}",
                    "card-theme-purple", "card-bg-dark", "?儭?), unsafe_allow_html=True)
        g2.markdown(make_card("???券尹 MTD 蝮賜???, f"NT$ {int(grand_total_rev):,}",
                    "card-theme-red", "card-bg-dark", "??"), unsafe_allow_html=True)

        st.markdown(
            "<br><hr style='margin: 5px 0; border: 1px dashed #ddd;'>", unsafe_allow_html=True)
        st.write("##### ?儭?擗輒??蝝航? (MTD)")

        # MTD 擗輒閮?
        mtd_bf_theme = df_mtd['bf_theme_act'].sum(
        ) if 'bf_theme_act' in df_mtd.columns else 0
        mtd_bf_zq = df_mtd['bf_zq_act'].sum(
        ) if 'bf_zq_act' in df_mtd.columns else 0
        mtd_af_theme = df_mtd['af_theme_act'].sum(
        ) if 'af_theme_act' in df_mtd.columns else 0
        mtd_af_zq = df_mtd['af_zq_act'].sum(
        ) if 'af_zq_act' in df_mtd.columns else 0

        # ?祆??湧?蝮賢?
        mtd_total_bf_act = df_mtd['bf_total_act'].sum(
        ) if 'bf_total_act' in df_mtd.columns else 0
        mtd_total_af_act = df_mtd['af_total_act'].sum(
        ) if 'af_total_act' in df_mtd.columns else 0

        # ?箔??渡移蝣綽??閮??摯摰Ｘ????撖阡?摰Ｘ???亙??箏極雿嚗?摰??仿???????冽 0 ?靘予?賂?
        if 'bf_total_act' in df_mtd.columns:
            active_bf_days = len(
                df_mtd[(df_mtd['bf_total_est'] > 0) | (df_mtd['bf_total_act'] > 0)])
        else:
            active_bf_days = 0

        if 'af_total_act' in df_mtd.columns:
            active_af_days = len(
                df_mtd[(df_mtd['af_total_est'] > 0) | (df_mtd['af_total_act'] > 0)])
        else:
            active_af_days = 0

        total_bf_days = active_bf_days if active_bf_days > 0 else 1
        total_af_days = active_af_days if active_af_days > 0 else 1

        mtd_avg_bf = mtd_total_bf_act / total_bf_days
        mtd_avg_af = mtd_total_af_act / total_af_days
        mtd_avg_total = mtd_avg_bf + mtd_avg_af

        # ?脣?擗輒?漲蝮賜?
        # ?寧?敺?蝑??潛?閮?雿蝯??潘??虜瘥?皞Ⅱ (?身?梯”?舐敞閮???)
        rest_month_rev = rest_mrev  # ?撌脰?蝞?
        rest_avg_spent = 0
        if not df_mtd.empty and 'rest_avg_spent' in df_mtd.columns:
            valid_aspent = df_mtd[df_mtd['rest_avg_spent'] > 0]
            if not valid_aspent.empty:
                rest_avg_spent = valid_aspent.iloc[-1]['rest_avg_spent']

        st.markdown(
            "<h6 style='color:#555; margin-top:15px;'>?????尹?TD 蝝航?</h6>", unsafe_allow_html=True)
        sz1, sz2, sz3 = st.columns(3)
        sz1.markdown(make_card(
            "?拚? (撖阡?)", f"{int(mtd_bf_zq)} 鈭?, "card-theme-orange", "", "??"), unsafe_allow_html=True)
        sz2.markdown(make_card(
            "銝???(撖阡?)", f"{int(mtd_af_zq)} 鈭?, "card-theme-purple", "", "?"), unsafe_allow_html=True)
        sz3.markdown(make_card(
            "蝡??? (撖阡?)", f"{int(mtd_bf_zq + mtd_af_zq)} 鈭?, "card-theme-blue", "", "?"), unsafe_allow_html=True)

        st.markdown(
            "<h6 style='color:#555; margin-top:20px;'>???蜓憿尹?TD 蝝航?</h6>", unsafe_allow_html=True)
        st1, st2, st3 = st.columns(3)
        st1.markdown(make_card(
            "?拚? (撖阡?)", f"{int(mtd_bf_theme)} 鈭?, "card-theme-orange", "", "??"), unsafe_allow_html=True)
        st2.markdown(make_card(
            "銝???(撖阡?)", f"{int(mtd_af_theme)} 鈭?, "card-theme-purple", "", "?"), unsafe_allow_html=True)
        st3.markdown(make_card(
            "銝駁??? (撖阡?)", f"{int(mtd_bf_theme + mtd_af_theme)} 鈭?, "card-theme-blue", "", "?"), unsafe_allow_html=True)

        st.markdown(
            "<h6 style='color:#555; margin-top:20px;'>???擗典?雿萇蜇閬賬?/h6>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.markdown(make_card("?拚尹?拚? (撖阡?)", f"{int(mtd_total_bf_act)} 鈭?,
                    "card-theme-orange", "card-bg-dark", "??"), unsafe_allow_html=True)
        m2.markdown(make_card("?拚尹銝???(撖阡?)", f"{int(mtd_total_af_act)} 鈭?,
                    "card-theme-purple", "card-bg-dark", "?"), unsafe_allow_html=True)
        m3.markdown(make_card("?冽?蝯??", f"NT$ {int(rest_month_rev):,}",
                    "card-theme-green", "card-bg-dark", "?"), unsafe_allow_html=True)
        m4.markdown(make_card("撟喳?摰Ｗ??, f"NT$ {int(rest_avg_spent):,}",
                    "card-theme-red", "card-bg-dark", "?屁"), unsafe_allow_html=True)

        st.markdown(
            "<h6 style='color:#555; margin-top:20px;'>???擗冽撟喳?靘恥??/h6>", unsafe_allow_html=True)
        a1, a2, a3 = st.columns(3)
        a1.markdown(make_card("?拚尹?拚?撟喳?", f"{mtd_avg_bf:.1f} 鈭???,
                    "card-theme-orange", "", "??), unsafe_allow_html=True)
        a2.markdown(make_card(
            "?拚尹銝??嗅像??, f"{mtd_avg_af:.1f} 鈭???, "card-theme-purple", "", "??), unsafe_allow_html=True)
        a3.markdown(make_card(
            "?拚尹?湧?蝮賢像??, f"{mtd_avg_total:.1f} 鈭???, "card-theme-blue", "", "??"), unsafe_allow_html=True)

    else:
        st.info("? 鞈?摨思葉?桀?撠??????)

with tab_m:
    st.header("?? ?????")

    # 1. ?????豢? (M-2, M-1, M, M+1)
    prev_prev_m_date = get_month_delta(selected_date, -2)
    prev_m_date = get_month_delta(selected_date, -1)
    next_m_date = get_month_delta(selected_date, 1)

    m_prev_prev = fetch_month_summary(
        prev_prev_m_date.year, prev_prev_m_date.month)
    m_prev = fetch_month_summary(prev_m_date.year, prev_m_date.month)
    m_curr = fetch_month_summary(selected_date.year, selected_date.month)
    m_next = fetch_month_summary(next_m_date.year, next_m_date.month)

    # ???餃僑???豢? (YoY)
    m_curr_ly = fetch_month_summary(
        selected_date.year - 1, selected_date.month)

    st.markdown("#### ?? ?祆?蝮質汗?撟游???頛?(YoY)")
    if not m_curr['df'].empty and not m_curr_ly['df'].empty:
        col1, col2, col3 = st.columns(3)

        adr_diff = m_curr['avg_adr'] - m_curr_ly['avg_adr']
        adr_pct = (adr_diff / m_curr_ly['avg_adr']
                   * 100) if m_curr_ly['avg_adr'] > 0 else 0
        adr_color = "#2ecc71" if adr_diff >= 0 else "#e74c3c"
        adr_sign = "+" if adr_diff >= 0 else ""
        col1.markdown(
            f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {adr_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>?嗆?撟喳? ADR</p><strong style='font-size:22px;'>NT$ {int(m_curr['avg_adr']):,}</strong><p style='margin:5px 0 0 0; font-size:13px; color:{adr_color}; font-weight:bold;'>頛撟游???{adr_sign}NT$ {int(adr_diff):,} ({adr_sign}{adr_pct:.1f}%)</p></div>", unsafe_allow_html=True)

        occ_diff = m_curr['avg_occ'] - m_curr_ly['avg_occ']
        occ_color = "#2ecc71" if occ_diff >= 0 else "#e74c3c"
        occ_sign = "+" if occ_diff >= 0 else ""
        col2.markdown(
            f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {occ_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>?嗆?撟喳? OCC</p><strong style='font-size:22px;'>{m_curr['avg_occ']:.1f}%</strong><p style='margin:5px 0 0 0; font-size:13px; color:{occ_color}; font-weight:bold;'>頛撟游???{occ_sign}{occ_diff:.1f}%</p></div>", unsafe_allow_html=True)

        rev_diff = m_curr['rev'] - m_curr_ly['rev']
        rev_pct = (rev_diff / m_curr_ly['rev']
                   * 100) if m_curr_ly['rev'] > 0 else 0
        rev_color = "#2ecc71" if rev_diff >= 0 else "#e74c3c"
        rev_sign = "+" if rev_diff >= 0 else ""
        col3.markdown(
            f"<div style='background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid {rev_color}; height:100%;'><p style='margin:0; font-size:13px; color:#666;'>?嗆?蝮賢恥?輻???/p><strong style='font-size:22px;'>NT$ {int(m_curr['rev']):,}</strong><p style='margin:5px 0 0 0; font-size:13px; color:{rev_color}; font-weight:bold;'>頛撟游???{rev_sign}NT$ {int(rev_diff):,} ({rev_sign}{rev_pct:.1f}%)</p></div>", unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom:20px;'></div>",
                    unsafe_allow_html=True)
    else:
        if m_curr['df'].empty:
            st.info("? ?祆?撠?豢?嚗瘜??餃僑??瘥???)
        elif m_curr_ly['df'].empty:
            st.info("? ?餃僑??撠甇瑕撠?鞈???)

    # ???啣??之瘣餃?鞈?
    taipei_events_df = fetch_taipei_events()

    # --- A. 瘥雿??瘜?(??撠?) ---
    st.subheader("?? 瘥雿??瘜?頛?(??)")
    col_chart1, col_chart2, col_chart3, col_chart4 = st.columns(4)

    def render_occ_chart(month_data, title_suffix):
        df = month_data['df'].copy()
        if df.empty:
            st.info(f"? {month_data['month_label']} 撠?豢???)
            return

        # ???啣??冽?撟喳? ADR ?箸?蝺?雿??蜓鞈???蝢?典?銝鞈?靘?隞亥圾瘙箏偕摨血?鋆?憿?
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
                       'adr_baseline_text'] = f"${int(avg_adr):,} (??"
            if y_adr > 0:
                df.loc[df.index[-1], 'y_adr_text'] = f"${int(y_adr):,} (撟?"
            if y_pure_adr > 0:
                df.loc[df.index[-1],
                       'y_pure_adr_text'] = f"${int(y_pure_adr):,} (蝝像)"

        dt = pd.to_datetime(df['date'])
        df['day'] = dt.dt.day
        weekday_map = {0: '銝', 1: '鈭?, 2: '銝?, 3: '??, 4: '鈭?, 5: '??, 6: '??}
        df['weekday'] = dt.dt.weekday.map(weekday_map)
        df['label'] = df['day'].astype(str) + " (" + df['weekday'] + ")"

        df['color_category'] = df['occ_rate'].apply(
            lambda x: '>=90' if x >= 90.0 else ('>=80' if x >= 80.0 else '<80'))

        if not df.empty:
            y_str, m_str = df['date'].iloc[0].split('-')[:2]
            holidays_dict = fetch_holidays_for_month(int(y_str), int(m_str))

            # ?蔥???暑??蝐?
            def get_combined_flags_list(d_str):
                import re
                h_f_str = holidays_dict.get(d_str, {}).get('flags', '')
                h_flags = re.findall(r'\[.*?\]|??', h_f_str)

                e_flags = []
                if not taipei_events_df.empty:
                    day_events = taipei_events_df[taipei_events_df['date'] == d_str]
                    for _, row in day_events.iterrows():
                        e_label = EVENT_TYPE_LABELS.get(
                            row['event_type'], '[瘣蒸')
                        if e_label not in e_flags:
                            e_flags.append(e_label)
                return h_flags + e_flags

            # 撱箇?憭惜璅惜鞈? (?憭??5 撅文??游????踹??漲??)
            for i in range(5):
                df[f'flag_{i}'] = df['date'].apply(lambda d: get_combined_flags_list(
                    d)[i] if len(get_combined_flags_list(d)) > i else '')
        else:
            for i in range(5):
                df[f'flag_{i}'] = ''

        # ==========================================
        # 1. 撱箇? OCC 摮? (?瑟???+ 雿?曉?瘥?摮?蝐?+ 瘣餃?/蝭??
        # ==========================================
        base_occ = alt.Chart(df).encode(
            x=alt.X('label:O',
                    title='?交?',
                    sort=df['label'].tolist(),
                    axis=alt.Axis(labelAngle=0)),
            tooltip=['date', 'occ_rate', 'adr']
        )

        bars = base_occ.mark_bar().encode(
            y=alt.Y('occ_rate:Q', title='雿??(%)',
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

        # 雿??摮?蝐?(?芰蝜潭 OCC 頠賂?銝??急頠?
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

        # 撱箇?憭惜?璅惜
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

        # 閮??嗆? ADR ??????蝯曹???Y 頠豢?靘偕 domain嚗???Altair 憭???皞偕摨衣蝡??渡??臭? Bug
        valid_adrs = df[df['adr'] >
                        0]['adr'] if not df.empty else pd.Series([])
        if not valid_adrs.empty:
            adr_min = max(0, int(valid_adrs.min() * 0.9))
            adr_max = int(valid_adrs.max() * 1.1)
        else:
            adr_min = 2000
            adr_max = 8000

        avg_adr = month_data.get('avg_adr', 0)
        y_adr, y_pure_adr = fetch_yearly_metrics(
            int(month_data['month_label'].split('-')[0]))

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
        # 2. 撱箇? ADR 摮? (????+ 鞈?暺?+ 蝝撟喳??踹?箸?蝺?+ 蝝???詨潭?閮?
        # ==========================================
        base_adr = alt.Chart(df).encode(
            x=alt.X('label:O', sort=df['label'].tolist()),  # ?芰??OCC X 頠詨?雿?
            tooltip=['date', 'occ_rate', 'adr']
        )

        adr_line = base_adr.mark_line(color='#ff9f43', strokeWidth=3, interpolate='monotone').encode(
            y=alt.Y('adr:Q', title='撟喳??踹 (NT$)', axis=alt.Axis(
                titleColor='#ff9f43', format='$,.0f'), scale=adr_scale)
        )
        adr_points = base_adr.mark_circle(color='black', size=100, stroke='white', strokeWidth=1.5).encode(
            y=alt.Y('adr:Q', scale=adr_scale)
        )

        adr_layers = [adr_line, adr_points]

        # 蝜芾ˊ?冽?撟喳? ADR 蝝?箸?蝺??喳?詨潭?閮?
        if avg_adr > 0:
            # 撱箇?瘞游像蝝?? (?梁?詨? df 閫?捱撠箏漲?函? bug嚗?????X 蝺函Ⅳ??霅偌撟?
            baseline_rule = alt.Chart(df).mark_rule(
                color='#e74c3c',
                strokeWidth=1.5,
                strokeDash=[5, 5]
            ).encode(
                y=alt.Y('adr_baseline:Q', scale=adr_scale)
            )

            # 撱箇?蝝??璅惜 (?梁?詨? df嚗?冽?敺?憭拍鼓鋆賣?摮?摰?撠?)
            baseline_text = alt.Chart(df).mark_text(
                align='right',     # ?寧?撠?嚗???敺?”?折 (撌血) 撱嗡撓
                baseline='middle',
                dx=-8,             # ?椰?宏 8 ??嚗??憭?僑 ADR ?詨潮???
                color='#000000',
                fontSize=12,
                fontWeight='bold'
            ).encode(
                x=alt.X('label:O', sort=df['label'].tolist()),
                y=alt.Y('adr_baseline:Q', scale=adr_scale),
                text='text:N' if 'text' in df.columns else 'adr_baseline_text:N'
            )
            adr_layers.extend([baseline_rule, baseline_text])

        # 蝜芾ˊ撟?ADR 暺?箸?蝺?
        if df.get('y_adr_baseline', pd.Series()).max() > 0:
            y_adr_rule = alt.Chart(df).mark_rule(color='#f1c40f', strokeWidth=1.5, strokeDash=[
                5, 5]).encode(y=alt.Y('y_adr_baseline:Q', scale=adr_scale))
            y_adr_text = alt.Chart(df).mark_text(
                align='left', baseline='middle', dx=8, dy=-14, color='#000000', fontSize=11, fontWeight='bold'
            ).encode(
                x=alt.X('label:O', sort=df['label'].tolist()), y=alt.Y('y_adr_baseline:Q', scale=adr_scale), text='y_adr_text:N'
            )
            adr_layers.extend([y_adr_rule, y_adr_text])

        # 蝜芾ˊ撟渡?撟單 ADR 暺?箸?蝺?
        if df.get('y_pure_adr_baseline', pd.Series()).max() > 0:
            yp_adr_rule = alt.Chart(df).mark_rule(color='#000000', strokeWidth=1.5, strokeDash=[
                5, 5]).encode(y=alt.Y('y_pure_adr_baseline:Q', scale=adr_scale))
            yp_adr_text = alt.Chart(df).mark_text(align='left', baseline='middle', dx=8, dy=14, color='#000000', fontSize=11, fontWeight='bold').encode(
                x=alt.X('label:O', sort=df['label'].tolist()), y=alt.Y('y_pure_adr_baseline:Q', scale=adr_scale), text='y_pure_adr_text:N'
            )
            adr_layers.extend([yp_adr_rule, yp_adr_text])

        adr_chart = alt.layer(*adr_layers)

        # ==========================================
        # 3. 蝯??拙???摰?? Y 頠貊?函??遘嚗祕?曉?蝢?朣?
        # ==========================================
        chart = alt.layer(occ_chart, adr_chart).resolve_scale(
            y='independent'
        ).properties(title=f"{month_data['month_label']} {title_suffix}", height=400)

        st.altair_chart(chart, use_container_width=True)

    with col_chart1:
        render_occ_chart(m_prev_prev, "(????")
    with col_chart2:
        render_occ_chart(m_prev, "(銝?)")
    with col_chart3:
        render_occ_chart(m_curr, "(?祆?)")
    with col_chart4:
        render_occ_chart(m_next, "(銝?)")

    # --- A2. ?餃僑??頠楚撠? (YoY Daily Comparison) ---
    st.markdown("#### ?? ?餃僑??頠楚撠? (YoY Daily Comparison)")
    if not m_curr['df'].empty and not m_curr_ly['df'].empty:
        df_ty = m_curr['df'].copy()
        df_ly = m_curr_ly['df'].copy()

        if 'day' not in df_ty.columns:
            df_ty['day'] = pd.to_datetime(df_ty['date']).dt.day
        if 'day' not in df_ly.columns:
            df_ly['day'] = pd.to_datetime(df_ly['date']).dt.day

        df_ty['year'] = '隞僑'
        df_ly['year'] = '?餃僑'

        df_yoy = pd.concat([df_ty[['day', 'adr', 'year']],
                           df_ly[['day', 'adr', 'year']]], ignore_index=True)
        df_yoy['adr'] = pd.to_numeric(df_yoy['adr'], errors='coerce').fillna(0)

        # 閮剖? Y 頠豢?靘偕
        yoy_adr_min = max(0, int(df_yoy['adr'].min() * 0.9))
        yoy_adr_max = int(df_yoy['adr'].max() * 1.1)
        if yoy_adr_min == yoy_adr_max:
            yoy_adr_max += 1000

        yoy_chart = alt.Chart(df_yoy).mark_line(point=True, strokeWidth=3).encode(
            x=alt.X('day:O', title='?交? (Day of Month)'),
            y=alt.Y('adr:Q', title='撟喳??踹 (NT$)', scale=alt.Scale(
                domain=[yoy_adr_min, yoy_adr_max], zero=False)),
            color=alt.Color('year:N',
                            scale=alt.Scale(domain=['隞僑', '?餃僑'], range=[
                                            '#ff9f43', '#bdc3c7']),
                            legend=alt.Legend(title="撟港遢", orient="top-left")
                            ),
            strokeDash=alt.condition(
                alt.datum.year == '?餃僑', alt.value([5, 5]), alt.value([0])),
            tooltip=['day', 'year', 'adr']
        ).properties(height=350)

        st.altair_chart(yoy_chart, use_container_width=True)

    st.markdown("<div style='margin-bottom:30px;'></div>",
                unsafe_allow_html=True)

    # --- B. ?銵函?豢??? ---
    st.markdown("#### ?? ?銵函?豢???")

    def calc_key_metrics(m_data):
        df = m_data.get('df', pd.DataFrame())
        res = {'high_adr_days': 0, 'top20_rev_avg': 0, 'bot20_rev_avg': 0,
               'dual_match_days': 0, 'month_label': m_data.get('month_label', '')}
        if df is None or df.empty:
            return res

        avg_adr = m_data.get('avg_adr', 0)

        # 蝣箔??詨潭迤蝣?
        df['adr_val'] = pd.to_numeric(df['adr'], errors='coerce').fillna(0)
        df['rev_val'] = pd.to_numeric(df['revenue'], errors='coerce').fillna(0)

        # 擃?嗆?撟喳? ADR 憭拇
        df['is_high_adr'] = df['adr_val'] > avg_adr
        res['high_adr_days'] = int(df['is_high_adr'].sum())

        # ?思?瘜? (??20% ?? 20%)
        n_days = len(df)
        n_top = max(1, int(round(n_days * 0.2)))

        df_sorted = df.sort_values('rev_val', ascending=False)
        top20_df = df_sorted.head(n_top)
        bot20_df = df_sorted.tail(n_top)

        res['top20_rev_avg'] = top20_df['rev_val'].mean(
        ) if not top20_df.empty else 0
        res['bot20_rev_avg'] = bot20_df['rev_val'].mean(
        ) if not bot20_df.empty else 0

        # ??憭拇嚗? 20% ??乩葉嚗DR 銋之?潛?像??ADR ?予??
        dual_match_df = top20_df[top20_df['is_high_adr']]
        res['dual_match_days'] = int(len(dual_match_df))
        res['dual_match_dates'] = dual_match_df['date'].sort_values(
        ).tolist() if not dual_match_df.empty else []

        return res

    curr_metrics = calc_key_metrics(m_curr)
    prev_metrics = calc_key_metrics(m_prev)
    pprev_metrics = calc_key_metrics(m_prev_prev)
    next_metrics = calc_key_metrics(m_next)

    def metric_diff_card(label, diff, target_label, unit="憭?):
        color = '#2ecc71' if diff >= 0 else '#e74c3c'
        status = '?祆?憭? if diff > 0 else '頛?' if diff < 0 else '?像'
        return f'<div style="flex: 1; min-width: 150px; background: #fff; padding: 12px; border-radius: 8px; border: 1px solid #eee; margin-bottom: 10px;"><p style="margin:0; font-size:12px; color:#999;">??{target_label} ?豢?</p><div style="display: flex; align-items: baseline; gap: 8px; margin-top: 5px;"><strong style="font-size:18px; color:{color};">{abs(diff)} {unit}</strong><span style="font-size:11px; color:#666;">({status})</span></div></div>'

    # 憭拇撌桃
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
            <p style="margin:0; font-size:14px; color:#666;">?? <strong>擃?嗆?撟喳? ADR 憭拇 (?祆?: {curr_metrics['high_adr_days']} 憭?</strong></p>
            <div style="display: flex; gap: 15px; margin-top: 10px; flex-wrap: wrap;">
                {metric_diff_card("????, diff_adr_pp, pprev_metrics['month_label'])}
                {metric_diff_card("銝?", diff_adr_p, prev_metrics['month_label'])}
                {metric_diff_card("銝???", diff_adr_n, next_metrics['month_label'])}
            </div>
            <p style="margin:15px 0 0 0; font-size:14px; color:#666;">?? <strong>??憭拇嚗? 20% ?銝? ADR (?祆?: {curr_metrics['dual_match_days']} 憭?</strong></p>
            <div style="display: flex; gap: 15px; margin-top: 10px; flex-wrap: wrap;">
                {metric_diff_card("????, diff_dual_pp, pprev_metrics['month_label'])}
                {metric_diff_card("銝?", diff_dual_p, prev_metrics['month_label'])}
                {metric_diff_card("銝???", diff_dual_n, next_metrics['month_label'])}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with kp_col2:
        st.markdown(f"""
        <div style="background: #fffcf5; padding: 15px; border-radius: 10px; border-left: 5px solid #f39c12; margin-bottom: 20px; height: 100%;">
            <p style="margin:0; font-size:14px; color:#666;">?? <strong>擃???? (?祆?)</strong></p>
            <div style="margin-top: 20px;">
                <p style="margin:0; font-size:13px; color:#999;">? ??20% ???(Top 20%) 撟喳??</p>
                <h3 style="margin: 5px 0 15px 0; color: #d35400;">NT$ {int(curr_metrics['top20_rev_avg']):,}</h3>
                <p style="margin:0; font-size:13px; color:#999;">?? 敺?20% ???(Bottom 20%) 撟喳??</p>
                <h3 style="margin: 5px 0 15px 0; color: #7f8c8d;">NT$ {int(curr_metrics['bot20_rev_avg']):,}</h3>
                <hr style="border: 0; border-top: 1px dashed #eee; margin: 15px 0;">
                <p style="margin:0; font-size:12px; color:#888;">? <strong>閫??</strong>嚗?? 20% ?像???嗅榆頝憭扳?嚗誨銵冽楚?箸?平蝮曉榆頝之嚗??瘛⊥?撥靽??/p>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # --- B3. OCC vs ADR ?情???寡那?瑕? ---
    st.markdown("#### ? 摰瘞港?閮箸嚗??輻? vs 撟喳??踹 ?情????隞亙僑蝝像 ADR ?箏?蝺皞?")
    scatter_df = m_curr['df'].copy()
    if not scatter_df.empty:
        scatter_df['occ_val'] = pd.to_numeric(
            scatter_df['occ_rate'], errors='coerce').fillna(0)
        scatter_df['adr_val'] = pd.to_numeric(
            scatter_df['adr'], errors='coerce').fillna(0)
        scatter_df['day'] = pd.to_datetime(scatter_df['date']).dt.day

        # 隞乓僑蝝像 ADR????Y 頠詨????摰Ｚ??ㄧ撖血?摨?嚗??楚?交?雿?
        y_adr_s, y_pure_adr_s = fetch_yearly_metrics(selected_date.year)
        adr_anchor = y_pure_adr_s if y_pure_adr_s > 0 else m_curr.get(
            'avg_adr', scatter_df['adr_val'].mean())
        anchor_label = f'撟渡?撟?ADR ${int(adr_anchor):,}'
        anchor_color = '#000000'
        occ_threshold = 85.0  # 擃??輻??瑼?

        def classify_quadrant(row):
            hi_occ = row['occ_val'] >= occ_threshold
            hi_adr = row['adr_val'] >= adr_anchor
            if hi_occ and hi_adr:
                return '?? ?嚗?OCC+擃DR嚗?
            if hi_occ and not hi_adr:
                return '? 鞈方都嚗?OCC+雿DR嚗?
            if not hi_occ and hi_adr:
                return '? 摰??嚗?OCC+擃DR嚗?
            return '? 瘛∪迤甇餅偌嚗?OCC+雿DR嚗?

        scatter_df['鞊⊿?'] = scatter_df.apply(classify_quadrant, axis=1)

        color_map = {
            '?? ?嚗?OCC+擃DR嚗?: '#ff9f43',
            '? 鞈方都嚗?OCC+雿DR嚗?: '#e74c3c',
            '? 摰??嚗?OCC+擃DR嚗?: '#f1c40f',
            '? 瘛∪迤甇餅偌嚗?OCC+雿DR嚗?: '#3498db',
        }

        scatter_chart = alt.Chart(scatter_df).mark_circle(size=100, opacity=0.8).encode(
            x=alt.X('occ_val:Q', title='雿??(%)',
                    scale=alt.Scale(domain=[0, 105])),
            y=alt.Y('adr_val:Q', title='撟喳??踹 ADR (NT$)',
                    scale=alt.Scale(zero=False)),
            color=alt.Color('鞊⊿?:N',
                            scale=alt.Scale(
                                domain=list(color_map.keys()),
                                range=list(color_map.values())
                            ),
                            legend=alt.Legend(
                                title="鞊⊿???", orient="bottom", columns=2)
                            ),
            tooltip=[
                alt.Tooltip('date:N', title='?交?'),
                alt.Tooltip('occ_val:Q', title='雿??(%)', format='.1f'),
                alt.Tooltip('adr_val:Q', title='ADR (NT$)', format=',.0f'),
                alt.Tooltip('鞊⊿?:N', title='鞊⊿?'),
            ]
        )

        # 撟渡?撟?ADR 瘞游像頛蝺?
        adr_rule = alt.Chart(pd.DataFrame({'y': [adr_anchor]})).mark_rule(
            color=anchor_color, strokeDash=[6, 3], strokeWidth=2
        ).encode(y='y:Q')
        adr_label = alt.Chart(pd.DataFrame({'y': [adr_anchor], 'x': [105], 'text': [anchor_label]})).mark_text(
            align='right', dx=-4, dy=-8, color=anchor_color, fontSize=11, fontWeight='bold'
        ).encode(x='x:Q', y='y:Q', text='text:N')

        # 85% OCC ?頛蝺?
        occ_rule = alt.Chart(pd.DataFrame({'x': [occ_threshold]})).mark_rule(
            color='#7f8c8d', strokeDash=[6, 3], strokeWidth=1.5
        ).encode(x='x:Q')
        occ_label = alt.Chart(pd.DataFrame({'x': [occ_threshold], 'y': [scatter_df['adr_val'].max() * 1.05], 'text': ['85% OCC ?瑼?]})).mark_text(
            align='left', dx=4, color='#7f8c8d', fontSize=11, fontWeight='bold'
        ).encode(x='x:Q', y='y:Q', text='text:N')

        final_chart = alt.layer(scatter_chart, adr_rule, adr_label, occ_rule, occ_label).properties(
            height=380,
            title=f"{m_curr['month_label']} 瘥摰瘞港?閮箸嚗???隞?”銝憭抬?隞亙僑蝝像 ADR ?箏?蝺?"
        )
        st.altair_chart(final_chart, use_container_width=True)

        # ?情?予?豢?閬?
        q_counts = scatter_df['鞊⊿?'].value_counts()
        q_cols = st.columns(4)
        for i, (q_name, color) in enumerate(color_map.items()):
            cnt = q_counts.get(q_name, 0)
            q_cols[i].markdown(
                f"<div style='background:{color}22; border-left:4px solid {color}; padding:10px; border-radius:6px; text-align:center;'>"
                f"<p style='margin:0; font-size:12px; color:#555;'>{q_name}</p>"
                f"<strong style='font-size:22px;'>{cnt} 憭?/strong></div>",
                unsafe_allow_html=True
            )
        st.write("")

        # --- 摰????(Pricing Success Rate) ---
        ideal_cnt = q_counts.get('?? ?嚗?OCC+擃DR嚗?, 0)
        cheap_cnt = q_counts.get('? 鞈方都嚗?OCC+雿DR嚗?, 0)
        high_occ_total = ideal_cnt + cheap_cnt
        success_rate = (ideal_cnt / high_occ_total *
                        100) if high_occ_total > 0 else 0

        # 閮?銝????寞???雿撠?
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
            verdict = '? 摰?賢??芰?'
        elif success_rate >= 60:
            bar_color = '#f39c12'
            verdict = '? 摰?賢?撠'
        else:
            bar_color = '#e74c3c'
            verdict = '? 摰?賢?敺??

        st.markdown(f"""
        <div style="background:#f8f9fa; border-radius:10px; padding:20px; margin-top:10px; border-left: 5px solid {bar_color};">
            <p style="margin:0 0 8px 0; font-size:14px; color:#555;">
                ?? <strong>擃??踵摰????/strong>
                <span style="font-size:12px; color:#aaa; margin-left:8px;">擃CC 憭拇??{high_occ_total} 憭抬??嗡葉 {int(ideal_cnt)} 憭?ADR 頞?撟渡?撟喳皞?/span>
            </p>
            <div style="display:flex; align-items:baseline; gap:15px; flex-wrap:wrap;">
                <strong style="font-size:40px; color:{bar_color};">{success_rate:.1f}%</strong>
                <span style="font-size:14px;">{verdict}</span>
                <span style="font-size:14px; color:{rate_color}; font-weight:bold;">vs 銝? {prev_success_rate:.1f}% ({rate_sign}{rate_diff:.1f}%)</span>
            </div>
            <div style="background:#e0e0e0; border-radius:999px; height:10px; margin-top:10px;">
                <div style="background:{bar_color}; width:{min(success_rate, 100):.1f}%; height:10px; border-radius:999px; transition: width 0.5s;"></div>
            </div>
            <p style="margin:8px 0 0 0; font-size:12px; color:#888;">? ?格?嚗??酗鞈?予?詻???撠?1-2 憭抬???撠????典? 80%</p>
        </div>
        """, unsafe_allow_html=True)
        st.write("")

    # --- B2. ?喳??唬???憭扳暑???霅血 ---
    st.markdown("#### ? ?喳??唬???憭扳暑???霅血 (?芯? 30 憭?")
    upcoming_holidays = fetch_upcoming_holidays(selected_date, 30)

    # ?蔥?啣??之瘣餃??唾郎?勗?銵?(?????暑??
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
                e_list.append(f"??儭?{row['event_name']}{v_suffix}")
                e_labels.append(EVENT_TYPE_LABELS.get(
                    row['event_type'], '[瘣蒸'))

        if h_info or e_list:
            all_flags = (h_info['flags'] if h_info else "") + \
                "".join(sorted(list(set(e_labels))))
            details_html = ""
            if h_info:
                details_html += f"<div style='margin-bottom:4px; color:#856404;'>?? {h_info['details']}</div>"
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
        st.info("?芯? 30 憭拙?⊿?憭批??交??啣?瘣餃???)

    st.divider()

    # --- C. ??暑?蜀????---
    st.markdown("#### ?? 蝮暹?鞎Ｙ摨虫漱????)
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
                f"<div style='background:#f1f8ff; padding:10px; border-radius:5px; border-left:3px solid #3498db; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>??蝐?({len(holiday_days)}憭?</p><strong style='font-size:16px;'>{h_occ:.1f}% / NT$ {int(h_adr):,}</strong></div>", unsafe_allow_html=True)
            c2.markdown(
                f"<div style='background:#f8f9fa; padding:10px; border-radius:5px; border-left:3px solid #ccc; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>?⊥?蝐?({len(non_holiday_days)}憭?</p><strong style='font-size:16px;'>{nh_occ:.1f}% / NT$ {int(nh_adr):,}</strong></div>", unsafe_allow_html=True)
            color = "#2ecc71" if diff_occ >= 0 else "#e74c3c"
            c3.markdown(
                f"<div style='background:#f0fff4; padding:10px; border-radius:5px; border-left:3px solid #2ecc71; height:100%;'><p style='margin:0; font-size:12px; color:#666;'>撣嗅???</p><strong style='font-size:16px; color:{color};'>{diff_occ:+.1f}% / NT$ {int(diff_adr):+,}</strong></div>", unsafe_allow_html=True)
            st.markdown("<div style='margin-bottom:15px;'></div>",
                        unsafe_allow_html=True)

        def render_exclusive_matrix(df, title_suffix=""):
            st.markdown(f"**?? ?情??隞找漱????{title_suffix}**")

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

            # 鞊⊿? 4: 蝝像??
            with col1:
                st.markdown(
                    f"<div style='background:#1e293b; padding:15px; border-radius:8px; border-left:4px solid #94a3b8; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#94a3b8; font-weight:bold;'>?情??4??撟單 ({days_pw}憭?</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>?箸?撠蝯?/p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_pw:.1f}% / NT$ {int(adr_pw):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:11px; color:#64748b;'>?⊥暑?蝭?嗥??箸?蝺?/p>"
                    f"</div>",
                    unsafe_allow_html=True
                )

            # 鞊⊿? 1: 蝝暑?
            with col2:
                diff_occ = occ_pe - occ_pw if days_pe > 0 and days_pw > 0 else 0
                diff_adr = adr_pe - adr_pw if days_pe > 0 and days_pw > 0 else 0
                color = "#10b981" if diff_adr >= 0 else "#ef4444"
                bg_style = "background:#0f172a; border-left:4px solid #3b82f6;"
                desc = f"瘛冽??? <span style='color:{color}; font-weight:bold;'>{format_diff(diff_occ, True)} / {format_diff(diff_adr)}</span>" if days_pe > 0 else "?⊥??
                st.markdown(
                    f"<div style='{bg_style} padding:15px; border-radius:8px; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#3b82f6; font-weight:bold;'>?情??1??瘣餃???({days_pe}憭?</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>???憭扳暑??/p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_pe:.1f}% / NT$ {int(adr_pe):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:12px; color:#cbd5e1;'>{desc}</p>"
                    f"</div>",
                    unsafe_allow_html=True
                )

            # 鞊⊿? 2: 蝝??嗆
            with col3:
                diff_occ = occ_ph - occ_pw if days_ph > 0 and days_pw > 0 else 0
                diff_adr = adr_ph - adr_pw if days_ph > 0 and days_pw > 0 else 0
                color = "#10b981" if diff_adr >= 0 else "#ef4444"
                bg_style = "background:#0f172a; border-left:4px solid #eab308;"
                desc = f"瘛冽??? <span style='color:{color}; font-weight:bold;'>{format_diff(diff_occ, True)} / {format_diff(diff_adr)}</span>" if days_ph > 0 else "?⊥??
                st.markdown(
                    f"<div style='{bg_style} padding:15px; border-radius:8px; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#eab308; font-weight:bold;'>?情??2??蝭?嗆 ({days_ph}憭?</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>??????/p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_ph:.1f}% / NT$ {int(adr_ph):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:12px; color:#cbd5e1;'>{desc}</p>"
                    f"</div>",
                    unsafe_allow_html=True
                )

            # 鞊⊿? 3: 暺?????
            with col4:
                diff_occ = occ_di - occ_pw if days_di > 0 and days_pw > 0 else 0
                diff_adr = adr_di - adr_pw if days_di > 0 and days_pw > 0 else 0
                color = "#10b981" if diff_adr >= 0 else "#ef4444"
                bg_style = "background:#1e1b4b; border-left:4px solid #a855f7;"
                desc = f"瘛冽??? <span style='color:{color}; font-weight:bold;'>{format_diff(diff_occ, True)} / {format_diff(diff_adr)}</span>" if days_di > 0 else "?⊥??
                st.markdown(
                    f"<div style='{bg_style} padding:15px; border-radius:8px; height:100%; min-height:140px; color:#f8fafc;'>"
                    f"<p style='margin:0; font-size:12px; color:#a855f7; font-weight:bold;'>?情??3????? ({days_di}憭?</p>"
                    f"<p style='margin:5px 0 0 0; font-size:12px; color:#cbd5e1;'>瘣餃? 嚗?蝭?嗥???/p>"
                    f"<strong style='font-size:18px; color:#f1f5f9;'>{occ_di:.1f}% / NT$ {int(adr_di):,}</strong>"
                    f"<p style='margin:8px 0 0 0; font-size:12px; color:#cbd5e1;'>{desc}</p>"
                    f"</div>",
                    unsafe_allow_html=True
                )
            st.markdown("<div style='margin-bottom:20px;'></div>",
                        unsafe_allow_html=True)

        render_impact_row(curr_df, 'is_any', "蝬??? (? + ?啣?瘣餃?)", "??")
        render_impact_row(curr_df, 'is_h', "?????嗅???, "??")
        render_impact_row(curr_df, 'is_e', "???憭扳暑????, "??儭?)

        st.markdown("<div style='margin-bottom:20px;'></div>",
                    unsafe_allow_html=True)
        render_exclusive_matrix(curr_df, "(?嗆?)")

        st.divider()

        # --- C2. ?銝???蝮暹??? (?瑟?頞典) ---
        st.markdown("#### ???銝???蝮暹??? (?瑟?頞典)")
        # ???????交?
        m1_date = get_month_delta(selected_date, -1)
        m2_date = get_month_delta(selected_date, -2)
        m3_date = get_month_delta(selected_date, -3)

        m1_sum = fetch_month_summary(m1_date.year, m1_date.month)
        m2_sum = fetch_month_summary(m2_date.year, m2_date.month)
        m3_sum = fetch_month_summary(m3_date.year, m3_date.month)

        hist_df = pd.concat([m1_sum['df'], m2_sum['df'],
                            m3_sum['df']], ignore_index=True)

        if not hist_df.empty:
            # 皞?甇瑕鞈???蝐?
            def get_hist_flags(row):
                d = row['date']
                y, m = int(d.split('-')[0]), int(d.split('-')[1])
                h_f = fetch_holidays_for_month(
                    y, m).get(d, {}).get('flags', '')
                e_f = ""
                if not taipei_events_df.empty:
                    de = taipei_events_df[taipei_events_df['date'] == d]
                    for _, r in de.iterrows():
                        e_f += EVENT_TYPE_LABELS.get(r['event_type'], '[瘣蒸')
                return (h_f != ''), (e_f != '')

            # ?箔??嚗????嗾?????亥???
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

            render_impact_row(hist_df, 'is_any', "蝬??? (?銝???)", "??")
            render_impact_row(hist_df, 'is_h', "?????嗅???(?銝???)", "??")
            render_impact_row(hist_df, 'is_e', "???憭扳暑????(?銝???)", "??儭?)

            st.markdown("<div style='margin-bottom:20px;'></div>",
                        unsafe_allow_html=True)
            render_exclusive_matrix(hist_df, "(?銝???)")
        else:
            st.info("撠頞喳??風?脫?脰??瑟?頞典????)

        with st.expander("?? ?亦??祆?????亥??啣?瘣餃?閰喟敦皜"):
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
                            e_info += f", ??儭?{r['event_name']}{v_suffix} ({r['event_type']})"

                    if h_info['flags'] or e_info:
                        st.markdown(
                            f"- **{d}** {h_info['flags']}{e_info}: {', '.join(h_info['details'])}")
                        has_any = True
            if not has_any:
                st.write("?祆??∩遙雿?憭扳暑?????)
    else:
        st.info("?祆?撠???豢??臭?????)

    st.divider()

    # --- D. ?漲???? (??撠?) ---
    st.subheader("?? ?漲????撠?")

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)

    def render_metric_col(month_data, label):
        st.markdown(
            f"<p style='text-align:center; color:#777; margin-bottom:10px;'>{label} ({month_data['month_label']})</p>", unsafe_allow_html=True)
        if not month_data['df'].empty:
            st.markdown(make_card(
                "?嗆?蝮賜???, f"NT$ {int(month_data['rev']):,}", "card-theme-orange", "card-bg-dark"), unsafe_allow_html=True)
            st.markdown(make_card(
                "?嗆?撟喳??踹", f"NT$ {int(month_data['avg_adr']):,}", "card-theme-green", "card-bg-dark"), unsafe_allow_html=True)
            st.markdown(make_card(
                "?嗆?雿??, f"{month_data['avg_occ']:.1f}%", "card-theme-blue", "card-bg-dark"), unsafe_allow_html=True)
            st.markdown(make_card(
                "?嗆? RevPAR", f"NT$ {int(month_data['revpar']):,}", "card-theme-purple", "card-bg-dark"), unsafe_allow_html=True)
        else:
            st.info("?怎?豢?")

    with col_m1:
        render_metric_col(m_prev_prev, "??????)
    with col_m2:
        render_metric_col(m_prev, "?儭?銝?")
    with col_m3:
        render_metric_col(m_curr, "???祆?")
    with col_m4:
        render_metric_col(m_next, "?塚? 銝?")

    # --- D. ?漲???? - ?撌桃 ---
    st.markdown("#### ?? ?漲????嚗??萄榆?啣?瘥?(?祆? vs ?嗡??遢)")

    def calculate_diff_row(current_val, compare_val, is_currency=True, is_percent=False):
        if compare_val == 0:
            return "<span style='color:#777;'>-</span>"
        diff = current_val - compare_val
        if is_currency:
            diff_str = f"{'?? if diff >= 0 else '??} NT$ {abs(int(diff)):,}"
        elif is_percent:
            diff_str = f"{'?? if diff >= 0 else '??} {abs(diff):.1f}%"
        else:
            diff_str = f"{'?? if diff >= 0 else '??} {abs(diff):.1f}"

        color = "#2ecc71" if diff >= 0 else "#e74c3c"  # 憓??箇??莎?皜??箇???
        return f"<span style='color:{color}; font-weight:bold;'>{diff_str}</span>"

    diff_table_html = f"""
    <table style="width:100%; border-collapse: collapse; margin-top: 10px; font-size: 15px;">
        <tr style="background-color: #f1f3f6; text-align: left;">
            <th style="padding: 12px; border: 1px solid #ddd;">???</th>
            <th style="padding: 12px; border: 1px solid #ddd;">???? ({m_prev_prev['month_label']}) ?豢?</th>
            <th style="padding: 12px; border: 1px solid #ddd;">????({m_prev['month_label']}) ?豢?</th>
            <th style="padding: 12px; border: 1px solid #ddd;">????({m_next['month_label']}) ?豢?</th>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">?嗆?蝮賜???/td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['rev'], m_prev_prev['rev'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['rev'], m_prev['rev'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['rev'], m_next['rev'])}</td>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">?嗆?撟喳??踹 (ADR)</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_adr'], m_prev_prev['avg_adr'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_adr'], m_prev['avg_adr'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_adr'], m_next['avg_adr'])}</td>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">?嗆?雿??(%)</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_occ'], m_prev_prev['avg_occ'], False, True)}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_occ'], m_prev['avg_occ'], False, True)}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['avg_occ'], m_next['avg_occ'], False, True)}</td>
        </tr>
        <tr>
            <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">?嗆? RevPAR</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['revpar'], m_prev_prev['revpar'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['revpar'], m_prev['revpar'])}</td>
            <td style="padding: 12px; border: 1px solid #ddd;">{calculate_diff_row(m_curr['revpar'], m_next['revpar'])}</td>
        </tr>
    </table>
    """
    st.write(diff_table_html, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("閮鳴?RevPAR 閮??孵??箝?像???輻? ? ?嗆?撟喳??踹??撌桃撠?銝???隞?”?祆?頛?嚗 隞?”?祆?頛???)

    st.divider()

    # --- 3. ????? ---
    st.subheader("? ?????")

    # ?脣???摮璅?(????豢?隞?
    month_key = selected_date.strftime('%Y-%m')
    current_target = get_monthly_target(month_key)
    m_rev = m_curr['rev']  # 雿輻??閮?憟賜??祆??

    t_col1, t_col2 = st.columns([1, 2])
    with t_col1:
        new_target = st.number_input(f"閮剖? {month_key} ?格?璆剔蜀 (NT$)", min_value=0,
                                     step=10000, value=current_target, key=f"target_input_{month_key}")
        if new_target != current_target:
            save_monthly_target(month_key, new_target)
            st.toast(f"撌脫??{month_key} ?格?璆剔蜀嚗?)
            time.sleep(0.5)
            st.rerun()

    if new_target > 0:
        gap = new_target - m_rev
        stretch_goal = new_target * 1.1
        stretch_gap = stretch_goal - m_rev
        progress = min(1.0, m_rev / new_target)
        st.progress(progress, text=f"?格????? {progress*100:.1f}%")

        # ??脣漲憭?摯 (Run-Rate Forecast)
        active_days = m_curr['df'][m_curr['df']['revenue'] > 0]
        elapsed_days = len(active_days)

        if elapsed_days > 0:
            import calendar
            total_days = calendar.monthrange(
                selected_date.year, selected_date.month)[1]
            daily_avg = m_rev / elapsed_days
            projected_rev = daily_avg * total_days
            projected_progress = projected_rev / new_target

            # ?郎憿??摮?
            status_color = "#2ecc71" if projected_rev >= new_target else "#ef4444"
            status_icon = "??" if projected_rev >= new_target else "??"
            status_text = "靘?脣漲嚗?閮?*?舫??拚?璅?*嚗? if projected_rev >= new_target else "靘?脣漲嚗?*???航?摨?*嚗遣霅啗矽?游????寞??撥靽嚗?

            st.markdown(
                f"<div style='background: #1e293b; padding: 15px; border-radius: 8px; border-left: 5px solid {status_color}; margin-top: 10px; margin-bottom: 15px; color: #f8fafc;'>"
                f"<p style='margin:0; font-size:13px; color:#94a3b8;'>? <strong>?嗆???脣漲憭?摯 (Pacing Forecast)</strong></p>"
                f"<div style='display: flex; gap: 20px; align-items: center; margin-top: 5px; flex-wrap: wrap; font-size: 13px;'>"
                f"<div>撌脩絞閮予?? <strong style='color:#f1f5f9;'>{elapsed_days} / {total_days} 憭?/strong></div>"
                f"<div>?嗅??亙??: <strong style='color:#f1f5f9;'>NT$ {int(daily_avg):,}</strong></div>"
                f"<div>?摯??蝮賜??? <strong style='color:{status_color}; font-size: 15px;'>NT$ {int(projected_rev):,}</strong></div>"
                f"<div>?摯?蝯???: <strong style='color:{status_color}; font-size: 15px;'>{projected_progress*100:.1f}%</strong></div>"
                f"</div>"
                f"<p style='margin: 8px 0 0 0; font-size: 12px; color: #cbd5e1;'>{status_icon} {status_text}</p>"
                f"</div>",
                unsafe_allow_html=True
            )

        a_col1, a_col2, a_col3 = st.columns(3)
        if gap <= 0:
            t_card = make_card("?格????瘜?, "?? 撌脤?璅?", "card-theme-green", "", "??)
        else:
            t_card = make_card(
                "頝?格??榆", f"NT$ {int(gap):,}", "card-theme-red", "", "?")
        a_col1.markdown(t_card, unsafe_allow_html=True)
        a_col2.markdown(make_card(
            "頞??格? (+10%)", f"NT$ {int(stretch_goal):,}", "card-theme-orange", "", "??"), unsafe_allow_html=True)
        if stretch_gap <= 0:
            s_card = make_card("頞????瘜?, "? 撌脰?璅???",
                               "card-theme-green", "card-bg-dark", "??")
        else:
            s_card = make_card(
                "頝頞??榆", f"NT$ {int(stretch_gap):,}", "card-theme-purple", "", "??)
        a_col3.markdown(s_card, unsafe_allow_html=True)
    else:
        st.info("? 隢銝頛詨?祆??格?璆剔蜀嚗頂蝯勗??芸??箸閮???撌株???)

with tab3:
    st.header("?完 ?踹??豢?")
    st.number_input("隞蝮賣?瘨??, min_value=0, step=1,
                    key="input_cleaned", on_change=on_input_change)
    st.number_input("?/蝥??, min_value=0, step=1,
                    key="input_hk_co", on_change=on_input_change)
    st.number_input("瘥犖撟喳????, min_value=0.0, step=0.1,
                    key="input_hk_avg", on_change=on_input_change)
    st.number_input("?踹?隢頃鞎餌", min_value=0, step=100,
                    key="input_hk_exp", on_change=on_input_change)

with tab4:
    st.header("?儭?擗輒?豢?")
    st.subheader("?? ?豢??梯”銝")
    rest_file = st.file_uploader("銝擗輒?梯” (Excel)嚗??芸??隞賢銵典神?亥??澈嚗?, type=[
                                 "xls", "xlsx"], key="rest_uploader")

    if rest_file:
        # ?典神?亙?憓??汗?
        try:
            # ?急??瑁?閫?? (銝??亥??澈)
            # ?箔??????ｇ???ㄐ?陛???汗
            df_preview = pd.read_excel(rest_file, header=None)
            st.info("?? **?梯”?批捆?郊??嚗?*")

            p_month_rev = 0
            p_avg_spent = 0
            found_days = 0

            for i, row in df_preview.iterrows():
                row_str = " ".join([str(v) for v in row if pd.notna(v)])
                if ('撌脩?蝞??? in row_str or '???? in row_str) and '?拚?' not in row_str and '銝??? not in row_str:
                    for v in row:
                        if any(c.isdigit() for c in str(v)) and not any(k in str(v) for k in ['撌脩?蝞???, '????]):
                            try:
                                p_month_rev = int(float(str(v).replace(
                                    'NT$', '').replace('$', '').replace(',', '').strip()))
                                break
                            except:
                                pass
                if '摰Ｗ?? in row_str:
                    for v in row:
                        if any(c.isdigit() for c in str(v)) and '摰Ｗ?? not in str(v):
                            try:
                                p_avg_spent = int(float(str(v).replace(
                                    'NT$', '').replace('$', '').replace(',', '').strip()))
                                break
                            except:
                                pass
                if re.search(r'\d{1,2}/\d{1,2}', str(row[0])):
                    found_days += 1

            pv_col1, pv_col2, pv_col3 = st.columns(3)
            pv_col1.metric("颲刻??箸?蝯??", f"NT$ {p_month_rev:,}")
            pv_col2.metric("颲刻??箏像?恥?桀", f"NT$ {p_avg_spent:,}")
            pv_col3.metric("颲刻??箸??交?蝝?, f"{found_days} 蝑?)

            if p_month_rev == 0:
                st.warning("?? 蝟餌絞?芾敺銵其葉?芸??曉??蝯????隢Ⅱ隤銵冽撘???瑼Ｘ??)

            if st.button("? 蝣箄??∟炊嚗神?亦頂蝯梯??澈", key="rest_btn"):
                saved_count = parse_and_save_restaurant(
                    rest_file, selected_date.year)
                if saved_count:
                    st.success(f"?????湔 {saved_count} 蝑??仿?撱唾???")
                    time.sleep(1)
                    st.rerun()
        except Exception as ex:
            st.error(f"?汗憭望?: {ex}")

    st.divider()
    st.subheader(f"擗輒??蝣箄?? ({date_str})")

    st.markdown("#### ?? ?拚??豢?")
    b1, b2, b3 = st.columns(3)
    b1.number_input("?蜓憿?隡唬?摰?, min_value=0, step=1,
                    key="input_bf_theme_est", on_change=on_input_change)
    b1.number_input("?蜓憿祕??摰?, min_value=0, step=1,
                    key="input_bf_theme_act", on_change=on_input_change)

    b2.number_input("????隡唬?摰?, min_value=0, step=1,
                    key="input_bf_zq_est", on_change=on_input_change)
    b2.number_input("???祕??摰?, min_value=0, step=1,
                    key="input_bf_zq_act", on_change=on_input_change)

    b3.number_input("?擗函蜇??隡?, min_value=0, step=1,
                    key="input_bf_total_est", on_change=on_input_change)
    b3.number_input("?擗函蜇?祕??, min_value=0, step=1,
                    key="input_bf_total_act", on_change=on_input_change)

    st.markdown("#### ? 銝??嗆??)
    a1, a2, a3 = st.columns(3)
    a1.number_input("?蜓憿?隡唬?摰?, min_value=0, step=1,
                    key="input_af_theme_est", on_change=on_input_change)
    a1.number_input("?蜓憿祕??摰?, min_value=0, step=1,
                    key="input_af_theme_act", on_change=on_input_change)

    a2.number_input("????隡唬?摰?, min_value=0, step=1,
                    key="input_af_zq_est", on_change=on_input_change)
    a2.number_input("???祕??摰?, min_value=0, step=1,
                    key="input_af_zq_act", on_change=on_input_change)

    a3.number_input("?擗函蜇??隡?, min_value=0, step=1,
                    key="input_af_total_est", on_change=on_input_change)
    a3.number_input("?擗函蜇?祕??, min_value=0, step=1,
                    key="input_af_total_act", on_change=on_input_change)

    st.markdown("#### ?? ?蝯?蝮賣????)
    c1, c2, c3 = st.columns(3)
    c1.number_input("撌脩?蝞???(?冽?)", min_value=0, step=100,
                    key="input_rest_mrev", on_change=on_input_change)
    c2.number_input("撟喳?摰Ｗ??, min_value=0, step=10,
                    key="input_rest_aspent", on_change=on_input_change)
    c3.number_input("THE PEAK 隢頃鞎餌", min_value=0, step=100,
                    key="input_rest_exp", on_change=on_input_change)

    col_rest1, col_rest2 = st.columns(2)
    col_rest1.number_input("The Peak ?嗆靘恥??, min_value=0,
                           step=1, key="input_peak_act", on_change=on_input_change)
    col_rest2.number_input("Happy Hour ?嗆靘恥??, min_value=0,
                           step=1, key="input_hh_act", on_change=on_input_change)

with tab5:
    st.header("? 撌亙??豢?")
    st.number_input("隞敺耨?踵", min_value=0, step=1,
                    key="input_repair", on_change=on_input_change)
    st.text_area("靽桃?蝝??, key="input_maint_rec", on_change=on_input_change)
    st.number_input("撌亙?隢頃鞎餌", min_value=0, step=100,
                    key="input_maint_exp", on_change=on_input_change)

with tab6:
    st.header("?? 瘥??蝝??)

    # --- ??梯”銝 + ??頛詨 (敺????唳?宏?? ---
    with st.expander("?? ??梯”銝 & ?嗆?詨???蝣箄?", expanded=False):
        jinxu_file = st.file_uploader(
            "銝??梯” (Excel/CSV)嚗??芸??隞賢銵典神?亥??澈嚗?, type=["csv", "xls", "xlsx"], key="jinxu_uploader")
        if jinxu_file:
            if st.button("? 撖怠蝟餌絞鞈?摨?):
                saved_count = parse_and_save_jinxu(jinxu_file)
                if saved_count:
                    st.success(f"????撠?{saved_count} 蝑??亥????亦頂蝯梯??澈嚗????航?矽?箝?)
                    time.sleep(1)
                    st.rerun()
        st.divider()
        st.subheader(f"?? ?嗆?詨???蝣箄? ({date_str})")
        rc1, rc2, rc3 = st.columns(3)
        rc1.number_input("閮??(%)", min_value=0.0, max_value=100.0,
                         step=0.1, key="input_occ", on_change=on_input_change)
        rc2.number_input("ADR (撟喳??踹)", min_value=0, step=10,
                         key="input_adr", on_change=on_input_change)
        rc3.number_input("蝮賜???, min_value=0, step=100,
                         key="input_rev", on_change=on_input_change)
        rc4, rc5 = st.columns(2)
        rc4.number_input("蝮賭??踵", min_value=0, step=1,
                         key="input_rooms", on_change=on_input_change)
        rc5.number_input("瑹隢頃鞎餌", min_value=0, step=100,
                         key="input_counter_exp", on_change=on_input_change)
        st.text_area("鞎?摰Ｚ迄", key="input_complaints", on_change=on_input_change)

    if selected_week != "--- ???梢?閬?---":
        import calendar
        _, last_day_of_month = calendar.monthrange(
            selected_date.year, selected_date.month)

        # 閫???豢?????
        week_idx = weekly_options.index(selected_week)
        start_d = (week_idx - 1) * 7 + 1
        if week_idx == 5:
            end_d = last_day_of_month
        else:
            end_d = min(start_d + 6, last_day_of_month)

        st.subheader(f"?? {selected_week} 敹恍祟閬芋撘?)
        st.info(
            f"甇??亦? {selected_date.year}撟游漲 {selected_date.month}?遢 ({start_d}????{end_d}?? ???渡???)

        # ?脣?閰脣???????
        c_month_str = selected_date.strftime('%Y-%m')

        for day in range(start_d, end_d + 1):
            target_date = f"{c_month_str}-{day:02d}"
            # ?ㄐ???get_daily_data
            d_data = get_daily_data(target_date)

            with st.expander(f"?? {target_date} ??蝝??, expanded=True):
                day_log = get_daily_log(target_date)
                if day_log:
                    st.markdown(f"**??交隤敦蝭??*\n\n{day_log}")
                    st.divider()
                    col_a, colb, colc = st.columns(3)
                    col_a.metric("雿??, f"{d_data.get('occ_rate', 0)}%")
                    colb.metric("ADR", f"NT$ {int(d_data.get('adr', 0)):,}")
                    colc.metric("?", f"NT$ {int(d_data.get('revenue', 0)):,}")
                else:
                    st.write("?? 甇斗???∩遙雿隤???)

        if st.button("漎? 餈?隞蝺刻摩璅∪?"):
            st.rerun()

    else:
        st.info(
            f"? 隢銝閰喟敦憛怠神 **{date_str}** ?????隤???撌乩???ㄐ?????芸??脣?嚗??????蝬脤?銋??冽?敹憭晞?)
        st.text_area("?? 隞撌乩????敦蝭?勗?嚗?, height=500, key="input_daily_log",
                     placeholder="?臭誑?券ㄐ閮?鈭斤???恥閮渡畾??IP ?亙?蝝啁??身?之靽桃???..蝑?, on_change=on_input_change)

with tab_p:
    st.header("? ?∟頃?梯祥??蝯梯?")

    current_month_str = selected_date.strftime('%Y-%m')

    try:
        # 霈?鞈潭??(?? TTL 隞亦Ⅱ靽?啣???
        possible_names = ["purchase data", "Purchase Data",
                          "purchase_data", "Purchase_Data"]
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
            # 皜?甈??迂 (蝘駁蝛箸)
            df_purchase.columns = df_purchase.columns.astype(str).str.strip()

            # 撠?甈? (?芸?霅?航??蝔梯?擃?
            date_col = next(
                (c for c in df_purchase.columns if '?交?' in c or 'Date' in c), None)
            dept_col = next(
                (c for c in df_purchase.columns if '?券?' in c or 'Dept' in c or '撌亙' in c), None)
            total_col = next(
                (c for c in df_purchase.columns if '撠?' in c or '??' in c or 'Total' in c), None)

            if not date_col or not dept_col or not total_col:
                missing = [c for c, found in [
                    ('?交?', date_col), ('?券?', dept_col), ('撠?', total_col)] if not found]
                st.error(f"???∟頃??蝻箏?敹?甈?嚗', '.join(missing)}")
                st.write("?桀??菜葫?啁?甈???", list(df_purchase.columns))
                st.stop()

            # 蝣箔??交?甈??箸????(?舀瘞?撟渲?銝?祈正?僑)
            def robust_date_parse(val):
                if pd.isna(val):
                    return None
                s = str(val).strip()
                # ?斗?臬?箸??僑?澆? (??/ 銝??撠?
                if '/' in s:
                    res = minguo_to_western(s)
                    if res:
                        return res
                # ?岫璅?閫??
                try:
                    return pd.to_datetime(val).date()
                except:
                    return None

            df_purchase['?交?'] = df_purchase[date_col].apply(robust_date_parse)

            # ???券?甈?蝛箏?(甇賊??啜????
            df_purchase[dept_col] = df_purchase[dept_col].fillna(
                "?芸?憿?).astype(str).str.strip()
            df_purchase.loc[df_purchase[dept_col] == "", dept_col] = "?芸?憿?

            # ?蕪 NaT/None
            df_purchase = df_purchase[df_purchase['?交?'].notna()]

            # ?蕪?嗆??豢?
            m_start = selected_date.replace(day=1)
            import calendar
            _, last_day = calendar.monthrange(
                selected_date.year, selected_date.month)
            m_end = selected_date.replace(day=last_day)

            df_month = df_purchase[(df_purchase['?交?'] >= m_start) & (
                df_purchase['?交?'] <= m_end)].copy()

            # --- ?啣?嚗?敺????豢??冽 MoM ?? ---
            prev_m_date = get_month_delta(selected_date, -1)
            pm_start = prev_m_date.replace(day=1)
            _, pm_last_day = calendar.monthrange(
                prev_m_date.year, prev_m_date.month)
            pm_end = prev_m_date.replace(day=pm_last_day)
            df_prev_month = df_purchase[(df_purchase['?交?'] >= pm_start) & (
                df_purchase['?交?'] <= pm_end)].copy()

            if not df_month.empty:
                # ?詨潭???
                df_month['撠?'] = pd.to_numeric(
                    df_month[total_col], errors='coerce').fillna(0)
                if not df_prev_month.empty:
                    df_prev_month['撠?'] = pd.to_numeric(
                        df_prev_month[total_col], errors='coerce').fillna(0)

                total_month_expense = df_month['撠?'].sum()
                total_prev_expense = df_prev_month['撠?'].sum(
                ) if not df_prev_month.empty else 0

                # 閮?憓??
                mom_delta = total_month_expense - total_prev_expense
                mom_pcnt = (mom_delta / total_prev_expense *
                            100) if total_prev_expense > 0 else 0

                # 1. ?祆?蝮賡??瑁? MoM
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #1f2c56 0%, #2e437c 100%); padding: 25px; border-radius: 15px; text-align: center; color: white; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
                    <p style="margin: 0; font-size: 1.1rem; opacity: 0.8;">?? {current_month_str} ?祆?蝮賡??琿?憿?/p>
                    <h1 style="margin: 10px 0 0 0; font-size: 3rem; font-weight: 800; letter-spacing: 1px;">NT$ {int(total_month_expense):,}</h1>
                </div>
                """, unsafe_allow_html=True)

                # 憿舐內 MoM ??
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.metric("銝???蝮賡?", f"NT$ {int(total_prev_expense):,}")
                with col_m2:
                    st.metric("???琿?憿?(MoM)", f"NT$ {int(mom_delta):,}", delta=int(
                        mom_delta), delta_color="inverse")
                with col_m3:
                    st.metric(
                        "???瑞??", f"{mom_pcnt:.1f}%", delta=f"{mom_pcnt:.1f}%", delta_color="inverse")

                st.divider()

                # --- ?啣虜?潛?改??曉憓?翰?? ---
                st.subheader("?? ?∟頃?啣虜?? (MoM Spikes)")
                # 閮????祆? vs 銝?
                curr_depts = df_month.groupby(
                    dept_col)['撠?'].sum().reset_index()
                curr_depts.columns = ['?券?', '撠?']

                if not df_prev_month.empty:
                    prev_depts = df_prev_month.groupby(
                        dept_col)['撠?'].sum().reset_index()
                    prev_depts.columns = ['?券?', '撠?']
                else:
                    prev_depts = pd.DataFrame(columns=['?券?', '撠?'])

                comparison = pd.merge(
                    curr_depts, prev_depts, on='?券?', how='left', suffixes=('_隞?, '_??)).fillna(0)

                # 摰閮?霈???(?踹? ZeroDivisionError ??Indexing 憿??梢)
                def calc_mom_ratio(row):
                    if row['撠?_??] > 0:
                        return (row['撠?_隞?] - row['撠?_??]) / row['撠?_??] * 100
                    return 100.0 if row['撠?_隞?] > 0 else 0.0

                comparison['霈???] = comparison.apply(calc_mom_ratio, axis=1)

                # ?曉霈??之??20% 銝?憿之?潔?摰?瑼餌? (靘? > 2000)
                spikes = comparison[(comparison['霈???] > 20) & (
                    comparison['撠?_隞?] > 2000)].sort_values('霈???, ascending=False)

                if not spikes.empty:
                    for _, row in spikes.iterrows():
                        st.warning(
                            f"? **{row['?券?']}** ?祆???啣虜嚗?銝?憓 **{row['霈???]:.1f}%** (NT$ {int(row['撠?_隞?]):,})")
                else:
                    st.success("???桀????∟頃??撟喟帘嚗?菜葫?啁撣詨之撟郭??)

                st.divider()

                # 2. ?券?雿?????
                st.subheader("?? ??隢頃雿???")
                dept_summary = df_month.groupby(
                    dept_col)['撠?'].sum().reset_index()
                dept_summary.columns = ['?券?', '撠?']

                # 蝜芾ˊ????(靘????)
                base = alt.Chart(dept_summary).encode(
                    theta=alt.Theta(
                        field="撠?", type="quantitative", stack=True),
                    color=alt.Color(
                        field="?券?",
                        type="nominal",
                        scale=alt.Scale(scheme='category10'),
                        legend=alt.Legend(title="?券?", orient="right"),
                        sort=alt.SortField("撠?", order="descending")
                    ),
                    order=alt.Order("撠?", sort="descending"),
                    tooltip=["?券?", alt.Tooltip(
                        "撠?", format=",.0f", title="蝮賡?憿?(NT$)")]
                ).properties(height=450)

                # ??銝駁?
                chart_arc = base.mark_arc(
                    innerRadius=60, outerRadius=120, stroke="#fff")

                # ?典?擗???憿舐內??
                chart_text = base.mark_text(radius=90, size=14, fontWeight="bold", color="white").encode(
                    text=alt.Text("撠?:Q", format=",.0f")
                )

                st.altair_chart(chart_arc + chart_text,
                                use_container_width=True)

                # --- ?啣?嚗?憌脩蜀????(The Peak & Happy Hour) ---
                st.divider()
                st.subheader("?儭?擗ㄡ蝮暹????祆楛摨血???(Cash-basis)")

                # ?脣??嗆?瘥?豢? (?恍?撱喃?摰Ｘ)
                m_data = fetch_month_summary(
                    selected_date.year, selected_date.month)
                df_daily_rest = m_data['df']

                if not df_daily_rest.empty:
                    # 蝣箔?敹?甈?摮銝西??箸??
                    target_cols = ['rest_day_guests', 'rest_hh_guests',
                                   'revenue', 'bf_total_act', 'af_total_act']
                    for c in target_cols:
                        if c in df_daily_rest.columns:
                            df_daily_rest[c] = pd.to_numeric(df_daily_rest[c].astype(
                                str).str.replace(',', ''), errors='coerce').fillna(0)
                        else:
                            df_daily_rest[c] = 0

                    # --- ?芸???頛荔?憒? The Peak 靘恥?貊 0嚗??芸??蜇 ?拚? + 銝???---
                    def calculate_peak_guests(row):
                        if row['rest_day_guests'] > 0:
                            return row['rest_day_guests']
                        return row['bf_total_act'] + row['af_total_act']

                    df_daily_rest['effective_peak_guests'] = df_daily_rest.apply(
                        calculate_peak_guests, axis=1)

                    # 蝭拚 The Peak ??Happy Hour ?∟頃 (撘瑕?璅∠??寥?)
                    all_depts_list = dept_summary['?券?'].astype(str).tolist()

                    # HH ?寥?嚗???'4'??HH' ??'HAPPY'
                    hh_matched = [d for d in all_depts_list if '4' in d or any(
                        k in d.upper() for k in ['HH', 'HAPPY', '甇⊥???'])]
                    # Peak ?寥?嚗???'PEAK' ??'擗輒'嚗?? HH ?券?
                    peak_matched = [d for d in all_depts_list if (any(k in d.upper(
                    ) for k in ['PEAK', '擗輒', 'THEPEAK', '擗ㄡ'])) and (d not in hh_matched)]

                    with st.expander("??儭??豢??寥??⊥???(?交??甇?Ⅱ隢???"):
                        st.info(f"?? ?菜葫?唬????: `{all_depts_list}`")
                        st.success(
                            f"? 甇賊???Happy Hour (HH) 銋?: `{hh_matched}`")
                        st.success(
                            f"? 甇賊???The Peak (擗輒) 銋?: `{peak_matched}`")

                        st.divider()
                        st.markdown("**?? ?祆??∟頃???敦 (DEBUG)**")
                        _debug_cols = [c for c in [date_col, dept_col, total_col] + [
                            c for c in df_month.columns if '??' in c or 'Item' in c or '?' in c
                        ] if c in df_month.columns]
                        st.caption(f"The Peak ????(??{len(df_month[df_month[dept_col].isin(peak_matched)])} 蝑??? NT$ {int(df_month[df_month[dept_col].isin(peak_matched)][total_col].apply(pd.to_numeric, errors='coerce').sum()):,})")
                        st.dataframe(df_month[df_month[dept_col].isin(peak_matched)][_debug_cols].sort_values(date_col), use_container_width=True)
                        st.caption(f"Happy Hour ????(??{len(df_month[df_month[dept_col].isin(hh_matched)])} 蝑??? NT$ {int(df_month[df_month[dept_col].isin(hh_matched)][total_col].apply(pd.to_numeric, errors='coerce').sum()):,})")
                        st.dataframe(df_month[df_month[dept_col].isin(hh_matched)][_debug_cols].sort_values(date_col), use_container_width=True)
                        st.caption(f"?? ?嗆???鞈澆???{len(df_month)} 蝑?| ?游?purchase data 銵典 {len(df_purchase)} 蝑?)

                    df_peak_purchase = df_month[df_month[dept_col].isin(
                        peak_matched)].copy()
                    df_hh_purchase = df_month[df_month[dept_col].isin(
                        hh_matched)].copy()

                    # --- ?脤??寥?嚗?券?????HH嚗?閰血????? ---
                    if df_hh_purchase.empty:
                        item_col = next((c for c in df_month.columns if any(
                            k in c for k in ['??', '?', 'Item'])), None)
                        if item_col:
                            df_hh_purchase = df_month[df_month[item_col].astype(
                                str).str.upper().str.contains('HH|HAPPY|甇⊥???', na=False)].copy()

                    # 閮?瘥?∟頃蝮賡?嚗?具誑?梁?桐???耨甇?鞈潭 vs 瘨憭梁?嚗?
                    df_daily_rest['?交?_obj'] = pd.to_datetime(
                        df_daily_rest['date']).dt.date
                    df_daily_rest['?交?_dt'] = pd.to_datetime(
                        df_daily_rest['date'])

                    def spread_weekly_cost(df_purchase, df_daily_base):
                        """撠鞈潸祥?其誑?梁?桐?嚗??文?園望?靘恥??銝憭?""
                        if df_purchase.empty or df_daily_base.empty:
                            return pd.Series(0, index=df_daily_base['?交?_obj'])

                        # ?? ISO ?勗
                        df_purchase = df_purchase.copy()
                        df_purchase['week'] = pd.to_datetime(
                            df_purchase['?交?']).dt.isocalendar().week.astype(int)
                        df_purchase['year'] = pd.to_datetime(
                            df_purchase['?交?']).dt.isocalendar().year.astype(int)
                        weekly_cost = df_purchase.groupby(['year', 'week'])[
                            '撠?'].sum().reset_index()

                        df_base = df_daily_base.copy()
                        df_base['week'] = df_base['?交?_dt'].dt.isocalendar(
                        ).week.astype(int)
                        df_base['year'] = df_base['?交?_dt'].dt.isocalendar(
                        ).year.astype(int)
                        df_base['has_guest'] = df_base['effective_peak_guests'] > 0

                        # 瘥望?靘恥?予??
                        days_per_week = df_base.groupby(['year', 'week'])[
                            'has_guest'].sum().reset_index()
                        days_per_week.columns = ['year', 'week', 'active_days']
                        days_per_week['active_days'] = days_per_week['active_days'].replace(
                            0, 1)  # ?脤??

                        # ?蔥?望???
                        df_base = pd.merge(df_base, weekly_cost, on=[
                                           'year', 'week'], how='left').fillna(0)
                        df_base = pd.merge(df_base, days_per_week, on=[
                                           'year', 'week'], how='left')
                        df_base['spread_cost'] = df_base['撠?'] / \
                            df_base['active_days']

                        return df_base.set_index('?交?_obj')['spread_cost']

                    # ?券勗??方?蝞??交???
                    peak_spread = spread_weekly_cost(
                        df_peak_purchase, df_daily_rest)
                    hh_spread = spread_weekly_cost(
                        df_hh_purchase, df_daily_rest)

                    # ?蔥靘恥?貉??勗??斗???
                    analysis_df = df_daily_rest[[
                        '?交?_obj', 'effective_peak_guests', 'rest_hh_guests', 'revenue']].copy()
                    analysis_df['peak_cost'] = analysis_df['?交?_obj'].map(
                        peak_spread).fillna(0)
                    analysis_df['hh_cost'] = analysis_df['?交?_obj'].map(
                        hh_spread).fillna(0)

                    # --- 蝝航????摩嚗?蝞?隞?蝝舐??豢? ---
                    analysis_df = analysis_df.sort_values('?交?_obj')
                    analysis_df['cum_peak_cost'] = analysis_df['peak_cost'].cumsum()
                    analysis_df['cum_peak_guests'] = analysis_df['effective_peak_guests'].cumsum(
                    )
                    analysis_df['cum_hh_cost'] = analysis_df['hh_cost'].cumsum()
                    analysis_df['cum_hh_guests'] = analysis_df['rest_hh_guests'].cumsum(
                    )

                    # 閮?蝝舐? CPG (???舐?撖衣?撟喳??韏啣)
                    analysis_df['cum_peak_cpg'] = analysis_df.apply(
                        lambda r: r['cum_peak_cost']/r['cum_peak_guests'] if r['cum_peak_guests'] > 0 else 0, axis=1)
                    analysis_df['cum_hh_cpg'] = analysis_df.apply(
                        lambda r: r['cum_hh_cost']/r['cum_hh_guests'] if r['cum_hh_guests'] > 0 else 0, axis=1)

                    # UI ? (?祆?蝮賜?)
                    total_peak_cost = analysis_df['cum_peak_cost'].iloc[-1] if not analysis_df.empty else 0
                    total_peak_guests = analysis_df['cum_peak_guests'].iloc[-1] if not analysis_df.empty else 0
                    final_peak_cpg = total_peak_cost / \
                        total_peak_guests if total_peak_guests > 0 else 0

                    total_hh_cost = analysis_df['cum_hh_cost'].iloc[-1] if not analysis_df.empty else 0
                    total_hh_guests = analysis_df['cum_hh_guests'].iloc[-1] if not analysis_df.empty else 0
                    final_hh_cpg = total_hh_cost / total_hh_guests if total_hh_guests > 0 else 0

                    # --- ?? The Peak CPG ?脩戌?蜀?? (CPG vs ???) ---
                    st.markdown("##### ?? The Peak CPG ?脩戌?蜀?? (CPG vs ???)")
                    st.caption(
                        "? **?脩戌?摰?*嚗蝝?蝺??)銝?嚗??祕蝺?CPG)?像????隞?”?∟頃?脩戌??嚗?)

                    trend_rows = []
                    for n_back in range(5, -1, -1):  # 敺?5 ????祆?
                        t_date = get_month_delta(selected_date, -n_back)
                        t_label = t_date.strftime('%Y-%m')

                        # ?府?鞈潭??
                        t_start = t_date.replace(day=1)
                        import calendar as _cal
                        _, t_last = _cal.monthrange(t_date.year, t_date.month)
                        t_end = t_date.replace(day=t_last)
                        df_t_purchase = df_purchase[(df_purchase['?交?'] >= t_start) & (
                            df_purchase['?交?'] <= t_end)].copy()

                        if not df_t_purchase.empty:
                            df_t_purchase['撠?'] = pd.to_numeric(
                                df_t_purchase[total_col], errors='coerce').fillna(0)

                        # ?府??摰Ｘ
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

                        # 蝭拚 The Peak ?∟頃
                        t_peak_cost = 0
                        if not df_t_purchase.empty and dept_col in df_t_purchase.columns:
                            t_all_depts = df_t_purchase[dept_col].astype(
                                str).unique().tolist()
                            t_hh = [d for d in t_all_depts if '4' in d or any(
                                k in d.upper() for k in ['HH', 'HAPPY', '甇⊥???'])]
                            t_peak_depts = [d for d in t_all_depts if any(
                                k in d.upper() for k in ['PEAK', '擗輒', 'THEPEAK', '擗ㄡ']) and d not in t_hh]
                            t_peak_cost = df_t_purchase[df_t_purchase[dept_col].isin(
                                t_peak_depts)]['撠?'].sum()

                        t_cpg = t_peak_cost / t_guests if t_guests > 0 else None
                        trend_rows.append(
                            {'?遢': t_label, 'CPG': t_cpg, '?格?': 150})

                    trend_df = pd.DataFrame(trend_rows).dropna(subset=['CPG'])

                    if not trend_df.empty and len(trend_df) >= 2:
                        # ??憭抒?銝西? trend_df ?蔥
                        sp_df = fetch_supplier_prices()
                        idx_df = get_market_index_df(sp_df)

                        if not idx_df.empty:
                            # ?亙???隞賣?憭??嚗?撟喳?
                            idx_monthly = idx_df.groupby('month_label')[
                                'index'].mean().reset_index()
                            trend_df = trend_df.merge(
                                idx_monthly, left_on='?遢', right_on='month_label', how='left')
                        else:
                            trend_df['index'] = None

                        base = alt.Chart(trend_df)

                        # 撌西遘嚗PG ?祕蝺?
                        cpg_line = base.mark_line(point=True, strokeWidth=3, color='#1f2c56').encode(
                            x=alt.X('?遢:N', title='?遢', sort=None),
                            y=alt.Y('CPG:Q', title='瘥恥? CPG (NT$)', scale=alt.Scale(
                                zero=False), axis=alt.Axis(titleColor='#1f2c56')),
                            tooltip=[
                                alt.Tooltip('?遢:N', title='?遢'),
                                alt.Tooltip(
                                    'CPG:Q', title='CPG (NT$)', format=',.0f')
                            ]
                        )

                        target_line = alt.Chart(pd.DataFrame({'y': [150]})).mark_rule(
                            color='#1f2c56', strokeDash=[6, 3], strokeWidth=1.5, opacity=0.5
                        ).encode(y='y:Q')
                        target_label = alt.Chart(pd.DataFrame({'y': [150], 'x': [trend_df['?遢'].iloc[-1]], 'text': ['?格? $150']})).mark_text(
                            align='right', dx=-4, dy=-8, color='#1f2c56', fontSize=11, fontWeight='bold', opacity=0.8
                        ).encode(x='x:N', y='y:Q', text='text:N')

                        cpg_layer = alt.layer(
                            cpg_line, target_line, target_label)

                        # ?唾遘嚗之?斗???蝝?蝺?
                        has_index_data = False
                        if 'index' in trend_df.columns and not trend_df['index'].isna().all():
                            valid_idx_count = trend_df['index'].notna().sum()
                            has_index_data = True

                            idx_line = base.mark_line(point={'color': '#e74c3c', 'size': 60}, strokeDash=[5, 5], strokeWidth=2, color='#e74c3c').encode(
                                x=alt.X('?遢:N', sort=None),
                                y=alt.Y('index:Q', title='??憭抒? (100=?箸?)', scale=alt.Scale(
                                    zero=False), axis=alt.Axis(titleColor='#e74c3c')),
                                tooltip=[
                                    alt.Tooltip('?遢:N', title='?遢'),
                                    alt.Tooltip(
                                        'index:Q', title='???', format=',.1f')
                                ]
                            )
                            chart = alt.layer(cpg_layer, idx_line).resolve_scale(
                                y='independent')

                            if valid_idx_count < 2:
                                st.info(
                                    "? ??嚗之?斗??賂?蝝?蝺?????隞賭?頞?2 ??嚗?甇斤??銵其???憿舐內銝???脤?暺??具??桀?????朣??餅?隞賜??鞈?隞仿＊蝷箏??湔?蝺?)
                        else:
                            chart = cpg_layer

                        st.altair_chart(chart.properties(
                            height=280), use_container_width=True)
                    else:
                        st.info("? ?閬撠?2 ??????賡＊蝷?CPG 頞典??)

                    st.divider()

                    # --- ?? ?∟頃?梯祥 vs ?拚?靘恥???賊??折?霅?隞仿梁?桐?嚗?--
                    st.markdown("##### ?? ?∟頃?梯祥 vs ?拚?靘恥???賊??折?霅??梧?")
                    st.caption(
                        "? ?拇?蝺?敶Ｙ??隅餈??氬?晞鞈潑? 靘恥???鞈潑? 靘恥??隞?”憌??抒恣?航??憿?)

                    corr_df = analysis_df[[
                        '?交?_obj', 'peak_cost', 'effective_peak_guests']].copy()
                    corr_df['?交?_dt'] = pd.to_datetime(corr_df['?交?_obj'])
                    corr_df['week'] = corr_df['?交?_dt'].dt.isocalendar(
                    ).week.astype(int)
                    corr_df['year'] = corr_df['?交?_dt'].dt.isocalendar(
                    ).year.astype(int)
                    corr_df['week_start'] = corr_df['?交?_dt'].apply(
                        lambda x: x - pd.Timedelta(days=x.dayofweek))

                    weekly_corr = corr_df.groupby('week_start').agg(
                        ?∟頃??=('peak_cost', 'sum'),
                        靘恥鈭箸=('effective_peak_guests', 'sum')
                    ).reset_index()
                    weekly_corr['?望活'] = weekly_corr['week_start'].dt.strftime(
                        'W%V\n%m/%d')

                    # 璅??? 0??00%嚗???憭批潘?
                    max_cost = weekly_corr['?∟頃??'].max()
                    max_guest = weekly_corr['靘恥鈭箸'].max()
                    weekly_corr['?∟頃(%)'] = (
                        weekly_corr['?∟頃??'] / max_cost * 100).round(1) if max_cost > 0 else 0
                    weekly_corr['靘恥(%)'] = (
                        weekly_corr['靘恥鈭箸'] / max_guest * 100).round(1) if max_guest > 0 else 0
                    weekly_corr['???有'] = (abs(
                        weekly_corr['?∟頃(%)'] - weekly_corr['靘恥(%)']) > 25).map({True: '?? ?啣虜', False: '??甇?虜'})

                    if not weekly_corr.empty and max_cost > 0 and max_guest > 0:
                        # 頧??瑟撘策 Altair
                        melt_df = weekly_corr.melt(
                            id_vars=['?望活', '???有', '?∟頃??', '靘恥鈭箸'],
                            value_vars=['?∟頃(%)', '靘恥(%)'],
                            var_name='??', value_name='璅????
                        )
                        color_map = {'?∟頃(%)': '#e67e22', '靘恥(%)': '#2980b9'}

                        corr_chart = alt.Chart(melt_df).mark_line(point=True, strokeWidth=2.5).encode(
                            x=alt.X('?望活:N', title='?望活', sort=None),
                            y=alt.Y('璅????Q', title='?詨?瘥? (% of max)',
                                    scale=alt.Scale(domain=[0, 110])),
                            color=alt.Color('??:N',
                                            scale=alt.Scale(domain=list(
                                                color_map.keys()), range=list(color_map.values())),
                                            legend=alt.Legend(
                                                title='??', orient='bottom')
                                            ),
                            tooltip=[
                                alt.Tooltip('?望活:N', title='?望活'),
                                alt.Tooltip('??:N', title='??'),
                                alt.Tooltip(
                                    '?∟頃??:Q', title='?∟頃?? (NT$)', format=',.0f'),
                                alt.Tooltip(
                                    '靘恥鈭箸:Q', title='靘恥鈭箸 (鈭?', format=',.0f'),
                                alt.Tooltip('???有:N', title='?亙熒???),
                            ]
                        ).properties(height=220)

                        st.altair_chart(corr_chart, use_container_width=True)

                        # 璅???有?望活
                        bad_weeks = weekly_corr[weekly_corr['???有'] == '?? ?啣虜']
                        if not bad_weeks.empty:
                            for _, bw in bad_weeks.iterrows():
                                diff = bw['?∟頃(%)'] - bw['靘恥(%)']
                                direction = "?∟頃??嚗?摰Ｗ?雿??眺憭芸?嚗? if diff > 0 else "靘恥??嚗?摰Ｗ?雿??眺憭芸?嚗?
                                st.warning(
                                    f"?? **{bw['?望活'].replace(chr(10), ' ')}** ?箇???有嚗direction}??∟頃 NT$ {int(bw['?∟頃??']):,} | 靘恥 {int(bw['靘恥鈭箸'])} 鈭?)
                        else:
                            st.success("???祆??望鞈潸鞎餉?靘恥鈭箸韏啣銝?湛?憌??抒恣?亙熒??)
                    else:
                        st.info("? ?祆?鞈?銝雲嚗瘜脰??賊??批???)

                    st.divider()
                    c_ana1, c_ana2 = st.columns(2)

                    with c_ana1:
                        st.markdown(
                            f"<div style='background:#f8f9fa; padding:15px; border-radius:10px; border-top:4px solid #1f2c56;'>", unsafe_allow_html=True)
                        st.markdown(f"**? The Peak (擗輒)**")
                        st.metric("?祆?蝮賣鞈潮?", f"NT$ {int(total_peak_cost):,}")
                        is_auto = "(?芸??蜇)" if (
                            df_daily_rest['rest_day_guests'].sum() == 0 and total_peak_guests > 0) else ""
                        st.metric(f"?祆?蝮賭?摰Ｘ {is_auto}",
                                  f"{int(total_peak_guests):,} 鈭?)

                        # CPG 憿霅衣內 (?格? $150)
                        peak_target = 150
                        delta_val = peak_target - final_peak_cpg
                        st.metric("撟喳?瘥恥? (CPG)", f"NT$ {int(final_peak_cpg):,}", delta=f"{int(delta_val)} (頝?格?)" if delta_val >=
                                  0 else f"{int(delta_val)} (撌脰?璅?", delta_color="normal" if delta_val >= 0 else "inverse")
                        st.markdown("</div>", unsafe_allow_html=True)

                        # --- ?啣?嚗瓷??皜祈??格??抒恣 ---
                        st.write("")
                        st.markdown("##### ? 鞎∪??格??抒恣")
                        # 1. ?雿? (擗ㄡ? / 蝮賜???
                        total_hotel_rev = m_data['rev']
                        cost_ratio = (
                            total_peak_cost / total_hotel_rev * 100) if total_hotel_rev > 0 else 0
                        st.write(f"?? ?桀??雿蜇?瘥?: **{cost_ratio:.1f}%**")

                        # 2. ???臬?葫
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
                                f"? ???摯蝮賣?? <span style='color:{forecast_color}; font-weight:bold;'>NT$ {int(forecast_total):,}</span>", unsafe_allow_html=True)
                            if final_peak_cpg > peak_target:
                                st.warning(
                                    f"?? 霅血?嚗??摰Ｘ???({int(final_peak_cpg)}) 撌脤??潛璅?{peak_target} ??隢炎閬脰疏??遢?蝞～?)
                        # ----------------------------
                    with c_ana2:
                        st.markdown(
                            f"<div style='background:#fff9f0; padding:15px; border-radius:10px; border-top:4px solid #ff9f43;'>", unsafe_allow_html=True)
                        st.markdown(f"**?? Happy Hour (HH)**")
                        st.metric("?祆?蝮賣鞈潮?", f"NT$ {int(total_hh_cost):,}")
                        st.metric("?祆?蝮賭?摰Ｘ", f"{int(total_hh_guests):,} 鈭?)
                        st.metric("撟喳?瘥恥???", f"NT$ {int(final_hh_cpg):,}")
                        if total_hh_cost > 0 and total_hh_guests == 0:
                            st.warning(
                                "?? ???HH ?∟頃鞎餌雿蜇靘恥?貊 0嚗??喋?踝? 擗輒?豢????颱誑閮?瘥恥??? (CPG)??)
                        st.markdown("</div>", unsafe_allow_html=True)

                    st.write("")
                    # 頞典?”
                    st.markdown("#### ?? ?祆?蝝航?瘥恥?頞典 (Monthly Cumulative CPG)")
                    analysis_df['?交?_str'] = analysis_df['?交?_obj'].astype(str)

                    # ?游??”
                    base_chart = alt.Chart(analysis_df).encode(
                        x=alt.X('?交?_str:O', title='?交?'))

                    peak_line = base_chart.mark_line(point=True, color='#1f2c56', strokeWidth=3).encode(
                        y=alt.Y('cum_peak_cpg:Q', title='蝝航?撟喳?? (NT$)'),
                        tooltip=['?交?_str', alt.Tooltip('cum_peak_guests', title='蝝航?靘恥'), alt.Tooltip(
                            'cum_peak_cost', title='蝝航??∟頃'), alt.Tooltip('cum_peak_cpg', format='.0f', title='蝝航? CPG')]
                    )

                    st.altair_chart(peak_line.properties(
                        title="The Peak 蝝航?撟喳??頞典", height=300), use_container_width=True)

                    if total_hh_guests > 0:
                        st.write("")
                        st.markdown("#### ?? Happy Hour 蝝航????")

                        # 憿舐內蝝航?鈭箸 vs 蝝航??
                        hh_chart_base = alt.Chart(analysis_df).encode(
                            x=alt.X('?交?_str:O', title='?交?'))

                        # ?瑟??＊蝷箇敞閮?CPG
                        hh_bar = hh_chart_base.mark_bar(color='#ff9f43', opacity=0.7).encode(
                            y=alt.Y('cum_hh_cpg:Q', title='蝝航?撟喳?? (NT$)'),
                            tooltip=[
                                '?交?_str',
                                alt.Tooltip('cum_hh_guests',
                                            title='蝝航?靘恥 (??)'),
                                alt.Tooltip('cum_hh_cost', title='蝝航??∟頃 (??)'),
                                alt.Tooltip(
                                    'cum_hh_cpg', format='.1f', title='蝝航? CPG')
                            ]
                        )

                        # ??銝璇?憿舐內蝝航?鈭箸????(蝣箔???甇?Ⅱ)
                        hh_guest_line = hh_chart_base.mark_line(color='#e67e22', strokeDash=[5, 5]).encode(
                            y=alt.Y('cum_hh_guests:Q', title='蝝航?鈭箸'),
                            tooltip=['?交?_str', alt.Tooltip(
                                'cum_hh_guests', title='蝝航?鈭箸')]
                        )

                        st.altair_chart(alt.layer(hh_bar, hh_guest_line).resolve_scale(y='independent').properties(
                            title="Happy Hour 蝝航?頞典 (?瑟?:?, ??:鈭箸)", height=300), use_container_width=True)

                    # --- ?啣?嚗??憌?瘨?瘥???(Dynamic CPG Analysis) ---
                    st.divider()
                    st.markdown("#### ? ????vs 銝?祆嚗?????瘥???)

                    # ?脣????交???
                    curr_metrics = calc_key_metrics(m_data)
                    dual_match_dates = curr_metrics.get('dual_match_dates', [])

                    if dual_match_dates:
                        # 撠??閮????
                        analysis_df['is_dual_match'] = analysis_df['?交?_str'].isin(
                            dual_match_dates)

                        df_dual = analysis_df[analysis_df['is_dual_match']]
                        df_normal = analysis_df[~analysis_df['is_dual_match']]

                        # 閮?????CPG
                        dual_peak_cost = df_dual['peak_cost'].sum()
                        dual_peak_guests = df_dual['effective_peak_guests'].sum(
                        )
                        dual_cpg = dual_peak_cost / dual_peak_guests if dual_peak_guests > 0 else 0

                        # 閮?銝?祆 CPG
                        normal_peak_cost = df_normal['peak_cost'].sum()
                        normal_peak_guests = df_normal['effective_peak_guests'].sum(
                        )
                        normal_cpg = normal_peak_cost / normal_peak_guests if normal_peak_guests > 0 else 0

                        cpg_col1, cpg_col2 = st.columns(2)

                        with cpg_col1:
                            st.markdown(f"""
                            <div style="background:#fff5e6; border-left:4px solid #e67e22; padding:15px; border-radius:8px;">
                                <p style="margin:0; font-size:13px; color:#e67e22; font-weight:bold;">?? ????(??{len(df_dual)} 憭?</p>
                                <h3 style="margin:5px 0;">NT$ {int(dual_cpg):,} / 摰?/h3>
                                <p style="margin:0; font-size:12px; color:#666;">蝮賡??鞎? NT$ {int(dual_peak_cost):,} | ??摰Ｘ: {int(dual_peak_guests):,} 鈭?/p>
                            </div>
                            """, unsafe_allow_html=True)

                        with cpg_col2:
                            st.markdown(f"""
                            <div style="background:#f8f9fa; border-left:4px solid #95a5a6; padding:15px; border-radius:8px;">
                                <p style="margin:0; font-size:13px; color:#7f8c8d; font-weight:bold;">?? 銝?祆 (??{len(df_normal)} 憭?</p>
                                <h3 style="margin:5px 0;">NT$ {int(normal_cpg):,} / 摰?/h3>
                                <p style="margin:0; font-size:12px; color:#666;">蝮賡??鞎? NT$ {int(normal_peak_cost):,} | ??摰Ｘ: {int(normal_peak_guests):,} 鈭?/p>
                            </div>
                            """, unsafe_allow_html=True)

                        # 憿舐內蝑撱箄降嚗?潭?靘???蝯?撌桀潘?
                        st.write("")
                        target_ratio = 1.10  # ????CPG ???唬??祆??110%
                        actual_ratio = (
                            dual_cpg / normal_cpg) if normal_cpg > 0 else 0
                        ratio_pct = actual_ratio * 100

                        if actual_ratio >= target_ratio:
                            st.success(
                                f"? **銝餃??蝑??嚗?* ???亦??桀恥?嚗T$ {int(dual_cpg):,}嚗??唬??祆嚗T$ {int(normal_cpg):,}嚗? **{ratio_pct:.0f}%**嚗???110% ?格??誨銵其??典之?亙???銝餃????游末???????踹敶Ｘ?甇??????)
                        elif actual_ratio >= 0.90:
                            diff_to_target = int(
                                normal_cpg * target_ratio - dual_cpg)
                            st.info(
                                f"?? **?∟頃撠銝餃?????* ????CPG ?箔??祆??{ratio_pct:.0f}%嚗璅???110%嚗?潮勗??方??箸嚗犖憭?憭拍憯? CPG嚗榆頝惇?澆???閬芋???遣霅啣???亦?勗?蝺典? NT$ {diff_to_target:,} / 鈭箏椰?喟??釭??嚗?擃垢摰Ｘ????啣榆?啜?)
                        else:
                            # 閮?銵????豢?
                            avg_normal_guests = (
                                normal_peak_guests / len(df_normal)) if len(df_normal) > 0 else 0
                            peak_target_cpg = 150  # ?格? CPG 銝?嚗??鞎∪??格?銝?湛?
                            # 撱箄降?望鞈潔????格? CPG ? 銝?祆撟喳?瘥靘恥??? 7 憭?
                            recommended_weekly_budget = int(
                                peak_target_cpg * avg_normal_guests * 7)
                            # ?祆?撖阡??勗??∟頃
                            total_weeks = max(1, round(len(df_normal) / 7))
                            actual_weekly_avg = int(
                                normal_peak_cost / total_weeks) if total_weeks > 0 else 0
                            overrun = actual_weekly_avg - recommended_weekly_budget

                            st.error(
                                f"?? **撟單憌???＊??嚗?? CPG ?銝?祆??{ratio_pct:.0f}%嚗?*\n\n"
                                f"?? **銝?祆?豢?**\n"
                                f"- 銝?祆撟喳?瘥靘恥?賂?**{avg_normal_guests:.1f} 鈭?*\n"
                                f"- 銝?祆?桀恥憌?? (CPG)嚗?*NT$ {int(normal_cpg):,}**\n\n"
                                f"? **?望鞈澆遣霅?*\n"
                                f"- 隞亦璅?CPG $150 閮?嚗遣霅唳???The Peak ?∟頃銝?嚗?*NT$ {recommended_weekly_budget:,}**\n"
                                f"- ?祆?撖阡??勗??∟頃嚗?*NT$ {actual_weekly_avg:,}**\n"
                                f"- {'? 頞撱箄降銝?嚗T$ ' + f'{overrun:,}' if overrun > 0 else '? ?函璅??'}\n\n"
                                f"?? **?航??嚗???餈賣嚗?*\n"
                                f"1. 撟單靘恥?訾???嚗◤餈怨蕭?鞈潘???嚗撠 OCC 蝣箄?嚗n"
                                f"2. 撟單????嚗??悅?勗誥嚗?瑼Ｚ?嚗n"
                                f"3. ??芰Ⅱ撖衣暺??蕭?伐?"
                            )
                    else:
                        st.info("? ?祆??桀??∠泵??隞嗥????伐??⊥??脰?撠?????)
                        st.caption(
                            "? ??隞?”蝝舐?靘恥?詻??璇??冽??蝛箇?嚗誨銵刻府?挾撠?Ｙ? HH ?賊??鞈潭?箝?)

                    st.info(
                        "? **??撠?甇?*嚗??摰Ｘ??研撣詨?擃?嚗?瑼Ｘ閰脫??行?憭批??∟頃?脣摨怠?嚗?靘恥?貉撓?交?行迤蝣箝?)

                else:
                    st.info("撠?菜葫?唳??擗輒靘恥?豢?嚗瘜脰????????)

                st.divider()

                # 3. ??閰喟敦蝯梯?
                st.subheader("? ??蝬祥??")

                # ?????
                departments = dept_summary.sort_values(
                    '撠?', ascending=False)['?券?'].tolist()

                for dept in departments:
                    dept_df = df_month[df_month[dept_col] == dept].copy()
                    dept_total = dept_df['撠?'].sum()

                    with st.expander(f"?? {dept} (蝮質?: NT$ {int(dept_total):,})", expanded=False):
                        # --- ?啣?嚗op 5 擃?????璁?---
                        item_name_col = next((c for c in dept_df.columns if any(
                            k in c for k in ['??', '?', 'Item'])), None)
                        if item_name_col:
                            st.markdown("##### ?? ????憿鞈澆???)
                            top_items = dept_df.groupby(item_name_col)['撠?'].sum(
                            ).sort_values(ascending=False).head(5).reset_index()
                            t_cols = st.columns(5)
                            for idx, row in top_items.iterrows():
                                with t_cols[idx]:
                                    st.metric(
                                        f"No.{idx+1} {row[item_name_col][:8]}", f"NT$ {int(row['撠?']):,}")
                        st.divider()

                        # --- ?啣?嚗?摨??---
                        sort_by = st.selectbox(f"???孵? ({dept})", [
                                               "?交? (?售???", "?? (擃?雿?", "?? (雿?擃?", "???迂"], key=f"sort_{dept}")

                        if sort_by == "?? (擃?雿?":
                            dept_df = dept_df.sort_values(
                                '撠?', ascending=False)
                        elif sort_by == "?? (雿?擃?":
                            dept_df = dept_df.sort_values('撠?', ascending=True)
                        elif sort_by == "?交? (?售???":
                            dept_df = dept_df.sort_values(
                                '?交?', ascending=False)
                        elif sort_by == "???迂" and item_name_col:
                            dept_df = dept_df.sort_values(item_name_col)

                        # 憿舐內閰脤?銵冽
                        cols_to_show = [c for c in [
                            '?交?', '靘???, '??', '閬', '?賊?', '?桐?', '?桀', '撠?'] if c in dept_df.columns]
                        if not cols_to_show:
                            cols_to_show = dept_df.columns.tolist()

                        st.dataframe(
                            dept_df[cols_to_show],
                            use_container_width=True,
                            hide_index=True
                        )

                # --- ? 4. ?桀?憌?瘨??移皞鞈潭獢???---
                if 'analysis_df' in locals() and not analysis_df.empty:
                    st.divider()
                    st.subheader("? ?桀?憌?瘨??移皞鞈潭獢???)
                    st.caption(
                        "???孵??憌???嚗?嚗???暻??絲擙桃?嚗?瘥恥撟喳?瘨?嚗蒂?芸??Ｗ蝎暹??怨疏??撱箄降??)

                    item_col = next((c for c in df_month.columns if any(
                        k in c for k in ['??', '?', 'Item'])), None)
                    qty_col = next((c for c in df_month.columns if any(
                        k in c for k in ['?賊?', 'Qty', 'Quantity'])), None)
                    unit_col = next((c for c in df_month.columns if any(
                        k in c for k in ['?桐?', 'Unit'])), None)
                    price_col = next((c for c in df_month.columns if any(
                        k in c for k in ['?桀', 'Price', 'Rate'])), None)

                    if item_col and qty_col:
                        # ?瑕?撣貊?摮??
                        all_items = df_month[item_col].dropna().astype(
                            str).str.strip()
                        all_items = all_items[all_items != ""]

                        common_keywords = ["??, "??, "??, "憟?,
                                           "蝐?, "暻?, "瘝?, "瘚琿悅", "??, "鞊?, "??, "擳?]
                        found_keywords = [k for k in common_keywords if any(
                            k in x for x in all_items)]
                        if not found_keywords:
                            found_keywords = ["??]

                        c_sel1, c_sel2 = st.columns([1, 1])
                        with c_sel1:
                            selected_keyword = st.selectbox(
                                "?? ?豢??????摮?,
                                options=found_keywords + ["(?芾?頛詨)"],
                                index=0,
                                key="item_analysis_keyword_select"
                            )
                        with c_sel2:
                            if selected_keyword == "(?芾?頛詨)":
                                search_term = st.text_input(
                                    "?? 頛詨?芾?憌??迂 (靘?: 擃???", "??, key="item_analysis_custom_input")
                            else:
                                search_term = selected_keyword

                        # 蝭拚?寥??鞈潮???
                        item_mask = df_month[item_col].astype(
                            str).str.contains(search_term, na=False, case=False)
                        item_df = df_month[item_mask].copy()

                        if not item_df.empty:
                            # ?詨潭???
                            item_df['cleaned_qty'] = pd.to_numeric(item_df[qty_col].astype(
                                str).str.replace(',', ''), errors='coerce').fillna(0)
                            item_df['cleaned_total'] = pd.to_numeric(item_df['撠?'].astype(
                                str).str.replace(',', ''), errors='coerce').fillna(0)

                            # ?桐??斗
                            most_common_unit = "?桐?"
                            if unit_col in item_df.columns:
                                most_common_unit = item_df[unit_col].mode(
                                ).iloc[0] if not item_df[unit_col].empty else "?桐?"

                            # 瘥?∟頃?游?
                            item_df['?交?_obj'] = pd.to_datetime(
                                item_df['?交?']).dt.date
                            daily_item_qty = item_df.groupby(
                                '?交?_obj')['cleaned_qty'].sum().reset_index()
                            daily_item_cost = item_df.groupby(
                                '?交?_obj')['cleaned_total'].sum().reset_index()

                            # ?蔥瘥靘恥
                            item_analysis_df = analysis_df[[
                                '?交?_obj', 'effective_peak_guests']].copy()
                            item_analysis_df = pd.merge(
                                item_analysis_df, daily_item_qty, on='?交?_obj', how='left').fillna(0)
                            item_analysis_df = pd.merge(
                                item_analysis_df, daily_item_cost, on='?交?_obj', how='left').fillna(0)

                            # ?勗?蝮質?蝞?
                            item_analysis_df['?交?_dt'] = pd.to_datetime(
                                item_analysis_df['?交?_obj'])
                            item_analysis_df['week_start'] = item_analysis_df['?交?_dt'].apply(
                                lambda x: x - pd.Timedelta(days=x.dayofweek))

                            weekly_item = item_analysis_df.groupby('week_start').agg(
                                蝮賣鞈潮?=('cleaned_qty', 'sum'),
                                蝮質祥??('cleaned_total', 'sum'),
                                靘恥鈭箸=('effective_peak_guests', 'sum')
                            ).reset_index()

                            weekly_item['?望活'] = pd.to_datetime(
                                weekly_item['week_start']).dt.strftime('W%V\n%m/%d')
                            weekly_item['瘥恥撟喳?瘨?'] = weekly_item.apply(
                                lambda r: r['蝮賣鞈潮?'] / r['靘恥鈭箸'] if r['靘恥鈭箸'] > 0 else 0, axis=1
                            )

                            # 閮??像??撟喳??桀
                            total_qty_month = weekly_item['蝮賣鞈潮?'].sum()
                            total_guests_month = weekly_item['靘恥鈭箸'].sum()
                            avg_rate_month = total_qty_month / \
                                total_guests_month if total_guests_month > 0 else 0
                            avg_unit_price = item_df['cleaned_total'].sum(
                            ) / item_df['cleaned_qty'].sum() if item_df['cleaned_qty'].sum() > 0 else 0

                            st.write("")
                            st.markdown(f"##### ?? **?search_term}?????璅?*")

                            c_m1, c_m2, c_m3 = st.columns(3)
                            c_m1.metric(
                                "?祆?蝮賣鞈潮?", f"{total_qty_month:,.1f} {most_common_unit}")
                            c_m2.metric(
                                "瘥恥撟喳?瘨? (雿輻??", f"{avg_rate_month:.2f} {most_common_unit}/鈭?, help="蝮賣鞈潮? / 蝮賭?摰Ｘ")
                            c_m3.metric(
                                "撟喳??∟頃?桀", f"NT$ {avg_unit_price:,.1f} /{most_common_unit}")

                            # ?”?
                            st.write("")
                            st.markdown(
                                f"###### ?? ?曹?摰Ｘ vs ?search_term}?鞈潮??詨?韏啣")

                            max_w_qty = weekly_item['蝮賣鞈潮?'].max()
                            max_w_guests = weekly_item['靘恥鈭箸'].max()
                            weekly_item['?∟頃??%)'] = (
                                weekly_item['蝮賣鞈潮?'] / max_w_qty * 100).round(1) if max_w_qty > 0 else 0
                            weekly_item['靘恥(%)'] = (
                                weekly_item['靘恥鈭箸'] / max_w_guests * 100).round(1) if max_w_guests > 0 else 0

                            melt_item_df = weekly_item.melt(
                                id_vars=['?望活', '蝮賣鞈潮?', '靘恥鈭箸', '蝮質祥??],
                                value_vars=['?∟頃??%)', '靘恥(%)'],
                                var_name='??', value_name='璅????
                            )

                            item_color_map = {
                                '?∟頃??%)': '#e67e22', '靘恥(%)': '#2980b9'}

                            item_chart = alt.Chart(melt_item_df).mark_line(point=True, strokeWidth=2.5).encode(
                                x=alt.X('?望活:N', title='?望活', sort=None),
                                y=alt.Y('璅????Q', title='?詨?瘥? (% of max)',
                                        scale=alt.Scale(domain=[0, 110])),
                                color=alt.Color('??:N',
                                                scale=alt.Scale(domain=list(item_color_map.keys()), range=list(
                                                    item_color_map.values())),
                                                legend=alt.Legend(
                                                    title='??', orient='bottom')
                                                ),
                                tooltip=[
                                    alt.Tooltip('?望活:N', title='?望活'),
                                    alt.Tooltip('??:N', title='??'),
                                    alt.Tooltip(
                                        '蝮賣鞈潮?:Q', title=f'蝮賣鞈潮? ({most_common_unit})', format=',.1f'),
                                    alt.Tooltip(
                                        '靘恥鈭箸:Q', title='靘恥鈭箸 (鈭?', format=',.0f'),
                                    alt.Tooltip(
                                        '蝮質祥??Q', title='蝮質祥??(NT$)', format=',.0f'),
                                ]
                            ).properties(height=200)

                            st.altair_chart(
                                item_chart, use_container_width=True)

                            # ? ?∟頃?寞?蝎曄?
                            st.write("")
                            st.markdown(f"##### ? ?search_term}?移皞鞈潭獢?蝞?")
                            st.write("閮剖??冽靘???靘恥?賂?蝟餌絞??鼠?冽蝞????鞈潮??鞎冽?蝔遣霅啜?)

                            col_calc1, col_calc2 = st.columns([1, 1])
                            with col_calc1:
                                input_guests = st.number_input(
                                    "?? ?芯?銝?梢?閮蜇靘恥??,
                                    min_value=10,
                                    max_value=5000,
                                    value=int(
                                        total_guests_month / 4) if total_guests_month > 0 else 500,
                                    step=50,
                                    key="item_calc_guests_input_widget"
                                )

                                # ???怨疏?望??內
                                vendor_type = "??"
                                if any(x in search_term for x in ["??, "??]):
                                    vendor_type = "??"
                                elif any(x in search_term for x in ["??, "??, "鞊?, "??, "擳?]):
                                    vendor_type = "??"
                                elif any(x in search_term for x in ["??, "瘝?, "蝐?, "暻?]):
                                    vendor_type = "?疏"

                                st.info(
                                    f"? **撱箄降???? ({vendor_type})**\n\n"
                                    f"- ?芸????舫甇Ｗ甈⊿脰疏??憭批??湔擙桀漲銝??撱Ｘ??n"
                                    f"- ?臭??曇?撖阡??怨疏?望?敶扯矽?游鞎具?
                                )

                            with col_calc2:
                                # ? 5% 摰摨怠?蝺抵?
                                recommended_qty = input_guests * avg_rate_month * 1.05
                                est_cost = recommended_qty * avg_unit_price

                                st.markdown(
                                    f"<div style='background:#2e437c15; border-left:4px solid #2e437c; padding:15px; border-radius:8px;'>"
                                    f"<h4 style='margin:0; color:#2e437c;'>撱箄降?∟頃蝮賡?</h4>"
                                    f"<h2 style='margin:5px 0; color:#2e437c;'>{recommended_qty:,.1f} {most_common_unit}</h2>"
                                    f"<p style='margin:0; font-size:12px; color:#666;'>撌脣???5% 摰摨怠?蝺抵?</p>"
                                    f"<hr style='margin:10px 0; border:none; border-top:1px solid #ddd;'>"
                                    f"<h5 style='margin:0; color:#333;'>?摯?∟頃鞎餌: <strong style='font-size:18px;'>NT$ {int(est_cost):,}</strong></h5>"
                                    f"</div>",
                                    unsafe_allow_html=True
                                )

                                # ?梢??望?瘥?撱箄降
                                st.markdown("?? **?怨疏?望???瘥??*")
                                if vendor_type == "??":
                                    st.markdown("- **?曹? (40%)**嚗遣霅唳鞈?**`{:.1f}`** {}嚗?隡啣甈∟祥?剁?**NT$ {:,}**嚗?.format(
                                        recommended_qty * 0.4, most_common_unit, int(est_cost * 0.4)))
                                    st.markdown("- **?曹? (30%)**嚗遣霅唳鞈?**`{:.1f}`** {}嚗?隡啣甈∟祥?剁?**NT$ {:,}**嚗?.format(
                                        recommended_qty * 0.3, most_common_unit, int(est_cost * 0.3)))
                                    st.markdown("- **?曹? (30%)**嚗遣霅唳鞈?**`{:.1f}`** {}嚗?隡啣甈∟祥?剁?**NT$ {:,}**嚗?.format(
                                        recommended_qty * 0.3, most_common_unit, int(est_cost * 0.3)))
                                elif vendor_type == "??":
                                    st.markdown("- **撟單瘥? (60%)**嚗?甈∪鞎典遣霅?**`{:.1f}`** {}嚗?隡啣甈∟祥?剁?**NT$ {:,}**嚗?.format(
                                        recommended_qty * 0.12, most_common_unit, int(est_cost * 0.12)))
                                    st.markdown("- **?曹??撥 (40%)**嚗?甈∪頞?**`{:.1f}`** {}嚗?隡啣甈∟祥?剁?**NT$ {:,}**嚗?.format(
                                        recommended_qty * 0.4, most_common_unit, int(est_cost * 0.4)))
                                else:
                                    st.markdown("- **?格活頞喲??∟頃 (100%)**嚗?曹???蝝鞎冽銝甈⊥扳鞈?**`{:.1f}`** {}嚗?隡啣甈∟祥?剁?**NT$ {:,}**嚗?.format(
                                        recommended_qty, most_common_unit, int(est_cost)))

                        else:
                            st.warning(
                                f"?? ?函???∟頃鞈?銝哨??曆??啣?search_term}?????迂??)
                            st.info("? 隢?閰阡?隞虜?券??萄?嚗??芾?頛詨?渡移蝣箇??摮?憒?????暻?嚗?)
                    else:
                        st.info("? ?∟頃??蝻箏????????雿??⊥??脰??桀?瘨?????)

            else:
                st.info(f"? {current_month_str} 撠?鞈潭????)
                st.write(
                    f"?對? ?具?*{used_name}**???葉蝮賢?潛 {len(df_purchase)} 蝑???雿??泵??{current_month_str} ????)
                with st.expander("??儭?暺迨?亦???銝剔???5 蝑?憪???(?日??"):
                    st.write(df_purchase.head(5))
        else:
            st.warning(
                f"?? ?⊥???Google Sheet 銝剜?唳鞈澆???(?岫?? {', '.join(possible_names)})??)
            st.info("? 隢Ⅱ隤???蝔望?行迤蝣綽?銝??葉?喳?撌脣‵?乩?銵???)

    except Exception as e:
        if "WorksheetNotFound" in str(e):
            st.error(f"???曆??唳鞈潛????隢Ⅱ隤?Google Sheet 銝剔????迂嚗? purchase data嚗?)
        else:
            st.error(f"霈?鞈潭??? {e}")
        import traceback
        st.expander("?航炊閰喟敦鞈?").code(traceback.format_exc())

# --- ?? ?祆??乩?靘??望鞈潮?憿遣霅堆??函??憛?銝?鞈湔鞈潭??---
with tab_p:
    st.divider()
    st.markdown("#### ?? ?祆??乩?靘??望鞈潮?憿遣霅?)

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
        # 1. ?芸??冽??撱單??
        fw_m_data = fetch_month_summary(
            selected_date.year, selected_date.month)
        fw_df = fw_m_data.get('df', pd.DataFrame())
        avg_fw = 0
        fw_label = ''
        if not fw_df.empty and 'bf_total_act' in fw_df.columns:
            fw_df = fw_df.copy()
            fw_df['_bf'] = pd.to_numeric(fw_df['bf_total_act'].astype(
                str).str.replace(',', ''), errors='coerce').fillna(0)
            active_fw = fw_df[fw_df['_bf'] > 0]['_bf']
            if not active_fw.empty:
                avg_fw = active_fw.mean()
                fw_label = f"?祆?撖阡? ({len(active_fw)} 憭抵???"

        # 2. ?嚗?其????拚?靘恥??
        if avg_fw == 0:
            fw_prev = fetch_month_summary(
                m_prev['year'], m_prev['month']) if 'year' in m_prev else {}
            fw_prev_df = fw_prev.get(
                'df', pd.DataFrame()) if fw_prev else pd.DataFrame()
            if not fw_prev_df.empty and 'bf_total_act' in fw_prev_df.columns:
                fw_prev_df = fw_prev_df.copy()
                fw_prev_df['_bf'] = pd.to_numeric(fw_prev_df['bf_total_act'].astype(
                    str).str.replace(',', ''), errors='coerce').fillna(0)
                prev_active = fw_prev_df[fw_prev_df['_bf'] > 0]['_bf']
                if not prev_active.empty:
                    avg_fw = prev_active.mean()
                    fw_label = f"?? 隞乩??像?隡堆??祆?撠擗輒鞈?嚗?

        # 3. ???交??殷?靘 tab_m ??calc_key_metrics嚗?
        fw_curr_metrics = calc_key_metrics(fw_m_data)
        fw_dual_dates = set(fw_curr_metrics.get('dual_match_dates', []))

        # 4. ?梁???
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
                wdates = [(ws + dt_timedelta(days=i)).strftime('%Y-%m-%d')
                          for i in range((we - ws).days + 1)]
                has_d = any(d in fw_dual_dates for d in wdates)
                days_cnt = len(wdates)
                fw_week_plans.append({
                    'label': f"{ws.strftime('%m/%d')} 嚚?{we.strftime('%m/%d')}",
                    'has_dual': has_d,
                    'recommended': int((150 * 1.15 if has_d else 150) * avg_fw * days_cnt),
                    'dual_labels': [d[5:] for d in wdates if d in fw_dual_dates],
                    'days_cnt': days_cnt,
                })
            fw_cursor = sun + dt_timedelta(days=1)

        if avg_fw > 0 and fw_week_plans:
            st.caption(
                f"? ?摯?箸?嚗??亙像??摰Ｘ **{avg_fw:.1f} 鈭?*嚗fw_label}嚗??望鞈潔????擃?15%??)
            for wp in fw_week_plans:
                color = '#e67e22' if wp['has_dual'] else '#2980b9'
                dual_note = f"?? ?恍??嚗', '.join(wp['dual_labels'])}" if wp['has_dual'] else ""
                c1, c2 = st.columns([2, 1])
                c1.markdown(f"**{wp['label']}**{dual_note}")
                c2.markdown(
                    f"<div style='background:{color}22; border-left:3px solid {color}; padding:8px 12px; border-radius:6px; text-align:center;'>"
                    f"<strong style='font-size:16px;'>NT$ {wp['recommended']:,}</strong>"
                    f"<br><span style='font-size:11px; color:#666;'>撱箄降?望鞈潔???({wp['days_cnt']}憭?</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
        else:
            st.info("? 撠頞喳???摰Ｘ??隡啁??望鞈潮?蝞?蝣箄?銝???撱單擗?摰Ｘ撌脣‵撖怒?)
    else:
        st.info("? ?望鞈澆遣霅啜??拍?潛???芯??遢??)

# =====================================================
# ?? ??? tab_s
# =====================================================
with tab_s:
    st.header("?? ???")
    sp_df = fetch_supplier_prices()

    if sp_df.empty:
        st.warning(
            "?? 撠霈??鞈???蝣箄? Google Sheets 銝剖歇撱箇? `supplier_prices` ??嚗? `period`?item_name`?unit`?price` 甈?撌脣‵撖怒?)
    else:
        periods_available = sorted(sp_df['period_dt'].unique())
        periods_str = [str(p) for p in periods_available]
        n_periods = len(periods_available)

        st.caption(
            f"?? ?桀??望? **{n_periods}** ???寡???{periods_str[0]} 嚚?{periods_str[-1]}嚗???{len(sp_df['item_name'].unique())} ????)

        # ?? A. ???拙? ??????????????????????????????
        if n_periods >= 2:
            st.markdown("#### ?? A. ???拙?")
            st.info(
                "? **隞暻潭憭抒?嚗?* 隞亦洵銝???箸????擃?寧 100 ??????貊 105嚗誨銵券ㄞ摨擃?憌??∟頃??鈭?5%??*?????隢?矽??CPG ????撘瑕恥閫靘?嚗?*")

            index_df = get_market_index_df(sp_df)
            base_period = periods_available[0]

            if not index_df.empty:
                latest_idx = index_df.iloc[-1]['index']
                prev_idx = index_df.iloc[-2]['index']
                diff_idx = latest_idx - prev_idx

                ic1, ic2 = st.columns([1, 3])
                with ic1:
                    st.metric(label="?祆?憭抒?", value=f"{latest_idx:.1f}",
                              delta=f"{diff_idx:+.1f} 暺?(vs銝?)", delta_color="inverse")
                    st.caption(f"?箸???{base_period} (=100)")
                with ic2:
                    # ??????銝???銝血???100
                    all_idx_vals = index_df['index'].tolist() + [100]
                    idx_min = max(0, int(min(all_idx_vals) * 0.98))
                    idx_max = int(max(all_idx_vals) * 1.02)

                    line_chart = alt.Chart(index_df).mark_line(point=True, strokeWidth=3, color='#e74c3c').encode(
                        x=alt.X('period_str:O', title='?',
                                axis=alt.Axis(labelAngle=-30)),
                        y=alt.Y('index:Q', title='?', scale=alt.Scale(
                            domain=[idx_min, idx_max], zero=False)),
                        tooltip=[
                            alt.Tooltip('period_str:N', title='?'),
                            alt.Tooltip('index:Q', title='憭抒?', format='.1f'),
                        ]
                    )

                    base_line = alt.Chart(pd.DataFrame({'y': [100]})).mark_rule(
                        strokeDash=[5, 5], color='gray').encode(y='y:Q')
                    st.altair_chart(
                        (base_line + line_chart).properties(height=250), use_container_width=True)
                st.divider()
            else:
                index_df = pd.DataFrame()

        # ?? B. ?祆?憌?摰??蝭? ??????????????????????????????
        st.markdown("#### ? B. ?祆?憌?摰??蝭?")

        # ???祆?蝮賣擗犖??(敺?m_curr)
        total_bf_guests = 0
        current_month_str = ""
        if 'm_curr' in locals() and m_curr.get('df') is not None and not m_curr['df'].empty:
            m_df = m_curr['df']
            current_month_str = m_curr.get('month_label', '')
            for _, r in m_df.iterrows():
                act = pd.to_numeric(r.get('bf_total_act', 0), errors='coerce')
                est = pd.to_numeric(r.get('bf_total_est', 0), errors='coerce')
                if pd.isna(act):
                    act = 0
                if pd.isna(est):
                    est = 0
                total_bf_guests += act if act > 0 else est

        if total_bf_guests > 0:
            bc1, bc2 = st.columns([1, 2])
            with bc1:
                target_cpg = st.number_input(
                    "? ?格? CPG (?桀恥???)", min_value=0, value=150, step=5, help="?身??150 ??銝餌恣?臭??之?斗??詨??扳撖祆?蝺葬??)
                total_budget = total_bf_guests * target_cpg
                st.metric(f"?祆??摯蝮賢?擗?({current_month_str})",
                          f"{int(total_bf_guests):,} 鈭?)
                st.metric("?祆?憌?蝮賡?蝞?(Budget)",
                          f"${int(total_budget):,}", help="蝮賢?擗犖??? ?格? CPG")

            with bc2:
                # ?寞? latest_idx 瘙箏???
                if n_periods >= 2:
                    if latest_idx > 105:
                        status = "? 撣??∪? (憭抒 > 105)"
                        def_pct, norm_pct, risk_pct = 0.70, 0.20, 0.10
                        advice = "撘瑞?撱箄降撠?70% ?????具??脩戌?憸冽葛憌?嚗?潮?蝮桅?憸券憌??∟頃??
                    elif latest_idx < 95:
                        status = "?? 撣雿? (憭抒 < 95)"
                        def_pct, norm_pct, risk_pct = 0.40, 0.30, 0.30
                        advice = "?桀??鞎瑟撣嚗? 30% ????摨鞎典??砍?鞎渡?擃◢?芷???
                    else:
                        status = "?? 撣撟喟帘"
                        def_pct, norm_pct, risk_pct = 0.50, 0.30, 0.20
                        advice = "撣瘜Ｗ?甇?虜嚗雁??皞?5:3:2 ?∟頃瘥???
                else:
                    status = "??
                    def_pct, norm_pct, risk_pct = 0.50, 0.30, 0.20
                    advice = "鞈?銝雲嚗雁??皞?靘?

                st.markdown(f"**撣?文?嚗status}**")
                st.caption(f"? ?啁撱箄降嚗advice}")

                # ?恍脣漲璇?
                st.markdown(f"""
                <div style='display:flex; height:24px; border-radius:12px; overflow:hidden; margin-bottom:10px;'>
                    <div style='width:{def_pct*100}%; background-color:#2ecc71; display:flex; align-items:center; justify-content:center; color:white; font-size:12px; font-weight:bold;'>?脩戌 {def_pct*100:.0f}%</div>
                    <div style='width:{norm_pct*100}%; background-color:#f1c40f; display:flex; align-items:center; justify-content:center; color:#333; font-size:12px; font-weight:bold;'>銝??{norm_pct*100:.0f}%</div>
                    <div style='width:{risk_pct*100}%; background-color:#e74c3c; display:flex; align-items:center; justify-content:center; color:white; font-size:12px; font-weight:bold;'>憸券 {risk_pct*100:.0f}%</div>
                </div>
                """, unsafe_allow_html=True)

                # ????
                col_d, col_n, col_r = st.columns(3)
                col_d.metric("?儭??輸◢皜舫?憿?, f"${int(total_budget * def_pct):,}")
                col_n.metric("?布 銝?祇?憿?, f"${int(total_budget * norm_pct):,}")
                col_r.metric("? 擃◢?芷?憿?, f"${int(total_budget * risk_pct):,}")

        else:
            st.info("? ?桀??亦?祆??摯蝮賣擗犖?賂??⊥?閰衣?蝮賡?蝞?蝣箄?銝???歇頛?嗆?鞈???)

        st.divider()

        # ?? C. ?祆??蝮質汗 ??????????????????????????????
        st.markdown("#### ?? C. ?祆??蝮質汗")
        latest_period = periods_available[-1]
        latest_df = sp_df[sp_df['period_dt'] == latest_period].copy()

        if n_periods >= 2:
            prev_period = periods_available[-2]
            prev_df = sp_df[sp_df['period_dt'] == prev_period][[
                'item_name', 'unit', 'price']].rename(columns={'price': 'prev_price'})
            latest_df = latest_df.merge(
                prev_df, on=['item_name', 'unit'], how='left')
            latest_df['change'] = latest_df['price'] - latest_df['prev_price']
            latest_df['change_pct'] = (
                latest_df['change'] / latest_df['prev_price'] * 100).round(1)

            # --- ?啣?嚗?蝞?撟渲隞??冽風????撟喳皞?璆萄?---
            ytd_stats = sp_df.groupby(['item_name', 'unit'])['price'].agg(
                ['mean', 'max', 'min']).reset_index()
            ytd_stats.rename(
                columns={'mean': 'ytd_avg', 'max': 'ytd_max', 'min': 'ytd_min'}, inplace=True)
            latest_df = latest_df.merge(
                ytd_stats, on=['item_name', 'unit'], how='left')

            def fmt_change(row):
                if pd.isna(row.get('change')):
                    return '??
                sign = '+' if row['change'] > 0 else ''
                color = '#e74c3c' if row['change'] > 0 else (
                    '#2ecc71' if row['change'] < 0 else '#888')
                arrow = '?? if row['change'] > 0 else (
                    '?? if row['change'] < 0 else '?')
                return f"<span style='color:{color};font-weight:bold;'>{arrow} {sign}{row['change_pct']:.1f}%</span>"

            latest_df['瞍脰?'] = latest_df.apply(fmt_change, axis=1)

            def fmt_ytd(row):
                if pd.isna(row.get('ytd_avg')):
                    return '??
                avg_p = row['ytd_avg']
                curr_p = row['price']
                if pd.isna(curr_p) or avg_p == 0:
                    return '??

                # 閮??僑撟喳??榆頝?%
                diff_pct = ((curr_p - avg_p) / avg_p * 100)
                sign = '+' if diff_pct > 0 else ''
                color = '#e74c3c' if diff_pct > 0 else (
                    '#2ecc71' if diff_pct < 0 else '#888')
                text = f"<span style='color:{color};'>{sign}{diff_pct:.1f}%</span>"

                # 璆萄澆噬蝡?(??芣?1??霈?????儔璅惜)
                badges = ""
                if row['ytd_max'] > row['ytd_min']:
                    if curr_p >= row['ytd_max']:
                        badges = " <span style='background:#e74c3c;color:white;font-size:10px;padding:2px 4px;border-radius:4px;margin-left:4px;'>甇瑕擃?</span>"
                    elif curr_p <= row['ytd_min']:
                        badges = " <span style='background:#2ecc71;color:white;font-size:10px;padding:2px 4px;border-radius:4px;margin-left:4px;'>甇瑕雿?</span>"

                return f"??{avg_p:.1f} ({text}){badges}"

            latest_df['蝝像?箸?撠'] = latest_df.apply(fmt_ytd, axis=1)
            display_cols = ['item_name', 'unit', 'price', '瞍脰?', '蝝像?箸?撠']
            col_rename = {'item_name': '??', 'unit': '?桐?',
                          'price': f'?祆??桀 ({latest_period})'}
        else:
            display_cols = ['item_name', 'unit', 'price']
            col_rename = {'item_name': '??', 'unit': '?桐?',
                          'price': f'?祆??桀 ({latest_period})'}

        show_df = latest_df[display_cols].rename(columns=col_rename)

        # ???蕪
        search_kw = st.text_input("?? ????", placeholder="頛詨?摮?憒?擃???)
        if search_kw:
            show_df = show_df[show_df['??'].str.contains(search_kw, na=False)]

        if n_periods >= 2:
            st.write(show_df.to_html(escape=False, index=False),
                     unsafe_allow_html=True)
        else:
            st.dataframe(show_df, use_container_width=True, hide_index=True)

        st.divider()

        # ?? D. ?祆? vs 銝?瞍脰??? ??????????????????????
        if n_periods >= 2:
            st.markdown("#### ?? D. ?祆? vs 銝?嚗撞頝?銵?)
            ranked = latest_df.dropna(subset=['change']).copy()
            ranked = ranked.sort_values('change_pct', ascending=False)

            bc1, bc2 = st.columns(2)
            with bc1:
                st.markdown("**? 瞍脣??憭?Top 5**")
                top_up = ranked.head(5)
                for _, r in top_up.iterrows():
                    if r['change_pct'] > 0:
                        st.markdown(
                            f"<div style='background:#fdf2f2; border-left:4px solid #e74c3c; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#e74c3c; font-size:13px;'>??+{r['change_pct']:.1f}%</span>"
                            f"<br><span style='font-size:12px; color:#888;'>{r['prev_price']:.0f} ??{r['price']:.0f} ??{r['unit']}</span></div>",
                            unsafe_allow_html=True
                        )
            with bc2:
                st.markdown("**? 頝??憭?Top 5**")
                top_down = ranked.tail(5).iloc[::-1]
                for _, r in top_down.iterrows():
                    if r['change_pct'] < 0:
                        st.markdown(
                            f"<div style='background:#f2fdf5; border-left:4px solid #2ecc71; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#2ecc71; font-size:13px;'>??{r['change_pct']:.1f}%</span>"
                            f"<br><span style='font-size:12px; color:#888;'>{r['prev_price']:.0f} ??{r['price']:.0f} ??{r['unit']}</span></div>",
                            unsafe_allow_html=True
                        )
            st.divider()

        # ?? E. ??頞典??????????????????????????????????
        if n_periods >= 2:
            st.markdown("#### ?? E. ??甇瑟?頞典")
            all_items = sorted(sp_df['item_name'].unique().tolist())
            selected_items = st.multiselect(
                "?豢???嚗憭嚗?,
                options=all_items,
                default=all_items[:3] if len(all_items) >= 3 else all_items,
                placeholder="隢??瘥?????
            )
            if selected_items:
                trend_df = sp_df[sp_df['item_name'].isin(
                    selected_items)].copy()
                trend_df['period_str'] = trend_df['period_dt'].astype(str)
                price_min = int(trend_df['price'].min() * 0.85)
                price_max = int(trend_df['price'].max() * 1.15)
                trend_chart = alt.Chart(trend_df).mark_line(point=True, strokeWidth=2).encode(
                    x=alt.X('period_str:O', title='?',
                            axis=alt.Axis(labelAngle=-30)),
                    y=alt.Y('price:Q', title='?桀', scale=alt.Scale(
                        domain=[price_min, price_max], zero=False)),
                    color=alt.Color(
                        'item_name:N', legend=alt.Legend(title='??')),
                    tooltip=[
                        alt.Tooltip('period_str:N', title='?'),
                        alt.Tooltip('item_name:N', title='??'),
                        alt.Tooltip('price:Q', title='?桀', format='.1f'),
                        alt.Tooltip('unit:N', title='?桐?'),
                    ]
                ).properties(height=380)
                st.altair_chart(trend_chart, use_container_width=True)
            st.divider()
        else:
            st.info("? ?桀??芣?銝????蝝舐?銝????桀??喳?亦?頞典??瞍脰?瘥???)
            st.divider()

        # ?? F. ?怨疏?啁撱箄降 ??????????????????????????????
        st.markdown("#### ? F. ?怨疏?啁撱箄降")
        if n_periods >= 2:
            ranked_all = latest_df.dropna(subset=['change_pct']).copy()
            # ??瞍脣嚗撞撟?> 5%
            alert_up = ranked_all[ranked_all['change_pct'] > 5].sort_values(
                'change_pct', ascending=False)
            # ?＊?嚗?撟?> 5%
            alert_down = ranked_all[ranked_all['change_pct']
                                    < -5].sort_values('change_pct')

            # 甇瑕憭拙霅血 (?桀??寞蝑?冽風?脫?擃嚗?甇瑕?郭??
            alert_all_time_high = ranked_all[(ranked_all['price'] >= ranked_all['ytd_max']) & (
                ranked_all['ytd_max'] > ranked_all['ytd_min'])]
            # 甇瑕雿?霅血
            alert_all_time_low = ranked_all[(ranked_all['price'] <= ranked_all['ytd_min']) & (
                ranked_all['ytd_max'] > ranked_all['ytd_min'])]

            if not alert_all_time_high.empty:
                high_items = '??.join(
                    alert_all_time_high['item_name'].head(5).tolist())
                st.error(
                    f"? **?風?脤?暺郎?晞?*嚗high_items} ?桀??箔?撟湔?擃嚗n\n?? 撘瑞?撱箄降嚗?Ｗ??冽????喲??脩戌(蝛拙?)憌?嚗?啣?澆??賬?)

            if not alert_up.empty:
                up_items = '??.join(alert_up['item_name'].head(5).tolist())
                st.warning(
                    f"?? **?剜?瞍脣?霅衣內嚗?{5}%嚗?*嚗up_items}\n\n?? 撱箄降嚗?隡唳隞???????Ⅱ隤梁??西蝮格???)

            if not alert_all_time_low.empty:
                low_items = '??.join(
                    alert_all_time_low['item_name'].head(5).tolist())
                st.success(
                    f"??**?風?脖?暺脣??*嚗low_items} ?桀?靘隞僑雿嚗n\n?? 撱箄降嚗?其?頞澈摮Ⅱ靽擙桃???銝??方疏嚗?雿?CPG??)
            elif not alert_down.empty:
                down_items = '??.join(alert_down['item_name'].head(5).tolist())
                st.success(
                    f"?? **?剜??璈?嚗?{5}%??**嚗down_items}\n\n?? 撱箄降嚗?詨?靘踹?嚗?拚?憭??)

            if alert_up.empty and alert_down.empty and alert_all_time_high.empty and alert_all_time_low.empty:
                st.info("???祆???湧?蝛拙?嚗?＊?啣虜瘜Ｗ?嚗??鞈潸??怠銵?胯?)

            # 敶??銵?
            with st.expander("?? 摰?啁??銵?):
                summary_rows = []
                for _, r in ranked_all.iterrows():
                    is_ath = (r['price'] >= r['ytd_max']) and (
                        r['ytd_max'] > r['ytd_min'])
                    is_atl = (r['price'] <= r['ytd_min']) and (
                        r['ytd_max'] > r['ytd_min'])

                    if is_ath:
                        strategy = "? 甇瑕擃?嚗撥?遣霅啣???
                    elif is_atl:
                        strategy = "??甇瑕雿?嚗遣霅啣鞎?
                    elif r['change_pct'] > 5:
                        strategy = "?? ?剜?瞍脣嚗?蹂誨"
                    elif r['change_pct'] < -5:
                        strategy = "?? ?剜??嚗憓?"
                    else:
                        strategy = "? 蝛拙?嚗撣詨鞎?

                    ytd_avg = r.get('ytd_avg', 0)
                    if pd.isna(ytd_avg) or ytd_avg == 0:
                        ytd_str = '??
                    else:
                        diff_pct = ((r['price'] - ytd_avg) / ytd_avg * 100)
                        ytd_str = f"??{ytd_avg:.1f} ({'+' if diff_pct>0 else ''}{diff_pct:.1f}%)"
                        if is_ath:
                            ytd_str += " [擃?]"
                        elif is_atl:
                            ytd_str += " [雿?]"

                    summary_rows.append({
                        '??': r['item_name'],
                        '?祆??桀': f"{r['price']:.0f} ??{r['unit']}",
                        '?剜?瞍脰?': f"{'+' if r['change_pct']>0 else ''}{r['change_pct']:.1f}%",
                        '蝝像?箸?撠': ytd_str,
                        '?啁撱箄降': strategy
                    })
                st.dataframe(pd.DataFrame(summary_rows),
                             use_container_width=True, hide_index=True)
        else:
            st.info("? ?怨疏?啁撱箄降?閬撠?????賣?撠?銝甈∟??格?啣?嚗票??`supplier_prices` ???喳?芸??Ｙ?撱箄降??)

        # ?? G. 憌?憸券?脩戌?? (?寞瘜Ｗ?摨血??? ??????????????????
        st.divider()
        st.markdown("#### ?儭?G. 憌?憸券?脩戌?? (?寞瘜Ｗ?摨血???")
        if 'ytd_stats' in locals() and not ytd_stats.empty:
            st.info(
                "? **瘜Ｗ???= (甇瑕?擃 - 甇瑕?雿) / 甇瑕?雿**?誨銵刻府憌??其?撟游?航?湔撞??憭批?摨艾n\n?? **?啁撱箄降**嚗蜓撱???交??⊿??輸??喳??憸券憌?嚗?雿輻撌血???脩戌憌?靘帘摰?CPG??)

            vol_df = ytd_stats[ytd_stats['ytd_min'] > 0].copy()
            vol_df['volatility'] = (
                vol_df['ytd_max'] - vol_df['ytd_min']) / vol_df['ytd_min'] * 100

            high_risk = vol_df[vol_df['volatility'] > 50].sort_values(
                'volatility', ascending=False)
            low_risk = vol_df[vol_df['volatility'] <= 20].sort_values(
                'volatility', ascending=True)

            vc1, vc2 = st.columns(2)
            with vc1:
                st.markdown("##### ?儭?擃蝳阡憸冽葛 (瘜Ｗ?璆萎? ??20%)")
                if not low_risk.empty:
                    for _, r in low_risk.head(10).iterrows():
                        st.markdown(
                            f"<div style='background:#f2fdf5; border-left:4px solid #2ecc71; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#888; font-size:12px;'>?憭扳撞撟? <span style='color:#2ecc71;'>{r['volatility']:.0f}%</span></span>"
                            f"<br><span style='font-size:12px; color:#666;'>??? {r['ytd_min']:.0f} ~ {r['ytd_max']:.0f} ??{r['unit']} (??{r['ytd_avg']:.0f})</span></div>",
                            unsafe_allow_html=True
                        )
                else:
                    st.write("?桀?撠雿郭????)

            with vc2:
                st.markdown("##### ? 擃◢?芸?瑕? (瘜Ｗ??? > 50%)")
                if not high_risk.empty:
                    for _, r in high_risk.head(10).iterrows():
                        st.markdown(
                            f"<div style='background:#fdf2f2; border-left:4px solid #e74c3c; padding:8px 12px; border-radius:6px; margin-bottom:6px;'>"
                            f"<strong>{r['item_name']}</strong> <span style='color:#888; font-size:12px;'>?憭扳撞撟? <span style='color:#e74c3c;'>{r['volatility']:.0f}%</span></span>"
                            f"<br><span style='font-size:12px; color:#666;'>??? {r['ytd_min']:.0f} ~ {r['ytd_max']:.0f} ??{r['unit']} (??{r['ytd_avg']:.0f})</span></div>",
                            unsafe_allow_html=True
                        )
                else:
                    st.write("?桀?撠擃郭????)
        else:
            st.info("? ?閬撠?????賢??????寞瘜Ｗ?摨艾?)

        # ?? H. ???銵???(Peak Demand Radar) ??????????????????
        st.divider()
        # ?望?閬???摨行???湔敺?m_curr (??tab_m 撌脣?頛??祆??豢?) 銝剝??啗?蝞?
        if 'm_curr' in locals() or 'm_curr' in globals():
            s_curr_metrics = calc_key_metrics(m_curr)
            if s_curr_metrics.get('dual_match_dates'):
                st.markdown("#### ? H. ???銵???(Peak Demand Radar)")
                st.info(
                    "? ?蝟餌絞?芸??芸?祆?蝚血??????????????乓?*隢?湔鞈潔犖?∠?交釣?嗾憭拍??????** ?予?拚?擃陸?犖?賊?隡啣?銝??航?∠?桀頛??迤?券??寧??蹂誨??嚗?蝬剜?擃?鞈芸?摰? CPG ?格???)

                # ?蔥?祆???????
                s_df_combined = pd.concat([m_curr['df'], m_next['df']], ignore_index=True) if 'm_next' in locals(
                ) and not m_next['df'].empty else m_curr['df'].copy()

                s_radar_cols = st.columns(
                    min(max(len(s_curr_metrics['dual_match_dates']), 1), 5))
                for i, d_date in enumerate(s_curr_metrics['dual_match_dates']):
                    # ?予?交?
                    next_day = (datetime.datetime.strptime(
                        d_date, '%Y-%m-%d') + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
                    next_day_row = s_df_combined[s_df_combined['date'] == next_day]

                    bf_count = 0
                    if not next_day_row.empty:
                        bf_col = 'bf_total_act' if 'bf_total_act' in next_day_row.columns and pd.to_numeric(
                            next_day_row['bf_total_act'].iloc[0], errors='coerce') > 0 else 'bf_total_est'
                        if bf_col in next_day_row.columns:
                            bf_count = pd.to_numeric(
                                next_day_row[bf_col], errors='coerce').fillna(0).iloc[0]

                    c = s_radar_cols[i % 5]
                    c.markdown(f"""
                    <div style="background: #fff; border: 2px solid #e74c3c; border-radius: 8px; padding: 15px; text-align: center; margin-bottom: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                        <p style="margin:0; font-size:10px; color:#aaa; letter-spacing:0.5px;">? ???乩???/p>
                        <h4 style="margin:4px 0; color:#e74c3c;">{d_date[5:]}</h4>
                        <hr style="border:0; border-top:1px dashed #eee; margin:8px 0;">
                        <p style="margin:0; font-size:10px; color:#aaa; letter-spacing:0.5px;">?? ???伐??拚?擃陸嚗?/p>
                        <p style="margin:4px 0; font-size:14px; color:#333; font-weight:bold;">{next_day[5:]}</p>
                        <p style="margin:4px 0 0 0; font-size:12px; color:#666;">?摯??: <strong style="color:#e74c3c;">{int(bf_count)}</strong> 鈭?/p>
                    </div>
                    """, unsafe_allow_html=True)

with tab7:
    st.header("? 鈭箔?璁?")

    # -- 鈭箔?蝞∠??賣 (Google Sheets ?? --
    def get_all_employees():
        try:
            df = conn.read(worksheet="employees", ttl="1m")
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

    # -- UI: ?啣??∪極? --
    with st.expander("???啣??圈脣撌亥?閮?, expanded=False):
        with st.form("add_employee_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            new_id = col1.text_input("?∪極蝺刻? (敹‵)")
            new_name = col2.text_input("憪? (敹‵)")

            new_dept = st.selectbox(
                "?撅祇?", ["頝臬?Plus銵?蝡?擗?, "瑹狗", "?踹?", "撌亙?", "The Peak"])
            new_pos = st.text_input("?瑚?")
            new_salary = st.number_input("?芾?", min_value=0, step=1000)

            submit_btn = st.form_submit_button("??蝣箄??啣?")
            if submit_btn:
                if not new_id or not new_name:
                    st.error("??隢‵撖怠撌亦楊??憪?嚗?)
                else:
                    res = add_employee(
                        new_id, new_name, new_dept, new_pos, new_salary)
                    if res == True:
                        st.success(f"?????啣??∪極嚗new_name}")
                        st.rerun()
                    elif res == "ID_EXISTS":
                        st.error("???∪極蝺刻?撌脣??剁?隢炎?交?阡?閬?)
                    else:
                        st.error(f"???啣?憭望?嚗res}")

    st.divider()

    # -- UI: ?∪極?”??摨?--
    df_emp = get_all_employees()

    # ?蕪?征?賜?鞈???(憒?Google Sheets 撣貉???撠曄征?質?)
    if not df_emp.empty and 'employee_id' in df_emp.columns:
        df_emp['employee_id'] = df_emp['employee_id'].astype(str).str.strip()
        # 蝘駁 pandas ?芸?撠摮???float ??Ｙ???.0 蝯偏
        df_emp['employee_id'] = df_emp['employee_id'].str.replace(
            r'\.0$', '', regex=True)
        df_emp = df_emp[df_emp['employee_id'] != '']
        df_emp = df_emp[df_emp['employee_id'].str.lower() != 'nan']

    if df_emp.empty:
        st.info("? ?桀?鞈?摨思葉撠?∪極鞈???)
    else:
        # 閮?蝮質鞈?(??瑚???PT ?犖)
        # 蝣箔? position 甈?摮銝??之撠神
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
            <p style="margin: 0; font-size: 14px; color: #666;">? 甇??∪極?芾?蝮質?</p>
            <h2 style="margin: 0; color: #2e437c;">NT$ {int(total_salary):,}</h2>
            <p style="margin: 5px 0 0 0; font-size: 12px; color: #999;">* 撌脰???方雿?蝔梁 "PT" ?犖?⊥??/p>
        </div>
        """, unsafe_allow_html=True)

        col_sort, col_search = st.columns([1, 1])
        sort_opt = col_sort.selectbox(
            "???孵?", ["?∪極蝺刻???", "?芾? (?梢??唬?)", "?芾? (?曹??圈?)", "????"])
        search_query = col_search.text_input("?? ??憪??楊??)

        # ???蕪
        if search_query:
            df_emp = df_emp[df_emp['name'].astype(str).str.contains(
                search_query, case=False) | df_emp['employee_id'].str.contains(search_query, case=False)]

        # 蝣箔? salary ?箸?潔誑靘踵?摨?
        if 'salary' in df_emp.columns:
            df_emp['salary'] = pd.to_numeric(
                df_emp['salary'], errors='coerce').fillna(0)

        # ???摩
        if sort_opt == "?∪極蝺刻???":
            df_emp = df_emp.sort_values("employee_id")
        elif sort_opt == "?芾? (?梢??唬?)":
            if 'salary' in df_emp.columns:
                df_emp = df_emp.sort_values("salary", ascending=False)
        elif sort_opt == "?芾? (?曹??圈?)":
            if 'salary' in df_emp.columns:
                df_emp = df_emp.sort_values("salary", ascending=True)
        elif sort_opt == "????":
            if 'dept' in df_emp.columns:
                df_emp = df_emp.sort_values(["dept", "employee_id"])

        # ?芸?蝢抵”?潮＊蝷?
        st.write(f"?? ?桀??望? {len(df_emp)} 雿撌?)

        # 璅???
        header_cols = st.columns([1.5, 1.5, 1.5, 1.5, 1.5, 1])
        header_cols[0].markdown("**?∪極蝺刻?**")
        header_cols[1].markdown("**憪?**")
        header_cols[2].markdown("**?券?**")
        header_cols[3].markdown("**?瑚?**")
        header_cols[4].markdown("**?芾?**")
        header_cols[5].markdown("**??**")

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

            # 雿輻 idx 靘?霅?key 蝯??臭?嚗??StreamlitDuplicateElementKey
            if row_cols[5].button("??儭?, key=f"del_{idx}_{row.get('employee_id', '')}", help="?芷甇文撌?):
                delete_employee(row.get('employee_id', ''))
                st.toast(f"撌脣?文撌? {row['name']}")
                time.sleep(0.5)
                st.rerun()

