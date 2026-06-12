import streamlit as st
import pandas as pd
import altair as alt
import io

def clean_nation_name(nation_str):
    if not isinstance(nation_str, str):
        return str(nation_str)
    # Remove English characters to leave only Chinese (e.g. KOR韓國 -> 韓國)
    import re
    cleaned = re.sub(r'[a-zA-Z\s]', '', nation_str).strip()
    return cleaned if cleaned else nation_str

def parse_tourism_bureau_excel(uploaded_file):
    try:
        # 讀取觀光署 Excel
        # 通常欄位在第 3 列 (index 2)，且有合併儲存格
        df = pd.read_excel(uploaded_file, header=2)
        
        # 找出國籍欄位 (可能叫 Nationality 或 居住地)
        nat_col = next((c for c in df.columns if 'Nationality' in str(c) or '國籍' in str(c) or '居住地' in str(c) or 'Unnamed: 0' in str(c)), None)
        
        # 找出當期與去年同期欄位 (通常包含年份或月份)
        # 假設結構：第1欄國籍，第2欄當期，第3欄同期，第4欄成長率
        if len(df.columns) >= 4:
            curr_col = df.columns[1]
            prev_col = df.columns[2]
            growth_col = df.columns[3]
            
            res_df = df[[nat_col, curr_col, prev_col, growth_col]].copy()
            res_df.columns = ['Nation_Raw', 'Curr_Arrivals', 'Prev_Arrivals', 'Growth_Rate_Pct']
            
            # 清理資料
            res_df = res_df.dropna(subset=['Nation_Raw'])
            res_df = res_df[~res_df['Nation_Raw'].astype(str).str.contains('Table|Total|Total|計|小計', na=False, case=False)]
            
            # 整理國籍名稱以便比對
            res_df['Nation_Clean'] = res_df['Nation_Raw'].apply(clean_nation_name)
            
            # 確保數值正確
            res_df['Curr_Arrivals'] = pd.to_numeric(res_df['Curr_Arrivals'], errors='coerce').fillna(0)
            res_df['Growth_Rate_Pct'] = pd.to_numeric(res_df['Growth_Rate_Pct'], errors='coerce').fillna(0)
            
            return res_df
        return None
    except Exception as e:
        st.error(f"解析觀光署報表失敗: {e}")
        return None

def render_nationality_tab():
    st.header("🌍 國籍客源分析專區")
    st.markdown("分析本飯店各國籍旅客分佈，並可結合交通部觀光署大盤數據進行交叉比對。")
    
    # 1. 讀取飯店資料
    df_hotel = None
    with st.spinner("載入飯店客源資料中..."):
        try:
            # 透過現有連線讀取 nationality_data
            conn = st.connection('gsheets', type='streamlit_gsheets.GSheetsConnection')
            df_hotel_raw = conn.read(worksheet="nationality_data")
            
            if not df_hotel_raw.empty:
                # 清理與過濾
                df_hotel = df_hotel_raw.copy()
                df_hotel.columns = [str(c).strip().lower() for c in df_hotel.columns]
                
                # 確保必要欄位存在
                if set(['nation', 'person', 'rate', 'nights']).issubset(set(df_hotel.columns)):
                    # 數值轉換
                    for c in ['person', 'rate', 'nights']:
                        # 處理可能的逗號千分位
                        df_hotel[c] = df_hotel[c].astype(str).str.replace(',', '', regex=False)
                        df_hotel[c] = pd.to_numeric(df_hotel[c], errors='coerce').fillna(0)
                    
                    # 過濾掉 0 房晚的國家
                    df_hotel = df_hotel[df_hotel['nights'] > 0].copy()
                    
                    if not df_hotel.empty:
                        # 擷取純中文國名
                        df_hotel['nation_clean'] = df_hotel['nation'].apply(clean_nation_name)
                        
                        # 計算 ADR (避免除以零)
                        df_hotel['adr'] = df_hotel.apply(lambda r: round(r['rate'] / r['nights']) if r['nights'] > 0 else 0, axis=1)
                        
                        # 若有多個月分，先以總和呈現
                        df_agg = df_hotel.groupby(['nation', 'nation_clean'], as_index=False).agg({
                            'nights': 'sum',
                            'person': 'sum',
                            'rate': 'sum'
                        })
                        df_agg['adr'] = df_agg.apply(lambda r: round(r['rate'] / r['nights']) if r['nights'] > 0 else 0, axis=1)
                        
                        # 計算佔比
                        total_nights = df_agg['nights'].sum()
                        df_agg['nights_pct'] = (df_agg['nights'] / total_nights * 100).round(1)
        except Exception as e:
            st.error(f"讀取 RTS_backup(nationality_data) 發生錯誤: {e}")
    
    if df_hotel is None or df_hotel.empty:
        st.warning("⚠️ 找不到客源資料，或資料庫中無大於 0 房晚的紀錄。請確認 Google Sheet 的 nationality_data 分頁。")
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
        st.subheader("🏆 主力客源分佈 (依房晚數)")
        
        # 處理 Top 10 與 Others
        df_pie = df_agg.sort_values('nights', ascending=False).copy()
        if len(df_pie) > 10:
            top10 = df_pie.head(10)
            others = pd.DataFrame([{
                'nation': '其他 (Others)', 'nation_clean': '其他',
                'nights': df_pie.iloc[10:]['nights'].sum(),
                'person': df_pie.iloc[10:]['person'].sum(),
                'rate': df_pie.iloc[10:]['rate'].sum(),
                'nights_pct': df_pie.iloc[10:]['nights_pct'].sum(),
                'adr': round(df_pie.iloc[10:]['rate'].sum() / df_pie.iloc[10:]['nights'].sum()) if df_pie.iloc[10:]['nights'].sum() > 0 else 0
            }])
            df_pie = pd.concat([top10, others], ignore_index=True)
        
        pie_chart = alt.Chart(df_pie).mark_arc(innerRadius=50).encode(
            theta=alt.Theta(field="nights", type="quantitative"),
            color=alt.Color(field="nation", type="nominal", sort='-q', legend=alt.Legend(title="國籍")),
            tooltip=['nation', 'nights', 'nights_pct', 'rate']
        ).properties(height=350)
        st.altair_chart(pie_chart, use_container_width=True)
        
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
        df_agg[['nation', 'nights', 'nights_pct', 'person', 'rate', 'adr']].sort_values('nights', ascending=False).rename(columns={
            'nation': '國籍 (代碼)',
            'nights': '總房晚數',
            'nights_pct': '佔比 (%)',
            'person': '總人數',
            'rate': '總營收',
            'adr': '平均房價 (ADR)'
        }),
        use_container_width=True,
        hide_index=True
    )
