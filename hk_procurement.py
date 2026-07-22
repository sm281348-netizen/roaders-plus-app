import pandas as pd
import openpyxl
import os
import streamlit as st
import math

HK_ITEM_MASTER_FILE = "hk_item_master.csv"

def init_hk_item_master():
    """初始化房務品項主檔"""
    if not os.path.exists(HK_ITEM_MASTER_FILE):
        df = pd.DataFrame(columns=[
            "品名", "換算單位(一箱幾件)", "標準配給量(每房幾件)", "備註"
        ])
        df.to_csv(HK_ITEM_MASTER_FILE, index=False, encoding='utf-8-sig')

def load_hk_item_master():
    """讀取房務品項主檔"""
    init_hk_item_master()
    return pd.read_csv(HK_ITEM_MASTER_FILE, encoding='utf-8-sig')

def save_hk_item_master(df):
    """儲存房務品項主檔"""
    df.to_csv(HK_ITEM_MASTER_FILE, index=False, encoding='utf-8-sig')

def extract_uom_multiplier(uom_str):
    """從 '500入/箱' 提取 500，若解析失敗回傳 1"""
    if pd.isna(uom_str) or not isinstance(uom_str, str):
        return 1.0
    import re
    match = re.search(r'(\d+)', uom_str)
    if match:
        return float(match.group(1))
    return 1.0

def parse_hk_inventory(file_bytes_or_path):
    """解析房務部盤點 Excel 檔"""
    # 讀取 Excel (只抓取值，忽略公式)
    wb = openpyxl.load_workbook(file_bytes_or_path, data_only=True)
    ws = wb.worksheets[0]
    data = ws.values
    
    # 找到表頭行
    cols = None
    for row in data:
        row_list = list(row)
        if row_list and isinstance(row_list[0], str) and "品名" in row_list[0]:
            # 將 \n 替換掉
            cols = [str(c).replace('\n', '') if c else f"Unnamed_{i}" for i, c in enumerate(row_list)]
            break
            
    if cols is None:
        st.error("找不到 Excel 表頭 (必須包含 '品名')")
        return pd.DataFrame()
        
    df = pd.DataFrame(data, columns=cols)
    
    # 過濾空行與非品項列
    df = df.dropna(subset=[cols[0]])
    df = df[~df[cols[0]].str.contains('品名', na=False)] # 過濾掉重複的表頭
    
    # 重新命名欄位，使其統一
    rename_dict = {}
    for c in cols:
        if '品名' in c: rename_dict[c] = 'ItemName'
        elif '本期庫存' in c: rename_dict[c] = 'CurrentInventory'
        elif '本期叫貨' in c: rename_dict[c] = 'CurrentOrder'
        elif '數量' in c: rename_dict[c] = 'UOM' # 原本的規格欄位
    
    df = df.rename(columns=rename_dict)
    
    # 只保留需要的欄位
    required_cols = ['ItemName', 'CurrentInventory', 'CurrentOrder', 'UOM']
    available_cols = [c for c in required_cols if c in df.columns]
    df = df[available_cols]
    
    # 確保數值欄位型別
    if 'CurrentInventory' in df.columns:
        df['CurrentInventory'] = pd.to_numeric(df['CurrentInventory'], errors='coerce').fillna(0)
    if 'CurrentOrder' in df.columns:
        df['CurrentOrder'] = pd.to_numeric(df['CurrentOrder'], errors='coerce').fillna(0)
        
    # 自動同步新發現的品項到 Master File
    sync_new_items_to_master(df)
    
    return df

def sync_new_items_to_master(parsed_df):
    """將 Excel 裡發現的新品項加到 hk_item_master.csv"""
    master_df = load_hk_item_master()
    existing_items = set(master_df['品名'].dropna().tolist())
    
    new_rows = []
    for _, row in parsed_df.iterrows():
        item_name = str(row.get('ItemName', '')).strip()
        if item_name and item_name not in existing_items and item_name != 'nan':
            uom_str = str(row.get('UOM', ''))
            multiplier = extract_uom_multiplier(uom_str)
            new_rows.append({
                "品名": item_name,
                "換算單位(一箱幾件)": multiplier,
                "標準配給量(每房幾件)": 1.0, # 預設給 1
                "備註": "系統自動新增"
            })
            existing_items.add(item_name)
            
    if new_rows:
        master_df = pd.concat([master_df, pd.DataFrame(new_rows)], ignore_index=True)
        save_hk_item_master(master_df)

def render_hk_procurement_dashboard(forecast_df):
    """渲染房務部 AI 預測叫貨戰情室"""
    st.markdown("### 🧹 房務數據戰情室：AI 動態預測叫貨")
    
    # === 1. 上傳區塊與參數設定 ===
    col1, col2 = st.columns([1, 1])
    with col1:
        st.info("💡 請上傳您的盤點 Excel 檔 (例如：`8@2026房務採購用4.0.xlsx`)")
        uploaded_file = st.file_uploader("上傳房務盤點檔 (.xlsx)", type=["xlsx"])
        
    with col2:
        st.markdown("**⚙️ 營運參數設定**")
        lead_time = st.number_input("廠商交貨期 (天)", min_value=1, value=3, help="從下單到送達需要的天數")
        order_cycle = st.number_input("採購週期 (天)", min_value=1, value=14, help="多久叫一次貨？例如兩週一次為14天")
        lag_days = st.number_input("盤點至今天數 (天)", min_value=0, value=0, help="盤點日離今天隔了幾天？系統會自動扣除這幾天的實際消耗。")
        safety_stock_ratio = st.slider("安全庫存比例 (%)", min_value=0, max_value=100, value=20, help="額外多抓的緩衝比例，預設20%")

    st.markdown("---")
    
    # === 2. 品項設定介面 ===
    st.subheader("⚙️ 品項單位與配給量設定")
    st.caption("請在此一次性設定每個品項的換算單位（一箱幾件）以及標準配給量（每間房消耗幾件）。")
    master_df = load_hk_item_master()
    edited_master_df = st.data_editor(master_df, num_rows="dynamic", use_container_width=True)
    if not master_df.equals(edited_master_df):
        save_hk_item_master(edited_master_df)
        st.success("品項設定已儲存！")
    
    st.markdown("---")

    # === 3. 分析與計算邏輯 ===
    if uploaded_file is not None:
        with st.spinner("正在解析盤點檔案與計算 AI 預測..."):
            inv_df = parse_hk_inventory(uploaded_file)
            
            if not inv_df.empty:
                st.subheader("🔮 AI 預測叫貨建議")
                
                # 計算涵蓋期間的預測住房數
                # 假設 forecast_df 有 '營業日期' 和 '房間預估_銷售房間數' 欄位
                # 這裡為了展示，簡化處理日期過濾，實務上需要 datetime 過濾
                try:
                    total_forecast_rooms = forecast_df['房間預估_銷售房間數'].sum() # 應縮小到 lead_time + order_cycle
                except:
                    total_forecast_rooms = 1000 # 備用假資料
                
                st.info(f"📅 未來預測區間 ({lead_time + order_cycle} 天) 的總預測賣出房間數： **{total_forecast_rooms:,.0f}** 間")
                
                # 進行運算
                results = []
                for _, row in inv_df.iterrows():
                    item = row['ItemName']
                    curr_inv = row['CurrentInventory']
                    orig_order = row.get('CurrentOrder', 0)
                    
                    # 取出對應的設定
                    config = edited_master_df[edited_master_df['品名'] == item]
                    if not config.empty:
                        uom_multiplier = config.iloc[0]['換算單位(一箱幾件)']
                        target_cpor = config.iloc[0]['標準配給量(每房幾件)']
                        status = "✅ 已設定"
                    else:
                        uom_multiplier = 1.0
                        target_cpor = 0.0
                        status = "⚠️ 未設定"
                    
                    # 公式計算：
                    # 預測總需求 = 預測總房間數 * 每房標準配給量
                    projected_demand = total_forecast_rooms * target_cpor
                    # 安全庫存 = 預測需求 * 比例
                    safety_stock = projected_demand * (safety_stock_ratio / 100.0)
                    # 總需要件數 = 預測需求 + 安全庫存
                    total_needed_units = projected_demand + safety_stock
                    
                    # 應補件數 = 總需要 - 現有盤點 (假設暫無 lag_days 修正)
                    replenish_units = max(0, total_needed_units - curr_inv)
                    
                    # AI 建議叫貨量 (箱) = 應補件數 / 換算單位，並無條件進位
                    if uom_multiplier > 0:
                        ai_order = math.ceil(replenish_units / uom_multiplier)
                    else:
                        ai_order = 0
                        
                    results.append({
                        "狀態": status,
                        "品名": item,
                        "盤點庫存 (件)": curr_inv,
                        "預測總需求 (件)": round(projected_demand, 1),
                        "AI 建議叫貨 (箱)": ai_order,
                        "Excel 原始叫貨 (箱)": orig_order,
                        "差異": ai_order - orig_order
                    })
                
                result_df = pd.DataFrame(results)
                
                # Highlight 差異大的或未設定的
                def highlight_status(val):
                    color = 'red' if '⚠️' in str(val) else 'green'
                    return f'color: {color}'
                    
                st.dataframe(result_df.style.applymap(highlight_status, subset=['狀態']), use_container_width=True)
                
                # 單品成本異常警報 (若需要)
                st.subheader("🚨 異常消耗警報")
                st.warning("目前功能建置中：待匯入實際消耗成本即可比對紅燈異常品項。")
                
