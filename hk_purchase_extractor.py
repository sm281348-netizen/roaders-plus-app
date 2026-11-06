# -*- coding: utf-8 -*-
r"""
房務請購單資料抽取模組 (HK Purchase Extractor - Production)
==========================================================
資料來源：Y:\房務\例行性表格\2026年\各月份\7@2026房務採購用4.0.xlsx
功能：
1. 自動定位 2026 最新月份累計活頁簿。
2. 遍歷 2026 年度所有請購單工作表（排除 2025 跨年單據、範本與非實體品項）。
3. 標準化欄位抽取：請購日期、品項名稱、當期單價、單位、請購數量、總價、建檔時間、類別(CAPEX/OPEX)、庫存量。
"""

import os
import re
import glob
import logging
import openpyxl
import pandas as pd
from datetime import datetime

logger = logging.getLogger("pms_sync")

DEFAULT_HK_DIR = r"Y:\房務\例行性表格\2026年"

EXCLUDED_SHEETS = [
    '20251223', '年前叫貨1223', '空白版-複製用', '(新增)站前房務常用品項', 
    '20260519-瀨口杯', '點貨清單3.0-空白板'
]

EXCLUDED_ITEMS = ['0', 'None', '品名', '小計', '合計', '總計', '零頭', '請購人', '規格說明']

def find_latest_hk_purchase_file(base_dir=DEFAULT_HK_DIR):
    """在 2026 年目錄中尋找最新月份的採購 Excel 檔案"""
    if not os.path.exists(base_dir):
        logger.warning(f"[HK_EXTRACT] 找不到房務目錄: {base_dir}")
        return None
        
    month_dirs = sorted(
        [d for d in glob.glob(os.path.join(base_dir, "2026-*月份")) if os.path.isdir(d)],
        reverse=True
    )
    
    for mdir in month_dirs:
        candidates = glob.glob(os.path.join(mdir, "*房務採購用*.xlsx"))
        valid_files = [f for f in candidates if not os.path.basename(f).startswith("~$")]
        if valid_files:
            latest_file = sorted(valid_files)[-1]
            logger.info(f"[HK_EXTRACT] 定位最新房務採購檔: {latest_file}")
            return latest_file
            
    return None

def extract_hk_purchases_2026(file_path=None):
    """
    從最新房務採購活頁簿中抽取全年度 2026 請購明細
    回傳 DataFrame: ['請購日期', '品項名稱', '當期單價', '單位', '請購數量', '總價', '建檔時間', '類別', '庫存量', '工作表', '代號']
    """
    if not file_path:
        file_path = find_latest_hk_purchase_file()
        
    if not file_path or not os.path.exists(file_path):
        logger.error(f"[HK_EXTRACT] 無法讀取房務採購檔: {file_path}")
        return pd.DataFrame()
        
    logger.info(f"[HK_EXTRACT] 開始讀取房務採購檔: {file_path}")
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    records = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for sname in wb.sheetnames:
        s_clean = sname.strip()
        if s_clean in EXCLUDED_SHEETS or s_clean.startswith('點貨清單'):
            continue
            
        ws = wb[sname]
        
        # 1. 抓取申請日期
        req_date_str = None
        for r in range(1, 5):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(r, c).value or '')
                if '申請日期' in val or '叫貨日' in val:
                    m = re.search(r'(\d{4})[/-]?(\d{1,2})[/-]?(\d{1,2})', val)
                    if m:
                        req_date_str = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                        break
            if req_date_str:
                break
                
        if not req_date_str:
            m = re.search(r'(\d{4})(\d{2})(\d{2})', s_clean)
            if m:
                req_date_str = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            elif '布巾' in s_clean:
                req_date_str = '2026-04-20'
            else:
                continue

        # 嚴格過濾 2026 年度資料
        if not req_date_str.startswith('2026'):
            continue

        # 2. 找表頭列
        header_row = None
        col_map = {}
        for r in range(1, 7):
            row_vals = [str(ws.cell(r, c).value or '').strip() for c in range(1, ws.max_column + 1)]
            if any('品名' in v for v in row_vals):
                header_row = r
                for c_idx, val in enumerate(row_vals, 1):
                    if '品名' in val: col_map['item'] = c_idx
                    elif '單價' in val: col_map['price'] = c_idx
                    elif '單位' in val: col_map['unit'] = c_idx
                    elif '數量' in val: col_map['qty'] = c_idx
                    elif '總金額' in val or '金額' in val: col_map['total'] = c_idx
                    elif '庫存' in val: col_map['inv'] = c_idx
                    elif '代號' in val: col_map['code'] = c_idx
                break
                
        if not header_row or 'item' not in col_map:
            continue
            
        # 3. 逐列提取品項
        for r in range(header_row + 1, ws.max_row + 1):
            item_val = ws.cell(r, col_map['item']).value
            if item_val is None:
                continue
            item_str = str(item_val).strip()
            if not item_str or item_str in EXCLUDED_ITEMS:
                continue
                
            qty_val = ws.cell(r, col_map['qty']).value if 'qty' in col_map else 0
            try:
                qty = float(qty_val or 0)
            except:
                qty = 0.0
                
            price_val = ws.cell(r, col_map['price']).value if 'price' in col_map else 0
            try:
                price = float(price_val or 0)
            except:
                price = 0.0
                
            # 若無數量且無單價，跳過空行/非實體列
            if qty == 0 and price == 0:
                continue
                
            total_val = ws.cell(r, col_map['total']).value if 'total' in col_map else 0
            try:
                total = float(total_val or (qty * price))
            except:
                total = qty * price
                
            unit = str(ws.cell(r, col_map['unit']).value or '').strip() if 'unit' in col_map else ''
            inv = str(ws.cell(r, col_map['inv']).value or '').strip() if 'inv' in col_map else ''
            code = str(ws.cell(r, col_map['code']).value or '').strip() if 'code' in col_map else ''
            
            is_capex = ('布巾' in s_clean or any(k in item_str for k in ['床單', '被套', '床包', '枕套', '羽絨被', '保潔墊']))
            category = '布巾資產(CAPEX)' if is_capex else '日常備品(OPEX)'
            
            records.append({
                '請購日期': req_date_str,
                '品項名稱': item_str,
                '當期單價': price,
                '單位': unit,
                '請購數量': qty,
                '總價': total,
                '建檔時間': now_str,
                '類別': category,
                '庫存量': inv,
                '工作表': s_clean,
                '代號': code
            })

    df = pd.DataFrame(records)
    logger.info(f"[HK_EXTRACT] 成功抽取 2026 房務請購明細共 {len(df)} 筆")
    return df
