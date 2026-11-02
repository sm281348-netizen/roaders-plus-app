# -*- coding: utf-8 -*-
r"""
餐飲部 (The Peak 與 4F 餐車) 採購單全自動解析引擎 (F&B Purchase Extractor)
========================================================================
支援全量歷史回補與即時監控：
1. 4F 餐車 (Happy Hour): Y:\餐務\{YYYY}(4F)\{M}月\ 採購單*.xlsx、菜商採購單*.xlsx
2. The Peak (37F): Y:\餐務\The peak\{YYYY}(37F)\ 1@THEPEAK採購用4.0.xlsx、3@thepeak菜商叫貨單.xlsx
3. 自動解析各工作表名稱（如 0261001, 10.5, 10 26, 11.2）對齊西元標準日期 (YYYY-MM-DD)。
4. 精準比對品名、單價、數量、小計金額，支援跨月與跨週自動分流。
"""

import os
import re
import datetime
import logging
import openpyxl
import pandas as pd

logger = logging.getLogger("pms_sync")

BASE_4F_DIR = r"Y:\餐務\2026(4F)"
BASE_PEAK_DIR = r"Y:\餐務\The peak\2026(37F)"

def parse_sheet_date(sheet_name: str, fallback_year: int = 2026) -> datetime.date:
    """
    從工作表名稱解析西元日期 (YYYY-MM-DD)
    支援格式:
    - '0261102', '261102', '20261102' -> 2026-11-02
    - '11.2', '10.5', '10.12', '10.21', '9.28' -> 2026-11-02, 2026-10-05...
    - '10 26', '8 24' -> 2026-10-26, 2026-08-24
    - '1010', '1015' -> 2026-10-10
    """
    s = str(sheet_name).strip()

    # 1. 026MMDD 或 26MMDD
    m_tw = re.match(r"^0?26(\d{2})(\d{2})", s)
    if m_tw:
        m, d = int(m_tw.group(1)), int(m_tw.group(2))
        return datetime.date(fallback_year, m, d)

    # 2. 2026MMDD
    m_full = re.match(r"^2026(\d{2})(\d{2})", s)
    if m_full:
        m, d = int(m_full.group(1)), int(m_full.group(2))
        return datetime.date(2026, m, d)

    # 3. M.D 或 M.DD 或 MM.DD (例如 11.2, 10.5, 10.19)
    m_dot = re.match(r"^(\d{1,2})\.(\d{1,2})", s)
    if m_dot:
        m, d = int(m_dot.group(1)), int(m_dot.group(2))
        return datetime.date(fallback_year, m, d)

    # 4. M D 或 MM DD (例如 '10 26', '8 24')
    m_space = re.match(r"^(\d{1,2})\s+(\d{1,2})", s)
    if m_space:
        m, d = int(m_space.group(1)), int(m_space.group(2))
        return datetime.date(fallback_year, m, d)

    # 5. MMDD 4位數字 (例如 '1008', '1012')
    m_4d = re.match(r"^(\d{2})(\d{2})$", s)
    if m_4d:
        m, d = int(m_4d.group(1)), int(m_4d.group(2))
        if 1 <= m <= 12 and 1 <= d <= 31:
            return datetime.date(fallback_year, m, d)

    return None

def extract_4f_purchases(target_year: int = 2026, target_month: int = None) -> pd.DataFrame:
    """
    抽取 4F 餐車 (Happy Hour) 叫貨單數據
    """
    records = []
    base_dir = rf"Y:\餐務\{target_year}(4F)"
    if not os.path.exists(base_dir):
        base_dir = BASE_4F_DIR

    if not os.path.exists(base_dir):
        logger.warning(f"4F 目錄不存在: {base_dir}")
        return pd.DataFrame()

    # 決定要掃描的月份子目錄
    month_folders = []
    if target_month:
        month_folders.append(f"{target_month}月")
    else:
        for item in os.listdir(base_dir):
            if os.path.isdir(os.path.join(base_dir, item)) and "月" in item:
                month_folders.append(item)

    for m_folder in month_folders:
        folder_path = os.path.join(base_dir, m_folder)
        if not os.path.isdir(folder_path):
            continue

        for fname in os.listdir(folder_path):
            if fname.startswith("~$") or not fname.endswith((".xlsx", ".xls")):
                continue
            if "採購" not in fname and "叫貨" not in fname:
                continue

            file_path = os.path.join(folder_path, fname)
            is_veg = "菜商" in fname
            category = "菜商生鮮" if is_veg else "常規採購"

            try:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                for sname in wb.sheetnames:
                    # 排除非訂單範本表
                    if any(k in sname for k in ["複製", "常用", "品項", "空白", "價格"]):
                        continue

                    d_obj = parse_sheet_date(sname, target_year)
                    if not d_obj:
                        continue
                    if target_month and d_obj.month != target_month:
                        continue

                    ws = wb[sname]
                    date_str = d_obj.strftime("%Y-%m-%d")

                    if not is_veg:
                        # 4F 常規請購單 (標準表頭在 Row 4)
                        for r in range(5, ws.max_row + 1):
                            item_name = ws.cell(r, 5).value
                            price = ws.cell(r, 6).value
                            unit = ws.cell(r, 7).value
                            qty = ws.cell(r, 8).value
                            amount = ws.cell(r, 9).value

                            if item_name and qty and qty not in [0, "0", None]:
                                try:
                                    q = float(qty)
                                    p = float(price) if price else 0.0
                                    amt = float(amount) if amount else (q * p)
                                    records.append({
                                        "date": date_str,
                                        "dept": "Happy Hour",
                                        "item_name": str(item_name).strip(),
                                        "quantity": q,
                                        "unit": str(unit).strip() if unit else "",
                                        "price": p,
                                        "total_amount": round(amt, 2),
                                        "category": category,
                                        "sheet_name": sname,
                                        "source_file": fname
                                    })
                                except Exception:
                                    pass
                    else:
                        # 4F 菜商叫貨單 (雙欄排版)
                        for r in range(5, ws.max_row + 1):
                            # 左半欄
                            n1, q1, u1 = ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value
                            if n1 and q1 and q1 not in [0, "0", None]:
                                try:
                                    records.append({
                                        "date": date_str,
                                        "dept": "Happy Hour",
                                        "item_name": str(n1).strip(),
                                        "quantity": float(q1),
                                        "unit": str(u1).strip() if u1 else "",
                                        "price": 0.0,
                                        "total_amount": 0.0,
                                        "category": category,
                                        "sheet_name": sname,
                                        "source_file": fname
                                    })
                                except Exception:
                                    pass
                            # 右半欄
                            n2, q2, u2 = ws.cell(r, 7).value, ws.cell(r, 8).value, ws.cell(r, 9).value
                            if n2 and q2 and q2 not in [0, "0", None]:
                                try:
                                    records.append({
                                        "date": date_str,
                                        "dept": "Happy Hour",
                                        "item_name": str(n2).strip(),
                                        "quantity": float(q2),
                                        "unit": str(u2).strip() if u2 else "",
                                        "price": 0.0,
                                        "total_amount": 0.0,
                                        "category": category,
                                        "sheet_name": sname,
                                        "source_file": fname
                                    })
                                except Exception:
                                    pass
                wb.close()
            except Exception as e:
                logger.error(f"解析 4F 檔案失敗 {file_path}: {e}")

    df = pd.DataFrame(records)
    return df

def extract_thepeak_purchases(target_year: int = 2026, target_month: int = None) -> pd.DataFrame:
    """
    抽取 The Peak (37F) 叫貨單數據
    """
    records = []
    base_dir = rf"Y:\餐務\The peak\{target_year}(37F)"
    if not os.path.exists(base_dir):
        base_dir = BASE_PEAK_DIR

    if not os.path.exists(base_dir):
        logger.warning(f"The Peak 目錄不存在: {base_dir}")
        return pd.DataFrame()

    # 1. 菜商價格對照表解析 (預先載入單價)
    veg_prices = {}
    veg_file = os.path.join(base_dir, "3@thepeak菜商叫貨單.xlsx")
    if os.path.exists(veg_file):
        try:
            wb_v = openpyxl.load_workbook(veg_file, data_only=True)
            for sname in wb_v.sheetnames:
                if "價格" in sname:
                    ws_p = wb_v[sname]
                    for r in range(6, ws_p.max_row + 1):
                        n1, p1 = ws_p.cell(r, 2).value, ws_p.cell(r, 5).value
                        if n1 and p1:
                            try:
                                veg_prices[str(n1).strip().replace(" ", "")] = float(p1)
                            except:
                                pass
                        n2, p2 = ws_p.cell(r, 7).value, ws_p.cell(r, 10).value
                        if n2 and p2:
                            try:
                                veg_prices[str(n2).strip().replace(" ", "")] = float(p2)
                            except:
                                pass
                    break
            wb_v.close()
        except Exception:
            pass

    # 2. 遍歷 The Peak 檔案
    target_files = [
        ("1@THEPEAK採購用4.0.xlsx", False, "常規採購"),
        ("3@thepeak菜商叫貨單.xlsx", True, "菜商生鮮")
    ]

    for fname, is_veg, category in target_files:
        fp = os.path.join(base_dir, fname)
        if not os.path.exists(fp):
            continue

        try:
            wb = openpyxl.load_workbook(fp, data_only=True)
            for sname in wb.sheetnames:
                if any(k in sname for k in ["複製", "常用", "品項", "空白", "價格", "工作表", "班表"]):
                    continue

                d_obj = parse_sheet_date(sname, target_year)
                if not d_obj:
                    continue
                if target_month and d_obj.month != target_month:
                    continue

                ws = wb[sname]
                date_str = d_obj.strftime("%Y-%m-%d")

                if not is_veg:
                    # 常規請購單
                    for r in range(5, ws.max_row + 1):
                        item_name = ws.cell(r, 5).value
                        price = ws.cell(r, 6).value
                        unit = ws.cell(r, 7).value
                        qty = ws.cell(r, 8).value
                        amount = ws.cell(r, 9).value

                        if item_name and qty and qty not in [0, "0", None]:
                            try:
                                q = float(qty)
                                p = float(price) if price else 0.0
                                amt = float(amount) if amount else (q * p)
                                records.append({
                                    "date": date_str,
                                    "dept": "The Peak",
                                    "item_name": str(item_name).strip(),
                                    "quantity": q,
                                    "unit": str(unit).strip() if unit else "",
                                    "price": p,
                                    "total_amount": round(amt, 2),
                                    "category": category,
                                    "sheet_name": sname,
                                    "source_file": fname
                                })
                            except Exception:
                                pass
                else:
                    # 菜商叫貨單
                    for r in range(5, ws.max_row + 1):
                        # 左半欄
                        n1, q1, u1 = ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value
                        if n1 and q1 and q1 not in [0, "0", None]:
                            try:
                                q = float(q1)
                                clean_n = str(n1).strip().replace(" ", "")
                                p = veg_prices.get(clean_n, 0.0)
                                amt = round(q * p, 2)
                                records.append({
                                    "date": date_str,
                                    "dept": "The Peak",
                                    "item_name": str(n1).strip(),
                                    "quantity": q,
                                    "unit": str(u1).strip() if u1 else "",
                                    "price": p,
                                    "total_amount": amt,
                                    "category": category,
                                    "sheet_name": sname,
                                    "source_file": fname
                                })
                            except Exception:
                                pass
                        # 右半欄
                        n2, q2, u2 = ws.cell(r, 7).value, ws.cell(r, 8).value, ws.cell(r, 9).value
                        if n2 and q2 and q2 not in [0, "0", None]:
                            try:
                                q = float(q2)
                                clean_n = str(n2).strip().replace(" ", "")
                                p = veg_prices.get(clean_n, 0.0)
                                amt = round(q * p, 2)
                                records.append({
                                    "date": date_str,
                                    "dept": "The Peak",
                                    "item_name": str(n2).strip(),
                                    "quantity": q,
                                    "unit": str(u2).strip() if u2 else "",
                                    "price": p,
                                    "total_amount": amt,
                                    "category": category,
                                    "sheet_name": sname,
                                    "source_file": fname
                                })
                            except Exception:
                                pass
            wb.close()
        except Exception as e:
            logger.error(f"解析 The Peak 檔案失敗 {fp}: {e}")

    df = pd.DataFrame(records)
    return df

def scan_october_audit():
    """專門掃描 10 月份漏更請購單診斷報告"""
    df_4f = extract_4f_purchases(2026, 10)
    df_peak = extract_thepeak_purchases(2026, 10)
    df_all = pd.concat([df_4f, df_peak], ignore_index=True)
    return df_all

if __name__ == "__main__":
    df_oct = scan_october_audit()
    print(f"Total October orders extracted: {len(df_oct)} rows")
    if not df_oct.empty:
        summary = df_oct.groupby(["dept", "date"]).agg({
            "total_amount": "sum",
            "item_name": "count"
        }).reset_index()
        print(summary.to_string(index=False))
