# -*- coding: utf-8 -*-
r"""
菜商報價 (supplier_prices) 與請購日報表動態校準與全自動同步工具
=============================================================
功能特色：
1. 強制直連 Google Sheets supplier_prices，動態獲取最新生效期菜商核定報價。
2. 嚴格白名單隔離：只比對與更新來自 3@thepeak菜商叫貨單.xlsx 的生鮮蔬果，絕不誤傷常規採購品項。
3. 階層式品名比對：精準名稱 -> 蔬果別名字典 -> 長度安全模糊匹配，防誤配翻倍。
4. 批次原子寫入：透過 ws.batch_update 毫秒級更新已入帳品項，避免觸發 Google Sheets API 頻率限制。
5. 增量補單：自動追加當月未入帳之新叫貨單。
6. 本地備份與 Git 推播：自動同步本地 CSV 庫並推播 GitHub，觸發 Streamlit Cloud 即刻重載。
"""

import os
import sys
import argparse
import datetime
import logging
import subprocess
import tomllib
import pandas as pd
import gspread

# 確保輸出編碼正確
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("supplier_price_sync")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

SECRETS_PATH = os.path.join(APP_DIR, ".streamlit", "secrets.toml")
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")
os.makedirs(BACKUP_DIR, exist_ok=True)

from fb_purchase_extractor import (
    load_supplier_prices_df,
    extract_thepeak_purchases,
    lookup_veg_price
)

def safe_float(val, default=0.0):
    try:
        if val is None or str(val).strip() in ["", "-", "None"]:
            return default
        v_clean = str(val).replace(",", "").replace("$", "").replace("NT$", "").replace("NT", "").strip()
        return float(v_clean)
    except Exception:
        return default

def get_gsheet_client():
    if not os.path.exists(SECRETS_PATH):
        logger.error(f"找不到 secrets.toml: {SECRETS_PATH}")
        return None
    try:
        with open(SECRETS_PATH, "rb") as f:
            sec = tomllib.load(f)
        c = sec.get("connections", {}).get("gsheets_station", {})
        if c.get("type") != "service_account":
            logger.error("gsheets_station 未設定為 service_account")
            return None
        gc = gspread.service_account_from_dict(c)
        sh = gc.open_by_url(c["spreadsheet"])
        return sh
    except Exception as e:
        logger.error(f"連接 Google Sheets 失敗: {e}", exc_info=True)
        return None

def sync_veg_prices(target_year: int = None, target_month: int = None, dry_run: bool = False, auto_git_push: bool = True):
    """
    執行菜商報價與請購日報表同步核心流程
    """
    today = datetime.date.today()
    if target_year is None:
        target_year = today.year
    if target_month is None:
        target_month = today.month

    logger.info(f"========== 啟動菜商報價與採購品項動態校準 (目標: {target_year}年 {target_month}月) ==========")
    if dry_run:
        logger.info("🔍 [DRY RUN] 模擬測試模式：不會實際寫入 Google Sheets 或推播 Git")

    # 1. 強制重載 Google Sheets supplier_prices 最新報價
    logger.info("步驟 1/5: 正在從 Google Sheets 強制載入最新 supplier_prices 官方報價單...")
    df_sp = load_supplier_prices_df(force_reload=True)
    if df_sp.empty:
        logger.error("無法取得 supplier_prices 資料，終止同步。")
        return False
    latest_quote_date = df_sp['period_dt'].max()
    logger.info(f"✅ 報價單載入成功！共 {len(df_sp):,} 筆報價，最新核定生效期為: {latest_quote_date}")

    # 2. 解析目標月份的菜商叫貨單 (白名單隔離，僅解析 3@thepeak菜商叫貨單.xlsx)
    logger.info(f"步驟 2/5: 正在解析網路磁碟菜商叫貨單 (鎖定 {target_year} 年 {target_month} 月)...")
    df_peak_excel = extract_thepeak_purchases(target_year=target_year, target_month=target_month, force_reload_prices=True, veg_only=True)
    if df_peak_excel.empty:
        logger.warning(f"The Peak 叫貨單無 {target_year} 年 {target_month} 月資料。")
        return True

    # 嚴格過濾出來自 3@thepeak菜商叫貨單.xlsx 的菜商品項 (白名單機制)
    veg_excel = df_peak_excel[df_peak_excel["category"] == "菜商生鮮"].copy()
    logger.info(f"✅ 叫貨單解析成功: 總筆數 {len(df_peak_excel)} 筆，其中菜商生鮮品項共 {len(veg_excel)} 筆")

    # 3. 讀取 Google Sheets thepeak_daily_purchase_report 現有線上紀錄
    logger.info("步驟 3/5: 正在讀取 Google Sheets thepeak_daily_purchase_report...")
    sh = get_gsheet_client()
    if sh is None:
        return False
    ws = sh.worksheet("thepeak_daily_purchase_report")
    all_values = ws.get_all_values()
    if not all_values:
        logger.error("thepeak_daily_purchase_report 為空！")
        return False

    headers = all_values[0]
    # 建立線上現有資料索引: (日期, 去空格品名) -> 字典
    existing_map = {}
    for r_idx, row in enumerate(all_values[1:], start=2):
        d_val = row[0].strip() if len(row) > 0 else ""
        n_raw = row[1].strip() if len(row) > 1 else ""
        n_clean = n_raw.replace(" ", "").replace("　", "")
        if d_val and n_clean:
            p_val = safe_float(row[2]) if len(row) > 2 else 0.0
            u_val = row[3].strip() if len(row) > 3 else ""
            q_val = safe_float(row[4]) if len(row) > 4 else 0.0
            tot_val = safe_float(row[5]) if len(row) > 5 else 0.0
            existing_map[(d_val, n_clean)] = {
                "row_idx": r_idx,
                "price": p_val,
                "unit": u_val,
                "qty": q_val,
                "total": tot_val,
                "raw_name": n_raw,
                "raw_row": row
            }

    # 4. 嚴格比對：只針對白名單中的菜商品項進行單價與總額校對
    logger.info("步驟 4/5: 正在進行品項白名單比對與價格差異分析...")
    updates_batch = []
    missing_rows = []
    identical_count = 0
    total_diff_amt = 0.0

    for _, r in veg_excel.iterrows():
        d_val = str(r["date"]).strip()
        n_clean = str(r["item_name"]).strip().replace(" ", "").replace("　", "")
        key = (d_val, n_clean)

        if key in existing_map:
            ex = existing_map[key]
            # 優先使用 Google Sheets 線上實際記錄的請購數量 (尊重現場退換貨微調)
            actual_qty = ex["qty"] if ex["qty"] > 0 else r["quantity"]
            expected_price = r["price"]
            expected_total = round(expected_price * actual_qty, 2) if expected_price > 0 else 0.0

            # 檢查單價或總額是否需要更新
            price_mismatch = abs(ex["price"] - expected_price) > 1e-4
            tot_mismatch = abs(ex["total"] - expected_total) > 1e-4

            if price_mismatch or tot_mismatch:
                row_idx = ex["row_idx"]
                diff_amt = expected_total - ex["total"]
                total_diff_amt += diff_amt
                unit_to_use = ex["unit"] or r["unit"]
                now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                updates_batch.append({
                    "range": f"C{row_idx}:G{row_idx}",
                    "values": [[expected_price, unit_to_use, actual_qty, expected_total, now_str]],
                    "row_idx": row_idx,
                    "date": d_val,
                    "item_name": r["item_name"],
                    "unit": unit_to_use,
                    "qty": actual_qty,
                    "old_price": ex["price"],
                    "new_price": expected_price,
                    "old_tot": ex["total"],
                    "new_tot": expected_total,
                    "diff_amt": diff_amt
                })
            else:
                identical_count += 1
        else:
            # 叫貨單上有，但 Google Sheets 尚未入帳的新項目
            if r["quantity"] > 0:
                missing_rows.append(r)

    # 5. 輸出比對統計報表
    print("\n" + "=" * 95)
    print(f"📊 菜商報價比對結果摘要 ({target_year} 年 {target_month} 月)")
    print("=" * 95)
    print(f"• 叫貨單菜商品項總數: {len(veg_excel)} 筆")
    print(f"• 價格已一致 (無須更新): {identical_count} 筆")
    print(f"• 單價/總價有變動 (待校準): {len(updates_batch)} 筆 (合計影響金額: NT$ {total_diff_amt:+,.2f})")
    print(f"• 尚未入帳之新叫貨明細: {len(missing_rows)} 筆")
    print("-" * 95)

    if updates_batch:
        print(f"\n🚨 【待校準價格之明細清單 (共 {len(updates_batch)} 筆)】:")
        fmt = "{:<12} {:<16} {:>8} {:>10} {:>10} {:>10} {:>12} {:>10}"
        print(fmt.format("請購日期", "品項名稱", "數量/單位", "舊單價", "新單價", "原總額", "新總額", "差額"))
        print("-" * 95)
        for u in updates_batch:
            qty_unit = f"{u['qty']}{u['unit']}"
            diff_str = f"{u['diff_amt']:+,.1f}"
            print(fmt.format(
                u["date"],
                u["item_name"],
                qty_unit,
                f"${u['old_price']:.1f}",
                f"${u['new_price']:.1f}",
                f"${u['old_tot']:,.1f}",
                f"${u['new_tot']:,.1f}",
                diff_str
            ))
        print("-" * 95)

    if missing_rows:
        print(f"\n➕ 【待追加之新叫貨明細 (共 {len(missing_rows)} 筆)】:")
        for m in missing_rows:
            print(f"  • [{m['date']}] {m['item_name']} | 叫貨: {m['quantity']}{m['unit']} | 核定單價: ${m['price']:.1f} | 總額: ${m['total_amount']:,.1f}")

    # 6. 執行寫入
    if not dry_run:
        logger.info("步驟 5/5: 正在原子性更新 Google Sheets...")
        
        # 6.1 批次更新價格
        if updates_batch:
            batch_payload = [{"range": u["range"], "values": u["values"]} for u in updates_batch]
            ws.batch_update(batch_payload)
            logger.info(f"🎉 成功批次更新 {len(updates_batch)} 筆品項之單價與總額！")

        # 6.2 追加新單
        if missing_rows:
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            append_rows_data = []
            for m in missing_rows:
                append_rows_data.append([
                    str(m["date"]),
                    str(m["item_name"]),
                    str(m["price"]),
                    str(m["unit"]),
                    str(m["quantity"]),
                    str(m["total_amount"]),
                    now_str
                ])
            ws.append_rows(append_rows_data, value_input_option='USER_ENTERED')
            logger.info(f"🎉 成功追加 {len(append_rows_data)} 列新叫貨單明細！")

        if not updates_batch and not missing_rows:
            logger.info("✅ 所有菜商品項價格已是最新核定價，無須異動。")

        # 6.3 更新本地備份 CSV
        logger.info("正在刷新本地 CSV 備份庫...")
        all_updated_values = ws.get_all_values()
        if all_updated_values:
            df_updated = pd.DataFrame(all_updated_values[1:], columns=all_updated_values[0])
            backup_csv = os.path.join(BACKUP_DIR, "thepeak_daily_purchase_report_latest.csv")
            df_updated.to_csv(backup_csv, index=False, encoding="utf-8-sig")
            logger.info(f"✅ 本地備份已更新: {backup_csv}")

        # 6.4 自動 Git Push
        if auto_git_push and (updates_batch or missing_rows):
            logger.info("正在自動推播變更至 GitHub 倉庫 (main 分支)...")
            try:
                subprocess.run(["git", "add", "data/pms_backup/"], cwd=APP_DIR, check=True)
                diff_res = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=APP_DIR)
                if diff_res.returncode != 0:
                    commit_msg = f"chore(sync): update veg prices from supplier_prices ({target_year}-{target_month:02d}) [skip ci]"
                    subprocess.run(["git", "commit", "-m", commit_msg], cwd=APP_DIR, check=True)
                    push_res = subprocess.run(["git", "push", "origin", "main"], cwd=APP_DIR, capture_output=True, text=True, encoding="utf-8")
                    if push_res.returncode == 0:
                        logger.info("🚀 [SUCCESS] 數據已成功推播至 GitHub！Streamlit Cloud 將自動同步呈現最新成本。")
                    else:
                        logger.warning(f"Git 推播警告: {push_res.stderr}")
                else:
                    logger.info("本地庫無檔案異動，略過 Git commit。")
            except Exception as e_git:
                logger.error(f"Git 推播發生例外: {e_git}")

    logger.info("========== 菜商報價同步作業圓滿完成 ==========\n")
    return True

def main():
    parser = argparse.ArgumentParser(description="菜商報價 (supplier_prices) 與請購單同步更新工具")
    parser.add_argument("--year", type=int, default=None, help="目標年份 (預設今年)")
    parser.add_argument("--month", type=int, default=None, help="目標月份 (預設當月)")
    parser.add_argument("--dry-run", action="store_true", help="模擬測試模式，不寫入 Google Sheets")
    parser.add_argument("--no-git", action="store_true", help="跳過 Git 推播")

    args = parser.parse_args()
    sync_veg_prices(
        target_year=args.year,
        target_month=args.month,
        dry_run=args.dry_run,
        auto_git_push=not args.no_git
    )

if __name__ == "__main__":
    main()
