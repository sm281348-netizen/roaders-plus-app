---
name: pms-ht2000-readonly
description: >-
  Use this skill whenever querying Roaders Plus Station Hotel's live JinXu PMS SQL Server database (HT2000)
  for room-type inventory, room-number-to-type mapping, future OTB bookings, or OOO room blocks,
  or when refreshing cloud snapshots under data/pms_backup/.
---

# 金旭 PMS (HT2000) 唯讀查詢與房型庫存快照技能

> **最高安全鐵則**：所有對 `HT2000` 之 SQL 查詢必須為 **100% 唯讀 `SELECT`**，絕對禁止執行任何寫入、修改或鎖表語句。

## 1. 核心資料表字典（實戰驗證版）
詳細欄位定義請見 [ht2000_schema.md](./references/ht2000_schema.md)。快速索引：
- **`HTRT`**：現行房型主檔（`RT01`=房型代碼、`RT03`=總房數；`RT03 > 0` 合計 128 間，已排除 2026-03 停用之 `DR`）。
- **`HTMSB`**：每日房態歷史（取 `MAX(MSB01)` 即可獲得全部 128 間房號 `MSB03` 對應現行房型 `MSB04`）。
- **`HTAC` JOIN `HTRV`**：未來每日分房型在手預訂（`AC02`=入住日 YYYYMMDD、`AC03`=房型、`SUM(AC06)`=已售間數，`WHERE v.RV06 = 'ACTV'`）。
- **`HTRB`**：現行與未來維修鎖房（`RB03`=日期、`RB04`=房號、`RB05='OOO'`、`RB06`=原因；常態布巾房為 `2509`、`2510`）。

## 2. 標準執行步驟
1. **確認連線模組**：複用專案根目錄 `pms_future_extractor.py` 內的 `get_sql_connection()`（自動從 `.streamlit/secrets.toml` 讀取設定，切勿硬寫帳密）。
2. **一鍵刷新全館房型剩餘庫存快照**（同時支援壁紙排程與未來 SiteMinder 定價比對）：
   ```powershell
   C:\Users\88691\AppData\Local\Python\pythoncore-3.14-64\python.exe ".agents\skills\pms-ht2000-readonly\scripts\refresh_roomtype_snapshot.py"
   ```
3. **驗證輸出**：確認 `data/pms_backup/wallpaper_roomtype_inventory.json` 已更新，且 12 個房型總房數加總等於 `128`。

## 3. 雲端環境降級原則
若在非飯店區網或 Streamlit Cloud 上無法連到 SQL Server，一律自動降級讀取 `data/pms_backup/station_occ_data.csv` 與 `data/pms_backup/wallpaper_roomtype_inventory.json`，並在 UI 標註資料快照時間。
