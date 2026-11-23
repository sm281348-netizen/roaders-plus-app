# Hotel Master (roaders-plus-app) 專案憲法與 AI 營運參謀準則

> 本專案為「路徒 PLUS 行旅（站前館／主題館）」經營決策與營運稽核中樞（Streamlit + 金旭 PMS + Google Sheets）。任何新對話與子代理在處理本工作區時，必須嚴格遵守以下鐵則。

## 1. 雙館主權與事實基準（Single Source of Truth）
- **站前館（主理館）**：
  - 總房數 `128 間`；常態布巾房 OOO 為 `2509 (SD)`、`2510 (SD)` → 常態可售 `126 間`。
  - 現行 12 種房型以金旭 `HTRT` 為主：`DF(8), F(8), DTP(4), LDT(4), DT(19), E(8), ER(8), PD(22), PT(12), SP(20), S(10), SD(5)`（註：`DR` 已於 2026-03 併入 `DT`）。
  - 未來在手 OTB 以每日排程更新之 `data/pms_backup/station_occ_data.csv`（或即時唯讀金旭 `HT2000`）為準；兩年歷史實績與月度目標存於 Google Sheets `RTS_backup`（`occ_data`、`targets`、`taipei_events`）。
- **主題館（無完整底層數據館）**：
  - 主題館缺乏各部門採購、人資與部分專案細部數據。**嚴禁留白報錯、嚴禁塞假數據**，無數據頁面一律呼叫 `app.py` 的 `_render_theme_unavailable_notice()`，顯示「暫無數據可提供此服務」並簡述該頁面能帶來的戰略價值。

## 2. 數據真實性與零假資料紅線（Zero-Mock Policy）
- 任何看板、圖表、PPTX（如店務會議簡報）或估算報告**絕對禁止寫死假資料**；若底層來源缺漏或無法計算，必須明確告知阻斷原因，絕不捏造。
- 凡涉及 Google Sheets、金旭 PMS、採購金額或房數邏輯修改，交付前**必須與原始表單對帳驗證**（如日期排序、當月合計、房數加總）。

## 3. 生產環境唯讀安全鐵則（Production Safety）
- **金旭 PMS (`HT2000`)**：100% 僅限 `SELECT` 唯讀查詢，**嚴禁任何 `INSERT` / `UPDATE` / `DELETE` / `DROP` / `ALTER`**。
- **SiteMinder 與外部 OTA 後台**：100% 僅限唯讀盤點與報表匯出分析，**嚴禁修改房價、庫存、Stop Sell、MinLOS 或點擊 Save / Publish**。
- **機密保護**：嚴禁讀取輸出或提交 `.streamlit/secrets.toml`、`rodersplusapp-*.json` 等憑證檔。

## 4. 程式碼防膨脹與專案紀律
- **`app.py` 禁止直接膨脹**：`app.py` 已逾 13,800 行（~800 KB）。所有新專案功能或新分頁**一律抽離為獨立 Python 模組**（比照 `mobile_ui.py`、`hk_dashboard.py`、`wallpaper_planner.py`），`app.py` 僅負責側邊欄 `menu_options` 註冊與 5～15 行路由掛載。
- **雲端與區網雙軌容錯**：凡依賴飯店區網金旭 PMS 或 `Y:` 槽的功能，必須同時支援讀取 `data/pms_backup/` 本地快照，確保 Streamlit Cloud 雲端部署零報錯。
- **禁止污染根目錄**：一次性除錯、爬蟲或資料驗證腳本一律放在對話 `scratch/` 目錄，禁止在專案根目錄新增 `fix_*.py`、`test_*.py`、`check_*.txt`。
- **禁止 `git add .`**：提交前必須以**白名單精準加檔**並通過 `py_compile` 語法檢核（詳見 `.agents/skills/safe-deploy-streamlit/SKILL.md`）。

## 5. 經理參謀溝通與紅隊自檢標準
- **語言與格式**：一律使用**繁體中文**，商業結論先行，金額與數據須附來源與計算公式。
- **主動紅隊思維（Red-Team）**：提出重大方案、採購評估、排程或架構改版時，主動檢視並揭露「邏輯漏洞、隱性成本（稅金/雜費/工時打折）與邊際風險」。
- **壓房與營運排程鐵則**：凡涉及關房／壓房施工規劃，必須同時驗證四個維度：**① 全館剩餘空房、② 單一房型（12 房型）零超賣緩衝、③ 去年同期實績與常態 OOO、④ 台北大型活動與節慶（`taipei_events`／跨年／農曆春節位移）**。

## 6. 本機環境備忘
- Python 路徑：`C:\Users\88691\AppData\Local\Python\pythoncore-3.14-64\python.exe`；PowerShell 5.1 請用 `;` 而非 `&&` 串接指令。
- 專屬技能（按需觸發）：`pms-ht2000-readonly`、`safe-deploy-streamlit`、`vendor-quote-auditor`、`hotel-ui-ux-module`、`siteminder-pricing-audit`。
