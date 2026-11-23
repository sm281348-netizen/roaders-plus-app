---
name: hotel-ui-ux-module
description: >-
  Use this skill when building a new interactive Streamlit page/module for Hotel Master,
  upgrading UI/UX dashboards, adding export features, or maintaining mobile_ui.py consistency.
---

# Hotel Master 獨立分頁 UI/UX 建置規範

## 1. 模組架構鐵則（禁止塞進 `app.py`）
- 新專案頁面一律建立獨立檔案（如 `wallpaper_planner.py`、`siteminder_audit.py`），透過 `render_<module_name>()` 對外暴露。
- `app.py` 僅於 `menu_options` 新增頁籤名稱，並於底部加入路由分支；若 `current_hotel == "主題館"` 且無該館數據，必須呼叫 `_render_theme_unavailable_notice()`。

## 2. 視覺與資訊層級標準
- **頂部 Hero Banner**：深色高對比漸層（`#1e293b → #0f172a`）配左側強調色條（`#f59e0b` 小鹿金橘），一句話交代規格、資料基準日與範圍。
- **全域參數模擬控制台**：凡涉及報價、預算或分期規則的模組，頂部必設可收合 `st.expander` 參數面板，所有分頁圖表與表格即時連動重算。
- **4 欄 Executive KPI 卡**：白底圓角（`12px`）＋頂部 4px 語意色條（藍規模／橘耗量／綠預算／紫單位造價）。
- **狀態紅綠燈一致性**：🔴 緊急/超賣/超支 (`#ef4444`)、🟠 重度/警告 (`#f97316`)、🟡 緊繃/中度 (`#eab308`)、🟢 安全/達標 (`#10b981`)。

## 3. 跨部門落地標配功能
- **可互動篩選＋勾選試算**：優先以 `st.data_editor` 提供勾選欄，讓經理自由組合房號或品項即時看成本。
- **一鍵交辦匯出**：至少提供多頁簽專業 Excel（`.xlsx`，`openpyxl` 記憶體產出）；涉及現場照片時，提供 `ImageColumn` 表格縮圖預覽與**自動分資料夾 ZIP 打包下載**（檔名自動帶入房號、房型與拆裝標籤）。
- **媒體輕量化**：提交至 Git 的參考照片須經 `ImageOps.exif_transpose` 轉正並壓縮至 `max 1080px / Q82`（單張 ~100 KB），兼顧清晰度與雲端載入速度。
- **手機版相容**：若為日常高頻營運指標，同步評估是否將摘要 KPI 納入 `mobile_ui.py`（`?view=mobile`）。
