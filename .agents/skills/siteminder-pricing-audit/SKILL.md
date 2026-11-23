---
name: siteminder-pricing-audit
description: >-
  Use this skill when auditing SiteMinder inventory grids, channel rates, competitor rates (Insights),
  or building the Revenue Strategy Pricing Audit module in Hotel Master. Strictly READ-ONLY.
---

# SiteMinder 定價稽核與競品比對技能（100% 唯讀）

> **最高安全紅線**：無論是透過瀏覽器代理檢視 SiteMinder 後台，或解析匯出報表，**全程 100% 唯讀**。嚴禁修改任何房價、庫存、Stop Sell、MinLOS，嚴禁點擊 Save / Publish / Update，亦不可點閱使用者瀏覽器中的其他非 SiteMinder 分頁。

## 1. 六大收益策略部定價疏漏檢核規則
將 SiteMinder 掛牌價與限制，與金旭 PMS（`pms-ht2000-readonly`）、去年同期 ADR 及 `taipei_events` 交叉比對：
1. **🔥 錯失漲價（Underpricing）**：當日 OTB ≥ 75% 或撞到大型活動／連假，但 SiteMinder 掛牌價低於競品均價 ≥ 15% 或低於去年同期成交 ADR。
2. **❄️ 虛高滯銷（Overpricing）**：距入住 ≤ 14 天、該房型空房仍多且 OTB 落後去年同期 > 15%，但掛牌價高於競品中位數 ≥ 20%。
3. **📉 12 房型價差倒掛（Rate Inversion）**：依 `HTRT` 定價階梯檢查是否出現高階房型（如 `DF / F / DTP / PT`）售價低於低階房型（如 `S / SP / PD`）之異常。
4. **🚫 幽靈關房／誤鎖限制（Ghost Stop-Sell / MinLOS）**：金旭 PMS 顯示該房型剩餘 ≥ 4 間且非旺日，卻在 SiteMinder 被設為 Stop Sell 或 MinLOS ≥ 2。
5. **⚖️ 渠道破價（Rate Parity Leak）**：官網直訂價高於 OTA 售價，或高抽成渠道售價低於低抽成渠道。
6. **⚡ 誤輸入防呆（Fat-Finger）**：相鄰日或同週基準下，單日售價無預警暴跌 > 40%（如少打一個零）。

## 2. 串接方式優先序
- **方式 A（最穩）**：由 SiteMinder 匯出 CSV / Excel 上傳至 Hotel Master 解析。
- **方式 B（免匯出鍵）**：本機安裝 Node.js LTS 並啟用 Chrome DevTools 遠端偵錯（`127.0.0.1:9222`）後，以唯讀瀏覽器代理擷取 Inventory Grid 與 Insights 競品數據。
