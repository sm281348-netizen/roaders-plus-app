# 金旭 PMS (HT2000) 站前館核心資料表字典（唯讀參考）

## 1. `HTRT`（房型主檔）
| 欄位 | 說明 | 備註 |
| :--- | :--- | :--- |
| `RT01` | 房型代碼 | 現行 12 種：`DF, F, DTP, LDT, DT, E, ER, PD, PT, SP, S, SD`（`DR` 總數為 0，已併入 `DT`） |
| `RT02` | 顯示排序 | `'01'` ～ `'14'` |
| `RT03` | 全館總房數 | 12 種現行房型加總 ＝ `128` 間 |
| `RT04` | 定價 (Rack Rate) | 依高至低反映房型定價階梯 |
| `RT05` | 標準入住人數 | 2 / 3 / 4 人 |

## 2. `HTMSB`（每日房態主檔 — 取得 128 房號對應房型）
```sql
SELECT DISTINCT LTRIM(RTRIM(MSB03)) AS room_no, LTRIM(RTRIM(MSB04)) AS room_type
FROM HTMSB
WHERE MSB01 = (SELECT MAX(MSB01) FROM HTMSB)
  AND LTRIM(RTRIM(MSB03)) <> ''
ORDER BY room_no;
```

## 3. `HTRV`（預約主檔）＋ `HTGR`（渠道／合約主檔）＋ `HTAC`（每日房租明細）
| 資料表.欄位 | 說明 | 實戰驗證要點 |
| :--- | :--- | :--- |
| `HTRV.RV01` | 預約單號 | JOIN `HTAC.AC01` |
| `HTRV.RV001` | **下訂日期 (YYYYMMDD)** | 計算 Lead Time (`DATEDIFF(day, RV001, RV03)`) 與 STLY Pace |
| `HTRV.RV005` | **最後異動／取消日期** | 還原去年今日 STLY Pace：`RV06='CXNL' AND RV005 > as_of_ly` |
| `HTRV.RV03` / `RV04` | 入住日 / 退房日 | 計算 ALOS (`DATEDIFF(day, RV03, RV04)`)；注意須單表算避免 JOIN `HTAC` 膨脹 |
| `HTRV.RV06` | 訂單狀態 | `'ACTV'` 有效、`'CXNL'` 已取消、`'NOSH'` No-Show |
| `HTRV.RV10` | **通路代碼** | JOIN `HTGR.GR01` 取得全名 `GR03` 與付款別 `GR07` (`CARD`預付淨價 / `OWN`到付毛價) |
| `HTAC.AC02` / `AC03` | 入住日 / 房型代碼 | 資料完整涵蓋 `20240101` 至 `202707+` |
| `HTAC.AC06` / `AC07` | **有效房數 / 取消保留房數** | `ACTV` 存於 `AC06`；取消轉為 `CXNL` 時 `AC06` 歸零並完整轉存至 **`AC07`**！ |
| `HTAC.AC08` | 每日房租金額 | 即使 `CXNL` 仍保留原房租；`預付`代碼為扣抽成 Net、`到付`為毛價 Gross |

### 站前館核心通路代碼（實測與 SiteMinder 100% 吻合）
- `4251987901` (`【AGODA預付】Net`) / `4251987902` (`【AGODA到付】Gross`)
- `ctrip01` (`【CTRIP預付】Net`) / `Ctrip02` (`【CTRIP到付】Gross`)
- `Booking` (`【BOOKING.COM】Gross 到付`)
- `TBB` (`【BookingButton】官網直訂 Direct`)
- `EXPEDIA01` (`【EXPEDIA預付】Net`) / `EXPEDIA02` (`【EXPEDIA到付】Gross`)
- `9030352601` (`【RAKUTEN】`) / `82883867` (`【AsiaYo】`) / `24763920` (`【FunNow】`)

## 4. `HTRB`（維修與鎖房 OOO）
```sql
SELECT LTRIM(RTRIM(RB03)) AS block_date,
       LTRIM(RTRIM(RB04)) AS room_no,
       LTRIM(RTRIM(RB06)) AS reason
FROM HTRB
WHERE RB03 >= '20261120' AND RB03 <= '20270331'
  AND RB05 = 'OOO';
```
- 常態固定 OOO：`2509 (SD)`、`2510 (SD)`（布巾房），故常態可售為 `126 間`。
- 單日特定房型剩餘可售公式：`Remaining(date, rt) = Total(rt) - OOO(date, rt) - Sold(date, rt)`。
