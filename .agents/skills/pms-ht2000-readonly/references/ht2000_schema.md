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

## 3. `HTAC` JOIN `HTRV`（未來在手預訂 OTB）
```sql
SELECT LTRIM(RTRIM(a.AC02)) AS stay_date,
       LTRIM(RTRIM(a.AC03)) AS room_type,
       SUM(a.AC06) AS sold_rooms,
       SUM(a.AC08) AS room_revenue
FROM HTAC a
JOIN HTRV v ON a.AC01 = v.RV01
WHERE a.AC02 >= '20261120' AND a.AC02 <= '20270331'
  AND v.RV06 = 'ACTV'
GROUP BY a.AC02, a.AC03;
```

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
