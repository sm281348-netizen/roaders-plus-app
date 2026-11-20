# -*- coding: utf-8 -*-
"""
路徒 PLUS 行旅（站前館）客房壁紙修繕規劃、報價試算與房型庫存排程模組
====================================================================
功能模組：
1. 內建 44 間房（55 面有效牆面，已剔除重複拍攝之 3109-3.jpg）完整實測數據、嚴重度、換牆數、電視拆裝與金旭 PMS 房型。
2. 頂部「報價與預算動態模擬器」：可即時調整單捲壁紙價、師傅出班費、單/雙師傅產能、算料模式（實測 FFD / 同樓層 / 廠商每面牆固定支數）、備品與含稅設定。
3. Tab 1：44 房優先順序多維度篩選、排序與自訂勾選即時試算。
4. Tab 2：報價對比圖表與「每月預算上限」分期規劃模擬。
5. Tab 3：金旭 PMS 即時房型庫存熱力圖 × What-If 壓房衝突即時檢核器（含區網一鍵重拉 PMS 按鈕）。
6. Tab 4：同樓層重災熱區與工程/房務防坑警示總覽。
7. Tab 5：一鍵匯出 4 分頁 Excel（.xlsx）報價與派工工作簿。
"""

import os
import io
import json
import math
import base64
import zipfile
import datetime
from collections import defaultdict

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image

APP_DIR = os.path.dirname(os.path.abspath(__file__))
INV_JSON_PATH = os.path.join(APP_DIR, "data", "pms_backup", "wallpaper_roomtype_inventory.json")
PHOTO_DIR = os.path.join(APP_DIR, "data", "wallpaper_photos")

# ==============================================================================
# 1. 經逐張照片驗證之 55 面有效待修牆面主檔（44 間房，已剔除重複照 3109-3.jpg）
# ==============================================================================
WALLS_DATA = [
    # 24F
    {"file": "2401.jpg",   "room": "2401", "floor": 24, "raw": "2.49x1.66m", "h": 249, "w": 166, "loc": "一般牆面(行李架旁)",    "tv": False, "desk": False, "note": "靠地垂直細裂刮傷"},
    {"file": "2403.jpg",   "room": "2403", "floor": 24, "raw": "2.53x1.73m", "h": 253, "w": 173, "loc": "書桌後方牆面",          "tv": False, "desk": True,  "note": "插座上方水平細裂、4組插座面板"},
    # 25F
    {"file": "2501.jpg",   "room": "2501", "floor": 25, "raw": "2.75*2.5m",  "h": 250, "w": 275, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視右下至插座斜向裂痕、拆電視+2組面板"},
    {"file": "2503.jpg",   "room": "2503", "floor": 25, "raw": "1.7*2.78m",  "h": 278, "w": 170, "loc": "書桌後方挑高牆",        "tv": False, "desk": True,  "note": "挑高278cm、插座上方水平裂痕、4組面板"},
    {"file": "2505.jpg",   "room": "2505", "floor": 25, "raw": "3.4*2.57m",  "h": 257, "w": 340, "loc": "行李架/矮櫃後方長牆",   "tv": False, "desk": False, "note": "多處撞擊點狀破損、拆壁掛飾品+2組面板"},
    {"file": "2508.jpg",   "room": "2508", "floor": 25, "raw": "2.57*3.26m", "h": 257, "w": 326, "loc": "電視牆連書桌牆",        "tv": True,  "desk": True,  "note": "電視左上轉角裂痕+下方破損、拆電視+書桌+4組面板"},
    {"file": "2508-2.jpg", "room": "2508", "floor": 25, "raw": "0.7*2.3m",   "h": 230, "w": 70,  "loc": "床頭側窗邊窄牆(發霉)",  "tv": False, "desk": False, "note": "★牆面發霉須先除霉＋防霉底塗乾燥"},
    # 26F
    {"file": "2601.jpg",   "room": "2601", "floor": 26, "raw": "2.57x2.7m",  "h": 257, "w": 270, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視左側水平裂+右下長條損傷、拆電視"},
    {"file": "2603.jpg",   "room": "2603", "floor": 26, "raw": "2.1x2.41m",  "h": 210, "w": 241, "loc": "走道/客房牆面",         "tv": False, "desk": False, "note": "地毯側長條刮痕碰傷、1組面板(高寬建議複核)"},
    {"file": "2603-2.jpg", "room": "2603", "floor": 26, "raw": "2.1mX87cm",  "h": 210, "w": 87,  "loc": "玄關控制面板窄牆",      "tv": False, "desk": False, "note": "★玄關面板旁垂直水漬污漬、4組溫控插卡面板"},
    {"file": "2605.jpg",   "room": "2605", "floor": 26, "raw": "1.4*2.58m",  "h": 258, "w": 140, "loc": "吊管電視區靠窗牆",      "tv": True,  "desk": False, "note": "天花板吊管電視下方斜裂、1組插座"},
    {"file": "2606.jpg",   "room": "2606", "floor": 26, "raw": "0.35*2.48m", "h": 248, "w": 35,  "loc": "窗邊突柱正面",          "tv": False, "desk": False, "note": "35cm窄柱底部磨損（極小免拆家具可順修）"},
    # 27F
    {"file": "2702.jpg",   "room": "2702", "floor": 27, "raw": "2.48*1.6m",  "h": 248, "w": 160, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視下方大片髒污刮痕+右側裂縫、拆電視+1面板"},
    {"file": "2702-2.jpg", "room": "2702", "floor": 27, "raw": "2.5*2.19m",  "h": 219, "w": 250, "loc": "床側至書桌樑下凹牆",    "tv": False, "desk": True,  "note": "中段與牆底兩處破損、移書桌+4組面板"},
    {"file": "2707.jpg",   "room": "2707", "floor": 27, "raw": "1.96mX88cm", "h": 196, "w": 88,  "loc": "行李架上方主牆",        "tv": False, "desk": False, "note": "自插座貫穿至牆角斜向長裂縫、2組面板"},
    {"file": "2707-2.jpg", "room": "2707", "floor": 27, "raw": "1.96mX53cm", "h": 196, "w": 53,  "loc": "玄關房門旁轉角牆",      "tv": False, "desk": False, "note": "下方刮痕破損、4組溫控插卡面板"},
    {"file": "2707-3.jpg", "room": "2707", "floor": 27, "raw": "2.19mX64cm", "h": 219, "w": 64,  "loc": "書桌旁突出窄柱牆",      "tv": False, "desk": True,  "note": "椅背高度撞傷刮痕、緊鄰書桌收邊"},
    {"file": "2711.jpg",   "room": "2711", "floor": 27, "raw": "3*2.3m",     "h": 230, "w": 300, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視下方近地毯污損刮痕、拆電視+1面板"},
    # 28F
    {"file": "2806.jpg",   "room": "2806", "floor": 28, "raw": "1.45mx10cm", "h": 10,  "w": 145, "loc": "落地窗下方矮窗台條",    "tv": False, "desk": False, "note": "★極小局部10x145cm破皮（餘料順手修免壓2天）"},
    {"file": "2810.jpg",   "room": "2810", "floor": 28, "raw": "2.29x1.79m", "h": 229, "w": 179, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視左下刮痕+牆底多處破皮、拆電視+1面板"},
    {"file": "2811.jpg",   "room": "2811", "floor": 28, "raw": "2.07x1.59m", "h": 207, "w": 159, "loc": "玄關門旁牆面",          "tv": False, "desk": False, "note": "近地毯多點破損、3組面板+確認壁掛物"},
    {"file": "2811-2.jpg", "room": "2811", "floor": 28, "raw": "2.21x2.68m", "h": 221, "w": 268, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視正下方明顯裂縫、拆電視"},
    {"file": "2812.jpg",   "room": "2812", "floor": 28, "raw": "2.48x4.06m", "h": 248, "w": 406, "loc": "書桌側長牆面",          "tv": False, "desk": True,  "note": "4.06m長牆自開關向右長裂縫、移書桌"},
    # 29F
    {"file": "2902.jpg",   "room": "2902", "floor": 29, "raw": "3.5*2m",     "h": 200, "w": 350, "loc": "床側至書桌主牆",        "tv": False, "desk": True,  "note": "左上接縫破損+右下磨損、高度200cm複核樑高+3面板"},
    {"file": "2903.jpg",   "room": "2903", "floor": 29, "raw": "2.4x2.5m",   "h": 250, "w": 240, "loc": "客房牆面",              "tv": False, "desk": False, "note": "牆面明顯起泡、刮傷與剝落"},
    {"file": "2905.jpg",   "room": "2905", "floor": 29, "raw": "1.3*2.5m",   "h": 250, "w": 130, "loc": "窗簾側茶水架/書桌牆",   "tv": False, "desk": True,  "note": "中上段水平摺痕/接縫異常、4組面板"},
    {"file": "2906.jpg",   "room": "2906", "floor": 29, "raw": "2.4*2.55m",  "h": 255, "w": 240, "loc": "洗手台右側至茶水檯牆",  "tv": False, "desk": True,  "note": "左側下緣近地毯局部破損"},
    {"file": "2907.jpg",   "room": "2907", "floor": 29, "raw": "2.5*1.25m",  "h": 250, "w": 125, "loc": "窗戶右側至牆角",        "tv": False, "desk": False, "note": "窗台下角小裂痕與牆底磨損（無障礙物）"},
    {"file": "2910.jpg",   "room": "2910", "floor": 29, "raw": "2.29x1.79m", "h": 229, "w": 179, "loc": "玄關門旁掛衣板牆",      "tv": False, "desk": False, "note": "開關周圍刮痕髒污、拆掛衣板+1開關"},
    # 30F
    {"file": "3001.jpg",   "room": "3001", "floor": 30, "raw": "2.6*2.6m",   "h": 260, "w": 260, "loc": "陽台門側電視主牆",      "tv": True,  "desk": False, "note": "左下方近地毯破損、拆壁掛電視+壁掛掛鉤"},
    # 31F (已剔除重複照 3109-3.jpg)
    {"file": "3103.jpg",   "room": "3103", "floor": 31, "raw": "2.3*2.55m",  "h": 255, "w": 230, "loc": "造型掛衣板牆面",        "tv": False, "desk": False, "note": "掛衣板下方近地毯局部破損、拆2片掛衣板"},
    {"file": "3105.jpg",   "room": "3105", "floor": 31, "raw": "1.3*2.5m",   "h": 250, "w": 130, "loc": "窗簾側書桌/層架牆",     "tv": False, "desk": True,  "note": "中段橫向起皺/裂痕、4組開關插座"},
    {"file": "3106.jpg",   "room": "3106", "floor": 31, "raw": "2.4*2.55m",  "h": 255, "w": 240, "loc": "電視/書桌主牆",         "tv": True,  "desk": True,  "note": "最左側上下貫穿長裂（若只換左側1~2條可免拆電視!）"},
    {"file": "3109.jpg",   "room": "3109", "floor": 31, "raw": "1.9*0.9m",   "h": 190, "w": 90,  "loc": "窗邊書桌旁窄牆",        "tv": False, "desk": True,  "note": "書桌椅旁牆底破損（高度190cm建議現場複尺）"},
    {"file": "3109-2.jpg", "room": "3109", "floor": 31, "raw": "2.27*1.87m", "h": 227, "w": 187, "loc": "床尾獨立隔間牆",        "tv": False, "desk": False, "note": "中段「1」字刮痕+底部破皮（已剔除重複照3109-3）"},
    {"file": "3111.jpg",   "room": "3111", "floor": 31, "raw": "2.13x2.59m", "h": 213, "w": 259, "loc": "床尾電視主牆",          "tv": True,  "desk": False, "note": "電視正下方汙損與細裂痕、拆電視+1開關"},
    {"file": "3111-2.jpg", "room": "3111", "floor": 31, "raw": "2.13x1.69m", "h": 213, "w": 169, "loc": "玄關溫控/掛衣板牆",     "tv": False, "desk": False, "note": "★玄關開關水漬脫膠髒污、4組面板+2掛衣板"},
    {"file": "3112.jpg",   "room": "3112", "floor": 31, "raw": "2.15x1.49m", "h": 215, "w": 149, "loc": "玄關溫控開關牆",        "tv": False, "desk": False, "note": "開關向右橫向刮裂、4組面板+鄰全身鏡"},
    # 32F
    {"file": "3211.jpg",   "room": "3211", "floor": 32, "raw": "2.09x2.75m", "h": 209, "w": 275, "loc": "床尾電視主牆",          "tv": True,  "desk": False, "note": "電視上方V字痕+正下方深色汙痕+右側長裂縫、拆電視"},
    # 33F
    {"file": "3301.jpg",   "room": "3301", "floor": 33, "raw": "2.45x2.35m", "h": 245, "w": 235, "loc": "素面主牆(靠窗簾)",      "tv": False, "desk": False, "note": "中段橫向裂痕（無障礙物）"},
    {"file": "3301-2.jpg", "room": "3301", "floor": 33, "raw": "2.22mX45cm", "h": 222, "w": 45,  "loc": "窗框與門框間窄牆",      "tv": False, "desk": False, "note": "開關上方污漬+斜向裂痕、2組面板"},
    {"file": "3302.jpg",   "room": "3302", "floor": 33, "raw": "2.07x1.56m", "h": 207, "w": 156, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "上方整排橫裂+電視左下方直達地毯破損、拆電視（現修冷氣OOO）"},
    {"file": "3303.jpg",   "room": "3303", "floor": 33, "raw": "2.11mX53cm", "h": 211, "w": 53,  "loc": "玄關門旁控制面板窄牆",  "tv": False, "desk": False, "note": "面板群周邊橫裂污損、4組溫控插卡面板"},
    {"file": "3303-1.jpg", "room": "3303", "floor": 33, "raw": "2.11x2.28m", "h": 211, "w": 228, "loc": "掛衣區牆面",            "tv": False, "desk": False, "note": "牆底多處刮傷破損、拆2片掛衣板+1開關"},
    {"file": "3305.jpg",   "room": "3305", "floor": 33, "raw": "2.25x1.41m", "h": 225, "w": 141, "loc": "窗簾與書桌/吧台側牆",   "tv": False, "desk": True,  "note": "中段橫向裂痕、4組開關插座"},
    {"file": "3307.jpg",   "room": "3307", "floor": 33, "raw": "2.49x4.93m", "h": 249, "w": 493, "loc": "客房長側牆(掛衣區)",    "tv": False, "desk": False, "note": "全館最大單牆4.93m、自天花貫穿至地毯直向大裂縫+2掛衣板"},
    {"file": "3311.jpg",   "room": "3311", "floor": 33, "raw": "2.29x2.45m", "h": 229, "w": 245, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視正下方中央直向裂痕、拆電視"},
    # 34F
    {"file": "3406.jpg",   "room": "3406", "floor": 34, "raw": "2.36x4.86m", "h": 236, "w": 486, "loc": "客房長向大面牆",        "tv": False, "desk": False, "note": "4.86m長牆近地毯下緣刮痕脫皮、1組插座"},
    {"file": "3408.jpg",   "room": "3408", "floor": 34, "raw": "2.37x3.82m", "h": 237, "w": 382, "loc": "電視暨掛衣板主牆",      "tv": True,  "desk": True,  "note": "電視左上直裂+掛衣下方磨損、拆電視+2掛衣板+書桌溫控"},
    {"file": "3411.jpg",   "room": "3411", "floor": 34, "raw": "2.16x2.61m", "h": 216, "w": 261, "loc": "浴廁門旁至電視主牆",    "tv": True,  "desk": False, "note": "門框上方直裂+電視上方直裂、拆電視（複核門洞寬）"},
    {"file": "3412.jpg",   "room": "3412", "floor": 34, "raw": "2.23x3.14m", "h": 223, "w": 314, "loc": "電視主牆",              "tv": True,  "desk": False, "note": "電視上方長條斜向裂縫、拆電視"},
    # 35F
    {"file": "3503.jpg",   "room": "3503", "floor": 35, "raw": "2.35x2.4m",  "h": 235, "w": 240, "loc": "客房牆面",              "tv": False, "desk": False, "note": "近地板下緣多處刮傷破皮、1組雙切開關"},
    {"file": "3506.jpg",   "room": "3506", "floor": 35, "raw": "2.06mx52cm", "h": 206, "w": 52,  "loc": "窗簾與床旁窄柱面",      "tv": False, "desk": False, "note": "52cm窄柱近地面破損脫皮"},
    {"file": "3506-2.jpg", "room": "3506", "floor": 35, "raw": "2.2x4.18m",  "h": 220, "w": 418, "loc": "掛衣板暨書桌/電視長牆", "tv": True,  "desk": True,  "note": "4.18m主牆髒污刮損、拆掛衣板+書桌+電視"},
    {"file": "3511.jpg",   "room": "3511", "floor": 35, "raw": "1.98x2.44m", "h": 198, "w": 244, "loc": "電視主牆(水漬起泡)",    "tv": True,  "desk": False, "note": "★牆底大面積水漬起泡（須查滲水）+高度198cm複尺、拆電視"},
]

# 44 房嚴重度等級與電視拆裝細分類
ROOM_SEVERITY = {
    # S 級（發霉 / 水漬起泡脫膠）
    "2508": "S", "3111": "S", "2603": "S", "3511": "S",
    # A 級（貫穿大裂縫 / 剝落 / 多牆重災 / 電視正中明顯汙損）
    "2707": "A", "3506": "A", "2702": "A", "2811": "A", "3408": "A",
    "3106": "A", "3211": "A", "2810": "A", "3302": "A", "3307": "A", "2903": "A",
    # B 級（主牆/電視牆視線高度中度裂痕或撞傷）
    "3301": "B", "3109": "B", "3303": "B", "3412": "B", "2601": "B",
    "2711": "B", "2501": "B", "3001": "B", "3411": "B", "3311": "B",
    "2605": "B", "2812": "B", "2505": "B", "2902": "B", "3105": "B", "3305": "B",
    # C 級（牆底磨損 / 插座細裂 / 小柱窗台輕微破損）
    "3406": "C", "2906": "C", "3103": "C", "3503": "C", "2503": "C",
    "2403": "C", "2401": "C", "2910": "C", "2905": "C", "3112": "C",
    "2907": "C", "2606": "C", "2806": "C",
}

SEV_ORDER = {"S": 1, "A": 2, "B": 3, "C": 4}
SEV_LABEL = {
    "S": "🔴 S級(發霉/水漬急症)",
    "A": "🟠 A級(貫穿裂/多牆重災)",
    "B": "🟡 B級(主牆/電視牆中度)",
    "C": "🟢 C級(牆底/小柱輕微)",
}

# 預設 PMS 房型對照（備援，會優先由 live JSON 覆蓋）
DEFAULT_ROOM_TO_RT = {
    "3001": "DF", "3301": "DF",
    "2508": "F",  "3109": "F",
    "2401": "DTP","2501": "DTP","2601": "DTP",
    "2505": "DT", "2606": "DT", "2707": "DT", "2806": "DT", "2906": "DT", "2907": "DT", "3106": "DT",
    "2905": "E",  "3105": "E",  "3305": "E",
    "2902": "ER", "3302": "ER",
    "2403": "PD", "2503": "PD", "2605": "PD", "2702": "PD", "2812": "PD", "3112": "PD", "3412": "PD",
    "3307": "PT", "3406": "PT", "3408": "PT", "3506": "PT",
    "2711": "SP", "2810": "SP", "2811": "SP", "2910": "SP", "3111": "SP", "3211": "SP", "3311": "SP", "3411": "SP", "3511": "SP",
    "2603": "S",  "2903": "S",  "3103": "S",  "3303": "S",  "3503": "S",
}

DEFAULT_RT_TOTAL = {
    "DF": 8, "F": 8, "DTP": 4, "LDT": 4, "DT": 19, "E": 8,
    "ER": 8, "PD": 22, "PT": 12, "SP": 20, "S": 10, "SD": 5,
}

# 已通過房型零超賣驗證的 6 個標準批次（對應「每月 NT$30,000 × 三個月完修」）
VERIFIED_PLAN_3MO = [
    {"batch": "12月-第1批 (急件首發)", "month": "2026-12", "d1": "2026-12-07", "d2": "2026-12-08",
     "rooms": ["3111", "3211", "2603", "2903", "3106", "2702", "3302"],
     "note": "避開 PT(12/7滿房) 與 DTP(緊繃)，全房型壓後皆留 ≥2 間餘裕"},
    {"batch": "12月-第2批 (發霉/水漬/PT)", "month": "2026-12", "d1": "2026-12-15", "d2": "2026-12-16",
     "rooms": ["2508", "3511", "3506", "3408", "2810", "2707"],
     "note": "F(剩2/3)、PT(剩5/3)、SP(剩10/14) 同步大解放，搶修發霉2508與水漬3511"},
    {"batch": "1月-第1批 (淡季谷底)", "month": "2027-01", "d1": "2027-01-18", "d2": "2027-01-19",
     "rooms": ["2811", "3411", "3001", "2605", "2812", "2505", "2902", "3105"],
     "note": "1 月最深低谷週（去年空房 18~23 間），清光中重度房"},
    {"batch": "1月-第2批 (DTP安全窗)", "month": "2027-01", "d1": "2027-01-25", "d2": "2027-01-26",
     "rooms": ["3307", "3109", "3301", "3303", "3412", "2601", "2711", "3311"],
     "note": "DTP 剩 2/2 可安全壓 2601，農曆年前客訴點歸零"},
    {"batch": "2月-第1批 (年後收尾)", "month": "2027-02", "d1": "2027-02-22", "d2": "2027-02-23",
     "rooms": ["2501", "3305", "3406", "3503", "3103", "3112", "2806"],
     "note": "避開 2/5~14 春節與 228 連假，DTP 剩 2/3 安全壓 2501"},
    {"batch": "2月-第2批 (輕傷完修)", "month": "2027-02", "d1": "2027-02-24", "d2": "2027-02-25",
     "rooms": ["2401", "2606", "2906", "2907", "2905", "2910", "2503", "2403"],
     "note": "0 台電視！DTP 剩 3/3 壓 2401，趕在 3 月旺季前 100% 完工"},
]

VERIFIED_PLAN_DEC20 = [
    {"batch": "12月衝刺-第1批", "month": "2026-12", "d1": "2026-12-07", "d2": "2026-12-08",
     "rooms": ["3111", "3211", "2603", "2903", "3106", "2702", "3302"],
     "note": "全房型壓後餘 ≥2 間"},
    {"batch": "12月衝刺-第2批", "month": "2026-12", "d1": "2026-12-14", "d2": "2026-12-15",
     "rooms": ["3307", "3408", "3412", "3301", "3303", "2810", "2707"],
     "note": "全房型壓後餘 ≥1 間；12/09 確認 12/14 OTB ≤78%"},
    {"batch": "12月衝刺-第3批", "month": "2026-12", "d1": "2026-12-16", "d2": "2026-12-17",
     "rooms": ["2508", "3511", "3506", "2811", "3411", "2711"],
     "note": "2508(F)於12/17壓後餘0須立即鎖房；SP剩14/11極寬裕"},
]


# ==============================================================================
# 2. 演算法引擎：直條排料 FFD、房型庫存載入與即時 PMS 重拉
# ==============================================================================

def calc_ffd_rolls(strip_lengths_cm, roll_len_cm=1000):
    """First-Fit Decreasing 直條排料演算法（不橫向拼接）"""
    if not strip_lengths_cm:
        return 0, 0
    bins = []
    for length in sorted(strip_lengths_cm, reverse=True):
        placed = False
        for i in range(len(bins)):
            if bins[i] >= length:
                bins[i] -= length
                placed = True
                break
        if not placed:
            bins.append(roll_len_cm - length)
    return len(bins), sum(bins)


def compute_selection_metrics(selected_rooms, walls_df, roll_mode="ffd_cross", fixed_rolls_per_wall=2.5,
                              trim_cm=10, roll_price=441, worker_fee=4500, cap_1=3.0, cap_2=8.0,
                              spare_rolls=0, misc_fee=0, include_tax=False):
    """根據任意勾選房號與報價參數，即時計算支數、工資與總預算"""
    sub = walls_df[walls_df["room"].isin(selected_rooms)]
    n_rooms = len(set(selected_rooms))
    n_walls = len(sub)
    if n_rooms == 0 or n_walls == 0:
        return {
            "n_rooms": 0, "n_walls": 0, "area_m2": 0.0, "area_ping": 0.0, "strips": 0,
            "rolls_ffd": 0, "rolls_floor": 0, "rolls_area15": 0, "rolls_fixed": 0,
            "rolls_used": 0, "rolls_total": 0, "mat_cost": 0,
            "days_2m": 0, "labor_2m": 0, "days_1m": 0, "labor_1m": 0,
            "subtotal": 0, "tax": 0, "grand_total": 0,
            "cost_per_room": 0, "cost_per_ping": 0, "tv_rooms": 0, "multi_wall_rooms": 0,
        }

    area_m2 = float(sub["area_m2"].sum())
    area_ping = area_m2 * 0.3025
    all_strips = []
    for _, row in sub.iterrows():
        st_cnt = int(row["strips"])
        all_strips.extend([int(row["h"]) + trim_cm] * st_cnt)

    rolls_ffd, _ = calc_ffd_rolls(all_strips, 1000)

    # 同樓層共料 FFD
    rolls_floor = 0
    for _, grp in sub.groupby("floor"):
        fl_strips = []
        for _, row in grp.iterrows():
            fl_strips.extend([int(row["h"]) + trim_cm] * int(row["strips"]))
        r_cnt, _ = calc_ffd_rolls(fl_strips, 1000)
        rolls_floor += r_cnt

    rolls_area15 = int(math.ceil((area_m2 / 5.3) * 1.15))
    rolls_fixed = int(math.ceil(n_walls * fixed_rolls_per_wall))

    if roll_mode == "ffd_cross":
        rolls_used = rolls_ffd
    elif roll_mode == "ffd_floor":
        rolls_used = rolls_floor
    elif roll_mode == "area_15":
        rolls_used = rolls_area15
    else:
        rolls_used = rolls_fixed

    rolls_total = rolls_used + int(spare_rolls)
    mat_cost = rolls_total * roll_price

    # 雙師傅 vs 單師傅出班天數與工資（若尾數 ≤ cap_1 可混搭 1 名師傅省半天）
    days_2m_full = int(math.ceil(n_rooms / max(cap_2, 1.0)))
    labor_2m_only = days_2m_full * 2 * worker_fee

    # 混搭最優化：例如 10 間 = 1 天雙師傅(7~8) + 1 天單師傅(3) = NT$13,500 比 2 天雙師傅(NT$18,000) 省!
    best_labor = labor_2m_only
    best_plan_str = f"雙師傅 {days_2m_full} 天"
    for d2 in range(days_2m_full + 1):
        rem_r = max(0.0, n_rooms - d2 * cap_2)
        d1 = int(math.ceil(rem_r / max(cap_1, 1.0))) if rem_r > 0.01 else 0
        cost = (d2 * 2 + d1) * worker_fee
        if cost < best_labor:
            best_labor = cost
            if d2 > 0 and d1 > 0:
                best_plan_str = f"雙師傅 {d2} 天 + 單師傅 {d1} 天"
            elif d2 > 0:
                best_plan_str = f"雙師傅 {d2} 天"
            else:
                best_plan_str = f"單師傅 {d1} 天"

    days_1m = int(math.ceil(n_rooms / max(cap_1, 1.0)))
    labor_1m = days_1m * worker_fee

    subtotal = mat_cost + best_labor + int(misc_fee)
    tax = int(round(subtotal * 0.05)) if include_tax else 0
    grand_total = subtotal + tax

    # 計算電視房與多牆房數
    rm_grp = sub.groupby("room")
    tv_rooms = int(sum(1 for _, g in rm_grp if g["tv"].any()))
    multi_wall_rooms = int(sum(1 for _, g in rm_grp if len(g) >= 2))

    return {
        "n_rooms": n_rooms, "n_walls": n_walls,
        "area_m2": area_m2, "area_ping": area_ping, "strips": len(all_strips),
        "rolls_ffd": rolls_ffd, "rolls_floor": rolls_floor,
        "rolls_area15": rolls_area15, "rolls_fixed": rolls_fixed,
        "rolls_used": rolls_used, "rolls_total": rolls_total, "mat_cost": mat_cost,
        "days_2m": days_2m_full, "labor_2m": best_labor, "labor_plan_str": best_plan_str,
        "days_1m": days_1m, "labor_1m": labor_1m,
        "subtotal": subtotal, "tax": tax, "grand_total": grand_total,
        "cost_per_room": grand_total / n_rooms if n_rooms else 0,
        "cost_per_ping": grand_total / area_ping if area_ping else 0,
        "tv_rooms": tv_rooms, "multi_wall_rooms": multi_wall_rooms,
    }


@st.cache_data(ttl=300)
def load_roomtype_inventory():
    """載入本地保存的金旭 PMS 房型庫存快照 JSON"""
    if os.path.exists(INV_JSON_PATH):
        try:
            with open(INV_JSON_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "as_of": "未載入",
        "rt_master": {k: {"total_inv": v} for k, v in DEFAULT_RT_TOTAL.items()},
        "room_to_rt": DEFAULT_ROOM_TO_RT,
        "ooo_rooms": {},
        "daily_avail": {},
    }


def refresh_pms_roomtype_inventory():
    """在區網環境下直連金旭 PMS (HT2000) 更新即時房型剩餘庫存並寫入快照檔"""
    from pms_future_extractor import get_sql_connection
    conn = get_sql_connection()
    cur = conn.cursor()

    cur.execute("SELECT RT01, RT02, RT03, RT04, RT05 FROM HTRT ORDER BY RT02")
    rt_master = {}
    for r in cur.fetchall():
        code = str(r[0]).strip()
        inv = int(float(r[2] or 0))
        if inv > 0:
            rt_master[code] = {
                "order": str(r[1]).strip(),
                "total_inv": inv,
                "rack_rate": int(float(r[3] or 0)),
                "pax": int(float(r[4] or 0)),
            }

    cur.execute("SELECT MAX(MSB01) FROM HTMSB")
    max_msb_date = str(cur.fetchone()[0]).strip()
    cur.execute("""
        SELECT DISTINCT LTRIM(RTRIM(MSB03)) AS rm, LTRIM(RTRIM(MSB04)) AS rt
        FROM HTMSB WHERE MSB01 = ? AND LTRIM(RTRIM(MSB03)) <> ''
    """, (max_msb_date,))
    room_to_rt = {rm: rt for rm, rt in cur.fetchall() if rm.isdigit()}

    today = datetime.date.today()
    end_d = datetime.date(2027, 3, 31)
    s_str, e_str = today.strftime("%Y%m%d"), end_d.strftime("%Y%m%d")

    cur.execute("""
        SELECT LTRIM(RTRIM(RB03)), LTRIM(RTRIM(RB04)), LTRIM(RTRIM(RB06))
        FROM HTRB WHERE RB03 >= ? AND RB03 <= ? AND RB05 = 'OOO'
    """, (s_str, e_str))
    ooo_by_date_rt = defaultdict(lambda: defaultdict(int))
    ooo_sample = {}
    for dt, rm, reason in cur.fetchall():
        rt = room_to_rt.get(rm, "UNK")
        ooo_by_date_rt[dt][rt] += 1
        if rm not in ooo_sample:
            ooo_sample[rm] = [rt, reason]

    cur.execute("""
        SELECT LTRIM(RTRIM(a.AC02)), LTRIM(RTRIM(a.AC03)), SUM(a.AC06)
        FROM HTAC a JOIN HTRV v ON a.AC01 = v.RV01
        WHERE a.AC02 >= ? AND a.AC02 <= ? AND v.RV06 = 'ACTV'
        GROUP BY a.AC02, a.AC03
    """, (s_str, e_str))
    sold_by_date_rt = defaultdict(lambda: defaultdict(int))
    for dt, rt, sold in cur.fetchall():
        sold_by_date_rt[dt][rt] = int(round(float(sold or 0)))

    daily_avail = {}
    d = today
    while d <= end_d:
        ds = d.strftime("%Y%m%d")
        day_d = {}
        for rt, info in rt_master.items():
            tot = info["total_inv"]
            ooo = ooo_by_date_rt[ds].get(rt, 0)
            sold = sold_by_date_rt[ds].get(rt, 0)
            day_d[rt] = {"tot": tot, "ooo": ooo, "sold": sold, "rem": tot - ooo - sold}
        daily_avail[d.strftime("%Y-%m-%d")] = day_d
        d += datetime.timedelta(days=1)

    conn.close()

    payload = {
        "as_of": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "msb_date": max_msb_date,
        "rt_master": rt_master,
        "room_to_rt": room_to_rt,
        "ooo_rooms": ooo_sample,
        "daily_avail": daily_avail,
    }
    os.makedirs(os.path.dirname(INV_JSON_PATH), exist_ok=True)
    with open(INV_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    load_roomtype_inventory.clear()
    return payload


def build_dataframes(inv_data):
    """建構 55 面牆 DataFrame 與 44 房彙總優先順序 DataFrame"""
    r2rt = inv_data.get("room_to_rt", DEFAULT_ROOM_TO_RT)
    rt_master = inv_data.get("rt_master", {})

    # 建議壓房日期對照（依月預算3萬 6 批次）
    rec_date_map = {}
    for b in VERIFIED_PLAN_3MO:
        lbl = f"{b['d1'][5:]}→{b['d2'][5:]} ({b['month'][5:]}月)"
        for r in b["rooms"]:
            rec_date_map[r] = lbl

    w_rows = []
    for item in WALLS_DATA:
        area_m2 = round((item["h"] / 100.0) * (item["w"] / 100.0), 2)
        strips = int(math.ceil(item["w"] / 53.0))
        lin_m = round(strips * (item["h"] + 10) / 100.0, 2)
        rm = item["room"]
        rt = r2rt.get(rm, DEFAULT_ROOM_TO_RT.get(rm, "UNK"))
        w_rows.append({
            **item,
            "rt": rt,
            "sev": ROOM_SEVERITY.get(rm, "C"),
            "area_m2": area_m2,
            "area_ping": round(area_m2 * 0.3025, 2),
            "strips": strips,
            "lin_m": lin_m,
        })
    df_walls = pd.DataFrame(w_rows)

    # 44 房彙總表
    rm_rows = []
    for rm, grp in df_walls.groupby("room"):
        sev = ROOM_SEVERITY.get(rm, "C")
        rt = grp["rt"].iloc[0]
        rt_tot = rt_master.get(rt, {}).get("total_inv", DEFAULT_RT_TOTAL.get(rt, 0))
        has_tv = bool(grp["tv"].any())
        has_desk = bool(grp["desk"].any())
        if rm == "3106":
            tv_status = "⚠️ 可視色差免拆"
            tv_rank = 2
        elif rm == "2605":
            tv_status = "✅ 吊管電視"
            tv_rank = 1
        elif has_tv:
            tv_status = "✅ 必拆電視"
            tv_rank = 1
        else:
            tv_status = "❌ 免拆電視"
            tv_rank = 3

        rm_strips = []
        for _, rr in grp.iterrows():
            rm_strips.extend([int(rr["h"]) + 10] * int(rr["strips"]))
        rm_rolls_iso, _ = calc_ffd_rolls(rm_strips, 1000)

        rm_rows.append({
            "room": rm,
            "floor": int(grp["floor"].iloc[0]),
            "rt": rt,
            "rt_label": f"{rt} ({rt_tot}間)",
            "rt_tot": rt_tot,
            "sev": sev,
            "sev_rank": SEV_ORDER[sev],
            "sev_label": SEV_LABEL[sev],
            "wall_count": int(len(grp)),
            "has_tv": has_tv,
            "tv_status": tv_status,
            "tv_rank": tv_rank,
            "has_desk": has_desk,
            "area_m2": round(float(grp["area_m2"].sum()), 2),
            "area_ping": round(float(grp["area_m2"].sum()) * 0.3025, 2),
            "strips": int(grp["strips"].sum()),
            "lin_m": round(float(grp["lin_m"].sum()), 2),
            "rolls_est_room": round(float(grp["lin_m"].sum()) / 9.2, 1),
            "rolls_iso": rm_rolls_iso,
            "locations": " ＋ ".join(grp["loc"].tolist()),
            "notes": " ｜ ".join(grp["note"].tolist()),
            "rec_slot": rec_date_map.get(rm, "待排"),
            "is_special": rm in ("2508", "3511", "3302", "2806", "2606"),
        })

    df_rooms = pd.DataFrame(rm_rows)
    df_rooms = df_rooms.sort_values(
        by=["sev_rank", "wall_count", "tv_rank", "area_m2"],
        ascending=[True, False, True, False],
    ).reset_index(drop=True)
    df_rooms.insert(0, "priority", range(1, len(df_rooms) + 1))
    return df_walls, df_rooms


def split_tv_rooms(rooms, df_rooms):
    """將一組房號拆分為『📺 需拆電視房號』與『⬜ 免拆電視房號』字串"""
    rm_map = df_rooms.set_index("room").to_dict("index")
    tv_list, no_tv_list = [], []
    for r in rooms:
        info = rm_map.get(r, {})
        if info.get("has_tv", False):
            if r == "3106":
                tv_list.append("3106(可視色差免拆)")
            elif r == "2605":
                tv_list.append("2605(吊管)")
            else:
                tv_list.append(r)
        else:
            no_tv_list.append(r)
    return (", ".join(tv_list) if tv_list else "無"), (", ".join(no_tv_list) if no_tv_list else "無")


@st.cache_data(ttl=3600)
def get_photo_thumb_uri(filename, max_px=180):
    """將指定照片轉為輕量 Base64 Data URI 縮圖，供 Streamlit ImageColumn 直接嵌入表格"""
    p = os.path.join(PHOTO_DIR, filename)
    if not os.path.exists(p):
        return ""
    try:
        im = Image.open(p).convert("RGB")
        im.thumbnail((max_px, max_px), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=72, optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        return f"data:image/jpeg;base64,{b64}"
    except Exception:
        return ""


def _build_subset_excel_bytes(rooms, df_walls, df_rooms):
    """產生單一批次之 Excel 估價與拆裝明細表（放進 ZIP 資料夾內一併給廠商）"""
    sub_w = df_walls[df_walls["room"].isin(rooms)].copy()
    sub_r = df_rooms[df_rooms["room"].isin(rooms)].copy()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        s1 = sub_r[["priority", "sev_label", "room", "floor", "rt_label", "wall_count", "tv_status", "area_m2", "area_ping", "strips", "lin_m", "notes"]].copy()
        s1.columns = ["優先順位", "嚴重度", "房號", "樓層", "PMS房型", "換牆數", "電視拆裝", "面積(m²)", "坪數", "條數", "下料總長(m)", "拆裝與破損備註"]
        s1.to_excel(writer, sheet_name="本批客房總表", index=False)

        s2 = sub_w[["floor", "room", "rt", "file", "raw", "h", "w", "area_m2", "area_ping", "strips", "lin_m", "loc", "note"]].copy()
        s2.columns = ["樓層", "房號", "房型", "照片檔名", "照片原始標註", "校正高(cm)", "校正寬(cm)", "淨面積(m²)", "坪數", "條數", "下料總長(m)", "牆面區域", "施工備註"]
        s2.to_excel(writer, sheet_name="本批逐牆尺寸明細", index=False)
    return buf.getvalue()


def build_batch_zip_bytes(rooms, df_walls, df_rooms, folder_name):
    """將指定房號之全部現場照片＋Excel 明細打包為可直接寄給廠商的 ZIP 資料夾"""
    sub_w = df_walls[df_walls["room"].isin(rooms)].sort_values(by=["room", "file"])
    rm_map = df_rooms.set_index("room").to_dict("index")
    buf = io.BytesIO()
    safe_folder = folder_name.replace("/", "-").replace("\\", "-").replace(":", "").replace(" ", "_")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, row in sub_w.iterrows():
            fname = row["file"]
            src = os.path.join(PHOTO_DIR, fname)
            if os.path.exists(src):
                rm = row["room"]
                info = rm_map.get(rm, {})
                tv_tag = "拆電視" if info.get("has_tv") else "免拆電視"
                arcname = f"{safe_folder}/{rm}_{row['rt']}_{tv_tag}_{fname}"
                zf.write(src, arcname=arcname)
        xlsx_bytes = _build_subset_excel_bytes(rooms, df_walls, df_rooms)
        zf.writestr(f"{safe_folder}/本批報價與拆裝明細.xlsx", xlsx_bytes)
    return buf.getvalue()


def build_all_batches_zip_bytes(plan_batches, df_walls, df_rooms):
    """將整個藍圖的所有批次一次打包成『已按批次分類好子資料夾』的總包 ZIP"""
    rm_map = df_rooms.set_index("room").to_dict("index")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, b in enumerate(plan_batches, 1):
            d1_s = b["d1"][5:].replace("-", "")
            d2_s = b["d2"][5:].replace("-", "")
            b_clean = b["batch"].replace("/", "-").replace(" ", "").replace("(", "_").replace(")", "")
            folder = f"{idx:02d}_{b_clean}_{d1_s}-{d2_s}"
            sub_w = df_walls[df_walls["room"].isin(b["rooms"])].sort_values(by=["room", "file"])
            for _, row in sub_w.iterrows():
                fname = row["file"]
                src = os.path.join(PHOTO_DIR, fname)
                if os.path.exists(src):
                    rm = row["room"]
                    info = rm_map.get(rm, {})
                    tv_tag = "拆電視" if info.get("has_tv") else "免拆電視"
                    zf.write(src, arcname=f"{folder}/{rm}_{row['rt']}_{tv_tag}_{fname}")
            xlsx_bytes = _build_subset_excel_bytes(b["rooms"], df_walls, df_rooms)
            zf.writestr(f"{folder}/本批報價與拆裝明細.xlsx", xlsx_bytes)
    return buf.getvalue()


def build_excel_bytes(df_rooms, df_walls, inv_data, params):
    """產出 4 分頁專業 Excel 報價與派工工作簿"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Sheet 1: 44 房優先順序總表
        s1 = df_rooms[[
            "priority", "sev_label", "room", "floor", "rt_label", "wall_count",
            "tv_status", "area_m2", "area_ping", "strips", "lin_m", "rec_slot", "notes"
        ]].copy()
        s1.columns = [
            "優先順位", "嚴重度", "房號", "樓層", "PMS房型(總數)", "換牆面數",
            "電視拆裝", "面積(m²)", "面積(坪)", "條數(53cm)", "下料總長(m)", "建議安全壓房日", "現場破損與拆裝備註"
        ]
        s1.to_excel(writer, sheet_name="44房優先順序總表", index=False)

        # Sheet 2: 55 面牆發包明細
        s2 = df_walls[[
            "floor", "room", "rt", "sev", "file", "raw", "h", "w",
            "area_m2", "area_ping", "strips", "lin_m", "loc", "note"
        ]].copy()
        s2.columns = [
            "樓層", "房號", "房型", "嚴重度", "照片檔名", "照片原始標註", "校正高(cm)", "校正寬(cm)",
            "淨面積(m²)", "淨面積(坪)", "條數", "下料總長(m)", "牆面區域", "施工備註"
        ]
        s2.to_excel(writer, sheet_name="55面牆發包明細", index=False)

        # Sheet 3: 6 批次零超賣排程與預算
        sched_rows = []
        for b in VERIFIED_PLAN_3MO:
            m = compute_selection_metrics(
                b["rooms"], df_walls,
                roll_mode=params["roll_mode"],
                fixed_rolls_per_wall=params["fixed_per_wall"],
                roll_price=params["roll_price"],
                worker_fee=params["worker_fee"],
                cap_1=params["cap_1"],
                cap_2=params["cap_2"],
                spare_rolls=0,
                misc_fee=0,
                include_tax=params["include_tax"],
            )
            r2rt = inv_data.get("room_to_rt", DEFAULT_ROOM_TO_RT)
            tv_str, no_tv_str = split_tv_rooms(b["rooms"], df_rooms)
            rm_str = ", ".join(f"{r}({r2rt.get(r,'?')})" for r in b["rooms"])
            sched_rows.append({
                "批次": b["batch"],
                "月份": b["month"],
                "卸家具日(Day1)": b["d1"],
                "施工日(Day2)": b["d2"],
                "房數": m["n_rooms"],
                "牆面數": m["n_walls"],
                "總坪數": round(m["area_ping"], 1),
                "總條數": m["strips"],
                "需拆電視": f"{m['tv_rooms']} 台",
                "📺 需拆電視房號": tv_str,
                "⬜ 免拆電視房號": no_tv_str,
                "實算支數": m["rolls_used"],
                "材料費(NT$)": m["mat_cost"],
                "雙師傅工資(NT$)": m["labor_2m"],
                "批次小計(NT$)": m["grand_total"],
                "全部房號(房型)": rm_str,
                "庫存檢核備註": b["note"],
            })
        pd.DataFrame(sched_rows).to_excel(writer, sheet_name="6批次零超賣排程與預算", index=False)

        # Sheet 4: PMS 房型庫存快照
        daily = inv_data.get("daily_avail", {})
        if daily:
            rts = ["DF", "F", "DTP", "LDT", "DT", "E", "ER", "PD", "PT", "SP", "S", "SD"]
            inv_rows = []
            for ds in sorted(daily.keys()):
                row_d = {"日期": ds}
                tot_rem = 0
                for rt in rts:
                    rem = daily[ds].get(rt, {}).get("rem", 0)
                    row_d[rt] = rem
                    tot_rem += rem
                row_d["全館剩餘總數"] = tot_rem
                inv_rows.append(row_d)
            pd.DataFrame(inv_rows).to_excel(writer, sheet_name="PMS房型剩餘庫存", index=False)

    return buf.getvalue()


# ==============================================================================
# 3. 主畫面渲染函式 (UI/UX 互動看板)
# ==============================================================================

def render_wallpaper_planner():
    """渲染 Hotel Master 獨立新分頁：🧱 專案：客房壁紙修繕規劃"""
    st.markdown("""
    <style>
    .wp-hero {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border-radius: 14px;
        padding: 18px 22px;
        color: #f8fafc;
        margin-bottom: 14px;
        border-left: 6px solid #f59e0b;
    }
    .wp-hero h3 { margin: 0 0 6px 0; color: #ffffff; font-size: 1.35rem; }
    .wp-hero p { margin: 0; color: #cbd5e1; font-size: 0.9rem; }
    .wp-kpi {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 14px 16px;
        box-shadow: 0 2px 6px rgba(15,23,42,0.04);
    }
    .wp-kpi-label { font-size: 0.82rem; color: #64748b; font-weight: 600; margin-bottom: 4px; }
    .wp-kpi-val { font-size: 1.55rem; font-weight: 800; color: #0f172a; line-height: 1.2; }
    .wp-kpi-sub { font-size: 0.78rem; color: #475569; margin-top: 4px; }
    .wp-badge-s { background:#fee2e2; color:#991b1b; padding:2px 8px; border-radius:999px; font-weight:700; font-size:0.78rem; }
    .wp-badge-a { background:#ffedd5; color:#9a3412; padding:2px 8px; border-radius:999px; font-weight:700; font-size:0.78rem; }
    .wp-badge-b { background:#fef9c3; color:#854d0e; padding:2px 8px; border-radius:999px; font-weight:700; font-size:0.78rem; }
    .wp-badge-c { background:#dcfce7; color:#166534; padding:2px 8px; border-radius:999px; font-weight:700; font-size:0.78rem; }
    </style>
    """, unsafe_allow_html=True)

    inv_data = load_roomtype_inventory()
    df_walls, df_rooms = build_dataframes(inv_data)

    st.markdown(f"""
    <div class="wp-hero">
        <h3>🧱 站前館客房壁紙修繕評估、報價模擬與房型壓房排程</h3>
        <p>壁紙規格：<b>寶齊來（明欣壁紙）耐燃三級 95220（批號 24 / 53｜寬 53cm × 長 10m，素面細布紋免對花）</b> ｜ 
        實測範圍：<b>44 間房（55 面有效牆，已剔除重複照 3109-3）</b> ｜ 金旭 PMS 房型庫存基準：<b>{inv_data.get('as_of', '2026-11-20')}</b></p>
    </div>
    """, unsafe_allow_html=True)

    # ==========================================================================
    # 全域報價與算料參數控制台（所有分頁即時連動！）
    # ==========================================================================
    with st.expander("⚙️ 廠商報價與算料參數設定（即時連動全頁試算，點擊收合／展開）", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            roll_price = st.number_input("壁紙單價 (NT$/支)", min_value=100, max_value=2000, value=441, step=10, key="wp_roll_price")
            worker_fee = st.number_input("師傅單人出班費 (NT$/人·天)", min_value=2000, max_value=10000, value=4500, step=500, key="wp_worker_fee")
        with c2:
            cap_1 = st.number_input("單師傅日產能 (間/天)", min_value=1.0, max_value=6.0, value=3.0, step=0.5, key="wp_cap1")
            cap_2 = st.slider("雙師傅日產能 (間/天)", min_value=5.0, max_value=8.5, value=8.0, step=0.5,
                              help="廠商報 7~8 間/天（預設以 8.0 間/天計，6~8 間批次皆為 1 天雙師傅 NT$9,000）；若擔心飯店管制時段打折，可往左調至 6.0 或 7.0 壓測預算", key="wp_cap2")
        with c3:
            mode_label = st.selectbox(
                "壁紙支數計算基準",
                [
                    "📐 實測直條跨房排料 FFD（最準，全案 64 支）",
                    "📏 同樓層共料排料（實務基準，全案 70 支）",
                    "📊 面積＋15% 損耗（粗估參考，全案 59 支）",
                    "⚠️ 固定每面牆支數（如廠商粗估 2.5 支/面）",
                ],
                index=0,
                key="wp_mode_sel",
            )
            mode_map = {
                "📐 實測直條跨房排料 FFD（最準，全案 64 支）": "ffd_cross",
                "📏 同樓層共料排料（實務基準，全案 70 支）": "ffd_floor",
                "📊 面積＋15% 損耗（粗估參考，全案 59 支）": "area_15",
                "⚠️ 固定每面牆支數（如廠商粗估 2.5 支/面）": "fixed_wall",
            }
            roll_mode = mode_map[mode_label]
            fixed_per_wall = st.number_input("固定每面牆支數（僅選固定模式生效）", min_value=1.0, max_value=4.0, value=2.5, step=0.1, key="wp_fixed_pw")
        with c4:
            monthly_budget = st.number_input("單月預算上限 (NT$/月)", min_value=10000, max_value=150000, value=30000, step=5000, key="wp_monthly_budget")
            sc1, sc2 = st.columns(2)
            spare_rolls = sc1.number_input("備品捲 (支)", min_value=0, max_value=15, value=4, step=1, help="同批號 24/53 預留備用捲", key="wp_spare")
            misc_fee = sc2.number_input("耗材雜費 ($)", min_value=0, max_value=20000, value=0, step=500, help="膠水、批土、除霉底塗等（若點工已含填 0）", key="wp_misc")
            include_tax = st.checkbox("計入 5% 營業稅", value=False, key="wp_tax")

    params = {
        "roll_mode": roll_mode, "fixed_per_wall": fixed_per_wall,
        "roll_price": roll_price, "worker_fee": worker_fee,
        "cap_1": cap_1, "cap_2": cap_2, "monthly_budget": monthly_budget,
        "spare_rolls": spare_rolls, "misc_fee": misc_fee, "include_tax": include_tax,
    }

    # 全案 44 間 KPI 總覽
    m_all = compute_selection_metrics(
        df_rooms["room"].tolist(), df_walls,
        roll_mode=roll_mode, fixed_rolls_per_wall=fixed_per_wall,
        roll_price=roll_price, worker_fee=worker_fee, cap_1=cap_1, cap_2=cap_2,
        spare_rolls=spare_rolls, misc_fee=misc_fee, include_tax=include_tax,
    )

    k1, k2, k3, k4 = st.columns(4)
    tax_txt = "含 5% 稅" if include_tax else "未稅"
    k1.markdown(f"""
    <div class="wp-kpi" style="border-top:4px solid #3b82f6;">
        <div class="wp-kpi-label">待修規模（全館 24F–35F）</div>
        <div class="wp-kpi-val">{m_all['n_rooms']} 間 / {m_all['n_walls']} 面牆</div>
        <div class="wp-kpi-sub">淨面積 {m_all['area_m2']:.1f} m²（{m_all['area_ping']:.1f} 坪）｜ {m_all['strips']} 條</div>
    </div>""", unsafe_allow_html=True)

    k2.markdown(f"""
    <div class="wp-kpi" style="border-top:4px solid #f59e0b;">
        <div class="wp-kpi-label">預估壁紙總支數（含備品 {spare_rolls} 支）</div>
        <div class="wp-kpi-val">{m_all['rolls_total']} 支</div>
        <div class="wp-kpi-sub">材料費 NT$ {m_all['mat_cost']:,}（@NT$ {roll_price:,}/支）</div>
    </div>""", unsafe_allow_html=True)

    k3.markdown(f"""
    <div class="wp-kpi" style="border-top:4px solid #10b981;">
        <div class="wp-kpi-label">預估全案總花費（{tax_txt}）</div>
        <div class="wp-kpi-val">NT$ {m_all['grand_total']:,}</div>
        <div class="wp-kpi-sub">{m_all['labor_plan_str']} 工資 NT$ {m_all['labor_2m']:,}（省單人 NT$ {m_all['labor_1m'] - m_all['labor_2m']:,}）</div>
    </div>""", unsafe_allow_html=True)

    k4.markdown(f"""
    <div class="wp-kpi" style="border-top:4px solid #8b5cf6;">
        <div class="wp-kpi-label">平均單房 / 每坪造價</div>
        <div class="wp-kpi-val">NT$ {int(m_all['cost_per_room']):,} / 間</div>
        <div class="wp-kpi-sub">約 NT$ {int(m_all['cost_per_ping']):,} / 坪 ｜ 依月預算需 {math.ceil(m_all['grand_total']/monthly_budget)} 個月</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ==========================================================================
    # 五大互動分頁
    # ==========================================================================
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🎯 優先順序篩選與自訂試算",
        "💰 報價對比與月預算分期",
        "📅 PMS 房型庫存 × 壓房模擬器",
        "🏢 樓層熱區與現場防坑",
        "📥 Excel 匯出與 PMS 同步",
    ])

    # --------------------------------------------------------------------------
    # TAB 1: 多維度篩選、排序與自訂勾選即時試算
    # --------------------------------------------------------------------------
    with tab1:
        st.markdown("#### 🔍 多維度分類篩選與優先順序切換")
        st.caption("自由切換嚴重度、換牆數、電視拆裝、樓層與 PMS 房型；也可直接勾選下方表格房號，即時試算該批組合的壁紙支數與預算！")

        # 快捷情境按鈕
        p_cols = st.columns(6)
        if "wp_preset_rooms" not in st.session_state:
            st.session_state["wp_preset_rooms"] = df_rooms["room"].tolist()

        if p_cols[0].button("📋 全部 44 間", use_container_width=True):
            st.session_state["wp_preset_rooms"] = df_rooms["room"].tolist()
        if p_cols[1].button("🔥 S+A 急件 (15間)", use_container_width=True):
            st.session_state["wp_preset_rooms"] = df_rooms[df_rooms["sev"].isin(["S", "A"])]["room"].tolist()
        if p_cols[2].button("🗓️ 12月急件 (13間)", use_container_width=True):
            st.session_state["wp_preset_rooms"] = VERIFIED_PLAN_3MO[0]["rooms"] + VERIFIED_PLAN_3MO[1]["rooms"]
        if p_cols[3].button("⚡ 12月中前 (20間)", use_container_width=True):
            st.session_state["wp_preset_rooms"] = VERIFIED_PLAN_DEC20[0]["rooms"] + VERIFIED_PLAN_DEC20[1]["rooms"] + VERIFIED_PLAN_DEC20[2]["rooms"]
        if p_cols[4].button("📺 需拆電視 (19間)", use_container_width=True):
            st.session_state["wp_preset_rooms"] = df_rooms[df_rooms["has_tv"]]["room"].tolist()
        if p_cols[5].button("✨ 免拆電視 (25間)", use_container_width=True):
            st.session_state["wp_preset_rooms"] = df_rooms[~df_rooms["has_tv"]]["room"].tolist()

        # 篩選列
        fc1, fc2, fc3, fc4, fc5 = st.columns([2.2, 1.5, 1.5, 1.4, 1.8])
        sev_opts = ["S", "A", "B", "C"]
        sel_sevs = fc1.multiselect("嚴重程度", sev_opts, default=sev_opts, format_func=lambda x: SEV_LABEL[x])
        wall_opts = [3, 2, 1]
        sel_wall_cnt = fc2.multiselect("換牆面數", wall_opts, default=wall_opts, format_func=lambda x: f"{x} 面牆")
        tv_filter = fc3.selectbox("拆電視條件", ["全部", "僅需拆電視 (含3106/2605)", "僅必拆電視 (18間)", "免拆電視 (25間)"])
        all_rts = sorted(df_rooms["rt"].unique().tolist())
        sel_rts = fc4.multiselect("PMS 房型", all_rts, default=all_rts)
        sort_by = fc5.selectbox("排序依據", [
            "綜合優先順序 (#1→#44)",
            "同樓層動線 (35F→24F)",
            "換牆面數 (多→少)",
            "施工面積 (大→小)",
            "PMS 房型分組 (稀缺優先)",
        ])

        all_floors = sorted(df_rooms["floor"].unique().tolist(), reverse=True)
        sel_floors = st.multiselect("樓層篩選（預設全選 35F→24F）", all_floors, default=all_floors, format_func=lambda x: f"{x}F")

        # 套用篩選
        df_view = df_rooms[
            df_rooms["sev"].isin(sel_sevs)
            & df_rooms["wall_count"].isin(sel_wall_cnt)
            & df_rooms["rt"].isin(sel_rts)
            & df_rooms["floor"].isin(sel_floors)
        ].copy()

        if tv_filter == "僅需拆電視 (含3106/2605)":
            df_view = df_view[df_view["has_tv"]]
        elif tv_filter == "僅必拆電視 (18間)":
            df_view = df_view[df_view["tv_status"].isin(["✅ 必拆電視", "✅ 吊管電視"])]
        elif tv_filter == "免拆電視 (25間)":
            df_view = df_view[~df_view["has_tv"]]

        # 套用排序
        if sort_by == "同樓層動線 (35F→24F)":
            df_view = df_view.sort_values(by=["floor", "sev_rank", "room"], ascending=[False, True, True])
        elif sort_by == "換牆面數 (多→少)":
            df_view = df_view.sort_values(by=["wall_count", "sev_rank", "area_m2"], ascending=[False, True, False])
        elif sort_by == "施工面積 (大→小)":
            df_view = df_view.sort_values(by=["area_m2", "sev_rank"], ascending=[False, True])
        elif sort_by == "PMS 房型分組 (稀缺優先)":
            df_view = df_view.sort_values(by=["rt_tot", "rt", "sev_rank"], ascending=[True, True, True])
        else:
            df_view = df_view.sort_values(by=["priority"], ascending=[True])

        # 加入互動勾選欄
        preset_set = set(st.session_state["wp_preset_rooms"])
        df_edit_in = df_view[[
            "priority", "sev_label", "room", "floor", "rt_label",
            "wall_count", "tv_status", "area_ping", "strips", "rec_slot", "notes"
        ]].copy()
        df_edit_in.insert(0, "勾選試算", df_edit_in["room"].isin(preset_set))
        df_edit_in.columns = [
            "勾選試算", "順位", "嚴重度", "房號", "樓層", "PMS房型",
            "牆數", "電視拆裝", "面積(坪)", "條數", "建議壓房日", "破損與拆裝備註"
        ]

        edited_df = st.data_editor(
            df_edit_in,
            use_container_width=True,
            hide_index=True,
            height=460,
            column_config={
                "勾選試算": st.column_config.CheckboxColumn("勾選", width="small"),
                "順位": st.column_config.NumberColumn("#", format="#%d", width="small"),
                "樓層": st.column_config.NumberColumn("樓層", format="%dF", width="small"),
                "牆數": st.column_config.NumberColumn("牆數", format="%d面", width="small"),
                "面積(坪)": st.column_config.NumberColumn("坪數", format="%.2f", width="small"),
                "條數": st.column_config.NumberColumn("條數", format="%d", width="small"),
            },
            disabled=["順位", "嚴重度", "房號", "樓層", "PMS房型", "牆數", "電視拆裝", "面積(坪)", "條數", "建議壓房日", "破損與拆裝備註"],
            key="wp_room_editor",
        )

        # 即時計算「已勾選房號」的專屬報價與工量卡
        checked_rooms = edited_df[edited_df["勾選試算"]]["房號"].tolist()
        m_sel = compute_selection_metrics(
            checked_rooms, df_walls,
            roll_mode=roll_mode, fixed_rolls_per_wall=fixed_per_wall,
            roll_price=roll_price, worker_fee=worker_fee, cap_1=cap_1, cap_2=cap_2,
            spare_rolls=0, misc_fee=0, include_tax=include_tax,
        )

        st.markdown("##### 🧮 目前勾選組合：即時報價與工程負荷試算（不含額外備品與雜費）")
        sc1, sc2, sc3, sc4, sc5 = st.columns(5)
        sc1.metric("已勾選房數 / 牆數", f"{m_sel['n_rooms']} 間 / {m_sel['n_walls']} 牆", f"{m_sel['area_ping']:.1f} 坪｜{m_sel['strips']} 條")
        sc2.metric("實算壁紙支數", f"{m_sel['rolls_used']} 支", f"材料費 NT$ {m_sel['mat_cost']:,}")
        sc3.metric("最優派工與工資", f"NT$ {m_sel['labor_2m']:,}", m_sel["labor_plan_str"])
        budget_delta = monthly_budget - m_sel["grand_total"]
        sc4.metric(
            f"此批小計 ({tax_txt})",
            f"NT$ {m_sel['grand_total']:,}",
            f"預算內餘 NT$ {budget_delta:,}" if budget_delta >= 0 else f"超月預算 NT$ {abs(budget_delta):,}",
            delta_color="normal" if budget_delta >= 0 else "inverse",
        )
        sc5.metric("需拆電視 / 多牆房", f"{m_sel['tv_rooms']} 台電視", f"多牆房 {m_sel['multi_wall_rooms']} 間")

        if m_sel["n_rooms"] > 0:
            sel_floors_list = sorted(df_rooms[df_rooms["room"].isin(checked_rooms)]["floor"].unique().tolist(), reverse=True)
            specials = []
            if "2508" in checked_rooms:
                specials.append("🔴 **2508 發霉牆**：須提前一天撕紙除霉防霉底塗")
            if "3511" in checked_rooms:
                specials.append("🔴 **3511 牆底水漬起泡**：須提前檢測滲水乾燥")
            if "3106" in checked_rooms:
                specials.append("⚠️ **3106 破損在最左側**：若新舊無色差僅換左側 1~2 條可免拆電視")
            st.info(f"🏢 涵蓋樓層（共 {len(sel_floors_list)} 層）：`{', '.join(f'{f}F' for f in sel_floors_list)}` ｜ 跨房 FFD `{m_sel['rolls_ffd']} 支` vs 同樓層 `{m_sel['rolls_floor']} 支` vs 固定 {fixed_per_wall} 支/面 `{m_sel['rolls_fixed']} 支`")
            if specials:
                st.warning(" ｜ ".join(specials))
            custom_zip = build_batch_zip_bytes(checked_rooms, df_walls, df_rooms, f"自訂選取_{len(checked_rooms)}間")
            st.download_button(
                label=f"📦 打包下載【目前勾選 {len(checked_rooms)} 間房（{m_sel['n_walls']} 張照片＋明細表）】給廠商 (.zip)",
                data=custom_zip,
                file_name=f"站前館壁紙_自訂{len(checked_rooms)}間照片資料夾.zip",
                mime="application/zip",
                use_container_width=True,
                key="wp_dl_custom_zip",
            )

    # --------------------------------------------------------------------------
    # TAB 2: 報價試算對比與月預算分期器
    # --------------------------------------------------------------------------
    with tab2:
        st.markdown("#### 💰 廠商報價拆解：工資最優化 × 算料防坑對比")
        ch1, ch2 = st.columns(2)

        with ch1:
            # 工資比較圖
            df_lab_cmp = pd.DataFrame([
                {"模式": f"全程雙師傅 ({m_all['labor_plan_str']})", "工資": m_all["labor_2m"], "單房工資": m_all["labor_2m"] / 44},
                {"模式": f"雙師傅 + 預留1人收尾日", "工資": m_all["labor_2m"] + worker_fee, "單房工資": (m_all["labor_2m"] + worker_fee) / 44},
                {"模式": f"全程單師傅 ({m_all['days_1m']} 天)", "工資": m_all["labor_1m"], "單房工資": m_all["labor_1m"] / 44},
            ])
            fig_lab = px.bar(
                df_lab_cmp, x="模式", y="工資", text="工資",
                color="模式", color_discrete_sequence=["#10b981", "#3b82f6", "#ef4444"],
                title="全案 44 間工資方案對比 (NT$)",
            )
            fig_lab.update_traces(texttemplate="NT$ %{y:,.0f}", textposition="outside")
            fig_lab.update_layout(showlegend=False, height=340, yaxis_title="總工資 (NT$)", xaxis_title="")
            st.plotly_chart(fig_lab, use_container_width=True)

        with ch2:
            # 算料模式材料費防坑對比
            df_mat_cmp = pd.DataFrame([
                {"算料方式": f"面積+15% ({m_all['rolls_area15']}支⚠️會缺料)", "支數": m_all["rolls_area15"], "材料費": m_all["rolls_area15"] * roll_price},
                {"算料方式": f"實測跨房FFD ({m_all['rolls_ffd']}支⭐最準)", "支數": m_all["rolls_ffd"], "材料費": m_all["rolls_ffd"] * roll_price},
                {"算料方式": f"同樓層共料 ({m_all['rolls_floor']}支⭐實務)", "支數": m_all["rolls_floor"], "材料費": m_all["rolls_floor"] * roll_price},
                {"算料方式": f"每面牆{fixed_per_wall}支 ({m_all['rolls_fixed']}支🚨高估)", "支數": m_all["rolls_fixed"], "材料費": m_all["rolls_fixed"] * roll_price},
            ])
            fig_mat = px.bar(
                df_mat_cmp, x="算料方式", y="材料費", text="材料費",
                color="算料方式", color_discrete_sequence=["#94a3b8", "#10b981", "#0ea5e9", "#f43f5e"],
                title=f"全案 55 面牆材料費防坑對比 (@NT$ {roll_price}/支)",
            )
            fig_mat.update_traces(texttemplate="NT$ %{y:,.0f}", textposition="outside")
            fig_mat.update_layout(showlegend=False, height=340, yaxis_title="壁紙材料費 (NT$)", xaxis_title="")
            st.plotly_chart(fig_mat, use_container_width=True)

        overbuy_diff = (m_all["rolls_fixed"] - m_all["rolls_floor"]) * roll_price
        if overbuy_diff > 0:
            st.error(f"🚨 **防坑提醒**：若廠商按「每面牆固定 {fixed_per_wall} 支（共 {m_all['rolls_fixed']} 支）」報價，比實測同樓層共料（{m_all['rolls_floor']} 支）**多買 {m_all['rolls_fixed'] - m_all['rolls_floor']} 支＝多花 NT$ {overbuy_diff:,}**！務必要求按實際領用支數實報實銷。")

        st.markdown("---")
        st.markdown(f"#### 📅 依「每月預算上限 NT$ {monthly_budget:,}」之 3 個月分期試算表（零超賣 6 批次）")
        plan_choice = st.radio(
            "切換預覽排程藍圖：",
            ["⭐ 藍圖 A：每月 3 萬 × 三個月穩健完修 (12月13間 / 1月16間 / 2月15間)", "⚡ 藍圖 B：12 月中旬前先衝刺 20 間急件 (3天雙師傅)"],
            horizontal=True,
        )
        active_plan = VERIFIED_PLAN_3MO if "藍圖 A" in plan_choice else VERIFIED_PLAN_DEC20

        rows_plan = []
        month_totals = defaultdict(int)
        for b in active_plan:
            mb = compute_selection_metrics(
                b["rooms"], df_walls,
                roll_mode=roll_mode, fixed_rolls_per_wall=fixed_per_wall,
                roll_price=roll_price, worker_fee=worker_fee, cap_1=cap_1, cap_2=cap_2,
                spare_rolls=0, misc_fee=0, include_tax=include_tax,
            )
            month_totals[b["month"]] += mb["grand_total"]
            tv_str, no_tv_str = split_tv_rooms(b["rooms"], df_rooms)
            rows_plan.append({
                "月份": b["month"],
                "批次": b["batch"],
                "卸家具 → 施工日": f"{b['d1'][5:]} → {b['d2'][5:]}",
                "房數": mb["n_rooms"],
                "牆數": mb["n_walls"],
                "坪數": round(mb["area_ping"], 1),
                "條數": mb["strips"],
                "需拆電視": f"{mb['tv_rooms']} 台",
                "📺 需拆電視房號": tv_str,
                "⬜ 免拆電視房號": no_tv_str,
                "實算支數": mb["rolls_used"],
                "材料費": f"NT$ {mb['mat_cost']:,}",
                "點工工資": f"NT$ {mb['labor_2m']:,}",
                f"批次小計 ({tax_txt})": f"NT$ {mb['grand_total']:,}",
            })
        st.dataframe(pd.DataFrame(rows_plan), use_container_width=True, hide_index=True)

        # 月度預算達成率卡片
        m_cols = st.columns(len(month_totals))
        for idx, (ym, m_amt) in enumerate(sorted(month_totals.items())):
            diff = monthly_budget - m_amt
            status_lbl = f"✅ 預算內（餘 NT$ {diff:,}）" if diff >= 0 else f"⚠️ 超出 NT$ {abs(diff):,}"
            m_cols[idx].metric(f"📆 {ym} 月度預估合計 ({tax_txt})", f"NT$ {m_amt:,}", status_lbl,
                               delta_color="normal" if diff >= 0 else "inverse")

        # 各批次「現場照片縮圖明細表」＋「一鍵打包分批資料夾給廠商 (.zip)」
        st.markdown("---")
        hd_c1, hd_c2 = st.columns([2.2, 1.8])
        with hd_c1:
            st.markdown("#### 🖼️ 各批次現場照片彙整與「廠商資料夾打包區」")
            st.caption("直接在此檢視每批對應房號的現場照片（已自動剔除重複照 3109-3），並可一鍵下載「已分好資料夾的 ZIP」直接轉傳廠商估價！")
        with hd_c2:
            all_zip_bytes = build_all_batches_zip_bytes(active_plan, df_walls, df_rooms)
            plan_tag = "3個月6批次" if "藍圖 A" in plan_choice else "12月中前20間"
            st.download_button(
                label=f"📦 一鍵打包下載【全部 {len(active_plan)} 批已分類資料夾】(.zip)",
                data=all_zip_bytes,
                file_name=f"站前館壁紙_{plan_tag}_全部分批資料夾.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key=f"wp_dl_all_batches_{plan_tag}",
            )

        batch_labels = [f"{b['batch']}（{b['d1'][5:]}卸 → {b['d2'][5:]}貼｜{len(b['rooms'])}間）" for b in active_plan]
        bc1, bc2 = st.columns([2.4, 1.6])
        sel_b_lbl = bc1.selectbox("選擇要檢視照片與下載的批次：", batch_labels, index=0, key=f"wp_photo_batch_sel_{plan_tag}")
        sel_b = active_plan[batch_labels.index(sel_b_lbl)]
        sel_b_walls = df_walls[df_walls["room"].isin(sel_b["rooms"])].sort_values(by=["room", "file"]).copy()
        sel_b_tv, sel_b_notv = split_tv_rooms(sel_b["rooms"], df_rooms)

        with bc2:
            st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
            one_zip = build_batch_zip_bytes(sel_b["rooms"], df_walls, df_rooms, f"{sel_b['batch']}_{sel_b['d1'][5:]}_{sel_b['d2'][5:]}")
            st.download_button(
                label=f"📁 下載本批資料夾 ({len(sel_b_walls)} 張照片＋Excel)",
                data=one_zip,
                file_name=f"站前館壁紙_{sel_b['d1'][5:].replace('-','')}_{len(sel_b['rooms'])}間.zip",
                mime="application/zip",
                use_container_width=True,
                key=f"wp_dl_one_batch_{sel_b['d1']}",
            )

        st.info(f"**🔹 {sel_b['batch']}（`{sel_b['d1']}` 卸家具 → `{sel_b['d2']}` 雙師傅施工）**　｜　📺 需拆電視：`{sel_b_tv}`　｜　⬜ 免拆電視：`{sel_b_notv}`")

        # 含嵌入圖片縮圖的逐牆對照表（一牆一圖，完全對應照片檔名！）
        rm_info_map = df_rooms.set_index("room").to_dict("index")
        tbl_rows = []
        for _, wr in sel_b_walls.iterrows():
            rm = wr["room"]
            rinfo = rm_info_map.get(rm, {})
            tbl_rows.append({
                "現場照片": get_photo_thumb_uri(wr["file"]),
                "房號": rm,
                "照片檔名": wr["file"],
                "PMS房型": rinfo.get("rt_label", wr["rt"]),
                "嚴重度": rinfo.get("sev_label", wr["sev"]),
                "電視拆裝": rinfo.get("tv_status", ""),
                "照片標註尺寸": wr["raw"],
                "高×寬(cm)": f"{wr['h']}×{wr['w']}",
                "坪數 / 條數": f"{wr['area_ping']:.2f}坪 / {wr['strips']}條",
                "牆面區域與施工備註": f"{wr['loc']}｜{wr['note']}",
            })

        st.dataframe(
            pd.DataFrame(tbl_rows),
            use_container_width=True,
            hide_index=True,
            column_config={
                "現場照片": st.column_config.ImageColumn("🖼️ 現場縮圖", width="medium"),
                "房號": st.column_config.TextColumn("房號", width="small"),
                "照片檔名": st.column_config.TextColumn("照片檔名", width="small"),
            },
        )

        # 大圖預覽畫廊（點擊展開看清晰大圖與紅框細節）
        with st.expander(f"🔎 展開【{sel_b['batch']}】全部 {len(sel_b_walls)} 張清晰大圖檢視（4 欄畫廊）", expanded=False):
            g_cols = st.columns(4)
            for idx_w, (_, wr) in enumerate(sel_b_walls.iterrows()):
                col_box = g_cols[idx_w % 4]
                p_img = os.path.join(PHOTO_DIR, wr["file"])
                rinfo = rm_info_map.get(wr["room"], {})
                cap = f"{wr['room']} ({wr['rt']})｜{wr['file']}｜{wr['raw']}｜{rinfo.get('tv_status','')}"
                if os.path.exists(p_img):
                    col_box.image(p_img, caption=cap, use_container_width=True)
                else:
                    col_box.warning(f"找不到圖片：{wr['file']}")

    # --------------------------------------------------------------------------
    # TAB 3: 金旭 PMS 即時房型庫存熱力圖 × What-If 壓房衝突模擬器
    # --------------------------------------------------------------------------
    with tab3:
        st.markdown("#### 🚦 金旭 PMS 即時房型剩餘庫存熱力圖（剩餘 ＝ 總數 － OOO － 已訂）")
        daily = inv_data.get("daily_avail", {})
        if not daily:
            st.warning("⚠️ 尚無房型庫存快照，請至「📥 Excel 匯出與 PMS 同步」分頁拉取金旭 PMS 即時數據。")
        else:
            range_opt = st.radio(
                "選擇檢視區間：",
                ["🎯 精選低谷候選日（推薦）", "2026 年 12 月全月", "2027 年 1 月全月", "2027 年 2 月全月"],
                horizontal=True,
                key="wp_inv_range",
            )
            key_rts = ["DTP", "F", "PT", "DF", "SP", "S", "DT", "PD", "ER", "E"]
            if "精選" in range_opt:
                target_dates = [
                    "2026-12-01", "2026-12-02", "2026-12-07", "2026-12-08",
                    "2026-12-14", "2026-12-15", "2026-12-16", "2026-12-17",
                    "2026-12-21", "2026-12-22",
                    "2027-01-04", "2027-01-05", "2027-01-06", "2027-01-07",
                    "2027-01-12", "2027-01-13", "2027-01-18", "2027-01-19",
                    "2027-01-20", "2027-01-21", "2027-01-25", "2027-01-26",
                    "2027-02-22", "2027-02-23", "2027-02-24", "2027-02-25",
                ]
                target_dates = [d for d in target_dates if d in daily]
            elif "12 月" in range_opt:
                target_dates = sorted(d for d in daily.keys() if d.startswith("2026-12"))
            elif "1 月" in range_opt:
                target_dates = sorted(d for d in daily.keys() if d.startswith("2027-01"))
            else:
                target_dates = sorted(d for d in daily.keys() if d.startswith("2027-02"))

            dows_zh = ["一", "二", "三", "四", "五", "六", "日"]
            y_labels = [f"{d[5:]} ({dows_zh[datetime.date.fromisoformat(d).weekday()]})" for d in target_dates]
            z_mat = [[daily[d].get(rt, {}).get("rem", 0) for rt in key_rts] for d in target_dates]
            x_labels = [f"{rt} ({DEFAULT_RT_TOTAL.get(rt,0)})" for rt in key_rts]

            fig_hm = go.Figure(data=go.Heatmap(
                z=z_mat, x=x_labels, y=y_labels,
                text=z_mat, texttemplate="%{text}",
                colorscale=[
                    [0.0, "#ef4444"],   # 0 間：紅（已滿房禁壓）
                    [0.07, "#f59e0b"],  # 1 間：黃（緊繃）
                    [0.20, "#86efac"],  # 3~4 間：淺綠
                    [1.0, "#15803d"],   # 10+ 間：深綠
                ],
                zmin=0, zmax=15,
                colorbar=dict(title="剩餘間數"),
            ))
            fig_hm.update_layout(
                title="各日期 × 10 種待修房型「即時剩餘可售間數」（紅 0＝滿房禁壓｜黃 1＝緊繃｜綠 ≥2＝安全）",
                height=max(420, len(target_dates) * 24),
                yaxis=dict(autorange="reversed"),
                margin=dict(l=20, r=20, t=50, b=20),
            )
            st.plotly_chart(fig_hm, use_container_width=True)

            st.markdown("---")
            st.markdown("#### 🛠️ What-If 壓房衝突即時檢核器（自選日期 × 自選房號，馬上驗算會不會超賣！）")

            # 一鍵載入標準批次
            preset_batch_names = ["自訂組合"] + [f"{b['batch']} ({b['d1'][5:]}→{b['d2'][5:]})" for b in VERIFIED_PLAN_3MO]
            sel_preset_b = st.selectbox("快速套用已驗證零超賣批次（或選「自訂組合」手動調配）：", preset_batch_names, index=1, key="wp_sim_preset")

            if sel_preset_b != "自訂組合":
                b_idx = preset_batch_names.index(sel_preset_b) - 1
                default_d1 = VERIFIED_PLAN_3MO[b_idx]["d1"]
                default_rms = VERIFIED_PLAN_3MO[b_idx]["rooms"]
            else:
                default_d1 = "2026-12-07"
                default_rms = ["3111", "3211", "2603", "2903", "3106", "2702", "3302"]

            all_avail_dates = sorted(daily.keys())
            sim_c1, sim_c2 = st.columns([1.2, 2.8])
            with sim_c1:
                d1_idx = all_avail_dates.index(default_d1) if default_d1 in all_avail_dates else 0
                sim_d1 = st.selectbox("選擇「Day 1 卸家具日」", all_avail_dates, index=d1_idx, key="wp_sim_d1")
                sim_d2_dt = datetime.date.fromisoformat(sim_d1) + datetime.timedelta(days=1)
                sim_d3_dt = datetime.date.fromisoformat(sim_d1) + datetime.timedelta(days=2)
                sim_d2 = sim_d2_dt.strftime("%Y-%m-%d")
                st.caption(f"➡️ **Day 2 雙師傅施工日**：`{sim_d2}`")
                st.caption(f"➡️ **Day 3 上午復歸開賣**：`{sim_d3_dt.strftime('%Y-%m-%d')}`（禁 Early CI）")

            with sim_c2:
                all_rm_opts = df_rooms["room"].tolist()
                rm_disp = {r["room"]: f"#{r['priority']} {r['room']} ({r['rt']} | {r['sev']}級 | {r['wall_count']}牆{'|TV' if r['has_tv'] else ''})" for _, r in df_rooms.iterrows()}
                sim_rooms = st.multiselect(
                    "勾選此批次要壓的房號（建議雙師傅每日 6~8 間）",
                    options=all_rm_opts,
                    default=[r for r in default_rms if r in all_rm_opts],
                    format_func=lambda x: rm_disp.get(x, x),
                    key=f"wp_sim_rooms_{sel_preset_b}",
                )

            if sim_rooms and sim_d1 in daily and sim_d2 in daily:
                r2rt = inv_data.get("room_to_rt", DEFAULT_ROOM_TO_RT)
                rt_need = defaultdict(list)
                for rm in sim_rooms:
                    rt_need[r2rt.get(rm, DEFAULT_ROOM_TO_RT.get(rm, "UNK"))].append(rm)

                chk_rows = []
                any_over = False
                any_tight = False
                for rt in sorted(rt_need.keys()):
                    need = len(rt_need[rt])
                    rem1 = daily[sim_d1].get(rt, {}).get("rem", 0)
                    rem2 = daily[sim_d2].get(rt, {}).get("rem", 0)
                    left1, left2 = rem1 - need, rem2 - need
                    if left1 < 0 or left2 < 0:
                        st_txt = "🔴 超賣！不可壓"
                        any_over = True
                    elif left1 == 0 or left2 == 0:
                        st_txt = "🟡 零緩衝（壓完即滿房）"
                        any_tight = True
                    else:
                        st_txt = "🟢 安全可壓"
                    chk_rows.append({
                        "PMS房型": f"{rt} (全館{DEFAULT_RT_TOTAL.get(rt,0)}間)",
                        "擬壓房號": ", ".join(rt_need[rt]),
                        "需壓間數": need,
                        f"Day1 ({sim_d1[5:]}) 剩餘": rem1,
                        "Day1 壓後餘": left1,
                        f"Day2 ({sim_d2[5:]}) 剩餘": rem2,
                        "Day2 壓後餘": left2,
                        "庫存檢核": st_txt,
                    })

                m_sim = compute_selection_metrics(
                    sim_rooms, df_walls,
                    roll_mode=roll_mode, fixed_rolls_per_wall=fixed_per_wall,
                    roll_price=roll_price, worker_fee=worker_fee, cap_1=cap_1, cap_2=cap_2,
                    spare_rolls=0, misc_fee=0, include_tax=include_tax,
                )

                if any_over:
                    st.error("🚨 **檢核失敗：出現房型超賣！** 請將標紅房型的房號改移至綠燈日期，或替換成其他庫存充裕房型。")
                elif any_tight:
                    st.warning("🟡 **檢核通過但有房型壓後歸零（0 緩衝）**：鎖房當下須同步確認 PMS 關閉該房型銷售。")
                else:
                    st.success(f"🟢 **檢核 100% 通過！零房型超賣、所有房型壓後皆留有安全可售緩衝！**（本批 {m_sim['n_rooms']} 間 / {m_sim['strips']} 條 / {m_sim['area_ping']:.1f} 坪 / {m_sim['tv_rooms']} 台TV / 預估花費 NT$ {m_sim['grand_total']:,}）")

                st.dataframe(pd.DataFrame(chk_rows), use_container_width=True, hide_index=True)
                sim_tv_str, sim_notv_str = split_tv_rooms(sim_rooms, df_rooms)
                st.info(f"📺 **本批需拆電視房號（{m_sim['tv_rooms']} 台）**：`{sim_tv_str}`　｜　⬜ **免拆電視房號**：`{sim_notv_str}`")

    # --------------------------------------------------------------------------
    # TAB 4: 同樓層重災熱區與現場施工防坑總覽
    # --------------------------------------------------------------------------
    with tab4:
        st.markdown("#### 🏢 同樓層（35F → 24F）重災嚴重度分布與工程拆裝熱區")
        fl_sev = df_rooms.groupby(["floor", "sev"]).size().reset_index(name="房數")
        fl_sev["樓層"] = fl_sev["floor"].astype(str) + "F"
        fig_fl = px.bar(
            fl_sev, x="樓層", y="房數", color="sev",
            category_orders={"樓層": [f"{f}F" for f in range(35, 23, -1)], "sev": ["S", "A", "B", "C"]},
            color_discrete_map={"S": "#dc2626", "A": "#f97316", "B": "#eab308", "C": "#22c55e"},
            title="各樓層待修房數與嚴重度堆疊分布（紅 S / 橘 A 為優先重災樓層）",
            text="房數",
        )
        fig_fl.update_layout(height=360, xaxis_title="樓層 (高→低)", yaxis_title="待修房數")
        st.plotly_chart(fig_fl, use_container_width=True)

        # 樓層彙總表
        fl_summary = []
        for fl in range(35, 23, -1):
            sub_r = df_rooms[df_rooms["floor"] == fl]
            sub_w = df_walls[df_walls["floor"] == fl]
            if sub_r.empty:
                continue
            fl_strips = []
            for _, rr in sub_w.iterrows():
                fl_strips.extend([int(rr["h"]) + 10] * int(rr["strips"]))
            fl_rolls, _ = calc_ffd_rolls(fl_strips, 1000)
            fl_summary.append({
                "樓層": f"{fl}F",
                "房數": len(sub_r),
                "牆數": len(sub_w),
                "S/A/B/C": f"{(sub_r['sev']=='S').sum()} / {(sub_r['sev']=='A').sum()} / {(sub_r['sev']=='B').sum()} / {(sub_r['sev']=='C').sum()}",
                "多牆房": int((sub_r["wall_count"] >= 2).sum()),
                "需拆電視": int(sub_r["has_tv"].sum()),
                "面積(坪)": round(float(sub_r["area_ping"].sum()), 2),
                "同層排料(支)": fl_rolls,
                "房號明細(房型/等級)": "、".join(f"{r['room']}({r['rt']}/{r['sev']})" for _, r in sub_r.iterrows()),
            })
        st.dataframe(pd.DataFrame(fl_summary), use_container_width=True, hide_index=True)

        st.markdown("#### ⚠️ 開工前現場防坑 6 大特別備註")
        a1, a2 = st.columns(2)
        with a1:
            st.error("**1. `2508-2` 牆面發霉（S 級首修）**：務必於 Day 1 上午先撕除舊紙、殺菌除霉＋塗刷防霉底塗並開除濕乾燥一晚，切忌當天撕當天貼。")
            st.error("**2. `3511` 牆底大面積水漬起泡（S 級）**：請工程部先確認無管線滲水並除濕乾燥，且該房標註高 198cm 偏低，開工前須雷射複尺。")
            st.warning("**3. `3109-3.jpg` 已剔除重複計價**：經比對牆面「1」字刮痕與破皮點，確認與 `3109-2.jpg` 為同一面床尾牆，已剔除以防廠商重複報價。")
        with a2:
            st.info("**4. `3106` 可視色差免拆電視**：破損全在最左側靠大理石洗手台，若新紙 `24/53` 與舊牆無色差，只沿接縫換左側 1~2 條即可免拆電視與書桌。")
            st.info("**5. `3302 (ER)` 現正修冷氣 OOO**：金旭 PMS 顯示該房目前因冷氣維修鎖房中，可順手檢查壁面並確認開工排程。")
            st.success("**6. `2806`（10cm 窗台條）與 `2606`（35cm 窄柱）**：免拆任何家具，可於同樓層施工時用餘料 20 分鐘順手修補，不必浪費 2 天壓房。")

    # --------------------------------------------------------------------------
    # TAB 5: 一鍵匯出 Excel 與金旭 PMS 即時重拉
    # --------------------------------------------------------------------------
    with tab5:
        st.markdown("#### 📥 一鍵匯出報價與派工 Excel（4 分頁完整工作簿）")
        st.caption("工作簿包含：① 44 房優先順序總表（含 PMS 房型與建議壓房日）、② 55 面牆發包明細、③ 6 批次零超賣排程與動態報價、④ 金旭 PMS 房型剩餘庫存表。")

        excel_bytes = build_excel_bytes(df_rooms, df_walls, inv_data, params)
        st.download_button(
            label="📥 下載《站前館客房壁紙修繕評估與派工總表.xlsx》",
            data=excel_bytes,
            file_name=f"站前館壁紙修繕評估與派工_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.markdown("---")
        st.markdown("#### 🔄 金旭 PMS 房型庫存即時同步（區網內電腦專用）")
        st.caption(f"目前載入之 PMS 庫存快照時間：**`{inv_data.get('as_of', '未知')}`**（檔案：`data/pms_backup/wallpaper_roomtype_inventory.json`）。若在飯店區網內使用，可點擊下方按鈕即時重拉最新訂房庫存。")
        if st.button("🔄 立即連線金旭 PMS 重新同步各房型剩餘庫存", use_container_width=True):
            with st.spinner("正在唯讀查詢金旭 PMS (HT2000) 各房型最新訂房與 OOO..."):
                try:
                    new_payload = refresh_pms_roomtype_inventory()
                    st.success(f"✅ 同步完成！已更新至 `{new_payload['as_of']}`（共 {len(new_payload['daily_avail'])} 天分房型即時庫存），請重新檢視「📅 PMS 房型庫存」分頁。")
                    st.rerun()
                except Exception as e:
                    st.warning(f"⚠️ 無法連線至飯店區網金旭 PMS（雲端版屬正常現象，系統將繼續使用最近一次快照 `{inv_data.get('as_of')}`）。詳細原因：{e}")
