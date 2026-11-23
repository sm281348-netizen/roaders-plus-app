# -*- coding: utf-8 -*-
"""唯讀同步金旭 PMS (HT2000) 站前館房型庫存快照至 data/pms_backup/wallpaper_roomtype_inventory.json"""
import os
import sys

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import wallpaper_planner

if __name__ == "__main__":
    payload = wallpaper_planner.refresh_pms_roomtype_inventory()
    tot = sum(v["total_inv"] for v in payload["rt_master"].values())
    print(f"[OK] Refreshed PMS room-type snapshot as_of={payload['as_of']} | total_inv={tot} | days={len(payload['daily_avail'])}")
