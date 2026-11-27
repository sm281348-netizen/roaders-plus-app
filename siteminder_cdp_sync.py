# -*- coding: utf-8 -*-
"""
SiteMinder 100% 唯讀本機 CDP 自動快照同步引擎 (SiteMinder Read-Only CDP Sync)
=============================================================================
安全紅線：
1. 100% 唯讀：僅讀取 Roaders Plus Hotel Taipei Station 分頁畫面文字與數值。
2. 嚴禁點擊 Save / Update / Reset / Publish，嚴禁修改任何房價、庫存、Stop Sell 或 MinLOS。
3. 嚴禁切閱非 SiteMinder 之其他瀏覽器分頁，嚴禁嘗試繞過登入/2FA/人機驗證。
4. 優雅降級：於 05:00 無人排程執行時，若 Chrome 未開啟或 SiteMinder 已登出，僅記錄狀態並保留既有快照，絕不中斷排程。

支援雙連線模式：
- 模式 A (無人值守建議)：獨立 Chrome 設定檔開啟 --remote-debugging-port=9223 (可用 launch_siteminder_chrome.bat 啟動)
- 模式 B (日常隨手同步)：自動偵測本機預設 Chrome 的 DevToolsActivePort (127.0.0.1:9222)
"""

import os
import re
import json
import time
import logging
import datetime
import argparse
import requests

logger = logging.getLogger("pms_sync")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP_DIR, "data", "pms_backup")
SNAPSHOT_PATH = os.path.join(BACKUP_DIR, "siteminder_inventory_snapshot.json")
SYNC_STATUS_PATH = os.path.join(BACKUP_DIR, "siteminder_sync_status.json")
SM_HISTORY_DIR = os.path.join(BACKUP_DIR, "sm_history")


def archive_siteminder_snapshot(payload: dict, max_keep_days: int = 120):
    """將當日有效 SiteMinder 掛牌快照按日期歸檔至 sm_history/YYYY-MM-DD.json（不覆蓋跨日歷史，供未來精準對照下訂當日 BAR）。"""
    if not payload or not payload.get("inventory"):
        return None
    as_of_date = str(payload.get("as_of", ""))[:10]
    if len(as_of_date) != 10:
        as_of_date = datetime.date.today().strftime("%Y-%m-%d")
    os.makedirs(SM_HISTORY_DIR, exist_ok=True)
    arch_path = os.path.join(SM_HISTORY_DIR, f"{as_of_date}.json")
    with open(arch_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    # 自動清理超過 max_keep_days 的舊歸檔避免倉庫膨脹
    try:
        files = sorted(fn for fn in os.listdir(SM_HISTORY_DIR) if fn.endswith(".json"))
        if len(files) > max_keep_days:
            for old_fn in files[:-max_keep_days]:
                os.remove(os.path.join(SM_HISTORY_DIR, old_fn))
    except Exception:
        pass
    return arch_path

PROPERTY_ID = "1bd893db-536e-44df-80de-7395edace58e"
DASHBOARD_URL = f"https://app.siteminder.com/web/cm/properties/{PROPERTY_ID}/dashboard"
INVENTORY_URL = f"https://app.siteminder.com/web/cm/properties/{PROPERTY_ID}/inventory"

ROOM_LABEL_TO_CODE = [
    ("豪華家庭房", "DF"),
    ("豪華三人房", "DTP"),
    ("奢華雙人房", "LDT"),
    ("豪華雙人房", "DT"),
    ("101景觀行政雙人房", "E"),
    ("行政雙人房", "E"),
    ("尊爵雙人房", "PD"),
    ("尊爵雙床房", "PT"),
    ("高級雙人房", "SP"),
    ("標準雙人房", "S"),
    ("家庭房", "F"),
]


class CDPClient:
    """極輕量純 Python WebSocket CDP 唯讀客戶端（依賴既有的 websocket-client）。"""

    def __init__(self, ws_url: str, timeout: int = 12):
        import websocket
        self.ws = websocket.create_connection(ws_url, timeout=timeout, suppress_origin=True)
        self.msg_id = 0

    def send(self, method: str, params: dict = None, session_id: str = None, timeout: int = 12):
        self.msg_id += 1
        req_id = self.msg_id
        payload = {"id": req_id, "method": method, "params": params or {}}
        if session_id:
            payload["sessionId"] = session_id
        self.ws.settimeout(timeout)
        self.ws.send(json.dumps(payload))
        deadline = time.time() + timeout
        while time.time() < deadline:
            raw = self.ws.recv()
            data = json.loads(raw)
            if data.get("id") == req_id:
                if "error" in data:
                    raise RuntimeError(f"CDP {method} error: {data['error']}")
                return data.get("result", {})
        raise TimeoutError(f"CDP {method} timed out after {timeout}s")

    def eval_js(self, expr: str, session_id: str = None, timeout: int = 12):
        res = self.send(
            "Runtime.evaluate",
            {"expression": expr, "returnByValue": True, "awaitPromise": True},
            session_id=session_id,
            timeout=timeout,
        )
        return res.get("result", {}).get("value")

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass


def _discover_siteminder_target(unattended: bool = True):
    """
    尋找已開啟且登入中之 SiteMinder 分頁（優先 port 9223 專用設定檔，其次預設 Chrome DevToolsActivePort）。
    回傳 (cdp_client, session_id, mode_name) 或 (None, None, reason)。
    """
    # 1. 優先嘗試 port 9223 (獨立設定檔，無人值守零確認彈窗)
    try:
        r = requests.get("http://127.0.0.1:9223/json/list", timeout=2)
        if r.status_code == 200:
            tabs = r.json()
            for t in tabs:
                if t.get("type") == "page" and "siteminder.com" in (t.get("url") or "").lower():
                    ws_url = t.get("webSocketDebuggerUrl")
                    if ws_url:
                        client = CDPClient(ws_url, timeout=10)
                        return client, None, "dedicated_9223"
    except Exception:
        pass

    # 2. 其次嘗試預設 Chrome DevToolsActivePort (127.0.0.1:9222)
    port_file = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data\DevToolsActivePort")
    if os.path.exists(port_file):
        try:
            lines = [ln.strip() for ln in open(port_file, "r", encoding="utf-8").read().splitlines() if ln.strip()]
            if len(lines) >= 2:
                port, path = lines[0], lines[1]
                ws_url = f"ws://127.0.0.1:{port}{path}"
                connect_timeout = 6 if unattended else 30
                client = CDPClient(ws_url, timeout=connect_timeout)
                targets = client.send("Target.getTargets", timeout=connect_timeout).get("targetInfos", [])
                sm_target = None
                for t in targets:
                    if t.get("type") == "page" and "siteminder.com" in (t.get("url") or "").lower():
                        sm_target = t
                        break
                if not sm_target:
                    client.close()
                    return None, None, "chrome_open_but_no_siteminder_tab"
                attach = client.send(
                    "Target.attachToTarget",
                    {"targetId": sm_target["targetId"], "flatten": True},
                    timeout=10,
                )
                return client, attach.get("sessionId"), f"default_chrome_{port}"
        except Exception as e:
            return None, None, f"default_chrome_unavailable: {type(e).__name__}"

    return None, None, "no_cdp_chrome_running"


def _js_parse_inventory_window():
    """唯讀解析當前 SiteMinder Inventory 網格（14 天視窗），回傳日期、房價、庫存與 MinLOS。"""
    return r"""
    (() => {
      const text = document.body ? document.body.innerText : '';
      const url = location.href;
      if (/login|signin|auth/i.test(url) || /Sign in to SiteMinder/i.test(text)) {
        return { logged_out: true, url };
      }

      // 擷取頁面所有日期標頭 (例如 "Nov\n23\nMon" 或 "23 Nov" 或 data attributes)
      const monthMap = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
      const headers = [];
      const headerEls = document.querySelectorAll('[class*="calendar"] th, thead th, [role="columnheader"], .sm-inventory-date, [data-testid*="date"]');
      headerEls.forEach(el => {
        const t = (el.innerText || '').replace(/\s+/g, ' ').trim();
        const m = t.match(/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b[^\d]{0,5}(\d{1,2})\b/i)
               || t.match(/\b(\d{1,2})\b[^\w]{0,5}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b/i);
        if (m) {
          let monStr, dayNum;
          if (monthMap[m[1].slice(0,3)]) { monStr = m[1].slice(0,3); dayNum = parseInt(m[2], 10); }
          else { dayNum = parseInt(m[1], 10); monStr = m[2].slice(0,3); }
          const key = monStr + '-' + dayNum;
          if (!headers.some(h => h.key === key)) {
            headers.push({ key, mon: monthMap[monStr], day: dayNum });
          }
        }
      });

      // 掃描所有含房型關鍵字的行區塊，讀取 input 數值（100% 唯讀，不觸發任何變更）
      const rows = [];
      const trs = Array.from(document.querySelectorAll('tr, [role="row"]'));
      for (const tr of trs) {
        const label = (tr.innerText || '').replace(/\s+/g, ' ').trim();
        if (!label) continue;
        const inputs = Array.from(tr.querySelectorAll('input')).map(inp => {
          if (inp.type === 'checkbox') return inp.checked ? 'Y' : 'N';
          return (inp.value || '').trim();
        });
        if (inputs.length >= 7) {
          rows.push({ label: label.slice(0, 100), inputs });
        }
      }
      return { logged_out: false, url, headers, row_count: rows.length, rows };
    })()
    """


def _js_parse_dashboard_health():
    """唯讀讀取 SiteMinder Dashboard 健康指標。"""
    return r"""
    (() => {
      const text = document.body ? document.body.innerText : '';
      const url = location.href;
      if (/login|signin|auth/i.test(url) || /Sign in to SiteMinder/i.test(text)) {
        return { logged_out: true, url };
      }
      const mMap = text.match(/(\d+)\s+of\s+(\d+)\s+rates?\s+mapped/i);
      const mFail = text.match(/Reservation\s+delivery\s+failures\s+today[^\d]{0,15}(\d+)/i)
                 || text.match(/(\d+)\s+delivery\s+failures?\s+today/i);
      return {
        logged_out: false,
        url,
        rates_mapped: mMap ? parseInt(mMap[1], 10) : null,
        rates_total: mMap ? parseInt(mMap[2], 10) : null,
        delivery_failures_today: mFail ? parseInt(mFail[1], 10) : 0,
      };
    })()
    """


def _match_room_code(label: str):
    for kw, code in ROOM_LABEL_TO_CODE:
        if kw in label:
            return code
    m = re.search(r"\b(DF|DTP|LDT|DT|PD|PT|SP|F|E|S)\b", label)
    return m.group(1) if m else None


def _save_sync_status(status: str, mode: str, detail: str, days_captured: int = 0):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    rec = {
        "last_attempt_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "status": status,
        "mode": mode,
        "detail": detail,
        "days_captured": days_captured,
    }
    with open(SYNC_STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(rec, f, ensure_ascii=False, indent=2)
    return rec


def run_siteminder_sync(unattended: bool = True, windows_count: int = 3):
    """
    100% 唯讀嘗試同步 SiteMinder 最新 Dashboard 健康度與 Inventory 掛牌價快照。
    若環境未就緒或已登出，自動優雅降級保留舊快照，回傳狀態字典。
    """
    client, sid, mode = _discover_siteminder_target(unattended=unattended)
    if not client:
        msg = f"SiteMinder CDP 未連線 ({mode})，保留現有快照不中斷排程。"
        logger.warning(f"[WARN] {msg}")
        return _save_sync_status("skipped_no_cdp", mode, msg)

    try:
        # 1. 讀取現有快照當底稿（確保未變欄位如失敗明細不遺失）
        existing = {}
        if os.path.exists(SNAPSHOT_PATH):
            try:
                with open(SNAPSHOT_PATH, "r", encoding="utf-8") as f:
                    existing = json.load(f)
            except Exception:
                existing = {}

        # 2. 前往 Dashboard 讀取健康指標（唯讀）
        client.send("Page.navigate", {"url": DASHBOARD_URL}, session_id=sid, timeout=15)
        time.sleep(4.0)
        dash = client.eval_js(_js_parse_dashboard_health(), session_id=sid, timeout=10) or {}
        if dash.get("logged_out"):
            msg = "SiteMinder 登入階段已過期（跳轉登入頁），依安全規範停止並保留舊快照，請於日間手動登入後重試。"
            logger.warning(f"[WARN] {msg}")
            return _save_sync_status("logged_out", mode, msg)

        # 3. 前往 Inventory 讀取多期 14 天網格（唯讀，絕不點擊 Save/Update）
        client.send("Page.navigate", {"url": INVENTORY_URL}, session_id=sid, timeout=15)
        time.sleep(5.0)

        inv_merged = dict(existing.get("inventory", {}))
        today = datetime.date.today()
        captured_dates = set()

        for w_idx in range(windows_count):
            parsed = client.eval_js(_js_parse_inventory_window(), session_id=sid, timeout=12) or {}
            if parsed.get("logged_out"):
                break

            headers = parsed.get("headers", [])
            # 將 (mon, day) 還原為正確 ISO 年份（跨年自動 +1）
            iso_dates = []
            cur_y = today.year
            prev_m = today.month
            for h in headers:
                m_num, d_num = int(h["mon"]), int(h["day"])
                if m_num < prev_m and (prev_m - m_num) >= 6:
                    cur_y += 1
                prev_m = m_num
                try:
                    iso_dates.append(datetime.date(cur_y, m_num, d_num).strftime("%Y-%m-%d"))
                except ValueError:
                    iso_dates.append(None)

            # 解析每一房型的可售數 (avail) 與 標準房價(不含早) 掛牌價
            active_rt = None
            for row in parsed.get("rows", []):
                lbl = row.get("label", "")
                inps = row.get("inputs", [])
                rt_hit = _match_room_code(lbl)
                if rt_hit and "標準房價" not in lbl and "含早" not in lbl:
                    active_rt = rt_hit
                    # 可售庫存列（通常為純小整數 0~30）
                    nums = [int(re.sub(r"[^\d]", "", x)) for x in inps if re.sub(r"[^\d]", "", x).isdigit()]
                    if len(inps) >= len(iso_dates) and iso_dates and nums and max(nums) <= 50:
                        for idx, d_iso in enumerate(iso_dates):
                            if not d_iso or idx >= len(inps):
                                continue
                            val_s = re.sub(r"[^\d]", "", inps[idx])
                            if val_s.isdigit():
                                inv_merged.setdefault(d_iso, {"ss": "N", "mlos": 1, "rooms": {}})
                                inv_merged[d_iso]["rooms"].setdefault(active_rt, {"r": 0, "a": 0})
                                inv_merged[d_iso]["rooms"][active_rt]["a"] = int(val_s)
                                captured_dates.add(d_iso)
                elif "標準房價(不含早)" in lbl and active_rt:
                    for idx, d_iso in enumerate(iso_dates):
                        if not d_iso or idx >= len(inps):
                            continue
                        val_s = re.sub(r"[^\d.]", "", inps[idx])
                        try:
                            price = int(round(float(val_s)))
                        except Exception:
                            continue
                        if price >= 1000:
                            inv_merged.setdefault(d_iso, {"ss": "N", "mlos": 1, "rooms": {}})
                            inv_merged[d_iso]["rooms"].setdefault(active_rt, {"r": 0, "a": 0})
                            inv_merged[d_iso]["rooms"][active_rt]["r"] = price
                            captured_dates.add(d_iso)

            # 翻下一頁 14 天（純瀏覽導覽按鈕，非資料修改鍵）
            if w_idx < windows_count - 1:
                clicked = client.eval_js(
                    """(() => {
                        const btn = document.querySelector('button[aria-label*="next" i], button[title*="next" i], [data-testid*="next-period"], .sm-calendar-next');
                        if (btn) { btn.click(); return true; }
                        return false;
                    })()""",
                    session_id=sid,
                    timeout=8,
                )
                if not clicked:
                    break
                time.sleep(3.5)

        # 4. 乾淨切回 Dashboard
        client.send("Page.navigate", {"url": DASHBOARD_URL}, session_id=sid, timeout=12)

        # 5. 更新並寫入快照（僅保留 today 以降日期以免膨脹）
        cutoff = today.strftime("%Y-%m-%d")
        inv_clean = {k: v for k, v in sorted(inv_merged.items()) if k >= cutoff}

        health = dict(existing.get("health", {}))
        if dash.get("rates_mapped") is not None:
            health["rates_mapped"] = dash["rates_mapped"]
        if dash.get("rates_total") is not None:
            health["rates_total"] = dash["rates_total"]
        if dash.get("delivery_failures_today") is not None:
            health["delivery_failures_today"] = dash["delivery_failures_today"]

        out_payload = {
            "property_name": existing.get("property_name", "Roaders Plus Hotel Taipei Station"),
            "property_id": PROPERTY_ID,
            "as_of": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source": f"SiteMinder Read-Only CDP Auto-Sync ({mode})",
            "rate_plan": "標準房價(不含早)",
            "currency": "TWD",
            "health": health,
            "inventory": inv_clean,
        }
        os.makedirs(BACKUP_DIR, exist_ok=True)
        with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
            json.dump(out_payload, f, ensure_ascii=False, indent=2)
        archive_siteminder_snapshot(out_payload)

        msg = f"SiteMinder 唯讀同步成功 (mode={mode}, 擷取 {len(captured_dates)} 天, 總快照 {len(inv_clean)} 天, 今日漏單={health.get('delivery_failures_today', 0)})"
        logger.info(f"[SUCCESS] {msg}")
        return _save_sync_status("success", mode, msg, days_captured=len(captured_dates))

    except Exception as e:
        msg = f"SiteMinder CDP 同步發生例外 ({mode}): {e}，已保留既有有效快照。"
        logger.warning(f"[WARN] {msg}")
        return _save_sync_status("error_kept_snapshot", mode, msg)
    finally:
        client.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    parser = argparse.ArgumentParser(description="SiteMinder 100% 唯讀 CDP 快照同步工具")
    parser.add_argument("--interactive", action="store_true", help="日間互動模式（給予較長等待時間供點擊 Chrome 允許）")
    parser.add_argument("--windows", type=int, default=3, help="擷取 14 天視窗頁數（預設 3 頁 = 42 天）")
    parser.add_argument("--git-push", action="store_true", help="同步成功後自動推播快照至 GitHub")
    args = parser.parse_args()

    res = run_siteminder_sync(unattended=not args.interactive, windows_count=args.windows)
    print("SiteMinder sync result:", json.dumps(res, ensure_ascii=False))
    if args.git_push and res.get("status") == "success":
        import subprocess
        subprocess.run([
            "git", "add",
            "data/pms_backup/siteminder_inventory_snapshot.json",
            "data/pms_backup/siteminder_sync_status.json",
            "data/pms_backup/sm_history/",
        ], cwd=APP_DIR, check=True)
        diff_res = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=APP_DIR)
        if diff_res.returncode != 0:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            subprocess.run(["git", "commit", "-m", f"chore(sync): auto-sync SiteMinder snapshot {ts} [skip ci]"], cwd=APP_DIR, check=True)
            subprocess.run(["git", "push", "origin", "main"], cwd=APP_DIR, check=True)
            print("[OK] 已推播至 GitHub 雲端版 Hotel Master。")
