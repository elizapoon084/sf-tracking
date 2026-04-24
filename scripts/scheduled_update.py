# -*- coding: utf-8 -*-
"""
scheduled_update.py — 定時自動更新順丰狀態 + 同步 Google Sheets

執行方式:
  python scripts/scheduled_update.py

Windows Task Scheduler 設定:
  動作: python scripts/scheduled_update.py
  起始路徑: C:/Users/user/Desktop/順丰E順递
  觸發: 每天 09:00 + 18:00
"""
import sys
import os
import subprocess
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import browser_utils
browser_utils.HEADLESS = True  # run silently in background, no Chrome window

from excel_manager import ExcelManager
from status_updater import update_all_statuses
from gsheets_sync import sync_excel_to_sheets
from logger import get_logger

log = get_logger("scheduled_update")


def _pull_from_github() -> None:
    """Pull latest changes from GitHub (picks up tax edits made on Streamlit Cloud)."""
    repo_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    result = subprocess.run(
        ["git", "pull"],
        cwd=repo_dir, capture_output=True, text=True, encoding="utf-8"
    )
    msg = result.stdout.strip() or result.stderr.strip()
    if result.returncode == 0:
        log.info("git pull: %s", msg)
        print(f"      git pull: {msg}")
    else:
        log.warning("git pull failed: %s", msg)
        print(f"      ⚠️  git pull 失敗: {msg}")


def _push_excel_to_github() -> None:
    """Commit tracking.xlsx and push to GitHub so Streamlit Cloud updates."""
    repo_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    excel_rel = os.path.join("data", "tracking.xlsx")
    now_str   = datetime.now().strftime("%Y-%m-%d %H:%M")

    def _git(*args):
        return subprocess.run(
            ["git"] + list(args),
            cwd=repo_dir, capture_output=True, text=True, encoding="utf-8"
        )

    # Check if there's anything to commit
    status = _git("status", "--porcelain", excel_rel)
    if not status.stdout.strip():
        print("      ℹ️  tracking.xlsx 無變化，略過推送")
        log.info("Git push skipped — no changes in tracking.xlsx")
        return

    _git("add", excel_rel)
    commit = _git("commit", "-m", f"auto: 更新運單狀態 {now_str}")
    if commit.returncode != 0:
        print(f"      ⚠️  git commit 失敗: {commit.stderr.strip()}")
        log.error("git commit failed: %s", commit.stderr)
        return

    push = _git("push")
    if push.returncode == 0:
        print("      ✅ 已推送去 GitHub，Streamlit 將在數分鐘內更新")
        log.info("Git push success → Streamlit Cloud will update")
    else:
        print(f"      ⚠️  git push 失敗: {push.stderr.strip()}")
        log.error("git push failed: %s", push.stderr)


def _close_chrome_if_running() -> None:
    """Kill all Chrome processes and wait until they're fully gone."""
    import time
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], capture_output=True)
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe", "/T"], capture_output=True)
    # Wait until no chrome.exe remains
    for _ in range(15):
        chk = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq chrome.exe"],
            capture_output=True, text=True
        )
        if "chrome.exe" not in chk.stdout:
            break
        time.sleep(1)
    time.sleep(2)  # extra buffer for OS cleanup
    log.info("Chrome processes cleared")


def run() -> None:
    start = datetime.now()
    log.info("=" * 55)
    log.info("定時更新開始  %s", start.strftime("%Y-%m-%d %H:%M:%S"))
    print(f"\n{'='*55}")
    print(f"定時更新開始  {start:%Y-%m-%d %H:%M:%S}")

    # ── 0. Pull latest from GitHub (picks up cloud tax edits) ─────────────────
    print("\n[0/3] 同步 GitHub 最新資料（包括雲端稅金更新）…")
    _pull_from_github()

    # ── 1. Scrape SF HK waybill list (3 pages = ~30 waybills) + 電子存根 ────────
    print("\n[1/3] 正在爬取順豐香港網站（3頁，最多30個運單）…")
    _close_chrome_if_running()
    excel   = ExcelManager()
    updated = 0
    try:
        results = update_all_statuses(excel)
        updated = len(results)
        print(f"      共更新 {updated} 個運單：")
        for wb, status in results.items():
            print(f"      {wb} → {status}")
    except Exception as e:
        log.error("爬取失敗: %s", e)
        print(f"      ⚠️  爬取失敗: {e}")

    # ── 2. Sync to Google Sheets ───────────────────────────────────────────────
    print("\n[2/3] 同步到 Google Sheets…")
    try:
        ws = excel.ws
        all_rows = [
            [cell.value for cell in row]
            for row in ws.iter_rows(min_row=2)
            if any(cell.value is not None for cell in row)
        ]
        ok = sync_excel_to_sheets(all_rows)
        if ok:
            print(f"      ✅ Google Sheets 已更新（{len(all_rows)} 行）")
        else:
            print("      ℹ️  Google Sheets 未設定，略過（見 README_GSHEETS）")
    except Exception as e:
        log.error("Sheets sync error: %s", e)
        print(f"      ⚠️  Sheets 同步失敗: {e}")

    # ── 3. Push tracking.xlsx to GitHub → Streamlit Cloud auto-updates ────────
    print("\n[3/3] 推送 tracking.xlsx 去 GitHub…")
    _push_excel_to_github()

    # ── Done ───────────────────────────────────────────────────────────────────
    elapsed = (datetime.now() - start).seconds
    print(f"\n{'='*55}")
    print(f"完成！耗時 {elapsed} 秒  ({datetime.now():%H:%M:%S})")
    log.info("定時更新完成，耗時 %ds，更新 %d 個運單", elapsed, updated)


if __name__ == "__main__":
    run()
