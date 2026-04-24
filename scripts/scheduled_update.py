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

from excel_manager import ExcelManager
from sf_china_scraper import scrape_all_waybills
from gsheets_sync import sync_excel_to_sheets
from logger import get_logger

log = get_logger("scheduled_update")


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

    # ── 1. Scrape SF China waybill list ────────────────────────────────────────
    print("\n[1/3] 正在爬取順豐中國網站運單列表…")
    _close_chrome_if_running()
    try:
        scraped = scrape_all_waybills()
        print(f"      共爬取到 {len(scraped)} 個運單")
    except Exception as e:
        log.error("爬取失敗: %s", e)
        print(f"      ⚠️  爬取失敗: {e}")
        scraped = []

    # ── 2. Update local Excel statuses ─────────────────────────────────────────
    print("\n[2/3] 更新本地 Excel 狀態…")
    excel   = ExcelManager()
    updated = 0
    skipped = 0

    if scraped:
        # Build lookup: {waybill: {status, freight, status_time}}
        scraped_map = {r["waybill"]: r for r in scraped}

        for _row_num, wb in excel.get_all_waybills():
            if wb in scraped_map:
                info = scraped_map[wb]
                status = info.get("status") or "狀態不明"
                excel.update_status(
                    wb,
                    status,
                    freight=info.get("freight", ""),
                    status_time=info.get("status_time", ""),
                )
                print(f"      {wb} → {status}")
                updated += 1
            else:
                skipped += 1

        print(f"      已更新 {updated} 個，未匹配 {skipped} 個")
    else:
        print("      （無爬取資料，跳過 Excel 更新）")

    # ── 3. Sync to Google Sheets ───────────────────────────────────────────────
    print("\n[3/3] 同步到 Google Sheets…")
    try:
        # Read all Excel rows (skip header)
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

    # ── 4. Push tracking.xlsx to GitHub → Streamlit Cloud auto-updates ────────
    print("\n[4/4] 推送 tracking.xlsx 去 GitHub…")
    _push_excel_to_github()

    # ── Done ───────────────────────────────────────────────────────────────────
    elapsed = (datetime.now() - start).seconds
    print(f"\n{'='*55}")
    print(f"完成！耗時 {elapsed} 秒  ({datetime.now():%H:%M:%S})")
    log.info("定時更新完成，耗時 %ds，更新 %d 個運單", elapsed, updated)


if __name__ == "__main__":
    run()
