# -*- coding: utf-8 -*-
"""
CLI wrapper for status_updater — called by tracking_dashboard.py via subprocess.
Prints progress to stdout so the dashboard can capture it.
"""
import sys
import os
import subprocess
import time
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_manager import ExcelManager
from status_updater import update_all_statuses


def _kill_chrome():
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], capture_output=True)
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe", "/T"], capture_output=True)
    for _ in range(12):
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq chrome.exe"],
                           capture_output=True, text=True)
        if "chrome.exe" not in r.stdout:
            break
        time.sleep(1)
    time.sleep(2)


if __name__ == "__main__":
    print("=== 開始查詢順豐狀態 ===")
    _kill_chrome()
    excel = ExcelManager()
    print("正在連接順豐網站（最多 3 分鐘）…")
    results = update_all_statuses(excel)
    for wb, status in results.items():
        print(f"  {wb} → {status}")
    print(f"\n=== 完成：已更新 {len(results)} 個運單 ===")
