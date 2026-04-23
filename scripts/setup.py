# -*- coding: utf-8 -*-
"""
One-time setup script. Run this FIRST before anything else.

  python setup.py

Does:
  1. Create all required directories
  2. Install Python packages
  3. Install Playwright Chromium
  4. Create the Chrome profile directory
  5. Create a blank tracking.xlsx
  6. Print next steps
"""
import os
import subprocess
import sys
from pathlib import Path

BASE_DIR = r"C:\Users\user\Desktop\順丰E順递"

DIRS = [
    BASE_DIR + r"\Images",
    BASE_DIR + r"\data",
    BASE_DIR + r"\logs",
    BASE_DIR + r"\scripts",
    r"C:\ChromeAutomation",
]

PACKAGES = [
    "playwright",
    "openpyxl",
    "zhconv",
    "win10toast",
    "Pillow",   # for screenshots
]


def main():
    print("=" * 60)
    print("順丰寄件自動化 — 環境設置")
    print("=" * 60)

    # 1. Create directories
    print("\n[1/4] 建立資料夾…")
    for d in DIRS:
        Path(d).mkdir(parents=True, exist_ok=True)
        print(f"  ✅ {d}")

    # 2. Install packages
    print("\n[2/4] 安裝 Python packages…")
    for pkg in PACKAGES:
        print(f"  → pip install {pkg}")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", pkg, "--quiet"],
            stdout=subprocess.DEVNULL,
        )
    print("  ✅ 所有 packages 安裝完成")

    # 3. Install Playwright browser
    print("\n[3/4] 安裝 Playwright Chromium…")
    subprocess.check_call(
        [sys.executable, "-m", "playwright", "install", "chromium"]
    )
    print("  ✅ Playwright Chromium 安裝完成")

    # 4. Create tracking.xlsx
    print("\n[4/4] 建立 tracking.xlsx…")
    sys.path.insert(0, os.path.join(BASE_DIR, "scripts"))
    from excel_manager import ExcelManager
    ExcelManager()  # creates file if not exists
    print(f"  ✅ {BASE_DIR}\\data\\tracking.xlsx")

    # Done
    print("\n" + "=" * 60)
    print("設置完成！")
    print("=" * 60)
    print("""
下一步:
  1. 手動用 Chrome 開以下頁面並登入（保存 session）:
       - POS:   https://online-store-99126206.web.app/
       - 順丰:  https://hk.sf-express.com/hk/tc/ship/home
     ⚠️  用呢個 profile 開 Chrome:
       chrome.exe --user-data-dir=C:\\ChromeAutomation

  2. 更新產品資料庫:
       python product_scraper.py

  3. 啟動主程式:
       python main_gui.py
""")


if __name__ == "__main__":
    main()
