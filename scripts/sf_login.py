# -*- coding: utf-8 -*-
"""
sf_login.py — 開啟 ChromeAutomation profile，手動登入 SF HK 網站
登入完成後關閉視窗即可
"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
from playwright.sync_api import sync_playwright

CHROME_PROFILE = r"C:\ChromeAutomation"
SF_URL = "https://hk.sf-express.com/hk/tc/"

print("正在開啟 Chrome...")
print("請在瀏覽器登入 SF HK 帳號，完成後按這裡 Enter 關閉")

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE,
        channel="chrome",
        headless=False,
        args=["--disable-blink-features=AutomationControlled",
              "--disable-session-crashed-bubble",
              "--disable-infobars"],
        no_viewport=True,
    )
    page = ctx.new_page()
    page.goto(SF_URL, wait_until="domcontentloaded", timeout=30000)
    print("\n✅ 瀏覽器已開啟，請手動登入 SF HK 網站")
    print("   登入完成後回來按 Enter...")
    input()
    ctx.close()
    print("✅ 完成！可以重新跑 clearance_upload.py 了")
