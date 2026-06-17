# -*- coding: utf-8 -*-
"""
camp_login.py — 手動登入 camp.sf-express.com，儲存 session 到 ChromeAutomation
登入完成後按 Enter 關閉即可，之後可以正常跑 demo_full_flow_v62.py
"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
from playwright.sync_api import sync_playwright

CHROME_PROFILE = r"C:\ChromeAutomation"
CAMP_URL = "https://camp.sf-express.com/MonthCard"

print("正在開啟 Chrome，請在瀏覽器登入 camp.sf-express.com...")
print("（用你的月結帳號 / 企業帳號登入）")

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
    page.goto(CAMP_URL, wait_until="domcontentloaded", timeout=30000)
    print("\n✅ 瀏覽器已開啟 camp.sf-express.com")
    print("   請在瀏覽器完成登入，登入後回來按 Enter 關閉...")
    input()
    ctx.close()
    print("✅ 完成！現在可以跑 demo_full_flow_v62.py 了")
