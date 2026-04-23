# -*- coding: utf-8 -*-
"""
Diagnostic: open POS, take screenshots, dump page info.
Run: python inspect_pos.py
"""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE,
        channel="chrome",
        headless=False,
        args=BROWSER_ARGS,
        slow_mo=200,
        viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()

    print("Opening POS...")
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    # Screenshot of landing page
    shot1 = os.path.join(LOGS_DIR, "pos_01_landing.png")
    page.screenshot(path=shot1, full_page=True)
    print(f"Screenshot saved: {shot1}")

    # Print all visible buttons and links
    print("\n=== Buttons on page ===")
    for el in page.locator("button").all()[:20]:
        try:
            txt = el.inner_text().strip()
            if txt:
                print(f"  BUTTON: {txt!r}")
        except:
            pass

    print("\n=== Links on page ===")
    for el in page.locator("a").all()[:20]:
        try:
            txt = el.inner_text().strip()
            href = el.get_attribute("href") or ""
            if txt:
                print(f"  LINK: {txt!r}  href={href!r}")
        except:
            pass

    print("\n=== Input fields ===")
    for el in page.locator("input").all()[:10]:
        try:
            t = el.get_attribute("type") or ""
            ph = el.get_attribute("placeholder") or ""
            nm = el.get_attribute("name") or ""
            print(f"  INPUT type={t!r} placeholder={ph!r} name={nm!r}")
        except:
            pass

    print(f"\nCurrent URL: {page.url}")
    print("\nBrowser stays open — press Ctrl+C to close when done inspecting.")
    input("Press Enter to close browser...")
    ctx.close()
