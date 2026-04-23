# -*- coding: utf-8 -*-
"""Deep inspection of POS product cards and detail page."""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=150, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)
    time.sleep(2)

    # ── Inspect product cards ──────────────────────────────────────────────────
    print("=== PRODUCT CARDS (first 5) ===")
    cards = page.locator(".product-card, .card, [class*='product'], [class*='item']").all()
    print(f"Found {len(cards)} elements matching product selectors")

    # Try to get inner HTML of first card
    if cards:
        print("\nFirst card HTML (truncated):")
        try:
            html = cards[0].inner_html()
            print(html[:800])
        except:
            pass

    # Get all text blocks that look like SKUs (numbers)
    print("\n=== Possible SKU/code text on page ===")
    all_text = page.inner_text("body")
    import re
    codes = re.findall(r'\b10\d{5}\b', all_text)
    print("Codes matching 10xxxxx pattern:", list(set(codes))[:15])

    # ── Click first 詳細介紹 button ────────────────────────────────────────────
    print("\n=== Clicking first 詳細介紹 button ===")
    detail_btns = page.locator("button:has-text('詳細介紹')").all()
    print(f"Found {len(detail_btns)} 詳細介紹 buttons")

    if detail_btns:
        detail_btns[0].click()
        time.sleep(2)
        shot = os.path.join(LOGS_DIR, "pos_02_product_detail.png")
        page.screenshot(path=shot, full_page=True)
        print(f"Screenshot: {shot}")

        print("\nDetail page buttons:")
        for el in page.locator("button").all():
            try:
                txt = el.inner_text().strip()
                if txt: print(f"  {txt!r}")
            except: pass

        print("\nDetail page inputs:")
        for el in page.locator("input, textarea").all():
            try:
                t = el.get_attribute("type") or ""
                ph = el.get_attribute("placeholder") or ""
                val = el.input_value() if t != "password" else "***"
                print(f"  type={t!r} placeholder={ph!r} value={val!r}")
            except: pass

        print("\nAll text on detail page (relevant excerpts):")
        detail_text = page.inner_text("body")
        # Print lines with numbers or keywords
        for line in detail_text.split('\n'):
            line = line.strip()
            if line and (re.search(r'\d{4,}', line) or
                         any(kw in line for kw in ['規格','材質','產地','價','SKU','貨號','成分'])):
                print(f"  {line!r}")

        codes2 = re.findall(r'\b10\d{5}\b', detail_text)
        print(f"\nSKU-like codes in detail: {list(set(codes2))}")

    # ── Check backend ──────────────────────────────────────────────────────────
    print("\n=== Navigating to backend ===")
    page.goto(POS_URL, wait_until="domcontentloaded")
    time.sleep(1)
    backend_btn = page.locator("button:has-text('后台管理')").first
    if backend_btn.is_visible(timeout=3000):
        backend_btn.click()
        time.sleep(1)
        shot3 = os.path.join(LOGS_DIR, "pos_03_backend.png")
        page.screenshot(path=shot3)
        print(f"Backend screenshot: {shot3}")
        print("Backend buttons:")
        for el in page.locator("button").all()[:15]:
            try:
                txt = el.inner_text().strip()
                if txt: print(f"  {txt!r}")
            except: pass

    ctx.close()
    print("\nDone.")
