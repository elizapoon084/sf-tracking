# -*- coding: utf-8 -*-
"""Inspect product HTML structure and VIP prices."""
import os, sys, time, re
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR, POS_VIP_PASS

os.makedirs(LOGS_DIR, exist_ok=True)

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=150, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    # ── Dump first product card HTML ──────────────────────────────────────────
    print("=== First product card HTML (outer) ===")
    # Try to find product container by looking at children of main content
    main_html = page.locator("main, #app, #root, .products, [class*='grid'], [class*='product']").first
    try:
        print(main_html.inner_html()[:2000])
    except:
        # Just dump a portion of body HTML to see structure
        body_html = page.locator("body").inner_html()
        # Find the section with product buttons
        idx = body_html.find('加入')
        if idx > 0:
            print("HTML around '加入' button:")
            print(body_html[max(0,idx-500):idx+200])

    # ── Normal prices on cards ────────────────────────────────────────────────
    print("\n=== Price elements on page ===")
    for sel in ['[class*="price"]', '[class*="Price"]', 'span.price', '.product-price',
                'span[style*="color"]', '[class*="amount"]', '[class*="Amount"]']:
        els = page.locator(sel).all()
        if els:
            print(f"Selector {sel!r} → {len(els)} elements")
            for el in els[:3]:
                try:
                    print(f"  text={el.inner_text().strip()!r}  class={el.get_attribute('class')!r}")
                except: pass
            break

    # ── Click VIP button, enter password ──────────────────────────────────────
    print("\n=== Activating VIP mode ===")
    vip_btn = page.locator("button:has-text('VIP')").first
    vip_btn.click()
    time.sleep(1)
    shot_vip = os.path.join(LOGS_DIR, "pos_04_vip_dialog.png")
    page.screenshot(path=shot_vip)
    print(f"VIP dialog screenshot: {shot_vip}")

    print("VIP dialog inputs:")
    for el in page.locator("input").all():
        try:
            t = el.get_attribute("type") or ""
            ph = el.get_attribute("placeholder") or ""
            print(f"  type={t!r} placeholder={ph!r}")
        except: pass

    # Enter VIP password
    pwd = page.locator("input[type='password'], input[type='text']").last
    try:
        pwd.fill(POS_VIP_PASS)
        page.keyboard.press("Enter")
        time.sleep(1.5)
    except Exception as e:
        print(f"Password entry error: {e}")

    shot_after = os.path.join(LOGS_DIR, "pos_05_after_vip.png")
    page.screenshot(path=shot_after)
    print(f"After VIP screenshot: {shot_after}")

    # ── Check prices after VIP ─────────────────────────────────────────────────
    print("\n=== Prices after VIP mode (first 5 products) ===")
    all_text = page.inner_text("body")
    # Find price patterns: $xxx or HK$xxx
    prices = re.findall(r'HK?\$\s*(\d+\.?\d*)', all_text)
    print("Price values found:", prices[:20])

    # ── Click first 詳細介紹 in VIP mode ─────────────────────────────────────
    print("\n=== Product modal in VIP mode ===")
    detail_btns = page.locator("button:has-text('詳細介紹')").all()
    if detail_btns:
        detail_btns[0].click()
        time.sleep(1.5)
        shot_modal = os.path.join(LOGS_DIR, "pos_06_modal_vip.png")
        page.screenshot(path=shot_modal)
        print(f"Modal VIP screenshot: {shot_modal}")

        modal_text = page.inner_text("body")
        # Show lines with price info
        for line in modal_text.split('\n'):
            line = line.strip()
            if line and re.search(r'\$\d+|價|Price|HK', line):
                print(f"  {line!r}")

        # Also dump modal HTML
        print("\n=== Modal HTML ===")
        for sel in ['[class*="modal"]', '[class*="dialog"]', '[class*="popup"]',
                    '[class*="overlay"]', '[role="dialog"]', '.detail']:
            modal = page.locator(sel).first
            try:
                if modal.is_visible(timeout=1000):
                    print(f"Modal selector: {sel!r}")
                    print(modal.inner_html()[:1500])
                    break
            except: pass

    ctx.close()
    print("\nDone.")
