# -*- coding: utf-8 -*-
"""Inspect POS backend (后台管理) for cashier/收銀台 interface."""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, POS_ADMIN_PASS, POS_VIP_PASS, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

def shot(page, name):
    p = os.path.join(LOGS_DIR, f"backend_{name}.png")
    page.screenshot(path=p, full_page=False)
    print(f"  📸 {p}")

def dump_buttons(page):
    seen = set()
    for el in page.locator("button, a[href]").all():
        try:
            t = el.inner_text().strip()
            if t and t not in seen and len(t) < 40:
                seen.add(t)
                print(f"    {t!r}")
        except: pass

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=200, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)
    shot(page, "00_landing")

    # ── Click 后台管理 ─────────────────────────────────────────────────────────
    print("\n=== Click 后台管理 ===")
    backend_btn = page.locator("button:has-text('后台管理')").first
    backend_btn.click()
    time.sleep(1.5)
    shot(page, "01_after_backend_click")

    # Check if password prompt appeared
    pwd_inputs = page.locator("input[type='password']").all()
    print(f"Password inputs visible: {len(pwd_inputs)}")
    for el in pwd_inputs:
        ph = el.get_attribute("placeholder") or ""
        print(f"  placeholder={ph!r}")

    if pwd_inputs:
        print(f"Entering admin password: {POS_ADMIN_PASS!r}")
        pwd_inputs[0].fill(POS_ADMIN_PASS)
        page.keyboard.press("Enter")
        time.sleep(1.5)
        shot(page, "02_after_admin_login")

    print("\nButtons after backend login:")
    dump_buttons(page)
    print(f"\nCurrent URL: {page.url}")

    # ── Look for 收銀台 or cashier ──────────────────────────────────────────────
    print("\n=== Looking for 收銀台 ===")
    for label in ['收銀台', '收銀', 'POS', '點餐', '銷售', '下單', '新增訂單', '創建訂單']:
        btns = page.locator(f"button:has-text('{label}'), a:has-text('{label}')").all()
        if btns:
            print(f"  Found: {label!r} × {len(btns)}")

    # Navigate to each backend section
    nav_items = page.locator("nav a, nav button, .sidebar a, .menu-item, [class*='nav'] a").all()
    print(f"\nNavigation items: {len(nav_items)}")
    for el in nav_items:
        try:
            t = el.inner_text().strip()
            h = el.get_attribute("href") or ""
            if t: print(f"  {t!r}  href={h!r}")
        except: pass

    # ── Full page text for clues ────────────────────────────────────────────────
    print("\n=== Backend page text ===")
    body_text = page.inner_text("body")
    for line in body_text.split('\n'):
        l = line.strip()
        if l and l not in ('MANLEE', 'Health & Wellness'):
            print(f"  {l!r}")

    shot(page, "03_backend_full")

    # ── Try VIP mode in backend ────────────────────────────────────────────────
    print("\n=== Try VIP button in backend ===")
    vip_btn = page.locator("button:has-text('VIP')").first
    try:
        if vip_btn.is_visible(timeout=2000):
            vip_btn.click()
            time.sleep(0.8)
            shot(page, "04_vip_in_backend")
            dump_buttons(page)
            pwd = page.locator("input[type='password']").first
            if pwd.is_visible(timeout=1000):
                pwd.fill(POS_VIP_PASS)
                page.keyboard.press("Enter")
                time.sleep(1.5)
                shot(page, "05_after_vip")
                print("VIP in backend activated")
                dump_buttons(page)
    except Exception as e:
        print(f"VIP: {e}")

    shot(page, "06_final")
    print(f"\nDone. Screenshots: {LOGS_DIR}")
    input("Press Enter to close...")
    ctx.close()
