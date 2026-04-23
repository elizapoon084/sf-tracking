# -*- coding: utf-8 -*-
"""Inspect correct add-to-cart + checkout flow."""
import os, sys, time, re
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, POS_VIP_PASS, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

def shot(page, name):
    p = os.path.join(LOGS_DIR, f"cashier2_{name}.png")
    page.screenshot(path=p, full_page=False)  # viewport only – faster
    print(f"  📸 {p}")
    return p

def dump_buttons(page):
    for el in page.locator("button").all():
        try:
            t = el.inner_text().strip()
            if t and t not in ('🏪 前台商店','🔐 后台管理','🚚 送貨','📋 退貨','👤 Poon','登出'):
                print(f"    BTN: {t!r}")
        except: pass

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=200, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    # Activate VIP
    page.locator("button:has-text('VIP')").first.click()
    time.sleep(0.8)
    page.locator("input[type='password']").first.fill(POS_VIP_PASS)
    page.keyboard.press("Enter")
    time.sleep(1.5)
    print("VIP activated")

    # ── Test 1: Search by NAME ─────────────────────────────────────────────────
    print("\n=== Test 1: Search by NAME '瓜拿那' ===")
    search = page.locator("input[placeholder='🔍 搜尋商品...']").first
    search.fill("瓜拿那")
    time.sleep(1.5)
    shot(page, "01_search_name")
    add_btns = page.locator("button:has-text('+ 加入')").all()
    print(f"  '+ 加入' buttons visible: {len(add_btns)}")
    dump_buttons(page)

    # ── Test 2: Click '+ 加入' and see what happens ────────────────────────────
    if add_btns:
        print("\n=== Test 2: Click '+ 加入' ===")
        add_btns[0].click()
        time.sleep(1.0)
        shot(page, "02_after_add")
        print("  After clicking + 加入:")
        dump_buttons(page)
        cart_btn = page.locator("button:has-text('購物車')").first
        print(f"  Cart button text: {cart_btn.inner_text()!r}")

    # ── Test 3: Add same item again (simulate qty=2) ───────────────────────────
    print("\n=== Test 3: Add same item again (qty=2 total) ===")
    add_btns2 = page.locator("button:has-text('+ 加入')").all()
    if add_btns2:
        add_btns2[0].click()
        time.sleep(0.8)
        cart_btn = page.locator("button:has-text('購物車')").first
        print(f"  Cart button text after 2nd add: {cart_btn.inner_text()!r}")

    # ── Test 4: Open cart ─────────────────────────────────────────────────────
    print("\n=== Test 4: Open cart ===")
    page.locator("button:has-text('購物車')").first.click()
    time.sleep(1.5)
    shot(page, "03_cart")
    print("  Cart content:")
    dump_buttons(page)

    # Close any popup with ✕ first
    close_btns = page.locator("button:has-text('✕'), button:has-text('×')").all()
    if close_btns:
        print(f"  Closing {len(close_btns)} popup(s)...")
        for b in close_btns:
            try:
                if b.is_visible(timeout=500):
                    b.click()
                    time.sleep(0.5)
            except: pass

    shot(page, "04_cart_after_close_popup")
    print("  After closing popups:")
    dump_buttons(page)
    print("\n  Full cart text (relevant lines):")
    for line in page.inner_text("body").split('\n'):
        l = line.strip()
        if l and any(kw in l for kw in ['$','件','qty','數量','合計','小計',
                                         '結帳','付款','確認','現金','收據',
                                         '瓜拿那','1000043','購物車']):
            print(f"    {l!r}")

    # ── Test 5: Find and click checkout ───────────────────────────────────────
    print("\n=== Test 5: Find checkout button ===")
    for label in ['結帳', '立即結帳', '付款', 'Checkout', '確認訂單']:
        b = page.locator(f"button:has-text('{label}')").first
        if b.is_visible(timeout=500):
            print(f"  Found: {label!r}")
            b.click()
            time.sleep(1.5)
            shot(page, f"05_checkout_{label}")
            print(f"  After clicking '{label}':")
            dump_buttons(page)
            print("  Checkout text:")
            for line in page.inner_text("body").split('\n'):
                l = line.strip()
                if l and any(kw in l for kw in ['$','現金','PayMe','付款',
                                                  '確認','取消','小票','收據']):
                    print(f"    {l!r}")
            break

    # ── Test 6: Payment options after checkout ────────────────────────────────
    print("\n=== Test 6: Payment options ===")
    for label in ['現金', 'PayMe', '信用卡', 'Cash', '完成', '確認付款']:
        b = page.locator(f"button:has-text('{label}'), label:has-text('{label}')").first
        try:
            if b.is_visible(timeout=500):
                print(f"  Payment option visible: {label!r}")
        except: pass

    shot(page, "06_final_state")
    print("\n\nDone. Browser stays open for manual inspection.")
    print(f"Screenshots in: {LOGS_DIR}")
    input("Press Enter to close...")
    ctx.close()
