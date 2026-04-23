# -*- coding: utf-8 -*-
"""
Inspect POS cashier flow:
  search → add to cart → cart page → checkout → receipt
"""
import os, sys, time, re
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, POS_VIP_PASS, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

TEST_SKU = "1000043"   # 瓜拿那 — change to any SKU you want to test

def shot(page, name):
    p = os.path.join(LOGS_DIR, f"cashier_{name}.png")
    page.screenshot(path=p, full_page=True)
    print(f"  📸 {p}")

def dump_buttons(page, label):
    print(f"\n--- Buttons [{label}] ---")
    for el in page.locator("button").all():
        try:
            t = el.inner_text().strip()
            if t: print(f"  {t!r}")
        except: pass

def dump_inputs(page, label):
    print(f"\n--- Inputs [{label}] ---")
    for el in page.locator("input, textarea").all():
        try:
            t  = el.get_attribute("type") or ""
            ph = el.get_attribute("placeholder") or ""
            v  = el.input_value() if t != "password" else "***"
            if ph or v: print(f"  type={t!r} ph={ph!r} val={v!r}")
        except: pass

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=200, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    # ── Step 1: Activate VIP ──────────────────────────────────────────────────
    print("\n=== Step 1: Activate VIP ===")
    page.locator("button:has-text('VIP')").first.click()
    time.sleep(0.8)
    page.locator("input[type='password']").first.fill(POS_VIP_PASS)
    page.keyboard.press("Enter")
    time.sleep(1.5)
    shot(page, "01_after_vip")
    print("VIP activated")

    # ── Step 2: Search for SKU ────────────────────────────────────────────────
    print(f"\n=== Step 2: Search for SKU {TEST_SKU} ===")
    search = page.locator("input[placeholder='🔍 搜尋商品...']").first
    search.click()
    search.fill(TEST_SKU)
    time.sleep(1.5)
    shot(page, "02_after_search")
    dump_buttons(page, "after search")

    # Check if search filtered results or showed dropdown
    all_text = page.inner_text("body")
    print("\nLines with SKU after search:")
    for line in all_text.split('\n'):
        if TEST_SKU in line or '瓜拿那' in line:
            print(f"  {line.strip()!r}")

    # ── Step 3: Add to cart ───────────────────────────────────────────────────
    print("\n=== Step 3: Add to cart ===")
    # Find visible '+ 加入' buttons
    add_btns = page.locator("button:has-text('+ 加入')").all()
    print(f"  Found {len(add_btns)} '+ 加入' buttons visible")
    if add_btns:
        add_btns[0].click()
        time.sleep(1.0)
        shot(page, "03_after_add")

    # Check cart count
    cart_btn = page.locator("button:has-text('購物車')").first
    cart_text = cart_btn.inner_text() if cart_btn.is_visible(timeout=2000) else ""
    print(f"  Cart button text: {cart_text!r}")

    # ── Step 4: Open cart ─────────────────────────────────────────────────────
    print("\n=== Step 4: Open cart ===")
    cart_btn.click()
    time.sleep(1.5)
    shot(page, "04_cart_open")
    dump_buttons(page, "cart open")
    dump_inputs(page, "cart open")

    print("\nCart page text (relevant):")
    cart_text = page.inner_text("body")
    for line in cart_text.split('\n'):
        line = line.strip()
        if line and any(kw in line for kw in
                        ['$','數量','件','qty','total','合計','小計','結帳',
                         '付款','現金','確認','SKU','貨號', TEST_SKU, '瓜拿那']):
            print(f"  {line!r}")

    shot(page, "04b_cart_detail")

    # ── Step 5: Look for checkout button ─────────────────────────────────────
    print("\n=== Step 5: Checkout button ===")
    for text in ['結帳', '付款', 'Checkout', '確認', '下單']:
        btns = page.locator(f"button:has-text('{text}')").all()
        if btns:
            print(f"  Found button: {text!r} × {len(btns)}")

    # ── Step 6: Click checkout ────────────────────────────────────────────────
    checkout_candidates = ['結帳', '付款', 'Checkout']
    for label in checkout_candidates:
        btn = page.locator(f"button:has-text('{label}')").first
        if btn.is_visible(timeout=1000):
            print(f"\n=== Step 6: Clicking '{label}' ===")
            btn.click()
            time.sleep(1.5)
            shot(page, "05_checkout")
            dump_buttons(page, "checkout screen")
            dump_inputs(page, "checkout screen")
            print("\nCheckout text:")
            for line in page.inner_text("body").split('\n'):
                line = line.strip()
                if line and any(kw in line for kw in
                                ['現金','PayMe','信用卡','付款','方式','金額',
                                 '$','確認','取消','收據','小票','完成']):
                    print(f"  {line!r}")
            break

    # ── Step 7: Look for payment options ─────────────────────────────────────
    print("\n=== Step 7: Payment options ===")
    for text in ['現金', 'PayMe', '信用卡', 'Cash']:
        els = page.locator(f"button:has-text('{text}'), label:has-text('{text}')").all()
        if els:
            print(f"  Payment option: {text!r} × {len(els)}")

    print("\n\nBrowser stays open for manual inspection.")
    print("Look at the screenshots in:", LOGS_DIR)
    input("Press Enter to close...")
    ctx.close()
