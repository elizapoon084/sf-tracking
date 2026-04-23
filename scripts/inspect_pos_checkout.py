# -*- coding: utf-8 -*-
"""Inspect the 填寫資料及付款 checkout form and receipt flow."""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, POS_VIP_PASS, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

def shot(page, name):
    p = os.path.join(LOGS_DIR, f"checkout_{name}.png")
    page.screenshot(path=p, full_page=False)
    print(f"  📸 {p}")

def dump_all(page, label):
    print(f"\n  -- Buttons [{label}] --")
    for el in page.locator("button").all():
        try:
            t = el.inner_text().strip()
            if t and len(t) < 30:
                print(f"    {t!r}")
        except: pass
    print(f"  -- Inputs [{label}] --")
    for el in page.locator("input, textarea, select").all():
        try:
            t  = el.get_attribute("type") or ""
            ph = el.get_attribute("placeholder") or ""
            nm = el.get_attribute("name") or ""
            v  = "" if t == "password" else el.input_value()
            print(f"    type={t!r} ph={ph!r} name={nm!r} val={v[:30]!r}")
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

    # Add 2 items to cart
    search = page.locator("input[placeholder='🔍 搜尋商品...']").first
    search.fill("瓜拿那")
    time.sleep(1.2)
    add_btn = page.locator("button:has-text('+ 加入')").first
    add_btn.click(); time.sleep(0.5)
    add_btn.click(); time.sleep(0.5)  # qty = 2
    print(f"Cart: {page.locator('button:has-text(\"購物車\")').first.inner_text()!r}")

    # Open cart
    page.locator("button:has-text('購物車')").first.click()
    time.sleep(1.5)
    shot(page, "01_cart_open")
    print("\n=== Cart open ===")
    dump_all(page, "cart")

    # ── Click 填寫資料及付款 ──────────────────────────────────────────────────
    checkout_btn = page.locator("button:has-text('填寫資料及付款')").first
    if checkout_btn.is_visible(timeout=3000):
        print("\n=== Clicking 填寫資料及付款 → ===")
        checkout_btn.click()
        time.sleep(2)
        shot(page, "02_checkout_form")
        dump_all(page, "checkout form")

        print("\n  Checkout form text (all lines):")
        for line in page.inner_text("body").split('\n'):
            l = line.strip()
            if l and l not in ('🏪 前台商店','🔐 后台管理','MANLEE','Health & Wellness'):
                print(f"    {l!r}")
    else:
        print("  '填寫資料及付款' button NOT visible after cart open")

    # ── Try filling checkout form ──────────────────────────────────────────────
    print("\n=== Filling checkout form ===")
    inputs = page.locator("input, textarea").all()
    for el in inputs:
        try:
            ph = el.get_attribute("placeholder") or ""
            nm = el.get_attribute("name") or ""
            t  = el.get_attribute("type") or ""
            print(f"  Input: type={t!r} ph={ph!r} name={nm!r}")
        except: pass

    # Fill name field if present
    for ph in ['姓名', '客人', '名字', 'name', 'Name']:
        try:
            f = page.locator(f"input[placeholder*='{ph}'], input[name*='{ph.lower()}']").first
            if f.is_visible(timeout=500):
                f.fill("測試客人")
                print(f"  Filled name field ({ph!r})")
                break
        except: pass

    shot(page, "03_filled_form")

    # ── Look for confirm/submit/pay ───────────────────────────────────────────
    print("\n=== Submit / Confirm buttons ===")
    for label in ['確認', '提交', '付款', '完成訂單', '下單', '確認付款',
                  'Submit', 'Confirm', '現金', 'PayMe']:
        b = page.locator(f"button:has-text('{label}')").first
        try:
            if b.is_visible(timeout=500):
                print(f"  FOUND: {label!r}")
        except: pass

    # ── Click 確認 / first payment option ────────────────────────────────────
    for label in ['確認', '現金', '確認付款', '完成訂單']:
        b = page.locator(f"button:has-text('{label}')").first
        try:
            if b.is_visible(timeout=500):
                print(f"\n=== Clicking '{label}' ===")
                b.click()
                time.sleep(2)
                shot(page, f"04_after_{label}")
                dump_all(page, f"after {label}")
                print("\n  Post-confirm text:")
                for line in page.inner_text("body").split('\n'):
                    l = line.strip()
                    if l and any(kw in l for kw in ['訂單','單號','Order','成功','完成',
                                                     '小票','收據','列印','Print','PDF']):
                        print(f"    {l!r}")
                break
        except: pass

    shot(page, "05_final")
    print(f"\nScreenshots saved in: {LOGS_DIR}")
    input("Press Enter to close...")
    ctx.close()
