# -*- coding: utf-8 -*-
"""Inspect backend cart: click product → cart → 結帳 → order number → receipt."""
import os, sys, time, re
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, POS_ADMIN_PASS, POS_VIP_PASS, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

def shot(page, name):
    p = os.path.join(LOGS_DIR, f"backend2_{name}.png")
    page.screenshot(path=p, full_page=False)
    print(f"  📸 {p}")

def dump_buttons(page):
    seen = set()
    for el in page.locator("button").all():
        try:
            t = el.inner_text().strip()
            if t and t not in seen and len(t) < 50 and \
               t not in ('🏪 前台商店','🔐 后台管理','🚚 送貨','📋 退貨','👤 Poon','登出'):
                seen.add(t)
                print(f"    {t!r}")
        except: pass

TEST_SKU = "1000043"  # 瓜拿那

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=200, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    # Login backend
    page.locator("button:has-text('后台管理')").first.click()
    time.sleep(1)
    page.locator("input[type='password']").first.fill(POS_ADMIN_PASS)
    page.keyboard.press("Enter")
    time.sleep(1.5)

    # Activate VIP
    page.locator("button:has-text('VIP價')").first.click()
    time.sleep(0.8)
    page.locator("input[type='password']").first.fill(POS_VIP_PASS)
    page.keyboard.press("Enter")
    time.sleep(1.5)
    print("Backend + VIP ready")

    # ── Click product by SKU 3 times (qty=3) ──────────────────────────────────
    print(f"\n=== Add {TEST_SKU} × 3 to cart ===")
    prod_btn = page.locator(f"button:has-text('{TEST_SKU}')").first
    print(f"  Product button text: {prod_btn.inner_text()!r}")
    for i in range(3):
        prod_btn.click()
        time.sleep(0.4)
        cart_text = page.locator("text=合計").first
        try:
            print(f"  After click {i+1}: cart area = {page.locator('text=合計').locator('..').inner_text()[:60]!r}")
        except: pass

    shot(page, "01_after_add")

    # Check cart area
    print("\n=== Cart area ===")
    cart_lines = []
    for line in page.inner_text("body").split('\n'):
        l = line.strip()
        if l and any(kw in l for kw in ['購物車','合計','$','小計','結帳',TEST_SKU,'瓜拿那','件','×','數量']):
            cart_lines.append(l)
            print(f"  {l!r}")

    # Check if cart has quantity controls
    print("\n  Buttons in cart area:")
    dump_buttons(page)

    # ── Click 結帳 (VIP) ───────────────────────────────────────────────────────
    print("\n=== Click 結帳 (VIP) ===")
    checkout_btn = page.locator("button:has-text('結帳')").first
    checkout_text = checkout_btn.inner_text()
    print(f"  Checkout button: {checkout_text!r}")
    checkout_btn.click()
    time.sleep(2)
    shot(page, "02_checkout")

    print("\n  Checkout screen buttons:")
    dump_buttons(page)

    print("\n  Checkout screen text (all):")
    for line in page.inner_text("body").split('\n'):
        l = line.strip()
        if l and l not in ('🏪 前台商店','🔐 后台管理','MANLEE','Health & Wellness',
                            '後台管理','全部','Activitae','ENERLAB','Monbélac','BELSNTE'):
            print(f"  {l!r}")

    # ── Inspect checkout inputs ────────────────────────────────────────────────
    print("\n  Checkout inputs:")
    for el in page.locator("input, textarea, select").all():
        try:
            t  = el.get_attribute("type") or ""
            ph = el.get_attribute("placeholder") or ""
            v  = "" if t == "password" else el.input_value()
            print(f"    type={t!r} ph={ph!r} val={v[:40]!r}")
        except: pass

    # ── Try to confirm / complete order ───────────────────────────────────────
    print("\n=== Try confirm order ===")
    for label in ['確認', '完成', '下單', '確認訂單', '現金', '確認付款']:
        b = page.locator(f"button:has-text('{label}')").first
        try:
            if b.is_visible(timeout=500):
                print(f"  Clicking: {label!r}")
                b.click()
                time.sleep(1.5)
                shot(page, f"03_after_{label}")
                print(f"  After {label!r}:")
                for line in page.inner_text("body").split('\n'):
                    l = line.strip()
                    if l and any(kw in l for kw in ['訂單','單號','#','Order','成功',
                                                     '小票','收據','列印','Print','完成']):
                        print(f"    {l!r}")
                dump_buttons(page)
                break
        except: pass

    shot(page, "04_final")
    print(f"\nScreenshots: {LOGS_DIR}")
    input("Press Enter to close...")
    ctx.close()
