# -*- coding: utf-8 -*-
"""Check receipt content and order number after 確認，出小票."""
import os, sys, time, re
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import POS_URL, POS_ADMIN_PASS, POS_VIP_PASS, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

def shot(page, name):
    p = os.path.join(LOGS_DIR, f"receipt_{name}.png")
    page.screenshot(path=p, full_page=False)
    print(f"  📸 {p}")

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=200, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    # Backend + VIP
    page.locator("button:has-text('后台管理')").first.click(); time.sleep(1)
    page.locator("input[type='password']").first.fill(POS_ADMIN_PASS)
    page.keyboard.press("Enter"); time.sleep(1.5)
    page.locator("button:has-text('VIP價')").first.click(); time.sleep(0.8)
    page.locator("input[type='password']").first.fill(POS_VIP_PASS)
    page.keyboard.press("Enter"); time.sleep(1.5)
    print("Ready")

    # Add 1000043 × 2
    btn = page.locator("button:has-text('1000043')").first
    btn.click(); time.sleep(0.4)
    btn.click(); time.sleep(0.4)

    # Checkout
    page.locator("button:has-text('結帳')").first.click(); time.sleep(1.5)
    shot(page, "01_checkout_screen")

    # Select 現金
    cash = page.locator("button:has-text('現金')").first
    if cash.is_visible(timeout=2000):
        cash.click(); time.sleep(0.5)
        print("Selected 現金")

    shot(page, "02_payment_selected")

    # Click 確認，出小票
    confirm = page.locator("button:has-text('確認，出小票')").first
    if confirm.is_visible(timeout=2000):
        confirm.click(); time.sleep(2)
        shot(page, "03_after_confirm")
        print("Clicked 確認，出小票")

    # ── Read FULL receipt text ─────────────────────────────────────────────────
    print("\n=== FULL receipt page text ===")
    full = page.inner_text("body")
    for line in full.split('\n'):
        l = line.strip()
        if l: print(f"  {l!r}")

    # Look for order number patterns
    print("\n=== Order number candidates ===")
    for pattern in [r'#\d+', r'訂單[號碼編]?\s*[：:]\s*(\w+)',
                     r'Order\s*#?\s*(\w+)', r'\bORD\w+', r'\b\d{6,}\b']:
        matches = re.findall(pattern, full)
        if matches:
            print(f"  Pattern {pattern!r}: {matches[:5]}")

    shot(page, "04_receipt_full")

    # ── Try to print (see what happens) ───────────────────────────────────────
    print("\n=== 列印 button ===")
    print_btn = page.locator("button:has-text('列印')").first
    if print_btn.is_visible(timeout=2000):
        print(f"  Found 列印 button: {print_btn.inner_text()!r}")
        # Check if page has printable content
        # Use CDP to capture PDF without showing dialog
        try:
            pdf_bytes = page.pdf(format="A5", print_background=True)
            test_pdf = os.path.join(LOGS_DIR, "test_receipt.pdf")
            with open(test_pdf, "wb") as f:
                f.write(pdf_bytes)
            print(f"  PDF saved via CDP: {test_pdf}")
        except Exception as e:
            print(f"  CDP PDF failed: {e}")
            print("  (Would need to use print dialog + pyautogui as fallback)")

    # ── Click 完成 ────────────────────────────────────────────────────────────
    done_btn = page.locator("button:has-text('完成')").first
    if done_btn.is_visible(timeout=2000):
        print(f"\n完成 button found: {done_btn.inner_text()!r}")

    print(f"\nScreenshots: {LOGS_DIR}")
    input("Press Enter to close...")
    ctx.close()
