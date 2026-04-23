# -*- coding: utf-8 -*-
"""
POS automation — backend cashier flow:
  后台管理 → 0000 → VIP價 → 941196 →
  click product(SKU)×qty → 結帳(VIP) → 現金 → 確認，出小票 →
  save PDF → 完成
"""
import os
import re
import sys
import time
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

from config import (
    POS_URL, POS_ADMIN_PASS, POS_VIP_PASS,
    CHROME_PROFILE, BROWSER_ARGS, PLAYWRIGHT_SLOW_MO, IMAGES_DIR,
)
from logger import get_logger, screenshot_on_error, toast_error

log = get_logger(__name__)


def _clear_chrome_locks() -> None:
    for lf in ["lockfile", "SingletonLock", "SingletonSocket", "SingletonCookie"]:
        try:
            os.remove(os.path.join(CHROME_PROFILE, lf))
        except Exception:
            pass


def run_pos_checkout(order: dict) -> tuple[str, str]:
    """
    Full POS backend checkout.
    Returns (pos_order_no, pdf_path).
    On failure: screenshot → log → toast → raise (browser stays open).
    """
    _clear_chrome_locks()
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            CHROME_PROFILE, channel="chrome", headless=False,
            args=BROWSER_ARGS, slow_mo=PLAYWRIGHT_SLOW_MO,
            viewport={"width": 1280, "height": 900},
        )
        page = ctx.new_page()
        try:
            page.goto(POS_URL, wait_until="domcontentloaded", timeout=20_000)
            time.sleep(2)

            _enter_backend(page)
            _activate_vip(page)
            _add_items(page, order["items"])
            _checkout(page)
            _select_payment_cash(page)
            _confirm_and_print_screen(page)

            order_no = _extract_order_no(page)
            pdf_path = _save_pdf(page, order, order_no)

            _click_done(page)
            log.info("POS done — order %s, pdf %s", order_no, pdf_path)
            ctx.close()
            return order_no, pdf_path

        except Exception as e:
            path = screenshot_on_error(page, "pos_automation")
            toast_error("POS 下單", str(e)[:120])
            log.exception("POS failed (screenshot: %s)", path)
            # Keep browser open for manual recovery
            raise


# ─── Step functions ───────────────────────────────────────────────────────────

def _enter_backend(page: Page) -> None:
    page.locator("button:has-text('后台管理')").first.click()
    time.sleep(0.8)
    pwd = page.locator("input[type='password'][placeholder='輸入密碼']").first
    pwd.wait_for(state="visible", timeout=5000)
    pwd.fill(POS_ADMIN_PASS)
    page.keyboard.press("Enter")
    time.sleep(1.5)
    log.debug("Backend entered")


def _activate_vip(page: Page) -> None:
    page.locator("button:has-text('VIP價')").first.click()
    time.sleep(0.8)
    pwd = page.locator("input[type='password'][placeholder='輸入密碼']").first
    pwd.wait_for(state="visible", timeout=5000)
    pwd.fill(POS_VIP_PASS)
    page.keyboard.press("Enter")
    time.sleep(1.5)
    log.debug("VIP activated")


def _add_items(page: Page, items: list) -> None:
    """Click each product button qty times. Each click = +1 unit."""
    for item in items:
        sku = item["sku"]
        qty = item["qty"]
        # Product buttons contain SKU in their text
        btn = page.locator(f"button:has-text('{sku}')").first
        try:
            btn.wait_for(state="visible", timeout=5000)
        except PWTimeout:
            raise ValueError(f"Product button for SKU {sku!r} not found in POS")

        log.debug("Adding %s × %d", sku, qty)
        for _ in range(qty):
            btn.click()
            time.sleep(0.25)

        # Verify cart total updated
        try:
            total_text = page.locator("text=合計").locator("..").inner_text(timeout=2000)
            log.debug("  Cart after %s: %s", sku, total_text[:40])
        except Exception:
            pass


def _checkout(page: Page) -> None:
    checkout_btn = page.locator("button:has-text('結帳')").first
    checkout_btn.wait_for(state="visible", timeout=5000)
    checkout_btn.click()
    time.sleep(1.5)
    log.debug("Checkout screen opened")


def _select_payment_cash(page: Page) -> None:
    cash_btn = page.locator("button:has-text('現金')").first
    cash_btn.wait_for(state="visible", timeout=5000)
    cash_btn.click()
    time.sleep(0.5)
    log.debug("Payment: 現金 selected")


def _confirm_and_print_screen(page: Page) -> None:
    confirm_btn = page.locator("button:has-text('確認，出小票')").first
    confirm_btn.wait_for(state="visible", timeout=5000)
    confirm_btn.click()
    time.sleep(2)
    log.debug("Receipt screen shown")


def _extract_order_no(page: Page) -> str:
    """Extract ORD-XXXXXX from receipt page."""
    text = page.inner_text("body")
    m = re.search(r"ORD-\d+", text)
    if m:
        order_no = m.group(0)
        log.info("Order number: %s", order_no)
        return order_no
    # Fallback: generate from date
    fallback = f"ORD-{date.today():%Y%m%d}"
    log.warning("Order number not found, using fallback: %s", fallback)
    return fallback


def _save_pdf(page: Page, order: dict, order_no: str) -> str:
    """Use CDP page.pdf() to save receipt without print dialog."""
    name = order["name"]
    today = date.today().strftime("%Y%m%d")
    # Prefer existing folder (original name), fall back to simplified
    person_dir = os.path.join(IMAGES_DIR, name)
    if not os.path.isdir(person_dir):
        simplified = order.get("name_simplified", name)
        alt_dir = os.path.join(IMAGES_DIR, simplified)
        if os.path.isdir(alt_dir):
            person_dir = alt_dir
    os.makedirs(person_dir, exist_ok=True)

    base = f"{today}_{name}_{order_no}"
    pdf_path = _unique_path(person_dir, base)

    try:
        page.pdf(
            path=pdf_path,
            format="A5",
            print_background=True,
            margin={"top": "10mm", "bottom": "10mm",
                    "left": "8mm", "right": "8mm"},
        )
        log.info("PDF saved: %s", pdf_path)
        return pdf_path
    except Exception as e:
        log.warning("CDP PDF failed (%s) — clicking 列印 as fallback", e)
        # Fallback: click 列印 button (opens browser print dialog)
        try:
            page.locator("button:has-text('列印')").first.click()
            time.sleep(2)
        except Exception:
            pass
        return pdf_path  # Return expected path (may not exist yet)


def _click_done(page: Page) -> None:
    try:
        done = page.locator("button:has-text('完成')").first
        if done.is_visible(timeout=3000):
            done.click()
            time.sleep(0.8)
    except Exception:
        pass


def _unique_path(directory: str, base: str) -> str:
    candidate = os.path.join(directory, f"{base}.pdf")
    if not os.path.exists(candidate):
        return candidate
    for i in range(1, 100):
        candidate = os.path.join(directory, f"{base}_{i:02d}.pdf")
        if not os.path.exists(candidate):
            return candidate
    return os.path.join(directory, f"{base}_dup.pdf")


# ─── Standalone test ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    from order_parser import parse_order
    sample = "1084043-3件，1084065-4件，1000044-2件，寄黃业偉 18125989028 廣東省深圳市龍崗區南聯劉屋村南段74號1樓"
    order = parse_order(sample)
    print("Parsed order:", order["name"], order["items"])
    order_no, pdf = run_pos_checkout(order)
    print("Done:", order_no, pdf)
