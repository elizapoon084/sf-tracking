# -*- coding: utf-8 -*-
"""
Scrape all products from the POS and cache to products.json.

Strategy: all product data is pre-loaded in the DOM — parse in one pass.
No need to click modals.

Run standalone:  python product_scraper.py
Or call:         scrape_products(force_refresh=True)
"""
import json
import os
import re
import sys
import time
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright, Page

from config import (
    POS_URL, POS_VIP_PASS, CHROME_PROFILE, BROWSER_ARGS,
    PLAYWRIGHT_SLOW_MO, PRODUCTS_JSON,
)
from logger import get_logger, screenshot_on_error, toast_error

log = get_logger(__name__)
_CACHE_TTL_HOURS = 24


# ─── Public API ───────────────────────────────────────────────────────────────

def load_products() -> dict:
    if os.path.exists(PRODUCTS_JSON):
        with open(PRODUCTS_JSON, encoding='utf-8') as f:
            return json.load(f)
    return {}


def scrape_products(force_refresh: bool = False) -> dict:
    if not force_refresh and _cache_fresh():
        log.info("products.json cache is fresh — skipping scrape")
        return load_products()

    log.info("Scraping products from POS…")
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            CHROME_PROFILE, channel="chrome", headless=False,
            args=BROWSER_ARGS, slow_mo=PLAYWRIGHT_SLOW_MO,
            viewport={"width": 1280, "height": 900},
        )
        page = ctx.new_page()
        try:
            page.goto(POS_URL, wait_until="domcontentloaded", timeout=20_000)
            time.sleep(3)

            _activate_vip(page)

            # Extract all text + prices in one pass
            full_text = page.inner_text("body")
            vip_prices = _get_vip_prices_from_dom(page)

            products = _parse_products(full_text, vip_prices)
            _save(products)
            log.info("Scraped %d products → %s", len(products), PRODUCTS_JSON)
            return products

        except Exception as e:
            path = screenshot_on_error(page, "product_scraper")
            toast_error("Product Scraper", str(e))
            log.exception("Product scrape failed (screenshot: %s)", path)
            raise
        finally:
            ctx.close()


# ─── VIP activation ───────────────────────────────────────────────────────────

def _activate_vip(page: Page) -> None:
    try:
        page.locator("button:has-text('VIP')").first.click()
        time.sleep(0.8)
        pwd = page.locator("input[type='password'][placeholder='輸入密碼']").first
        pwd.wait_for(state="visible", timeout=4000)
        pwd.fill(POS_VIP_PASS)
        page.keyboard.press("Enter")
        time.sleep(1.5)
        log.info("VIP mode activated")
    except Exception as e:
        log.warning("VIP activation skipped: %s", e)


# ─── Price extraction from DOM ────────────────────────────────────────────────

def _get_vip_prices_from_dom(page: Page) -> list:
    """
    Extract all HK$X prices from page text in order.
    Prices appear as pairs (VIP price, original price) for each product.
    Returns list where index i = VIP price for product i.
    """
    try:
        full_text = page.inner_text("body")
        # Find all HK$XX values in the order they appear
        all_prices = re.findall(r'HK\$(\d+(?:\.\d+)?)', full_text)
        raw = [float(p) for p in all_prices]

        # Prices come in pairs: (VIP, original). VIP is always lower.
        # Group into pairs and take the smaller (VIP) value.
        vip = []
        i = 0
        while i < len(raw) - 1:
            a, b = raw[i], raw[i + 1]
            # VIP price is the smaller one
            vip.append(min(a, b))
            i += 2
        # Handle odd trailing price
        if i < len(raw):
            vip.append(raw[i])

        log.info("Collected %d VIP prices from page text", len(vip))
        return vip
    except Exception as e:
        log.warning("Price extraction failed: %s", e)
        return []


# ─── Text parser ──────────────────────────────────────────────────────────────

def _parse_products(full_text: str, vip_prices: list) -> dict:
    """
    Parse all product blocks from full page text.

    Each product block:
      [SKU]            ← 6-7 digit line
      [description]    ← may span many lines
      品牌：X 材質：X 規格型號：X  ← always at end of description
    """
    products: dict = {}

    # Split into lines; find lines that are pure SKUs
    lines = full_text.split('\n')

    sku_indices = []  # (line_index, sku)
    for i, line in enumerate(lines):
        s = line.strip()
        if re.match(r'^\d{6,7}$', s):
            sku_indices.append((i, s))

    log.info("Found %d potential SKU lines", len(sku_indices))

    for pos, (line_idx, sku) in enumerate(sku_indices):
        # Collect lines until next SKU or end
        if pos + 1 < len(sku_indices):
            next_idx = sku_indices[pos + 1][0]
        else:
            next_idx = len(lines)

        block = '\n'.join(lines[line_idx + 1: next_idx])

        name      = _extract_name(block)
        brand     = _extract_field(block, r'品牌[：:]\s*(.+?)(?:\s+材質|$|\n)')
        material  = _extract_field(block, r'材質[：:]\s*(.+?)(?:\s+規格|品牌|$)')
        spec      = _extract_field(block, r'規格型號[：:]\s*(.+?)(?:\s+品牌|$|\n)')

        vip_price = vip_prices[pos] if pos < len(vip_prices) else 0.0

        # Skip navigation/UI text blocks
        if not material and not spec and not name:
            continue

        products[sku] = {
            "sku":       sku,
            "name":      name[:60],
            "brand":     brand,
            "material":  material,
            "spec":      spec,
            "origin":    "台灣",
            "vip_price": vip_price,
        }

    return products


def _extract_name(block: str) -> str:
    """First meaningful line in the block that isn't metadata."""
    skip = re.compile(
        r'HK\$|原價|品牌：|材質：|規格|加入|前台|后台|VIP|購物車|搜尋|#', re.I
    )
    for line in block.split('\n'):
        line = line.strip()
        if not line or skip.search(line):
            continue
        if re.match(r'^\d+$', line):
            continue
        return line
    return ""


def _extract_field(text: str, pattern: str) -> str:
    m = re.search(pattern, text, re.DOTALL)
    if not m:
        return ""
    return m.group(1).strip().split('\n')[0].strip()


# ─── Cache helpers ────────────────────────────────────────────────────────────

def _cache_fresh() -> bool:
    if not os.path.exists(PRODUCTS_JSON):
        return False
    age = (datetime.now().timestamp() - os.path.getmtime(PRODUCTS_JSON)) / 3600
    return age < _CACHE_TTL_HOURS


def _save(products: dict) -> None:
    os.makedirs(os.path.dirname(PRODUCTS_JSON), exist_ok=True)
    with open(PRODUCTS_JSON, "w", encoding="utf-8") as f:
        json.dump(products, f, ensure_ascii=False, indent=2)


# ─── Standalone ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    result = scrape_products(force_refresh=True)
    print(f"\nTotal: {len(result)} products saved to {PRODUCTS_JSON}")
    for sku, p in list(result.items())[:5]:
        print(f"  {sku}: {p['name']} | 材質:{p['material']} | 規格:{p['spec']} | HKD{p['vip_price']}")
