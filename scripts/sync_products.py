# -*- coding: utf-8 -*-
"""
sync_products.py — 自動從 POS 網站抓取產品資料，更新 products.json
執行後會登入後台 → 啟動 VIP 價 → 等你瀏覽所有產品 → 掃描 → 更新 products.json
"""
import os
import sys
import json
import time

sys.stdout.reconfigure(encoding="utf-8")
from playwright.sync_api import sync_playwright

CHROME_PROFILE  = r"C:\ChromeAutomation"
POS_URL         = "https://online-store-99126206.web.app/"
POS_PASS        = "0000"
VIP_PASS        = "941196"
PRODUCTS_JSON   = r"C:\Users\user\Desktop\順丰E順递\data\products.json"
LOGS_DIR        = r"C:\Users\user\Desktop\順丰E順递\logs"

os.makedirs(LOGS_DIR, exist_ok=True)

SCAN_JS = """() => {
    const results = [];
    const seen = new Set();

    // 掃描所有元素（包括唔可見的），搵含有7位SKU的最小元素
    for (const el of document.querySelectorAll('*')) {
        const text = (el.textContent || '').trim();

        // 搵 7-10 位 SKU
        const skuMatch = text.match(/\\b(\\d{7,10})\\b/);
        if (!skuMatch) continue;
        const sku = skuMatch[1];
        if (seen.has(sku)) continue;

        // 唔要太大的容器（可能係整頁）
        if (text.length > 800) continue;

        // 搵所有 $ 價格，最後一個係 VIP 價
        const allPrices = [];
        const priceRe = /\\$\\s*([\\d.]+)/g;
        let pm;
        while ((pm = priceRe.exec(text)) !== null) {
            const val = parseFloat(pm[1]);
            if (val > 0 && val < 5000) allPrices.push(val);
        }
        if (allPrices.length === 0) continue;  // 冇價格就跳過

        const vipPrice = allPrices[allPrices.length - 1];

        seen.add(sku);
        results.push({
            sku,
            vipPrice,
            rawText: text.replace(/\\s+/g, ' ').substring(0, 120),
        });
    }
    return results;
}"""

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=["--disable-blink-features=AutomationControlled"],
        slow_mo=100, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()

    print("▶ 開啟 POS 網站...")
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    print("▶ 登入後台管理...")
    page.locator("button:has-text('后台管理')").first.click()
    time.sleep(0.8)
    page.locator("input[type='password']").first.fill(POS_PASS)
    page.keyboard.press("Enter")
    time.sleep(2)

    print("▶ 啟動 VIP 價...")
    page.locator("button:has-text('VIP價')").first.click()
    time.sleep(0.8)
    page.locator("input[type='password']").first.fill(VIP_PASS)
    page.keyboard.press("Enter")
    time.sleep(2)

    print("\n" + "="*60)
    print("  請喺瀏覽器裡面，手動瀏覽所有產品類別，")
    print("  確保所有產品都曾出現過，再返嚟按 Enter 掃描。")
    print("="*60)
    input("  按 Enter 開始掃描所有產品...")

    print("\n▶ 掃描產品資料...")
    products = page.evaluate(SCAN_JS)
    print(f"  共找到 {len(products)} 個產品")

    if not products:
        print("\n⚠️  未能抓取任何產品，請確認 VIP 價已啟動")
    else:
        print("\n找到以下產品：")
        print("-" * 70)
        for p in products:
            print(f"  SKU: {p['sku']}  VIP價: ${p['vipPrice']}")
            print(f"       {p['rawText']}")
            print()

        # 載入現有 products.json
        existing = {}
        if os.path.exists(PRODUCTS_JSON):
            with open(PRODUCTS_JSON, encoding="utf-8") as f:
                existing = json.load(f)

        updated = 0
        matched = 0
        not_in_json = []

        for p in products:
            sku = p["sku"]
            if sku in existing:
                matched += 1
                old_price = existing[sku].get("vip_price", 0)
                if round(old_price, 1) != round(p["vipPrice"], 1):
                    print(f"  🔄 更新 {sku}  {existing[sku].get('name','')[:20]}  ${old_price} → ${p['vipPrice']}")
                    existing[sku]["vip_price"] = p["vipPrice"]
                    updated += 1
            else:
                not_in_json.append(sku)

        print(f"\n{'='*60}")
        if updated:
            with open(PRODUCTS_JSON, "w", encoding="utf-8") as f:
                json.dump(existing, f, ensure_ascii=False, indent=2)
            print(f"✅ 已更新 {updated} 個產品 VIP 價")
        else:
            print(f"✅ 已核對 {matched} 個產品，全部 VIP 價正確，無需更新")

        if not_in_json:
            print(f"\n⚠️  以下 SKU 喺網站有但 products.json 冇，需要手動新增：")
            for sku in not_in_json:
                print(f"     {sku}")
        print("="*60)

    input("\n按 Enter 關閉瀏覽器...")
    ctx.close()
