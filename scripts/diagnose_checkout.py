# -*- coding: utf-8 -*-
"""
diagnose_checkout.py
逐個 SKU 加入購物車 → 試 checkout → 找出令橙色按鈕失效的 SKU
"""
import sys, time, json
sys.stdout.reconfigure(encoding='utf-8')
from playwright.sync_api import sync_playwright

CHROME_PROFILE = r"C:\ChromeAutomation"
POS_URL  = "https://online-store-99126206.web.app/"
POS_PASS = "0000"
VIP_PASS = "941196"

# 方遜訂單的貨品（逐個加入測試）
TEST_SKUS = [
    '1084085',  # Activitae 長青護心高鈣脫脂奶粉
    '0300442',  # SENBEL 再注氧潔面啫喱
    '0300530',  # SENBEL 柔和泡沫潔面乳
    '0300408',  # SENBEL 緊緻精華面膜
    '1084065',  # Activitae 女士寶第三代升級版
    '1000044',  # Activitae 蟻木/樟木
    '1084067',  # Activitae 雄風寶加強版
    '1000458',  # Activitae 螺旋藻
    '0300409',  # SENBEL 淨膚精華面膜
]

def can_click_confirm(page) -> bool:
    """試 click 結帳 → 返回橙色按鈕係咪可以按"""
    try:
        page.locator("button:has-text('結帳')").first.click()
        time.sleep(2)
        btn = page.locator("button:has-text('確認，出小票')").first
        btn.wait_for(state="visible", timeout=5000)
        # 滾到底
        page.evaluate("""() => {
            document.querySelectorAll('*').forEach(el => {
                if (el.scrollHeight > el.clientHeight + 5 &&
                    getComputedStyle(el).overflowY !== 'visible' &&
                    el.clientHeight > 50 && el.clientHeight < 500)
                    el.scrollTop = el.scrollHeight;
            });
        }""")
        time.sleep(0.5)
        # 試 click 並看 console errors
        errors = []
        page.on('console', lambda msg: errors.append(msg.text) if msg.type == 'error' else None)
        btn.click(force=True)
        time.sleep(1.5)
        # 橙色按鈕消失 = checkout 成功
        still_showing = btn.is_visible(timeout=500)
        # 返回結帳頁
        page.locator("button:has-text('返回')").first.click(force=True) if still_showing else None
        time.sleep(1)
        return not still_showing, errors
    except Exception as e:
        print(f"    [ERR] {e}")
        return False, []

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=["--disable-blink-features=AutomationControlled",
              "--disable-infobars", "--disable-session-crashed-bubble"],
        slow_mo=100, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    # 登入 + VIP
    page.locator("button:has-text('后台管理')").first.click(); time.sleep(0.8)
    page.locator("input[type='password']").first.fill(POS_PASS); page.keyboard.press("Enter"); time.sleep(1.5)
    page.locator("button:has-text('VIP價')").first.click(); time.sleep(0.8)
    page.locator("input[type='password']").first.fill(VIP_PASS); page.keyboard.press("Enter"); time.sleep(2)
    print("✅ 登入完成\n")

    bad_sku = None
    for i, sku in enumerate(TEST_SKUS):
        print(f"▶ 加入第 {i+1} 個 SKU: {sku}")
        try:
            btn = page.locator(f"button:has-text('{sku}')").first
            btn.wait_for(state="visible", timeout=10000)
            btn.click()
            time.sleep(0.5)
        except Exception as e:
            print(f"  [跳過] {sku} 搵唔到: {e}")
            continue

        ok, errors = can_click_confirm(page)
        status = "✅ 橙色按鈕正常" if ok else "❌ 橙色按鈕失效"
        print(f"  {status}")
        if errors:
            print(f"  Console errors: {errors[:3]}")
        if not ok:
            bad_sku = sku
            print(f"\n🎯 第 {i+1} 個 SKU [{sku}] 加入後出問題！")
            break
        print()

    if bad_sku:
        print(f"\n結論：問題出自 SKU [{bad_sku}]")
    else:
        print("\n結論：所有 SKU 均正常，問題可能係 item 總數/總額")

    input("\n按 Enter 結束...")
    ctx.close()
