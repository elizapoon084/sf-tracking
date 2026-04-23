# -*- coding: utf-8 -*-
"""Inspect SF Express ship form — find 智慧填寫 + 識明 selectors."""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import SF_SHIP_URL, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

def shot(page, name):
    p = os.path.join(LOGS_DIR, f"sf_{name}.png")
    page.screenshot(path=p, full_page=False)
    print(f"  📸 {p}")

def dump_buttons(page):
    seen = set()
    for el in page.locator("button, a[role='button'], span[role='button']").all():
        try:
            t = el.inner_text().strip()
            if t and t not in seen and len(t) < 60:
                seen.add(t)
                print(f"    btn: {t!r}")
        except: pass

def dump_inputs(page):
    for el in page.locator("input, textarea").all():
        try:
            t  = el.get_attribute("type") or "text"
            ph = el.get_attribute("placeholder") or ""
            nm = el.get_attribute("name") or ""
            cl = el.get_attribute("class") or ""
            if t != "hidden":
                print(f"    input type={t!r} ph={ph!r} name={nm!r} class={cl[:40]!r}")
        except: pass

TEST_TEXT = "黄业伟 18125989028 廣東省深圳市龍崗區南聯劉屋村南段74號1樓"

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=200, viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(SF_SHIP_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)
    shot(page, "01_landing")

    print("\n=== Initial buttons ===")
    dump_buttons(page)
    print("\n=== Initial inputs ===")
    dump_inputs(page)

    # Look for 智慧填寫 button
    print("\n=== Looking for 智慧填寫 ===")
    for kw in ['智慧', '智能', 'Smart', '識明', '智慧填寫']:
        els = page.locator(f"button:has-text('{kw}'), a:has-text('{kw}'), "
                           f"span:has-text('{kw}')").all()
        if els:
            print(f"  FOUND {kw!r}: {len(els)} elements")
            for el in els[:3]:
                try:
                    print(f"    text={el.inner_text()!r}  tag={el.evaluate('e=>e.tagName')}")
                except: pass

    # Click RECIPIENT 智慧填寫 directly (index 1, 收 side)
    print("\n=== Click RECIPIENT 智慧填寫 (index 1, 收 side) ===")
    smart_span = page.locator("span:has-text('智慧填寫')").nth(1)
    smart_span.wait_for(state="visible", timeout=5000)
    smart_span.click()
    time.sleep(1)
    shot(page, "02_recipient_dialog")

    # Show all buttons inside dialog
    print("  Buttons now visible:")
    dump_buttons(page)

    # Fill the intel-address textarea
    print("\n=== Fill smart textarea ===")
    smart_ta = page.locator(
        "textarea[class*='intelAddr'], textarea[placeholder*='陳先生']"
    ).first
    try:
        smart_ta.wait_for(state="visible", timeout=4000)
        smart_ta.click()
        # Use type() to trigger proper React input events
        smart_ta.type(TEST_TEXT, delay=30)
        time.sleep(0.8)
        shot(page, "03_filled")
        print(f"  Typed: {TEST_TEXT!r}")

        # JS click 識別 (it's a SPAN not a button)
        print("\n  JS clicking 識別...")
        page.evaluate("""() => {
            for (const el of document.querySelectorAll('*')) {
                if (el.childNodes.length===1 && el.firstChild.nodeType===3
                    && el.firstChild.textContent.trim()==='識別') {
                    el.click(); return;
                }
            }
        }""")
        time.sleep(2.5)
        shot(page, "04_after_identify")

        # Dump field values after 識別
        print("\n  Input values after 識別:")
        for el in page.locator("input, textarea").all():
            try:
                t  = el.get_attribute("type") or "text"
                ph = el.get_attribute("placeholder") or ""
                v  = el.input_value() if t != "password" else ""
                if v: print(f"    ph={ph!r}  val={v!r}")
            except: pass

    except Exception as e:
        print(f"  Error: {e}")

# ── Step 2: Select 自寄 ────────────────────────────────────────────────────────
print("\n=== Select 自寄 ===")
for kw in ['自寄', '自行送']:
    btn = page.locator(f"label:has-text('{kw}'), button:has-text('{kw}'), "
                       f"span:has-text('{kw}')").first
    try:
        if btn.is_visible(timeout=2000):
            print(f"  Clicking {kw!r}")
            btn.click()
            time.sleep(0.8)
            shot(page, "05_ziji")
            print("  Selected 自寄")
            break
    except: pass

print("  Buttons after 自寄 selection:")
dump_buttons(page)

# ── Step 3: Click 下一步 ───────────────────────────────────────────────────────
print("\n=== Click 下一步 ===")
for kw in ['下一步', '繼續', 'Next', '確認']:
    btn = page.locator(f"button:has-text('{kw}')").first
    try:
        if btn.is_visible(timeout=2000):
            print(f"  Clicking {kw!r}")
            btn.click()
            time.sleep(2)
            shot(page, f"06_after_{kw}")
            break
    except: pass

# ── Step 4: 報關資料 page ──────────────────────────────────────────────────────
print("\n=== 報關資料 page ===")
print("  Page text:")
for line in page.inner_text("body").split('\n'):
    l = line.strip()
    if l and len(l) < 80 and l not in ('MANLEE','Health & Wellness'):
        print(f"  {l!r}")
print("\n  Buttons:")
dump_buttons(page)
print("\n  Inputs:")
dump_inputs(page)
shot(page, "07_customs")
shot(page, "08_final")
print(f"\nDone. Screenshots: {LOGS_DIR}")
input("Press Enter to close...")
ctx.close()
