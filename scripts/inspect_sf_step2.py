# -*- coding: utf-8 -*-
"""SF inspect v2 — 全自動版,由頭到尾跑一次"""
import os, sys, time, traceback
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import SF_SHIP_URL, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

for lf in ["lockfile", "SingletonLock", "SingletonSocket", "SingletonCookie"]:
    p = os.path.join(CHROME_PROFILE, lf)
    try:
        os.remove(p)
    except Exception:
        pass

SENDER_TEXT = "潘正儀 66832382 香港九龍新蒲崗大有街33號佳力工業大廈603室"
RECIP_TEXT  = "黄业伟 18125989028 廣東省深圳市龍崗區南聯劉屋村南段74號1樓"


def shot(page, name):
    p = os.path.join(LOGS_DIR, f"sf2_{name}.png")
    page.screenshot(path=p, full_page=False)
    print(f"  📸 {p}")


def step(title):
    print(f"\n{'='*60}\n  {title}\n{'='*60}")


def react_fill(page, selector_index, value):
    """填 React controlled input,觸發 onChange"""
    page.evaluate(f"""() => {{
        const inputs = document.querySelectorAll("input[name='contactName']");
        const input = inputs[{selector_index}];
        if (!input) return;
        const setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, {repr(value)});
        input.dispatchEvent(new Event('input',  {{ bubbles: true }}));
        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
        input.blur();
    }}""")


def click_identify(page):
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.childNodes.length===1 && el.firstChild.nodeType===3
                && el.firstChild.textContent.trim()==='識別') { el.click(); return; }
        }
    }""")


def main():
    with sync_playwright() as pw:
        try:
            ctx = pw.chromium.launch_persistent_context(
                CHROME_PROFILE, channel="chrome", headless=False,
                args=BROWSER_ARGS, slow_mo=150,
                viewport={"width": 1280, "height": 900},
            )
        except Exception as e:
            print(f"❌ Chrome 開唔到: {e}")
            return

        page = ctx.new_page()

        try:
            # ── Step 0: 開頁 ─────────────────────────────────────
            step("Step 0 — 開順丰寄件頁")
            page.goto(SF_SHIP_URL, wait_until="domcontentloaded", timeout=20000)
            time.sleep(3)
            shot(page, "00_landing")
            print("✅ 頁面載入")

            # ── Step 1: 寄件人智慧填寫 ───────────────────────────
            step("Step 1 — 寄件人智慧填寫")
            page.locator("span:has-text('智慧填寫')").nth(0).click()
            time.sleep(0.8)
            ta = page.locator("textarea[class*='intelAddr'], textarea[placeholder*='陳先生']").first
            ta.wait_for(state="visible", timeout=5000)
            ta.fill(SENDER_TEXT)
            time.sleep(0.5)
            click_identify(page)
            time.sleep(2.5)
            shot(page, "01_sender_done")
            print("✅ 寄件人識別完成")
            time.sleep(3)

            # ── Step 2: 收件人智慧填寫 ───────────────────────────
            step("Step 2 — 收件人智慧填寫")
            page.locator("span:has-text('智慧填寫')").nth(1).click()
            time.sleep(0.8)
            ta = page.locator("textarea[class*='intelAddr'], textarea[placeholder*='陳先生']").first
            ta.wait_for(state="visible", timeout=5000)
            ta.fill(RECIP_TEXT)
            time.sleep(0.5)
            click_identify(page)
            time.sleep(2.5)
            shot(page, "02_recip_done")
            print("✅ 收件人識別完成")

            # ── Step 2b: 補填寄件人姓名 ──────────────────────────
            step("Step 2b — 補填寄件人姓名（順丰 bug）")
            react_fill(page, 0, "潘正儀")
            time.sleep(0.5)
            shot(page, "02b_all_filled")
            print("✅ 寄件人姓名補填完成")
            time.sleep(1)

            # ── Step 3: Click 自寄 ───────────────────────────────
            step("Step 3 — Click 自寄")
            clicked = page.evaluate("""() => {
                for (const el of document.querySelectorAll('*')) {
                    if (el.textContent.trim()==='自寄' && el.offsetParent!==null) {
                        const r = el.getBoundingClientRect();
                        if (r.width>40 && r.height>20) { el.click(); return true; }
                    }
                }
                return false;
            }""")
            print(f"  自寄 clicked: {clicked}")
            time.sleep(1.5)
            shot(page, "03_after_ziji")

            # ── Step 4: 開選自寄點 ───────────────────────────────
            step("Step 4 — 開選自寄點")
            page.evaluate("""() => {
                for (const el of document.querySelectorAll('*')) {
                    if (el.textContent.trim()==='選自寄點' && el.offsetParent!==null) {
                        el.click(); return;
                    }
                }
            }""")
            time.sleep(2)
            shot(page, "04_station_picker")

            # ── Step 5: 揀第一個自寄站 ──────────────────────────
            step("Step 5 — 揀第一個自寄站")
            station = page.evaluate("""() => {
                // try <li> rows first
                for (const el of document.querySelectorAll('li')) {
                    if (el.offsetParent !== null) {
                        const r = el.getBoundingClientRect();
                        if (r.width > 100 && r.height > 20 && r.height < 120) {
                            el.click();
                            return {tag: 'LI', text: el.textContent.trim().slice(0,60)};
                        }
                    }
                }
                return null;
            }""")
            print(f"  Station clicked: {station}")
            time.sleep(1.5)
            shot(page, "05_station_selected")

            # ── Step 6: Click +新增物品 via Playwright locator ──────────────────
            step("Step 6 — Click +新增物品")
            time.sleep(1)
            add_btn = page.locator("button:has-text('新增物品')").first
            add_btn.scroll_into_view_if_needed(timeout=8000)
            time.sleep(0.5)
            shot(page, "06_before_add")
            add_btn.click(timeout=8000)
            print("✅ Clicked +新增物品")

            time.sleep(2.5)
            shot(page, "07_item_dialog")

            # ── Step 7: Dump 物品 dialog fields ─────────────────
            step("Step 7 — Dump 物品 dialog fields")
            print("  Inputs:")
            for el in page.locator("input, textarea, select").all():
                try:
                    if el.is_visible(timeout=200):
                        t  = el.get_attribute("type") or "text"
                        ph = el.get_attribute("placeholder") or ""
                        nm = el.get_attribute("name") or ""
                        v  = el.input_value() if t not in ("file","hidden") else ""
                        print(f"    type={t!r} ph={ph[:40]!r} name={nm!r} val={v[:20]!r}")
                except Exception:
                    pass
            print("  Buttons/Labels:")
            seen = set()
            for el in page.locator("button, label, [role='button']").all():
                try:
                    txt = el.inner_text().strip()
                    if txt and txt not in seen and len(txt) < 50 and el.is_visible(timeout=200):
                        seen.add(txt)
                        print(f"    {txt!r}")
                except Exception:
                    pass

            print("\n✅ 全部步驟完成")
            print(f"截圖: {LOGS_DIR}")
            time.sleep(999999)

        except Exception as e:
            print(f"\n❌ 錯誤:\n{e}")
            traceback.print_exc()
            shot(page, "ERROR")
            time.sleep(999999)
        finally:
            ctx.close()


if __name__ == "__main__":
    main()
