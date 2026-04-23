# -*- coding: utf-8 -*-
"""
SF v24 — 基於 DevTools 睇到嘅真實 class:
  - Checkbox container: checkout-panel_agreedCheckbox__TiO9C
  - 下單 button: <div role="button"> 含 'submitBtn'
"""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright
from config import SF_SHIP_URL, CHROME_PROFILE, BROWSER_ARGS, LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)

for lf in ["lockfile", "SingletonLock", "SingletonSocket", "SingletonCookie"]:
    try:
        os.remove(os.path.join(CHROME_PROFILE, lf))
    except Exception:
        pass

SENDER_TEXT = "潘正儀 66832382 香港九龍新蒲崗大有街33號佳力工業大廈603室"
RECIP_TEXT  = "黄业伟 18125989028 廣東省深圳市龍崗區南聯劉屋村南段74號1樓"

DEMO_ITEMS = [
    {"name": "女士寶", "brand": "INOVITAL", "material": "膠囊",
     "spec": "60粒", "unit_price": 23, "qty": 3},
    {"name": "雄風寶", "brand": "INOVITAL", "material": "膠囊",
     "spec": "60粒", "unit_price": 20, "qty": 3},
]

MONTHLY_ACCOUNT = "8526937071"


def shot(page, name):
    p = os.path.join(LOGS_DIR, f"sf_{name}.png")
    page.screenshot(path=p, full_page=False)
    print(f"  📸 {p}")


def smart_fill(page, text, which):
    page.locator("span:has-text('智慧填寫')").nth(which).click()
    time.sleep(0.8)
    ta = page.locator("textarea[class*='intelAddr'], textarea[placeholder*='陳先生']").first
    ta.wait_for(state="visible", timeout=5000)
    ta.click()
    ta.type(text, delay=20)
    time.sleep(0.5)
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.childNodes.length===1 && el.firstChild.nodeType===3
                && el.firstChild.textContent.trim()==='識別') {
                el.click(); return;
            }
        }
    }""")
    time.sleep(2.5)


def fill_by_label(page, label_text, value):
    info = page.evaluate(f"""() => {{
        const targets = ['{label_text}', '{label_text}：', '{label_text}:'];
        const labels = [];
        for (const el of document.querySelectorAll('*')) {{
            if (el.offsetParent === null) continue;
            const t = el.textContent.trim();
            if (targets.includes(t)) {{ labels.push(el); continue; }}
            if (t.length < 10 && t.includes('{label_text}')) {{ labels.push(el); }}
        }}
        if (labels.length === 0) return null;
        labels.sort((a, b) => a.children.length - b.children.length);
        for (const lbl of labels) {{
            let parent = lbl.parentElement;
            for (let depth = 0; depth < 5; depth++) {{
                if (!parent) break;
                const inputs = parent.querySelectorAll(
                    'input:not([type="hidden"]):not([type="radio"]):not([type="checkbox"])'
                );
                for (const inp of inputs) {{
                    if (inp.offsetParent === null) continue;
                    const r = inp.getBoundingClientRect();
                    if (r.width === 0) continue;
                    return {{x: r.left + r.width / 2, y: r.top + r.height / 2}};
                }}
                parent = parent.parentElement;
            }}
        }}
        return null;
    }}""")
    if not info:
        return False
    page.mouse.click(info["x"], info["y"])
    time.sleep(0.3)
    page.keyboard.press("Control+A")
    time.sleep(0.1)
    page.keyboard.press("Delete")
    time.sleep(0.1)
    page.keyboard.type(str(value), delay=30)
    time.sleep(0.3)
    return True


def click_wupin_radio(page):
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.textContent.trim() === '物品' && el.offsetParent !== null) {
                if (el.children.length > 0) continue;
                const r = el.getBoundingClientRect();
                if (r.width < 10 || r.height < 5) continue;
                el.click(); return true;
            }
        }
        return false;
    }""")


def click_dialog_confirm(page):
    result = page.evaluate("""() => {
        for (const el of document.querySelectorAll('[class*="package-declaration_confirm"]')) {
            if (el.offsetParent === null) continue;
            if (el.textContent.trim() !== '確認') continue;
            const r = el.getBoundingClientRect();
            if (r.width < 30 || r.height < 15) continue;
            el.click();
            return {ok: true, tag: el.tagName};
        }
        return {ok: false};
    }""")
    print(f"     確認 dialog: {result}")
    time.sleep(1)
    return result.get("ok", False)


def wait_dialog_closed(page, timeout=5):
    for _ in range(timeout * 2):
        gone = page.evaluate("""() => {
            return document.querySelectorAll('[role="dialog"][data-state="open"]').length === 0;
        }""")
        if gone:
            return True
        time.sleep(0.5)
    return False


def fill_one_item(page, item, idx):
    print(f"\n  ─── 物品 {idx + 1}: {item['name']} ───")
    click_wupin_radio(page)
    time.sleep(1.5)
    fill_by_label(page, "物品名稱", item["name"])
    print(f"  ✅ 物品名稱")
    page.keyboard.press("Tab")
    time.sleep(4)
    fill_by_label(page, "品牌", item["brand"]); print(f"  ✅ 品牌"); time.sleep(2)
    fill_by_label(page, "材質", item["material"]); print(f"  ✅ 材質"); time.sleep(2)
    fill_by_label(page, "規格型號", item["spec"]); print(f"  ✅ 規格"); time.sleep(2)
    fill_by_label(page, "物品單價", item["unit_price"]); print(f"  ✅ 單價"); time.sleep(0.5)
    fill_by_label(page, "物品數量", item["qty"]); print(f"  ✅ 數量"); time.sleep(0.5)
    shot(page, f"item{idx+1}_filled")
    time.sleep(2)
    click_dialog_confirm(page)
    time.sleep(2)
    closed = wait_dialog_closed(page)
    print(f"  Dialog closed: {closed}")


def click_add_item_button(page):
    info = page.evaluate("""() => {
        const candidates = ['+新增物品', '新增物品', '+ 新增物品'];
        for (const el of document.querySelectorAll('*')) {
            const t = el.textContent.trim();
            if (!candidates.includes(t)) continue;
            if (el.offsetParent === null || el.children.length > 0) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 10) continue;
            return { absY: r.top + window.scrollY, h: r.height, x: r.left + r.width / 2 };
        }
        return null;
    }""")
    if not info:
        return False
    page.evaluate(f"window.scrollTo(0, {info['absY'] - 450 + info['h'] / 2})")
    time.sleep(0.8)
    click_info = page.evaluate("""() => {
        const candidates = ['+新增物品', '新增物品', '+ 新增物品'];
        for (const el of document.querySelectorAll('*')) {
            const t = el.textContent.trim();
            if (!candidates.includes(t)) continue;
            if (el.offsetParent === null || el.children.length > 0) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 10) continue;
            return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
        }
        return null;
    }""")
    if click_info:
        page.mouse.click(click_info['x'], click_info['y'])
        time.sleep(2)
        return True
    return False


def select_monthly_payment(page):
    result = page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.offsetParent === null || el.children.length > 0) continue;
            const t = el.textContent.trim();
            if (t !== '月結' && t !== '寄付月結') continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 15) continue;
            el.click();
            return {ok: true, text: t};
        }
        return {ok: false};
    }""")
    print(f"  月結: {result}")
    time.sleep(1.5)
    return result.get("ok", False)


def fill_monthly_account(page, account):
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.textContent.trim() === '付款方式' && el.offsetParent !== null) {
                el.scrollIntoView({block: 'center'});
                return;
            }
        }
    }""")
    time.sleep(0.8)

    info = page.evaluate("""() => {
        for (const inp of document.querySelectorAll('input')) {
            if (inp.offsetParent === null) continue;
            const ph = inp.placeholder || '';
            const rect = inp.getBoundingClientRect();
            if (rect.width === 0 || rect.height === 0) continue;
            if (ph.includes('月結') || ph.includes('卡號')) {
                return {x: rect.left + rect.width/2, y: rect.top + rect.height/2, ph};
            }
        }
        return null;
    }""")
    if not info:
        print("     ❌ 揾唔到月結卡號 input")
        return False

    print(f"     揾到月結欄: {info}")
    page.mouse.click(info["x"], info["y"])
    time.sleep(0.3)
    page.mouse.click(info["x"], info["y"])
    time.sleep(0.5)
    page.keyboard.press("Control+A")
    page.keyboard.press("Delete")
    page.keyboard.type(account, delay=50)
    time.sleep(0.5)

    current = page.evaluate("""() => {
        for (const inp of document.querySelectorAll('input')) {
            if (inp.offsetParent === null) continue;
            const ph = inp.placeholder || '';
            if (ph.includes('月結') || ph.includes('卡號')) return inp.value;
        }
        return null;
    }""")
    print(f"     value: {current!r}")

    # 有 dropdown option 就揀
    time.sleep(0.8)
    clicked = page.evaluate(f"""() => {{
        for (const el of document.querySelectorAll('*')) {{
            if (el.offsetParent === null || el.children.length > 0) continue;
            const t = el.textContent.trim();
            if (t.includes('{account}') && t.length < 50) {{
                const r = el.getBoundingClientRect();
                if (r.width < 30) return null;
                el.click();
                return t;
            }}
        }}
        return null;
    }}""")
    if clicked:
        print(f"     揀咗 option: {clicked!r}")

    page.keyboard.press("Tab")
    time.sleep(0.5)
    print(f"  ✅ 月結卡號已處理")
    return True


def click_agree_checkbox(page):
    """
    ⭐ v24 真實 class:checkout-panel_agreedCheckbox__TiO9C
    成個 container 可 click,直接撳中心就得
    """
    # 先 scroll 到 checkbox
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('[class*="agreedCheckbox"]')) {
            if (el.offsetParent !== null) {
                el.scrollIntoView({block: 'center'});
                return;
            }
        }
    }""")
    time.sleep(0.8)

    result = page.evaluate("""() => {
        for (const el of document.querySelectorAll('[class*="agreedCheckbox"]')) {
            if (el.offsetParent === null) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 30) continue;
            el.click();
            return {
                ok: true,
                tag: el.tagName,
                cls: (el.className || '').toString().slice(0, 100),
                x: r.left + r.width/2,
                y: r.top + r.height/2
            };
        }
        // Fallback:揾 checkbox_checkbox class
        for (const el of document.querySelectorAll('[class*="checkbox_checkbox"]')) {
            if (el.offsetParent === null) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 30) continue;
            // 要附近有「閱讀並同意」text
            const parent = el.parentElement;
            if (!parent) continue;
            const txt = parent.textContent || '';
            if (!txt.includes('閱讀') || !txt.includes('同意')) continue;
            el.click();
            return {ok: true, fallback: true, cls: (el.className || '').toString().slice(0, 100)};
        }
        return {ok: false};
    }""")
    print(f"     閱讀同意 checkbox: {result}")
    time.sleep(1)
    return result.get("ok", False)


def click_terms_dialog_agree(page):
    """條款 dialog 出咗之後,撳紅色「同意本條款,下次不再提示」"""
    time.sleep(1.5)
    shot(page, "terms_dialog")

    result = page.evaluate("""() => {
        const candidates = [
            '同意本條款,下次不再提示',
            '同意本條款,下次不再提示',
            '同意本條款',
            '同意並繼續',
            '同意',
        ];
        for (const target of candidates) {
            for (const el of document.querySelectorAll('*')) {
                if (el.offsetParent === null) continue;
                const t = el.textContent.trim();
                if (t !== target) continue;
                if (el.children.length > 0) continue;
                const r = el.getBoundingClientRect();
                if (r.width < 50 || r.height < 20) continue;
                el.click();
                return {ok: true, text: t};
            }
        }
        // Fallback
        for (const el of document.querySelectorAll('*')) {
            if (el.offsetParent === null) continue;
            const t = el.textContent.trim();
            if (!t.includes('同意') || t.length > 30) continue;
            if (!t.includes('不再提示') && !t.includes('繼續') && t !== '同意本條款') continue;
            if (el.children.length > 0) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 50 || r.height < 20) continue;
            el.click();
            return {ok: true, text: t, fallback: true};
        }
        return {ok: false};
    }""")
    print(f"     條款同意: {result}")
    time.sleep(1.5)
    return result.get("ok", False)


def click_submit_order(page):
    """
    ⭐ v24 真實:<div role="button" class="...submitBtn...">下單</div>
    """
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(0.5)

    result = page.evaluate("""() => {
        // 試 1:精準 — class 含 submitBtn
        for (const el of document.querySelectorAll('[class*="submitBtn"]')) {
            if (el.offsetParent === null) continue;
            if (!el.textContent.trim().includes('下單')) continue;
            // 確保唔係 disabled
            const cls = (el.className || '').toString();
            if (cls.includes('disabled')) {
                return {ok: false, reason: 'disabled', cls: cls.slice(0, 100)};
            }
            const r = el.getBoundingClientRect();
            if (r.width < 40) continue;
            el.click();
            return {ok: true, method: 'submitBtn', cls: cls.slice(0, 100)};
        }
        // 試 2:揾 role="button" 有「下單」
        for (const el of document.querySelectorAll('[role="button"]')) {
            if (el.offsetParent === null) continue;
            if (el.textContent.trim() !== '下單') continue;
            const cls = (el.className || '').toString();
            if (cls.includes('disabled')) return {ok: false, reason: 'role_disabled'};
            el.click();
            return {ok: true, method: 'role_button'};
        }
        // 試 3:任何 text 「下單」
        for (const el of document.querySelectorAll('*')) {
            if (el.offsetParent === null || el.children.length > 0) continue;
            if (el.textContent.trim() !== '下單') continue;
            const r = el.getBoundingClientRect();
            if (r.width < 40 || r.height < 20) continue;
            el.click();
            return {ok: true, method: 'text'};
        }
        return {ok: false};
    }""")
    print(f"  下單: {result}")
    time.sleep(2)
    return result.get("ok", False)


with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=BROWSER_ARGS, slow_mo=150,
        viewport={"width": 1280, "height": 900},
    )
    page = ctx.new_page()
    page.goto(SF_SHIP_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(3)

    print("\n=== Step 1: 寄件人 ===")
    smart_fill(page, SENDER_TEXT, 0)
    page.evaluate("""() => {
        const inputs = document.querySelectorAll("input[name='contactName']");
        const input = inputs[0];
        if (!input) return;
        const setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, '潘正儀');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.blur();
    }""")
    time.sleep(0.5)

    print("\n=== Step 2: 收件人 ===")
    smart_fill(page, RECIP_TEXT, 1)

    print("\n=== Step 3: 自寄 ===")
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.textContent.trim() === '自寄' && el.offsetParent !== null) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 40 && rect.height > 20) {
                    el.click(); return true;
                }
            }
        }
        return false;
    }""")
    time.sleep(1.5)

    print(f"\n=== Step 4-5: Loop {len(DEMO_ITEMS)} 件物品 ===")
    for idx, item in enumerate(DEMO_ITEMS):
        print(f"\n>>> 撳「+新增物品」({idx + 1}/{len(DEMO_ITEMS)})")
        if not click_add_item_button(page):
            print("❌ 撳唔到")
            break
        time.sleep(1.5)
        fill_one_item(page, item, idx)

    shot(page, "after_all_items")

    print("\n=== Step 6: 揀月結 ===")
    select_monthly_payment(page)
    time.sleep(1)
    shot(page, "06_monthly")

    print(f"\n=== Step 7: 填月結卡號 {MONTHLY_ACCOUNT} ===")
    fill_monthly_account(page, MONTHLY_ACCOUNT)
    time.sleep(0.5)
    shot(page, "07_account")

    print("\n=== Step 8a: 撳閱讀並同意 ===")
    click_agree_checkbox(page)
    time.sleep(0.5)
    shot(page, "08a_agreed")

    print("\n=== Step 8b: 條款 dialog 撳「同意本條款」 ===")
    click_terms_dialog_agree(page)
    time.sleep(0.5)
    shot(page, "08b_terms_agreed")

    print("\n=== Step 9: 撳下單 ===")
    click_submit_order(page)
    time.sleep(3)
    shot(page, "09_submitted")

    print(f"\n完成。截圖喺: {LOGS_DIR}")
    input("按 Enter 關瀏覽器...")
    ctx.close()
