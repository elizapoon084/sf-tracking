# -*- coding: utf-8 -*-
"""
SF Express automation:
  1. 收件人 智慧填寫 → 識別  (auto-fills name/phone/city/address)
  2. 選 自寄
  3. +新增物品 × each item  (name, price HKD, qty, origin 台灣)
  4. Summary popup — user confirms
  5. Submit → capture waybill
  6. Clearance document upload
"""
import threading
import time
import tkinter as tk
from tkinter import ttk
from pathlib import Path

from playwright.sync_api import Page, TimeoutError as PWTimeout

from browser_utils import new_page
from config import (
    SF_SHIP_URL, SF_CLEARANCE_URL,
    SF_SENDER_NAME, SF_PAYMENT_MODE, SF_ACCOUNT_NO,
    IMAGES_DIR,
)
from logger import get_logger, screenshot_on_error, toast_error

log = get_logger(__name__)


class SubmissionCancelledError(Exception):
    pass


def run_sf_submission(order: dict, pdf_path: str) -> str:
    """
    Full SF shipping flow.
    Returns waybill number string.
    Raises SubmissionCancelledError if user cancels.
    """
    page = new_page(SF_SHIP_URL)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15_000)
        time.sleep(2)

        _fill_recipient_smart(page, order)
        _select_self_drop(page)
        _fill_items(page, order["items"])

        _select_monthly_payment(page)
        _fill_monthly_account(page, SF_ACCOUNT_NO)
        _click_agree_checkbox(page)
        _click_terms_dialog_agree(page)
        _submit(page)
        waybill = _extract_waybill(page)
        log.info("SF waybill: %s", waybill)

        # Clearance upload is optional — skip if no ID photos
        id_front = order.get("id_front", "")
        id_back  = order.get("id_back", "")
        if id_front and id_back:
            try:
                _upload_clearance(order, waybill, id_front, id_back, pdf_path)
            except Exception as e:
                log.warning("Clearance upload failed (non-fatal): %s", e)

        return waybill

    except SubmissionCancelledError:
        raise
    except Exception as e:
        path = screenshot_on_error(page, "sf_automation")
        toast_error("順丰填單", str(e)[:100])
        log.exception("SF automation failed (screenshot: %s)", path)
        raise


# ─── Step 1: 智慧填寫 recipient ───────────────────────────────────────────────

def _fill_recipient_smart(page: Page, order: dict) -> None:
    """Click 收件人 智慧填寫 (index 1), paste text, JS-click 識別."""
    name       = order.get("name_simplified") or order["name"]
    smart_text = f"{name} {order['phone']} {order['address']}"

    try:
        # Click the RECIPIENT 智慧填寫 span (index 1 — after sender's at index 0)
        smart_spans = page.locator("span:has-text('智慧填寫')")
        smart_spans.nth(1).wait_for(state="visible", timeout=6000)
        smart_spans.nth(1).click()
        page.wait_for_timeout(800)

        # Fill the smart textarea (class contains 'intelAddr')
        ta = page.locator(
            "textarea[class*='intelAddr'], textarea[placeholder*='陳先生']"
        ).first
        ta.wait_for(state="visible", timeout=5000)
        ta.click()
        ta.type(smart_text, delay=20)
        page.wait_for_timeout(500)

        # 識別 is a SPAN (button_text__agCPY), not a <button> — JS click required
        page.evaluate("""() => {
            for (const el of document.querySelectorAll('*')) {
                if (el.childNodes.length === 1 &&
                    el.firstChild.nodeType === 3 &&
                    el.firstChild.textContent.trim() === '識別') {
                    el.click(); return;
                }
            }
        }""")
        page.wait_for_timeout(2500)
        log.info("Smart-fill recipient done: %s", name)

    except Exception as e:
        log.warning("Smart-fill failed (%s) — falling back to manual fill", e)
        _fill_recipient_manual(page, order)


def _fill_recipient_manual(page: Page, order: dict) -> None:
    """Fallback: fill recipient fields one by one."""
    name = order.get("name_simplified") or order["name"]
    # Recipient fields are the SECOND set (after sender's identical placeholders)
    _fill_nth(page, "input[placeholder='請填寫收件人姓名']", 0, name)
    _fill_nth(page, "input[placeholder='請填寫手機號碼或固話']", 1, order["phone"])
    _fill_nth(page, "input[placeholder='請填寫詳細地址'][name='detailAddress']",
              1, order["address"])


# ─── Step 2: 自寄 ─────────────────────────────────────────────────────────────

def _select_self_drop(page: Page) -> None:
    """Select 自寄, then click first station in the popup."""
    # JS: click exact-text '自寄' with sufficient size (avoids tiny spans)
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.textContent.trim() === '自寄' && el.offsetParent !== null) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 40 && rect.height > 20) { el.click(); return; }
            }
        }
    }""")
    page.wait_for_timeout(1200)

    # Click the red station name in the 自寄點搜索 popup
    try:
        page.get_by_text("新蒲崗福和工廠順豐站", exact=True).click(timeout=3000)
        page.wait_for_timeout(800)
        log.debug("Station clicked: 新蒲崗")
    except Exception:
        try:
            page.evaluate("""() => {
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
            let node;
            while (node = walker.nextNode()) {
                if (node.textContent.includes('新蒲崗福和')) {
                    let el = node.parentElement;
                    for (let i = 0; i < 5; i++) {
                        const rect = el.getBoundingClientRect();
                        if (rect.width > 200 && rect.height > 40) { el.click(); return; }
                        if (el.parentElement) el = el.parentElement;
                    }
                    el.click(); return;
                }
            }
        }""")
            page.wait_for_timeout(800)
            log.debug("Station card clicked via fallback")
        except Exception:
            pass


# ─── Step 3: 物品申報 ──────────────────────────────────────────────────────────

def _fill_items(page: Page, items: list) -> None:
    """For each item: click +新增物品 → fill dialog → click 確認."""
    for idx, item in enumerate(items):
        log.debug("Adding item %d: %s ×%d", idx + 1, item["name"], item["qty"])
        _add_one_item(page, item)
        page.wait_for_timeout(400)


def _add_one_item(page: Page, item: dict) -> None:
    # Click +新增物品 via JS (the element is a span, not a button)
    _js_click_add_item(page)
    page.wait_for_timeout(2000)

    # Select 物品 radio via JS
    _js_click_wupin_radio(page)
    page.wait_for_timeout(1500)

    # Fill each field via label-proximity JS click + keyboard type
    _js_fill_by_label(page, "物品名稱", item["name"])
    page.keyboard.press("Tab")
    page.wait_for_timeout(4000)  # SF site validates name and unlocks other fields

    _js_fill_by_label(page, "品牌", item.get("brand", ""))
    page.wait_for_timeout(2000)
    _js_fill_by_label(page, "材質", item.get("material", ""))
    page.wait_for_timeout(2000)
    _js_fill_by_label(page, "規格型號", item.get("spec", ""))
    page.wait_for_timeout(2000)
    _js_fill_by_label(page, "物品單價", int(item["unit_price"]))
    page.wait_for_timeout(500)
    _js_fill_by_label(page, "物品數量", item["qty"])
    page.wait_for_timeout(500)

    # 確認 via JS (class-based, not button locator)
    _js_click_confirm(page)
    page.wait_for_timeout(2000)
    _wait_dialog_closed(page)
    log.debug("Item added: %s × %d @ HKD%s", item["name"], item["qty"], item["unit_price"])


def _js_click_add_item(page: Page) -> None:
    info = page.evaluate("""() => {
        const candidates = ['+新增物品', '新增物品', '+ 新增物品'];
        for (const el of document.querySelectorAll('*')) {
            const t = el.textContent.trim();
            if (!candidates.includes(t)) continue;
            if (el.offsetParent === null || el.children.length > 0) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 10) continue;
            return { absY: r.top + window.scrollY, h: r.height };
        }
        return null;
    }""")
    if info:
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
    else:
        raise RuntimeError("Cannot find +新增物品 button on SF page")


def _js_click_wupin_radio(page: Page) -> None:
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.textContent.trim() === '物品' && el.offsetParent !== null) {
                if (el.children.length > 0) continue;
                const r = el.getBoundingClientRect();
                if (r.width < 10 || r.height < 5) continue;
                el.click(); return;
            }
        }
    }""")


def _js_fill_by_label(page: Page, label_text: str, value) -> bool:
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


def _js_click_confirm(page: Page) -> None:
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('[class*="package-declaration_confirm"]')) {
            if (el.offsetParent === null) continue;
            if (el.textContent.trim() !== '確認') continue;
            const r = el.getBoundingClientRect();
            if (r.width < 30 || r.height < 15) continue;
            el.click(); return;
        }
    }""")


def _wait_dialog_closed(page: Page, timeout: int = 5) -> bool:
    for _ in range(timeout * 2):
        gone = page.evaluate("""() => {
            return document.querySelectorAll('[role="dialog"][data-state="open"]').length === 0;
        }""")
        if gone:
            return True
        time.sleep(0.5)
    return False


# ─── Summary popup ────────────────────────────────────────────────────────────

def _show_summary_popup(order: dict) -> bool:
    """Show modal on main tkinter thread. Returns True if user confirms."""
    event  = threading.Event()
    result = [False]

    def _build_popup():
        popup = tk.Toplevel()
        popup.title("確認寄件資料")
        popup.grab_set()
        popup.resizable(False, False)

        pad = {"padx": 10, "pady": 4}
        tk.Label(popup, text="請確認以下資料", font=("", 13, "bold")).pack(**pad)

        info = tk.LabelFrame(popup, text="收件人")
        info.pack(fill="x", padx=10, pady=5)
        tk.Label(info, text=f"姓名: {order['name']}").pack(anchor="w", **pad)
        tk.Label(info, text=f"電話: {order['phone']}").pack(anchor="w", **pad)
        tk.Label(info, text=f"地址: {order['address']}",
                 wraplength=420, justify="left").pack(anchor="w", **pad)

        items_frame = tk.LabelFrame(popup, text="申報物品")
        items_frame.pack(fill="x", padx=10, pady=5)
        cols = ("物品名稱", "數量", "單價 HKD", "小計")
        tree = ttk.Treeview(items_frame, columns=cols, show="headings",
                            height=min(8, len(order["items"])))
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100, anchor="center")
        tree.column("物品名稱", width=200)
        for it in order["items"]:
            tree.insert("", "end", values=(
                it["name"][:20], it["qty"],
                f"${it['unit_price']:.0f}", f"${it['subtotal']:.0f}",
            ))
        tree.pack(fill="x", padx=5, pady=4)

        tk.Label(popup, text=f"申報總額: HKD {order['total']:.0f}",
                 font=("", 11, "bold")).pack(**pad)

        btns = tk.Frame(popup)
        btns.pack(pady=10)

        def on_confirm():
            result[0] = True
            popup.destroy()
            event.set()

        def on_cancel():
            popup.destroy()
            event.set()

        tk.Button(btns, text="✅ 確認提交", command=on_confirm,
                  bg="#27ae60", fg="white", width=14, height=2).pack(side="left", padx=8)
        tk.Button(btns, text="❌ 取消", command=on_cancel,
                  bg="#e74c3c", fg="white", width=10, height=2).pack(side="left", padx=8)
        popup.protocol("WM_DELETE_WINDOW", on_cancel)

    root = tk._default_root
    if root:
        root.after(0, _build_popup)
    event.wait(timeout=300)
    return result[0]


# ─── Steps 6-9: 付款→月結→同意→下單 ─────────────────────────────────────────

def _select_monthly_payment(page: Page) -> None:
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.offsetParent === null || el.children.length > 0) continue;
            const t = el.textContent.trim();
            if (t !== '月結' && t !== '寄付月結') continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 15) continue;
            el.click(); return;
        }
    }""")
    page.wait_for_timeout(1500)
    log.debug("Monthly payment selected")


def _fill_monthly_account(page: Page, account: str) -> None:
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('*')) {
            if (el.textContent.trim() === '付款方式' && el.offsetParent !== null) {
                el.scrollIntoView({block: 'center'}); return;
            }
        }
    }""")
    page.wait_for_timeout(800)
    info = page.evaluate("""() => {
        for (const inp of document.querySelectorAll('input')) {
            if (inp.offsetParent === null) continue;
            const ph = inp.placeholder || '';
            const r = inp.getBoundingClientRect();
            if (r.width === 0 || r.height === 0) continue;
            if (ph.includes('月結') || ph.includes('卡號')) {
                return {x: r.left + r.width/2, y: r.top + r.height/2};
            }
        }
        return null;
    }""")
    if not info:
        log.warning("月結卡號 input not found")
        return
    page.mouse.click(info["x"], info["y"])
    time.sleep(0.3)
    page.mouse.click(info["x"], info["y"])
    time.sleep(0.3)
    page.keyboard.press("Control+A")
    page.keyboard.press("Delete")
    page.keyboard.type(account, delay=50)
    time.sleep(0.8)
    page.evaluate(f"""() => {{
        for (const el of document.querySelectorAll('*')) {{
            if (el.offsetParent === null || el.children.length > 0) continue;
            const t = el.textContent.trim();
            if (t.includes('{account}') && t.length < 50) {{
                const r = el.getBoundingClientRect();
                if (r.width < 30) return;
                el.click(); return;
            }}
        }}
    }}""")
    page.keyboard.press("Tab")
    page.wait_for_timeout(500)
    log.debug("Monthly account filled: %s", account)


def _click_agree_checkbox(page: Page) -> None:
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('[class*="agreedCheckbox"]')) {
            if (el.offsetParent !== null) { el.scrollIntoView({block: 'center'}); return; }
        }
    }""")
    page.wait_for_timeout(800)
    page.evaluate("""() => {
        for (const el of document.querySelectorAll('[class*="agreedCheckbox"]')) {
            if (el.offsetParent === null) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 30) continue;
            el.click(); return;
        }
        for (const el of document.querySelectorAll('[class*="checkbox_checkbox"]')) {
            if (el.offsetParent === null) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 30) continue;
            const txt = (el.parentElement || el).textContent || '';
            if (!txt.includes('閱讀') || !txt.includes('同意')) continue;
            el.click(); return;
        }
    }""")
    page.wait_for_timeout(2000)
    log.debug("Agree checkbox clicked")


def _click_terms_dialog_agree(page: Page) -> None:
    page.wait_for_timeout(1500)
    page.evaluate("""() => {
        const candidates = [
            '同意本條款,下次不再提示', '同意本條款，下次不再提示',
            '同意本條款', '同意並繼續', '同意',
        ];
        for (const target of candidates) {
            for (const el of document.querySelectorAll('*')) {
                if (el.offsetParent === null) continue;
                const t = el.textContent.trim();
                if (t !== target) continue;
                if (el.children.length > 0) continue;
                const r = el.getBoundingClientRect();
                if (r.width < 50 || r.height < 20) continue;
                el.click(); return;
            }
        }
    }""")
    page.wait_for_timeout(1500)
    log.debug("Terms dialog agreed")


# ─── Submit & waybill ─────────────────────────────────────────────────────────

def _submit(page: Page) -> None:
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(500)

    _JS_CLICK_SUBMIT = """() => {
        for (const el of document.querySelectorAll('[class*="submitBtn"]')) {
            if (el.offsetParent === null) continue;
            if (!el.textContent.trim().includes('下單')) continue;
            const cls = (el.className || '').toString();
            if (cls.includes('disabled')) return {ok: false, reason: 'disabled'};
            const r = el.getBoundingClientRect();
            if (r.width < 40) continue;
            el.click();
            return {ok: true, method: 'submitBtn', cls: cls.slice(0,80)};
        }
        for (const el of document.querySelectorAll('[role="button"]')) {
            if (el.offsetParent === null) continue;
            if (el.textContent.trim() !== '下單') continue;
            const cls = (el.className || '').toString();
            if (cls.includes('disabled')) return {ok: false, reason: 'role_disabled'};
            el.click();
            return {ok: true, method: 'role_button'};
        }
        for (const el of document.querySelectorAll('*')) {
            if (el.offsetParent === null || el.children.length > 0) continue;
            if (el.textContent.trim() !== '下單') continue;
            const r = el.getBoundingClientRect();
            if (r.width < 40 || r.height < 20) continue;
            el.click();
            return {ok: true, method: 'text'};
        }
        return {ok: false, reason: 'not_found'};
    }"""

    # Retry up to 8 times (16 seconds) waiting for button to become enabled
    for attempt in range(8):
        result = page.evaluate(_JS_CLICK_SUBMIT)
        if result.get("ok"):
            log.info("SF form submitted (attempt %d): %s", attempt + 1, result)
            break
        log.debug("Submit attempt %d: %s", attempt + 1, result)
        time.sleep(2)
    else:
        raise RuntimeError(f"下單 button not found or still disabled after retries: {result}")

    page.wait_for_load_state("domcontentloaded", timeout=20_000)
    time.sleep(2)


def _extract_waybill(page: Page) -> str:
    """Find SF waybill number (starts with SF) on confirmation page."""
    import re
    for attempt in range(5):
        content = page.content()
        m = re.search(r"SF\d{10,}", content)
        if m:
            return m.group(0)
        time.sleep(1)
    return "SF_UNKNOWN"


# ─── Clearance upload ─────────────────────────────────────────────────────────

def _upload_clearance(order: dict, waybill: str,
                      id_front: str, id_back: str, pdf_path: str) -> None:
    if not id_front or not id_back:
        log.warning("ID photos missing — skipping clearance upload")
        return

    page = new_page(SF_CLEARANCE_URL)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=12_000)
        time.sleep(1)

        # Enter waybill
        wb_input = page.locator(
            "input[placeholder*='運單'], input[placeholder*='單號']"
        ).first
        wb_input.wait_for(state="visible", timeout=6000)
        wb_input.fill(waybill)
        page.wait_for_timeout(400)

        # Upload files
        inputs = page.locator("input[type='file']").all()
        files = [id_front, id_back, pdf_path]
        for i, fpath in enumerate(files):
            if i < len(inputs) and Path(fpath).exists():
                inputs[i].set_input_files(fpath)
                page.wait_for_timeout(400)

        # Submit
        submit = page.locator(
            "button[type='submit'], button:has-text('提交'), button:has-text('上載')"
        ).first
        submit.wait_for(state="visible", timeout=6000)
        submit.click()
        page.wait_for_load_state("domcontentloaded", timeout=15_000)
        log.info("Clearance uploaded for waybill %s", waybill)

    except Exception as e:
        path = screenshot_on_error(page, "sf_clearance")
        log.exception("Clearance upload failed (screenshot: %s)", path)
        raise


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _fill_nth(page: Page, selector: str, nth: int, value: str) -> None:
    try:
        field = page.locator(selector).nth(nth)
        if field.is_visible(timeout=2000):
            field.triple_click()
            field.fill(value)
    except Exception:
        pass
