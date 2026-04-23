# -*- coding: utf-8 -*-
"""
Scrape the SF HK waybill list (https://hk.sf-express.com/hk/tc/waybill/list).

Logic:
  1. Scrape up to MAX_LIST_PAGES pages of the sent-waybills list.
  2. For each waybill found:
       - If not in Excel → add a new row (append_from_sf).
       - If status changed → update status + timestamp.
  3. For any waybill that JUST became 已簽收 → click in, open 電子存根 tab,
     scrape all receipt fields, save to Excel.
"""
import re
import time

from playwright.sync_api import Page

from browser_utils import new_page, close_all
from config import SF_WAYBILL_URL, COL_STATUS, COL_FREIGHT, COL_NAME
from excel_manager import ExcelManager

# Direct waybill detail URL — avoids fragile list-row clicking
_SF_DETAIL_URL = "https://hk.sf-express.com/hk/tc/waybill/waybill-detail/{}"
from logger import get_logger, screenshot_on_error, toast_error, toast_ok

log = get_logger(__name__)

MAX_LIST_PAGES = 3

_STATUS_KEYWORDS = [
    "已簽收", "派送中", "待派送", "待寄出", "運送中", "攬收成功", "攬收",
    "已發出", "到達", "退回", "異常", "卡關", "問題件", "攔截", "已取消",
]


# ══════════════════════════════════════════════════════════════════════════════
# Public entry point
# ══════════════════════════════════════════════════════════════════════════════

def update_all_statuses(excel: ExcelManager) -> dict:
    """
    Main entry point called by update_status_cli.py and tracking_dashboard.py.
    Returns {waybill: status_str} for GUI display.
    """
    page = new_page(SF_WAYBILL_URL)
    results: dict = {}

    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
        time.sleep(2)
        _switch_to_sent_tab(page)
        _set_page_size_50(page)   # show 50 per page so we get all at once

        # ── Pass 1: scrape list (single pass after enlarging page size) ───────
        scraped: dict[str, dict] = {}
        for pnum in range(1, MAX_LIST_PAGES + 1):
            items = _scrape_list_page(page)
            for item in items:
                scraped[item["waybill"]] = item
            log.info("List page %d: %d waybills (total %d)", pnum, len(items), len(scraped))
            if not items or not _click_next_page(page):
                break
            time.sleep(1.5)

        log.info("Total from HK SF list: %d waybills", len(scraped))
        if not scraped:
            return {}

        # ── Pass 2: update Excel, collect newly-delivered waybills ────────────
        needs_receipt: list[str] = []

        for waybill, info in scraped.items():
            status = info.get("status") or "狀態不明"
            existing_row = excel.find_row_by_waybill(waybill)
            old_status = ""

            if existing_row:
                old_status = str(excel.ws.cell(existing_row, COL_STATUS).value or "")
            else:
                # Brand-new waybill — add to Excel
                excel.append_from_sf(waybill, info)
                log.info("New waybill added: %s → %s", waybill, status)

            excel.update_status(waybill, status, status_time=info.get("date", ""))
            results[waybill] = status

            # Fetch receipt if: 已簽收 AND (freight missing OR name is wrong/empty)
            if "已簽收" in status:
                row = excel.find_row_by_waybill(waybill)
                if row:
                    freight_val = str(excel.ws.cell(row, COL_FREIGHT).value or "")
                    name_val    = str(excel.ws.cell(row, COL_NAME).value or "")
                    freight_missing = not freight_val or freight_val.strip() in ("", "nan", "None")
                    # "Eliza poon" is the sender — name is wrong if it's English/sender
                    name_wrong = not name_val or name_val.strip() in ("", "nan", "None", "Eliza poon")
                    if freight_missing or name_wrong:
                        needs_receipt.append(waybill)

        # ── Pass 3: fetch 電子存根 for 已簽收 with missing data (max 8/run) ────
        for waybill in needs_receipt[:8]:
            log.info("Fetching 電子存根 for %s", waybill)
            try:
                detail = _get_electronic_receipt(page, waybill)
                if detail:
                    excel.update_receipt_detail(waybill, detail)
                    if detail.get("freight"):
                        excel.update_status(waybill, "已簽收",
                                           freight=detail["freight"],
                                           status_time=detail.get("delivery_time", ""))
                    results[waybill] = "已簽收 (存根已存)"
                    log.info("電子存根 saved for %s", waybill)
            except Exception as e:
                log.warning("電子存根 failed for %s: %s", waybill, e)
        if len(needs_receipt) > 5:
            log.info("Remaining %d receipts will be fetched on next run", len(needs_receipt) - 5)

        toast_ok(f"已更新 {len(results)} 個運單狀態")
        return results

    except Exception as e:
        path = screenshot_on_error(page, "status_updater")
        toast_error("狀態更新", str(e)[:100])
        log.exception("Status update failed (screenshot: %s)", path)
        raise
    finally:
        close_all()


# ══════════════════════════════════════════════════════════════════════════════
# List-page helpers
# ══════════════════════════════════════════════════════════════════════════════

def _switch_to_sent_tab(page: Page) -> None:
    clicked = page.evaluate("""() => {
        const targets = ['我寄的', '寄出'];
        for (const el of document.querySelectorAll('*')) {
            if (el.offsetParent === null || el.children.length > 0) continue;
            const t = el.textContent.trim();
            if (!targets.includes(t)) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 10) continue;
            el.click();
            return true;
        }
        return false;
    }""")
    if clicked:
        time.sleep(1.5)


def _set_page_size_50(page: Page) -> None:
    """Change the per-page selector from 10 to 50 so all waybills appear at once."""
    # Open the dropdown
    page.evaluate("""() => {
        const t = document.querySelector("[class*=select-new_trigger]");
        if (t) t.click();
    }""")
    time.sleep(0.8)
    # Click the 50 option
    clicked = page.evaluate("""() => {
        for (const el of document.querySelectorAll("[class*=select-new_option]")) {
            if ((el.textContent || "").includes("50")) {
                el.click();
                return true;
            }
        }
        return false;
    }""")
    if clicked:
        try:
            page.wait_for_load_state("networkidle", timeout=8_000)
        except Exception:
            pass
        time.sleep(1.5)
        log.info("Per-page set to 50")
    else:
        log.warning("Could not set per-page to 50 — will use pagination")


def _scrape_list_page(page: Page) -> list[dict]:
    STATUS = _STATUS_KEYWORDS
    raw = page.evaluate(f"""() => {{
        const STATUS = {STATUS!r};
        const WB_RE = /\\b(SF\\d{{10,}}|\\d{{15,18}})\\b/g;
        const results = [];
        const seen = new Set();

        const containers = Array.from(document.querySelectorAll(
            '[class*="order"],[class*="waybill"],[class*="item"],[class*="record"],tr,li'
        )).filter(el => {{
            if (el.offsetParent === null) return false;
            const t = el.innerText || '';
            WB_RE.lastIndex = 0;
            return t.length > 15 && t.length < 2000 && WB_RE.test(t);
        }});

        for (const el of containers) {{
            const text = (el.innerText || '').replace(/\\s+/g, ' ').trim();
            WB_RE.lastIndex = 0;
            const wbs = [...text.matchAll(/\\b(SF\\d{{10,}}|\\d{{15,18}})\\b/g)];
            if (!wbs.length) continue;
            const wb = wbs[0][1];
            if (seen.has(wb)) continue;
            seen.add(wb);
            let status = '';
            for (const kw of STATUS) {{ if (text.includes(kw)) {{ status = kw; break; }} }}
            const dateM = text.match(/\\d{{4}}[-/]\\d{{1,2}}[-/]\\d{{1,2}}(?:[T\\s]\\d{{2}}:\\d{{2}})?/);
            results.push({{ waybill: wb, status, date: dateM ? dateM[0] : '', text: text.slice(0, 400) }});
        }}

        if (results.length === 0) {{
            const body = document.body.innerText || '';
            WB_RE.lastIndex = 0;
            for (const m of body.matchAll(/\\b(SF\\d{{10,}}|\\d{{15,18}})\\b/g)) {{
                const wb = m[1];
                if (seen.has(wb)) continue;
                seen.add(wb);
                const idx = body.indexOf(wb);
                const snippet = body.slice(Math.max(0, idx - 150), idx + 500);
                let status = '';
                for (const kw of STATUS) {{ if (snippet.includes(kw)) {{ status = kw; break; }} }}
                const dateM = snippet.match(/\\d{{4}}[-/]\\d{{1,2}}[-/]\\d{{1,2}}/);
                results.push({{ waybill: wb, status, date: dateM ? dateM[0] : '', text: snippet }});
            }}
        }}

        return results;
    }}""")
    return raw or []


def _click_next_page(page: Page) -> bool:
    clicked = page.evaluate("""() => {
        const NEXT = ['下一頁', '下一页', '>', '›', '»', 'Next'];
        for (const el of document.querySelectorAll(
            'button,a,[role="button"],[class*="next"],[class*="page"]'
        )) {
            if (el.offsetParent === null) continue;
            const t = (el.textContent || el.getAttribute('aria-label') || '').trim();
            if (!NEXT.includes(t)) continue;
            if (el.disabled || el.getAttribute('aria-disabled') === 'true') return false;
            if (el.classList.contains('disabled') || el.classList.contains('is-disabled')) return false;
            el.click();
            return true;
        }
        return false;
    }""")
    if clicked:
        try:
            page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            pass
        time.sleep(1.5)
    return bool(clicked)


# ══════════════════════════════════════════════════════════════════════════════
# 電子存根 (electronic receipt) helpers
# ══════════════════════════════════════════════════════════════════════════════

def _get_electronic_receipt(page: Page, waybill: str) -> dict:
    """
    1. Load waybill detail (運單資訊 is default tab).
    2. Extract customer name/phone/address via CSS classes.
    3. Click 電子存根 tab and scrape all receipt fields.
    """
    detail_url = _SF_DETAIL_URL.format(waybill)
    log.info("Loading detail page: %s", detail_url)
    try:
        page.goto(detail_url, wait_until="networkidle", timeout=25_000)
    except Exception:
        page.goto(detail_url, timeout=25_000)
    time.sleep(3)

    # Step 1: Get recipient from 運單資訊 using real CSS classes
    recv = _scrape_recipient_css(page)
    log.info("  運單資訊 recipient: name=%s phone=%s", recv.get("name"), recv.get("phone"))

    # Step 2: Click 電子存根 tab
    _click_tab(page, "電子存根")
    time.sleep(2.5)

    detail = _scrape_receipt_fields(page)

    # CSS fills in ONLY what the regex couldn't find (regex is more reliable for Chinese names)
    if recv.get("phone") and not detail.get("recipient_phone"):
        detail["recipient_phone"] = recv["phone"]
    if recv.get("address") and not detail.get("recipient_address"):
        detail["recipient_address"] = recv["address"]

    return detail


def _scrape_recipient_css(page: Page) -> dict:
    """
    Extract recipient (收) name/phone/address from 運單資訊.
    Strategy: walk ALL matching elements in DOM order; after seeing recvIcon,
    the very next nameTelBox is the recipient's — regardless of how many
    other nameTelBox elements exist (profile nav, sender, etc.).
    """
    result = page.evaluate("""() => {
        // Collect all relevant elements in DOM order
        const els = Array.from(document.querySelectorAll(
            '[class*="recvIcon"],[class*="nameTelBox"],[class*="detailAddrBox"]'
        ));
        let seenRecv = false, nameBox = null, addrBox = null;
        for (const el of els) {
            const cls = el.className || '';
            if (cls.includes('recvIcon')) {
                seenRecv = true;
            } else if (seenRecv && !nameBox && cls.includes('nameTelBox')) {
                nameBox = el;
            } else if (seenRecv && nameBox && !addrBox && cls.includes('detailAddrBox')) {
                addrBox = el;
            }
        }
        if (!nameBox) return {};
        const nameRaw = (nameBox.innerText || '').trim();
        const addrRaw = addrBox ? (addrBox.innerText || '').trim() : '';
        const lines = nameRaw.split(/[\\n\\r]+/).map(s => s.trim()).filter(Boolean);
        return { name: lines[0] || '', phone: lines[1] || '', address: addrRaw };
    }""")
    return result or {}


def _click_waybill_row(page: Page, waybill: str) -> bool:
    """Find the waybill entry on the list and click it open."""
    clicked = page.evaluate(f"""() => {{
        const wb = '{waybill}';
        for (const el of document.querySelectorAll('*')) {{
            if (el.offsetParent === null) continue;
            const t = (el.textContent || '').trim();
            if (!t.startsWith(wb) && t !== wb) continue;
            if (el.children.length > 3) continue;
            const row = el.closest('tr,[class*="order"],[class*="waybill"],[class*="item"],[class*="record"]') || el;
            row.click();
            return true;
        }}
        return false;
    }}""")
    if clicked:
        time.sleep(1.5)
        return True

    # Fallback: use search box
    searched = page.evaluate(f"""() => {{
        for (const inp of document.querySelectorAll('input')) {{
            if (inp.offsetParent === null) continue;
            const ph = (inp.placeholder || '').toLowerCase();
            if (ph.includes('運單') || ph.includes('單號') || ph.includes('search') || ph.includes('waybill')) {{
                const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                setter.call(inp, '{waybill}');
                inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                inp.focus();
                return true;
            }}
        }}
        return false;
    }}""")
    if searched:
        page.keyboard.press("Enter")
        time.sleep(2)
        try:
            page.wait_for_load_state("networkidle", timeout=8_000)
        except Exception:
            pass
        clicked = page.evaluate(f"""() => {{
            const wb = '{waybill}';
            for (const el of document.querySelectorAll('*')) {{
                if (el.offsetParent === null) continue;
                if (!(el.textContent || '').includes(wb)) continue;
                if (el.children.length > 4) continue;
                const row = el.closest('tr,[class*="order"],[class*="waybill"],[class*="item"]') || el;
                row.click();
                return true;
            }}
            return false;
        }}""")
        if clicked:
            time.sleep(1.5)
    return bool(clicked)


def _click_tab(page: Page, tab_name: str) -> None:
    """Click a tab by name. Tries the known CSS class first, then falls back."""
    clicked = page.evaluate(f"""() => {{
        const name = '{tab_name}';
        // Primary: use the known CSS class from DevTools
        for (const el of document.querySelectorAll('[class*="waybill-tabs_tabItem"]')) {{
            if ((el.textContent || '').trim() === name) {{
                el.click();
                return true;
            }}
        }}
        // Fallback: any visible leaf element with exact text
        for (const el of document.querySelectorAll('*')) {{
            if (el.offsetParent === null || el.children.length > 2) continue;
            if ((el.textContent || '').trim() !== name) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 5) continue;
            el.click();
            return true;
        }}
        return false;
    }}""")
    if clicked:
        time.sleep(1.5)
    else:
        log.warning("Tab '%s' not found", tab_name)


def _scrape_receipt_fields(page: Page) -> dict:
    """Extract all 電子存根 fields from the current page text."""
    try:
        text = page.inner_text("body", timeout=10_000)
    except Exception:
        return {}

    if not text or "電子存根" not in text:
        log.warning("電子存根 not found in page text (len=%d)", len(text) if text else 0)
        return {}

    def _find(patterns: list) -> str:
        for pat in patterns:
            m = re.search(pat, text, re.DOTALL)
            if m:
                return (m.group(1) if m.lastindex else m.group(0)).strip()
        return ""

    # ── Recipient name — read from the address panel only (avoids body noise) ──
    # "已簽收\n香港\n→\n深圳市" in the body causes false regex matches.
    # Scope to [class*="send-receive"] which only has sender + recipient lines.
    recipient_name = ""
    try:
        panel_text = page.inner_text('[class*="send-receive"]', timeout=5_000)
        # panel_text: "寄\nEliza poon\n****2382\n香港...\n收\n黃業偉\n181****9028\n廣東省..."
        # Find everything after the "收" line
        pm = re.search(r'收\s+(\S{2,10})(?:\s|$)', panel_text)
        if pm:
            recipient_name = pm.group(1).strip()
    except Exception:
        pass
    if not recipient_name:
        recipient_name = _find([r"收件[人方][：:]\s*(\S{2,10})"])

    # ── Masked phone (keep as reference only) ─────────────────────────────────
    recipient_phone = _find([
        r"(\d{3}\*{4}\d{4})",      # 181****9028
        r"(\+?852[-\s]?\d{4}[-\s]?\d{4})",
    ])

    # ── Recipient address ─────────────────────────────────────────────────────
    recipient_address = _find([
        r"(廣東省[^\n\t]{5,100})",
        r"(廣東[^\n\t]{5,80})",
        r"(深圳[^\n\t]{5,80})",
        r"(香港[^\n\t]{10,100}(?:室|樓|層|號)[^\n]{0,30})",
        r"(九龍[^\n\t]{5,80})",
        r"(新界[^\n\t]{5,80})",
    ])

    # ── Items (no colon in label on some pages) ───────────────────────────────
    items = _find([
        r"託寄物\s*[：:]?\s*\n?\s*([^\n\t產品數量件]{2,300})",
        r"貨物名稱\s*[：:]\s*([^\n]{2,200})",
    ])
    # Clean up trailing whitespace / field names that bled in
    if items:
        items = re.split(r"\s{3,}|產品類型|數量：", items)[0].strip()

    qty             = _find([r"數量\s*[：:]\s*(\d+)"])
    pieces          = _find([r"件數\s*[：:]\s*(\d+)"])
    actual_weight   = _find([r"實際重量\s*[：:]\s*([\d.]+\s*[kK][gG]?)",
                              r"實重\s*[：:]\s*([\d.]+\s*[kK][gG]?)"])
    chg_weight      = _find([r"計算重量\s*[：:]\s*([\d.]+\s*[kK][gG]?)",   # actual label
                              r"計費重量\s*[：:]\s*([\d.]+\s*[kK][gG]?)",
                              r"計重\s*[：:]\s*([\d.]+\s*[kK][gG]?)"])
    freight         = _find([
        r"費用合計\s*[：:]?\s*HKD\s*([\d.]+)",
        r"運費\s*[：:]?\s*HKD\s*([\d.]+)",
        r"HKD\s*(\d+\.?\d*)",
    ])
    product_type    = _find([r"產品類型\s*[：:]\s*(\S+)"])
    payment         = _find([r"付款方式\s*[：:]\s*([^\n\t]{2,20})"])
    delivery_person = _find([r"收件員\s*[：:]\s*(\d+)"])
    # Actual label on page is 收時間, NOT 收件時間
    delivery_time   = _find([
        r"收時間\s*[：:]\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})",
        r"收件時間\s*[：:]\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})",
        r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})",
    ])

    log.info("  Receipt — name:%s  items:%s  freight:%s  time:%s",
             recipient_name, (items or "")[:40], freight, delivery_time)

    return {
        "recipient_name":     recipient_name,
        "recipient_phone":    recipient_phone,
        "recipient_address":  recipient_address,
        "items":              items,
        "qty":                qty,
        "pieces":             pieces,
        "actual_weight":      actual_weight,
        "chargeable_weight":  chg_weight,
        "freight":            freight,
        "product_type":       product_type,
        "payment":            payment,
        "delivery_person":    delivery_person,
        "delivery_time":      delivery_time,
    }
