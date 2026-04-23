# -*- coding: utf-8 -*-
"""
Scrape waybills and statuses from SF Express China waybill list.
URL: https://www.sf-express.com/chn/sc/waybill/list

Returns structured records ready to merge into tracking.xlsx / Google Sheets.

Standalone usage:
    python sf_china_scraper.py          # scrape 3 pages (default)
    python sf_china_scraper.py 5        # scrape 5 pages
    python sf_china_scraper.py 0        # scrape all pages
"""
import re
import sys
import time

from playwright.sync_api import Page

from browser_utils import new_page, close_all
from config import SF_CHINA_LIST_URL
from logger import get_logger, screenshot_on_error

log = get_logger(__name__)

_STATUS_KEYWORDS = [
    # Traditional Chinese (HK site)
    "已簽收", "派送中", "待派送", "待寄出", "運送中", "攬收成功", "攬收",
    "已發出", "到達", "退回", "異常", "卡關", "問題件", "攔截", "已取消",
    # Simplified Chinese (China site)
    "已签收", "快件已签收", "签收", "派送中", "运输中", "待寄出", "已取消",
    "已揽收", "揽收成功", "待揽收", "已发出", "到达派送站", "到达", "退回",
    "异常", "问题件", "拦截", "转寄", "待取", "已入库",
]

# JS regex strings (escaped for Python f-string injection)
_WB_RE_JS = r"\b(\d{15,18}|SF\d{12,})\b"


def scrape_all_waybills(retries: int = 2, max_pages: int = 0) -> list[dict]:
    """
    Open the SF China waybill list, scrape pages, return list of dicts.
    Each dict: {waybill, status, status_time, recipient, address, freight, date, items}
    max_pages=0 means unlimited; set to e.g. 3 to limit scraping to 3 pages.
    Retries up to `retries` times on browser-closed errors.
    """
    for attempt in range(1, retries + 2):
        try:
            return _scrape_attempt(max_pages=max_pages)
        except Exception as e:
            err = str(e)
            if attempt <= retries and ("TargetClosed" in err or "Target page" in err
                                       or "browser has been closed" in err.lower()):
                log.warning("Browser closed unexpectedly (attempt %d), retrying in 3s…", attempt)
                close_all()
                time.sleep(3)
            else:
                raise
    return []


def _scrape_attempt(max_pages: int = 0) -> list[dict]:
    page = new_page(SF_CHINA_LIST_URL)
    results: list[dict] = []

    try:
        page.wait_for_load_state("networkidle", timeout=25_000)
        time.sleep(3)

        if _is_logged_out(page):
            log.warning("SF China site — not logged in, skipping scrape")
            return []

        try:
            _switch_sent_tab(page)
        except Exception:
            pass  # tab switch is optional; default view may already show sent items

        page_idx = 1
        while True:
            log.info("Scraping SF China list — page %d", page_idx)
            items = _scrape_page(page)
            if not items:
                break
            results.extend(items)
            log.info("  page %d → %d items (running total %d)", page_idx, len(items), len(results))

            if max_pages and page_idx >= max_pages:
                log.info("Reached max_pages=%d, stopping.", max_pages)
                break
            if not _next_page(page):
                break
            page_idx += 1
            time.sleep(2)

        log.info("SF China scrape complete: %d waybills", len(results))
        return results

    except Exception as e:
        path = screenshot_on_error(page, "sf_china_scraper")
        log.exception("SF China scrape failed (screenshot %s): %s", path, e)
        raise
    finally:
        close_all()


# ─── helpers ─────────────────────────────────────────────────────────────────

def _is_logged_out(page: Page) -> bool:
    url  = page.url.lower()
    text = ""
    try:
        text = page.inner_text("body", timeout=5_000)[:800]
    except Exception:
        pass
    if "login" in url or "sign" in url:
        return True
    for phrase in ["請登入", "登入賬號", "登录", "Login", "sign in"]:
        if phrase.lower() in text.lower():
            return True
    return False


def _switch_sent_tab(page: Page) -> None:
    clicked = page.evaluate("""() => {
        const want = ['我寄的', '寄出', '已寄出'];
        for (const el of document.querySelectorAll('*')) {
            if (el.offsetParent === null || el.children.length > 0) continue;
            const t = (el.textContent || '').trim();
            if (!want.includes(t)) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 20 || r.height < 8) continue;
            el.click();
            return true;
        }
        return false;
    }""")
    if clicked:
        time.sleep(1.5)


def _scrape_page(page: Page) -> list[dict]:
    """Extract all waybill rows from the currently visible list page."""
    raw = page.evaluate(f"""() => {{
        const WB_RE  = /{_WB_RE_JS}/g;
        const STATUS = {_STATUS_KEYWORDS!r};
        const results = [];
        const seen = new Set();

        // ── Strategy 1: find container elements that hold exactly one waybill ──
        const containers = Array.from(document.querySelectorAll(
            '[class*="order"],[class*="waybill"],[class*="item"],[class*="row"],[class*="card"],tr,li'
        )).filter(el => {{
            if (el.offsetParent === null) return false;
            const t = el.innerText || '';
            if (t.length < 10 || t.length > 3000) return false;
            return WB_RE.test(t);
        }});

        for (const el of containers) {{
            const text = (el.innerText || '').replace(/\\s+/g,' ').trim();
            WB_RE.lastIndex = 0;
            const wbs = [...text.matchAll(/{_WB_RE_JS}/g)];
            if (!wbs.length) continue;
            const wb = wbs[0][1];
            if (seen.has(wb)) continue;
            seen.add(wb);

            let status = '';
            for (const kw of STATUS) {{ if (text.includes(kw)) {{ status = kw; break; }} }}

            const dateM  = text.match(/\\d{{4}}[-/年]\\d{{1,2}}[-/月]\\d{{1,2}}(?:[日])?(?:[\\s T]\\d{{2}}:\\d{{2}})?/);
            const amtM   = text.match(/[¥￥]\\s*(\\d+\\.?\\d*)/i) || text.match(/(\\d+\\.\\d{{2}})\\s*[元円]/);
            const phone  = text.match(/1[3-9]\\d{{9}}|(\\+?852)?\\d{{8}}/);

            results.push({{
                waybill:     wb,
                status:      status,
                status_time: dateM ? dateM[0] : '',
                freight:     amtM  ? amtM[1]  : '',
                phone:       phone ? phone[0] : '',
                text:        text.slice(0, 600),
            }});
        }}

        // ── Strategy 2 fallback: scan full page text ──
        if (results.length === 0) {{
            const body = document.body ? (document.body.innerText || '') : '';
            const allWb = [...body.matchAll(/{_WB_RE_JS}/g)];
            for (const m of allWb) {{
                const wb = m[1];
                if (seen.has(wb)) continue;
                seen.add(wb);
                const idx = body.indexOf(wb);
                const snippet = body.slice(Math.max(0, idx - 200), idx + 800);
                let status = '';
                for (const kw of STATUS) {{ if (snippet.includes(kw)) {{ status = kw; break; }} }}
                const dateM = snippet.match(/\\d{{4}}[-/年]\\d{{1,2}}[-/月]\\d{{1,2}}/);
                const amtM  = snippet.match(/[¥￥]\\s*(\\d+\\.?\\d*)/i);
                results.push({{
                    waybill:     wb,
                    status:      status,
                    status_time: dateM ? dateM[0] : '',
                    freight:     amtM ? amtM[1] : '',
                    phone:       '',
                    text:        snippet,
                }});
            }}
        }}

        return results;
    }}""")

    return raw or []


def _next_page(page: Page) -> bool:
    """Click the next-page button. Returns True if clicked."""
    clicked = page.evaluate("""() => {
        const NEXT = ['下一頁','下一页','>','›','»','Next'];
        for (const el of document.querySelectorAll(
            'button,a,[role="button"],[class*="next"],[class*="page"]'
        )) {
            if (el.offsetParent === null) continue;
            const t = (el.textContent || el.getAttribute('aria-label') || '').trim();
            if (!NEXT.includes(t) && !NEXT.some(n => t === n)) continue;
            if (el.disabled || el.getAttribute('aria-disabled') === 'true') return false;
            if (el.classList.contains('disabled') || el.classList.contains('is-disabled')) return false;
            el.click();
            return true;
        }
        return false;
    }""")
    if clicked:
        try:
            page.wait_for_load_state("networkidle", timeout=12_000)
        except Exception:
            pass
        time.sleep(1.5)
    return bool(clicked)


# ─── standalone runner ────────────────────────────────────────────────────────

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    max_p = int(sys.argv[1]) if len(sys.argv) > 1 else 3
    label = f"前 {max_p} 頁" if max_p else "全部"
    print(f"▶ 開始爬取順豐運單列表（{label}）…\n")

    records = scrape_all_waybills(max_pages=max_p)

    if not records:
        print("⚠ 未取得任何資料。請確認已在 Chrome 登入順豐帳號。")
    else:
        print(f"✅ 共取得 {len(records)} 條運單記錄：\n")
        print(f"{'#':<4} {'運單號':<20} {'狀態':<12} {'狀態時間':<20} {'運費':<8} {'電話'}")
        print("─" * 80)
        for i, r in enumerate(records, 1):
            print(f"{i:<4} {r['waybill']:<20} {r['status']:<12} {r['status_time']:<20} {r['freight']:<8} {r['phone']}")
