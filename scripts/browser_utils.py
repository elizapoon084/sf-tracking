# -*- coding: utf-8 -*-
"""Shared Playwright browser helpers used by all automation modules."""
from playwright.sync_api import sync_playwright, BrowserContext, Page

from config import CHROME_PROFILE, BROWSER_ARGS, PLAYWRIGHT_TIMEOUT, PLAYWRIGHT_SLOW_MO
from logger import get_logger

log = get_logger(__name__)

_pw_instance  = None
_browser_ctx: BrowserContext | None = None


def _clear_locks() -> None:
    import os
    for lf in ["lockfile", "SingletonLock", "SingletonSocket", "SingletonCookie"]:
        try:
            os.remove(os.path.join(CHROME_PROFILE, lf))
        except Exception:
            pass


def get_context() -> BrowserContext:
    """Return (or create) the shared persistent Chrome context."""
    global _pw_instance, _browser_ctx
    if _browser_ctx is not None:
        return _browser_ctx

    _clear_locks()
    log.info("Launching Chrome with profile: %s", CHROME_PROFILE)
    _pw_instance = sync_playwright().start()
    _browser_ctx = _pw_instance.chromium.launch_persistent_context(
        CHROME_PROFILE,
        channel="chrome",
        headless=False,
        args=BROWSER_ARGS,
        slow_mo=PLAYWRIGHT_SLOW_MO,
        viewport={"width": 1280, "height": 900},
        accept_downloads=True,
    )
    _browser_ctx.set_default_timeout(PLAYWRIGHT_TIMEOUT)
    return _browser_ctx


def new_page(url: str | None = None) -> Page:
    """Open a new tab (or reuse a blank one). Optionally navigate to url."""
    ctx = get_context()
    # Reuse an existing blank page if one exists
    for p in ctx.pages:
        if p.url in ("about:blank", "chrome://newtab/", ""):
            page = p
            break
    else:
        page = ctx.new_page()

    if url:
        page.goto(url, wait_until="domcontentloaded")
    return page


def close_all() -> None:
    """Close browser — only call on clean exit, NOT on error."""
    global _pw_instance, _browser_ctx
    try:
        if _browser_ctx:
            _browser_ctx.close()
        if _pw_instance:
            _pw_instance.stop()
    except Exception:
        pass
    _browser_ctx = None
    _pw_instance = None
