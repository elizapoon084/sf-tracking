# -*- coding: utf-8 -*-
import logging
import os
from datetime import datetime
from pathlib import Path

from config import LOGS_DIR

os.makedirs(LOGS_DIR, exist_ok=True)


def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s — %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    # Console
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # Daily file
    log_file = os.path.join(LOGS_DIR, f"{datetime.now():%Y-%m-%d}.log")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


def screenshot_on_error(page, label: str) -> Path:
    """Save error screenshot; page may be None if browser never opened."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = Path(LOGS_DIR) / f"{ts}_{label}.png"
    try:
        if page is not None:
            page.screenshot(path=str(path))
    except Exception:
        pass
    return path


def toast(title: str, msg: str, icon: str = "✅") -> None:
    """Windows toast notification — silently skipped if win10toast unavailable."""
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(
            f"{icon} {title}", msg,
            duration=8, threaded=True,
        )
    except Exception:
        pass  # non-fatal — log already written


def toast_error(label: str, detail: str = "") -> None:
    toast("順丰自動化", f"❌ 卡住喺 {label}\n{detail}"[:200], icon="❌")


def toast_ok(msg: str) -> None:
    toast("順丰自動化", msg, icon="✅")
