# -*- coding: utf-8 -*-
"""
Parse WhatsApp order text into a structured OrderDict.

Expected format (flexible):
  1084043-3件，1084065-4件，1000044-2件，寄黃业偉 18125989028 廣東省深圳市龍崗區南聯劉屋村南段74號1樓
"""
import json
import os
import re
from datetime import date
from typing import TypedDict

try:
    import zhconv
    _HAS_ZHCONV = True
except ImportError:
    _HAS_ZHCONV = False

from config import PRODUCTS_JSON
from logger import get_logger

log = get_logger(__name__)

# ─── TypedDicts ───────────────────────────────────────────────────────────────

class ItemDict(TypedDict):
    sku:        str
    qty:        int
    name:       str
    brand:      str
    material:   str
    spec:       str
    origin:     str
    unit_price: float
    subtotal:   float


class OrderDict(TypedDict):
    raw_text:       str
    name:           str   # original (as in WhatsApp) — used for ID folder
    name_simplified: str  # simplified Chinese — used for SF form & Excel
    phone:          str
    address:        str
    items:          list  # list[ItemDict]
    total:          float
    date:           str   # YYYY-MM-DD
    notes:          str


# ─── Compiled patterns ────────────────────────────────────────────────────────

# e.g. "1084043-3件" or "1084043×3件" or "1084043 x3件"
_RE_ITEM = re.compile(
    r"(\d{5,10})"                       # SKU
    r"[\s\-×xX*]+"                      # separator
    r"(\d+)\s*件",
    re.UNICODE,
)

# "寄" followed by a name (1-6 CJK or letters), then whitespace
_RE_AFTER_JI = re.compile(r"寄\s*([^\s\d，,、]{1,10})\s+")

# 11-digit mainland phone or 8-digit HK phone
_RE_PHONE = re.compile(r"1[3-9]\d{9}|[2-9]\d{7}")

# Everything after the phone number to end of meaningful text
_RE_ADDRESS_AFTER_PHONE = re.compile(
    r"(?:1[3-9]\d{9}|[2-9]\d{7})\s*(.+)", re.DOTALL
)


# ─── Public API ───────────────────────────────────────────────────────────────

def parse_order(raw_text: str) -> OrderDict:
    """Raises ValueError with a human-readable message on parse failure."""
    text = raw_text.strip()

    items = _extract_items(text)
    if not items:
        raise ValueError("找唔到貨品 — 請確認格式係「貨號-數量件」例如 1084043-3件")

    name = _extract_name(text)
    if not name:
        raise ValueError("找唔到收件人名 — 請確認格式係「寄[姓名] 電話 地址」")

    phone = _extract_phone(text)
    if not phone:
        raise ValueError("找唔到電話號碼")

    address = _extract_address(text)
    if not address:
        raise ValueError("找唔到地址")

    total = sum(it["subtotal"] for it in items)

    return OrderDict(
        raw_text=raw_text,
        name=name,
        name_simplified=to_simplified(name),
        phone=phone,
        address=address,
        items=items,
        total=round(total, 2),
        date=date.today().isoformat(),
        notes="",
    )


def to_simplified(text: str) -> str:
    if _HAS_ZHCONV:
        return zhconv.convert(text, "zh-hans")
    return text  # fallback: keep as-is


# ─── Private helpers ──────────────────────────────────────────────────────────

def _load_products() -> dict:
    if os.path.exists(PRODUCTS_JSON):
        with open(PRODUCTS_JSON, encoding="utf-8") as f:
            return json.load(f)
    return {}


def _extract_items(text: str) -> list:
    products = _load_products()
    items = []
    for m in _RE_ITEM.finditer(text):
        sku, qty = m.group(1), int(m.group(2))
        prod = products.get(sku, {})
        unit_price = float(prod.get("vip_price", 0))
        items.append(ItemDict(
            sku=sku,
            qty=qty,
            name=prod.get("name", sku),
            brand=prod.get("brand", ""),
            material=prod.get("material", ""),
            spec=prod.get("spec", ""),
            origin=prod.get("origin", "台灣"),
            unit_price=unit_price,
            subtotal=round(unit_price * qty, 2),
        ))
    return items


def _extract_name(text: str) -> str:
    m = _RE_AFTER_JI.search(text)
    if not m:
        return ""
    return m.group(1).strip("，,、 \t")


def _extract_phone(text: str) -> str:
    m = _RE_PHONE.search(text)
    return m.group(0) if m else ""


def _extract_address(text: str) -> str:
    m = _RE_ADDRESS_AFTER_PHONE.search(text)
    if not m:
        return ""
    addr = m.group(1).strip()
    # Trim trailing noise
    addr = re.split(r"[。\n]", addr)[0].strip()
    return addr


# ─── CLI test ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    sample = "1084043-3件，1084065-4件，1000044-2件，寄黃业偉 18125989028 廣東省深圳市龍崗區南聯劉屋村南段74號1樓"
    import pprint
    pprint.pprint(parse_order(sample))
