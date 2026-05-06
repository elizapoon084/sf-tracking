# -*- coding: utf-8 -*-
import os

# ─── Base paths ───────────────────────────────────────────────────────────────
BASE_DIR     = r"C:\Users\user\Desktop\順丰E順递"
SCRIPTS_DIR  = os.path.join(BASE_DIR, "scripts")
IMAGES_DIR   = os.path.join(BASE_DIR, "Images")
DATA_DIR     = os.path.join(BASE_DIR, "data")
LOGS_DIR     = os.path.join(BASE_DIR, "logs")

PRODUCTS_JSON = os.path.join(DATA_DIR, "products.json")
EXCEL_PATH    = os.path.join(DATA_DIR, "tracking.xlsx")

# ─── Chrome ───────────────────────────────────────────────────────────────────
CHROME_PROFILE = r"C:\ChromeAutomation"
BROWSER_ARGS   = [
    "--disable-blink-features=AutomationControlled",
]
PLAYWRIGHT_TIMEOUT  = 30_000   # ms
PLAYWRIGHT_SLOW_MO  = 150      # ms – reduces flakiness on fast machines

# ─── POS ──────────────────────────────────────────────────────────────────────
POS_URL           = "https://online-store-99126206.web.app/"
POS_ADMIN_PASS    = "0000"
POS_VIP_PASS      = "941196"

# ─── SF Express ───────────────────────────────────────────────────────────────
SF_SHIP_URL      = "https://hk.sf-express.com/hk/tc/ship/home"
SF_CLEARANCE_URL = "https://hk.sf-express.com/hk/tc/clearance"
SF_WAYBILL_URL   = "https://hk.sf-express.com/hk/tc/waybill/list"
SF_CHINA_LIST_URL = "https://www.sf-express.com/chn/sc/waybill/list"

# Fixed sender
SF_SENDER_NAME = "Eliza poon"
SF_ACCOUNT_NO  = "8526937071"

# Payment mode: "cod" (到付, testing) or "monthly" (寄付月結, production)
SF_PAYMENT_MODE = "cod"

# ─── Excel columns (1-indexed for openpyxl) ───────────────────────────────────
COL_DATE         = 1
COL_NAME         = 2
COL_POS_ORDER    = 3
COL_WAYBILL      = 4
COL_RECIPIENT    = 5
COL_PHONE        = 6
COL_ADDRESS      = 7
COL_ITEMS        = 8
COL_QTY          = 9
COL_VIP_TOTAL    = 10
COL_PAYMENT      = 11
COL_FREIGHT      = 12
COL_STATUS       = 13
COL_STATUS_TIME  = 14
COL_ANOMALY      = 15
COL_PDF_PATH     = 16
COL_NOTES        = 17
COL_TAX          = 18
COL_RECEIVED     = 19

EXCEL_HEADERS = [
    "日期", "客人名", "POS訂單號", "順丰運單號",
    "收件人", "收件電話", "收件地址",
    "貨品摘要", "件數", "VIP總額(HKD)",
    "付款方式", "運費(HKD)", "最新狀態",
    "狀態更新時間", "異常標記", "小票檔案路徑", "備註", "稅金(HKD)", "收件狀態",
]
EXCEL_SHEET = "追蹤表"

# Anomaly statuses that trigger red highlight
ANOMALY_KEYWORDS = ["退回", "異常", "卡關", "攔截", "問題件",
                    "异常", "问题件", "拦截"]  # simplified Chinese variants

# ─── Google Sheets (optional) ─────────────────────────────────────────────────
# After completing Google Sheets setup, put credentials JSON here and set sheet name.
GSHEETS_CREDENTIALS = os.path.join(DATA_DIR, "gsheets_credentials.json")
GSHEETS_SPREADSHEET = "順丰寄件追蹤"   # exact name of your Google Sheet
