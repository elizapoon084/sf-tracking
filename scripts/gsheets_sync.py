# -*- coding: utf-8 -*-
"""
Google Sheets sync — pushes tracking.xlsx data to an online Google Sheet.

Setup (one-time):
  1. pip install gspread google-auth
  2. Create a Google Cloud project → enable Google Sheets API + Google Drive API
  3. Create a Service Account → download JSON key → save as data/gsheets_credentials.json
  4. Create a Google Sheet named "順丰寄件追蹤" (or set GSHEETS_SPREADSHEET in config.py)
  5. Share that sheet with the service account email (Editor access)

After setup, this module auto-syncs every time scheduled_update.py runs.
If credentials file is missing, it silently skips (no crash).
"""

import os
from datetime import datetime
from typing import Optional

from logger import get_logger
from config import GSHEETS_CREDENTIALS, GSHEETS_SPREADSHEET, EXCEL_HEADERS

log = get_logger(__name__)

# Extra column added to Google Sheet (not in local Excel)
_GS_HEADERS = EXCEL_HEADERS + ["最後爬取時間"]

# How to find the right row: match on waybill number (column index, 0-based)
_WB_COL_IDX = 3   # 4th column = 順丰運單號


def _open_sheet():
    """Return (gspread.Worksheet, header_row_values) or raise."""
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(GSHEETS_CREDENTIALS, scopes=scopes)
    gc    = gspread.authorize(creds)

    try:
        sh = gc.open(GSHEETS_SPREADSHEET)
    except gspread.SpreadsheetNotFound:
        log.info("Google Sheet '%s' not found — creating it", GSHEETS_SPREADSHEET)
        sh = gc.create(GSHEETS_SPREADSHEET)

    ws = sh.sheet1
    existing = ws.row_values(1)
    if existing != _GS_HEADERS:
        ws.clear()
        ws.append_row(_GS_HEADERS, value_input_option="USER_ENTERED")
    return ws


def sync_excel_to_sheets(excel_rows: list[list]) -> bool:
    """
    Push all Excel data rows to Google Sheets, upserting by waybill number.
    excel_rows: list of value lists (one per data row, matching EXCEL_HEADERS order).
    Returns True on success, False if Sheets not configured / unavailable.
    """
    if not os.path.exists(GSHEETS_CREDENTIALS):
        log.debug("gsheets_credentials.json not found — skipping Sheets sync")
        return False

    try:
        import gspread  # noqa: F401 (import check)
    except ImportError:
        log.warning("gspread not installed — run: pip install gspread google-auth")
        return False

    try:
        ws = _open_sheet()
        now = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Read existing sheet data for upsert logic
        all_vals = ws.get_all_values()  # includes header
        # Build {waybill: row_number_1indexed}
        wb_to_row: dict[str, int] = {}
        for i, row in enumerate(all_vals[1:], start=2):  # skip header
            wb = row[_WB_COL_IDX] if len(row) > _WB_COL_IDX else ""
            if wb:
                wb_to_row[wb] = i

        updates_batch = []   # (row_number, [values])
        appends       = []   # [values] rows to append

        for excel_row in excel_rows:
            values = list(excel_row) + [now]  # add 最後爬取時間
            # Pad to full width
            while len(values) < len(_GS_HEADERS):
                values.append("")
            values = [str(v) if v is not None else "" for v in values]

            wb = values[_WB_COL_IDX]
            if wb and wb in wb_to_row:
                updates_batch.append((wb_to_row[wb], values))
            else:
                appends.append(values)

        # Perform updates
        if updates_batch:
            import gspread
            from gspread.utils import rowcol_to_a1
            cell_list = []
            for row_num, values in updates_batch:
                for col_idx, val in enumerate(values, start=1):
                    cell_list.append(gspread.Cell(row_num, col_idx, val))
            ws.update_cells(cell_list, value_input_option="USER_ENTERED")

        # Append new rows
        if appends:
            ws.append_rows(appends, value_input_option="USER_ENTERED")

        log.info("Google Sheets sync: updated %d rows, appended %d new rows",
                 len(updates_batch), len(appends))
        return True

    except Exception as e:
        log.error("Google Sheets sync failed: %s", e)
        return False
