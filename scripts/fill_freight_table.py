# -*- coding: utf-8 -*-
import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")

FREIGHT_TABLE = r"C:\Users\user\Desktop\順丰E順递\裝箱單給珠海模版\運費表.xlsx"

freight_map = {
    "1084093": 21,
    "1084066": 13,
    "1084090": 13,
    "1084099": 13,
    "1084092": 21,
    "0100345": 23,
    "0100346": 23,
    "1084064": 13,
    "1084065": 13,
    "0300525": 13,
    "0300523": 13,
    "0300648": 13,
    "0300432": 13,
    "0300270": 13,
    "0300283": 13,
    "1084088": 13,
    "1084043": 18,
    "1084080": 13,
    "1084063": 13,
    "1084085": 38,
    "1084078": 13,
    "1084074": 13,
    "1084083": 21,
    "1084067": 13,
    "1084086": 13,
    "1084069": 13,
    "1084082": 13,
    "0300616": 13,
    "0300614": 13,
    "0300416": 13,
    "1000044": 13,
    "1000458": 13,
    "1084077": 13,
    "1084094": 13,
    "1084084": 21,
}

wb = openpyxl.load_workbook(FREIGHT_TABLE)
ws = wb.active

filled = 0
for r in range(2, ws.max_row + 1):
    code = str(ws.cell(r, 1).value or "").strip()
    if code in freight_map:
        ws.cell(r, 3).value = freight_map[code]
        filled += 1

wb.save(FREIGHT_TABLE)
print(f"✅ 完成：{filled} 個產品已填入運費")

import os
os.startfile(FREIGHT_TABLE)
input("\n按 Enter 結束...")
