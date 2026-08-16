#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dump_sheet_values.py
====================
ดึงค่าทั้งชีต "MA&Optimize NER" ออกมาเป็น JSON เพื่อใช้จำลอง ExcelScript.Workbook
ตอนทดสอบ office-scripts/buildCardPayload.ts นอก Excel

    python3 tools/dump_sheet_values.py <path-to-xlsm>
    -> tools/.sheet-values.json
"""

import json
import os
import sys
import datetime

try:
    import openpyxl
except ImportError:
    sys.exit("ต้องติดตั้ง openpyxl ก่อน:  pip install openpyxl")

SHEET = "MA&Optimize NER"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".sheet-values.json")


def cell(v):
    """แปลงค่าให้เหมือนที่ ExcelScript.Range.getValues() คืนมา: string | number | boolean"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def main():
    if len(sys.argv) < 2:
        sys.exit("ใช้: python3 tools/dump_sheet_values.py <path-to-xlsm>")

    wb = openpyxl.load_workbook(sys.argv[1], data_only=True, read_only=True)
    ws = wb[SHEET]

    rows = [[cell(c) for c in r] for r in ws.iter_rows(values_only=True)]
    wb.close()

    # getUsedRange() ของ Excel เริ่มที่แถวแรกที่มีข้อมูล — ชีตนี้มีข้อมูลตั้งแต่แถว 1
    payload = {"sheet": SHEET, "firstRowIndex": 0, "values": rows}

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)

    print("wrote {}  ({} rows x {} cols)".format(
        OUT, len(rows), max(len(r) for r in rows) if rows else 0))


if __name__ == "__main__":
    main()
