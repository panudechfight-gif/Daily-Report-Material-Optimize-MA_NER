#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_sample_data.py
====================
อ่านไฟล์ Optimize_HiringMA_Full_Rebuild_VBA_Macro_V2.xlsm (ชีต "MA&Optimize NER")
แล้วสร้างไฟล์ JSON data payload สำหรับ Adaptive Card 3 แบบ:

    adaptive-cards/data/sample-data-main.json    -> การ์ดหลัก (จัดกลุ่มตาม Zone)
    adaptive-cards/data/sample-data-flat.json    -> การ์ดแบบง่าย (ไม่จัดกลุ่ม)
    adaptive-cards/data/sample-data-empty.json   -> การ์ดตอนไม่มีข้อมูล (empty state)

สคริปต์นี้คือ "สัญญาข้อมูล" (data contract) ตัวจริง — ตรรกะเดียวกันนี้ถูกพอร์ต
ไปไว้ใน office-scripts/buildCardPayload.ts เพื่อให้ Power Automate เรียกใช้ได้

วิธีใช้:
    python3 tools/build_sample_data.py <path-to-xlsm> [--period "MA (17 Aug 26)"]
"""

import argparse
import json
import math
import os
import sys
from collections import OrderedDict

try:
    import openpyxl
except ImportError:
    openpyxl = None       # ยังใช้ --from-dump ได้ ถ้าไม่มี openpyxl

# --------------------------------------------------------------------------------------
# รหัสวัสดุที่การ์ดต้องแสดง — รายการตายตัว 10 รหัส
# ต้องตรงกับ TARGET_ITEMS ใน office-scripts/buildCardPayload.ts เสมอ
# แถวข้อมูลแสดงเป็น "icon · short" — 1 icon ต่อ 1 group เท่านั้น
# --------------------------------------------------------------------------------------
TARGET_ITEMS = {
    "50MT004BB": {"icon": "\U0001F3F7\uFE0F", "group": "NAME PLATE", "short": "Name Plate Aluminium",
                  "name": "Name Plate (Aluminium)", "about": "ป้ายชื่อติดอุปกรณ์"},

    "52CL003BB": {"icon": "\U0001F9F0", "group": "CLOSURE", "short": "Closure 12C",
                  "name": "CLOSURE 12 C", "about": "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic"},
    "52CL004BB": {"icon": "\U0001F9F0", "group": "CLOSURE", "short": "Closure 24C",
                  "name": "CLOSURE 24 C", "about": "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic"},
    "52CL006BB": {"icon": "\U0001F9F0", "group": "CLOSURE", "short": "Closure 48C",
                  "name": "CLOSURE 48 C", "about": "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic"},
    "52CL009BB": {"icon": "\U0001F9F0", "group": "CLOSURE", "short": "Closure 12C Inline",
                  "name": "CLOSURE 12 C FOR OFC DROP WIRE (IN LINE)",
                  "about": "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic"},
    "52CL010BB": {"icon": "\U0001F9F0", "group": "CLOSURE", "short": "Closure 60C",
                  "name": "CLOSURE 60 C", "about": "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic"},

    "53OF150BB": {"icon": "\U0001F50C", "group": "DROP CABLE", "short": "Drop Cable 1C SC/UPC · 3m",
                  "name": "OPTICAL FIBER DROP CABLE 1C, FLAT TYPE (G.657A) WITH 2 SC/UPC PRE-CONNECTOR, 3m.",
                  "about": "สาย Pigtail Patch สำหรับ Splice เชื่อมต่อ"},

    "53OF157BB": {"icon": "\U0001F9F6", "group": "ARSS FIBER CABLE", "short": "ARSS Fiber Cable 12C",
                  "name": "ARSS OPTICAL FIBER CABLE 12c-FIBRE3", "about": "สาย Cable Fiber Optic"},
    "53OF158BB": {"icon": "\U0001F9F6", "group": "ARSS FIBER CABLE", "short": "ARSS Fiber Cable 24C",
                  "name": "ARSS OPTICAL FIBER CABLE 24c-FIBRE3", "about": "สาย Cable Fiber Optic"},
    "53OF160BB": {"icon": "\U0001F9F6", "group": "ARSS FIBER CABLE", "short": "ARSS Fiber Cable 60C",
                  "name": "ARSS OPTICAL FIBER CABLE 60c-FIBRE3", "about": "สาย Cable Fiber Optic"},
}

# หัวตารางที่ยอมรับได้ของแต่ละคอลัมน์ (ชีตเคยเปลี่ยนชื่อมาแล้ว) — ชื่อใหม่อยู่หน้าสุด
COL_ALIASES = {
    "Onhand": ["Onhand", "OMC-Onhand"],
    "จัดสรร(กระจาย)": ["จัดสรร(กระจาย)", "จัดสรร"],
    "จัดสรร(เบิก Hub)": ["จัดสรร(เบิก Hub)", "เบิกจาก Hub"],
}


def pick(rec, std):
    """อ่านค่าจาก record ตามชื่อมาตรฐาน โดยลองชื่อเดิมให้ด้วย"""
    for alias in COL_ALIASES.get(std, [std]):
        if alias in rec:
            return rec.get(alias)
    return None

# คอลัมน์บล็อกตัวเลขที่ข้อมูลในชีตเลื่อนไปทางขวา 1 ช่องจากหัวตารางของตัวเอง
# (ชีตมีคอลัมน์ว่างไม่มีหัวแทรกอยู่หลัง OMC-Onhand)
# ตรวจจับด้วยกฎ  จัดสรร + เบิกจาก Hub + จัดสรรเพิ่ม = รวม 3 รายการ  ใน fix_header_shift()
SHIFT_COLS = ["จัดสรร(กระจาย)", "จัดสรร(เบิก Hub)", "จัดสรรเพิ่ม", "รวม 3 รายการ"]

SHEET_SOURCE = "MA&Optimize NER"
HEADER_ROW = 7  # ชีต MA&Optimize NER วางหัวตารางไว้แถวที่ 7

# emoji ประจำ Zone — ใช้ให้แยกกลุ่มออกจากกันด้วยสายตาบนมือถือ
ZONE_EMOJI = {
    "RC2-NMA": "🟦",
    "RC2-UBN": "🟩",
    "RC3-KKN": "🟨",
    "RC3-UDN": "🟧",
    "RC3-SNK": "🟪",
}

# Data_Set -> ชุดสี/emoji ของหัวการ์ด (emoji ชุดคลังวัสดุ)
DATASET_THEME = {
    "MA": {"emoji": "📦", "style": "accent", "color": "accent", "label": "MA — Weekly Allocation"},
    "Optimize": {"emoji": "🗃️", "style": "good", "color": "good", "label": "Optimize — Allocation Plan"},
}

# Status ในไฟล์ -> emoji + สีของ Adaptive Card
# ค่าที่พบจริงในไฟล์: "ปกติ", "จัดสรรเกิน (ไม่มีการใช้)",
#                    "ใช้เกินยอดรับ (ดึงสต็อกเดิม)", "ยอดไม่ตรง (รับนอกระบบ)",
#                    "- ไม่มีข้อมูลรอบก่อน"
# statusShort = "" แปลว่าไม่ต้องพิมพ์คำอธิบายซ้ำบนการ์ด (emoji สื่อความหมายพอแล้ว)
# ตัดคำอธิบายของ "จัดสรรเกิน" และ "ไม่มีรอบก่อน" ออก เพื่อเหลือที่ให้ข้อมูลจริง
STATUS_THEME = [
    ("ปกติ", "✅", "good", ""),
    ("จัดสรรเกิน", "🟡", "warning", ""),
    ("ใช้เกินยอดรับ", "🔴", "attention", "ใช้เกินยอดรับ"),
    ("ยอดไม่ตรง", "🟠", "warning", "ยอดไม่ตรง"),
    ("ไม่มีข้อมูลรอบก่อน", "⚪", "default", ""),
]

# ลำดับความรุนแรง ใช้เลือกสถานะตัวแทนตอนรวมหลายจังหวัดเป็นแถวเดียว
STATUS_RANK = [
    ("ใช้เกินยอดรับ", 4),
    ("ยอดไม่ตรง", 3),
    ("จัดสรรเกิน", 2),
    ("ปกติ", 1),
]

TH_MONTH = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
            "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

# เพดานแถวที่แสดงบนการ์ด — Teams ปฏิเสธ Adaptive Card ที่ใหญ่เกิน 28 KB
# การ์ดรวมแถวซ้ำของแต่ละ Zone แล้ว (1 Zone + 1 Item Code = 1 แถว)
# จำนวนแถวจริงจึงน้อยกว่าจำนวนแถวดิบมาก และครอบคลุมทั้งรอบได้ในการ์ดใบเดียว
TOP_ROWS_DEFAULT = 40
TOP_ROWS_MAX = 40


def is_target_item(code):
    """คืน True เมื่อรหัสวัสดุอยู่ในขอบเขตที่การ์ดต้องแสดง"""
    return str(code or "").strip().upper() in TARGET_ITEMS


def status_theme(status):
    s = str(status or "").strip()
    for key, emoji, color, short in STATUS_THEME:
        if key in s:
            return emoji, color, short
    return "⚪", "default", ""


def status_rank(status):
    s = str(status or "").strip()
    for key, rank in STATUS_RANK:
        if key in s:
            return rank
    return 0


def fix_header_shift(header, data_rows):
    """แก้แนวคอลัมน์เลื่อนในชีตต้นทาง

    ชีต MA&Optimize NER มีคอลัมน์ว่างที่ไม่มีหัวตารางแทรกอยู่หลัง OMC-Onhand
    ข้อมูลของบล็อก จัดสรร / เบิกจาก Hub / จัดสรรเพิ่ม / รวม 3 รายการ
    จึงไปอยู่ทางขวาของหัวตารางตัวเอง 1 ช่อง ถ้าอ่านตามชื่อหัวตารางตรง ๆ
    ค่าที่ได้ในช่อง "เบิกจาก Hub" จะเป็นค่าของ "จัดสรร" แทน

    ตรวจด้วยกฎที่ชีตคำนวณไว้เอง: จัดสรร + เบิกจาก Hub + จัดสรรเพิ่ม = รวม 3 รายการ
    ถ้าวันหนึ่งชีตถูกแก้ให้ตรงแล้ว ฟังก์ชันนี้จะคืนหัวตารางเดิมโดยไม่เลื่อนอะไร
    """
    idx = {}
    for i, name in enumerate(header):
        if name and name not in idx:
            idx[name] = i
    if not all(n in idx for n in SHIFT_COLS):
        return header

    for off in (0, 1):
        checked = agree = 0
        for r in data_rows[:300]:
            ia, ib = idx["จัดสรร(กระจาย)"] + off, idx["จัดสรร(เบิก Hub)"] + off
            ic, isum = idx["จัดสรรเพิ่ม"] + off, idx["รวม 3 รายการ"] + off
            if isum >= len(r):
                continue
            checked += 1
            if abs(num(r[ia]) + num(r[ib]) + num(r[ic]) - num(r[isum])) < 1e-9:
                agree += 1
        if checked and agree == checked:
            if off == 0:
                return header
            fixed = list(header)
            for n in SHIFT_COLS:            # ปลดชื่อเดิมออกก่อน
                fixed[idx[n]] = "_spacer_" + n
            for n in SHIFT_COLS:            # แล้วค่อยวางชื่อไว้ตำแหน่งที่ข้อมูลอยู่จริง
                if idx[n] + 1 < len(fixed):
                    fixed[idx[n] + 1] = n
            return fixed
    return header


def num(v):
    """แปลงค่าเป็นตัวเลข; ค่าว่าง/ข้อความ -> 0"""
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def fmt_num(v):
    """จัดรูปแบบตัวเลขให้อ่านง่ายบนการ์ด: มีคอมมา, ตัด .0 ทิ้ง, ทศนิยมไม่เกิน 2"""
    n = num(v)
    if abs(n - round(n)) < 1e-9:
        return "{:,}".format(int(round(n)))
    return "{:,.2f}".format(n)


def fmt_int(v):
    """จัดรูปแบบเป็นจำนวนเต็มเสมอ — ใช้กับ 'เบิกจาก Hub' ที่ต้องเป็นยอดจ่ายจริง

    ปัดครึ่งขึ้นแบบคณิตศาสตร์ให้ตรงกับ Math.round() ของ TypeScript
    (Python round() ปัดครึ่งไปหาเลขคู่ ซึ่งให้ผลต่างกันที่ .5)
    """
    n = num(v)
    rounded = int(math.floor(abs(n) + 0.5))
    return "{:,}".format(-rounded if n < 0 else rounded)


def thai_date(dt):
    if dt is None:
        return ""
    try:
        return "{} {} {}".format(dt.day, TH_MONTH[dt.month], dt.year + 543)
    except Exception:
        return str(dt)


def read_rows_from_dump(dump_path):
    """อ่านแถวจาก tools/.sheet-values.json แทนไฟล์ .xlsm

    ใช้เมื่ออยู่ในเครื่องที่ไม่มีไฟล์ Excel ต้นฉบับ (เช่น CI)
    ไฟล์ dump สร้างด้วย tools/dump_sheet_values.py
    """
    with open(dump_path, encoding="utf-8") as f:
        dump = json.load(f)

    values = dump["values"]
    hdr_idx = HEADER_ROW - 1 - dump["firstRowIndex"]
    if hdr_idx < 0 or hdr_idx >= len(values):
        sys.exit("หาหัวตารางแถวที่ {} ในไฟล์ dump ไม่เจอ".format(HEADER_ROW))

    header = [str(c).strip() if c is not None else "" for c in values[hdr_idx]]
    raw = values[hdr_idx + 1:]
    header = fix_header_shift(header, raw)
    rows = []
    for r in raw:
        rec = dict(zip(header, r))
        if rec.get("Item Code"):
            rows.append(rec)
    return rows


def read_rows(xlsm_path):
    if openpyxl is None:
        sys.exit("ต้องติดตั้ง openpyxl ก่อน:  pip install openpyxl")
    wb = openpyxl.load_workbook(xlsm_path, data_only=True, read_only=True)
    if SHEET_SOURCE not in wb.sheetnames:
        sys.exit("ไม่พบชีต '{}' ในไฟล์".format(SHEET_SOURCE))
    ws = wb[SHEET_SOURCE]

    header = None
    raw = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < HEADER_ROW:
            continue
        if i == HEADER_ROW:
            header = [str(c).strip() if c is not None else "" for c in r]
            continue
        raw.append(r)
    wb.close()

    header = fix_header_shift(header, raw)
    rows = []
    for r in raw:
        rec = dict(zip(header, r))
        if rec.get("Item Code"):
            rows.append(rec)
    return rows


def build_payload(rows, period=None, dashboard_url="", source_url="", top=TOP_ROWS_DEFAULT):
    """แปลงแถวดิบ -> payload ที่ Adaptive Card template ใช้ผูกข้อมูล

    top = จำนวนแถวสูงสุดที่จะ "แสดง" บนการ์ด (เรียงรายการเสี่ยงขึ้นก่อน)
          Adaptive Card ใน Teams มีเพดานขนาด 28 KB ถ้าใส่ครบทุกแถวการ์ดจะไม่ถูกส่ง
          ที่เหลือสรุปเป็นบรรทัด "และอีก N รายการ" พร้อมปุ่มไป Dashboard
    """

    picked = [r for r in rows if is_target_item(r.get("Item Code"))]

    # ถ้าไม่ระบุรอบ ให้ใช้รอบล่าสุดของชุด Optimize (รายงานหลัก) ตามลำดับ Period_Seq
    # ส่งชื่อชุดข้อมูลมาแทนชื่อรอบก็ได้ = "รอบล่าสุดของชุดนั้น"
    want_set = "Optimize"
    if period is not None and str(period).strip().lower() in ("ma", "optimize"):
        want_set = "MA" if str(period).strip().lower() == "ma" else "Optimize"
        period = None
    if period is None:
        pool = [r for r in picked if str(r.get("Data_Set")) == want_set] or picked
        if pool:
            period = max(pool, key=lambda r: (num(r.get("Period_Seq")),))["Distribution period"]

    picked = [r for r in picked if str(r.get("Distribution period")) == str(period)]

    data_set = str(picked[0].get("Data_Set")) if picked else "Optimize"
    theme = DATASET_THEME.get(data_set, DATASET_THEME["Optimize"])
    prev_period = ""
    for r in picked:
        if r.get("Prev_Period"):
            prev_period = str(r["Prev_Period"])
            break

    # ---- รวมแถวซ้ำ: 1 Zone + 1 Item Code = 1 แถว ----
    # ชีตต้นทางแตกแถวตามจังหวัด รหัสเดียวกันจึงโผล่หลายครั้งในโซนเดียว
    # การ์ดไม่ได้แสดงจังหวัด จึงบวกยอดทุกจังหวัดเข้าด้วยกันแล้วเก็บจำนวนจังหวัดไว้
    agg = OrderedDict()
    for r in picked:
        zone = str(r.get("Zone") or "-").strip()
        code = str(r.get("Item Code") or "").strip().upper()
        key = (zone, code)
        a = agg.get(key)
        if a is None:
            a = {
                "zone": zone, "itemCode": code,
                "itemName": (TARGET_ITEMS[code]["name"] if code in TARGET_ITEMS
                             else str(r.get("Description") or r.get("Art No") or "").strip()),
                "itemIcon": TARGET_ITEMS[code]["icon"] if code in TARGET_ITEMS else "\u25AB\uFE0F",
                "itemShort": TARGET_ITEMS[code]["short"] if code in TARGET_ITEMS else code,
                "unit": str(r.get("Unit") or "").strip(),
                "onhand": 0.0, "distribute": 0.0, "fromHub": 0.0,
                "prevBefore": 0.0, "prevReceived": 0.0,
                "provinces": [], "status": "", "statusRank": -1,
                "isRisk": False, "riskFlag": "",
            }
            agg[key] = a

        a["onhand"] += num(pick(r, "Onhand"))
        a["distribute"] += num(pick(r, "จัดสรร(กระจาย)"))
        a["fromHub"] += num(pick(r, "จัดสรร(เบิก Hub)"))
        a["prevBefore"] += num(r.get("Prev_Before"))
        a["prevReceived"] += num(r.get("Prev_Received"))

        prov = str(r.get("Province") or "").strip()
        if prov and prov not in a["provinces"]:
            a["provinces"].append(prov)

        st = str(r.get("Status") or "").strip()
        rank = status_rank(st)
        if rank > a["statusRank"]:
            a["statusRank"] = rank
            a["status"] = st

        risk = str(r.get("Risk_Flag") or "").strip()
        if "เสี่ยง" in risk:
            a["isRisk"] = True
            a["riskFlag"] = risk
        elif risk and not a["riskFlag"]:
            a["riskFlag"] = risk

    all_items = []
    risk_total = watch_total = ok_total = 0

    for a in agg.values():
        emoji, color, short = status_theme(a["status"])
        if a["isRisk"]:
            risk_total += 1
        elif a["riskFlag"]:
            watch_total += 1
        else:
            ok_total += 1

        n_prov = len(a["provinces"])
        item = OrderedDict([
            ("itemCode", a["itemCode"]),
            ("itemName", a["itemName"][:45]),
            ("itemIcon", a["itemIcon"]),
            ("itemShort", a["itemShort"]),
            ("unit", a["unit"]),
            ("province", "{} จังหวัด".format(n_prov) if n_prov else "-"),
            ("provinceCount", n_prov),
            # ต่อท้ายรหัสเฉพาะตอนที่ยอดนี้รวมมาจากหลายจังหวัด จะได้รู้ว่าตัวเลขมาจากไหน
            ("provinceTag", " ·{}จว.".format(n_prov) if n_prov > 1 else ""),
            ("dataSet", data_set),
            ("period", str(period)),
            ("prevPeriod", prev_period or "-"),
            ("onhand", fmt_num(a["onhand"])),
            ("distribute", fmt_int(a["distribute"])),
            ("fromHub", fmt_int(a["fromHub"])),
            ("numbersLine", "Onhand {} · กระจาย {} · Hub {}".format(
                fmt_num(a["onhand"]), fmt_int(a["distribute"]), fmt_int(a["fromHub"]))),
            ("prevBefore", fmt_num(a["prevBefore"])),
            ("prevReceived", fmt_num(a["prevReceived"])),
            ("status", str(a["status"] or "-").strip().lstrip("- ").strip() or "-"),
            ("statusShort", short),
            ("statusEmoji", emoji),
            ("statusColor", color),
            ("riskFlag", a["riskFlag"] or "-"),
            ("isRisk", a["isRisk"]),
            ("rowWeight", "bolder" if a["isRisk"] else "default"),
            ("zone", a["zone"]),
            ("zoneEmoji", ZONE_EMOJI.get(a["zone"], "🔷")),
        ])
        all_items.append(item)

    # เรียงลำดับความสำคัญ: เสี่ยงขาดก่อน -> Onhand น้อยก่อน -> รหัสวัสดุ -> Zone
    def rank_key(x):
        return (not x["isRisk"], num(x["onhand"].replace(",", "")), x["itemCode"])

    all_items.sort(key=lambda x: rank_key(x) + (x["zone"],))

    # นับจำนวนเต็มของแต่ละ Zone ไว้แสดงในหัวกลุ่ม (ไม่ใช่แค่จำนวนที่ถูกตัดมาแสดง)
    zone_totals = OrderedDict()
    zone_risk_totals = OrderedDict()
    for it in all_items:
        zone_totals[it["zone"]] = zone_totals.get(it["zone"], 0) + 1
        if it["isRisk"]:
            zone_risk_totals[it["zone"]] = zone_risk_totals.get(it["zone"], 0) + 1

    # เลือกแถววนรอบทีละ Zone — เอาอันดับ 1 ของทุก Zone ก่อน แล้วค่อยวนอันดับ 2
    # ทำให้ทุก Zone มีที่บนการ์ดเสมอ แม้รายการเสี่ยงจะกระจุกอยู่ Zone เดียว
    by_zone = OrderedDict()
    for it in all_items:
        by_zone.setdefault(it["zone"], []).append(it)
    zone_names = sorted(by_zone.keys())
    for z in zone_names:
        by_zone[z].sort(key=rank_key)

    limit = top if top and top > 0 else len(all_items)
    shown = []
    rank = 0
    while len(shown) < limit:
        added = False
        for z in zone_names:
            if len(shown) >= limit:
                break
            if rank < len(by_zone[z]):
                shown.append(by_zone[z][rank])
                added = True
        if not added:
            break
        rank += 1
    more_count = len(all_items) - len(shown)

    grouped = OrderedDict()
    for it in shown:
        grouped.setdefault(it["zone"], []).append(it)

    # ยอดรวมของแต่ละ Zone = บวกทุกจังหวัดในโซนนั้น (ไม่ใช่แค่แถวที่แสดง)
    zone_onhand = OrderedDict()
    zone_dist = OrderedDict()
    zone_hub = OrderedDict()
    for it in all_items:
        z = it["zone"]
        zone_onhand[z] = zone_onhand.get(z, 0.0) + num(it["onhand"].replace(",", ""))
        zone_dist[z] = zone_dist.get(z, 0.0) + num(it["distribute"].replace(",", ""))
        zone_hub[z] = zone_hub.get(z, 0.0) + num(it["fromHub"].replace(",", ""))

    zone_list = []
    for zone in sorted(grouped.keys()):
        items = grouped[zone]
        n_risk = zone_risk_totals.get(zone, 0)
        n_all = zone_totals.get(zone, len(items))
        t_onhand = fmt_num(zone_onhand.get(zone, 0))
        t_dist = fmt_int(zone_dist.get(zone, 0))
        t_hub = fmt_int(zone_hub.get(zone, 0))
        zone_list.append(OrderedDict([
            ("zone", zone),
            ("zoneEmoji", ZONE_EMOJI.get(zone, "🔷")),
            ("itemCount", n_all),
            ("shownCount", len(items)),
            ("riskCount", n_risk),
            # ไม่มีรายการเสี่ยง = ไม่ต้องขึ้นป้ายอะไรเลย ปล่อยว่างให้การ์ดโล่ง
            ("riskBadge", "⚠️ เสี่ยงขาด {}".format(n_risk) if n_risk else ""),
            ("riskColor", "attention" if n_risk else "good"),
            ("countLabel", "{} รายการ".format(n_all) if len(items) == n_all
                           else "แสดง {} จาก {} รายการ".format(len(items), n_all)),
            ("totalOnhand", t_onhand),
            ("totalDistribute", t_dist),
            ("totalHub", t_hub),
            ("totalLabel", "Onhand {} · กระจาย {} · เบิก Hub {}".format(
                t_onhand, t_dist, t_hub)),
            # หัวกลุ่มรวมเป็นบรรทัดเดียว — ประหยัดพื้นที่การ์ดได้ราว 0.5 KB ต่อ Zone
            ("zoneHeadline", "{} {} · {} · Onhand {} · กระจาย {} · เบิก Hub {}{}".format(
                ZONE_EMOJI.get(zone, "🔷"), zone,
                "{} รายการ".format(n_all) if len(items) == n_all
                else "แสดง {}/{}".format(len(items), n_all),
                t_onhand, t_dist, t_hub,
                " · ⚠️ เสี่ยงขาด {}".format(n_risk) if n_risk else "")),
            ("items", items),
        ]))

    # ตารางอธิบายท้ายการ์ด — ถอดรหัส icon + ชื่อย่อบนแถวข้อมูลกลับเป็น Item Code เต็ม
    # 1 บรรทัดต่อ "กลุ่ม" (CLOSURE / OPTICAL FIBER / NAME PLATE) ไม่ใช่ต่อรหัส
    legend = []
    legend_seen = {}
    for c in sorted({it["itemCode"] for it in all_items}):
        spec = TARGET_ITEMS.get(c)
        icon = spec["icon"] if spec else "▫️"
        group = spec["group"] if spec else "อื่น ๆ"
        about = spec["about"] if spec else ""
        short = spec["short"] if spec else c
        pair = c + "=" + short
        if group not in legend_seen:
            legend_seen[group] = len(legend)
            legend.append(OrderedDict([
                ("itemIcon", icon), ("groupLabel", group),
                ("itemAbout", about), ("pairs", pair)]))
        else:
            at = legend_seen[group]
            legend[at]["pairs"] = legend[at]["pairs"] + " · " + pair

    cutoff = ""

    meta = OrderedDict([
        ("title", "Report Stock Allocation Cable&Material NER"),
        ("subtitle", ""),
        ("hasLegend", len(legend) > 0),
        ("period", str(period or "-")),
        ("prevPeriod", prev_period or "-"),
        ("dataSet", data_set),
        ("dataSetLabel", theme["label"]),
        ("dataSetEmoji", "\U0001F3EC"),
        ("dataSetStyle", theme["style"]),
        ("dataSetColor", theme["color"]),
        ("cutoffDate", cutoff),
        ("generatedAt", ""),  # Power Automate เติมตอนรัน
        ("totalItems", len(all_items)),
        ("shownItems", len(shown)),
        ("totalZones", len(zone_totals)),
        ("riskCount", risk_total),
        ("watchCount", watch_total),
        ("okCount", ok_total),
        ("hasData", len(all_items) > 0),
        ("hasMore", more_count > 0),
        ("moreCount", more_count),
        ("moreText", ""),
        ("dashboardUrl", dashboard_url or "https://app.powerbi.com/CHANGE-ME"),
        ("sourceFileUrl", source_url or "https://mimotech-my.sharepoint.com/:x:/r/personal/panudeck_ais_co_th/Documents/Workflow%20Power%20Automate/JobMonitor_CMTRS_Job_Tracking.xlsx?d=we08d992568d54ab88ef622f150e1cd05&csf=1&web=1&e=8sLq79"),
    ])

    flat = [OrderedDict(it) for it in shown]

    return (
        OrderedDict([("meta", meta), ("zones", zone_list), ("legend", legend)]),
        OrderedDict([("meta", meta), ("items", flat), ("legend", legend)]),
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm", nargs="?",
                    help="ไฟล์ .xlsm ต้นทาง (เว้นได้ถ้าใช้ --from-dump)")
    ap.add_argument("--from-dump", default=None,
                    help="อ่านจาก tools/.sheet-values.json แทนไฟล์ .xlsm")
    ap.add_argument("--period", default=None)
    ap.add_argument("--risk-period", default="MA (10 Aug 26)",
                    help="รอบที่ใช้ทำตัวอย่าง 'มีรายการเสี่ยงขาด' สำหรับทดสอบสี")
    ap.add_argument("--top", type=int, default=TOP_ROWS_DEFAULT)
    ap.add_argument("--outdir", default="adaptive-cards/data")
    args = ap.parse_args()

    if args.from_dump:
        rows = read_rows_from_dump(args.from_dump)
    elif args.xlsm:
        rows = read_rows(args.xlsm)
    else:
        ap.error("ต้องระบุไฟล์ .xlsm หรือใช้ --from-dump")
    main_payload, flat_payload = build_payload(rows, args.period, top=args.top)
    risk_payload, _ = build_payload(rows, args.risk_period, top=args.top)

    os.makedirs(args.outdir, exist_ok=True)

    empty = json.loads(json.dumps(main_payload))
    empty["meta"].update({
        "hasData": False, "totalItems": 0, "shownItems": 0, "totalZones": 0,
        "riskCount": 0, "watchCount": 0, "okCount": 0,
        "hasMore": False, "moreCount": 0, "moreText": "",
    })
    empty["meta"]["hasLegend"] = False
    empty["zones"] = []
    empty["legend"] = []

    out = {
        "sample-data-main.json": main_payload,
        "sample-data-flat.json": flat_payload,
        "sample-data-risk.json": risk_payload,
        "sample-data-empty.json": empty,
    }
    for name, payload in out.items():
        path = os.path.join(args.outdir, name)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print("wrote {}".format(path))

    m = main_payload["meta"]
    print("\nรอบ            : {}".format(m["period"]))
    print("รอบก่อนหน้า     : {}".format(m["prevPeriod"]))
    print("Data_Set       : {}".format(m["dataSet"]))
    print("จำนวนรายการ     : {}  ({} Zone)".format(m["totalItems"], m["totalZones"]))
    print("เสี่ยงขาด/พอใช้/ปกติ : {}/{}/{}".format(m["riskCount"], m["watchCount"], m["okCount"]))


if __name__ == "__main__":
    main()
