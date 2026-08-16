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
# --------------------------------------------------------------------------------------
TARGET_ITEMS = {
    "50MT004BB",   # Name Plate (Aluminium)
    "52CL003BB",   # CLOSURE 12 C
    "52CL004BB",   # CLOSURE 24 C
    "52CL006BB",   # CLOSURE 48 C
    "52CL009BB",   # CLOSURE 12 C FOR OFC DROP WIRE (IN LINE)
    "52CL010BB",   # CLOSURE 60 C
    "53OF150BB",   # OPTICAL FIBER DROP CABLE 1C, FLAT TYPE
    "53OF157BB",   # ARSS OPTICAL FIBER CABLE 12c-FIBRE3
    "53OF158BB",   # ARSS OPTICAL FIBER CABLE 24c-FIBRE3
    "53OF160BB",   # ARSS OPTICAL FIBER CABLE 60c-FIBRE3
}

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
    "Optimize": {"emoji": "🗃️", "style": "good", "color": "good", "label": "OPTIMIZE — Allocation Plan"},
}

# Status ในไฟล์ -> emoji + สีของ Adaptive Card
# ค่าที่พบจริงในไฟล์: "ปกติ", "จัดสรรเกิน (ไม่มีการใช้)",
#                    "ใช้เกินยอดรับ (ดึงสต็อกเดิม)", "ยอดไม่ตรง (รับนอกระบบ)",
#                    "- ไม่มีข้อมูลรอบก่อน"
# statusShort = ป้ายสั้นสำหรับตารางในการ์ด (ข้อความเต็มยาวเกินจนการ์ดชนเพดาน 28 KB)
STATUS_THEME = [
    ("ปกติ", "✅", "good", "ปกติ"),
    ("จัดสรรเกิน", "🟡", "warning", "จัดสรรเกิน"),
    ("ใช้เกินยอดรับ", "🔴", "attention", "ใช้เกินยอดรับ"),
    ("ยอดไม่ตรง", "🟠", "warning", "ยอดไม่ตรง"),
    ("ไม่มีข้อมูลรอบก่อน", "⚪", "default", "ไม่มีรอบก่อน"),
]

TH_MONTH = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
            "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

# เพดานแถวที่แสดงบนการ์ด — Teams ปฏิเสธ Adaptive Card ที่ใหญ่เกิน 28 KB
# วัดจริงด้วย tools/test_render_script.js จากข้อมูลในไฟล์:
#   MA       10 แถว = 23.2 KB | 12 แถว = 25.6 KB
#   Optimize 10 แถว = 25.0 KB   <- กรณีหนักสุดที่ยังปลอดภัย (5 Zone)
#   Optimize 11 แถว = 26.3 KB   <- เริ่มเฉียด
#   Optimize 12 แถว = 27.5 KB   <- เฉียดเพดานมาก
TOP_ROWS_DEFAULT = 10
TOP_ROWS_MAX = 10


def is_target_item(code):
    """คืน True เมื่อรหัสวัสดุอยู่ในขอบเขตที่การ์ดต้องแสดง"""
    return str(code or "").strip().upper() in TARGET_ITEMS


def status_theme(status):
    s = str(status or "").strip()
    for key, emoji, color, short in STATUS_THEME:
        if key in s:
            return emoji, color, short
    return "⚪", "default", "ไม่ระบุ"


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
    rows = []
    for r in values[hdr_idx + 1:]:
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
    rows = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < HEADER_ROW:
            continue
        if i == HEADER_ROW:
            header = [str(c).strip() if c is not None else "" for c in r]
            continue
        rec = dict(zip(header, r))
        if rec.get("Item Code"):
            rows.append(rec)
    wb.close()
    return rows


def build_payload(rows, period=None, dashboard_url="", source_url="", top=TOP_ROWS_DEFAULT):
    """แปลงแถวดิบ -> payload ที่ Adaptive Card template ใช้ผูกข้อมูล

    top = จำนวนแถวสูงสุดที่จะ "แสดง" บนการ์ด (เรียงรายการเสี่ยงขึ้นก่อน)
          Adaptive Card ใน Teams มีเพดานขนาด 28 KB ถ้าใส่ครบทุกแถวการ์ดจะไม่ถูกส่ง
          ที่เหลือสรุปเป็นบรรทัด "และอีก N รายการ" พร้อมปุ่มไป Dashboard
    """

    picked = [r for r in rows if is_target_item(r.get("Item Code"))]

    # ถ้าไม่ระบุรอบ ให้ใช้รอบล่าสุดของฝั่ง MA ตามลำดับ Period_Seq
    if period is None:
        ma = [r for r in picked if str(r.get("Data_Set")) == "MA"]
        if ma:
            period = max(ma, key=lambda r: (num(r.get("Period_Seq")),))["Distribution period"]

    picked = [r for r in picked if str(r.get("Distribution period")) == str(period)]

    data_set = str(picked[0].get("Data_Set")) if picked else "MA"
    theme = DATASET_THEME.get(data_set, DATASET_THEME["MA"])
    prev_period = ""
    for r in picked:
        if r.get("Prev_Period"):
            prev_period = str(r["Prev_Period"])
            break

    # ---- แปลงทุกแถวเป็น item ก่อน แล้วค่อยตัด top N ----
    all_items = []
    risk_total = watch_total = ok_total = 0

    for r in picked:
        zone = str(r.get("Zone") or "-").strip()
        emoji, color, short = status_theme(r.get("Status"))
        risk = str(r.get("Risk_Flag") or "").strip()
        is_risk = "เสี่ยง" in risk
        if is_risk:
            risk_total += 1
        elif risk:
            watch_total += 1
        else:
            ok_total += 1

        item = OrderedDict([
            ("itemCode", str(r.get("Item Code") or "").strip()),
            ("itemName", str(r.get("Description") or r.get("Art No") or "").strip()[:45]),
            ("unit", str(r.get("Unit") or "").strip()),
            ("province", str(r.get("Province") or "-").strip() or "-"),
            ("dataSet", data_set),
            ("period", str(period)),
            ("prevPeriod", str(r.get("Prev_Period") or "-")),
            ("onhand", fmt_num(r.get("OMC-Onhand"))),
            ("fromHub", fmt_int(r.get("เบิกจาก Hub"))),
            ("prevBefore", fmt_num(r.get("Prev_Before"))),
            ("prevReceived", fmt_num(r.get("Prev_Received"))),
            ("status", str(r.get("Status") or "-").strip().lstrip("- ").strip() or "-"),
            ("statusShort", short),
            ("statusEmoji", emoji),
            ("statusColor", color),
            ("riskFlag", risk or "-"),
            ("isRisk", is_risk),
            ("rowWeight", "bolder" if is_risk else "default"),
            ("zone", zone),
            ("zoneEmoji", ZONE_EMOJI.get(zone, "🔷")),
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

    zone_list = []
    for zone in sorted(grouped.keys()):
        items = grouped[zone]
        n_risk = zone_risk_totals.get(zone, 0)
        n_all = zone_totals.get(zone, len(items))
        zone_list.append(OrderedDict([
            ("zone", zone),
            ("zoneEmoji", ZONE_EMOJI.get(zone, "🔷")),
            ("itemCount", n_all),
            ("shownCount", len(items)),
            ("riskCount", n_risk),
            ("riskBadge", "⚠️ เสี่ยงขาด {} รายการ".format(n_risk) if n_risk else "✅ ไม่มีรายการเสี่ยง"),
            ("riskColor", "attention" if n_risk else "good"),
            ("countLabel", "{} รายการ".format(n_all) if len(items) == n_all
                           else "แสดง {} จาก {} รายการ".format(len(items), n_all)),
            ("items", items),
        ]))

    cutoff = ""

    meta = OrderedDict([
        ("title", "Report Material Optimize&MA_NER"),
        ("subtitle", ""),
        ("period", str(period or "-")),
        ("prevPeriod", prev_period or "-"),
        ("dataSet", data_set),
        ("dataSetLabel", theme["label"]),
        ("dataSetEmoji", theme["emoji"]),
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
        ("moreText", "…และอีก {} รายการ — กด “📊 Dashboard” เพื่อดูทั้งหมด".format(more_count)
                     if more_count > 0 else ""),
        ("dashboardUrl", dashboard_url or "https://app.powerbi.com/CHANGE-ME"),
        ("sourceFileUrl", source_url or "https://contoso.sharepoint.com/CHANGE-ME"),
    ])

    flat = [OrderedDict(it) for it in shown]

    return (
        OrderedDict([("meta", meta), ("zones", zone_list)]),
        OrderedDict([("meta", meta), ("items", flat)]),
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
    empty["zones"] = []

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
