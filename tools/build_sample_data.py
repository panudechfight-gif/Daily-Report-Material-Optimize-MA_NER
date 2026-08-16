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
import os
import re
import sys
from collections import OrderedDict

try:
    import openpyxl
except ImportError:
    sys.exit("ต้องติดตั้ง openpyxl ก่อน:  pip install openpyxl")

# --------------------------------------------------------------------------------------
# กติกาการคัดเลือก Item Code — ตามที่ผู้ใช้กำหนด: เฉพาะ 53OFxxx และ 52CLx00xBB
# 53OF... = สายเคเบิลใยแก้ว (Optical Fiber Cable / Drop Cable / ADSS / ARSS)
# 52CL... = ตู้พักสาย (Closure)
# ในไฟล์จริงรหัส Closure เขียนเป็น 52CL + 3 หลัก + BB/AS เช่น 52CL009BB
# จึงใช้ regex ที่ครอบคลุมทั้งสองแบบ
# --------------------------------------------------------------------------------------
RE_CABLE = re.compile(r"^53OF\d{3}(BB|AS)$", re.IGNORECASE)
RE_CLOSURE = re.compile(r"^52CL\d{3}(BB|AS)$", re.IGNORECASE)

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

# Data_Set -> ชุดสี/emoji ของหัวการ์ด
DATASET_THEME = {
    "MA": {"emoji": "🛠️", "style": "accent", "color": "accent", "label": "MA (งานซ่อมบำรุง)"},
    "Optimize": {"emoji": "⚙️", "style": "good", "color": "good", "label": "Optimize (ปรับปรุงโครงข่าย)"},
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
#   MA       12 แถว = 23.0 KB
#   Optimize 12 แถว = 24.6 KB   <- กรณีหนักสุดที่ยังปลอดภัย
#   Optimize 13 แถว = 27.6 KB   <- เฉียดเพดาน
#   Optimize 14 แถว = 28.9 KB   <- Teams ปฏิเสธ
TOP_ROWS_DEFAULT = 12
TOP_ROWS_MAX = 12


def is_target_item(code):
    """คืน True เมื่อรหัสวัสดุอยู่ในขอบเขตที่การ์ดต้องแสดง"""
    c = str(code or "").strip()
    return bool(RE_CABLE.match(c) or RE_CLOSURE.match(c))


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


def thai_date(dt):
    if dt is None:
        return ""
    try:
        return "{} {} {}".format(dt.day, TH_MONTH[dt.month], dt.year + 543)
    except Exception:
        return str(dt)


def read_rows(xlsm_path):
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
            ("fromHub", fmt_num(r.get("เบิกจาก Hub"))),
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

    # เรียงลำดับความสำคัญ: เสี่ยงขาดก่อน -> Onhand น้อยก่อน -> Zone -> รหัสวัสดุ
    all_items.sort(key=lambda x: (not x["isRisk"], num(x["onhand"].replace(",", "")),
                                  x["zone"], x["itemCode"]))
    shown = all_items[:top] if top and top > 0 else all_items
    more_count = len(all_items) - len(shown)

    # นับจำนวนเต็มของแต่ละ Zone ไว้แสดงในหัวกลุ่ม (ไม่ใช่แค่จำนวนที่ถูกตัดมาแสดง)
    zone_totals = OrderedDict()
    zone_risk_totals = OrderedDict()
    for it in all_items:
        zone_totals[it["zone"]] = zone_totals.get(it["zone"], 0) + 1
        if it["isRisk"]:
            zone_risk_totals[it["zone"]] = zone_risk_totals.get(it["zone"], 0) + 1

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
        ("title", "รายงานวัสดุคงคลัง MA & Optimize — NER"),
        ("subtitle", "เฉพาะสายเคเบิล 53OF และตู้พักสาย 52CL"),
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
    ap.add_argument("xlsm")
    ap.add_argument("--period", default=None)
    ap.add_argument("--risk-period", default="MA (10 Aug 26)",
                    help="รอบที่ใช้ทำตัวอย่าง 'มีรายการเสี่ยงขาด' สำหรับทดสอบสี")
    ap.add_argument("--top", type=int, default=TOP_ROWS_DEFAULT)
    ap.add_argument("--outdir", default="adaptive-cards/data")
    args = ap.parse_args()

    rows = read_rows(args.xlsm)
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
