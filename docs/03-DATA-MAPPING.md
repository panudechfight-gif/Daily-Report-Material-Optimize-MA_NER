# ตารางเทียบคอลัมน์ และบั๊กที่พบในไฟล์ V2

---

## 1. ข้อมูลที่การ์ดแสดง

โจทย์กำหนดให้การ์ดแสดง: `Data_Set`, `Distribution period`, `Zone`,
`Item Code` (เฉพาะ 53OF และ 52CL), `OMC-Onhand`, `Prev_Period`,
`Prev_Before`, `Prev_Received` และ *เบิกจาก Hub = Distribution period ในรอบนั้น ๆ*

| ชื่อในการ์ด | คอลัมน์ต้นทาง | คอลัมน์ | แสดงตรงไหนบนการ์ด |
|---|---|---|---|
| ชุดข้อมูล | `Data_Set` | AC | หัวการ์ด — กำหนดสีทั้งใบ |
| รอบจัดสรร | `Distribution period` | A | หัวการ์ด แถบที่ 1 |
| รอบก่อนหน้า | `Prev_Period` | Q | หัวการ์ด แถบที่ 2 |
| Zone | `Zone` | C | หัวกลุ่ม (แบ่งตาราง) |
| จังหวัด | `Province` | D | ใต้ Item Code |
| รหัสวัสดุ | `Item Code` | E | คอลัมน์ที่ 1 |
| ชื่อวัสดุ | `Description` | G | ใต้ Item Code (จอกว้างเท่านั้น) |
| OMC-Onhand | `OMC-Onhand` | I | คอลัมน์ที่ 2 |
| เบิกจาก Hub | `เบิกจาก Hub` | K | คอลัมน์ที่ 3 |
| Prev_Before | `Prev_Before` | S | คอลัมน์ที่ 4 |
| Prev_Received | `Prev_Received` | T | คอลัมน์ที่ 5 |
| สถานะ | `Status` | AA | สี + emoji ของแถว |
| ความเสี่ยง | `Risk_Flag` | AB | ตัวหนา + ป้ายเตือนที่หัว Zone |

> **หมายเหตุเรื่อง "เบิกจาก Hub"**
> ตามโจทย์ ค่าในคอลัมน์นี้คือยอดจ่ายของรอบ (Distribution) ในรอบนั้น ๆ
> การ์ดจึงตั้งหัวคอลัมน์ว่า **"เบิกจาก Hub"** และผูกกับรอบที่ระบุในหัวการ์ด

---

## 2. เกณฑ์คัดรหัสวัสดุ

```
53OF + ตัวเลข 3 หลัก + BB หรือ AS     เช่น 53OF150BB, 53OF152AS
52CL + ตัวเลข 3 หลัก + BB หรือ AS     เช่น 52CL003BB, 52CL001AS
```

> 📌 โจทย์เขียนรูปแบบไว้ว่า `52CLx00xBB` แต่ข้อมูลจริงในไฟล์เป็น
> **52CL + 3 หลัก** (`52CL003BB`, `52CL009BB`, `52CL015BB`)
> จึงใช้เกณฑ์ 3 หลักเพื่อให้ครอบคลุมรหัสที่มีอยู่จริงทั้งหมด

**รหัสที่พบในไฟล์ (25 รหัส)**

| กลุ่ม | รหัส |
|---|---|
| ตู้พักสาย (52CL) | 52CL001AS, 52CL002AS, 52CL003AS, 52CL003BB, 52CL004AS, 52CL004BB, 52CL006BB, 52CL009BB, 52CL010BB, 52CL015BB |
| สายเคเบิล (53OF) | 53OF005AS, 53OF043BB, 53OF111BB, 53OF112BB, 53OF121BB, 53OF131BB, 53OF150BB, 53OF151BB, 53OF152AS, 53OF152BB, 53OF155BB, 53OF156BB, 53OF157BB, 53OF158BB, 53OF160BB |

**จำนวนแถวที่เข้าเกณฑ์ แยกตามรอบ**

| รอบ | จำนวน | เสี่ยงขาด |
|---|---|---|
| MA (17 Aug 26) | 126 | 0 |
| Optimize(3 Aug 26) | 105 | 14 |
| Optimize(6 July 26) | 100 | 0 |
| MA (13 July 26) | 33 | 0 |
| MA (20 July 26) | 32 | 2 |
| MA (27 July 26) | 21 | 3 |
| MA (3 Aug 26) | 21 | 3 |
| MA (10 Aug 26) | 15 | 9 |

---

## 3. สี / emoji ที่ใช้

### 3.1 ตาม Data_Set (กำหนดสีหัวการ์ดทั้งใบ)

| Data_Set | emoji | container style | ความหมาย |
|---|---|---|---|
| `MA` | 🛠️ | `accent` (ฟ้า) | งานซ่อมบำรุง |
| `Optimize` | ⚙️ | `good` (เขียว) | ปรับปรุงโครงข่าย |

### 3.2 ตาม Zone (หัวกลุ่มตาราง)

| Zone | emoji | จังหวัด |
|---|---|---|
| RC2-NMA | 🟦 | Buriram, Chaiyaphum, NakhonRatchasima, Surin |
| RC2-UBN | 🟩 | AmnatCharoen, Sisaket, UbonRatchathani, Yasothon |
| RC3-KKN | 🟨 | Kalasin, KhonKaen, MahaSarakham, RoiEt, SakonNakhon, Mukdahan, NakhonPhanom |
| RC3-UDN | 🟧 | BuengKan, Loei, NongBuaLamphu, NongKhai, UdonThani |
| RC3-SNK | 🟪 | (พบในข้อมูลแต่ไม่มีในชีต Zone) |

### 3.3 ตาม Status (สีของแถว)

| Status ในไฟล์ | emoji | สี AC | ป้ายสั้นบนการ์ด |
|---|---|---|---|
| ปกติ | ✅ | `good` | ปกติ |
| จัดสรรเกิน (ไม่มีการใช้) | 🟡 | `warning` | จัดสรรเกิน |
| ใช้เกินยอดรับ (ดึงสต็อกเดิม) | 🔴 | `attention` | ใช้เกินยอดรับ |
| ยอดไม่ตรง (รับนอกระบบ) | 🟠 | `warning` | ยอดไม่ตรง |
| - ไม่มีข้อมูลรอบก่อน | ⚪ | `default` | ไม่มีรอบก่อน |

> ใช้ **ป้ายสั้น** บนการ์ดเพราะข้อความเต็มยาวจนทำให้การ์ดชนเพดาน 28 KB
> ข้อความเต็มยังอยู่ในฟิลด์ `status` ของ payload และแสดงในการ์ดแบบมือถือ

### 3.4 ตาม Risk_Flag

| Risk_Flag | ผลบนการ์ด |
|---|---|
| `เสี่ยงขาด` | แถวเป็นตัวหนา + ขึ้นบนสุด + นับใน ⚠️ + ป้ายแดงที่หัว Zone |
| `พอใช้` | นับในกล่อง 🟡 พอใช้ |
| *(ว่าง)* | นับในกล่อง ✅ ปกติ |

**ลำดับการเรียงแถว:** เสี่ยงขาดก่อน → คงเหลือน้อยก่อน → Zone → รหัสวัสดุ
รายการที่สำคัญที่สุดจึงอยู่ใน 12 แถวแรกที่การ์ดแสดงเสมอ

---

## 4. ⚠️ บั๊กที่พบในไฟล์ V2 — หัวตาราง PA_Export เลื่อน 1 คอลัมน์

### 4.1 อาการ

ชีต `PA_Export` (ตาราง `tblPAExport`) ถูกสร้างไว้สำหรับ Power Automate โดยเฉพาะ
แต่ **หัวตารางตั้งแต่คอลัมน์ J ถึง AC เลื่อนไป 1 ช่อง** เทียบกับข้อมูลจริง

ถ้าเอา `tblPAExport` เข้า Power Automate โดยไม่แก้ **การ์ดจะแสดงข้อมูลผิดคอลัมน์
ทั้งใบโดยไม่มี error ให้เห็นเลย**

### 4.2 สาเหตุ

ฟังก์ชัน `BuildHeaders` ใน `modExport` ประกาศชื่อคอลัมน์ ASCII ไว้ **30 ชื่อ**
โดยมี `"OnhandAfter"` แทรกอยู่ตำแหน่งที่ 10:

```vba
def = Array("Period", "RO", "Zone", "Province", "ItemCode", "ArtNo", _
            "Description", "Unit", "OnhandBefore", "OnhandAfter", "Alloc", _
                                                    ^^^^^^^^^^^^^ ตัวปัญหา
```

แต่ตารางต้นทาง `tblMA_RO2RO3` มี **29 คอลัมน์** และ **ไม่มีคอลัมน์ OnhandAfter**
ชื่อทุกตัวหลังจากนั้นจึงเลื่อนไปเขียนทับตำแหน่งของคอลัมน์ถัดไป

### 4.3 ตารางเทียบ — หัวที่เขียนไว้ vs ข้อมูลที่อยู่จริง

| คอลัมน์ | หัวที่เขียนไว้ | ข้อมูลที่อยู่จริง | ผลถ้าไม่แก้ |
|---|---|---|---|
| A–I | ✅ ถูกต้อง | ✅ ถูกต้อง | ปกติ |
| J | `OnhandAfter` | จัดสรร (Alloc) | |
| K | `Alloc` | เบิกจาก Hub | **เบิกจาก Hub หาไม่เจอ** |
| L | `FromHub` | จัดสรรเพิ่ม | **ดึง FromHub ได้ค่าผิด** |
| M | `ExtraAlloc` | รวม 3 รายการ | |
| N | `Total3` | เหตุผล | |
| O | `Reason` | Control_Key | |
| P | `ControlKey` | Period_Seq | |
| Q | `PeriodSeq` | Prev_Period | |
| R | `PrevPeriod` | Prev_Exist | **Prev_Period ได้ค่า 0/1** |
| S | `PrevExist` | Prev_Before | |
| T | `PrevBefore` | Prev_Received | **สลับกันสองคอลัมน์** |
| U | `PrevReceived` | Expected_Before | |
| V | `ExpectedBefore` | Net_Used | |
| W | `NetUsed` | Days_Period | **ได้ค่า 7 = จำนวนวัน ไม่ใช่ยอดใช้** |
| X | `DaysPeriod` | Burn_Rate_Day | |
| Y | `BurnRateDay` | Days_Cover | |
| Z | `DaysCover` | Suggest_Next_Alloc | |
| AA | `SuggestNextAlloc` | Status | |
| AB | `Status` | Risk_Flag | **Status ได้ค่าความเสี่ยง** |
| AC | `RiskFlag` | Data_Set | **RiskFlag ได้ค่า "MA"/"Optimize"** |
| AD | `CutoffDate` | ✅ CutoffDate | ปกติ |
| AE | `ExportedAt` | ✅ ExportedAt | ปกติ |

**หลักฐานยืนยันจากข้อมูลจริง** (แถวที่ 1200 ของ PA_Export):

- คอลัมน์ `NetUsed` มีค่า `7` — ซึ่งคือ `Days_Period` (รอบรายสัปดาห์ = 7 วัน)
  ไม่ใช่ปริมาณการใช้
- คอลัมน์ `RiskFlag` มีค่า `"MA"` — ซึ่งเป็นค่าของ `Data_Set`
  (`Risk_Flag` จริง ๆ มีแค่ `เสี่ยงขาด` / `พอใช้` / ว่าง)

### 4.4 วิธีแก้ — เลือกอย่างใดอย่างหนึ่ง

**ทางเลือกที่ 1: ใช้ Office Script (ทางที่คู่มือนี้แนะนำ)**

`office-scripts/renderAdaptiveCard.ts` **อ่านชีต `MA&Optimize NER` โดยตรง**
และหาตำแหน่งคอลัมน์จากชื่อหัวตารางจริงในแถวที่ 7 จึงไม่โดนบั๊กนี้เลย
ไม่ต้องแก้อะไรในไฟล์ Excel

**ทางเลือกที่ 2: แก้มาโครแล้วใช้ PA_Export ต่อ**

ติดตั้ง `vba/modExport_FIXED.bas` แทน `modExport` เดิม:

1. เปิด VBA Editor (`Alt` + `F11`)
2. คลิกขวาที่ `modExport` → **Remove modExport** → ตอบ **No** ตอนถามว่าจะ export ไหม
3. **File → Import File…** → เลือก `vba/modExport_FIXED.bas`
4. รัน `BuildExportTable_Click` ใหม่

การแก้มี 2 จุด:

- ตัด `"OnhandAfter"` ออกจากลิสต์ เหลือ 29 ชื่อ ตรงกับตารางต้นทาง
- เพิ่ม `VerifyExportHeaders()` ที่เช็กจำนวนคอลัมน์ก่อน export ทุกครั้ง
  ถ้าวันหน้ามีคนเพิ่ม/ลบคอลัมน์ในชีต จะเตือนทันทีแทนที่จะเงียบแล้วเพี้ยน

---

## 5. โครงสร้าง payload ที่การ์ดใช้

```jsonc
{
  "meta": {
    "period": "MA (17 Aug 26)",      // Distribution period
    "prevPeriod": "MA (10 Aug 26)",  // Prev_Period
    "dataSet": "MA",                 // Data_Set
    "dataSetStyle": "accent",        // สีหัวการ์ด
    "totalItems": 126,               // จำนวนที่เข้าเกณฑ์ทั้งหมด
    "shownItems": 12,                // จำนวนที่แสดงจริงบนการ์ด
    "riskCount": 0, "watchCount": 0, "okCount": 126,
    "hasData": true,                 // ใช้กับ $when (empty state)
    "hasMore": true,                 // มีรายการเกินที่แสดง
    "moreText": "…และอีก 114 รายการ — กด “📊 Dashboard” เพื่อดูทั้งหมด",
    "generatedAt": "17 Aug 2569 08:00 น.",
    "dashboardUrl": "...", "sourceFileUrl": "..."
  },
  "zones": [
    {
      "zone": "RC2-NMA", "zoneEmoji": "🟦",
      "itemCount": 27,                       // ยอดเต็มของ Zone นี้
      "shownCount": 3,                       // ที่แสดงบนการ์ด
      "countLabel": "แสดง 3 จาก 27 รายการ",
      "riskBadge": "✅ ไม่มีรายการเสี่ยง", "riskColor": "good",
      "items": [
        {
          "itemCode": "52CL003BB",
          "itemName": "Closure OFC 12 core",
          "province": "Chaiyaphum", "unit": "PC",
          "onhand": "0", "fromHub": "0",
          "prevBefore": "24", "prevReceived": "0",
          "status": "ใช้เกินยอดรับ (ดึงสต็อกเดิม)",
          "statusShort": "ใช้เกินยอดรับ",
          "statusEmoji": "🔴", "statusColor": "attention",
          "riskFlag": "เสี่ยงขาด", "isRisk": true, "rowWeight": "bolder",
          "zone": "RC2-NMA", "zoneEmoji": "🟦"
        }
      ]
    }
  ],
  "items": [ /* รายการชุดเดียวกันแบบไม่แยก Zone — ใช้กับ 02_card-simple-flat.json */ ]
}
```
