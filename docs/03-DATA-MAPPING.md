# ตารางเทียบคอลัมน์ และบั๊กที่พบในไฟล์ V2

---

## 1. ข้อมูลที่การ์ดแสดง

โจทย์กำหนดให้การ์ดแสดง: `Data_Set`, `Distribution period`, `Zone`,
`Item Code` (เฉพาะ 8 รหัสที่กำหนด), `OMC-Onhand`, `Prev_Period`,
`Prev_Before`, `Prev_Received` และ *เบิกจาก Hub = Distribution period ในรอบนั้น ๆ*

| ชื่อในการ์ด | คอลัมน์ต้นทาง | คอลัมน์ | แสดงตรงไหนบนการ์ด |
|---|---|---|---|
| ชุดข้อมูล | `Data_Set` | AC | หัวการ์ด — กำหนดสีทั้งใบ |
| รอบจัดสรร | `Distribution period` | A | หัวการ์ด แถบที่ 1 |
| รอบก่อนหน้า | `Prev_Period` | Q | หัวการ์ด แถบที่ 2 |
| Zone | `Zone` | C | หัวกลุ่ม พร้อมยอดรวมทั้งโซน — แสดงครบทุก Zone |
| รหัสวัสดุ | `Item Code` | E | คอลัมน์ที่ 1 |
| ชื่อวัสดุ | `TARGET_ITEMS` | — | ตารางอธิบายท้ายการ์ด (พิมพ์ครั้งเดียว ไม่ซ้ำทุกแถว) |
| OMC-Onhand | `OMC-Onhand` | I | คอลัมน์ที่ 2 |
| เบิกจาก Hub | `เบิกจาก Hub` | **L** | คอลัมน์ที่ 3 — **ปัดเป็นจำนวนเต็ม** (ดูข้อ 4.0 เรื่องคอลัมน์เลื่อน) |
| Prev_Before | `Prev_Before` | S | คอลัมน์ที่ 4 |
| Prev_Received | `Prev_Received` | T | คอลัมน์ที่ 5 |
| สถานะ | `Status` | AA | สี + emoji ของแถว |
| ความเสี่ยง | `Risk_Flag` | AB | ตัวหนา + ป้ายเตือนที่หัว Zone |

> **หมายเหตุเรื่อง "เบิกจาก Hub"**
> ตามโจทย์ ค่าในคอลัมน์นี้คือยอดจ่ายของรอบ (Distribution) ในรอบนั้น ๆ
> การ์ดจึงตั้งหัวคอลัมน์ว่า **"เบิกจาก Hub"** และผูกกับรอบที่ระบุในหัวการ์ด
>
> ค่านี้แสดงเป็น **จำนวนเต็มเสมอ** (ฟังก์ชัน `fmtInt`) เพราะเป็นยอดจ่ายจริงที่นับเป็นชิ้น
> ต่างจาก `OMC-Onhand` / `Prev_Before` / `Prev_Received` ที่ยังคงทศนิยม 2 ตำแหน่ง
> เมื่อค่าต้นทางมีเศษ (ฟังก์ชัน `fmtNum`)

> **หมายเหตุเรื่อง `Province`**
> การ์ด **ไม่แสดงชื่อจังหวัด** แต่ใช้จังหวัดเป็นตัวรวมแถว —
> รหัสเดียวกันในโซนเดียวกันถูกบวกยอดเข้าด้วยกันเหลือแถวเดียว (ดูข้อ 3.5)
> payload ยังเก็บ `province` (จำนวนจังหวัดเป็นข้อความ), `provinceCount` (ตัวเลข)
> และ `provinceTag` (ป้ายเล็กต่อท้ายรหัส เช่น ` ·3จว.`) ไว้ให้ใช้ต่อ

---

## 2. รหัสวัสดุที่รายงาน (8 รหัส)

กำหนดเป็นรายการตายตัว ไม่ได้ใช้ pattern matching
ดูที่ `TARGET_ITEMS` ใน `office-scripts/buildCardPayload.ts`
ค่าของแต่ละรหัสคือชื่อที่ใช้แสดงในตารางอธิบายท้ายการ์ด

| Item Code | ชื่อวัสดุ |
|---|---|
| `53OF150BB` | OPTICAL FIBER DROP CABLE 1C, FLAT TYPE (G.657A) WITH 2 SC/UPC PRE-CONNECTOR, 3m. |
| `53OF157BB` | ARSS OPTICAL FIBER CABLE 12c-FIBRE3 |
| `53OF158BB` | ARSS OPTICAL FIBER CABLE 24c-FIBRE3 |
| `53OF160BB` | ARSS OPTICAL FIBER CABLE 60c-FIBRE3 |
| `50MT004BB` | Name Plate (Aluminium) |
| `52CL003BB` | CLOSURE 12 C |
| `52CL004BB` | CLOSURE 24 C |
| `52CL010BB` | CLOSURE 60 C |

> ⚠️ ถ้าแก้รายการนี้ ต้องแก้ที่ `tools/build_sample_data.py` ให้ตรงกันด้วย
> มิฉะนั้น `npm test` จะเทียบผลไม่ตรงและฟ้อง error

**จำนวนแถวแยกตามรอบ** — "แถวดิบ" คือแถวในชีต "หลังรวม" คือแถวบนการ์ด
(การ์ดรวม 1 Zone + 1 Item Code = 1 แถว โดยบวกยอดของทุกจังหวัดเข้าด้วยกัน)

| รอบ | ชุด | แถวดิบ | หลังรวม | Zone | เสี่ยงขาด |
|---|---|---|---|---|---|
| MA (10 Aug 26) | MA | 10 | 7 | 3 | 6 |
| MA (13 July 26) | MA | 22 | 14 | 4 | 0 |
| MA (17 Aug 26) | MA | 74 | 24 | 4 | 0 |
| MA (20 July 26) | MA | 18 | 10 | 3 | 0 |
| MA (27 July 26) | MA | 10 | 7 | 3 | 0 |
| MA (3 Aug 26) | MA | 13 | 9 | 4 | 1 |
| Optimize(3 Aug 26) | Optimize | 35 | 35 | 5 | 11 |
| Optimize(6 July 26) | Optimize | 35 | 35 | 5 | 0 |

> ชุด Optimize ไม่มีค่าในคอลัมน์ `Province` แถวดิบจึงเท่ากับแถวหลังรวมอยู่แล้ว
> ส่วนชุด MA แตกตามจังหวัด การรวมจึงลดจำนวนแถวลงมาก

---

## 3. สี / emoji ที่ใช้

### 3.1 ตาม Data_Set (กำหนดสีหัวการ์ดทั้งใบ)

| Data_Set | emoji | container style | ป้ายบนการ์ด |
|---|---|---|---|
| `MA` | 📦 | `accent` (ฟ้า) | MA — Weekly Allocation |
| `Optimize` | 🗃️ | `good` (เขียว) | OPTIMIZE — Allocation Plan |

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
| ปกติ | ✅ | `good` | *(ไม่แสดง)* |
| จัดสรรเกิน (ไม่มีการใช้) | 🟡 | `warning` | *(ไม่แสดง)* |
| ใช้เกินยอดรับ (ดึงสต็อกเดิม) | 🔴 | `attention` | ใช้เกินยอดรับ |
| ยอดไม่ตรง (รับนอกระบบ) | 🟠 | `warning` | ยอดไม่ตรง |
| - ไม่มีข้อมูลรอบก่อน | ⚪ | `default` | *(ไม่แสดง)* |

> ป้ายที่เขียนว่า *(ไม่แสดง)* คือกรณีที่ **emoji สื่อความหมายพอแล้ว**
> จึงตัดคำอธิบายทิ้งเพื่อเหลือที่ให้ข้อมูลตัวเลข
> ข้อความเต็มยังอยู่ในฟิลด์ `status` ของ payload สำหรับการ์ดแบบอื่น

### 3.4 ตาม Risk_Flag

| Risk_Flag | ผลบนการ์ด |
|---|---|
| `เสี่ยงขาด` | แถวเป็นตัวหนา + ขึ้นบนสุดของ Zone + นับใน ⚠️ + ป้ายแดงที่หัว Zone |
| `พอใช้` | นับในกล่อง 🟡 พอใช้ |
| *(ว่าง)* | นับในกล่อง ✅ ปกติ |

**ลำดับการเรียงภายในแต่ละ Zone:** เสี่ยงขาดก่อน → คงเหลือน้อยก่อน → รหัสวัสดุ

**การเลือกแถวลงการ์ด:** วนรอบทีละ Zone — หยิบอันดับ 1 ของทุก Zone ก่อน
แล้วค่อยวนกลับมาหยิบอันดับ 2 จนครบโควตา `topRows`

ผลคือ **ทุก Zone มีที่บนการ์ดเสมอ** และรายการเสี่ยงขาดของแต่ละ Zone
ขึ้นก่อนรายการอื่นใน Zone เดียวกันเสมอ

---

## 3.5 การรวมแถว — 1 Zone + 1 Item Code = 1 แถว

ชีตต้นทางแตกแถวตามจังหวัด รหัสเดียวกันจึงโผล่หลายครั้งในโซนเดียว
เช่น `RC2-NMA` + `53OF157BB` มีอยู่ 3 แถว (3 จังหวัด) ในรอบ MA (17 Aug 26)

การ์ดไม่แสดงชื่อจังหวัด การเห็นรหัสเดิมซ้ำติดกันจึงอ่านไม่รู้เรื่อง
จึงรวมเป็นแถวเดียวก่อนขึ้นการ์ด

| ค่า | วิธีรวม |
|---|---|
| `OMC-Onhand` | บวกทุกจังหวัดในโซน |
| `เบิกจาก Hub` | บวกทุกจังหวัดในโซน |
| `Prev_Before` | บวกทุกจังหวัดในโซน |
| `Prev_Received` | บวกทุกจังหวัดในโซน |
| `Status` | เลือกอันที่ **รุนแรงที่สุด** ในกลุ่ม |
| `Risk_Flag` | ถ้ามีจังหวัดใดเสี่ยงขาด ทั้งแถวนับเป็นเสี่ยงขาด |
| จำนวนจังหวัด | ต่อท้ายรหัสเป็นป้ายเล็ก เช่น `53OF157BB ·3จว.` |

**ลำดับความรุนแรงของ Status** (ใช้เลือกตัวแทน)

```
ใช้เกินยอดรับ (4)  >  ยอดไม่ตรง (3)  >  จัดสรรเกิน (2)  >  ปกติ (1)  >  ไม่มีรอบก่อน (0)
```

**ยอดรวมของ Zone** บนหัวกลุ่มคือผลบวกของทุกจังหวัดในโซนนั้น ไม่ใช่แค่แถวที่แสดง

---

## 4. ⚠️ บั๊กที่พบในไฟล์ V2 — หัวตารางเลื่อน 1 คอลัมน์

พบ **2 จุด** ที่มีอาการเดียวกันแต่คนละชีต

| # | ชีต | กระทบอะไร | สถานะ |
|---|---|---|---|
| 4.0 | `MA&Optimize NER` (ชีตต้นทาง) | `เบิกจาก Hub` ได้ค่าของ `จัดสรร` | ✅ สคริปต์ตรวจและแก้เองแล้ว |
| 4.1–4.4 | `PA_Export` | ทั้งบล็อก J–AC เลื่อน | ⚠️ ต้องติดตั้ง VBA ถ้าจะใช้ชีตนี้ |

### 4.0 ชีตต้นทาง `MA&Optimize NER` — คอลัมน์ว่างไม่มีหัว

มีคอลัมน์ว่างที่ไม่มีหัวตารางแทรกอยู่หลัง `OMC-Onhand`
ข้อมูลของบล็อกถัดไปจึงอยู่ทางขวาของหัวตารางตัวเอง 1 ช่อง

| หัวตารางเขียนว่า | ข้อมูลจริงอยู่ที่ |
|---|---|
| `OMC-Onhand` | ตรงตำแหน่ง ✅ |
| *(ไม่มีหัว)* | คอลัมน์ว่าง — ว่างทั้ง 2,365 แถว |
| `จัดสรร` | คอลัมน์ถัดไป ➡️ |
| `เบิกจาก Hub` | คอลัมน์ถัดไป ➡️ |
| `จัดสรรเพิ่ม` | คอลัมน์ถัดไป ➡️ |
| `รวม 3 รายการ` | คอลัมน์ถัดไป ➡️ |
| `เหตุผล` | ไม่มีข้อมูล (หัวตารางที่เหลือค้าง) |
| `Control_Key` เป็นต้นไป | ตรงตำแหน่ง ✅ |

**ผลกระทบ:** สคริปต์รุ่นก่อนแสดงค่าของ `จัดสรร` ในช่อง `เบิกจาก Hub`
ผิด **179 จาก 217 แถว** ที่เข้าเกณฑ์

**วิธีตรวจที่สคริปต์ใช้** — กฎที่ชีตคำนวณไว้เอง

```
จัดสรร + เบิกจาก Hub + จัดสรรเพิ่ม = รวม 3 รายการ
```

| อ่านแบบ | กฎเป็นจริงกี่แถว |
|---|---|
| ตามหัวตารางตรง ๆ | 847 / 2,365 |
| เลื่อน 1 ช่อง | **2,365 / 2,365** ✅ |

สคริปต์ลองทั้งสองแบบแล้วเลือกแบบที่กฎเป็นจริงทุกแถว
ถ้าชีตถูกแก้ให้ตรงในภายหลัง ตัวตรวจจะเลือกแบบ "ไม่เลื่อน" ให้เองโดยไม่ต้องแก้โค้ด

โค้ด: `fix_header_shift()` ใน `tools/build_sample_data.py`
และบล็อก `shiftNames` ใน `office-scripts/buildCardPayload.ts`

> ✅ `OMC-Onhand`, `Prev_Before`, `Prev_Received` ตรวจแล้วอยู่ตรงตำแหน่ง ไม่โดนผลกระทบ
> ยืนยันด้วยสูตรในชีตเอง: `Expected_Before = Prev_Before + Prev_Received`
> และ `Net_Used = Expected_Before − OMC-Onhand`

---

## 4.1 บั๊กชีต PA_Export — หัวตารางเลื่อน 1 คอลัมน์

### อาการ

ชีต `PA_Export` (ตาราง `tblPAExport`) ถูกสร้างไว้สำหรับ Power Automate โดยเฉพาะ
แต่ **หัวตารางตั้งแต่คอลัมน์ J ถึง AC เลื่อนไป 1 ช่อง** เทียบกับข้อมูลจริง

ถ้าเอา `tblPAExport` เข้า Power Automate โดยไม่แก้ **การ์ดจะแสดงข้อมูลผิดคอลัมน์
ทั้งใบโดยไม่มี error ให้เห็นเลย**

### สาเหตุ

ฟังก์ชัน `BuildHeaders` ใน `modExport` ประกาศชื่อคอลัมน์ ASCII ไว้ **30 ชื่อ**
โดยมี `"OnhandAfter"` แทรกอยู่ตำแหน่งที่ 10:

```vba
def = Array("Period", "RO", "Zone", "Province", "ItemCode", "ArtNo", _
            "Description", "Unit", "OnhandBefore", "OnhandAfter", "Alloc", _
                                                    ^^^^^^^^^^^^^ ตัวปัญหา
```

แต่ตารางต้นทาง `tblMA_RO2RO3` มี **29 คอลัมน์** และ **ไม่มีคอลัมน์ OnhandAfter**
ชื่อทุกตัวหลังจากนั้นจึงเลื่อนไปเขียนทับตำแหน่งของคอลัมน์ถัดไป

### ตารางเทียบ — หัวที่เขียนไว้ vs ข้อมูลที่อยู่จริง

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

### วิธีแก้ — เลือกอย่างใดอย่างหนึ่ง

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
