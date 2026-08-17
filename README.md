# Report Stock Allocation Cable&Material NER
### Adaptive Card สำหรับ Power Automate → Microsoft Teams

ส่งรายงานวัสดุคงคลังเข้าห้อง Teams อัตโนมัติ
โดยดึงข้อมูลจากไฟล์ `Optimize_HiringMA_Full_Rebuild_VBA_Macro_V2.xlsm`
เฉพาะรหัสวัสดุ **10 รหัส** ที่กำหนดไว้ (ดูตารางข้างล่าง)

รายงานหลักคือชุด **OPTIMIZE — Allocation Plan** (รอบเดือน)
ชุด **MA — Weekly Allocation** (รอบสัปดาห์) ส่งเพิ่มได้จาก Flow เดียวกัน

---

## 🚀 เริ่มตรงนี้

👉 **[คู่มือสร้าง Flow ฉบับมือใหม่ ทำตามได้เอง](docs/01-POWER-AUTOMATE-FLOW-GUIDE.md)** — ใช้เวลาประมาณ 1 ชั่วโมง

👉 **[เช็กลิสต์ก่อนใช้งานจริง](docs/02-CHECKLIST.md)** — 6 ระยะ ติ๊กทีละข้อ

---

## รหัสวัสดุที่รายงาน (10 รหัส แบ่ง 4 กลุ่ม)

แถวข้อมูลบนการ์ดแสดง **`icon · ชื่อย่อ`** แทน Item Code เต็ม
(เช่น `🧰 · Closure 12C` แทน `52CL003BB`) อ่านง่ายกว่าและกินที่ต่อแถวน้อยลง
เพราะรวมกับตัวเลข 3 ค่าไว้บรรทัดเดียวด้วย รหัสเต็มไปอยู่ในตารางอธิบาย
ท้ายการ์ดแทน (ดูหัวข้อถัดไป)

| icon | Item Code | ชื่อวัสดุเต็ม | ชื่อย่อบนการ์ด |
|---|---|---|---|
| 🏷️ | `50MT004BB` | Name Plate (Aluminium) | `Name Plate Aluminium` |
| 🧰 | `52CL003BB` | CLOSURE 12 C | `Closure 12C` |
| 🧰 | `52CL004BB` | CLOSURE 24 C | `Closure 24C` |
| 🧰 | `52CL006BB` | CLOSURE 48 C | `Closure 48C` |
| 🧰 | `52CL009BB` | CLOSURE 12 C FOR OFC DROP WIRE (IN LINE) | `Closure 12C Inline` |
| 🧰 | `52CL010BB` | CLOSURE 60 C | `Closure 60C` |
| 🔌 | `53OF150BB` | OPTICAL FIBER DROP CABLE 1C, FLAT TYPE (G.657A) … | `Drop Cable 1C SC/UPC · 3m` |
| 🧶 | `53OF157BB` | ARSS OPTICAL FIBER CABLE 12c-FIBRE3 | `ARSS Fiber Cable 12C` |
| 🧶 | `53OF158BB` | ARSS OPTICAL FIBER CABLE 24c-FIBRE3 | `ARSS Fiber Cable 24C` |
| 🧶 | `53OF160BB` | ARSS OPTICAL FIBER CABLE 60c-FIBRE3 | `ARSS Fiber Cable 60C` |

**สี่กลุ่ม (แบ่งด้วย icon)** — 🏷️ NAME PLATE · 🧰 CLOSURE · 🔌 DROP CABLE · 🧶 ARSS FIBER CABLE
รหัสในกลุ่มเดียวกันใช้ icon เดียวกันเสมอ กวาดสายตาแยกกลุ่มได้ทันทีโดยไม่ต้องอ่านชื่อย่อ
(สาย Drop Cable แยก icon จาก ARSS Fiber Cable เพราะใช้งานคนละแบบ — Drop Cable
เป็นสาย Pigtail Patch สำหรับ Splice ส่วน ARSS เป็นสาย Trunk)

ตารางอธิบายท้ายการ์ด **1 บรรทัดต่อกลุ่ม** ไม่ใช่ต่อรหัส เช่น
`🧰 CLOSURE — หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic` แล้วตามด้วย
`52CL003BB=Closure 12C · 52CL004BB=Closure 24C · 52CL006BB=Closure 48C · …`
เป็นจุดเดียวที่ถอดชื่อย่อกลับเป็น Item Code เต็ม

แก้รายการนี้ได้ที่ `TARGET_ITEMS` ใน `office-scripts/buildCardPayload.ts`
(และให้แก้ที่ `tools/build_sample_data.py` ตามกันไว้ เพื่อให้ชุดทดสอบยังเทียบได้)
เพิ่มรหัสใหม่แล้วอย่าลืมตั้ง `group` ให้ตรงกลุ่มเดิม ไม่งั้นจะได้ icon "▫️" (ไม่รู้จัก)

---

## ⚠️ บั๊กหัวตารางในชีตต้นทาง — แก้ให้แล้ว

ชีต `MA&Optimize NER` มีคอลัมน์ว่างไม่มีหัวตารางแทรกอยู่หลัง `Onhand`
ทำให้บล็อก `จัดสรร(กระจาย)` / `จัดสรร(เบิก Hub)` / `จัดสรรเพิ่ม` / `รวม 3 รายการ`
มีข้อมูลอยู่ทางขวาของหัวตารางตัวเอง 1 ช่อง

ถ้าอ่านตามชื่อหัวตารางตรง ๆ ช่อง **จัดสรร(เบิก Hub)** จะได้ค่าของ **จัดสรร(กระจาย)** แทน
กระทบ **179 จาก 217 แถว** ที่เข้าเกณฑ์

สคริปต์ตรวจเองด้วยกฎ `จัดสรร(กระจาย) + จัดสรร(เบิก Hub) + จัดสรรเพิ่ม = รวม 3 รายการ`
แล้วเลือกตำแหน่งที่ถูกให้อัตโนมัติ — ถ้าชีตถูกแก้ให้ตรงในภายหลังก็ยังอ่านถูก

`Onhand`, `Prev_Before`, `Prev_Received` ตรวจแล้วอยู่ตรงตำแหน่ง ไม่โดนผลกระทบ

รายละเอียด: [docs/04-TROUBLESHOOTING.md ข้อ 7.1](docs/04-TROUBLESHOOTING.md)

---

## สิ่งที่การ์ดแสดง

| ข้อมูล | ที่มา |
|---|---|
| ชุดข้อมูล (Data_Set) | กำหนดสีของทั้งการ์ด — MA ฟ้า / Optimize เขียว (หัวการ์ดใช้ 🏬 ทั้งคู่) |
| รอบจัดสรร (Distribution period) | หัวการ์ด |
| รอบก่อนหน้า (Prev_Period) | หัวการ์ด |
| Zone | หัวกลุ่มตาราง พร้อมยอดรวมทั้งโซน — **แสดงครบทุก Zone เสมอ** |
| Item Code | 10 รหัสตามตารางข้างบน — แสดงเป็น icon กลุ่ม + ชื่อย่อ |
| Onhand | รวมในบรรทัดเดียวกับอีก 2 ค่า (คอลัมน์ขวา) · รวมต่อ Zone บนหัวกลุ่มด้วย |
| จัดสรร(กระจาย) | รวมในบรรทัดเดียวกัน — **ชุด Optimize ลงยอดไว้ที่นี่** |
| จัดสรร(เบิก Hub) | รวมในบรรทัดเดียวกัน — **ชุด MA ลงยอดไว้ที่นี่** — จำนวนเต็ม |
| Status / Risk_Flag | สี emoji และการเรียงลำดับแถว |

### การรวมแถว — 1 Zone + 1 Item Code = 1 แถว

ชีตต้นทางแตกแถวตามจังหวัด รหัสเดียวกันจึงโผล่หลายครั้งในโซนเดียว
การ์ดไม่แสดง `Province` การเห็นรหัสซ้ำติดกันจึงอ่านไม่รู้เรื่องและกินที่เปล่า ๆ

การ์ดจึง **บวกยอดของทุกจังหวัดในโซนเข้าด้วยกัน** แล้วแสดงรหัสละแถวเดียว

- `Onhand`, `จัดสรร(กระจาย)`, `จัดสรร(เบิก Hub)` = ผลบวกทุกจังหวัด
- ยอดรวมของ Zone บนหัวกลุ่ม = ผลบวกทั้งโซน
- สถานะของแถว = สถานะที่ **รุนแรงที่สุด** ในกลุ่ม (ใช้เกินยอดรับ > ยอดไม่ตรง > จัดสรรเกิน > ปกติ)
- แถวที่รวมมาจากหลายจังหวัดจะมีป้ายเล็ก ๆ ต่อท้ายชื่อย่อ เช่น `🧶 · ARSS Fiber Cable 12C ·3จว.`

ผลคือทั้งรอบลงการ์ดใบเดียวได้หมด — MA (17 Aug 26) จาก **91 แถวดิบ เหลือ 29 แถว**
(นับตามรหัสวัสดุชุด 10 รหัสล่าสุด)

> การ์ดไม่แสดงชื่อ `Province` เพื่อให้เหลือที่พอสำหรับข้อมูลตัวเลข

**ปุ่มบนการ์ด** — 📊 Dashboard และ 📁 Open the source file. (Excel / SharePoint)

---

## โครงสร้างโปรเจกต์

```
adaptive-cards/
  01_card-main.json              ★ การ์ดหลัก จัดกลุ่มตาม Zone + responsive
  02_card-simple-flat.json         การ์ดแบบง่าย ตารางเดียวไม่แยก Zone
  03_card-empty-state.json         การ์ด "ไม่มีข้อมูล" แบบแยกไฟล์
  04_card-mobile-stacked.json      การ์ดแนวตั้งสำหรับมือถือโดยเฉพาะ
  data/                            ข้อมูลตัวอย่างจากไฟล์จริง (ใช้กับ Designer)
  expanded/                        ผลลัพธ์หลัง expand (สร้างจากการทดสอบ)

office-scripts/
  renderAdaptiveCard.ts          ★ สคริปต์ที่เอาไปวางใน Excel (สร้างอัตโนมัติ)
  buildCardPayload.ts              ตรรกะเตรียมข้อมูล (ใช้ถ้าอยากได้แค่ payload)
  _renderAdaptiveCard.skeleton.ts  โครงที่ใช้ประกอบไฟล์ข้างบน

vba/
  modExport_FIXED.bas              แก้บั๊กหัวตาราง PA_Export เลื่อน 1 คอลัมน์

docs/
  01-POWER-AUTOMATE-FLOW-GUIDE.md  คู่มือทีละขั้น
  02-CHECKLIST.md                  เช็กลิสต์ 6 ระยะ + รายสัปดาห์
  03-DATA-MAPPING.md               ตารางเทียบคอลัมน์ + รายละเอียดบั๊ก
  04-TROUBLESHOOTING.md            รวมปัญหาและวิธีแก้

tools/
  build_sample_data.py             สร้างข้อมูลตัวอย่างจากไฟล์ .xlsm
  dump_sheet_values.py             ดึงค่าทั้งชีตไว้ใช้ทดสอบ
  build_office_script.js           ประกอบ renderAdaptiveCard.ts
  expand_and_validate.js           ทดสอบ template ด้วย adaptivecards-templating
  test_office_script.js            ทดสอบตรรกะ Office Script
  test_render_script.js            ทดสอบว่าการ์ดตรงกับไลบรารีทุก byte
```

---

## ⚠️ สองเรื่องที่ต้องรู้ก่อนเริ่ม

### 1. Power Automate ไม่ expand template ให้

ไฟล์ `01_card-main.json` มี `${...}`, `$data`, `$when` ซึ่งใช้ได้ที่
[adaptivecards.io/designer](https://adaptivecards.io/designer/) สำหรับออกแบบเท่านั้น

action **Post adaptive card** ของ Teams รับได้เฉพาะ JSON ที่แทนค่าแล้ว
คู่มือนี้จึงให้ **Office Script** เป็นคนแทนค่าให้ในแอ็กชันเดียว

ผลลัพธ์ถูกตรวจแล้วว่า **ตรงกับไลบรารี `adaptivecards-templating` ทุก byte**
ในทุกเคสทดสอบ — สิ่งที่เห็นใน Designer คือสิ่งที่ Teams ได้จริง

### 2. ไฟล์ V2 มีบั๊กหัวตารางในชีต PA_Export

ชีต `PA_Export` ที่ทำไว้ให้ Power Automate มี**หัวตารางเลื่อน 1 คอลัมน์**
ตั้งแต่คอลัมน์ J ถึง AC ถ้าดึงไปใช้ตรง ๆ ข้อมูลจะผิดคอลัมน์ทั้งใบ
โดยไม่มี error ให้เห็น เช่น `Status` จะได้ค่าของ `Risk_Flag`
และ `RiskFlag` จะได้ค่า `"MA"` / `"Optimize"`

- **ทางที่แนะนำ:** ใช้ `renderAdaptiveCard.ts` ซึ่งอ่านชีตต้นทางโดยตรง — ไม่ต้องแก้อะไร
- **ถ้าต้องใช้ PA_Export:** ติดตั้ง `vba/modExport_FIXED.bas`

รายละเอียด: [docs/03-DATA-MAPPING.md](docs/03-DATA-MAPPING.md#4-️-บั๊กที่พบในไฟล์-v2--หัวตาราง-pa_export-เลื่อน-1-คอลัมน์)

---

## เพดาน 28 KB ของ Teams

Teams ปฏิเสธข้อความที่ใหญ่เกิน 28 KB ด้วย error **413 RequestEntityTooLarge**

สคริปต์จึง **วัดขนาดการ์ดหลังประกอบเสร็จ** ถ้าเกินงบก็ลดจำนวนแถวแล้วประกอบใหม่
วนได้สูงสุด 8 รอบ (ดู `CARD_SAFE_BYTES`) ไม่ต้องมาไล่ปรับ `topRows` เอง

### ทำไมงบถึงเป็น 24,500 ไบต์ ทั้งที่เพดานคือ 28 KB

เพดาน 28 KB ของ Teams นับ **ทั้งซองข้อความ** ไม่ใช่แค่ตัวการ์ด และตอนเดินสาย
อักขระนอก ASCII อาจถูก escape เป็น `\uXXXX` — ตัวอักษรไทยจาก 3 ไบต์เป็น 6 ไบต์
emoji จาก 4 ไบต์เป็น 12 ไบต์ การ์ดใบนี้เป็นภาษาไทยเกือบทั้งใบ
การวัดแบบ UTF-8 ตรง ๆ จึงต่ำกว่าความจริง

เคยตั้งงบไว้ 26,000 ไบต์แล้ววัดแบบ UTF-8 — การ์ดรอบ Optimize ที่วัดได้ 24.9 KB
(worst case 26,859 ไบต์) **ถูก Teams ปฏิเสธจริงด้วย RequestEntityTooLarge**
ตอนนี้จึงวัดแบบแย่ที่สุด (`wireBytes`) และตั้งงบไว้ที่ 24,500 ไบต์

การ์ดถูกทำให้เล็กลง 2 ชั้น:

1. ตัดคอลัมน์ `Prev_Before` / `Prev_Received` ออก
2. รวม 3 ตัวเลข (Onhand / กระจาย / เบิก Hub) ของแต่ละแถวเป็น **บรรทัดเดียว**
   แทนที่จะแยก 3 คอลัมน์ ลด JSON ต่อแถวได้มาก — จำเป็นขึ้นเมื่อรหัสวัสดุ
   เพิ่มเป็น 10 รหัส (จาก 8) ทำให้ Optimize ที่มี 5 Zone พุ่งไปแตะ 45 แถว
   ซึ่งด้วยหน้าตาเดิม (4 คอลัมน์แยก) กินไป 26.9 KB เหลือขอบแค่ 1.1 KB จาก 28 KB

ทุกรอบจึงยังลงการ์ดใบเดียว **ครบทุกรหัส ไม่มีการตัดแถวทิ้ง** และมีขอบปลอดภัยเผื่อไว้มาก:

| รอบ | Zone | รหัสวัสดุ (10 รหัส) | UTF-8 | worst case | เหลือถึงเพดาน 28 KB |
|---|---|---|---|---|---|
| Optimize (3 Aug 26) | 5 | **45 ครบ** | 18.0 KB | 20.7 KB ✅ | 7.3 KB |
| MA (17 Aug 26) | 4 | **29 ครบ** | 13.9 KB | 16.1 KB ✅ | 11.9 KB |
| MA (10 Aug 26) | 3 | **9 ครบ** | 8.1 KB | 9.4 KB ✅ | 18.6 KB |

บรรทัด *"…และอีก N รายการ"* ถูกตัดออกจากการ์ดแล้ว เพราะทุกรหัสลงครบเสมอ
ถ้าวันหนึ่งข้อมูลโตจนเกินงบจริง ๆ จะเห็นเป็น `แสดง n/N` บนหัวกลุ่ม Zone แทน

**การเลือกแถว** ทำแบบวนรอบทีละ Zone (เอาอันดับ 1 ของทุก Zone ก่อน แล้วค่อยวนอันดับ 2)
ทุก Zone จึงมีที่บนการ์ดเสมอ และภายในแต่ละ Zone เรียง **รายการเสี่ยงขาดขึ้นก่อน**
ตามด้วยของที่เหลือน้อยสุด

---

## การทดสอบ

```bash
npm install adaptivecards-templating typescript

python3 tools/build_sample_data.py <path-to-xlsm>
python3 tools/dump_sheet_values.py <path-to-xlsm>

node tools/expand_and_validate.js   # template expand ผ่าน 6 เคส
node tools/build_office_script.js   # ประกอบ renderAdaptiveCard.ts
node tools/test_office_script.js    # ตรรกะตรงกับสคริปต์อ้างอิง Python
node tools/test_render_script.js    # การ์ดตรงกับไลบรารีทุก byte 5 เคส
```

**ผลการทดสอบล่าสุด** (ข้อมูลจริงจากไฟล์ V2, 2,365 แถว, เข้าเกณฑ์ 217 แถว)

| ชุดทดสอบ | ผล |
|---|---|
| expand template ด้วย adaptivecards-templating | ✅ 6/6 เคส |
| ไม่มี `${...}` ค้างหลัง expand | ✅ |
| ขนาดการ์ดไม่เกิน 28 KB | ✅ |
| `$when` แสดง/ซ่อน empty state | ✅ |
| ปุ่ม Action.OpenUrl ครบ 2 ปุ่ม | ✅ |
| Office Script ให้ผลตรงกับสคริปต์อ้างอิง Python | ✅ 35/35 แถว ทุก field |
| การ์ดจาก Office Script ตรงกับไลบรารีทุก byte | ✅ 5/5 เคส |
| กฎ `จัดสรร(กระจาย) + จัดสรร(เบิก Hub) + จัดสรรเพิ่ม = รวม 3 รายการ` | ✅ 2,365/2,365 แถว |
