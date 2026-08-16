# Report Material Optimize&MA_NER
### Adaptive Card สำหรับ Power Automate → Microsoft Teams

ส่งรายงานวัสดุคงคลังเข้าห้อง Teams อัตโนมัติ
โดยดึงข้อมูลจากไฟล์ `Optimize_HiringMA_Full_Rebuild_VBA_Macro_V2.xlsm`
เฉพาะรหัสวัสดุ **8 รหัส** ที่กำหนดไว้ (ดูตารางข้างล่าง)

รายงานหลักคือชุด **OPTIMIZE — Allocation Plan** (รอบเดือน)
ชุด **MA — Weekly Allocation** (รอบสัปดาห์) ส่งเพิ่มได้จาก Flow เดียวกัน

---

## 🚀 เริ่มตรงนี้

👉 **[คู่มือสร้าง Flow ฉบับมือใหม่ ทำตามได้เอง](docs/01-POWER-AUTOMATE-FLOW-GUIDE.md)** — ใช้เวลาประมาณ 1 ชั่วโมง

👉 **[เช็กลิสต์ก่อนใช้งานจริง](docs/02-CHECKLIST.md)** — 6 ระยะ ติ๊กทีละข้อ

---

## รหัสวัสดุที่รายงาน (8 รหัส)

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

แก้รายการนี้ได้ที่ `TARGET_ITEMS` ใน `office-scripts/buildCardPayload.ts`
(และให้แก้ที่ `tools/build_sample_data.py` ตามกันไว้ เพื่อให้ชุดทดสอบยังเทียบได้)

ชื่อวัสดุมาจากรายการนี้โดยตรง และพิมพ์ไว้ครั้งเดียวเป็นตารางอธิบายท้ายการ์ด
แทนที่จะซ้ำต่อท้ายทุกแถว

---

## ⚠️ บั๊กหัวตารางในชีตต้นทาง — แก้ให้แล้ว

ชีต `MA&Optimize NER` มีคอลัมน์ว่างไม่มีหัวตารางแทรกอยู่หลัง `OMC-Onhand`
ทำให้บล็อก `จัดสรร` / `เบิกจาก Hub` / `จัดสรรเพิ่ม` / `รวม 3 รายการ`
มีข้อมูลอยู่ทางขวาของหัวตารางตัวเอง 1 ช่อง

ถ้าอ่านตามชื่อหัวตารางตรง ๆ ช่อง **เบิกจาก Hub** จะได้ค่าของ **จัดสรร** แทน
กระทบ **179 จาก 217 แถว** ที่เข้าเกณฑ์

สคริปต์ตรวจเองด้วยกฎ `จัดสรร + เบิกจาก Hub + จัดสรรเพิ่ม = รวม 3 รายการ`
แล้วเลือกตำแหน่งที่ถูกให้อัตโนมัติ — ถ้าชีตถูกแก้ให้ตรงในภายหลังก็ยังอ่านถูก

`OMC-Onhand`, `Prev_Before`, `Prev_Received` ตรวจแล้วอยู่ตรงตำแหน่ง ไม่โดนผลกระทบ

รายละเอียด: [docs/04-TROUBLESHOOTING.md ข้อ 7.1](docs/04-TROUBLESHOOTING.md)

---

## สิ่งที่การ์ดแสดง

| ข้อมูล | ที่มา |
|---|---|
| ชุดข้อมูล (Data_Set) | กำหนดสีของทั้งการ์ด — 📦 MA ฟ้า / 🗃️ OPTIMIZE เขียว |
| รอบจัดสรร (Distribution period) | หัวการ์ด |
| รอบก่อนหน้า (Prev_Period) | หัวการ์ด |
| Zone | หัวกลุ่มตาราง พร้อมยอดรวมทั้งโซน — **แสดงครบทุก Zone เสมอ** |
| Item Code | 8 รหัสตามตารางข้างบน |
| OMC-Onhand | คอลัมน์ตาราง |
| เบิกจาก Hub | คอลัมน์ตาราง (= ยอดจ่ายของรอบนั้น) — **จำนวนเต็ม ไม่มีทศนิยม** |
| Prev_Before / Prev_Received | คอลัมน์ตาราง |
| Status / Risk_Flag | สี emoji และการเรียงลำดับแถว |

### การรวมแถว — 1 Zone + 1 Item Code = 1 แถว

ชีตต้นทางแตกแถวตามจังหวัด รหัสเดียวกันจึงโผล่หลายครั้งในโซนเดียว
การ์ดไม่แสดง `Province` การเห็นรหัสซ้ำติดกันจึงอ่านไม่รู้เรื่องและกินที่เปล่า ๆ

การ์ดจึง **บวกยอดของทุกจังหวัดในโซนเข้าด้วยกัน** แล้วแสดงรหัสละแถวเดียว

- `OMC-Onhand`, `เบิกจาก Hub`, `Prev_Before`, `Prev_Received` = ผลบวกทุกจังหวัด
- ยอดรวมของ Zone บนหัวกลุ่ม = ผลบวกทั้งโซน
- สถานะของแถว = สถานะที่ **รุนแรงที่สุด** ในกลุ่ม (ใช้เกินยอดรับ > ยอดไม่ตรง > จัดสรรเกิน > ปกติ)
- แถวที่รวมมาจากหลายจังหวัดจะมีป้ายเล็ก ๆ ต่อท้ายรหัส เช่น `53OF157BB ·3จว.`

ผลคือทั้งรอบลงการ์ดใบเดียวได้หมด — MA (17 Aug 26) จาก **74 แถวดิบ เหลือ 24 แถว**

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

Teams ปฏิเสธการ์ดที่ใหญ่เกิน 28 KB **เงียบ ๆ** โดย Flow ยังขึ้นเขียว

สคริปต์จึง **วัดขนาดการ์ดเป็นไบต์จริงหลังประกอบเสร็จ** ถ้าเกินงบ 26,000 ไบต์
ก็ลดจำนวนแถวแล้วประกอบใหม่ วนได้สูงสุด 8 รอบ (ดู `CARD_SAFE_BYTES`)
ไม่ต้องมาไล่ปรับ `topRows` เองอีกต่อไป และปลอดภัยไม่ว่ารอบนั้นจะมีกี่ Zone

ขนาดจริงที่วัดจากข้อมูลในไฟล์ — **ทุกรอบลงได้ครบ ไม่มีการตัดทิ้ง**:

| รอบ | Zone | แถวดิบ | แถวบนการ์ด | ขนาด |
|---|---|---|---|---|
| Optimize (3 Aug 26) | 5 | 35 | 35 | 24.9 KB ✅ |
| MA (17 Aug 26) | 4 | 74 | 24 | 19.4 KB ✅ |
| MA (10 Aug 26) | 3 | 10 | 7 | 10.1 KB ✅ |

ถ้ารอบไหนใหญ่จนต้องตัดจริง ๆ ส่วนที่เหลือจะสรุปเป็นบรรทัด
*"…และอีก N รายการ"* พร้อมปุ่มไป Dashboard

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
| กฎ `จัดสรร + เบิกจาก Hub + จัดสรรเพิ่ม = รวม 3 รายการ` | ✅ 2,365/2,365 แถว |
