# รายงานวัสดุคงคลัง MA & Optimize — NER
### Adaptive Card สำหรับ Power Automate → Microsoft Teams

ส่งรายงานวัสดุคงคลังรายสัปดาห์เข้าห้อง Teams อัตโนมัติทุกวันจันทร์ 08:00 น.
โดยดึงข้อมูลจากไฟล์ `Optimize_HiringMA_Full_Rebuild_VBA_Macro_V2.xlsm`
เฉพาะรหัส **53OF** (สายเคเบิลใยแก้ว) และ **52CL** (ตู้พักสาย)

---

## 🚀 เริ่มตรงนี้

👉 **[คู่มือสร้าง Flow ฉบับมือใหม่ ทำตามได้เอง](docs/01-POWER-AUTOMATE-FLOW-GUIDE.md)** — ใช้เวลาประมาณ 1 ชั่วโมง

👉 **[เช็กลิสต์ก่อนใช้งานจริง](docs/02-CHECKLIST.md)** — 6 ระยะ ติ๊กทีละข้อ

---

## สิ่งที่การ์ดแสดง

| ข้อมูล | ที่มา |
|---|---|
| ชุดข้อมูล (Data_Set) | กำหนดสีของทั้งการ์ด — 🛠️ MA ฟ้า / ⚙️ Optimize เขียว |
| รอบจัดสรร (Distribution period) | หัวการ์ด |
| รอบก่อนหน้า (Prev_Period) | หัวการ์ด |
| Zone | หัวกลุ่มตาราง พร้อม emoji ประจำโซน |
| Item Code | เฉพาะ 53OFxxx และ 52CLxxx |
| OMC-Onhand | คอลัมน์ตาราง |
| เบิกจาก Hub | คอลัมน์ตาราง (= ยอดจ่ายของรอบนั้น) |
| Prev_Before / Prev_Received | คอลัมน์ตาราง |
| Status / Risk_Flag | สี emoji และการเรียงลำดับแถว |

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
การ์ดจึงแสดงได้สูงสุด **12 แถว** แล้วสรุปที่เหลือเป็นบรรทัด
*"…และอีก N รายการ"* พร้อมปุ่มไป Dashboard

ขนาดจริงที่วัดจากข้อมูลในไฟล์:

| แถว | MA | Optimize |
|---|---|---|
| 12 | 23.0 KB ✅ | 24.6 KB ✅ |
| 13 | 24.3 KB | 27.6 KB ⚠️ |
| 14 | 25.6 KB | 28.9 KB ❌ |

แถวถูกเรียงให้ **รายการเสี่ยงขาดขึ้นก่อนเสมอ** ตามด้วยของที่เหลือน้อยสุด
เรื่องที่ต้องรีบรู้จึงไม่มีทางตกไปอยู่นอก 12 แถวแรก

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

**ผลการทดสอบล่าสุด** (ข้อมูลจริงจากไฟล์ V2, 2,365 แถว, เข้าเกณฑ์ 453 แถว)

| ชุดทดสอบ | ผล |
|---|---|
| expand template ด้วย adaptivecards-templating | ✅ 6/6 เคส |
| ไม่มี `${...}` ค้างหลัง expand | ✅ |
| ขนาดการ์ดไม่เกิน 28 KB | ✅ |
| `$when` แสดง/ซ่อน empty state | ✅ |
| ปุ่ม Action.OpenUrl ครบ 2 ปุ่ม | ✅ |
| Office Script ให้ผลตรงกับสคริปต์อ้างอิง Python | ✅ 12/12 แถว ทุก field |
| การ์ดจาก Office Script ตรงกับไลบรารีทุก byte | ✅ 5/5 เคส |
