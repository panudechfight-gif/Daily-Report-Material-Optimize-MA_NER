# ✅ ระบบเช็กลิสต์ — สร้าง Flow ส่ง Adaptive Card เข้า Teams

พิมพ์หน้านี้ออกมาแล้วติ๊กทีละข้อ หรือคัดลอกไปวางใน Planner / Teams Task ก็ได้
ทุกข้อเขียนไว้ให้ **ตรวจได้จริง** — มีวิธีตรวจและอาการที่แปลว่าไม่ผ่านกำกับไว้

**ความคืบหน้า:** ระยะ A ☐ · B ☐ · C ☐ · D ☐ · E ☐ · F ☐

---

## ระยะ A — เตรียมของ (ก่อนเริ่ม)

| ☐ | รายการ | วิธีตรวจว่าผ่าน |
|---|---|---|
| ☐ | A1. มีสิทธิ์ใช้ Power Automate | เปิด make.powerautomate.com แล้วกด Create ได้ |
| ☐ | A2. มีสิทธิ์เขียนไซต์ SharePoint | อัปโหลดไฟล์ทดสอบเข้าไลบรารีได้ |
| ☐ | A3. มีสิทธิ์โพสต์ในห้อง Teams ปลายทาง | พิมพ์ข้อความในห้องนั้นได้ |
| ☐ | A4. องค์กรเปิดใช้ Office Scripts | เปิด Excel for the web → เห็นแท็บ **Automate** |
| ☐ | A5. มีลิงก์ Dashboard ที่จะใส่ในปุ่ม 📊 | เปิดลิงก์แล้วเข้าได้จริง |
| ☐ | A6. มีไฟล์ `.xlsm` ต้นฉบับ | เปิดได้ ไม่ติด Protected View |

> ❗ **ถ้า A4 ไม่ผ่าน** ให้ข้ามไปใช้เส้นทางสำรอง
> ([ภาคผนวก ข ในคู่มือ](01-POWER-AUTOMATE-FLOW-GUIDE.md#ภาคผนวก-ข-เส้นทางสำรอง-ไม่ใช้-office-script))
> และต้องติดตั้ง `vba/modExport_FIXED.bas` ก่อนเสมอ

---

## ระยะ B — เตรียมไฟล์ Excel

| ☐ | รายการ | วิธีตรวจว่าผ่าน |
|---|---|---|
| ☐ | B1. รีเฟรชข้อมูลใน `.xlsm` แล้ว | ชีต `MA&Optimize NER` แสดงรอบล่าสุดที่ต้องการ |
| ☐ | B2. เซลล์ C2 / C4 เลือก Sheet ต้นทางถูกรอบ | C2 = รอบ MA, C4 = รอบ Optimize ที่ต้องการ |
| ☐ | B3. Save As เป็น `.xlsx` บน SharePoint แล้ว | เห็นไฟล์ `.xlsx` ในไลบรารี |
| ☐ | B4. ชื่อชีตในไฟล์ `.xlsx` ยังเป็น `MA&Optimize NER` | เปิดใน Excel for the web แล้วดูแท็บล่าง |
| ☐ | B5. เซลล์ **A7** = `Distribution period` | คลิก A7 แล้วดูแถบสูตร |
| ☐ | B6. แถวที่ 8 เป็นข้อมูลแถวแรก | A8 ควรเป็นชื่อรอบ เช่น `MA (13 July 26)` |
| ☐ | B7. คัดลอกลิงก์ไฟล์ `.xlsx` เก็บไว้แล้ว | วางลิงก์ในเบราว์เซอร์แล้วไฟล์เปิดขึ้น |

> ❗ **B5 ไม่ผ่าน** = สคริปต์จะ error `หาหัวตารางแถวที่ 7 ไม่เจอ`

---

## ระยะ C — ติดตั้ง Office Script

| ☐ | รายการ | วิธีตรวจว่าผ่าน |
|---|---|---|
| ☐ | C1. เปิดไฟล์ `.xlsx` ใน Excel for the web | URL ขึ้นต้นด้วย `.sharepoint.com` |
| ☐ | C2. Automate → New Script | Code Editor เปิดทางขวา |
| ☐ | C3. ลบโค้ดตัวอย่างออกหมด | Editor ว่างเปล่า |
| ☐ | C4. วางเนื้อ `office-scripts/renderAdaptiveCard.ts` ครบทั้งไฟล์ | บรรทัดสุดท้ายต้องเป็น `}` ปิด `main` |
| ☐ | C5. ตั้งชื่อสคริปต์ `renderAdaptiveCard` | ชื่อที่หัว Editor ตรงกัน |
| ☐ | C6. กด Save script | ไม่มีข้อความ error ค้าง |
| ☐ | C7. กด **Run** แล้วได้ JSON | Output ขึ้นต้น `{"type":"AdaptiveCard"` |
| ☐ | C8. JSON ที่ได้ไม่มี `${` เหลืออยู่ | ค้นหา `${` ใน Output ต้องไม่เจอ |

> ❗ **C7 error "ไม่พบชีต"** → กลับไป B4
> ❗ **C7 error "ไม่พบคอลัมน์: ..."** → กลับไป B5 / B6

---

## ระยะ D — สร้าง Flow

| ☐ | รายการ | วิธีตรวจว่าผ่าน |
|---|---|---|
| ☐ | D1. สร้าง Scheduled cloud flow แล้ว | เห็น action `Recurrence` |
| ☐ | D2. Recurrence: Frequency = `Week`, Interval = `1` | |
| ☐ | D3. Recurrence: **Time zone = Bangkok** | ไม่ใช่ UTC — ถ้าพลาด Flow จะยิงบ่าย 3 |
| ☐ | D4. Recurrence: On these days = `Monday` | |
| ☐ | D5. Recurrence: At these hours = `8` | |
| ☐ | D6. Recurrence: **At these minutes = `0`** | ถ้าเว้นว่าง Flow จะยิง 60 ครั้ง |
| ☐ | D7. เพิ่ม Compose ชื่อ `GeneratedAt` พร้อมสูตรเวลาไทย | ทดสอบแล้วได้รูปแบบ `17 Aug 2569 08:00 น.` |
| ☐ | D8. เพิ่ม **Excel Online (Business) → Run script** | |
| ☐ | D9. Run script: เลือกไฟล์ `.xlsx` ถูกไฟล์ | ชื่อไฟล์ตรงกับที่ทำใน B3 |
| ☐ | D10. Run script: Script = `renderAdaptiveCard` | มีช่องพารามิเตอร์ 5 ช่องโผล่มา |
| ☐ | D11. `topRows` = `12` | **ห้ามเกิน 12** |
| ☐ | D12. `dashboardUrl` ใส่ลิงก์จริงแล้ว | ไม่ใช่ `CHANGE-ME` |
| ☐ | D13. `sourceFileUrl` ใส่ลิงก์จาก B7 แล้ว | ไม่ใช่ `CHANGE-ME` |
| ☐ | D14. `generatedAt` ผูกกับ Outputs ของ `GeneratedAt` | |
| ☐ | D15. เพิ่ม **Teams → Post adaptive card in a chat or channel** | |
| ☐ | D16. Post as = `Flow bot`, Post in = `Channel` | |
| ☐ | D17. เลือก Team และ Channel ถูกห้อง | |
| ☐ | D18. ช่อง Adaptive Card = `string(body('Run_script')?['result'])` | สูตรไม่ขึ้นแดง |
| ☐ | D19. กด Save แล้วไม่มี error | |

> ❗ **D18 ขึ้นแดง** = ชื่อ action ไม่ใช่ `Run_script`
> ดูชื่อจริงที่หัว action แล้วเปลี่ยนช่องว่างเป็น `_` ในสูตร

---

## ระยะ E — ทดสอบ

| ☐ | รายการ | วิธีตรวจว่าผ่าน |
|---|---|---|
| ☐ | E1. กด Test → Manually → Run flow | ทุก action เขียวหมด |
| ☐ | E2. การ์ดโผล่ในห้อง Teams | เห็นการ์ดจริง ไม่ใช่ข้อความเปล่า |
| ☐ | E3. หัวการ์ดมีสีตาม Data_Set | 🛠️ MA = ฟ้า / ⚙️ Optimize = เขียว |
| ☐ | E4. แถบบนแสดง รอบจัดสรร / รอบก่อนหน้า / ชุดข้อมูล ครบ | ค่าไม่เป็น `-` ทั้งหมด |
| ☐ | E5. กล่องตัวเลข 4 กล่องแสดงถูก | ⚠️ + 🟡 + ✅ = 📦 รวม |
| ☐ | E6. หัวกลุ่ม Zone มี emoji และจำนวนรายการ | เช่น `🟦 Zone RC2-NMA` |
| ☐ | E7. ตารางครบ 5 คอลัมน์ | Item Code / OMC-Onhand / เบิกจาก Hub / Prev_Before / Prev_Received |
| ☐ | E8. รายการเสี่ยงขาดอยู่บนสุดและเป็นตัวหนา | ทดสอบด้วย `period` = `MA (10 Aug 26)` (มี 9 รายการเสี่ยง) |
| ☐ | E9. บรรทัด "…และอีก N รายการ" ขึ้นถูก | N = ยอดรวม − 12 |
| ☐ | E10. ปุ่ม **📊 Dashboard** กดแล้วเปิดถูกที่ | |
| ☐ | E11. ปุ่ม **📁 Open the source file** กดแล้วเปิดไฟล์ Excel | |
| ☐ | E12. **เปิดดูบนมือถือ** — เลย์เอาต์สลับเป็นแนวตั้ง | ตัวเลขย้ายมาอยู่ใต้ชื่อรายการ |
| ☐ | E13. ทดสอบ empty state | ตั้ง `period` = `ไม่มีรอบนี้` ชั่วคราว → ต้องเห็นการ์ด 🗂️ พร้อมคำแนะนำ 3 ข้อ |
| ☐ | E14. คืนค่า `period` กลับเป็นว่างแล้ว | ไม่ลืมข้อนี้! |

---

## ระยะ F — ส่งมอบ

| ☐ | รายการ | วิธีตรวจว่าผ่าน |
|---|---|---|
| ☐ | F1. เปิด Flow ให้ทำงานจริง (สถานะ On) | หน้า My flows แสดง `On` |
| ☐ | F2. ตั้ง Owner เพิ่มอย่างน้อย 1 คน | Flow → Share → เพิ่ม co-owner |
| ☐ | F3. แจ้งทีมว่าต้อง Save As `.xlsx` ทุกครั้งหลัง Refresh | มีคนรับผิดชอบชัดเจน |
| ☐ | F4. บันทึกว่าใครดูแล Flow นี้ | เขียนไว้ในห้อง Teams หรือ Wiki |
| ☐ | F5. รอดูรอบจริงรอบแรก (จันทร์ถัดไป 08:00) | การ์ดมาตามเวลา |
| ☐ | F6. ตรวจ Run history หลังรอบแรก | 28-day run history ไม่มีสีแดง |

---

## 🔄 เช็กลิสต์ประจำสัปดาห์ (ทำก่อนเช้าวันจันทร์)

| ☐ | รายการ |
|---|---|
| ☐ | เปิด `.xlsm` → กด **Refresh Data New Sheet** |
| ☐ | ตรวจว่า C2 / C4 เลือกรอบล่าสุดแล้ว |
| ☐ | **Save As ทับไฟล์ `.xlsx` บน SharePoint** ← ข้อที่ลืมกันบ่อยที่สุด |
| ☐ | ปิดไฟล์ให้เรียบร้อย (ไฟล์ที่เปิดค้างอาจล็อกไม่ให้ Flow อ่าน) |

> 💡 ถ้าลืมข้อ "Save As ทับ `.xlsx`" การ์ดจะมาตามปกติ **แต่เป็นข้อมูลของสัปดาห์ก่อน**
> สังเกตได้จากค่า "รอบจัดสรร" บนหัวการ์ดที่ไม่เปลี่ยน

---

## 🧪 เช็กลิสต์สำหรับนักพัฒนา (เมื่อแก้ template การ์ด)

| ☐ | คำสั่ง | ต้องได้ผลว่า |
|---|---|---|
| ☐ | `python3 tools/build_sample_data.py <xlsm>` | เขียนไฟล์ตัวอย่าง 4 ไฟล์ |
| ☐ | `python3 tools/dump_sheet_values.py <xlsm>` | เขียน `tools/.sheet-values.json` |
| ☐ | `node tools/expand_and_validate.js` | ✅ ผ่านทั้งหมด |
| ☐ | `node tools/build_office_script.js` | เขียน `renderAdaptiveCard.ts` ใหม่ |
| ☐ | `node tools/test_office_script.js` | ✅ ตรงกับสคริปต์อ้างอิง |
| ☐ | `node tools/test_render_script.js` | ✅ ตรงกับ adaptivecards-templating ทุก byte |
| ☐ | ทุกเคสขนาดการ์ด < 28 KB | ไม่มีคำเตือนเรื่องขนาด |
| ☐ | คัดลอก `renderAdaptiveCard.ts` ไปวางทับสคริปต์ใน Excel | รัน Test flow แล้วการ์ดยังถูกต้อง |
