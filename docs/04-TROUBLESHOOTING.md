# 🔧 รวมปัญหาและวิธีแก้

เรียงจากที่เจอบ่อยที่สุด

---

## 1. การ์ดไม่ขึ้นใน Teams เลย ทั้งที่ Flow เขียวหมด

**สาเหตุที่พบบ่อยที่สุด: การ์ดใหญ่เกิน 28 KB**

Teams ปฏิเสธ Adaptive Card ที่เกิน 28 KB **เงียบ ๆ** โดย Power Automate
ยังรายงานว่าสำเร็จ

**วิธีตรวจ**

1. เปิด Run history → กดที่ action `Run script`
2. คัดลอกค่า **OUTPUTS → result** ทั้งก้อน
3. วางลงในโปรแกรมนับตัวอักษร หรือเซฟเป็นไฟล์แล้วดูขนาด
4. ถ้าเกิน 28,672 ไบต์ = เจอสาเหตุแล้ว

**วิธีแก้**

ลด `topRows` ใน action `Run script` เหลือ `10`

**ขนาดจริงที่วัดจากข้อมูลในไฟล์นี้**

| แถว | ชุด MA | ชุด Optimize |
|---|---|---|
| 10 | 20.5 KB | 20.3 KB |
| 11 | 21.8 KB | 21.6 KB |
| 12 | 23.0 KB ✅ | 24.6 KB ✅ |
| 13 | 24.3 KB | 27.6 KB ⚠️ |
| 14 | 25.6 KB | **28.9 KB ❌** |

**สาเหตุอื่นที่เป็นไปได้**

- ช่อง Adaptive Card ใส่ค่าไม่ถูก → ดูข้อ 2
- Flow bot ถูกบล็อกในห้องนั้น → ลองโพสต์ในห้องอื่นเทียบดู

---

## 2. การ์ดขึ้นเป็นข้อความ `${meta.period}` ตรง ๆ

**สาเหตุ:** เอา template ดิบจาก `adaptive-cards/01_card-main.json`
ไปวางในช่อง Adaptive Card โดยตรง

Power Automate **ไม่รู้จัก** `${...}`, `$data`, `$when` — มันแค่ส่งข้อความไปตามที่พิมพ์

**วิธีแก้:** ช่อง Adaptive Card ต้องใส่สูตรนี้เท่านั้น

```
string(body('Run_script')?['result'])
```

---

## 3. สูตร `body('Run_script')` ขึ้นสีแดง

ชื่อ action ในสูตรต้องตรงกับชื่อจริง โดยเปลี่ยน **ช่องว่างเป็นขีดล่าง**

| ชื่อ action ที่เห็นบนจอ | เขียนในสูตร |
|---|---|
| `Run script` | `body('Run_script')` |
| `Run script 2` | `body('Run_script_2')` |
| `เรียกสคริปต์` | เปลี่ยนชื่อ action เป็นภาษาอังกฤษก่อน |

**ทางลัด:** กดแท็บ **Dynamic content** แล้วเลือก **result** จากรายการ
Power Automate จะเติมสูตรให้เอง จากนั้นค่อยครอบด้วย `string(...)`

---

## 4. Run script error: `ไม่พบชีต MA&Optimize NER`

**ตรวจตามลำดับ**

1. เปิดไฟล์ `.xlsx` ใน Excel for the web → ดูชื่อแท็บล่างสุด
2. ชื่อต้องเป็น `MA&Optimize NER` **เป๊ะ ๆ** — รวมเครื่องหมาย `&` และช่องว่าง
3. ระวังชื่อที่ดูเหมือนกันแต่ไม่เหมือน:
   - `MA & Optimize NER` (มีช่องว่างรอบ `&`) ❌
   - `MA&Optimize NER ` (มีช่องว่างท้าย) ❌
4. ถ้าชื่อเปลี่ยนไปแล้วจริง ๆ ให้แก้บรรทัดแรกของสคริปต์:
   ```typescript
   const SHEET_SOURCE = "ชื่อชีตใหม่";
   ```

---

## 5. Run script error: `ไม่พบคอลัมน์: ...`

สคริปต์อ่านหัวตารางจาก **แถวที่ 7**

**ตรวจ**

- เซลล์ `A7` ต้องเป็น `Distribution period`
- ถ้าหัวตารางย้ายไปแถวอื่น ให้แก้ในสคริปต์:
  ```typescript
  const HEADER_ROW = 7;   // เปลี่ยนเป็นแถวที่ถูกต้อง
  ```
- ชื่อคอลัมน์ที่ต้องมีครบ: `Distribution period`, `Zone`, `Item Code`,
  `OMC-Onhand`, `เบิกจาก Hub`, `Prev_Period`, `Prev_Before`, `Prev_Received`,
  `Status`, `Risk_Flag`, `Data_Set`
- ระวังคอลัมน์ไทยที่มีช่องว่างท้ายชื่อ เช่น `เบิกจาก Hub ` — สคริปต์ตัดช่องว่างให้แล้ว
  แต่ถ้าสะกดต่างจริง ๆ จะหาไม่เจอ

---

## 6. การ์ดขึ้น empty state ทั้งที่ในไฟล์มีข้อมูล

**สาเหตุอันดับ 1: ลืม Save As ทับไฟล์ `.xlsx`**

Flow อ่านไฟล์ `.xlsx` ไม่ใช่ `.xlsm` — รีเฟรชใน `.xlsm` อย่างเดียวไม่พอ

**สาเหตุอื่น**

| สาเหตุ | วิธีตรวจ |
|---|---|
| รอบที่ระบุใน `period` สะกดผิด | ต้องตรงเป๊ะ เช่น `MA (17 Aug 26)` — มีวงเล็บและช่องว่าง |
| รอบนั้นไม่มีรหัส 53OF/52CL จริง ๆ | เปิดชีตแล้วกรองคอลัมน์ `Item Code` ดู |
| ยังไม่ได้กด Refresh Data New Sheet | ดูค่าในคอลัมน์ `Distribution period` |

**วิธีตรวจเร็ว:** เปิดชีต `MA&Optimize NER` → กรอง `Item Code`
ด้วย "begins with" `53OF` แล้วดูว่ามีแถวของรอบที่ต้องการหรือไม่

---

## 7. Flow ยิงผิดเวลา

| อาการ | สาเหตุ | แก้ |
|---|---|---|
| ยิงบ่าย 3 แทน 8 โมง | Time zone เป็น UTC | Recurrence → Time zone = `(UTC+07:00) Bangkok` |
| ยิง 60 ครั้งติดกัน | ไม่ได้ใส่ At these minutes | ใส่ `0` |
| ยิงทุกวัน | ไม่ได้เลือก On these days | เลือก `Monday` |
| ไม่ยิงเลย | Flow ถูกปิด | My flows → เปิดสวิตช์เป็น `On` |
| ไม่ยิงหลังไม่มีคนใช้ 90 วัน | Microsoft ปิด Flow อัตโนมัติ | เข้าไปเปิดใหม่ |

---

## 8. ตัวเลขบนการ์ดผิดคอลัมน์ (Prev_Before ได้ค่าของ Prev_Received)

**สาเหตุ:** ไปดึงข้อมูลจากชีต `PA_Export` ซึ่งมีบั๊กหัวตารางเลื่อน 1 คอลัมน์

**วิธีแก้ที่ง่ายที่สุด:** ใช้ `renderAdaptiveCard` (อ่านชีตต้นทางโดยตรง ไม่โดนบั๊ก)

**ถ้าจำเป็นต้องใช้ `PA_Export`:** ติดตั้ง `vba/modExport_FIXED.bas`
แล้วรัน `BuildExportTable_Click` ใหม่

รายละเอียดเต็ม: [03-DATA-MAPPING.md ข้อ 4](03-DATA-MAPPING.md#4-️-บั๊กที่พบในไฟล์-v2--หัวตาราง-pa_export-เลื่อน-1-คอลัมน์)

---

## 9. เลย์เอาต์มือถือไม่สลับ ยังเป็นตารางแคบ ๆ

การ์ดหลักใช้คุณสมบัติ `targetWidth` ซึ่งต้องใช้ Teams เวอร์ชันที่รองรับ

**ถ้าไคลเอนต์ไม่รองรับ** จะแสดงทั้งตารางและบรรทัดสรุปพร้อมกัน (ซ้ำกัน แต่ยังอ่านได้)

**วิธีแก้:** เปลี่ยนไปใช้ `adaptive-cards/04_card-mobile-stacked.json`
ซึ่งเป็นเลย์เอาต์แนวตั้งเสมอ ไม่พึ่ง `targetWidth`

วิธีเปลี่ยน:
1. แก้ `tools/build_office_script.js` บรรทัด `const CARD = ...`
   ให้ชี้ไปที่ `04_card-mobile-stacked.json`
2. รัน `node tools/build_office_script.js`
3. คัดลอกสคริปต์ใหม่ไปวางทับใน Excel

---

## 10. Excel Online เปิดไฟล์ไม่ได้ / Run script ค้าง

| สาเหตุ | แก้ |
|---|---|
| ไฟล์ยังเปิดค้างอยู่บนเครื่องใครสักคน | ปิดไฟล์ให้หมดแล้วลองใหม่ |
| ไฟล์เป็น `.xlsm` | Office Scripts ไม่รองรับ ต้องใช้ `.xlsx` |
| ไฟล์ใหญ่เกิน | ลบชีตรอบเก่าที่ไม่ใช้ออกจากสำเนา `.xlsx` |
| ไฟล์อยู่ใน OneDrive ส่วนตัว | ย้ายไป SharePoint ไลบรารีทีม |

---

## 11. อยากดูว่าสคริปต์ส่งอะไรออกไปกันแน่

1. Run history → กด action **Run script**
2. ดูส่วน **OUTPUTS → result** — นี่คือ JSON การ์ดทั้งใบ
3. คัดลอกไปวางใน [adaptivecards.io/designer](https://adaptivecards.io/designer/)
   ช่อง **Card Payload Editor** (ช่อง Sample Data ปล่อยว่าง เพราะแทนค่ามาแล้ว)
4. เลือก Host app = **Microsoft Teams** แล้วดูว่าการ์ดหน้าตาเป็นยังไง

วิธีนี้แยกได้ทันทีว่าปัญหาอยู่ที่ **ข้อมูล** หรือที่ **การส่งเข้า Teams**

---

## 12. ทดสอบทั้งระบบก่อนแก้อะไรในโค้ด

```bash
python3 tools/build_sample_data.py <path-to-xlsm>
python3 tools/dump_sheet_values.py <path-to-xlsm>

node tools/expand_and_validate.js     # template expand ได้จริงไหม
node tools/build_office_script.js     # สร้าง renderAdaptiveCard.ts ใหม่
node tools/test_office_script.js      # ตรรกะข้อมูลตรงกับตัวอ้างอิงไหม
node tools/test_render_script.js      # การ์ดตรงกับไลบรารีจริงทุก byte ไหม
```

ทั้ง 4 คำสั่งต้องขึ้น ✅ ก่อนเอาสคริปต์ไปวางใน Excel
