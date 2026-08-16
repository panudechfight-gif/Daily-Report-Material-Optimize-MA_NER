# คู่มือสร้าง Flow ส่ง Adaptive Card เข้า Microsoft Teams
### ฉบับมือใหม่ ทำตามได้เอง ทีละขั้น

Report Material Optimize&MA_NER
เฉพาะรหัสวัสดุ **8 รหัส** ที่กำหนดไว้ (ดู [ข้อ 5.2](#52-เปลี่ยนรหัสวัสดุที่แสดง))

> 📄 มีฉบับหน้าเว็บแบบพับเก็บได้ทีละ Action ที่ [`flow-guide-artifact.html`](flow-guide-artifact.html)

---

## สารบัญ

| ตอน | เรื่อง | เวลาที่ใช้ |
|---|---|---|
| [ตอนที่ 0](#ตอนที่-0-เข้าใจภาพรวมก่อนลงมือ) | เข้าใจภาพรวมก่อนลงมือ | 5 นาที |
| [ตอนที่ 1](#ตอนที่-1-เตรียมไฟล์-excel-บน-sharepoint) | เตรียมไฟล์ Excel บน SharePoint | 15 นาที |
| [ตอนที่ 2](#ตอนที่-2-ติดตั้ง-office-script) | ติดตั้ง Office Script | 10 นาที |
| [ตอนที่ 3](#ตอนที่-3-สร้าง-flow-ทีละ-action) | สร้าง Flow ทีละ Action | 30 นาที |
| [ตอนที่ 4](#ตอนที่-4-ทดสอบ) | ทดสอบ | 10 นาที |
| [ตอนที่ 5](#ตอนที่-5-ปรับแต่ง) | ปรับแต่งให้เข้ากับงานจริง | ตามต้องการ |
| [ภาคผนวก ก](#ภาคผนวก-ก-ใช้-adaptive-cards-designer) | ใช้ Adaptive Cards Designer | — |
| [ภาคผนวก ข](#ภาคผนวก-ข-เส้นทางสำรอง-ไม่ใช้-office-script) | เส้นทางสำรอง (ไม่ใช้ Office Script) | — |

---

## ตอนที่ 0: เข้าใจภาพรวมก่อนลงมือ

### 0.1 Flow นี้ทำอะไร

```
   ⏰ ทุกวันจันทร์ 08:00 น.
            │
            ▼
   📗 เปิดไฟล์ Excel บน SharePoint
      อ่านชีต "MA&Optimize NER"
            │
            ▼
   🔍 คัดเฉพาะ 8 รหัสที่กำหนด
      ของรอบล่าสุด แล้วจัดกลุ่มตาม Zone
            │
            ▼
   🎨 ประกอบเป็น Adaptive Card
      (สี/emoji ตาม Data_Set และ Status)
            │
            ▼
   💬 โพสต์เข้าห้อง Teams
      พร้อมปุ่ม 📊 Dashboard และ 📁 เปิดไฟล์ต้นทาง
```

### 0.2 เรื่องสำคัญที่ต้องเข้าใจก่อน — เรื่องนี้ทำให้คนติดกันเยอะที่สุด

> **Power Automate ไม่รู้จัก `${...}`**
>
> ไฟล์ `adaptive-cards/01_card-main.json` เป็น **template** ที่ยังมีตัวแปร
> `${meta.period}`, `$data`, `$when` อยู่ ซึ่ง **ใช้ได้เฉพาะใน
> adaptivecards.io/designer** สำหรับออกแบบและดูตัวอย่างเท่านั้น
>
> action **"Post adaptive card in a chat or channel"** ของ Teams
> รับได้เฉพาะ JSON ที่ **แทนค่าเสร็จแล้ว** ถ้าวาง template ดิบ ๆ ลงไป
> การ์ดจะขึ้นเป็นข้อความ `${meta.period}` ตรง ๆ หรือส่งไม่ออกเลย

**แล้วใครเป็นคนแทนค่า?**

คู่มือนี้ให้ **Office Script** เป็นคนทำ — สคริปต์ `renderAdaptiveCard.ts`
รวมทั้งการอ่านข้อมูล การแทนค่า และการคืน JSON การ์ดที่สมบูรณ์ ไว้ในแอ็กชันเดียว

ผลลัพธ์ที่สคริปต์นี้สร้าง **ตรงกับไลบรารี `adaptivecards-templating` ทุก byte**
(พิสูจน์ด้วย `node tools/test_render_script.js`) แปลว่า **สิ่งที่เห็นใน Designer
คือสิ่งที่ Teams จะได้จริง**

### 0.3 สิ่งที่ต้องมี

- [ ] บัญชี Microsoft 365 ที่ใช้ Power Automate ได้
- [ ] สิทธิ์เขียนไฟล์ในไซต์ SharePoint ที่จะเก็บ Excel
- [ ] สิทธิ์โพสต์ในห้อง Teams ปลายทาง
- [ ] ไฟล์ `Optimize_HiringMA_Full_Rebuild_VBA_Macro_V2.xlsm`

---

## ตอนที่ 1: เตรียมไฟล์ Excel บน SharePoint

### 1.1 แปลงไฟล์เป็น .xlsx

> ⚠️ **Office Scripts ไม่รองรับไฟล์ .xlsm** (ไฟล์ที่มีมาโคร)
> ต้องมีไฟล์ `.xlsx` อีกหนึ่งไฟล์ให้ Flow อ่าน

**ทำไมต้องมี 2 ไฟล์**

| ไฟล์ | ใช้ทำอะไร | ใครใช้ |
|---|---|---|
| `.xlsm` (ต้นฉบับ) | รันมาโคร Refresh / BuildExportTable | คนทำงาน |
| `.xlsx` (สำเนา) | ให้ Flow อ่านอัตโนมัติ | Power Automate |

**ขั้นตอน**

1. เปิดไฟล์ `.xlsm` ด้วย Excel บนเครื่อง
2. กดปุ่ม **Refresh Data New Sheet** ในชีต `MA&Optimize NER` ให้ข้อมูลเป็นรอบล่าสุด
3. **File → Save As**
4. เลือกที่เก็บเป็นไซต์ SharePoint เช่น
   `https://<องค์กร>.sharepoint.com/sites/NER/Shared Documents/`
5. **Save as type** เลือก **Excel Workbook (\*.xlsx)**
6. ตั้งชื่อ `Optimize_HiringMA_Full_Rebuild_V2.xlsx`
7. กด Save → Excel จะเตือนว่ามาโครจะหายไป ให้กด **Yes** (ถูกต้องแล้ว ไฟล์นี้ไม่ต้องใช้มาโคร)

### 1.2 คัดลอกลิงก์ไฟล์เก็บไว้

เปิดไฟล์ `.xlsx` บน SharePoint → กด **Copy link** → เก็บลิงก์ไว้
ลิงก์นี้จะใช้กับปุ่ม **📁 Open the source file** บนการ์ด

### 1.3 ตรวจว่าชีตถูกต้อง

เปิดไฟล์ `.xlsx` ใน Excel for the web แล้วตรวจว่า:

- [ ] มีชีตชื่อ **`MA&Optimize NER`** (ชื่อต้องตรงเป๊ะ รวมเครื่องหมาย `&`)
- [ ] **แถวที่ 7** คือหัวตาราง — เซลล์ A7 ต้องเป็นคำว่า `Distribution period`
- [ ] แถวที่ 8 เป็นต้นไปเป็นข้อมูล
- [ ] มีคอลัมน์ครบ: `Zone`, `Item Code`, `OMC-Onhand`, `เบิกจาก Hub`,
      `Prev_Period`, `Prev_Before`, `Prev_Received`, `Status`, `Risk_Flag`, `Data_Set`

---

## ตอนที่ 2: ติดตั้ง Office Script

### 2.1 เปิดตัวแก้ไขสคริปต์

1. เปิดไฟล์ `.xlsx` ใน **Excel for the web** (เปิดจาก SharePoint โดยตรง)
2. แถบเมนูบนสุด → แท็บ **Automate**
3. กด **New Script**
4. จะมีหน้าต่าง Code Editor เปิดขึ้นมาทางขวา พร้อมโค้ดตัวอย่าง

### 2.2 วางโค้ด

1. **ลบโค้ดตัวอย่างทิ้งให้หมด** (Ctrl+A แล้ว Delete)
2. เปิดไฟล์ `office-scripts/renderAdaptiveCard.ts` จากโปรเจกต์นี้
3. คัดลอก **ทั้งไฟล์** แล้ววางลงใน Code Editor
4. กด **Save script**
5. เปลี่ยนชื่อสคริปต์ (คลิกที่ชื่อด้านบน) เป็น **`renderAdaptiveCard`**

> 💡 ถ้าขึ้นขีดหยักแดงเรื่อง `ExcelScript` ให้กด Save แล้วรอสักครู่ —
> ตัวตรวจไวยากรณ์บางครั้งโหลดช้ากว่าโค้ด

### 2.3 ทดสอบสคริปต์ทันที

1. กด **Run** ในตัวแก้ไข
2. ดูช่อง Output ด้านล่าง — ต้องได้ JSON ก้อนใหญ่ที่ขึ้นต้นด้วย
   `{"type":"AdaptiveCard","$schema":...`
3. ถ้า error ให้ดู [ตอนที่ 4.3](#43-เจอปัญหา-ดูตรงนี้)

> ✅ **ถึงตรงนี้แล้ว ส่วนที่ยากที่สุดผ่านไปแล้ว** ที่เหลือเป็นการต่อ Flow

---

## ตอนที่ 3: สร้าง Flow ทีละ Action

ไปที่ [make.powerautomate.com](https://make.powerautomate.com)
→ **Create** → **Scheduled cloud flow**

**ตั้งค่าหน้าแรก**

| ช่อง | ใส่ค่า |
|---|---|
| Flow name | `รายงานวัสดุคงคลัง NER รายสัปดาห์` |
| Starting | วันที่จะเริ่ม |
| Repeat every | `1` `Week` |
| On these days | เลือก **M** (จันทร์) |

กด **Create**

---

### Action 1 — Recurrence (ตัวจับเวลา)

Flow สร้าง action นี้ให้อัตโนมัติแล้ว ให้กดเข้าไปตั้งค่าเพิ่ม:

1. กดที่ **Recurrence**
2. กด **Show advanced options**

| ช่อง | ใส่ค่า | หมายเหตุ |
|---|---|---|
| Frequency | `Week` | |
| Interval | `1` | |
| Time zone | `(UTC+07:00) Bangkok, Hanoi, Jakarta` | **สำคัญ** ถ้าไม่ตั้ง Flow จะยิงตามเวลา UTC (บ่าย 3 บ้านเรา) |
| On these days | `Monday` | |
| At these hours | `8` | |
| At these minutes | `0` | |

> ⚠️ ถ้าไม่ใส่ **At these minutes = 0** Flow จะยิงทุกนาทีในชั่วโมงนั้น (60 ครั้ง!)

---

### Action 2 — Compose: เวลาที่สร้างรายงาน

กด **+ New step** → ค้นหา `Compose` → เลือก **Compose** (Data Operation)

- เปลี่ยนชื่อ action เป็น **`GeneratedAt`** (คลิกจุดสามจุด → Rename)
- ช่อง **Inputs** กดไอคอน **fx** แล้ววางสูตรนี้:

```
concat(
  formatDateTime(convertFromUtc(utcNow(),'SE Asia Standard Time'),'d MMM '),
  string(add(int(formatDateTime(convertFromUtc(utcNow(),'SE Asia Standard Time'),'yyyy')),543)),
  ' ',
  formatDateTime(convertFromUtc(utcNow(),'SE Asia Standard Time'),'HH:mm'),
  ' น.'
)
```

จะได้ผลลัพธ์แบบ `17 Aug 2569 08:00 น.` (ปี พ.ศ.)

> 💡 อยากได้ปี ค.ศ. แบบสั้น ๆ ใช้แค่บรรทัดเดียวพอ:
> `formatDateTime(convertFromUtc(utcNow(),'SE Asia Standard Time'),'dd/MM/yyyy HH:mm')`

---

### Action 3 — Run script (หัวใจของ Flow)

กด **+ New step** → ค้นหา `Run script` → เลือก
**Excel Online (Business) → Run script**

| ช่อง | ใส่ค่า |
|---|---|
| Location | `SharePoint Sites` |
| Document Library | `Documents` (หรือไลบรารีที่เก็บไฟล์) |
| File | กดไอคอนโฟลเดอร์ แล้วเลือก `Optimize_HiringMA_Full_Rebuild_V2.xlsx` |
| Script | `renderAdaptiveCard` |

**พอเลือก Script แล้ว จะมีช่องพารามิเตอร์โผล่มาเพิ่ม 5 ช่อง:**

| พารามิเตอร์ | ใส่ค่า | ความหมาย |
|---|---|---|
| `period` | `-` | `-` = รอบ **Optimize** ล่าสุดอัตโนมัติ (ดูกล่องข้างล่าง) |
| `topRows` | `40` | เพดานแถว — สคริปต์วัดขนาดแล้วลดให้เองถ้าจำเป็น |
| `dashboardUrl` | ลิงก์ Dashboard ของคุณ | ใช้กับปุ่ม 📊 |
| `sourceFileUrl` | ลิงก์ไฟล์ Excel จากข้อ 1.2 | ใช้กับปุ่ม 📁 |
| `generatedAt` | เลือก **Outputs** ของ `GeneratedAt` | เวลาที่สร้างรายงาน |

> ⚠️ **ช่อง `period` เว้นว่างไม่ได้**
>
> Power Automate บังคับให้พารามิเตอร์ของ Run script ต้องมีค่า ถ้าเว้นว่างจะขึ้น
> `Invalid parameter for 'Run script'. Error: 'ScriptParameters/period' is required.`
>
> ค่าที่ใส่ได้ในช่องนี้:
>
> | ใส่ | ได้อะไร |
> |---|---|
> | `-` `auto` `latest` `null` | รอบ **Optimize** ล่าสุด (รายงานหลัก) |
> | `Optimize` | รอบ Optimize ล่าสุด |
> | `MA` | รอบ **MA** ล่าสุด |
> | `Optimize(3 Aug 26)` | เจาะจงรอบนั้น |
> | `MA (17 Aug 26)` | เจาะจงรอบนั้น |
>
> พิมพ์เล็กใหญ่ไม่สำคัญสำหรับคำสั่ง (`-`, `auto`, `ma`, `optimize`)
> แต่ชื่อรอบต้องตรงเป๊ะรวมวงเล็บและช่องว่าง
>
> > ⚠️ อย่าใช้ `fx` แล้วใส่ `''` — Power Automate ยังนับเป็นค่าว่างและ error เหมือนเดิม
> > ให้พิมพ์ `-` ลงในช่องตรง ๆ เป็นข้อความธรรมดา

> 💡 **`topRows` ไม่ต้องคำนวณเองแล้ว**
>
> Teams ปฏิเสธ Adaptive Card ที่ใหญ่เกิน **28 KB** เงียบ ๆ โดยไม่บอกสาเหตุ
> สคริปต์จึงประกอบการ์ดแล้ว **วัดขนาดเป็นไบต์จริง** ถ้าเกินงบ 26,000 ไบต์
> ก็ลดแถวแล้วประกอบใหม่ให้เอง
>
> ใส่ `40` ไว้เป็นเพดานก็พอ ขนาดจริงที่วัดได้:
>
> | รอบ | Zone | แถวดิบ | แถวบนการ์ด | ขนาด |
> |---|---|---|---|---|
> | Optimize (3 Aug 26) | 5 | 35 | 35 | 24.9 KB ✅ |
> | MA (17 Aug 26) | 4 | 74 | 24 | 19.4 KB ✅ |
>
> **ทุกรอบลงการ์ดใบเดียวได้ครบ ไม่มีการตัดทิ้ง** เพราะการ์ดรวมแถวซ้ำให้แล้ว
> (1 Zone + 1 Item Code = 1 แถว — ยอดของทุกจังหวัดถูกบวกเข้าด้วยกัน)
>
> ถ้ารอบไหนใหญ่จนต้องตัดจริง ๆ ส่วนที่เหลือจะสรุปเป็นบรรทัด
> *"…และอีก N รายการ"* พร้อมปุ่มไป Dashboard

---

### Action 4 — Post adaptive card in a chat or channel

กด **+ New step** → ค้นหา `Post adaptive card` → เลือก
**Microsoft Teams → Post adaptive card in a chat or channel**

| ช่อง | ใส่ค่า |
|---|---|
| Post as | `Flow bot` |
| Post in | `Channel` |
| Team | เลือกทีมของคุณ |
| Channel | เลือกห้องปลายทาง |
| Adaptive Card | ดูข้างล่าง ⬇️ |

**ช่อง Adaptive Card** — กดไอคอน **fx** แล้ววาง:

```
string(body('Run_script')?['result'])
```

> 💡 **ถ้าสูตรขึ้นแดง** แปลว่าชื่อ action ของคุณไม่ใช่ `Run_script`
> ให้ดูชื่อจริงที่หัว action แล้วแทนที่ในสูตร โดยเปลี่ยนช่องว่างเป็น `_`
> เช่น action ชื่อ `Run script` → `body('Run_script')`
>
> อีกวิธีคือกดแท็บ **Dynamic content** แล้วเลือก **result** จากรายการ

> ⚠️ **อย่าลืม `string(...)`** ถ้าใส่แค่ `body('Run_script')?['result']`
> Power Automate จะส่งเป็น object แล้วการ์ดอาจไม่ขึ้น

กด **Save**

---

### ✅ Flow ที่เสร็จแล้วมี 4 Action

```
1. Recurrence            ⏰ ทุกวันจันทร์ 08:00 (Bangkok)
2. GeneratedAt           📅 Compose เวลาไทย
3. Run script            📗 renderAdaptiveCard -> JSON การ์ด
4. Post adaptive card    💬 โพสต์เข้า Teams
```

> 💬 **ไม่ต้องมี Condition เช็ก "ไม่มีข้อมูล"**
> การ์ดจัดการเองด้วย `$when` — ถ้ารอบนั้นไม่มีรหัสที่กำหนดเลย
> จะแสดง **empty state** พร้อมคำแนะนำ 3 ข้อว่าต้องไปตรวจอะไรบ้าง

---

## ตอนที่ 4: ทดสอบ

### 4.1 รันด้วยมือ

1. หน้า Flow กด **Test** (มุมขวาบน)
2. เลือก **Manually** → **Test**
3. กด **Run flow** → **Done**
4. รอสักครู่ ทุก action ควรขึ้นเครื่องหมายถูกสีเขียว

### 4.2 ตรวจการ์ดใน Teams

เปิดห้อง Teams ที่ตั้งไว้ ต้องเห็นการ์ดที่มี:

- [ ] หัวการ์ดเขียนว่า **Report Material Optimize&MA_NER**
- [ ] ป้ายชุดข้อมูลถูกต้อง — 📦 **MA — Weekly Allocation** (ฟ้า)
      หรือ 🗃️ **OPTIMIZE — Allocation Plan** (เขียว)
- [ ] แถบ 3 ช่อง: รอบจัดสรร / รอบก่อนหน้า / ชุดข้อมูล
- [ ] แถบตัวเลข 4 กล่อง: ⚠️ เสี่ยงขาด, 🟡 พอใช้, ✅ ปกติ, 📦 รวม
- [ ] **หัวกลุ่มขึ้นครบทุก Zone** พร้อม emoji
      (🟦 RC2-NMA, 🟩 RC2-UBN, 🟨 RC3-KKN, 🟧 RC3-UDN, 🟪 RC3-SNK — เท่าที่รอบนั้นมีข้อมูล)
- [ ] ตาราง 5 คอลัมน์: Item Code / OMC-Onhand / เบิกจาก Hub / Prev_Before / Prev_Received
      — **ไม่มี Province**
- [ ] คอลัมน์ **เบิกจาก Hub เป็นจำนวนเต็ม** ไม่มีจุดทศนิยม
- [ ] บรรทัด "…และอีก N รายการ"
- [ ] ปุ่ม **📊 Dashboard** และ **📁 Open the source file. (Excel / SharePoint)**

**เปิดดูบนมือถือด้วย** — Teams บนมือถือจะสลับเป็นเลย์เอาต์แนวตั้ง
ตัวเลขทั้ง 4 ค่าจะย้ายมาอยู่ใต้ชื่อรายการแทนที่จะเป็นคอลัมน์

### 4.3 เจอปัญหา? ดูตรงนี้

| อาการ | สาเหตุ | วิธีแก้ |
|---|---|---|
| กด Save แล้วขึ้น `'ScriptParameters/period' is required` | ช่อง `period` ว่าง | ใส่ `-` ลงไปตรง ๆ |
| การ์ดไม่ขึ้นเลย แต่ Flow เขียว | ใช้สคริปต์เก่าที่ยังไม่วัดขนาดเอง | คัดลอก `renderAdaptiveCard.ts` ตัวล่าสุดไปวางทับใน Excel |
| Zone บางโซนหายไปทั้งกลุ่ม | ใช้สคริปต์เวอร์ชันเก่า | คัดลอก `renderAdaptiveCard.ts` ตัวล่าสุดไปวางทับใน Excel |
| เบิกจาก Hub ยังมีทศนิยม | ใช้สคริปต์เวอร์ชันเก่า | คัดลอก `renderAdaptiveCard.ts` ตัวล่าสุดไปวางทับใน Excel |
| การ์ดขึ้นเป็นข้อความ `${meta.period}` | เอา template ดิบไปวางในช่อง Adaptive Card | ต้องใช้ `string(body('Run_script')?['result'])` |
| `Run script` แดง: ไม่พบชีต | ชื่อชีตไม่ตรง | ต้องเป็น `MA&Optimize NER` เป๊ะ ๆ |
| `Run script` แดง: ไม่พบคอลัมน์ | หัวตารางไม่ได้อยู่แถว 7 | ตรวจว่า A7 = `Distribution period` |
| การ์ดขึ้น empty state ทั้งที่มีข้อมูล | ไฟล์ .xlsx ยังเป็นข้อมูลเก่า | Refresh ใน .xlsm แล้ว Save As ทับ .xlsx ใหม่ |
| Flow ยิงตอนบ่าย 3 | ไม่ได้ตั้ง Time zone | Recurrence → Time zone = Bangkok |
| Flow ยิง 60 ครั้ง | ไม่ได้ตั้ง At these minutes | ใส่ `0` |
| ตัวเลข Prev_Before ผิดคอลัมน์ | ไปดึงจากชีต PA_Export ที่มีบั๊ก | ดู [docs/03-DATA-MAPPING.md](03-DATA-MAPPING.md) |

รายละเอียดเพิ่มเติม: [docs/04-TROUBLESHOOTING.md](04-TROUBLESHOOTING.md)

---

## ตอนที่ 5: ปรับแต่ง

### 5.1 เปลี่ยนวัน/เวลาที่ส่ง

Recurrence → แก้ **On these days** / **At these hours**
ส่งทุกวัน: เอาเครื่องหมายถูกออกจาก On these days ให้หมด

### 5.2 เปลี่ยนรหัสวัสดุที่แสดง

รหัสที่รายงานถูกกำหนดเป็นรายการตายตัวใน `TARGET_ITEMS` — แก้ที่นี่ที่เดียว
ค่าของแต่ละรหัสคือชื่อที่ใช้แสดงในตารางอธิบายท้ายการ์ด:

```typescript
const TARGET_ITEMS: { [k: string]: string } = {
  "53OF150BB": "OPTICAL FIBER DROP CABLE 1C, FLAT TYPE (G.657A) WITH 2 SC/UPC PRE-CONNECTOR, 3m.",
  "53OF157BB": "ARSS OPTICAL FIBER CABLE 12c-FIBRE3",
  "53OF158BB": "ARSS OPTICAL FIBER CABLE 24c-FIBRE3",
  "53OF160BB": "ARSS OPTICAL FIBER CABLE 60c-FIBRE3",
  "50MT004BB": "Name Plate (Aluminium)",
  "52CL003BB": "CLOSURE 12 C",
  "52CL004BB": "CLOSURE 24 C",
  "52CL010BB": "CLOSURE 60 C",
};
```

**เพิ่มรหัส** ให้ใส่บรรทัดใหม่พร้อมชื่อ เช่น `"53OF155BB": "ชื่อวัสดุ",`
**ลบรหัส** ให้ลบบรรทัดนั้นทิ้ง

> 💡 ถ้าแก้ในโปรเจกต์ (ไม่ใช่แก้ในกล่อง Excel ตรง ๆ) ให้แก้ที่
> `office-scripts/buildCardPayload.ts` และ `tools/build_sample_data.py`
> ให้ตรงกัน แล้วรัน `node tools/build_office_script.js` กับ `npm test`

### 5.3 ส่งรอบที่ต้องการแบบเจาะจง

พารามิเตอร์ `period` ใส่ชื่อรอบตรง ๆ เช่น `MA (10 Aug 26)` หรือ `Optimize(3 Aug 26)`

### 5.4 ส่งทั้ง OPTIMIZE (เดือน) และ MA (สัปดาห์)

**ไม่ต้องแก้ชื่อรอบทุกครั้ง** — ใส่ชื่อชุดข้อมูลลงในช่อง `period` ได้เลย
สคริปต์จะไปหยิบรอบล่าสุดของชุดนั้นให้เอง

| Flow | Recurrence | `period` | ได้อะไร |
|---|---|---|---|
| รายเดือน | วันที่ 1 เวลา 08:00 | `Optimize` | รอบ Optimize ล่าสุด |
| รายสัปดาห์ | ทุกจันทร์ 08:00 | `MA` | รอบ MA ล่าสุด |

**วิธีที่แนะนำ:** แยกเป็น **2 Flow** ไปเลย เพราะรอบเวลาไม่เท่ากัน
(เดือนละครั้ง กับสัปดาห์ละครั้ง) การใส่รวมใน Flow เดียวจะทำให้
Optimize ถูกส่งซ้ำทุกสัปดาห์โดยไม่จำเป็น

ถ้าอยากได้ทั้งสองใบในทีเดียวจริง ๆ ให้เพิ่ม **Run script** + **Post adaptive card**
อีกคู่หนึ่งใน Flow เดิม แล้วตั้ง `period` ของคู่ที่สองเป็น `MA`

> 📖 แนวทางเลือกจังหวะและเนื้อหาของรายงานทั้งสองชุด
> ดู [ข้อ 5.6](#56-แนวทางรายงาน-optimize-เดือน-กับ-ma-สัปดาห์)

### 5.5 แก้หน้าตาการ์ด

1. แก้ไฟล์ `adaptive-cards/01_card-main.json`
2. ทดสอบใน Designer (ดู [ภาคผนวก ก](#ภาคผนวก-ก-ใช้-adaptive-cards-designer))
3. รัน `node tools/build_office_script.js` เพื่อสร้าง `renderAdaptiveCard.ts` ใหม่
4. รัน `node tools/test_render_script.js` ให้ผ่าน
5. คัดลอกไปวางทับสคริปต์เดิมใน Excel

> ⚠️ **อย่าแก้ `renderAdaptiveCard.ts` โดยตรง** — ไฟล์นี้ถูกสร้างอัตโนมัติ
> แก้ไปก็จะหายตอนสร้างใหม่ ให้แก้ที่ไฟล์ JSON เท่านั้น

### 5.6 แนวทางรายงาน OPTIMIZE (เดือน) กับ MA (สัปดาห์)

สองชุดนี้ตอบคำถามคนละข้อ ถ้าส่งเหมือนกันทั้งคู่ คนอ่านจะเริ่มข้ามทั้งสองใบ

| | 🗃️ OPTIMIZE — รอบเดือน | 📦 MA — รอบสัปดาห์ |
|---|---|---|
| ตอบคำถาม | "เดือนหน้าต้องเตรียมของเท่าไร" | "สัปดาห์นี้มีอะไรจะขาด" |
| คนอ่าน | หัวหน้าโซน / คนวางแผนจัดซื้อ | หน้างาน / คนเบิกของ |
| สิ่งที่ควรเด่น | ยอดรวมทั้งโซน แนวโน้มเทียบรอบก่อน | รายการเสี่ยงขาด ต้องเติมด่วน |
| จังหวะที่แนะนำ | วันที่ 1 ของเดือน 08:00 | ทุกวันจันทร์ 08:00 |
| ควรทำอะไรต่อ | อนุมัติแผนจัดสรรรอบถัดไป | เร่งเบิก/ย้ายของระหว่างจังหวัด |

**ข้อแนะนำที่ใช้ได้จริง 5 ข้อ**

1. **แยก Flow กัน** อย่ายัดรวมใบเดียว — คนละกลุ่มผู้อ่าน คนละความถี่
   ถ้ารวมกัน คนหน้างานจะต้องเลื่อนผ่านข้อมูลรายเดือนทุกสัปดาห์

2. **แยกห้อง Teams หรืออย่างน้อยแยก thread** ให้ MA อยู่ห้องปฏิบัติการ
   ส่วน OPTIMIZE อยู่ห้องวางแผน คนจะได้ไม่ต้องกรองเอง

3. **MA ควรสั้น** ถ้าสัปดาห์ไหนไม่มีรายการเสี่ยงเลย การ์ดที่ขึ้นว่า
   "✅ ปกติทั้งหมด" หนึ่งบรรทัดมีค่ากว่าตารางยาว ๆ —
   ทำได้ด้วย Condition ใน Flow เช็ก `riskCount` จาก Run script ก่อนโพสต์

4. **OPTIMIZE ควรมีการเทียบ** ตัวเลข `Prev_Before` / `Prev_Received` บนการ์ด
   คือจุดที่บอกว่า "รอบก่อนใช้ไปเท่าไร" ให้ผู้อ่านดูคู่กับ `OMC-Onhand` เสมอ
   ถ้าอยากได้กราฟแนวโน้มจริง ๆ ให้ลิงก์ไป Dashboard แทนการยัดลงการ์ด

5. **ให้ทั้งสองใบชี้ไป Dashboard เดียวกัน** การ์ดคือ "สัญญาณเตือน" ไม่ใช่รายงานฉบับเต็ม
   อะไรที่ต้องกรอง/ไล่ดูรายจังหวัด ให้จบที่ Dashboard

> 💡 **เรื่องที่คนมักพลาด:** ถ้าจะเทียบ MA กับ OPTIMIZE ตรง ๆ ระวังว่า
> ยอดบนการ์ดเป็น **ผลรวมทุกจังหวัดในโซน** แล้ว การเทียบรายจังหวัดต้องไปดูที่ Dashboard

---

## ภาคผนวก ก: ใช้ Adaptive Cards Designer

ใช้ตอนอยากลองปรับหน้าตาการ์ดโดยไม่ต้องรัน Flow จริง

1. เปิด [adaptivecards.io/designer](https://adaptivecards.io/designer/)
2. **Select host app** เลือก **Microsoft Teams** (สำคัญ — สีและระยะห่างจะตรงกับของจริง)
3. ช่อง **Card Payload Editor** วางเนื้อไฟล์ `adaptive-cards/01_card-main.json`
4. ช่อง **Sample Data Editor** วางเนื้อไฟล์ `adaptive-cards/data/sample-data-main.json`
5. กด **Preview mode** → การ์ดจะแสดงข้อมูลจริง

**ไฟล์ตัวอย่างที่ให้มา**

| ไฟล์ข้อมูล | ใช้ดูอะไร |
|---|---|
| `sample-data-main.json` | รอบล่าสุด Optimize (3 Aug 26) — 35 รายการ 5 Zone แสดงครบ |
| `sample-data-risk.json` | รอบ MA (10 Aug 26) — 7 รายการ มี 6 รายการเสี่ยงขาด ดูสีแดง/ส้ม |
| `sample-data-empty.json` | ไม่มีข้อมูล — ดู empty state |
| `sample-data-flat.json` | ใช้กับ `02_card-simple-flat.json` |

**การ์ดทั้ง 4 แบบ**

| ไฟล์ | ใช้เมื่อไหร่ |
|---|---|
| `01_card-main.json` | **ตัวหลัก** จัดกลุ่มตาม Zone + ปรับตามความกว้างจอ |
| `02_card-simple-flat.json` | อยากได้ตารางเดียวไม่แยก Zone (เบากว่า ใส่ได้หลายแถวกว่า) |
| `03_card-empty-state.json` | แยกการ์ด "ไม่มีข้อมูล" ออกมาต่างหาก (ถ้าใช้ Condition ใน Flow) |
| `04_card-mobile-stacked.json` | บังคับเลย์เอาต์แนวตั้งเสมอ สำหรับทีมที่ดูผ่านมือถือเป็นหลัก |

---

## ภาคผนวก ข: เส้นทางสำรอง (ไม่ใช้ Office Script)

ใช้เมื่อองค์กรปิด Office Scripts หรืออยากดึงจาก **SharePoint List** แทน Excel

### ข.1 ถ้าดึงจาก Excel Table

ต้องแก้บั๊กหัวตารางในไฟล์ก่อน มิฉะนั้นข้อมูลจะผิดคอลัมน์ทั้งใบ

1. ติดตั้ง `vba/modExport_FIXED.bas` แทน `modExport` เดิม (วิธีติดตั้งอยู่ในหัวไฟล์)
2. รันมาโคร `BuildExportTable_Click`
3. Save As เป็น `.xlsx`
4. ใน Flow ใช้ **Excel Online (Business) → List rows present in a table**
   - Table: `tblPAExport`
   - **Filter Query**: `Period eq 'MA (17 Aug 26)'`
   - **Top Count**: `5000`

> รายละเอียดบั๊กและตารางเทียบคอลัมน์: [docs/03-DATA-MAPPING.md](03-DATA-MAPPING.md)

### ข.2 ถ้าดึงจาก SharePoint List

สร้าง List ที่มีคอลัมน์ตามนี้ (ชื่อภายในเป็น ASCII เพื่อให้เขียนสูตรง่าย):

| คอลัมน์ | ชนิด | มาจากคอลัมน์ในชีต |
|---|---|---|
| `Period` | Single line of text | Distribution period |
| `Zone` | Single line of text | Zone |
| `Province` | Single line of text | Province |
| `ItemCode` | Single line of text | Item Code |
| `Description` | Single line of text | Description |
| `OnhandBefore` | Number | OMC-Onhand |
| `FromHub` | Number | เบิกจาก Hub |
| `PrevPeriod` | Single line of text | Prev_Period |
| `PrevBefore` | Number | Prev_Before |
| `PrevReceived` | Number | Prev_Received |
| `Status` | Single line of text | Status |
| `RiskFlag` | Single line of text | Risk_Flag |
| `DataSet` | Choice (MA / Optimize) | Data_Set |

**Flow: Get items** → Filter Query:

```
Period eq 'MA (17 Aug 26)' and (startswith(ItemCode,'53OF') or startswith(ItemCode,'52CL'))
```

### ข.3 ประกอบการ์ดด้วยสูตร Power Automate

เส้นทางนี้ยาวกว่ามาก สรุปเป็นขั้น:

1. **Filter array** — คัดเฉพาะ 53OF/52CL
2. **Select** — แปลงแต่ละแถวเป็น object ของ ColumnSet หนึ่งแถว
   (ต้องเขียน JSON ของ ColumnSet ทั้งก้อนในช่อง Map แบบ Text mode)
3. **Compose** — ประกอบการ์ด โดยเอาผลจาก Select ไปวางในคีย์ `body`
   ด้วย `union()` หรือ `addProperty()`
4. **Post adaptive card** — `string(outputs('Compose'))`

> 💡 **ข้อแนะนำตามตรง:** วิธีนี้แก้ยากและ debug ยากกว่า Office Script มาก
> เพราะ Power Automate ไม่มีตัวช่วยตรวจ JSON ให้เลย
> ถ้าองค์กรเปิด Office Scripts ได้ ให้ใช้ทางหลัก (ตอนที่ 2–3) จะคุ้มกว่ามาก

---

## ถัดไป

- ✅ [ระบบเช็กลิสต์ก่อนใช้งานจริง](02-CHECKLIST.md) — เช็กทีละข้อก่อนส่งมอบ
- 📊 [ตารางเทียบคอลัมน์ + บั๊กที่พบในไฟล์](03-DATA-MAPPING.md)
- 🔧 [รวมปัญหาและวิธีแก้](04-TROUBLESHOOTING.md)
