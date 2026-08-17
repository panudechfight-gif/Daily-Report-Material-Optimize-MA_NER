// ***  ไฟล์นี้ถูกสร้างอัตโนมัติโดย tools/build_office_script.js  ***
// ***  ห้ามแก้ด้วยมือ — แก้ที่ adaptive-cards/01_card-main.json  ***
//     สร้างเมื่อ: (รัน node tools/build_office_script.js เพื่ออัปเดต)

/**
 * renderAdaptiveCard.ts — Office Script ที่คืน "Adaptive Card ฉบับสมบูรณ์"
 * =======================================================================
 *  *** ไฟล์นี้ถูกสร้างอัตโนมัติ — ห้ามแก้ด้วยมือ ***
 *  แก้ที่ adaptive-cards/01_card-main.json แล้วรัน:
 *      node tools/build_office_script.js
 *
 * ทำไมต้องมีสคริปต์นี้
 * -------------------
 * Power Automate action "Post adaptive card in a chat or channel" รับได้เฉพาะ
 * JSON ของการ์ดที่ "แทนค่าเรียบร้อยแล้ว" เท่านั้น — มันไม่รู้จัก ${...} / $data / $when
 * (การ expand ด้วย adaptivecards-templating เกิดที่ adaptivecards.io/designer
 *  ตอนออกแบบ ไม่ได้เกิดตอน Teams แสดงการ์ด)
 *
 * สคริปต์นี้จึงทำ 3 อย่างในแอ็กชันเดียว:
 *      1. อ่านข้อมูลจากชีต MA&Optimize NER แล้วสร้าง payload
 *      2. เอา payload ไป expand ใส่ template ที่ฝังไว้ข้างล่าง
 *      3. คืน JSON การ์ดที่พร้อมส่งเข้า Teams ทันที
 *
 * ใน Power Automate จึงเหลือแค่:
 *      Recurrence -> Run script -> Post adaptive card (body('Run_script')?['result'])
 *
 * ข้อจำกัด: Office Scripts ทำงานกับ .xlsx เท่านั้น (ไม่รองรับ .xlsm)
 */

/* eslint-disable */

// ===========================================================================
// ส่วนที่ 1 — ตรรกะเตรียมข้อมูล (สำเนาเดียวกับ buildCardPayload.ts)
// ===========================================================================
const SHEET_SOURCE = "MA&Optimize NER";
const HEADER_ROW = 7;           // หัวตารางอยู่แถวที่ 7 (แถว 1-6 เป็นปุ่ม/คำอธิบาย)
// เพดานแถวขั้นสูงสุดที่เป็นไปได้จริง = จำนวน Zone x จำนวนรหัสเป้าหมาย
// (10 รหัส x 5 Zone = 50 แถวเป็นอย่างมาก) เพดานนี้กันไม่ให้พารามิเตอร์ topRows
// จาก Power Automate ถูกตั้งเกินจริงเท่านั้น — ตัวที่บังคับให้การ์ดพอดี 28 KB
// จริง ๆ คือ CARD_SAFE_BYTES ใน _renderAdaptiveCard.skeleton.ts
// (การ์ดรวมแถวซ้ำของแต่ละ Zone แล้ว จำนวนแถวจริงจึงน้อยกว่าจำนวนแถวดิบมาก
//  เช่น MA (17 Aug 26) 91 แถวดิบ -> 29 แถวหลังรวม)
const TOP_ROWS_MAX = 60;

// ---------------------------------------------------------------------------
// รหัสวัสดุที่การ์ดต้องแสดง — กำหนดเป็นรายการตายตัว 10 รหัส
// เพิ่ม/ลดรหัส ให้แก้ที่นี่ที่เดียว (แล้วแก้ TARGET_ITEMS ใน tools/build_sample_data.py ตามกัน)
//
//   icon  = สัญลักษณ์ประจำ "กลุ่ม" — 🏷️ ป้ายชื่อ / 🧰 CLOSURE / 🔌 Drop Cable / 🧶 ARSS Fiber Cable
//           ของกลุ่มเดียวกันใช้ icon เดียวกันเสมอ ให้กวาดสายตาแยกกลุ่มได้ทันที
//   group = ชื่อกลุ่มภาษาอังกฤษ (หัวข้อในตารางอธิบายท้ายการ์ด) — 1 icon ต่อ 1 group เท่านั้น
//   short = ชื่อย่อที่ใช้แทน Item Code บนแถวข้อมูล แสดงเป็น "icon · short"
//   name  = ชื่อทางการเต็ม (ไม่ได้ใช้บนการ์ดแล้ว เก็บไว้เผื่ออ้างอิง)
//   about = คำอธิบายภาษาไทยของทั้งกลุ่ม (พิมพ์ครั้งเดียวต่อกลุ่มในตารางอธิบาย)
//
// แถวข้อมูลแสดง "icon · short" (เช่น 🧰 · Closure 12C) ส่วน Item Code เต็มไปอยู่ใน
// ตารางอธิบายท้ายการ์ดแทน (เช่น 52CL003BB=Closure 12C) ใครจะกลับไปหาในไฟล์ Excel
// ก็เปิดตารางอธิบายดูได้ ไม่ต้องพิมพ์รหัสเต็มซ้ำทุกแถว
// ---------------------------------------------------------------------------
interface TargetItem { icon: string; group: string; short: string; name: string; about: string; }

const TARGET_ITEMS: { [k: string]: TargetItem } = {
  "50MT004BB": { icon: "🏷️", group: "NAME PLATE", short: "Name Plate Aluminium",
    name: "Name Plate (Aluminium)", about: "ป้ายชื่อติดอุปกรณ์" },

  "52CL003BB": { icon: "🧰", group: "CLOSURE", short: "Closure 12C",
    name: "CLOSURE 12 C", about: "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic" },
  "52CL004BB": { icon: "🧰", group: "CLOSURE", short: "Closure 24C",
    name: "CLOSURE 24 C", about: "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic" },
  "52CL006BB": { icon: "🧰", group: "CLOSURE", short: "Closure 48C",
    name: "CLOSURE 48 C", about: "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic" },
  "52CL009BB": { icon: "🧰", group: "CLOSURE", short: "Closure 12C Inline",
    name: "CLOSURE 12 C FOR OFC DROP WIRE (IN LINE)", about: "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic" },
  "52CL010BB": { icon: "🧰", group: "CLOSURE", short: "Closure 60C",
    name: "CLOSURE 60 C", about: "หัวต่อ CLOSURE สำหรับตัดต่อ Cable Optic" },

  "53OF150BB": { icon: "🔌", group: "DROP CABLE", short: "Drop Cable 1C SC/UPC · 3m",
    name: "OPTICAL FIBER DROP CABLE 1C, FLAT TYPE (G.657A) WITH 2 SC/UPC PRE-CONNECTOR, 3m.",
    about: "สาย Pigtail Patch สำหรับ Splice เชื่อมต่อ" },

  "53OF157BB": { icon: "🧶", group: "ARSS FIBER CABLE", short: "ARSS Fiber Cable 12C",
    name: "ARSS OPTICAL FIBER CABLE 12c-FIBRE3", about: "สาย Cable Fiber Optic" },
  "53OF158BB": { icon: "🧶", group: "ARSS FIBER CABLE", short: "ARSS Fiber Cable 24C",
    name: "ARSS OPTICAL FIBER CABLE 24c-FIBRE3", about: "สาย Cable Fiber Optic" },
  "53OF160BB": { icon: "🧶", group: "ARSS FIBER CABLE", short: "ARSS Fiber Cable 60C",
    name: "ARSS OPTICAL FIBER CABLE 60c-FIBRE3", about: "สาย Cable Fiber Optic" },
};

function isTargetItem(code: string): boolean {
  const c = (code || "").trim().toUpperCase();
  return TARGET_ITEMS[c] !== undefined;
}

const ZONE_EMOJI: { [k: string]: string } = {
  "RC2-NMA": "🟦",
  "RC2-UBN": "🟩",
  "RC3-KKN": "🟨",
  "RC3-UDN": "🟧",
  "RC3-SNK": "🟪",
};

interface StatusTheme { emoji: string; color: string; short: string; }

/**
 * สีและ emoji ประจำสถานะ
 * short = "" แปลว่าไม่ต้องพิมพ์คำอธิบายซ้ำบนการ์ด (emoji สื่อความหมายพอแล้ว)
 * ตัดคำอธิบายของ "จัดสรรเกิน" และ "ไม่มีรอบก่อน" ออก เพื่อเหลือที่ให้ข้อมูลจริง
 */
function statusTheme(status: string): StatusTheme {
  const s = status || "";
  if (s.indexOf("ปกติ") >= 0) return { emoji: "✅", color: "good", short: "" };
  if (s.indexOf("จัดสรรเกิน") >= 0) return { emoji: "🟡", color: "warning", short: "" };
  if (s.indexOf("ใช้เกินยอดรับ") >= 0) return { emoji: "🔴", color: "attention", short: "ใช้เกินยอดรับ" };
  if (s.indexOf("ยอดไม่ตรง") >= 0) return { emoji: "🟠", color: "warning", short: "ยอดไม่ตรง" };
  if (s.indexOf("ไม่มีข้อมูลรอบก่อน") >= 0) return { emoji: "⚪", color: "default", short: "" };
  return { emoji: "⚪", color: "default", short: "" };
}

/** ลำดับความรุนแรงของสถานะ ใช้เลือกสถานะตัวแทนตอนรวมหลายจังหวัดเป็นแถวเดียว */
function statusRank(status: string): number {
  const s = status || "";
  if (s.indexOf("ใช้เกินยอดรับ") >= 0) return 4;
  if (s.indexOf("ยอดไม่ตรง") >= 0) return 3;
  if (s.indexOf("จัดสรรเกิน") >= 0) return 2;
  if (s.indexOf("ปกติ") >= 0) return 1;
  return 0;
}

function toNum(v: (string | number | boolean)): number {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v));
  return isNaN(n) ? 0 : n;
}

/** ใส่คอมมาคั่นหลักพันให้ข้อความตัวเลข */
function addThousandSep(text: string): string {
  const parts = text.split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return parts.join(".");
}

/** จัดรูปแบบตัวเลขให้อ่านง่าย: มีคอมมา, ตัด .0, ทศนิยมไม่เกิน 2 */
function fmtNum(v: (string | number | boolean)): string {
  const n = toNum(v);
  const isWhole = Math.abs(n - Math.round(n)) < 1e-9;
  // จำนวนเต็มแสดงไม่มีทศนิยม, มีเศษแสดง 2 ตำแหน่งเสมอ (1.2 -> "1.20")
  const text = isWhole ? String(Math.round(n)) : (Math.round(n * 100) / 100).toFixed(2);
  return addThousandSep(text);
}

/**
 * จัดรูปแบบเป็นจำนวนเต็มเสมอ — ใช้กับ "เบิกจาก Hub" ที่ต้องเป็นยอดจ่ายจริง
 * ปัดครึ่งขึ้นแบบคณิตศาสตร์ (0.5 -> 1, -0.5 -> -1) ไม่ให้ค่าติดลบเพี้ยน
 */
function fmtInt(v: (string | number | boolean)): string {
  const n = toNum(v);
  const rounded = n < 0 ? -Math.round(Math.abs(n)) : Math.round(n);
  return addThousandSep(String(rounded));
}

interface CardItem {
  itemCode: string; itemName: string; itemIcon: string; itemShort: string;
  unit: string; province: string;
  provinceCount: number; provinceTag: string;
  dataSet: string; period: string; prevPeriod: string;
  onhand: string; distribute: string; fromHub: string; numbersLine: string;
  prevBefore: string; prevReceived: string;
  status: string; statusShort: string; statusEmoji: string; statusColor: string;
  riskFlag: string; isRisk: boolean; rowWeight: string;
  zone: string; zoneEmoji: string;
}

interface CardZone {
  zone: string; zoneEmoji: string; itemCount: number; shownCount: number;
  riskCount: number; riskBadge: string; riskColor: string; countLabel: string;
  totalOnhand: string; totalDistribute: string; totalHub: string;
  totalLabel: string; zoneHeadline: string;
  items: CardItem[];
}

/**
 * ตารางอธิบายท้ายการ์ด — 1 บรรทัดต่อ "กลุ่ม" (CLOSURE / OPTICAL FIBER / NAME PLATE)
 * ไม่ใช่ 1 บรรทัดต่อรหัส เพราะแถวข้อมูลใช้ icon + ชื่อย่อแทน Item Code ไปแล้ว
 * ตารางนี้จึงทำหน้าที่ "ถอดรหัส" กลับเป็น Item Code เต็มไว้ที่เดียว
 * เช่น 🔗 CLOSURE — 52CL003BB=12C · 52CL004BB=24C · 52CL006BB=48C · …
 */
interface CardLegend { itemIcon: string; groupLabel: string; itemAbout: string; pairs: string; }

interface CardMeta {
  title: string; subtitle: string; hasLegend: boolean;
  period: string; prevPeriod: string;
  dataSet: string; dataSetLabel: string; dataSetEmoji: string;
  dataSetStyle: string; dataSetColor: string;
  generatedAt: string;
  totalItems: number; shownItems: number; totalZones: number;
  riskCount: number; watchCount: number; okCount: number;
  hasData: boolean; hasMore: boolean; moreCount: number; moreText: string;
  dashboardUrl: string; sourceFileUrl: string;
}

interface CardPayload {
  meta: CardMeta; zones: CardZone[]; items: CardItem[]; legend: CardLegend[];
}

// ---------------------------------------------------------------------------
// ตัวแปลงค่าพารามิเตอร์ที่ Power Automate ส่งเข้ามา
//
// พารามิเตอร์ของ main() ประกาศเป็นแบบไม่บังคับ (`period?: string`) เพื่อไม่ให้
// Power Automate ขึ้นดอกจันบังคับกรอก ผลคือช่องที่เว้นว่างจะได้ค่า undefined
// หรือ "" เข้ามา ตัวช่วย 2 ตัวนี้จึงแปลงให้เป็นค่าที่ตรรกะข้างล่างใช้ได้เสมอ
//
// หมายเหตุ: ประกาศ `?` พร้อม default value ในตัวเดียวกันไม่ได้ใน TypeScript
//          การเติมค่าตั้งต้นจึงย้ายมาทำที่นี่แทน
// ---------------------------------------------------------------------------

/** ข้อความจากช่องที่อาจถูกเว้นว่าง -> คืนสตริงเสมอ (ไม่มี undefined) */
function resolveText(v: (string | undefined)): string {
  return v === undefined || v === null ? "" : String(v).trim();
}

/** จำนวนแถวจากช่องที่อาจถูกเว้นว่างหรือกรอกมาผิดรูปแบบ -> คืนตัวเลขที่ใช้ได้เสมอ */
function resolveTopRows(v: (number | string | undefined)): number {
  if (v === undefined || v === null || v === "") return TOP_ROWS_MAX;
  const n = typeof v === "number" ? v : parseFloat(String(v));
  if (isNaN(n) || n < 1) return TOP_ROWS_MAX;
  return n > TOP_ROWS_MAX ? TOP_ROWS_MAX : Math.floor(n);
}

// ---------------------------------------------------------------------------
// ตรรกะหลัก แยกออกจาก main() เพื่อให้ renderAdaptiveCard.ts นำไปใช้ซ้ำได้
// (tools/build_office_script.js คัดลอกทุกอย่างเหนือ main() ไปประกอบร่าง)
// ---------------------------------------------------------------------------
function buildPayload(
  workbook: ExcelScript.Workbook,
  period: string = "",
  topRows: number = 12,
  dashboardUrl: string = "",
  sourceFileUrl: string = ""
): CardPayload {

  if (topRows > TOP_ROWS_MAX) topRows = TOP_ROWS_MAX;
  if (topRows < 1) topRows = TOP_ROWS_MAX;

  const sheet = workbook.getWorksheet(SHEET_SOURCE);
  if (!sheet) throw new Error("ไม่พบชีต " + SHEET_SOURCE);

  const used = sheet.getUsedRange();
  if (!used) throw new Error("ชีต " + SHEET_SOURCE + " ว่างเปล่า");

  const values = used.getValues();
  const firstRow = used.getRowIndex();          // 0-based
  const hdrIdx = HEADER_ROW - 1 - firstRow;
  if (hdrIdx < 0 || hdrIdx >= values.length) throw new Error("หาหัวตารางแถวที่ 7 ไม่เจอ");

  // แผนที่ ชื่อหัวตาราง -> ตำแหน่งคอลัมน์ (อ่านจากหัวจริง ไม่ใช้ตำแหน่งตายตัว)
  const col: { [k: string]: number } = {};
  const hdr = values[hdrIdx];
  for (let c = 0; c < hdr.length; c++) {
    const name = String(hdr[c] === null || hdr[c] === undefined ? "" : hdr[c]).trim();
    if (name !== "" && col[name] === undefined) col[name] = c;
  }

  // -------------------------------------------------------------------------
  // หัวตารางบางคอลัมน์เคยถูกเปลี่ยนชื่อมาแล้วในไฟล์ต้นทาง
  //     OMC-Onhand   -> Onhand
  //     จัดสรร        -> จัดสรร(กระจาย)
  //     เบิกจาก Hub   -> จัดสรร(เบิก Hub)
  //
  // จึงไม่ยึดชื่อเดียวตายตัว แต่ไล่หาจากรายชื่อที่ยอมรับได้ แล้วผูกกลับเข้า
  // ชื่อมาตรฐานที่ตรรกะข้างล่างใช้ ถ้าวันหนึ่งชีตเปลี่ยนชื่ออีก ให้เติมชื่อใหม่
  // ไว้ "ข้างหน้า" ในอาร์เรย์ของคอลัมน์นั้น — ของเดิมยังอ่านได้เหมือนเดิม
  // -------------------------------------------------------------------------
  const ALIASES: { [k: string]: string[] } = {
    "Onhand": ["Onhand", "OMC-Onhand"],
    "จัดสรร(กระจาย)": ["จัดสรร(กระจาย)", "จัดสรร"],
    "จัดสรร(เบิก Hub)": ["จัดสรร(เบิก Hub)", "เบิกจาก Hub"],
    "จัดสรรเพิ่ม": ["จัดสรรเพิ่ม"],
    "รวม 3 รายการ": ["รวม 3 รายการ"],
  };
  for (const std of Object.keys(ALIASES)) {
    for (const alias of ALIASES[std]) {
      if (col[alias] !== undefined) { col[std] = col[alias]; break; }
    }
  }

  const need = ["Distribution period", "Zone", "Item Code", "Onhand",
                "จัดสรร(เบิก Hub)", "Prev_Period", "Prev_Before", "Prev_Received",
                "Status", "Risk_Flag", "Data_Set"];
  const missing = need.filter(n => col[n] === undefined);
  if (missing.length > 0) throw new Error("ไม่พบคอลัมน์: " + missing.join(", "));

  // -------------------------------------------------------------------------
  // แก้แนวคอลัมน์เลื่อนในชีตต้นทาง
  //
  // ชีต MA&Optimize NER มีข้อมูลของบล็อก จัดสรร(กระจาย) / จัดสรร(เบิก Hub) /
  // จัดสรรเพิ่ม / รวม 3 รายการ วางเลื่อนไปทางขวาของหัวตารางตัวเอง 1 ช่อง
  // ถ้าอ่านตามชื่อหัวตารางตรง ๆ ค่าที่ได้ในช่อง "จัดสรร(เบิก Hub)"
  // จะเป็นค่าของ "จัดสรร(กระจาย)" แทน
  //
  // ตรวจด้วยกฎที่ชีตคำนวณไว้เอง:
  //     จัดสรร(กระจาย) + จัดสรร(เบิก Hub) + จัดสรรเพิ่ม = รวม 3 รายการ
  // ลองทั้งแบบไม่เลื่อนและเลื่อน 1 ช่อง แล้วเลือกแบบที่กฎนี้เป็นจริงทุกแถวที่สุ่มมา
  // ถ้าวันหนึ่งชีตต้นทางถูกแก้ให้ตรงแล้ว ตัวตรวจนี้จะเลือกแบบไม่เลื่อนเองอัตโนมัติ
  //
  // (ยืนยันกับไฟล์ V2 ล่าสุดแล้ว: off=1 ทำให้กฎเป็นจริง 2,365/2,365 แถว
  //  ส่วน off=0 ตรงแค่ 847 แถว — บั๊กนี้ยังอยู่แม้หัวตารางจะถูกเปลี่ยนชื่อแล้ว)
  // -------------------------------------------------------------------------
  const shiftNames = ["จัดสรร(กระจาย)", "จัดสรร(เบิก Hub)", "จัดสรรเพิ่ม", "รวม 3 รายการ"];
  const haveAll = shiftNames.every(n => col[n] !== undefined);
  if (haveAll) {
    let bestOffset = 0;
    for (let off = 0; off <= 1; off++) {
      let checked = 0, agree = 0;
      for (let r = hdrIdx + 1; r < values.length && checked < 300; r++) {
        const row = values[r];
        if (row === undefined) continue;
        const ia = col["จัดสรร(กระจาย)"] + off, ib = col["จัดสรร(เบิก Hub)"] + off;
        const ic = col["จัดสรรเพิ่ม"] + off, is = col["รวม 3 รายการ"] + off;
        if (is >= row.length) continue;
        checked++;
        const total = toNum(row[ia]) + toNum(row[ib]) + toNum(row[ic]);
        if (Math.abs(total - toNum(row[is])) < 1e-9) agree++;
      }
      if (checked > 0 && agree === checked) { bestOffset = off; break; }
    }
    if (bestOffset > 0) {
      for (const n of shiftNames) col[n] = col[n] + bestOffset;
    }
  }

  const get = (row: (string | number | boolean)[], name: string): string => {
    const i = col[name];
    if (i === undefined) return "";
    const v = row[i];
    return v === null || v === undefined ? "" : String(v).trim();
  };

  // ---- 1. คัดเฉพาะรหัส 53OF / 52CL ----
  const picked: (string | number | boolean)[][] = [];
  for (let r = hdrIdx + 1; r < values.length; r++) {
    if (isTargetItem(get(values[r], "Item Code"))) picked.push(values[r]);
  }

  // ---- 2. เลือกรอบ: ถ้าไม่ระบุ ใช้รอบ Optimize ที่ Period_Seq สูงสุด ----
  // Power Automate บังคับให้พารามิเตอร์ต้องมีค่า ส่งค่าว่างจริง ๆ ไม่ได้
  // จึงรับค่าตัวแทน "ไม่ระบุ" ได้หลายแบบ แล้วแปลงเป็นค่าว่างให้เอง
  //
  // ค่าตั้งต้นคือชุด Optimize (รอบเดือน) ตามที่กำหนดให้เป็นรายงานหลัก
  // ถ้าอยากได้ชุด MA (รอบสัปดาห์) ให้ส่ง period = "MA" หรือชื่อรอบ MA ตรง ๆ
  let targetPeriod = (period || "").trim();
  const blank = targetPeriod.toLowerCase();
  if (blank === "-" || blank === "auto" || blank === "latest" || blank === "null") {
    targetPeriod = "";
  }
  // ส่งชื่อชุดข้อมูลมาแทนชื่อรอบได้ = "รอบล่าสุดของชุดนั้น"
  let wantSet = "Optimize";
  if (blank === "ma") { wantSet = "MA"; targetPeriod = ""; }
  else if (blank === "optimize") { wantSet = "Optimize"; targetPeriod = ""; }

  if (targetPeriod === "") {
    let bestSeq = -1;
    for (const row of picked) {
      if (get(row, "Data_Set") !== wantSet) continue;
      const seq = toNum(col["Period_Seq"] !== undefined ? row[col["Period_Seq"]] : 0);
      if (seq > bestSeq) { bestSeq = seq; targetPeriod = get(row, "Distribution period"); }
    }
    // ถ้าไม่มีชุดที่ขอเลย ถอยไปใช้รอบล่าสุดเท่าที่มี เพื่อไม่ให้การ์ดว่างโดยไม่จำเป็น
    if (targetPeriod === "") {
      for (const row of picked) {
        const seq = toNum(col["Period_Seq"] !== undefined ? row[col["Period_Seq"]] : 0);
        if (seq > bestSeq) { bestSeq = seq; targetPeriod = get(row, "Distribution period"); }
      }
    }
  }
  const inPeriod = picked.filter(row => get(row, "Distribution period") === targetPeriod);

  // ---- 3. รวมแถวซ้ำ: 1 Zone + 1 Item Code = 1 แถว ----
  // ชีตต้นทางแตกแถวตามจังหวัด รหัสเดียวกันจึงโผล่หลายครั้งในโซนเดียว
  // การ์ดไม่ได้แสดงจังหวัด การเห็นรหัสซ้ำจึงอ่านไม่รู้เรื่องและกินที่เปล่า ๆ
  // จึงบวกยอดของทุกจังหวัดเข้าด้วยกัน แล้วเก็บจำนวนจังหวัดไว้บอกที่มา
  const numOf = (row: (string | number | boolean)[], name: string): number => {
    const i = col[name];
    return i === undefined ? 0 : toNum(row[i]);
  };

  interface Agg {
    zone: string; itemCode: string; itemName: string; itemIcon: string; itemShort: string;
    unit: string;
    onhand: number; distribute: number; fromHub: number;
    prevBefore: number; prevReceived: number;
    provinces: { [k: string]: boolean }; provinceCount: number;
    status: string; statusRank: number; isRisk: boolean; riskFlag: string;
  }

  const aggMap: { [k: string]: Agg } = {};
  const aggOrder: string[] = [];
  let dataSet = "Optimize";
  let prevPeriod = "";

  for (const row of inPeriod) {
    const ds = get(row, "Data_Set") || "Optimize";
    dataSet = ds;
    const pp = get(row, "Prev_Period");
    if (pp !== "" && prevPeriod === "") prevPeriod = pp;

    const zone = get(row, "Zone") || "-";
    const code = get(row, "Item Code").toUpperCase();
    const key = zone + "|" + code;

    if (aggMap[key] === undefined) {
      const spec = TARGET_ITEMS[code];
      aggMap[key] = {
        zone: zone, itemCode: code,
        itemName: spec !== undefined ? spec.name
          : (get(row, "Description") || get(row, "Art No")),
        itemIcon: spec !== undefined ? spec.icon : "▫️",
        itemShort: spec !== undefined ? spec.short : code,
        unit: get(row, "Unit"),
        onhand: 0, distribute: 0, fromHub: 0, prevBefore: 0, prevReceived: 0,
        provinces: {}, provinceCount: 0,
        status: "", statusRank: -1, isRisk: false, riskFlag: "",
      };
      aggOrder.push(key);
    }
    const a = aggMap[key];

    a.onhand += numOf(row, "Onhand");
    // ชุด Optimize ลงยอดไว้ที่ "จัดสรร(กระจาย)" ส่วนชุด MA ลงที่ "จัดสรร(เบิก Hub)"
    // การ์ดจึงต้องแสดงทั้งคู่ ไม่งั้นรอบ Optimize จะขึ้นศูนย์ทั้งใบ
    a.distribute += numOf(row, "จัดสรร(กระจาย)");
    a.fromHub += numOf(row, "จัดสรร(เบิก Hub)");
    a.prevBefore += numOf(row, "Prev_Before");
    a.prevReceived += numOf(row, "Prev_Received");

    const prov = get(row, "Province");
    if (prov !== "" && a.provinces[prov] === undefined) {
      a.provinces[prov] = true;
      a.provinceCount++;
    }

    // สถานะตัวแทน = อันที่รุนแรงที่สุดในกลุ่ม
    const stText = get(row, "Status");
    const rank = statusRank(stText);
    if (rank > a.statusRank) { a.statusRank = rank; a.status = stText; }

    const risk = get(row, "Risk_Flag");
    if (risk.indexOf("เสี่ยง") >= 0) { a.isRisk = true; a.riskFlag = risk; }
    else if (risk !== "" && a.riskFlag === "") { a.riskFlag = risk; }
  }

  const all: CardItem[] = [];
  let riskCount = 0, watchCount = 0, okCount = 0;

  for (const key of aggOrder) {
    const a = aggMap[key];
    const st = statusTheme(a.status);
    if (a.isRisk) riskCount++; else if (a.riskFlag !== "") watchCount++; else okCount++;

    all.push({
      itemCode: a.itemCode,
      itemName: a.itemName.length > 45 ? a.itemName.substring(0, 45) : a.itemName,
      itemIcon: a.itemIcon,
      itemShort: a.itemShort,
      unit: a.unit,
      province: a.provinceCount > 0 ? String(a.provinceCount) + " จังหวัด" : "-",
      provinceCount: a.provinceCount,
      // ต่อท้ายรหัสเฉพาะตอนที่ยอดนี้รวมมาจากหลายจังหวัด จะได้รู้ว่าตัวเลขมาจากไหน
      provinceTag: a.provinceCount > 1 ? " ·" + String(a.provinceCount) + "จว." : "",
      dataSet: dataSet,
      period: targetPeriod,
      prevPeriod: prevPeriod || "-",
      onhand: fmtNum(a.onhand),
      distribute: fmtInt(a.distribute),
      fromHub: fmtInt(a.fromHub),
      // 3 ตัวเลขรวมเป็นบรรทัดเดียว — ประหยัดโครงสร้าง JSON ของการ์ดได้มาก
      // (ตัด TextBlock ต่อคอลัมน์ที่ซ้ำกันทุกแถวออก 2 ก้อน) จำเป็นมากขึ้นเมื่อ
      // รายการวัสดุมี 10 รหัส x 5 Zone แล้ว ไม่งั้นการ์ด Optimize ชนเพดาน 28 KB
      numbersLine: "Onhand " + fmtNum(a.onhand)
        + " · กระจาย " + fmtInt(a.distribute)
        + " · Hub " + fmtInt(a.fromHub),
      prevBefore: fmtNum(a.prevBefore),
      prevReceived: fmtNum(a.prevReceived),
      status: (a.status || "-").replace(/^-\s*/, "") || "-",
      statusShort: st.short,
      statusEmoji: st.emoji,
      statusColor: st.color,
      riskFlag: a.riskFlag || "-",
      isRisk: a.isRisk,
      rowWeight: a.isRisk ? "bolder" : "default",
      zone: a.zone,
      zoneEmoji: ZONE_EMOJI[a.zone] || "🔷",
    });
  }

  // ---- 4. เรียงความสำคัญ: เสี่ยงขาดก่อน -> คงเหลือน้อยก่อน -> รหัส ----
  const cmpItem = (a: CardItem, b: CardItem): number => {
    if (a.isRisk !== b.isRisk) return a.isRisk ? -1 : 1;
    const na = toNum(a.onhand.replace(/,/g, ""));
    const nb = toNum(b.onhand.replace(/,/g, ""));
    if (na !== nb) return na - nb;
    return a.itemCode < b.itemCode ? -1 : (a.itemCode > b.itemCode ? 1 : 0);
  };
  all.sort((a, b) => {
    const c = cmpItem(a, b);
    if (c !== 0) return c;
    return a.zone < b.zone ? -1 : (a.zone > b.zone ? 1 : 0);
  });

  // ---- 5. จัดกลุ่มตาม Zone (นับยอดเต็มของแต่ละ Zone ไว้แสดงในหัวกลุ่ม) ----
  const zoneTotal: { [k: string]: number } = {};
  const zoneRisk: { [k: string]: number } = {};
  for (const it of all) {
    zoneTotal[it.zone] = (zoneTotal[it.zone] || 0) + 1;
    if (it.isRisk) zoneRisk[it.zone] = (zoneRisk[it.zone] || 0) + 1;
  }

  // ---- 6. เลือกแถววนรอบทีละ Zone ----
  // เลือกแบบ "เอาอันดับ 1 ของทุก Zone ก่อน แล้วค่อยวนอันดับ 2"
  // ทำให้ทุก Zone มีที่บนการ์ดเสมอ แม้รายการเสี่ยงจะกระจุกอยู่ Zone เดียว
  const byZone: { [k: string]: CardItem[] } = {};
  const zoneNames: string[] = [];
  for (const it of all) {
    if (byZone[it.zone] === undefined) { byZone[it.zone] = []; zoneNames.push(it.zone); }
    byZone[it.zone].push(it);
  }
  zoneNames.sort();
  for (const z of zoneNames) byZone[z].sort((a, b) => cmpItem(a, b));

  const shown: CardItem[] = [];
  let rank = 0;
  while (shown.length < topRows) {
    let addedThisRound = false;
    for (const z of zoneNames) {
      if (shown.length >= topRows) break;
      const list = byZone[z];
      if (rank < list.length) { shown.push(list[rank]); addedThisRound = true; }
    }
    if (!addedThisRound) break;      // ทุก Zone หมดรายการแล้ว
    rank++;
  }
  const moreCount = all.length - shown.length;

  const grouped: { [k: string]: CardItem[] } = {};
  for (const it of shown) {
    if (!grouped[it.zone]) grouped[it.zone] = [];
    grouped[it.zone].push(it);
  }

  // ยอดรวมของแต่ละ Zone = บวกทุกจังหวัดในโซนนั้น (ไม่ใช่แค่แถวที่แสดง)
  // ตรงตามโจทย์: Onhand และ จัดสรร(เบิก Hub) รวมต่อ Zone ของรอบที่เลือกอยู่
  const zoneOnhand: { [k: string]: number } = {};
  const zoneDist: { [k: string]: number } = {};
  const zoneHub: { [k: string]: number } = {};
  for (const it of all) {
    zoneOnhand[it.zone] = (zoneOnhand[it.zone] || 0) + toNum(it.onhand.replace(/,/g, ""));
    zoneDist[it.zone] = (zoneDist[it.zone] || 0) + toNum(it.distribute.replace(/,/g, ""));
    zoneHub[it.zone] = (zoneHub[it.zone] || 0) + toNum(it.fromHub.replace(/,/g, ""));
  }

  const zones: CardZone[] = Object.keys(grouped).sort().map(z => {
    const items = grouped[z];
    const nRisk = zoneRisk[z] || 0;
    const nAll = zoneTotal[z] || items.length;
    return {
      zone: z,
      zoneEmoji: ZONE_EMOJI[z] || "🔷",
      itemCount: nAll,
      shownCount: items.length,
      riskCount: nRisk,
      // ไม่มีรายการเสี่ยง = ไม่ต้องขึ้นป้ายอะไรเลย ปล่อยว่างให้การ์ดโล่ง
      riskBadge: nRisk > 0 ? "เสี่ยงขาด " + nRisk : "",
      riskColor: nRisk > 0 ? "attention" : "good",
      countLabel: items.length === nAll
        ? nAll + " รายการ"
        : "แสดง " + items.length + " จาก " + nAll + " รายการ",
      // ยอดรวมทั้งโซน แสดงคู่กับหัวกลุ่ม
      totalOnhand: fmtNum(zoneOnhand[z] || 0),
      totalDistribute: fmtInt(zoneDist[z] || 0),
      totalHub: fmtInt(zoneHub[z] || 0),
      totalLabel: "Onhand " + fmtNum(zoneOnhand[z] || 0)
        + " · กระจาย " + fmtInt(zoneDist[z] || 0)
        + " · เบิก Hub " + fmtInt(zoneHub[z] || 0),
      // หัวกลุ่มรวมเป็นบรรทัดเดียว — ประหยัดพื้นที่การ์ดได้ราว 0.5 KB ต่อ Zone
      //
      // ไม่ใส่ emoji ประจำ Zone ตรงนี้แล้ว ตามนโยบาย "emoji มีเฉพาะหัวการ์ด"
      // การแยก Zone อาศัยตัวหนา + สีของบรรทัด (riskColor) และเส้นคั่นแทน
      // ฟิลด์ zoneEmoji ยังคงอยู่ใน payload เพื่อให้การ์ดใบอื่นเรียกใช้ได้เหมือนเดิม
      zoneHeadline: z
        + " · " + (items.length === nAll ? nAll + " รายการ" : "แสดง " + items.length + "/" + nAll)
        + " · Onhand " + fmtNum(zoneOnhand[z] || 0)
        + " · กระจาย " + fmtInt(zoneDist[z] || 0)
        + " · เบิก Hub " + fmtInt(zoneHub[z] || 0)
        + (nRisk > 0 ? " · เสี่ยงขาด " + nRisk : ""),
      items: items,
    };
  });

  // ตารางอธิบายท้ายการ์ด — ถอดรหัส icon + ชื่อย่อบนแถวข้อมูลกลับเป็น Item Code เต็ม
  //
  // 1 บรรทัดต่อ "กลุ่ม" (CLOSURE / OPTICAL FIBER / NAME PLATE) ไม่ใช่ต่อรหัส
  // เพราะแถวข้อมูลเปลี่ยนไปแสดง icon + ชื่อย่อ (เช่น 🔗 12C) แทน Item Code แล้ว
  // ใครต้องการรหัสเต็มไปค้นในไฟล์ Excel ก็ไล่ดูที่นี่ที่เดียว
  const legendCodes: string[] = [];
  for (const it of all) {
    if (legendCodes.indexOf(it.itemCode) < 0) legendCodes.push(it.itemCode);
  }
  legendCodes.sort();

  const legend: CardLegend[] = [];
  const legendSeen: { [k: string]: number } = {};   // group -> ตำแหน่งใน legend
  for (const c of legendCodes) {
    const spec = TARGET_ITEMS[c];
    const icon = spec !== undefined ? spec.icon : "▫️";
    const group = spec !== undefined ? spec.group : "อื่น ๆ";
    const about = spec !== undefined ? spec.about : "";
    const short = spec !== undefined ? spec.short : c;
    const pair = c + "=" + short;
    if (legendSeen[group] === undefined) {
      legendSeen[group] = legend.length;
      legend.push({ itemIcon: icon, groupLabel: group, itemAbout: about, pairs: pair });
    } else {
      const at = legendSeen[group];
      legend[at].pairs = legend[at].pairs + " · " + pair;
    }
  }

  const isMA = dataSet === "MA";
  const meta: CardMeta = {
    title: "Report Stock Allocation Cable&Material NER",
    subtitle: "",                          // ตัดบรรทัดคำอธิบายรหัสวัสดุออกแล้ว
    hasLegend: legend.length > 0,
    period: targetPeriod || "-",
    prevPeriod: prevPeriod || "-",
    dataSet: dataSet,
    dataSetLabel: isMA ? "MA — Weekly Allocation" : "Optimize — Allocation Plan",
    // 🏬 = คลังวัสดุ (Store) ใช้ทั้งสองชุดข้อมูล ให้การ์ดสื่อว่าเป็นรายงานคลัง
    // ความต่างของ MA / Optimize อ่านได้จากป้ายใต้ชื่อและสีของทั้งการ์ดอยู่แล้ว
    dataSetEmoji: "🏬",
    dataSetStyle: isMA ? "accent" : "good",
    dataSetColor: isMA ? "accent" : "good",
    generatedAt: "",                       // Power Automate เติมด้วย convertTimeZone()
    totalItems: all.length,
    shownItems: shown.length,
    totalZones: Object.keys(zoneTotal).length,
    riskCount: riskCount,
    watchCount: watchCount,
    okCount: okCount,
    hasData: all.length > 0,
    // เป้าหมายคือให้ทุกรหัสลงการ์ดครบเสมอ บรรทัด "…และอีก N รายการ" จึงถูกตัดออก
    // จากหน้าการ์ดแล้ว แต่ยังคำนวณค่าไว้ เผื่อ Flow อื่นหรือการ์ดแบบอื่นเรียกใช้
    hasMore: moreCount > 0,
    moreCount: moreCount,
    moreText: "",
    dashboardUrl: dashboardUrl || "https://app.powerbi.com/CHANGE-ME",
    sourceFileUrl: sourceFileUrl || "https://mimotech-my.sharepoint.com/:x:/r/personal/panudeck_ais_co_th/Documents/Workflow%20Power%20Automate/JobMonitor_CMTRS_Job_Tracking.xlsx?d=we08d992568d54ab88ef622f150e1cd05&csf=1&web=1&e=8sLq79",
  };

  // คืนทั้ง zones (การ์ดหลัก) และ items (การ์ดแบบง่าย) เผื่อเลือกใช้ได้ทั้งสองแบบ
  return { meta: meta, zones: zones, items: shown, legend: legend };
}

// ===========================================================================
// ส่วนที่ 2 — template ของการ์ด (ฝังมาจาก adaptive-cards/01_card-main.json)
// ===========================================================================
const CARD_TEMPLATE: object = {
  "type": "AdaptiveCard",
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "version": "1.5",
  "msteams": {
    "width": "Full"
  },
  "fallbackText": "Report Stock Allocation Cable&Material NER (อุปกรณ์ของคุณแสดงการ์ดนี้ไม่ได้ กรุณาเปิดจาก Dashboard)",
  "speak": "Report Stock Allocation ${meta.dataSet} รอบ ${meta.period} มี ${meta.totalItems} รายการ เสี่ยงขาด ${meta.riskCount} รายการ",
  "body": [
    {
      "type": "Container",
      "style": "${meta.dataSetStyle}",
      "bleed": true,
      "items": [
        {
          "type": "ColumnSet",
          "columns": [
            {
              "type": "Column",
              "width": "auto",
              "verticalContentAlignment": "center",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "${meta.dataSetEmoji}",
                  "size": "large",
                  "wrap": false,
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "width": "stretch",
              "verticalContentAlignment": "center",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "${meta.title}",
                  "weight": "bolder",
                  "size": "medium",
                  "wrap": true,
                  "spacing": "none"
                },
                {
                  "type": "TextBlock",
                  "text": "${meta.dataSetLabel}",
                  "size": "small",
                  "isSubtle": true,
                  "wrap": true,
                  "spacing": "none"
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "ColumnSet",
      "spacing": "small",
      "columns": [
        {
          "type": "Column",
          "width": "stretch",
          "items": [
            {
              "type": "TextBlock",
              "text": "รอบจัดสรร",
              "size": "small",
              "isSubtle": true,
              "wrap": true,
              "spacing": "none"
            },
            {
              "type": "TextBlock",
              "text": "**${meta.period}**",
              "size": "small",
              "color": "accent",
              "wrap": true,
              "spacing": "none"
            }
          ]
        },
        {
          "type": "Column",
          "width": "stretch",
          "items": [
            {
              "type": "TextBlock",
              "text": "รอบก่อนหน้า",
              "size": "small",
              "isSubtle": true,
              "wrap": true,
              "spacing": "none"
            },
            {
              "type": "TextBlock",
              "text": "**${meta.prevPeriod}**",
              "size": "small",
              "isSubtle": true,
              "wrap": true,
              "spacing": "none"
            }
          ]
        }
      ]
    },
    {
      "type": "Container",
      "$when": "${meta.hasData == false}",
      "spacing": "medium",
      "style": "emphasis",
      "items": [
        {
          "type": "TextBlock",
          "text": "ไม่มีรายการวัสดุในรอบนี้",
          "size": "large",
          "weight": "bolder",
          "horizontalAlignment": "center",
          "wrap": true,
          "spacing": "small"
        },
        {
          "type": "TextBlock",
          "text": "ไม่พบรหัสวัสดุที่กำหนดไว้ในรอบ **${meta.period}**",
          "horizontalAlignment": "center",
          "isSubtle": true,
          "wrap": true,
          "spacing": "small"
        },
        {
          "type": "TextBlock",
          "text": "สิ่งที่ควรตรวจสอบ:\n\n1. กดปุ่ม **Refresh Data New Sheet** ในไฟล์ Excel แล้วรัน **BuildExportTable** ใหม่\n2. ตรวจว่าเลือก Sheet ต้นทางถูกรอบหรือยัง (เซลล์ C2 / C4 ของชีต MA&Optimize NER)\n3. ตรวจว่าไฟล์ถูกบันทึกกลับขึ้น SharePoint แล้ว",
          "wrap": true,
          "spacing": "medium"
        }
      ]
    },
    {
      "type": "Container",
      "$when": "${meta.hasData}",
      "spacing": "small",
      "items": [
        {
          "type": "ColumnSet",
          "style": "emphasis",
          "bleed": true,
          "spacing": "small",
          "columns": [
            {
              "type": "Column",
              "width": "stretch",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "เสี่ยงขาด",
                  "size": "small",
                  "horizontalAlignment": "center",
                  "wrap": true,
                  "spacing": "none"
                },
                {
                  "type": "TextBlock",
                  "text": "${string(meta.riskCount)}",
                  "size": "large",
                  "weight": "bolder",
                  "color": "attention",
                  "horizontalAlignment": "center",
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "width": "stretch",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "พอใช้",
                  "size": "small",
                  "horizontalAlignment": "center",
                  "wrap": true,
                  "spacing": "none"
                },
                {
                  "type": "TextBlock",
                  "text": "${string(meta.watchCount)}",
                  "size": "large",
                  "weight": "bolder",
                  "color": "warning",
                  "horizontalAlignment": "center",
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "width": "stretch",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "ปกติ",
                  "size": "small",
                  "horizontalAlignment": "center",
                  "wrap": true,
                  "spacing": "none"
                },
                {
                  "type": "TextBlock",
                  "text": "${string(meta.okCount)}",
                  "size": "large",
                  "weight": "bolder",
                  "color": "good",
                  "horizontalAlignment": "center",
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "width": "stretch",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "รวม",
                  "size": "small",
                  "horizontalAlignment": "center",
                  "wrap": true,
                  "spacing": "none"
                },
                {
                  "type": "TextBlock",
                  "text": "${string(meta.totalItems)}",
                  "size": "large",
                  "weight": "bolder",
                  "horizontalAlignment": "center",
                  "spacing": "none"
                }
              ]
            }
          ]
        },
        {
          "type": "ColumnSet",
          "spacing": "medium",
          "separator": true,
          "columns": [
            {
              "type": "Column",
              "width": 42,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Item",
                  "weight": "bolder",
                  "size": "small",
                  "color": "accent",
                  "wrap": true,
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "width": 58,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Onhand · กระจาย · เบิก Hub",
                  "weight": "bolder",
                  "size": "small",
                  "color": "accent",
                  "horizontalAlignment": "right",
                  "wrap": true,
                  "spacing": "none"
                }
              ]
            }
          ]
        },
        {
          "type": "Container",
          "$data": "${zones}",
          "spacing": "none",
          "items": [
            {
              "type": "TextBlock",
              "text": "${zoneHeadline}",
              "size": "small",
              "weight": "bolder",
              "color": "${riskColor}",
              "wrap": true,
              "spacing": "small",
              "separator": true
            },
            {
              "type": "ColumnSet",
              "spacing": "none",
              "columns": [
                {
                  "type": "Column",
                  "width": 42,
                  "items": [
                    {
                      "type": "TextBlock",
                      "$data": "${items}",
                      "text": "${itemShort}${provinceTag}",
                      "size": "small",
                      "weight": "${rowWeight}",
                      "color": "${statusColor}",
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 58,
                  "items": [
                    {
                      "type": "TextBlock",
                      "$data": "${items}",
                      "text": "${numbersLine}",
                      "size": "small",
                      "horizontalAlignment": "right",
                      "spacing": "none"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "Container",
      "$when": "${meta.hasLegend}",
      "spacing": "small",
      "separator": true,
      "items": [
        {
          "type": "TextBlock",
          "text": "ชื่อย่อบนการ์ด · Item Code เต็ม",
          "size": "small",
          "weight": "bolder",
          "color": "accent",
          "wrap": true,
          "spacing": "none"
        },
        {
          "type": "Container",
          "$data": "${legend}",
          "spacing": "small",
          "items": [
            {
              "type": "TextBlock",
              "text": "**${groupLabel}** — ${itemAbout}",
              "size": "small",
              "wrap": true,
              "spacing": "none"
            },
            {
              "type": "TextBlock",
              "text": "${pairs}",
              "size": "small",
              "isSubtle": true,
              "wrap": true,
              "spacing": "none"
            }
          ]
        }
      ]
    },
    {
      "type": "TextBlock",
      "text": "สร้างรายงานเมื่อ ${meta.generatedAt} • ที่มา: ${meta.dataSet} / ${meta.period}",
      "size": "small",
      "isSubtle": true,
      "wrap": true,
      "spacing": "small",
      "separator": true
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Dashboard",
      "url": "${meta.dashboardUrl}"
    },
    {
      "type": "Action.OpenUrl",
      "title": "Open the source file. (Excel / SharePoint)",
      "url": "${meta.sourceFileUrl}"
    }
  ]
};

// ===========================================================================
// ส่วนที่ 3 — ตัว expand template ขนาดเล็ก
//
// รองรับเฉพาะไวยากรณ์ที่การ์ดในโปรเจกต์นี้ใช้จริง:
//     ${path.to.value}       แทนค่าลงในข้อความ
//     ${string(x)}           แปลงตัวเลขเป็นข้อความ
//     ${x == false}          เงื่อนไขเทียบค่า
//     ${$root.meta.x}        อ้างถึงข้อมูลชั้นบนสุดจากใน $data
//     "$data": "${array}"    ทำซ้ำ element ตามจำนวนสมาชิกในอาร์เรย์
//     "$when": "${cond}"     ตัด element ทิ้งเมื่อเงื่อนไขเป็นเท็จ
//
// ผลลัพธ์ถูกตรวจแล้วว่าตรงกับไลบรารี adaptivecards-templating ทุก byte
// (ดู tools/test_office_script.js)
// ===========================================================================

type Scope = { data: object; root: object };

function acResolvePath(path: string, scope: Scope): (string | number | boolean | object) {
  const trimmed = path.trim();
  let cur: object;
  let rest: string;

  if (trimmed.indexOf("$root.") === 0) {
    cur = scope.root;
    rest = trimmed.substring(6);
  } else if (trimmed === "$root") {
    return scope.root;
  } else {
    cur = scope.data;
    rest = trimmed;
  }

  const parts = rest.split(".");
  let value: (string | number | boolean | object) = cur;
  for (const p of parts) {
    if (p === "") continue;
    if (value === null || value === undefined || typeof value !== "object") return "";
    value = (value as { [k: string]: (string | number | boolean | object) })[p];
  }
  return value === null || value === undefined ? "" : value;
}

/** ประเมินนิพจน์ภายใน ${...} */
function acEval(expr: string, scope: Scope): (string | number | boolean | object) {
  let e = expr.trim();

  // string(...)
  if (e.indexOf("string(") === 0 && e.charAt(e.length - 1) === ")") {
    const inner = e.substring(7, e.length - 1);
    const v = acResolvePath(inner, scope);
    return v === null || v === undefined ? "" : String(v);
  }

  // เปรียบเทียบ  x == false / x == true / x != false
  const cmp = e.match(/^(.+?)\s*(==|!=)\s*(true|false)$/);
  if (cmp) {
    const left = acResolvePath(cmp[1], scope);
    const want = cmp[3] === "true";
    const truthy = left === true || (left !== false && left !== "" && left !== 0 &&
                                     left !== null && left !== undefined);
    return cmp[2] === "==" ? truthy === want : truthy !== want;
  }

  return acResolvePath(e, scope);
}

/** แทนค่า ${...} ทุกจุดในข้อความหนึ่งบรรทัด */
function acExpandString(text: string, scope: Scope): (string | boolean | number) {
  // ทั้งสตริงเป็น ${...} ตัวเดียว -> คืนค่าดิบ (เพื่อให้ $when ได้ boolean จริง)
  const whole = text.match(/^\$\{([^}]*)\}$/);
  if (whole) {
    const v = acEval(whole[1], scope);
    if (typeof v === "object") return String(v);
    return v;
  }

  let out = "";
  let i = 0;
  while (i < text.length) {
    const start = text.indexOf("${", i);
    if (start < 0) { out += text.substring(i); break; }
    out += text.substring(i, start);
    const end = text.indexOf("}", start);
    if (end < 0) { out += text.substring(start); break; }
    const v = acEval(text.substring(start + 2, end), scope);
    out += (v === null || v === undefined) ? "" : String(v);
    i = end + 1;
  }
  return out;
}

function acIsTruthy(v: (string | number | boolean | object)): boolean {
  if (v === true) return true;
  if (v === false || v === null || v === undefined) return false;
  if (v === "" || v === 0) return false;
  if (v === "false") return false;
  return true;
}

/** expand ทั้งก้อน — คืน null เมื่อ $when เป็นเท็จ (ให้ผู้เรียกตัดทิ้ง) */
function acExpandNode(node: (string | number | boolean | object), scope: Scope):
    (string | number | boolean | object) {

  if (typeof node === "string") return acExpandString(node, scope);
  if (typeof node === "number" || typeof node === "boolean") return node;
  if (node === null || node === undefined) return node;

  if (Array.isArray(node)) {
    const arr: (string | number | boolean | object)[] = [];
    for (const child of (node as (string | number | boolean | object)[])) {
      // element ที่มี $data ต้องขยายเป็นหลาย element
      if (child !== null && typeof child === "object" && !Array.isArray(child) &&
          (child as { [k: string]: object })["$data"] !== undefined) {
        const spec = (child as { [k: string]: string })["$data"];
        const bound = typeof spec === "string" ? acExpandStringRaw(spec, scope) : spec;
        if (Array.isArray(bound)) {
          for (const row of (bound as object[])) {
            const one = acExpandNode(acStripKey(child, "$data"), { data: row, root: scope.root });
            if (one !== null) arr.push(one);
          }
          continue;
        }
      }
      const v = acExpandNode(child, scope);
      if (v !== null) arr.push(v);
    }
    return arr;
  }

  const obj = node as { [k: string]: (string | number | boolean | object) };

  // $when — ประเมินก่อน ถ้าเท็จให้ตัดทิ้งทั้ง element
  if (obj["$when"] !== undefined) {
    const cond = typeof obj["$when"] === "string"
      ? acExpandString(obj["$when"] as string, scope)
      : obj["$when"];
    if (!acIsTruthy(cond)) return null;
  }

  const out: { [k: string]: (string | number | boolean | object) } = {};
  for (const key of Object.keys(obj)) {
    if (key === "$when" || key === "$data") continue;
    const v = acExpandNode(obj[key], scope);
    if (v !== null) out[key] = v;
  }
  return out;
}

/** คืนค่าดิบของ ${...} (ใช้กับ $data ที่ต้องได้อาร์เรย์ ไม่ใช่ข้อความ) */
function acExpandStringRaw(text: string, scope: Scope): (string | number | boolean | object) {
  const whole = text.match(/^\$\{([^}]*)\}$/);
  if (whole) return acEval(whole[1], scope);
  return acExpandString(text, scope);
}

function acStripKey(obj: object, key: string): object {
  const src = obj as { [k: string]: (string | number | boolean | object) };
  const out: { [k: string]: (string | number | boolean | object) } = {};
  for (const k of Object.keys(src)) if (k !== key) out[k] = src[k];
  return out;
}

// ===========================================================================
// ส่วนที่ 4 — main
// ===========================================================================
/**
 * ขนาดของการ์ดตอนเดินสายจริง (worst case)
 *
 * ตอนส่งขึ้น Teams การ์ดถูกหุ้มอยู่ในซองข้อความอีกชั้น และอักขระนอก ASCII
 * อาจถูก escape เป็น `\uXXXX` ทำให้ตัวอักษรไทย 1 ตัวที่ปกติกิน 3 ไบต์
 * กลายเป็น 6 ไบต์ และ emoji 1 ตัวจาก 4 ไบต์กลายเป็น 12 ไบต์
 *
 * การ์ดใบนี้เป็นภาษาไทยเกือบทั้งใบ การวัดแบบ UTF-8 ตรง ๆ จึงต่ำกว่าความจริง
 * จนเคยหลุดเพดานไปแล้ว (Teams ตอบกลับ 413 RequestEntityTooLarge)
 * ฟังก์ชันนี้จึงนับแบบแย่ที่สุดที่เป็นไปได้ เพื่อให้งบข้างล่างเชื่อถือได้จริง
 */
function wireBytes(s: string): number {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) n += 1;                                   // ASCII ส่งดิบ 1 ไบต์
    else if (c >= 0xD800 && c <= 0xDBFF) { n += 12; i++; }  // emoji = \uXXXX 2 ชุด
    else n += 6;                                            // ไทย/สัญลักษณ์ = \uXXXX
  }
  return n;
}

// ---------------------------------------------------------------------------
// 🔧 งบขนาดการ์ด — ปรับได้ที่บรรทัดเดียวนี้
//
// Teams ปฏิเสธข้อความที่ใหญ่เกิน 28 KB ด้วย error 413 RequestEntityTooLarge
// แต่ 28 KB นั้นนับ "ทั้งซองข้อความ" ไม่ใช่แค่ตัวการ์ด และส่วนที่หุ้มอยู่
// สคริปต์มองไม่เห็น จึงต้องเผื่อที่ไว้
//
// ตัวเลขที่วัดจากไฟล์จริง (หน่วยเป็นไบต์แบบ worst case ตาม wireBytes ข้างบน):
//     26,859  การ์ดหน้าตาแรกสุด (5 คอลัมน์แยก, 8 รหัส) -> Teams ปฏิเสธจริง ❌
//     27,519  หน้าตาเดิม (4 คอลัมน์แยก) พอเพิ่มเป็น 10 รหัส/45 แถว -> เหลือขอบ
//             แค่ 1.1 KB เสี่ยงมาก จึงรวม Onhand/กระจาย/เบิก Hub เป็นบรรทัดเดียว
//     22,528  หน้าตาปัจจุบัน (2 คอลัมน์ ตัวเลขรวมบรรทัดเดียว) รอบ Optimize
//             45 แถวครบทุกรหัส ✅ — เหลือขอบถึงเพดาน 28 KB ราว 6 KB
//     17,510  รอบ MA 29 แถวครบทุกรหัส ✅
//
// ตั้งไว้ 24,500 เพื่อให้ "ทุกรหัสลงการ์ดครบ" ตามที่ต้องการ โดยยังต่ำกว่า
// เพดาน 28 KB ราว 3.5 KB และมีขอบเหลือจากขนาดจริงของรอบที่ใหญ่ที่สุดอีกราว 2 KB
//
// ถ้าเจอ RequestEntityTooLarge อีก ให้ลดเลขนี้ลงทีละ 2,000 (แลกกับแถวที่หายไป)
// ถ้ารอบไหนข้อมูลโตจนถูกตัดแถว จะเห็นเป็น "แสดง n/N" บนหัวกลุ่ม Zone
// ---------------------------------------------------------------------------
const CARD_SAFE_BYTES = 24500;

/**
 * พารามิเตอร์ทุกตัวหลัง workbook ประกาศเป็นแบบไม่บังคับ (`?`) โดยตั้งใจ
 *
 * Power Automate ถือว่าพารามิเตอร์ที่ประกาศด้วย default value (`period: string = ""`)
 * เป็นช่อง "บังคับกรอก" (มีดอกจัน) ปล่อยว่างแล้วขึ้น error
 *      Invalid parameter for 'Run script'. Error: 'ScriptParameters/period' is required.
 * มีแต่พารามิเตอร์ที่ประกาศด้วย `?` เท่านั้นที่ Flow ยอมให้เว้นว่างได้
 * ค่าตั้งต้นจึงย้ายไปเติมด้วย resolveText() / resolveTopRows() แทน
 *
 * ⚠️ หลังวางสคริปต์เวอร์ชันนี้ทับของเดิม ต้องลบแล้วเพิ่ม action "Run script"
 *    ใน Flow ใหม่ (หรือเลือกสคริปต์ซ้ำอีกครั้ง) เพื่อให้ Flow อ่านรายการ
 *    พารามิเตอร์ชุดใหม่ — ของเดิมจะยังจำว่าเป็นช่องบังคับอยู่
 */
function main(
  workbook: ExcelScript.Workbook,
  period?: string,
  topRows?: number,
  dashboardUrl?: string,
  sourceFileUrl?: string,
  generatedAt?: string,
  reportType?: string
): (string | number | boolean | object) {

  // reportType ไม่ใช้ในรอบนี้ แต่คงไว้เพื่อให้ Flow เดิมที่ตั้งค่าไว้แล้วยังใช้ได้
  const periodText = resolveText(period);
  const dashUrl = resolveText(dashboardUrl);
  const srcUrl = resolveText(sourceFileUrl);
  const genAt = resolveText(generatedAt);

  // สร้างการ์ดแล้ววัดขนาดจริง ถ้าเกินงบก็ลดแถวแล้วสร้างใหม่
  // ทำให้การ์ดปลอดภัยเองไม่ว่ารอบนั้นจะมีกี่ Zone กี่รายการ โดยไม่ต้องมาไล่ปรับ topRows ทีหลัง
  let rows = resolveTopRows(topRows);
  let card: (string | number | boolean | object) = {};

  for (let attempt = 0; attempt < 8; attempt++) {
    const payload = buildPayload(workbook, periodText, rows, dashUrl, srcUrl);
    payload.meta.generatedAt = genAt;
    card = acExpandNode(CARD_TEMPLATE, { data: payload, root: payload });

    const size = wireBytes(JSON.stringify(card));
    const shown = payload.meta.shownItems;
    if (size <= CARD_SAFE_BYTES || shown <= 4) break;

    // ลดตามสัดส่วนที่เกินงบ แล้วบังคับให้ลดลงอย่างน้อย 1 แถวเสมอ กันวนไม่จบ
    let next = Math.floor(shown * CARD_SAFE_BYTES / size);
    if (next >= shown) next = shown - 1;
    if (next < 1) next = 1;
    rows = next;
  }

  return card;
}
