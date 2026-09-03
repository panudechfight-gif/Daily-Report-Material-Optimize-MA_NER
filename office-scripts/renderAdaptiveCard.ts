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
// คงพารามิเตอร์ topRows ไว้เพื่อให้ Flow เดิมเรียก Script นี้ได้โดยไม่เสียโครงสร้าง
// แต่รายการเป้าหมายจะไม่ถูกตัดด้วยค่านี้อีกแล้ว เพราะการ์ดแบบบรรทัดรวมมีขนาดเล็กลงมาก
const TOP_ROWS_MAX = 40;

// ---------------------------------------------------------------------------
// รหัสวัสดุที่การ์ดต้องแสดง — กำหนดเป็นรายการตายตัว 6 รหัสตามลำดับบน Card
// เพิ่ม/ลดรหัส ให้แก้ที่นี่ที่เดียว (แล้วแก้ TARGET_ITEMS ใน tools/build_sample_data.py ตามกัน)
// 52CL004BB, 52CL010BB และรหัสอื่นที่ไม่อยู่ในชุดนี้จะไม่แสดง
// ---------------------------------------------------------------------------
const TARGET_ITEM_ORDER: string[] = [
  "50MT004BB",
  "53OF150BB",
  "52CL003BB",
  "52CL009BB",
  "53OF157BB",
  "53OF158BB",
];

const TARGET_ITEMS: { [k: string]: string } = {
  "50MT004BB": "Name Plate(A)",
  "53OF150BB": "SC/UPC DW 1C",
  "52CL003BB": "Closure 12C",
  "52CL004BB": "Closure 24C",
  "52CL009BB": "Closure 12C(In)",
  "53OF157BB": "Cable ARSS 12C",
  "53OF158BB": "Cable ARSS 24C",
  "53OF160BB": "Cable ARSS 60C",

};

function isTargetItem(code: string): boolean {
  const c = (code || "").trim().toUpperCase();
  return TARGET_ITEMS[c] !== undefined;
}

// ---------------------------------------------------------------------------
// 🔧 เกณฑ์ตัดสินสถานะ — ⚠️ เสี่ยงขาด / 🟡 พอใช้ / ✅ ปกติ
//
// สถานะทั้ง 3 ตัวคิดจาก "ปริมาณของรหัส 53OF157BB" เท่านั้น
// รหัสอื่นอีก 5 ตัวยังแสดงตัวเลขบนการ์ดครบตามเดิม แต่ไม่มีผลกับสถานะ
// 1 Zone มี 53OF157BB ได้ 1 ค่า -> 1 Zone จึงมีสถานะเดียว
//
// โหมดการคิด เลือกได้ที่ STATUS_USE_QTY_THRESHOLD:
//   false (ค่าตั้งต้น) = ใช้คอลัมน์ Risk_Flag ของแถว 53OF157BB ที่ชีตคำนวณไว้แล้ว
//                        (ถ้าแถวนั้น Risk_Flag ว่าง จะถอยไปใช้เกณฑ์ปริมาณข้างล่างเอง)
//   true               = ไม่สนใจ Risk_Flag ใช้ Onhand เทียบเกณฑ์ปริมาณอย่างเดียว
// ---------------------------------------------------------------------------
const STATUS_ITEM_CODE = "53OF157BB";
const STATUS_USE_QTY_THRESHOLD = false;

// เกณฑ์ปริมาณคงเหลือ (Onhand) ของ 53OF157BB — หน่วยเดียวกับในชีต
//   Onhand <= STATUS_RISK_MAX   -> ⚠️ เสี่ยงขาด
//   Onhand <= STATUS_WATCH_MAX  -> 🟡 พอใช้
//   มากกว่า STATUS_WATCH_MAX    -> ✅ ปกติ
// ปรับตัวเลขได้ที่ 2 บรรทัดนี้ที่เดียว ไม่ต้องแก้ที่อื่น
const STATUS_RISK_MAX = 2000;
const STATUS_WATCH_MAX = 6000;

/**
 * ระดับสถานะของโซน คิดจากข้อมูลแถว 53OF157BB
 * คืน 2 = ⚠️ เสี่ยงขาด, 1 = 🟡 พอใช้, 0 = ✅ ปกติ
 */
function statusLevelOf(onhand: number, riskFlag: string): number {
  if (!STATUS_USE_QTY_THRESHOLD) {
    const f = riskFlag || "";
    if (f.indexOf("เสี่ยง") >= 0) return 2;
    if (f.indexOf("พอใช้") >= 0 || f.indexOf("เฝ้าระวัง") >= 0) return 1;
    if (f !== "" && f.indexOf("ปกติ") >= 0) return 0;
    if (f !== "") return 1;      // มีธงอย่างอื่นที่ไม่ใช่ "ปกติ" = ยังไม่ถึงขั้นเสี่ยง
    // Risk_Flag ว่าง -> ตกลงไปใช้เกณฑ์ปริมาณข้างล่าง
  }
  if (onhand <= STATUS_RISK_MAX) return 2;
  if (onhand <= STATUS_WATCH_MAX) return 1;
  return 0;
}

/** ป้ายสถานะสำหรับหัวกลุ่ม Zone (-1 = ไม่มีรหัส 53OF157BB ในโซนนั้น) */
function statusBadgeOf(level: number): string {
  if (level === 2) return "⚠️ เสี่ยงขาด";
  if (level === 1) return "🟡 พอใช้";
  if (level === 0) return "✅ ปกติ";
  return "";
}

/** สีของหัวกลุ่ม Zone ตามระดับสถานะ */
function statusColorOf(level: number): string {
  if (level === 2) return "attention";
  if (level === 1) return "warning";
  if (level === 0) return "good";
  return "default";
}

/** ลำดับ Item Code บนการ์ดต้องตรงกับ TARGET_ITEM_ORDER เสมอ */
function targetItemRank(code: string): number {
  const i = TARGET_ITEM_ORDER.indexOf((code || "").trim().toUpperCase());
  return i < 0 ? TARGET_ITEM_ORDER.length : i;
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
 * จัดรูปแบบเป็นจำนวนเต็มเสมอ — ใช้กับ "จัดสรร(เบิก Hub)" ที่ต้องเป็นยอดจ่ายจริง
 * ปัดครึ่งขึ้นแบบคณิตศาสตร์ (0.5 -> 1, -0.5 -> -1) ไม่ให้ค่าติดลบเพี้ยน
 */
function fmtInt(v: (string | number | boolean)): string {
  const n = toNum(v);
  const rounded = n < 0 ? -Math.round(Math.abs(n)) : Math.round(n);
  return addThousandSep(String(rounded));
}

// ---------------------------------------------------------------------------
// 🔧 กัน Markdown ของ Teams ตีความบรรทัดในคอลัมน์ตัวเลขผิด
//
// การ์ดนี้รวมค่าของทุกรายการในโซนเป็นข้อความก้อนเดียวคั่นด้วย "\n" แล้ววางใน
// TextBlock เดียว (ทำให้การ์ดเล็กพอจะใส่ได้ทุก Zone) แต่ Teams อ่านข้อความนั้น
// เป็น Markdown ก่อนเสมอ บรรทัดที่เป็นขีดล้วน ๆ จึงกลายเป็นไวยากรณ์:
//
//      0            <- บรรทัดข้อความ
//      -            <- Teams อ่านเป็น "เส้นใต้หัวข้อ" (setext heading)
//                      => บรรทัด 0 ข้างบนกลายเป็นหัวข้อ H2 ตัวใหญ่และหนา
//
//      -            <- ขีดที่ขึ้นต้นบล็อก => กลายเป็น bullet list (•)
//
// นี่คือสาเหตุที่ตัวเลข Onhand / จัดสรร(กระจาย) แสดงใหญ่และหนาผิดปกติบน
// Teams for Windows ส่วนคอลัมน์ที่มีข้อมูลครบทุกบรรทัด (ไม่มีขีด) แสดงปกติ
//
// ทางแก้: ใช้ en dash (–) แทนขีดปกติสำหรับช่องที่ไม่มีข้อมูล — หน้าตาเหมือนเดิม
// แต่ Markdown ไม่ถือเป็นไวยากรณ์ พร้อมกรองบรรทัดที่เป็นสัญลักษณ์ Markdown ล้วน
// ทิ้งอีกชั้นหนึ่ง เผื่อค่าจากชีตต้นทางหลุดมาในรูปแบบอื่น
// ---------------------------------------------------------------------------
const NO_DATA = "–";           // en dash (U+2013) — Markdown ไม่ตีความ

/** ทำให้ข้อความ 1 บรรทัดปลอดภัยต่อ Markdown ก่อนเอาไปต่อกันด้วย "\n" */
function mdSafeLine(v: string): string {
  const s = v === null || v === undefined ? "" : String(v);
  const t = s.trim();
  if (t === "") return NO_DATA;
  // บรรทัดที่มีแต่สัญลักษณ์ Markdown (- = _ * + # > ~) คือ heading / hr / list
  let onlyMarks = true;
  for (let i = 0; i < t.length; i++) {
    if ("-=_*+#>~".indexOf(t.charAt(i)) < 0) { onlyMarks = false; break; }
  }
  return onlyMarks ? NO_DATA : s;
}

interface CardItem {
  itemCode: string; itemName: string; unit: string; province: string;
  provinceCount: number; provinceTag: string;
  dataSet: string; period: string; prevPeriod: string;
  onhand: string; fromHub: string; distributed: string;
  // ยอด "จัดสรร(กระจาย)" ของรอบก่อนหน้า (รหัสเดียวกัน โซนเดียวกัน)
  // ใช้เทียบให้เห็นว่ารอบก่อนกระจายไปเท่าไหร่ รอบนี้ได้เพิ่มมาอีกเท่าไหร่
  prevDistributed: string;
  status: string; statusShort: string; statusEmoji: string; statusColor: string;
  riskFlag: string; isRisk: boolean; rowWeight: string;
  zone: string; zoneEmoji: string;
}

interface CardZone {
  zone: string; zoneEmoji: string; itemCount: number; shownCount: number;
  riskCount: number; riskBadge: string; riskColor: string; countLabel: string;
  totalOnhand: string; totalHub: string; totalLabel: string; zoneHeadline: string;
  // ยอดรวมของ "จัดสรร(กระจาย)" ทั้งโซน — รอบก่อนหน้า กับ รอบนี้
  totalPrevDistributed: string; totalDistributed: string; deltaLabel: string;
  itemLines: string; onhandLines: string; fromHubLines: string;
  prevDistributedLines: string; distributedLines: string;
  items: CardItem[];
}

/** ตารางอธิบายรหัสวัสดุท้ายการ์ด — พิมพ์ครั้งเดียวแทนการซ้ำชื่อยาวทุกแถว */
interface CardLegend { itemCode: string; itemName: string; }

interface CardMeta {
  title: string; subtitle: string; hasLegend: boolean;
  period: string; prevPeriod: string;
  dataSet: string; dataSetLabel: string; dataSetEmoji: string;
  dataSetStyle: string; dataSetColor: string;
  generatedAt: string;
  totalItems: number; shownItems: number; totalZones: number;
  riskCount: number; watchCount: number; okCount: number;
  // คำอธิบายว่าสถานะบนการ์ดคิดมาจากรหัสไหน
  statusBasis: string;
  // แสดงคอลัมน์ "จัดสรร(เบิก Hub)" หรือไม่ — ใช้เฉพาะรายงาน MA
  showHub: boolean;
  hasData: boolean; hasPrevData: boolean;
  dashboardUrl: string; sourceFileUrl: string;
}

interface CardPayload {
  meta: CardMeta; zones: CardZone[]; items: CardItem[];
  legend: CardLegend[]; legendText: string;
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

  const need = ["Distribution period", "Zone", "Item Code", "Onhand",
    "จัดสรร(กระจาย)", "จัดสรร(เบิก Hub)", "Prev_Period",
    "Status", "Risk_Flag", "Data_Set"];
  const missing = need.filter(n => col[n] === undefined);
  if (missing.length > 0) throw new Error("ไม่พบคอลัมน์: " + missing.join(", "));

  // -------------------------------------------------------------------------
  // แก้แนวคอลัมน์เลื่อนในชีตต้นทาง
  //
  // ชีต MA&Optimize NER มีคอลัมน์ว่างที่ไม่มีหัวตารางแทรกอยู่หลัง Onhand
  // ทำให้ข้อมูลของบล็อก จัดสรร(กระจาย) / จัดสรร(เบิก Hub) / จัดสรรเพิ่ม / รวม 3 รายการ
  // ไปอยู่ทางขวาของหัวตารางตัวเองอยู่ 1 ช่อง ถ้าอ่านตามชื่อหัวตารางตรง ๆ
  // ค่าที่ได้ในช่อง "จัดสรร(เบิก Hub)" จะเป็นค่าของ "จัดสรร(กระจาย)" แทน
  //
  // ตรวจด้วยกฎที่ชีตคำนวณไว้เอง:
  // จัดสรร(กระจาย) + จัดสรร(เบิก Hub) + จัดสรรเพิ่ม = รวม 3 รายการ
  // ลองทั้งแบบไม่เลื่อนและเลื่อน 1 ช่อง แล้วเลือกแบบที่กฎนี้เป็นจริงทุกแถวที่สุ่มมา
  // ถ้าวันหนึ่งชีตต้นทางถูกแก้ให้ตรงแล้ว ตัวตรวจนี้จะเลือกแบบไม่เลื่อนเองอัตโนมัติ
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

  // ---- 1. คัดเฉพาะ 6 Item Code ที่กำหนดใน TARGET_ITEM_ORDER ----
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
  // เก็บทุกแถวของรอบไว้หา Zone ให้ครบ แม้บาง Zone จะขาด Item Code เป้าหมายในต้นทาง
  const periodRows: (string | number | boolean)[][] = [];
  for (let r = hdrIdx + 1; r < values.length; r++) {
    if (get(values[r], "Distribution period") === targetPeriod) periodRows.push(values[r]);
  }
  const inPeriod = periodRows.filter(row => isTargetItem(get(row, "Item Code")));

  // ---- 3. รวมแถวซ้ำ: 1 Zone + 1 Item Code = 1 แถว ----
  // ชีตต้นทางแตกแถวตามจังหวัด รหัสเดียวกันจึงโผล่หลายครั้งในโซนเดียว
  // การ์ดไม่ได้แสดงจังหวัด การเห็นรหัสซ้ำจึงอ่านไม่รู้เรื่องและกินที่เปล่า ๆ
  // จึงบวกยอดของทุกจังหวัดเข้าด้วยกัน แล้วเก็บจำนวนจังหวัดไว้บอกที่มา
  const numOf = (row: (string | number | boolean)[], name: string): number => {
    const i = col[name];
    return i === undefined ? 0 : toNum(row[i]);
  };

  interface Agg {
    zone: string; itemCode: string; itemName: string; unit: string;
    onhand: number; fromHub: number; distributed: number;
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
      aggMap[key] = {
        zone: zone, itemCode: code,
        itemName: TARGET_ITEMS[code] || get(row, "Description") || get(row, "Art No"),
        unit: get(row, "Unit"),
        onhand: 0, fromHub: 0, distributed: 0,
        provinces: {}, provinceCount: 0,
        status: "", statusRank: -1, isRisk: false, riskFlag: "",
      };
      aggOrder.push(key);
    }
    const a = aggMap[key];

    a.onhand += numOf(row, "Onhand");
    a.fromHub += numOf(row, "จัดสรร(เบิก Hub)");
    a.distributed += numOf(row, "จัดสรร(กระจาย)");

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

  // ---- 3.1 ยอด "จัดสรร(กระจาย)" ของรอบก่อนหน้า ----
  // ไม่ต้องพึ่งคอลัมน์สรุปใด ๆ เพิ่ม — อ่านซ้ำจากชีตเดิมโดยใช้ชื่อรอบที่ได้จาก
  // คอลัมน์ Prev_Period ของรอบปัจจุบัน แล้วรวมยอดด้วยกติกาเดียวกับรอบนี้
  // (1 Zone + 1 Item Code = 1 ค่า, บวกทุกจังหวัดในโซน)
  const prevDistMap: { [k: string]: number } = {};      // "Zone|ItemCode" -> ยอดรอบก่อน
  const prevZoneDist: { [k: string]: number } = {};     // "Zone"          -> ยอดรวมรอบก่อน
  let hasPrevData = false;
  if (prevPeriod !== "") {
    for (let r = hdrIdx + 1; r < values.length; r++) {
      const row = values[r];
      if (row === undefined) continue;
      if (get(row, "Distribution period") !== prevPeriod) continue;
      if (!isTargetItem(get(row, "Item Code"))) continue;
      const z = get(row, "Zone") || "-";
      const c = get(row, "Item Code").toUpperCase();
      const k = z + "|" + c;
      const v = numOf(row, "จัดสรร(กระจาย)");
      prevDistMap[k] = (prevDistMap[k] || 0) + v;
      prevZoneDist[z] = (prevZoneDist[z] || 0) + v;
      hasPrevData = true;
    }
  }

  /** ยอดรอบก่อนของคู่ Zone|Item — ไม่มีข้อมูลจริงให้แสดงขีด ไม่ใช่ 0 */
  const prevDistText = (zoneName: string, code: string): string => {
    const v = prevDistMap[zoneName + "|" + code];
    return v === undefined ? NO_DATA : fmtNum(v);
  };

  const all: CardItem[] = [];
  let riskCount = 0, watchCount = 0, okCount = 0;

  // สถานะของแต่ละ Zone — คิดจากแถว STATUS_ITEM_CODE (53OF157BB) เท่านั้น
  // "Zone" -> 2 เสี่ยงขาด / 1 พอใช้ / 0 ปกติ (ไม่มีคีย์ = โซนนั้นไม่มีรหัสนี้)
  const zoneStatusLevel: { [k: string]: number } = {};

  for (const key of aggOrder) {
    const a = aggMap[key];
    const st = statusTheme(a.status);
    // ตัวนับ ⚠️/🟡/✅ นับเฉพาะ 53OF157BB — รหัสอื่นไม่นับ แต่ยังแสดงตัวเลขบนการ์ดครบ
    if (a.itemCode === STATUS_ITEM_CODE) {
      const lv = statusLevelOf(a.onhand, a.riskFlag);
      zoneStatusLevel[a.zone] = lv;
      if (lv === 2) riskCount++; else if (lv === 1) watchCount++; else okCount++;
    }

    all.push({
      itemCode: a.itemCode,
      itemName: a.itemName.length > 45 ? a.itemName.substring(0, 45) : a.itemName,
      unit: a.unit,
      province: a.provinceCount > 0 ? String(a.provinceCount) + " จังหวัด" : "-",
      provinceCount: a.provinceCount,
      // ต่อท้ายรหัสเฉพาะตอนที่ยอดนี้รวมมาจากหลายจังหวัด จะได้รู้ว่าตัวเลขมาจากไหน
      provinceTag: a.provinceCount > 1 ? " ·" + String(a.provinceCount) + "จว." : "",
      dataSet: dataSet,
      period: targetPeriod,
      prevPeriod: prevPeriod || "-",
      onhand: fmtNum(a.onhand),
      fromHub: fmtInt(a.fromHub),
      distributed: fmtNum(a.distributed),
      prevDistributed: prevDistText(a.zone, a.itemCode),
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

  // ---- 4. เรียง Item Code ตามลำดับที่ผู้ใช้กำหนด ----
  const cmpItem = (a: CardItem, b: CardItem): number => {
    const ra = targetItemRank(a.itemCode);
    const rb = targetItemRank(b.itemCode);
    if (ra !== rb) return ra - rb;
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

  // ---- 6. แสดงทุกรายการของทุก Zone โดยไม่ตัดด้วย topRows ----
  // สาเหตุเดิมที่ข้อมูลตกหล่นคือใช้ topRows เป็นเพดานรวมทั้ง Card และ main()
  // ลดแถวซ้ำเมื่อขนาด JSON สูงเกินงบ การ์ดเวอร์ชันนี้รวมค่าแต่ละคอลัมน์เป็น
  // TextBlock หลายบรรทัด จึงแสดงครบได้โดยไม่สร้าง element ซ้ำ 4 ชุดต่อ Item
  const byZone: { [k: string]: CardItem[] } = {};
  const zoneNames: string[] = [];
  const zoneSeen: { [k: string]: boolean } = {};
  for (const row of periodRows) {
    const z = get(row, "Zone") || "-";
    if (zoneSeen[z] === undefined) {
      zoneSeen[z] = true;
      zoneNames.push(z);
    }
  }
  for (const it of all) {
    if (byZone[it.zone] === undefined) byZone[it.zone] = [];
    if (zoneSeen[it.zone] === undefined) {
      zoneSeen[it.zone] = true;
      zoneNames.push(it.zone);
    }
    byZone[it.zone].push(it);
  }
  zoneNames.sort();
  for (const z of zoneNames) {
    const sourceItems = byZone[z] || [];
    sourceItems.sort((a, b) => cmpItem(a, b));

    // สร้างช่องว่างด้วย "-" เมื่อชีตต้นทางไม่มีรหัสนั้นจริง เพื่อให้แต่ละ Zone
    // ยังคงมี 6 บรรทัดตามลำดับเดียวกัน และแยกออกจากค่าศูนย์ที่เป็นข้อมูลจริง
    const itemByCode: { [k: string]: CardItem } = {};
    for (const it of sourceItems) itemByCode[it.itemCode] = it;
    const completeItems: CardItem[] = [];
    for (const code of TARGET_ITEM_ORDER) {
      const found = itemByCode[code];
      if (found !== undefined) {
        completeItems.push(found);
      } else {
        completeItems.push({
          itemCode: code,
          itemName: TARGET_ITEMS[code] || "",
          unit: "",
          province: "-",
          provinceCount: 0,
          provinceTag: "",
          dataSet: dataSet,
          period: targetPeriod,
          prevPeriod: prevPeriod || "-",
          onhand: NO_DATA,
          fromHub: NO_DATA,
          distributed: NO_DATA,
          // รหัสที่หายไปรอบนี้ อาจเคยมีรอบก่อน — แสดงไว้ให้เห็นว่าเคยกระจายเท่าไหร่
          prevDistributed: prevDistText(z, code),
          status: "ไม่มีข้อมูลในรอบนี้",
          statusShort: "",
          statusEmoji: "⚪",
          statusColor: "default",
          riskFlag: "-",
          isRisk: false,
          rowWeight: "default",
          zone: z,
          zoneEmoji: ZONE_EMOJI[z] || "🔷",
        });
      }
    }
    byZone[z] = completeItems;
  }

  const shown: CardItem[] = [];
  for (const z of zoneNames) {
    for (const it of byZone[z]) shown.push(it);
  }
  const grouped: { [k: string]: CardItem[] } = {};
  for (const it of shown) {
    if (!grouped[it.zone]) grouped[it.zone] = [];
    grouped[it.zone].push(it);
  }

  // ยอดรวมของแต่ละ Zone = บวกทุกจังหวัดในโซนนั้น (ไม่ใช่แค่แถวที่แสดง)
  const zoneOnhand: { [k: string]: number } = {};
  const zoneHub: { [k: string]: number } = {};
  const zoneDist: { [k: string]: number } = {};
  for (const it of all) {
    zoneOnhand[it.zone] = (zoneOnhand[it.zone] || 0) + toNum(it.onhand.replace(/,/g, ""));
    zoneHub[it.zone] = (zoneHub[it.zone] || 0) + toNum(it.fromHub.replace(/,/g, ""));
    zoneDist[it.zone] = (zoneDist[it.zone] || 0) + toNum(it.distributed.replace(/,/g, ""));
  }

  const zones: CardZone[] = Object.keys(grouped).sort().map(z => {
    const items = grouped[z];
    const nRisk = zoneRisk[z] || 0;
    const nAll = TARGET_ITEM_ORDER.length;
    // เทียบยอดกระจายทั้งโซน: รอบก่อน -> รอบนี้ (แสดงเฉพาะตอนมีข้อมูลรอบก่อนจริง)
    const zPrev = prevZoneDist[z];
    const zNow = zoneDist[z] || 0;
    const zDelta = zNow - (zPrev === undefined ? 0 : zPrev);
    const deltaText = zPrev === undefined
      ? ""
      : " · กระจาย " + fmtNum(zPrev) + " → " + fmtNum(zNow)
      + " (" + (zDelta > 0 ? "+" : "") + fmtNum(zDelta) + ")";
    // ป้ายสถานะของโซน คิดจากปริมาณของ 53OF157BB เพียงรหัสเดียว
    // โซนที่ไม่มีรหัสนี้ในรอบนั้น = ไม่ขึ้นป้าย ปล่อยหัวกลุ่มโล่งไว้
    const zLevel = zoneStatusLevel[z] === undefined ? -1 : zoneStatusLevel[z];
    const zBadge = statusBadgeOf(zLevel);
    return {
      zone: z,
      zoneEmoji: ZONE_EMOJI[z] || "🔷",
      itemCount: nAll,
      shownCount: items.length,
      riskCount: nRisk,
      riskBadge: zBadge,
      riskColor: statusColorOf(zLevel),
      countLabel: items.length === nAll
        ? nAll + " รายการ"
        : "แสดง " + items.length + " จาก " + nAll + " รายการ",
      // ยอดรวมทั้งโซน แสดงคู่กับหัวกลุ่ม
      totalOnhand: fmtNum(zoneOnhand[z] || 0),
      totalHub: fmtInt(zoneHub[z] || 0),
      totalLabel: "คงเหลือ " + fmtNum(zoneOnhand[z] || 0),
      totalPrevDistributed: zPrev === undefined ? NO_DATA : fmtNum(zPrev),
      totalDistributed: fmtNum(zNow),
      deltaLabel: deltaText,
      // หัวกลุ่มรวมเป็นบรรทัดเดียว — ประหยัดพื้นที่การ์ดได้ราว 0.5 KB ต่อ Zone
      zoneHeadline: (ZONE_EMOJI[z] || "🔷") + " " + z
        + " · " + (items.length === nAll ? nAll + " รายการ" : "แสดง " + items.length + "/" + nAll)
        + (zBadge !== "" ? " · " + zBadge : ""),
      // แสดงเฉพาะ Description ใต้แต่ละ Zone โดย itemName ยังคงอ้างอิง
      // TARGET_ITEMS ตาม Item Code ภายในระบบ ไม่แสดง Item Code ซ้ำบน Card
      // ทุกบรรทัดต้องผ่าน mdSafeLine() ก่อนต่อกันด้วย "\n" เสมอ
      // ไม่งั้น Teams จะอ่านบรรทัดขีดเป็น heading/bullet แล้วดันตัวเลขให้ใหญ่หนา
      itemLines: items.map(it => mdSafeLine(it.itemName)).join("\n"),
      // ไม่ใช้ markdown **bold** ที่นี่ เพราะ Teams บน Windows เรนเดอร์ตัวหนา
      // ในบรรทัดที่ถูกรวมหลายบรรทัดด้วยขนาดใหญ่ผิดปกติ (ไม่แคร์ "size": "small")
      // ทำให้ตัวเลขล้นคอลัมน์จนต้องเลื่อนดูทีละส่วน จึงใช้สัญลักษณ์แทนการทำตัวหนา
      onhandLines: items.map(it => mdSafeLine(it.isRisk ? "⚠ " + it.onhand : it.onhand)).join("\n"),
      fromHubLines: items.map(it => mdSafeLine(it.fromHub)).join("\n"),
      prevDistributedLines: items.map(it => mdSafeLine(it.prevDistributed)).join("\n"),
      distributedLines: items.map(it => mdSafeLine(it.distributed)).join("\n"),
      items: items,
    };
  });

  // ตารางอธิบายรหัสวัสดุ — พิมพ์ครั้งเดียวท้ายการ์ด แทนที่จะซ้ำทุกแถว
  const itemNameByCode: { [k: string]: string } = {};
  for (const it of all) {
    if (itemNameByCode[it.itemCode] === undefined && it.itemName !== "") {
      itemNameByCode[it.itemCode] = it.itemName;
    }
  }
  const legend: CardLegend[] = TARGET_ITEM_ORDER.map(c => {
    return { itemCode: c, itemName: TARGET_ITEMS[c] || itemNameByCode[c] || "" };
  });
  const legendText = legend.map(x => "**" + x.itemCode + "** · " + x.itemName).join("\n");

  const isMA = dataSet === "MA";
  const meta: CardMeta = {
    title: "Report Stock Allocation Cable&Material NER",
    subtitle: "",                          // ตัดบรรทัดคำอธิบายรหัสวัสดุออกแล้ว
    // ปิด Legend ท้าย Card เพราะ Description แสดงอยู่ใต้แต่ละ Zone แล้ว
    // ตัว expand จะตัด Container ของ Legend ออกจาก Card ทั้งก้อน
    hasLegend: false,
    period: targetPeriod || "-",
    prevPeriod: prevPeriod || "-",
    dataSet: dataSet,
    dataSetLabel: isMA ? "MA — Weekly Allocation" : "Optimize — Allocation Plan",
    // ใช้สัญลักษณ์ Store เดียวกันทั้งรายงาน MA และ OPTIMIZE
    dataSetEmoji: "🏬",
    dataSetStyle: isMA ? "accent" : "good",
    dataSetColor: isMA ? "accent" : "good",
    generatedAt: "",                       // Power Automate เติมด้วย convertTimeZone()
    totalItems: all.length,
    shownItems: shown.length,
    totalZones: zoneNames.length,
    riskCount: riskCount,
    watchCount: watchCount,
    okCount: okCount,
    // บอกให้ชัดบนการ์ดว่าสถานะ 3 ตัวข้างบนมาจากรหัสไหน จะได้ไม่เข้าใจผิดว่านับทุกรหัส
    statusBasis: "สถานะ ⚠️/🟡/✅ คำนวณจากรหัส " + STATUS_ITEM_CODE + " เท่านั้น",
    // คอลัมน์ "จัดสรร(เบิก Hub)" ใช้เฉพาะรายงาน MA — รอบ Optimize ซ่อนทิ้ง
    // main() ยังทับค่านี้ได้อีกชั้นด้วยพารามิเตอร์ reportType
    showHub: isMA,
    hasData: all.length > 0,
    hasPrevData: hasPrevData,
    dashboardUrl: dashboardUrl || "https://app.powerbi.com/CHANGE-ME",
    sourceFileUrl: sourceFileUrl || "https://mimotech-my.sharepoint.com/:x:/r/personal/panudeck_ais_co_th/Documents/Workflow%20Power%20Automate/Optimize_HiringMA_Full_Rebuild_V2.xlsx?d=w07489ab067624e578d50d67069db8e92&csf=1&web=1&e=7qCiSy",
  };

  // คืนทั้ง zones (การ์ดหลัก) และ items (การ์ดแบบง่าย) เผื่อเลือกใช้ได้ทั้งสองแบบ
  return { meta: meta, zones: zones, items: shown, legend: legend, legendText: legendText };
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
              "text": "🗓️ รอบจัดสรร",
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
              "text": "⏮️ รอบก่อนหน้า",
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
          "text": "📭",
          "size": "extraLarge",
          "horizontalAlignment": "center",
          "wrap": false
        },
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
                  "text": "⚠️ เสี่ยงขาด",
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
                  "text": "🟡 พอใช้",
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
                  "text": "✅ ปกติ",
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
            }
          ]
        },
        {
          "type": "TextBlock",
          "text": "${meta.statusBasis}",
          "size": "small",
          "isSubtle": true,
          "wrap": true,
          "spacing": "small"
        },
        {
          "type": "ColumnSet",
          "spacing": "medium",
          "separator": true,
          "columns": [
            {
              "type": "Column",
              "width": 40,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Zone",
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
              "width": 15,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Onhand",
                  "weight": "bolder",
                  "size": "small",
                  "color": "accent",
                  "horizontalAlignment": "right",
                  "wrap": true,
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "$when": "${$root.meta.showHub}",
              "width": 18,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "จัดสรร(เบิก Hub)",
                  "weight": "bolder",
                  "size": "small",
                  "color": "accent",
                  "horizontalAlignment": "right",
                  "wrap": true,
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "width": 20,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "จัดสรร(กระจาย) รอบก่อน",
                  "weight": "bolder",
                  "size": "small",
                  "color": "accent",
                  "horizontalAlignment": "right",
                  "wrap": true,
                  "spacing": "none"
                }
              ]
            },
            {
              "type": "Column",
              "width": 20,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "จัดสรร(กระจาย) รอบนี้",
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
                  "width": 40,
                  "items": [
                    {
                      "type": "TextBlock",
                      "text": "${itemLines}",
                      "size": "small",
                      "wrap": true,
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 15,
                  "items": [
                    {
                      "type": "TextBlock",
                      "text": "${onhandLines}",
                      "size": "small",
                      "weight": "default",
                      "horizontalAlignment": "right",
                      "wrap": true,
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "$when": "${$root.meta.showHub}",
                  "width": 18,
                  "items": [
                    {
                      "type": "TextBlock",
                      "text": "${fromHubLines}",
                      "size": "small",
                      "horizontalAlignment": "right",
                      "wrap": true,
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 20,
                  "items": [
                    {
                      "type": "TextBlock",
                      "text": "${prevDistributedLines}",
                      "size": "small",
                      "weight": "default",
                      "isSubtle": true,
                      "horizontalAlignment": "right",
                      "wrap": true,
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 20,
                  "items": [
                    {
                      "type": "TextBlock",
                      "text": "${distributedLines}",
                      "size": "small",
                      "weight": "default",
                      "horizontalAlignment": "right",
                      "wrap": true,
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
          "text": " 🏷️Item Code · Description ",
          "size": "small",
          "weight": "bolder",
          "color": "accent",
          "wrap": true,
          "spacing": "none"
        },
        {
          "type": "TextBlock",
          "text": "${legendText}",
          "size": "small",
          "isSubtle": true,
          "wrap": true,
          "spacing": "none"
        }
      ]
    },
    {
      "type": "TextBlock",
      "text": "🔗 Source of information • ที่มา: ${meta.dataSet} / ${meta.period} • ${meta.generatedAt}",
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
      "title": "📁 Open the source file. (Excel / SharePoint)",
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
// สคริปต์มองไม่เห็น จึงต้องเผื่อที่ไว้ให้มากพอ
//
// ค่าเดิม 26,000 หลวมเกินไป — การ์ดรอบ Optimize ที่วัดได้ 25 KB ถูกปฏิเสธจริง
// ค่านี้จึงตั้งไว้ที่ราว 70% ของเพดาน โดยวัดแบบ worst case ข้างบน
//
// ถ้ายังเจอ RequestEntityTooLarge อีก ให้ลดเลขนี้ลงทีละ 2,000
// ถ้าอยากให้การ์ดจุแถวได้มากขึ้นและ Flow ผ่านสบาย ๆ แล้ว ค่อยเพิ่มขึ้นทีละ 1,000
// ---------------------------------------------------------------------------
const CARD_SAFE_BYTES = 20000;

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

  const periodText = resolveText(period);
  // reportType ใช้ตัดสินว่าจะแสดงคอลัมน์ "จัดสรร(เบิก Hub)" หรือไม่
  //   "MA"       -> แสดง
  //   "Optimize" -> ซ่อน
  //   เว้นว่าง    -> ตัดสินจากชุดข้อมูลจริงที่อ่านได้ (Data_Set)
  const reportKind = resolveText(reportType).toLowerCase();
  const dashUrl = resolveText(dashboardUrl);
  const srcUrl = resolveText(sourceFileUrl);
  const genAt = resolveText(generatedAt);

  // สร้างการ์ดแบบกะทัดรัดครั้งเดียว โดยไม่ลดจำนวน Item Code ของ Zone ใด
  // topRows ยังคงรับไว้เพื่อให้ Run script ใน Flow เดิมไม่เสีย parameter mapping
  const rows = resolveTopRows(topRows);
  const payload = buildPayload(workbook, periodText, rows, dashUrl, srcUrl);
  payload.meta.generatedAt = genAt;

  // ทับค่า showHub ด้วย reportType ที่ Flow ส่งมา (ถ้าส่งมา)
  if (reportKind === "optimize") payload.meta.showHub = false;
  else if (reportKind === "ma") payload.meta.showHub = true;

  let card = acExpandNode(CARD_TEMPLATE, { data: payload, root: payload });

  // ถ้าข้อความรอบใดยาวผิดปกติ ให้ตัดเฉพาะ Legend ซึ่งเป็นข้อมูลประกอบ
  // รายการหลัก 6 Item Code ของทุก Zone จะยังอยู่ครบเสมอ
  if (wireBytes(JSON.stringify(card)) > CARD_SAFE_BYTES && payload.meta.hasLegend) {
    payload.meta.hasLegend = false;
    card = acExpandNode(CARD_TEMPLATE, { data: payload, root: payload });
  }

  return card;
}
