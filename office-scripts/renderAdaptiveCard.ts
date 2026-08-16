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
// เพดานแถวที่ปลอดภัย วัดจริงจากข้อมูลในไฟล์ (ดู tools/test_render_script.js)
// การ์ดรวมแถวซ้ำของแต่ละ Zone แล้ว จำนวนแถวจริงจึงน้อยกว่าจำนวนแถวดิบมาก
// (เช่น MA (17 Aug 26) 74 แถวดิบ -> 24 แถวหลังรวม) เพดานนี้จึงครอบคลุมทั้งรอบได้
const TOP_ROWS_MAX = 40;

// ---------------------------------------------------------------------------
// รหัสวัสดุที่การ์ดต้องแสดง — กำหนดเป็นรายการตายตัว 8 รหัส
// เพิ่ม/ลดรหัส ให้แก้ที่นี่ที่เดียว (แล้วแก้ TARGET_ITEMS ใน tools/build_sample_data.py ตามกัน)
// ค่าของแต่ละรหัสคือชื่อที่ใช้แสดงในตารางอธิบายท้ายการ์ด
// ---------------------------------------------------------------------------
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
  itemCode: string; itemName: string; unit: string; province: string;
  provinceCount: number; provinceTag: string;
  dataSet: string; period: string; prevPeriod: string;
  onhand: string; fromHub: string; prevBefore: string; prevReceived: string;
  status: string; statusShort: string; statusEmoji: string; statusColor: string;
  riskFlag: string; isRisk: boolean; rowWeight: string;
  zone: string; zoneEmoji: string;
}

interface CardZone {
  zone: string; zoneEmoji: string; itemCount: number; shownCount: number;
  riskCount: number; riskBadge: string; riskColor: string; countLabel: string;
  totalOnhand: string; totalHub: string; totalLabel: string; zoneHeadline: string;
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
  hasData: boolean; hasMore: boolean; moreCount: number; moreText: string;
  dashboardUrl: string; sourceFileUrl: string;
}

interface CardPayload {
  meta: CardMeta; zones: CardZone[]; items: CardItem[]; legend: CardLegend[];
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

  const need = ["Distribution period", "Zone", "Item Code", "OMC-Onhand",
                "เบิกจาก Hub", "Prev_Period", "Prev_Before", "Prev_Received",
                "Status", "Risk_Flag", "Data_Set"];
  const missing = need.filter(n => col[n] === undefined);
  if (missing.length > 0) throw new Error("ไม่พบคอลัมน์: " + missing.join(", "));

  // -------------------------------------------------------------------------
  // แก้แนวคอลัมน์เลื่อนในชีตต้นทาง
  //
  // ชีต MA&Optimize NER มีคอลัมน์ว่างที่ไม่มีหัวตารางแทรกอยู่หลัง OMC-Onhand
  // ทำให้ข้อมูลของบล็อก จัดสรร / เบิกจาก Hub / จัดสรรเพิ่ม / รวม 3 รายการ
  // ไปอยู่ทางขวาของหัวตารางตัวเองอยู่ 1 ช่อง ถ้าอ่านตามชื่อหัวตารางตรง ๆ
  // ค่าที่ได้ในช่อง "เบิกจาก Hub" จะเป็นค่าของ "จัดสรร" แทน
  //
  // ตรวจด้วยกฎที่ชีตคำนวณไว้เอง:  จัดสรร + เบิกจาก Hub + จัดสรรเพิ่ม = รวม 3 รายการ
  // ลองทั้งแบบไม่เลื่อนและเลื่อน 1 ช่อง แล้วเลือกแบบที่กฎนี้เป็นจริงทุกแถวที่สุ่มมา
  // ถ้าวันหนึ่งชีตต้นทางถูกแก้ให้ตรงแล้ว ตัวตรวจนี้จะเลือกแบบไม่เลื่อนเองอัตโนมัติ
  // -------------------------------------------------------------------------
  const shiftNames = ["จัดสรร", "เบิกจาก Hub", "จัดสรรเพิ่ม", "รวม 3 รายการ"];
  const haveAll = shiftNames.every(n => col[n] !== undefined);
  if (haveAll) {
    let bestOffset = 0;
    for (let off = 0; off <= 1; off++) {
      let checked = 0, agree = 0;
      for (let r = hdrIdx + 1; r < values.length && checked < 300; r++) {
        const row = values[r];
        if (row === undefined) continue;
        const ia = col["จัดสรร"] + off, ib = col["เบิกจาก Hub"] + off;
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
    zone: string; itemCode: string; itemName: string; unit: string;
    onhand: number; fromHub: number; prevBefore: number; prevReceived: number;
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
        onhand: 0, fromHub: 0, prevBefore: 0, prevReceived: 0,
        provinces: {}, provinceCount: 0,
        status: "", statusRank: -1, isRisk: false, riskFlag: "",
      };
      aggOrder.push(key);
    }
    const a = aggMap[key];

    a.onhand += numOf(row, "OMC-Onhand");
    a.fromHub += numOf(row, "เบิกจาก Hub");
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
  const zoneOnhand: { [k: string]: number } = {};
  const zoneHub: { [k: string]: number } = {};
  for (const it of all) {
    zoneOnhand[it.zone] = (zoneOnhand[it.zone] || 0) + toNum(it.onhand.replace(/,/g, ""));
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
      riskBadge: nRisk > 0 ? "⚠️ เสี่ยงขาด " + nRisk : "",
      riskColor: nRisk > 0 ? "attention" : "good",
      countLabel: items.length === nAll
        ? nAll + " รายการ"
        : "แสดง " + items.length + " จาก " + nAll + " รายการ",
      // ยอดรวมทั้งโซน แสดงคู่กับหัวกลุ่ม
      totalOnhand: fmtNum(zoneOnhand[z] || 0),
      totalHub: fmtInt(zoneHub[z] || 0),
      totalLabel: "คงเหลือ " + fmtNum(zoneOnhand[z] || 0) + " · Hub " + fmtInt(zoneHub[z] || 0),
      // หัวกลุ่มรวมเป็นบรรทัดเดียว — ประหยัดพื้นที่การ์ดได้ราว 0.5 KB ต่อ Zone
      zoneHeadline: (ZONE_EMOJI[z] || "🔷") + " " + z
        + " · " + (items.length === nAll ? nAll + " รายการ" : "แสดง " + items.length + "/" + nAll)
        + " · คงเหลือ " + fmtNum(zoneOnhand[z] || 0)
        + " · Hub " + fmtInt(zoneHub[z] || 0)
        + (nRisk > 0 ? " · ⚠️ เสี่ยงขาด " + nRisk : ""),
      items: items,
    };
  });

  // ตารางอธิบายรหัสวัสดุ — พิมพ์ครั้งเดียวท้ายการ์ด แทนที่จะซ้ำทุกแถว
  const legendCodes: string[] = [];
  for (const it of all) {
    if (legendCodes.indexOf(it.itemCode) < 0) legendCodes.push(it.itemCode);
  }
  legendCodes.sort();
  const legend: CardLegend[] = legendCodes.map(c => {
    return { itemCode: c, itemName: TARGET_ITEMS[c] || "" };
  });

  const isMA = dataSet === "MA";
  const meta: CardMeta = {
    title: "Report Material Optimize&MA_NER",
    subtitle: "",                          // ตัดบรรทัดคำอธิบายรหัสวัสดุออกแล้ว
    hasLegend: legend.length > 0,
    period: targetPeriod || "-",
    prevPeriod: prevPeriod || "-",
    dataSet: dataSet,
    dataSetLabel: isMA ? "MA — Weekly Allocation" : "OPTIMIZE — Allocation Plan",
    dataSetEmoji: isMA ? "📦" : "🗃️",
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
    hasMore: moreCount > 0,
    moreCount: moreCount,
    moreText: moreCount > 0
      ? "…และอีก " + moreCount + " รายการ — กด “📊 Dashboard” เพื่อดูทั้งหมด"
      : "",
    dashboardUrl: dashboardUrl || "https://app.powerbi.com/CHANGE-ME",
    sourceFileUrl: sourceFileUrl || "https://contoso.sharepoint.com/CHANGE-ME",
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
  "fallbackText": "Report Material Optimize&MA_NER (อุปกรณ์ของคุณแสดงการ์ดนี้ไม่ได้ กรุณาเปิดจาก Dashboard)",
  "speak": "Report Material ${meta.dataSet} รอบ ${meta.period} มี ${meta.totalItems} รายการ เสี่ยงขาด ${meta.riskCount} รายการ",
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
            },
            {
              "type": "Column",
              "width": "stretch",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "📦 รวม",
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
              "width": 32,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Item Code",
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
              "width": 17,
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
              "width": 17,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "เบิก Hub",
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
              "width": 17,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Prev_Before",
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
              "width": 17,
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Prev_Recv",
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
                  "width": 32,
                  "items": [
                    {
                      "type": "TextBlock",
                      "$data": "${items}",
                      "text": "${statusEmoji} ${itemCode}${provinceTag}",
                      "size": "small",
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 17,
                  "items": [
                    {
                      "type": "TextBlock",
                      "$data": "${items}",
                      "text": "${onhand}",
                      "size": "small",
                      "weight": "${rowWeight}",
                      "color": "${statusColor}",
                      "horizontalAlignment": "right",
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 17,
                  "items": [
                    {
                      "type": "TextBlock",
                      "$data": "${items}",
                      "text": "${fromHub}",
                      "size": "small",
                      "horizontalAlignment": "right",
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 17,
                  "items": [
                    {
                      "type": "TextBlock",
                      "$data": "${items}",
                      "text": "${prevBefore}",
                      "size": "small",
                      "horizontalAlignment": "right",
                      "spacing": "none"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": 17,
                  "items": [
                    {
                      "type": "TextBlock",
                      "$data": "${items}",
                      "text": "${prevReceived}",
                      "size": "small",
                      "horizontalAlignment": "right",
                      "spacing": "none"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "TextBlock",
          "$when": "${meta.hasMore}",
          "text": "${meta.moreText}",
          "size": "small",
          "isSubtle": true,
          "horizontalAlignment": "center",
          "wrap": true,
          "spacing": "small",
          "separator": true
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
          "text": "รหัสวัสดุในรายงานนี้",
          "size": "small",
          "weight": "bolder",
          "color": "accent",
          "wrap": true,
          "spacing": "none"
        },
        {
          "type": "TextBlock",
          "$data": "${legend}",
          "text": "**${itemCode}** · ${itemName}",
          "size": "small",
          "isSubtle": true,
          "wrap": true,
          "spacing": "none"
        }
      ]
    },
    {
      "type": "TextBlock",
      "text": "🕗 สร้างรายงานเมื่อ ${meta.generatedAt} • ที่มา: ${meta.dataSet} / ${meta.period}",
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
      "title": "📊 Dashboard",
      "url": "${meta.dashboardUrl}"
    },
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
/** ความยาวจริงเป็นไบต์แบบ UTF-8 — ข้อความไทย 1 ตัวกิน 3 ไบต์ emoji กิน 4 ไบต์
 *  (String.length นับเป็น UTF-16 code unit จึงต่ำกว่าความจริงมากจนวัดขนาดการ์ดไม่ได้) */
function utf8Bytes(s: string): number {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) n += 1;
    else if (c < 0x800) n += 2;
    else if (c >= 0xD800 && c <= 0xDBFF) { n += 4; i++; }   // surrogate pair
    else n += 3;
  }
  return n;
}

// Teams ปฏิเสธการ์ดเกิน 28,672 ไบต์แบบเงียบ ๆ — เผื่อไว้ให้เหลือที่หายใจ
const CARD_SAFE_BYTES = 26000;

function main(
  workbook: ExcelScript.Workbook,
  period: string = "",
  topRows: number = 40,
  dashboardUrl: string = "",
  sourceFileUrl: string = "",
  generatedAt: string = "",
  reportType: string = ""
): (string | number | boolean | object) {

  // reportType ไม่ใช้ในรอบนี้ แต่ API ต้องการเพื่อความเข้ากันได้
  // สร้างการ์ดแล้ววัดขนาดจริง ถ้าเกินงบก็ลดแถวแล้วสร้างใหม่
  // ทำให้การ์ดปลอดภัยเองไม่ว่ารอบนั้นจะมีกี่ Zone กี่รายการ โดยไม่ต้องมาไล่ปรับ topRows ทีหลัง
  let rows = topRows;
  let card: (string | number | boolean | object) = {};

  for (let attempt = 0; attempt < 8; attempt++) {
    const payload = buildPayload(workbook, period, rows, dashboardUrl, sourceFileUrl);
    payload.meta.generatedAt = generatedAt || "";
    card = acExpandNode(CARD_TEMPLATE, { data: payload, root: payload });

    const size = utf8Bytes(JSON.stringify(card));
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
