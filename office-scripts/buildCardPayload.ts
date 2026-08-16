/**
 * buildCardPayload.ts — Office Script สำหรับ Power Automate
 * =========================================================
 * อ่านชีต "MA&Optimize NER" แล้วคืน JSON payload ที่พร้อมเสียบเข้า Adaptive Card
 * (โครงสร้างเดียวกับ adaptive-cards/data/sample-data-main.json)
 *
 * ทำไมถึงอ่านชีตต้นทางโดยตรง ไม่อ่าน PA_Export
 * ---------------------------------------------
 * ชีต PA_Export ในไฟล์ V2 มีบั๊กหัวตารางเลื่อน 1 คอลัมน์ตั้งแต่คอลัมน์ J
 * (ดู vba/modExport_FIXED.bas และ docs/03-DATA-MAPPING.md)
 * สคริปต์นี้อ่านหัวตารางภาษาไทยจากชีตต้นทางแถวที่ 7 โดยตรง จึงไม่โดนบั๊กนั้น
 * ถ้าติดตั้ง modExport_FIXED.bas แล้วจะเปลี่ยนไปอ่าน PA_Export ก็ได้ ไม่บังคับ
 *
 * วิธีใช้ใน Power Automate
 * ------------------------
 *   Action: Excel Online (Business) -> Run script
 *   Script: buildCardPayload
 *   Parameters:
 *       period        (เว้นว่าง = ใช้รอบล่าสุดอัตโนมัติ)
 *       topRows       12   <- อย่าเกิน 15 การ์ด Teams จำกัด 28 KB
 *       dashboardUrl  https://...
 *       sourceFileUrl https://...
 *
 * ผลลัพธ์: body('Run_script')?['result'] คือ payload ที่เอาไปวางในการ์ดได้เลย
 *
 * ข้อจำกัด: Office Scripts ทำงานกับไฟล์ .xlsx เท่านั้น (ไม่รองรับ .xlsm)
 *          ให้ Save As เป็น .xlsx ไว้บน SharePoint อีกไฟล์หนึ่ง
 */

const SHEET_SOURCE = "MA&Optimize NER";
const HEADER_ROW = 7;           // หัวตารางอยู่แถวที่ 7 (แถว 1-6 เป็นปุ่ม/คำอธิบาย)
// เพดานแถวที่ปลอดภัย วัดจริงจากข้อมูลในไฟล์ (ดู tools/test_render_script.js):
//   ชุด Optimize 10 แถว = 25.0 KB | 11 แถว = 26.3 KB | 12 แถว = 27.5 KB (เฉียดเพดาน)
//   ชุด MA      10 แถว = 23.2 KB | 12 แถว = 25.6 KB
// การ์ดกระจายแถวให้ทุก Zone จึงมีหัวกลุ่ม Zone หลายชุด ซึ่งกินที่คงที่ราว 1.5 KB ต่อ Zone
// ตั้งไว้ที่ 10 เพื่อกันชนเพดาน 28 KB ของ Teams ในรอบที่มี 5 Zone
const TOP_ROWS_MAX = 10;

// ---------------------------------------------------------------------------
// รหัสวัสดุที่การ์ดต้องแสดง — กำหนดเป็นรายการตายตัว 10 รหัส
// เพิ่ม/ลดรหัส ให้แก้ที่นี่ที่เดียว
// ---------------------------------------------------------------------------
const TARGET_ITEMS: { [k: string]: boolean } = {
  "50MT004BB": true,   // Name Plate (Aluminium)
  "52CL003BB": true,   // CLOSURE 12 C
  "52CL004BB": true,   // CLOSURE 24 C
  "52CL006BB": true,   // CLOSURE 48 C
  "52CL009BB": true,   // CLOSURE 12 C FOR OFC DROP WIRE (IN LINE)
  "52CL010BB": true,   // CLOSURE 60 C
  "53OF150BB": true,   // OPTICAL FIBER DROP CABLE 1C, FLAT TYPE
  "53OF157BB": true,   // ARSS OPTICAL FIBER CABLE 12c-FIBRE3
  "53OF158BB": true,   // ARSS OPTICAL FIBER CABLE 24c-FIBRE3
  "53OF160BB": true,   // ARSS OPTICAL FIBER CABLE 60c-FIBRE3
};

function isTargetItem(code: string): boolean {
  const c = (code || "").trim().toUpperCase();
  return TARGET_ITEMS[c] === true;
}

const ZONE_EMOJI: { [k: string]: string } = {
  "RC2-NMA": "🟦",
  "RC2-UBN": "🟩",
  "RC3-KKN": "🟨",
  "RC3-UDN": "🟧",
  "RC3-SNK": "🟪",
};

interface StatusTheme { emoji: string; color: string; short: string; }

function statusTheme(status: string): StatusTheme {
  const s = status || "";
  if (s.indexOf("ปกติ") >= 0) return { emoji: "✅", color: "good", short: "ปกติ" };
  if (s.indexOf("จัดสรรเกิน") >= 0) return { emoji: "🟡", color: "warning", short: "จัดสรรเกิน" };
  if (s.indexOf("ใช้เกินยอดรับ") >= 0) return { emoji: "🔴", color: "attention", short: "ใช้เกินยอดรับ" };
  if (s.indexOf("ยอดไม่ตรง") >= 0) return { emoji: "🟠", color: "warning", short: "ยอดไม่ตรง" };
  if (s.indexOf("ไม่มีข้อมูลรอบก่อน") >= 0) return { emoji: "⚪", color: "default", short: "ไม่มีรอบก่อน" };
  return { emoji: "⚪", color: "default", short: "ไม่ระบุ" };
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
  dataSet: string; period: string; prevPeriod: string;
  onhand: string; fromHub: string; prevBefore: string; prevReceived: string;
  status: string; statusShort: string; statusEmoji: string; statusColor: string;
  riskFlag: string; isRisk: boolean; rowWeight: string;
  zone: string; zoneEmoji: string;
}

interface CardZone {
  zone: string; zoneEmoji: string; itemCount: number; shownCount: number;
  riskCount: number; riskBadge: string; riskColor: string; countLabel: string;
  items: CardItem[];
}

interface CardMeta {
  title: string; subtitle: string; period: string; prevPeriod: string;
  dataSet: string; dataSetLabel: string; dataSetEmoji: string;
  dataSetStyle: string; dataSetColor: string;
  generatedAt: string;
  totalItems: number; shownItems: number; totalZones: number;
  riskCount: number; watchCount: number; okCount: number;
  hasData: boolean; hasMore: boolean; moreCount: number; moreText: string;
  dashboardUrl: string; sourceFileUrl: string;
}

interface CardPayload { meta: CardMeta; zones: CardZone[]; items: CardItem[]; }

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

  // ---- 2. เลือกรอบ: ถ้าไม่ระบุ ใช้รอบ MA ที่ Period_Seq สูงสุด ----
  // Power Automate บังคับให้พารามิเตอร์ต้องมีค่า ส่งค่าว่างจริง ๆ ไม่ได้
  // จึงรับค่าตัวแทน "ไม่ระบุ" ได้หลายแบบ แล้วแปลงเป็นค่าว่างให้เอง
  let targetPeriod = (period || "").trim();
  const blank = targetPeriod.toLowerCase();
  if (blank === "-" || blank === "auto" || blank === "latest" || blank === "null") {
    targetPeriod = "";
  }
  if (targetPeriod === "") {
    let bestSeq = -1;
    for (const row of picked) {
      if (get(row, "Data_Set") !== "MA") continue;
      const seq = toNum(col["Period_Seq"] !== undefined ? row[col["Period_Seq"]] : 0);
      if (seq > bestSeq) { bestSeq = seq; targetPeriod = get(row, "Distribution period"); }
    }
  }
  const inPeriod = picked.filter(row => get(row, "Distribution period") === targetPeriod);

  // ---- 3. แปลงเป็น item + นับสถานะ ----
  const all: CardItem[] = [];
  let riskCount = 0, watchCount = 0, okCount = 0;
  let dataSet = "MA";
  let prevPeriod = "";

  for (const row of inPeriod) {
    const ds = get(row, "Data_Set") || "MA";
    dataSet = ds;
    const pp = get(row, "Prev_Period");
    if (pp !== "" && prevPeriod === "") prevPeriod = pp;

    const st = statusTheme(get(row, "Status"));
    const risk = get(row, "Risk_Flag");
    const isRisk = risk.indexOf("เสี่ยง") >= 0;
    if (isRisk) riskCount++; else if (risk !== "") watchCount++; else okCount++;

    const zone = get(row, "Zone") || "-";
    const desc = get(row, "Description") || get(row, "Art No");

    all.push({
      itemCode: get(row, "Item Code"),
      itemName: desc.length > 45 ? desc.substring(0, 45) : desc,
      unit: get(row, "Unit"),
      province: get(row, "Province") || "-",
      dataSet: ds,
      period: targetPeriod,
      prevPeriod: pp || "-",
      onhand: fmtNum(col["OMC-Onhand"] !== undefined ? row[col["OMC-Onhand"]] : 0),
      fromHub: fmtInt(col["เบิกจาก Hub"] !== undefined ? row[col["เบิกจาก Hub"]] : 0),
      prevBefore: fmtNum(col["Prev_Before"] !== undefined ? row[col["Prev_Before"]] : 0),
      prevReceived: fmtNum(col["Prev_Received"] !== undefined ? row[col["Prev_Received"]] : 0),
      status: (get(row, "Status") || "-").replace(/^-\s*/, "") || "-",
      statusShort: st.short,
      statusEmoji: st.emoji,
      statusColor: st.color,
      riskFlag: risk || "-",
      isRisk: isRisk,
      rowWeight: isRisk ? "bolder" : "default",
      zone: zone,
      zoneEmoji: ZONE_EMOJI[zone] || "🔷",
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
      riskBadge: nRisk > 0 ? "⚠️ เสี่ยงขาด " + nRisk + " รายการ" : "✅ ไม่มีรายการเสี่ยง",
      riskColor: nRisk > 0 ? "attention" : "good",
      countLabel: items.length === nAll
        ? nAll + " รายการ"
        : "แสดง " + items.length + " จาก " + nAll + " รายการ",
      items: items,
    };
  });

  const isMA = dataSet === "MA";
  const meta: CardMeta = {
    title: "Report Material Optimize&MA_NER",
    subtitle: "",                          // ตัดบรรทัดคำอธิบายรหัสวัสดุออกแล้ว
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
  return { meta: meta, zones: zones, items: shown };
}


// ---------------------------------------------------------------------------
// จุดเริ่มที่ Power Automate เรียก
// ---------------------------------------------------------------------------
function main(
  workbook: ExcelScript.Workbook,
  period: string = "",
  topRows: number = 12,
  dashboardUrl: string = "",
  sourceFileUrl: string = ""
): CardPayload {
  return buildPayload(workbook, period, topRows, dashboardUrl, sourceFileUrl);
}
