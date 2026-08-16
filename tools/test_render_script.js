#!/usr/bin/env node
/**
 * test_render_script.js
 * =====================
 * ตรวจข้อที่สำคัญที่สุดของโปรเจกต์นี้:
 *
 *   การ์ดที่ office-scripts/renderAdaptiveCard.ts สร้างขึ้น
 *   ต้องเหมือนกับการ์ดที่ไลบรารี adaptivecards-templating ตัวจริง expand ออกมา
 *   ทุก byte
 *
 * ถ้าผ่าน แปลว่าสิ่งที่เห็นใน adaptivecards.io/designer คือสิ่งที่ Teams จะได้จริง
 *
 *   python3 tools/dump_sheet_values.py <xlsm>
 *   node tools/build_office_script.js
 *   node tools/test_render_script.js
 */

const fs = require("fs");
const path = require("path");
const ts = require("typescript");
const { Template } = require("adaptivecards-templating");

const ROOT = path.resolve(__dirname, "..");
const VALUES = path.join(__dirname, ".sheet-values.json");

if (!fs.existsSync(VALUES)) {
  console.error("รัน: python3 tools/dump_sheet_values.py <xlsm> ก่อน");
  process.exit(1);
}

function load(tsFile) {
  const src = fs.readFileSync(tsFile, "utf8");
  const js = ts.transpileModule(src, {
    compilerOptions: { target: ts.ScriptTarget.ES2018, module: ts.ModuleKind.None },
  }).outputText;
  return new Function(js + "\nreturn { main: main };")();
}

const sheetValues = JSON.parse(fs.readFileSync(VALUES, "utf8"));
function makeWorkbook() {
  const range = {
    getValues: () => sheetValues.values,
    getRowIndex: () => sheetValues.firstRowIndex,
  };
  const sheet = { getUsedRange: () => range };
  return { getWorksheet: (n) => (n === sheetValues.sheet ? sheet : undefined) };
}

let failures = 0;
const pass = (m) => console.log("  ✅ " + m);
const fail = (m) => { failures++; console.log("  ❌ " + m); };

const renderScript = load(path.join(ROOT, "office-scripts", "renderAdaptiveCard.ts"));
const payloadScript = load(path.join(ROOT, "office-scripts", "buildCardPayload.ts"));
const templateJson = JSON.parse(
  fs.readFileSync(path.join(ROOT, "adaptive-cards", "01_card-main.json"), "utf8")
);

const GEN_AT = "16 ส.ค. 2569 08:00 น.";
const DASH = "https://app.powerbi.com/groups/me/reports/abc123";
const FILE = "https://contoso.sharepoint.com/sites/NER/Shared%20Documents/Optimize.xlsx";

// เคสทดสอบ: รอบปกติ, รอบที่มีรายการเสี่ยง, รอบที่ไม่มีข้อมูล (empty state)
const CASES = [
  { name: "รอบล่าสุด (MA 17 Aug 26)", period: "", top: 12 },
  { name: "รอบที่มีรายการเสี่ยงขาด (MA 10 Aug 26)", period: "MA (10 Aug 26)", top: 12 },
  { name: "รอบ Optimize", period: "Optimize(3 Aug 26)", top: 12 },
  { name: "รอบที่ไม่มีข้อมูล -> empty state", period: "ไม่มีรอบนี้จริง", top: 12 },
  // ขอ 20 แถว แต่สคริปต์ต้องบีบลงเหลือ TOP_ROWS_MAX (12) เอง ไม่ให้การ์ดชนเพดาน
  { name: "ขอเกินเพดาน (top=20) ต้องถูกบีบเหลือ 10", period: "Optimize(3 Aug 26)", top: 20 },
];

for (const c of CASES) {
  console.log("\n▶ " + c.name);

  // ก. การ์ดจาก Office Script (ตัว expand ที่เขียนเอง)
  let mine;
  try {
    mine = renderScript.main(makeWorkbook(), c.period, c.top, DASH, FILE, GEN_AT);
    pass("renderAdaptiveCard.ts รันผ่าน");
  } catch (e) {
    fail("renderAdaptiveCard.ts error: " + e.message);
    continue;
  }

  // ข. การ์ดจากไลบรารีตัวจริง โดยใช้ payload ชุดเดียวกัน
  const payload = payloadScript.main(makeWorkbook(), c.period, c.top, DASH, FILE);
  payload.meta.generatedAt = GEN_AT;
  const theirs = new Template(templateJson).expand({ $root: payload });

  const a = JSON.stringify(mine);
  const b = JSON.stringify(theirs);

  if (a === b) {
    pass(`ตรงกับ adaptivecards-templating ทุก byte (${(Buffer.byteLength(a) / 1024).toFixed(1)} KB)`);
  } else {
    fail("ผลลัพธ์ไม่ตรงกับไลบรารี");
    // ชี้จุดต่างจุดแรกให้ดู
    let i = 0;
    while (i < a.length && i < b.length && a[i] === b[i]) i++;
    console.log("     ตำแหน่งที่ต่าง: " + i);
    console.log("     ของเรา   : ..." + a.substring(Math.max(0, i - 60), i + 90));
    console.log("     ไลบรารี  : ..." + b.substring(Math.max(0, i - 60), i + 90));
  }

  // ตรวจว่าเป็นการ์ดที่ส่งเข้า Teams ได้จริง
  if (mine.type === "AdaptiveCard" && mine.version) pass(`เป็น AdaptiveCard v${mine.version}`);
  else fail("ผลลัพธ์ไม่ใช่ AdaptiveCard");

  if (/\$\{[^}]*\}/.test(a)) fail("มี ${...} ค้างในการ์ด");
  else pass("ไม่มี binding ค้าง");

  const bytes = Buffer.byteLength(a, "utf8");
  if (bytes > 28 * 1024) fail(`ขนาด ${(bytes / 1024).toFixed(1)} KB เกินเพดาน Teams 28 KB`);
  else pass(`ขนาด ${(bytes / 1024).toFixed(1)} KB`);

  const isEmptyCase = c.period === "ไม่มีรอบนี้จริง";
  if (isEmptyCase && a.includes("ไม่มีรายการวัสดุในรอบนี้")) pass("แสดง empty state ถูกต้อง");
  else if (isEmptyCase) fail("ควรแสดง empty state");
}

console.log("\n" + "=".repeat(60));
if (failures === 0) {
  console.log("✅ Office Script สร้างการ์ดได้ตรงกับ adaptivecards-templating ทุกเคส");
  process.exit(0);
}
console.log(`❌ ไม่ผ่าน ${failures} รายการ`);
process.exit(1);
