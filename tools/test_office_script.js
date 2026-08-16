#!/usr/bin/env node
/**
 * test_office_script.js
 * =====================
 * รัน office-scripts/buildCardPayload.ts นอก Excel เพื่อพิสูจน์ว่าตรรกะถูกต้อง
 * โดยจำลอง ExcelScript.Workbook จากค่าจริงในไฟล์ .xlsm
 *
 * ขั้นตอน:
 *   1. python3 tools/dump_sheet_values.py <xlsm>   -> tools/.sheet-values.json
 *   2. node tools/test_office_script.js
 *
 * เทียบผลลัพธ์กับ adaptive-cards/data/sample-data-main.json (สร้างจาก Python)
 * ถ้าตรงกัน แปลว่า Office Script กับสคริปต์อ้างอิงให้ผลเหมือนกันจริง
 */

const fs = require("fs");
const path = require("path");
const ts = require("typescript");

const ROOT = path.resolve(__dirname, "..");
const VALUES = path.join(__dirname, ".sheet-values.json");

if (!fs.existsSync(VALUES)) {
  console.error("ไม่พบ " + VALUES + "\nรัน: python3 tools/dump_sheet_values.py <xlsm> ก่อน");
  process.exit(1);
}

// ---- transpile TypeScript -> JavaScript แล้วดึงฟังก์ชัน main ออกมา ----
const src = fs.readFileSync(path.join(ROOT, "office-scripts", "buildCardPayload.ts"), "utf8");
const js = ts.transpileModule(src, {
  compilerOptions: { target: ts.ScriptTarget.ES2018, module: ts.ModuleKind.None },
}).outputText;

const factory = new Function(js + "\nreturn main;");
const main = factory();

// ---- จำลอง ExcelScript.Workbook เท่าที่สคริปต์ใช้จริง ----
const sheetValues = JSON.parse(fs.readFileSync(VALUES, "utf8"));

function makeWorkbook(name, values, firstRowIndex) {
  const range = {
    getValues: () => values,
    getRowIndex: () => firstRowIndex,
  };
  const sheet = { getUsedRange: () => range };
  return { getWorksheet: (n) => (n === name ? sheet : undefined) };
}

const wb = makeWorkbook(sheetValues.sheet, sheetValues.values, sheetValues.firstRowIndex);

let failures = 0;
const pass = (m) => console.log("  ✅ " + m);
const fail = (m) => { failures++; console.log("  ❌ " + m); };

console.log("▶ รัน buildCardPayload (จำลอง ExcelScript)");
let out;
try {
  out = main(wb, "", 12, "https://dash.example/x", "https://sp.example/f.xlsx");
  pass("สคริปต์รันผ่าน ไม่มี exception");
} catch (e) {
  fail("สคริปต์ error: " + e.message);
  process.exit(1);
}

// ---- เทียบกับผลลัพธ์อ้างอิงจาก Python ----
const ref = JSON.parse(
  fs.readFileSync(path.join(ROOT, "adaptive-cards", "data", "sample-data-main.json"), "utf8")
);

const metaKeys = ["period", "prevPeriod", "dataSet", "totalItems", "shownItems",
                  "totalZones", "riskCount", "watchCount", "okCount",
                  "hasData", "hasMore", "moreCount"];
for (const k of metaKeys) {
  if (JSON.stringify(out.meta[k]) === JSON.stringify(ref.meta[k])) {
    pass(`meta.${k} = ${JSON.stringify(out.meta[k])}`);
  } else {
    fail(`meta.${k}: Office Script = ${JSON.stringify(out.meta[k])} แต่ Python = ${JSON.stringify(ref.meta[k])}`);
  }
}

// เทียบรายการทีละแถว (รหัส + ตัวเลข 4 ช่อง + สถานะ)
const flat = (p) => (p.zones || []).reduce((a, z) => a.concat(z.items), []);
const a = flat(out), b = flat(ref);
if (a.length !== b.length) {
  fail(`จำนวนแถวไม่ตรง: ${a.length} vs ${b.length}`);
} else {
  const fields = ["itemCode", "zone", "province", "onhand", "fromHub",
                  "prevBefore", "prevReceived", "statusShort", "statusColor", "isRisk"];
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    for (const f of fields) {
      if (JSON.stringify(a[i][f]) !== JSON.stringify(b[i][f])) {
        if (diff < 5) fail(`แถว ${i + 1} field ${f}: ${JSON.stringify(a[i][f])} vs ${JSON.stringify(b[i][f])}`);
        diff++;
      }
    }
  }
  if (diff === 0) pass(`ข้อมูลทั้ง ${a.length} แถวตรงกับผลลัพธ์อ้างอิงทุก field`);
  else fail(`พบความต่างรวม ${diff} จุด`);
}

// ---- payload จาก Office Script ต้อง expand เข้าการ์ดได้จริง ----
const { Template } = require("adaptivecards-templating");
out.meta.generatedAt = "16 ส.ค. 2569 08:00 น.";
for (const cardFile of ["01_card-main.json", "02_card-simple-flat.json", "04_card-mobile-stacked.json"]) {
  try {
    const tpl = JSON.parse(fs.readFileSync(path.join(ROOT, "adaptive-cards", cardFile), "utf8"));
    const expanded = JSON.stringify(new Template(tpl).expand({ $root: out }));
    if (/\$\{[^}]*\}/.test(expanded)) fail(`${cardFile}: มี binding ค้าง`);
    else pass(`${cardFile}: expand ผ่าน (${(Buffer.byteLength(expanded) / 1024).toFixed(1)} KB)`);
  } catch (e) {
    fail(`${cardFile}: ${e.message}`);
  }
}

console.log("\n" + "=".repeat(60));
if (failures === 0) {
  console.log("✅ Office Script ให้ผลตรงกับสคริปต์อ้างอิง และ payload ใช้กับการ์ดได้จริง");
  process.exit(0);
}
console.log(`❌ ไม่ผ่าน ${failures} รายการ`);
process.exit(1);
