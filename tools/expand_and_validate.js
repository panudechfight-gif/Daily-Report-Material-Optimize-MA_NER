#!/usr/bin/env node
/**
 * expand_and_validate.js
 * ======================
 * ทดสอบว่า Adaptive Card template + data payload "expand ได้จริง"
 * ด้วยไลบรารีตัวจริงที่ adaptivecards.io/designer ใช้ คือ adaptivecards-templating
 *
 *   npm i adaptivecards-templating
 *   node tools/expand_and_validate.js
 *
 * สิ่งที่ตรวจ:
 *   1. template + data expand ผ่านโดยไม่มี error
 *   2. ไม่มี ${...} ที่ผูกข้อมูลไม่ติดหลงเหลืออยู่ในผลลัพธ์
 *   3. ขนาดการ์ดหลัง expand ไม่เกิน 28 KB (เพดานของ Microsoft Teams)
 *   4. $when ทำงาน (การ์ด empty state ต้องไม่มีตารางข้อมูล / การ์ดปกติต้องไม่มี empty state)
 *   5. ค่าที่ต้องอยู่บนการ์ดจริง ๆ (รอบ, Zone, Item Code) ถูกแทนที่แล้ว
 */

const fs = require("fs");
const path = require("path");
const { Template } = require("adaptivecards-templating");

const ROOT = path.resolve(__dirname, "..");
const CARDS = path.join(ROOT, "adaptive-cards");
const DATA = path.join(CARDS, "data");
const OUT = path.join(CARDS, "expanded");

const TEAMS_LIMIT = 28 * 1024;   // เพดานจริงของ Microsoft Teams
const SAFE_LIMIT = 24 * 1024;    // เกินนี้ถือว่าเสี่ยง — ควรลดจำนวนแถว

const CASES = [
  { card: "01_card-main.json", data: "sample-data-main.json", expectEmpty: false },
  { card: "01_card-main.json", data: "sample-data-risk.json", expectEmpty: false, out: "01_card-main.risk.json" },
  { card: "01_card-main.json", data: "sample-data-empty.json", expectEmpty: true, out: "01_card-main.empty.json" },
  { card: "02_card-simple-flat.json", data: "sample-data-flat.json", expectEmpty: false },
  { card: "03_card-empty-state.json", data: "sample-data-empty.json", expectEmpty: true },
  { card: "04_card-mobile-stacked.json", data: "sample-data-risk.json", expectEmpty: false },
];

let failures = 0;
const pass = (m) => console.log("  ✅ " + m);
const fail = (m) => { failures++; console.log("  ❌ " + m); };

function walk(node, visit) {
  if (Array.isArray(node)) return node.forEach((n) => walk(n, visit));
  if (node && typeof node === "object") {
    visit(node);
    Object.values(node).forEach((v) => walk(v, visit));
  }
}

function countType(card, type) {
  let n = 0;
  walk(card, (o) => { if (o.type === type) n++; });
  return n;
}

fs.mkdirSync(OUT, { recursive: true });

for (const c of CASES) {
  const cardPath = path.join(CARDS, c.card);
  const dataPath = path.join(DATA, c.data);
  const label = `${c.card}  <-  ${c.data}`;
  console.log("\n▶ " + label);

  if (!fs.existsSync(cardPath)) { fail("ไม่พบไฟล์ template: " + c.card); continue; }
  if (!fs.existsSync(dataPath)) { fail("ไม่พบไฟล์ data: " + c.data); continue; }

  const templateJson = JSON.parse(fs.readFileSync(cardPath, "utf8"));
  const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));

  // ค่าที่ Power Automate จะเติมตอนรันจริง — ใส่ให้ครบก่อนทดสอบ
  data.meta.generatedAt = data.meta.generatedAt || "16 ส.ค. 2569 08:00 น.";
  data.meta.dashboardUrl = "https://app.powerbi.com/groups/me/reports/demo";
  data.meta.sourceFileUrl = "https://contoso.sharepoint.com/sites/NER/Shared%20Documents/Optimize_HiringMA.xlsx";

  let expanded;
  try {
    expanded = new Template(templateJson).expand({ $root: data });
    pass("expand ผ่าน (adaptivecards-templating)");
  } catch (e) {
    fail("expand ล้มเหลว: " + e.message);
    continue;
  }

  const text = JSON.stringify(expanded);

  // 2. ไม่มี binding ค้าง
  const leftovers = text.match(/\$\{[^}]*\}/g);
  if (leftovers) fail("มี binding ที่ยังไม่ถูกแทนค่า: " + [...new Set(leftovers)].slice(0, 5).join(", "));
  else pass("ไม่มี ${...} ค้างอยู่ในผลลัพธ์");

  // ค่าที่แทนไม่ติดจะกลายเป็น "undefined" หรือ null ใน text
  if (/"text"\s*:\s*"undefined"/.test(text)) fail('พบ TextBlock ที่มีค่า "undefined"');
  else pass('ไม่มี TextBlock ที่เป็น "undefined"');

  // 3. ขนาด
  const bytes = Buffer.byteLength(text, "utf8");
  const kb = (bytes / 1024).toFixed(1);
  if (bytes > TEAMS_LIMIT) fail(`ขนาดการ์ด ${kb} KB เกินเพดาน Teams 28 KB — ลดจำนวนแถว (--top)`);
  else if (bytes > SAFE_LIMIT) console.log(`  ⚠️  ขนาดการ์ด ${kb} KB ใกล้เพดาน 28 KB — ควรลด --top`);
  else pass(`ขนาดการ์ด ${kb} KB (เพดาน Teams 28 KB)`);

  // 4. $when
  const hasEmptyMarker = text.includes("ไม่มีรายการวัสดุในรอบนี้");

  // นับแถวจาก "รหัสวัสดุที่โผล่จริงในผลลัพธ์" ไม่ผูกกับชนิด element
  // (การ์ดแต่ละใบวางแถวคนละแบบ: ColumnSet ต่อแถว, คอลัมน์ละหลาย TextBlock, หรือบรรทัดเดียวจบ)
  const rowsExpanded = (() => {
    const codes = new Set((data.items || []).map(i => i.itemCode));
    for (const z of data.zones || []) for (const i of z.items || []) codes.add(i.itemCode);
    let n = 0;
    for (const code of codes) if (text.includes(code)) n++;
    return n;
  })();

  if (c.expectEmpty) {
    if (hasEmptyMarker) pass("$when แสดง empty state ถูกต้อง");
    else fail("ควรแสดง empty state แต่ไม่พบ");
    if (!text.includes("Zone RC")) pass("$when ซ่อนตารางข้อมูลแล้ว");
    else fail("empty state แต่ยังมีตาราง Zone หลงเหลือ");
  } else {
    if (!hasEmptyMarker) pass("$when ซ่อน empty state ถูกต้อง");
    else fail("มีข้อมูลแต่ยังแสดง empty state");
    if (rowsExpanded > 0) pass(`ขยายแถวข้อมูลแล้ว (รหัสวัสดุ ${rowsExpanded} รหัส)`);
    else fail("ไม่พบแถวข้อมูลหลัง expand");
  }

  // 5. ค่าที่ต้องปรากฏจริง
  if (!c.expectEmpty) {
    const checks = [
      [data.meta.period, "Distribution period"],
      [data.meta.prevPeriod, "Prev_Period"],
      [data.meta.dataSet, "Data_Set"],
    ];
    const firstItem = data.zones ? data.zones[0].items[0] : data.items[0];
    checks.push([firstItem.itemCode, "Item Code"]);
    checks.push([firstItem.zone, "Zone"]);
    for (const [val, name] of checks) {
      if (val && text.includes(val)) pass(`พบค่า ${name}: ${val}`);
      else fail(`ไม่พบค่า ${name}: ${val}`);
    }
  }

  // ตรวจ Action.OpenUrl 2 ปุ่มตามสเปก
  const urls = [];
  walk(expanded, (o) => { if (o.type === "Action.OpenUrl") urls.push(o.title + " -> " + o.url); });
  if (urls.length === 2 && urls.some((u) => u.includes("Dashboard")) && urls.some((u) => u.includes("source file"))) {
    pass("มีปุ่ม Action.OpenUrl ครบ 2 ปุ่ม");
  } else {
    fail("ปุ่ม Action.OpenUrl ไม่ครบ: " + JSON.stringify(urls));
  }
  if (urls.some((u) => u.includes("${") || u.includes("CHANGE-ME"))) {
    fail("URL ของปุ่มยังไม่ถูกแทนค่า");
  }

  const outName = c.out || c.card.replace(".json", ".expanded.json");
  fs.writeFileSync(path.join(OUT, outName), JSON.stringify(expanded, null, 2), "utf8");
  console.log("  → " + path.relative(ROOT, path.join(OUT, outName)));
}

console.log("\n" + "=".repeat(60));
if (failures === 0) {
  console.log("✅ ผ่านทั้งหมด — template expand ได้จริงทุกเคส");
  process.exit(0);
} else {
  console.log(`❌ ไม่ผ่าน ${failures} รายการ`);
  process.exit(1);
}
