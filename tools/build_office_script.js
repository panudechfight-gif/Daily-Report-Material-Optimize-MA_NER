#!/usr/bin/env node
/**
 * build_office_script.js
 * ======================
 * ประกอบไฟล์ office-scripts/renderAdaptiveCard.ts จาก 2 แหล่ง:
 *
 *   1. office-scripts/buildCardPayload.ts        (ตรรกะเตรียมข้อมูล)
 *   2. adaptive-cards/01_card-main.json          (หน้าตาการ์ด)
 *
 * เข้าไปแทนที่ marker ใน office-scripts/_renderAdaptiveCard.skeleton.ts
 *
 * ทำแบบนี้เพื่อให้ template การ์ดมี "แหล่งความจริงเดียว" คือไฟล์ JSON
 * แก้การ์ดที่ไฟล์ JSON แล้วรันคำสั่งนี้ ไฟล์ Office Script จะอัปเดตตาม
 *
 *   node tools/build_office_script.js
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const SKELETON = path.join(ROOT, "office-scripts", "_renderAdaptiveCard.skeleton.ts");
const PAYLOAD = path.join(ROOT, "office-scripts", "buildCardPayload.ts");
const CARD = path.join(ROOT, "adaptive-cards", "01_card-main.json");
const OUT = path.join(ROOT, "office-scripts", "renderAdaptiveCard.ts");

// ---- 1. ดึงตรรกะเตรียมข้อมูล: ทุกอย่างยกเว้นหัวคอมเมนต์และฟังก์ชัน main() ----
let payloadSrc = fs.readFileSync(PAYLOAD, "utf8");

const mainMarker = "// จุดเริ่มที่ Power Automate เรียก";
const cut = payloadSrc.indexOf(mainMarker);
if (cut < 0) {
  console.error("หา main() ใน buildCardPayload.ts ไม่เจอ — marker เปลี่ยนไปหรือเปล่า");
  process.exit(1);
}
payloadSrc = payloadSrc.slice(0, cut).replace(/\/\/ -+\s*$/, "").trimEnd();

// ตัดคอมเมนต์หัวไฟล์ (block comment ก้อนแรก) ออก เพราะ skeleton มีหัวของตัวเองแล้ว
payloadSrc = payloadSrc.replace(/^\/\*\*[\s\S]*?\*\/\s*/, "");

// ---- 2. อ่าน template การ์ด ----
const cardJson = JSON.parse(fs.readFileSync(CARD, "utf8"));
const cardLiteral = JSON.stringify(cardJson, null, 2)
  .split("\n")
  .map((l, i) => (i === 0 ? l : "" + l))
  .join("\n");

// ---- 3. ประกอบร่าง ----
let out = fs.readFileSync(SKELETON, "utf8");
out = out.replace("/*__PAYLOAD_LOGIC__*/", payloadSrc);
out = out.replace("/*__CARD_TEMPLATE__*/", cardLiteral);

out =
  "// ***  ไฟล์นี้ถูกสร้างอัตโนมัติโดย tools/build_office_script.js  ***\n" +
  "// ***  ห้ามแก้ด้วยมือ — แก้ที่ adaptive-cards/01_card-main.json  ***\n" +
  "//     สร้างเมื่อ: (รัน node tools/build_office_script.js เพื่ออัปเดต)\n\n" +
  out;

fs.writeFileSync(OUT, out, "utf8");

const kb = (Buffer.byteLength(out, "utf8") / 1024).toFixed(1);
console.log(`wrote ${path.relative(ROOT, OUT)}  (${kb} KB)`);
console.log("คัดลอกเนื้อไฟล์นี้ไปวางใน Excel -> Automate -> New Script");
