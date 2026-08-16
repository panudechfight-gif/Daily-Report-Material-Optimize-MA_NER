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
/*__PAYLOAD_LOGIC__*/

// ===========================================================================
// ส่วนที่ 2 — template ของการ์ด (ฝังมาจาก adaptive-cards/01_card-main.json)
// ===========================================================================
const CARD_TEMPLATE: object = /*__CARD_TEMPLATE__*/;

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
