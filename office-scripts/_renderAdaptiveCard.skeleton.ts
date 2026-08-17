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
//     26,859  การ์ดหน้าตาเดิม (5 คอลัมน์) -> Teams ปฏิเสธจริง ❌
//     23,538  การ์ดหน้าตาปัจจุบัน รอบ Optimize 35 แถวครบทุกรหัส ✅
//     18,692  รอบ MA 24 แถวครบทุกรหัส ✅
//
// ตั้งไว้ 24,500 เพื่อให้ "ทุกรหัสลงการ์ดครบ" ตามที่ต้องการ โดยยังต่ำกว่า
// ขนาดที่เคยถูกปฏิเสธจริงอยู่ราว 2,400 ไบต์ และต่ำกว่าเพดาน 28 KB ราว 4 KB
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
