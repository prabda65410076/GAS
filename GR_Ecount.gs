/**
 * GR Ecount — ตารางรายวัน 3 โหมด แยกกันคนละชีต + แยก Normal / SN
 *
 *   จำนวนชิ้น         -> ชีต 'Matchine Use'
 *   ชั่วโมงเดินเครื่อง  -> ชีต 'ชม.เดินเครื่อง'
 *   อัตราใช้งาน %      -> ชีต 'อัตราใช้งาน'
 *
 * แต่ละเครื่องจะมี 2 แถว: Normal และ SN
 *   - SN     = แถวในชีต GR CT FM BD ที่คอลัมน์ S มีข้อความ (หาเจอใน SN LIST)
 *   - Normal = คอลัมน์ S ว่าง หรือขึ้น "ไม่พบ"
 *
 * คอลัมน์ Q-U ของชีต GR CT FM BD คำนวณด้วยสคริปต์แทนสูตร (เร็วกว่าและไม่ทำให้ไฟล์หน่วง)
 *   Q = MID(A,6,2)        เดือน
 *   R = MID(A,9,2)        วัน
 *   S = XLOOKUP(D, 'SN LIST'!A:A, 'SN LIST'!B:B, "ไม่พบ")
 *   T = VLOOKUP(L, 'Mc List'!D:E, 2) ถ้าไม่เจอ -> VLOOKUP(L, 'Mc List'!E:F, 2) ถ้าไม่เจอ -> "ไม่พบ"
 *   U = LEFT(A,10)        วันที่ (yyyy-mm-dd)
 * เมนู  เครื่องมือ GR > คำนวณคอลัมน์ Q-U  (และจะคำนวณให้อัตโนมัติก่อนสร้างตาราง ถ้า AUTO_FILL_HELPERS = true)
 *
 * รายชื่อเครื่องเก็บไว้ในชีตซ่อน 'รายชื่อเครื่อง' (สร้างอัตโนมัติครั้งแรก
 * โดยคัดลอก A1:F90 จาก Matchine Use) เพื่อไม่ให้รายชื่อเสียหายเมื่อชีตถูกเขียนทับ
 * ถ้าจะเพิ่ม/แก้เครื่อง ใช้เมนู  เครื่องมือ GR > แก้ไขรายชื่อเครื่อง
 *
 * โครงคอลัมน์ผลลัพธ์:  A-F = ข้อมูลเครื่อง,  G = ประเภท (Normal/SN),
 *                      H-AL = วันที่ 1-31,  AM = รวม/เฉลี่ย
 *
 * ติดตั้ง: ส่วนขยาย > Apps Script > ลบโค้ดเดิม > วางไฟล์นี้ > บันทึก > รีเฟรชชีต
 */

// ===================== ตั้งค่า =====================
const SCRIPT_VERSION = 'v4';   // ใช้เช็กว่าชีตกำลังรันโค้ดเวอร์ชันนี้อยู่

const CFG = {
  SRC_SHEET:    'GR CT FM BD',     // ชีตข้อมูลดิบ
  TPL_SHEET:    'Matchine Use',    // ชีตที่ใช้คัดลอกรายชื่อเครื่องครั้งแรก
  MASTER_SHEET: 'รายชื่อเครื่อง',    // ชีตซ่อน เก็บรายชื่อเครื่องถาวร

  OUT_SHEET: {
    qty:   'Matchine Use',
    hours: 'ชม.เดินเครื่อง',
    util:  'อัตราใช้งาน',
  },

  SRC_START_ROW: 2,
  SRC_END_ROW:   200000,      // 0 = ถึงแถวสุดท้ายที่มีข้อมูล

  SRC_START_TIME_COL: 2,      // B = start Time
  SRC_END_TIME_COL:   3,      // C = end Time
  SRC_QTY_COL:        6,      // F = qty
  SRC_KEY_COL:       12,      // L = m/C
  SRC_SN_COL:        19,      // S = ผลจาก SN LIST
  SRC_DATE_COL:      21,      // U = date-No.

  CHUNK_ROWS: 50000,

  // ---- คอลัมน์ช่วย Q-U (แทนสูตร) ----
  AUTO_FILL_HELPERS: true,    // true = คำนวณ Q-U ใหม่ทุกครั้งก่อนสร้างตาราง
  SRC_DOC_COL:   1,           // A = ข้อความวันที่/เวลา (ใช้กับ Q, R, U)
  SRC_PART_COL:  4,           // D = รหัสที่ใช้หาใน SN LIST
  HELPER_FIRST_COL: 17,       // Q (เขียน Q, R, S, T, U ต่อกัน 5 คอลัมน์)
  SN_SHEET: 'SN LIST',        // A = รหัส, B = ผลลัพธ์
  MC_SHEET: 'Mc List',        // หา L ใน D -> คืน E, ถ้าไม่เจอหาใน E -> คืน F
  NOT_FOUND_TEXT: 'ไม่พบ',

  MONTH_CELL: 'B1',
  YEAR_CELL:  'C1',

  // ---- แยก Normal / SN ----
  SPLIT_SN: true,
  TYPE_LABELS: ['Normal', 'SN'],
  SN_NOT_FOUND: ['ไม่พบ', '#N/A', '#REF!', '#VALUE!', '#ERROR!'],  // ถือเป็น Normal
  SN_ROW_COLOR: '#fff2cc',    // สีพื้นแถว SN ('' = ไม่ใส่สี)

  // ---- โครงตาราง ----
  TPL_FIRST_ROW: 3,
  TPL_LAST_ROW:  90,          // ใช้ตอนคัดลอกรายชื่อครั้งแรกเท่านั้น
  INFO_FIRST_COL: 1,          // A
  INFO_LAST_COL:  6,          // F
  KEY_COL:        6,          // F = รหัสที่จับคู่กับ L
  DAY_ROW: 2,

  ADD_TOTAL: true,
  BLANK_AS: '',
  COLOR_BLOCKS: true,

  // ---- เฉพาะโหมด hours / util ----
  SHIFT_HOURS: 8,
  OVERNIGHT_OK: true,
  MAX_SPAN_HOURS: 24,
  UTIL_TOTAL_BASIS: 'active_days',   // หรือ 'all_days'

  FMT_QTY:   '#,##0',
  FMT_HOURS: '0.00',
  FMT_UTIL:  '0.0%',
};
// ==================================================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('เครื่องมือ GR ' + SCRIPT_VERSION)
    .addItem('จำนวนชิ้น — เดือนเดียว', 'runMonthQty')
    .addItem('จำนวนชิ้น — ทั้งปี', 'runYearQty')
    .addSeparator()
    .addItem('ชั่วโมงเดินเครื่อง — เดือนเดียว', 'runMonthHours')
    .addItem('ชั่วโมงเดินเครื่อง — ทั้งปี', 'runYearHours')
    .addSeparator()
    .addItem('อัตราใช้งาน % — เดือนเดียว', 'runMonthUtil')
    .addItem('อัตราใช้งาน % — ทั้งปี', 'runYearUtil')
    .addSeparator()
    .addItem('คำนวณคอลัมน์ Q-U (แทนสูตร)', 'fillHelperCols')
    .addSeparator()
    .addItem('สร้างชีตปลายทาง', 'setupSheets')
    .addItem('แก้ไขรายชื่อเครื่อง', 'editMaster')
    .addItem('ล้างแถว "เดือน..." ออกจากรายชื่อเครื่อง', 'cleanMaster')
    .addItem('ตรวจการอ่านเวลา', 'debugTimes')
    .addItem('ตรวจรหัสที่หาไม่เจอ', 'debugMissingKeys')
    .addToUi();
}


/* ============ ฟังก์ชันที่เรียกจากเมนู ============ */

const ALL_MONTHS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

function runMonthQty()   { runOne_('qty');   }
function runYearQty()    { runAll_('qty');   }
function runMonthHours() { runOne_('hours'); }
function runYearHours()  { runAll_('hours'); }
function runMonthUtil()  { runOne_('util');  }
function runYearUtil()   { runAll_('util');  }

function runOne_(mode) {
  const ctx = openSheets_(mode);
  const my = getMonthYear_(ctx.dst);
  if (CFG.AUTO_FILL_HELPERS) fillHelperCols_(ctx.ss, ctx.src);
  render_(ctx, my.year, [my.month], mode);
  ctx.dst.activate();
}

function runAll_(mode) {
  const ctx = openSheets_(mode);
  const year = getYear_(ctx.dst);
  if (CFG.AUTO_FILL_HELPERS) fillHelperCols_(ctx.ss, ctx.src);
  render_(ctx, year, ALL_MONTHS, mode);
  ctx.dst.activate();
}

/** เมนู: คำนวณคอลัมน์ Q-U ของชีตข้อมูลดิบ */
function fillHelperCols() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CFG.SRC_SHEET);
  if (!src) throw new Error('ไม่พบชีตข้อมูลดิบ: ' + CFG.SRC_SHEET);
  const msg = fillHelperCols_(ss, src);
  SpreadsheetApp.getUi().alert(msg);
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const master = getMaster_(ss);
  const made = [];
  ['qty', 'hours', 'util'].forEach(function (mode) {
    const name = CFG.OUT_SHEET[mode];
    const before = !!ss.getSheetByName(name);
    ensureSheet_(ss, name, master);
    made.push(name + (before ? ' (มีอยู่แล้ว)' : ' (สร้างใหม่)'));
  });
  SpreadsheetApp.getUi().alert('เรียบร้อย:\n' + made.join('\n'));
}

/** เปิดชีตรายชื่อเครื่องขึ้นมาให้แก้ไข */
function editMaster() {
  const sh = getMaster_(SpreadsheetApp.getActive());
  sh.showSheet();
  sh.activate();
  SpreadsheetApp.getUi().alert(
    'แก้ไขรายชื่อเครื่องได้ตั้งแต่แถว ' + CFG.TPL_FIRST_ROW + ' ลงไป (คอลัมน์ A-F)\n' +
    'คอลัมน์ F ต้องเป็นรหัสเครื่องที่ตรงกับคอลัมน์ L ของ ' + CFG.SRC_SHEET + '\n\n' +
    'แก้เสร็จแล้ว คลิกขวาที่แท็บชีตแล้วเลือก "ซ่อนชีต" ได้');
}


/** ลบแถวหัวเดือน ("เดือน 1 / 2026 ...") ที่หลุดเข้ามาในชีตรายชื่อเครื่อง */
function cleanMaster() {
  const ss = SpreadsheetApp.getActive();
  const m = getMaster_(ss);
  const first = CFG.TPL_FIRST_ROW;
  const last = m.getLastRow();
  if (last < first) return;

  const nInfo = CFG.INFO_LAST_COL - CFG.INFO_FIRST_COL + 1;
  const vals = m.getRange(first, CFG.INFO_FIRST_COL, last - first + 1, nInfo).getValues();
  let removed = 0;
  // ลบจากล่างขึ้นบน เลขแถวจะได้ไม่เลื่อน
  for (let i = vals.length - 1; i >= 0; i--) {
    if (isBandRow_(vals[i])) { m.deleteRow(first + i); removed++; }
  }

  m.showSheet();
  m.activate();
  SpreadsheetApp.getUi().alert(
    'ลบแถว "เดือน..." ออกแล้ว ' + removed + ' แถว\n\n' +
    'โปรดเช็กท้ายรายชื่อว่าเครื่องครบหรือไม่ (ดูเลขลำดับในคอลัมน์ A)\n' +
    'ถ้าขาด ให้พิมพ์เพิ่มต่อท้าย หรือคัดลอกจากประวัติเวอร์ชันของไฟล์');
}

/**
 * แถวนี้เป็นแถบหัวเดือนที่สคริปต์สร้าง ไม่ใช่เครื่องจักร
 * ตรวจทุกช่องในแถว และตัดช่องว่างพิเศษ/อักขระล่องหนออกก่อน
 */
function isBandRow_(row) {
  const txt = row.map(function (v) { return String(v); }).join(' ')
                 .replace(/[  -‍ ﻿]/g, ' ');
  return /เดือน\s*\d{1,2}\s*\/\s*\d{2,4}/.test(txt);
}


/* ============ คอลัมน์ช่วย Q-U (แทนสูตร) ============ */

/**
 * คำนวณ Q-U ของชีตข้อมูลดิบแล้วเขียนเป็นค่าตายตัว (ทับสูตรเดิม)
 * คืนข้อความสรุปผล
 */
function fillHelperCols_(ss, src) {
  const t0 = Date.now();
  const tz = ss.getSpreadsheetTimeZone();
  const NF = CFG.NOT_FOUND_TEXT;
  const startRow = CFG.SRC_START_ROW;
  const nCols = 5;   // Q R S T U

  // ---- หาแถวสุดท้ายที่มีข้อมูลจริง (ดูจากคอลัมน์ A, D, L) ----
  const sheetLast = src.getLastRow();
  if (sheetLast < startRow) return 'ไม่มีข้อมูลในชีต ' + CFG.SRC_SHEET;
  const nAll = sheetLast - startRow + 1;

  ss.toast('กำลังคำนวณคอลัมน์ Q-U...', 'รอสักครู่', 30);

  const colA = readColumn_(src, startRow, nAll, CFG.SRC_DOC_COL);
  const colD = readColumn_(src, startRow, nAll, CFG.SRC_PART_COL);
  const colL = readColumn_(src, startRow, nAll, CFG.SRC_KEY_COL);

  let n = nAll;
  while (n > 0 && isBlank_(colA[n - 1]) && isBlank_(colD[n - 1]) && isBlank_(colL[n - 1])) n--;

  // ---- ตารางค้นหา ----
  const snMap  = buildLookup_(ss, CFG.SN_SHEET, 1, 2);   // SN LIST  A -> B
  const mcMap1 = buildLookup_(ss, CFG.MC_SHEET, 4, 5);   // Mc List  D -> E
  const mcMap2 = buildLookup_(ss, CFG.MC_SHEET, 5, 6);   // Mc List  E -> F

  // ---- คำนวณ ----
  const out = new Array(n);
  let snHit = 0, mcHit = 0;
  for (let i = 0; i < n; i++) {
    const a = docText_(colA[i], tz);

    const q = a.substr(5, 2);       // MID(A,6,2)
    const r = a.substr(8, 2);       // MID(A,9,2)
    const u = a.substr(0, 10);      // LEFT(A,10)

    let s = NF;
    const kD = norm_(colD[i]);
    if (kD !== '' && snMap.has(kD)) { s = snMap.get(kD); snHit++; }

    let t = NF;
    const kL = norm_(colL[i]);
    if (kL !== '') {
      if (mcMap1.has(kL))      { t = mcMap1.get(kL); mcHit++; }
      else if (mcMap2.has(kL)) { t = mcMap2.get(kL); mcHit++; }
    }

    out[i] = [q, r, s, t, u];
  }

  // ---- เขียน (Q, R, U เป็นข้อความเหมือนผลของ MID/LEFT) ----
  const qCol = CFG.HELPER_FIRST_COL;
  if (src.getMaxColumns() < qCol + nCols - 1) {
    src.insertColumnsAfter(src.getMaxColumns(), qCol + nCols - 1 - src.getMaxColumns());
  }
  if (n > 0) {
    src.getRange(startRow, qCol,     n, 2).setNumberFormat('@');   // Q, R
    src.getRange(startRow, qCol + 4, n, 1).setNumberFormat('@');   // U
    let done = 0;
    while (done < n) {
      const take = Math.min(CFG.CHUNK_ROWS, n - done);
      src.getRange(startRow + done, qCol, take, nCols).setValues(out.slice(done, done + take));
      done += take;
    }
  }

  // ล้างสูตร/ค่าเก่าที่ค้างอยู่ใต้ข้อมูล
  const maxRows = src.getMaxRows();
  const tailStart = startRow + n;
  if (maxRows >= tailStart) {
    src.getRange(tailStart, qCol, maxRows - tailStart + 1, nCols).clearContent();
  }
  SpreadsheetApp.flush();

  const msg = 'คำนวณ Q-U แล้ว ' + n.toLocaleString() + ' แถว' +
              ' • เจอใน SN LIST ' + snHit.toLocaleString() +
              ' • เจอใน Mc List ' + mcHit.toLocaleString() +
              ' • ' + ((Date.now() - t0) / 1000).toFixed(1) + ' วิ';
  ss.toast(msg, 'คอลัมน์ Q-U', 10);
  Logger.log(msg);
  return msg;
}

/** สร้าง Map จากชีต: คีย์ในคอลัมน์ keyCol -> ค่าในคอลัมน์ valCol (เจอตัวแรกก่อน เหมือน XLOOKUP/VLOOKUP) */
function buildLookup_(ss, sheetName, keyCol, valCol) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('ไม่พบชีต: ' + sheetName);
  const map = new Map();
  const last = sh.getLastRow();
  if (last < 1) return map;

  const c1 = Math.min(keyCol, valCol);
  const width = Math.abs(valCol - keyCol) + 1;
  const vals = sh.getRange(1, c1, last, width).getValues();
  const ki = keyCol - c1, vi = valCol - c1;
  for (let i = 0; i < vals.length; i++) {
    const k = norm_(vals[i][ki]);
    if (k !== '' && !map.has(k)) map.set(k, vals[i][vi]);
  }
  return map;
}

/** ค่าในคอลัมน์ A เป็นข้อความ (ถ้าชีตแปลงเป็นวันที่ไปแล้ว ให้จัดรูปเป็น yyyy-MM-dd HH:mm:ss) */
function docText_(v, tz) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}

function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}


/* ================== แกนกลาง ================== */

function openSheets_(mode) {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CFG.SRC_SHEET);
  if (!src) throw new Error('ไม่พบชีตข้อมูลดิบ: ' + CFG.SRC_SHEET);

  // ต้องเรียกก่อนเขียนอะไรลงชีต เพื่อเก็บรายชื่อเครื่องไว้ก่อน
  const master = getMaster_(ss);
  const dst = ensureSheet_(ss, CFG.OUT_SHEET[mode], master);

  return { ss: ss, src: src, master: master, dst: dst };
}

/** ชีตรายชื่อเครื่อง ถ้ายังไม่มีจะคัดลอก A1:F90 จากชีตต้นแบบแล้วซ่อนไว้ */
function getMaster_(ss) {
  let m = ss.getSheetByName(CFG.MASTER_SHEET);
  if (m) return m;

  const tpl = ss.getSheetByName(CFG.TPL_SHEET);
  if (!tpl) throw new Error('ไม่พบชีตต้นแบบ: ' + CFG.TPL_SHEET);

  const nInfo = CFG.INFO_LAST_COL - CFG.INFO_FIRST_COL + 1;
  const vals = tpl.getRange(1, CFG.INFO_FIRST_COL, CFG.TPL_LAST_ROW, nInfo).getValues()
    .filter(function (r, i) { return i < CFG.TPL_FIRST_ROW - 1 || !isBandRow_(r); });

  m = ss.insertSheet(CFG.MASTER_SHEET);
  m.getRange(1, CFG.INFO_FIRST_COL, vals.length, nInfo).setValues(vals);
  for (let c = CFG.INFO_FIRST_COL; c <= CFG.INFO_LAST_COL; c++) {
    m.setColumnWidth(c, tpl.getColumnWidth(c));
  }
  tpl.activate();
  m.hideSheet();
  ss.toast('สร้างชีตซ่อน "' + CFG.MASTER_SHEET + '" เก็บรายชื่อเครื่องแล้ว', 'แจ้งเตือน', 8);
  return m;
}

/** หาชีตตามชื่อ ถ้าไม่มีให้สร้าง พร้อมคัดลอกหัวตาราง (แถว 1-2) จากรายชื่อเครื่อง */
function ensureSheet_(ss, name, master) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;

  sh = ss.insertSheet(name);
  const nInfo = CFG.INFO_LAST_COL - CFG.INFO_FIRST_COL + 1;
  const nHead = CFG.TPL_FIRST_ROW - 1;
  const head = master.getRange(1, CFG.INFO_FIRST_COL, nHead, nInfo).getValues();
  sh.getRange(1, CFG.INFO_FIRST_COL, nHead, nInfo).setValues(head);

  const needCols = dayFirstCol_() + 32;
  if (sh.getMaxColumns() < needCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  }
  for (let c = CFG.INFO_FIRST_COL; c <= CFG.INFO_LAST_COL; c++) {
    sh.setColumnWidth(c, master.getColumnWidth(c));
  }
  return sh;
}

function dayFirstCol_() {
  return CFG.INFO_LAST_COL + 1 + (CFG.SPLIT_SN ? 1 : 0);
}

/** true = แถวนี้เป็นงาน SN (คอลัมน์ S มีค่า และไม่ใช่ "ไม่พบ"/ค่า error) */
function isSN_(v) {
  if (v === '' || v === null || v === undefined) return false;
  const s = String(v).trim().toUpperCase();
  if (s === '') return false;
  for (let i = 0; i < CFG.SN_NOT_FOUND.length; i++) {
    if (s === String(CFG.SN_NOT_FOUND[i]).toUpperCase()) return false;
  }
  return true;
}

function parseTime_(v) {
  if (v === '' || v === null || v === undefined) return null;

  if (v instanceof Date && !isNaN(v)) {
    return v.getHours() * 60 + v.getMinutes() + v.getSeconds() / 60;
  }
  if (typeof v === 'number') {
    if (v >= 0 && v < 1) return v * 1440;
    if (v >= 0 && v <= 24) return v * 60;
    return null;
  }

  const s = String(v).trim();
  const m = s.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if (!m) return null;
  const hh = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  const ss = m[3] ? parseInt(m[3], 10) : 0;
  if (hh > 24 || mm > 59) return null;
  return hh * 60 + mm + ss / 60;
}

function spanMinutes_(rawStart, rawEnd) {
  const a = parseTime_(rawStart);
  const b = parseTime_(rawEnd);
  if (a === null || b === null) return null;

  let diff = b - a;
  if (diff < 0) {
    if (!CFG.OVERNIGHT_OK) return null;
    diff += 1440;
  }
  if (diff <= 0 || diff > CFG.MAX_SPAN_HOURS * 60) return null;
  return diff;
}

function render_(ctx, year, months, mode) {
  const t0 = Date.now();
  const ss = ctx.ss, src = ctx.src, master = ctx.master, dst = ctx.dst;
  const stacked = months.length > 1;
  const isTime = (mode === 'hours' || mode === 'util');
  const split  = CFG.SPLIT_SN;
  const nTypes = split ? 2 : 1;

  const nInfoCols = CFG.INFO_LAST_COL - CFG.INFO_FIRST_COL + 1;
  const nTypeCols = split ? 1 : 0;
  const dayOff    = nInfoCols + nTypeCols;           // ตำแหน่งวันที่ 1 ใน line[]
  const nDayCols  = 31 + (CFG.ADD_TOTAL ? 1 : 0);
  const totalIdx  = 31;
  const nOutCols  = dayOff + nDayCols;
  const dayCol    = CFG.INFO_FIRST_COL + dayOff;     // คอลัมน์จริงของวันที่ 1

  // ---- รายชื่อเครื่องจากชีตซ่อน ----
  const mLast = master.getLastRow();
  if (mLast < CFG.TPL_FIRST_ROW) throw new Error('ชีต "' + CFG.MASTER_SHEET + '" ไม่มีรายชื่อเครื่อง');
  const rawInfo = master.getRange(CFG.TPL_FIRST_ROW, CFG.INFO_FIRST_COL,
                                  mLast - CFG.TPL_FIRST_ROW + 1, nInfoCols).getValues();
  const info = rawInfo.filter(function (r) { return !isBandRow_(r); });   // ข้ามแถว "เดือน..."
  const bandSkipped = rawInfo.length - info.length;
  const nTplRows = info.length;

  // ---- อ่านข้อมูลดิบ ----
  const srcLast = src.getLastRow();
  const endRow = CFG.SRC_END_ROW > 0 ? Math.min(CFG.SRC_END_ROW, srcLast) : srcLast;
  const nRows = endRow - CFG.SRC_START_ROW + 1;
  if (nRows <= 0) throw new Error('ช่วงแถวต้นทางไม่ถูกต้อง');

  ss.toast('กำลังอ่าน ' + nRows.toLocaleString() + ' แถว...', 'รอสักครู่', 30);

  const colDate = readColumn_(src, CFG.SRC_START_ROW, nRows, CFG.SRC_DATE_COL);
  const colKey  = readColumn_(src, CFG.SRC_START_ROW, nRows, CFG.SRC_KEY_COL);
  const colA = isTime ? readColumn_(src, CFG.SRC_START_ROW, nRows, CFG.SRC_START_TIME_COL)
                      : readColumn_(src, CFG.SRC_START_ROW, nRows, CFG.SRC_QTY_COL);
  const colB = isTime ? readColumn_(src, CFG.SRC_START_ROW, nRows, CFG.SRC_END_TIME_COL)
                      : null;
  const colSN = split ? readColumn_(src, CFG.SRC_START_ROW, nRows, CFG.SRC_SN_COL) : null;

  const slotOf = {};
  for (let i = 0; i < months.length; i++) slotOf[months[i]] = i;

  const byKey = new Map();
  let used = 0, badDate = 0, badTime = 0, skipped = 0, cntN = 0, cntSN = 0;

  for (let i = 0; i < nRows; i++) {
    const rawDate = colDate[i];
    if (rawDate === '' || rawDate === null) continue;

    let y, m, d;
    if (rawDate instanceof Date) {
      y = rawDate.getFullYear(); m = rawDate.getMonth() + 1; d = rawDate.getDate();
    } else {
      const p = parseDate_(rawDate);
      if (!p) { badDate++; continue; }
      y = p.y; m = p.m; d = p.d;
    }

    if (y !== year) { skipped++; continue; }
    const slot = slotOf[m];
    if (slot === undefined) { skipped++; continue; }

    const key = norm_(colKey[i]);
    if (key === '') continue;

    let val;
    if (isTime) {
      val = spanMinutes_(colA[i], colB[i]);
      if (val === null) { badTime++; continue; }
    } else {
      const raw = colA[i];
      val = (typeof raw === 'number') ? raw : num_(raw);
      if (!isFinite(val)) continue;
    }

    const t = (split && isSN_(colSN[i])) ? 1 : 0;
    if (t) cntSN++; else cntN++;

    let arr = byKey.get(key);
    if (!arr) { arr = new Float64Array(months.length * nTypes * 31); byKey.set(key, arr); }
    arr[(slot * nTypes + t) * 31 + d - 1] += val;
    used++;
  }

  // ---- ประกอบผลลัพธ์ ----
  const shiftMin = CFG.SHIFT_HOURS * 60;
  const bandColor = (mode === 'qty') ? '#d2e3fc'
                  : (mode === 'hours') ? '#d9ead3' : '#fce8b2';

  const out = [], bg = [], fw = [];
  let hits = 0;

  for (let s = 0; s < months.length; s++) {
    const mNum = months[s];
    const dim  = new Date(year, mNum, 0).getDate();

    if (stacked) {
      const band = new Array(nOutCols).fill('');
      band[0] = 'เดือน ' + mNum + ' / ' + year + '  (' + modeLabel_(mode) + ')';
      out.push(band);
      bg.push(new Array(nOutCols).fill(CFG.COLOR_BLOCKS ? bandColor : null));
      fw.push(new Array(nOutCols).fill('bold'));
    }

    for (let r = 0; r < nTplRows; r++) {
      const key = norm_(info[r][CFG.KEY_COL - CFG.INFO_FIRST_COL]);
      const arr = key === '' ? null : byKey.get(key);
      const types = (split && key !== '') ? nTypes : 1;

      for (let t = 0; t < types; t++) {
        const line = new Array(nOutCols).fill('');
        for (let c = 0; c < nInfoCols; c++) line[c] = info[r][c];
        if (split && key !== '') line[nInfoCols] = CFG.TYPE_LABELS[t];

        const base = (s * nTypes + t) * 31;
        let sum = 0, activeDays = 0;

        for (let d = 0; d < 31; d++) {
          if (d >= dim) continue;
          const v = arr ? arr[base + d] : 0;
          if (!v) { line[dayOff + d] = CFG.BLANK_AS; continue; }

          sum += v;
          activeDays++;
          hits++;

          if (mode === 'qty')        line[dayOff + d] = v;
          else if (mode === 'hours') line[dayOff + d] = v / 60;
          else                       line[dayOff + d] = v / shiftMin;
        }

        if (CFG.ADD_TOTAL) {
          if (!activeDays) {
            line[dayOff + totalIdx] = CFG.BLANK_AS;
          } else if (mode === 'qty') {
            line[dayOff + totalIdx] = sum;
          } else if (mode === 'hours') {
            line[dayOff + totalIdx] = sum / 60;
          } else {
            const basis = (CFG.UTIL_TOTAL_BASIS === 'all_days') ? dim : activeDays;
            line[dayOff + totalIdx] = sum / (shiftMin * basis);
          }
        }

        out.push(line);
        bg.push(new Array(nOutCols).fill((t === 1 && CFG.SN_ROW_COLOR) ? CFG.SN_ROW_COLOR : null));
        fw.push(new Array(nOutCols).fill('normal'));
      }
    }
  }

  const nOutRows = out.length;

  // ---- เตรียมพื้นที่ ----
  const firstRow = CFG.TPL_FIRST_ROW;
  const needRows = firstRow + nOutRows - 1;
  if (dst.getMaxRows() < needRows) {
    dst.insertRowsAfter(dst.getMaxRows(), needRows - dst.getMaxRows());
  }
  const needCols = CFG.INFO_FIRST_COL + nOutCols - 1;
  if (dst.getMaxColumns() < needCols) {
    dst.insertColumnsAfter(dst.getMaxColumns(), needCols - dst.getMaxColumns());
  }

  const wipe = dst.getRange(firstRow, CFG.INFO_FIRST_COL,
                            dst.getMaxRows() - firstRow + 1, nOutCols);
  wipe.clearContent();
  wipe.setBackground(null);
  wipe.setFontWeight('normal');

  // ---- เขียน ----
  const outRange = dst.getRange(firstRow, CFG.INFO_FIRST_COL, nOutRows, nOutCols);
  outRange.setValues(out);
  outRange.setBackgrounds(bg);
  outRange.setFontWeights(fw);

  const fmt = (mode === 'qty') ? CFG.FMT_QTY
            : (mode === 'hours') ? CFG.FMT_HOURS : CFG.FMT_UTIL;
  dst.getRange(firstRow, dayCol, nOutRows, nDayCols).setNumberFormat(fmt);

  const dayHead = [];
  for (let d = 1; d <= 31; d++) dayHead.push(d);
  if (CFG.ADD_TOTAL) dayHead.push(mode === 'util' ? 'เฉลี่ย' : 'รวม');
  dst.getRange(CFG.DAY_ROW, dayCol, 1, nDayCols).setValues([dayHead]);
  if (split) dst.getRange(CFG.DAY_ROW, CFG.INFO_FIRST_COL + nInfoCols).setValue('ประเภท');

  if (CFG.YEAR_CELL) dst.getRange(CFG.YEAR_CELL).setValue(year);
  if (CFG.MONTH_CELL && !stacked) dst.getRange(CFG.MONTH_CELL).setValue(months[0]);

  if (dst.getFrozenRows() < 2) dst.setFrozenRows(2);
  dst.setFrozenColumns(dayOff);

  let msg = modeLabel_(mode) + ' -> ชีต "' + dst.getName() + '" • ' +
            (stacked ? ('ทั้งปี ' + year) : ('เดือน ' + months[0] + '/' + year)) +
            ' • อ่าน ' + nRows.toLocaleString() + ' แถว • ใช้ได้ ' + used + ' แถว' +
            (split ? (' (Normal ' + cntN + ' / SN ' + cntSN + ')') : '') +
            ' • เติม ' + hits + ' ช่อง • ' +
            ((Date.now() - t0) / 1000).toFixed(1) + ' วิ';
  if (badTime) msg += ' • เวลาใช้ไม่ได้ ' + badTime;
  if (badDate) msg += ' • วันที่อ่านไม่ออก ' + badDate;
  if (skipped) msg += ' • นอกช่วง ' + skipped;
  if (bandSkipped) msg += ' • ข้ามแถว "เดือน..." ในรายชื่อเครื่อง ' + bandSkipped + ' แถว';
  ss.toast(msg, 'เสร็จแล้ว (' + SCRIPT_VERSION + ')', 15);
  Logger.log(msg);
}

function modeLabel_(mode) {
  return mode === 'qty' ? 'จำนวนชิ้น'
       : mode === 'hours' ? 'ชั่วโมงเดินเครื่อง'
       : 'อัตราใช้งาน %';
}


/* ==================== ตัวช่วย ==================== */

function norm_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim().toUpperCase();
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return NaN;
  if (typeof v === 'number') return v;
  if (v instanceof Date) return NaN;
  return parseFloat(String(v).replace(/,/g, '').trim());
}

function readColumn_(sheet, startRow, nRows, col) {
  const out = new Array(nRows);
  let done = 0;
  while (done < nRows) {
    const take = Math.min(CFG.CHUNK_ROWS, nRows - done);
    const vals = sheet.getRange(startRow + done, col, take, 1).getValues();
    for (let j = 0; j < take; j++) out[done + j] = vals[j][0];
    done += take;
  }
  return out;
}

function parseDate_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
  }
  if (v === '' || v === null || v === undefined) return null;

  const parts = String(v).trim().split(/[\/\-.\s]+/);
  if (parts.length < 3) return null;

  const a = parseInt(parts[0], 10);
  const b = parseInt(parts[1], 10);
  const c = parseInt(parts[2], 10);
  if (!isFinite(a) || !isFinite(b) || !isFinite(c)) return null;

  let y, m, d;
  if (a > 31) { y = a; m = b; d = c; }
  else        { d = a; m = b; y = c; }

  if (y > 2400) y -= 543;
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  return { y: y, m: m, d: d };
}

function getYear_(dst) {
  if (CFG.YEAR_CELL) {
    const raw = dst.getRange(CFG.YEAR_CELL).getValue();
    if (raw instanceof Date) return raw.getFullYear();
    let y = num_(raw);
    if (isFinite(y)) {
      if (y > 2400) y -= 543;
      if (y >= 1900 && y <= 2200) return Math.round(y);
    }
  }
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('เลือกปี', 'พิมพ์ปี เช่น  2026  หรือ  2569', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error('ยกเลิกแล้ว');
  let y = parseInt(res.getResponseText().trim(), 10);
  if (!isFinite(y)) throw new Error('ปีไม่ถูกต้อง');
  if (y > 2400) y -= 543;
  return y;
}

function getMonthYear_(dst) {
  let month = null, year = null;

  if (CFG.MONTH_CELL) {
    const raw = dst.getRange(CFG.MONTH_CELL).getValue();
    const full = parseDate_(raw);
    if (full) { month = full.m; year = full.y; }
    else {
      const n = num_(raw);
      if (isFinite(n) && n >= 1 && n <= 12) month = Math.round(n);
    }
  }

  if (month === null) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.prompt('เลือกเดือน', 'พิมพ์เดือน เช่น  9  หรือ  2026/09',
                          ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) throw new Error('ยกเลิกแล้ว');
    const txt = res.getResponseText().trim();
    const bits = txt.split(/[\/\-.]/).length;
    const p = parseDate_(bits === 2 ? txt + '/1' : txt);
    if (p) { month = p.m; year = p.y; }
    else {
      const m2 = parseInt(txt, 10);
      if (!(m2 >= 1 && m2 <= 12)) throw new Error('เดือนไม่ถูกต้อง: ' + txt);
      month = m2;
    }
  }

  if (CFG.YEAR_CELL) {
    const yv = num_(dst.getRange(CFG.YEAR_CELL).getValue());
    if (isFinite(yv)) year = yv > 2400 ? yv - 543 : yv;
  }
  if (!year) year = new Date().getFullYear();

  return { month: month, year: year };
}


/* ================ เครื่องมือตรวจสอบ ================ */

function debugTimes() {
  const src = SpreadsheetApp.getActive().getSheetByName(CFG.SRC_SHEET);
  const n = 25;
  const a = src.getRange(CFG.SRC_START_ROW, CFG.SRC_START_TIME_COL, n, 1).getValues();
  const b = src.getRange(CFG.SRC_START_ROW, CFG.SRC_END_TIME_COL, n, 1).getValues();

  const lines = [];
  for (let i = 0; i < n; i++) {
    const min = spanMinutes_(a[i][0], b[i][0]);
    lines.push('แถว ' + (CFG.SRC_START_ROW + i) + ': ' + a[i][0] + ' -> ' + b[i][0] +
               '  =  ' + (min === null ? 'ใช้ไม่ได้'
                          : (min.toFixed(0) + ' นาที (' + (min / 60).toFixed(2) + ' ชม., ' +
                             (min / (CFG.SHIFT_HOURS * 60) * 100).toFixed(1) + '%)')));
  }
  Logger.log(lines.join('\n'));
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

function debugMissingKeys() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CFG.SRC_SHEET);
  const master = getMaster_(ss);

  const srcLast = src.getLastRow();
  const endRow = CFG.SRC_END_ROW > 0 ? Math.min(CFG.SRC_END_ROW, srcLast) : srcLast;
  const nRows = endRow - CFG.SRC_START_ROW + 1;

  const colKey = readColumn_(src, CFG.SRC_START_ROW, nRows, CFG.SRC_KEY_COL);
  const srcKeys = new Set();
  for (let i = 0; i < colKey.length; i++) srcKeys.add(norm_(colKey[i]));

  const nTpl = master.getLastRow() - CFG.TPL_FIRST_ROW + 1;
  const keys = master.getRange(CFG.TPL_FIRST_ROW, CFG.KEY_COL, nTpl, 1).getValues();

  const missing = [];
  keys.forEach(function (r, i) {
    const k = norm_(r[0]);
    if (k !== '' && !srcKeys.has(k)) missing.push('แถว ' + (CFG.TPL_FIRST_ROW + i) + ': ' + r[0]);
  });

  const out = missing.length
    ? 'รหัสที่ไม่พบในต้นทาง ' + missing.length + ' ตัว:\n' + missing.join('\n')
    : 'รหัสครบทุกตัว';
  Logger.log(out);
  SpreadsheetApp.getUi().alert(out.substring(0, 1000));
}
