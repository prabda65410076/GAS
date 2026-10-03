// ============================================================
//   รับข้อมูลจาก PAD (POST เป็น CSV) แล้วเขียนทับข้อมูลในชีตของ "เดือนปัจจุบัน"
//   มีแค่ 2 ชุดข้อมูล: Sales และ PD (แยกปลายทางด้วย query parameter "type")
//
//   วิธีเรียกจาก PAD:
//     - ข้อมูล Sales -> ยิงไปที่ ...exec?type=sales
//     - ข้อมูล PD    -> ยิงไปที่ ...exec?type=pd
//
//   Request body (Content type: text/plain): ใส่แค่ค่า CSV ล้วนๆ ไม่ต้องมี key=value
//
//   การทำงานรายเดือน:
//     - ภายในเดือนเดียวกัน: ทุกครั้งที่ PAD ส่งข้อมูลมา จะลบของเก่าในชีต "Sales" / "PD"
//       แล้วเขียนข้อมูลใหม่ทับทั้งหมด
//     - เมื่อขึ้นเดือนใหม่ (ครั้งแรกที่ส่งข้อมูลมาในเดือนใหม่):
//         1) เปลี่ยนชื่อชีตเดิมเป็น "ชื่อชีต + เดือน" เช่น "Sales 2026-09" (เก็บไว้เป็นประวัติ)
//         2) สร้างชีตใหม่ชื่อ "Sales" สำหรับเดือนใหม่ แล้วค่อยเขียนข้อมูลลงไป
//     - เดือนของแต่ละชีตถูกจำไว้ใน Document Properties (key: CURRENT_MONTH_<type>)
// ============================================================

// แผนที่ type (query param) -> ชื่อชีตปลายทาง (type ไม่สนตัวพิมพ์เล็ก/ใหญ่)
var SHEET_MAP = {
  "sales": "Sales",
  "pd": "PD"
};

// คอลัมน์ที่ต้องแปลงเป็นตัวเลข (นับคอลัมน์ A=1) — แก้ start/end ให้ตรงกับข้อมูลจริง
// ถ้าไม่มีคอลัมน์ตัวเลข ให้ตั้งเป็น null
var NUMBER_COLS_MAP = {
  "sales": { start: 9, end: 9 },  // TODO: ปรับให้ตรงกับไฟล์ Sales
  "pd": { start: 9, end: 9 }      // TODO: ปรับให้ตรงกับไฟล์ PD
};

var NUMBER_FORMAT = "#,##0.00";  // รูปแบบตัวเลข
var DATE_FORMAT = "yyyy-mm-dd";  // รูปแบบแสดงผลของคอลัมน์ U (Date Only)
var MONTH_KEY_FORMAT = "yyyy-MM"; // รูปแบบเดือนที่ต่อท้ายชื่อชีตเก่า เช่น "Sales 2026-09"
var ARCHIVE_SEPARATOR = " ";      // ตัวคั่นระหว่างชื่อชีตกับเดือน

// ──────────────────────────────────────────────────────────
//   แปลงข้อความวันที่ -> Date object (คืน null ถ้าแปลงไม่ได้)
// ──────────────────────────────────────────────────────────
function parseDateOnly(str) {
  if (!str) return null;
  var s = String(str).trim();

  // รูปแบบ yyyy-mm-dd หรือ yyyy/mm/dd
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  // รูปแบบ dd/mm/yyyy หรือ dd-mm-yyyy (เผื่อไว้)
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  return null;
}

// ──────────────────────────────────────────────────────────
//   เดือนปัจจุบันตาม timezone ของไฟล์ Spreadsheet เช่น "2026-10"
// ──────────────────────────────────────────────────────────
function getCurrentMonthKey(ss) {
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), MONTH_KEY_FORMAT);
}

// ──────────────────────────────────────────────────────────
//   คืนชีตของเดือนปัจจุบันสำหรับ type นี้
//   ถ้าขึ้นเดือนใหม่แล้ว: เปลี่ยนชื่อชีตเดิมเป็น "ชื่อชีต + เดือนเก่า" แล้วสร้างชีตใหม่
//   คืนค่า { sheet: Sheet, archivedAs: ชื่อชีตที่ถูกเก็บเป็นประวัติ (หรือ null) }
// ──────────────────────────────────────────────────────────
function getSheetForCurrentMonth(ss, type, sheetName) {
  var props = PropertiesService.getDocumentProperties();
  var propKey = "CURRENT_MONTH_" + type;
  var currentMonth = getCurrentMonthKey(ss);
  var storedMonth = props.getProperty(propKey);

  var sheet = ss.getSheetByName(sheetName);
  var archivedAs = null;

  if (!sheet) {
    // ยังไม่มีชีตนี้ สร้างใหม่สำหรับเดือนปัจจุบัน
    sheet = ss.insertSheet(sheetName);
  } else if (storedMonth && storedMonth !== currentMonth) {
    // ขึ้นเดือนใหม่: เก็บชีตเดิมเป็นประวัติ แล้วสร้างชีตใหม่ไว้ตำแหน่งเดิม
    archivedAs = getUniqueSheetName(ss, sheetName + ARCHIVE_SEPARATOR + storedMonth);
    var oldIndex = sheet.getIndex();
    sheet.setName(archivedAs);
    sheet = ss.insertSheet(sheetName, oldIndex - 1);
  }
  // กรณี storedMonth ว่าง (รันครั้งแรก) จะถือว่าชีตที่มีอยู่เป็นของเดือนปัจจุบัน

  props.setProperty(propKey, currentMonth);
  return { sheet: sheet, archivedAs: archivedAs };
}

// ถ้าชื่อซ้ำกับชีตที่มีอยู่แล้ว ให้ต่อท้าย (2), (3), ...
function getUniqueSheetName(ss, baseName) {
  var name = baseName;
  var n = 2;
  while (ss.getSheetByName(name)) {
    name = baseName + " (" + n + ")";
    n++;
  }
  return name;
}

// ──────────────────────────────────────────────────────────
//   Entry point: PAD ส่ง CSV มาแบบ text/plain (POST) พร้อม query param ?type=...
// ──────────────────────────────────────────────────────────
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getDocumentLock();
  try {
    if (!e || !e.parameter || !e.parameter.type) {
      return textOutput("ERROR: ไม่ได้ระบุ type ใน query string (เช่น ?type=sales หรือ ?type=pd) | ตัวเลือกที่มี: " + Object.keys(SHEET_MAP).join(", "));
    }

    var type = String(e.parameter.type).trim().toLowerCase();
    var sheetName = SHEET_MAP[type];

    if (!sheetName) {
      return textOutput("ERROR: ไม่รู้จัก type '" + e.parameter.type + "' | ตัวเลือกที่มี: " + Object.keys(SHEET_MAP).join(", "));
    }

    if (!e.postData || !e.postData.contents) {
      return textOutput("ERROR: No data received (type=" + type + ")");
    }

    // 1) แปลง CSV ที่ได้จาก PAD
    var csvText = e.postData.contents;
    var debugRawLength = csvText.length;
    var debugNewlineCount = (csvText.match(/\n/g) || []).length;

    var newRows = Utilities.parseCsv(csvText);
    var debugParsedRowCount = newRows.length;

    if (newRows.length === 0) {
      return textOutput("ERROR: No data received | type=" + type + " | rawLength=" + debugRawLength);
    }

    // 2) ลบแถวที่ซ้ำกับ header ออก (เผื่อ PAD ส่ง header ซ้ำมาจากหลายหน้า)
    var headerRow = newRows[0];
    var headerKey = headerRow.join("|");
    var filteredRows = [headerRow];
    for (var i = 1; i < newRows.length; i++) {
      if (newRows[i].join("|") !== headerKey) filteredRows.push(newRows[i]);
    }
    newRows = filteredRows;

    // 2.1) ทำให้ทุกแถวมีจำนวนคอลัมน์เท่ากัน (setValues ต้องการตารางสี่เหลี่ยม)
    var TARGET_COL_U = 20; // Index ของคอลัมน์ U (A=0, B=1, ... U=20)
    var maxCols = TARGET_COL_U + 1;
    for (var r0 = 0; r0 < newRows.length; r0++) {
      if (newRows[r0].length > maxCols) maxCols = newRows[r0].length;
    }
    for (var r1 = 0; r1 < newRows.length; r1++) {
      while (newRows[r1].length < maxCols) newRows[r1].push("");
    }

    // 2.2) ดึงวันที่จาก Col B ไปใส่ Col U แบบเป็น Date จริง
    for (var r = 0; r < newRows.length; r++) {
      if (r === 0) {
        newRows[r][TARGET_COL_U] = "Date Only";
      } else {
        var colB_Val = newRows[r][1];
        if (colB_Val) {
          var dateOnly = String(colB_Val).trim().split(" ")[0];
          var dateObj = parseDateOnly(dateOnly);
          newRows[r][TARGET_COL_U] = dateObj ? dateObj : dateOnly;
        }
      }
    }

    var numRows = newRows.length;
    var numCols = maxCols;

    // 2.3) แปลงคอลัมน์ตัวเลขของ type นี้ (ถ้ามี) จากข้อความ -> ตัวเลขจริง
    var numberCols = NUMBER_COLS_MAP[type];
    var lastNumberColIdx = 0;
    if (numberCols) {
      lastNumberColIdx = Math.min(numberCols.end, numCols);
      if (lastNumberColIdx >= numberCols.start) {
        for (var r2 = 1; r2 < newRows.length; r2++) {
          for (var c = numberCols.start; c <= lastNumberColIdx; c++) {
            var cellVal = newRows[r2][c - 1];
            if (cellVal === "" || cellVal === null || typeof cellVal === "undefined") continue;
            var cleaned = String(cellVal).replace(/,/g, "").trim();
            var num = Number(cleaned);
            if (!isNaN(num) && cleaned !== "") {
              newRows[r2][c - 1] = num;
            }
          }
        }
      }
    }

    // 3) ล็อกไว้กันการเรียกซ้อนกัน (เช่น PAD ยิง Sales กับ PD พร้อมกันตอนขึ้นเดือนใหม่)
    lock.waitLock(30000);

    // 3.1) หา/สร้างชีตของเดือนปัจจุบัน (ถ้าขึ้นเดือนใหม่จะเก็บชีตเดิมเป็นประวัติให้อัตโนมัติ)
    var result = getSheetForCurrentMonth(ss, type, sheetName);
    var sheet = result.sheet;

    // 3.2) ลบข้อมูลเก่าของเดือนนี้ แล้วเขียนข้อมูลใหม่ทับ
    sheet.clear();
    var targetRange = sheet.getRange(1, 1, numRows, numCols);
    targetRange.setNumberFormat("@"); // กันรหัสที่ขึ้นต้นด้วย 0 ถูกแปลงเป็นตัวเลข

    // ตั้งฟอร์แมตคอลัมน์ U เป็นวันที่ "ก่อน" setValues
    if (numRows > 1) {
      sheet.getRange(2, TARGET_COL_U + 1, numRows - 1, 1).setNumberFormat(DATE_FORMAT);
    }

    // ตั้งฟอร์แมตคอลัมน์ตัวเลข "ก่อน" setValues
    if (numberCols && lastNumberColIdx >= numberCols.start && numRows > 1) {
      sheet.getRange(2, numberCols.start, numRows - 1, lastNumberColIdx - numberCols.start + 1)
        .setNumberFormat(NUMBER_FORMAT);
    }

    targetRange.setValues(newRows);
    sheet.getRange(numRows + 2, 1).setValue("อัปเดตล่าสุด: " + new Date());
    SpreadsheetApp.flush();

    var msg = "OK: " + (numRows - 1) + " rows written to '" + sheetName + "' (type=" + type + ")";
    if (result.archivedAs) {
      msg += " | ขึ้นเดือนใหม่: เก็บข้อมูลเดือนก่อนไว้ที่ชีต '" + result.archivedAs + "'";
    }
    msg += " | DEBUG: rawLength=" + debugRawLength +
      " | newlineCount=" + debugNewlineCount +
      " | parsedRows(beforeDedupe)=" + debugParsedRowCount;

    return textOutput(msg);

  } catch (err) {
    return textOutput("ERROR: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

function textOutput(msg) {
  return ContentService.createTextOutput(msg).setMimeType(ContentService.MimeType.TEXT);
}
