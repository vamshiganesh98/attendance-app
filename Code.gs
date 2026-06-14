// ── TAKSHASHILA — Attendance Backend (v10) ──
// Deploy as Web App: Execute as "Me", Access "Anyone"
//
// Layout: one tab per year ("2025", "2026" …)
// Grid: dates as ROWS, students as COLUMNS (flipped from v9)
// Cells: P/A stored but font colour matches background — looks like a heatmap

const SECRET_KEY     = "takshashila2025";
const STUDENTS_SHEET = "Students";

function getSpreadsheet(sheetId) {
  if (!sheetId) throw new Error("No Spreadsheet ID provided. Set it in app Settings.");
  return SpreadsheetApp.openById(sheetId);
}

// Returns the year tab (e.g. "2025"), creating it if needed
function getYearSheet(ss, dateStr) {
  const year  = String(dateStr).split('-')[0];
  let   sheet = ss.getSheetByName(year);
  if (!sheet) {
    sheet = ss.insertSheet(year);
    // Row 1, Col 1 = "Date" header
    sheet.getRange(1, 1).setValue("Date");
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }
  return sheet;
}

// All tabs whose name is a 4-digit year
function getAllYearSheets(ss) {
  return ss.getSheets().filter(s => /^\d{4}$/.test(s.getName()));
}

// ── COLUMN helpers (students are columns, starting col 2) ──
function getOrCreateStudentCol(sheet, name) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === name) return i + 1;
  }
  const newCol = lastCol + 1;
  sheet.getRange(1, newCol).setValue(name);
  sheet.setColumnWidth(newCol, 60);
  return newCol;
}

// ── ROW helpers (dates are rows, starting row 2) ──
function getOrCreateDateRow(sheet, dateStr) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow >= 2) {
    const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (formatDate(dates[i][0]) === dateStr) return i + 2;
    }
  }
  const newRow = lastRow + 1;
  const [y, m, d] = dateStr.split('-').map(Number);
  const cell = sheet.getRange(newRow, 1);
  cell.setValue(new Date(y, m - 1, d));
  cell.setNumberFormat("d/M/yy");
  return newRow;
}

function doGet(e) {
  try {
    const key      = e.parameter.key;
    const action   = e.parameter.action;
    const callback = e.parameter.callback;
    const sheetId  = e.parameter.sheetId;

    if (key !== SECRET_KEY) return json({ ok: false, error: "Invalid key" }, callback);
    if (action === "ping")  return json({ ok: true, message: "Takshashila backend active" }, callback);

    const ss = getSpreadsheet(sheetId);

    // ── Save attendance ──
    if (action === "save") {
      const data    = JSON.parse(e.parameter.data);
      const records = data.records;
      const dateStr = records[0].date;

      const sheet  = getYearSheet(ss, dateStr);
      const dateRow = getOrCreateDateRow(sheet, dateStr);

      records.forEach(r => {
        const col = getOrCreateStudentCol(sheet, r.name);
        sheet.getRange(dateRow, col).setValue(r.status === "present" ? "P" : "A");
      });

      styleSheet(sheet);
      return json({ ok: true, saved: records.length }, callback);
    }

    // ── Bulk import ──
    if (action === "importBulk") {
      let raw = e.parameter.data;
      // Decode base64 if sent with enc=b64
      if (e.parameter.enc === 'b64') {
        raw = Utilities.newBlob(Utilities.base64Decode(raw)).getDataAsString();
      }
      let records;
      if (raw && raw[0] !== '{' && raw[0] !== '[') {
        records = raw.split(',').filter(Boolean).map(r => {
          const [date, name, s] = r.split('|');
          return { date: date.trim(), name: name.trim(), status: s.trim() === 'P' ? 'present' : 'absent' };
        });
      } else {
        records = JSON.parse(raw).records;
      }

      // Group by year so we only touch each sheet once
      const byYear = {};
      records.forEach(r => {
        const yr = r.date.split('-')[0];
        if (!byYear[yr]) byYear[yr] = [];
        byYear[yr].push(r);
      });

      Object.entries(byYear).forEach(([yr, recs]) => {
        const sheet = getYearSheet(ss, yr + '-01-01');

        // Pre-create all date rows and student cols
        const dates = [...new Set(recs.map(r => r.date))].sort();
        const names = [...new Set(recs.map(r => r.name))];
        const dateRowMap = {};
        const nameColMap = {};
        dates.forEach(d => { dateRowMap[d] = getOrCreateDateRow(sheet, d); });
        names.forEach(n => { nameColMap[n] = getOrCreateStudentCol(sheet, n); });

        // Bulk-read then bulk-write
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow < 2 || lastCol < 2) return;
        const range  = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
        const values = range.getValues();

        recs.forEach(r => {
          const row = dateRowMap[r.date] - 2;
          const col = nameColMap[r.name]  - 2;
          if (row >= 0 && col >= 0 && row < values.length && col < values[0].length) {
            values[row][col] = r.status === "present" ? "P" : "A";
          }
        });

        range.setValues(values);
        styleSheet(sheet);
      });

      return json({ ok: true, imported: records.length }, callback);
    }

    // ── Save students list ──
    if (action === "saveStudents") {
      const data    = JSON.parse(e.parameter.data);
      const active  = data.students || [];
      const alumnis = data.alumni   || [];
      const sheet   = getOrCreate(ss, STUDENTS_SHEET, ["name", "alumni"]);
      sheet.clearContents();
      sheet.appendRow(["name", "alumni"]);
      const maxRows = Math.max(active.length, alumnis.length);
      for (let i = 0; i < maxRows; i++) {
        sheet.appendRow([active[i] || "", alumnis[i] || ""]);
      }
      return json({ ok: true }, callback);
    }

    // ── Fetch students ──
    if (action === "students") {
      const sheet = ss.getSheetByName(STUDENTS_SHEET);
      if (!sheet) return json({ ok: true, students: [], alumni: [] }, callback);
      const rows     = sheet.getDataRange().getValues();
      const students = [], alumni = [];
      rows.slice(1).forEach(r => {
        const name   = String(r[0]).trim();
        const alumno = String(r[1]).trim();
        if (name)   students.push(name);
        if (alumno) alumni.push(alumno);
      });
      return json({ ok: true, students, alumni }, callback);
    }

    // ── Fetch records (reads all year tabs) ──
    if (action === "records") {
      const yearSheets = getAllYearSheets(ss);
      if (!yearSheets.length) return json({ ok: true, records: [] }, callback);

      const records = [];
      yearSheets.forEach(sheet => {
        const data    = sheet.getDataRange().getValues();
        const headers = data[0]; // Row 0: "Date", student1, student2, ...
        for (let r = 1; r < data.length; r++) {
          const dateStr = data[r][0] ? formatDate(data[r][0]) : null;
          if (!dateStr) continue;
          for (let c = 1; c < headers.length; c++) {
            const name = String(headers[c]).trim();
            const val  = String(data[r][c]).trim().toUpperCase();
            if (name && (val === "P" || val === "A")) {
              records.push({ date: dateStr, name, status: val === "P" ? "present" : "absent" });
            }
          }
        }
      });

      return json({ ok: true, records }, callback);
    }

    // ── Delete all data ──
    if (action === "clearAll") {
      getAllYearSheets(ss).forEach(s => s.clear());
      const stuSheet = ss.getSheetByName(STUDENTS_SHEET);
      if (stuSheet) stuSheet.clear();
      return json({ ok: true }, callback);
    }

    return json({ ok: false, error: "Unknown action" }, callback);

  } catch (err) {
    return json({ ok: false, error: err.message + " | action=" + e.parameter.action + " | sheetId=" + (e.parameter.sheetId || "none") }, e.parameter.callback);
  }
}

function doPost(e) {
  return json({ ok: false, error: "Use GET" }, null);
}

// Style a year sheet: header row dark, date col bold, P=green, A=red, empty=-
function styleSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  // Header row (student names)
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground("#1A1208");
  headerRange.setFontColor("#F0E8D8");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");

  // Date column
  sheet.getRange(1, 1, lastRow, 1).setFontWeight("bold");
  sheet.setColumnWidth(1, 80);

  if (lastRow > 1 && lastCol > 1) {
    const range  = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
    range.setHorizontalAlignment("center");
    range.setFontSize(10);
    range.setFontWeight("bold");

    const values     = range.getValues();
    const bgColors   = [];
    const fontColors = [];

    for (let r = 0; r < values.length; r++) {
      const bgRow = [], fcRow = [];
      for (let c = 0; c < values[r].length; c++) {
        const val = String(values[r][c]).trim().toUpperCase();
        if (val === "P") {
          bgRow.push("#E8F2EB"); fcRow.push("#2D6A3F"); // green bg, dark green text
        } else if (val === "A") {
          bgRow.push("#F5E8E8"); fcRow.push("#8B2020"); // red bg, dark red text
        } else {
          values[r][c] = "-";                           // fill empty with dash
          bgRow.push("#F8F8F8"); fcRow.push("#BBBBBB"); // light gray, muted
        }
      }
      bgColors.push(bgRow);
      fontColors.push(fcRow);
    }

    range.setValues(values);
    range.setBackgrounds(bgColors);
    range.setFontColors(fontColors);
  }
}

function json(obj, callback) {
  const text = callback
    ? callback + '(' + JSON.stringify(obj) + ')'
    : JSON.stringify(obj);
  const mime = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(text).setMimeType(mime);
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return String(val).trim();
}

function getOrCreate(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}
