// ============================================================
// ManualAttendance.gs — temporary supervisor attendance backfill (22–31 Aug)
// ============================================================
// A no-login supervisor page collects Present / Absent / Day Off per
// employee per day into a STAGING tab (Manual_Attendance_Aug). Supervisors
// only collect; nothing hits pay until HR merges it into Payroll Review.
//
// Absent  -> unpaid deduction (Basic/26 x absent days) in the period.
// Day Off -> paid (no deduction).  Present -> normal.
//
// ENDPOINTS: getManualAttnEmployees, saveManualAttendance,
//            getManualAttendance, getAbsenceDeductionsForPeriod
// ============================================================

const MANUAL_ATTN_TAB = 'Manual_Attendance_Aug';
const MANUAL_ATTN_LOCATIONS = ['Kaifan', 'Jabriya', 'Ardiya', 'Qurain', 'Central Kitchen', 'Khairan', 'Head Office'];

function ensureManualAttnTab_(ss) {
  let sheet = ss.getSheetByName(MANUAL_ATTN_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(MANUAL_ATTN_TAB);
    sheet.appendRow(['Employee ID', 'Name', 'Location', 'Shift', 'Date', 'Status', 'Submitted By', 'Timestamp']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setNumberFormat('@'); // Date as text
  }
  return sheet;
}

// Employee list + locations for the supervisor page dropdowns.
function getManualAttnEmployees(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const d = hr.getDataRange().getValues();
    const H = d[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');
    const out = [];
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][cId] || '').trim();
      if (!id) continue;
      const status = cStatus !== -1 ? String(d[i][cStatus] || '').trim().toLowerCase() : '';
      if (status === 'inactive') continue;
      out.push({ id: id, name: cName !== -1 ? String(d[i][cName] || '').trim() : id });
    }
    out.sort((a, b) => a.name.localeCompare(b.name));
    // Locations from the Branch_Locations tab (active only), managed on the Locations page.
    let locations = [];
    try { locations = (getLocations({}).data || []).filter(l => l.active).map(l => l.name); } catch (e) {}
    if (!locations.length) locations = MANUAL_ATTN_LOCATIONS; // fallback
    // Active shift names for the page dropdown.
    let shiftNames = [];
    try { shiftNames = (getShifts({}).data || []).filter(s => s.active).map(s => s.name); } catch (e) {}
    return { success: true, employees: out, locations: locations, shifts: shiftNames };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Save one attendance mark (latest overwrites the same employee+date).
// params: { empId, location, date (dd-mm-yyyy), status (Present|Absent|Day Off), by }
function saveManualAttendance(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const empId = String(params.empId || '').trim();
    const location = String(params.location || '').trim();
    const shift = String(params.shift || '').trim();
    const dateStr = String(params.date || '').trim();
    const status = String(params.status || '').trim();
    if (!empId || !dateStr || !status) return { success: false, error: 'empId, date and status are required.' };
    if (['Present', 'Absent', 'Day Off'].indexOf(status) === -1) return { success: false, error: 'status must be Present, Absent or Day Off.' };

    // Date guard: only 22–31 Aug 2026.
    const d = parseFlexibleDate_(dateStr);
    if (!d) return { success: false, error: 'Bad date.' };
    const lo = new Date(2026, 7, 22), hi = new Date(2026, 7, 31, 23, 59, 59);
    if (d < lo || d > hi) return { success: false, error: 'Date must be between 22-08-2026 and 31-08-2026.' };

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureManualAttnTab_(ss);
    const name = empNameById_(ss, empId) || empId;
    const dd = formatDdMmYyyy_(d);

    // Latest overwrites: find an existing row for this emp+date (Date is col 5).
    const data = sheet.getDataRange().getValues();
    let rowNum = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === empId && String(data[i][4]).trim() === dd) { rowNum = i + 1; break; }
    }
    const stamp = Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm:ss');
    const row = [empId, name, location, shift, dd, status, String(params.by || 'Supervisor').trim(), stamp];
    if (rowNum === -1) { sheet.appendRow(row); sheet.getRange(sheet.getLastRow(), 5).setNumberFormat('@'); }
    else sheet.getRange(rowNum, 1, 1, 8).setValues([row]);

    return { success: true, message: name + ' marked ' + status + ' on ' + dd + (location ? ' @ ' + location : '') + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Read all manual attendance (for HR review / merge).
function getManualAttendance(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureManualAttnTab_(ss);
    const d = sheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      if (!String(d[i][0]).trim()) continue;
      rows.push({ empId: d[i][0], name: d[i][1], location: d[i][2], shift: d[i][3], date: d[i][4], status: d[i][5], by: d[i][6], at: d[i][7] });
    }
    // Tally per employee
    const tally = {};
    rows.forEach(r => {
      const k = r.empId;
      tally[k] = tally[k] || { name: r.name, present: 0, absent: 0, dayoff: 0 };
      const s = String(r.status).toLowerCase();
      if (s === 'present') tally[k].present++;
      else if (s === 'absent') tally[k].absent++;
      else if (s === 'day off') tally[k].dayoff++;
    });
    return { success: true, count: rows.length, rows: rows, tally: tally };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// NOTE: getAbsenceDeductionsForPeriod_ used to be defined HERE, reading
// Manual_Attendance_Aug. It has been REMOVED.
//
// Apps Script shares ONE global namespace across every .gs file, so having
// two functions with this name — this one, and the Attendance_Daily version
// in AbsenceDeductions.gs — meant whichever file loaded last silently won.
// Absence deductions therefore depended on file load order, which is not
// something you control. If this version won, only Aug 22-31 backfill rows
// produced deductions and everything from September onward deducted nothing.
//
// Attendance_Daily is now the SINGLE SOURCE OF TRUTH for attendance.
// The live definition lives in AbsenceDeductions.gs.
//
// Manual_Attendance_Aug is still the input for the Aug 22-31 supervisor
// backfill: writeDayToAttendanceDaily_ reads it as priority 1 and folds those
// rows into Attendance_Daily, so they still count — through one path, not two.

// One-time: mark Head Office staff Present for 22–30 Aug 2026, skipping
// their weekly off (Fri/Sat). Writes into Manual_Attendance_Aug (latest-wins).
// Head Office staff = employees whose Location is 'Head Office' OR whose Shift
// is an auto-present shift. params.dryRun=true reports without writing.
function backfillHeadOfficeAug(params) {
  try {
    params = params || {};
    const dry = String(params.dryRun).toLowerCase() === 'true';
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureManualAttnTab_(ss);

    const hr = ss.getSheetByName('HR Maya');
    const d = hr.getDataRange().getValues();
    const H = d[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    let cLoc = H.findIndex(h => String(h).trim().toLowerCase() === 'location');
    if (cLoc === -1) cLoc = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cShift = H.findIndex(h => String(h).trim().toLowerCase() === 'shift');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');

    // Auto-present shift names (so shift-based Head Office also qualifies).
    const autoShifts = {};
    try { (getShifts({}).data || []).forEach(s => { if (s.autoPresent) autoShifts[s.name.toLowerCase()] = true; }); } catch (e) {}

    // Existing marks (emp|date) to avoid overwrite.
    const existing = {};
    const md = sheet.getDataRange().getValues();
    for (let i = 1; i < md.length; i++) existing[String(md[i][0]).trim() + '|' + String(md[i][4]).trim()] = true;

    // Head Office employees.
    const staff = [];
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][cId] || '').trim(); if (!id) continue;
      if (cStatus !== -1 && String(d[i][cStatus]).trim().toLowerCase() === 'inactive') continue;
      const loc = cLoc !== -1 ? String(d[i][cLoc] || '').trim() : '';
      const shift = cShift !== -1 ? String(d[i][cShift] || '').trim() : '';
      const isHO = loc.toLowerCase() === 'head office' || (shift && autoShifts[shift.toLowerCase()]);
      if (isHO) staff.push({ id: id, name: cName !== -1 ? String(d[i][cName]).trim() : id, loc: loc || 'Head Office', shift: shift });
    }

    // Days 22..30 Aug 2026, skip Fri(5)/Sat(6).
    const rows = []; let skippedOff = 0, skippedExisting = 0;
    for (let day = 22; day <= 30; day++) {
      const dt = new Date(2026, 7, day);
      const dow = dt.getDay(); // 0=Sun..6=Sat
      if (dow === 5 || dow === 6) { skippedOff += staff.length; continue; } // Fri/Sat off
      const dd = formatDdMmYyyy_(dt);
      staff.forEach(s => {
        if (existing[s.id + '|' + dd]) { skippedExisting++; return; }
        rows.push([s.id, s.name, s.loc, s.shift, dd, 'Present', 'System (Head Office backfill)', Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm:ss')]);
      });
    }

    if (!dry && rows.length) {
      const start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, rows.length, 8).setValues(rows);
      sheet.getRange(start, 5, rows.length, 1).setNumberFormat('@');
    }
    return {
      success: true, dryRun: dry, headOfficeStaff: staff.length, marked: rows.length,
      skippedWeeklyOff: skippedOff, skippedAlreadyMarked: skippedExisting,
      message: (dry ? 'DRY RUN — ' : '') + 'Head Office backfill: ' + staff.length + ' staff, ' + rows.length + ' Present day-rows (22–30 Aug, Fri/Sat skipped)' + (skippedExisting ? ', ' + skippedExisting + ' already marked' : '') + '.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}