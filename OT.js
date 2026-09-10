// ============================================================
// OT.gs — Maya HR : daily OT entry, summary, and report
// ============================================================
// Standalone Apps Script file. Shares global scope with Code.gs,
// Payroll.gs, LeaveEncashment.gs, PayrollDocs.gs, SharedUtils.gs.
// Reuses KD (Payroll.gs) and the date helpers (SharedUtils.gs).
// Do NOT redeclare them.
//
// OT lives in its own 'OT_Entries' tab, keyed by real DATE (dd-mm-yyyy):
//   Date | Employee ID | Name | OT Hours | Remark | Added By
// Payroll (without-attendance mode) sums the OT rows whose date falls
// inside the run's period — see getOtHoursForPeriod_() used by Payroll.gs.
//
// ENDPOINTS (added to Code.gs handleRequest):
//   addOtEntry, getOtSummary, generateOtReport
// ============================================================

const OT_ENTRIES_TAB = 'OT_Entries';

// ------------------------------------------------------------
// Ensure the OT tab exists
// ------------------------------------------------------------
function ensureOtTab_(ss) {
  let sheet = ss.getSheetByName(OT_ENTRIES_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(OT_ENTRIES_TAB);
    sheet.appendRow(['Date', 'Employee ID', 'Name', 'OT Hours', 'Remark', 'Added By']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@'); // Date column as text dd-mm-yyyy
    sheet.autoResizeColumns(1, 6);
  }
  return sheet;
}

// ------------------------------------------------------------
// ADD one OT entry (Date, Employee, OT Hours)
// params: { date (dd-mm-yyyy), empId, hours, remark?, by? }
// ------------------------------------------------------------
function addOtEntry(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const d = parseFlexibleDate_(params.date);
    if (!d) return { success: false, error: 'date must be a valid date (dd-mm-yyyy)' };
    const empId = String(params.empId || '').trim();
    if (!empId) return { success: false, error: 'empId is required' };
    const hours = Number(params.hours);
    if (!(hours > 0)) return { success: false, error: 'OT Hours must be greater than 0' };

    const name = empNameById_(ss, empId);
    if (!name) return { success: false, error: 'Employee ID "' + empId + '" not found in HR Maya' };

    const sheet = ensureOtTab_(ss);
    const dateStr = formatDdMmYyyy_(d);

    // If a row already exists for this employee+date, overwrite it (one OT figure per day).
    const data = sheet.getDataRange().getValues();
    let rowNum = -1;
    for (let i = 1; i < data.length; i++) {
      const rd = parseFlexibleDate_(data[i][0]);
      if (rd && sameCalendarDate_(rd, d) && String(data[i][1]).trim().toUpperCase() === empId.toUpperCase()) { rowNum = i + 1; break; }
    }

    const row = [dateStr, empId, name, hours, String(params.remark || '').trim(), String(params.by || 'Dashboard').trim()];
    if (rowNum !== -1) sheet.getRange(rowNum, 1, 1, 6).setValues([row]);
    else sheet.appendRow(row);
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat('@');
    sheet.autoResizeColumns(1, 6);

    return { success: true, date: dateStr, empId: empId, name: name, hours: hours,
      message: (rowNum !== -1 ? 'Updated' : 'Added') + ' OT: ' + name + ' — ' + hours + ' hrs on ' + dateStr };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Per-employee OT totals over a date range (inclusive).
// params: { from (dd-mm-yyyy), to (dd-mm-yyyy) }  — omit for all-time.
// Returns rows with days-with-OT, total hours, and OT amount (hrs * Price Per Hour).
// ------------------------------------------------------------
function getOtSummary(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const from = params.from ? startOfDay_(parseFlexibleDate_(params.from)) : null;
    const to   = params.to   ? endOfDay_(parseFlexibleDate_(params.to))     : null;
    if (params.from && !from) return { success: false, error: 'from must be a valid date (dd-mm-yyyy)' };
    if (params.to && !to) return { success: false, error: 'to must be a valid date (dd-mm-yyyy)' };

    const rate = ratesById_(ss); // { empId: pricePerHour }
    const agg = {}; // empId -> {name, days, hours}
    const sheet = ss.getSheetByName(OT_ENTRIES_TAB);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const rd = parseFlexibleDate_(data[i][0]);
        const id = String(data[i][1] || '').trim();
        const hrs = Number(data[i][3]) || 0;
        if (!rd || !id || hrs <= 0) continue;
        if (from && rd < from) continue;
        if (to && rd > to) continue;
        if (!agg[id]) agg[id] = { name: String(data[i][2] || '').trim(), days: 0, hours: 0 };
        agg[id].days += 1;
        agg[id].hours += hrs;
      }
    }

    const rows = Object.keys(agg).map(id => {
      const hrs = Math.round(agg[id].hours * 100) / 100;
      const amt = KD(agg[id].hours * (rate[id] || 0));
      return { empId: id, name: agg[id].name, days: agg[id].days, hours: hrs, rate: KD(rate[id] || 0), amount: amt };
    }).sort((a, b) => a.empId.localeCompare(b.empId));

    const totalHours = Math.round(rows.reduce((s, r) => s + r.hours, 0) * 100) / 100;
    const totalAmount = KD(rows.reduce((s, r) => s + r.amount, 0));
    return { success: true, from: from ? formatDdMmYyyy_(from) : '', to: to ? formatDdMmYyyy_(to) : '',
      data: rows, totalHours: totalHours, totalAmount: totalAmount, employees: rows.length };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Write the OT summary to an 'OT Report <from>-<to>' tab.
// params: { from, to }
// ------------------------------------------------------------
function generateOtReport(params) {
  try {
    const summary = getOtSummary(params);
    if (!summary.success) return summary;
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    const label = (summary.from || 'start') + ' to ' + (summary.to || 'now');
    const safe = label.replace(/\//g, '-');
    const tabName = 'OT Report ' + safe;
    let out = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
    out.clear();

    const headers = ['Employee ID', 'Name', 'Days with OT', 'Total OT Hours', 'Price Per Hour', 'OT Amount'];
    const rows = [headers];
    summary.data.forEach(r => rows.push([r.empId, r.name, r.days, r.hours, r.rate, r.amount]));
    rows.push(['', 'TOTALS (' + summary.employees + ' staff)', '', summary.totalHours, '', summary.totalAmount]);

    out.getRange(1, 1, rows.length, headers.length).setValues(rows);
    out.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    out.getRange(rows.length, 1, 1, headers.length).setFontWeight('bold').setBackground('#f2f4f7');
    if (rows.length > 2) {
      out.getRange(2, 4, rows.length - 1, 1).setNumberFormat('0.00'); // hours
      [5, 6].forEach(c => out.getRange(2, c, rows.length - 1, 1).setNumberFormat('0.000')); // KD
    }
    out.setFrozenRows(1);
    out.autoResizeColumns(1, headers.length);

    return { success: true, tab: tabName, employees: summary.employees,
      totalHours: summary.totalHours, totalAmount: summary.totalAmount,
      message: 'Wrote "' + tabName + '": ' + summary.employees + ' staff, ' + summary.totalHours + ' OT hours, ' + summary.totalAmount + ' KD.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Used by Payroll.gs (without-attendance mode): total OT hours per
// employee whose entry date falls within [periodStart, periodEnd].
// Returns { empId: totalHours }
// ------------------------------------------------------------
function getOtHoursForPeriod_(ss, periodStart, periodEnd) {
  const out = {};
  const sheet = ss.getSheetByName(OT_ENTRIES_TAB);
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rd = parseFlexibleDate_(data[i][0]);
    const id = String(data[i][1] || '').trim();
    const hrs = Number(data[i][3]) || 0;
    if (!rd || !id || hrs <= 0) continue;
    if (rd < startOfDay_(periodStart) || rd > endOfDay_(periodEnd)) continue;
    out[id] = (out[id] || 0) + hrs;
  }
  return out;
}

// Price Per Hour by Employee ID from HR Maya.
function ratesById_(ss) {
  const out = {};
  const hr = ss.getSheetByName('HR Maya');
  if (!hr) return out;
  const d = hr.getDataRange().getValues();
  const H = d[0];
  const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
  const cRate = H.findIndex(h => String(h).trim().toLowerCase() === 'price per hour');
  for (let i = 1; i < d.length; i++) {
    const id = String(d[i][cId] || '').trim();
    if (id) out[id] = cRate !== -1 ? Number(d[i][cRate]) || 0 : 0;
  }
  return out;
}

// ============================================================
// ONE-SHOT SETUP — force-create every tab + the payslip folder.
// Run once from the editor (setupMayaHR) after pasting all files.
// Safe to re-run: only creates what's missing.
// ============================================================
function setupMayaHR() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const made = [];

  // Leave tabs (seeded rulebook + balances + log)
  ensureLeaveTabs();
  made.push('Leave_Types', 'Leave_Balances', 'Encashment_Log');

  // OT tab
  ensureOtTab_(ss);
  made.push('OT_Entries');

  // Payroll_Inputs
  if (!ss.getSheetByName(PAYROLL_INPUTS_TAB)) {
    const s = ss.insertSheet(PAYROLL_INPUTS_TAB);
    s.appendRow(['Employee ID', 'Name (English)', 'Period End', 'Deductions', 'Remark']);
    s.getRange(1, 1, 1, 5).setFontWeight('bold'); s.setFrozenRows(1);
    made.push(PAYROLL_INPUTS_TAB);
  }

  // Payroll_Control
  if (!ss.getSheetByName(PAYROLL_CONTROL_TAB)) {
    const s = ss.insertSheet(PAYROLL_CONTROL_TAB);
    s.appendRow(['Run Date', 'Period Start', 'Period End', 'Employees', 'Total OT Hours', 'Tab Name', 'Status']);
    s.getRange(1, 1, 1, 7).setFontWeight('bold'); s.setFrozenRows(1);
    made.push(PAYROLL_CONTROL_TAB);
  }

  // Attendance_Amendments
  if (!ss.getSheetByName('Attendance_Amendments')) {
    const s = ss.insertSheet('Attendance_Amendments');
    s.appendRow(['When', 'Admin', 'Employee ID', 'Employee Name', 'Date', 'Stage', 'Old Value', 'New Value']);
    s.getRange(1, 1, 1, 8).setFontWeight('bold'); s.setFrozenRows(1);
    made.push('Attendance_Amendments');
  }

  // Payslip Drive folder
  const folder = getOrCreateFolder_(PAYSLIP_FOLDER_NAME);

  const msg = 'Maya HR setup complete. Tabs ready: ' + made.join(', ') +
    '. Payslip folder: ' + folder.getName() + ' (' + folder.getUrl() + ').';
  Logger.log(msg);
  return { success: true, created: made, folderUrl: folder.getUrl(), message: msg };
}