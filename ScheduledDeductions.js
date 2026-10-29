// ============================================================
// ScheduledDeductions.gs — log deductions in advance; auto-applied by payroll
// ============================================================
// A Scheduled_Deductions tab: one row per planned deduction.
//   Employee ID · Name · Type · Amount · Month · Remark · Status · Applied In
// Status: Pending -> (payroll run in that month) -> Applied (locked).
// Month is a label like "Sep-2026". getPayrollReview reads Pending rows whose
// Month matches the run's period-end month, pre-fills them as itemised
// deductions, and runPayrollFromReview marks them Applied so they never repeat.
//
// ENDPOINTS: getScheduledDeductions, saveScheduledDeduction,
//            cancelScheduledDeduction
// ============================================================

const SCHED_DED_TAB = 'Scheduled_Deductions';

function ensureSchedDedTab_(ss) {
  let sheet = ss.getSheetByName(SCHED_DED_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SCHED_DED_TAB);
    sheet.appendRow(['ID', 'Employee ID', 'Name', 'Type', 'Amount', 'Month', 'Remark', 'Status', 'Applied In']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Month label from a Date, e.g. "Sep-2026".
function monthLabel_(d) { return Utilities.formatDate(d, tz_(), 'MMM-yyyy'); }

// ── THE MONTH CELL (fixed 22-09-2026) ───────────────────────────────
// The page sends "Sep-2026", but Google Sheets turns that text into a
// DATE (01-09-2026) the moment it is written. Every Month cell in the tab
// is a Date. The old match compared String(date) — "Tue Sep 01 2026 …" —
// with "sep-2026", so NO scheduled deduction ever matched a payroll run.
// Now both sides go through this one function, and the Month column is
// written as plain text so it stays "Sep-2026".
function schedMonthKey_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : monthLabel_(v).toLowerCase();
  const s = String(v).trim();
  const m = s.match(/^([A-Za-z]{3})[A-Za-z]*[-\s\/]+(\d{4})$/);
  if (m) return (m[1].charAt(0).toUpperCase() + m[1].slice(1, 3).toLowerCase() + '-' + m[2]).toLowerCase();
  const d = (typeof parseFlexibleDate_ === 'function') ? parseFlexibleDate_(s) : null;
  return d ? monthLabel_(d).toLowerCase() : s.toLowerCase();
}
// Display form of a Month cell: "Sep-2026".
function schedMonthLabel_(v) {
  const k = schedMonthKey_(v);
  return k ? k.charAt(0).toUpperCase() + k.slice(1) : '';
}
// Write a row with the Month column kept as text.
function schedWriteRow_(sheet, rowNum, values, monthCol) {
  sheet.getRange(rowNum, monthCol).setNumberFormat('@');
  sheet.getRange(rowNum, 1, 1, values.length).setValues([values]);
}

function getScheduledDeductions(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSchedDedTab_(ss);
    const d = sheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      if (!String(d[i][1] || '').trim()) continue;
      rows.push({ id: d[i][0], empId: d[i][1], name: d[i][2], type: d[i][3], amount: Number(d[i][4]) || 0,
        month: schedMonthLabel_(d[i][5]), remark: d[i][6], status: d[i][7], appliedIn: d[i][8] });
    }
    if (params.status) return { success: true, data: rows.filter(r => String(r.status).toLowerCase() === String(params.status).toLowerCase()) };
    return { success: true, data: rows };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Add or edit a scheduled deduction. Pending rows are editable; Applied are locked.
// params: { id?, empId, type, amount, month ("Sep-2026"), remark }
function saveScheduledDeduction(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSchedDedTab_(ss);
    const empId = String(params.empId || '').trim();
    const type = String(params.type || '').trim();
    const amount = Number(params.amount) || 0;
    const month = schedMonthLabel_(String(params.month || '').trim());
    if (!empId || !type || !amount || !month) return { success: false, error: 'empId, type, amount and month are all required.' };
    const name = empNameById_(ss, empId) || empId;

    const data = sheet.getDataRange().getValues();
    if (params.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(params.id)) {
          if (String(data[i][7]).toLowerCase() === 'applied') return { success: false, error: 'This deduction is already Applied and cannot be edited.' };
          schedWriteRow_(sheet, i + 1, [params.id, empId, name, type, amount, month, String(params.remark || ''), 'Pending', ''], 6);
          return { success: true, message: 'Updated scheduled deduction for ' + name + '.' };
        }
      }
    }
    const id = 'SD-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');
    schedWriteRow_(sheet, sheet.getLastRow() + 1, [id, empId, name, type, amount, month, String(params.remark || ''), 'Pending', ''], 6);
    return { success: true, id: id, message: 'Scheduled ' + amount.toFixed(3) + ' KD ' + type + ' for ' + name + ' in ' + month + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function cancelScheduledDeduction(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSchedDedTab_(ss);
    const d = sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]) === String(params.id)) {
        if (String(d[i][7]).toLowerCase() === 'applied') return { success: false, error: 'Already applied — cannot cancel.' };
        sheet.getRange(i + 1, 8).setValue('Cancelled');
        return { success: true, message: 'Cancelled.' };
      }
    }
    return { success: false, error: 'Not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Pending scheduled deductions due in a given period-end date's month.
// Returns { empId: { total, items:[{type,amount,remark,rowIndex,id}] } }
function getScheduledDeductionsForPeriod_(ss, periodEnd) {
  const out = {};
  const sheet = ss.getSheetByName(SCHED_DED_TAB);
  if (!sheet) return out;
  const wantMonth = schedMonthKey_(periodEnd);
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const empId = String(d[i][1] || '').trim();
    const status = String(d[i][7] || '').trim().toLowerCase();
    const month = schedMonthKey_(d[i][5]);
    if (!empId || status !== 'pending' || month !== wantMonth) continue;
    if (!out[empId]) out[empId] = { total: 0, items: [] };
    const amt = Number(d[i][4]) || 0;
    out[empId].total += amt;
    out[empId].items.push({ id: d[i][0], type: d[i][3], amount: amt, remark: d[i][6], rowIndex: i + 1 });
  }
  return out;
}

// Mark scheduled-deduction rows Applied after a payroll run consumes them.
function markScheduledDeductionsApplied_(ss, rowIndexes, tabName) {
  const sheet = ss.getSheetByName(SCHED_DED_TAB);
  if (!sheet) return;
  (rowIndexes || []).forEach(function (r) {
    sheet.getRange(r, 8).setValue('Applied');
    sheet.getRange(r, 9).setValue(tabName || formatDdMmYyyy_(new Date()));
  });
}