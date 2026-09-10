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
        month: d[i][5], remark: d[i][6], status: d[i][7], appliedIn: d[i][8] });
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
    const month = String(params.month || '').trim();
    if (!empId || !type || !amount || !month) return { success: false, error: 'empId, type, amount and month are all required.' };
    const name = empNameById_(ss, empId) || empId;

    const data = sheet.getDataRange().getValues();
    if (params.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(params.id)) {
          if (String(data[i][7]).toLowerCase() === 'applied') return { success: false, error: 'This deduction is already Applied and cannot be edited.' };
          sheet.getRange(i + 1, 1, 1, 9).setValues([[params.id, empId, name, type, amount, month, String(params.remark || ''), 'Pending', '']]);
          return { success: true, message: 'Updated scheduled deduction for ' + name + '.' };
        }
      }
    }
    const id = 'SD-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');
    sheet.appendRow([id, empId, name, type, amount, month, String(params.remark || ''), 'Pending', '']);
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
  const wantMonth = monthLabel_(periodEnd).toLowerCase();
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const empId = String(d[i][1] || '').trim();
    const status = String(d[i][7] || '').trim().toLowerCase();
    const month = String(d[i][5] || '').trim().toLowerCase();
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