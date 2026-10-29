// ============================================================
// ScheduledExtras.gs — log bonuses / extras in advance; auto-applied by payroll
// ------------------------------------------------------------
// The mirror of Scheduled_Deductions, for money ADDED to pay.
// Nothing about who gets what is in the code: every extra is a row in the
// Scheduled_Extras tab, entered from the Scheduled Deductions & Extras page.
//
//   Scheduled_Extras:  ID · Employee ID · Name · Type · Amount · Month ·
//                      Remark · Status · Applied In
//   Status: Pending -> (payroll run for that month) -> Applied (locked).
//   Month:  "Sep-2026" (kept as text — see schedMonthKey_ in
//           ScheduledDeductions.gs, which both tabs share).
//
// getPayrollReview reads Pending rows whose Month matches the run's
// period-end month and adds them to Bonus/Additional; runPayrollFromReview
// marks them Applied so they are never paid twice.
//
// ENDPOINTS: getScheduledExtras, saveScheduledExtra, cancelScheduledExtra
// ============================================================

const SCHED_EXTRA_TAB = 'Scheduled_Extras';
const SCHED_EXTRA_HEADERS = ['ID', 'Employee ID', 'Name', 'Type', 'Amount', 'Month', 'Remark', 'Status', 'Applied In'];

function ensureSchedExtraTab_(ss) {
  let sheet = ss.getSheetByName(SCHED_EXTRA_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SCHED_EXTRA_TAB);
    sheet.appendRow(SCHED_EXTRA_HEADERS);
    sheet.getRange(1, 1, 1, SCHED_EXTRA_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1).setNumberFormat('@');   // Month stays text
  }
  return sheet;
}

// Column positions by header — never by fixed index.
function schedExtraCols_(sheet) {
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                 .map(h => String(h).trim().toLowerCase());
  const at = n => H.indexOf(n.toLowerCase());
  return { id: at('ID'), emp: at('Employee ID'), name: at('Name'), type: at('Type'), amount: at('Amount'),
           month: at('Month'), remark: at('Remark'), status: at('Status'), applied: at('Applied In'),
           width: H.length };
}

function getScheduledExtras(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSchedExtraTab_(ss);
    const c = schedExtraCols_(sheet);
    const d = sheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      if (!String(d[i][c.emp] || '').trim()) continue;
      rows.push({ id: d[i][c.id], empId: d[i][c.emp], name: d[i][c.name], type: d[i][c.type],
        amount: Number(d[i][c.amount]) || 0, month: schedMonthLabel_(d[i][c.month]),
        remark: d[i][c.remark], status: d[i][c.status], appliedIn: d[i][c.applied] });
    }
    // Types already used — the page offers them as suggestions; any new type is allowed.
    const types = {};
    rows.forEach(r => { const t = String(r.type || '').trim(); if (t) types[t] = true; });
    const data = params.status
      ? rows.filter(r => String(r.status).toLowerCase() === String(params.status).toLowerCase()) : rows;
    return { success: true, data: data, types: Object.keys(types).sort() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Add or edit a scheduled extra. Pending rows are editable; Applied are locked.
// params: { id?, empId, type, amount, month ("Sep-2026"), remark }
function saveScheduledExtra(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSchedExtraTab_(ss);
    const c = schedExtraCols_(sheet);
    const empId = String(params.empId || '').trim();
    const type = String(params.type || '').trim();
    const amount = Number(params.amount) || 0;
    const month = schedMonthLabel_(String(params.month || '').trim());
    if (!empId || !type || !(amount > 0) || !month)
      return { success: false, error: 'Employee, type, a positive amount and month are all required.' };
    const name = empNameById_(ss, empId) || empId;

    const build = (id) => {
      const row = new Array(c.width).fill('');
      row[c.id] = id; row[c.emp] = empId; row[c.name] = name; row[c.type] = type; row[c.amount] = amount;
      row[c.month] = month; row[c.remark] = String(params.remark || ''); row[c.status] = 'Pending'; row[c.applied] = '';
      return row;
    };
    const data = sheet.getDataRange().getValues();
    if (params.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][c.id]) === String(params.id)) {
          if (String(data[i][c.status]).toLowerCase() === 'applied')
            return { success: false, error: 'This extra is already Applied and cannot be edited.' };
          schedWriteRow_(sheet, i + 1, build(params.id), c.month + 1);
          return { success: true, message: 'Updated scheduled extra for ' + name + '.' };
        }
      }
    }
    const id = 'SE-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 1000);
    schedWriteRow_(sheet, sheet.getLastRow() + 1, build(id), c.month + 1);
    return { success: true, id: id, message: 'Scheduled ' + amount.toFixed(3) + ' KD ' + type + ' for ' + name + ' in ' + month + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function cancelScheduledExtra(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSchedExtraTab_(ss);
    const c = schedExtraCols_(sheet);
    const d = sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c.id]) === String(params.id)) {
        if (String(d[i][c.status]).toLowerCase() === 'applied') return { success: false, error: 'Already applied — cannot cancel.' };
        sheet.getRange(i + 1, c.status + 1).setValue('Cancelled');
        return { success: true, message: 'Cancelled.' };
      }
    }
    return { success: false, error: 'Not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Pending scheduled extras due in a given period-end date's month.
// Returns { empId: { total, items:[{id,type,amount,remark,rowIndex}] } }
function getScheduledExtrasForPeriod_(ss, periodEnd) {
  const out = {};
  const sheet = ss.getSheetByName(SCHED_EXTRA_TAB);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const c = schedExtraCols_(sheet);
  const wantMonth = schedMonthKey_(periodEnd);
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const empId = String(d[i][c.emp] || '').trim();
    const status = String(d[i][c.status] || '').trim().toLowerCase();
    if (!empId || status !== 'pending' || schedMonthKey_(d[i][c.month]) !== wantMonth) continue;
    const amt = Number(d[i][c.amount]) || 0;
    if (!out[empId]) out[empId] = { total: 0, items: [] };
    out[empId].total += amt;
    out[empId].items.push({ id: d[i][c.id], type: d[i][c.type], amount: amt, remark: d[i][c.remark], rowIndex: i + 1 });
  }
  return out;
}

// Mark scheduled-extra rows Applied after a payroll run pays them.
function markScheduledExtrasApplied_(ss, rowIndexes, tabName) {
  const sheet = ss.getSheetByName(SCHED_EXTRA_TAB);
  if (!sheet) return;
  const c = schedExtraCols_(sheet);
  (rowIndexes || []).forEach(function (r) {
    sheet.getRange(r, c.status + 1).setValue('Applied');
    sheet.getRange(r, c.applied + 1).setValue(tabName || formatDdMmYyyy_(new Date()));
  });
}