// ============================================================
// LeaveLifecycle.gs — Maya HR : grant leave + date-driven status flips
// ============================================================
// Builds on LeaveEncashment.gs (Leave_Types, Leave_Balances) and
// SharedUtils.gs date helpers. Do NOT redeclare those.
//
// FLOW:
//   grantLeave -> writes a dated approved leave to Leave_Grants and
//     decrements the leave balance (Taken += days, Remaining -= days).
//   applyLeaveStatusChanges (daily trigger) -> on a grant's START date,
//     auto-flips the employee to "On Leave" (so payroll excludes them);
//     after the END date, adds them to a REACTIVATION reminder (does NOT
//     auto-flip back — admin confirms). Only leave types flagged
//     "Excludes From Payroll" cause the On Leave flip.
//   getRemainingVacationTally -> reconciles HR Maya "Remaining Vacation"
//     against the Leave_Balances Annual figure, flags mismatches.
//
// ENDPOINTS: grantLeave, getLeaveGrants, applyLeaveStatusChanges,
//            getLeaveStatusReminders, getRemainingVacationTally
// ============================================================

const LEAVE_GRANTS_TAB = 'Leave_Grants';

function ensureLeaveGrantsTab_(ss) {
  let sheet = ss.getSheetByName(LEAVE_GRANTS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(LEAVE_GRANTS_TAB);
    sheet.appendRow(['Grant ID', 'Employee ID', 'Name', 'Type', 'Start', 'End', 'Days', 'Status', 'Granted By', 'Granted On']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Ensure Leave_Types has an "Excludes From Payroll" column; seed sensible
// defaults (Annual & Unpaid = Yes, others = No) when creating it.
function ensureExcludeColumn_(ss) {
  const types = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (!types) return;
  const H = types.getRange(1, 1, 1, types.getLastColumn()).getValues()[0];
  let idx = H.findIndex(h => String(h).trim().toLowerCase() === 'excludes from payroll');
  if (idx !== -1) return;
  const col = types.getLastColumn() + 1;
  types.getRange(1, col).setValue('Excludes From Payroll').setFontWeight('bold');
  const data = types.getDataRange().getValues();
  const cT = H.findIndex(h => String(h).trim().toLowerCase() === 'type');
  for (let i = 1; i < data.length; i++) {
    const t = String(data[i][cT] || '').trim().toLowerCase();
    const def = (t === 'annual' || t === 'unpaid') ? 'Yes' : 'No';
    types.getRange(i + 1, col).setValue(def);
  }
}

function excludesFromPayroll_(ss, type) {
  const types = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (!types) return (String(type).toLowerCase() === 'annual' || String(type).toLowerCase() === 'unpaid');
  const data = types.getDataRange().getValues();
  const H = data[0];
  const cT = H.findIndex(h => String(h).trim().toLowerCase() === 'type');
  const cE = H.findIndex(h => String(h).trim().toLowerCase() === 'excludes from payroll');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cT] || '').trim().toLowerCase() === String(type).trim().toLowerCase()) {
      if (cE === -1) { const t = String(type).toLowerCase(); return t === 'annual' || t === 'unpaid'; }
      return String(data[i][cE]).trim().toLowerCase() === 'yes';
    }
  }
  return false;
}

// ------------------------------------------------------------
// GRANT LEAVE — record a dated approved leave + decrement balance.
// params: { empId, type, start (dd-mm-yyyy), end (dd-mm-yyyy), days?, by? }
// ------------------------------------------------------------
function grantLeave(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTabs(); ensureExcludeColumn_(ss);
    const grants = ensureLeaveGrantsTab_(ss);

    const empId = String(params.empId || '').trim();
    const type = String(params.type || '').trim();
    if (!empId || !type) return { success: false, error: 'empId and type are required' };
    const start = parseFlexibleDate_(params.start);
    const end = parseFlexibleDate_(params.end);
    if (!start || !end) return { success: false, error: 'start and end must be valid dates (dd-mm-yyyy)' };
    if (endOfDay_(end) < startOfDay_(start)) return { success: false, error: 'End date is before start date.' };

    const name = empNameById_(ss, empId);
    if (!name) return { success: false, error: 'Employee ID "' + empId + '" not found in HR Maya' };

    const types = getLeaveTypesMap_(ss);
    if (!types[type.toLowerCase()]) return { success: false, error: 'Unknown leave type "' + type + '"' };

    // Inclusive calendar days unless overridden.
    let days = Number(params.days) || 0;
    if (days <= 0) days = Math.round((startOfDay_(end) - startOfDay_(start)) / 86400000) + 1;

    // Decrement balance for balance-tracked types (skip if entitlement 0, e.g. Unpaid).
    const entitlement = types[type.toLowerCase()].entitlement || 0;
    let newRemaining = null;
    if (entitlement > 0) {
      const balSheet = ss.getSheetByName(LEAVE_BALANCES_TAB);
      const rowNum = findBalanceRow_(balSheet, empId, type);
      if (rowNum === -1) return { success: false, error: 'No ' + type + ' balance for ' + name + '. Set a balance first.' };
      const b = balSheet.getRange(rowNum, 1, 1, 7).getValues()[0];
      const remaining = Number(b[6]) || 0;
      // Allow going negative (no block). Balance can drop below zero.
      newRemaining = remaining - days;
      balSheet.getRange(rowNum, 5).setValue((Number(b[4]) || 0) + days); // Taken
      balSheet.getRange(rowNum, 7).setValue(newRemaining);               // Remaining (may be negative)
    }

    const grantId = 'LG-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');
    grants.appendRow([grantId, empId, name, types[type.toLowerCase()].type, formatDdMmYyyy_(start), formatDdMmYyyy_(end), days, 'Approved', String(params.by || 'Dashboard').trim(), formatDdMmYyyy_(new Date())]);

    const willFlip = excludesFromPayroll_(ss, type);
    const negativeWarning = (newRemaining !== null && newRemaining < 0)
      ? ' ⚠ ' + name + ' is now at ' + newRemaining + ' ' + type + ' (negative balance).' : '';
    return {
      success: true, grantId: grantId, empId: empId, name: name, type: types[type.toLowerCase()].type,
      start: formatDdMmYyyy_(start), end: formatDdMmYyyy_(end), days: days, excludesFromPayroll: willFlip,
      newRemaining: newRemaining, negative: (newRemaining !== null && newRemaining < 0),
      message: 'Granted ' + days + ' day(s) of ' + type + ' to ' + name + ' (' + formatDdMmYyyy_(start) + ' to ' + formatDdMmYyyy_(end) + ').' + negativeWarning +
        (willFlip ? ' They will be set to On Leave (excluded from payroll) on ' + formatDdMmYyyy_(start) + '.' : ' Status unaffected (this type is paid in payroll).')
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function getLeaveGrants(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLeaveGrantsTab_(ss);
    const data = sheet.getDataRange().getValues();
    const filterId = String(params.empId || '').trim().toUpperCase();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][1] || '').trim();
      if (!id) continue;
      if (filterId && id.toUpperCase() !== filterId) continue;
      out.push({ grantId: data[i][0], empId: id, name: data[i][2], type: data[i][3], start: data[i][4], end: data[i][5], days: data[i][6], status: data[i][7], grantedBy: data[i][8], grantedOn: data[i][9] });
    }
    return { success: true, data: out };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// DAILY STATUS ENGINE — flip to On Leave on start date; list reactivations.
// Auto-flips Active -> On Leave when a payroll-excluding grant covers today.
// Does NOT auto-flip back; returns a reactivation list for the admin.
// Run on a daily time-based trigger, or manually / via URL.
// ------------------------------------------------------------
function applyLeaveStatusChanges(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grantsSheet = ensureLeaveGrantsTab_(ss);
    const today = startOfDay_(new Date());

    // Which employees SHOULD be on leave today (a payroll-excluding, approved grant covers today)?
    const shouldBeOnLeave = {}; const coveringGrant = {};
    const gdata = grantsSheet.getDataRange().getValues();
    for (let i = 1; i < gdata.length; i++) {
      const empId = String(gdata[i][1] || '').trim();
      const type = String(gdata[i][3] || '').trim();
      const status = String(gdata[i][7] || '').trim().toLowerCase();
      if (!empId || status !== 'approved') continue;
      const s = parseFlexibleDate_(gdata[i][4]), e = parseFlexibleDate_(gdata[i][5]);
      if (!s || !e) continue;
      if (today >= startOfDay_(s) && today <= endOfDay_(e) && excludesFromPayroll_(ss, type)) {
        shouldBeOnLeave[empId] = true;
        coveringGrant[empId] = { type: type, end: formatDdMmYyyy_(e) };
      }
    }

    // Compare with current HR Maya status.
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');
    if (cStatus === -1) return { success: false, error: 'HR Maya has no Status column.' };

    const flippedToLeave = []; const toReactivate = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      const name = cName !== -1 ? String(hd[i][cName] || '').trim() : id;
      const status = String(hd[i][cStatus] || '').trim().toLowerCase();

      if (shouldBeOnLeave[id] && status !== 'on leave') {
        hr.getRange(i + 1, cStatus + 1).setValue('On Leave');
        flippedToLeave.push({ id: id, name: name, until: coveringGrant[id].end, type: coveringGrant[id].type });
      } else if (!shouldBeOnLeave[id] && status === 'on leave') {
        toReactivate.push({ id: id, name: name });
      }
    }

    return {
      success: true, ranOn: formatDdMmYyyy_(new Date()),
      flippedToLeave: flippedToLeave, toReactivate: toReactivate,
      message: 'Set ' + flippedToLeave.length + ' employee(s) to On Leave. ' +
        toReactivate.length + ' whose leave has ended need reactivation (not auto-changed): ' +
        (toReactivate.length ? toReactivate.map(r => r.name + ' (' + r.id + ')').join(', ') : 'none') + '.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Read-only: who currently needs reactivating (for the dashboard/Grant page).
function getLeaveStatusReminders(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grantsSheet = ensureLeaveGrantsTab_(ss);
    const today = startOfDay_(new Date());
    const shouldBeOnLeave = {};
    const gdata = grantsSheet.getDataRange().getValues();
    for (let i = 1; i < gdata.length; i++) {
      const empId = String(gdata[i][1] || '').trim();
      const type = String(gdata[i][3] || '').trim();
      const status = String(gdata[i][7] || '').trim().toLowerCase();
      if (!empId || status !== 'approved') continue;
      const s = parseFlexibleDate_(gdata[i][4]), e = parseFlexibleDate_(gdata[i][5]);
      if (s && e && today >= startOfDay_(s) && today <= endOfDay_(e) && excludesFromPayroll_(ss, type)) shouldBeOnLeave[empId] = true;
    }
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');
    const toReactivate = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      if (String(hd[i][cStatus] || '').trim().toLowerCase() === 'on leave' && !shouldBeOnLeave[id]) {
        toReactivate.push({ id: id, name: cName !== -1 ? String(hd[i][cName] || '').trim() : id });
      }
    }
    return { success: true, toReactivate: toReactivate };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Reconcile HR Maya "Remaining Vacation" vs Leave_Balances (Annual).
// ------------------------------------------------------------
function getRemainingVacationTally(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cRem = H.findIndex(h => String(h).trim().toLowerCase() === 'remaining vacation');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');

    // Leave_Balances Annual remaining by empId.
    const balAnnual = {};
    const bal = ss.getSheetByName(LEAVE_BALANCES_TAB);
    if (bal) {
      const bd = bal.getDataRange().getValues();
      for (let i = 1; i < bd.length; i++) {
        if (String(bd[i][2] || '').trim().toLowerCase() === 'annual') balAnnual[String(bd[i][0] || '').trim()] = Number(bd[i][6]) || 0;
      }
    }

    const rows = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      const hrRem = cRem !== -1 ? (Number(hd[i][cRem]) || 0) : 0;
      const balRem = (id in balAnnual) ? balAnnual[id] : null;
      rows.push({
        empId: id, name: cName !== -1 ? String(hd[i][cName] || '').trim() : id,
        status: cStatus !== -1 ? String(hd[i][cStatus] || '').trim() : '',
        hrRemaining: hrRem, balanceRemaining: balRem,
        mismatch: (balRem !== null && Math.abs(balRem - hrRem) > 0.001)
      });
    }
    const mism = rows.filter(r => r.mismatch).length;
    return { success: true, data: rows, mismatches: mism,
      message: mism ? (mism + ' employee(s) have a mismatch between HR Maya Remaining Vacation and Leave_Balances.') : 'HR Maya and Leave_Balances agree on Annual remaining.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}