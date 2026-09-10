// ============================================================
// LeaveEncashment.gs — Maya HR : typed leave balances + encashment
// ============================================================
// Standalone Apps Script file. Shares global scope with Code.gs,
// Payroll.gs, PayrollDocs.gs, SharedUtils.gs. Reuses KD (Payroll.gs)
// and date helpers (SharedUtils.gs). Do NOT redeclare them.
//
// All leave data lives in the HR Maya sheet (MAIN_SPREADSHEET_ID),
// keyed by Employee ID — the old Vacation / Sick Leave tabs in the
// separate payroll spreadsheet are retired (migrate once, then ignore).
//
// TABS THIS FILE OWNS:
//   Leave_Types      Type | Can Encash | Can Carry Forward | Annual Entitlement | Notes
//   Leave_Balances   Employee ID | Name | Type | Entitled | Taken | Encashed | Remaining
//   Encashment_Log   Employee ID | Name | Type | Days | Amount | Period End | Encash Date | By
//
// ENDPOINTS (added to Code.gs handleRequest):
//   ensureLeaveTabs, getLeaveTypes, getLeaveBalances, upsertLeaveBalance,
//   bulkUploadLeaveBalances, encashLeave, migrateLeaveBalances
// ============================================================

const LEAVE_TYPES_TAB    = 'Leave_Types';
const LEAVE_BALANCES_TAB = 'Leave_Balances';
const ENCASHMENT_LOG_TAB = 'Encashment_Log';

// ------------------------------------------------------------
// Tab setup + seeded rulebook
// ------------------------------------------------------------
function ensureLeaveTabs(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    let types = ss.getSheetByName(LEAVE_TYPES_TAB);
    if (!types) {
      types = ss.insertSheet(LEAVE_TYPES_TAB);
      types.appendRow(['Type', 'Can Encash', 'Can Carry Forward', 'Annual Entitlement', 'Notes']);
      types.getRange(1, 1, 1, 5).setFontWeight('bold');
      types.setFrozenRows(1);
      // Seeded defaults — editable by HR afterwards.
      types.getRange(2, 1, 4, 5).setValues([
        ['Annual',    'Yes', 'Yes', 30, 'Paid annual leave; encashable and carries forward'],
        ['Sick',      'No',  'No',  15, 'Cannot be encashed or carried forward'],
        ['Emergency', 'No',  'No',  0,  'As approved; not encashable'],
        ['Unpaid',    'No',  'No',  0,  'Unpaid; no balance value']
      ]);
      types.autoResizeColumns(1, 5);
    }

    let bal = ss.getSheetByName(LEAVE_BALANCES_TAB);
    if (!bal) {
      bal = ss.insertSheet(LEAVE_BALANCES_TAB);
      bal.appendRow(['Employee ID', 'Name', 'Type', 'Entitled', 'Taken', 'Encashed', 'Remaining']);
      bal.getRange(1, 1, 1, 7).setFontWeight('bold');
      bal.setFrozenRows(1);
    }

    let log = ss.getSheetByName(ENCASHMENT_LOG_TAB);
    if (!log) {
      log = ss.insertSheet(ENCASHMENT_LOG_TAB);
      log.appendRow(['Employee ID', 'Name', 'Type', 'Days', 'Amount', 'Period End', 'Encash Date', 'By', 'Voucher No']);
      log.getRange(1, 1, 1, 8).setFontWeight('bold');
      log.setFrozenRows(1);
    }

    return { success: true, message: 'Leave tabs ready (Leave_Types, Leave_Balances, Encashment_Log).' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ------------------------------------------------------------
// Rulebook reader — { TypeName: {canEncash, canCarry, entitlement} }
// ------------------------------------------------------------
function getLeaveTypesMap_(ss) {
  const out = {};
  const sheet = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  const H = data[0];
  const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const cT = c('Type'), cE = c('Can Encash'), cF = c('Can Carry Forward'), cA = c('Annual Entitlement');
  const yes = (v) => String(v).trim().toLowerCase() === 'yes' || String(v).trim().toLowerCase() === 'true';
  for (let i = 1; i < data.length; i++) {
    const t = String(data[i][cT] || '').trim();
    if (!t) continue;
    out[t.toLowerCase()] = {
      type: t,
      canEncash: cE !== -1 ? yes(data[i][cE]) : false,
      canCarry:  cF !== -1 ? yes(data[i][cF]) : false,
      entitlement: cA !== -1 ? Number(data[i][cA]) || 0 : 0
    };
  }
  return out;
}

function getLeaveTypes(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTabs();
    const map = getLeaveTypesMap_(ss);
    return { success: true, data: Object.keys(map).map(k => map[k]) };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Balances reader
// ------------------------------------------------------------
function getLeaveBalances(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTabs();
    const sheet = ss.getSheetByName(LEAVE_BALANCES_TAB);
    const data = sheet.getDataRange().getValues();
    const filterId = String(params.empId || '').trim().toUpperCase();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0] || '').trim();
      if (!id) continue;
      if (filterId && id.toUpperCase() !== filterId) continue;
      out.push({
        empId: id, name: String(data[i][1] || '').trim(), type: String(data[i][2] || '').trim(),
        entitled: Number(data[i][3]) || 0, taken: Number(data[i][4]) || 0,
        encashed: Number(data[i][5]) || 0, remaining: Number(data[i][6]) || 0
      });
    }
    const types = getLeaveTypesMap_(ss);
    return { success: true, data: out, types: Object.keys(types).map(k => types[k]) };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Locate a balance row (1-indexed sheet row) for empId+type, or -1.
function findBalanceRow_(sheet, empId, type) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === empId.toUpperCase() &&
        String(data[i][2]).trim().toLowerCase() === type.toLowerCase()) {
      return i + 1;
    }
  }
  return -1;
}

function empNameById_(ss, empId) {
  const hr = ss.getSheetByName('HR Maya');
  if (!hr) return '';
  const d = hr.getDataRange().getValues();
  const H = d[0];
  const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
  const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
  for (let i = 1; i < d.length; i++) {
    if (cId !== -1 && String(d[i][cId]).trim().toUpperCase() === empId.toUpperCase()) {
      return cName !== -1 ? String(d[i][cName]).trim() : '';
    }
  }
  return '';
}

// ------------------------------------------------------------
// Add / edit one balance row (Remaining auto = Entitled - Taken - Encashed)
// params: { empId, type, entitled, taken?, encashed? }
// ------------------------------------------------------------
function upsertLeaveBalance(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTabs();
    const empId = String(params.empId || '').trim();
    const type  = String(params.type || '').trim();
    if (!empId || !type) return { success: false, error: 'empId and type are required' };

    const types = getLeaveTypesMap_(ss);
    if (!types[type.toLowerCase()]) return { success: false, error: 'Unknown leave type "' + type + '". Add it to Leave_Types first.' };

    const name = empNameById_(ss, empId);
    if (!name) return { success: false, error: 'Employee ID "' + empId + '" not found in HR Maya' };

    const sheet = ss.getSheetByName(LEAVE_BALANCES_TAB);
    const rowNum = findBalanceRow_(sheet, empId, type);
    const existing = rowNum !== -1 ? sheet.getRange(rowNum, 1, 1, 7).getValues()[0] : null;

    const entitled = params.entitled !== undefined && params.entitled !== '' ? Number(params.entitled) || 0 : (existing ? Number(existing[3]) || 0 : (types[type.toLowerCase()].entitlement || 0));
    const taken    = params.taken !== undefined && params.taken !== '' ? Number(params.taken) || 0 : (existing ? Number(existing[4]) || 0 : 0);
    const encashed = params.encashed !== undefined && params.encashed !== '' ? Number(params.encashed) || 0 : (existing ? Number(existing[5]) || 0 : 0);
    const remaining = entitled - taken - encashed;

    const row = [empId, name, types[type.toLowerCase()].type, entitled, taken, encashed, remaining];
    if (rowNum !== -1) sheet.getRange(rowNum, 1, 1, 7).setValues([row]);
    else sheet.appendRow(row);
    sheet.autoResizeColumns(1, 7);

    return { success: true, message: (rowNum !== -1 ? 'Updated' : 'Added') + ' ' + type + ' balance for ' + name + ' (Remaining ' + remaining + ').' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Bulk upload — rows: [{empId, type, entitled, taken?, encashed?}, ...]
// Sent as JSON in POST body (e.parameter.payload or postData).
// ------------------------------------------------------------
function bulkUploadLeaveBalances(params) {
  try {
    const raw = params && params.rows ? params.rows : '[]';
    const rows = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (!Array.isArray(rows) || !rows.length) return { success: false, error: 'No rows provided' };

    let ok = 0; const errors = [];
    rows.forEach((r, i) => {
      const res = upsertLeaveBalance(r);
      if (res.success) ok++; else errors.push('Row ' + (i + 1) + ' (' + (r.empId || '?') + '/' + (r.type || '?') + '): ' + res.error);
    });
    return { success: true, uploaded: ok, failed: errors.length, errors: errors,
      message: 'Uploaded ' + ok + ' balance rows' + (errors.length ? ', ' + errors.length + ' failed.' : '.') };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// ENCASH LEAVE
// params: { empId, type, days, encashDate (dd-mm-yyyy), by? }
// - blocked unless the type's Can Encash = Yes
// - blocked if days > Remaining
// - amount = Basic/30 * days (month=30 days) ; adds to Encashed, subtracts Remaining
// - logs to Encashment_Log with the Period End it will be paid in
//   (first run whose Period End is on/after encashDate)
// ------------------------------------------------------------
function encashLeave(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTabs();

    const empId = String(params.empId || '').trim();
    const type  = String(params.type || '').trim();
    const days  = Number(params.days) || 0;
    const by    = String(params.by || 'Dashboard').trim();
    if (!empId || !type) return { success: false, error: 'empId and type are required' };
    if (days <= 0) return { success: false, error: 'Days to encash must be greater than 0' };

    const encashDate = parseFlexibleDate_(params.encashDate) || new Date();

    const types = getLeaveTypesMap_(ss);
    const rule = types[type.toLowerCase()];
    if (!rule) return { success: false, error: 'Unknown leave type "' + type + '"' };
    if (!rule.canEncash) return { success: false, error: type + ' leave cannot be encashed (per Leave_Types rules).' };

    const balSheet = ss.getSheetByName(LEAVE_BALANCES_TAB);
    const rowNum = findBalanceRow_(balSheet, empId, type);
    if (rowNum === -1) return { success: false, error: 'No ' + type + ' balance found for ' + empId };
    const b = balSheet.getRange(rowNum, 1, 1, 7).getValues()[0];
    const name = String(b[1] || '').trim();
    const remaining = Number(b[6]) || 0;
    if (days > remaining) return { success: false, error: 'Cannot encash ' + days + ' days — only ' + remaining + ' remaining for ' + name + ' (' + type + ').' };

    // Encashment rate = Basic / 30 (a month = 30 days), so 30 days = 1 basic,
    // 60 days = 2 basics. NOTE: this 30-day divisor is ONLY for leave
    // encashment; OT, unpaid leave, absence and half-day all use Basic/26.
    const basic = basicById_(ss, empId);
    if (basic <= 0) return { success: false, error: 'No basic salary found for ' + empId };
    const amount = KD(basic / 30 * days);

    // Which payroll period does this attach to? First Period End on/after encashDate.
    const periodEnd = resolveEncashPeriodEnd_(ss, encashDate);

    // Apply to balance: Encashed += days, Remaining -= days
    const newEncashed = (Number(b[5]) || 0) + days;
    const newRemaining = remaining - days;
    balSheet.getRange(rowNum, 6).setValue(newEncashed);
    balSheet.getRange(rowNum, 7).setValue(newRemaining);

    // Log it (with an auto Voucher No so a printable voucher can be generated).
    const voucherNo = 'ENC-' + Utilities.formatDate(new Date(), tz_(), 'yyyy') + '-' + Utilities.formatDate(new Date(), tz_(), 'MMddHHmmss');
    // Ensure the log has a "Voucher No" column (upgrade old 8-col tabs).
    const logSheet = ss.getSheetByName(ENCASHMENT_LOG_TAB);
    const logH = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    if (logH.findIndex(h => String(h).trim().toLowerCase() === 'voucher no') === -1) {
      logSheet.getRange(1, logSheet.getLastColumn() + 1).setValue('Voucher No').setFontWeight('bold');
    }
    logSheet.appendRow([
      empId, name, rule.type, days, amount,
      periodEnd ? formatDdMmYyyy_(periodEnd) : '(next run)',
      formatDdMmYyyy_(encashDate), by, voucherNo
    ]);

    return {
      success: true, empId: empId, name: name, type: rule.type, days: days, amount: amount,
      periodEnd: periodEnd ? formatDdMmYyyy_(periodEnd) : '',
      remaining: newRemaining, voucherNo: voucherNo,
      message: 'Encashed ' + days + ' ' + rule.type + ' day(s) for ' + name + ' = ' + amount +
        ' KD. Remaining ' + newRemaining + '. Will be paid in the run ending ' +
        (periodEnd ? formatDdMmYyyy_(periodEnd) : 'on/after ' + formatDdMmYyyy_(encashDate)) + '.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function basicById_(ss, empId) {
  const hr = ss.getSheetByName('HR Maya');
  if (!hr) return 0;
  const d = hr.getDataRange().getValues();
  const H = d[0];
  const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
  const cBasic = H.findIndex(h => String(h).trim().toLowerCase() === 'basic salary');
  for (let i = 1; i < d.length; i++) {
    if (cId !== -1 && String(d[i][cId]).trim().toUpperCase() === empId.toUpperCase()) {
      return cBasic !== -1 ? Number(d[i][cBasic]) || 0 : 0;
    }
  }
  return 0;
}

// First Payroll_Control Period End on/after the encash date; else null (unpaid yet).
function resolveEncashPeriodEnd_(ss, encashDate) {
  const sheet = ss.getSheetByName(PAYROLL_CONTROL_TAB);
  const target = startOfDay_(encashDate);
  let best = null;
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const pe = parseFlexibleDate_(data[i][2]);
      if (!pe) continue;
      const peEnd = endOfDay_(pe);
      if (peEnd >= target && (!best || pe < best)) best = pe;
    }
  }
  return best; // may be null: no run yet covers it — encashLog notes "(next run)"
}

// ------------------------------------------------------------
// Encashment total for a given Period End (used by payroll), per employee.
// Returns { empId: {days, amount, types:[...]} }
// ------------------------------------------------------------
function getEncashmentsForPeriod_(ss, periodEndKey) {
  const out = {};
  const sheet = ss.getSheetByName(ENCASHMENT_LOG_TAB);
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  const H = data[0];
  const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const cId = c('Employee ID'), cType = c('Type'), cDays = c('Days'), cAmt = c('Amount'), cPer = c('Period End');
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][cId] || '').trim();
    if (!id) continue;
    if (formatKey_(data[i][cPer]) !== periodEndKey) continue;
    if (!out[id]) out[id] = { days: 0, amount: 0, types: [] };
    out[id].days += Number(data[i][cDays]) || 0;
    out[id].amount += Number(data[i][cAmt]) || 0;
    out[id].types.push(String(data[i][cType] || '').trim());
  }
  Object.keys(out).forEach(id => { out[id].amount = KD(out[id].amount); });
  return out;
}

// ------------------------------------------------------------
// MIGRATION (preview then commit) from the old Vacation / Sick Leave tabs
// (in PAYROLL_SPREADSHEET_ID), matching names -> Employee IDs in HR Maya.
// params: { commit?: 'true' }
// ------------------------------------------------------------
function migrateLeaveBalances(params) {
  try {
    params = params || {};
    const commit = String(params.commit || '').toLowerCase() === 'true';
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTabs();

    // Build name -> {id, name} from HR Maya
    const hr = ss.getSheetByName('HR Maya');
    const d = hr.getDataRange().getValues();
    const H = d[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const byName = {};
    for (let i = 1; i < d.length; i++) {
      const nm = String(d[i][cName] || '').trim().toLowerCase();
      if (nm) byName[nm] = { id: String(d[i][cId] || '').trim(), name: String(d[i][cName] || '').trim() };
    }

    const matched = [], unmatched = [];
    const src = SpreadsheetApp.openById(PAYROLL_SPREADSHEET_ID);

    function pull(tabName, typeName, remainingHeaderCandidates) {
      const sheet = src.getSheetByName(tabName);
      if (!sheet) return;
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return;
      const SH = rows[0];
      const findCol = (cands) => {
        for (const cand of cands) {
          const idx = SH.findIndex(h => String(h).trim().toLowerCase().includes(cand.toLowerCase()));
          if (idx !== -1) return idx;
        }
        return -1;
      };
      const nCol = findCol(['name in english', 'english', 'name']);
      const rCol = findCol(remainingHeaderCandidates);
      for (let i = 1; i < rows.length; i++) {
        const nm = String(rows[i][nCol] || '').trim();
        if (!nm) continue;
        const remaining = rCol !== -1 ? Number(rows[i][rCol]) || 0 : 0;
        const hit = byName[nm.toLowerCase()];
        if (hit && hit.id) matched.push({ empId: hit.id, name: hit.name, type: typeName, remaining: remaining });
        else unmatched.push({ name: nm, type: typeName, remaining: remaining });
      }
    }

    pull('Vacation', 'Annual', ['remaining vacation', 'remaining', 'left']);
    pull('Sick Leave', 'Sick', ['remaining', 'balance', 'left']);

    if (!commit) {
      return { success: true, preview: true, matchedCount: matched.length, unmatchedCount: unmatched.length,
        matched: matched, unmatched: unmatched,
        message: 'PREVIEW: ' + matched.length + ' matched, ' + unmatched.length + ' unmatched. ' +
          'Re-run with commit=true to write the matched rows into Leave_Balances.' };
    }

    // Commit: write matched as balances (Entitled defaults from rulebook, Taken=0, Encashed=0, Remaining=migrated)
    const types = getLeaveTypesMap_(ss);
    let written = 0;
    matched.forEach(m => {
      const entitlement = types[m.type.toLowerCase()] ? types[m.type.toLowerCase()].entitlement : m.remaining;
      const res = upsertLeaveBalance({ empId: m.empId, type: m.type, entitled: entitlement, taken: Math.max(0, entitlement - m.remaining), encashed: 0 });
      if (res.success) written++;
    });
    return { success: true, preview: false, written: written, unmatchedCount: unmatched.length, unmatched: unmatched,
      message: 'Migrated ' + written + ' balances into Leave_Balances. ' + unmatched.length + ' unmatched (resolve manually).' };
  } catch (err) { return { success: false, error: err.toString() }; }
}