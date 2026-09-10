// ============================================================
// LeaveTypes.gs — Maya HR : Leave Type management + Allocation
// ============================================================
// Extends the Leave_Types rulebook (LeaveEncashment.gs) to the full
// attribute set and adds CRUD + allocation, so leave types are DATA,
// never hard-coded. The "Is Paid" flag drives payroll: Unpaid types are
// deducted (Basic/26 x days) in the run.
//
// Leave_Types columns:
//   Type | Can Encash | Can Carry Forward | Annual Entitlement | Notes |
//   Excludes From Payroll | Is Paid | Active
//
// ENDPOINTS: getLeaveTypesFull, saveLeaveType, deleteLeaveType,
//            allocateLeave, migrateRemainingToAnnual
// ============================================================

function ensureLeaveTypeAttrs_(ss) {
  ensureLeaveTabs();
  const sheet = ss.getSheetByName(LEAVE_TYPES_TAB);
  let H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const has = (name) => H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase()) !== -1;
  const addCol = (name, filler) => {
    const col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue(name).setFontWeight('bold');
    const n = sheet.getLastRow() - 1;
    if (n > 0) {
      const cT = H.findIndex(h => String(h).trim().toLowerCase() === 'type');
      const types = sheet.getRange(2, cT + 1, n, 1).getValues();
      const vals = types.map(r => [filler(String(r[0] || '').trim())]);
      sheet.getRange(2, col, n, 1).setValues(vals);
    }
    H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  };
  if (!has('Excludes From Payroll')) addCol('Excludes From Payroll', t => (/annual|unpaid/i.test(t) ? 'Yes' : 'No'));
  if (!has('Is Paid')) addCol('Is Paid', t => (/unpaid/i.test(t) ? 'No' : 'Yes'));
  if (!has('Active')) addCol('Active', () => 'Yes');
  return sheet;
}

function yesNo_(v) { const s = String(v).trim().toLowerCase(); return (s === 'yes' || s === 'true' || s === '1') ? 'Yes' : 'No'; }
function isYes_(v) { const s = String(v).trim().toLowerCase(); return s === 'yes' || s === 'true' || s === '1'; }

function getLeaveTypesFull(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLeaveTypeAttrs_(ss);
    const data = sheet.getDataRange().getValues();
    const H = data[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cT = c('Type'), cEnc = c('Can Encash'), cCar = c('Can Carry Forward'), cEnt = c('Annual Entitlement'),
          cNote = c('Notes'), cExc = c('Excludes From Payroll'), cPaid = c('Is Paid'), cAct = c('Active');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const t = String(data[i][cT] || '').trim();
      if (!t) continue;
      out.push({
        type: t,
        canEncash: cEnc !== -1 ? isYes_(data[i][cEnc]) : false,
        canCarry: cCar !== -1 ? isYes_(data[i][cCar]) : false,
        entitlement: cEnt !== -1 ? Number(data[i][cEnt]) || 0 : 0,
        notes: cNote !== -1 ? String(data[i][cNote] || '') : '',
        excludesFromPayroll: cExc !== -1 ? isYes_(data[i][cExc]) : false,
        isPaid: cPaid !== -1 ? isYes_(data[i][cPaid]) : true,
        active: cAct !== -1 ? isYes_(data[i][cAct]) : true
      });
    }
    return { success: true, data: out };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Add or edit a leave type by name. Optionally allocate its entitlement to all.
// params: name, isPaid, entitlement, canEncash, canCarry, excludesFromPayroll, active, notes, allocateAll
function saveLeaveType(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLeaveTypeAttrs_(ss);
    const name = String(params.name || '').trim();
    if (!name) return { success: false, error: 'Leave type name is required.' };

    const data = sheet.getDataRange().getValues();
    const H = data[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cT = c('Type'), cEnc = c('Can Encash'), cCar = c('Can Carry Forward'), cEnt = c('Annual Entitlement'),
          cNote = c('Notes'), cExc = c('Excludes From Payroll'), cPaid = c('Is Paid'), cAct = c('Active');

    let rowNum = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cT] || '').trim().toLowerCase() === name.toLowerCase()) { rowNum = i + 1; break; }
    }
    const width = H.length;
    const row = rowNum === -1 ? new Array(width).fill('') : data[rowNum - 1].slice();
    const put = (idx, val) => { if (idx !== -1) row[idx] = val; };
    const prov = (k) => params[k] !== undefined && params[k] !== '';

    put(cT, name);
    if (rowNum === -1 || prov('canEncash')) put(cEnc, yesNo_(params.canEncash));
    if (rowNum === -1 || prov('canCarry')) put(cCar, yesNo_(params.canCarry));
    if (rowNum === -1 || prov('entitlement')) put(cEnt, Number(params.entitlement) || 0);
    if (rowNum === -1 || prov('notes')) put(cNote, String(params.notes || ''));
    if (rowNum === -1 || prov('excludesFromPayroll')) put(cExc, yesNo_(params.excludesFromPayroll));
    if (rowNum === -1 || prov('isPaid')) put(cPaid, params.isPaid === undefined ? 'Yes' : yesNo_(params.isPaid));
    if (rowNum === -1 || prov('active')) put(cAct, params.active === undefined ? 'Yes' : yesNo_(params.active));

    if (rowNum === -1) sheet.appendRow(row);
    else sheet.getRange(rowNum, 1, 1, width).setValues([row]);

    let allocated = 0;
    const ent = Number(params.entitlement) || 0;
    if (isYes_(params.allocateAll) && ent > 0) {
      const res = allocateLeave({ type: name, scope: 'all', days: ent, mode: 'set' });
      allocated = res.allocated || 0;
    }
    return { success: true, type: name, allocated: allocated,
      message: (rowNum === -1 ? 'Created' : 'Updated') + ' leave type "' + name + '"' + (allocated ? (', allocated ' + ent + ' days to ' + allocated + ' employees') : '') + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Soft-delete (Active=No) by default; hard=true removes the row.
function deleteLeaveType(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLeaveTypeAttrs_(ss);
    const name = String(params.name || '').trim();
    if (!name) return { success: false, error: 'name required' };
    const data = sheet.getDataRange().getValues();
    const H = data[0];
    const cT = H.findIndex(h => String(h).trim().toLowerCase() === 'type');
    const cAct = H.findIndex(h => String(h).trim().toLowerCase() === 'active');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cT] || '').trim().toLowerCase() === name.toLowerCase()) {
        if (isYes_(params.hard)) { sheet.deleteRow(i + 1); return { success: true, message: 'Deleted leave type "' + name + '".' }; }
        sheet.getRange(i + 1, cAct + 1).setValue('No');
        return { success: true, message: 'Retired leave type "' + name + '" (Active = No). History preserved.' };
      }
    }
    return { success: false, error: 'Leave type "' + name + '" not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// ALLOCATION — give employees a balance of a type.
// params: type, days, mode ('set'|'topup'), scope ('all'|'branch'|'ids'),
//         branch?, ids? (comma-separated)
// ------------------------------------------------------------
function allocateLeave(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTypeAttrs_(ss);
    const type = String(params.type || '').trim();
    const days = Number(params.days) || 0;
    const mode = String(params.mode || 'set').toLowerCase();
    const scope = String(params.scope || 'all').toLowerCase();
    if (!type) return { success: false, error: 'type required' };

    const typesMap = getLeaveTypesMap_(ss);
    if (!typesMap[type.toLowerCase()]) return { success: false, error: 'Unknown leave type "' + type + '".' };
    const canonicalType = typesMap[type.toLowerCase()].type;

    // Build the target employee set from HR Maya.
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    let cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'location'); if (cBranch === -1) cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const idSet = scope === 'ids'
      ? String(params.ids || '').split(',').map(s => s.trim().toUpperCase()).filter(Boolean)
      : null;
    const branch = scope === 'branch' ? String(params.branch || '').trim().toLowerCase() : null;

    const bal = ss.getSheetByName(LEAVE_BALANCES_TAB);
    let allocated = 0;
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      if (idSet && idSet.indexOf(id.toUpperCase()) === -1) continue;
      if (branch && String(hd[i][cBranch] || '').trim().toLowerCase() !== branch) continue;
      const name = cName !== -1 ? String(hd[i][cName] || '').trim() : id;
      allocateOne_(bal, id, name, canonicalType, days, mode);
      allocated++;
    }
    return { success: true, type: canonicalType, days: days, mode: mode, allocated: allocated,
      message: (mode === 'topup' ? 'Added ' : 'Set ') + days + ' day(s) of ' + canonicalType + ' for ' + allocated + ' employee(s).' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Upsert one Leave_Balances row. mode 'set' overwrites Entitled; 'topup' adds.
function allocateOne_(bal, empId, name, type, days, mode) {
  const rowNum = findBalanceRow_(bal, empId, type);
  if (rowNum === -1) {
    // new row: Entitled=days, Taken=0, Encashed=0, Remaining=days
    bal.appendRow([empId, name, type, days, 0, 0, days]);
    return;
  }
  const r = bal.getRange(rowNum, 1, 1, 7).getValues()[0];
  const taken = Number(r[4]) || 0, enc = Number(r[5]) || 0;
  let entitled;
  if (mode === 'topup') entitled = (Number(r[3]) || 0) + days;
  else entitled = days; // 'set'
  bal.getRange(rowNum, 4).setValue(entitled);
  bal.getRange(rowNum, 7).setValue(entitled - taken - enc); // Remaining
}

// ------------------------------------------------------------
// ONE-SHOT MIGRATION: HR Maya "Remaining Vacation" -> Annual balance.
// preview=true reports without writing.
// ------------------------------------------------------------
function migrateRemainingToAnnual(params) {
  try {
    params = params || {};
    const preview = isYes_(params.preview);
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTypeAttrs_(ss);
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cRem = H.findIndex(h => String(h).trim().toLowerCase() === 'remaining vacation');
    if (cRem === -1) return { success: false, error: 'HR Maya has no "Remaining Vacation" column.' };

    const bal = ss.getSheetByName(LEAVE_BALANCES_TAB);
    let migrated = 0; const rows = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      const name = cName !== -1 ? String(hd[i][cName] || '').trim() : id;
      const rem = Number(hd[i][cRem]) || 0;
      rows.push({ empId: id, name: name, remaining: rem });
      if (!preview) { allocateOne_(bal, id, name, 'Annual', rem, 'set'); migrated++; }
    }
    return { success: true, preview: preview, migrated: preview ? rows.length : migrated, sample: rows.slice(0, 8),
      message: (preview ? 'PREVIEW — ' : '') + 'Annual balance set from HR Maya Remaining Vacation for ' + (preview ? rows.length : migrated) + ' employee(s).' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// All employees x all types: balance matrix for the Grant page, so the
// per-row type dropdown can switch instantly without a round-trip.
// Returns { employees:[{id,name,status,branch}], balances:{empId:{type:remaining}} }
// ------------------------------------------------------------
function getGrantMatrix(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTypeAttrs_(ss);

    // Employees from HR Maya
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');
    let cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'location'); if (cBranch === -1) cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const employees = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      employees.push({
        id: id,
        name: cName !== -1 ? String(hd[i][cName] || '').trim() : id,
        status: cStatus !== -1 ? String(hd[i][cStatus] || '').trim() : '',
        branch: cBranch !== -1 ? String(hd[i][cBranch] || '').trim() : ''
      });
    }

    // Balances: empId -> { type(lowercase) : remaining }
    const balances = {};
    const bal = ss.getSheetByName(LEAVE_BALANCES_TAB);
    if (bal) {
      const bd = bal.getDataRange().getValues();
      for (let i = 1; i < bd.length; i++) {
        const id = String(bd[i][0] || '').trim();
        const type = String(bd[i][2] || '').trim();
        if (!id || !type) continue;
        if (!balances[id]) balances[id] = {};
        balances[id][type.toLowerCase()] = Number(bd[i][6]) || 0; // Remaining
      }
    }

    // Active types (for the dropdown)
    const typesRes = getLeaveTypesFull({});
    const types = (typesRes.data || []).filter(t => t.active).map(t => ({ type: t.type, isPaid: t.isPaid, excludesFromPayroll: t.excludesFromPayroll }));

    return { success: true, employees: employees, balances: balances, types: types };
  } catch (err) { return { success: false, error: err.toString() }; }
}