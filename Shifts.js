// ============================================================
// Shifts.gs — Maya HR : Shift definitions + assignment
// ============================================================
// Foundation for the attendance/OT engine. A Shift defines the expected
// working block so the system can compute OT and flag late/early/break.
//
// Shifts tab columns:
//   Shift Name | Start | End | Normal Hours | Break Minutes | Grace Minutes | Active
//
// Model (locked): Normal Hours = 9 TOTAL (break included). OT = presence - 9.
// Start drives the late-in flag (with Grace). End is informational. Break
// Minutes is the allowed break, used only for the break-overrun flag.
// Times are stored as text "HH:mm" (24h).
//
// ENDPOINTS: getShifts, saveShift, deleteShift, assignShift
// ============================================================

const SHIFTS_TAB = 'Shifts';

function ensureShiftsTab_(ss) {
  let sheet = ss.getSheetByName(SHIFTS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SHIFTS_TAB);
    sheet.appendRow(['Shift Name', 'Start', 'End', 'Normal Hours', 'Break Minutes', 'Grace Minutes', 'Auto Present', 'Weekly Off', 'Active']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // No seeded shifts — the user defines all shifts on the Shifts page.
  }
  // Ensure the two attendance-behaviour columns exist on older Shifts tabs.
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const need = (name) => H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase()) === -1;
  if (need('Auto Present')) sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Auto Present').setFontWeight('bold');
  if (need('Weekly Off')) sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Weekly Off').setFontWeight('bold');
  return sheet;
}

// Add a "Shift" column to HR Maya if missing.
function ensureHrShiftColumn_(ss) {
  const hr = ss.getSheetByName('HR Maya');
  const H = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0];
  if (H.findIndex(h => String(h).trim().toLowerCase() === 'shift') === -1) {
    hr.getRange(1, hr.getLastColumn() + 1).setValue('Shift').setFontWeight('bold');
  }
  return hr;
}

function isYesShift_(v) { const s = String(v).trim().toLowerCase(); return s === 'yes' || s === 'true' || s === '1'; }
function yesNoShift_(v) { return isYesShift_(v) ? 'Yes' : 'No'; }
function validHHmm_(s) { return /^([01]?\d|2[0-3]):[0-5]\d$/.test(String(s).trim()); }
function padHHmm_(s) { const m = String(s).trim().match(/^(\d{1,2}):(\d{2})$/); return m ? (m[1].padStart(2, '0') + ':' + m[2]) : String(s).trim(); }

function getShifts(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureShiftsTab_(ss);
    const data = sheet.getDataRange().getValues();
    const H = data[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cN = c('Shift Name'), cS = c('Start'), cE = c('End'), cNH = c('Normal Hours'), cB = c('Break Minutes'), cG = c('Grace Minutes'), cAP = c('Auto Present'), cWO = c('Weekly Off'), cA = c('Active');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][cN] || '').trim();
      if (!name) continue;
      out.push({
        name: name,
        start: padHHmm_(data[i][cS]), end: padHHmm_(data[i][cE]),
        normalHours: Number(data[i][cNH]) || 9,
        breakMinutes: Number(data[i][cB]) || 0,
        graceMinutes: Number(data[i][cG]) || 0,
        autoPresent: cAP !== -1 ? isYesShift_(data[i][cAP]) : false,
        weeklyOff: cWO !== -1 ? String(data[i][cWO] || '').trim() : '',
        active: cA !== -1 ? isYesShift_(data[i][cA]) : true
      });
    }
    return { success: true, data: out };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Add/edit a shift by name.
// params: name, start, end, normalHours, breakMinutes, graceMinutes, active
function saveShift(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureShiftsTab_(ss);
    const name = String(params.name || '').trim();
    if (!name) return { success: false, error: 'Shift name is required.' };
    const start = padHHmm_(params.start), end = padHHmm_(params.end);
    if (params.start !== undefined && params.start !== '' && !validHHmm_(start)) return { success: false, error: 'Start must be HH:mm (24h), e.g. 05:30.' };
    if (params.end !== undefined && params.end !== '' && !validHHmm_(end)) return { success: false, error: 'End must be HH:mm (24h), e.g. 14:30.' };

    const data = sheet.getDataRange().getValues();
    const H = data[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cN = c('Shift Name'), cS = c('Start'), cE = c('End'), cNH = c('Normal Hours'), cB = c('Break Minutes'), cG = c('Grace Minutes'), cAP = c('Auto Present'), cWO = c('Weekly Off'), cA = c('Active');

    let rowNum = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cN] || '').trim().toLowerCase() === name.toLowerCase()) { rowNum = i + 1; break; }
    }
    const width = H.length;
    const row = rowNum === -1 ? new Array(width).fill('') : data[rowNum - 1].slice();
    const prov = (k) => params[k] !== undefined && params[k] !== '';
    row[cN] = name;
    if (rowNum === -1 || prov('start')) row[cS] = start;
    if (rowNum === -1 || prov('end')) row[cE] = end;
    row[cNH] = (params.normalHours !== undefined && params.normalHours !== '') ? Number(params.normalHours) : (rowNum === -1 ? 9 : row[cNH]);
    row[cB] = (params.breakMinutes !== undefined && params.breakMinutes !== '') ? Number(params.breakMinutes) : (rowNum === -1 ? 60 : row[cB]);
    row[cG] = (params.graceMinutes !== undefined && params.graceMinutes !== '') ? Number(params.graceMinutes) : (rowNum === -1 ? 10 : row[cG]);
    if (cAP !== -1 && (rowNum === -1 || params.autoPresent !== undefined)) row[cAP] = yesNoShift_(params.autoPresent);
    if (cWO !== -1 && (rowNum === -1 || params.weeklyOff !== undefined)) row[cWO] = String(params.weeklyOff || '').trim();
    if (cA !== -1) row[cA] = (params.active === undefined) ? (rowNum === -1 ? 'Yes' : row[cA]) : yesNoShift_(params.active);

    if (rowNum === -1) { sheet.appendRow(row); rowNum = sheet.getLastRow(); }
    else sheet.getRange(rowNum, 1, 1, width).setValues([row]);
    sheet.getRange(rowNum, cS + 1, 1, 2).setNumberFormat('@'); // keep times as text

    return { success: true, name: name, message: (rowNum === sheet.getLastRow() && data.findIndex((r, i) => i > 0 && String(r[cN]).trim().toLowerCase() === name.toLowerCase()) === -1 ? 'Created' : 'Saved') + ' shift "' + name + '".' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function deleteShift(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureShiftsTab_(ss);
    const name = String(params.name || '').trim();
    const data = sheet.getDataRange().getValues();
    const H = data[0];
    const cN = H.findIndex(h => String(h).trim().toLowerCase() === 'shift name');
    const cA = H.findIndex(h => String(h).trim().toLowerCase() === 'active');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cN] || '').trim().toLowerCase() === name.toLowerCase()) {
        if (isYesShift_(params.hard)) { sheet.deleteRow(i + 1); return { success: true, message: 'Deleted shift "' + name + '".' }; }
        sheet.getRange(i + 1, cA + 1).setValue('No');
        return { success: true, message: 'Retired shift "' + name + '".' };
      }
    }
    return { success: false, error: 'Shift "' + name + '" not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Assign a shift to employees. scope: 'all' | 'branch' | 'ids'
// params: shift, scope, branch?, ids?
function assignShift(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureShiftsTab_(ss);
    const hr = ensureHrShiftColumn_(ss);
    const shift = String(params.shift || '').trim();
    if (!shift) return { success: false, error: 'shift required' };
    const scope = String(params.scope || 'all').toLowerCase();

    const shifts = getShifts({}).data || [];
    if (!shifts.some(s => s.name.toLowerCase() === shift.toLowerCase())) return { success: false, error: 'Unknown shift "' + shift + '".' };

    const data = hr.getDataRange().getValues();
    const H = data[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    let cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'location');
    if (cBranch === -1) cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cShift = H.findIndex(h => String(h).trim().toLowerCase() === 'shift');
    const idSet = scope === 'ids' ? String(params.ids || '').split(',').map(s => s.trim().toUpperCase()).filter(Boolean) : null;
    const branch = scope === 'branch' ? String(params.branch || '').trim().toLowerCase() : null;

    let assigned = 0;
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][cId] || '').trim();
      if (!id) continue;
      if (idSet && idSet.indexOf(id.toUpperCase()) === -1) continue;
      if (branch && String(data[i][cBranch] || '').trim().toLowerCase() !== branch) continue;
      hr.getRange(i + 1, cShift + 1).setValue(shift);
      assigned++;
    }
    return { success: true, shift: shift, assigned: assigned,
      message: 'Assigned shift "' + shift + '" to ' + assigned + ' employee(s).' };
  } catch (err) { return { success: false, error: err.toString() }; }
}