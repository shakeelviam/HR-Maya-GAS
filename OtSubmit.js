// ============================================================
// OtSubmit.gs — staff self-submit OT (kiosk add-on) -> pending -> compiler approves
// ============================================================
// submitOtRequest — validates Employee ID + PIN, checks no OT already exists
//   for that date (in OT_Pending OR OT_Entries), writes a PENDING row.
// getPendingOt / approvePendingOt / rejectPendingOt — compiler review.
// Approved pending rows are posted into OT_Entries (which payroll sums).
//
// OT_Pending tab: ID · Employee ID · Name · Date · Hours · Source · Status · Submitted
//
// ENDPOINTS: submitOtRequest, getPendingOt, approvePendingOt, rejectPendingOt
// ============================================================

const OT_PENDING_TAB = 'OT_Pending';

function ensureOtPendingTab_(ss) {
  let sheet = ss.getSheetByName(OT_PENDING_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(OT_PENDING_TAB);
    sheet.appendRow(['ID', 'Employee ID', 'Name', 'Date', 'Hours', 'Source', 'Status', 'Submitted', 'Geo Flag', 'Computed OT']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normDayStr_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[1].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[1];
  var d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}

// Does the employee already have OT (pending or approved) for this date?
function otExistsForDate_(ss, empId, dd) {
  // OT_Entries (approved)
  const ent = ss.getSheetByName(OT_ENTRIES_TAB);
  if (ent) {
    const d = ent.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][1]).trim() === empId && normDayStr_(d[i][0]) === dd && (Number(d[i][3]) || 0) > 0) {
        return { where: 'approved', hours: Number(d[i][3]) || 0 };
      }
    }
  }
  // OT_Pending (awaiting approval)
  const pen = ss.getSheetByName(OT_PENDING_TAB);
  if (pen) {
    const d = pen.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][1]).trim() === empId && normDayStr_(d[i][3]) === dd && String(d[i][6]).trim().toLowerCase() === 'pending') {
        return { where: 'pending', hours: Number(d[i][4]) || 0 };
      }
    }
  }
  return null;
}

// Compute OT from attendance for an employee on a date: presence - shift hours.
// Returns { hasPunch, presence, ot } — ot null if no complete punch pair.
function computeOtFromPunches_(ss, empId, dd) {
  const log = ss.getSheetByName('Attendance_Log');
  if (!log) return { hasPunch: false, presence: null, ot: null };
  const data = log.getDataRange().getValues();
  const H = data[0];
  const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
  const cDate = H.findIndex(h => String(h).trim().toLowerCase() === 'date');
  const cStage = H.findIndex(h => String(h).trim().toLowerCase() === 'stage');
  const cTs = H.findIndex(h => String(h).trim().toLowerCase() === 'timestamp');
  let ci = null, co = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cId]).trim() !== empId) continue;
    if (normDayStr_(data[i][cDate]) !== dd) continue;
    const stage = String(data[i][cStage]).trim();
    const ts = (typeof parseLogTimestamp === 'function') ? parseLogTimestamp(data[i][cTs]) : null;
    if (!ts) continue;
    if (stage === 'Check In' && (!ci || ts < ci)) ci = ts;
    if (stage === 'Check Out' && (!co || ts > co)) co = ts;
  }
  if (!ci || !co) return { hasPunch: (!!ci || !!co), presence: null, ot: null };
  let mins = (co - ci) / 60000; if (mins < 0) mins += 1440;
  const presence = Math.round((mins / 60) * 100) / 100;
  // Shift normal hours: employee's assigned shift, else 9.
  let normalHours = 9;
  try {
    const shifts = getShiftMap_(ss);
    const hr = ss.getSheetByName('HR Maya'); const hd = hr.getDataRange().getValues();
    const hH = hd[0]; const hcId = hH.findIndex(h=>String(h).trim().toLowerCase()==='employee id'); const hcSh = hH.findIndex(h=>String(h).trim().toLowerCase()==='shift');
    for (let i=1;i<hd.length;i++){ if(String(hd[i][hcId]).trim()===empId){ const sn=hcSh!==-1?String(hd[i][hcSh]||'').trim().toLowerCase():''; if(sn&&shifts[sn]) normalHours=shifts[sn].normalHours||9; break; } }
  } catch (e) {}
  const ot = Math.max(0, Math.round((presence - normalHours) * 100) / 100);
  return { hasPunch: true, presence: presence, ot: ot };
}

// Staff submit: { empId, pin, date (dd-mm-yyyy), hours, lat, lng }
function submitOtRequest(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const empId = String(params.empId || '').trim();
    const pin = String(params.pin || '').trim();
    const hours = Number(params.hours) || 0;
    const dateRaw = String(params.date || '').trim();
    if (!empId || !pin) return { success: false, error: 'Employee ID and PIN required.' };
    if (!dateRaw) return { success: false, error: 'Date required (dd-mm-yyyy).' };
    if (hours <= 0) return { success: false, error: 'Enter OT hours greater than 0.' };
    if (hours > 16) return { success: false, error: 'Hours look too high (' + hours + '). Please check.' };

    // Validate ID + PIN via the kiosk login check.
    const auth = loginStaffKiosk({ empId: empId, pin: pin });
    if (!auth || !auth.success) return { success: false, error: 'Invalid Employee ID or PIN.' };

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureOtPendingTab_(ss);
    const parsed = parseFlexibleDate_(dateRaw);
    if (!parsed) return { success: false, error: 'Bad date. Use dd-mm-yyyy.' };
    const dd = formatDdMmYyyy_(parsed);

    const exists = otExistsForDate_(ss, empId, dd);
    if (exists) {
      return { success: false, error: 'OT for ' + dd + ' already exists for you: ' + exists.hours + ' hours (' + exists.where + '). Contact HR for any amendment.' };
    }

    const name = empNameById_(ss, empId) || empId;
    // Geofence the submission (they submit at the location right after punch-out).
    // Geofence using the employee's OWN location (so Drivers/non-geofenced get
    // "nearest, not fenced" instead of a false "outside").
    let empLoc = '';
    try {
      const hr = ss.getSheetByName('HR Maya'); const hd = hr.getDataRange().getValues(); const hH = hd[0];
      let cL = hH.findIndex(h => String(h).trim().toLowerCase() === 'location');
      if (cL === -1) cL = hH.findIndex(h => String(h).trim().toLowerCase() === 'branch');
      const cI = hH.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
      for (let i = 1; i < hd.length; i++) { if (String(hd[i][cI]).trim() === empId) { empLoc = cL !== -1 ? String(hd[i][cL] || '').trim() : ''; break; } }
    } catch (e) {}
    const geo = checkGeofence_(ss, empLoc, params.lat, params.lng);
    // Cross-check: OT computed from their punches (presence - shift hours).
    const comp = computeOtFromPunches_(ss, empId, dd);
    const compStr = comp.ot == null ? (comp.hasPunch ? 'incomplete punch' : 'no punch') : (comp.ot + 'h (presence ' + comp.presence + 'h)');

    const pen = ss.getSheetByName(OT_PENDING_TAB);
    const id = 'OTP-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');
    pen.appendRow([id, empId, name, dd, hours, 'Kiosk (self)', 'Pending', Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm:ss'), geo.flag, compStr]);
    pen.getRange(pen.getLastRow(), 4).setNumberFormat('@');
    return { success: true, geoFlag: geo.flag, computedOt: comp.ot,
      message: 'Recorded ' + hours + 'h OT for ' + dd + ' — pending HR review. Thank you.' +
        (geo.flag && geo.flag.indexOf('OK') !== 0 ? ' (location: ' + geo.flag + ')' : '') };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function getPendingOt(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const pen = ensureOtPendingTab_(ss);
    const d = pen.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim() && String(d[i][6]).trim().toLowerCase() === 'pending') {
        rows.push({ row: i + 1, id: d[i][0], empId: d[i][1], name: d[i][2], date: d[i][3], hours: Number(d[i][4]) || 0, source: d[i][5], submitted: d[i][7],
          geoFlag: d[i][8] || '', computedOt: d[i][9] || '' });
      }
    }
    return { success: true, rows: rows };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Approve: post to OT_Entries, mark pending row Approved. params: { row, hours?, by }
function approvePendingOt(params) {
  try {
    params = params || {};
    const rowNum = Number(params.row);
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const pen = ss.getSheetByName(OT_PENDING_TAB);
    if (!pen || rowNum < 2) return { success: false, error: 'Valid row required.' };
    const r = pen.getRange(rowNum, 1, 1, 8).getValues()[0];
    if (String(r[6]).trim().toLowerCase() !== 'pending') return { success: false, error: 'Already ' + r[6] + '.' };
    const empId = String(r[1]).trim(), name = String(r[2]).trim(), dd = String(r[3]).trim();
    const hours = (params.hours !== undefined && params.hours !== '') ? Number(params.hours) : Number(r[4]) || 0;

    let ent = ss.getSheetByName(OT_ENTRIES_TAB);
    if (!ent) { ent = ss.insertSheet(OT_ENTRIES_TAB); ent.appendRow(['Date','Employee ID','Name','OT Hours','Remark','Added By']); }
    ent.appendRow([dd, empId, name, hours, 'Self-submitted (approved)', 'OT approval (' + String(params.by || 'compiler') + ')']);
    ent.getRange(ent.getLastRow(), 1).setNumberFormat('@');
    pen.getRange(rowNum, 5).setValue(hours);
    pen.getRange(rowNum, 7).setValue('Approved');
    return { success: true, message: 'Approved ' + hours + 'h for ' + name + ' (' + dd + ') — posted to OT_Entries.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function rejectPendingOt(params) {
  try {
    params = params || {};
    const rowNum = Number(params.row);
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const pen = ss.getSheetByName(OT_PENDING_TAB);
    if (!pen || rowNum < 2) return { success: false, error: 'Valid row required.' };
    pen.getRange(rowNum, 7).setValue('Rejected');
    return { success: true, message: 'Rejected.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}