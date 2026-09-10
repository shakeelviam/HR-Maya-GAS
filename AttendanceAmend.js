// ============================================================
// AttendanceAmend.gs — view + admin-amend the Attendance_Log
// ============================================================
// getAttendanceLog   — read punches (filter by date / employee / location).
// addAttendancePunch — admin adds a punch (flagged MANUAL); duplicate stage
//                      per employee+date blocked (same rule as the kiosk).
// editAttendancePunch— admin edits an existing punch by row.
// deleteAttendancePunch — admin removes a punch by row.
// Every change is written to Attendance_Amendments (who/when/old->new).
//
// FIX: 'Day Off' is now an accepted stage. Previously ATTN_STAGES only
// listed the four punch stages, so an admin could not manually add or
// edit a Day Off — the same gap that was already fixed in
// markKioskAttendance (Code.gs) but never here. A Day Off does not
// need a time; it is stored at 00:00 like every other Day Off row.
//
// ENDPOINTS: getAttendanceLog, addAttendancePunch, editAttendancePunch,
//            deleteAttendancePunch
// ============================================================

const ATTN_AMEND_TAB = 'Attendance_Amendments';
const ATTN_STAGES = ['Check In', 'Break Out', 'Break In', 'Check Out', 'Day Off'];

function ensureAttnAmendTab_(ss) {
  let sheet = ss.getSheetByName(ATTN_AMEND_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(ATTN_AMEND_TAB);
    sheet.appendRow(['When', 'By', 'Action', 'Employee ID', 'Date', 'Stage', 'Detail (old \u2192 new)']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function logAmend_(ss, by, action, empId, dateStr, stage, detail) {
  ensureAttnAmendTab_(ss).appendRow([Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm:ss'),
    String(by || 'Admin'), action, empId, dateStr, stage, detail || '']);
}
function attnCols_(H) {
  const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  return { ts: c('Timestamp'), id: c('Employee ID'), name: c('Employee Name'), date: c('Date'),
    stage: c('Stage'), auth: c('Auth') !== -1 ? c('Auth') : c('PIN Used'),
    branch: c('Branch'), lat: c('Latitude'), lng: c('Longitude'), dist: c('Distance (m)'), flag: c('Geo Flag') };
}

// Read the log with optional filters.
function getAttendanceLog(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName('Attendance_Log');
    if (!log) return { success: true, rows: [] };
    const data = log.getDataRange().getValues();
    const H = data[0]; const C = attnCols_(H);
    const fDate = String(params.date || '').trim();
    const fEmp = String(params.empId || '').trim().toUpperCase();
    const fLoc = String(params.location || '').trim().toLowerCase();
    const normDay = (v) => {
      if (v == null || v === '') return '';
      if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
      var s = String(v).trim();
      var m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
      if (m) return m[1].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[3];
      m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
      if (m) return m[3].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[1];
      var d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
    };
    const fDateNorm = normDay(fDate);
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][C.id] || '').trim();
      if (!id) continue;
      const dateStr = normDay(data[i][C.date]);
      const loc = C.branch !== -1 ? String(data[i][C.branch] || '').trim() : '';
      if (fDateNorm && dateStr !== fDateNorm) continue;
      if (fEmp && id.toUpperCase() !== fEmp) continue;
      if (fLoc && loc.toLowerCase() !== fLoc) continue;
      rows.push({ row: i + 1, ts: String(data[i][C.ts] || ''), empId: id,
        name: C.name !== -1 ? String(data[i][C.name] || '') : '', date: dateStr,
        stage: String(data[i][C.stage] || ''), auth: C.auth !== -1 ? String(data[i][C.auth] || '') : '',
        branch: loc, flag: C.flag !== -1 ? String(data[i][C.flag] || '') : '' });
    }
    rows.sort((a, b) => (a.date + a.empId + a.stage).localeCompare(b.date + b.empId + b.stage));
    return { success: true, rows: rows };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Admin adds a punch. params: { empId, date (dd-mm-yyyy), stage, time ("HH:mm"), location, by }
// 'Day Off' is accepted and does not require a time (stored at 00:00).
function addAttendancePunch(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const empId = String(params.empId || '').trim();
    const stage = String(params.stage || '').trim();
    const dateStr = String(params.date || '').trim();
    const time = String(params.time || '').trim();
    if (!empId || !stage || !dateStr) return { success: false, error: 'empId, date and stage required.' };
    if (ATTN_STAGES.indexOf(stage) === -1) return { success: false, error: 'Invalid stage.' };
    const dd = formatDdMmYyyy_(parseFlexibleDate_(dateStr) || new Date());

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    let log = ss.getSheetByName('Attendance_Log');
    if (!log) { log = ss.insertSheet('Attendance_Log'); log.appendRow(['Timestamp','Employee ID','Employee Name','Date','Stage','Auth','Branch','Latitude','Longitude','Distance (m)','Geo Flag']); }
    const data = log.getDataRange().getValues();
    const H = data[0]; const C = attnCols_(H);

    // Duplicate stage guard (same rule as kiosk).
    const normDay = (v) => { if (v instanceof Date) return formatDdMmYyyy_(v); const m = String(v).trim().match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/); return m ? (m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[3]) : String(v).trim(); };
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][C.id]).trim() === empId && normDay(data[i][C.date]) === dd && String(data[i][C.stage]).trim() === stage) {
        return { success: false, error: 'This employee already has "' + stage + '" on ' + dd + '.' };
      }
    }
    const name = empNameById_(ss, empId) || empId;
    // Timestamp: date + time. Day Off has no meaningful time — store 00:00.
    // Stored like the kiosk (dd/MM/yyyy HH:mm:ss).
    const defaultTime = (stage === 'Day Off') ? '00:00' : '00:00';
    const tparts = (time.match(/^(\d{1,2}):(\d{2})$/)) ? time : defaultTime;
    const tsStr = Utilities.formatDate(parseFlexibleDate_(dateStr), tz_(), 'dd/MM/yyyy') + ' ' + tparts + ':00';
    const loc = String(params.location || '').trim();
    const rowVals = new Array(H.length).fill('');
    rowVals[C.ts] = tsStr; rowVals[C.id] = empId; if (C.name!==-1) rowVals[C.name]=name;
    rowVals[C.date] = dd; rowVals[C.stage] = stage;
    if (C.auth !== -1) rowVals[C.auth] = 'Manual';
    if (C.branch !== -1) rowVals[C.branch] = loc;
    if (C.flag !== -1) rowVals[C.flag] = 'MANUAL (admin)';
    log.appendRow(rowVals);
    log.getRange(log.getLastRow(), C.date + 1).setNumberFormat('@');

    logAmend_(ss, params.by, 'ADD', empId, dd, stage,
      (stage === 'Day Off' ? 'added Day Off' : 'added @ ' + tparts) + (loc ? ' / ' + loc : ''));
    return { success: true, message: 'Added ' + stage + ' for ' + name + ' on ' + dd +
      (stage === 'Day Off' ? ' (MANUAL).' : ' at ' + tparts + ' (MANUAL).') };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Admin edits a punch by row. params: { row, stage?, time?, location?, by }
function editAttendancePunch(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const rowNum = Number(params.row);
    if (!rowNum || rowNum < 2) return { success: false, error: 'Valid row required.' };
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName('Attendance_Log');
    const H = log.getDataRange().getValues()[0]; const C = attnCols_(H);
    const cur = log.getRange(rowNum, 1, 1, H.length).getValues()[0];
    const empId = String(cur[C.id] || '').trim();
    const dd = String(cur[C.date] || '').trim();
    const changes = [];

    if (params.stage !== undefined && params.stage !== '' && String(params.stage) !== String(cur[C.stage])) {
      if (ATTN_STAGES.indexOf(String(params.stage)) === -1) return { success: false, error: 'Invalid stage.' };
      changes.push('stage ' + cur[C.stage] + ' \u2192 ' + params.stage);
      log.getRange(rowNum, C.stage + 1).setValue(params.stage);
    }
    if (params.time !== undefined && params.time !== '' && /^(\d{1,2}):(\d{2})$/.test(params.time)) {
      const d = parseFlexibleDate_(dd);
      const newTs = Utilities.formatDate(d || new Date(), tz_(), 'dd/MM/yyyy') + ' ' + params.time + ':00';
      changes.push('time \u2192 ' + params.time);
      log.getRange(rowNum, C.ts + 1).setValue(newTs);
    }
    if (params.location !== undefined && C.branch !== -1 && String(params.location) !== String(cur[C.branch])) {
      changes.push('location ' + (cur[C.branch] || '(blank)') + ' \u2192 ' + params.location);
      log.getRange(rowNum, C.branch + 1).setValue(params.location);
    }
    if (!changes.length) return { success: true, message: 'No changes.' };
    if (C.flag !== -1) log.getRange(rowNum, C.flag + 1).setValue('MANUAL (edited)');
    logAmend_(ss, params.by, 'EDIT', empId, dd, String(cur[C.stage]), changes.join('; '));
    return { success: true, message: 'Updated: ' + changes.join('; ') };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Admin deletes a punch by row.
function deleteAttendancePunch(params) {
  try {
    params = params || {};
    const rowNum = Number(params.row);
    if (!rowNum || rowNum < 2) return { success: false, error: 'Valid row required.' };
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName('Attendance_Log');
    const H = log.getDataRange().getValues()[0]; const C = attnCols_(H);
    const cur = log.getRange(rowNum, 1, 1, H.length).getValues()[0];
    const empId = String(cur[C.id] || '').trim();
    const dd = String(cur[C.date] || '').trim();
    const stage = String(cur[C.stage] || '').trim();
    if (!empId) return { success: false, error: 'Empty row.' };
    log.deleteRow(rowNum);
    logAmend_(ss, params.by, 'DELETE', empId, dd, stage, 'removed punch');
    return { success: true, message: 'Deleted ' + stage + ' for ' + empId + ' on ' + dd + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}