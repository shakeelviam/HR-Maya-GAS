// ============================================================
// AttendanceOT.gs — compute OT from attendance, review, post to OT_Entries
// ============================================================
// Reads Attendance_Log per employee-day, matches each employee's assigned
// Shift, and computes (per day):
//   presence   = Check Out - Check In  (total, break included)
//   OT         = presence - Normal Hours (shift total; e.g. 9)   [never <0]
//   lateIn     = Check In later than (shift Start + Grace)       [flag]
//   earlyOut   = Check Out earlier than shift End                [flag]
//   breakTaken = SUM of all Break Out->Break In pairs (minutes)
//   breakOver  = breakTaken - shift Break Minutes (e.g. 60)      [flag]
// Bad punches (missing/again, unpaired breaks, no shift) are flagged.
//
// FLOW: getAttendanceOtDraft (reporting) -> compiler edits -> postOtToEntries.
// Nothing touches OT_Entries until the compiler posts.
//
// ENDPOINTS: getAttendanceOtDraft, postOtToEntries, scrubAttendancePins
// ============================================================

// One-shot: remove real PINs already stored in Attendance_Log.
function scrubAttendancePins() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName('Attendance_Log');
    if (!log) return { success: false, error: 'No Attendance_Log.' };
    const H = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0];
    let col = H.findIndex(h => String(h).trim().toLowerCase() === 'pin used');
    if (col === -1) col = H.findIndex(h => String(h).trim().toLowerCase() === 'auth');
    if (col === -1) return { success: true, message: 'No PIN column found — nothing to scrub.' };
    log.getRange(1, col + 1).setValue('Auth');
    const n = log.getLastRow() - 1;
    if (n > 0) log.getRange(2, col + 1, n, 1).setValue('Verified');
    return { success: true, message: 'Scrubbed PINs from ' + n + ' Attendance_Log row(s); column now shows "Verified".' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Shift lookup by name -> { start(min), end(min), normalHours, breakMinutes, graceMinutes }
function getShiftMap_(ss) {
  const map = {};
  const sh = ss.getSheetByName(SHIFTS_TAB);
  if (!sh) return map;
  const d = sh.getDataRange().getValues();
  const H = d[0];
  const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const cN = c('Shift Name'), cS = c('Start'), cE = c('End'), cNH = c('Normal Hours'), cB = c('Break Minutes'), cG = c('Grace Minutes'), cAP = c('Auto Present'), cWO = c('Weekly Off'), cA = c('Active');
  const toMin = (s) => { const m = String(s).trim().match(/^(\d{1,2}):(\d{2})$/); return m ? (parseInt(m[1],10)*60 + parseInt(m[2],10)) : null; };
  for (let i = 1; i < d.length; i++) {
    const name = String(d[i][cN] || '').trim();
    if (!name) continue;
    map[name.toLowerCase()] = {
      name: name, startMin: toMin(d[i][cS]), endMin: toMin(d[i][cE]),
      normalHours: Number(d[i][cNH]) || 9, breakMinutes: Number(d[i][cB]) || 0,
      graceMinutes: Number(d[i][cG]) || 0,
      autoPresent: cAP !== -1 && String(d[i][cAP]).trim().toLowerCase() === 'yes',
      weeklyOff: cWO !== -1 ? String(d[i][cWO] || '').trim().toLowerCase() : '',
      active: cA === -1 || String(d[i][cA]).trim().toLowerCase() === 'yes'
    };
  }
  return map;
}

// Compute the per-day OT draft for a date range (reporting only).
// params: { from:'dd-mm-yyyy', to:'dd-mm-yyyy' }
function getAttendanceOtDraft(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const from = params.from ? startOfDay_(parseFlexibleDate_(params.from)) : null;
    const to = params.to ? endOfDay_(parseFlexibleDate_(params.to)) : null;

    // Employee -> shift name (HR Maya).
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cShift = H.findIndex(h => String(h).trim().toLowerCase() === 'shift');
    const cDesig = H.findIndex(h => String(h).trim().toLowerCase() === 'designation');
    const empShift = {}, empName = {}, empDesig = {};
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim(); if (!id) continue;
      empShift[id] = cShift !== -1 ? String(hd[i][cShift] || '').trim() : '';
      empName[id] = cName !== -1 ? String(hd[i][cName] || '').trim() : id;
      empDesig[id] = cDesig !== -1 ? String(hd[i][cDesig] || '').trim() : '';
    }
    const shifts = getShiftMap_(ss);

    // Group punches by employee+date.
    const log = ss.getSheetByName('Attendance_Log');
    if (!log) return { success: false, error: 'No Attendance_Log.' };
    const L = log.getDataRange().getValues();
    const LH = L[0];
    const li = (n) => LH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const ciTs = li('Timestamp'), ciId = li('Employee ID'), ciDate = li('Date'), ciStage = li('Stage'), ciBranch = li('Branch');

    const days = {}; // key empId|date -> { punches:[{stage,ts,loc}], dateStr }
    for (let i = 1; i < L.length; i++) {
      const id = String(L[i][ciId] || '').trim();
      const stage = String(L[i][ciStage] || '').trim();
      if (!id || !stage) continue;
      const ts = parseLogTimestamp(L[i][ciTs]);
      if (!ts) continue;
      if (from && ts < from) continue;
      if (to && ts > to) continue;
      const dateStr = String(L[i][ciDate] || '').trim() || formatDdMmYyyy_(ts);
      const loc = ciBranch !== -1 ? String(L[i][ciBranch] || '').trim() : '';
      const key = id + '|' + dateStr;
      (days[key] = days[key] || { empId: id, dateStr: dateStr, punches: [] }).punches.push({ stage: stage, ts: ts, loc: loc });
    }

    // Shift map + location-shifts for resolution.
    const shiftMapFull = getShiftMap_(ss);
    const empLoc = {}; // employee assigned Location (fallback)
    (function () {
      const cLoc = H.findIndex(h => String(h).trim().toLowerCase() === 'location');
      const cB2 = cLoc !== -1 ? cLoc : H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
      for (let i = 1; i < hd.length; i++) { const id = String(hd[i][cId] || '').trim(); if (id) empLoc[id] = cB2 !== -1 ? String(hd[i][cB2] || '').trim() : ''; }
    })();

    const rows = [];
    Object.keys(days).forEach(key => {
      const d = days[key];
      const p = d.punches.sort((a, b) => a.ts - b.ts);
      const flags = [];

      const ci = p.find(x => x.stage === 'Check In');
      const co = [...p].reverse().find(x => x.stage === 'Check Out');

      // Resolve the day's shift. Employee's assigned Location is the reliable
      // source (staff are driven to their branch by company car); the punch's
      // geofenced location is a fallback when the assigned Location has no shift.
      let shiftName = '';
      const assignedLoc = empLoc[d.empId] || '';
      const punchLoc = (ci && ci.loc) ? ci.loc : '';
      if (ci) {
        const ciMin = ci.ts.getHours() * 60 + ci.ts.getMinutes();
        shiftName = resolveShiftForCheckin_(ss, assignedLoc, ciMin, shiftMapFull);
        if (!shiftName && punchLoc) shiftName = resolveShiftForCheckin_(ss, punchLoc, ciMin, shiftMapFull);
      }
      if (!shiftName) shiftName = empShift[d.empId] || ''; // fallback: employee's own Shift
      if (!shiftName && ci) flags.push('No shift resolved (check Location→Shifts)');
      const shift = shifts[shiftName.toLowerCase()];

      // Auto-present shifts (e.g. Head Office): counted present, no punch expected.
      // EXCEPTION: Designation = Driver — auto-present for presence, but OT is
      // entered manually (no fixed punch location). Flag them so the compiler
      // adds OT as OT_Entries rows (which payroll already sums).
      if (shift && shift.autoPresent) {
        const isDriver = String(empDesig[d.empId] || '').trim().toLowerCase() === 'driver';
        rows.push({
          empId: d.empId, name: empName[d.empId] || d.empId, date: d.dateStr, shift: shiftName,
          checkIn: ci ? Utilities.formatDate(ci.ts, tz_(), 'HH:mm') : 'auto',
          checkOut: co ? Utilities.formatDate(co.ts, tz_(), 'HH:mm') : 'auto',
          presence: shift.normalHours, ot: 0, dayType: 'Full', breakMin: 0, breakOver: 0, lateIn: 0, earlyOut: 0,
          flags: isDriver ? 'Auto-present (Driver — add OT manually in OT_Entries)' : 'Auto-present (no OT)'
        });
        return;
      }

      // Sum break pairs (Break Out -> next Break In).
      let breakMin = 0, openBreak = null, breakPairs = 0;
      p.forEach(x => {
        if (x.stage === 'Break Out') { if (openBreak) flags.push('Break Out without Break In'); openBreak = x.ts; }
        else if (x.stage === 'Break In') { if (openBreak) { breakMin += (x.ts - openBreak) / 60000; breakPairs++; openBreak = null; } else flags.push('Break In without Break Out'); }
      });
      if (openBreak) flags.push('Break Out not closed');

      let presenceH = null, otH = 0, lateInMin = 0, earlyOutMin = 0, breakOverMin = 0;
      if (!ci) flags.push('No Check In');
      if (!co) flags.push('No Check Out');
      if (!shiftName) flags.push('No shift assigned');
      else if (!shift) flags.push('Shift "' + shiftName + '" not found');

      if (ci && co) {
        let mins = (co.ts - ci.ts) / 60000;
        if (mins < 0) mins += 24 * 60; // crossed midnight (overnight)
        presenceH = Math.round((mins / 60) * 100) / 100;
        if (shift) {
          otH = Math.max(0, Math.round((presenceH - shift.normalHours) * 100) / 100);
          // late-in vs shift start + grace (compare minutes-of-day)
          if (shift.startMin != null) {
            const ciMin = ci.ts.getHours() * 60 + ci.ts.getMinutes();
            const lateBy = ciMin - (shift.startMin + shift.graceMinutes);
            if (lateBy > 0) { lateInMin = lateBy; flags.push('Late in ' + lateBy + 'm'); }
          }
          if (shift.endMin != null) {
            const coMin = co.ts.getHours() * 60 + co.ts.getMinutes();
            const earlyBy = shift.endMin - coMin;
            if (earlyBy > 0 && earlyBy < 12 * 60) { earlyOutMin = earlyBy; flags.push('Early out ' + earlyBy + 'm'); }
          }
          breakOverMin = Math.max(0, Math.round(breakMin - shift.breakMinutes));
          if (breakOverMin > 0) flags.push('Break over ' + breakOverMin + 'm');
        }
        if (presenceH != null && presenceH <= 6) flags.push('Half day (' + presenceH + 'h ≤ 6h — needs approval)');
      }

      rows.push({
        empId: d.empId, name: empName[d.empId] || d.empId, date: d.dateStr, shift: shiftName,
        location: assignedLoc || punchLoc,
        checkIn: ci ? Utilities.formatDate(ci.ts, tz_(), 'HH:mm') : '',
        checkOut: co ? Utilities.formatDate(co.ts, tz_(), 'HH:mm') : '',
        presence: presenceH, ot: otH,
        dayType: (presenceH == null) ? '' : (presenceH > 6 ? 'Full' : 'Half'),
        breakMin: Math.round(breakMin), breakOver: breakOverMin,
        lateIn: lateInMin, earlyOut: earlyOutMin,
        flags: flags.join('; ')
      });
    });

    rows.sort((a, b) => (a.name + a.date).localeCompare(b.name + b.date));
    const totalOt = Math.round(rows.reduce((s, r) => s + (r.ot || 0), 0) * 100) / 100;
    return { success: true, from: params.from || '', to: params.to || '', count: rows.length, totalOt: totalOt, rows: rows };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Post approved OT rows into OT_Entries. The compiler may have edited OT hours.
// params: { rows: [{empId, name, date, ot, remark?}], source? }
// Replaces any existing "Attendance" rows for the same emp+date to avoid dupes,
// then appends. Manual OT_Entries rows are left untouched.
function postOtToEntries(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const rows = params.rows || [];
    if (!rows.length) return { success: false, error: 'No rows to post.' };
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    let sheet = ss.getSheetByName(OT_ENTRIES_TAB);
    if (!sheet) { sheet = ss.insertSheet(OT_ENTRIES_TAB); sheet.appendRow(['Date','Employee ID','Name','OT Hours','Remark','Added By']); }

    // Remove prior auto rows for the same emp+date (Added By starts with 'Attendance').
    const data = sheet.getDataRange().getValues();
    const toDelete = [];
    const targetKeys = {};
    rows.forEach(r => { targetKeys[String(r.empId).trim() + '|' + String(r.date).trim()] = true; });
    for (let i = data.length - 1; i >= 1; i--) {
      const key = String(data[i][1]).trim() + '|' + String(data[i][0]).trim();
      const addedBy = String(data[i][5] || '');
      if (targetKeys[key] && addedBy.indexOf('Attendance') === 0) toDelete.push(i + 1);
    }
    toDelete.forEach(rn => sheet.deleteRow(rn));

    const by = 'Attendance (' + String(params.by || 'approved') + ')';
    let posted = 0;
    const append = [];
    rows.forEach(r => {
      const ot = Number(r.ot) || 0;
      if (ot <= 0) return; // only positive OT rows
      append.push([String(r.date).trim(), String(r.empId).trim(), String(r.name || '').trim(), ot, String(r.remark || 'Auto from attendance').trim(), by]);
      posted++;
    });
    if (append.length) {
      const start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, append.length, 6).setValues(append);
      sheet.getRange(start, 1, append.length, 1).setNumberFormat('@');
    }
    return { success: true, posted: posted, replaced: toDelete.length,
      message: 'Posted ' + posted + ' OT row(s) to OT_Entries' + (toDelete.length ? ' (replaced ' + toDelete.length + ' prior attendance row(s))' : '') + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Standalone: archive the current OT_Entries and start a fresh empty tab.
// Run once from the editor before going live with attendance. Optional label
// via ?label=... ; defaults to today's date.
function archiveOtEntriesNow(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const label = String(params.label || '').trim() || Utilities.formatDate(new Date(), tz_(), 'dd-MMM-yyyy');
    const sheet = ss.getSheetByName(OT_ENTRIES_TAB);
    if (!sheet) {
      // Nothing to archive; just make a fresh tab.
      const fresh = ss.insertSheet(OT_ENTRIES_TAB);
      fresh.appendRow(['Date', 'Employee ID', 'Name', 'OT Hours', 'Remark', 'Added By']);
      fresh.getRange(1, 1, 1, 6).setFontWeight('bold'); fresh.setFrozenRows(1);
      return { success: true, message: 'No existing OT_Entries. Created a fresh empty OT_Entries.' };
    }
    const rowsArchived = Math.max(0, sheet.getLastRow() - 1);
    let archiveName = 'OT_Entries ' + label;
    let n = 2;
    while (ss.getSheetByName(archiveName)) { archiveName = 'OT_Entries ' + label + ' (' + n + ')'; n++; }
    sheet.setName(archiveName);
    const fresh = ss.insertSheet(OT_ENTRIES_TAB);
    fresh.appendRow(['Date', 'Employee ID', 'Name', 'OT Hours', 'Remark', 'Added By']);
    fresh.getRange(1, 1, 1, 6).setFontWeight('bold'); fresh.setFrozenRows(1);
    return { success: true, archived: archiveName, rowsArchived: rowsArchived,
      message: 'Archived ' + rowsArchived + ' OT row(s) to "' + archiveName + '" and started a fresh empty OT_Entries.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Purge Attendance_Log and rebuild the full, correct header.
// params.keepHeader ignored — always rewrites the standard header.
function resetAttendanceLog(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Attendance_Log');
    const header = ['Timestamp', 'Employee ID', 'Employee Name', 'Date', 'Stage', 'Auth', 'Branch', 'Latitude', 'Longitude', 'Distance (m)', 'Geo Flag'];
    if (!sheet) sheet = ss.insertSheet('Attendance_Log');
    const hadRows = Math.max(0, sheet.getLastRow() - 1);
    sheet.clear();
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return { success: true, purged: hadRows, message: 'Attendance_Log purged (' + hadRows + ' rows removed) and headers rebuilt: ' + header.join(', ') + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}