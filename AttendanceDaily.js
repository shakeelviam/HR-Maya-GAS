// ============================================================
// AttendanceDaily.gs  — v3
// ------------------------------------------------------------
// rebuildAttendanceDaily()
//   • Part 1 : 22–31 Aug from Manual_Attendance_Aug
//   • Part 2 : 1 Sep → YESTERDAY from Attendance_Log
//   • Leave  : "Granted Leaves" tab — Approved grants override
//              Default Day Off with "On Leave" status.
//
// Priority order (highest → lowest):
//   1. Manual_Aug entry       (supervisor submitted)
//   2. Kiosk completed session (Check In + Check Out)
//   3. Approved Granted Leave  (new)
//   4. Auto_HeadOffice         (Head Office shift, working day)
//   5. Weekend                 (Head Office shift, Fri/Sat only)
//   6. Default Day Off         (no data, not auto-present)
//
// WEEKEND RULE: Fri/Sat off applies ONLY to Head Office
// (Auto Present = Yes). All other locations work 7 days.
//
// OPEN SESSIONS: not captured. Second-shift employees who
// haven't started yet would be wrongly flagged. Only complete
// days (up to yesterday) are compiled.
// ============================================================

function adColMap_(sheet) {
  const hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map  = {};
  hdrs.forEach((h, i) => { const k = String(h).trim(); if (k) map[k] = i; });
  return map;
}

function adNormDd_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return adPad_(val.getDate()) + '-' + adPad_(val.getMonth() + 1) + '-' + val.getFullYear();
  }
  const s = String(val).trim();
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) return adPad_(m2[1]) + '-' + adPad_(m2[2]) + '-' + m2[3];
  return s;
}

function adPad_(n) { return String(n).padStart(2, '0'); }

function adDateToStr_(d) {
  return adPad_(d.getDate()) + '-' + adPad_(d.getMonth() + 1) + '-' + d.getFullYear();
}

function adParseDate_(s) {
  const m = String(s).match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
}

function adParseTs_(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const s = String(val).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):?(\d{0,2})/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1], +m[4], +m[5], m[6] ? +m[6] : 0);
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function adIsWeekend_(d) { const w = d.getDay(); return w === 5 || w === 6; }

function adHHmm_(d) {
  if (!d || isNaN(d.getTime())) return '';
  return adPad_(d.getHours()) + ':' + adPad_(d.getMinutes());
}

function adHoursDiff_(a, b) {
  if (!a || !b) return 0;
  return (b - a) / 3600000;
}

function adLocFromFlag_(flag) {
  const m = String(flag || '').match(/\(([^)]+)\)/);
  return m ? m[1].trim() : '';
}

// ── Iterate through a dd-mm-yyyy date range, call fn(dateStr) for each day ──
function adEachDate_(startStr, endStr, fn) {
  const s = adParseDate_(startStr);
  const e = adParseDate_(endStr);
  if (!s || !e || e < s) return;
  for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
    fn(adDateToStr_(new Date(d)));
  }
}

// ============================================================
// MASTER FUNCTION
// ============================================================

function rebuildAttendanceDaily() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const adSheet    = ss.getSheetByName('Attendance_Daily');
  const manSheet   = ss.getSheetByName('Manual_Attendance_Aug');
  const hrSheet    = ss.getSheetByName('HR Maya');
  const shSheet    = ss.getSheetByName('Shifts');
  const logSheet   = ss.getSheetByName('Attendance_Log');
  const lvSheet    = ss.getSheetByName('Leave_Grants');

  if (!adSheet)  throw new Error('Sheet not found: Attendance_Daily');
  if (!manSheet) throw new Error('Sheet not found: Manual_Attendance_Aug');
  if (!hrSheet)  throw new Error('Sheet not found: HR Maya');
  if (!logSheet) throw new Error('Sheet not found: Attendance_Log');

  // ── Headers ──────────────────────────────────────────────────────────
  const HEADERS = [
    'Employee ID', 'Name', 'Date', 'Location', 'Shift',
    'Check In', 'Break Out', 'Break In', 'Check Out',
    'Presence (hrs)', 'Late In (min)', 'Early Out (min)', 'Break Over (min)',
    'OT Hours', 'Day Type', 'Source', 'Status', 'Flags'
  ];
  adSheet.clearContents();
  adSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  adSheet.setFrozenRows(1);

  // ── HR Maya roster ───────────────────────────────────────────────────
  const hrCols = adColMap_(hrSheet);
  const HR_ID  = hrCols['Employee ID']    ?? 0;
  const HR_NM  = hrCols['Name (English)'] ?? hrCols['Name'] ?? 1;
  const HR_LOC = hrCols['Location']       ?? hrCols['Branch'] ?? -1;
  const HR_SH  = hrCols['Shift']          ?? -1;
  const HR_ST  = hrCols['Status']         ?? -1;

  const hrData = hrSheet.getLastRow() > 1
    ? hrSheet.getRange(2, 1, hrSheet.getLastRow() - 1, hrSheet.getLastColumn()).getValues() : [];

  const employees = [];
  const empMap    = {};
  hrData.forEach(row => {
    const id = String(row[HR_ID] || '').trim();
    const st = HR_ST >= 0 ? String(row[HR_ST] || '').trim() : 'Active';
    if (!id || id === 'Employee ID' || st === 'Inactive') return;
    const emp = {
      id,
      name     : String(row[HR_NM]  || '').trim(),
      location : HR_LOC >= 0 ? String(row[HR_LOC] || '').trim() : '',
      shift    : HR_SH  >= 0 ? String(row[HR_SH]  || '').trim() : '',
    };
    employees.push(emp);
    empMap[id] = emp;
  });
  if (!employees.length) throw new Error('No active employees in HR Maya.');

  // ── Shifts — auto-present + normal hours ─────────────────────────────
  const autoShifts = new Set();
  const shiftHours = {};
  if (shSheet && shSheet.getLastRow() > 1) {
    const sc    = adColMap_(shSheet);
    const SH_NM = sc['Shift']        ?? sc['Name'] ?? 0;
    const SH_AU = sc['Auto Present'] ?? sc['AutoPresent'] ?? -1;
    const SH_NH = sc['Normal Hours'] ?? sc['Normal Hrs']  ?? -1;
    shSheet.getRange(2, 1, shSheet.getLastRow() - 1, shSheet.getLastColumn()).getValues()
      .forEach(r => {
        const name = String(r[SH_NM] || '').trim();
        if (!name) return;
        if (SH_AU >= 0 && String(r[SH_AU] || '').toLowerCase() === 'yes') autoShifts.add(name);
        if (SH_NH >= 0) { const h = parseFloat(r[SH_NH]); if (!isNaN(h) && h > 0) shiftHours[name] = h; }
      });
  }

  // ── Granted Leaves — build per-employee per-date lookup ─────────────
  // leaveMap["EMPID|dd-mm-yyyy"] → leave type (Annual / Sick / etc.)
  const leaveMap = {};
  if (lvSheet && lvSheet.getLastRow() > 1) {
    const lc      = adColMap_(lvSheet);
    const LV_EMP  = lc['Employee ID'] ?? 1;
    const LV_TYPE = lc['Type']        ?? 3;
    const LV_ST   = lc['Start']       ?? 4;
    const LV_EN   = lc['End']         ?? 5;
    const LV_STAT = lc['Status']      ?? 7;

    lvSheet.getRange(2, 1, lvSheet.getLastRow() - 1, lvSheet.getLastColumn()).getValues()
      .forEach(row => {
        const stat  = String(row[LV_STAT] || '').trim();
        if (stat !== 'Approved') return;                // only approved grants
        const empId = String(row[LV_EMP] || '').trim();
        const type  = String(row[LV_TYPE] || '').trim();
        const start = adNormDd_(row[LV_ST]);
        const end   = adNormDd_(row[LV_EN]);
        if (!empId || !start || !end) return;
        adEachDate_(start, end, dateStr => {
          leaveMap[empId + '|' + dateStr] = type;
        });
      });
  } else {
    Logger.log('WARNING: Leave_Grants sheet not found — leave status not applied.');
  }

  // ╔══════════════════════════════════════════════════════════════╗
  // ║  PART 1 — 22–31 Aug from Manual_Attendance_Aug              ║
  // ╚══════════════════════════════════════════════════════════════╝

  const manCols  = adColMap_(manSheet);
  const M_ID     = manCols['Employee ID'] ?? 0;
  const M_LOC    = manCols['Location']    ?? 1;
  const M_SHIFT  = manCols['Shift']       ?? 2;
  const M_DATE   = manCols['Date']        ?? 3;
  const M_STATUS = manCols['Status']      ?? 4;

  const manData = manSheet.getLastRow() > 1
    ? manSheet.getRange(2, 1, manSheet.getLastRow() - 1, manSheet.getLastColumn()).getValues() : [];

  const manLookup = {};
  manData.forEach(row => {
    const empId  = String(row[M_ID]     || '').trim();
    const date   = adNormDd_(row[M_DATE]);
    const status = String(row[M_STATUS] || '').trim();
    if (!empId || !date || !status) return;
    manLookup[empId + '|' + date] = {
      location : String(row[M_LOC]   || '').trim(),
      shift    : String(row[M_SHIFT] || '').trim(),
      status,
    };
  });

  const augRows = [];
  const AUG_START = new Date(2026, 7, 22);
  const AUG_END   = new Date(2026, 7, 31);

  for (let d = new Date(AUG_START); d <= AUG_END; d.setDate(d.getDate() + 1)) {
    const dateStr = adDateToStr_(new Date(d));
    const weekend = adIsWeekend_(new Date(d));

    employees.forEach(emp => {
      const key      = emp.id + '|' + dateStr;
      const manual   = manLookup[key];
      const isAuto   = autoShifts.has(emp.shift);
      const leaveType= leaveMap[key];

      let location, shift, status, dayType, source, flags = '';

      if (manual) {
        // 1 — Manual supervisor entry always wins
        location = manual.location || emp.location;
        shift    = manual.shift    || emp.shift;
        status   = manual.status;
        dayType  = status === 'Present' ? 'Full Day' : status === 'Absent' ? 'Absent' : 'Day Off';
        source   = 'Manual_Aug';

      } else if (leaveType) {
        // 2 — Approved leave (no manual entry)
        location = emp.location; shift = emp.shift;
        status   = 'On Leave'; dayType = 'On Leave';
        source   = 'Granted_Leave'; flags = leaveType;

      } else if (isAuto && !weekend) {
        // 3 — Head Office working day
        location = emp.location; shift = emp.shift;
        status   = 'Present'; dayType = 'Full Day'; source = 'Auto_HeadOffice';

      } else if (isAuto && weekend) {
        // 4 — Head Office weekend
        location = emp.location; shift = emp.shift;
        status   = 'Day Off'; dayType = 'Day Off'; source = 'Weekend';

      } else {
        // 5 — Default fallback
        location = emp.location; shift = emp.shift;
        if (emp.location === 'Driver Location') {
          // Drivers work every day — no automatic day off
          status = 'Present'; dayType = 'Full Day'; source = 'Driver_Default';
        } else {
          status = 'Day Off'; dayType = 'Day Off'; source = 'Default';
        }
      }

      augRows.push([
        emp.id, emp.name, dateStr, location, shift,
        '', '', '', '', '', '', '', '', 0, dayType, source, status, flags
      ]);
    });
  }

  // ╔══════════════════════════════════════════════════════════════╗
  // ║  PART 2 — 1 Sep → YESTERDAY from Attendance_Log            ║
  // ║  Today excluded: sessions are incomplete (second shift      ║
  // ║  may not have started). Daily trigger adds yesterday each   ║
  // ║  morning.                                                   ║
  // ╚══════════════════════════════════════════════════════════════╝

  const logCols = adColMap_(logSheet);
  const L_TS    = logCols['Timestamp']     ?? 0;
  const L_ID    = logCols['Employee ID']   ?? 1;
  const L_DATE  = logCols['Date']          ?? 3;
  const L_STAGE = logCols['Stage']         ?? 4;
  const L_BRANCH= logCols['Branch']        ?? logCols['Location'] ?? -1;
  const L_FLAG  = logCols['Location Flag'] ?? logCols['Geo Flag'] ?? -1;

  const logData = logSheet.getLastRow() > 1
    ? logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues() : [];

  const SEP_START = new Date(2026, 8, 1);

  // Compile up to YESTERDAY only — today's sessions are incomplete
  const YESTERDAY = new Date();
  YESTERDAY.setDate(YESTERDAY.getDate() - 1);
  YESTERDAY.setHours(23, 59, 59, 0);

  const sepDates = [];
  const _seen    = new Set();
  for (let d = new Date(SEP_START); d <= YESTERDAY; d.setDate(d.getDate() + 1)) {
    const ds = adDateToStr_(new Date(d));
    if (!_seen.has(ds)) { _seen.add(ds); sepDates.push(ds); }
  }

  // Group kiosk punches by employee (Sep onwards only)
  const punchMap = {};
  logData.forEach(row => {
    const ts = adParseTs_(row[L_TS]);
    if (!ts || ts < SEP_START) return;
    const empId = String(row[L_ID]    || '').trim();
    const stage = String(row[L_STAGE] || '').trim();
    if (!empId || !stage) return;
    let location = '';
    if (L_BRANCH >= 0) location = String(row[L_BRANCH] || '').trim();
    if (!location && L_FLAG >= 0) location = adLocFromFlag_(String(row[L_FLAG] || ''));
    if (!location && empMap[empId])  location = empMap[empId].location;
    const dateStr = adNormDd_(row[L_DATE]) || adDateToStr_(ts);
    if (!punchMap[empId]) punchMap[empId] = [];
    punchMap[empId].push({ ts, date: dateStr, stage, location });
  });

  const kioskRows  = [];
  const MAX_HRS    = 20; // flag presence > 20h as suspect

  employees.forEach(emp => {
    const punches = (punchMap[emp.id] || []).sort((a, b) => a.ts - b.ts);
    const isAuto  = autoShifts.has(emp.shift);
    const normHrs = shiftHours[emp.shift] || 9;

    // Build completed sessions map
    const sessions = {};
    let open = null;

    punches.forEach(p => {
      if (p.stage === 'Check In') {
        open = { date: p.date, location: p.location || emp.location,
                 checkIn: p.ts, breaks: [], checkOut: null };
      } else if (p.stage === 'Break Out' && open) {
        open.breaks.push({ out: p.ts, in: null });
      } else if (p.stage === 'Break In' && open && open.breaks.length > 0) {
        const last = open.breaks[open.breaks.length - 1];
        if (!last.in) last.in = p.ts;
      } else if (p.stage === 'Check Out' && open) {
        open.checkOut = p.ts;
        if (open.checkOut > open.checkIn) sessions[open.date] = open;
        open = null;
      } else if (p.stage === 'Day Off') {
        sessions[p.date] = { dayOff: true, date: p.date, location: p.location || emp.location };
        open = null;
      }
    });
    // Open sessions intentionally not captured — see header comment.

    sepDates.forEach(dateStr => {
      const sess     = sessions[dateStr];
      const d        = adParseDate_(dateStr);
      const weekend  = d ? adIsWeekend_(d) : false;
      const key      = emp.id + '|' + dateStr;
      const leaveType= leaveMap[key];

      if (sess && sess.dayOff) {
        // Explicit kiosk Day Off punch
        kioskRows.push([emp.id, emp.name, dateStr,
          sess.location, emp.shift, '', '', '', '',
          '', '', '', '', 0, 'Day Off', 'Kiosk', 'Day Off', '']);

      } else if (sess && sess.checkIn && sess.checkOut) {
        // Completed kiosk session — compute presence + OT
        const loc   = sess.location || emp.location;
        const ciStr = adHHmm_(sess.checkIn);
        const coStr = adHHmm_(sess.checkOut);

        let breakHrs = 0, boStr = '', biStr = '';
        sess.breaks.forEach((br, i) => {
          if (br.out && br.in && br.in > br.out) {
            breakHrs += adHoursDiff_(br.out, br.in);
            if (i === 0) { boStr = adHHmm_(br.out); biStr = adHHmm_(br.in); }
          }
        });

        const rawHrs      = adHoursDiff_(sess.checkIn, sess.checkOut);
        const presenceHrs = Math.max(0, rawHrs - breakHrs);
        const isSuspect   = presenceHrs > MAX_HRS;
        const presenceFin = isSuspect ? 0 : presenceHrs;
        const otHrs       = isSuspect ? 0 : Math.max(0, presenceHrs - normHrs);
        const dayType     = isSuspect    ? 'Check Punches'
                          : presenceHrs >= normHrs * 0.5 ? 'Full Day' : 'Half Day';
        const flags       = isSuspect
          ? 'REVIEW: computed ' + rawHrs.toFixed(1) + 'h — check punch order' : '';

        kioskRows.push([emp.id, emp.name, dateStr, loc, emp.shift,
          ciStr, boStr, biStr, coStr,
          +presenceFin.toFixed(3), '', '', '',
          +otHrs.toFixed(3),
          dayType, 'Kiosk', 'Present', flags]);

      } else if (leaveType) {
        // No kiosk data, but approved leave covers this date
        kioskRows.push([emp.id, emp.name, dateStr, emp.location, emp.shift,
          '', '', '', '', '', '', '', '', 0,
          'On Leave', 'Granted_Leave', 'On Leave', leaveType]);

      } else if (isAuto && !weekend) {
        // Head Office — auto-present on working days
        kioskRows.push([emp.id, emp.name, dateStr, emp.location, emp.shift,
          '', '', '', '', normHrs, '', '', '', 0,
          'Full Day', 'Auto_HeadOffice', 'Present', '']);

      } else if (isAuto && weekend) {
        // Head Office — Fri/Sat off
        kioskRows.push([emp.id, emp.name, dateStr, emp.location, emp.shift,
          '', '', '', '', '', '', '', '', 0,
          'Day Off', 'Weekend', 'Day Off', '']);

      } else {
        // Default — no data, not auto-present, not on leave
        if (emp.location === 'Driver Location') {
          // Drivers work every day — no automatic day off
          kioskRows.push([emp.id, emp.name, dateStr, emp.location, emp.shift,
            '', '', '', '', '', '', '', '', 0,
            'Full Day', 'Driver_Default', 'Present', '']);
        } else {
          kioskRows.push([emp.id, emp.name, dateStr, emp.location, emp.shift,
            '', '', '', '', '', '', '', '', 0,
            'Day Off', 'Default', 'Day Off', '']);
        }
      }
    });
  });

  // ── Write to Attendance_Daily ─────────────────────────────────
  const allRows = [...augRows, ...kioskRows];
  if (allRows.length > 0) {
    adSheet.getRange(2, 1, allRows.length, HEADERS.length).setValues(allRows);
  }

  // ── Summary ──────────────────────────────────────────────────
  const leaveRowsAug   = augRows.filter(r => r[15] === 'Granted_Leave').length;
  const leaveRowsSep   = kioskRows.filter(r => r[15] === 'Granted_Leave').length;
  const driverRowsAug  = augRows.filter(r => r[15] === 'Driver_Default').length;
  const driverRowsSep  = kioskRows.filter(r => r[15] === 'Driver_Default').length;
  const reviewFlags    = kioskRows.filter(r => String(r[17]).includes('REVIEW')).length;

  const msg =
    'Attendance_Daily rebuilt (v3).\n\n' +
    '  22–31 Aug  :  ' + augRows.length + ' rows\n' +
    '    Manual_Aug      : ' + augRows.filter(r => r[15] === 'Manual_Aug').length + '\n' +
    '    Granted_Leave   : ' + leaveRowsAug + '\n' +
    '    Driver_Default  : ' + driverRowsAug + '\n' +
    '    Auto_HeadOffice : ' + augRows.filter(r => r[15] === 'Auto_HeadOffice').length + '\n' +
    '    Weekend (HO)    : ' + augRows.filter(r => r[15] === 'Weekend').length + '\n' +
    '    Default Day Off : ' + augRows.filter(r => r[15] === 'Default').length + '\n\n' +
    '  1 Sep – yesterday :  ' + kioskRows.length + ' rows\n' +
    '    Kiosk Present   : ' + kioskRows.filter(r => r[16] === 'Present' && r[15] === 'Kiosk').length + '\n' +
    '    Granted_Leave   : ' + leaveRowsSep + '\n' +
    '    Driver_Default  : ' + driverRowsSep + '\n' +
    '    Auto_HeadOffice : ' + kioskRows.filter(r => r[15] === 'Auto_HeadOffice').length + '\n' +
    '    Day Off         : ' + kioskRows.filter(r => r[16] === 'Day Off').length + '\n' +
    '    REVIEW flags    : ' + reviewFlags + '\n\n' +
    '  Total  : ' + allRows.length + ' rows  |  ' + employees.length +
    ' employees  |  ' + sepDates.length + ' Sep days (up to yesterday)\n' +
    '  Leave rows: ' + (leaveRowsAug + leaveRowsSep) +
    '  |  Driver default Present: ' + (driverRowsAug + driverRowsSep);

  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    allRows.length + ' rows. Leave rows: ' + (leaveRowsAug + leaveRowsSep) +
    '. Review flags: ' + reviewFlags,
    'Attendance_Daily rebuilt', 15);
}