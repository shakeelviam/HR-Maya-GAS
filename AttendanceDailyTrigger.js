// ============================================================
// AttendanceDailyTrigger.gs  — v4 (Roster-integrated)
// ------------------------------------------------------------
// Daily trigger: runs at 4 AM, writes YESTERDAY's attendance
// into Attendance_Daily.
//
// SETUP (run once from the editor):
//   1. Open Apps Script editor
//   2. Select setupDailyTrigger → Run
//   3. Approve permissions when prompted
//   That installs a daily 4 AM trigger on dailyAttendanceTrigger().
//
// Priority order per employee per day (highest → lowest):
//   1. Manual_Aug entry        (22–31 Aug backfill only)
//   2. Kiosk Day Off self-report
//        → checked against published roster:
//          roster says Working → Absent + discrepancy logged
//          roster says Day Off / no roster → Day Off (as self-reported)
//   3. Kiosk completed session (Check In + Check Out — punch wins)
//        → checked against published roster:
//          roster said Day Off, or shift mismatch → Present stands,
//          discrepancy logged for supervisor review
//   4. Approved Granted Leave
//   5. Auto_HeadOffice (working day)
//   6. Weekend (Head Office only, Fri/Sat)
//   7. Roster says Working, no punch  → Absent  (Roster_Absent)
//   8. Roster says Day Off            → Day Off (Roster_DayOff)
//   9. Driver_Default (Driver Location, no roster needed)
//  10. No roster entry at all         → "No Roster" (No_Roster)
//        — replaces the old blind "Default Day Off" guess.
//
// Depends on RosterBackend.gs being deployed in the same project:
//   rbGetPublishedRosterForDate_(dateStr) -> { empId: {status, shift, location} }
//   rbLogDiscrepancy_(empId, name, location, date, rosterSaid, actual)
// ============================================================

// ── Install the 4 AM daily trigger ────────────────────────────────────
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyAttendanceTrigger') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('dailyAttendanceTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  Logger.log('Daily trigger installed: dailyAttendanceTrigger runs at 4 AM every day.');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily 4 AM trigger installed.', 'Trigger set up', 8
  );
}

// ── The trigger function — called automatically at 4 AM ────────────────
function dailyAttendanceTrigger() {
  // Apply any approved transfers whose effective date has arrived, and
  // expire any still-pending requests whose date has passed unactioned.
  // Runs FIRST — before compiling attendance — so a transfer effective
  // today is already live by the time today's data gets processed.
  if (typeof applyPendingTransfers_ === 'function') {
    Logger.log('dailyAttendanceTrigger: applying pending transfers');
    applyPendingTransfers_();
  }

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const dateStr = adDateToStr_(yesterday);
  Logger.log('dailyAttendanceTrigger: processing ' + dateStr);
  writeDayToAttendanceDaily_(dateStr);
  Logger.log('dailyAttendanceTrigger: done for ' + dateStr);
}

// ── Execute due transfers, expire stale pending requests ────────────
// Depends on RosterBackend.gs's rbApplyTransfer_ being in the same
// project (same global scope — Apps Script shares all .gs files).
function applyPendingTransfers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pending_Transfers');
  if (!sheet || sheet.getLastRow() < 2) return;

  const rc = adColMap_(sheet);
  const todayStr = adDateToStr_(new Date());
  const todayObj = adParseDate_(todayStr);
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();

  let executed = 0, expired = 0;

  data.forEach((row, i) => {
    const rowNum = i + 2;
    const status = String(row[rc['Status']]||'').trim();
    const effDateStr = adNormDd_(row[rc['Effective Date']]);
    const effObj = adParseDate_(effDateStr);
    if (!effObj) return;

    if (status === 'Approved' && effObj <= todayObj) {
      const empId    = String(row[rc['Employee ID']]||'').trim();
      const empName  = String(row[rc['Name']]||'').trim();
      const fromLoc  = String(row[rc['From Location']]||'').trim();
      const toLoc    = String(row[rc['To Location']]||'').trim();
      const reason   = String(row[rc['Reason']]||'').trim();

      if (typeof rbApplyTransfer_ === 'function') {
        const result = rbApplyTransfer_(empId, empName, fromLoc, toLoc, effDateStr, 'System (scheduled)', reason);
        if (result.success) {
          sheet.getRange(rowNum, rc['Status']+1).setValue('Completed');
          executed++;
        } else {
          Logger.log('applyPendingTransfers_: failed for ' + empId + ' — ' + result.error);
        }
      } else {
        Logger.log('applyPendingTransfers_: rbApplyTransfer_ not found — is RosterBackend.gs deployed?');
      }

    } else if (status === 'Pending Approval' && effObj < todayObj) {
      // Effective date passed with nobody having approved/rejected it.
      sheet.getRange(rowNum, rc['Status']+1).setValue('Expired');
      expired++;
    }
  });

  Logger.log('applyPendingTransfers_: ' + executed + ' executed, ' + expired + ' expired.');
}

// ── Manual backfill: run for any past date from the editor ────────────
function manualDayWrite() {
  const dateStr = '05-09-2026'; // ← change this date and run
  Logger.log('manualDayWrite: processing ' + dateStr);
  writeDayToAttendanceDaily_(dateStr);
  Logger.log('manualDayWrite: done for ' + dateStr);
}

// ══════════════════════════════════════════════════════════════════════
// CORE: write one day's attendance rows into Attendance_Daily
// ══════════════════════════════════════════════════════════════════════
function writeDayToAttendanceDaily_(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const adSheet  = ss.getSheetByName('Attendance_Daily');
  const hrSheet  = ss.getSheetByName('HR Maya');
  const shSheet  = ss.getSheetByName('Shifts');
  const logSheet = ss.getSheetByName('Attendance_Log');
  const lvSheet  = ss.getSheetByName('Leave_Grants');

  if (!adSheet || !hrSheet || !logSheet) {
    throw new Error('Required sheet missing (Attendance_Daily / HR Maya / Attendance_Log)');
  }

  const HEADERS = [
    'Employee ID', 'Name', 'Date', 'Location', 'Shift',
    'Check In', 'Break Out', 'Break In', 'Check Out',
    'Presence (hrs)', 'Late In (min)', 'Early Out (min)', 'Break Over (min)',
    'OT Hours', 'Day Type', 'Source', 'Status', 'Flags'
  ];

  if (adSheet.getLastRow() === 0) {
    adSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    adSheet.setFrozenRows(1);
  }

  // ── Step 1: Delete existing rows for this date (idempotent) ─────────
  const lastRow = adSheet.getLastRow();
  if (lastRow > 1) {
    const dateCol = adColMap_(adSheet)['Date'] ?? 2;
    const data    = adSheet.getRange(2, 1, lastRow - 1, adSheet.getLastColumn()).getValues();
    const toDelete = [];
    data.forEach((row, i) => {
      if (adNormDd_(row[dateCol]) === dateStr) toDelete.push(i + 2);
    });
    for (let i = toDelete.length - 1; i >= 0; i--) {
      adSheet.deleteRow(toDelete[i]);
    }
  }

  // ── Step 2: Read HR Maya roster (employee master, not shift-roster) ──
  const hrCols = adColMap_(hrSheet);
  const HR_ID  = hrCols['Employee ID']    ?? 0;
  const HR_NM  = hrCols['Name (English)'] ?? hrCols['Name'] ?? 1;
  const HR_LOC = hrCols['Location']       ?? hrCols['Branch'] ?? -1;
  const HR_SH  = hrCols['Shift']          ?? -1;
  const HR_ST  = hrCols['Status']         ?? -1;
  const hrData = hrSheet.getLastRow() > 1
    ? hrSheet.getRange(2, 1, hrSheet.getLastRow()-1, hrSheet.getLastColumn()).getValues() : [];

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

  // ── Step 3: Shifts ────────────────────────────────────────────────────
  const autoShifts = new Set();
  const shiftHours = {};
  if (shSheet && shSheet.getLastRow() > 1) {
    const sc    = adColMap_(shSheet);
    const SH_NM = sc['Shift'] ?? sc['Name'] ?? 0;
    const SH_AU = sc['Auto Present'] ?? sc['AutoPresent'] ?? -1;
    const SH_NH = sc['Normal Hours'] ?? sc['Normal Hrs']  ?? -1;
    shSheet.getRange(2,1,shSheet.getLastRow()-1,shSheet.getLastColumn()).getValues()
      .forEach(r => {
        const name = String(r[SH_NM]||'').trim(); if (!name) return;
        if (SH_AU>=0 && String(r[SH_AU]||'').toLowerCase()==='yes') autoShifts.add(name);
        if (SH_NH>=0) { const h=parseFloat(r[SH_NH]); if(!isNaN(h)&&h>0) shiftHours[name]=h; }
      });
  }

  // ── Step 4: Leave grants covering this date ───────────────────────────
  const leaveOnDate = {};
  if (lvSheet && lvSheet.getLastRow() > 1) {
    const lc      = adColMap_(lvSheet);
    const LV_EMP  = lc['Employee ID'] ?? 1;
    const LV_TYPE = lc['Type']        ?? 3;
    const LV_ST   = lc['Start']       ?? 4;
    const LV_EN   = lc['End']         ?? 5;
    const LV_STAT = lc['Status']      ?? 7;
    lvSheet.getRange(2,1,lvSheet.getLastRow()-1,lvSheet.getLastColumn()).getValues()
      .forEach(row => {
        if (String(row[LV_STAT]||'').trim() !== 'Approved') return;
        const empId = String(row[LV_EMP]||'').trim();
        const type  = String(row[LV_TYPE]||'').trim();
        const start = adNormDd_(row[LV_ST]);
        const end   = adNormDd_(row[LV_EN]);
        if (!empId || !start || !end) return;
        if (adDateInRange_(dateStr, start, end)) leaveOnDate[empId] = type;
      });
  }

  // ── Step 4b: Published roster for this date (NEW) ──────────────────
  // rbGetPublishedRosterForDate_ lives in RosterBackend.gs, same project.
  let rosterMap = {};
  if (typeof rbGetPublishedRosterForDate_ === 'function') {
    rosterMap = rbGetPublishedRosterForDate_(dateStr);
  }

  // ── Step 5: Kiosk punches for this date ──────────────────────────────
  const logCols = adColMap_(logSheet);
  const L_TS    = logCols['Timestamp']     ?? 0;
  const L_ID    = logCols['Employee ID']   ?? 1;
  const L_DATE  = logCols['Date']          ?? 3;
  const L_STAGE = logCols['Stage']         ?? 4;
  const L_BRANCH= logCols['Branch']        ?? logCols['Location'] ?? -1;
  const L_FLAG  = logCols['Location Flag'] ?? logCols['Geo Flag'] ?? -1;

  const punchMap = {};
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2,1,logSheet.getLastRow()-1,logSheet.getLastColumn()).getValues()
      .forEach(row => {
        const rowDate = adNormDd_(row[L_DATE]);
        if (rowDate !== dateStr) return;
        const empId = String(row[L_ID]||'').trim();
        const stage = String(row[L_STAGE]||'').trim();
        if (!empId || !stage) return;
        let location = '';
        if (L_BRANCH>=0) location = String(row[L_BRANCH]||'').trim();
        if (!location && L_FLAG>=0) location = adLocFromFlag_(String(row[L_FLAG]||''));
        if (!location && empMap[empId]) location = empMap[empId].location;
        const ts = adParseTs_(row[L_TS]);
        if (!ts) return;
        if (!punchMap[empId]) punchMap[empId] = [];
        punchMap[empId].push({ ts, stage, location });
      });
  }

  // ── Step 6: Build rows ────────────────────────────────────────────────
  const targetDate = adParseDate_(dateStr);
  const weekend    = targetDate ? adIsWeekend_(targetDate) : false;
  const MAX_HRS    = 20;
  const rows       = [];

  employees.forEach(emp => {
    const isAuto     = autoShifts.has(emp.shift);
    const normHrs    = shiftHours[emp.shift] || 9;
    const leaveType  = leaveOnDate[emp.id];
    const rosterInfo = rosterMap[emp.id]; // { status, shift, location } | undefined
    const punches    = (punchMap[emp.id] || []).sort((a,b) => a.ts - b.ts);

    let session = null;
    let open    = null;
    punches.forEach(p => {
      if (p.stage === 'Check In') {
        open = { location: p.location || emp.location, checkIn: p.ts, breaks: [], checkOut: null };
      } else if (p.stage === 'Break Out' && open) {
        open.breaks.push({ out: p.ts, in: null });
      } else if (p.stage === 'Break In' && open && open.breaks.length > 0) {
        const last = open.breaks[open.breaks.length-1]; if (!last.in) last.in = p.ts;
      } else if (p.stage === 'Check Out' && open) {
        open.checkOut = p.ts;
        if (open.checkOut > open.checkIn) session = open;
        open = null;
      } else if (p.stage === 'Day Off') {
        session = { dayOff: true, location: p.location || emp.location };
        open = null;
      }
    });

    let row;

    // ── 1: Kiosk Day Off self-report — check against roster ───────────
    if (session && session.dayOff) {
      if (rosterInfo && rosterInfo.status === 'Working') {
        // Roster says they were scheduled — self-report doesn't override.
        row = [emp.id, emp.name, dateStr, emp.location, rosterInfo.shift || emp.shift,
          '','','','','','','','',0,'Absent','Roster_Absent','Absent',
          'Self-reported Day Off via kiosk, but roster scheduled Working'];
        if (typeof rbLogDiscrepancy_ === 'function') {
          rbLogDiscrepancy_(emp.id, emp.name, rosterInfo.location || emp.location, dateStr,
            'Working', 'Self-reported Day Off (kiosk)');
        }
      } else {
        row = [emp.id, emp.name, dateStr, session.location, emp.shift,
          '','','','','','','','',0,'Day Off','Kiosk','Day Off',''];
      }

    // ── 2: Completed kiosk session — punch always wins ─────────────────
    } else if (session && session.checkIn && session.checkOut) {
      const loc   = session.location || emp.location;
      const ciStr = adHHmm_(session.checkIn);
      const coStr = adHHmm_(session.checkOut);
      let breakHrs=0, boStr='', biStr='';
      session.breaks.forEach((br,i) => {
        if (br.out && br.in && br.in > br.out) {
          breakHrs += adHoursDiff_(br.out, br.in);
          if (i===0) { boStr=adHHmm_(br.out); biStr=adHHmm_(br.in); }
        }
      });
      const rawHrs      = adHoursDiff_(session.checkIn, session.checkOut);
      const presenceHrs = Math.max(0, rawHrs - breakHrs);
      const isSuspect   = presenceHrs > MAX_HRS;
      const presenceFin = isSuspect ? 0 : presenceHrs;
      const otHrs       = isSuspect ? 0 : Math.max(0, presenceHrs - normHrs);
      const dayType     = isSuspect ? 'Check Punches'
                        : presenceHrs >= normHrs*0.5 ? 'Full Day' : 'Half Day';
      let flags         = isSuspect ? 'REVIEW: ' + rawHrs.toFixed(1) + 'h — check punches' : '';

      // Roster mismatch check (informational — punch still wins)
      if (rosterInfo) {
        if (rosterInfo.status === 'Day Off') {
          if (typeof rbLogDiscrepancy_ === 'function') {
            rbLogDiscrepancy_(emp.id, emp.name, rosterInfo.location || emp.location, dateStr,
              'Day Off', 'Present ' + ciStr + '–' + coStr);
          }
          flags = (flags ? flags + '; ' : '') + 'Roster said Day Off — punch overrides';
        } else if (rosterInfo.shift && rosterInfo.shift !== emp.shift) {
          if (typeof rbLogDiscrepancy_ === 'function') {
            rbLogDiscrepancy_(emp.id, emp.name, rosterInfo.location || emp.location, dateStr,
              'Working (' + rosterInfo.shift + ')', 'Present, shift on file: ' + emp.shift);
          }
        }
      }

      row = [emp.id, emp.name, dateStr, loc, emp.shift,
        ciStr, boStr, biStr, coStr,
        +presenceFin.toFixed(3), '','','', +otHrs.toFixed(3),
        dayType, 'Kiosk', 'Present', flags];

    // ── 3: Approved leave ────────────────────────────────────────────
    } else if (leaveType) {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'On Leave','Granted_Leave','On Leave', leaveType];

    // ── 4/5: Head Office auto-present ────────────────────────────────
    } else if (isAuto && !weekend) {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','',normHrs,'','','',0,'Full Day','Auto_HeadOffice','Present',''];

    } else if (isAuto && weekend) {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'Day Off','Weekend','Day Off',''];

    // ── 6/7: Published roster, no punch ──────────────────────────────
    } else if (rosterInfo && rosterInfo.status === 'Working') {
      row = [emp.id, emp.name, dateStr, emp.location, rosterInfo.shift || emp.shift,
        '','','','','','','','',0,'Absent','Roster_Absent','Absent',
        'Rostered to work — no punch recorded'];

    } else if (rosterInfo && rosterInfo.status === 'Day Off') {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'Day Off','Roster_DayOff','Day Off',''];

    // ── 8: Driver default (no roster needed — works every day) ───────
    } else if (emp.location === 'Driver Location') {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'Full Day','Driver_Default','Present',''];

    // ── 9: No roster published at all — explicit, not a guess ────────
    } else {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'No Roster','No_Roster','No Roster',
        'No roster published for this date'];
    }

    rows.push(row);
  });

  // ── Step 7: Append rows ───────────────────────────────────────────────
  if (rows.length > 0) {
    const startRow = adSheet.getLastRow() + 1;
    adSheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
  }

  Logger.log('writeDayToAttendanceDaily_: ' + dateStr + ' — ' + rows.length + ' rows written.');
}

// ── Helpers (shared with AttendanceDaily.gs) ───────────────────────────

function adColMap_(sheet) {
  const hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map  = {};
  hdrs.forEach((h,i) => { const k=String(h).trim(); if(k) map[k]=i; });
  return map;
}

function adNormDd_(val) {
  if (!val) return '';
  if (val instanceof Date) return adPad_(val.getDate())+'-'+adPad_(val.getMonth()+1)+'-'+val.getFullYear();
  const s = String(val).trim();
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return m[3]+'-'+m[2]+'-'+m[1];
  const m2= s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if(m2) return adPad_(m2[1])+'-'+adPad_(m2[2])+'-'+m2[3];
  return s;
}
function adPad_(n) { return String(n).padStart(2,'0'); }
function adDateToStr_(d) { return adPad_(d.getDate())+'-'+adPad_(d.getMonth()+1)+'-'+d.getFullYear(); }
function adParseDate_(s) {
  const m = String(s).match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
}
function adParseTs_(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const s = String(val).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):?(\d{0,2})/);
  if (m) return new Date(+m[3],+m[2]-1,+m[1],+m[4],+m[5],m[6]?+m[6]:0);
  const d = new Date(s); return isNaN(d.getTime()) ? null : d;
}
function adIsWeekend_(d) { const w=d.getDay(); return w===5||w===6; }
function adHHmm_(d) { if(!d||isNaN(d.getTime())) return ''; return adPad_(d.getHours())+':'+adPad_(d.getMinutes()); }
function adHoursDiff_(a,b) { if(!a||!b) return 0; return (b-a)/3600000; }
function adLocFromFlag_(flag) { const m=String(flag||'').match(/\(([^)]+)\)/); return m?m[1].trim():''; }

function adDateInRange_(dateStr, startStr, endStr) {
  function toIso(s) { const p=s.split('-'); return p[2]+'-'+p[1]+'-'+p[0]; }
  const d=toIso(dateStr), s=toIso(startStr), e=toIso(endStr);
  return d>=s && d<=e;
}