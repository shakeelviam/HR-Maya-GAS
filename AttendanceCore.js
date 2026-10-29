// ============================================================
// AttendanceCore.gs — THE single definition of attendance logic
// ============================================================
// WHY THIS FILE EXISTS
// Before this, SEVEN files each carried their own copy of "a Check In
// opens a session, a Check Out closes it, a second Check In orphans the
// first":
//   AttendanceDaily.js, AttendanceDailyTrigger.js, AttendanceOT.js,
//   BackfillCheckouts.js, Code.js, OtSubmit.js, Payroll.js
// Seven copies meant seven chances to disagree — and they did.
// AttendanceDaily.js guessed "Default Day Off" where
// AttendanceDailyTrigger.js said "No Roster": same punches, different
// answer, and which one you got depended on which function was called.
//
// There were also SEVEN near-identical date normalisers under different
// names (normDayStr_, adNormDd_, vsNormDd_, rbNormDd_, sc_normDay_,
// ad_normDate_, otr_norm_). A header named "Shift Name" instead of
// "Shift" broke silently because each file guessed differently.
//
// EVERYTHING in this file is behaviour-preserving. It is lifted from
// AttendanceDailyTrigger.js, which is the most evolved of the seven and
// the one the 4 AM compile actually uses. Migrating a caller to these
// functions must NOT change its output. The test is simple: recompile a
// date range before and after, and diff. An empty diff proves it.
//
// NAMING: every public function is ac*. Nothing here collides with the
// existing helpers, so files can be migrated ONE AT A TIME and the old
// helpers removed only once nothing calls them.
// ============================================================

// ── Constants: previously hardcoded in several files, often differently ──
const AC_MAX_SESSION_HOURS = 20;   // above this, punches are wrong, not heroic
const AC_DEFAULT_NORM_HRS  = 9;    // when a shift has no Normal Hours on file
const AC_HALF_DAY_RATIO    = 0.5;  // below half the shift = Half Day

const AC_STAGES = ['Check In', 'Break Out', 'Break In', 'Check Out', 'Day Off'];
const AC_PUNCH_STAGES = ['Check In', 'Break Out', 'Break In', 'Check Out'];

// ════════════════════════════════════════════════════════════
// 0. REPEATED CHECK-IN TAPS — the rule (Shakeel, 21-09-2026)
// ════════════════════════════════════════════════════════════
// "Wherever there is a duplicate check-in we take the first one which is
//  aligning with a shift timing."
//
// A Check In while a session is still open used to ORPHAN the earlier one
// and start the day again. So a man who checked in at 12:26 and then tapped
// Check In again at 20:10 and 21:11 was measured from 21:11 — 0.31 hours,
// Half Day, for a full day's work (Ajay Gurung, 20-09).
//
// Now: a Check In less than AC_REPEAT_CHECKIN_HOURS after the open one is a
// REPEAT TAP on the same shift, not a new shift. All the taps form one run,
// and the session starts at the FIRST tap in the run that lines up with a
// shift start — the start of any shift at his location, or his own shift.
// "Lines up" = from AC_ALIGN_BEFORE_MIN before the start to
// AC_ALIGN_AFTER_MIN after it. If no tap lines up, the first tap wins.
//
// Beyond AC_REPEAT_CHECKIN_HOURS the earlier check-in is a forgotten
// check-out from a previous shift, and it is orphaned exactly as before.
const AC_REPEAT_CHECKIN_HOURS = 16;
const AC_ALIGN_BEFORE_MIN     = 60;
const AC_ALIGN_AFTER_MIN      = 90;

// A completed session shorter than this is tap noise, not work: a man
// fumbling the kiosk at 18:32 → 18:33 after a full day. It used to REPLACE
// the real session as "the" session of the day (Laxman, 20-09: 0.01h).
const AC_MIN_SESSION_MIN      = 5;

function acIsRepeatCheckIn_(openStartTs, newTs) {
  if (!openStartTs || !newTs) return false;
  const h = (newTs - openStartTs) / 3600000;
  return h >= 0 && h < AC_REPEAT_CHECKIN_HOURS;
}

// Does a clock time line up with any of the given shift starts (minutes
// after midnight)? Wraps midnight, so 23:40 lines up with a 00:30 start.
function acAlignsWithShift_(ts, startsMin) {
  if (!ts || !startsMin || !startsMin.length) return false;
  const t = ts.getHours() * 60 + ts.getMinutes();
  return startsMin.some(s => {
    let d = t - s;                       // minutes AFTER the start
    if (d > 720) d -= 1440;
    if (d < -720) d += 1440;
    return d >= -AC_ALIGN_BEFORE_MIN && d <= AC_ALIGN_AFTER_MIN;
  });
}

// run: [{ts, row?, ...}] in time order. Returns the element to start from.
function acPickRunStart_(run, startsMin) {
  if (!run || !run.length) return null;
  for (let i = 0; i < run.length; i++) {
    if (acAlignsWithShift_(run[i].ts, startsMin)) return run[i];
  }
  return run[0];
}

// Shift starts (minutes) that count as "a shift timing" for one man: every
// shift belonging to his location — by the location→shift map when one is
// deployed, else by name prefix ("Jabriya First Shift" belongs to
// "Jabriya") — plus his own shift from HR Maya. Also any location the punch
// itself was made at, so a man working another branch is judged by that
// branch's shifts.
//   shifts: acShifts_(ss)   locMap: optional { locationLower: [shiftName] }
function acShiftStartsFor_(shifts, locMap, locations, ownShift) {
  const out = [], seen = {};
  const add = (name) => {
    const s = shifts && shifts[name];
    if (!s || !s.start) return;
    const m = s.start.h * 60 + s.start.m;
    if (!seen[m]) { seen[m] = true; out.push(m); }
  };
  if (ownShift) add(String(ownShift).trim());
  (locations || []).forEach(loc => {
    const l = String(loc || '').trim().toLowerCase();
    if (!l) return;
    if (locMap && locMap[l]) locMap[l].forEach(add);
    Object.keys(shifts || {}).forEach(n => { if (n.toLowerCase().indexOf(l) === 0) add(n); });
  });
  return out;
}

// Build a lookup once per run: empId -> function(extraLocations) -> starts.
function acShiftStartsIndex_(ss) {
  const shifts = acShifts_(ss);
  let locMap = null;
  try { if (typeof getLocationShiftsMap_ === 'function') locMap = getLocationShiftsMap_(ss); } catch (e) {}
  const emps = {};
  acEmployees_(ss, true).forEach(e => { emps[e.id] = e; });
  return function (empId, extraLocations) {
    const e = emps[empId] || {};
    const locs = [e.location].concat(extraLocations || []);
    return acShiftStartsFor_(shifts, locMap, locs, e.shift);
  };
}

// ════════════════════════════════════════════════════════════
// 1. DATES AND TIMES — one implementation, not seven
// ════════════════════════════════════════════════════════════

// Any date cell → 'dd-mm-yyyy'. Handles Date objects, dd-mm-yyyy,
// dd/mm/yyyy, yyyy-mm-dd, and anything Date can parse.
function acNormDate_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return '';
    return acPad_(v.getDate()) + '-' + acPad_(v.getMonth() + 1) + '-' + v.getFullYear();
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return acPad_(m[1]) + '-' + acPad_(m[2]) + '-' + m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2,'0') + '-' + acPad_(m[2]) + '-' + m[1];
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : acNormDate_(d);
}

function acPad_(n) { return String(n).padStart(2, '0'); }

// 'dd-mm-yyyy' → Date at midnight. Returns null if unparseable —
// callers MUST handle null rather than assuming a date came back.
function acParseDate_(s) {
  if (Object.prototype.toString.call(s) === '[object Date]') {
    return isNaN(s.getTime()) ? null : new Date(s.getFullYear(), s.getMonth(), s.getDate());
  }
  const m = String(s).trim().match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
}

// Any timestamp cell → Date. Attendance_Log stores 'dd/MM/yyyy HH:mm:ss'.
function acParseTs_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? null : v;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):?(\d{0,2})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], m[6] ? +m[6] : 0);
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function acDateStr_(d) { return acNormDate_(d); }
function acHHmm_(d) {
  if (!d || isNaN(d.getTime())) return '';
  return acPad_(d.getHours()) + ':' + acPad_(d.getMinutes());
}
function acHoursBetween_(a, b) { return (!a || !b) ? 0 : (b - a) / 3600000; }
function acIsWeekend_(d) { const w = d.getDay(); return w === 5 || w === 6; }  // Fri/Sat
function acDateInRange_(dateStr, fromStr, toStr) {
  const iso = (s) => { const p = String(s).split('-'); return p[2] + p[1] + p[0]; };
  return iso(dateStr) >= iso(fromStr) && iso(dateStr) <= iso(toStr);
}

// Header lookup that tolerates case, padding and non-breaking spaces —
// and accepts ALTERNATIVES, because 'Shift' vs 'Shift Name' and
// 'Location' vs 'Branch' have each broken something silently.
// Returns -1 when not found. NEVER silently falls back to column 0.
function acCol_(headerRow, candidates) {
  const norm = (v) => String(v == null ? '' : v)
    .replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  const H = headerRow.map(norm);
  for (const c of candidates) {
    const i = H.indexOf(norm(c));
    if (i !== -1) return i;
  }
  return -1;
}

function acColMap_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// ════════════════════════════════════════════════════════════
// 2. SESSIONIZING — the one rule
// ════════════════════════════════════════════════════════════
// Lifted verbatim in behaviour from AttendanceDailyTrigger.js.
//
// THE RULE, stated once so it can never drift again:
//   Check In   → opens a session. If one is already open it is ORPHANED
//                (the employee never checked out) and replaced.
//   Break Out  → starts a break, only if a session is open.
//   Break In   → closes the most recent break that has no end yet.
//   Check Out  → closes the session. Only counts if it is AFTER the
//                check-in; a checkout before its check-in is corrupt.
//   Day Off    → ends any open session and marks the day as Day Off.
//
// Punches MUST be sorted by timestamp. Pass them unsorted and you get
// nonsense, so this sorts defensively rather than trusting the caller.
//
// Returns:
//   { completed: [...], open: session|null, orphaned: [...], dayOff: {...}|null }
//
// 'completed' — Check In and Check Out both present, checkout after checkin
// 'open'      — a Check In with no Check Out yet (still working, or forgot)
// 'orphaned'  — a Check In superseded by another Check In or a Day Off
// 'dayOff'    — an explicit Day Off punch
// startsMin (optional): shift starts for this man — see acShiftStartsFor_.
// Without it, a run of repeated taps starts at its first tap.
function acSessionize_(punches, startsMin) {
  const now = new Date();
  // A punch in the future has not happened. The only way one gets into the
  // log is a check-out written at a shift end that had not come yet — and
  // it made the kiosk believe the man had already left.
  const p = (punches || []).filter(x => !x.ts || x.ts <= now)
                           .slice().sort((a, b) => a.ts - b.ts);
  const completed = [], orphaned = [], noise = [];
  let open = null, dayOff = null;

  const settleStart = (o) => {
    const st = acPickRunStart_(o.run, startsMin) || o.run[0];
    o.checkIn = st.ts; o.checkInRow = st.row;
    if (st.location) o.location = st.location;
    o.repeatTaps = o.run.length - 1;
    // Breaks taken before the chosen start belong to nothing.
    o.breaks = o.breaks.filter(b => !b.out || b.out >= st.ts);
  };

  p.forEach(x => {
    const stage = String(x.stage || '').trim();

    if (stage === 'Check In') {
      if (open && acIsRepeatCheckIn_(open.run[0].ts, x.ts)) {
        open.run.push({ ts: x.ts, row: x.row, location: x.location });   // repeat tap
        return;
      }
      if (open) { settleStart(open); orphaned.push(open); }            // never checked out
      open = { checkIn: x.ts, checkInRow: x.row, location: x.location,
               run: [{ ts: x.ts, row: x.row, location: x.location }],
               breaks: [], checkOut: null, checkOutRow: null };

    } else if (stage === 'Break Out') {
      if (open) open.breaks.push({ out: x.ts, outRow: x.row, in: null, inRow: null });

    } else if (stage === 'Break In') {
      if (open && open.breaks.length) {
        const last = open.breaks[open.breaks.length - 1];
        if (!last.in) { last.in = x.ts; last.inRow = x.row; }
      }

    } else if (stage === 'Check Out') {
      if (open) {
        settleStart(open);
        open.checkOut = x.ts;
        open.checkOutRow = x.row;
        // A checkout BEFORE its checkin is corrupt data, not a session.
        if (open.checkOut <= open.checkIn) orphaned.push(open);
        else if ((open.checkOut - open.checkIn) / 60000 < AC_MIN_SESSION_MIN) noise.push(open);
        else completed.push(open);
        open = null;
      }

    } else if (stage === 'Day Off') {
      // A Day Off tap does not end a shift that began the same day — if he
      // then checks out, the completed session outranks the tap, as it
      // always has. Only an open session from an EARLIER day is closed.
      if (open && acNormDate_(open.run[0].ts) !== acNormDate_(x.ts)) {
        settleStart(open); orphaned.push(open); open = null;
      }
      dayOff = { ts: x.ts, row: x.row, location: x.location };
    }
    // Break Out / Break In with no open session are ignored — the same
    // convention every one of the seven copies already used.
  });

  if (open) settleStart(open);
  return { completed: completed, open: open, orphaned: orphaned, dayOff: dayOff, noise: noise };
}

// ── Presence for one completed session ──────────────────────────────
// Returns { rawHours, breakHours, presence, suspect, flags }
//
// 'suspect' means the computed span exceeds AC_MAX_SESSION_HOURS. That is
// almost always a wrong DATE rather than a heroic shift — it is how the
// 62-hour sessions got in. Presence is forced to 0 so the number never
// silently reaches payroll, and the row is flagged for a human.
function acPresence_(session) {
  if (!session || !session.checkIn || !session.checkOut) {
    return { rawHours: 0, breakHours: 0, presence: 0, suspect: false, flags: '' };
  }
  const raw = acHoursBetween_(session.checkIn, session.checkOut);

  let breakHours = 0;
  (session.breaks || []).forEach(br => {
    if (br.out && br.in && br.in > br.out) breakHours += acHoursBetween_(br.out, br.in);
  });

  // BREAK IS NOT DEDUCTED. Policy (20-09-2026): the shift's normal hours
  // are hours ON SITE, already inclusive of the break, so taking it off
  // again charges the man twice. breakHours is still returned — the break
  // times are a record of when it was taken, not a deduction.
  const presence = raw;
  const suspect  = presence > AC_MAX_SESSION_HOURS;

  return {
    rawHours: raw,
    breakHours: breakHours,
    presence: suspect ? 0 : presence,
    suspect: suspect,
    flags: suspect ? ('REVIEW: ' + raw.toFixed(1) + 'h — check punches') : '',
  };
}

// ── Overtime and day type from presence ─────────────────────────────
function acOvertime_(presence, normalHours) {
  const norm = (normalHours && normalHours > 0) ? normalHours : AC_DEFAULT_NORM_HRS;
  return Math.max(0, presence - norm);
}

function acDayType_(presence, normalHours, suspect) {
  if (suspect) return 'Check Punches';
  const norm = (normalHours && normalHours > 0) ? normalHours : AC_DEFAULT_NORM_HRS;
  return presence >= norm * AC_HALF_DAY_RATIO ? 'Full Day' : 'Half Day';
}

// ════════════════════════════════════════════════════════════
// 3. READING Attendance_Log — one reader
// ════════════════════════════════════════════════════════════
// Returns { empId: [ {ts,row,stage,location,auth,date}, ... ] }
// Optional from/to ('dd-mm-yyyy') filter by the punch TIMESTAMP.
//
// IMPORTANT: detection of dangling sessions must read FULL history, not
// a filtered window — a real Check Out sitting just outside the window
// would otherwise look like a missing one. Pass no range for that.
function acReadPunches_(ss, fromStr, toStr) {
  const log = ss.getSheetByName('Attendance_Log');
  const out = {};
  if (!log || log.getLastRow() < 2) return out;

  const H = acColMap_(log);
  const cId    = acCol_(H, ['Employee ID']);
  const cTs    = acCol_(H, ['Timestamp']);
  const cDate  = acCol_(H, ['Date']);
  const cStage = acCol_(H, ['Stage']);
  const cAuth  = acCol_(H, ['Auth', 'PIN Used']);
  const cBr    = acCol_(H, ['Branch', 'Location']);
  const cFlag  = acCol_(H, ['Geo Flag']);
  if (cId < 0 || cTs < 0 || cStage < 0) {
    Logger.log('acReadPunches_: Attendance_Log missing Employee ID / Timestamp / Stage.');
    return out;
  }

  const from = fromStr ? acParseDate_(fromStr) : null;
  const to   = toStr   ? acParseDate_(toStr)   : null;

  const data = log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn()).getValues();
  data.forEach((r, i) => {
    const id = String(r[cId] || '').trim();
    if (!id) return;
    const ts = acParseTs_(r[cTs]);
    if (!ts) return;
    if (from && ts < from) return;
    if (to) { const end = new Date(to.getTime() + 86399999); if (ts > end) return; }

    let location = cBr >= 0 ? String(r[cBr] || '').trim() : '';
    if (!location && cFlag >= 0) location = acLocFromFlag_(String(r[cFlag] || ''));

    (out[id] = out[id] || []).push({
      row: i + 2,
      ts: ts,
      date: cDate >= 0 ? acNormDate_(r[cDate]) : acNormDate_(ts),
      stage: String(r[cStage] || '').trim(),
      auth: cAuth >= 0 ? String(r[cAuth] || '').trim() : '',
      location: location,
    });
  });

  Object.keys(out).forEach(k => out[k].sort((a, b) => a.ts - b.ts));
  return out;
}

// Geo flag strings look like 'OK (Ardiya)' — pull the branch out.
function acLocFromFlag_(flag) {
  const m = String(flag || '').match(/\(([^)]+)\)/);
  return m ? m[1].trim() : '';
}

// ════════════════════════════════════════════════════════════
// 4. WRITING a punch — one guarded path
// ════════════════════════════════════════════════════════════
// Every writer should go through this. Today the duplicate-stage check,
// the 20-hour sanity check and the valid-stage list are copy-pasted into
// some writers and missing from others — which is how a 62-hour session
// and 21 junk punches got in, and how three separate files rejected
// 'Day Off' as an invalid stage.
//
// opts: { empId, date:'dd-mm-yyyy', stage, time:'HH:mm', location,
//         auth, by, allowDuplicate:false, skipSanity:false }
// Returns { success, message|error, row }
function acWritePunch_(ss, opts) {
  try {
    const empId = String(opts.empId || '').trim();
    const stage = String(opts.stage || '').trim();
    const dateStr = acNormDate_(opts.date);

    if (!empId)   return { success: false, error: 'Employee ID required.' };
    if (!dateStr) return { success: false, error: 'Valid date required (dd-mm-yyyy).' };
    if (AC_STAGES.indexOf(stage) === -1) {
      return { success: false, error: 'Invalid stage "' + stage + '". Valid: ' + AC_STAGES.join(', ') };
    }

    // Day Off carries no meaningful time; punches do.
    const time = (stage === 'Day Off') ? '00:00' : String(opts.time || '').trim();
    if (stage !== 'Day Off' && !/^\d{1,2}:\d{2}$/.test(time)) {
      return { success: false, error: 'Time required as HH:MM for ' + stage + '.' };
    }

    const log = ss.getSheetByName('Attendance_Log');
    if (!log) return { success: false, error: 'Attendance_Log not found.' };
    const H = acColMap_(log);
    const cId    = acCol_(H, ['Employee ID']);
    const cName  = acCol_(H, ['Employee Name']);
    const cTs    = acCol_(H, ['Timestamp']);
    const cDate  = acCol_(H, ['Date']);
    const cStage = acCol_(H, ['Stage']);
    const cAuth  = acCol_(H, ['Auth', 'PIN Used']);
    const cBr    = acCol_(H, ['Branch', 'Location']);
    const cFlag  = acCol_(H, ['Geo Flag']);

    // ── Guard 1: no duplicate stage on the same employee-day ──
    if (!opts.allowDuplicate) {
      const data = log.getRange(2, 1, Math.max(log.getLastRow() - 1, 1), log.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][cId] || '').trim() === empId &&
            acNormDate_(data[i][cDate]) === dateStr &&
            String(data[i][cStage] || '').trim() === stage) {
          return { success: false, error: empId + ' already has "' + stage + '" on ' + dateStr + '.' };
        }
      }
    }

    // Build the timestamp
    const d = acParseDate_(dateStr);
    const tp = time.split(':');
    const ts = new Date(d.getFullYear(), d.getMonth(), d.getDate(), +tp[0], +tp[1], 0);

    // ── Guard 2: would this create an implausible session? ──
    // This is what a 62-hour session looks like before it exists.
    if (!opts.skipSanity && (stage === 'Check Out' || stage === 'Check In')) {
      const punches = acReadPunches_(ss)[empId] || [];
      const probe = punches.concat([{ ts: ts, stage: stage, row: -1, location: '' }]);
      const s = acSessionize_(probe);
      const bad = s.completed.filter(x => acHoursBetween_(x.checkIn, x.checkOut) > AC_MAX_SESSION_HOURS);
      if (bad.length) {
        const hrs = acHoursBetween_(bad[0].checkIn, bad[0].checkOut).toFixed(1);
        return { success: false, error: 'This would create a ' + hrs +
          '-hour session — almost certainly the wrong DATE rather than a real shift. ' +
          'Check the date, not just the time.' };
      }
    }

    const row = new Array(H.length).fill('');
    row[cId] = empId;
    if (cName >= 0) row[cName] = acEmpName_(ss, empId);
    if (cTs   >= 0) row[cTs]   = Utilities.formatDate(ts, acTz_(), 'dd/MM/yyyy HH:mm:ss');
    if (cDate >= 0) row[cDate] = dateStr;
    row[cStage] = stage;
    if (cAuth >= 0) row[cAuth] = opts.auth || 'Manual';
    if (cBr   >= 0) row[cBr]   = opts.location || '';
    if (cFlag >= 0) row[cFlag] = opts.geoFlag || 'MANUAL';

    log.appendRow(row);
    if (cDate >= 0) log.getRange(log.getLastRow(), cDate + 1).setNumberFormat('@');

    if (typeof logAmend_ === 'function') {
      logAmend_(ss, opts.by || 'System', 'ADD', empId, dateStr, stage,
                'via acWritePunch_ at ' + time);
    }

    return { success: true, row: log.getLastRow(),
             message: stage + ' recorded for ' + empId + ' on ' + dateStr +
                      (stage === 'Day Off' ? '.' : ' at ' + time + '.') };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════
// 5. SHARED LOOKUPS
// ════════════════════════════════════════════════════════════

function acTz_() { return Session.getScriptTimeZone() || 'Asia/Kuwait'; }

function acEmpName_(ss, empId) {
  const hr = ss.getSheetByName('HR Maya');
  if (!hr) return empId;
  const H = acColMap_(hr);
  const cId = acCol_(H, ['Employee ID']);
  const cNm = acCol_(H, ['Name (English)', 'Name', 'Employee Name']);
  if (cId < 0 || cNm < 0) return empId;
  const data = hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues();
  const r = data.find(x => String(x[cId] || '').trim() === String(empId).trim());
  return r ? String(r[cNm] || '').trim() : empId;
}

// Every active employee with the fields attendance logic needs.
function acEmployees_(ss, includeInactive) {
  const hr = ss.getSheetByName('HR Maya');
  const out = [];
  if (!hr || hr.getLastRow() < 2) return out;
  const H = acColMap_(hr);
  const cId = acCol_(H, ['Employee ID']);
  const cNm = acCol_(H, ['Name (English)', 'Name']);
  const cLoc = acCol_(H, ['Location', 'Branch']);
  const cSh = acCol_(H, ['Shift']);
  const cSt = acCol_(H, ['Status']);
  if (cId < 0) { Logger.log('acEmployees_: no Employee ID column in HR Maya.'); return out; }

  hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().forEach(r => {
    const id = String(r[cId] || '').trim();
    if (!id) return;
    const status = cSt >= 0 ? String(r[cSt] || '').trim() : 'Active';
    if (!includeInactive && status.toLowerCase() === 'inactive') return;
    out.push({
      id: id,
      name: cNm >= 0 ? String(r[cNm] || '').trim() : id,
      location: cLoc >= 0 ? String(r[cLoc] || '').trim() : '',
      shift: cSh >= 0 ? String(r[cSh] || '').trim() : '',
      status: status,
    });
  });
  return out;
}

// Shift definitions. NOTE the header is 'Shift Name' — several files
// looked for 'Shift' and silently fell back to column 0.
// Returns { name: { start:{h,m}, end:{h,m}, overnight, normalHours, autoPresent } }
function acShifts_(ss) {
  const sh = ss.getSheetByName('Shifts');
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const H = acColMap_(sh);
  const cNm = acCol_(H, ['Shift Name', 'Shift', 'Name']);
  const cSt = acCol_(H, ['Start']);
  const cEn = acCol_(H, ['End']);
  const cNh = acCol_(H, ['Normal Hours', 'Normal Hrs']);
  const cAu = acCol_(H, ['Auto Present', 'AutoPresent']);
  if (cNm < 0) { Logger.log('acShifts_: no shift-name column found.'); return out; }

  const t = (v) => {
    if (v == null || v === '') return null;
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return isNaN(v.getTime()) ? null : { h: v.getHours(), m: v.getMinutes() };
    }
    const m = String(v).trim().match(/^(\d{1,2}):(\d{2})/);
    return m ? { h: +m[1], m: +m[2] } : null;
  };

  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(r => {
    const name = String(r[cNm] || '').trim();
    if (!name) return;
    const start = cSt >= 0 ? t(r[cSt]) : null;
    const end   = cEn >= 0 ? t(r[cEn]) : null;
    const nh    = cNh >= 0 ? parseFloat(r[cNh]) : NaN;
    out[name] = {
      start: start, end: end,
      overnight: (start && end) ? ((end.h * 60 + end.m) < (start.h * 60 + start.m)) : false,
      normalHours: (!isNaN(nh) && nh > 0) ? nh : AC_DEFAULT_NORM_HRS,
      autoPresent: cAu >= 0 && String(r[cAu] || '').trim().toLowerCase() === 'yes',
    };
  });
  return out;
}

// ════════════════════════════════════════════════════════════
// 6. SELF-TEST — prove the core matches the old behaviour
// ════════════════════════════════════════════════════════════
// Run acSelfTest() after deploying. It checks the sessionizer against
// cases drawn from real incidents in this system.
function acSelfTest() {
  const L = [];
  let pass = 0, fail = 0;
  const T = (name, got, want) => {
    const ok = String(got) === String(want);
    ok ? pass++ : fail++;
    L.push((ok ? '  PASS  ' : '  FAIL  ') + name +
           (ok ? '' : '   got=' + got + ' want=' + want));
  };
  const ts = (h, m, day) => new Date(2026, 8, day || 3, h, m, 0);

  // Normal shift with a break
  let s = acSessionize_([
    { ts: ts(8,0),  stage: 'Check In' },
    { ts: ts(12,0), stage: 'Break Out' },
    { ts: ts(13,0), stage: 'Break In' },
    { ts: ts(17,0), stage: 'Check Out' },
  ]);
  T('normal shift → 1 completed', s.completed.length, 1);
  T('normal shift → presence 9h (break NOT deducted)', acPresence_(s.completed[0]).presence, 9);
  T('break still reported separately', acPresence_(s.completed[0]).breakHours, 1);

  // REPEATED CHECK-IN TAPS — the 21-09 rule.
  // Ajay Gurung, Jabriya, 20-09: 12:26 in, 20:10 in, 21:11 in, 21:30 out.
  // He used to be measured from 21:11 — 0.31h, Half Day.
  const JAB = [12*60+30, 18*60+30];          // Jabriya First / Second starts
  s = acSessionize_([
    { ts: ts(12,26,20), stage: 'Check In' },
    { ts: ts(20,10,20), stage: 'Check In' },
    { ts: ts(21,11,20), stage: 'Check In' },
    { ts: ts(21,30,20), stage: 'Check Out' },
  ], JAB);
  T('repeat taps are one session, not three', s.completed.length + '/' + s.orphaned.length, '1/0');
  T('starts at the first tap that lines up (12:26)', acHHmm_(s.completed[0].checkIn), '12:26');
  T('Ajay: 12:26 → 21:30 is 9h 04m', Math.round(acPresence_(s.completed[0]).presence*60), 544);
  T('repeat taps counted', s.completed[0].repeatTaps, 2);

  // First tap does NOT line up, a later one does → the later one.
  s = acSessionize_([
    { ts: ts(15,5), stage: 'Check In' },           // nowhere near a shift start
    { ts: ts(18,40), stage: 'Check In' },          // Second Shift +10m
    { ts: ts(3,30,4), stage: 'Check Out' },
  ], JAB);
  T('skips a tap that lines up with no shift', acHHmm_(s.completed[0].checkIn), '18:40');

  // No tap lines up → the first tap.
  s = acSessionize_([
    { ts: ts(15,5), stage: 'Check In' }, { ts: ts(16,0), stage: 'Check In' },
    { ts: ts(23,0), stage: 'Check Out' },
  ], JAB);
  T('nothing lines up → first tap', acHHmm_(s.completed[0].checkIn), '15:05');

  // More than 16h apart is a forgotten check-out, not a repeat tap.
  s = acSessionize_([
    { ts: ts(12,30,3), stage: 'Check In' },
    { ts: ts(12,30,4), stage: 'Check In' },
    { ts: ts(21,30,4), stage: 'Check Out' },
  ], JAB);
  T('24h later still orphans the old check-in', s.orphaned.length + '/' + s.completed.length, '1/1');

  // Tap noise: a 1-minute session after the real one is not "the" session.
  s = acSessionize_([
    { ts: ts(8,56), stage: 'Check In' },  { ts: ts(18,26), stage: 'Check Out' },
    { ts: ts(18,32), stage: 'Check In' }, { ts: ts(18,33), stage: 'Check Out' },
  ], [9*60]);
  T('1-minute session is noise', s.completed.length + '/' + s.noise.length, '1/1');

  // A punch in the future has not happened.
  s = acSessionize_([
    { ts: ts(12,30), stage: 'Check In' },
    { ts: new Date(Date.now() + 3*3600000), stage: 'Check Out' },
  ], JAB);
  T('future check-out ignored → still open', (s.open ? 'open' : 'closed') + '/' + s.completed.length, 'open/0');

  // A Day Off tap does not end a shift that started the same day…
  s = acSessionize_([
    { ts: ts(11,32), stage: 'Check In' },
    { ts: ts(17,46), stage: 'Day Off' },
    { ts: ts(20,30), stage: 'Check Out' },
  ], [11*60+30]);
  T('same-day Day Off does not kill the shift', s.completed.length, 1);
  T('Day Off still detected', s.dayOff ? 'yes' : 'no', 'yes');
  // …but it does close one left open from an earlier day.
  s = acSessionize_([
    { ts: ts(8,0,2), stage: 'Check In' },
    { ts: ts(9,0,3), stage: 'Day Off' },
  ]);
  T('Day Off closes an EARLIER day\'s open session', s.orphaned.length, 1);

  // Overnight: check out on the NEXT day is still one session
  s = acSessionize_([
    { ts: ts(17,30,9),  stage: 'Check In' },
    { ts: ts(2,30,10),  stage: 'Check Out' },
  ]);
  T('overnight shift → 1 completed', s.completed.length, 1);
  T('overnight presence 9h', acPresence_(s.completed[0]).presence, 9);

  // The live case that set the policy: Kristal Ghalan, Jabriya, 03-09-2026.
  s = acSessionize_([
    { ts: new Date(2026,8,3,12,39), stage: 'Check In' },
    { ts: new Date(2026,8,3,15,0),  stage: 'Break Out' },
    { ts: new Date(2026,8,3,15,50), stage: 'Break In' },
    { ts: new Date(2026,8,4,3,46),  stage: 'Check Out' },
  ]);
  const kp = acPresence_(s.completed[0]);
  T('12:39 → 03:46 is 15.117h on site', Math.round(kp.presence*1000)/1000, 15.117);
  T('50m break recorded, not deducted', Math.round(kp.breakHours*60), 50);
  T('OT against a 9h shift = 6.117h', Math.round(acOvertime_(kp.presence, 9)*1000)/1000, 6.117);

  // The 62-hour incident: must be flagged, presence forced to 0
  s = acSessionize_([
    { ts: new Date(2026,8,1,5,8),  stage: 'Check In' },
    { ts: new Date(2026,8,3,19,25), stage: 'Check Out' },
  ]);
  const p = acPresence_(s.completed[0]);
  T('62h session flagged suspect', p.suspect, true);
  T('62h session presence forced 0', p.presence, 0);

  // Checkout BEFORE checkin is corrupt, not a session
  s = acSessionize_([
    { ts: ts(17,0), stage: 'Check In' },
    { ts: ts(14,0), stage: 'Check Out' },
  ]);
  T('checkout before checkin → not completed', s.completed.length, 0);

  // Date normalisation — the seven-helper problem
  T('dd-mm-yyyy',  acNormDate_('03-09-2026'), '03-09-2026');
  T('dd/mm/yyyy',  acNormDate_('3/9/2026'),   '03-09-2026');
  T('yyyy-mm-dd',  acNormDate_('2026-09-03'), '03-09-2026');
  T('Date object', acNormDate_(new Date(2026,8,3)), '03-09-2026');

  // Header lookup must tolerate the real-world variants
  T('header "Shift Name"', acCol_(['Shift Name','Start'], ['Shift Name','Shift']), 0);
  T('header padded/upper', acCol_([' EMPLOYEE ID '], ['Employee ID']), 0);
  T('header missing → -1', acCol_(['Foo'], ['Employee ID']), -1);

  // Day type
  T('half day', acDayType_(3, 9, false), 'Half Day');
  T('full day', acDayType_(8, 9, false), 'Full Day');
  T('suspect → Check Punches', acDayType_(0, 9, true), 'Check Punches');

  const out = '\n════════ AttendanceCore self-test ════════\n' + L.join('\n') +
    '\n\n  ' + pass + ' passed, ' + fail + ' failed\n' +
    (fail ? '  FIX BEFORE MIGRATING ANY CALLER.\n' : '  Core is behaving correctly.\n');
  Logger.log(out);
  return out;
}