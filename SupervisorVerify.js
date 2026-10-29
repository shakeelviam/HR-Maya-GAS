// ============================================================
// SupervisorVerify.gs — period verification by supervisors
// ------------------------------------------------------------
// WHY THIS EXISTS
// OT claims sit unpaid because the punch data cannot settle most of them,
// and a run of days cannot be resolved from punches either. Only the person
// who was there knows what actually happened. This asks each supervisor to
// confirm, for their own location, ONLY the things that are in dispute.
//
// ------------------------------------------------------------
// ONE RECORD, ONE PLACE  (this is the whole point of this rewrite)
//
// There is no Supervisor_Verification tab any more. A verdict is written
// onto the record it is a verdict ABOUT:
//
//   OT verdict        -> onto the OT_Pending ROW ITSELF, in four columns
//                        (Supervisor Verdict / Hours Confirmed /
//                         Verified By / Verified On).
//                        The claim and the supervisor's answer are one row.
//                        Status stays 'Pending' — a supervisor does not pay
//                        anything. The OT compiler still decides, but now
//                        with the statement attached.
//
//   "Worked" verdict  -> real punches into Attendance_Log
//                        (Check In / Check Out). Attendance_Log is already
//                        the source of truth; Attendance_Daily is derived
//                        from it. Nothing new to look at.
//
//   "Day Off" verdict -> a 'Day Off' punch into Attendance_Log, exactly
//                        the row the kiosk writes when a man taps Day Off.
//
//   "Absent" verdict  -> a row in Manual_Attendance_Aug (Status = Absent),
//                        the tab supervisors already use to state a day.
//                        Absent is also what the compile produces when
//                        nothing is recorded, so this verdict cannot change
//                        the outcome — it only stops the asking.
//
// NOTHING ALREADY ANSWERED IS SHOWN AGAIN. The list is recomputed from live
// data on every load — there is no "mark as done" step anyone can forget:
//   an OT claim disappears once its verdict column is filled;
//   a day disappears once Attendance_Log holds a supervisor-authored row
//   for that man on that date, or Manual_Attendance_Aug states it.
// Days a supervisor already fixed with the correction tool below therefore
// disappear too — they were answered by the fix.
//
// ------------------------------------------------------------
// TWO PERIODS, DELIBERATELY DIFFERENT
//   22-31 Aug : supervisors marked Present/Absent only; no punch times
//               exist at all. OT CLAIMS ONLY — there is no attendance
//               discrepancy to show, because there is no punch data to
//               disagree with.
//   01 Sep on : real kiosk punches. OT claims AND attendance exceptions.
//
// ROUTER CASES: getSupervisorVerifyList, saveSupervisorVerdicts
// Run svSelfCheck() once after deploying.
// ============================================================

const SV_MANUAL_UNTIL = '31-08-2026';   // no punch times exist on or before this
const SV_FROM_DEFAULT = '22-08-2026';
const SV_AUTH         = 'Supervisor verified';
const SV_MANUAL_TAB   = 'Manual_Attendance_Aug';
const SV_RECOMPILE_BUDGET_MS = 200000;  // stop recompiling before the 6-min wall

// Auth values that mean "a human already stated this day". Any of these on a
// punch row removes that day from the supervisor's list.
const SV_HUMAN_AUTH = ['supervisor verified', 'supervisor corrected',
                       'supervisor added', 'manual', 'backfilled'];

// Sources that record attendance WITHOUT punch times — the man never punched
// and was never expected to, so there is no missing check-out to chase.
const SV_TIMELESS_SOURCES = ['Manual_Aug', 'Auto_HeadOffice', 'Driver_Default'];

// ── WHO MAY ANSWER (21-09-2026) ──────────────────────────────────────
// Before: anyone who opened the page picked a name from a list and answered
// as that person. 101 answers were a supervisor answering for HIMSELF —
// Keshab confirmed 10 of his own OT claims, Bibek 8, Nilam 7.
// Now: the supervisor proves who he is with his kiosk PIN, he may answer
// only for a location he supervises (Branch_Locations → Supervisors).
// HR may answer at any location.
// 21-09-2026 (Shakeel): a supervisor MAY answer for himself. Every answer
// carries his PIN-proven name, and HR still decides every OT claim on
// OT Review, so his own verdict is a statement, not a payment.
const SV_HR_IDS = ['MT-00060'];

// ── NOTHING IS ANSWERED BEFORE IT HAS HAPPENED (21-09-2026) ──────────
// On 20-09 supervisors answered that same day between 14:29 and 19:19,
// while the shifts were still running. "Worked" wrote check-outs at the
// shift end — 20:30, 21:30, 03:30 the next morning — hours before they
// happened. Men who stayed on lost those hours, and the kiosk, seeing a
// check-out already on file, recorded their real check-out as a Check In.
// A day can now be answered only once its shift is over, plus this grace.
const SV_SHIFT_END_GRACE_MIN = 30;
// When nobody knows his shift: the latest shift in the Shifts sheet ends
// at 03:30, so nothing is still running at 04:00 the next morning.
const SV_UNKNOWN_SHIFT_END = '04:00';

// ── small helpers ────────────────────────────────────────────────────
function sv_pad_(n) { return String(n).padStart(2, '0'); }

function sv_norm_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[1];
  const d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}
function sv_iso_(dStr) { const p = String(dStr).split('-'); return (p[2]||'')+(p[1]||'')+(p[0]||''); }

// dd-mm-yyyy -> Date at midnight local.
function sv_toDate_(dStr) {
  const p = String(dStr).split('-');
  if (p.length !== 3) return null;
  const d = new Date(+p[2], +p[1] - 1, +p[0]);
  return isNaN(d.getTime()) ? null : d;
}
function sv_dateToStr_(d) {
  return sv_pad_(d.getDate()) + '-' + sv_pad_(d.getMonth()+1) + '-' + d.getFullYear();
}

// Hours -> "7h 58m". Carries properly: 7.996h is 8h, NOT "7h 60m".
// (The old version floored the hours and rounded the minutes separately,
//  so 59.76 minutes rounded to 60 and printed beside the floored hour.)
function sv_hm_(h) {
  if (h == null || h === '' || isNaN(Number(h))) return '—';
  h = Number(h);
  const neg = h < 0;
  let mins = Math.round(Math.abs(h) * 60);
  const hh = Math.floor(mins / 60), mm = mins % 60;
  if (!hh && !mm) return '0h';
  return (neg ? '-' : '') + (hh ? hh + 'h' : '') + (mm ? (hh ? ' ' : '') + mm + 'm' : '');
}

// A time-only cell comes back as a Date anchored to 1899-12-30. Take the clock.
function sv_clock_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]')
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz_(), 'HH:mm');
  const m = String(v).trim().match(/(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2,'0') + ':' + m[2] : String(v).trim();
}
function sv_hmToMin_(hm) {
  const m = String(hm || '').match(/^(\d{1,2}):(\d{2})/);
  return m ? (+m[1]) * 60 + (+m[2]) : null;
}

function sv_headerIndex_(sheet) {
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  H.forEach((h, i) => { const k = String(h).trim().toLowerCase(); if (k && !(k in map)) map[k] = i; });
  return { headers: H, at: (name) => (name.toLowerCase() in map) ? map[name.toLowerCase()] : -1 };
}

// ── OT_Pending verdict columns ───────────────────────────────────────
// The verdict belongs on the claim. Four columns are appended once, at the
// end, so the positional reads elsewhere (OtCompile's OTC_COL 0..9) are
// untouched. Returns 0-based indexes.
const SV_OT_VERDICT_HEADERS = ['Supervisor Verdict', 'Hours Confirmed',
                               'Verified By', 'Verified On'];

function sv_otCols_(pen) {
  let hi = sv_headerIndex_(pen);
  const missing = SV_OT_VERDICT_HEADERS.filter(h => hi.at(h) === -1);
  if (missing.length) {
    const start = pen.getLastColumn() + 1;
    pen.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold');
    SpreadsheetApp.flush();
    hi = sv_headerIndex_(pen);
  }
  return {
    verdict: hi.at('Supervisor Verdict'),
    hours:   hi.at('Hours Confirmed'),
    by:      hi.at('Verified By'),
    on:      hi.at('Verified On'),
    width:   pen.getLastColumn(),
  };
}

// ── Shifts ───────────────────────────────────────────────────────────
// { shiftNameLower: { name, start:'HH:mm', end:'HH:mm', normal } }
function sv_shiftMap_(ss) {
  const out = {};
  const sh = ss.getSheetByName('Shifts');
  if (!sh || sh.getLastRow() < 2) return out;
  const hi = sv_headerIndex_(sh);
  let cN = hi.at('Shift Name'); if (cN === -1) cN = hi.at('Shift'); if (cN === -1) cN = hi.at('Name');
  const cS = hi.at('Start'), cE = hi.at('End');
  let cH = hi.at('Normal Hours'); if (cH === -1) cH = hi.at('Normal Hrs');
  if (cN === -1 || cS === -1 || cE === -1) return out;
  sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues().forEach(r => {
    const name = String(r[cN] || '').trim(); if (!name) return;
    const start = sv_clock_(r[cS]), end = sv_clock_(r[cE]);
    if (!start || !end) return;
    const nh = cH >= 0 ? parseFloat(r[cH]) : NaN;
    out[name.toLowerCase()] = { name: name, start: start, end: end,
                                normal: (!isNaN(nh) && nh > 0) ? nh : 9 };
  });
  return out;
}

function sv_locShifts_(ss, location) {
  try {
    if (typeof getLocationShiftsMap_ === 'function') {
      const lm = getLocationShiftsMap_(ss);
      return lm[String(location).toLowerCase()] || [];
    }
  } catch (e) {}
  return [];
}

// Work out which shift a day belongs to. Order: the shift Attendance_Daily
// already recorded; else the location's only shift; else, if there is a real
// check-in, the location shift whose start is nearest to it. Never guesses
// between two shifts with nothing to go on — returns null and says so.
function sv_resolveShift_(shiftMap, locShiftNames, recordedShift, checkInHm) {
  const pick = (n) => shiftMap[String(n || '').toLowerCase()] || null;
  const direct = pick(recordedShift);
  if (direct) return direct;
  const cands = locShiftNames.map(pick).filter(Boolean);
  if (cands.length === 1) return cands[0];
  if (cands.length > 1) {
    const ci = sv_hmToMin_(checkInHm);
    if (ci == null) return null;
    let best = null, bestGap = 1e9;
    cands.forEach(c => {
      const s = sv_hmToMin_(c.start); if (s == null) return;
      let gap = Math.abs(s - ci); if (gap > 720) gap = 1440 - gap;   // wrap midnight
      if (gap < bestGap) { bestGap = gap; best = c; }
    });
    return bestGap <= 240 ? best : null;   // within 4h of a shift start, or don't guess
  }
  return null;
}

// ── When is a day over? ──────────────────────────────────────────────
// The moment the shift on that date ends (next morning for an overnight
// shift), plus the grace. Unknown shift → 04:00 the next morning.
function sv_shiftEndTs_(dStr, shift) {
  const base = sv_toDate_(dStr); if (!base) return null;
  let endMin, overnight;
  if (shift && shift.end) {
    endMin = sv_hmToMin_(shift.end);
    overnight = endMin <= sv_hmToMin_(shift.start);
  } else {
    endMin = sv_hmToMin_(SV_UNKNOWN_SHIFT_END);
    overnight = true;
  }
  return new Date(base.getTime() + (overnight ? 86400000 : 0) + endMin * 60000);
}
function sv_dayEnded_(dStr, shift, now) {
  const end = sv_shiftEndTs_(dStr, shift);
  if (!end) return false;
  return (now || new Date()).getTime() >= end.getTime() + SV_SHIFT_END_GRACE_MIN * 60000;
}
function sv_yesterday_() {
  const t = new Date(); t.setHours(0, 0, 0, 0);
  return sv_dateToStr_(new Date(t.getTime() - 86400000));
}

// "Keshab Alemagar (MT-00028)" → "MT-00028"
function sv_idFromBy_(by) {
  const m = String(by || '').match(/\((MT-\d+)\)\s*$/i);
  return m ? m[1].toUpperCase() : '';
}

// Supervisor IDs per location, from Branch_Locations → Supervisors.
function sv_supervisorsByLocation_(ss) {
  const out = {};
  const sh = ss.getSheetByName('Branch_Locations');
  if (!sh || sh.getLastRow() < 2) return out;
  const hi = sv_headerIndex_(sh);
  let cL = hi.at('Branch'); if (cL === -1) cL = hi.at('Location');
  const cS = hi.at('Supervisors');
  if (cL === -1 || cS === -1) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(r => {
    const loc = String(r[cL] || '').trim().toLowerCase(); if (!loc) return;
    out[loc] = String(r[cS] || '').split(/[,;\s]+/).map(x => x.trim().toUpperCase()).filter(Boolean);
  });
  return out;
}

// Proves who is answering. Returns { ok, id, name, isHr, by } or { ok:false, error }.
// The PIN is the kiosk PIN in HR Maya — the same check loginStaffKiosk makes.
function sv_authSupervisor_(ss, supId, supPin, location) {
  supId = String(supId || '').trim().toUpperCase();
  supPin = String(supPin || '').trim();
  if (!supId || !supPin) return { ok: false, error: 'Select your name and enter your kiosk PIN.' };
  const hr = ss.getSheetByName('HR Maya');
  if (!hr) return { ok: false, error: 'HR Maya not found.' };
  const hi = sv_headerIndex_(hr);
  const cId = hi.at('Employee ID'), cPin = hi.at('PIN'), cSt = hi.at('Status');
  let cNm = hi.at('Name (English)'); if (cNm === -1) cNm = hi.at('Name');
  if (cId === -1 || cPin === -1) return { ok: false, error: 'HR Maya is missing Employee ID or PIN.' };
  const data = hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues();
  let who = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][cId] || '').trim().toUpperCase() !== supId) continue;
    let pin = String(data[i][cPin] == null ? '' : data[i][cPin]).trim();
    if (pin.endsWith('.0')) pin = pin.slice(0, -2);
    if (pin !== supPin) break;
    if (cSt !== -1 && String(data[i][cSt] || '').trim().toLowerCase() === 'inactive') break;
    who = { id: supId, name: cNm !== -1 ? String(data[i][cNm] || '').trim() : supId };
    break;
  }
  if (!who) return { ok: false, error: 'Wrong Employee ID or PIN.' };
  const isHr = SV_HR_IDS.indexOf(supId) !== -1;
  if (!isHr && location) {
    const sups = sv_supervisorsByLocation_(ss)[String(location).trim().toLowerCase()] || [];
    if (sups.indexOf(supId) === -1)
      return { ok: false, error: who.name + ' is not a supervisor at ' + location + '.' };
  }
  return { ok: true, id: supId, name: who.name, isHr: isHr, by: who.name + ' (' + supId + ')' };
}

// Active staff at one location, by Employee ID (upper-case).
function sv_staffAt_(ss, location) {
  const out = {};
  const hr = ss.getSheetByName('HR Maya'); if (!hr || hr.getLastRow() < 2) return out;
  const hi = sv_headerIndex_(hr);
  const cId = hi.at('Employee ID'), cSt = hi.at('Status');
  let cL = hi.at('Location'); if (cL === -1) cL = hi.at('Branch');
  if (cId === -1 || cL === -1) return out;
  const want = String(location || '').trim().toLowerCase();
  hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().forEach(r => {
    if (cSt !== -1 && String(r[cSt] || '').trim().toLowerCase() === 'inactive') return;
    if (String(r[cL] || '').trim().toLowerCase() !== want) return;
    const id = String(r[cId] || '').trim().toUpperCase(); if (id) out[id] = true;
  });
  return out;
}

// ── Web entry points (router) ────────────────────────────────────────
// The list, as the page asks for it: identity proven by PIN.
function svWebList(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const location = String(params.location || '').trim();
    const auth = sv_authSupervisor_(ss, params.supId, params.supPin, location);
    if (!auth.ok) return { success: false, error: auth.error };
    const p = {}; Object.keys(params).forEach(k => { if (k !== 'supPin') p[k] = params[k]; });
    p.viewerId = auth.id;
    const r = getSupervisorVerifyList(p);
    if (r && r.success) r.viewer = { id: auth.id, name: auth.name, isHr: auth.isHr };
    return r;
  } catch (e) { return { success: false, error: e.toString() }; }
}

// The correction table (correctSession / addSession / deleteSession) obeys
// the same three rules. Called by the router BEFORE the real function; on
// success it rewrites body.by from the proven identity.
function svGuardCorrection(body, kind) {
  try {
    body = body || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const empId = String(body.empId || '').trim().toUpperCase();
    let location = String(body.location || '').trim();
    if (!location && empId) {
      const hr = ss.getSheetByName('HR Maya'), hi = sv_headerIndex_(hr);
      const cId = hi.at('Employee ID'); let cL = hi.at('Location'); if (cL === -1) cL = hi.at('Branch');
      hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().some(r => {
        if (String(r[cId] || '').trim().toUpperCase() !== empId) return false;
        location = String(r[cL] || '').trim(); return true;
      });
    }
    const auth = sv_authSupervisor_(ss, body.supId, body.supPin, location);
    if (!auth.ok) return { ok: false, error: auth.error };

    if (kind !== 'deleteSession') {
      const d = sv_norm_(body.date);
      const now = new Date();
      const at = (hm, nextDay) => {
        const base = sv_toDate_(d), m = sv_hmToMin_(hm);
        return (base && m != null) ? new Date(base.getTime() + (nextDay ? 86400000 : 0) + m * 60000) : null;
      };
      const tin = at(body.checkIn, false);
      const nextDay = body.coNextDay === true || String(body.coNextDay) === 'true';
      const tout = at(body.checkOut, nextDay);
      if ((tin && tin > now) || (tout && tout > now))
        return { ok: false, error: 'That time has not happened yet (' +
          Utilities.formatDate((tout && tout > now) ? tout : tin, tz_(), 'dd-MM HH:mm') +
          '). Enter it once he has actually checked out.' };
    }
    body.by = auth.by;           // never the name the page typed
    delete body.supPin;
    return { ok: true };
  } catch (e) { return { ok: false, error: e.toString() }; }
}

// For the page: who may pick themselves from the list as HR.
function svRoles() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const hi = sv_headerIndex_(hr);
    const cId = hi.at('Employee ID'); let cNm = hi.at('Name (English)'); if (cNm === -1) cNm = hi.at('Name');
    const names = {};
    hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues()
      .forEach(r => { names[String(r[cId] || '').trim().toUpperCase()] = String(r[cNm] || '').trim(); });
    return { success: true, hr: SV_HR_IDS.map(id => ({ id: id, name: names[id] || id })) };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Attendance_Log index ─────────────────────────────────────────────
// Sessions keyed by the day they STARTED (overnight shifts belong to the day
// they began). Also records whether a human has already stated this day.
//
// REPEATED CHECK-IN TAPS: a Check In while one is open, less than 16h later,
// is a repeat tap on the same shift (AttendanceCore §0). The session starts
// at the first tap that lines up with a shift — the same answer the 4 AM
// compile gives — so "Worked" closes the right check-in. Before this it
// closed the LAST tap: Ajay Gurung, answered "Worked", still came out 0.31h.
//   startsFor(empId, locations) → shift starts in minutes (optional)
function sv_logIndex_(ss, empIds, startsFor) {
  const idx = {}, human = {};
  const log = ss.getSheetByName('Attendance_Log');
  const meta = { cols: null, sheet: log };
  if (!log || log.getLastRow() < 2) return { sessions: idx, human: human, meta: meta };

  const hi = sv_headerIndex_(log);
  const cId = hi.at('Employee ID'), cDate = hi.at('Date'), cStage = hi.at('Stage'),
        cTs = hi.at('Timestamp'), cAuth = hi.at('Auth');
  meta.cols = { id: cId, name: hi.at('Employee Name'), date: cDate, stage: cStage,
                ts: cTs, auth: cAuth, branch: hi.at('Branch'),
                flag: hi.at('Geo Flag') !== -1 ? hi.at('Geo Flag') : hi.at('Location Flag'),
                width: log.getLastColumn() };
  if (cId === -1 || cStage === -1 || cTs === -1) return { sessions: idx, human: human, meta: meta };

  const data = log.getRange(2, 1, log.getLastRow()-1, log.getLastColumn()).getValues();
  const cBranch = meta.cols.branch;
  const now = new Date();
  const byEmp = {};
  data.forEach((r, i) => {
    const id = String(r[cId] || '').trim(); if (!id || !empIds[id]) return;
    const stage = String(r[cStage] || '').trim();
    const dateCol = cDate >= 0 ? sv_norm_(r[cDate]) : '';
    const auth = cAuth >= 0 ? String(r[cAuth] || '').trim().toLowerCase() : '';
    if (auth && SV_HUMAN_AUTH.indexOf(auth) !== -1 && dateCol) human[id + '|' + dateCol] = true;
    if (stage === 'Day Off') { if (dateCol) idx[id + '|' + dateCol] = { dayOff: true }; return; }
    if (stage !== 'Check In' && stage !== 'Check Out') return;
    const ts = (typeof parseLogTimestamp === 'function') ? parseLogTimestamp(r[cTs]) : null;
    if (!ts || ts > now) return;                       // a future punch has not happened
    (byEmp[id] = byEmp[id] || []).push({ row: i + 2, ts: ts, stage: stage,
      location: cBranch >= 0 ? String(r[cBranch] || '').trim() : '' });
  });

  const repeat = (a, b) => (typeof acIsRepeatCheckIn_ === 'function') ? acIsRepeatCheckIn_(a, b) : false;
  Object.keys(byEmp).forEach(id => {
    const p = byEmp[id].sort((a, b) => a.ts - b.ts);
    const locs = {}; p.forEach(x => { if (x.location) locs[x.location] = true; });
    const starts = startsFor ? startsFor(id, Object.keys(locs)) : [];
    let cur = null;
    const close = (s) => {
      const st = (typeof acPickRunStart_ === 'function') ? (acPickRunStart_(s.run, starts) || s.run[0])
                                                         : s.run[s.run.length - 1];
      s.ciRow = st.row; s.ciTs = st.ts;
      s.taps = s.run.map(x => Utilities.formatDate(x.ts, tz_(), 'HH:mm'));
      const k = id + '|' + sv_dateToStr_(s.ciTs);
      const prev = idx[k];
      // A real session outranks a bare Day Off tap on the same date.
      if (!prev || prev.dayOff) idx[k] = s;
    };
    p.forEach(x => {
      if (x.stage === 'Check In') {
        if (cur && !cur.coRow && repeat(cur.run[0].ts, x.ts)) { cur.run.push(x); return; }
        if (cur) close(cur);
        cur = { run: [x], ciRow: x.row, ciTs: x.ts, coRow: null, coTs: null };
      }
      else if (cur) { cur.coRow = x.row; cur.coTs = x.ts; close(cur); cur = null; }
    });
    if (cur) close(cur);
  });

  return { sessions: idx, human: human, meta: meta };
}

// ── Manual_Attendance_Aug index ──────────────────────────────────────
function sv_manualIndex_(ss, empIds) {
  const out = {};
  const sh = ss.getSheetByName(SV_MANUAL_TAB);
  if (!sh || sh.getLastRow() < 2) return out;
  const hi = sv_headerIndex_(sh);
  const cId = hi.at('Employee ID'), cDate = hi.at('Date'), cSt = hi.at('Status');
  if (cId === -1 || cDate === -1 || cSt === -1) return out;
  sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues().forEach(r => {
    const id = String(r[cId] || '').trim(); if (!id || !empIds[id]) return;
    const d = sv_norm_(r[cDate]); if (!d) return;
    out[id + '|' + d] = String(r[cSt] || '').trim();
  });
  return out;
}

// ============================================================
// BUILD THE LIST
// ============================================================
function getSupervisorVerifyList(params) {
  try {
    params = params || {};
    const location = String(params.location || '').trim();
    if (!location) return { success:false, error:'location required' };
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    const from = params.from ? sv_norm_(params.from) : SV_FROM_DEFAULT;
    // Never today. A day is not over while it is still being worked — see
    // SV_SHIFT_END_GRACE_MIN. Yesterday's night shift is held back below
    // until it has actually finished.
    const yday = sv_yesterday_();
    let to = params.to ? sv_norm_(params.to) : yday;
    if (sv_iso_(to) > sv_iso_(yday)) to = yday;
    const now = new Date();
    const inRange = (d) => d && sv_iso_(d) >= sv_iso_(from) && sv_iso_(d) <= sv_iso_(to);
    const isManualPeriod = (d) => sv_iso_(d) <= sv_iso_(SV_MANUAL_UNTIL);

    // ── staff at this location ────────────────────────────────────
    const hr = ss.getSheetByName('HR Maya');
    const hi = sv_headerIndex_(hr);
    const hId = hi.at('Employee ID'), hSt = hi.at('Status');
    let hNm = hi.at('Name (English)'); if (hNm === -1) hNm = hi.at('Name');
    let hLoc = hi.at('Location'); if (hLoc === -1) hLoc = hi.at('Branch');
    let hSh = hi.at('Shift');
    const mine = {}, empShift = {};
    hr.getRange(2,1,hr.getLastRow()-1,hr.getLastColumn()).getValues().forEach(r => {
      const id = String(r[hId]||'').trim(); if (!id) return;
      if (hSt !== -1 && String(r[hSt]||'').trim().toLowerCase() === 'inactive') return;
      if (hLoc === -1 || String(r[hLoc]||'').trim().toLowerCase() !== location.toLowerCase()) return;
      mine[id] = hNm !== -1 ? String(r[hNm]||'').trim() : id;
      empShift[id] = hSh !== -1 ? String(r[hSh]||'').trim() : '';
    });

    // A supervisor sees and answers his own days too (21-09-2026).
    const ownHidden = false;

    // ── attendance context ────────────────────────────────────────
    const ctx = {};
    const ad = ss.getSheetByName('Attendance_Daily');
    if (ad && ad.getLastRow() > 1) {
      const ai = sv_headerIndex_(ad);
      const aId=ai.at('Employee ID'), aDate=ai.at('Date'), aIn=ai.at('Check In'),
            aOut=ai.at('Check Out'), aPres=ai.at('Presence (hrs)'), aOt=ai.at('OT Hours'),
            aSt=ai.at('Status'), aType=ai.at('Day Type'), aSrc=ai.at('Source'),
            aShift=ai.at('Shift'), aFlags=ai.at('Flags');
      ad.getRange(2,1,ad.getLastRow()-1,ad.getLastColumn()).getValues().forEach(r => {
        const id = String(r[aId]||'').trim(); if (!id || !mine[id]) return;
        const d = sv_norm_(r[aDate]); if (!inRange(d)) return;
        ctx[id + '|' + d] = {
          checkIn: aIn>=0 ? sv_clock_(r[aIn]) : '', checkOut: aOut>=0 ? sv_clock_(r[aOut]) : '',
          presence: aPres>=0 ? (Number(r[aPres])||0) : 0,
          computed: aOt>=0 ? (Number(r[aOt])||0) : 0,
          status: aSt>=0 ? String(r[aSt]||'').trim() : '',
          dayType: aType>=0 ? String(r[aType]||'').trim() : '',
          source: aSrc>=0 ? String(r[aSrc]||'').trim() : '',
          shift: aShift>=0 ? String(r[aShift]||'').trim() : '',
          flags: aFlags>=0 ? String(r[aFlags]||'').trim() : '',
        };
      });
    }

    const startsFor = (typeof acShiftStartsIndex_ === 'function') ? acShiftStartsIndex_(ss) : null;
    const li = sv_logIndex_(ss, mine, startsFor);
    const manual = sv_manualIndex_(ss, mine);

    // OT claims per man and date, with WHEN and WHERE he made them. A claim
    // tapped at the kiosk at 02:35 with "OK (Kaifan)" means he was at Kaifan
    // at 02:35 — the supervisor should see that before he answers the day.
    const claimsAt = {};
    const penAll = ss.getSheetByName('OT_Pending');
    if (penAll && penAll.getLastRow() > 1) {
      penAll.getRange(2, 1, penAll.getLastRow() - 1, Math.min(9, penAll.getLastColumn())).getValues().forEach(r => {
        const id = String(r[1] || '').trim(); if (!mine[id]) return;
        const d = sv_norm_(r[3]); if (!d) return;
        const sub = (typeof parseLogTimestamp === 'function') ? parseLogTimestamp(r[7]) : null;
        (claimsAt[id + '|' + d] = claimsAt[id + '|' + d] || []).push({
          hours: Number(r[4]) || 0,
          at: sub ? Utilities.formatDate(sub, tz_(), 'dd-MM HH:mm') : '',
          onSite: /^OK\b/i.test(String(r[8] || '')),
        });
      });
    }

    // Which shift does an OT claim belong to? A second shift finishes at
    // 03:30, so the man taps OT on a date that has no session of its own.
    // otcSessionIndex_ / otcResolveClaim_ live in OtCompile.gs and sit on
    // AttendanceCore — ONE answer, shared with the OT compiler's console,
    // rather than this file forming a second opinion.
    const otIdx = (typeof otcSessionIndex_ === 'function') ? otcSessionIndex_(ss) : null;
    const shiftMap = sv_shiftMap_(ss);
    const locShifts = sv_locShifts_(ss, location);

    // ── 1. OT claims ──────────────────────────────────────────────
    const otRows = [];
    let otAnswered = 0;
    const pen = ss.getSheetByName('OT_Pending');
    if (pen && pen.getLastRow() > 1) {
      const vc = sv_otCols_(pen);
      const data = pen.getRange(2,1,pen.getLastRow()-1,pen.getLastColumn()).getValues();
      data.forEach((r, i) => {
        if (String(r[6]||'').trim() !== 'Pending') return;      // already decided by HR
        const id = String(r[1]||'').trim(); if (!mine[id]) return;
        const d = sv_norm_(r[3]); if (!inRange(d)) return;

        // ALREADY ANSWERED — the verdict is on the claim itself.
        if (vc.verdict >= 0 && String(r[vc.verdict]||'').trim()) { otAnswered++; return; }

        const cx = ctx[id + '|' + d] || {};
        const sess = li.sessions[id + '|' + d];
        const claimed = Number(r[4])||0;

        // ── IS HE CLAIMING MORE THAN HE WORKED? ────────────────────────
        // The commonest overclaim is not a lie, it is ROUNDING UP. The
        // punches show 5h 17m over the shift and the claim comes in at a
        // flat 6h. Forty-three minutes, every time, on every man, every
        // week. Seen live on 03-09: one claim of 6h against SIX MINUTES
        // of overtime in the punches.
        //
        // The page has always shown the real figure. It never did the
        // subtraction, so nobody saw the gap. Now it does, in words, and
        // the exact punch figure is one tap away.
        //
        // Only where the punches can actually be trusted: real times, a
        // completed session, and not the August period where no punch
        // times exist at all.
        const hasPunchEvidence = !isManualPeriod(d) &&
          SV_TIMELESS_SOURCES.indexOf(cx.source || '') === -1 &&
          !!(cx.checkIn && cx.checkOut) && (cx.presence || 0) > 0;
        const computedOt = Number(cx.computed) || 0;
        const overBy = hasPunchEvidence ? (claimed - computedOt) : 0;
        // A whole-hour claim sitting less than an hour above the punches is
        // rounding, not a dispute about what happened.
        const roundedUp = hasPunchEvidence && overBy > 0.02 && overBy < 1 &&
                          Math.abs(claimed - Math.round(claimed)) < 0.001;
        const overclaim = hasPunchEvidence && overBy > 0.02;

        // What the supervisor most needs to know about this row. Never print
        // "worked 0h" beside two real punch times — that is a broken session,
        // and saying so is the useful thing.
        const resolved = (otIdx && typeof otcResolveClaim_ === 'function')
                           ? otcResolveClaim_(otIdx, id, d,
                               (typeof otc_submittedTs_ === 'function') ? otc_submittedTs_(r[7]) : null) : null;
        const spill = !!(resolved && resolved.spillover && resolved.session && !resolved.session.suspect);
        const belongsTo = spill ? resolved.ownerDate : d;

        let context;
        if (spill) {
          // Never ask a supervisor about an empty day that is really the
          // back half of the night before.
          const oc = ctx[id + '|' + belongsTo] || {};
          context = 'Night shift — started ' + belongsTo + ' at ' + resolved.session.inHm +
                    ' and finished ' + d + ' at ' + resolved.session.outHm + ', ' +
                    sv_hm_(resolved.session.presence) + ' worked' +
                    (oc.computed ? ' (' + sv_hm_(oc.computed) + ' over the shift)' : '') +
                    '. The overtime belongs to that shift.';
        } else if (isManualPeriod(d)) {
          context = 'No punch times exist for this period — you marked Present/Absent only.';
        } else if (cx.source && SV_TIMELESS_SOURCES.indexOf(cx.source) !== -1) {
          context = 'Marked ' + (cx.status || 'Present') + ' with no punch times — ' +
                    'the punches can neither support nor contradict this claim.';
        } else if (cx.checkIn && cx.checkOut && cx.presence > 0) {
          const base = 'Punched ' + cx.checkIn + ' → ' + cx.checkOut +
                       ', worked ' + sv_hm_(cx.presence);
          if (computedOt <= 0 && claimed > 0) {
            context = base + ' — that is within the ' + (cx.shift || 'normal') +
                      ' shift, so NO overtime was earned. A claim of ' + sv_hm_(claimed) +
                      ' would mean he stayed about ' + sv_hm_(cx.presence + claimed) + '.';
          } else if (overclaim) {
            context = base + '. The punches show ' + sv_hm_(computedOt) +
                      ' of overtime — he is claiming ' + sv_hm_(claimed) + ', which is ' +
                      sv_hm_(overBy) + ' MORE than he worked.' +
                      (roundedUp ? ' Overtime is paid on the actual time, not rounded up to the next hour.' : '');
          } else {
            context = base + (computedOt > 0 ? ' (' + sv_hm_(computedOt) + ' over shift)' : '');
          }
        } else if (cx.checkIn && cx.checkOut) {
          context = 'Punched ' + cx.checkIn + ' → ' + cx.checkOut +
                    ' but the session does not add up — the times need fixing below.';
        } else if (cx.checkIn) {
          context = 'Punched in at ' + cx.checkIn + ' and never punched out.';
        } else if (sess && sess.ciRow && !sess.coRow) {
          // The day itself is not answered yet. Settle the attendance first;
          // the OT then has a session to be measured against.
          context = 'Checked in at ' + Utilities.formatDate(sess.ciTs, tz_(), 'HH:mm') +
                    ' and never checked out — answer his attendance for ' + d +
                    ' first (Attendance section), then this claim.';
        } else if (sess && sess.dayOff) {
          context = 'He tapped Day Off on this date — claiming a shift worked on a day off.';
        } else {
          context = 'No punches recorded on this day.';
        }

        otRows.push({
          row: i + 2, empId: id, name: mine[id], date: d,
          claimed: claimed, claimedHm: sv_hm_(claimed),
          shift: cx.shift || empShift[id] || '',
          status: cx.status || '(no record)',
          checkIn: cx.checkIn || '', checkOut: cx.checkOut || '',
          presenceHm: cx.presence ? sv_hm_(cx.presence) : '',
          computedHm: sv_hm_(cx.computed || 0),
          manualPeriod: isManualPeriod(d),
          hasPunchEvidence: hasPunchEvidence,
          computed: computedOt,
          computedOtHm: sv_hm_(computedOt),
          overBy: Math.round(overBy * 100) / 100,
          overByHm: overclaim ? sv_hm_(overBy) : '',
          overclaim: overclaim,
          roundedUp: roundedUp,
          belongsTo: belongsTo,
          spillover: spill,
          shiftTimes: spill ? (resolved.session.inHm + ' → ' + resolved.session.outHm) : '',
          context: context,
        });
      });
    }

    // ── 2. Attendance exceptions (Sept onward only) ───────────────
    // In the manual period there is nothing to dispute: no punch data exists,
    // so the supervisor's own Present/Absent mark IS the record.
    //
    // THE SUPERVISOR IS NOT EXPECTED TO KNOW WHO IS MISSING. This list names
    // them. It asks about every day that is not SETTLED, rather than about a
    // hand-written list of bad statuses — so a status nobody anticipated
    // ('No Roster', or anything added later) surfaces instead of disappearing.
    // A day is settled only when it is a clean, positive record.
    const settled = (cx) => {
      if (cx.status === 'Day Off' || cx.status === 'On Leave') return true;
      if (cx.status !== 'Present') return false;
      if (cx.dayType === 'Half Day' || cx.dayType === 'Check Punches') return false;
      if (SV_TIMELESS_SOURCES.indexOf(cx.source) !== -1) return true;   // never punched, never expected to
      return !!(cx.checkIn && cx.checkOut && cx.presence > 0);
    };

    const attRows = [];
    let attAnswered = 0, stillRunning = 0;

    // Dates this location was actually compiled on, and the window each man
    // appears in. Used below to spot a day with no row at all — a man who
    // fell out of the compile entirely — without inventing days before he
    // joined or after he left.
    const compiledDates = {}, empSpan = {};
    Object.keys(ctx).forEach(k => {
      const parts = k.split('|'), id = parts[0], d = parts[1];
      if (isManualPeriod(d)) return;
      compiledDates[d] = true;
      const s = empSpan[id] || (empSpan[id] = { first: d, last: d });
      if (sv_iso_(d) < sv_iso_(s.first)) s.first = d;
      if (sv_iso_(d) > sv_iso_(s.last))  s.last  = d;
    });

    Object.keys(ctx).forEach(k => {
      const parts = k.split('|'), id = parts[0], d = parts[1];
      if (isManualPeriod(d)) return;
      const cx = ctx[k];

      if (settled(cx)) return;

      const sess = li.sessions[k] || {};

      // ALREADY ANSWERED — a human has stated this day, here or by correcting
      // the punches with the tool below. BUT a human who touched the day and
      // left it OPEN has not answered it: 5 claims in 01-03 Sep sat on
      // check-ins a supervisor corrected with no check-out after them, and
      // were never asked about again. Open means not answered.
      const stillOpen = !!(sess.ciRow && !sess.coRow);
      if ((li.human[k] || manual[k]) && !stillOpen) { attAnswered++; return; }
      const openIn = (sess.ciRow && !sess.coRow) ? Utilities.formatDate(sess.ciTs, tz_(), 'HH:mm') : '';
      const inHm = cx.checkIn || openIn;
      const shift = sv_resolveShift_(shiftMap, locShifts, cx.shift || empShift[id], inHm);

      // Not over yet — his shift on this date is still running.
      if (!sv_dayEnded_(d, shift, now)) { stillRunning++; return; }

      attRows.push({
        empId: id, name: mine[id], date: d,
        shift: (shift && shift.name) || cx.shift || '',
        shiftStart: shift ? shift.start : '', shiftEnd: shift ? shift.end : '',
        knowsShift: !!shift,
        missing: false,
        status: cx.status || 'No record', dayType: cx.dayType,
        checkIn: inHm, checkOut: cx.checkOut,
        presenceHm: cx.presence ? sv_hm_(cx.presence) : '—',
        hasOpenSession: !!(sess.ciRow && !sess.coRow),
        context: sv_attContext_(cx, sess, openIn, claimsAt[k]),
        // What "Worked" will write, so the supervisor is told before he taps.
        // With no punch at all there is nothing to complete: the day is
        // recorded as a stated full day, not as invented punch times.
        willWrite: inHm
          ? (shift ? 'keeps ' + inHm + ' in, sets ' + shift.end + ' out' : '')
          : 'recorded as a full day worked (no punch times invented)',
      });
    });

    // ── 2b. Men with NO ROW AT ALL on a day the rest of the location
    // was compiled. The exception list above can only ask about rows that
    // exist; a man who dropped out of the compile entirely would otherwise
    // be silently unpaid and nobody would be asked about him.
    // Bounded to the days between his first and last appearance, so a new
    // joiner is not queried about the weeks before he started.
    const dateList = Object.keys(compiledDates).sort((a,b) => sv_iso_(a).localeCompare(sv_iso_(b)));
    Object.keys(mine).forEach(id => {
      const span = empSpan[id]; if (!span) return;        // never appears in the period at all
      dateList.forEach(d => {
        if (sv_iso_(d) < sv_iso_(span.first) || sv_iso_(d) > sv_iso_(span.last)) return;
        const k = id + '|' + d;
        if (ctx[k]) return;                                // has a row — handled above
        if (li.human[k] || manual[k]) { attAnswered++; return; }
        const sess = li.sessions[k] || {};
        const shift = sv_resolveShift_(shiftMap, locShifts, empShift[id],
                        sess.ciTs ? Utilities.formatDate(sess.ciTs, tz_(), 'HH:mm') : '');
        if (!sv_dayEnded_(d, shift, now)) { stillRunning++; return; }
        attRows.push({
          empId: id, name: mine[id], date: d,
          shift: (shift && shift.name) || empShift[id] || '',
          shiftStart: shift ? shift.start : '', shiftEnd: shift ? shift.end : '',
          knowsShift: !!shift,
          missing: true,
          status: 'Missing', dayType: '',
          checkIn: '', checkOut: '', presenceHm: '—',
          hasOpenSession: false,
          context: 'He is not on the attendance sheet for this day at all, ' +
                   'though the rest of the location is. Say what he did.',
          // A missing day has no punch by definition, so "Worked" records a
          // stated full day. No shift needs resolving and no times are made up.
          willWrite: 'recorded as a full day worked (no punch times invented)',
        });
      });
    });

    // How much is being claimed above the punches, in total. A supervisor
    // who can see "14 claims, 9h 40m more than the punches show" answers
    // differently from one looking at a list of separate small roundings.
    let overCount = 0, overHours = 0, roundedCount = 0;
    otRows.forEach(r => {
      if (!r.overclaim) return;
      overCount++; overHours += r.overBy;
      if (r.roundedUp) roundedCount++;
    });

    const bydate = (a,b) => (sv_iso_(a.date)+a.name).localeCompare(sv_iso_(b.date)+b.name);
    otRows.sort(bydate); attRows.sort(bydate);

    return { success:true, location:location, from:from, to:to,
             ot:otRows, attendance:attRows,
             otCount:otRows.length, attCount:attRows.length,
             otAnswered:otAnswered, attAnswered:attAnswered,
             overCount: overCount, roundedCount: roundedCount,
             overHours: Math.round(overHours * 100) / 100,
             overHoursHm: sv_hm_(overHours),
             missingCount: attRows.filter(r => r.missing).length,
             staffCount: Object.keys(mine).length,
             daysCovered: dateList.length,
             stillRunning: stillRunning,     // held back: shift not finished yet
             ownHidden: ownHidden };         // his own days are for someone else
  } catch (err) { return { success:false, error: err.toString() }; }
}

// What the supervisor needs to know about an attendance day, in one line.
function sv_attContext_(cx, sess, openIn, claims) {
  let s;
  const taps = (sess && sess.taps && sess.taps.length > 1) ? sess.taps : null;
  if (openIn) {
    s = 'Checked in at ' + openIn + ' and never checked out.';
  } else if (cx.checkIn && !cx.checkOut) {
    s = 'Punched in at ' + cx.checkIn + ' and never punched out.';
  } else if (cx.status === 'Absent') {
    s = 'Marked ABSENT — this deducts a day\'s pay unless you say otherwise.';
  } else if (cx.status === 'Not Recorded') {
    s = 'Nothing recorded for this day at all.';
  } else if (cx.status === 'No Roster') {
    s = 'No roster was published — and he has no punch either.';
  } else if (cx.dayType === 'Half Day') {
    s = 'Worked ' + sv_hm_(cx.presence) + ' — less than the shift.';
  } else if (cx.checkIn || cx.checkOut) {
    s = 'Punches look wrong (' + sv_hm_(cx.presence) + ').';
  } else {
    s = 'Marked "' + (cx.status || 'nothing') + '" with no punches.';
  }
  if (taps) s += ' He tapped Check In ' + taps.length + ' times (' + taps.join(', ') +
                 ') — the day is counted from ' + (openIn || cx.checkIn || taps[0]) + '.';
  // A kiosk OT claim is a timestamp of him standing at the kiosk. If he made
  // one on site late at night, "Worked" at the shift end would cut off the
  // hours he is claiming — say so, so the supervisor enters the real time.
  (claims || []).forEach(c => {
    if (!c.at) return;
    s += ' He claimed ' + sv_hm_(c.hours) + ' OT from the kiosk at ' + c.at +
         (c.onSite ? ' (on site)' : '') + '.';
  });
  if ((claims || []).some(c => c.onSite) && (openIn || cx.checkIn))
    s += ' If he stayed that late, enter his real check-out below instead of tapping Worked.';
  return s;
}

// ============================================================
// RECORD VERDICTS
// verdicts: [{ type:'OT'|'ATT', empId, name, date, claimed, verdict, hours, note }]
//   OT  verdict: 'yes' | 'partly' | 'no'
//   ATT verdict: 'present' | 'dayoff' | 'absent'
// ============================================================
function saveSupervisorVerdicts(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const location = String(params.location || '').trim();

    // ── WHO IS ANSWERING — proven, not picked from a list ──────────
    const auth = sv_authSupervisor_(ss, params.supId, params.supPin, location);
    if (!auth.ok) return { success:false, error: auth.error };
    const by = auth.by;                      // never the text the page sent
    const me = auth.id;

    let list;
    try { list = typeof params.verdicts === 'string' ? JSON.parse(params.verdicts) : params.verdicts; }
    catch (e) { return { success:false, error:'Invalid verdicts payload.' }; }
    if (!list || !list.length) return { success:false, error:'Nothing submitted.' };

    const nowD = new Date();
    const today = formatDdMmYyyy_(nowD);
    const now = Utilities.formatDate(nowD, tz_(), 'dd-MM-yyyy HH:mm:ss');
    const notes = [], skipped = [];
    let otSaved = 0, attSaved = 0;

    // Nothing dated today or later — it has not finished happening.
    list = list.filter(v => {
      const d = sv_norm_(v.date);
      if (d && sv_iso_(d) < sv_iso_(today)) return true;
      skipped.push((v.name || v.empId) + ' ' + (v.date || '') + ': today is not over yet — answer it tomorrow');
      return false;
    });
    // Only men at the location he supervises (HR: anywhere).
    if (!auth.isHr) {
      const staff = sv_staffAt_(ss, location);
      list = list.filter(v => {
        if (staff[String(v.empId || '').trim().toUpperCase()]) return true;
        skipped.push((v.name || v.empId) + ' ' + (v.date || '') + ': not on the staff at ' + location);
        return false;
      });
    }

    // ── OT verdicts: written onto the claim row ───────────────────
    const otList = list.filter(v => String(v.type).toUpperCase() === 'OT');
    if (otList.length) {
      const pen = ss.getSheetByName('OT_Pending');
      if (!pen) return { success:false, error:'OT_Pending not found.' };
      const vc = sv_otCols_(pen);
      const last = pen.getLastRow();

      otList.forEach(v => {
        const r = parseInt(v.row, 10);
        if (!r || r < 2 || r > last) { skipped.push(v.empId + ' ' + v.date + ': row out of range'); return; }
        const cur = pen.getRange(r, 1, 1, vc.width).getValues()[0];
        // Guard against a stale page: the row must still be the same claim.
        if (String(cur[1]||'').trim() !== String(v.empId).trim() ||
            sv_norm_(cur[3]) !== sv_norm_(v.date)) {
          skipped.push(v.empId + ' ' + v.date + ': the list moved on — reload and try again'); return;
        }
        if (String(cur[6]||'').trim() !== 'Pending') {
          skipped.push(v.empId + ' ' + v.date + ': already decided by HR'); return;
        }
        if (String(cur[vc.verdict]||'').trim()) {
          skipped.push(v.empId + ' ' + v.date + ': already answered'); return;
        }

        const claimed = Number(cur[4]) || 0;
        let word, hours;
        if (v.verdict === 'no')          { word = 'Not worked'; hours = 0; }
        else if (v.verdict === 'partly') { word = 'Partly';     hours = Number(v.hours); }
        else                             { word = 'Confirmed';  hours = (v.hours == null || v.hours === '') ? claimed : Number(v.hours); }
        if (isNaN(hours) || hours < 0) { skipped.push(v.empId + ' ' + v.date + ': invalid hours'); return; }
        if (hours > claimed) hours = claimed;     // a supervisor may cut a claim, never raise it

        pen.getRange(r, vc.verdict + 1, 1, 4).setValues([[word, hours, by, now]]);
        if (typeof logAmend_ === 'function')
          logAmend_(ss, by, 'OT VERIFY', v.empId, sv_norm_(v.date), 'OT',
                    'claimed ' + claimed + 'h → supervisor says ' + word +
                    ' (' + hours + 'h)' + (v.note ? ' — ' + v.note : ''));
        otSaved++;
      });
    }

    // ── Attendance verdicts: written where attendance already lives ──
    const attList = list.filter(v => String(v.type).toUpperCase() === 'ATT');
    const touchedDates = {};
    if (attList.length) {
      const empIds = {}; attList.forEach(v => empIds[String(v.empId).trim()] = true);
      const startsFor = (typeof acShiftStartsIndex_ === 'function') ? acShiftStartsIndex_(ss) : null;
      const li = sv_logIndex_(ss, empIds, startsFor);
      // A day already stated in Manual_Attendance_Aug has been answered just
      // as firmly as one answered with a punch. Without this, a stale page
      // could re-post a verdict over the top of one already given.
      const manualIdx = sv_manualIndex_(ss, empIds);
      const log = li.meta.sheet, C = li.meta.cols;
      if (!log || !C || C.id === -1 || C.stage === -1 || C.ts === -1)
        return { success:false, error:'Attendance_Log is missing its Employee ID / Stage / Timestamp columns.' };

      const shiftMap = sv_shiftMap_(ss);
      const locShifts = sv_locShifts_(ss, location);
      const ctxShift = sv_dailyShiftLookup_(ss, empIds);

      const mkTs = (dStr, hm, nextDay) => {
        const base = sv_toDate_(dStr); if (!base) return null;
        const d = new Date(base.getTime() + (nextDay ? 86400000 : 0));
        return Utilities.formatDate(d, tz_(), 'dd/MM/yyyy') + ' ' + hm + ':00';
      };
      // Never write a punch that has not happened yet.
      const isFuture = (dStr, hm, nextDay) => {
        const base = sv_toDate_(dStr); if (!base) return true;
        return base.getTime() + (nextDay ? 86400000 : 0) + (sv_hmToMin_(hm) || 0) * 60000 > nowD.getTime();
      };
      const appendPunch = (empId, name, dStr, stage, tsStr) => {
        const row = new Array(C.width).fill('');
        row[C.id] = empId;
        if (C.name >= 0) row[C.name] = name || '';
        if (C.date >= 0) row[C.date] = dStr;
        row[C.stage] = stage;
        row[C.ts] = tsStr;
        if (C.auth >= 0) row[C.auth] = SV_AUTH;
        if (C.branch >= 0) row[C.branch] = location;
        if (C.flag >= 0) row[C.flag] = 'VERIFIED (supervisor)';
        log.appendRow(row);
        if (C.date >= 0) log.getRange(log.getLastRow(), C.date + 1).setNumberFormat('@');
      };

      attList.forEach(v => {
        const empId = String(v.empId || '').trim();
        const dStr = sv_norm_(v.date);
        const name = String(v.name || '').trim();
        if (!empId || !dStr) { skipped.push((empId||'?') + ': bad row'); return; }
        const k = empId + '|' + dStr;

        if (li.human[k] || manualIdx[k]) { skipped.push(name + ' ' + dStr + ': already answered'); return; }

        const sess = li.sessions[k] || {};

        // Is his shift on this date over? Not before it has ended.
        const shiftNow = sv_resolveShift_(shiftMap, locShifts, v.shift || ctxShift[k],
                           sess.ciTs ? Utilities.formatDate(sess.ciTs, tz_(), 'HH:mm') : '');
        if (!sv_dayEnded_(dStr, shiftNow, nowD)) {
          const end = sv_shiftEndTs_(dStr, shiftNow);
          skipped.push(name + ' ' + dStr + ': his shift has not finished — answer after ' +
            (end ? Utilities.formatDate(new Date(end.getTime() + SV_SHIFT_END_GRACE_MIN * 60000), tz_(), 'dd-MM HH:mm') : 'it ends'));
          return;
        }

        // ── A VERDICT ON A DAY WITH NO PUNCH STAGES IS A STATEMENT ──────
        // Not a punch. Day Off and Absent have no stages by definition, and
        // Present has none when the man never touched the kiosk. On those
        // days the supervisor's word is the whole record, so it is written
        // where supervisor statements live — Manual_Attendance_Aug — which
        // the compile reads at priority 0, above the roster and above
        // everything else.
        //
        // This replaces three different strengths of answer with one:
        //   Day Off used to be written as a kiosk-shaped 'Day Off' punch,
        //   which the compile then re-judged against the roster. If the
        //   roster said Working, the compile OVERTURNED the supervisor and
        //   wrote Absent — docking a day's pay after a supervisor had
        //   explicitly said it was a day off, and telling nobody.
        //
        //   Present with no punches used to be written as an invented
        //   check-in and check-out at the shift times. That fabricated
        //   punch evidence: the invented session made hasPunchEvidence
        //   true, so an OT claim on that day could be flagged as an
        //   over-claim against times a supervisor had made up.
        const statement = (status, verb) => {
          const ok = sv_writeManualDay_(ss, empId, name, location, dStr, status, by);
          if (!ok) { skipped.push(name + ' ' + dStr + ': could not write to ' + SV_MANUAL_TAB); return false; }
          if (typeof logAmend_ === 'function')
            logAmend_(ss, by, verb, empId, dStr, 'attendance', v.note || '');
          // Attendance_Daily is derived, so it is stale the moment this is
          // written. Rebuild the date. (The Absent path used to skip this
          // and leave the sheet wrong until the 4 AM compile.)
          touchedDates[dStr] = true; attSaved++;
          return true;
        };

        if (v.verdict === 'dayoff') {
          statement('Day Off', 'VERIFY DAY OFF');
          if (sess.ciRow && !sess.coRow)
            notes.push(name + ' ' + dStr + ': recorded as Day Off, but he has a check-in ' +
                       'with no check-out on that date — clear it in the correction table below.');
          return;
        }

        if (v.verdict === 'absent') {
          statement('Absent', 'VERIFY ABSENT');
          return;
        }

        // 'present' — he worked.
        // No punch at all on this date: nothing to complete, and nothing to
        // invent. State the day and let the compile pay it at the shift.
        if (!sess.ciRow && !sess.coRow) {
          if (statement('Present', 'VERIFY WORKED'))
            notes.push(name + ' ' + dStr + ': recorded as a full day worked (no punches exist ' +
                       'for this date, so none were invented).');
          return;
        }

        // A real punch DOES exist — there is a stage here, so complete it
        // rather than overwrite it with a statement.
        // v.shift is the shift the page showed him when he tapped Worked, so
        // what gets written is exactly what he was told would be written.
        const shift = sv_resolveShift_(shiftMap, locShifts, v.shift || ctxShift[k], sess.ciTs ?
                        Utilities.formatDate(sess.ciTs, tz_(), 'HH:mm') : '');
        if (!shift) {
          skipped.push(name + ' ' + dStr + ': this location has more than one shift and ' +
                       'nothing says which one — add the session in the table below instead.');
          return;
        }
        const overnight = sv_hmToMin_(shift.end) <= sv_hmToMin_(shift.start);
        const outDay = sess.ciTs ? sv_dateToStr_(sess.ciTs) : dStr;
        if (isFuture(outDay, shift.end, overnight)) {
          skipped.push(name + ' ' + dStr + ': ' + shift.end + ' has not come yet — answer after the shift ends');
          return;
        }
        let replacedOut = '';   // a real check-out time this verdict overwrites — kept in the log

        if (sess.ciRow && !sess.coRow) {
          // Real check-in, missing check-out: add only the check-out.
          const ciDay = sv_dateToStr_(sess.ciTs);   // the check-in row is left alone
          appendPunch(empId, name, ciDay, 'Check Out', mkTs(ciDay, shift.end, overnight));
          notes.push(name + ' ' + dStr + ': check-out set to ' + shift.end + ' (shift end).');
        } else if (sess.ciRow && sess.coRow) {
          // Complete but short. Keep the real check-in; move the check-out to
          // the shift end, which is what "he worked the full shift" means.
          const ciDay = sv_dateToStr_(sess.ciTs);
          const oldOut = Utilities.formatDate(sess.coTs, tz_(), 'dd-MM-yyyy HH:mm');
          replacedOut = oldOut;
          log.getRange(sess.coRow, C.ts + 1).setValue(mkTs(ciDay, shift.end, overnight));
          if (C.auth >= 0) log.getRange(sess.coRow, C.auth + 1).setValue(SV_AUTH);
          notes.push(name + ' ' + dStr + ': check-out moved from ' + oldOut + ' to ' + shift.end + '.');
        } else {
          // Check-out with no check-in — a real stage exists but it is the
          // wrong end. Add the missing check-in at the shift start; the
          // man's own punch is left exactly as he made it.
          appendPunch(empId, name, dStr, 'Check In', mkTs(dStr, shift.start, false));
          replacedOut = Utilities.formatDate(sess.coTs, tz_(), 'dd-MM-yyyy HH:mm');
          // His own check-out may sit before the shift start (an orphan from
          // the night before). Moving it to the shift end is the same thing
          // the complete-but-short branch above does, and it keeps the
          // session from failing to close.
          log.getRange(sess.coRow, C.ts + 1).setValue(mkTs(dStr, shift.end, overnight));
          if (C.auth >= 0) log.getRange(sess.coRow, C.auth + 1).setValue(SV_AUTH);
          notes.push(name + ' ' + dStr + ': check-in added at ' + shift.start +
                     ', check-out set to ' + shift.end + '.');
        }

        if (typeof logAmend_ === 'function')
          logAmend_(ss, by, 'VERIFY WORKED', empId, dStr, 'attendance',
                    shift.name + ' ' + shift.start + '–' + shift.end +
                    (replacedOut ? ' (his own check-out was ' + replacedOut + ')' : '') +
                    (v.note ? ' — ' + v.note : ''));
        touchedDates[dStr] = true; attSaved++;
      });
    }

    // ── Refresh Attendance_Daily for the dates we changed ─────────
    // Attendance_Log is the source of truth; Attendance_Daily is derived.
    // Best effort with a time budget — whatever does not fit is picked up by
    // the 4 AM compile, and the supervisor is told which.
    const dates = Object.keys(touchedDates).sort((a,b) => sv_iso_(a).localeCompare(sv_iso_(b)));
    const rebuilt = [], deferred = [];
    const t0 = Date.now();
    dates.forEach(d => {
      if (Date.now() - t0 > SV_RECOMPILE_BUDGET_MS) { deferred.push(d); return; }
      try {
        if (typeof writeDayToAttendanceDaily_ === 'function') { writeDayToAttendanceDaily_(d); rebuilt.push(d); }
        else deferred.push(d);
      } catch (e) { deferred.push(d); }
    });

    const saved = otSaved + attSaved;
    let message = saved + ' answer(s) recorded';
    if (otSaved) message += ' — ' + otSaved + ' OT claim(s) marked for the OT compiler';
    if (attSaved) message += ' — ' + attSaved + ' attendance day(s) written';
    message += '.';
    if (rebuilt.length) message += ' Attendance rebuilt for ' + rebuilt.join(', ') + '.';
    if (deferred.length) message += ' ' + deferred.join(', ') + ' will refresh on tonight\'s run.';
    if (skipped.length) message += ' ' + skipped.length + ' skipped.';

    return { success:true, saved:saved, ot:otSaved, attendance:attSaved,
             rebuilt:rebuilt, deferred:deferred, notes:notes, skipped:skipped,
             message:message };
  } catch (err) { return { success:false, error: err.toString() }; }
}

// Shift recorded in Attendance_Daily, keyed emp|date — used when deciding
// which shift a "Worked" verdict should be written against.
function sv_dailyShiftLookup_(ss, empIds) {
  const out = {};
  const ad = ss.getSheetByName('Attendance_Daily');
  if (!ad || ad.getLastRow() < 2) return out;
  const ai = sv_headerIndex_(ad);
  const aId = ai.at('Employee ID'), aDate = ai.at('Date'), aShift = ai.at('Shift');
  if (aId === -1 || aDate === -1 || aShift === -1) return out;
  ad.getRange(2,1,ad.getLastRow()-1,ad.getLastColumn()).getValues().forEach(r => {
    const id = String(r[aId]||'').trim(); if (!id || !empIds[id]) return;
    out[id + '|' + sv_norm_(r[aDate])] = String(r[aShift]||'').trim();
  });
  return out;
}

// Write (or update) one day in Manual_Attendance_Aug — the tab supervisors
// already use to state a day, and the one the compile reads first.
// ============================================================
// UNDO VERDICTS GIVEN BEFORE THE DAY WAS OVER  (21-09-2026)
// ------------------------------------------------------------
// Run from the editor:
//   svPreviewUndoEarly()   — says exactly what would change. Writes nothing.
//   svCommitUndoEarly()    — does it, then rebuilds Attendance_Daily.
//
// A verdict is EARLY if it was given before that man's shift on that date
// had ended. Whatever an early verdict wrote is removed:
//   • check-ins / check-outs it APPENDED to Attendance_Log
//     (Geo Flag "VERIFIED (supervisor)" or "MANUAL (supervisor)")
//   • Present / Absent / Day Off statements in Manual_Attendance_Aug
// The man's own kiosk punches are never touched. The day then comes back
// onto the supervisor's list, to be answered now that it is over.
//
// (Until 21-09-2026 it also cleared OT verdicts a supervisor gave on his
// own claim. Supervisors may now answer for themselves — it no longer does,
// and svCommitRestoreSelfOt() puts back any it cleared.)
//
// August (≤ 31-08) is never touched: there, statements are the record.
// ============================================================
function svPreviewUndoEarly() { return sv_undoEarly_(false); }
function svCommitUndoEarly()  { return sv_undoEarly_(true); }

// asData: return the findings instead of a report (used by checkDay()).
function sv_undoEarly_(commit, asData) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const lock = LockService.getScriptLock();
  if (commit && !lock.tryLock(30000)) { Logger.log('Busy — try again in a minute.'); return 'busy'; }
  try {
    const out = [];
    const shiftMap = sv_shiftMap_(ss);
    const parseTs = (v) => (typeof parseLogTimestamp === 'function') ? parseLogTimestamp(v)
                          : (v instanceof Date ? v : null);
    const fmt = (d) => d ? Utilities.formatDate(d, tz_(), 'dd-MM HH:mm') : '?';

    // His shift from HR Maya, for when the verdict itself does not say.
    const hr = ss.getSheetByName('HR Maya'), hhi = sv_headerIndex_(hr);
    const hId = hhi.at('Employee ID'), hSh = hhi.at('Shift');
    let hNm = hhi.at('Name (English)'); if (hNm === -1) hNm = hhi.at('Name');
    let hLoc = hhi.at('Location'); if (hLoc === -1) hLoc = hhi.at('Branch');
    const emp = {};
    hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().forEach(r => {
      const id = String(r[hId] || '').trim().toUpperCase(); if (!id) return;
      emp[id] = { name: hNm >= 0 ? String(r[hNm] || '').trim() : id,
                  loc: hLoc >= 0 ? String(r[hLoc] || '').trim() : '',
                  shift: hSh >= 0 ? shiftMap[String(r[hSh] || '').trim().toLowerCase()] || null : null };
    });
    const endFor = (id, d, detail) => {
      // "Kaifan First Shift 11:30–20:30" — the shift the verdict itself used.
      const m = String(detail || '').match(/(\d{1,2}:\d{2})\s*[–-]\s*(\d{1,2}:\d{2})/);
      const sh = m ? { start: m[1], end: m[2] } : (emp[id] && emp[id].shift);
      return sv_shiftEndTs_(d, sh);
    };
    const afterAug = (d) => d && sv_iso_(d) > sv_iso_(SV_MANUAL_UNTIL);

    // ── 1. Which days were answered before they were over? ───────
    // Attendance_Amendments rows: When | By | Verb | Employee | Date | Stage | Detail
    const early = {};            // emp|date -> { when, by, verb, end }
    const selfAtt = [];
    const am = ss.getSheetByName('Attendance_Amendments');
    const VERDICT_VERBS = ['VERIFY WORKED', 'VERIFY DAY OFF', 'VERIFY ABSENT', 'ADD'];
    if (am && am.getLastRow() > 1) {
      am.getRange(2, 1, am.getLastRow() - 1, Math.max(7, am.getLastColumn())).getValues().forEach(r => {
        const verb = String(r[2] || '').trim().toUpperCase();
        const id = String(r[3] || '').trim().toUpperCase();
        const d = sv_norm_(r[4]);
        const when = parseTs(r[0]);
        if (!id || !afterAug(d) || !when) return;
        const byId = sv_idFromBy_(r[1]);
        if (VERDICT_VERBS.indexOf(verb) === -1) return;
        const end = endFor(id, d, r[6]);
        if (!end || when.getTime() >= end.getTime()) return;
        const k = id + '|' + d;
        if (!early[k] || when < early[k].when) early[k] = { when: when, by: String(r[1]), verb: verb, end: end };
      });
    }

    // ── 2. What those verdicts wrote into Attendance_Log ─────────
    const log = ss.getSheetByName('Attendance_Log');
    const lhi = sv_headerIndex_(log);
    const lId = lhi.at('Employee ID'), lDate = lhi.at('Date'), lStage = lhi.at('Stage'),
          lTs = lhi.at('Timestamp'), lAuth = lhi.at('Auth');
    const lFlag = lhi.at('Geo Flag') !== -1 ? lhi.at('Geo Flag') : lhi.at('Location Flag');
    const logDel = [], cannot = [];
    const logData = log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn()).getValues();
    logData.forEach((r, i) => {
      const id = String(r[lId] || '').trim().toUpperCase();
      const d = sv_norm_(r[lDate]);
      const k = id + '|' + d;
      if (!early[k]) return;
      const flag = lFlag >= 0 ? String(r[lFlag] || '').trim() : '';
      const auth = lAuth >= 0 ? String(r[lAuth] || '').trim().toLowerCase() : '';
      const written = /^(VERIFIED|MANUAL) \(supervisor\)$/i.test(flag);
      if (written) {
        logDel.push({ row: i + 2, id: id, d: d, stage: String(r[lStage] || ''), ts: parseTs(r[lTs]) });
      } else if (auth === 'supervisor verified' || auth === 'supervisor added') {
        // His own punch, whose time a verdict overwrote. The old time was not
        // kept before 21-09, so it cannot be put back automatically.
        cannot.push(k + '  ' + String(r[lStage]) + ' now ' + fmt(parseTs(r[lTs])) +
                    ' — overwritten by the verdict; his own time is lost. Correct it by hand.');
      }
    });

    // ── 3. Statements in Manual_Attendance_Aug made too early ────
    const man = ss.getSheetByName(SV_MANUAL_TAB);
    const manDel = [];
    if (man && man.getLastRow() > 1) {
      const mhi = sv_headerIndex_(man);
      const mId = mhi.at('Employee ID'), mDate = mhi.at('Date'), mSt = mhi.at('Status'),
            mTs = mhi.at('Timestamp'), mBy = mhi.at('Submitted By');
      man.getRange(2, 1, man.getLastRow() - 1, man.getLastColumn()).getValues().forEach((r, i) => {
        const id = String(r[mId] || '').trim().toUpperCase();
        const d = sv_norm_(r[mDate]);
        if (!id || !afterAug(d)) return;
        const when = mTs >= 0 ? parseTs(r[mTs]) : null;
        if (!when) return;
        const k = id + '|' + d;
        const end = early[k] ? early[k].end : endFor(id, d, '');
        if (!end || when.getTime() >= end.getTime()) return;
        manDel.push({ row: i + 2, id: id, d: d, status: String(r[mSt] || ''), when: when,
                      by: mBy >= 0 ? String(r[mBy] || '') : '' });
        if (!early[k]) early[k] = { when: when, by: mBy >= 0 ? String(r[mBy]) : '', verb: 'STATEMENT', end: end };
      });
    }

    // ── 4. (retired 21-09-2026) OT verdicts on his own claim stay —
    //       a supervisor may answer for himself. See svPreviewRestoreSelfOt().
    const otClear = [];

    if (asData) return { early: early, logDel: logDel, manDel: manDel, otClear: otClear, selfAtt: selfAtt };

    // ── Report ───────────────────────────────────────────────────
    const keys = Object.keys(early).sort((a, b) => {
      const [ia, da] = a.split('|'), [ib, db] = b.split('|');
      return (sv_iso_(da) + ia).localeCompare(sv_iso_(db) + ib);
    });
    out.push((commit ? 'UNDONE' : 'PREVIEW — nothing written yet') + ' — verdicts given before the day was over');
    out.push('');
    out.push(keys.length + ' man-day(s) answered early:');
    keys.forEach(k => {
      const [id, d] = k.split('|'), e = early[k];
      const w = logDel.filter(x => x.id === id && x.d === d).map(x => x.stage + ' ' + fmt(x.ts));
      const m = manDel.filter(x => x.id === id && x.d === d).map(x => 'stated ' + x.status);
      out.push('  ' + d + '  ' + sv_pad2_(id, 9) + sv_pad2_((emp[id] || {}).name || '', 20) +
               'answered ' + fmt(e.when) + ' by ' + e.by + ', shift ended ' + fmt(e.end) +
               (w.length || m.length ? '  → remove: ' + w.concat(m).join(', ') : '  → nothing left to remove'));
    });
    out.push('');
    out.push('Attendance_Log rows to remove : ' + logDel.length);
    out.push('Statements to remove          : ' + manDel.length);
    if (cannot.length) {
      out.push('');
      out.push('CANNOT UNDO AUTOMATICALLY (' + cannot.length + '):');
      cannot.forEach(s => out.push('   ' + s));
    }

    if (!commit) {
      out.push('');
      out.push('Nothing has been changed. If this is right, run svCommitUndoEarly().');
      return sv_report_(out);
    }

    // ── Commit ───────────────────────────────────────────────────
    const who = 'System (undo early verdicts)';
    logDel.sort((a, b) => b.row - a.row).forEach(x => log.deleteRow(x.row));
    manDel.sort((a, b) => b.row - a.row).forEach(x => man.deleteRow(x.row));
    if (typeof logAmend_ === 'function') {
      logDel.forEach(x => logAmend_(ss, who, 'UNDO EARLY', x.id, x.d, x.stage,
        'removed ' + x.stage + ' ' + fmt(x.ts) + ' written before the shift ended'));
      manDel.forEach(x => logAmend_(ss, who, 'UNDO EARLY', x.id, x.d, 'attendance',
        'removed statement "' + x.status + '" made ' + fmt(x.when) + ' by ' + x.by));
    }
    SpreadsheetApp.flush();

    const dates = {};
    keys.forEach(k => { dates[k.split('|')[1]] = true; });
    const dl = Object.keys(dates).sort((a, b) => sv_iso_(a).localeCompare(sv_iso_(b)));
    if (dl.length && typeof adTouchDates_ === 'function') {
      const r = adTouchDates_(dl);
      out.push('');
      out.push('Attendance_Daily rebuilt for ' + dl.join(', ') +
               (r && r.skipped && r.skipped.length ? ' — NOT finished: ' + r.skipped.join(', ') +
                ' (the 4 AM compile will finish them)' : ''));
    }
    out.push('');
    out.push('Done. Those days are back on each supervisor\'s list.');
    return sv_report_(out);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
// ============================================================
// PUT BACK SUPERVISORS' ANSWERS ON THEIR OWN OT CLAIMS  (21-09-2026)
// ------------------------------------------------------------
// svCommitUndoEarly() used to clear the verdict a supervisor gave on his
// own OT claim. That rule is gone: a supervisor may answer for himself.
// This writes those verdicts back, exactly as they were given.
//
// Source 1: the verdicts as they stood in OT_Pending at 13:15 on 21-09
//           (34 of them — below, word for word).
// Source 2: Attendance_Amendments "cleared ... on his own claim" lines,
//           for any cleared after that. Those lines kept the verdict and
//           who gave it, not the hours: a "Confirmed" gets the hours he
//           claimed; anything else is listed for him to answer again.
//
// Only a claim that is still Pending AND has an empty verdict is written.
// Anything answered since is left alone. Safe to run more than once.
// ============================================================
const SV_SELF_OT_AS_GIVEN = [
  // [claim ID, verdict, hours confirmed, verified by, verified on]
  ['OTP-20260902-142343', 'Confirmed', 1, 'Dipendra Thapa (MT-00017)', '20-09-2026 15:20:43'],
  ['OTP-20260903-195822', 'Confirmed', 5, 'Laxman Bahadur Airi (MT-00032)', '20-09-2026 14:51:45'],
  ['OTP-20260904-023309', 'Confirmed', 6, 'Nilam Tamang (MT-00039)', '20-09-2026 17:33:22'],
  ['OTP-20260905-023115', 'Confirmed', 6, 'Dipendra Thapa (MT-00017)', '20-09-2026 16:45:23'],
  ['OTP-20260905-033211', 'Confirmed', 6, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260905-092724', 'Confirmed', 13, 'Laxman Bahadur Airi (MT-00032)', '20-09-2026 14:51:45'],
  ['OTP-20260906-033130', 'Confirmed', 6, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260906-033203', 'Confirmed', 6, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260906-033204', 'Confirmed', 6, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260906-033215', 'Confirmed', 6, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260907-023249', 'Confirmed', 6, 'Nilam Tamang (MT-00039)', '20-09-2026 17:33:22'],
  ['OTP-20260908-023343', 'Confirmed', 6, 'Dipendra Thapa (MT-00017)', '20-09-2026 16:45:23'],
  ['OTP-20260909-032603', 'Partly', 1, 'Dipendra Thapa (MT-00017)', '20-09-2026 16:45:23'],
  ['OTP-20260909-032925', 'Confirmed', 7, 'Nilam Tamang (MT-00039)', '20-09-2026 17:33:22'],
  ['OTP-20260909-032943', 'Confirmed', 1, 'Dipendra Thapa (MT-00017)', '20-09-2026 16:45:23'],
  ['OTP-20260909-041443', 'Confirmed', 7, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260909-043009', 'Confirmed', 7, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260910-000740', 'Confirmed', 2, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260911-023453', 'Confirmed', 6, 'Nilam Tamang (MT-00039)', '20-09-2026 17:33:22'],
  ['OTP-20260911-025307', 'Confirmed', 4, 'Suraj Kumal (MT-00056)', '20-09-2026 17:11:05'],
  ['OTP-20260911-033755', 'Confirmed', 6, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260911-035512', 'Confirmed', 6, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260912-003957', 'Confirmed', 2, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260912-023135', 'Confirmed', 6, 'Dipendra Thapa (MT-00017)', '20-09-2026 16:45:23'],
  ['OTP-20260913-041921', 'Confirmed', 7, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260913-233220', 'Confirmed', 2, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260914-023343', 'Confirmed', 6, 'Nilam Tamang (MT-00039)', '20-09-2026 17:33:22'],
  ['OTP-20260916-033021', 'Confirmed', 7, 'Nilam Tamang (MT-00039)', '20-09-2026 17:33:22'],
  ['OTP-20260916-040512', 'Confirmed', 7, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260916-045008', 'Confirmed', 7, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260917-034248', 'Confirmed', 6, 'Bibek Nepali (MT-00009)', '20-09-2026 18:13:31'],
  ['OTP-20260918-023148', 'Confirmed', 6, 'Nilam Tamang (MT-00039)', '20-09-2026 17:33:22'],
  ['OTP-20260918-033143', 'Confirmed', 6, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
  ['OTP-20260919-000520', 'Confirmed', 2, 'Keshab Alemagar (MT-00028)', '20-09-2026 18:43:44'],
];

function svPreviewRestoreSelfOt() { return sv_restoreSelfOt_(false); }
function svCommitRestoreSelfOt()  { return sv_restoreSelfOt_(true); }

function sv_restoreSelfOt_(commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const lock = LockService.getScriptLock();
  if (commit && !lock.tryLock(30000)) { Logger.log('Busy — try again in a minute.'); return 'busy'; }
  try {
    const out = [];
    const pen = ss.getSheetByName('OT_Pending');
    if (!pen || pen.getLastRow() < 2) return sv_report_(['OT_Pending is empty — nothing to restore.']);
    const vc = sv_otCols_(pen);
    const rows = pen.getRange(2, 1, pen.getLastRow() - 1, vc.width).getValues();
    const byId = {}, byEmpDate = {};
    rows.forEach((r, i) => {
      const o = { row: i + 2, r: r };
      byId[String(r[0] || '').trim()] = o;
      const k = String(r[1] || '').trim().toUpperCase() + '|' + sv_norm_(r[3]);
      (byEmpDate[k] = byEmpDate[k] || []).push(o);
    });
    const open = (o) => o && String(o.r[6] || '').trim() === 'Pending' && !String(o.r[vc.verdict] || '').trim();

    const writes = [], already = [], again = [];
    const seen = {};
    SV_SELF_OT_AS_GIVEN.forEach(g => {
      const o = byId[g[0]];
      if (!o) { already.push(g[0] + ' — claim no longer in OT_Pending'); return; }
      seen[o.row] = true;
      if (!open(o)) { already.push(g[0] + '  ' + o.r[1] + ' ' + sv_norm_(o.r[3]) + ' — ' +
        (String(o.r[6]).trim() !== 'Pending' ? 'already decided (' + o.r[6] + ')' : 'has a verdict: ' + o.r[vc.verdict])); return; }
      writes.push({ row: o.row, id: String(o.r[1]), d: sv_norm_(o.r[3]), claimed: o.r[4],
                    vals: [g[1], g[2], g[3], g[4]], from: 'as given' });
    });

    // Amendment lines for anything cleared after the 13:15 snapshot.
    const am = ss.getSheetByName('Attendance_Amendments');
    if (am && am.getLastRow() > 1) {
      am.getRange(2, 1, am.getLastRow() - 1, am.getLastColumn()).getValues().forEach(r => {
        const txt = r.map(c => String(c == null ? '' : c)).join(' | ');
        const m = txt.match(/cleared "([^"]*)" given by (.+?) on his own claim/);
        if (!m) return;
        const idm = txt.match(/MT-\d+/i);
        const d = r.map(c => sv_norm_(c)).find(v => /^\d{2}-\d{2}-\d{4}$/.test(v) && sv_iso_(v) > sv_iso_(SV_MANUAL_UNTIL))
               || r.map(c => sv_norm_(c)).find(v => /^\d{2}-\d{2}-\d{4}$/.test(v));
        if (!idm || !d) return;
        (byEmpDate[idm[0].toUpperCase() + '|' + d] || []).forEach(o => {
          if (seen[o.row] || !open(o)) return;
          const giver = sv_idFromBy_(m[2]);
          if (giver !== String(o.r[1]).trim().toUpperCase()) return;
          seen[o.row] = true;
          if (/^confirmed$/i.test(m[1].trim())) {
            writes.push({ row: o.row, id: String(o.r[1]), d: d, claimed: o.r[4],
                          vals: [m[1].trim(), o.r[4], m[2].trim(),
                                  r[0] instanceof Date ? Utilities.formatDate(r[0], tz_(), 'dd-MM-yyyy HH:mm:ss') : String(r[0] || '')],
                          from: 'amendment log' });
          } else {
            again.push(d + '  ' + o.r[1] + '  ' + o.r[2] + ' — "' + m[1] + '" was cleared and its hours were not kept; ' +
                       'he answers it again on the supervisor page');
          }
        });
      });
    }

    out.push((commit ? 'RESTORED' : 'PREVIEW — nothing written yet') + ' — supervisors\' answers on their own OT claims');
    out.push('');
    out.push('To put back (' + writes.length + '):');
    writes.forEach(w => out.push('   ' + w.d + '  ' + sv_pad2_(w.id, 9) + sv_pad2_(w.claimed + 'h claimed', 13) +
                                 '"' + w.vals[0] + '" ' + w.vals[1] + 'h  by ' + w.vals[2] + '  (' + w.from + ')'));
    if (again.length) { out.push(''); out.push('TO ANSWER AGAIN (' + again.length + '):'); again.forEach(x => out.push('   ' + x)); }
    out.push('');
    out.push('Left alone (' + already.length + ') — never cleared, or answered/decided since.');

    if (!commit) {
      out.push('');
      out.push(writes.length ? 'Nothing has been changed. If this is right, run svCommitRestoreSelfOt().'
                             : 'Nothing to restore — the verdicts were never cleared.');
      return sv_report_(out);
    }
    writes.forEach(w => pen.getRange(w.row, vc.verdict + 1, 1, 4).setValues([w.vals]));
    if (typeof logAmend_ === 'function')
      writes.forEach(w => logAmend_(ss, 'System (restore own-claim verdicts)', 'RESTORE SELF OT VERIFY', w.id, w.d, 'OT',
        'restored "' + w.vals[0] + '" ' + w.vals[1] + 'h given by ' + w.vals[2]));
    SpreadsheetApp.flush();
    out.push('');
    out.push('Done.');
    return sv_report_(out);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sv_pad2_(s, n) { s = String(s); return s.length >= n ? s.slice(0, n - 1) + ' ' : s + ' '.repeat(n - s.length); }
function sv_report_(lines) { const t = lines.join('\n'); Logger.log(t); return t; }

function sv_writeManualDay_(ss, empId, name, location, dStr, status, by) {
  let sh = ss.getSheetByName(SV_MANUAL_TAB);
  if (!sh) {
    sh = ss.insertSheet(SV_MANUAL_TAB);
    sh.appendRow(['Employee ID','Name','Location','Shift','Date','Status','Submitted By','Timestamp']);
    sh.getRange(1,1,1,8).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const hi = sv_headerIndex_(sh);
  const cId = hi.at('Employee ID'), cDate = hi.at('Date'), cSt = hi.at('Status');
  if (cId === -1 || cDate === -1 || cSt === -1) return false;
  const cNm = hi.at('Name'), cLoc = hi.at('Location'),
        cBy = hi.at('Submitted By'), cTs = hi.at('Timestamp');
  const width = sh.getLastColumn();
  const now = Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm:ss');

  let target = -1;
  if (sh.getLastRow() > 1) {
    const data = sh.getRange(2,1,sh.getLastRow()-1,width).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][cId]||'').trim() === empId && sv_norm_(data[i][cDate]) === dStr) { target = i + 2; break; }
    }
  }
  const row = new Array(width).fill('');
  row[cId] = empId; row[cDate] = dStr; row[cSt] = status;
  if (cNm >= 0) row[cNm] = name || '';
  if (cLoc >= 0) row[cLoc] = location || '';
  if (cBy >= 0) row[cBy] = by || '';
  if (cTs >= 0) row[cTs] = now;

  if (target === -1) { sh.appendRow(row); target = sh.getLastRow(); }
  else sh.getRange(target, 1, 1, width).setValues([row]);
  sh.getRange(target, cDate + 1).setNumberFormat('@');
  return true;
}

// ============================================================
// Run once after deploying. Reports what this file can see and do,
// without writing anything.
// ============================================================
function svSelfCheck() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const out = [];
  const ok = (b) => b ? 'OK  ' : 'FAIL';

  const pen = ss.getSheetByName('OT_Pending');
  out.push(ok(!!pen) + ' OT_Pending present');
  if (pen) {
    const vc = sv_otCols_(pen);
    out.push(ok(vc.verdict >= 0 && vc.hours >= 0 && vc.by >= 0 && vc.on >= 0) +
             ' verdict columns at ' + (vc.verdict+1) + '-' + (vc.on+1));
  }
  const log = ss.getSheetByName('Attendance_Log');
  out.push(ok(!!log) + ' Attendance_Log present');
  if (log) {
    const hi = sv_headerIndex_(log);
    ['Employee ID','Date','Stage','Timestamp','Auth'].forEach(h =>
      out.push(ok(hi.at(h) >= 0) + ' Attendance_Log."' + h + '"'));
  }
  const shifts = sv_shiftMap_(ss);
  out.push(ok(Object.keys(shifts).length > 0) + ' Shifts readable (' +
           Object.keys(shifts).length + ' shifts: ' +
           Object.keys(shifts).map(k => shifts[k].name + ' ' + shifts[k].start + '-' + shifts[k].end).join(', ') + ')');
  out.push(ok(typeof writeDayToAttendanceDaily_ === 'function') + ' writeDayToAttendanceDaily_ callable');
  out.push(ok(typeof parseLogTimestamp === 'function') + ' parseLogTimestamp callable');
  out.push(ok(typeof logAmend_ === 'function') + ' logAmend_ callable');
  out.push(ok(!!ss.getSheetByName(SV_MANUAL_TAB)) + ' ' + SV_MANUAL_TAB + ' present');
  out.push('');
  out.push('hm carry: 7.996h -> ' + sv_hm_(7.996) + '   (must be 8h, not "7h 60m")');
  out.push('hm carry: 0.9999h -> ' + sv_hm_(0.9999) + '  (must be 1h)');
  out.push('hm:       7.5h   -> ' + sv_hm_(7.5));

  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}