// ============================================================
// BackfillCheckouts.gs — bulk-fix dangling Check Ins across
// Attendance_Log by synthesizing a missing Check Out.
// ============================================================
// PROBLEM: many Check Ins across history never got a matching
// Check Out (kiosk not tapped, network drop, forgotten, etc.).
//
// FIX APPROACH (confirmed):
//   - The existing Check In is NEVER touched — it's real, recorded
//     data, not a gap. Only the missing Check Out is added.
//   - Check Out = the employee's assigned shift's actual defined
//     END TIME from the Shifts tab (not a duration computed from
//     whenever they happened to check in — an early or late
//     check-in should not silently become "on time" by construction,
//     and it isn't, since Check In is never rewritten).
//   - Overnight shifts (End time earlier than Start time) roll the
//     checkout to the next calendar day automatically.
//   - Any Pending OT claim for the same date is SHOWN in the preview
//     for context, but is NOT added to the checkout. Adding it would
//     double-pay: once via inflated presence feeding the OT draft, and
//     again when the claim itself is approved into OT_Entries.
//     The claim keeps its own approval path — the only route to payroll.
//   - No shift on record / shift has no End time → falls back to
//     Check In + 9h duration, clearly flagged as a fallback so it's
//     visible which rows used which method.
//
// SAFETY:
//   - Preview → Commit. Nothing is ever written by preview.
//   - Commit re-detects fresh (never trusts a stale client list).
//   - Any session still legitimately IN PROGRESS (final computed
//     checkout, including any OT, would land in the future) is
//     automatically excluded — shown separately, never touched.
//   - Detection always scans the FULL Attendance_Log regardless of
//     the requested date filter, so a real Check Out sitting just
//     outside the filter window can never cause a false positive.
//     The date filter only controls which results are shown/acted on.
//   - Every synthesized row is tagged Auth='Backfilled' with a Geo
//     Flag describing exactly which basis was used — never looks
//     like a real punch, fully audited to Attendance_Amendments.
//   - After commit, Attendance_Daily is automatically recompiled
//     for every date actually touched.
//
// ENDPOINTS: previewMissingCheckouts, commitMissingCheckouts
// ============================================================
// VERSION MARKER — bump this whenever this file changes.
// previewMissingCheckoutsLog() prints it, so there is never any doubt
// about WHICH copy of this file actually ran. Tier 0 appearing as zero
// while the code clearly implements it means an older copy is live.
const BACKFILL_VERSION = 'v6-exclude-exempt-locations';

// ── DO NOT TOUCH ANYTHING ON OR BEFORE THIS DATE ────────────────────
// Supervisors manually reviewed and corrected every punch through
// 03-09-2026 after the kiosk failure on 02-03 Sept. Those rows are the
// authoritative record — a human looked at each one. Synthesising a
// checkout over them would overwrite real corrections with a guess.
//
// Anything dangling on or before this date has already been decided:
// either it was fixed, or it was deliberately left as-is.
const BACKFILL_SKIP_ON_OR_BEFORE = '03-09-2026';

// A session longer than this is not a real shift — it is a phantom
// check-in (the 02:35 kiosk-failure punches) paired with a genuine
// later checkout. Refuse to create one.
const BACKFILL_MAX_SESSION_HRS = 16;

// Punch Auth values that mean "a human already decided this".
// Never synthesise a checkout for a check-in carrying one of these.
const BACKFILL_HUMAN_AUTH = ['Supervisor corrected', 'Supervisor added', 'Manual', 'Backfilled'];

// ── Locations the backfill must never touch ─────────────────────────
// These are auto-marked Present by writeDayToAttendanceDaily_ and their
// staff are not expected to punch:
//   Driver Location — Driver_Default marks them present every day
//   Head Office     — Auto_HeadOffice marks them present on working days
// They also have no shift definition, so any checkout here falls to the
// weakest 9h fallback and invents data for a day that is already
// correctly recorded. Same exemption the roster planner uses.
const BACKFILL_EXEMPT_LOCATIONS = ['Driver Location', 'Head Office'];
// ============================================================

// ── Build employee lookup: id -> { name, location, shift } ──────────
function bc_buildEmpMap_(ss) {
  const hr = ss.getSheetByName('HR Maya');
  const hd = hr.getDataRange().getValues();
  const H = hd[0];
  const cId   = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
  const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
  let cLoc    = H.findIndex(h => String(h).trim().toLowerCase() === 'location');
  if (cLoc === -1) cLoc = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
  const cShift = H.findIndex(h => String(h).trim().toLowerCase() === 'shift');

  const map = {};
  for (let i = 1; i < hd.length; i++) {
    const id = String(hd[i][cId] || '').trim();
    if (!id) continue;
    map[id] = {
      name: cName !== -1 ? String(hd[i][cName] || '').trim() : id,
      location: cLoc !== -1 ? String(hd[i][cLoc] || '').trim() : '',
      shift: cShift !== -1 ? String(hd[i][cShift] || '').trim() : '',
    };
  }
  return map;
}

// ── Parse a Shifts-tab time cell (may be a Date/time value or an
// "HH:mm" string, depending on how the sheet cell is formatted). ────
function bc_parseTimeOfDay_(val) {
  if (val == null || val === '') return null;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    if (isNaN(val.getTime())) return null;
    return { h: val.getHours(), m: val.getMinutes() };
  }
  const s = String(val).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return { h: parseInt(m[1], 10), m: parseInt(m[2], 10) };
  return null;
}

// ── Build shift -> { start, end, overnight, normalHours } map using
// the Shifts tab's actual Start/End times (not a computed duration).
// FIX: the real header is "Shift Name" — matching only "Shift" found
// nothing, which silently emptied this entire map and forced every
// row to the 9h fallback. Matches "shift name" primarily, "shift" as
// a defensive fallback.
function bc_buildShiftTimesMap_(ss) {
  const sh = ss.getSheetByName('Shifts');
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cNm = H.findIndex(h => {
    const s = String(h).trim().toLowerCase();
    return s === 'shift name' || s === 'shift';
  });
  const cSt = H.findIndex(h => String(h).trim().toLowerCase() === 'start');
  const cEn = H.findIndex(h => String(h).trim().toLowerCase() === 'end');
  const cNh = H.findIndex(h => {
    const s = String(h).trim().toLowerCase();
    return s === 'normal hours' || s === 'normal hrs';
  });

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  data.forEach(row => {
    const name = String(row[cNm] || '').trim();
    if (!name) return;
    const start = cSt !== -1 ? bc_parseTimeOfDay_(row[cSt]) : null;
    const end   = cEn !== -1 ? bc_parseTimeOfDay_(row[cEn]) : null;
    const nh    = cNh !== -1 ? parseFloat(row[cNh]) : NaN;
    out[name] = {
      start: start,
      end: end,
      overnight: (start && end) ? ((end.h * 60 + end.m) < (start.h * 60 + start.m)) : false,
      normalHours: (!isNaN(nh) && nh > 0) ? nh : 9,
    };
  });
  return out;
}

// ── Location -> [shift names] map — mirrors RosterBackend.gs's
// rbGetShiftsForLocation_ exactly: reads Branch_Locations' "Shifts"
// column (comma-separated), falls back to Shifts-tab rows whose name
// contains the location string. Needed because per-employee Shift
// assignment on HR Maya is unreliable — nobody's been through the
// roster yet, so shift assignment still effectively lives at the
// location level, not the employee level. ────────────────────────────
function bc_getLocationShiftNames_(ss, location) {
  const locSheet = ss.getSheetByName('Branch_Locations');
  if (locSheet && locSheet.getLastRow() > 1) {
    const H = locSheet.getRange(1, 1, 1, locSheet.getLastColumn()).getValues()[0];
    const cB  = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cSh = H.findIndex(h => String(h).trim().toLowerCase() === 'shifts');
    if (cB !== -1 && cSh !== -1) {
      const data = locSheet.getRange(2, 1, locSheet.getLastRow() - 1, locSheet.getLastColumn()).getValues();
      for (const row of data) {
        if (String(row[cB] || '').trim().toLowerCase() === String(location).trim().toLowerCase()) {
          const list = String(row[cSh] || '').split(',').map(s => s.trim()).filter(Boolean);
          if (list.length) return list;
        }
      }
    }
  }
  // Fallback: Shifts tab rows whose name contains the location string.
  const sh = ss.getSheetByName('Shifts');
  const out = [];
  if (sh && sh.getLastRow() > 1) {
    const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const cNm = H.findIndex(h => {
      const s = String(h).trim().toLowerCase();
      return s === 'shift name' || s === 'shift';
    });
    const cAc = H.findIndex(h => String(h).trim().toLowerCase() === 'active');
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(row => {
      const name = String(row[cNm] || '').trim();
      if (!name) return;
      if (cAc !== -1 && String(row[cAc]).trim().toLowerCase() !== 'yes') return;
      if (name.toLowerCase().includes(String(location).trim().toLowerCase())) out.push(name);
    });
  }
  return out;
}

// ── Circular (wraparound-safe) distance in minutes between two
// times-of-day — so a 23:30 start and a 00:15 check-in read as 45
// minutes apart, not ~23 hours apart. ─────────────────────────────
function bc_circularMinuteDist_(aH, aM, bH, bM) {
  const a = aH * 60 + aM, b = bH * 60 + bM;
  const diff = Math.abs(a - b);
  return Math.min(diff, 1440 - diff);
}

// ── Pick the best-matching shift for a Check-In: among a location's
// candidate shifts, the one whose Start time is closest to the actual
// Check-In time. Returns { name, info } or null if no usable candidate. ─
function bc_pickBestShiftForCheckIn_(shiftTimesMap, candidateNames, ciH, ciM) {
  let best = null;
  candidateNames.forEach(name => {
    const info = shiftTimesMap[name];
    if (!info || !info.start || !info.end) return;
    const dist = bc_circularMinuteDist_(ciH, ciM, info.start.h, info.start.m);
    if (!best || dist < best.dist) best = { name, info, dist };
  });
  return best;
}

// ── Build empId|date -> hours map from OT_Pending, Status='Pending'
// rows only. Column layout confirmed: ID · Employee ID · Name · Date ·
// Hours · Source · Status · Submitted · Geo Flag · Computed OT
// (indices: 0..9). Uses the existing global normDayStr_ from
// OtSubmit.gs for date normalization — not redeclared here. ─────────
function bc_getPendingOtMap_(ss) {
  const pen = ss.getSheetByName('OT_Pending');
  const map = {};
  if (!pen || pen.getLastRow() < 2) return map;
  const data = pen.getRange(2, 1, pen.getLastRow() - 1, pen.getLastColumn()).getValues();
  data.forEach(row => {
    const empId  = String(row[1] || '').trim();   // Employee ID
    const status = String(row[6] || '').trim().toLowerCase(); // Status
    if (!empId || status !== 'pending') return;
    const dateStr = (typeof normDayStr_ === 'function') ? normDayStr_(row[3]) : String(row[3] || '').trim();
    const hours = Number(row[4]) || 0; // Hours
    if (!dateStr || hours <= 0) return;
    map[empId + '|' + dateStr] = hours;
  });
  return map;
}

// ── Peer checkouts: what did everyone ELSE at this location do that day?
// ------------------------------------------------------------
// Crews clock out together. On 12-09 at Central Kitchen Hawally the whole
// shift checked out at 14:13, not the 14:00 shift end. So the most accurate
// checkout for someone who forgot is what their COLLEAGUES actually did —
// not a generic shift-end guess.
//
// Returns { 'location|dd-mm-yyyy': [Date, ...] } of real Check Out
// timestamps so a median can be taken. Median, not mean: one person
// leaving hours late must not drag the whole estimate.
function bc_buildPeerCheckouts_(ss) {
  const log = ss.getSheetByName('Attendance_Log');
  const out = {};
  if (!log || log.getLastRow() < 2) return out;

  const H = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0];
  const col = (names) => {
    for (const n of names) {
      const i = H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
      if (i !== -1) return i;
    }
    return -1;
  };
  const cTs = col(['Timestamp']), cStage = col(['Stage']),
        cBr = col(['Branch','Location']), cDate = col(['Date']);
  if (cTs < 0 || cStage < 0 || cBr < 0) return out;

  log.getRange(2, 1, log.getLastRow()-1, log.getLastColumn()).getValues().forEach(r => {
    if (String(r[cStage] || '').trim() !== 'Check Out') return;
    const ts = parseLogTimestamp(r[cTs]);
    if (!ts) return;
    const loc = String(r[cBr] || '').trim();
    if (!loc) return;
    const dStr = cDate >= 0 ? normDayStr_(r[cDate])
                            : Utilities.formatDate(ts, tz_(), 'dd-MM-yyyy');
    const k = loc + '|' + dStr;
    (out[k] = out[k] || []).push(ts);
  });
  return out;
}

// Median clock time from a list of Date objects → { h, m, n }
function bc_medianTime_(list) {
  if (!list || !list.length) return null;
  const mins = list.map(d => d.getHours() * 60 + d.getMinutes()).sort((a, b) => a - b);
  const mid = Math.floor(mins.length / 2);
  const m = mins.length % 2 ? mins[mid] : Math.round((mins[mid-1] + mins[mid]) / 2);
  return { h: Math.floor(m / 60), m: m % 60, n: mins.length };
}

// ── Core detector: scans the FULL Attendance_Log, returns every
// dangling Check In (no matching Check Out before the next Check
// In / Day Off / end of data). Always full-history — callers filter
// by date afterward. ─────────────────────────────────────────────
function bc_findDanglingCheckIns_(ss) {
  const log = ss.getSheetByName('Attendance_Log');
  if (!log || log.getLastRow() < 2) return [];

  const data = log.getDataRange().getValues();
  const H = data[0];
  const cId    = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
  const cStage = H.findIndex(h => String(h).trim().toLowerCase() === 'stage');
  const cTs    = H.findIndex(h => String(h).trim().toLowerCase() === 'timestamp');
  const cBranch= H.findIndex(h => String(h).trim().toLowerCase() === 'branch');

  const byEmp = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][cId] || '').trim();
    if (!id) continue;
    const ts = (typeof parseLogTimestamp === 'function') ? parseLogTimestamp(data[i][cTs]) : null;
    if (!ts) continue;
    (byEmp[id] = byEmp[id] || []).push({
      row: i + 1, ts: ts,
      stage: String(data[i][cStage] || '').trim(),
      location: cBranch !== -1 ? String(data[i][cBranch] || '').trim() : '',
    });
  }

  const dangling = [];
  Object.keys(byEmp).forEach(empId => {
    const punches = byEmp[empId].sort((a, b) => a.ts - b.ts);
    let open = null;
    punches.forEach(p => {
      if (p.stage === 'Check In') {
        // A new Check In always supersedes any still-open prior session —
        // same convention already used throughout this codebase
        // (kioskSessionState_, SupervisorCorrect's getSupervisorSessions).
        if (open) dangling.push({ empId, ciRow: open.row, ciTs: open.ts, location: open.location, auth: open.auth });
        open = { row: p.row, ts: p.ts, location: p.location };
      } else if (p.stage === 'Check Out') {
        if (open) open = null; // closed cleanly
      } else if (p.stage === 'Day Off') {
        if (open) dangling.push({ empId, ciRow: open.row, ciTs: open.ts, location: open.location, auth: open.auth });
        open = null;
      }
      // Break Out / Break In with no open session: ignored (matches
      // existing convention elsewhere — nothing to orphan).
    });
    if (open) dangling.push({ empId, ciRow: open.row, ciTs: open.ts, location: open.location, auth: open.auth });
  });

  return dangling;
}

// ── Shared: detect + compute what WOULD be fixed, without writing.
// Returns { willFix: [...], skippedInProgress: [...] }. ─────────────
function bc_computeFixPlan_(ss, fromStr, toStr) {
  const empMap       = bc_buildEmpMap_(ss);
  const shiftTimes    = bc_buildShiftTimesMap_(ss);
  const pendingOtMap = bc_getPendingOtMap_(ss);
  const peerCheckouts = bc_buildPeerCheckouts_(ss);
  const dangling      = bc_findDanglingCheckIns_(ss);
  const now           = new Date();

  const willFix = [], skippedInProgress = [];

  dangling.forEach(d => {
    const emp = empMap[d.empId] || { name: d.empId, location: d.location, shift: '' };
    const ciDateStr = Utilities.formatDate(d.ciTs, tz_(), 'dd-MM-yyyy');

    // Date-range filter (only affects what's shown/acted on, not detection)
    if (fromStr || toStr) {
      const ciDay = new Date(d.ciTs.getFullYear(), d.ciTs.getMonth(), d.ciTs.getDate());
      if (fromStr) { const f = parseFlexibleDate_(fromStr); if (f && ciDay < f) return; }
      if (toStr)   { const t = parseFlexibleDate_(toStr);   if (t && ciDay > t) return; }
    }

    // ── SKIP 0: exempt locations are auto-present; never backfill them.
    const empLoc = emp.location || d.location || '';
    if (BACKFILL_EXEMPT_LOCATIONS.indexOf(empLoc) !== -1) {
      skippedInProgress.push({
        empId: d.empId, name: emp.name, location: empLoc,
        date: ciDateStr, ciRow: d.ciRow,
        checkIn: Utilities.formatDate(d.ciTs, tz_(), 'HH:mm'),
        shift: emp.shift || '', basis: '', otHours: 0,
        computedCheckOut: '', computedCheckOutFull: '', computedCheckOutDate: '',
        reason: empLoc + ' is auto-marked Present — no checkout needed',
      });
      return;
    }

    // ── SKIP 1: supervisors already corrected everything up to the cutoff.
    // Those rows are the authoritative record; do not guess over them.
    const cutoff = parseFlexibleDate_(BACKFILL_SKIP_ON_OR_BEFORE);
    const ciDay  = new Date(d.ciTs.getFullYear(), d.ciTs.getMonth(), d.ciTs.getDate());
    if (cutoff && ciDay <= cutoff) {
      skippedInProgress.push({
        empId: d.empId, name: emp.name, location: emp.location || d.location,
        date: ciDateStr, ciRow: d.ciRow,
        checkIn: Utilities.formatDate(d.ciTs, tz_(), 'HH:mm'),
        shift: emp.shift || '', basis: '', otHours: 0,
        computedCheckOut: '', computedCheckOutFull: '', computedCheckOutDate: '',
        reason: 'On or before ' + BACKFILL_SKIP_ON_OR_BEFORE +
                ' — already reviewed and corrected by supervisors',
      });
      return;
    }

    // ── SKIP 2: a human already touched this check-in.
    if (d.auth && BACKFILL_HUMAN_AUTH.indexOf(d.auth) !== -1) {
      skippedInProgress.push({
        empId: d.empId, name: emp.name, location: emp.location || d.location,
        date: ciDateStr, ciRow: d.ciRow,
        checkIn: Utilities.formatDate(d.ciTs, tz_(), 'HH:mm'),
        shift: emp.shift || '', basis: '', otHours: 0,
        computedCheckOut: '', computedCheckOutFull: '', computedCheckOutDate: '',
        reason: 'Check-in is "' + d.auth + '" — a human already decided this row',
      });
      return;
    }

    const shiftInfo = shiftTimes[emp.shift] || null;

    // TIER 0 (peer checkouts) was REMOVED. The idea was sound for a
    // single-shift site — the Central Kitchen crew all leave together at
    // 14:13 — but every other location runs two shifts, so the peer pool
    // mixes First Shift (ends 21:30) with Second Shift (ends 03:30) and
    // the median lands nowhere real: Ardiya came out at 04:10, Kaifan at
    // 11:26. Correcting it would mean grouping peers by shift as well as
    // location and date, which is real complexity to gain ~13 minutes
    // over the shift end time. The shift-based tiers below are accurate
    // enough and far easier to reason about.

    let computedCheckOut, basis, shiftUsed;

    if (!computedCheckOut && emp.shift && shiftInfo && shiftInfo.end) {
      // Explicit HR Maya assignment exists and matches a real shift — trust it.
      const base = new Date(d.ciTs.getFullYear(), d.ciTs.getMonth(), d.ciTs.getDate());
      if (shiftInfo.overnight) base.setDate(base.getDate() + 1);
      computedCheckOut = new Date(base.getFullYear(), base.getMonth(), base.getDate(),
        shiftInfo.end.h, shiftInfo.end.m, 0);
      basis = emp.shift + ' end time (assigned)';
      shiftUsed = emp.shift;
    } else {
      // No reliable per-employee assignment (expected pre-roster — shift
      // assignment still effectively lives at the location, not the
      // employee, until someone's been through the roster). Infer the
      // best-matching shift for this location by nearest Start time to
      // the actual Check In.
      const candidates = bc_getLocationShiftNames_(ss, emp.location || d.location);
      const ciH = d.ciTs.getHours(), ciM = d.ciTs.getMinutes();
      const pick = bc_pickBestShiftForCheckIn_(shiftTimes, candidates, ciH, ciM);

      if (pick) {
        const base = new Date(d.ciTs.getFullYear(), d.ciTs.getMonth(), d.ciTs.getDate());
        if (pick.info.overnight) base.setDate(base.getDate() + 1);
        computedCheckOut = new Date(base.getFullYear(), base.getMonth(), base.getDate(),
          pick.info.end.h, pick.info.end.m, 0);
        basis = pick.name + ' end time (inferred from location + check-in time)';
        shiftUsed = pick.name;
      } else {
        // No shift data available for this location at all — last resort.
        computedCheckOut = new Date(d.ciTs.getTime() + 9 * 3600000);
        basis = 'no shift data available for this location — used 9h from check-in';
        shiftUsed = '(none)';
      }
    }

    // Add any Pending self-submitted OT for this date, on top of the
    // shift-end-based checkout. Rolls the day forward automatically
    // if the OT pushes past midnight.
    // ── DOUBLE-COUNT FIX ────────────────────────────────────────────
    // This used to ADD the employee's pending OT hours onto the
    // synthesized checkout. That paid the same overtime twice:
    //
    //   checkout = shift end + claimed OT
    //     → presence inflated in Attendance_Daily
    //     → getAttendanceOtDraft computes presence - normalHours
    //     → that inflated OT gets posted to OT_Entries
    //   AND the original OT_Pending row is still Pending
    //     → approving it posts the SAME hours to OT_Entries again
    //   getAllOtHours_ sums both. Paid twice.
    //
    // The checkout is now the plain shift end time. The OT claim keeps
    // its own approval path into OT_Entries, which is the only route
    // that should ever reach payroll.
    //
    // The claim is still READ here, but only to show it in the preview
    // so you can see what the employee said about the same day.
    const otHours = pendingOtMap[d.empId + '|' + ciDateStr] || 0;

    const entry = {
      empId: d.empId, name: emp.name, location: emp.location || d.location,
      date: ciDateStr,
      ciRow: d.ciRow,
      checkIn: Utilities.formatDate(d.ciTs, tz_(), 'HH:mm'),
      shift: shiftUsed,
      basis: basis,
      otHours: otHours,
      computedCheckOut: Utilities.formatDate(computedCheckOut, tz_(), 'HH:mm'),
      computedCheckOutFull: Utilities.formatDate(computedCheckOut, tz_(), 'dd/MM/yyyy HH:mm:ss'),
      computedCheckOutDate: Utilities.formatDate(computedCheckOut, tz_(), 'dd-MM-yyyy'),
    };

    // ── REJECT: implausibly long session ──
    // An 02:35 check-in paired with a 20:30 shift end is 18 hours. Those
    // 02:35 punches are phantoms from the kiosk failure, and the employee
    // usually has a second, real check-in later the same day which is
    // matched correctly. Refuse to write the phantom one.
    if (computedCheckOut && (computedCheckOut - d.ciTs) / 3600000 > BACKFILL_MAX_SESSION_HRS) {
      skippedInProgress.push({
        empId: d.empId, name: emp.name, location: emp.location || d.location,
        date: ciDateStr, ciRow: d.ciRow,
        checkIn: Utilities.formatDate(d.ciTs, tz_(), 'HH:mm'),
        shift: shiftUsed, basis: basis, otHours: 0,
        computedCheckOut: Utilities.formatDate(computedCheckOut, tz_(), 'HH:mm'),
        computedCheckOutFull: '', computedCheckOutDate: '',
        reason: 'Would be ' + ((computedCheckOut - d.ciTs)/3600000).toFixed(1) +
                'h — likely a phantom check-in, not a real shift',
      });
      return;
    }

    if (computedCheckOut >= now) {
      entry.reason = 'Still in progress — computed checkout would be in the future';
      skippedInProgress.push(entry);
    } else {
      willFix.push(entry);
    }
  });

  willFix.sort((a, b) => (a.date + a.name).localeCompare(b.date + b.name));
  skippedInProgress.sort((a, b) => (a.date + a.name).localeCompare(b.date + b.name));

  return { willFix, skippedInProgress };
}

// ── ENDPOINT: previewMissingCheckouts — read-only. ───────────────────
// params: { from (dd-mm-yyyy, optional), to (dd-mm-yyyy, optional) }
function previewMissingCheckouts(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const plan = bc_computeFixPlan_(ss, params.from || '', params.to || '');
    return {
      success: true,
      willFix: plan.willFix,
      skippedInProgress: plan.skippedInProgress,
      willFixCount: plan.willFix.length,
      skippedCount: plan.skippedInProgress.length,
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ── ENDPOINT: commitMissingCheckouts — writes + recompiles. ──────────
// Re-detects fresh (ignores anything the client saw in preview).
// params: { from, to, by }
function commitMissingCheckouts(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const plan = bc_computeFixPlan_(ss, params.from || '', params.to || '');

    if (!plan.willFix.length) {
      return { success: true, fixed: 0, skipped: plan.skippedInProgress.length,
        recompiledDates: [], message: 'Nothing to fix in this range.' };
    }

    const log = ss.getSheetByName('Attendance_Log');
    const H = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0];
    const cId=H.findIndex(h=>String(h).trim().toLowerCase()==='employee id');
    const cName=H.findIndex(h=>String(h).trim().toLowerCase()==='employee name');
    const cDate=H.findIndex(h=>String(h).trim().toLowerCase()==='date');
    const cStage=H.findIndex(h=>String(h).trim().toLowerCase()==='stage');
    const cTs=H.findIndex(h=>String(h).trim().toLowerCase()==='timestamp');
    const cAuth=H.findIndex(h=>String(h).trim().toLowerCase()==='auth');
    const cBranch=H.findIndex(h=>String(h).trim().toLowerCase()==='branch');
    const cFlag=H.findIndex(h=>String(h).trim().toLowerCase()==='geo flag');

    const touchedDates = new Set();

    plan.willFix.forEach(f => {
      const row = new Array(H.length).fill('');
      row[cId] = f.empId;
      if (cName !== -1) row[cName] = f.name;
      row[cDate] = f.computedCheckOutDate; // dated by when it actually falls (handles overnight)
      row[cStage] = 'Check Out';
      row[cTs] = f.computedCheckOutFull;
      if (cAuth !== -1) row[cAuth] = 'Backfilled';
      if (cBranch !== -1) row[cBranch] = f.location;
      if (cFlag !== -1) {
        row[cFlag] = 'BACKFILLED (' + f.basis + (f.otHours > 0 ? '; claimed ' + f.otHours + 'h OT separately' : '') + ')';
      }
      log.appendRow(row);
      log.getRange(log.getLastRow(), cDate + 1).setNumberFormat('@');

      if (typeof logAmend_ === 'function') {
        logAmend_(ss, params.by || 'System (backfill)', 'ADD', f.empId, f.date, 'Check Out',
          'synthesized: ' + f.checkIn + ' -> ' + f.computedCheckOut + ' (' + f.basis +
          (f.otHours > 0 ? '; employee separately claimed ' + f.otHours + 'h OT' : '') + '). Check In left untouched.');
      }

      touchedDates.add(f.date); // Check-In date — the date Attendance_Daily sessionizes this under
    });

    // Recompile Attendance_Daily for every date actually touched.
    const recompiled = [];
    if (typeof writeDayToAttendanceDaily_ === 'function') {
      touchedDates.forEach(dateStr => {
        try {
          writeDayToAttendanceDaily_(dateStr);
          recompiled.push(dateStr);
        } catch (e) {
          Logger.log('commitMissingCheckouts: recompile failed for ' + dateStr + ' — ' + e.message);
        }
      });
    }

    return {
      success: true,
      fixed: plan.willFix.length,
      skipped: plan.skippedInProgress.length,
      recompiledDates: recompiled,
      message: 'Added ' + plan.willFix.length + ' checkout(s). ' +
        plan.skippedInProgress.length + ' skipped (still in progress). ' +
        'Recompiled Attendance_Daily for ' + recompiled.length + ' date(s).',
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}
// ── Editor-friendly preview ─────────────────────────────────────────
// previewMissingCheckouts() returns JSON for the web UI and prints
// nothing, so running it from the Apps Script editor looks like it did
// nothing. This prints a readable report to the log instead.
//
// Writes nothing. Run: previewMissingCheckoutsLog()
function previewMissingCheckoutsLog() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const plan = bc_computeFixPlan_(ss, '', '');   // whole history

  const L = [];
  L.push('══════════════════════════════════════════════');
  L.push(' BACKFILL PREVIEW — nothing written');
  L.push('   CODE VERSION : ' + BACKFILL_VERSION);
  L.push('   (if this does not match the const in BackfillCheckouts.gs, an older copy is running)');
  L.push('   Would add : ' + plan.willFix.length + ' checkout(s)');
  L.push('   Skipped   : ' + plan.skippedInProgress.length + ' (still in progress)');
  L.push('══════════════════════════════════════════════');
  L.push('');

  // Group by which rule produced the answer, so weak guesses stand out.
  const byTier = { peer: [], assigned: [], inferred: [], fallback: [] };
  plan.willFix.forEach(f => {
    if (f.basis.indexOf('colleague') !== -1)      byTier.peer.push(f);
    else if (f.basis.indexOf('(assigned)') !== -1) byTier.assigned.push(f);
    else if (f.basis.indexOf('inferred') !== -1)   byTier.inferred.push(f);
    else                                           byTier.fallback.push(f);
  });

  // If Tier 0 produced nothing, show the peer evidence for the first few
  // rows so the reason is visible in the same run rather than needing
  // another diagnostic pass.
  if (!byTier.peer.length) {
    const peersDbg = bc_buildPeerCheckouts_(ss);
    L.push('!! TIER 0 MATCHED NOTHING — peer evidence for the first 5 rows:');
    plan.willFix.slice(0, 5).forEach(f => {
      const nd = (function(){
        const p = f.date.split('-');
        const dd = new Date(+p[2], +p[1]-1, +p[0] + 1);
        return Utilities.formatDate(dd, tz_(), 'dd-MM-yyyy');
      })();
      const kSame = f.location + '|' + f.date, kNext = f.location + '|' + nd;
      const aSame = peersDbg[kSame] || [], aNext = peersDbg[kNext] || [];
      const mSame = bc_medianTime_(aSame), mNext = bc_medianTime_(aNext);
      L.push('   ' + f.empId + '  ' + f.date + '  in ' + f.checkIn);
      L.push('      "' + kSame + '" → ' + aSame.length + ' peer(s)' +
             (mSame ? '  median ' + String(mSame.h).padStart(2,'0') + ':' + String(mSame.m).padStart(2,'0') : ''));
      L.push('      "' + kNext + '" → ' + aNext.length + ' peer(s)' +
             (mNext ? '  median ' + String(mNext.h).padStart(2,'0') + ':' + String(mNext.m).padStart(2,'0') : ''));
    });
    L.push('');
  }

  L.push('HOW EACH CHECKOUT WAS DECIDED');
  L.push('   Tier 0  colleagues at same location : ' + byTier.peer.length + '   (most reliable)');
  L.push('   Tier 1  employee\'s assigned shift   : ' + byTier.assigned.length);
  L.push('   Tier 2  shift inferred from location: ' + byTier.inferred.length);
  L.push('   Tier 3  check-in + 9h fallback      : ' + byTier.fallback.length + '   (review these)');
  L.push('');

  const dump = (title, arr) => {
    L.push('── ' + title + '  (' + arr.length + ') ──');
    if (!arr.length) { L.push('   none'); L.push(''); return; }
    arr.forEach(f => {
      L.push('   ' + f.empId + '  ' + (f.name || '').substring(0, 22).padEnd(22) +
             f.date + '  in ' + f.checkIn + ' → out ' + f.computedCheckOut +
             (f.computedCheckOutDate !== f.date ? ' (' + f.computedCheckOutDate + ')' : '') +
             (f.otHours > 0 ? '   [claimed ' + f.otHours + 'h OT separately]' : ''));
      L.push('        ' + f.basis);
    });
    L.push('');
  };

  dump('TIER 3 — weakest basis, check these first', byTier.fallback);
  dump('TIER 2 — shift inferred', byTier.inferred);
  dump('TIER 1 — assigned shift', byTier.assigned);
  dump('TIER 0 — matched colleagues', byTier.peer);

  if (plan.skippedInProgress.length) {
    L.push('── SKIPPED: still in progress ──');
    plan.skippedInProgress.forEach(f =>
      L.push('   ' + f.empId + '  ' + (f.name || '') + '  ' + f.date + '  in ' + f.checkIn));
    L.push('');
  }

  L.push('══════════════════════════════════════════════');
  L.push(' If this looks right, commit from the Attendance page,');
  L.push(' or run commitMissingCheckouts({by:"you@..."}) here.');
  L.push('══════════════════════════════════════════════');

  const out = L.join('\n');
  Logger.log(out);
  return out;
}

// ── Why is the peer (Tier 0) rule not firing? ───────────────────────
// Every preview row is coming back as Tier 2 ("inferred from location +
// check-in time"), which means bc_buildPeerCheckouts_ is not producing a
// match for any dangling check-in. This prints exactly what the peer map
// contains versus what is being looked up, so the mismatch is visible
// rather than guessed at.
//
// Writes nothing. Run: diagnosePeerLookup()
function diagnosePeerLookup() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const peers = bc_buildPeerCheckouts_(ss);
  const keys = Object.keys(peers);
  const L = [];

  L.push('══════════════════════════════════════════════');
  L.push(' PEER LOOKUP DIAGNOSTIC');
  L.push('══════════════════════════════════════════════');
  L.push('');
  L.push('Peer map entries built: ' + keys.length);
  if (!keys.length) {
    L.push('');
    L.push('EMPTY. bc_buildPeerCheckouts_ found no Check Out rows with a');
    L.push('usable Branch/Location value. Check that Attendance_Log has a');
    L.push('"Branch" (or "Location") column and that Check Out rows carry it.');
    Logger.log(L.join('\n'));
    return L.join('\n');
  }

  L.push('');
  L.push('Sample of keys IN THE MAP (location|date → how many checkouts):');
  keys.slice(0, 15).forEach(k => L.push('   "' + k + '"  → ' + peers[k].length));
  L.push('');

  // Now show what the fix plan is actually looking UP
  const empMap = bc_buildEmpMap_(ss);
  const dangling = bc_findDanglingCheckIns_(ss);
  L.push('Dangling check-ins found: ' + dangling.length);
  L.push('');
  L.push('First 15 LOOKUPS the plan attempts:');
  dangling.slice(0, 15).forEach(d => {
    const emp = empMap[d.empId] || {};
    const loc = emp.location || d.location || '(blank)';
    const dStr = Utilities.formatDate(d.ciTs, tz_(), 'dd-MM-yyyy');
    const nextD = Utilities.formatDate(new Date(d.ciTs.getTime() + 86400000), tz_(), 'dd-MM-yyyy');
    const k1 = loc + '|' + dStr, k2 = loc + '|' + nextD;
    const n1 = peers[k1] ? peers[k1].length : 0;
    const n2 = peers[k2] ? peers[k2].length : 0;
    L.push('   ' + d.empId + '  emp.location="' + (emp.location || '') +
           '"  punch.location="' + (d.location || '') + '"');
    L.push('        looking up "' + k1 + '" → ' + n1 + ' peer(s)');
    L.push('        looking up "' + k2 + '" → ' + n2 + ' peer(s)');
    if (!n1 && !n2) {
      // Find what keys DO exist for that date, to expose the difference
      const same = keys.filter(k => k.indexOf('|' + dStr) !== -1);
      if (same.length) {
        L.push('        !! no match. Keys that DO exist for ' + dStr + ':');
        same.slice(0, 6).forEach(k => L.push('             "' + k + '"  (' + peers[k].length + ')'));
      } else {
        L.push('        !! no peer checkouts exist for ' + dStr + ' at any location');
      }
    }
  });

  L.push('');
  L.push('══════════════════════════════════════════════');
  const out = L.join('\n');
  Logger.log(out);
  return out;
}