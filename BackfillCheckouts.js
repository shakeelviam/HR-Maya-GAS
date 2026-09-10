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
//   - If the employee has a Pending row in OT_Pending for that same
//     date, those hours are added ON TOP of the shift's end time —
//     they told us they worked that OT, so the attendance record
//     shouldn't be blind to it. This does NOT touch OT_Pending's
//     status or write to OT_Entries — that OT claim still goes
//     through the normal Approve/Reject review separately.
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
        if (open) dangling.push({ empId, ciRow: open.row, ciTs: open.ts, location: open.location });
        open = { row: p.row, ts: p.ts, location: p.location };
      } else if (p.stage === 'Check Out') {
        if (open) open = null; // closed cleanly
      } else if (p.stage === 'Day Off') {
        if (open) dangling.push({ empId, ciRow: open.row, ciTs: open.ts, location: open.location });
        open = null;
      }
      // Break Out / Break In with no open session: ignored (matches
      // existing convention elsewhere — nothing to orphan).
    });
    if (open) dangling.push({ empId, ciRow: open.row, ciTs: open.ts, location: open.location });
  });

  return dangling;
}

// ── Shared: detect + compute what WOULD be fixed, without writing.
// Returns { willFix: [...], skippedInProgress: [...] }. ─────────────
function bc_computeFixPlan_(ss, fromStr, toStr) {
  const empMap       = bc_buildEmpMap_(ss);
  const shiftTimes    = bc_buildShiftTimesMap_(ss);
  const pendingOtMap = bc_getPendingOtMap_(ss);
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

    const shiftInfo = shiftTimes[emp.shift] || null;
    let computedCheckOut, basis, shiftUsed;

    if (emp.shift && shiftInfo && shiftInfo.end) {
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
    const otHours = pendingOtMap[d.empId + '|' + ciDateStr] || 0;
    if (otHours > 0) {
      computedCheckOut = new Date(computedCheckOut.getTime() + otHours * 3600000);
    }

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
        row[cFlag] = 'BACKFILLED (' + f.basis + (f.otHours > 0 ? ' + ' + f.otHours + 'h pending OT' : '') + ')';
      }
      log.appendRow(row);
      log.getRange(log.getLastRow(), cDate + 1).setNumberFormat('@');

      if (typeof logAmend_ === 'function') {
        logAmend_(ss, params.by || 'System (backfill)', 'ADD', f.empId, f.date, 'Check Out',
          'synthesized: ' + f.checkIn + ' -> ' + f.computedCheckOut + ' (' + f.basis +
          (f.otHours > 0 ? ', +' + f.otHours + 'h pending OT' : '') + '). Check In left untouched.');
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