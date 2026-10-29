// ============================================================
// RecompileRange.gs — rebuild Attendance_Daily across a date range
// ------------------------------------------------------------
// manualDayWrite() handles one date at a time and needs the date
// edited in code each run. This does a range in one go.
//
// SAFE TO RE-RUN: writeDayToAttendanceDaily_ deletes each date's rows
// before rewriting them, so recompiling the same range twice produces
// the same result. It reads Attendance_Log, Roster, Leave_Grants,
// Shifts and HR Maya — it never writes to any of them.
//
// TWO PRESETS:
//   recompileCleanupDates()  → 01-09, 02-09, 03-09
//                              (after the junk punch deletions)
//   recompileFullRange()     → RANGE_FROM → RANGE_TO
//                              (fixes the 329 rows still carrying
//                               source 'Default' from the retired
//                               compile — they were never rebuilt
//                               with roster/transfer/No-Roster logic)
//
// APPS SCRIPT HAS A 6-MINUTE EXECUTION LIMIT. A long range can hit it.
// If that happens the log tells you the last date completed — set
// RANGE_FROM to the next day and run again. Nothing is corrupted by
// stopping partway; each date is written independently.
// ============================================================

const RANGE_FROM = '22-08-2026';
const RANGE_TO   = '20-09-2026';         // blank = yesterday

// ── Preset 1: just the three dates affected by the junk cleanup ──────
function recompileCleanupDates() {
  return rc_run_(['01-09-2026', '02-09-2026', '03-09-2026'],
                 'cleanup dates (after junk punch deletion)');
}

// ── Preset 2: the whole range ────────────────────────────────────────
function recompileFullRange() {
  const from = parseFlexibleDate_(RANGE_FROM);
  let to;
  if (RANGE_TO) to = parseFlexibleDate_(RANGE_TO);
  else { to = new Date(); to.setDate(to.getDate() - 1); }
  if (!from || !to) throw new Error('Bad RANGE_FROM / RANGE_TO.');

  const dates = [];
  for (let d = startOfDay_(from); d <= startOfDay_(to); d.setDate(d.getDate() + 1)) {
    dates.push(formatDdMmYyyy_(d));
  }
  return rc_run_(dates, RANGE_FROM + ' → ' + formatDdMmYyyy_(to));
}

// ── Shared runner ────────────────────────────────────────────────────
function rc_run_(dates, label) {
  const started = new Date();
  const LIMIT_MS = 5 * 60 * 1000;        // stop before the 6-minute cap
  const done = [], failed = [];
  let stoppedEarly = null;

  Logger.log('Recompiling ' + dates.length + ' date(s): ' + label);

  for (let i = 0; i < dates.length; i++) {
    if (new Date() - started > LIMIT_MS) {
      stoppedEarly = dates[i];
      break;
    }
    try {
      writeDayToAttendanceDaily_(dates[i]);
      done.push(dates[i]);
      Logger.log('  ✓ ' + dates[i]);
    } catch (err) {
      failed.push(dates[i] + ' — ' + err.message);
      Logger.log('  ✗ ' + dates[i] + ' — ' + err.message);
    }
  }

  const L = [];
  L.push('');
  L.push('══════════════════════════════════════════════');
  L.push(' RECOMPILE COMPLETE');
  L.push('  Recompiled : ' + done.length + ' date(s)');
  L.push('  Failed     : ' + failed.length);
  if (failed.length) failed.forEach(f => L.push('     ' + f));
  if (stoppedEarly) {
    L.push('');
    L.push('  STOPPED EARLY to stay inside the 6-minute limit.');
    L.push('  Last completed : ' + (done.length ? done[done.length - 1] : '(none)'));
    L.push('  Set RANGE_FROM = ' + stoppedEarly + ' and run again.');
  }
  L.push('══════════════════════════════════════════════');

  const msg = L.join('\n');
  Logger.log(msg);
  try {
    SpreadsheetApp.openById(MAIN_SPREADSHEET_ID)
      .toast('Recompiled ' + done.length + ' date(s)' +
             (stoppedEarly ? ' — more to do, see log' : ''), 'Recompile', 10);
  } catch (e) {}

  return { success: true, recompiled: done.length, failed: failed,
           stoppedEarly: stoppedEarly, dates: done };
}