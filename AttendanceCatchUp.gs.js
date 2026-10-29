// ============================================================
// AttendanceCatchUp.gs — run the attendance compile by hand
// ------------------------------------------------------------
// The nightly trigger (dailyAttendanceTrigger, 04:00) rebuilds a
// rolling 14-day window ending YESTERDAY. Two things it cannot do:
// show you today, and recover a morning when it did not fire.
//
// These do. All three take no arguments, so the editor's Run button
// works on them directly, and no date is ever typed into the code.
//
//   compileToday()            — today only. Safe to run any time; a
//                               shift still running shows what it has
//                               so far and corrects itself later.
//   compileCatchUp()          — the same 14-day window the nightly
//                               run does, but ending TODAY. This is
//                               the one to press when 04:00 missed.
//   showAttendanceTriggers()  — lists what is actually scheduled, so
//                               you can see whether the 04:00 trigger
//                               still exists.
//
// Rebuilding a date REPLACES its rows, it never duplicates them, so
// any of these can be run twice with no harm.
// ============================================================

function acDay_(offsetDays) {
  const d = new Date();
  if (offsetDays) d.setDate(d.getDate() + offsetDays);
  if (typeof adDateToStr_ !== 'function') {
    throw new Error('AttendanceDailyTrigger.gs is not in this project — adDateToStr_ is missing.');
  }
  return adDateToStr_(d);
}

// ------------------------------------------------------------
// Today, on its own
// ------------------------------------------------------------
function compileToday() {
  if (typeof rebuildAttendanceDaily_ !== 'function') {
    Logger.log('FAILED: AttendanceDailyTrigger.gs is not in this project.');
    return { success: false, error: 'rebuildAttendanceDaily_ is missing.' };
  }
  const today = acDay_(0);
  Logger.log('compileToday: rebuilding ' + today);
  const res = rebuildAttendanceDaily_(today, today);
  Logger.log('compileToday: ' + (res && res.message ? res.message : JSON.stringify(res)));
  return res;
}

// ------------------------------------------------------------
// The nightly window, but ending today instead of yesterday —
// what to run when the 04:00 trigger did not fire
// ------------------------------------------------------------
function compileCatchUp() {
  if (typeof rebuildAttendanceDaily_ !== 'function') {
    Logger.log('FAILED: AttendanceDailyTrigger.gs is not in this project.');
    return { success: false, error: 'rebuildAttendanceDaily_ is missing.' };
  }
  const days = (typeof AD_NIGHTLY_WINDOW_DAYS === 'number' && AD_NIGHTLY_WINDOW_DAYS > 0)
                 ? AD_NIGHTLY_WINDOW_DAYS : 14;
  const from = acDay_(-(days - 1));
  const to   = acDay_(0);
  Logger.log('compileCatchUp: rebuilding ' + from + ' → ' + to + ' (' + days + ' days, ending today)');
  const res = rebuildAttendanceDaily_(from, to);
  Logger.log('compileCatchUp: ' + (res && res.message ? res.message : JSON.stringify(res)));
  return res;
}

// ------------------------------------------------------------
// What is actually scheduled right now
// ------------------------------------------------------------
function showAttendanceTriggers() {
  const all = ScriptApp.getProjectTriggers();
  if (!all.length) {
    Logger.log('There are NO triggers installed on this project at all.');
    return { success: true, count: 0, handlers: [] };
  }
  const handlers = [];
  Logger.log(all.length + ' trigger(s) installed:');
  all.forEach(function (t) {
    const h = t.getHandlerFunction();
    handlers.push(h);
    Logger.log('   ' + h + '   (' + t.getEventType() + ')');
  });
  const hasDaily = handlers.indexOf('dailyAttendanceTrigger') !== -1;
  Logger.log(hasDaily
    ? 'dailyAttendanceTrigger IS installed. If a morning was missed, it failed or was skipped — check Executions, not this.'
    : 'dailyAttendanceTrigger is NOT installed. Nothing is compiling attendance overnight. Run setupDailyTrigger() to put it back.');
  return { success: true, count: all.length, handlers: handlers, dailyInstalled: hasDaily };
}