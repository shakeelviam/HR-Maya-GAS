// ============================================================
// RETIRED — AttendanceDaily.gs
// ============================================================
// This was the ORIGINAL attendance compile. It has been replaced by
// AttendanceDailyTrigger.gs (the 4 AM compile) which is roster-aware,
// transfer-aware, and uses AttendanceCore.gs for sessionizing.
//
// WHY IT HAD TO GO, not just be left alone:
//
//  1. IT DISAGREED WITH THE LIVE COMPILE. Given identical punches this
//     file wrote  Status='Day Off', Source='Default'  as a blind guess
//     for anyone with no punch. The live compile writes 'No Roster' or
//     'Not Recorded' — honest about not knowing. 329 rows in
//     Attendance_Daily still carry Source='Default' from this file.
//
//  2. IT WAS STILL CALLABLE. Routed in Code.gs as 'buildAttendanceDaily'.
//     Anyone running it would silently overwrite good roster-aware rows
//     with the old guesswork.
//
//  3. IT CARRIED ITS OWN SESSIONIZER — one of seven copies of the same
//     logic, each free to drift from the others.
//
// The stub below returns a clear error rather than deleting the
// function, so an old bookmark, router case or scheduled trigger fails
// LOUDLY instead of quietly doing the wrong thing.
//
// TO REBUILD ATTENDANCE: use recompileFullRange() in RecompileRange.gs,
// which calls writeDayToAttendanceDaily_ — the single live compile.
// ============================================================

function buildAttendanceDaily(params) {
  return {
    success: false,
    retired: true,
    error: 'buildAttendanceDaily has been RETIRED. It used the old ' +
           'pre-roster logic and would overwrite correct rows with ' +
           'guessed "Default Day Off" values. Use recompileFullRange() ' +
           '(RecompileRange.gs) or the 4 AM dailyAttendanceTrigger instead.'
  };
}

function rebuildAttendanceDaily(params) {
  return buildAttendanceDaily(params);
}