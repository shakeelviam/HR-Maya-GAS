// ============================================================
// AttendanceDailyTrigger.gs  — v5 (range rebuild)
// ------------------------------------------------------------
// Daily trigger: runs at 4 AM, writes attendance into Attendance_Daily.
//
// SETUP (run once from the editor):
//   1. Open Apps Script editor
//   2. Select setupDailyTrigger → Run
//   3. Approve permissions when prompted
//   That installs a daily 4 AM trigger on dailyAttendanceTrigger().
//
// ════════════════════════════════════════════════════════════
// WHAT CHANGED IN v5, AND WHY
//
// Attendance_Daily is not a record. It is a CONCLUSION drawn from six
// inputs — Attendance_Log, Leave_Grants, Roster, Manual_Attendance_Aug,
// Shifts and HR Maya. Every one of those can change after the conclusion
// was drawn: leave granted a week late, a punch corrected, a roster
// published after the fact. When that happens the conclusion is stale and
// every reader downstream — the supervisor page, OT verification, payroll
// — believes a lie.
//
// Real case, 20-09-2026: sick leave for MT-00033 covering 31-08 → 16-09
// was granted on the 20th. Those days had been compiled days earlier, so
// they still read "No Roster", and his supervisor was being asked to
// explain four days the man was signed off for.
//
// The cure is to rebuild whenever an input changes. That was impossible
// before because a rebuild was too expensive to call: v4 deleted the
// target date's rows ONE AT A TIME (adSheet.deleteRow in a loop). Measured
// live on this sheet: ~24 seconds per date for 68 rows. Recompiling three
// weeks took 8+ minutes and blew the 6-minute limit, so nobody called it,
// so the sheet stayed stale.
//
// v5 rebuilds a RANGE in one pass:
//   • every input is read ONCE, not once per date
//   • rows are computed in memory for every date in the range
//   • the sheet is rewritten with ONE setValues and at most ONE
//     deleteRows — no per-row deletion anywhere
// The same three weeks now cost about what a single date used to.
//
// IDENTICAL OUTPUT IS THE WHOLE POINT. Every decision branch below is
// v4's, unchanged. The only things that changed are WHEN inputs are read
// and HOW rows reach the sheet. The per-date punch admission test is
// still applied per date (see adPunchesForDate_) precisely so that a
// range rebuild produces byte-identical rows to running v4 date by date.
// Verified against v4 by differential test before shipping.
//
// writeDayToAttendanceDaily_(dateStr) keeps its exact old signature and
// behaviour, so RecompileRange.gs, SupervisorVerify.gs, manualDayWrite()
// and the 4 AM trigger all keep working untouched.
//
// NEW ENTRY POINTS for keeping the sheet honest:
//   rebuildAttendanceDaily_(from, to)  — rebuild a date range, fast
//   adTouchRange_(from, to)            — same, but never throws and
//                                        obeys a time budget; this is
//                                        what other files call after
//                                        changing an input
//   adTouchDates_([d1, d2, ...])       — same for a scattered set
//
// Call adTouchRange_ from anything that changes an input:
//   grant / amend / cancel leave, add / correct / delete a punch,
//   publish a roster, write a manual day, change a shift's timings.
// One line each. Past dates only — future dates have no rows yet and get
// built when the day arrives.
// ════════════════════════════════════════════════════════════
//
// Priority order per employee per day (highest → lowest):
//   0. Manual_Aug entry        (22-31 Aug supervisor backfill — highest
//                               authority; a human stated what happened)
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
//   9. Driver_Default (Driver Location, no roster needed) — only up to
//      DRIVER_AUTO_PRESENT_UNTIL; after that drivers are punch-based
//  10. No roster entry at all         → "No Roster" (No_Roster)
//        — replaces the old blind "Default Day Off" guess.
//
// Depends on RosterBackend.gs being deployed in the same project:
//   rbGetPublishedRosterForDate_(dateStr) -> { empId: {status, shift, location} }
//   rbLogDiscrepancy_(empId, name, location, date, rosterSaid, actual)
// ============================================================

// ── Roster go-live ───────────────────────────────────────────────────
// The roster feature did not exist before this date. For any earlier
// day, "No Roster" would be flagging the absence of a system that had
// not been built yet — meaningless noise that hides the days that do
// need attention.
//
// Before this date : no punch / leave / auto-present  →  "Not Recorded"
//                    (source Pre_Roster). Honest: nobody knows what they
//                    did, and no roster could have said. NOT deducted.
// On or after      : the same case  →  "No Roster" (source No_Roster),
//                    which correctly tells a supervisor to go and plan.
//                    Also NOT deducted — an admin gap is not the
//                    employee's fault.
//
// Neither status deducts pay. The distinction is about the sheet telling
// the truth rather than crying wolf on 3 weeks of history.
const ROSTER_GO_LIVE = '08-09-2026';

// ── DRIVERS: auto-present up to this date, punch-based after it ──────
// Until 23-09-2026 drivers had no kiosk credentials, so Driver Location
// resolved to Present every day regardless of punches (see the block in
// adBuildRowsForDate_). From 24-09-2026 the drivers have their own
// credentials and punch like everyone else, so their real punches decide
// the day. Drivers still have NO roster, so a day with no punch and no
// Day Off tap lands on "No Roster" — not Absent — and is not deducted.
// Change this one date if the switchover moves; nothing else.
const DRIVER_AUTO_PRESENT_UNTIL = '23-09-2026';

function driverAutoPresentOn_(dateStr) {
  const d = adParseDate_(dateStr);
  const last = adParseDate_(DRIVER_AUTO_PRESENT_UNTIL);
  if (!d || !last) return true;   // unparseable → behave as before
  return d <= last;
}

// Stop a rebuild before Apps Script kills it. Whatever is left is
// reported, not lost — each date is independent and re-running is safe.
const AD_REBUILD_BUDGET_MS = 280000;   // ~4m40s of a 6m limit

const AD_HEADERS = [
  'Employee ID', 'Name', 'Date', 'Location', 'Shift',
  'Check In', 'Break Out', 'Break In', 'Check Out',
  'Presence (hrs)', 'Late In (min)', 'Early Out (min)', 'Break Over (min)',
  'OT Hours', 'Day Type', 'Source', 'Status', 'Flags'
];

function rosterExistedOn_(dateStr) {
  const d = adParseDate_(dateStr);
  const go = adParseDate_(ROSTER_GO_LIVE);
  if (!d || !go) return true;   // unparseable → behave as before
  return d >= go;
}

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
// Rebuilds a ROLLING WINDOW, not just yesterday. Anything that went stale
// during the day — a backdated leave grant, a corrected punch, a roster
// published late — self-heals overnight even if nothing called
// adTouchRange_ at the time. With the range rebuild this costs about what
// one date used to.
const AD_NIGHTLY_WINDOW_DAYS = 14;

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
  const to = adDateToStr_(yesterday);

  const windowStart = new Date(yesterday.getTime() - (AD_NIGHTLY_WINDOW_DAYS - 1) * 86400000);
  const from = adDateToStr_(windowStart);

  Logger.log('dailyAttendanceTrigger: rebuilding ' + from + ' → ' + to);
  const res = rebuildAttendanceDaily_(from, to);
  Logger.log('dailyAttendanceTrigger: ' + res.message);
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
// PUBLIC: rebuild one day. Same signature and same result as v4 —
// everything that already calls this keeps working.
// ══════════════════════════════════════════════════════════════════════
function writeDayToAttendanceDaily_(dateStr) {
  const res = rebuildAttendanceDaily_(dateStr, dateStr);
  if (!res.success) throw new Error(res.error);
  return res;
}

// ══════════════════════════════════════════════════════════════════════
// PUBLIC: rebuild a DATE RANGE in one pass.
// Inputs read once, rows computed in memory, sheet rewritten once.
// Safe to re-run: the range's existing rows are replaced, not duplicated.
// ══════════════════════════════════════════════════════════════════════
function rebuildAttendanceDaily_(fromStr, toStr, opts) {
  opts = opts || {};
  const t0 = Date.now();
  const budget = opts.budgetMs || AD_REBUILD_BUDGET_MS;

  const from = adNormDd_(fromStr);
  const to   = adNormDd_(toStr || fromStr);
  const fromObj = adParseDate_(from), toObj = adParseDate_(to);
  if (!fromObj) return { success:false, error:'rebuildAttendanceDaily_: bad from date ' + fromStr };
  if (!toObj)   return { success:false, error:'rebuildAttendanceDaily_: bad to date ' + toStr };
  if (toObj < fromObj) return { success:false, error:'rebuildAttendanceDaily_: to is before from' };

  // Every date in the range, ascending — the same order N sequential
  // per-date calls would have produced.
  const dates = [];
  for (let d = new Date(fromObj.getTime()); d <= toObj; d = new Date(d.getTime() + 86400000)) {
    dates.push(adDateToStr_(d));
  }

  const ss  = adSpreadsheet_();
  const ctx = adLoadContext_(ss);            // ← every input, read ONCE

  const built = [], skipped = [];
  let allRows = [];
  for (let i = 0; i < dates.length; i++) {
    if (Date.now() - t0 > budget) { skipped.push.apply(skipped, dates.slice(i)); break; }
    const dateStr = dates[i];
    const rows = adBuildDayRows_(ctx, dateStr);
    allRows = allRows.concat(rows);
    built.push(dateStr);
    Logger.log('writeDayToAttendanceDaily_: ' + dateStr + ' — ' + rows.length + ' rows written.');
  }

  // One rewrite for the whole range.
  adCommitRows_(ctx, built, allRows);

  const message = built.length + ' date(s) rebuilt, ' + allRows.length + ' rows' +
    (skipped.length ? ' — STOPPED at the time limit, ' + skipped.length +
                      ' date(s) not done (re-run from ' + skipped[0] + ')' : '') +
    ' in ' + Math.round((Date.now() - t0) / 1000) + 's';
  Logger.log('rebuildAttendanceDaily_: ' + from + ' → ' + to + ' — ' + message);

  return { success:true, from:from, to:to, dates:built, skipped:skipped,
           rows:allRows.length, seconds:Math.round((Date.now()-t0)/1000),
           message:message };
}

// ══════════════════════════════════════════════════════════════════════
// PUBLIC: "an input changed — make the sheet honest again."
// Never throws, obeys a time budget, and only rebuilds dates that are in
// the past (a future date has no rows yet; it gets built when it arrives).
// This is what leave grants, punch edits, roster publishes and manual
// entries should call. One line at each call site:
//
//     if (typeof adTouchRange_ === 'function') adTouchRange_(start, end);
//
// Whatever does not fit the budget is picked up by the nightly rolling
// rebuild, so nothing stays stale for more than a day.
// ══════════════════════════════════════════════════════════════════════
function adTouchRange_(fromStr, toStr, opts) {
  try {
    const today = adDateToStr_(new Date());
    const todayObj = adParseDate_(today);
    let fromObj = adParseDate_(adNormDd_(fromStr));
    let toObj   = adParseDate_(adNormDd_(toStr || fromStr));
    if (!fromObj || !toObj) return { success:false, error:'adTouchRange_: bad dates' };
    if (toObj > todayObj) toObj = todayObj;            // nothing to rebuild in the future
    if (toObj < fromObj) return { success:true, dates:[], rows:0,
                                  message:'nothing in the past to rebuild' };
    return rebuildAttendanceDaily_(adDateToStr_(fromObj), adDateToStr_(toObj), opts);
  } catch (err) {
    // Never let a refresh break the thing that triggered it. The nightly
    // rolling rebuild is the safety net.
    Logger.log('adTouchRange_ failed (non-fatal): ' + err);
    return { success:false, error: err.toString() };
  }
}

function adTouchDates_(dates, opts) {
  try {
    const list = (dates || []).map(adNormDd_).filter(Boolean).sort(
      (a,b) => adIsoKey_(a).localeCompare(adIsoKey_(b)));
    if (!list.length) return { success:true, dates:[], rows:0, message:'no dates' };
    return adTouchRange_(list[0], list[list.length-1], opts);
  } catch (err) {
    Logger.log('adTouchDates_ failed (non-fatal): ' + err);
    return { success:false, error: err.toString() };
  }
}

// ══════════════════════════════════════════════════════════════════════
// CONTEXT — every input read exactly once.
// v4 re-read all of this for each date. That was the real cost alongside
// the per-row deletes.
// ══════════════════════════════════════════════════════════════════════
function adLoadContext_(ss) {
  const adSheet  = ss.getSheetByName('Attendance_Daily');
  const hrSheet  = ss.getSheetByName('HR Maya');
  const shSheet  = ss.getSheetByName('Shifts');
  const logSheet = ss.getSheetByName('Attendance_Log');
  const lvSheet  = ss.getSheetByName('Leave_Grants');
  const manSheet = ss.getSheetByName('Manual_Attendance_Aug');

  if (!adSheet || !hrSheet || !logSheet) {
    throw new Error('Required sheet missing (Attendance_Daily / HR Maya / Attendance_Log)');
  }

  if (adSheet.getLastRow() === 0) {
    adSheet.getRange(1, 1, 1, AD_HEADERS.length).setValues([AD_HEADERS]);
    adSheet.setFrozenRows(1);
  }

  // ── HR Maya (employee master, not shift-roster) ─────────────────────
  const hrCols = adColMap_(hrSheet);
  const HR_ID  = hrCols['Employee ID']    ?? 0;
  const HR_NM  = hrCols['Name (English)'] ?? hrCols['Name'] ?? 1;
  const HR_LOC = hrCols['Location']       ?? hrCols['Branch'] ?? -1;
  const HR_SH  = hrCols['Shift']          ?? -1;
  const HR_ST  = hrCols['Status']         ?? -1;
  const HR_DOJ = hrCols['Date of Join']   ?? hrCols['Joining Date'] ?? -1;
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
      // A man who has not joined yet has no attendance to compile. Kept
      // as the raw cell; adParseDate_ reads it when the date is judged.
      joinedOn : HR_DOJ >= 0 ? adNormDd_(row[HR_DOJ]) : '',
    };
    employees.push(emp);
    empMap[id] = emp;
  });

  // ── Shifts ──────────────────────────────────────────────────────────
  // Start and End are read as minutes past midnight. They are what tells
  // us, on a day that is still running, whether a man is genuinely absent
  // or simply not due yet.
  const autoShifts = {};
  const shiftHours = {};
  const shiftStart = {};
  const shiftEnd   = {};
  if (shSheet && shSheet.getLastRow() > 1) {
    const sc    = adColMap_(shSheet);
    const SH_NM = sc['Shift'] ?? sc['Name'] ?? 0;
    const SH_AU = sc['Auto Present'] ?? sc['AutoPresent'] ?? -1;
    const SH_NH = sc['Normal Hours'] ?? sc['Normal Hrs']  ?? -1;
    const SH_ST = sc['Start'] ?? sc['Shift Start'] ?? -1;
    const SH_EN = sc['End']   ?? sc['Shift End']   ?? -1;
    shSheet.getRange(2,1,shSheet.getLastRow()-1,shSheet.getLastColumn()).getValues()
      .forEach(r => {
        const name = String(r[SH_NM]||'').trim(); if (!name) return;
        if (SH_AU>=0 && String(r[SH_AU]||'').toLowerCase()==='yes') autoShifts[name] = true;
        if (SH_NH>=0) { const h=parseFloat(r[SH_NH]); if(!isNaN(h)&&h>0) shiftHours[name]=h; }
        if (SH_ST>=0) { const t=adHHmmToMin_(r[SH_ST]); if(t!=null) shiftStart[name]=t; }
        if (SH_EN>=0) { const t=adHHmmToMin_(r[SH_EN]); if(t!=null) shiftEnd[name]=t; }
      });
  }

  // ── Approved leave grants (kept as ranges; matched per date) ────────
  const leaveGrants = [];
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
        leaveGrants.push({ empId:empId, type:type, start:start, end:end });
      });
  }

  // ── Manual_Attendance_Aug — the 22-31 Aug supervisor backfill ───────
  // THIS WAS ONCE MISSING ENTIRELY. The header comment claimed Manual_Aug
  // was priority 0, but v3 never read the sheet — only the now-retired
  // AttendanceDaily.gs did. Recompiling any August date therefore WIPED
  // ~680 employee-days of supervisor-entered attendance and replaced it
  // with a fallback status. Caught by compareToBaseline().
  //
  // A manual entry is the highest authority for its date: a human looked
  // at it and said what happened. Nothing overrides it.
  //
  // Indexed by date here. Later rows still overwrite earlier ones for the
  // same employee+date, exactly as the per-date scan did.
  const manualByDate = {};
  if (manSheet && manSheet.getLastRow() > 1) {
    const mc = adColMap_(manSheet);
    const M_ID     = mc['Employee ID'] ?? 0;
    const M_LOC    = mc['Location']    ?? 1;
    const M_SHIFT  = mc['Shift']       ?? 2;
    const M_DATE   = mc['Date']        ?? 3;
    const M_STATUS = mc['Status']      ?? 4;
    manSheet.getRange(2,1,manSheet.getLastRow()-1,manSheet.getLastColumn()).getValues()
      .forEach(row => {
        const empId  = String(row[M_ID] || '').trim();
        const date   = adNormDd_(row[M_DATE]);
        const status = String(row[M_STATUS] || '').trim();
        if (!empId || !date || !status) return;
        (manualByDate[date] = manualByDate[date] || {})[empId] = {
          location: String(row[M_LOC]   || '').trim(),
          shift:    String(row[M_SHIFT] || '').trim(),
          status:   status,
        };
      });
  }

  // ── Kiosk punches, parsed once, grouped by employee ─────────────────
  // NOT filtered here. The admission test is per date (window OR filed-for
  // -that-date) and is applied in adPunchesForDate_, so a range rebuild
  // sees exactly the punches v4 would have seen for each date and no more.
  const logCols = adColMap_(logSheet);
  const L_TS    = logCols['Timestamp']     ?? 0;
  const L_ID    = logCols['Employee ID']   ?? 1;
  const L_DATE  = logCols['Date']          ?? 3;
  const L_STAGE = logCols['Stage']         ?? 4;
  const L_BRANCH= logCols['Branch']        ?? logCols['Location'] ?? -1;
  const L_FLAG  = logCols['Location Flag'] ?? logCols['Geo Flag'] ?? -1;

  const punchesByEmp = {};
  // A punch whose time has not come yet has not happened. The only way one
  // gets into the log is a check-out written at a shift end in advance —
  // supervisors verifying 20-09 while the shifts were still running wrote
  // 23 of them — and counting it would close a session that is still open.
  const NOW = new Date();
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2,1,logSheet.getLastRow()-1,logSheet.getLastColumn()).getValues()
      .forEach(row => {
        const ts = adParseTs_(row[L_TS]);
        if (!ts || ts > NOW) return;
        const empId = String(row[L_ID]||'').trim();
        const stage = String(row[L_STAGE]||'').trim();
        if (!empId || !stage) return;
        let location = '';
        if (L_BRANCH>=0) location = String(row[L_BRANCH]||'').trim();
        if (!location && L_FLAG>=0) location = adLocFromFlag_(String(row[L_FLAG]||''));
        if (!location && empMap[empId]) location = empMap[empId].location;
        if (!punchesByEmp[empId]) punchesByEmp[empId] = [];
        // dateCol is kept because Day Off is matched by it, not by the
        // timestamp — see the Day Off branch in adBuildDayRows_.
        punchesByEmp[empId].push({ ts:ts, stage:stage, location:location,
                                   dateCol: adNormDd_(row[L_DATE]) });
      });
    Object.keys(punchesByEmp).forEach(id => punchesByEmp[id].sort((a,b) => a.ts - b.ts));
  }

  // Shift starts per man, for the repeated-check-in rule (AttendanceCore).
  // Without AttendanceCore every run simply starts at its first tap.
  let startsFor = null;
  try { if (typeof acShiftStartsIndex_ === 'function') startsFor = acShiftStartsIndex_(ss); }
  catch (e) { Logger.log('adLoadContext_: shift starts unavailable — ' + e.message); }

  // Roster_Discrepancies already on file, so a recompile does not append
  // the same discrepancy again. (It did: 1,589 rows, 242 distinct — one
  // man-day logged up to 16 times, once per recompile.)
  const discSeen = {};
  const discSheet = ss.getSheetByName('Roster_Discrepancies');
  if (discSheet && discSheet.getLastRow() > 1) {
    discSheet.getRange(2, 1, discSheet.getLastRow() - 1, Math.min(6, discSheet.getLastColumn())).getValues()
      .forEach(r => { discSeen[[String(r[0]).trim(), adNormDd_(r[3]), String(r[4]).trim(), String(r[5]).trim()].join('|')] = true; });
  }

  return {
    ss: ss, adSheet: adSheet, discSeen: discSeen,
    employees: employees, empMap: empMap,
    autoShifts: autoShifts, shiftHours: shiftHours,
    shiftStart: shiftStart, shiftEnd: shiftEnd,
    leaveGrants: leaveGrants, manualByDate: manualByDate,
    punchesByEmp: punchesByEmp,
    startsFor: startsFor,
    rosterCache: {},
  };
}

// Log a roster discrepancy only if it is not already on file.
function adLogDiscrepancyOnce_(ctx, empId, name, location, dateStr, rosterSaid, actual) {
  if (typeof rbLogDiscrepancy_ !== 'function') return;
  const k = [empId, dateStr, String(rosterSaid).trim(), String(actual).trim()].join('|');
  if (ctx.discSeen && ctx.discSeen[k]) return;
  if (ctx.discSeen) ctx.discSeen[k] = true;
  rbLogDiscrepancy_(empId, name, location, dateStr, rosterSaid, actual);
}

// Published roster for a date, fetched at most once per rebuild.
function adRosterForDate_(ctx, dateStr) {
  if (ctx.rosterCache[dateStr] === undefined) {
    let m = {};
    if (typeof rbGetPublishedRosterForDate_ === 'function') {
      m = rbGetPublishedRosterForDate_(dateStr) || {};
    }
    ctx.rosterCache[dateStr] = m;
  }
  return ctx.rosterCache[dateStr];
}

function adLeaveForDate_(ctx, dateStr) {
  const out = {};
  ctx.leaveGrants.forEach(g => {
    if (adDateInRange_(dateStr, g.start, g.end)) out[g.empId] = g.type;
  });
  return out;
}

// The punches v4 would have admitted for this one date. Kept per-date on
// purpose: admitting the union of a range would let a backdated punch
// filed for date X take part in sessionizing date Y, which v4 never did.
function adPunchesForDate_(ctx, empId, dateStr, winStart, winEnd) {
  const all = ctx.punchesByEmp[empId];
  if (!all || !all.length) return [];
  const out = [];
  for (let i = 0; i < all.length; i++) {
    const p = all[i];
    // Admit a punch if EITHER:
    //   (a) its timestamp falls in the window — normal punches, and
    //       the overnight case where a Check Out lands the next day, OR
    //   (b) its Date column says it is FOR the date being built — a
    //       BACKDATED Day Off, whose timestamp is the submission time
    //       and can be days later.
    // Without (b), "Log a previous Day Off" was invisible to the
    // compile: MT-00003 filed a Day Off for 02-09 on 08-09 at 14:23,
    // the timestamp sat outside the 01-09→04-09 window, and 02-09
    // silently lost its Day Off.
    const inWindow      = (p.ts >= winStart && p.ts <= winEnd);
    const filedForToday = (p.dateCol === dateStr);
    if (!inWindow && !filedForToday) continue;
    out.push(p);
  }
  return out;   // already in timestamp order
}

// ══════════════════════════════════════════════════════════════════════
// CORE: build one day's rows. Every branch is v4's, unchanged.
// ══════════════════════════════════════════════════════════════════════
function adBuildDayRows_(ctx, dateStr) {
  const targetDate = adParseDate_(dateStr);
  if (!targetDate) throw new Error('adBuildDayRows_: bad date ' + dateStr);

  const leaveOnDate  = adLeaveForDate_(ctx, dateStr);
  const manualLookup = ctx.manualByDate[dateStr] || {};
  const rosterMap    = adRosterForDate_(ctx, dateStr);

  // OVERNIGHT-SHIFT FIX (v4). Punches are collected across a WINDOW (the
  // day before through the day after), the whole window is sessionized,
  // and only the session whose CHECK-IN falls on the date being built is
  // kept. A session belongs to the day it STARTED, which is also how the
  // shift itself is defined. Filtering by the date column instead — as
  // v3 did — silently destroyed every overnight session, and every second
  // shift at every location is overnight.
  const winStart = new Date(targetDate.getTime() - 86400000);
  const winEnd   = new Date(targetDate.getTime() + 2 * 86400000);

  const weekend = adIsWeekend_(targetDate);
  const MAX_HRS = 20;
  const rows    = [];

  // Is this the day we are standing in? A day still running cannot be
  // judged the way a finished day can: a man with no punch may simply
  // not be due yet, and a man with no check-out is just at work.
  const adNow    = new Date();
  const isToday  = (dateStr === adDateToStr_(adNow));
  const nowMin   = adNow.getHours() * 60 + adNow.getMinutes();

  ctx.employees.forEach(emp => {
    // Not joined yet on this date — he does not belong in attendance at
    // all. A blank joining date means we do not know, so he is included.
    if (!adJoinedBy_(emp.joinedOn, targetDate)) return;

    const isAuto     = !!ctx.autoShifts[emp.shift];
    const normHrs    = ctx.shiftHours[emp.shift] || 9;
    const leaveType  = leaveOnDate[emp.id];
    const rosterInfo = rosterMap[emp.id]; // { status, shift, location } | undefined
    const punches    = adPunchesForDate_(ctx, emp.id, dateStr, winStart, winEnd);

    // Sessionize the WHOLE window, then keep only what belongs to this date.
    // A session belongs to the day its CHECK-IN falls on; a Day Off belongs
    // to the day it was punched.
    // REPEATED CHECK-IN TAPS (21-09-2026). A Check In while a session is
    // open no longer restarts the day: taps less than 16h apart are one run,
    // and the session starts at the first tap that lines up with a shift.
    // See AttendanceCore §0 for the rule and the real cases behind it.
    const locs = {}; punches.forEach(p => { if (p.location) locs[p.location] = true; });
    const starts = ctx.startsFor ? ctx.startsFor(emp.id, Object.keys(locs)) : [];
    const repeat = (a, b) => (typeof acIsRepeatCheckIn_ === 'function')
                               ? acIsRepeatCheckIn_(a, b) : false;
    const pickStart = (run) => (typeof acPickRunStart_ === 'function')
                               ? (acPickRunStart_(run, starts) || run[0]) : run[run.length - 1];
    const MIN_SESSION_MIN = (typeof AC_MIN_SESSION_MIN === 'number') ? AC_MIN_SESSION_MIN : 0;

    let session = null;        // what this date resolves to
    let open    = null;
    const done  = [];          // completed sessions that STARTED on this date
    let dayOffHere = null;
    // A check-in on this date that never got a check-out. Kept, so the day
    // shows he came in instead of "no punch recorded" (21-09-2026: Nilam,
    // Nishan, Ram Bahadur, Dhirendra all checked in at Kaifan on 20-09 and
    // came out Absent).
    let openHere = null;
    const keepIfOpenHere = (o) => {
      if (!o) return;
      const st = pickStart(o.run);
      if (adDateToStr_(st.ts) === dateStr && !openHere) openHere = { run: o.run, start: st, breaks: o.breaks,
                                                                   location: st.location || o.location };
    };
    punches.forEach(p => {
      if (p.stage === 'Check In') {
        if (open && repeat(open.run[0].ts, p.ts)) { open.run.push(p); return; }
        keepIfOpenHere(open);        // a new day began with this one still open
        open = { location: p.location || emp.location, run: [p], breaks: [], checkOut: null };
      } else if (p.stage === 'Break Out' && open) {
        open.breaks.push({ out: p.ts, in: null });
      } else if (p.stage === 'Break In' && open && open.breaks.length > 0) {
        const last = open.breaks[open.breaks.length-1]; if (!last.in) last.in = p.ts;
      } else if (p.stage === 'Check Out' && open) {
        const st = pickStart(open.run);
        open.checkIn = st.ts;
        if (st.location) open.location = st.location;
        open.repeatTaps = open.run.length - 1;
        open.breaks = open.breaks.filter(b => b.out >= st.ts);
        open.checkOut = p.ts;
        // Only claim this session if it STARTED on the date being built,
        // and it is longer than a fumble at the kiosk.
        if (open.checkOut > open.checkIn && adDateToStr_(open.checkIn) === dateStr &&
            (open.checkOut - open.checkIn) / 60000 >= MIN_SESSION_MIN) done.push(open);
        open = null;
      } else if (p.stage === 'Day Off') {
        // DAY OFF IS MATCHED BY ITS DATE COLUMN, NOT ITS TIMESTAMP.
        // The kiosk's "Log a previous Day Off" writes the submission time
        // as the timestamp and the day it APPLIES TO as the Date column.
        // Real example — MT-00003 row 1085:
        //     timestamp 08/09 14:23   date-col 02-09
        // He checked out at 14:12 on the 8th and then backdated a Day Off
        // for the 2nd. Matching on the timestamp attributed it to the 8th,
        // conflicting with a day he actually worked AND losing the Day Off
        // on the 2nd. Sessions match on timestamp (overnight shifts cross
        // midnight); Day Off matches on the date it was filed for.
        //
        // A COMPLETED SESSION STILL BEATS A DAY OFF for the same date: if
        // real work is recorded, the Day Off tap is the mistake, not the work.
        if (p.dateCol === dateStr) dayOffHere = { dayOff: true, location: p.location || emp.location };
        // A Day Off does NOT close an open session here — an overnight
        // shift may still be running when an unrelated backdated Day Off
        // is filed. A Day Off for THIS date ends an open session left over
        // from an earlier day — but NOT one that began today. Ram Bahadur
        // Sirmal, 20-09: checked in 11:32, tapped Day Off at 17:46, worked
        // to the end. The tap used to kill the shift and he came out Absent.
        if (p.dateCol === dateStr && open && adDateToStr_(open.run[0].ts) !== dateStr) open = null;
      }
    });
    keepIfOpenHere(open);

    // One session: that is the day. Several: a split shift — the hours are
    // added up, the day runs from the first check-in to the last check-out,
    // and it is flagged so nobody reads the span as the presence.
    if (done.length === 1) {
      session = done[0];
    } else if (done.length > 1) {
      const first = done[0], last = done[done.length - 1];
      session = {
        location: first.location, checkIn: first.checkIn, checkOut: last.checkOut,
        breaks: first.breaks, repeatTaps: done.reduce((n, s) => n + (s.repeatTaps || 0), 0),
        splitHours: done.reduce((h, s) => h + adHoursDiff_(s.checkIn, s.checkOut), 0),
        splitCount: done.length,
      };
    } else if (openHere) {
      // Checked in, never checked out. He was here — that beats a Day Off
      // tap and the roster's "no punch". Hours are not known until the
      // supervisor answers the day (supervisor page), so no presence, no OT.
      session = { openOnly: true, location: openHere.location, checkIn: openHere.start.ts,
                  breaks: openHere.breaks, repeatTaps: openHere.run.length - 1,
                  dayOffTap: !!dayOffHere };
    } else if (dayOffHere) {
      // A COMPLETED SESSION STILL BEATS A DAY OFF for the same date.
      session = dayOffHere;
    }

    let row;
    const manual = manualLookup[emp.id];

    // ── DRIVERS ARE ALWAYS AUTO-PRESENT ──────────────────────────────
    // Drivers now have kiosk access, but their punches must NOT drive
    // presence. Two reasons:
    //   1. They are on the road; a punch records where they happened to
    //      be, not a shift being worked. Driver Location has no shift
    //      definition and no start/end times to measure against.
    //   2. Punched badly, they produce WORSE results than not punching.
    //      Seen live on 14-09: MT-00049 punched Break In BEFORE Break Out
    //      (02:33 in, 19:13 out) and the sessionizer read the inverted
    //      pairs as ~10 hours of break — presence came out 6.56h across a
    //      16h40m span. Meanwhile MT-00064, who did not punch at all, got
    //      a clean Driver_Default Present.
    //
    // So: Driver Location resolves to Present regardless of punches. The
    // punches still sit in Attendance_Log for reference — nothing is
    // deleted — they simply do not compute presence or OT.
    if (emp.location === 'Driver Location' && driverAutoPresentOn_(dateStr)) {
      // Auto-present does NOT override things that genuinely mean the
      // driver was not working. Approved leave and a self-reported Day
      // Off both outrank it — otherwise a driver on leave would be
      // written as Present and the unpaid-leave deduction would never
      // fire in payroll.
      if (leaveType) {
        row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
          '','','','','','','','',0,'On Leave','Granted_Leave','On Leave', leaveType];
      } else if (session && session.dayOff) {
        row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
          '','','','','','','','',0,'Day Off','Kiosk','Day Off',
          'Driver self-reported Day Off'];
      } else {
        row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
          '','','','','','','','',0,'Full Day','Driver_Default','Present',
          'Driver — auto-present; kiosk punches not used for presence'];
      }
      rows.push(row);
      return;
    }

    // ── 0: Manual supervisor entry — highest authority, nothing overrides.
    // A human looked at this day and stated what happened.
    //
    // THERE ARE NO PUNCH STAGES ON THESE DAYS AND THERE NEVER WILL BE.
    // The supervisor posted Present / Absent / Day Off and that IS the
    // record — 22-31 Aug by backfill, and from Sept onward every verdict
    // he gives on a day nobody punched. So this branch must produce a
    // COMPLETE day on its own. It cannot leave the hours blank and wait
    // for punches to fill them in, because nothing is coming.
    //
    // Presence for a stated Present is the shift itself. This is exactly
    // what Auto_HeadOffice (branch 5) does for the identical shape — a
    // man who is present and does not punch — and leaving it blank here
    // while filling it there was the inconsistency. A stated Present with
    // no hours is a day that looks worked and pays nothing.
    //
    // OT stays 0. A day with no punches cannot evidence overtime; if he
    // worked over, that goes through an OT claim a supervisor verifies.
    if (manual) {
      const mStatus  = manual.status;
      const mDayType = mStatus === 'Present' ? 'Full Day'
                     : mStatus === 'Absent'  ? 'Absent'
                     : 'Day Off';
      const mPresence = mStatus === 'Present' ? normHrs : '';
      row = [emp.id, emp.name, dateStr,
        manual.location || emp.location,
        manual.shift    || emp.shift,
        '','','','', mPresence, '','','',0, mDayType, 'Manual_Aug', mStatus, ''];

    // ── 1: Kiosk Day Off self-report — check against roster ───────────
    } else if (session && session.dayOff) {
      if (rosterInfo && rosterInfo.status === 'Working') {
        // Roster says they were scheduled — self-report doesn't override.
        row = [emp.id, emp.name, dateStr, emp.location, rosterInfo.shift || emp.shift,
          '','','','','','','','',0,'Absent','Roster_Absent','Absent',
          'Self-reported Day Off via kiosk, but roster scheduled Working'];
        if (true) {
          adLogDiscrepancyOnce_(ctx, emp.id, emp.name, rosterInfo.location || emp.location, dateStr,
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
      // ── BREAK IS NOT DEDUCTED FROM PRESENCE ──────────────────────
      // POLICY (Shakeel, 20-09-2026): a 9-hour shift means 9 hours ON SITE,
      // already including the 60-minute break. So the break must not come
      // off a second time. A man on site 12:39 → 03:46 has done 15h 07m,
      // and 6h 07m of that is overtime.
      //
      // Deducting it again was both wrong and perverse. Two men on the same
      // shift at Jabriya on 03-09, check-ins one minute apart:
      //   Kristal Ghalan  12:39 → 03:46, punched a 50m break → 5h 17m OT
      //   Yogendra Oli    12:38 → 03:46, punched no break    → 6h 08m OT
      // Punching your break cost you 51 minutes of overtime. Not punching
      // it cost you nothing — so before long nobody punches a break at all
      // and the break data is worthless.
      //
      // breakHrs is still computed, because the Break Out / Break In times
      // are still WRITTEN to the sheet. They are a record of when he took
      // his break, not a deduction from his pay.
      // A split shift is the SUM of its sessions, not the span across them.
      const rawHrs      = (session.splitHours != null) ? session.splitHours
                                                       : adHoursDiff_(session.checkIn, session.checkOut);
      const presenceHrs = rawHrs;
      const isSuspect   = presenceHrs > MAX_HRS;
      const presenceFin = isSuspect ? 0 : presenceHrs;
      const otHrs       = isSuspect ? 0 : Math.max(0, presenceHrs - normHrs);
      const dayType     = isSuspect ? 'Check Punches'
                        : presenceHrs >= normHrs*0.5 ? 'Full Day' : 'Half Day';
      let flags         = isSuspect ? 'REVIEW: ' + rawHrs.toFixed(1) + 'h — check punches' : '';
      if (session.splitCount > 1)
        flags = (flags ? flags + '; ' : '') + 'Split shift: ' + session.splitCount + ' sessions added up';
      if (session.repeatTaps > 0)
        flags = (flags ? flags + '; ' : '') + 'Repeated check-in (' + session.repeatTaps +
                ' extra tap' + (session.repeatTaps > 1 ? 's' : '') + ') — started at ' + ciStr;

      // Roster mismatch check (informational — punch still wins)
      if (rosterInfo) {
        if (rosterInfo.status === 'Day Off') {
          if (true) {
            adLogDiscrepancyOnce_(ctx, emp.id, emp.name, rosterInfo.location || emp.location, dateStr,
              'Day Off', 'Present ' + ciStr + '–' + coStr);
          }
          flags = (flags ? flags + '; ' : '') + 'Roster said Day Off — punch overrides';
        } else if (rosterInfo.shift && rosterInfo.shift !== emp.shift) {
          if (true) {
            adLogDiscrepancyOnce_(ctx, emp.id, emp.name, rosterInfo.location || emp.location, dateStr,
              'Working (' + rosterInfo.shift + ')', 'Present, shift on file: ' + emp.shift);
          }
        }
      }

      // Shift for display/OT: prefer the employee's own, fall back to what
      // the roster said for this date. A blank shift silently drops OT to
      // the 9h default, so an empty value here is not harmless.
      const shiftUsed = emp.shift || (rosterInfo && rosterInfo.shift) || '';
      row = [emp.id, emp.name, dateStr, loc, shiftUsed,
        ciStr, boStr, biStr, coStr,
        +presenceFin.toFixed(3), '','','', +otHrs.toFixed(3),
        dayType, 'Kiosk', 'Present', flags];

    // ── 2b: Checked in, no check-out — present, hours to be answered ──
    // Status Present so he is not counted absent; Day Type 'Check Punches'
    // so the supervisor page asks for his check-out; presence and OT stay
    // empty until that answer is in. Nothing is invented.
    } else if (session && session.openOnly) {
      const ciStr = adHHmm_(session.checkIn);
      // On today, an open session is a man at work, not a question for
      // the supervisor. The ask comes back on the nightly rebuild if he
      // really never checked out.
      let flags = isToday
        ? ('Checked in ' + ciStr + ' — on shift now')
        : ('Checked in ' + ciStr + ', no check-out — supervisor to answer');
      if (session.repeatTaps > 0)
        flags += '; Repeated check-in (' + session.repeatTaps + ' extra tap' +
                 (session.repeatTaps > 1 ? 's' : '') + ') — started at ' + ciStr;
      if (session.dayOffTap) flags += '; Day Off tapped after checking in — ignored';
      const shiftUsed = (rosterInfo && rosterInfo.shift) || emp.shift || '';
      row = [emp.id, emp.name, dateStr, session.location || emp.location, shiftUsed,
        ciStr, '', '', '', '', '', '', '', 0, (isToday ? 'Working' : 'Check Punches'), 'Kiosk', 'Present', flags];

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
      const shName = rosterInfo.shift || emp.shift || '';
      const stMin  = ctx.shiftStart[shName];
      const enMin  = ctx.shiftEnd[shName];
      // Absent is a conclusion, and on today it can only be drawn once
      // the shift has actually finished. Before that he is not due, or
      // he is late, or he is on his way — none of which is absence.
      let stillDue = false;
      if (isToday) {
        if (stMin == null || enMin == null) stillDue = true;   // no times on file — never accuse
        else if (enMin <= stMin)            stillDue = true;   // overnight shift: it ends tomorrow
        else if (nowMin < enMin)            stillDue = true;   // the shift is still running
      }
      if (stillDue) {
        const when = (stMin == null)
              ? 'start time not set on the Shifts tab'
              : (nowMin < stMin
                   ? 'shift starts ' + adMinToHHmm_(stMin) + ' — not due yet'
                   : 'shift started ' + adMinToHHmm_(stMin) + ' — no punch yet, day still running');
        row = [emp.id, emp.name, dateStr, emp.location, shName,
          '','','','','','','','',0,'Not Started','Roster_NotStarted','Not Started',
          'Rostered to work; ' + when];
      } else {
        row = [emp.id, emp.name, dateStr, emp.location, shName,
          '','','','','','','','',0,'Absent','Roster_Absent','Absent',
          'Rostered to work — no punch recorded'];
      }

    } else if (rosterInfo && rosterInfo.status === 'Day Off') {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'Day Off','Roster_DayOff','Day Off',''];

    // ── 8: (Driver Location up to DRIVER_AUTO_PRESENT_UNTIL is handled at
    // the TOP of this loop. From 24-09-2026 drivers fall through here like
    // everyone else: punches make the day, a Day Off tap makes a Day Off,
    // and a day with neither lands on "No Roster" below — drivers have no
    // roster, and that status deducts nothing.)

    // ── 9: Nothing matched. What this MEANS depends on whether the
    // roster feature existed yet — see ROSTER_GO_LIVE above.
    } else if (rosterExistedOn_(dateStr)) {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'No Roster','No_Roster','No Roster',
        'No roster published for this date'];
    } else {
      row = [emp.id, emp.name, dateStr, emp.location, emp.shift,
        '','','','','','','','',0,'Not Recorded','Pre_Roster','Not Recorded',
        'Before the roster feature existed — no punch, no day off, nothing on record'];
    }

    rows.push(row);
  });

  return rows;
}

// ══════════════════════════════════════════════════════════════════════
// COMMIT — the part that used to cost 24 seconds a date.
//
// v4: delete the date's rows ONE AT A TIME (~0.35s each on this sheet),
//     then append. 68 rows per date = ~24s of pure sheet API calls.
// v5: read the body once, drop the rebuilt dates in memory, write
//     everything back with ONE setValues and at most ONE deleteRows for
//     whatever slack is left at the bottom.
//
// Net effect is identical to running v4 date by date in ascending order:
// rows for untouched dates keep their order and position, rebuilt dates
// end up after them in date order.
// ══════════════════════════════════════════════════════════════════════
function adCommitRows_(ctx, rebuiltDates, newRows) {
  const sheet = ctx.adSheet;
  const target = {};
  rebuiltDates.forEach(d => target[d] = true);

  const lastRow = sheet.getLastRow();
  const width   = Math.max(sheet.getLastColumn(), AD_HEADERS.length);

  let kept = [];
  if (lastRow > 1) {
    const dateCol = adColMap_(sheet)['Date'] ?? 2;
    const body = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    for (let i = 0; i < body.length; i++) {
      if (!target[adNormDd_(body[i][dateCol])]) kept.push(body[i]);
    }
  }

  // Pad the new rows out to the sheet's real width so setValues never
  // throws on a sheet that has grown extra columns.
  const padded = newRows.map(r => {
    if (r.length === width) return r;
    const out = r.slice(0, width);
    while (out.length < width) out.push('');
    return out;
  });

  const out = kept.concat(padded);
  if (out.length) sheet.getRange(2, 1, out.length, width).setValues(out);

  const oldBody = Math.max(0, lastRow - 1);
  const slack   = oldBody - out.length;
  if (slack > 0) sheet.deleteRows(out.length + 2, slack);   // ONE call, not a loop
}

// ── Helpers ───────────────────────────────────────────────────────────

// The spreadsheet to work on.
//
// The compile has always used getActiveSpreadsheet(), which is correct for
// a container-bound script — including from a trigger or a web app. But the
// rest of the project opens MAIN_SPREADSHEET_ID by hand, and a rebuild is
// now called from BOTH worlds: the 4 AM trigger, the editor, grantLeave,
// and the supervisor page's doGet. If the active spreadsheet is ever
// unavailable, fall back to the ID rather than fail in one context and work
// in another — a rebuild that silently does nothing is worse than a slow one.
function adSpreadsheet_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss && typeof MAIN_SPREADSHEET_ID === 'string' && MAIN_SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  }
  if (!ss) throw new Error('adSpreadsheet_: no spreadsheet available');
  return ss;
}

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
// '13:30', a Date from a time cell, or a fraction of a day -> minutes past
// midnight. Anything it cannot read comes back null, never 0 — a 0 would
// read as midnight and make every shift look long finished.
function adHHmmToMin_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) { if (isNaN(v.getTime())) return null; return v.getHours()*60 + v.getMinutes(); }
  if (typeof v === 'number') {
    if (v > 0 && v < 1) return Math.round(v * 1440);           // Sheets time serial
    return null;
  }
  const m = String(v).trim().match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  const h = Number(m[1]), mi = Number(m[2]);
  if (h < 0 || h > 23 || mi < 0 || mi > 59) return null;
  return h*60 + mi;
}
// True when the employee had joined on or before the date being built.
// A blank or unreadable joining date returns true: we never drop a man
// from attendance because a cell was empty.
function adJoinedBy_(joinedOn, targetDate) {
  if (!joinedOn) return true;
  const j = adParseDate_(joinedOn);
  if (!j || isNaN(j.getTime())) return true;
  return j.getTime() <= targetDate.getTime();
}
function adMinToHHmm_(mins) {
  if (mins == null) return '';
  return adPad_(Math.floor(mins/60)) + ':' + adPad_(mins % 60);
}
function adDateToStr_(d) { return adPad_(d.getDate())+'-'+adPad_(d.getMonth()+1)+'-'+d.getFullYear(); }
function adIsoKey_(dateStr) { const p=String(dateStr).split('-'); return (p[2]||'')+(p[1]||'')+(p[0]||''); }
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

// ══════════════════════════════════════════════════════════════════════
// Run once after deploying: rebuilds a range with v5 and reports timing.
// Use compareToBaseline() around it for the real proof.
// ══════════════════════════════════════════════════════════════════════
function adRebuildSelfCheck() {
  const to = adDateToStr_(new Date());
  const from = adDateToStr_(new Date(Date.now() - 6 * 86400000));
  const res = rebuildAttendanceDaily_(from, to);
  Logger.log(res.message);
  return res.message;
}