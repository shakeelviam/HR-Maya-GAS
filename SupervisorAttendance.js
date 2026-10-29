/**
 * SupervisorAttendance.gs — backend for supervisor_attendance.html
 *
 * WHY THIS EXISTS
 *   The old page called three methods that cannot do the job:
 *     saveManualAttendance  — hardcoded August-only date guard, writes to
 *                             Manual_Attendance_Aug. Every submission between
 *                             18-10 and 20-10 silently failed.
 *     addAttendancePunch    — refuses a duplicate stage, so it cannot write the
 *                             second Check Out on a date that already has one.
 *     buildAttendanceDaily  — retired; returns { success:false, retired:true }.
 *   It also had no way to recompile Attendance_Daily, because nothing in the
 *   router rebuilds it.
 *
 * WHAT THE 20-10-2026 CLEAN-UP TAUGHT, encoded here
 *   1. "Absent" IS NOT A PUNCH. ATTN_STAGES has no Absent. Absence is expressed
 *      by the ROSTER saying Working while no punch arrives. So Absent is written
 *      to the roster, not the log.
 *   2. A Day Off punch LOSES to a roster row that says Working. Eight of that
 *      period's eleven absences were men who tapped Day Off and were marked
 *      Absent anyway. So Day Off is written to the ROSTER too.
 *   3. One Check Out per employee per date is NOT enough. A man who works a
 *      Second shift then a First shift the next day finishes twice on one
 *      calendar date. svSetPunch_ appends in that case instead of overwriting.
 *   4. A session belongs to the date it STARTED. An overnight finish carries the
 *      next day's date in the log but closes the day it began on.
 *
 * INSTALLING
 *   1. Add this file to the Apps Script project.
 *   2. Add these two cases to the router's switch, beside the other cases:
 *
 *        case 'svGetOpenDays':
 *          return ContentService.createTextOutput(JSON.stringify(
 *            svGetOpenDays((e.postData && e.postData.contents)
 *              ? JSON.parse(e.postData.contents) : e.parameter)))
 *            .setMimeType(ContentService.MimeType.JSON);
 *        case 'svSaveAnswers':
 *          return ContentService.createTextOutput(JSON.stringify(
 *            svSaveAnswers((e.postData && e.postData.contents)
 *              ? JSON.parse(e.postData.contents) : e.parameter)))
 *            .setMimeType(ContentService.MimeType.JSON);
 *
 *   3. Deploy a NEW VERSION of the web app. Without that the page sees nothing.
 *
 * NOT TESTED AGAINST THE LIVE SHEET. Before sending the link to anyone, sign in
 * as one supervisor, answer ONE day, save, and check Attendance_Daily yourself.
 */

// Days either side of the window that may be touched when a shift runs overnight.
const SV_EDGE_DAYS = 1;

// ────────────────────────────────────────────────────────────────────────────
// READ — the days at this supervisor's branch that still need an answer.
// ────────────────────────────────────────────────────────────────────────────
function svGetOpenDays(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const auth = vsAuthSupervisor_(params.supId, params.supPin);
    if (!auth.ok) return { success: false, error: auth.error };

    const from = parseFlexibleDate_(String(params.from || '').trim());
    const to   = parseFlexibleDate_(String(params.to   || '').trim());
    if (!svValidDate_(from) || !svValidDate_(to)) {
      return { success: false, error: 'from and to must be valid dates (dd-mm-yyyy).' };
    }
    if (to < from) return { success: false, error: 'The end date is before the start date.' };

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = ss.getSheetByName('Attendance_Daily');
    if (!sh) return { success: false, error: 'Attendance_Daily not found.' };

    const data = sh.getDataRange().getValues();
    const H = data[0].map(function (h) { return String(h).trim(); });
    const c = {
      id:    H.indexOf('Employee ID'),
      name:  H.indexOf('Name'),
      date:  H.indexOf('Date'),
      loc:   H.indexOf('Location'),
      shift: H.indexOf('Shift'),
      ci:    H.indexOf('Check In'),
      co:    H.indexOf('Check Out'),
      stat:  H.indexOf('Status'),
      src:   H.indexOf('Source')
    };
    for (var k in c) {
      if (c[k] < 0) return { success: false, error: 'Attendance_Daily is missing the "' + k + '" column.' };
    }

    // Today is never a question: a man may still be on shift.
    const todayStr = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    const out = [];

    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const d = parseFlexibleDate_(row[c.date]);
      if (!svValidDate_(d)) continue;
      if (d < from || d > to) continue;
      if (Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') === todayStr) continue;
      if (String(row[c.loc] || '').trim() !== auth.location) continue;

      const status = String(row[c.stat] || '').trim();
      const ci     = svHHmm_(row[c.ci]);
      const co     = svHHmm_(row[c.co]);
      var why = '';

      if (status === 'No Roster' && !ci && !co) {
        why = 'No roster was published for this day and he never punched. '
            + 'Nothing is deducted, so right now this day is paid with nothing behind it.';
      } else if (status === 'Absent') {
        why = 'Recorded absent: he was rostered to work and no punch arrived. '
            + 'If he was off, say so and the roster will be corrected.';
      } else if (status === 'Present' && ci && !co) {
        why = 'He started at ' + ci + ' and never finished. The hours cannot be '
            + 'worked out until someone says what time he left.';
      } else if (status === 'Not Started') {
        why = 'Rostered to work, but the day finished with no punch at all.';
      } else {
        continue;
      }

      out.push({
        id:     String(row[c.id] || '').trim(),
        name:   String(row[c.name] || '').trim(),
        date:   formatDdMmYyyy_(d),
        loc:    String(row[c.loc] || '').trim(),
        shift:  String(row[c.shift] || '').trim(),
        status: status,
        ci:     ci,
        co:     co,
        why:    why
      });
    }

    out.sort(function (a, b) {
      if (a.date === b.date) return a.name < b.name ? -1 : 1;
      return parseFlexibleDate_(a.date) - parseFlexibleDate_(b.date);
    });

    return {
      success: true,
      supervisor: auth.name,
      location: auth.location,
      from: formatDdMmYyyy_(from),
      to: formatDdMmYyyy_(to),
      count: out.length,
      days: out
    };
  } catch (err) {
    return { success: false, error: 'Server error: ' + err.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// WRITE — one call for the whole batch, then ONE rebuild.
//
// answers: [{ empId, date, verdict:'Present'|'Absent'|'Day Off',
//             ci:'HH:mm', co:'HH:mm', coNextDay:bool, shift:'' }]
// ────────────────────────────────────────────────────────────────────────────
function svSaveAnswers(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const auth = vsAuthSupervisor_(params.supId, params.supPin);
    if (!auth.ok) return { success: false, error: auth.error };

    const answers = params.answers || [];
    if (!answers.length) return { success: false, error: 'Nothing to save.' };

    const by = auth.name + ' (' + String(params.supId).trim() + ')';
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const results = [];
    var lo = null, hi = null;

    for (var i = 0; i < answers.length; i++) {
      const a = answers[i];
      const empId = String(a.empId || '').trim();
      const d = parseFlexibleDate_(String(a.date || '').trim());
      const verdict = String(a.verdict || '').trim();
      const key = empId + '|' + String(a.date || '');

      try {
        if (!empId || !svValidDate_(d)) throw new Error('Bad employee or date.');
        if (['Present', 'Absent', 'Day Off'].indexOf(verdict) < 0) throw new Error('Unknown verdict.');
        if (!vsTeamContains_(auth.location, empId)) throw new Error('Not on your team.');

        if (lo === null || d < lo) lo = d;
        if (hi === null || d > hi) hi = d;

        if (verdict === 'Present') {
          const ci = svNormTime_(a.ci), co = svNormTime_(a.co);
          if (!ci || !co) throw new Error('Both the start and the finish time are needed.');

          // The finish may land on the next calendar day. The session still
          // belongs to the date it STARTED.
          var coDate = d;
          if (a.coNextDay || co < ci) {
            coDate = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
            if (hi === null || coDate > hi) hi = coDate;
          }
          svSetPunch_(ss, empId, d,      'Check In',  ci, by, false);
          svSetPunch_(ss, empId, coDate, 'Check Out', co, by, true);
          svRosterSet_(ss, auth, empId, d, 'Working', String(a.shift || '').trim(),
                       'Supervisor confirmed he worked ' + ci + ' to ' + co + '.');
          results.push({ key: key, ok: true, msg: 'worked ' + ci + ' to ' + co });

        } else if (verdict === 'Day Off') {
          // Written to the ROSTER, not as a punch. A Day Off punch loses to a
          // roster row that says Working — that bug produced eight wrong
          // absences in the 21-09 to 20-10 period.
          svRosterSet_(ss, auth, empId, d, 'Day Off', '',
                       'Supervisor confirmed this was his day off.');
          results.push({ key: key, ok: true, msg: 'day off' });

        } else {
          // Absent is the roster saying Working with no punch. There is no
          // Absent punch stage, so this is the only way to express it.
          svRosterSet_(ss, auth, empId, d, 'Working', String(a.shift || '').trim(),
                       'Supervisor confirmed he did not turn up.');
          results.push({ key: key, ok: true, msg: 'absent' });
        }
      } catch (rowErr) {
        results.push({ key: key, ok: false, msg: rowErr.message });
      }
    }

    const saved = results.filter(function (r) { return r.ok; }).length;
    var rebuilt = null, rebuildError = null;

    if (saved && lo && hi) {
      // Widen by a day each side: an overnight finish touches the next date.
      const rFrom = new Date(lo.getFullYear(), lo.getMonth(), lo.getDate() - SV_EDGE_DAYS);
      const rTo   = new Date(hi.getFullYear(), hi.getMonth(), hi.getDate() + SV_EDGE_DAYS);
      try {
        const rb = rebuildAttendanceDaily_(formatDdMmYyyy_(rFrom), formatDdMmYyyy_(rTo));
        rebuilt = (rb && rb.message) ? rb.message
                : (formatDdMmYyyy_(rFrom) + ' to ' + formatDdMmYyyy_(rTo) + ' rebuilt');
      } catch (rbErr) {
        // The answers ARE written. Say so plainly so nobody re-enters them.
        rebuildError = rbErr.message;
      }
    }

    return {
      success: true,
      saved: saved,
      failed: results.length - saved,
      results: results,
      rebuilt: rebuilt,
      rebuildError: rebuildError
    };
  } catch (err) {
    return { success: false, error: 'Server error: ' + err.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────────────────

// An unparseable date comes back as an Invalid Date, which is TRUTHY, and every
// comparison against it is false. Test the time value, never just the object.
function svValidDate_(x) { return !!x && !isNaN(x.getTime()); }

function svHHmm_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz_(), 'HH:mm');
  }
  const s = String(v).trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (!m) return '';
  return ('0' + m[1]).slice(-2) + ':' + m[2];
}

function svNormTime_(v) {
  const s = svHHmm_(v);
  if (!s) return '';
  const p = s.split(':');
  const h = parseInt(p[0], 10), mi = parseInt(p[1], 10);
  if (isNaN(h) || isNaN(mi) || h < 0 || h > 23 || mi < 0 || mi > 59) return '';
  return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2);
}

/**
 * Write one punch.
 *
 * allowSecond=true is for a Check Out on a date that may already hold one,
 * because the man finished twice on that calendar date (a Second shift, then a
 * First shift the next day). amendAttendance cannot express this: it finds the
 * FIRST matching stage and overwrites it, which destroys the other day's close.
 * The log holds both rows happily and the compile reads them correctly — only
 * the old write path was too narrow.
 */
function svSetPunch_(ss, empId, dateObj, stage, hhmm, by, allowSecond) {
  const log = ss.getSheetByName('Attendance_Log');
  if (!log) throw new Error('Attendance_Log not found.');

  const data = log.getDataRange().getValues();
  const dateStr = formatDdMmYyyy_(dateObj);
  const parts = hhmm.split(':');
  const ts = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(),
                      parseInt(parts[0], 10), parseInt(parts[1], 10), 0);
  const tsStr = Utilities.formatDate(ts, tz_(), 'dd/MM/yyyy HH:mm:ss');

  var matches = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== empId) continue;
    if (String(data[i][4]).trim() !== stage) continue;
    const rd = parseFlexibleDate_(data[i][3]);
    if (!svValidDate_(rd) || !sameCalendarDate_(rd, dateObj)) continue;
    matches.push({ rowNum: i + 1, time: svHHmm_(data[i][0]) });
  }

  // Already exactly right — do nothing. This is what makes re-submitting safe.
  for (var j = 0; j < matches.length; j++) {
    if (matches[j].time === hhmm) return 'already set';
  }

  var empName = 'Unknown';
  if (typeof vsEmpName_ === 'function') {
    try { empName = vsEmpName_(empId) || empName; } catch (e) {}
  }

  if (matches.length === 0 || (allowSecond && matches.length >= 1)) {
    log.appendRow([tsStr, empId, empName, dateStr, stage, 'ADMIN']);
    svAudit_(ss, by, empId, empName, dateStr, stage, '(none)', tsStr);
    return 'added';
  }

  log.getRange(matches[0].rowNum, 1).setValue(tsStr);
  log.getRange(matches[0].rowNum, 4).setValue(dateStr);
  log.getRange(matches[0].rowNum, 6).setValue('ADMIN');
  svAudit_(ss, by, empId, empName, dateStr, stage, matches[0].time, tsStr);
  return 'updated';
}

/**
 * Set a roster row's Planned Status, creating the row if there isn't one.
 *
 * This is the lever that actually decides Absent vs Day Off. The compile reads
 * only PUBLISHED rows, so anything written here is published immediately —
 * a Draft row is invisible and the day falls through to "No Roster", which is
 * not deducted.
 */
function svRosterSet_(ss, auth, empId, dateObj, plannedStatus, shift, reason) {
  const sh = (typeof rbEnsureRosterSheet_ === 'function')
             ? rbEnsureRosterSheet_(ss) : ss.getSheetByName('Roster');
  if (!sh) throw new Error('Roster sheet not found.');

  const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
              .map(function (h) { return String(h).trim(); });
  const c = {};
  ['Employee ID','Name','Location','Date','Planned Status','Shift','Status',
   'Published By','Published On','Last Amended By','Last Amended On','Amend Reason']
    .forEach(function (k) { c[k] = H.indexOf(k); });
  if (c['Employee ID'] < 0 || c['Date'] < 0 || c['Planned Status'] < 0) {
    throw new Error('Roster is missing a required column.');
  }

  const now = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm:ss');
  const dateStr = formatDdMmYyyy_(dateObj);
  var empName = empId;
  if (typeof vsEmpName_ === 'function') {
    try { empName = vsEmpName_(empId) || empId; } catch (e) {}
  }

  var found = -1;
  if (sh.getLastRow() > 1) {
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][c['Employee ID']]).trim() !== empId) continue;
      const rd = parseFlexibleDate_(data[i][c['Date']]);
      if (svValidDate_(rd) && sameCalendarDate_(rd, dateObj)) { found = i + 2; break; }
    }
  }

  function put(col, val) { if (c[col] >= 0) sh.getRange(found, c[col] + 1).setValue(val); }

  if (found > -1) {
    put('Planned Status', plannedStatus);
    put('Shift', plannedStatus === 'Day Off' ? '' : (shift || ''));
    put('Status', 'Published');
    if (c['Published By'] >= 0 && !sh.getRange(found, c['Published By'] + 1).getValue()) {
      put('Published By', auth.name);
      put('Published On', now);
    }
    put('Last Amended By', auth.name);
    put('Last Amended On', now);
    put('Amend Reason', reason);
  } else {
    const row = new Array(H.length).fill('');
    function set(col, val) { if (c[col] >= 0) row[c[col]] = val; }
    set('Employee ID', empId);
    set('Name', empName);
    set('Location', auth.location);
    set('Date', dateStr);
    set('Planned Status', plannedStatus);
    set('Shift', plannedStatus === 'Day Off' ? '' : (shift || ''));
    set('Status', 'Published');
    set('Published By', auth.name);
    set('Published On', now);
    set('Last Amended By', auth.name);
    set('Last Amended On', now);
    set('Amend Reason', reason);
    sh.appendRow(row);
  }

  if (typeof vsAudit_ === 'function') {
    try {
      vsAudit_(empId, empName, dateStr,
               'ROSTER SET: ' + plannedStatus + ' — ' + reason, auth.name);
    } catch (e) {}
  }
}

function svAudit_(ss, admin, empId, empName, dateStr, stage, oldVal, newVal) {
  var a = ss.getSheetByName('Attendance_Amendments');
  if (!a) {
    a = ss.insertSheet('Attendance_Amendments');
    a.appendRow(['When','Admin','Employee ID','Employee Name','Date','Stage','Old Value','New Value']);
  }
  const when = formatDdMmYyyy_(new Date()) + ' ' +
               Utilities.formatDate(new Date(), tz_(), 'HH:mm:ss');
  a.appendRow([when, admin, empId, empName, dateStr, stage, oldVal, newVal]);
}