// ============================================================
// Notifications.gs — all HR email alerts, one place
// ------------------------------------------------------------
// Every notification in the system routes through here so there is a
// single place to change the recipient, the styling, or the schedule.
//
// SETUP (run ONCE from the editor):
//   setupNotificationTriggers()
//     → 07:00 daily  : dailyAttendanceAlert()
//     → 08:00 Monday : weeklyRosterGapAlert()
//   Then optionally send yourself a test:
//     sendTestNotification()
//
// Every function is also safe to run manually from the editor.
// ============================================================

const NOTIFY_TO      = 'shakeel.viam@gmail.com';
const NOTIFY_FROM    = 'Maya HR System';
const NOTIFY_ENABLED = true;   // master switch — set false to silence all email

// ── Shared HTML shell so every email looks the same ──────────────────
function nt_wrap_(title, subtitle, bodyHtml, accent) {
  accent = accent || '#0d9488';
  return '' +
  '<div style="font-family:system-ui,-apple-system,Segoe UI,sans-serif;background:#f6f7f8;padding:24px">' +
    '<div style="max-width:640px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08)">' +
      '<div style="background:#0f2027;padding:20px 24px">' +
        '<div style="color:#fff;font-size:18px;font-weight:700">' + title + '</div>' +
        (subtitle ? '<div style="color:#9fb3bd;font-size:13px;margin-top:4px">' + subtitle + '</div>' : '') +
      '</div>' +
      '<div style="padding:24px">' + bodyHtml + '</div>' +
      '<div style="padding:14px 24px;background:#fafbfc;border-top:1px solid #eceff1;color:#8a9aa5;font-size:11px">' +
        'Maya Tex Restaurants — HR System · sent ' +
        Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm') +
      '</div>' +
    '</div>' +
  '</div>';
}

function nt_stat_(label, value, color) {
  return '<td style="padding:10px 14px;background:#fafbfc;border:1px solid #eceff1;border-radius:8px;text-align:center">' +
    '<div style="font-size:22px;font-weight:700;color:' + (color || '#0f2027') + '">' + value + '</div>' +
    '<div style="font-size:11px;color:#8a9aa5;text-transform:uppercase;letter-spacing:.4px">' + label + '</div>' +
  '</td>';
}

function nt_table_(headers, rows) {
  if (!rows.length) return '';
  let h = '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:8px">' +
    '<thead><tr>' + headers.map(x =>
      '<th style="text-align:left;padding:8px;background:#f2f4f5;border-bottom:2px solid #e0e4e6;' +
      'font-size:11px;text-transform:uppercase;color:#61727d">' + x + '</th>').join('') +
    '</tr></thead><tbody>';
  rows.forEach(r => {
    h += '<tr>' + r.map(c =>
      '<td style="padding:8px;border-bottom:1px solid #eceff1">' + c + '</td>').join('') + '</tr>';
  });
  return h + '</tbody></table>';
}

function nt_send_(subject, html) {
  if (!NOTIFY_ENABLED) { Logger.log('Notifications disabled — would have sent: ' + subject); return false; }
  try {
    MailApp.sendEmail({ to: NOTIFY_TO, subject: subject, htmlBody: html, name: NOTIFY_FROM });
    Logger.log('Sent: ' + subject);
    return true;
  } catch (err) {
    Logger.log('MAIL FAILED (' + subject + '): ' + err.toString());
    return false;
  }
}

// ══════════════════════════════════════════════════════════════════
// 1. DAILY ATTENDANCE ALERT — 07:00, covers yesterday
// ══════════════════════════════════════════════════════════════════
function dailyAttendanceAlert() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const y  = new Date(); y.setDate(y.getDate() - 1);
  const dStr = formatDdMmYyyy_(y);

  const sheet = ss.getSheetByName('Attendance_Daily');
  if (!sheet || sheet.getLastRow() < 2) {
    nt_send_('⚠️ Maya HR — No attendance data for ' + dStr,
      nt_wrap_('Attendance Alert', dStr,
        '<p style="color:#b3261e"><b>Attendance_Daily has no rows.</b> The 4 AM compile may not have run. ' +
        'Check Apps Script → Executions for <code>dailyAttendanceTrigger</code>.</p>', '#b3261e'));
    return;
  }

  const C = ad_cols_(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  let present = 0, absent = 0, dayOff = 0, onLeave = 0, noRoster = 0;
  const absentList = [], reviewList = [], noRosterList = [];

  rows.forEach(r => {
    if (ad_normDate_(r[C.date]) !== dStr) return;
    const st   = String(r[C.status]  || '').trim();
    const dt   = C.dayType >= 0 ? String(r[C.dayType] || '').trim() : '';
    const name = C.name    >= 0 ? String(r[C.name]    || '').trim() : '';
    const loc  = C.location>= 0 ? String(r[C.location]|| '').trim() : '';
    const id   = String(r[C.id] || '').trim();

    if (st === 'Present')       present++;
    else if (st === 'Absent')  { absent++;  absentList.push([id, name, loc]); }
    else if (st === 'Day Off')  dayOff++;
    else if (st === 'On Leave') onLeave++;
    else if (st === 'No Roster'){ noRoster++; noRosterList.push([id, name, loc]); }

    if (dt === 'Half Day' || dt === 'Check Punches') {
      const pres = C.presence >= 0 ? (Number(r[C.presence]) || 0).toFixed(2) : '—';
      reviewList.push([id, name, loc, dt, pres + 'h']);
    }
  });

  if (!present && !absent && !dayOff && !onLeave && !noRoster) {
    nt_send_('⚠️ Maya HR — No rows compiled for ' + dStr,
      nt_wrap_('Attendance Alert', dStr,
        '<p style="color:#b3261e"><b>No attendance rows exist for ' + dStr + '.</b> ' +
        'The nightly compile likely did not run for this date.</p>', '#b3261e'));
    return;
  }

  let body =
    '<table style="width:100%;border-spacing:8px 0"><tr>' +
      nt_stat_('Present', present, '#0f7b4f') +
      nt_stat_('Absent', absent, absent ? '#b3261e' : '#0f2027') +
      nt_stat_('Day Off', dayOff) +
      nt_stat_('On Leave', onLeave) +
      nt_stat_('No Roster', noRoster, noRoster ? '#a06800' : '#0f2027') +
    '</tr></table>';

  if (absentList.length) {
    body += '<h3 style="font-size:14px;margin:22px 0 4px;color:#b3261e">Absent — will be deducted (Basic ÷ 26)</h3>' +
            nt_table_(['ID', 'Name', 'Location'], absentList);
  }
  if (reviewList.length) {
    body += '<h3 style="font-size:14px;margin:22px 0 4px;color:#a06800">Needs HR review</h3>' +
            '<p style="font-size:12px;color:#61727d;margin:0">Not deducted automatically — someone must decide.</p>' +
            nt_table_(['ID', 'Name', 'Location', 'Type', 'Presence'], reviewList);
  }
  if (noRosterList.length) {
    body += '<h3 style="font-size:14px;margin:22px 0 4px;color:#61727d">No roster published</h3>' +
            '<p style="font-size:12px;color:#61727d;margin:0">Not deducted. Supervisor has not planned these days.</p>' +
            nt_table_(['ID', 'Name', 'Location'], noRosterList.slice(0, 25)) +
            (noRosterList.length > 25 ? '<p style="font-size:12px;color:#8a9aa5">…and ' + (noRosterList.length - 25) + ' more</p>' : '');
  }

  const flag = absent || reviewList.length ? '⚠️ ' : '✅ ';
  nt_send_(flag + 'Maya HR — Attendance ' + dStr +
    ' (' + present + ' present, ' + absent + ' absent)',
    nt_wrap_('Daily Attendance', dStr, body));
}

// ══════════════════════════════════════════════════════════════════
// 2. WEEKLY ROSTER GAP ALERT — Monday 08:00, looks at the week ahead
// ══════════════════════════════════════════════════════════════════
function weeklyRosterGapAlert() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const roster = ss.getSheetByName('Roster');
  const hr     = ss.getSheetByName('HR Maya');
  if (!hr) return;

  const EXEMPT = ['Head Office', 'Driver Location'];

  const hrH = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0];
  const hc  = (n) => hrH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const HID = hc('employee id'), HNM = hc('name (english)');
  let HLOC = hc('location'); if (HLOC === -1) HLOC = hc('branch');
  const HST = hc('status');

  const staff = [];
  hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().forEach(r => {
    const id = String(r[HID] || '').trim();
    if (!id) return;
    if (HST !== -1 && String(r[HST] || '').trim().toLowerCase() === 'inactive') return;
    const loc = HLOC !== -1 ? String(r[HLOC] || '').trim() : '';
    if (EXEMPT.indexOf(loc) !== -1) return;   // these locations don't roster
    staff.push({ id: id, name: HNM !== -1 ? String(r[HNM] || '').trim() : id, loc: loc });
  });

  // Dates for the coming 7 days
  const days = [];
  for (let i = 0; i < 7; i++) { const d = new Date(); d.setDate(d.getDate() + i); days.push(formatDdMmYyyy_(d)); }

  // Who has a PUBLISHED roster on which day
  const have = {};
  if (roster && roster.getLastRow() > 1) {
    const RH = roster.getRange(1, 1, 1, roster.getLastColumn()).getValues()[0];
    const rc = (n) => RH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const RID = rc('employee id'), RDT = rc('date'), RST = rc('status');
    roster.getRange(2, 1, roster.getLastRow() - 1, roster.getLastColumn()).getValues().forEach(r => {
      if (RST !== -1 && String(r[RST] || '').trim() !== 'Published') return;
      have[String(r[RID] || '').trim() + '|' + ad_normDate_(r[RDT])] = true;
    });
  }

  const byLoc = {};
  staff.forEach(s => {
    const missing = days.filter(d => !have[s.id + '|' + d]).length;
    if (!missing) return;
    if (!byLoc[s.loc]) byLoc[s.loc] = [];
    byLoc[s.loc].push([s.id, s.name, missing + ' of 7']);
  });

  const locs = Object.keys(byLoc).sort();
  if (!locs.length) {
    nt_send_('✅ Maya HR — Roster complete for the week ahead',
      nt_wrap_('Weekly Roster Check', days[0] + ' → ' + days[6],
        '<p style="color:#0f7b4f"><b>Every location has a published roster for all 7 days.</b></p>'));
    return;
  }

  let body = '<p style="font-size:13px;color:#61727d">Staff below have no published roster for part of the week ahead. ' +
             'They are <b>not</b> deducted for these days, but coverage is unplanned.</p>';
  let total = 0;
  locs.forEach(loc => {
    total += byLoc[loc].length;
    body += '<h3 style="font-size:14px;margin:20px 0 4px">' + (loc || '(no location)') + '</h3>' +
            nt_table_(['ID', 'Name', 'Days missing'], byLoc[loc]);
  });

  nt_send_('⚠️ Maya HR — ' + total + ' staff without roster this week',
    nt_wrap_('Weekly Roster Check', days[0] + ' → ' + days[6], body, '#a06800'));
}

// ══════════════════════════════════════════════════════════════════
// 3. PRE-PAYROLL READINESS — call before running payroll
// ══════════════════════════════════════════════════════════════════
function payrollReadinessAlert(from, to) {
  if (!from || !to) { Logger.log('payrollReadinessAlert: from/to required.'); return; }
  const r = getPayrollReadiness({ from: from, to: to });
  if (!r.success) {
    nt_send_('❌ Maya HR — Payroll readiness check failed',
      nt_wrap_('Payroll Readiness', from + ' → ' + to,
        '<p style="color:#b3261e">' + r.error + '</p>', '#b3261e'));
    return;
  }

  const rows = r.issues.map(i => [
    i.severity === 'blocker' ? '<b style="color:#b3261e">BLOCKER</b>'
      : i.severity === 'review' ? '<b style="color:#a06800">REVIEW</b>'
      : '<span style="color:#61727d">INFO</span>',
    i.type, i.count, i.message,
  ]);

  const body =
    '<p style="font-size:15px;font-weight:600;color:' + (r.ready ? '#0f7b4f' : '#b3261e') + '">' +
      (r.ready ? '✅ Ready to run payroll' : '❌ ' + r.blockerCount + ' blocker(s) — do not run payroll yet') +
    '</p>' +
    (rows.length ? nt_table_(['Severity', 'Type', 'Count', 'Detail'], rows)
                 : '<p style="color:#0f7b4f">Nothing flagged.</p>');

  nt_send_((r.ready ? '✅' : '❌') + ' Maya HR — Payroll readiness ' + from + ' → ' + to,
    nt_wrap_('Payroll Readiness', from + ' → ' + to, body, r.ready ? '#0f7b4f' : '#b3261e'));
}

// ══════════════════════════════════════════════════════════════════
// 4. PAYROLL RUN COMPLETED — call after a successful run
// ══════════════════════════════════════════════════════════════════
function payrollCompletedAlert(result, from, to) {
  const body =
    '<table style="width:100%;border-spacing:8px 0"><tr>' +
      nt_stat_('Employees', result.count || 0) +
      nt_stat_('Total Net (KD)', (result.totalNet || 0).toFixed(3), '#0f7b4f') +
    '</tr></table>' +
    '<p style="font-size:13px;color:#61727d;margin-top:18px">' +
      'Sheet: <b>' + (result.sheetName || '—') + '</b><br>' +
      (result.archived ? 'OT_Entries archived to: <b>' + result.archived + '</b>' : 'OT_Entries not archived.') +
    '</p>';
  nt_send_('✅ Maya HR — Payroll run ' + from + ' → ' + to,
    nt_wrap_('Payroll Completed', from + ' → ' + to, body, '#0f7b4f'));
}

// ══════════════════════════════════════════════════════════════════
// 5. SYSTEM HEALTH — did the nightly compile actually run?
// ══════════════════════════════════════════════════════════════════
function systemHealthAlert() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const problems = [];

  const y = new Date(); y.setDate(y.getDate() - 1);
  const dStr = formatDdMmYyyy_(y);

  const ad = ss.getSheetByName('Attendance_Daily');
  if (!ad || ad.getLastRow() < 2) {
    problems.push('Attendance_Daily is empty.');
  } else {
    const C = ad_cols_(ad);
    const found = ad.getRange(2, 1, ad.getLastRow() - 1, ad.getLastColumn()).getValues()
      .some(r => ad_normDate_(r[C.date]) === dStr);
    if (!found) problems.push('No Attendance_Daily rows for yesterday (' + dStr + ') — nightly compile may have failed.');
  }

  // Encashment rows wrongly sitting in OT_Entries (the double-pay bug)
  const ot = ss.getSheetByName('OT_Entries');
  if (ot && ot.getLastRow() > 1) {
    const H = ot.getRange(1, 1, 1, ot.getLastColumn()).getValues()[0];
    const cSrc = H.findIndex(h => String(h).trim().toLowerCase() === 'source');
    if (cSrc !== -1) {
      const n = ot.getRange(2, 1, ot.getLastRow() - 1, ot.getLastColumn()).getValues()
        .filter(r => String(r[cSrc] || '').trim() === 'Leave Encashment').length;
      if (n) problems.push(n + ' encashment row(s) in OT_Entries — these would be PAID TWICE.');
    }
  }

  // Is the 4 AM trigger still installed?
  const hasTrigger = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'dailyAttendanceTrigger');
  if (!hasTrigger) problems.push('The 4 AM dailyAttendanceTrigger is NOT installed.');

  if (!problems.length) {
    nt_send_('✅ Maya HR — System healthy',
      nt_wrap_('System Health', dStr, '<p style="color:#0f7b4f">All checks passed.</p>', '#0f7b4f'));
    return;
  }
  nt_send_('🚨 Maya HR — ' + problems.length + ' system issue(s)',
    nt_wrap_('System Health', dStr,
      '<ul style="color:#b3261e;font-size:13px;line-height:1.7">' +
      problems.map(p => '<li>' + p + '</li>').join('') + '</ul>', '#b3261e'));
}

// ══════════════════════════════════════════════════════════════════
// SETUP — run once
// ══════════════════════════════════════════════════════════════════
function setupNotificationTriggers() {
  ['dailyAttendanceAlert', 'weeklyRosterGapAlert', 'systemHealthAlert'].forEach(fn => {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
    });
  });

  ScriptApp.newTrigger('dailyAttendanceAlert').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('systemHealthAlert').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('weeklyRosterGapAlert').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  Logger.log('Notification triggers installed:\n' +
    '  07:00 daily  → dailyAttendanceAlert + systemHealthAlert\n' +
    '  08:00 Monday → weeklyRosterGapAlert\n' +
    'All sent to ' + NOTIFY_TO);
}

function sendTestNotification() {
  const ok = nt_send_('✅ Maya HR — Test notification',
    nt_wrap_('Test', 'If you can read this, email is working',
      '<p>Notifications are configured and sending to <b>' + NOTIFY_TO + '</b>.</p>' +
      '<p style="font-size:13px;color:#61727d">Next: run <code>setupNotificationTriggers()</code> once to schedule them.</p>',
      '#0f7b4f'));
  Logger.log(ok ? 'Test sent to ' + NOTIFY_TO : 'Test FAILED — check quota/permissions.');
}