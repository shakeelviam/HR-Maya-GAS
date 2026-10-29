// ============================================================
// SubmissionSync.gs — staff document updates reach HR Maya by themselves
// ------------------------------------------------------------
// The staff form (https://shakeelviam.github.io/maya-tex-docs/) writes
// into the Staff Docs Collection workbook. Getting those answers into
// HR Maya used to need two functions run by hand:
//     pullSubmissionsFromDocs  →  syncStaffSubmissions
// so a renewed Civil ID could sit for days while the expiry reminders
// kept going out.
//
// This file runs that same pair on a schedule and tells you what came
// in. Nothing about the mapping is repeated here — it calls the existing
// refreshAndSyncSubmissions in PullSubmissions.gs.
//
// SETUP — run ONCE from the editor:
//   setupSubmissionSyncTrigger()     → every hour
// Useful any time:
//   syncSubmissionsNow()             → run it immediately
//   checkSubmissionSyncTrigger()     → is it actually scheduled?
//
// The email goes out ONLY when somebody has submitted since the last
// run, so an idle hour is silent. Each run remembers the newest
// submission it has seen in a script property, so nothing is announced
// twice and a missed hour is caught up on the next one.
// ============================================================

const SUBSYNC_EVERY_HOURS = 1;
const SUBSYNC_PROP = 'SUBSYNC_LAST_TS';          // newest submission already reported
const SUBSYNC_DOC_COLS = ['CID Expiry', 'Passport Expiry', 'Health Expiry', 'Civil ID', 'Passport No'];

function subSyncRecipient_() {
  if (typeof NOTIFY_TO === 'string' && NOTIFY_TO) return NOTIFY_TO;
  return Session.getEffectiveUser().getEmail();
}

function subSyncFmt_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return (typeof formatDdMmYyyy_ === 'function') ? formatDdMmYyyy_(v)
      : Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }
  return String(v == null ? '' : v).trim();
}

// A timestamp cell → milliseconds. Sheets gives a Date; a text cell goes
// through the shared parser. Checked by tag, not instanceof, so a Date
// from another execution context still counts.
function subSyncTime_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? 0 : v.getTime();
  }
  if (typeof parseFlexibleDate_ === 'function') {
    const d = parseFlexibleDate_(v);
    if (d) return d.getTime();
  }
  const d2 = new Date(String(v || ''));
  return isNaN(d2.getTime()) ? 0 : d2.getTime();
}

// Submissions newer than the last one we reported.
// Returns { rows:[{name, civil, when, cells:[[label,value]…]}], newestTs }
function subSyncNewSince_(ss, lastTs) {
  const tabName = (typeof STAFF_SUBMISSION_TAB === 'string') ? STAFF_SUBMISSION_TAB : 'Staff Submission';
  const sh = ss.getSheetByName(tabName);
  const out = { rows: [], newestTs: lastTs };
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const H = d[0].map(h => String(h).trim());
  const at = (n) => H.findIndex(h => h.toLowerCase() === String(n).toLowerCase());
  const cTs = at('Timestamp'), cName = at('Name'), cCivil = at('Civil ID');
  if (cTs === -1) return out;

  for (let i = 1; i < d.length; i++) {
    const ts = subSyncTime_(d[i][cTs]);
    if (!ts || ts <= lastTs) continue;
    if (ts > out.newestTs) out.newestTs = ts;
    const cells = [];
    SUBSYNC_DOC_COLS.forEach(name => {
      const c = at(name);
      if (c === -1) return;
      const v = subSyncFmt_(d[i][c]);
      if (v) cells.push([name, v]);
    });
    out.rows.push({
      name: cName !== -1 ? String(d[i][cName] || '').trim() : '(no name)',
      civil: cCivil !== -1 ? String(d[i][cCivil] || '').trim() : '',
      when: subSyncFmt_(d[i][cTs]),
      cells: cells
    });
  }
  return out;
}

// Pull + sync + report. Safe to run by hand; that is also what the
// hourly trigger calls.
function syncSubmissionsNow() {
  if (typeof refreshAndSyncSubmissions !== 'function') {
    const err = 'PullSubmissions.gs is missing — refreshAndSyncSubmissions not found.';
    Logger.log(err);
    return { success: false, error: err };
  }
  const props = PropertiesService.getScriptProperties();
  const lastTs = Number(props.getProperty(SUBSYNC_PROP) || 0);

  const res = refreshAndSyncSubmissions({});
  if (!res || !res.success) {
    Logger.log('Submission sync FAILED: ' + ((res && (res.error || res.sync)) || 'unknown'));
    return res || { success: false, error: 'no result' };
  }

  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const fresh = subSyncNewSince_(ss, lastTs);
  props.setProperty(SUBSYNC_PROP, String(fresh.newestTs));

  if (!fresh.rows.length) {
    Logger.log('Submission sync ran — nothing new. ' + res.message);
    return { success: true, newSubmissions: 0, message: res.message };
  }

  // Someone updated their documents — say who, and what they sent.
  const has = (typeof nt_wrap_ === 'function' && typeof nt_table_ === 'function' && typeof nt_send_ === 'function');
  const headers = ['Submitted', 'Name', 'Civil ID', 'What they sent'];
  const rows = fresh.rows.map(r => [
    r.when, r.name, r.civil || '—',
    r.cells.map(c => c[0] + ': <b>' + c[1] + '</b>').join('<br>') || '—'
  ]);
  let body = '<p style="font-size:13px;color:#61727d">These reached HR Maya just now, through the staff document form. ' +
    'Their expiry reminders restart from the new dates.</p>';
  body += has ? nt_table_(headers, rows) : subSyncPlainTable_(headers, rows);
  body += '<p style="font-size:12px;color:#61727d;margin-top:14px">' + res.message + '</p>';

  const subject = '📥 Maya HR — ' + fresh.rows.length + ' document update' +
                  (fresh.rows.length === 1 ? '' : 's') + ' synced';
  if (has) nt_send_(subject, nt_wrap_('Document Updates', subSyncFmt_(new Date()) + ' · staff form', body, '#0d9488'));
  else {
    try {
      MailApp.sendEmail({ to: subSyncRecipient_(), subject: subject, htmlBody: body, name: 'Maya HR System' });
    } catch (err) { Logger.log('MAIL FAILED: ' + err.toString()); }
  }
  Logger.log(subject + ' — ' + res.message);
  return { success: true, newSubmissions: fresh.rows.length, message: res.message };
}

function subSyncPlainTable_(headers, rows) {
  let h = '<table style="border-collapse:collapse;font-family:Arial;font-size:13px"><tr>' +
    headers.map(x => '<th style="text-align:left;padding:6px;border-bottom:2px solid #ccc">' + x + '</th>').join('') + '</tr>';
  rows.forEach(r => { h += '<tr>' + r.map(c => '<td style="padding:6px;border-bottom:1px solid #eee;vertical-align:top">' + c + '</td>').join('') + '</tr>'; });
  return h + '</table>';
}

// ── Triggers ─────────────────────────────────────────────────────────
function setupSubmissionSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncSubmissionsNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncSubmissionsNow').timeBased().everyHours(SUBSYNC_EVERY_HOURS).create();
  const msg = 'Staff document sync scheduled every ' + SUBSYNC_EVERY_HOURS + ' hour(s). ' +
              'Email goes to ' + subSyncRecipient_() + ' only when something new comes in.';
  Logger.log(msg);
  return msg;
}

function checkSubmissionSyncTrigger() {
  const found = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncSubmissionsNow');
  const last = PropertiesService.getScriptProperties().getProperty(SUBSYNC_PROP);
  const msg = (found.length
      ? 'Installed: ' + found.length + ' trigger(s) for syncSubmissionsNow, every ' + SUBSYNC_EVERY_HOURS + ' hour(s).'
      : 'NOT installed — run setupSubmissionSyncTrigger() once.') +
    ' Last submission reported: ' + (last && Number(last) ? subSyncFmt_(new Date(Number(last))) : 'none yet') + '.';
  Logger.log(msg);
  return msg;
}

// Start reporting from now on, ignoring everything already in the tab.
// Use it once after installing, so the first email is not a dump of
// every submission ever made.
function markSubmissionsSeen() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const fresh = subSyncNewSince_(ss, 0);
  PropertiesService.getScriptProperties().setProperty(SUBSYNC_PROP, String(fresh.newestTs));
  const msg = 'Marked ' + fresh.rows.length + ' existing submission(s) as already seen. ' +
              'Only newer ones will be emailed.';
  Logger.log(msg);
  return msg;
}