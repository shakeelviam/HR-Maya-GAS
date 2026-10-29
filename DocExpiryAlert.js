// ============================================================
// DocExpiryAlert.gs — document expiry reminders, one per milestone
// ------------------------------------------------------------
// Civil ID, passport and health card expiry dates live in HR Maya. This
// file watches them and emails HR when a document reaches a reminder
// point. Each document has its own schedule (Shakeel, 23-09-2026):
//
//   Passport     6 months ahead, then every month   → 180,150,120,90,60,30
//   Civil ID     1 month ahead, then 15 days later  → 30,15
//   Health card  2 months ahead, then every 15 days → 60,45,30,15
//   Anything already expired → chased again every 7 days
//
// The schedule is data, not logic: edit DOC_EXPIRY_RULES to change a
// lead time, a follow-up interval or to watch another column. Nothing
// else needs touching.
//
// EVERY REMINDER IS LOGGED in the Doc_Expiry_Log tab (one row per person
// per document per milestone), and a milestone already in that log is
// never emailed twice. That is what makes the follow-ups real follow-ups
// instead of the same list every morning — and it means a day the
// trigger does not run is caught up the next day rather than skipped.
//
// SETUP — run ONCE from the editor:
//   setupDocExpiryTrigger()      → 07:00 every day
// Useful any time:
//   previewDocumentExpiry()      → the full picture, sends nothing
//   documentExpiryAlert()        → send now (same as the daily run)
//   checkDocExpiryTrigger()      → is it actually scheduled?
//
// It reuses the shared email look (nt_wrap_ / nt_table_ / nt_send_) from
// Notifications.gs when that file is present, and falls back to plain
// MailApp if it is not, so this file works on its own.
// ============================================================

// header  = the HR Maya column holding the expiry date
// label   = what staff call it
// first   = first reminder, days before expiry
// every   = follow-up interval in days after the first reminder
// updateField = what the staff form calls it, for the reminder text
const DOC_EXPIRY_RULES = [
  { header: 'Passport Expiry', label: 'Passport',    first: 180, every: 30 },
  { header: 'CID Expiry',      label: 'Civil ID',    first: 30,  every: 15 },
  { header: 'Health Expiry',   label: 'Health card', first: 60,  every: 15 }
];

const DOC_EXPIRED_CHASE_DAYS = 7;     // expired documents: chase this often
const DOC_EXPIRY_HOUR = 7;            // 07:00 daily
const DOC_EXPIRY_LOG_TAB = 'Doc_Expiry_Log';
const DOC_EXPIRY_LOG_HEADERS = ['Sent On', 'Employee ID', 'Name', 'Location', 'Document', 'Expiry Date', 'Milestone', 'Days Left'];
const DOC_UPDATE_LINK = 'https://shakeelviam.github.io/maya-tex-docs/';

// Who gets it. Falls back to the Notifications.gs recipient, then to the
// account running the script, so it is never sent nowhere.
function docExpiryRecipient_() {
  if (typeof NOTIFY_TO === 'string' && NOTIFY_TO) return NOTIFY_TO;
  return Session.getEffectiveUser().getEmail();
}

// dd-mm-yyyy, yyyy-mm-dd or a real date cell → Date at midnight.
function docExpParse_(v) {
  if (v === '' || v == null) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function docExpFmt_(d) {
  const p = (n) => String(n).length < 2 ? '0' + n : String(n);
  return p(d.getDate()) + '-' + p(d.getMonth() + 1) + '-' + d.getFullYear();
}

// The reminder points for one rule, largest first: 180,150,…,30 etc.
function docExpMilestones_(rule) {
  const out = [];
  for (let d = rule.first; d >= 1; d -= rule.every) out.push(d);
  return out;
}

// Which milestone a document is AT today: the smallest milestone that
// daysLeft has already reached. 46 days before a health card expires is
// milestone 60 (the 45 one has not arrived yet); 44 days is milestone 45.
function docExpMilestoneFor_(rule, daysLeft) {
  const ms = docExpMilestones_(rule);
  let hit = null;
  ms.forEach(m => { if (daysLeft <= m && (hit === null || m < hit)) hit = m; });
  return hit;
}

// ── The log: what has already been sent ──────────────────────────────
function docExpLogSheet_(ss) {
  let sh = ss.getSheetByName(DOC_EXPIRY_LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(DOC_EXPIRY_LOG_TAB);
    sh.getRange(1, 1, 1, DOC_EXPIRY_LOG_HEADERS.length).setValues([DOC_EXPIRY_LOG_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// key: employee + document + THAT expiry date + milestone. Renewing the
// document changes the date, so the new one starts its reminders fresh.
function docExpKey_(id, doc, expiryStr, milestone) {
  return [id, doc, expiryStr, milestone].join('|');
}

function docExpSentIndex_(ss) {
  const sh = docExpLogSheet_(ss);
  const idx = {};
  if (sh.getLastRow() < 2) return idx;
  const d = sh.getDataRange().getValues();
  const H = d[0].map(h => String(h).trim().toLowerCase());
  const cId = H.indexOf('employee id'), cDoc = H.indexOf('document'),
        cExp = H.indexOf('expiry date'), cMs = H.indexOf('milestone');
  for (let i = 1; i < d.length; i++) {
    const id = String(d[i][cId] || '').trim();
    if (!id) continue;
    const exp = d[i][cExp] instanceof Date ? docExpFmt_(d[i][cExp]) : String(d[i][cExp] || '').trim();
    idx[docExpKey_(id, String(d[i][cDoc] || '').trim(), exp, String(d[i][cMs] || '').trim())] = true;
  }
  return idx;
}

// ── Read HR Maya and work out what is due ────────────────────────────
// Returns { due:[…], expired:[…], upcoming:[…], checked }
//   due      — a milestone reached and not yet emailed (this is what is sent)
//   expired  — past expiry, chased every DOC_EXPIRED_CHASE_DAYS days
//   upcoming — everything inside its first reminder window, sent or not
// Each row: { id, name, loc, doc, dateStr, days, milestone }
function collectDocExpiry_() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const hr = ss.getSheetByName('HR Maya');
  const empty = { due: [], expired: [], upcoming: [], checked: 0 };
  if (!hr || hr.getLastRow() < 2) return empty;

  const data = hr.getDataRange().getValues();
  const H = data[0].map(h => String(h).trim().toLowerCase());
  const at = (n) => H.indexOf(String(n).trim().toLowerCase());
  const cId = at('Employee ID'), cName = at('Name (English)'), cStatus = at('Status');
  let cLoc = at('Location'); if (cLoc === -1) cLoc = at('Branch');

  const rules = DOC_EXPIRY_RULES.map(r => {
    const copy = { header: r.header, label: r.label, first: r.first, every: r.every, col: at(r.header) };
    return copy;
  }).filter(r => r.col !== -1);

  const sent = docExpSentIndex_(ss);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const due = [], expired = [], upcoming = [];
  let checked = 0;

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][cId] || '').trim();
    if (!id) continue;
    // Leavers are not chased for documents.
    if (cStatus !== -1 && String(data[i][cStatus] || '').trim().toLowerCase() === 'inactive') continue;
    checked++;
    const name = cName !== -1 ? String(data[i][cName] || '').trim() : id;
    const loc = cLoc !== -1 ? String(data[i][cLoc] || '').trim() : '';

    rules.forEach(rule => {
      const d = docExpParse_(data[i][rule.col]);
      if (!d) return;                                  // blank or unreadable — nothing to claim
      const dateStr = docExpFmt_(d);
      const daysLeft = Math.round((d.getTime() - today.getTime()) / 86400000);
      const base = { id: id, name: name, loc: loc, doc: rule.label, dateStr: dateStr, days: daysLeft };

      if (daysLeft < 0) {
        // Expired: one reminder per DOC_EXPIRED_CHASE_DAYS-day block, so the
        // chase repeats weekly instead of every single morning.
        const block = Math.floor((-daysLeft) / DOC_EXPIRED_CHASE_DAYS);
        const ms = 'expired+' + (block * DOC_EXPIRED_CHASE_DAYS) + 'd';
        const row = Object.assign({ milestone: ms }, base);
        expired.push(row);
        if (!sent[docExpKey_(id, rule.label, dateStr, ms)]) due.push(row);
        return;
      }
      const milestone = docExpMilestoneFor_(rule, daysLeft);
      if (milestone === null) return;                  // still outside the first reminder
      const row = Object.assign({ milestone: String(milestone) }, base);
      upcoming.push(row);
      if (!sent[docExpKey_(id, rule.label, dateStr, String(milestone))]) due.push(row);
    });
  }

  const soonest = (a, b) => a.days - b.days;
  return { due: due.sort(soonest), expired: expired.sort(soonest), upcoming: upcoming.sort(soonest), checked: checked };
}

// Write what was emailed into Doc_Expiry_Log so it is not repeated.
function docExpRecord_(rows) {
  if (!rows.length) return;
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = docExpLogSheet_(ss);
  const today = docExpFmt_(new Date());
  const out = rows.map(x => [today, x.id, x.name, x.loc, x.doc, x.dateStr, x.milestone, x.days]);
  sh.getRange(sh.getLastRow() + 1, 1, out.length, DOC_EXPIRY_LOG_HEADERS.length).setValues(out);
}

// ── Plain-text picture for the editor; sends nothing, logs nothing ───
function previewDocumentExpiry() {
  const r = collectDocExpiry_();
  const line = (x) => '   ' + x.id + '  ' + x.name + '  (' + (x.loc || 'no location') + ')  ' +
    x.doc + '  ' + x.dateStr + '  ' +
    (x.days < 0 ? 'EXPIRED ' + (-x.days) + ' day(s) ago' : 'in ' + x.days + ' day(s)') +
    '  [milestone ' + x.milestone + ']';
  const out = ['Document expiry — ' + docExpFmt_(new Date()) + '  (' + r.checked + ' active staff)'];
  out.push('Schedule: ' + DOC_EXPIRY_RULES.map(x => x.label + ' ' + x.first + 'd then every ' + x.every + 'd').join(' · ') +
           ' · expired chased every ' + DOC_EXPIRED_CHASE_DAYS + 'd');
  out.push('');
  out.push('WOULD BE EMAILED NOW (' + r.due.length + ')');
  r.due.forEach(x => out.push(line(x)));
  out.push('');
  out.push('ALL EXPIRED (' + r.expired.length + ')');
  r.expired.forEach(x => out.push(line(x)));
  out.push('');
  out.push('ALL INSIDE THEIR REMINDER WINDOW (' + r.upcoming.length + ')');
  r.upcoming.forEach(x => out.push(line(x)));
  const text = out.join('\n');
  Logger.log(text);
  return text;
}

// ── The daily email ──────────────────────────────────────────────────
function documentExpiryAlert() {
  const r = collectDocExpiry_();
  if (!r.due.length) {
    Logger.log('Document expiry: nothing new to report today (' + r.expired.length + ' expired, ' +
               r.upcoming.length + ' in window — all already sent).');
    return false;
  }
  const has = (typeof nt_wrap_ === 'function' && typeof nt_table_ === 'function' && typeof nt_send_ === 'function');
  const rows = (list) => list.map(x => [
    x.id, x.name, x.loc || '—', x.doc, x.dateStr,
    x.days < 0 ? '<b style="color:#b3261e">expired ' + (-x.days) + ' day(s) ago</b>'
               : (x.days <= 15 ? '<b style="color:#b3261e">' + x.days + ' day(s)</b>' : x.days + ' day(s)')
  ]);
  const headers = ['ID', 'Name', 'Location', 'Document', 'Expires', 'When'];
  const dueExpired = r.due.filter(x => x.days < 0);
  const dueSoon = r.due.filter(x => x.days >= 0);
  const subtitle = docExpFmt_(new Date()) + ' · ' + r.checked + ' active staff · ' +
    DOC_EXPIRY_RULES.map(x => x.label + ' ' + x.first + 'd/' + x.every + 'd').join(' · ');

  let body = '';
  if (dueExpired.length) {
    body += '<h3 style="font-size:14px;margin:4px 0 4px;color:#b3261e">Expired — chase now (' + dueExpired.length + ')</h3>' +
      (has ? nt_table_(headers, rows(dueExpired)) : docExpPlainTable_(headers, rows(dueExpired)));
  }
  if (dueSoon.length) {
    body += '<h3 style="font-size:14px;margin:20px 0 4px">Reminder due (' + dueSoon.length + ')</h3>' +
      (has ? nt_table_(headers, rows(dueSoon)) : docExpPlainTable_(headers, rows(dueSoon)));
  }
  body += '<p style="font-size:12px;color:#61727d;margin-top:16px">' +
    'Staff update their own documents at <a href="' + DOC_UPDATE_LINK + '">' + DOC_UPDATE_LINK + '</a>; ' +
    'the new dates reach HR Maya through the submissions sync, and a renewed document starts its reminders again from the new date.<br>' +
    'Each person appears once per milestone — every reminder sent is recorded in the <b>' + DOC_EXPIRY_LOG_TAB + '</b> tab. ' +
    'A blank expiry cell is never reported, so an empty date is invisible here.</p>';

  const subject = '📄 Maya HR — ' + (dueExpired.length ? dueExpired.length + ' expired, ' : '') +
                  dueSoon.length + ' document reminder' + (dueSoon.length === 1 ? '' : 's');
  const ok = has
    ? nt_send_(subject, nt_wrap_('Document Expiry', subtitle, body, dueExpired.length ? '#b3261e' : '#a06800'))
    : docExpPlainSend_(subject, body);
  if (ok !== false) docExpRecord_(r.due);     // only record what actually went out
  return ok;
}

// Fallbacks used only when Notifications.gs is not in the project.
function docExpPlainTable_(headers, rows) {
  let h = '<table style="border-collapse:collapse;font-family:Arial;font-size:13px"><tr>' +
    headers.map(x => '<th style="text-align:left;padding:6px;border-bottom:2px solid #ccc">' + x + '</th>').join('') + '</tr>';
  rows.forEach(r => { h += '<tr>' + r.map(c => '<td style="padding:6px;border-bottom:1px solid #eee">' + c + '</td>').join('') + '</tr>'; });
  return h + '</table>';
}
function docExpPlainSend_(subject, html) {
  try {
    MailApp.sendEmail({ to: docExpiryRecipient_(), subject: subject, htmlBody: html, name: 'Maya HR System' });
    Logger.log('Sent: ' + subject);
    return true;
  } catch (err) { Logger.log('MAIL FAILED (' + subject + '): ' + err.toString()); return false; }
}

// ── WhatsApp helper ──────────────────────────────────────────────────
// One line per person to chase, ready to paste or to feed a sender:
//   MT-00009 | Bibek Nepali | +965… | Civil ID | 23-08-2026 | expired 31 day(s) ago
// Everyone inside their reminder window, whether or not HR has emailed.
function docExpiryChaseList() {
  const r = collectDocExpiry_();
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const hr = ss.getSheetByName('HR Maya');
  const data = hr.getDataRange().getValues();
  const H = data[0].map(h => String(h).trim().toLowerCase());
  const cId = H.indexOf('employee id'), cWa = H.indexOf('whatsapp'), cMob = H.indexOf('mobile');
  const phone = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][cId] || '').trim();
    if (!id) continue;
    phone[id] = String((cWa !== -1 ? data[i][cWa] : '') || (cMob !== -1 ? data[i][cMob] : '') || '').trim();
  }
  const all = r.expired.concat(r.upcoming).sort((a, b) => a.days - b.days);
  const text = all.map(x => [x.id, x.name, phone[x.id] || 'no number', x.doc, x.dateStr,
    x.days < 0 ? 'expired ' + (-x.days) + ' day(s) ago' : 'in ' + x.days + ' day(s)'].join(' | ')).join('\n');
  Logger.log(text || 'Nobody to chase.');
  return text;
}

// ── Triggers ─────────────────────────────────────────────────────────
function setupDocExpiryTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'documentExpiryAlert') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('documentExpiryAlert').timeBased().everyDays(1).atHour(DOC_EXPIRY_HOUR).create();
  const msg = 'Document expiry alert scheduled daily at ' + DOC_EXPIRY_HOUR + ':00, sent to ' + docExpiryRecipient_() + '.';
  Logger.log(msg);
  return msg;
}

// Is it actually scheduled? Answers the "why am I not getting mail" question.
function checkDocExpiryTrigger() {
  const found = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'documentExpiryAlert');
  const msg = found.length
    ? 'Installed: ' + found.length + ' trigger(s) for documentExpiryAlert. Mail goes to ' + docExpiryRecipient_() + '.'
    : 'NOT installed — run setupDocExpiryTrigger() once.';
  Logger.log(msg);
  return msg;
}

// ============================================================
// THE CHASE LIST, AS DATA
// ------------------------------------------------------------
// docExpiryChaseList() above prints one line per DOCUMENT, for a human
// to read. The WhatsApp sender needs something different: one row per
// MAN, with all of his expired documents together, a number that will
// actually reach him, and the language to write to him in.
//
// getDocChaseList is that, as JSON and as ready-made CSV, so the sender
// can pull the list itself instead of anybody retyping it.
// ============================================================

const DOC_LANG_HEADER = 'Language';     // optional column on HR Maya
const DOC_KW_PREFIX   = '965';

// Kuwait numbers are stored three ways in HR Maya: bare 8 digits,
// +965xxxxxxxx, and 00965xxxxxxxx. All three must come out the same.
function docNormNumber_(v) {
  let s = String(v == null ? '' : v).replace(/\D/g, '');
  if (!s) return '';
  if (s.length > 11 && s.indexOf('00') === 0) s = s.slice(2);
  if (s.length === 8) return DOC_KW_PREFIX + s;                 // bare Kuwait mobile
  if (s.length === 11 && s.indexOf(DOC_KW_PREFIX) === 0) return s;
  return s;                                                     // foreign, left as it is
}
function docIsKuwait_(n) { return String(n).length === 11 && String(n).indexOf(DOC_KW_PREFIX) === 0; }

// np / bn / en. Anything else, or blank, means English only.
function docNormLang_(v) {
  const s = String(v == null ? '' : v).trim().toLowerCase();
  if (!s) return '';
  if (s.indexOf('np') === 0 || s.indexOf('nep') === 0) return 'np';
  if (s.indexOf('bn') === 0 || s.indexOf('ban') === 0 || s.indexOf('bang') === 0) return 'bn';
  if (s.indexOf('en') === 0) return 'en';
  return '';
}

// Creates the optional Language column so it can be filled in once.
// Run it bare from the editor.
function addDocLanguageColumn() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const hr = ss.getSheetByName('HR Maya');
  if (!hr) return { success: false, error: 'No HR Maya tab.' };
  const H = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  if (H.indexOf(DOC_LANG_HEADER.toLowerCase()) !== -1) {
    const msg = 'HR Maya already has a "' + DOC_LANG_HEADER + '" column.';
    Logger.log(msg); return { success: true, message: msg };
  }
  hr.getRange(1, hr.getLastColumn() + 1).setValue(DOC_LANG_HEADER).setFontWeight('bold');
  const msg = 'Added a "' + DOC_LANG_HEADER + '" column to HR Maya. Put np, bn or en against each man. ' +
              'Blank means the message goes in English only.';
  Logger.log(msg);
  return { success: true, message: msg };
}

// ------------------------------------------------------------
// getDocChaseList({ scope })
//   scope 'expired' (default) — documents already past their date
//         'due'               — whatever today's email would cover
//         'all'               — expired plus everyone inside a reminder window
// ------------------------------------------------------------
function getDocChaseList(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const scope = String(params.scope || 'expired').trim().toLowerCase();

    const r = collectDocExpiry_();
    const rows = scope === 'all' ? r.expired.concat(r.upcoming)
               : scope === 'due' ? r.due
               : r.expired;

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const data = hr.getDataRange().getValues();
    const H = data[0].map(h => String(h).trim().toLowerCase());
    const at = (n) => H.indexOf(String(n).trim().toLowerCase());
    const cId = at('Employee ID'), cMob = at('Mobile'), cWa = at('WhatsApp'), cLang = at(DOC_LANG_HEADER);

    const info = {};
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][cId] || '').trim();
      if (!id) continue;
      const mob = docNormNumber_(cMob !== -1 ? data[i][cMob] : '');
      const wa  = docNormNumber_(cWa  !== -1 ? data[i][cWa]  : '');
      // A man working in Kuwait is reached on his Kuwait number. Several
      // rows carry a home-country number in the WhatsApp column, which is
      // where messages have been quietly going nowhere — so a Kuwait
      // number wins, whichever column it sits in.
      let number = '', note = '';
      if (docIsKuwait_(wa)) number = wa;
      else if (docIsKuwait_(mob)) { number = mob; if (wa) note = 'WhatsApp column holds a foreign number (+' + wa + ') — used the Kuwait mobile instead'; }
      else { number = wa || mob; if (number) note = 'no Kuwait number on file — using +' + number; }
      info[id] = { number: number, note: note, lang: docNormLang_(cLang !== -1 ? data[i][cLang] : '') };
    }

    // One row per man, his documents gathered together.
    const byMan = {};
    rows.forEach(x => {
      if (!byMan[x.id]) {
        const e = info[x.id] || { number: '', note: '', lang: '' };
        byMan[x.id] = { employee_id: x.id, name: x.name, location: x.loc,
                        number: e.number, numberNote: e.note, lang: e.lang,
                        docs: [], soonest: x.days };
      }
      byMan[x.id].docs.push({ doc: x.doc, date: x.dateStr, days: x.days });
      if (x.days < byMan[x.id].soonest) byMan[x.id].soonest = x.days;
    });

    const list = Object.keys(byMan).map(k => byMan[k]).sort((a, b) => a.soonest - b.soonest);
    list.forEach(m => {
      m.documents = m.docs.map(d => d.doc + '|' + d.date).join('; ');
      m.expiredCount = m.docs.length;
    });

    const sendable = list.filter(m => m.number);
    const noNumber = list.filter(m => !m.number);

    const q = (s) => '"' + String(s == null ? '' : s).replace(/"/g, '""') + '"';
    const csv = ['number,name,employee_id,location,lang,documents'].concat(
      sendable.map(m => [m.number, q(m.name), m.employee_id, q(m.location), m.lang, q(m.documents)].join(','))
    ).join('\n');

    return {
      success: true, scope: scope, data: sendable, csv: csv,
      noNumber: noNumber.map(m => ({ employee_id: m.employee_id, name: m.name, location: m.location })),
      warnings: sendable.filter(m => m.numberNote).map(m => m.name + ' (' + m.employee_id + '): ' + m.numberNote),
      noLanguage: sendable.filter(m => !m.lang).length,
      hasLanguageColumn: cLang !== -1,
      count: sendable.length,
      message: sendable.length + ' man/men to chase' +
               (noNumber.length ? ', ' + noNumber.length + ' with no number on file' : '') + '.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Editor-runnable: prints the CSV and drops a copy in Drive, for the
// times you would rather hand the sender a file than let it fetch.
function docExpiryChaseCsv() {
  const r = getDocChaseList({ scope: 'expired' });
  if (!r.success) { Logger.log('Failed: ' + r.error); return r; }
  Logger.log(r.message);
  if (r.warnings.length) { Logger.log(''); Logger.log('CHECK THESE NUMBERS:'); r.warnings.forEach(w => Logger.log('  ' + w)); }
  if (r.noNumber.length) { Logger.log(''); Logger.log('NO NUMBER ON FILE:'); r.noNumber.forEach(m => Logger.log('  ' + m.employee_id + ' ' + m.name + ' (' + m.location + ')')); }
  if (!r.hasLanguageColumn) Logger.log('\nHR Maya has no "' + DOC_LANG_HEADER + '" column — every message goes in English only. Run addDocLanguageColumn() to add one.');
  else if (r.noLanguage) Logger.log('\n' + r.noLanguage + ' man/men have no language set — those go in English only.');
  Logger.log('\n--- contacts-docs.csv ---\n' + r.csv);

  const folder = getOrCreateFolder_('Maya Document Chase');
  const name = 'contacts-docs ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd') + '.csv';
  folder.getFilesByName(name); // (a same-day re-run simply adds another copy)
  const file = folder.createFile(name, r.csv, MimeType.CSV);
  Logger.log('\nSaved to Drive: ' + file.getUrl());
  return { success: true, url: file.getUrl(), csv: r.csv, message: r.message };
}

// ============================================================
// "MY DOCUMENTS" — the kiosk hands a man straight to his own page
// ------------------------------------------------------------
// The staff document form (maya-tex-docs) makes a man find himself in a
// list of 68 names and then confirm he is who he says. At the kiosk that
// is wasted work: he has already signed in with his Employee ID and PIN,
// so the tablet knows him.
//
// The kiosk asks for a short-lived signed link, opens the form with it,
// and the form asks getMyDocuments who the man is and what is expired.
// The signature is what makes this safe to expose: without it, anybody
// could read any man's document dates by guessing MT-000xx.
// ============================================================

const DOC_LINK_TTL_MIN   = 30;                  // a kiosk link is good for half an hour
const DOC_LINK_SECRET_KEY = 'DOC_LINK_SECRET';

function docLinkSecret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty(DOC_LINK_SECRET_KEY);
  if (!s) {
    s = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty(DOC_LINK_SECRET_KEY, s);
  }
  return s;
}

// empId + the half-hour window it was minted in. A link stops working by
// itself, so one left open on a shared tablet is not a standing door.
function docLinkSig_(empId, win) {
  const raw = Utilities.computeHmacSha256Signature(
    String(empId).toUpperCase() + ':' + win, docLinkSecret_());
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').slice(0, 20);
}
function docLinkWindow_(offset) {
  return Math.floor(Date.now() / (DOC_LINK_TTL_MIN * 60000)) + (offset || 0);
}

// Called by the kiosk, for the man who is signed in at that moment.
function makeMyDocsLink(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const empId = String(params.empId || '').trim().toUpperCase();
    if (!empId) return { success: false, error: 'empId is required' };
    const win = docLinkWindow_(0);
    return { success: true, empId: empId, t: win + '.' + docLinkSig_(empId, win),
             minutes: DOC_LINK_TTL_MIN };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function docLinkValid_(empId, t) {
  const parts = String(t || '').split('.');
  if (parts.length !== 2) return false;
  const win = Number(parts[0]);
  if (!win) return false;
  const now = docLinkWindow_(0);
  // The window before this one is still accepted, so a link minted at
  // 10:29 does not die at 10:30 while the man is still typing.
  if (win !== now && win !== now - 1) return false;
  return parts[1] === docLinkSig_(empId, win);
}

// ------------------------------------------------------------
// getMyDocuments({ empId, t })
// Returns his name and every document the expiry rules watch, with its
// date and whether it has run out — so the form can open on his own page
// and show him exactly what needs doing.
// ------------------------------------------------------------
function getMyDocuments(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const empId = String(params.empId || '').trim().toUpperCase();
    if (!empId) return { success: false, error: 'empId is required' };
    if (!docLinkValid_(empId, params.t)) {
      return { success: false, expired: true,
               error: 'This link has expired. Sign in on the kiosk again and press Update my documents.' };
    }

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const data = hr.getDataRange().getValues();
    const H = data[0].map(h => String(h).trim().toLowerCase());
    const at = (n) => H.indexOf(String(n).trim().toLowerCase());
    const cId = at('Employee ID'), cName = at('Name (English)'), cDesig = at('Designation');
    let cLoc = at('Location'); if (cLoc === -1) cLoc = at('Branch');

    const rules = DOC_EXPIRY_RULES.map(r => ({ label: r.label, col: at(r.header) })).filter(r => r.col !== -1);
    const today = new Date(); today.setHours(0, 0, 0, 0);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cId] || '').trim().toUpperCase() !== empId) continue;
      const docs = rules.map(rule => {
        const d = docExpParse_(data[i][rule.col]);
        if (!d) return { doc: rule.label, date: '', status: 'missing', days: null };
        const days = Math.round((d.getTime() - today.getTime()) / 86400000);
        return { doc: rule.label, date: docExpFmt_(d), days: days,
                 status: days < 0 ? 'expired' : (days <= 60 ? 'soon' : 'ok') };
      });
      return {
        success: true,
        empId: empId,
        name: cName !== -1 ? String(data[i][cName] || '').trim() : empId,
        designation: cDesig !== -1 ? String(data[i][cDesig] || '').trim() : '',
        location: cLoc !== -1 ? String(data[i][cLoc] || '').trim() : '',
        documents: docs,
        needsAttention: docs.filter(d => d.status === 'expired' || d.status === 'missing').map(d => d.doc)
      };
    }
    return { success: false, error: 'No employee ' + empId + ' on HR Maya.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}