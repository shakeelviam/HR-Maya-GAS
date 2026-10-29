// ============================================================
// DateOfBirth.gs — Date of Birth, read out of the Civil ID
// ------------------------------------------------------------
// A Kuwait Civil ID carries the birthday in its first seven
// digits:
//
//     2 76 08 24 03264
//     |  |  |  |  |
//     |  |  |  |  +-- serial + check digit (nothing to do with us)
//     |  |  |  +----- day
//     |  |  +-------- month
//     |  +----------- year, two digits
//     +-------------- century: 2 = born 19xx, 3 = born 20xx
//
// so 276082403264 is 24-08-1976.
//
// The column is DERIVED, never typed. Correct a Civil ID and run
// refreshDateOfBirth() and the date follows. A new man gets his the
// moment his Civil ID is entered.
//
// Two functions, both safe to run from the editor with no arguments:
//
//   addDateOfBirthColumn()  — adds the column before Civil ID (once),
//                             then fills it. Safe to run twice.
//   refreshDateOfBirth()    — fills it again from today's Civil IDs.
//
// A row whose Civil ID cannot give a date is LEFT ALONE, never
// blanked — so anything entered by hand survives a refresh — and the
// reason is reported.
// ============================================================

const DOB_STAFF_TAB    = 'HR Maya';
const DOB_HEADER       = 'Date of Birth';
const DOB_CIVIL_HEADER = 'Civil ID';
const DOB_ID_HEADER    = 'Employee ID';
const DOB_NAME_HEADER  = 'Name (English)';

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------

function dobPad_(n) { return (n < 10 ? '0' : '') + n; }

// Header position by NAME, never by a fixed letter or number, so the
// column can sit anywhere and be moved later without touching code.
function dobCol_(headers, name) {
  const want = String(name).trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === want) return i;
  }
  return -1;
}

// A Civil ID may arrive as text or as a number, and Sheets likes to
// hand back numbers with a trailing ".0". Keep the digits only.
function dobDigits_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim();
  if (!s) return '';
  if (/e\+?\d+$/i.test(s)) {                 // 3.02082e+11 — a number that lost its digits
    const n = Number(s);
    if (!isFinite(n)) return '';
    s = n.toFixed(0);
  }
  s = s.replace(/\.\d*$/, '');               // 302082202776.0 — a number with a decimal tail
  return s.replace(/\D/g, '');
}

// The rule itself. Returns { ok, text, date } or { ok:false, reason }.
function dobFromCivilId_(civilId) {
  const s = dobDigits_(civilId);
  if (!s) return { ok: false, reason: 'no Civil ID on the row' };
  if (s.length !== 12) return { ok: false, reason: 'Civil ID is ' + s.length + ' digits, not 12' };

  const prefix = s.charAt(0);
  if (prefix !== '2' && prefix !== '3') {
    return { ok: false, reason: 'Civil ID starts with ' + prefix + ' — expected 2 (born 19xx) or 3 (born 20xx)' };
  }

  const year  = (prefix === '2' ? 1900 : 2000) + Number(s.substring(1, 3));
  const month = Number(s.substring(3, 5));
  const day   = Number(s.substring(5, 7));

  if (month < 1 || month > 12) return { ok: false, reason: 'Civil ID gives month ' + month };
  if (day < 1 || day > 31)     return { ok: false, reason: 'Civil ID gives day ' + day };

  // Catches 31-02 and the like: the Date object silently rolls over,
  // so compare it back against what we asked for.
  const d = new Date(year, month - 1, day);
  if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) {
    return { ok: false, reason: 'Civil ID gives ' + dobPad_(day) + '-' + dobPad_(month) + '-' + year + ', which is not a real date' };
  }

  return { ok: true, date: d, text: dobPad_(day) + '-' + dobPad_(month) + '-' + year };
}

// ------------------------------------------------------------
// Data validation
// ------------------------------------------------------------
// The Civil ID column carries a rule ("must be 12 characters"). When a
// column is inserted in front of it, Sheets can leave that rule sitting
// on the column letter rather than following the data — so the rule
// ends up guarding Date of Birth, and Civil ID is left unguarded. That
// rejects every date we write AND quietly removes the check that was
// protecting the Civil IDs.
//
// So: take any rule off Date of Birth, and if Civil ID has lost its
// rule, put that same rule back where it belongs.
function dobFixValidation_(sh, iDob, iCiv) {
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  const dobBody = sh.getRange(2, iDob + 1, rows, 1);
  const civBody = sh.getRange(2, iCiv + 1, rows, 1);

  let dobRule = null, civRule = null;
  try { dobRule = dobBody.getDataValidations()[0][0] || null; } catch (e) { dobRule = null; }
  try { civRule = civBody.getDataValidations()[0][0] || null; } catch (e) { civRule = null; }

  const out = { hadRuleOnDob: !!dobRule, restoredOnCivilId: false, civilIdUnguarded: false };

  if (dobRule && !civRule) {
    civBody.setDataValidation(dobRule);            // put it back on Civil ID
    out.restoredOnCivilId = true;
  } else if (!dobRule && !civRule) {
    out.civilIdUnguarded = true;                   // nothing to move — say so
  }

  // Date of Birth is written by this script, so it never wants a rule.
  sh.getRange(1, iDob + 1, sh.getMaxRows(), 1).clearDataValidations();
  return out;
}

// ------------------------------------------------------------
// Fill — shared by both entry points
// ------------------------------------------------------------
function dobFill_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { filled: 0, kept: 0, problems: [] };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const iDob  = dobCol_(headers, DOB_HEADER);
  const iCiv  = dobCol_(headers, DOB_CIVIL_HEADER);
  const iId   = dobCol_(headers, DOB_ID_HEADER);
  const iName = dobCol_(headers, DOB_NAME_HEADER);
  if (iDob < 0) throw new Error(DOB_STAFF_TAB + ' has no "' + DOB_HEADER + '" column — run addDateOfBirthColumn() first.');
  if (iCiv < 0) throw new Error(DOB_STAFF_TAB + ' has no "' + DOB_CIVIL_HEADER + '" column.');

  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = [], problems = [];
  let filled = 0, kept = 0;

  rows.forEach(function (r) {
    const existing = (iDob < r.length && r[iDob] !== undefined && r[iDob] !== null) ? r[iDob] : '';
    const empId = String(r[iId] === undefined ? '' : r[iId]).trim();

    // Spacer or summary rows: leave exactly as they are.
    if (!empId) { out.push([existing]); return; }

    const got = dobFromCivilId_(r[iCiv]);
    if (got.ok) { out.push([got.text]); filled++; return; }

    // Cannot derive it — keep whatever is in the cell (a hand-typed
    // date must survive) and say why.
    out.push([existing]);
    kept++;
    problems.push(empId + '  ' + String(r[iName] === undefined ? '' : r[iName]).trim() + ' — ' + got.reason +
                  (existing ? ' (left the existing "' + existing + '" alone)' : ''));
  });

  sh.getRange(2, iDob + 1, out.length, 1).setValues(out);
  return { filled: filled, kept: kept, problems: problems };
}

// ------------------------------------------------------------
// Add the column (once), then fill it
// ------------------------------------------------------------
function addDateOfBirthColumn() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = ss.getSheetByName(DOB_STAFF_TAB);
    if (!sh) return dobReport_({ success: false, error: DOB_STAFF_TAB + ' tab not found.' });

    let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const iCiv = dobCol_(headers, DOB_CIVIL_HEADER);
    if (iCiv < 0) return dobReport_({ success: false, error: DOB_STAFF_TAB + ' has no "' + DOB_CIVIL_HEADER + '" column.' });

    let added = false, dobAt, civAt;
    if (dobCol_(headers, DOB_HEADER) < 0) {
      sh.insertColumnBefore(iCiv + 1);                       // Date of Birth takes the Civil ID's old place
      dobAt = iCiv;                                          // ...and Civil ID moves one to the right
      civAt = iCiv + 1;
      added = true;
    } else {
      dobAt = dobCol_(headers, DOB_HEADER);
      civAt = dobCol_(headers, DOB_CIVIL_HEADER);
    }

    // BEFORE writing anything into the column. A rule left behind on the
    // column letter rejects the header text as readily as it rejects a
    // date, so it has to go first.
    const val = dobFixValidation_(sh, dobAt, civAt);

    if (added) {
      sh.getRange(1, dobAt + 1).setValue(DOB_HEADER).setFontWeight('bold');
      // Plain text, so Sheets does not turn 14-02-2001 into a date
      // object and print it back in some other format.
      sh.getRange(1, dobAt + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    }

    const res = dobFill_(sh);
    let note = '';
    if (val.restoredOnCivilId) note = ' The "12 characters" rule had moved onto the new column — put back on ' + DOB_CIVIL_HEADER + '.';
    else if (val.hadRuleOnDob) note = ' Cleared a leftover validation rule from the ' + DOB_HEADER + ' column.';
    else if (val.civilIdUnguarded) note = ' NOTE: ' + DOB_CIVIL_HEADER + ' has no validation rule on it — re-add it by hand if you want the 12-character check back.';

    return dobReport_({
      success: true,
      added: added,
      filled: res.filled,
      kept: res.kept,
      validation: val,
      problems: res.problems,
      message: (added ? 'Added "' + DOB_HEADER + '" before ' + DOB_CIVIL_HEADER + '. ' : '"' + DOB_HEADER + '" was already there. ') +
               res.filled + ' date(s) of birth written from the Civil ID.' +
               (res.kept ? ' ' + res.kept + ' row(s) could not be worked out — see below.' : '') + note
    });
  } catch (err) {
    return dobReport_({ success: false, error: err.toString() });
  }
}

// ------------------------------------------------------------
// Fill it again — after a Civil ID is corrected or a new man joins
// ------------------------------------------------------------
function refreshDateOfBirth() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = ss.getSheetByName(DOB_STAFF_TAB);
    if (!sh) return dobReport_({ success: false, error: DOB_STAFF_TAB + ' tab not found.' });

    // Never let a stray validation rule on the column block the write.
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const dobAt = dobCol_(headers, DOB_HEADER);
    const civAt = dobCol_(headers, DOB_CIVIL_HEADER);
    if (dobAt < 0) return dobReport_({ success: false, error: DOB_STAFF_TAB + ' has no "' + DOB_HEADER + '" column — run addDateOfBirthColumn() first.' });
    if (civAt >= 0) dobFixValidation_(sh, dobAt, civAt);

    const res = dobFill_(sh);
    return dobReport_({
      success: true,
      added: false,
      filled: res.filled,
      kept: res.kept,
      problems: res.problems,
      message: res.filled + ' date(s) of birth refreshed from the Civil ID.' +
               (res.kept ? ' ' + res.kept + ' row(s) could not be worked out — see below.' : '')
    });
  } catch (err) {
    return dobReport_({ success: false, error: err.toString() });
  }
}

// Print to the editor log as well as returning, because the Run
// button shows the log and not the return value.
function dobReport_(r) {
  if (r.success) {
    Logger.log(r.message);
    (r.problems || []).forEach(function (p) { Logger.log('   ' + p); });
  } else {
    Logger.log('FAILED: ' + r.error);
  }
  return r;
}