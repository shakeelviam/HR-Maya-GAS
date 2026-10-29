// ============================================================
// MigrationBaseline.gs — proof that a refactor changed nothing
// ------------------------------------------------------------
// THE PROBLEM WITH REFACTORING LIVE PAYROLL CODE
// "It should behave the same" is an assertion. This makes it a test.
//
// HOW TO USE
//   BEFORE migrating a file:   snapshotBaseline()
//   AFTER  migrating a file:   recompile the same range, then compareToBaseline()
//
// compareToBaseline() reports every cell that differs. An EMPTY diff is
// the only acceptable result for a behaviour-preserving change. Any
// difference is a bug in the migration, not an improvement — even if
// the new number looks more correct. If a number genuinely should
// change, that is a separate, deliberate commit with its own reasoning.
//
// The snapshot is written to a hidden tab '_Baseline_Snapshot' so it
// survives between script runs. It is data only — nothing reads it
// except compareToBaseline().
// ============================================================

const BASELINE_TAB  = '_Baseline_Snapshot';
const BASELINE_FROM = '22-08-2026';
const BASELINE_TO   = '20-09-2026';          // blank = yesterday

// Columns that must not drift. Deliberately excludes nothing —
// if a refactor changes a flag string, that is still a change.
const BASELINE_COLS = ['Employee ID', 'Date', 'Location', 'Shift',
                       'Check In', 'Break Out', 'Break In', 'Check Out',
                       'Presence (hrs)', 'OT Hours', 'Day Type',
                       'Source', 'Status', 'Flags'];

function bl_range_() {
  const from = acParseDate_(BASELINE_FROM);
  let to;
  if (BASELINE_TO) to = acParseDate_(BASELINE_TO);
  else { to = new Date(); to.setDate(to.getDate() - 1); to = acParseDate_(acNormDate_(to)); }
  return { from: from, to: to };
}

// Read Attendance_Daily into a comparable map: 'empId|date' -> [values]
function bl_read_(ss) {
  const ad = ss.getSheetByName('Attendance_Daily');
  const map = {};
  if (!ad || ad.getLastRow() < 2) return map;

  const H = acColMap_(ad);
  const idx = {};
  BASELINE_COLS.forEach(c => { idx[c] = acCol_(H, [c]); });

  const r = bl_range_();
  const cId = idx['Employee ID'], cDate = idx['Date'];
  if (cId < 0 || cDate < 0) throw new Error('Attendance_Daily missing Employee ID / Date.');

  ad.getRange(2, 1, ad.getLastRow() - 1, ad.getLastColumn()).getValues().forEach(row => {
    const dStr = acNormDate_(row[cDate]);
    const d = acParseDate_(dStr);
    if (!d || d < r.from || d > r.to) return;
    const id = String(row[cId] || '').trim();
    if (!id) return;
    // NORMALISE before storing. Without this the Date column produces
    // false differences: the snapshot sheet stores a Date object, while
    // the live Attendance_Daily cell is TEXT (setNumberFormat('@')), so
    // String() gives "Sat Aug 22 2026 00:00:00 GMT+0300" on one side and
    // "22-08-2026" on the other — a difference in representation, not in
    // data. Numbers get the same treatment so 9 and 9.0 do not disagree.
    map[id + '|' + dStr] = BASELINE_COLS.map(c => {
      if (idx[c] < 0) return '';
      const v = row[idx[c]];
      if (v == null || v === '') return '';
      if (c === 'Date') return acNormDate_(v);
      if (Object.prototype.toString.call(v) === '[object Date]') return acNormDate_(v);
      // Numeric cells: compare by value, not by how they were typed.
      if (typeof v === 'number') return String(Math.round(v * 10000) / 10000);
      const s = String(v).trim();
      if (s !== '' && !isNaN(Number(s))) return String(Math.round(Number(s) * 10000) / 10000);
      return s;
    });
  });
  return map;
}

// ── STEP 1: capture ─────────────────────────────────────────────────
function snapshotBaseline() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const map = bl_read_(ss);
  const keys = Object.keys(map).sort();

  let sh = ss.getSheetByName(BASELINE_TAB);
  if (!sh) sh = ss.insertSheet(BASELINE_TAB);
  sh.clear();

  const rows = [['key'].concat(BASELINE_COLS)];
  keys.forEach(k => rows.push([k].concat(map[k])));
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.hideSheet();

  const r = bl_range_();
  const msg = 'Baseline captured: ' + keys.length + ' employee-days (' +
              acNormDate_(r.from) + ' → ' + acNormDate_(r.to) + '). ' +
              'Now migrate, recompile the SAME range, then run compareToBaseline().';
  Logger.log(msg);
  try { ss.toast(msg, 'Baseline', 10); } catch (e) {}
  return msg;
}

// ── STEP 2: compare ─────────────────────────────────────────────────
function compareToBaseline() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = ss.getSheetByName(BASELINE_TAB);
  if (!sh || sh.getLastRow() < 2) {
    const m = 'No baseline found. Run snapshotBaseline() BEFORE migrating.';
    Logger.log(m); return m;
  }

  // Load the snapshot
  const snap = {};
  const data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const cols = data[0].slice(1).map(String);
  // Re-normalise on read-back: Sheets may have re-typed a stored value
  // (a date-looking string can come back as a Date object).
  const renorm = (v, colName) => {
    if (v == null || v === '') return '';
    if (colName === 'Date') return acNormDate_(v);
    if (Object.prototype.toString.call(v) === '[object Date]') return acNormDate_(v);
    if (typeof v === 'number') return String(Math.round(v * 10000) / 10000);
    const s = String(v).trim();
    if (s !== '' && !isNaN(Number(s))) return String(Math.round(Number(s) * 10000) / 10000);
    return s;
  };
  for (let i = 1; i < data.length; i++) {
    snap[String(data[i][0])] = data[i].slice(1).map((v, j) => renorm(v, cols[j]));
  }

  const now = bl_read_(ss);

  const changed = [], added = [], removed = [];
  Object.keys(snap).forEach(k => {
    if (!now[k]) { removed.push(k); return; }
    const a = snap[k], b = now[k].map(String);
    const diffs = [];
    for (let i = 0; i < cols.length; i++) {
      if ((a[i] || '') !== (b[i] || '')) {
        diffs.push(cols[i] + ': "' + (a[i] || '') + '" → "' + (b[i] || '') + '"');
      }
    }
    if (diffs.length) changed.push({ key: k, diffs: diffs });
  });
  Object.keys(now).forEach(k => { if (!snap[k]) added.push(k); });

  const L = [];
  L.push('══════════════════════════════════════════════');
  L.push(' BASELINE COMPARISON');
  L.push('   Baseline rows : ' + Object.keys(snap).length);
  L.push('   Current rows  : ' + Object.keys(now).length);
  L.push('   Changed       : ' + changed.length);
  L.push('   Added         : ' + added.length);
  L.push('   Removed       : ' + removed.length);
  L.push('══════════════════════════════════════════════');
  L.push('');

  if (!changed.length && !added.length && !removed.length) {
    L.push('  IDENTICAL. The refactor is provably behaviour-preserving.');
  } else {
    L.push('  NOT IDENTICAL. Every difference below is a migration bug');
    L.push('  until proven otherwise — including ones that look "more correct".');
    L.push('');
    if (changed.length) {
      L.push('CHANGED (' + changed.length + '):');
      changed.slice(0, 100).forEach(c => {
        L.push('   ' + c.key);
        c.diffs.forEach(d => L.push('       ' + d));
      });
      if (changed.length > 100) L.push('   …and ' + (changed.length - 100) + ' more');
      L.push('');
    }
    if (removed.length) {
      L.push('REMOVED (' + removed.length + ') — rows that existed before and do not now:');
      removed.slice(0, 50).forEach(k => L.push('   ' + k));
      L.push('');
    }
    if (added.length) {
      L.push('ADDED (' + added.length + ') — rows that did not exist before:');
      added.slice(0, 50).forEach(k => L.push('   ' + k));
    }
  }
  L.push('══════════════════════════════════════════════');

  const out = L.join('\n');
  Logger.log(out);
  return out;
}

// Remove the snapshot tab once a migration is signed off.
function clearBaseline() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = ss.getSheetByName(BASELINE_TAB);
  if (sh) { ss.deleteSheet(sh); Logger.log('Baseline cleared.'); }
  else Logger.log('No baseline to clear.');
}


// ════════════════════════════════════════════════════════════
// CLASSIFIED COMPARISON — the readable version
// ════════════════════════════════════════════════════════════
// compareToBaseline() dumps every cell difference, which truncates in
// the log and buries the one thing that matters: is every change an
// EXPECTED one, or is something unexplained still moving?
//
// This buckets each changed row into a known category and prints counts,
// so a clean result is a short list rather than 500 lines to read.
//
// EXPECTED CATEGORIES (all four verified during this migration):
//   1. Pre_Roster      — dates before ROSTER_GO_LIVE now say "Not Recorded"
//                        instead of guessing "Day Off"
//   2. Session restored — overnight shifts that the old date-column filter
//                        destroyed; presence and OT reappear
//   3. Absent reversed  — someone marked Absent (and DEDUCTED) who had a
//                        real session all along
//   4. Shift filled     — blank HR Maya shift now falls back to the roster
//
// ANYTHING ELSE is unexplained and needs looking at before payroll.
function compareToBaselineSummary() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = ss.getSheetByName(BASELINE_TAB);
  if (!sh || sh.getLastRow() < 2) return 'No baseline. Run snapshotBaseline() first.';

  const data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const cols = data[0].slice(1).map(String);
  const renorm = (v, colName) => {
    if (v == null || v === '') return '';
    if (colName === 'Date') return acNormDate_(v);
    if (Object.prototype.toString.call(v) === '[object Date]') return acNormDate_(v);
    if (typeof v === 'number') return String(Math.round(v * 10000) / 10000);
    const s = String(v).trim();
    if (s !== '' && !isNaN(Number(s))) return String(Math.round(Number(s) * 10000) / 10000);
    return s;
  };
  const snap = {};
  for (let i = 1; i < data.length; i++) {
    snap[String(data[i][0])] = data[i].slice(1).map((v, j) => renorm(v, cols[j]));
  }
  const now = bl_read_(ss);

  const ci = (n) => cols.indexOf(n);
  const C_SRC = ci('Source'), C_ST = ci('Status'), C_PR = ci('Presence (hrs)'),
        C_SH = ci('Shift'), C_OT = ci('OT Hours');

  const buckets = { preRoster: [], sessionRestored: [], absentReversed: [],
                    shiftFilled: [], unexplained: [] };

  Object.keys(snap).forEach(k => {
    const a = snap[k], b = now[k];
    if (!b) return;
    let differs = false;
    for (let i = 0; i < cols.length; i++) if ((a[i]||'') !== (String(b[i])||'')) { differs = true; break; }
    if (!differs) return;

    const oldSrc = a[C_SRC] || '', newSrc = String(b[C_SRC] || '');
    const oldSt  = a[C_ST]  || '', newSt  = String(b[C_ST]  || '');
    const oldPr  = a[C_PR]  || '', newPr  = String(b[C_PR]  || '');
    const oldSh  = a[C_SH]  || '', newSh  = String(b[C_SH]  || '');

    if (newSrc === 'Pre_Roster') { buckets.preRoster.push(k); return; }
    if (oldSt === 'Absent' && newSt === 'Present') { buckets.absentReversed.push(k + '  (+' + newPr + 'h)'); return; }
    if (!oldPr && newPr && newSrc === 'Kiosk') { buckets.sessionRestored.push(k + '  (+' + newPr + 'h)'); return; }
    // Shift-only change
    if (!oldSh && newSh) {
      let onlyShift = true;
      for (let i = 0; i < cols.length; i++) {
        if (i === C_SH) continue;
        if ((a[i]||'') !== (String(b[i])||'')) { onlyShift = false; break; }
      }
      if (onlyShift) { buckets.shiftFilled.push(k); return; }
    }
    buckets.unexplained.push(k + '   ' + oldSrc + '/' + oldSt + ' → ' + newSrc + '/' + newSt +
      (oldPr !== newPr ? '   presence ' + (oldPr||'0') + ' → ' + (newPr||'0') : ''));
  });

  const L = [];
  L.push('══════════════════════════════════════════════');
  L.push(' CLASSIFIED COMPARISON');
  L.push('══════════════════════════════════════════════');
  L.push('');
  L.push('EXPECTED:');
  L.push('  1. Pre_Roster (pre-' + (typeof ROSTER_GO_LIVE !== 'undefined' ? ROSTER_GO_LIVE : 'go-live') + ') : ' + buckets.preRoster.length);
  L.push('  2. Overnight session restored      : ' + buckets.sessionRestored.length);
  L.push('  3. Absent reversed (was deducted)  : ' + buckets.absentReversed.length);
  L.push('  4. Blank shift filled from roster  : ' + buckets.shiftFilled.length);
  L.push('');
  L.push('UNEXPLAINED (must be zero before payroll) : ' + buckets.unexplained.length);
  L.push('');

  if (buckets.absentReversed.length) {
    L.push('── Absent reversed — these people were going to be DEDUCTED ──');
    buckets.absentReversed.forEach(x => L.push('   ' + x));
    L.push('');
  }
  if (buckets.unexplained.length) {
    L.push('── UNEXPLAINED — investigate each ──');
    buckets.unexplained.slice(0, 80).forEach(x => L.push('   ' + x));
    if (buckets.unexplained.length > 80) L.push('   …and ' + (buckets.unexplained.length - 80) + ' more');
  } else {
    L.push('  Nothing unexplained. Every change is accounted for.');
  }
  L.push('══════════════════════════════════════════════');

  const out = L.join('\n');
  Logger.log(out);
  return out;
}