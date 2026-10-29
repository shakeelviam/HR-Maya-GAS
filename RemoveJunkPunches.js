// ============================================================
// RemoveJunkPunches.gs — remove the junk punches identified by
// classifyRapidPunches on 02–03 Sept 2026.
// ------------------------------------------------------------
// This tool does NOT search for anything. Every row it touches is
// listed explicitly below, with the employee, timestamp and stage it
// is expected to contain.
//
// WHY HARDCODED
// Row numbers shift the moment anything is deleted or inserted. A tool
// that re-detects at delete time can act on a different row than the
// one you reviewed. Instead every target is pinned here AND verified
// against its expected employee + stage + timestamp before removal.
// If a row does not match what is expected, that row is SKIPPED and
// reported — it never guesses.
//
// ORDER OF OPERATIONS
//   1. verifyJunkRows()  — checks every target still matches. No writes.
//   2. removeJunkRows()  — re-verifies, then deletes highest row first.
//
// AFTER RUNNING: recompile 01-09, 02-09 and 03-09.
// ============================================================

// Each entry pins a row AND what must be in it.
// Sourced directly from the classifyRapidPunches output.
const JUNK_TARGETS = [

  // ── BATCH 2 — the four partial bursts.
  // Row numbers REGENERATED after batch 1 deleted 17 rows and shifted
  // everything below. Each of these keeps a genuine session boundary,
  // which is exactly why they were held back rather than deleted blind.
  //
  //   MT-00001  Check In 20:11:39 → Break Out 20:11:45   keep row 276
  //   MT-00010  Check In 17:39:25 → Break Out 17:40:11   keep row 231
  //   MT-00013  Check In 18:56:56 → Break Out 18:57:11   keep row 250
  //   MT-00039  Break In 18:28:46 → Check Out 18:29:18   keep row 237

  { row: 277, id: 'MT-00001', stage: 'Break Out', time: '03/09 20:11:45', why: 'Accidental tap 6s after a REAL Check In (row 276 must survive)' },
  { row: 232, id: 'MT-00010', stage: 'Break Out', time: '03/09 17:40:11', why: 'Accidental tap 46s after a REAL Check In (row 231 must survive)' },
  { row: 251, id: 'MT-00013', stage: 'Break Out', time: '03/09 18:57:11', why: 'Accidental tap 15s after a REAL Check In (row 250 must survive)' },
  { row: 236, id: 'MT-00039', stage: 'Break In',  time: '03/09 18:28:46', why: 'Accidental tap 32s before a REAL Check Out (row 237 must survive)' },
];

// DELIBERATELY NOT INCLUDED:
//   Batch 1 (17 rows) — already deleted.
//   MT-00016 Dipak Lo — already corrected with the supervisor.
//   MT-00027 Kapil Koirala rows 77/78 and 155-158 — PARKED. Deleting these
//     leaves him with no attendance at all on 02-09 and 03-09, so a decision
//     is needed on what he actually worked before anything is removed.
//   MT-00041 Pujan Bhattarai rows 159-162 — PARKED, same reason.

function rj_cols_(sheet) {
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const c = (names) => {
    for (const n of names) {
      const i = H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
      if (i !== -1) return i;
    }
    return -1;
  };
  return { ts: c(['Timestamp']), id: c(['Employee ID']), name: c(['Employee Name']),
           date: c(['Date']), stage: c(['Stage']), auth: c(['Auth','PIN Used']) };
}

// Check every target against what is actually in the sheet right now.
function rj_verify_(ss) {
  const log = ss.getSheetByName('Attendance_Log');
  const C = rj_cols_(log);
  const last = log.getLastRow();
  const ok = [], bad = [];

  JUNK_TARGETS.forEach(t => {
    if (t.row < 2 || t.row > last) {
      bad.push({ t: t, found: 'row ' + t.row + ' is outside the sheet (last row ' + last + ')' });
      return;
    }
    const r = log.getRange(t.row, 1, 1, log.getLastColumn()).getValues()[0];
    const id = String(r[C.id] || '').trim();
    const stage = String(r[C.stage] || '').trim();
    const ts = parseLogTimestamp(r[C.ts]);
    const tsStr = ts ? Utilities.formatDate(ts, tz_(), 'dd/MM HH:mm:ss') : '(unparseable)';
    const name = C.name !== -1 ? String(r[C.name] || '').trim() : '';

    if (id === t.id && stage === t.stage && tsStr === t.time) {
      ok.push({ t: t, name: name, tsStr: tsStr });
    } else {
      bad.push({ t: t, found: 'row ' + t.row + ' contains ' + (id || '(blank)') + ' ' +
                 (stage || '(blank)') + ' ' + tsStr + '  — EXPECTED ' + t.id + ' ' +
                 t.stage + ' ' + t.time });
    }
  });
  return { ok: ok, bad: bad };
}

// ── STEP 1: verify only. Writes nothing. ────────────────────────────
function verifyJunkRows() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const v = rj_verify_(ss);
  const L = [];
  L.push('══════════════════════════════════════════════');
  L.push(' VERIFY JUNK ROWS — nothing deleted');
  L.push(' Targets: ' + JUNK_TARGETS.length);
  L.push('══════════════════════════════════════════════');
  L.push('');
  L.push('MATCHES EXPECTED (' + v.ok.length + ') — safe to remove:');
  v.ok.forEach(x => L.push('   row ' + String(x.t.row).padEnd(5) + x.t.id + '  ' +
                           (x.name || '').padEnd(24) + x.t.stage.padEnd(11) + x.tsStr +
                           '   ' + x.t.why));
  L.push('');
  if (v.bad.length) {
    L.push('DOES NOT MATCH (' + v.bad.length + ') — will be SKIPPED:');
    v.bad.forEach(x => L.push('   ' + x.found));
    L.push('');
    L.push('   Row numbers shift when rows are added or deleted. If these are');
    L.push('   wrong, re-run classifyRapidPunches() and update JUNK_TARGETS.');
  } else {
    L.push('All targets match. Run removeJunkRows() to delete them.');
  }
  L.push('══════════════════════════════════════════════');
  const out = L.join('\n');
  Logger.log(out);
  return out;
}

// ── STEP 2: delete. Re-verifies first; skips anything that moved. ──
function removeJunkRows() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const log = ss.getSheetByName('Attendance_Log');
  const v = rj_verify_(ss);

  if (!v.ok.length) {
    Logger.log('Nothing matched — nothing deleted. Run verifyJunkRows() to see why.');
    return { success: false, deleted: 0, skipped: v.bad.length };
  }

  // Audit BEFORE deleting, while the rows still exist.
  v.ok.forEach(x => {
    if (typeof logAmend_ === 'function') {
      logAmend_(ss, 'System (junk punch cleanup)', 'DELETE', x.t.id,
                x.tsStr.split(' ')[0], x.t.stage, x.t.why);
    }
  });

  // Highest row first so earlier indexes stay valid.
  const rows = v.ok.map(x => x.t.row).sort((a, b) => b - a);
  rows.forEach(r => log.deleteRow(r));

  const msg = 'Deleted ' + rows.length + ' junk row(s). Skipped ' + v.bad.length +
    ' that did not match. NOW RECOMPILE: 01-09-2026, 02-09-2026, 03-09-2026.';
  Logger.log(msg);
  if (v.bad.length) v.bad.forEach(x => Logger.log('   SKIPPED: ' + x.found));
  try { ss.toast(msg, 'Junk punch cleanup', 12); } catch (e) {}
  return { success: true, deleted: rows.length, skipped: v.bad.length, message: msg };
}