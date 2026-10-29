// ============================================================
// RapidPunchCleanup.gs — classify junk kiosk punches
// ------------------------------------------------------------
// THE PROBLEM
// On 02→03 Sept a run of punches was recorded seconds apart across
// many employees — someone tapping rapidly through the kiosk. Example:
//   MT-00015  Break Out 00:38:54 → Break In 00:39:18   (24s "break")
//   MT-00027  Check In 02:47:57 → ... → Check Out 02:48:26  (29s "shift")
//
// WHY A BLANKET RULE IS WRONG
// An earlier version of this tool proposed deleting every punch in a
// rapid burst. That would have destroyed real data:
//   MT-00001  Check In 20:11:39 → Break Out 20:11:45
// The Check In is genuine — they started work. Only the Break Out, six
// seconds later, is an accidental tap. Deleting both leaves the person
// with no session at all.
//
// THE RULE THIS VERSION USES — per burst, by shape:
//   Break Out → Break In        both junk      phantom break
//   Check In  → Check Out       both junk      phantom session (seconds long)
//   Check In  → Break Out       drop Break Out KEEP the Check In
//   Break In  → Check Out       drop Break In  KEEP the Check Out
//   anything else               NOTHING        flagged for a human
//
// SAFETY
//   - This file DELETES NOTHING. There is no commit function.
//     It only classifies and reports. Removal is a separate, later step
//     once you have reviewed the classification.
//   - Only Auth='Verified' bursts are considered. Any burst containing a
//     'Supervisor corrected' or 'Manual' row is skipped — those are
//     deliberate human entries.
//   - PROTECTED_IDS are reported separately and never included.
//
// USAGE:  classifyRapidPunches()   → read the log
// ============================================================

const RAPID_GAP_SECONDS = 120;        // 2 minutes
const RAPID_FROM = '02-09-2026';
const RAPID_TO   = '03-09-2026';

// Parked for a separate decision: these have ONLY 'Verified' punches on
// the affected dates and no Supervisor-corrected version, so any removal
// leaves them with no attendance at all that day.
const PROTECTED_IDS = ['MT-00027', 'MT-00041'];

function rp_cols_(sheet) {
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const c = (names) => {
    for (const n of names) {
      const i = H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
      if (i !== -1) return i;
    }
    return -1;
  };
  return {
    ts: c(['Timestamp']), id: c(['Employee ID']), name: c(['Employee Name']),
    date: c(['Date']), stage: c(['Stage']), auth: c(['Auth', 'PIN Used']),
    branch: c(['Branch', 'Location']),
  };
}

function rp_normDate_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[1];
  const d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}

// Decide what, if anything, is junk inside one burst.
// Returns { verdict, dropRows[], keepRows[], reason }
function rp_classify_(burst) {
  const stages = burst.map(x => x.stage);
  const seq = stages.join(' → ');

  const isPair = (a, b) => burst.length === 2 && stages[0] === a && stages[1] === b;

  if (isPair('Break Out', 'Break In')) {
    return { verdict: 'DROP BOTH', dropRows: burst.map(x => x.row), keepRows: [],
             reason: 'Phantom break — a break of ' + rp_span_(burst) + 's contributes nothing' };
  }
  if (isPair('Check In', 'Check Out')) {
    return { verdict: 'DROP BOTH', dropRows: burst.map(x => x.row), keepRows: [],
             reason: 'Phantom session — a whole shift in ' + rp_span_(burst) + 's is not real' };
  }
  if (isPair('Check In', 'Break Out')) {
    return { verdict: 'DROP ONE', dropRows: [burst[1].row], keepRows: [burst[0].row],
             reason: 'Check In is REAL — keep it. The Break Out ' + rp_span_(burst) +
                     's later is an accidental tap' };
  }
  if (isPair('Break In', 'Check Out')) {
    return { verdict: 'DROP ONE', dropRows: [burst[0].row], keepRows: [burst[1].row],
             reason: 'Check Out is REAL — keep it. The Break In ' + rp_span_(burst) +
                     's earlier is an accidental tap' };
  }
  if (isPair('Break Out', 'Check Out')) {
    return { verdict: 'DROP ONE', dropRows: [burst[0].row], keepRows: [burst[1].row],
             reason: 'Check Out is REAL — keep it. The unclosed Break Out is an accidental tap' };
  }
  if (isPair('Check In', 'Break In')) {
    return { verdict: 'DROP ONE', dropRows: [burst[1].row], keepRows: [burst[0].row],
             reason: 'Check In is REAL — keep it. A Break In with no Break Out is an accidental tap' };
  }

  // Full Check In → Break → Check Out cycle in seconds: entirely phantom.
  if (burst.length === 4 && seq === 'Check In → Break Out → Break In → Check Out') {
    return { verdict: 'DROP ALL', dropRows: burst.map(x => x.row), keepRows: [],
             reason: 'Complete shift cycle in ' + rp_span_(burst) + 's — entirely phantom' };
  }

  return { verdict: 'REVIEW', dropRows: [], keepRows: burst.map(x => x.row),
           reason: 'Unrecognised pattern (' + seq + ') — a human should decide' };
}

function rp_span_(burst) {
  return Math.round((burst[burst.length-1].ts - burst[0].ts) / 1000);
}

function rp_findBursts_(ss) {
  const log = ss.getSheetByName('Attendance_Log');
  if (!log || log.getLastRow() < 2) return { bursts: [], protectedBursts: [] };

  const C = rp_cols_(log);
  const data = log.getDataRange().getValues();
  const from = startOfDay_(parseFlexibleDate_(RAPID_FROM));
  const to   = endOfDay_(parseFlexibleDate_(RAPID_TO));

  const byEmp = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][C.id] || '').trim();
    if (!id) continue;
    const ts = parseLogTimestamp(data[i][C.ts]);
    if (!ts || ts < from || ts > to) continue;
    (byEmp[id] = byEmp[id] || []).push({
      row: i + 1, ts: ts,
      auth: C.auth !== -1 ? String(data[i][C.auth] || '').trim() : '',
      name: C.name !== -1 ? String(data[i][C.name] || '').trim() : id,
      date: rp_normDate_(data[i][C.date]),
      stage: String(data[i][C.stage] || '').trim(),
      branch: C.branch !== -1 ? String(data[i][C.branch] || '').trim() : '',
    });
  }

  const bursts = [], protectedBursts = [];
  Object.keys(byEmp).forEach(id => {
    const p = byEmp[id].sort((a, b) => a.ts - b.ts);
    let run = [p[0]];
    const runs = [];
    for (let i = 1; i < p.length; i++) {
      if ((p[i].ts - p[i-1].ts) / 1000 <= RAPID_GAP_SECONDS) run.push(p[i]);
      else { if (run.length > 1) runs.push(run); run = [p[i]]; }
    }
    if (run.length > 1) runs.push(run);

    runs.forEach(r => {
      if (!r.every(x => x.auth === 'Verified')) return;   // human entries untouched
      const cls = rp_classify_(r);
      const entry = {
        id: id, name: r[0].name, date: r[0].date, branch: r[0].branch,
        span: rp_span_(r),
        detail: r.map(x => x.stage + ' ' + Utilities.formatDate(x.ts, tz_(), 'HH:mm:ss')).join('  →  '),
        rows: r.map(x => x.row),
        verdict: cls.verdict, dropRows: cls.dropRows,
        keepRows: cls.keepRows, reason: cls.reason,
      };
      if (PROTECTED_IDS.indexOf(id) !== -1) protectedBursts.push(entry);
      else bursts.push(entry);
    });
  });

  bursts.sort((a,b) => (a.date + a.name).localeCompare(b.date + b.name));
  protectedBursts.sort((a,b) => (a.date + a.name).localeCompare(b.date + b.name));
  return { bursts: bursts, protectedBursts: protectedBursts };
}

// ── CLASSIFY — reports only. Deletes nothing. ───────────────────────
function classifyRapidPunches() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const f = rp_findBursts_(ss);
  const L = [];

  L.push('══════════════════════════════════════════════');
  L.push(' RAPID PUNCH CLASSIFICATION');
  L.push(' THIS TOOL DELETES NOTHING. Report only.');
  L.push(' Window: ' + RAPID_FROM + ' → ' + RAPID_TO +
         '   Rule: Verified punches ≤ ' + RAPID_GAP_SECONDS + 's apart');
  L.push('══════════════════════════════════════════════');
  L.push('');

  const group = (v) => f.bursts.filter(b => b.verdict === v);
  const both   = group('DROP BOTH').concat(group('DROP ALL'));
  const one    = group('DROP ONE');
  const review = group('REVIEW');

  let dropCount = 0;
  f.bursts.forEach(b => dropCount += b.dropRows.length);

  L.push('SUMMARY');
  L.push('  Bursts found                  : ' + f.bursts.length);
  L.push('  Rows that are junk            : ' + dropCount);
  L.push('  Bursts where BOTH are junk    : ' + both.length);
  L.push('  Bursts where ONE is real      : ' + one.length + '   ← do NOT delete the whole burst');
  L.push('  Bursts needing a human        : ' + review.length);
  L.push('  Protected (parked)            : ' + f.protectedBursts.length);
  L.push('');

  const dump = (title, arr) => {
    L.push('──────────────────────────────────────────────');
    L.push(title + '  (' + arr.length + ')');
    L.push('');
    arr.forEach(b => {
      L.push('  ' + b.id + '  ' + b.name + '  ' + b.date + '  (' + b.branch + ')  ' + b.span + 's');
      L.push('     ' + b.detail);
      L.push('     ' + b.reason);
      if (b.dropRows.length) L.push('     JUNK rows : ' + b.dropRows.join(', '));
      if (b.keepRows.length) L.push('     KEEP rows : ' + b.keepRows.join(', '));
      L.push('');
    });
  };

  if (both.length)   dump('A. ENTIRE BURST IS JUNK — safe to remove all rows', both);
  if (one.length)    dump('B. PART OF THE BURST IS REAL — remove ONLY the junk row', one);
  if (review.length) dump('C. NEEDS A HUMAN DECISION', review);
  if (f.protectedBursts.length) dump('D. PROTECTED — parked, not classified for removal', f.protectedBursts);

  L.push('══════════════════════════════════════════════');
  L.push(' Nothing has been changed. Review sections A and B,');
  L.push(' then decide what to remove.');
  L.push('══════════════════════════════════════════════');

  const out = L.join('\n');
  Logger.log(out);
  return out;
}