// ============================================================
// DiagnoseSession.gs — READ ONLY. Why did a session vanish?
// ------------------------------------------------------------
// compareToBaseline() showed kiosk sessions disappearing:
//   MT-00002 04-09  Presence 8.294 → ""   Kiosk → Pre_Roster
//   MT-00002 05-09  Presence 14.04 → ""
//   MT-00009 05-09  Presence 14.003 → ""
// and one appearing:
//   MT-00007 04-09  Default → Kiosk
//
// This shows the raw punches the compile sees for those employee-days,
// across a 3-day window, so the cause is visible rather than guessed.
//
// Writes nothing.
// ============================================================

const DIAG_TARGETS = [
  // MT-00031 Lakpa Dorje Lama — Qurain Second Shift 18:30 → 03:30.
  // EVERY session crosses midnight, so each checkout is dated the day
  // AFTER its check-in. Presence dropped 15.054 → 9.804 on 03-09 after
  // the window fix, and a manual repair hit the duplicate-stage guard
  // because 04-09 already had a checkout. Walk the whole run.
  { id: 'MT-00031', date: '01-09-2026' },
  { id: 'MT-00031', date: '03-09-2026' },
  { id: 'MT-00031', date: '05-09-2026' },
];

function diagnoseSessions() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const log = ss.getSheetByName('Attendance_Log');
  const L = [];

  const H = acColMap_(log);
  const cId    = acCol_(H, ['Employee ID']);
  const cTs    = acCol_(H, ['Timestamp']);
  const cDate  = acCol_(H, ['Date']);
  const cStage = acCol_(H, ['Stage']);
  const cAuth  = acCol_(H, ['Auth', 'PIN Used']);
  const cBr    = acCol_(H, ['Branch', 'Location']);

  L.push('══════════════════════════════════════════════');
  L.push(' SESSION DIAGNOSTIC — read only');
  L.push('══════════════════════════════════════════════');
  L.push('');
  L.push('KEY QUESTION: writeDayToAttendanceDaily_ selects punches by the');
  L.push('DATE COLUMN, not by the timestamp. If a Check Out was written with');
  L.push('a Date cell of the NEXT day (normal for an overnight shift), then');
  L.push('building the earlier date sees only a Check In and finds no');
  L.push('complete session. Compare "date-col" against the timestamp below.');
  L.push('');

  const data = log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn()).getValues();

  DIAG_TARGETS.forEach(t => {
    const centre = acParseDate_(t.date);
    // Wider window: this employee's sessions chain across days, so a
    // one-day view hides the pattern.
    const from = new Date(centre.getTime() - 2 * 86400000);
    const to   = new Date(centre.getTime() + 3 * 86400000);

    L.push('──────────────────────────────────────────────');
    L.push(t.id + '  around ' + t.date + '  (' + acEmpName_(ss, t.id) + ')');
    L.push('');

    const rows = [];
    data.forEach((r, i) => {
      if (String(r[cId] || '').trim() !== t.id) return;
      const ts = acParseTs_(r[cTs]);
      if (!ts || ts < from || ts > to) return;
      rows.push({
        row: i + 2, ts: ts,
        dateCol: cDate >= 0 ? acNormDate_(r[cDate]) : '',
        stage: String(r[cStage] || '').trim(),
        auth: cAuth >= 0 ? String(r[cAuth] || '').trim() : '',
        branch: cBr >= 0 ? String(r[cBr] || '').trim() : '',
      });
    });
    rows.sort((a, b) => a.ts - b.ts);

    if (!rows.length) { L.push('   (no punches in window)'); L.push(''); return; }

    rows.forEach(p => {
      const tsDate = acNormDate_(p.ts);
      const mismatch = (p.dateCol && p.dateCol !== tsDate) ? '   <<< date-col ≠ timestamp date' : '';
      L.push('   row ' + String(p.row).padEnd(5) +
             Utilities.formatDate(p.ts, acTz_(), 'dd/MM HH:mm:ss') + '  ' +
             p.stage.padEnd(10) + ' date-col=' + p.dateCol.padEnd(12) +
             p.auth.padEnd(22) + p.branch + mismatch);
    });

    // What the CORE makes of it
    const s = acSessionize_(rows.map(p => ({ ts: p.ts, stage: p.stage, row: p.row, location: p.branch })));
    L.push('');
    L.push('   Core sessionizer over the whole window:');
    L.push('      completed: ' + s.completed.length + '   orphaned: ' + s.orphaned.length +
           '   open: ' + (s.open ? 'yes' : 'no') + '   dayOff: ' + (s.dayOff ? 'yes' : 'no'));
    s.completed.forEach(c => {
      const p = acPresence_(c);
      L.push('      session ' + Utilities.formatDate(c.checkIn, acTz_(), 'dd/MM HH:mm') +
             ' → ' + Utilities.formatDate(c.checkOut, acTz_(), 'dd/MM HH:mm') +
             '   presence ' + p.presence.toFixed(2) + 'h' +
             (p.suspect ? '  SUSPECT' : ''));
    });

    // What the COMPILE sees — punches filtered to the single date column
    const sameDate = rows.filter(p => p.dateCol === t.date);
    L.push('');
    L.push('   Punches whose DATE COLUMN = ' + t.date + ': ' + sameDate.length);
    if (sameDate.length) {
      L.push('      ' + sameDate.map(p => p.stage).join(', '));
      const s2 = acSessionize_(sameDate.map(p => ({ ts: p.ts, stage: p.stage, row: p.row, location: p.branch })));
      L.push('      → completed: ' + s2.completed.length +
             (s2.completed.length ? '' : '   ← THIS is why the compile finds nothing'));
    }
    L.push('');
  });

  L.push('══════════════════════════════════════════════');
  const out = L.join('\n');
  Logger.log(out);
  return out;
}