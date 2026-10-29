// ============================================================
// AutoPresentCheck.gs — READ ONLY diagnostic.
// ------------------------------------------------------------
// Answers: "are Head Office staff and Drivers actually being
// auto-marked Present, and is the 4 AM trigger doing it?"
//
// Writes nothing. Run it, read the log.
//   Apps Script editor → checkAutoPresent → Run → View → Logs
//
// WHAT IT CHECKS
//  1. Is the 4 AM dailyAttendanceTrigger installed at all?
//  2. Shifts tab: which shifts are flagged Auto Present = Yes?
//  3. HR Maya: every Head Office / Driver employee, and whether
//     the compile WOULD mark them Present.
//  4. Attendance_Daily: what actually happened for them, by date.
//
// THE TWO WAYS THIS SILENTLY BREAKS
//  - Head Office presence depends on the employee's Shift matching a
//    shift flagged Auto Present. A BLANK Shift in HR Maya means no
//    match, so they fall through to "No Roster" instead of Present.
//  - Driver presence depends on Location being EXACTLY 'Driver
//    Location'. Any case or spacing difference breaks it.
// ============================================================

function checkAutoPresent() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const L = [];
  const say = (s) => { L.push(s); };

  say('══════════════════════════════════════════════');
  say(' AUTO-PRESENT DIAGNOSTIC');
  say('══════════════════════════════════════════════');
  say('');

  // ── 1. Trigger installed? ──────────────────────────────────
  const trigs = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyAttendanceTrigger');
  say('1. THE 4 AM TRIGGER');
  if (!trigs.length) {
    say('   NOT INSTALLED. Nothing is compiling automatically.');
    say('   Fix: run setupDailyTrigger() once.');
  } else {
    say('   Installed (' + trigs.length + ' trigger' + (trigs.length > 1 ? 's — duplicates!' : '') + ').');
    if (trigs.length > 1) say('   More than one will compile the same day repeatedly. Remove the extras.');
  }
  say('');

  // ── 2. Which shifts are Auto Present? ──────────────────────
  const sh = ss.getSheetByName('Shifts');
  const autoShifts = new Set();
  say('2. SHIFTS FLAGGED "Auto Present = Yes"');
  if (!sh || sh.getLastRow() < 2) {
    say('   Shifts tab missing or empty.');
  } else {
    const SH = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const c = (names) => {
      for (const n of names) {
        const i = SH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
        if (i !== -1) return i;
      }
      return -1;
    };
    // NOTE: the real header is 'Shift Name'. The compile looks for 'Shift'
    // then 'Name' then falls back to column 0 — it works by luck only
    // because Shift Name happens to BE column 0.
    const cNm = c(['Shift Name', 'Shift', 'Name']);
    const cAu = c(['Auto Present', 'AutoPresent']);
    say('   Header used for shift name: "' + (cNm >= 0 ? SH[cNm] : 'NOT FOUND') + '"' +
        (cNm !== 0 ? '' : '  (column 0 — the compile\'s fallback also lands here)'));
    if (cAu < 0) {
      say('   No "Auto Present" column found — NOTHING can be auto-present.');
    } else {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(r => {
        const nm = String(r[cNm] || '').trim();
        if (nm && String(r[cAu] || '').trim().toLowerCase() === 'yes') {
          autoShifts.add(nm);
          say('   ✓ ' + nm);
        }
      });
      if (!autoShifts.size) say('   NONE. No shift has Auto Present = Yes.');
    }
  }
  say('');

  // ── 3. HR Maya: would each exempt employee be marked Present? ──
  const hr = ss.getSheetByName('HR Maya');
  const hd = hr.getDataRange().getValues();
  const HH = hd[0];
  const hc = (n) => HH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const hId = hc('Employee ID'), hNm = hc('Name (English)'),
        hSt = hc('Status'), hShift = hc('Shift');
  let hLoc = hc('Location'); if (hLoc === -1) hLoc = hc('Branch');

  const broken = [], ok = [];
  const locSpellings = {};

  for (let i = 1; i < hd.length; i++) {
    const id = String(hd[i][hId] || '').trim();
    if (!id) continue;
    if (hSt !== -1 && String(hd[i][hSt] || '').trim().toLowerCase() === 'inactive') continue;
    const rawLoc = hLoc !== -1 ? String(hd[i][hLoc] || '') : '';
    const loc = rawLoc.trim();
    const lower = loc.toLowerCase();
    const name = hNm !== -1 ? String(hd[i][hNm] || '').trim() : id;
    const shift = hShift !== -1 ? String(hd[i][hShift] || '').trim() : '';

    if (lower.indexOf('head office') !== -1 || lower.indexOf('driver') !== -1) {
      locSpellings['"' + rawLoc + '"'] = (locSpellings['"' + rawLoc + '"'] || 0) + 1;

      if (lower.indexOf('driver') !== -1) {
        // Driver path: EXACT match on 'Driver Location'
        if (loc === 'Driver Location') ok.push(id + '  ' + name + '  (Driver — exact match OK)');
        else broken.push(id + '  ' + name + '  DRIVER but Location is "' + rawLoc +
                         '" — compile needs EXACTLY "Driver Location"');
      } else {
        // Head Office path: needs a Shift that is flagged Auto Present
        if (!shift) {
          broken.push(id + '  ' + name + '  HEAD OFFICE but Shift is BLANK — ' +
                      'will fall through to "No Roster", not Present');
        } else if (!autoShifts.has(shift)) {
          broken.push(id + '  ' + name + '  HEAD OFFICE, Shift "' + shift +
                      '" is NOT flagged Auto Present = Yes');
        } else {
          ok.push(id + '  ' + name + '  (Head Office, shift "' + shift + '" OK)');
        }
      }
    }
  }

  say('3. HEAD OFFICE / DRIVER STAFF — would the compile mark them Present?');
  say('   Location spellings found:');
  Object.keys(locSpellings).sort().forEach(k => say('      ' + k + '  x' + locSpellings[k]));
  say('');
  say('   WILL BE MARKED PRESENT: ' + ok.length);
  ok.forEach(s => say('      ✓ ' + s));
  say('');
  say('   WILL NOT: ' + broken.length);
  broken.forEach(s => say('      ✗ ' + s));
  say('');

  // ── 4. What actually happened in Attendance_Daily ──────────
  const ad = ss.getSheetByName('Attendance_Daily');
  say('4. WHAT ACTUALLY HAPPENED (Attendance_Daily)');
  if (!ad || ad.getLastRow() < 2) {
    say('   Attendance_Daily is empty.');
  } else {
    const add = ad.getDataRange().getValues();
    const AH = add[0];
    const ac = (n) => AH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const aId = ac('Employee ID'), aDate = ac('Date'),
          aSt = ac('Status'), aSrc = ac('Source'), aLoc = ac('Location');

    const bySource = {}, byStatusExempt = {};
    let exemptRows = 0;
    for (let i = 1; i < add.length; i++) {
      const loc = aLoc !== -1 ? String(add[i][aLoc] || '').trim().toLowerCase() : '';
      const src = aSrc !== -1 ? String(add[i][aSrc] || '').trim() : '';
      const st  = aSt  !== -1 ? String(add[i][aSt]  || '').trim() : '';
      bySource[src || '(blank)'] = (bySource[src || '(blank)'] || 0) + 1;
      if (loc.indexOf('head office') !== -1 || loc.indexOf('driver') !== -1) {
        exemptRows++;
        byStatusExempt[st || '(blank)'] = (byStatusExempt[st || '(blank)'] || 0) + 1;
      }
    }
    say('   All rows by Source:');
    Object.keys(bySource).sort().forEach(k => say('      ' + k + ': ' + bySource[k]));
    say('');
    say('   Head Office / Driver rows: ' + exemptRows);
    Object.keys(byStatusExempt).sort().forEach(k =>
      say('      ' + k + ': ' + byStatusExempt[k] +
          (k === 'No Roster' ? '   ← WRONG, should be Present' : '')));
    if (!bySource['Auto_HeadOffice']) say('');
    if (!bySource['Auto_HeadOffice']) say('   Auto_HeadOffice appears 0 times — Head Office auto-present is NOT firing.');
    if (!bySource['Driver_Default'])  say('   Driver_Default appears 0 times — Driver auto-present is NOT firing.');
  }

  say('');
  say('══════════════════════════════════════════════');
  const out = L.join('\n');
  Logger.log(out);
  return out;
}