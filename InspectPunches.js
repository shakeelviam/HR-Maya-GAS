// ============================================================
// InspectPunches.gs — READ ONLY. Shows the complete punch record
// for specific employees on specific dates, alongside their
// location's shift times, so a correction can be decided exactly
// rather than guessed.
// ------------------------------------------------------------
// Writes nothing. No deletes, no edits.
//
// USAGE: edit INSPECT_LIST below, then run inspectPunches()
// ============================================================

// Each entry: employee ID, and the date to centre on. The tool shows
// the day before and the day after too, because an overnight shift's
// Check In sits on the previous calendar day.
const INSPECT_LIST = [
  { id: 'MT-00045', date: '02-09-2026' },   // Ram Bahadur Thing
  { id: 'MT-00016', date: '03-09-2026' },   // Dipak Lo
];

function inspectPunches() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const L = [];

  // ---- HR Maya: who they are, where, what shift ----
  const hr = ss.getSheetByName('HR Maya');
  const hd = hr.getDataRange().getValues();
  const HH = hd[0];
  const hc = (n) => HH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const hId = hc('Employee ID'), hNm = hc('Name (English)'), hSh = hc('Shift');
  let hLoc = hc('Location'); if (hLoc === -1) hLoc = hc('Branch');

  const emp = {};
  for (let i = 1; i < hd.length; i++) {
    const id = String(hd[i][hId] || '').trim();
    if (id) emp[id] = {
      name:  hNm  !== -1 ? String(hd[i][hNm]  || '').trim() : id,
      loc:   hLoc !== -1 ? String(hd[i][hLoc] || '').trim() : '',
      shift: hSh  !== -1 ? String(hd[i][hSh]  || '').trim() : '',
    };
  }

  // ---- Shifts: name, start, end ----
  const sh = ss.getSheetByName('Shifts');
  const shifts = [];
  if (sh && sh.getLastRow() > 1) {
    const SH = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const sc = (names) => {
      for (const n of names) {
        const i = SH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
        if (i !== -1) return i;
      }
      return -1;
    };
    const cNm = sc(['Shift Name','Shift','Name']), cSt = sc(['Start']),
          cEn = sc(['End']), cNh = sc(['Normal Hours','Normal Hrs']);
    sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues().forEach(r => {
      const nm = String(r[cNm] || '').trim();
      if (!nm) return;
      const fmt = (v) => {
        if (Object.prototype.toString.call(v) === '[object Date]')
          return Utilities.formatDate(v, tz_(), 'HH:mm');
        return String(v || '').trim();
      };
      shifts.push({ name: nm, start: fmt(r[cSt]), end: fmt(r[cEn]),
                    hours: cNh !== -1 ? r[cNh] : '' });
    });
  }

  // ---- Attendance_Log ----
  const log = ss.getSheetByName('Attendance_Log');
  const ld = log.getDataRange().getValues();
  const LH = ld[0];
  const lc = (n) => LH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const lId = lc('Employee ID'), lTs = lc('Timestamp'), lDate = lc('Date'),
        lStage = lc('Stage'), lAuth = lc('Auth') !== -1 ? lc('Auth') : lc('PIN Used');
  let lBr = lc('Branch'); if (lBr === -1) lBr = lc('Location');
  const lFlag = lc('Geo Flag');

  // ---- OT_Pending: did they already claim OT? ----
  const pen = ss.getSheetByName('OT_Pending');
  const otPending = {};
  if (pen && pen.getLastRow() > 1) {
    pen.getRange(2, 1, pen.getLastRow()-1, pen.getLastColumn()).getValues().forEach(r => {
      const id = String(r[1] || '').trim();
      const dt = r[3];
      const dStr = Object.prototype.toString.call(dt) === '[object Date]'
        ? formatDdMmYyyy_(dt) : String(dt || '').trim();
      if (id) otPending[id + '|' + dStr] = { hours: r[4], status: String(r[6] || '').trim() };
    });
  }

  L.push('══════════════════════════════════════════════');
  L.push(' PUNCH INSPECTION — read only, nothing changed');
  L.push('══════════════════════════════════════════════');
  L.push('');
  L.push('SHIFTS ON FILE');
  shifts.forEach(s => L.push('   ' + s.name.padEnd(28) + s.start + ' → ' + s.end +
                             '   (' + s.hours + 'h)'));
  L.push('');

  INSPECT_LIST.forEach(target => {
    const e = emp[target.id] || { name: '?', loc: '?', shift: '' };
    const centre = parseFlexibleDate_(target.date);
    const from = startOfDay_(addDays_(centre, -1));
    const to   = endOfDay_(addDays_(centre, 1));

    L.push('──────────────────────────────────────────────');
    L.push(target.id + '  ' + e.name);
    L.push('   Location : ' + e.loc);
    L.push('   Shift on HR Maya : ' + (e.shift || '(BLANK)'));
    L.push('   Shifts available at this location:');
    shifts.filter(s => s.name.toLowerCase().indexOf(e.loc.toLowerCase().split(' ')[0]) !== -1)
          .forEach(s => L.push('      ' + s.name + '  ' + s.start + ' → ' + s.end));
    L.push('');
    L.push('   ALL PUNCHES ' + formatDdMmYyyy_(from) + ' → ' + formatDdMmYyyy_(to) + ':');

    const found = [];
    for (let i = 1; i < ld.length; i++) {
      if (String(ld[i][lId] || '').trim() !== target.id) continue;
      const ts = parseLogTimestamp(ld[i][lTs]);
      if (!ts || ts < from || ts > to) continue;
      found.push({
        row: i + 1, ts: ts,
        date: lDate !== -1 ? (Object.prototype.toString.call(ld[i][lDate]) === '[object Date]'
              ? formatDdMmYyyy_(ld[i][lDate]) : String(ld[i][lDate] || '').trim()) : '',
        stage: String(ld[i][lStage] || '').trim(),
        auth: lAuth !== -1 ? String(ld[i][lAuth] || '').trim() : '',
        branch: lBr !== -1 ? String(ld[i][lBr] || '').trim() : '',
        flag: lFlag !== -1 ? String(ld[i][lFlag] || '').trim() : '',
      });
    }
    found.sort((a, b) => a.ts - b.ts);
    if (!found.length) L.push('      (none)');
    found.forEach(p => {
      L.push('      row ' + String(p.row).padEnd(5) +
             Utilities.formatDate(p.ts, tz_(), 'dd/MM HH:mm:ss') + '  ' +
             p.stage.padEnd(10) + ' date-col=' + p.date.padEnd(12) +
             p.auth.padEnd(22) + p.branch);
    });

    L.push('');
    L.push('   OT claimed for these dates:');
    let anyOt = false;
    [-1, 0, 1].forEach(off => {
      const dStr = formatDdMmYyyy_(addDays_(centre, off));
      const k = target.id + '|' + dStr;
      if (otPending[k]) {
        anyOt = true;
        L.push('      ' + dStr + '  ' + otPending[k].hours + 'h  (' + otPending[k].status + ')');
      }
    });
    if (!anyOt) L.push('      (none in OT_Pending)');
    L.push('');
  });

  L.push('══════════════════════════════════════════════');
  const out = L.join('\n');
  Logger.log(out);
  return out;
}