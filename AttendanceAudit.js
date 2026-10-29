// ============================================================
// AttendanceAudit.gs — READ ONLY. Writes ONE output sheet.
// ------------------------------------------------------------
// Never touches Attendance_Log, Attendance_Daily, HR Maya or payroll.
// The only thing it writes is the 'Attendance_Audit' tab, which it
// clears and rebuilds each run.
//
// USAGE:  Apps Script editor → select auditAttendanceRange → Run
//         Then open the 'Attendance_Audit' tab.
//
// ── WHAT CHANGED AND WHY ─────────────────────────────────────
// The previous version judged a "silent day" purely from
// Attendance_Log. That was wrong. Head Office staff and Drivers are
// AUTO-MARKED PRESENT by writeDayToAttendanceDaily_ (Auto_HeadOffice
// and Driver_Default) and never punch at all — so the audit reported
// them as missing attendance when the system was working exactly as
// designed. Weekends and approved leave had the same problem.
//
// This version asks the only question that actually matters:
//
//     Does Attendance_Daily have a row for this employee-day,
//     and does that row resolve to something legitimate?
//
// A day is only reported as a REAL PROBLEM when Attendance_Daily
// either has no row at all, or has a row that nobody can act on
// (Absent with no explanation, No Roster, broken punches).
//
// Anything auto-resolved — Auto_HeadOffice, Driver_Default, Weekend,
// Granted_Leave, Roster_DayOff, Kiosk — is counted and shown in the
// summary, but NOT flagged as a problem.
// ============================================================

const AUDIT_FROM = '22-08-2026';
const AUDIT_TO   = '';                 // blank = yesterday
const AUDIT_TAB  = 'Attendance_Audit';

// Locations whose staff never punch — presence is automatic by design.
const AUDIT_EXEMPT_LOCATIONS = ['Head Office', 'Driver Location'];

function auditAttendanceRange() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

  const from = parseFlexibleDate_(AUDIT_FROM);
  let to;
  if (AUDIT_TO) to = parseFlexibleDate_(AUDIT_TO);
  else { to = new Date(); to.setDate(to.getDate() - 1); }
  if (!from || !to) throw new Error('Bad AUDIT_FROM / AUDIT_TO date.');
  const fromD = startOfDay_(from), toD = startOfDay_(to);

  const norm = (v) => {
    if (v == null || v === '') return '';
    if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (m) return m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[3];
    m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m) return m[3].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[1];
    const d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
  };
  const inRange = (dStr) => {
    const d = parseFlexibleDate_(dStr);
    return d && startOfDay_(d) >= fromD && startOfDay_(d) <= toD;
  };

  // ── Active staff ────────────────────────────────────────────
  const hr = ss.getSheetByName('HR Maya');
  const hd = hr.getDataRange().getValues();
  const HH = hd[0];
  const hc = (n) => HH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const hId = hc('Employee ID'), hNm = hc('Name (English)'), hSt = hc('Status');
  let hLoc = hc('Location'); if (hLoc === -1) hLoc = hc('Branch');

  const staff = {};
  for (let i = 1; i < hd.length; i++) {
    const id = String(hd[i][hId] || '').trim();
    if (!id) continue;
    if (hSt !== -1 && String(hd[i][hSt] || '').trim().toLowerCase() === 'inactive') continue;
    staff[id] = {
      name: hNm  !== -1 ? String(hd[i][hNm]  || '').trim() : id,
      loc:  hLoc !== -1 ? String(hd[i][hLoc] || '').trim() : '',
    };
  }
  const staffIds = Object.keys(staff);

  // ── Attendance_Log ──────────────────────────────────────────
  const log = ss.getSheetByName('Attendance_Log');
  const ld  = log.getDataRange().getValues();
  const LH  = ld[0];
  const lc  = (n) => LH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const lId = lc('Employee ID'), lDate = lc('Date'), lStage = lc('Stage'), lTs = lc('Timestamp');
  let lBranch = lc('Branch'); if (lBranch === -1) lBranch = lc('Location');

  const day = {};                       // 'id|date' -> { stages:{stage:[rows]}, hasDayOff }
  const byEmp = {};
  for (let i = 1; i < ld.length; i++) {
    const id = String(ld[i][lId] || '').trim();
    if (!id) continue;
    const dStr  = norm(ld[i][lDate]);
    const stage = String(ld[i][lStage] || '').trim();
    const k = id + '|' + dStr;
    if (!day[k]) day[k] = { stages: {}, hasDayOff: false };
    if (!day[k].stages[stage]) day[k].stages[stage] = [];
    day[k].stages[stage].push(i + 1);   // exact Attendance_Log row
    if (stage === 'Day Off') day[k].hasDayOff = true;

    const ts = parseLogTimestamp(ld[i][lTs]);
    if (ts) (byEmp[id] = byEmp[id] || []).push({
      ts: ts, stage: stage, row: i + 1,
      branch: lBranch !== -1 ? String(ld[i][lBranch] || '').trim() : ''
    });
  }

  // ── Attendance_Daily: the authority on whether a day resolved ──
  const ad = ss.getSheetByName('Attendance_Daily');
  const daily = {};                     // 'id|date' -> { status, source, dayType, presence, loc }
  const sourceCount = {}, statusCount = {};
  const dupDailyRows = [];
  let compiled = {};
  if (ad && ad.getLastRow() > 1) {
    const add = ad.getDataRange().getValues();
    const AH = add[0];
    const ac = (n) => AH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const aId = ac('Employee ID'), aDate = ac('Date'), aSt = ac('Status'),
          aSrc = ac('Source'), aTy = ac('Day Type'),
          aPr = ac('Presence (hrs)'), aLoc = ac('Location');
    const seen = {};
    for (let i = 1; i < add.length; i++) {
      const dStr = norm(add[i][aDate]);
      if (!inRange(dStr)) continue;
      const id = String(add[i][aId] || '').trim();
      if (!id) continue;
      compiled[dStr] = true;
      const k = id + '|' + dStr;
      seen[k] = (seen[k] || 0) + 1;
      const rec = {
        status:   aSt  !== -1 ? String(add[i][aSt]  || '').trim() : '',
        source:   aSrc !== -1 ? String(add[i][aSrc] || '').trim() : '',
        dayType:  aTy  !== -1 ? String(add[i][aTy]  || '').trim() : '',
        presence: aPr  !== -1 ? (Number(add[i][aPr]) || 0) : '',
        loc:      aLoc !== -1 ? String(add[i][aLoc] || '').trim() : '',
      };
      daily[k] = rec;
      sourceCount[rec.source || '(blank)'] = (sourceCount[rec.source || '(blank)'] || 0) + 1;
      statusCount[rec.status || '(blank)'] = (statusCount[rec.status || '(blank)'] || 0) + 1;
    }
    Object.keys(seen).forEach(k => {
      if (seen[k] > 1) {
        const p = k.split('|');
        dupDailyRows.push([p[0], (staff[p[0]]||{}).name || '?', p[1], seen[k],
                           'Compiled twice — recompiling this date clears it']);
      }
    });
  }

  // ── Dangling check-ins ──────────────────────────────────────
  const dangling = [];
  Object.keys(byEmp).forEach(id => {
    const p = byEmp[id].sort((a, b) => a.ts - b.ts);
    let open = null;
    p.forEach(x => {
      if (x.stage === 'Check In') { if (open) dangling.push({ id: id, p: open }); open = x; }
      else if (x.stage === 'Check Out') open = null;
      else if (x.stage === 'Day Off')  { if (open) dangling.push({ id: id, p: open }); open = null; }
    });
    if (open) dangling.push({ id: id, p: open });
  });
  const danglingRows = dangling
    .filter(x => { const d = startOfDay_(x.p.ts); return d >= fromD && d <= toD; })
    .map(x => [x.id, (staff[x.id]||{}).name || '?', formatDdMmYyyy_(x.p.ts),
               Utilities.formatDate(x.p.ts, tz_(), 'HH:mm'), x.p.branch, x.p.row,
               'Check In with no Check Out'])
    .sort((a, b) => (a[2] + a[0]).localeCompare(b[2] + b[0]));

  // ── Duplicates + contradictions ─────────────────────────────
  const dupRows = [], contraRows = [], overnightRows = [];
  Object.keys(day).forEach(k => {
    const parts = k.split('|'), id = parts[0], dStr = parts[1];
    if (!inRange(dStr)) return;
    const rec = day[k];
    Object.keys(rec.stages).forEach(stage => {
      const rows = rec.stages[stage];
      if (rows.length > 1) {
        dupRows.push([id, (staff[id]||{}).name || '?', dStr, stage, rows.length,
                      rows.join(', '), 'Same stage recorded ' + rows.length + ' times']);
      }
    });
    if (rec.hasDayOff) {
      const worked = ['Check In','Check Out','Break Out','Break In'].filter(s => rec.stages[s]);

      // OVERNIGHT SHIFTS ARE NOT CONTRADICTIONS.
      // A night shift starting 17:30 on the 9th and ending 03:30 on the 10th
      // writes its Check Out row dated the 10th. If the 10th is genuinely the
      // employee's day off, that date legitimately holds BOTH a Check Out
      // (belonging to the previous day's shift) and a Day Off.
      //
      // The tell is the absence of a Check In on that date: the session was
      // opened the day before. Only flag when a session actually STARTED on
      // the same date the employee declared off.
      const startedHere = !!rec.stages['Check In'];

      if (worked.length && startedHere) {
        const all = [];
        Object.keys(rec.stages).forEach(s => all.push.apply(all, rec.stages[s]));
        contraRows.push([id, (staff[id]||{}).name || '?', dStr,
                         'Day Off + ' + worked.join(', '), all.sort((a,b)=>a-b).join(', '),
                         'Shift STARTED on a declared Day Off — decide which is true']);
      } else if (worked.length) {
        // Check Out only, no Check In — overnight shift closing out.
        const all = [];
        Object.keys(rec.stages).forEach(s => all.push.apply(all, rec.stages[s]));
        overnightRows.push([id, (staff[id]||{}).name || '?', dStr,
                            worked.join(', '), all.sort((a,b)=>a-b).join(', '),
                            'Normal overnight shift closing — no action needed']);
      }
    }
  });
  dupRows.sort((a,b) => (a[2]+a[0]).localeCompare(b[2]+b[0]));
  contraRows.sort((a,b) => (a[2]+a[0]).localeCompare(b[2]+b[0]));
  overnightRows.sort((a,b) => (a[2]+a[0]).localeCompare(b[2]+b[0]));

  // ══ THE CORE FIX: classify every employee-day properly ══════
  // A day is only a PROBLEM if Attendance_Daily has no row, or has a
  // row that needs human action. Auto-resolved days are counted and
  // reported, never flagged.
  const AUTO_OK_SOURCES = ['Auto_HeadOffice','Driver_Default','Weekend',
                           'Granted_Leave','Roster_DayOff','Kiosk','Manual_Aug'];

  const notCompiled = [];   // no Attendance_Daily row at all
  const needsAction = [];   // compiled, but requires a decision
  const autoResolved = [];  // compiled and fine — informational only

  for (let d = new Date(fromD); d <= toD; d.setDate(d.getDate() + 1)) {
    const dStr = formatDdMmYyyy_(d);
    staffIds.forEach(id => {
      const k = id + '|' + dStr;
      const rec = daily[k];
      const loc = staff[id].loc;
      const exempt = AUDIT_EXEMPT_LOCATIONS.indexOf(loc) !== -1;
      const punched = !!day[k];

      if (!rec) {
        notCompiled.push([id, staff[id].name, dStr, loc,
          punched ? 'Has punches but no Attendance_Daily row'
                  : (exempt ? 'Exempt location — should have been auto-marked Present'
                            : 'No punch and no compiled row'),
          exempt ? 'CHECK COMPILE' : 'RECOMPILE THIS DATE']);
        return;
      }

      const st = rec.status, src = rec.source, ty = rec.dayType;

      if (st === 'Absent' || st === 'No Roster' || ty === 'Check Punches' || ty === 'Half Day') {
        needsAction.push([id, staff[id].name, dStr, loc, st, src, ty, rec.presence,
          st === 'Absent'        ? 'WILL be deducted at Basic/26' :
          st === 'No Roster'     ? 'Not deducted — supervisor published no roster' :
          ty === 'Check Punches' ? 'Broken presence — fix punches before payroll' :
                                   'Half day — needs an HR decision']);
      } else if (AUTO_OK_SOURCES.indexOf(src) !== -1 || st === 'Present' ||
                 st === 'Day Off' || st === 'On Leave') {
        autoResolved.push({ loc: loc, src: src, status: st });
      } else {
        needsAction.push([id, staff[id].name, dStr, loc, st, src, ty, rec.presence,
          'Unrecognised status/source — investigate']);
      }
    });
  }
  notCompiled.sort((a,b) => (a[2]+a[0]).localeCompare(b[2]+b[0]));
  needsAction.sort((a,b) => (a[2]+a[0]).localeCompare(b[2]+b[0]));

  // Auto-resolved breakdown, for the summary
  const autoBySource = {};
  autoResolved.forEach(x => {
    const key = x.src || '(blank)';
    autoBySource[key] = (autoBySource[key] || 0) + 1;
  });

  // ── WRITE THE SHEET ─────────────────────────────────────────
  let sh = ss.getSheetByName(AUDIT_TAB);
  if (!sh) sh = ss.insertSheet(AUDIT_TAB); else sh.clear();

  const out = [];
  const push = (a) => out.push(a);
  const blank = () => out.push(['']);

  const totalDays = staffIds.length * (Math.round((toD - fromD) / 86400000) + 1);

  push(['ATTENDANCE AUDIT', formatDdMmYyyy_(fromD) + ' → ' + formatDdMmYyyy_(toD),
        'generated ' + Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm')]);
  blank();
  push(['SUMMARY']);
  push(['Active staff', staffIds.length]);
  push(['Employee-days in range', totalDays]);
  push(['Dates compiled into Attendance_Daily', Object.keys(compiled).length]);
  blank();
  push(['PROBLEMS — these need action']);
  push(['1. Duplicate punches', dupRows.length, 'FIX FIRST — breaks presence and OT']);
  push(['2. Real contradictions (shift STARTED on a Day Off)', contraRows.length, 'FIX FIRST — needs a decision']);
  push(['3. Dangling check-ins', danglingRows.length, 'Backfill would close these']);
  push(['4. Not compiled at all', notCompiled.length, 'Recompile the affected dates']);
  push(['5. Needs a decision', needsAction.length, 'Absent / No Roster / Half Day / broken punches']);
  push(['6. Duplicate Attendance_Daily rows', dupDailyRows.length, 'Recompile clears these']);
  blank();
  push(['Overnight shifts closing on a Day Off', overnightRows.length, 'NORMAL — not a problem']);
  blank();
  push(['RESOLVED AUTOMATICALLY — not problems, shown for confidence']);
  push(['Total auto-resolved employee-days', autoResolved.length]);
  Object.keys(autoBySource).sort().forEach(src =>
    push(['   ' + src, autoBySource[src],
          src === 'Auto_HeadOffice' ? 'Head Office — present by default, never punches' :
          src === 'Driver_Default'  ? 'Drivers — present by default, never punch' :
          src === 'Weekend'         ? 'Head Office weekend' :
          src === 'Granted_Leave'   ? 'Approved leave' :
          src === 'Roster_DayOff'   ? 'Roster-planned day off' :
          src === 'Kiosk'           ? 'Real punches' :
          src === 'Manual_Aug'      ? 'August supervisor backfill' : '']));
  blank();
  push(['STATUS BREAKDOWN in Attendance_Daily']);
  Object.keys(statusCount).sort().forEach(s => push(['   ' + s, statusCount[s]]));

  const section = (title, headers, rows, note) => {
    push(['']); push([title + '  (' + rows.length + ')']);
    if (note) push([note]);
    if (!rows.length) { push(['None.']); return; }
    push(headers);
    rows.forEach(r => push(r));
  };

  section('1. DUPLICATE PUNCHES',
    ['Employee ID','Name','Date','Stage','Times','Attendance_Log rows','Note'], dupRows,
    'Delete the wrong row(s) in Attendance_Log. Row numbers are exact.');

  section('2. REAL CONTRADICTIONS — shift STARTED on a declared Day Off',
    ['Employee ID','Name','Date','What exists','Attendance_Log rows','Note'], contraRows,
    'A Check In exists on a date the employee declared off. Someone must decide which is true.');

  section('2b. OVERNIGHT SHIFTS closing on a Day Off — NORMAL, no action',
    ['Employee ID','Name','Date','What exists','Attendance_Log rows','Note'], overnightRows,
    'Check Out only, no Check In: the shift began the previous evening. Listed for transparency.');

  section('3. DANGLING CHECK-INS',
    ['Employee ID','Name','Date','Check In','Branch','Attendance_Log row','Note'], danglingRows,
    'Exactly what Backfill Missing Checkouts would close.');

  section('4. NOT COMPILED — no Attendance_Daily row',
    ['Employee ID','Name','Date','Location','Why','Action'], notCompiled,
    'These dates were never compiled, or the compile skipped this employee.');

  section('5. NEEDS A DECISION',
    ['Employee ID','Name','Date','Location','Status','Source','Day Type','Presence','Note'], needsAction,
    'Compiled successfully but a human has to decide what happens.');

  section('6. DUPLICATE Attendance_Daily ROWS',
    ['Employee ID','Name','Date','Times','Note'], dupDailyRows,
    'A compile fault, not bad punches.');

  const width = out.reduce((m, r) => Math.max(m, r.length), 1);
  const padded = out.map(r => { const c = r.slice(); while (c.length < width) c.push(''); return c; });
  sh.getRange(1, 1, padded.length, width).setValues(padded);

  sh.getRange(1, 1, 1, width).setFontWeight('bold').setFontSize(12);
  sh.setFrozenRows(1);
  for (let i = 0; i < padded.length; i++) {
    const f = String(padded[i][0]);
    if (/^\d\. [A-Z]/.test(f) || f === 'SUMMARY' || f.indexOf('PROBLEMS') === 0 ||
        f.indexOf('RESOLVED AUTOMATICALLY') === 0 || f.indexOf('STATUS BREAKDOWN') === 0) {
      sh.getRange(i + 1, 1, 1, width).setFontWeight('bold').setBackground('#e8eaed');
    }
  }
  sh.autoResizeColumns(1, Math.min(width, 10));

  const msg = 'Audit written. Problems: ' + dupRows.length + ' dup, ' + contraRows.length +
    ' contra, ' + danglingRows.length + ' dangling, ' + notCompiled.length + ' not compiled, ' +
    needsAction.length + ' need decision. Auto-resolved: ' + autoResolved.length + '.';
  Logger.log(msg);
  try { ss.toast(msg, 'Attendance Audit', 10); } catch (e) {}
  return msg;
}