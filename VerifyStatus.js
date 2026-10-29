/**
 * VerifyStatus.gs — what is still unanswered, and is the data sound.
 * Shakeel / Maya Tex — 20-09-2026
 *
 *   verificationStatus()        how much work is left, per location.
 *                               Run this repeatedly through the night.
 *
 *   attendanceIntegrityCheck()  proves the compile's arithmetic against
 *                               itself. Needs NO baseline.
 *
 * WHY NO BASELINE.
 * _Baseline_Snapshot was re-captured after the recompile, so it now holds
 * the post-change state and a diff against it would compare the data with
 * itself and report nothing. Rather than reconstruct it, these checks
 * re-derive the numbers from the punches and the Shifts sheet and say
 * whether the sheet agrees with them. That is a stronger test than a diff:
 * a diff only says "it changed", this says "it is right".
 */

const VS_FROM = '22-08-2026';
const VS_TO   = '20-09-2026';   // status + integrity range end (recompile runs to yesterday)

// OT at or below this on a normal-length day is a man working his rostered
// shift, not a man staying late. See the note in attendanceIntegrityCheck.
const VS_ROUTINE_OT = 1.25;


// ============================================================
// RECOMPILE THE PAYROLL RANGE — 22-08 up to YESTERDAY
// ------------------------------------------------------------
// Nothing to edit before running. Never touches today (today is not over).
// If Apps Script stops it at the 6-minute limit, run it again: it says
// which dates are left, and rebuilding a date twice is harmless.
// ============================================================
function recompileForPayroll() {
  const t = new Date(); t.setHours(0, 0, 0, 0);
  const y = new Date(t.getTime() - 86400000);
  const yStr = String(y.getDate()).padStart(2, '0') + '-' + String(y.getMonth() + 1).padStart(2, '0') + '-' + y.getFullYear();
  const res = rebuildAttendanceDaily_(VS_FROM, yStr);
  const out = ['RECOMPILE ' + VS_FROM + ' → ' + yStr, ''];
  if (!res || res.success === false) { out.push('FAILED: ' + ((res && res.error) || 'no result')); return vsReport_(out); }
  out.push('Dates rebuilt: ' + (res.dates || []).length);
  const left = res.skipped || [];
  if (left.length) {
    out.push('NOT FINISHED — ' + left.length + ' date(s) left: ' + left.join(', '));
    out.push('Run recompileForPayroll() again.');
  } else {
    out.push('Complete. Now run attendanceIntegrityCheck().');
  }
  return vsReport_(out);
}

// ============================================================
// HOW MUCH WORK IS LEFT — run this through the night
// ============================================================
function verificationStatus() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const out = [];
  const locs = vsLocations_(ss);
  if (!locs.length) return vsReport_(['No locations found in HR Maya.']);

  out.push('VERIFICATION OUTSTANDING   ' + VS_FROM + ' → ' + VS_TO);
  out.push('as at ' + Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm'));
  out.push('');
  out.push(vsPad_('LOCATION', 22) + vsPad_('OT LEFT', 10) + vsPad_('DAYS LEFT', 12) +
           vsPad_('DONE', 8) + 'OVER-CLAIMS');
  out.push(vsPad_('', 62, '-'));

  let tOt = 0, tAtt = 0, tDone = 0, tOver = 0;
  const idle = [];

  locs.forEach(loc => {
    let r;
    try { r = getSupervisorVerifyList({ location: loc, from: VS_FROM, to: VS_TO }); }
    catch (e) { out.push(vsPad_(loc, 22) + 'ERROR: ' + e.message); return; }
    if (!r || !r.success) { out.push(vsPad_(loc, 22) + 'ERROR: ' + ((r && r.error) || '?')); return; }

    const otLeft  = r.otCount  || 0;
    const attLeft = r.attCount || 0;
    const done    = (r.otAnswered || 0) + (r.attAnswered || 0);
    const over    = r.overCount || 0;

    tOt += otLeft; tAtt += attLeft; tDone += done; tOver += over;
    if (done === 0 && (otLeft + attLeft) > 0) idle.push(loc);

    out.push(vsPad_(loc, 22) + vsPad_(String(otLeft), 10) + vsPad_(String(attLeft), 12) +
             vsPad_(String(done), 8) + (over ? over + '  (' + (r.overHoursHm || '') + ')' : ''));
  });

  out.push(vsPad_('', 62, '-'));
  out.push(vsPad_('TOTAL', 22) + vsPad_(String(tOt), 10) + vsPad_(String(tAtt), 12) +
           vsPad_(String(tDone), 8) + String(tOver));
  out.push('');
  out.push((tOt + tAtt) + ' answer(s) still needed across ' + locs.length + ' location(s).');

  if (idle.length) {
    out.push('');
    out.push('NOT STARTED — nobody at these locations has answered anything yet:');
    idle.forEach(l => out.push('   ' + l));
    out.push('Chase these first. A location on zero has not opened the page.');
  }
  return vsReport_(out);
}


// ============================================================
// IS THE DATA SOUND — no baseline required
// ============================================================
function attendanceIntegrityCheck() {
  const ss  = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const ad  = ss.getSheetByName('Attendance_Daily');
  if (!ad) return vsReport_(['Attendance_Daily not found.']);

  const all = ad.getRange(1, 1, ad.getLastRow(), ad.getLastColumn()).getValues();
  const h   = all[0].map(x => String(x || '').trim());
  const at  = n => h.indexOf(n);
  const C = { id:at('Employee ID'), name:at('Name'), date:at('Date'), shift:at('Shift'),
              ci:at('Check In'), bo:at('Break Out'), bi:at('Break In'), co:at('Check Out'),
              pres:at('Presence (hrs)'), ot:at('OT Hours'),
              dayType:at('Day Type'), src:at('Source'), st:at('Status'), flags:at('Flags') };
  if (C.id < 0 || C.date < 0) return vsReport_(['Attendance_Daily has no recognisable header row.']);

  const norm = vsShiftHours_(ss);
  const out = [], bad = [];
  let checked=0, breakRows=0, statedPres=0, statedOther=0, routine=0, realOt=0, totalOt=0, open=0;
  const routineSample = [], breakSample = [];

  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const id = String(r[C.id] || '').trim();
    const d  = adNormDd_(r[C.date]);
    if (!id || !d || !adDateInRange_(d, VS_FROM, VS_TO)) continue;

    const nm   = C.name >= 0 ? String(r[C.name] || '') : '';
    // Times come back from the sheet as Date values (1899-12-30 HH:mm), not
    // text. Read them as clock times, or the arithmetic check silently skips
    // every punched day — which is exactly what the first version did.
    const ci   = vsClock_(r[C.ci]);
    const co   = vsClock_(r[C.co]);
    const bo   = vsClock_(r[C.bo]);
    const pres = vsNum_(r[C.pres]);
    const ot   = vsNum_(r[C.ot]);
    const src  = String(r[C.src] || '').trim();
    const st   = String(r[C.st]  || '').trim();
    const nh   = norm[String(r[C.shift] || '').trim()] || 9;
    const tag  = vsPad_(id, 10) + vsPad_(nm, 20) + vsPad_(d, 12);

    totalOt += ot;

    // ── 1. A punched day: presence must be raw clock time, break NOT deducted.
    // A split shift is the SUM of its sessions, so its span is not its
    // presence — those rows are checked by the compile, not here.
    const split = C.flags >= 0 && /Split shift/.test(String(r[C.flags] || ''));
    if (ci && co && st === 'Present' && !split) {
      const span = vsSpan_(ci, co);
      if (span != null) {
        checked++;
        if (Math.abs(span - pres) > 0.035 && pres !== 0) {   // HH:mm drops seconds: up to 2 min either way
          bad.push(tag + 'presence ' + pres + ' but ' + ci + '→' + co + ' is ' +
                   vsHm_(span) + (bo ? '  (break ' + bo + ' — is it being deducted?)' : ''));
        }
        const wantOt = Math.max(0, span - nh);
        if (Math.abs(wantOt - ot) > 0.035 && pres !== 0) {
          bad.push(tag + 'OT ' + ot + ' but ' + vsHm_(span) + ' on site against a ' +
                   nh + 'h shift is ' + vsHm_(wantOt));
        }
      }
      if (bo) { breakRows++; if (breakSample.length < 4) breakSample.push(tag + ci + '→' + co + '  ' + vsHm_(pres) + '  OT ' + vsHm_(ot)); }
      if (ot > 0 && ot <= VS_ROUTINE_OT) {
        routine++; if (routineSample.length < 6) routineSample.push(tag + ci + '→' + co + '  OT ' + vsHm_(ot) + '  (' + nh + 'h shift)');
      } else if (ot > VS_ROUTINE_OT) realOt++;
    }

    // ── 2. A stated day: supervisor said what happened, no punch stages.
    if (src === 'Manual_Aug') {
      if (st === 'Present') {
        statedPres++;
        if (Math.abs(pres - nh) > 0.02)
          bad.push(tag + 'stated Present but presence is ' + vsCell_(r[C.pres]) + ', expected ' + nh);
        if (ot !== 0)
          bad.push(tag + 'stated Present carries OT ' + ot + ' — a day with no punches cannot evidence overtime');
      } else {
        statedOther++;
        if (pres !== 0 && String(r[C.pres] || '').trim() !== '')
          bad.push(tag + 'stated ' + st + ' but carries presence ' + vsCell_(r[C.pres]));
      }
    }

    // ── 3. Still open: punched in, never out.
    if (ci && !co && st === 'Present') open++;
  }

  out.push('ATTENDANCE INTEGRITY   ' + VS_FROM + ' → ' + VS_TO);
  out.push('');
  out.push('Punched days with a break punch   : ' + breakRows);
  breakSample.forEach(s => out.push('      ' + s));
  out.push('Supervisor-stated Present         : ' + statedPres);
  out.push('Supervisor-stated Absent/Day Off  : ' + statedOther);
  out.push('Days punched in, never punched out: ' + open);
  out.push('');
  out.push('TOTAL OT IN RANGE                 : ' + vsHm_(totalOt));
  out.push('  of which routine (≤' + VS_ROUTINE_OT + 'h)        : ' + routine + ' row(s)');
  routineSample.forEach(s => out.push('      ' + s));
  out.push('  genuinely long days             : ' + realOt + ' row(s)');
  out.push('');

  if (!bad.length) {
    out.push('0 ARITHMETIC FAULTS. ' + checked + ' punched day(s) checked — every one agrees with its own punches and shift.');
  } else {
    out.push(bad.length + ' ROW(S) WHERE THE SHEET DISAGREES WITH ITS OWN PUNCHES:');
    bad.slice(0, 40).forEach(s => out.push('   ' + s));
    if (bad.length > 40) out.push('   ...and ' + (bad.length - 40) + ' more');
  }
  return vsReport_(out);
}


// ============================================================
// helpers
// ============================================================
// Any date cell → dd-mm-yyyy, including a Date that was written to the
// sheet as text ("Sat Sep 05 2026 00:00:00 GMT+0300 (Arabian Standard Time)").
function vsAnyDate_(v) {
  if (v == null || v === '') return '';
  const a = adNormDd_(v);
  if (/^\d{2}-\d{2}-\d{4}$/.test(a)) return a;
  const m = String(v).match(/^[A-Z][a-z]{2} ([A-Z][a-z]{2}) (\d{1,2}) (\d{4})/);
  if (!m) return '';
  const mon = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'].indexOf(m[1]) + 1;
  return mon ? String(m[2]).padStart(2, '0') + '-' + String(mon).padStart(2, '0') + '-' + m[3] : '';
}

function vsClock_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]')
    return isNaN(v.getTime()) ? '' : String(v.getHours()).padStart(2, '0') + ':' + String(v.getMinutes()).padStart(2, '0');
  const m = String(v).trim().match(/(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2, '0') + ':' + m[2] : '';
}

function vsLocations_(ss) {
  const sh = ss.getSheetByName('HR Maya');
  if (!sh || sh.getLastRow() < 2) return [];
  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const h = all[0].map(x => String(x || '').trim());
  const cL = h.indexOf('Location'), cS = h.indexOf('Status');
  if (cL < 0) return [];
  const seen = {}, out = [];
  for (let i = 1; i < all.length; i++) {
    if (cS >= 0 && String(all[i][cS] || '').trim().toLowerCase() === 'inactive') continue;
    const l = String(all[i][cL] || '').trim();
    if (!l || seen[l]) continue;
    seen[l] = true; out.push(l);
  }
  return out.sort();
}

function vsShiftHours_(ss) {
  const out = {};
  const sh = ss.getSheetByName('Shifts');
  if (!sh || sh.getLastRow() < 2) return out;
  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const h = all[0].map(x => String(x || '').trim());
  const cN = h.indexOf('Shift') >= 0 ? h.indexOf('Shift') : h.indexOf('Shift Name');
  const cH = h.indexOf('Normal Hours') >= 0 ? h.indexOf('Normal Hours') : h.indexOf('Normal Hrs');
  if (cN < 0 || cH < 0) return out;
  for (let i = 1; i < all.length; i++) {
    const n = String(all[i][cN] || '').trim();
    const v = parseFloat(all[i][cH]);
    if (n && !isNaN(v) && v > 0) out[n] = v;
  }
  return out;
}

// Hours between two HH:mm strings, crossing midnight when the second is earlier.
function vsSpan_(a, b) {
  const p = s => { const m = String(s).match(/^(\d{1,2}):(\d{2})/); return m ? (+m[1]) * 60 + (+m[2]) : null; };
  const x = p(a), y = p(b);
  if (x == null || y == null) return null;
  let d = y - x; if (d <= 0) d += 1440;
  return d / 60;
}

function vsHm_(h) { const s = h < 0 ? '-' : ''; h = Math.abs(h || 0);
                    return s + Math.floor(h) + 'h ' + Math.round((h % 1) * 60) + 'm'; }
function vsNum_(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function vsCell_(v) { return (v === '' || v == null) ? '(blank)' : String(v); }
function vsPad_(s, n, ch) { s = String(s); ch = ch || ' ';
  return s.length >= n ? s.slice(0, n - 1) + ' ' : s + ch.repeat(n - s.length); }

function vsReport_(lines) {
  const txt = lines.join('\n');
  Logger.log(txt);
  return txt;
}


// ============================================================
// IS THIS DAY FIXED?  — checkDay()
// ------------------------------------------------------------
// Set CHECK_DAY, run checkDay(). Writes nothing. Answers, for one date:
//   1. Is Attendance_Daily up to date with the rules? It rebuilds the day
//      in memory and compares with the sheet, row by row.
//   2. Does every active man have exactly one row?
//   3. What is still unanswered, per location (open check-ins, Absent
//      with punches, Half Days)?
//   4. Are there verdicts left that were given before the shift ended?
//   5. OT claims already covered by an OT_Entries row (his entry stands).
//   6. OT: every claim that belongs to this day's shifts — filed on the
//      wrong date, two claims for one shift, claims with no shift.
// The last line says either "CLEAN" or exactly what is left.
// checkYesterday() does the same for yesterday.
// ============================================================
const CHECK_DAY = '20-09-2026';

function checkDay()       { return vsCheckDay_(CHECK_DAY); }
function checkYesterday() {
  const t = new Date(); t.setHours(0, 0, 0, 0);
  const y = new Date(t.getTime() - 86400000);
  return vsCheckDay_(String(y.getDate()).padStart(2, '0') + '-' + String(y.getMonth() + 1).padStart(2, '0') + '-' + y.getFullYear());
}

function vsCheckDay_(D) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const out = [], todo = [];
  const pad = n => String(n).padStart(2, '0');
  const cell = v => {
    if (v == null || v === '') return '';
    if (Object.prototype.toString.call(v) === '[object Date]') return pad(v.getHours()) + ':' + pad(v.getMinutes());
    if (typeof v === 'number') return (Math.round(v * 100) / 100).toString();
    const n = Number(v); if (String(v).trim() !== '' && !isNaN(n)) return (Math.round(n * 100) / 100).toString();
    return String(v).trim();
  };
  out.push('CHECK ' + D + '   (as at ' + Utilities.formatDate(new Date(), tz_(), 'dd-MM HH:mm') + ')');
  out.push('');

  // ── 1 + 2. Sheet vs a fresh build of the same day ──────────────
  const ctx = adLoadContext_(ss);
  const fresh = adBuildDayRows_(ctx, D);
  const ad = ss.getSheetByName('Attendance_Daily');
  const all = ad.getRange(1, 1, ad.getLastRow(), ad.getLastColumn()).getValues();
  const onSheet = {}, dupes = [];
  for (let i = 1; i < all.length; i++) {
    if (adNormDd_(all[i][2]) !== D) continue;
    const id = String(all[i][0] || '').trim();
    if (onSheet[id]) dupes.push(id); else onSheet[id] = all[i];
  }
  const stale = [], missing = [];
  const COLS = [5, 6, 7, 8, 9, 13, 14, 15, 16];          // in/out, presence, OT, type, source, status
  fresh.forEach(f => {
    const s = onSheet[f[0]];
    if (!s) { missing.push(f[0] + ' ' + f[1]); return; }
    const diff = COLS.filter(c => cell(s[c]) !== cell(f[c]));
    if (diff.length) stale.push(sv_pad2_(f[0], 10) + sv_pad2_(f[1], 20) + 'sheet: ' +
      [cell(s[5]), cell(s[8]), cell(s[9]) + 'h', s[16]].join(' ') + '   rules now give: ' +
      [cell(f[5]), cell(f[8]), cell(f[9]) + 'h', f[16]].join(' '));
  });
  out.push('1. ATTENDANCE_DAILY UP TO DATE WITH THE RULES');
  if (!stale.length && !missing.length) out.push('   yes — every row matches a fresh build of ' + D);
  else {
    out.push('   NO — ' + stale.length + ' row(s) differ' + (missing.length ? ', ' + missing.length + ' missing' : '') + ':');
    stale.slice(0, 30).forEach(x => out.push('     ' + x));
    missing.forEach(x => out.push('     missing: ' + x));
    todo.push('Recompile — run recompileForPayroll() (the sheet is older than the rules).');
  }
  if (dupes.length) { out.push('   DUPLICATE rows for: ' + dupes.join(', ')); todo.push('Duplicate Attendance_Daily rows — recompile.'); }

  // Status breakdown, from what the rules give.
  const byStatus = {};
  fresh.forEach(f => { const k = String(f[16]); byStatus[k] = (byStatus[k] || 0) + 1; });
  out.push('   ' + fresh.length + ' staff: ' + Object.keys(byStatus).map(k => byStatus[k] + ' ' + k).join(', '));
  const rep = fresh.filter(f => /Repeated check-in/.test(String(f[17])));
  if (rep.length) out.push('   ' + rep.length + ' day(s) with repeated check-in taps, counted from the first tap that lines up: ' +
    rep.map(f => f[1] + ' ' + cell(f[5])).join(', '));
  out.push('');

  // ── 3. Still to answer, per location ─────────────────────────
  out.push('3. STILL TO ANSWER (supervisor page)');
  let open = 0;
  vsLocations_(ss).forEach(loc => {
    const r = getSupervisorVerifyList({ location: loc, from: D, to: D });
    if (!r || !r.success) return;
    const att = r.attendance || [], ot = r.ot || [];
    if (!att.length && !ot.length && !r.stillRunning) return;
    out.push('   ' + loc + ': ' + att.length + ' day(s), ' + ot.length + ' OT claim(s)' +
             (r.stillRunning ? ', ' + r.stillRunning + ' not finished yet' : ''));
    att.forEach(a => out.push('       ' + sv_pad2_(a.name, 20) + a.status + ' — ' + String(a.context).slice(0, 110)));
    open += att.length + ot.length;
  });
  if (!open) out.push('   nothing — every day and claim here has been answered');
  else todo.push(open + ' answer(s) still needed on the supervisor page.');
  out.push('');

  // ── 4. Early verdicts ────────────────────────────────────────
  const u = sv_undoEarly_(false, true);
  const earlyHere = Object.keys(u.early).filter(k => k.split('|')[1] === D);
  const delHere = u.logDel.filter(x => x.d === D).length + u.manDel.filter(x => x.d === D).length;
  out.push('4. VERDICTS GIVEN BEFORE THE SHIFT ENDED');
  if (!delHere) out.push('   none left');
  else { out.push('   ' + delHere + ' record(s) on ' + earlyHere.length + ' man-day(s) still in place'); todo.push('Run svPreviewUndoEarly() then svCommitUndoEarly().'); }
  out.push('');

  // ── 6. OT claims that belong to this day's shifts ────────────
  out.push('6. OT CLAIMS FOR SHIFTS THAT STARTED ON ' + D);
  const dd = s => { const p = s.split('-'); return new Date(+p[2], +p[1] - 1, +p[0]); };
  const around = d => { const t = dd(D).getTime() + d * 86400000, x = new Date(t);
                        return pad(x.getDate()) + '-' + pad(x.getMonth() + 1) + '-' + x.getFullYear(); };
  const L = getOtCompileList({ from: around(-1), to: around(1) });
  const mine = (L && L.rows ? L.rows : []).filter(r => (r.belongsTo || r.date) === D);
  const perShift = {};
  mine.forEach(r => { (perShift[r.empId] = perShift[r.empId] || []).push(r); });
  let misdated = 0, doubles = 0, unresolved = 0, posted = 0;
  Object.keys(perShift).forEach(id => {
    const list = perShift[id];
    if (list.length > 1) doubles++;
    list.forEach(r => {
      const moved = r.date !== D;
      if (moved) misdated++;
      if (r.flag === 'incomplete-punch' || r.flag === 'worked-day-off' || r.flag === 'no-attendance') unresolved++;
      if (r.flag === 'already-posted') posted++;
      out.push('   ' + sv_pad2_(r.name, 20) + sv_pad2_(r.claimedHm, 7) + 'punches ' + sv_pad2_(r.computedHm, 8) +
        (r.supVerdict ? 'supervisor ' + r.supVerdict + ' ' + r.supHoursHm + '  ' : '') +
        (moved ? '[dated ' + r.date + ' — belongs here] ' : '') +
        (list.length > 1 ? '[' + list.length + ' CLAIMS FOR ONE SHIFT] ' : '') + r.flag +
        (r.flag === 'already-posted' ? ' ' + r.postedHm + ' in OT_Entries' : ''));
    });
  });
  if (!mine.length) out.push('   no claims');
  if (misdated) todo.push(misdated + ' claim(s) filed on the wrong date — run commitOtDateFix().');
  if (doubles) todo.push(doubles + ' man(s) with two claims for one shift — reject the duplicate on the OT Review page.');
  if (unresolved) todo.push(unresolved + ' claim(s) on a day with no finished shift — answer the attendance first.');
  if (posted) todo.push(posted + ' claim(s) already covered by OT_Entries — run commitOtEntriesReconcile() to close them.');
  out.push('');

  out.push(todo.length ? D + ' IS NOT DONE — ' + todo.length + ' thing(s) left:' : D + ' IS CLEAN.');
  todo.forEach((t, i) => out.push('   ' + (i + 1) + '. ' + t));
  return vsReport_(out);
}


// ============================================================
// THE ONE OT CHECK — otCheck()
// ------------------------------------------------------------
// For every man and every shift in the range, puts the four OT numbers
// side by side, each from the ONE place it lives:
//   punches   Attendance_Daily "OT Hours" — presence − normal hours
//   claimed   OT_Pending, attached to the shift it belongs to
//   supervisor OT_Pending "Hours Confirmed" (same row as the claim)
//   paid      OT_Entries — the only OT payroll reads
// and lists only where they disagree in a way that costs money:
//   • paid more than the punches show, with punch evidence
//   • two claims for one shift
// ============================================================
function otCheck() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const out = [];
  const num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
  const hm = h => vsHm_(h);

  // punches
  const ad = ss.getSheetByName('Attendance_Daily');
  const A = ad.getRange(1, 1, ad.getLastRow(), ad.getLastColumn()).getValues();
  const punch = {}, evidence = {};
  for (let i = 1; i < A.length; i++) {
    const d = adNormDd_(A[i][2]); if (!adDateInRange_(d, VS_FROM, VS_TO)) continue;
    const k = String(A[i][0]).trim() + '|' + d;
    punch[k] = num(A[i][13]);
    evidence[k] = String(A[i][15]) === 'Kiosk' && String(A[i][16]) === 'Present';
  }
  // claims, by the shift they belong to
  const L = getOtCompileList({ from: VS_FROM, to: VS_TO });
  const claims = {};
  (L.rows || []).forEach(r => { const k = r.empId + '|' + (r.belongsTo || r.date); (claims[k] = claims[k] || []).push(r); });
  // paid
  const oe = ss.getSheetByName('OT_Entries');
  const paid = {}, names = {}, routes = {}, routeHours = {};
  let textDates = 0;
  if (oe && oe.getLastRow() > 1) {
    oe.getRange(2, 1, oe.getLastRow() - 1, Math.max(6, oe.getLastColumn())).getValues().forEach(r => {
      const d = vsAnyDate_(r[0]); if (!d || !adDateInRange_(d, VS_FROM, VS_TO)) return;
      const k = String(r[1]).trim() + '|' + d;
      const rem = String(r[4] || '') + ' ' + String(r[5] || '');
      const route = /compiler sheet/i.test(rem) ? 'sheet'
                  : /^auto from attendance/i.test(String(r[4] || '')) ? 'auto' : 'claim';
      paid[k] = (paid[k] || 0) + num(r[3]); names[String(r[1]).trim()] = String(r[2] || '');
      (routes[k] = routes[k] || {})[route] = true;
      routeHours[route] = (routeHours[route] || 0) + num(r[3]);
      if (typeof r[0] === 'string' && /^[A-Z][a-z]{2} [A-Z][a-z]{2} \d/.test(r[0])) textDates++;
    });
  }
  let tPunch = 0, tClaim = 0, tPaid = 0;
  Object.keys(punch).forEach(k => { tPunch += punch[k]; });
  Object.keys(claims).forEach(k => claims[k].forEach(r => { tClaim += num(r.claimed); }));
  Object.keys(paid).forEach(k => { tPaid += paid[k]; });

  const overpaid = [], doubles = [];
  Object.keys(paid).forEach(k => {
    if (evidence[k] && paid[k] > (punch[k] || 0) + 0.1)
      overpaid.push(sv_pad2_(k.replace('|', '  '), 24) + sv_pad2_(names[k.split('|')[0]] || '', 20) +
                    'paid ' + hm(paid[k]) + ', punches show ' + hm(punch[k] || 0));
  });
  Object.keys(claims).forEach(k => {
    if (claims[k].length > 1 && !claims[k].some(r => r.flag === 'already-posted')) doubles.push(sv_pad2_(k.replace('|', '  '), 24) + sv_pad2_(claims[k][0].name, 20) +
      claims[k].map(r => r.claimedHm + ' (dated ' + r.date + ')').join(' + '));
  });

  out.push('OT CHECK   ' + VS_FROM + ' → ' + VS_TO);
  out.push('');
  out.push('Punches show   : ' + hm(tPunch) + '   (Attendance_Daily — the fact)');
  out.push('Pending claims : ' + hm(tClaim) + '   (OT_Pending — the request)');
  out.push('Posted to pay  : ' + hm(tPaid) + '   (OT_Entries — the decision; payroll reads only this)');
  out.push('');
  out.push('  of which the OT compiler\'s sheet : ' + hm(routeHours.sheet || 0) + '   ← his entry wins');
  out.push('  of which from approved claims    : ' + hm(routeHours.claim || 0));
  out.push('  of which "Auto from attendance"  : ' + hm(routeHours.auto || 0) +
           '   ← posted by the OT compiler; his entry stands');
  const stillOpen = (L.rows || []).filter(r => r.flag === 'already-posted');
  if (stillOpen.length) out.push('  ' + stillOpen.length + ' Pending claim(s) are for days already in OT_Entries — ' +
           'run previewOtEntriesReconcile(), then commitOtEntriesReconcile(), so they are never paid twice');
  const both = Object.keys(routes).filter(k => Object.keys(routes[k]).length > 1);
  if (both.length) out.push('  TWO ROWS FOR ONE DAY — PAID TWICE (' + both.length + '): ' + both.map(k => k.replace('|', ' ')).join(', ') +
           '  — delete the wrong row in OT_Entries');
  if (textDates) out.push('  ' + textDates + ' OT_Entries row(s) have a date stored as text ("Sat Sep 05 2026 …") — read here, but other tools may miss them');
  out.push('');
  out.push('POSTED FOR PAYMENT ABOVE WHAT THE PUNCHES SHOW (' + overpaid.length + '):');
  overpaid.slice(0, 60).forEach(x => out.push('   ' + x));
  out.push('');
  out.push('TWO OR MORE CLAIMS FOR ONE SHIFT (' + doubles.length + ') — pay one, reject the rest:');
  doubles.slice(0, 60).forEach(x => out.push('   ' + x));
  return vsReport_(out);
}