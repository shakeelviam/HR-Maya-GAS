// ============================================================
// OtCompile.gs — OT verification console for the OT compiler
// ------------------------------------------------------------
// WHY THIS EXISTS
// OT claims are submitted by staff on the kiosk and land in OT_Pending.
// They are the employee's own statement of what they worked, so they are
// the starting point — not something to be silently overridden by a
// computed figure. But punch data has known gaps, so a human has to
// verify each claim before any of it reaches payroll.
//
// This is that human's console. It is NOT the supervisor: OT compiling is
// a separate role with its own allowlist and no access to the HR dashboard.
//
// WHAT THE COMPILER SEES PER CLAIM
//   the claim itself (the thing being verified)
//   punch context — check in/out, presence, computed OT (REFERENCE ONLY)
//   a flag describing how claim and punches relate
//   THE SUPERVISOR'S VERDICT, read off the same OT_Pending row
//   a decision: Approve / Adjust / Reject, pre-filled with the supervisor's
//   figure when he has given one, otherwise with the claim
//
// ONE RECORD, ONE PLACE
// The supervisor's answer is not in a tab of its own. SupervisorVerify.gs
// writes it onto the OT_Pending row — Supervisor Verdict / Hours Confirmed /
// Verified By / Verified On — so the claim, the punches and the statement of
// the man who was there are one row. Nothing here has to be cross-checked
// against a second sheet, and a paid hour traces back to whoever vouched for
// it through the OT_Entries remark.
//
// Only approved rows reach OT_Entries. Nothing else touches payroll.
//
// ACCESS
// Lives inside the HR dashboard (OT Review page) and inherits its Google
// sign-in — no separate login, no separate page. The 'by' email is passed
// through only so each decision is attributable in the audit trail.
//
// ROUTER CASES (add to doGet):
//   getOtCompileList, postApprovedOt
// ============================================================

const OTC_PENDING_TAB = 'OT_Pending';
const OTC_ENTRIES_TAB = 'OT_Entries';

// ── THE OT COMPILER'S ENTRY WINS  (21-09-2026, Shakeel) ─────────────
// The OT compiler collects and compiles all OT. Anything already in
// OT_Entries for a man and a day is HIS number and it stands:
//   • a claim for that same day is not paid a second time — it is closed
//     as OTC_SETTLED, with his figure and the claim noted in the audit log;
//   • OT Review shows such a claim as "Already in OT_Entries" so it is
//     never approved on top of his entry;
//   • postApprovedOt() and the "Post ticked to OT_Entries" button never
//     add a second row for a man and day already there.
// previewOtEntriesReconcile() / commitOtEntriesReconcile() apply this to
// everything filed so far.
const OTC_COMPILER_EMAILS = ['3abwd.3@gmail.com'];
const OTC_SETTLED = 'Settled (OT_Entries)';

// OT_Pending layout: 0 ID | 1 Employee ID | 2 Name | 3 Date | 4 Hours
//                    5 Source | 6 Status | 7 Submitted | 8 Geo Flag | 9 Computed OT
const OTC_COL = { id:0, empId:1, name:2, date:3, hours:4,
                  source:5, status:6, submitted:7, geo:8, computed:9 };

function otc_norm_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[1];
  const d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}

// Everything already in OT_Entries, by man and day.
// { 'MT-00007|05-09-2026': { hours, rows:[{row, hours, remark, by, compiler}] } }
// Dates are read whatever way they were stored — a date cell, "05-09-2026",
// or the text "Sat Sep 05 2026 00:00:00 GMT+0300 (…)".
function otc_entriesIndex_(ss) {
  const idx = {};
  const ot = ss.getSheetByName(OTC_ENTRIES_TAB);
  if (!ot || ot.getLastRow() < 2) return idx;
  const c = otc_entryCols_(ot);
  ot.getRange(2, 1, ot.getLastRow() - 1, ot.getLastColumn()).getValues().forEach((r, i) => {
    const id = String(r[c.empId] || '').trim().toUpperCase();
    const d = otc_anyDate_(r[c.date]);
    if (!id || !d) return;
    const by = c.by >= 0 ? String(r[c.by] || '') : '';
    const e = { row: i + 2, hours: Number(r[c.hours]) || 0,
                remark: c.remark >= 0 ? String(r[c.remark] || '') : '',
                by: by, compiler: otc_isCompiler_(by) };
    const k = id + '|' + d;
    (idx[k] = idx[k] || { hours: 0, rows: [] }).rows.push(e);
    idx[k].hours += e.hours;
  });
  return idx;
}
function otc_isCompiler_(by) {
  const s = String(by || '').toLowerCase();
  return OTC_COMPILER_EMAILS.some(e => s.indexOf(e.toLowerCase()) !== -1);
}
// OT_Entries columns by header. Live layout:
//   Date | Employee ID | Name | OT Hours | Remark | Added By
function otc_entryCols_(ot) {
  const H = ot.getRange(1, 1, 1, ot.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  const at = (...names) => { for (const n of names) { const i = H.indexOf(n); if (i !== -1) return i; } return -1; };
  const c = { date: at('date'), empId: at('employee id'), name: at('name'),
              hours: at('ot hours', 'hours'), remark: at('remark'), by: at('added by'),
              source: at('source'), width: H.length };
  if (c.date < 0) c.date = 0; if (c.empId < 0) c.empId = 1;
  if (c.name < 0) c.name = 2; if (c.hours < 0) c.hours = 3;
  return c;
}
function otc_anyDate_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : formatDdMmYyyy_(v);
  const s = String(v).trim();
  if (/^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}$/.test(s) || /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/.test(s)) return otc_norm_(s);
  const d = new Date(s.replace(/\s*\(.*\)\s*$/, ''));
  return isNaN(d.getTime()) ? '' : formatDdMmYyyy_(d);
}
function otc_entryText_(e) {
  return e.rows.map(x => otc_hm_(x.hours) + ' "' + (x.remark || '—') + '"' +
                         (x.by ? ' ' + x.by : '')).join(' + ');
}

// A time-only cell in Attendance_Daily comes back as a Date anchored to
// 1899-12-30. Stringifying it gives "Sat Dec 30 1899 18:26:00 GMT+0306".
// Pull out just the clock time.
function otc_clock_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, tz_(), 'HH:mm');
  }
  const s = String(v).trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2,'0') + ':' + m[2] : s;
}

// Hours -> "7h 58m". Carries properly: 7.996h is 8h, NOT "7h 60m".
// (Flooring the hours and rounding the minutes separately let 59.76 minutes
//  round to 60 and print beside the floored hour.)
function otc_hm_(h) {
  if (h == null || h === '' || isNaN(Number(h))) return '—';
  h = Number(h);
  const neg = h < 0;
  const mins = Math.round(Math.abs(h) * 60);
  const hh = Math.floor(mins / 60), mm = mins % 60;
  if (!hh && !mm) return '0h';
  return (neg ? '-' : '') + (hh ? hh + 'h' : '') + (mm ? (hh ? ' ' : '') + mm + 'm' : '');
}

// The supervisor's verdict lives on the OT_Pending row itself — written by
// SupervisorVerify.gs, read here. One row tells the whole story: what was
// claimed, what the punches show, and what the man's supervisor says.
// Locates the four columns by header; never creates them (that is the
// verifier's job) and never assumes a position.
function otc_verdictCols_(pen) {
  const H = pen.getRange(1, 1, 1, pen.getLastColumn()).getValues()[0];
  const at = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  return { verdict: at('Supervisor Verdict'), hours: at('Hours Confirmed'),
           by: at('Verified By'), on: at('Verified On') };
}

// Sources that record attendance WITHOUT punch times. A person marked
// Present this way never punched and was never expected to — so there is
// no check-out to be "missing". Calling that an incomplete punch is wrong:
//   Manual_Aug      supervisors stated Present/Absent for 22-31 Aug, no times
//   Auto_HeadOffice Head Office staff are present by default
//   Driver_Default  drivers are present by default
// For OT purposes these days have NO punch evidence either way, so the
// claim can only be settled by whoever knows what was worked.
const OTC_TIMELESS_SOURCES = ['Manual_Aug', 'Auto_HeadOffice', 'Driver_Default'];

// ════════════════════════════════════════════════════════════
// WHICH SESSION DOES AN OT CLAIM BELONG TO?
// ------------------------------------------------------------
// A second shift starts 17:30 and ends 03:30. The man finishes at half
// three in the morning, taps OT on the kiosk, and the date picker shows
// TODAY — the 14th. But the shift he is claiming for started on the 13th.
// The 14th has no session of its own, so every check against it came back
// "no punches recorded", and his supervisor was asked to explain a day
// that was really the back half of the previous night.
//
// This resolves a claim to the session that actually produced it:
//   a session that STARTED on the claim date          → its own day
//   else a session that ENDED on it, having started
//        the day before                               → OVERNIGHT SPILLOVER
//   else nothing                                      → genuinely no session
//
// Sessionizing goes through AttendanceCore (acReadPunches_ + acSessionize_)
// rather than a private copy. That file is the one definition of "a Check
// In opens a session"; this is its first caller.
// ════════════════════════════════════════════════════════════
function otcSessionIndex_(ss) {
  if (typeof acReadPunches_ !== 'function' || typeof acSessionize_ !== 'function') {
    Logger.log('otcSessionIndex_: AttendanceCore.gs not found — falling back to date-only matching.');
    return null;
  }
  const byEmp = acReadPunches_(ss);           // whole log, once, sorted
  const startedOn = {}, endedOn = {}, openOn = {}, doneByEmp = {};
  // Repeated check-in taps: the session starts at the first tap that lines
  // up with a shift (AttendanceCore). Same rule the 4 AM compile uses.
  const startsFor = (typeof acShiftStartsIndex_ === 'function') ? acShiftStartsIndex_(ss) : null;

  Object.keys(byEmp).forEach(empId => {
    const locs = {}; byEmp[empId].forEach(p => { if (p.location) locs[p.location] = true; });
    const s = acSessionize_(byEmp[empId], startsFor ? startsFor(empId, Object.keys(locs)) : null);

    s.completed.forEach(sess => {
      const sd = acNormDate_(sess.checkIn);
      const ed = acNormDate_(sess.checkOut);
      const p  = acPresence_(sess);
      const rec = {
        empId: empId, startDate: sd, endDate: ed,
        inHm: acHHmm_(sess.checkIn), outHm: acHHmm_(sess.checkOut),
        presence: p.presence, suspect: p.suspect,
        overnight: sd !== ed, location: sess.location || '',
        checkInTs: sess.checkIn, checkOutTs: sess.checkOut,
      };
      (doneByEmp[empId] = doneByEmp[empId] || []).push(rec);
      (startedOn[empId + '|' + sd] = startedOn[empId + '|' + sd] || []).push(rec);
      if (sd !== ed) (endedOn[empId + '|' + ed] = endedOn[empId + '|' + ed] || []).push(rec);
    });

    // A check-in with no check-out still means he was here — worth saying
    // so, rather than reporting the day as empty.
    // An ORPHANED check-in counts too: he checked in on 20-09, never
    // checked out, and checked in again on 21-09 — the 20-09 one is still
    // a day he started and nobody has closed.
    (s.orphaned || []).concat(s.open ? [s.open] : []).forEach(o => {
      if (!o || !o.checkIn) return;
      const sd = acNormDate_(o.checkIn);
      if (!startedOn[empId + '|' + sd]) openOn[empId + '|' + sd] = { inHm: acHHmm_(o.checkIn) };
    });
  });

  return { startedOn: startedOn, endedOn: endedOn, openOn: openOn, doneByEmp: doneByEmp };
}

// Returns { session, ownerDate, spillover, ambiguous } or null.
// AMBIGUOUS means more than one session could claim it — two overnight
// sessions ending the same morning. Those are never moved automatically;
// a human decides.
// submittedTs (optional): when he tapped the claim at the kiosk. It is the
// strongest evidence there is — he was standing there at that moment.
function otcResolveClaim_(idx, empId, dStr, submittedTs) {
  if (!idx) return null;
  const own = idx.startedOn[empId + '|' + dStr];
  // You cannot claim overtime for a shift that has not started yet. A claim
  // dated 19-09 but tapped at 00:05 on 19-09 is for the shift that had just
  // ENDED (18-09's), not for the 19-09 shift that starts at 12:30 — so its
  // own-date session does not own it. Without this, every nightly claimant
  // looked like he had two claims on one shift.
  const ownStarted = (own || []).filter(o => !submittedTs || o.checkInTs <= new Date(submittedTs.getTime() + 2 * 3600000));
  if (ownStarted.length) {
    return { session: ownStarted[0], ownerDate: dStr, spillover: false, ambiguous: ownStarted.length > 1 };
  }
  // He checked in on this date and has not checked out yet. The claim is
  // that shift's — it is NOT the tail of last night's shift just because
  // last night's is the only one that happens to be complete.
  // (Nishan Rai, 20-09: checked in 11:32, claimed 6h at 02:36. Without this
  //  the claim was read as belonging to 19-09 and would have been moved.)
  if (idx.openOn && idx.openOn[empId + '|' + dStr]) {
    return { session: null, ownerDate: dStr, spillover: false, ambiguous: false, open: true,
             openInHm: idx.openOn[empId + '|' + dStr].inHm };
  }
  // THE CLAIM BELONGS TO THE SHIFT HE HAD JUST FINISHED (21-09-2026).
  // A claim dated on a day with no shift of its own is attached to the
  // last shift that ended before he submitted it. Covers both ways the
  // kiosk date goes wrong:
  //   submitted after midnight → dated the NEXT day (Deb Bahadur: shift
  //     19-09, claimed 00:10 on 20-09, his day off)
  //   submitted late at night → dated the day AFTER, or picked the day
  //     BEFORE (Suman B.: worked 10-09, claimed 23:41 on 10-09 for "09-09",
  //     his day off)
  //   submitted on his day off for yesterday (Uttam: worked 11-09, claimed
  //     14:30 on 12-09, his day off)
  // Guard rails: the shift must have ended no more than 20h before the
  // claim (or up to 2h after — some tap OT just before Check Out), and must
  // have STARTED within a day of the claim date. A claim filed days later
  // is not guessed at.
  if (submittedTs && idx.doneByEmp && idx.doneByEmp[empId]) {
    const sub = submittedTs.getTime();
    const cd = otc_toDate_(dStr);
    let best = null;
    idx.doneByEmp[empId].forEach(rec => {
      const end = rec.checkOutTs.getTime();
      if (end > sub + 2 * 3600000 || end < sub - 20 * 3600000) return;
      const sd = otc_toDate_(rec.startDate);
      if (!cd || !sd || Math.abs(sd - cd) > 86400000 + 3600000) return;
      if (!best || end > best.checkOutTs.getTime()) best = rec;
    });
    if (best && best.startDate !== dStr) {
      return { session: best, ownerDate: best.startDate, spillover: true, ambiguous: false,
               byClaimTime: true };
    }
  }
  const spill = idx.endedOn[empId + '|' + dStr];
  if (spill && spill.length === 1) {
    return { session: spill[0], ownerDate: spill[0].startDate, spillover: true, ambiguous: false };
  }
  if (spill && spill.length > 1) {
    return { session: null, ownerDate: '', spillover: false, ambiguous: true };
  }
  return null;
}

// Previous calendar day, dd-mm-yyyy in and out.
function otc_fmtTs_(v) {
  const d = otc_submittedTs_(v);
  return d ? Utilities.formatDate(d, tz_(), 'dd-MM HH:mm') : String(v || '');
}
function otc_toDate_(dStr) {
  const p = String(dStr || '').split('-');
  if (p.length !== 3) return null;
  const d = new Date(+p[2], +p[1] - 1, +p[0]);
  return isNaN(d.getTime()) ? null : d;
}
// OT_Pending "Submitted" cell → Date (the kiosk writes dd/MM/yyyy HH:mm:ss).
function otc_submittedTs_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v;
  if (typeof parseLogTimestamp === 'function') return parseLogTimestamp(v);
  return null;
}

function otc_prevDay_(dStr) {
  const p = String(dStr).split('-');
  if (p.length !== 3) return '';
  const d = new Date(+p[2], +p[1] - 1, +p[0] - 1);
  const pad = (n) => String(n).padStart(2, '0');
  return pad(d.getDate()) + '-' + pad(d.getMonth() + 1) + '-' + d.getFullYear();
}

// ── Attribution ──────────────────────────────────────────────────────
function otc_by_(params) {
  const by = String((params && params.by) || '').trim();
  return by || 'Unknown';
}

// ── Build the verification list ──────────────────────────────────────
function getOtCompileList(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    const from = params.from ? otc_norm_(params.from) : '';
    const to   = params.to   ? otc_norm_(params.to)   : '';
    const inRange = (dStr) => {
      if (!from && !to) return true;
      const iso = (s) => { const p = s.split('-'); return p[2]+p[1]+p[0]; };
      if (from && iso(dStr) < iso(from)) return false;
      if (to   && iso(dStr) > iso(to))   return false;
      return true;
    };

    // Attendance_Daily gives punch CONTEXT — never the authority here.
    const ctx = {};
    const ad = ss.getSheetByName('Attendance_Daily');
    if (ad && ad.getLastRow() > 1) {
      const AH = ad.getRange(1,1,1,ad.getLastColumn()).getValues()[0];
      const ac = (n) => AH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
      const aId=ac('Employee ID'), aDate=ac('Date'), aIn=ac('Check In'), aOut=ac('Check Out'),
            aPres=ac('Presence (hrs)'), aOt=ac('OT Hours'), aStatus=ac('Status'),
            aShift=ac('Shift'), aLoc=ac('Location'), aSrc=ac('Source');
      ad.getRange(2,1,ad.getLastRow()-1,ad.getLastColumn()).getValues().forEach(r => {
        const id = String(r[aId]||'').trim(); if (!id) return;
        ctx[id + '|' + otc_norm_(r[aDate])] = {
          checkOutRaw: aOut >= 0 ? String(r[aOut] || '') : '',
          checkIn:  aIn   >= 0 ? otc_clock_(r[aIn])  : '',
          checkOut: aOut  >= 0 ? otc_clock_(r[aOut]) : '',
          presence: aPres >= 0 ? (Number(r[aPres]) || 0) : 0,
          computed: aOt   >= 0 ? (Number(r[aOt])   || 0) : 0,
          status:   aStatus>= 0 ? String(r[aStatus]|| '').trim() : '',
          shift:    aShift >= 0 ? String(r[aShift] || '').trim() : '',
          location: aLoc   >= 0 ? String(r[aLoc]   || '').trim() : '',
          source:   aSrc   >= 0 ? String(r[aSrc]   || '').trim() : '',
        };
      });
    }

    const pen = ss.getSheetByName(OTC_PENDING_TAB);
    if (!pen || pen.getLastRow() < 2) return { success:true, rows:[] };

    // One sessionize pass for the whole log, shared by every claim below.
    const sessIdx = otcSessionIndex_(ss);
    // What OT_Entries already holds — the compiler's entry wins.
    const entIdx = otc_entriesIndex_(ss);

    const vc = otc_verdictCols_(pen);
    const data = pen.getRange(2,1,pen.getLastRow()-1,pen.getLastColumn()).getValues();
    const rows = [];

    data.forEach((r, i) => {
      const status = String(r[OTC_COL.status] || '').trim();
      if (status !== 'Pending') return;              // already decided
      const empId = String(r[OTC_COL.empId] || '').trim();
      const dStr  = otc_norm_(r[OTC_COL.date]);
      if (!empId || !dStr || !inRange(dStr)) return;

      const claimed = Number(r[OTC_COL.hours]) || 0;
      const k = empId + '|' + dStr;
      // NOT const: a spillover claim is re-pointed at the session that
      // actually produced it, a few lines below.
      let cx = ctx[k] || { checkIn:'', checkOut:'', presence:0, computed:0,
                           status:'', shift:'', location:'' };

      // FLAG describes how the claim relates to the punches. It is
      // information for the compiler — it never decides anything.
      // An OT claim filed on a date with no session is usually one of two
      // things, and they need completely different handling:
      //
      //   OVERNIGHT SPILLOVER — the session STARTED the previous day and
      //   ended in the early hours of the claim date. The employee files
      //   the OT on the date he is standing there submitting it. Example:
      //   in 17:30 on 03-09, out 03:30 on 04-09, files OT for 04-09 — and
      //   04-09 is his day off, so it has no punches of its own. The OT is
      //   real; it belongs to the PREVIOUS day's session.
      //
      //   WORKED A DAY OFF — came in on a scheduled day off, never punched
      //   at all, and claimed the whole shift as OT. Nothing in the punch
      //   data can confirm or deny it. Only a supervisor can.
      // Resolve the claim to the session that produced it. When it spilled
      // over from the night before, JUDGE IT AGAINST THAT SESSION — not
      // against a date that has no punches because the man was asleep.
      const res = otcResolveClaim_(sessIdx, empId, dStr, otc_submittedTs_(r[OTC_COL.submitted]));
      const spill = !!(res && res.spillover && res.session && !res.session.suspect);
      const ownerDate = spill ? res.ownerDate : dStr;
      if (spill) {
        const oc = ctx[empId + '|' + ownerDate];
        if (oc) cx = oc;         // presence, computed OT, shift, in/out all move with it
      }

      let flag, flagWhy;
      if (spill) {
        flag = 'overnight-spillover';
        flagWhy = (res.byClaimTime
                    ? 'Claimed at the kiosk ' + otc_fmtTs_(r[OTC_COL.submitted]) + ', straight after the shift of ' +
                      ownerDate + ' (' + res.session.inHm + ' → ' + res.session.outHm + ')'
                    : 'The shift started ' + ownerDate + ' at ' + res.session.inHm +
                      ' and finished ' + res.session.endDate + ' at ' + res.session.outHm) +
                  ' — ' + otc_hm_(res.session.presence) + ' worked, ' + otc_hm_(cx.computed) +
                  ' over the shift. This claim belongs to that shift, not to ' + dStr +
                  '. Run commitOtDateFix() to file it on ' + ownerDate + '.';
      } else if (res && res.open) {
        flag = 'incomplete-punch';
        flagWhy = 'Checked in at ' + res.openInHm + ' on ' + dStr + ' and never checked out — ' +
                  'his attendance for this day has to be answered first (supervisor page), ' +
                  'then the OT can be measured.';
      } else if (res && res.ambiguous && !res.session) {
        flag = 'ambiguous-session';
        flagWhy = 'More than one overnight session ended on ' + dStr +
                  ' — which one this claim belongs to cannot be decided from the punches';
      } else if (cx.status === 'Day Off' || cx.status === 'Not Recorded') {
        flag = 'worked-day-off';
        flagWhy = 'Attendance says "' + cx.status + '" with no punches — claiming a ' +
                  'whole shift worked on a day off. Only a supervisor can confirm this';
      } else if (!cx.status) {
        flag = 'no-attendance';
        flagWhy = 'No Attendance_Daily row for this date at all';
      } else if (cx.status !== 'Present') {
        flag = 'not-present';
        flagWhy = 'Attendance says "' + cx.status + '", not Present';
      } else if (OTC_TIMELESS_SOURCES.indexOf(cx.source) !== -1) {
        flag = 'no-punch-data';
        flagWhy = 'Marked Present via ' + cx.source + ' with no punch times — ' +
                  'he never punched and was not expected to, so the punches can ' +
                  'neither support nor contradict this claim';
      } else if (!cx.checkOut) {
        flag = 'incomplete-punch';
        flagWhy = 'Check-in with no check-out — presence cannot be computed';
      } else if (cx.computed === 0 && claimed > 0) {
        flag = 'no-ot-in-punches';
        // Spell out the arithmetic. "Full Day" in Day Type only means
        // presence reached half the shift — it does NOT mean OT was earned.
        // OT is presence MINUS the shift's normal hours, so a day can be
        // Present and Full Day and still have zero overtime.
        flagWhy = 'Worked ' + otc_hm_(cx.presence) + ' against a ' +
                  (cx.shift || 'standard') + ' shift, so no overtime was earned. ' +
                  'The claim of ' + otc_hm_(claimed) + ' would mean about ' +
                  otc_hm_(cx.presence + claimed) + ' on shift.';
      } else if (claimed > cx.computed + 0.02 && (claimed - cx.computed) < 1 &&
                 Math.abs(claimed - Math.round(claimed)) < 0.001 && cx.computed > 0) {
        // ROUNDED UP TO THE NEXT WHOLE HOUR. The punches say 5h 17m and the
        // claim arrives as a flat 6h. It is not a dispute about what
        // happened — it is 43 minutes of arithmetic, on every man, every
        // week. Separated from a real over-claim because the remedy is
        // different: pay the punched figure.
        flag = 'rounded-up';
        flagWhy = 'Punches show ' + otc_hm_(cx.computed) + ' of overtime; the claim is a flat ' +
                  otc_hm_(claimed) + ' — ' + otc_hm_(claimed - cx.computed) +
                  ' more. Overtime is paid on actual time, not rounded up to the next hour.';
      } else if (claimed > cx.computed + 0.25) {
        flag = 'claim-above-punches';
        flagWhy = 'Claim exceeds computed by ' + otc_hm_(claimed - cx.computed);
      } else if (claimed < cx.computed - 0.25) {
        flag = 'claim-below-punches';
        flagWhy = 'Claim is ' + otc_hm_(cx.computed - claimed) + ' LESS than computed';
      } else {
        flag = 'consistent';
        flagWhy = 'Claim matches the punches';
      }

      // Already in OT_Entries for this man and this shift? Then that entry
      // stands and this claim must not be paid on top of it.
      const posted = entIdx[empId.toUpperCase() + '|' + ownerDate] || entIdx[empId.toUpperCase() + '|' + dStr] || null;
      if (posted) {
        flag = 'already-posted';
        flagWhy = 'OT_Entries already has ' + otc_hm_(posted.hours) + ' for this day — ' + otc_entryText_(posted) +
                  '. That entry stands; approving this claim would pay the day twice. ' +
                  'commitOtEntriesReconcile() closes it.';
      }

      // What the supervisor said, straight off this same row. When he has
      // answered, that is the best evidence there is — the punches could not
      // settle these claims, which is why he was asked.
      const supVerdict = vc.verdict >= 0 ? String(r[vc.verdict] || '').trim() : '';
      const supHours   = (vc.hours >= 0 && r[vc.hours] !== '' && r[vc.hours] != null)
                           ? Number(r[vc.hours]) : null;
      const supBy      = vc.by >= 0 ? String(r[vc.by] || '').trim() : '';
      const supOn      = vc.on >= 0 ? String(r[vc.on] || '').trim() : '';

      rows.push({
        row: i + 2,
        otId: String(r[OTC_COL.id] || ''),
        supVerdict: supVerdict,
        supHours: supHours,
        supHoursHm: supHours == null ? '' : otc_hm_(supHours),
        supBy: supBy,
        supOn: supOn,
        // Pre-fill the compiler's decision with the supervisor's figure when
        // there is one, otherwise with the claim.
        suggested: supHours != null ? supHours : claimed,
        empId: empId,
        name: String(r[OTC_COL.name] || '').trim(),
        date: dStr,
        location: cx.location,
        shift: cx.shift,
        claimed: claimed,
        claimedHm: otc_hm_(claimed),
        computed: cx.computed,
        computedHm: otc_hm_(cx.computed),
        presence: cx.presence,
        presenceHm: otc_hm_(cx.presence),
        checkIn: cx.checkIn,
        checkOut: cx.checkOut,
        attStatus: cx.status,
        variance: Math.round((claimed - cx.computed) * 100) / 100,
        varianceHm: otc_hm_(claimed - cx.computed),
        flag: flag,
        flagWhy: flagWhy,
        // The shift this claim really belongs to. Equal to the claim date
        // for an ordinary day; the night before for a spillover.
        belongsTo: ownerDate,
        spillover: spill,
        sessionIn:  spill ? res.session.inHm  : (cx.checkIn  || ''),
        sessionOut: spill ? res.session.outHm : (cx.checkOut || ''),
        prevDate: spill ? ownerDate : '',
        prevComputed: spill ? cx.computed : 0,
        prevComputedHm: spill ? otc_hm_(cx.computed) : '',
        submitted: String(r[OTC_COL.submitted] || ''),
        postedHours: posted ? posted.hours : null,
        postedHm: posted ? otc_hm_(posted.hours) : '',
        postedText: posted ? otc_entryText_(posted) : '',
      });
    });

    rows.sort((a,b) => {
      const iso = (s) => { const p = s.split('-'); return p[2]+p[1]+p[0]; };
      return (iso(a.date)+a.name).localeCompare(iso(b.date)+b.name);
    });

    const summary = {};
    rows.forEach(r => { summary[r.flag] = (summary[r.flag] || 0) + 1; });

    const verified = rows.filter(r => r.supVerdict).length;

    return { success:true, rows:rows, count:rows.length, summary:summary,
             verified: verified, unverified: rows.length - verified };
  } catch (err) {
    return { success:false, error: err.toString() };
  }
}

// ── Apply decisions ──────────────────────────────────────────────────
// decisions: [{ row, decision:'approve'|'reject'|'adjust', hours, note }]
// Approved/adjusted rows are written to OT_Entries and the OT_Pending row
// is marked. Rejected rows are marked only — nothing reaches payroll.
function postApprovedOt(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const by = otc_by_(params);   // who is making these decisions — for the audit trail

    let decisions;
    try { decisions = typeof params.decisions === 'string'
            ? JSON.parse(params.decisions) : params.decisions; }
    catch (e) { return { success:false, error:'Invalid decisions payload.' }; }
    if (!decisions || !decisions.length) return { success:false, error:'No decisions submitted.' };

    const pen = ss.getSheetByName(OTC_PENDING_TAB);
    let ot = ss.getSheetByName(OTC_ENTRIES_TAB);
    if (!ot) {
      ot = ss.insertSheet(OTC_ENTRIES_TAB);
      ot.appendRow(['Date','Employee ID','Name','OT Hours','Remark','Added By']);
    }
    // Written by header, into the layout OT_Entries actually has.
    const ec = otc_entryCols_(ot);
    const entIdx = otc_entriesIndex_(ss);
    let settled = 0;
    const settledList = [];

    const now = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm:ss');
    let approved = 0, rejected = 0, totalHours = 0;
    const errors = [];

    decisions.forEach(d => {
      const r = parseInt(d.row);
      if (!r || r < 2 || r > pen.getLastRow()) { errors.push('bad row ' + d.row); return; }

      const cur = pen.getRange(r, 1, 1, pen.getLastColumn()).getValues()[0];
      if (String(cur[OTC_COL.status] || '').trim() !== 'Pending') {
        errors.push(String(cur[OTC_COL.empId]) + ' ' + otc_norm_(cur[OTC_COL.date]) +
                    ' is no longer Pending — skipped');
        return;
      }

      const empId = String(cur[OTC_COL.empId] || '').trim();
      const name  = String(cur[OTC_COL.name]  || '').trim();
      const dStr  = otc_norm_(cur[OTC_COL.date]);

      // Carry the supervisor's statement into the audit trail, so a paid hour
      // can always be traced back to who vouched for it.
      const vcd = otc_verdictCols_(pen);
      const sup = vcd.verdict >= 0 ? String(cur[vcd.verdict] || '').trim() : '';
      const supNote = sup
        ? ' — supervisor: ' + sup +
          (vcd.hours >= 0 && cur[vcd.hours] !== '' ? ' (' + cur[vcd.hours] + 'h)' : '') +
          (vcd.by >= 0 && cur[vcd.by] ? ' per ' + String(cur[vcd.by]).trim() : '')
        : ' — not verified by a supervisor';

      if (d.decision === 'reject') {
        pen.getRange(r, OTC_COL.status + 1).setValue('Rejected');
        if (typeof logAmend_ === 'function')
          logAmend_(ss, by, 'OT REJECT', empId, dStr, 'OT',
                    'claimed ' + (cur[OTC_COL.hours]||0) + 'h — ' +
                    (d.note || 'no reason given') + supNote);
        rejected++;
        return;
      }

      // approve or adjust
      const hours = Number(d.hours);
      if (isNaN(hours) || hours <= 0) { errors.push(empId + ' ' + dStr + ': invalid hours'); return; }

      // Already in OT_Entries for this man and day? That entry wins — close
      // the claim against it instead of paying the day twice.
      const ek = empId.toUpperCase() + '|' + dStr;
      if (entIdx[ek]) {
        pen.getRange(r, OTC_COL.status + 1).setValue(OTC_SETTLED);
        if (typeof logAmend_ === 'function')
          logAmend_(ss, by, 'OT SETTLED', empId, dStr, 'OT',
                    'claimed ' + (cur[OTC_COL.hours]||0) + 'h — not posted: OT_Entries already has ' +
                    otc_entryText_(entIdx[ek]) + supNote);
        settledList.push(empId + ' ' + dStr + ': already in OT_Entries (' + otc_hm_(entIdx[ek].hours) +
                         ') — that entry stands, claim closed');
        settled++;
        return;
      }

      const remark = 'Kiosk claim (verified) on ' + now +
                     (d.decision === 'adjust' ? ' — ADJUSTED from ' + (cur[OTC_COL.hours]||0) + 'h' : '') +
                     (d.note ? ' — ' + d.note : '') + supNote;
      const out = new Array(Math.max(ec.width, 6)).fill('');
      out[ec.date] = dStr; out[ec.empId] = empId; out[ec.name] = name; out[ec.hours] = hours;
      if (ec.remark >= 0) out[ec.remark] = remark;
      if (ec.by >= 0) out[ec.by] = 'OT approval (' + by + ')';
      if (ec.source >= 0) out[ec.source] = 'Kiosk claim (verified)';
      ot.appendRow(out);
      entIdx[ek] = { hours: hours, rows: [{ hours: hours, remark: remark, by: 'OT approval (' + by + ')' }] };

      pen.getRange(r, OTC_COL.status + 1).setValue(
        d.decision === 'adjust' ? 'Approved (adjusted)' : 'Approved');

      if (typeof logAmend_ === 'function')
        logAmend_(ss, by, 'OT APPROVE', empId, dStr, 'OT',
                  'claimed ' + (cur[OTC_COL.hours]||0) + 'h → paid ' + hours + 'h' +
                  (d.note ? ' — ' + d.note : '') + supNote);

      approved++; totalHours += hours;
    });

    return {
      success: true,
      approved: approved, rejected: rejected,
      totalHours: Math.round(totalHours * 100) / 100,
      totalHoursHm: otc_hm_(totalHours),
      errors: errors,
      settled: settled,
      settledList: settledList,
      message: approved + ' approved (' + otc_hm_(totalHours) + '), ' +
               rejected + ' rejected.' +
               (settled ? ' ' + settled + ' already in OT_Entries — those entries stand, claims closed.' : '') +
               (errors.length ? ' ' + errors.length + ' skipped.' : ''),
    };
  } catch (err) {
    return { success:false, error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════
// MOVE A CLAIM ONTO THE SHIFT THAT PRODUCED IT
// ------------------------------------------------------------
// Showing the right session on screen is not enough. The claim is still
// DATED the 14th, so payroll files it under the 14th — and when the 14th
// and the 13th fall either side of a payroll cut-off, the hours land in
// the wrong month. The date on the row has to be right.
//
// This moves only the unambiguous ones:
//   the claim date has NO session of its own, AND
//   exactly ONE overnight session ended on it, AND
//   that session is not flagged suspect, AND
//   the employee has no other pending claim on the target date
// Anything else is left for a human — including two overnight sessions
// ending the same morning, which the punches cannot settle.
//
// The employee's own date is kept in "Original Claim Date" so his
// statement is never silently overwritten. Preview first, always.
// ════════════════════════════════════════════════════════════
const OTC_FIX_HEADERS = ['Original Claim Date', 'Date Moved By'];

function otc_fixCols_(pen, create) {
  const read = () => {
    const H = pen.getRange(1, 1, 1, pen.getLastColumn()).getValues()[0];
    const at = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    return { orig: at('Original Claim Date'), by: at('Date Moved By') };
  };
  let c = read();
  if (create && (c.orig === -1 || c.by === -1)) {
    const H = pen.getRange(1, 1, 1, pen.getLastColumn()).getValues()[0];
    const missing = OTC_FIX_HEADERS.filter(h =>
      H.every(x => String(x).trim().toLowerCase() !== h.toLowerCase()));
    if (missing.length) {
      pen.getRange(1, pen.getLastColumn() + 1, 1, missing.length)
         .setValues([missing]).setFontWeight('bold');
      SpreadsheetApp.flush();
    }
    c = read();
  }
  return c;
}

function previewOtDateFix(params) { return otc_dateFix_(params, false); }
function commitOtDateFix(params)  { return otc_dateFix_(params, true); }

function otc_dateFix_(params, commit) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const pen = ss.getSheetByName(OTC_PENDING_TAB);
    if (!pen || pen.getLastRow() < 2) return { success:true, count:0, message:'No pending OT to check.' };

    const idx = otcSessionIndex_(ss);
    if (!idx) return { success:false, error:'AttendanceCore.gs is not deployed — cannot resolve sessions.' };

    const fc = otc_fixCols_(pen, commit);
    const data = pen.getRange(2, 1, pen.getLastRow()-1, pen.getLastColumn()).getValues();

    // Pending claims already sitting on each date, so a move never creates
    // two claims for one session.
    const pendingOn = {};
    data.forEach(r => {
      if (String(r[OTC_COL.status]||'').trim() !== 'Pending') return;
      const id = String(r[OTC_COL.empId]||'').trim();
      const d  = otc_norm_(r[OTC_COL.date]);
      if (id && d) pendingOn[id + '|' + d] = (pendingOn[id + '|' + d] || 0) + 1;
    });

    const willMove = [], leftAlone = [];
    data.forEach((r, i) => {
      if (String(r[OTC_COL.status]||'').trim() !== 'Pending') return;
      const empId = String(r[OTC_COL.empId]||'').trim();
      const name  = String(r[OTC_COL.name]||'').trim();
      const dStr  = otc_norm_(r[OTC_COL.date]);
      const hours = Number(r[OTC_COL.hours]) || 0;
      if (!empId || !dStr) return;

      const res = otcResolveClaim_(idx, empId, dStr, otc_submittedTs_(r[OTC_COL.submitted]));
      if (!res) return;                                  // no session either way
      if (!res.spillover) return;                        // has its own day
      if (res.ambiguous || !res.session) {
        leftAlone.push({ empId, name, date: dStr, why: 'more than one overnight session ended that morning' });
        return;
      }
      if (res.session.suspect) {
        leftAlone.push({ empId, name, date: dStr, why: 'that session is flagged suspect — fix the punches first' });
        return;
      }
      willMove.push({ row: i + 2, empId, name, hours,
                      from: dStr, to: res.ownerDate,
                      shift: res.session.inHm + ' → ' + res.session.outHm,
                      presenceHm: otc_hm_(res.session.presence) });
    });

    // A man who claims every night files each claim one day late, so every
    // claim moves back one day together. Judge doubling on where claims END
    // UP, not where they sit now: only two claims landing on the same shift
    // is a real double (e.g. the same claim tapped twice in one minute).
    const moving = {}; willMove.forEach(m => { moving[m.row] = m.to; });
    const landing = {};
    data.forEach((r, i) => {
      if (String(r[OTC_COL.status]||'').trim() !== 'Pending') return;
      const id = String(r[OTC_COL.empId]||'').trim();
      const d  = moving[i + 2] || otc_norm_(r[OTC_COL.date]);
      if (id && d) landing[id + '|' + d] = (landing[id + '|' + d] || 0) + 1;
    });
    for (let j = willMove.length - 1; j >= 0; j--) {
      const m = willMove[j];
      if ((landing[m.empId + '|' + m.to] || 0) > 1) {
        leftAlone.push({ empId: m.empId, name: m.name, date: m.from,
          why: 'another claim would also land on the ' + m.to + ' shift — two claims for one shift, decide which stands' });
        willMove.splice(j, 1);
      }
    }

    if (commit) {
      const by = otc_by_(params);
      const now = Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm:ss');
      willMove.forEach(m => {
        if (fc.orig >= 0 && !String(pen.getRange(m.row, fc.orig + 1).getValue() || '').trim()) {
          pen.getRange(m.row, fc.orig + 1).setValue(m.from).setNumberFormat('@');
        }
        pen.getRange(m.row, OTC_COL.date + 1).setValue(m.to).setNumberFormat('@');
        if (fc.by >= 0) pen.getRange(m.row, fc.by + 1).setValue(by + ' ' + now);
        if (typeof logAmend_ === 'function')
          logAmend_(ss, by, 'OT REDATE', m.empId, m.to, 'OT',
                    'claim of ' + m.hours + 'h moved from ' + m.from + ' to ' + m.to +
                    ' — belongs to the shift ' + m.shift);
      });
    }

    const msg = (commit ? 'Moved ' : 'Would move ') + willMove.length +
      ' overnight claim(s) onto the shift that produced them.' +
      (leftAlone.length ? ' ' + leftAlone.length + ' left for a human.' : '') +
      (commit ? '' : ' Nothing changed — run commitOtDateFix() to apply.');

    // Run from the editor, the result has to be in the Execution log to be seen.
    Logger.log([msg, ''].concat(
      willMove.map(m => '  ' + (commit ? 'MOVED ' : 'MOVE  ') + m.from + ' → ' + m.to + '  ' + m.empId + '  ' +
                        m.name + '  ' + m.hours + 'h  (shift ' + m.shift + ')'),
      leftAlone.length ? ['', 'LEFT FOR YOU (' + leftAlone.length + '):'] : [],
      leftAlone.map(l => '  ' + l.date + '  ' + l.empId + '  ' + l.name + ' — ' + l.why)).join('\n'));
    return { success:true, committed: !!commit, count: willMove.length,
             willMove: willMove, leftAlone: leftAlone, message: msg };
  } catch (err) {
    return { success:false, error: err.toString() };
  }
}


// ============================================================
// GUARD FOR "Post ticked to OT_Entries"  (router: postOtToEntries)
// ------------------------------------------------------------
// Called by the router BEFORE postOtToEntries. Drops any ticked row for a
// man and day already in OT_Entries — what is there stands; posting again
// would pay the day twice. Also drops a second pick for the same man and
// day in one batch. Returns { body, skipped[] }.
// ============================================================
function otGuardPostEntries(body) {
  body = body || {};
  let rows = body.rows;
  try { if (typeof rows === 'string') rows = JSON.parse(rows); } catch (e) { rows = []; }
  if (!Array.isArray(rows) || !rows.length) return { body: body, skipped: [] };
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const idx = otc_entriesIndex_(ss);
  const seen = {}, keep = [], skipped = [];
  rows.forEach(r => {
    const id = String(r.empId || '').trim().toUpperCase();
    const d = otc_anyDate_(r.date);
    const k = id + '|' + d;
    if (idx[k]) { skipped.push((r.name || id) + ' ' + d + ' — already in OT_Entries (' + otc_hm_(idx[k].hours) + ')'); return; }
    if (seen[k]) { skipped.push((r.name || id) + ' ' + d + ' — ticked twice'); return; }
    seen[k] = true; keep.push(r);
  });
  body.rows = keep;
  return { body: body, skipped: skipped };
}

// ============================================================
// CROSS-CHECK THE OT COMPILER'S ENTRIES WITH THE CLAIMS
// previewOtEntriesReconcile()  — writes nothing
// commitOtEntriesReconcile()   — closes the claims his entries cover
// ------------------------------------------------------------
// Rule (Shakeel, 21-09-2026): the compiler's entry wins.
//   A. A Pending claim for a man and day that OT_Entries already covers is
//      closed as OTC_SETTLED. His hours are what is paid; the claim, the
//      supervisor's answer and his figure go into Attendance_Amendments.
//   B. His entries with no claim — stand as they are. Listed.
//   C. Days he compiled, where a man claimed and he has NO entry — NOT
//      touched: they stay on OT Review. Listed, because either he left the
//      man out or the claim is wrong, and only you know which.
//   D. A man and day with two rows in OT_Entries — payroll pays both.
//      Listed; delete the wrong one by hand.
// Nothing in OT_Entries is changed by this.
// ============================================================
function previewOtEntriesReconcile() { return otc_reconcile_(false); }
function commitOtEntriesReconcile()  { return otc_reconcile_(true); }

function otc_reconcile_(commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const lock = LockService.getScriptLock();
  if (commit && !lock.tryLock(30000)) { Logger.log('Busy — try again in a minute.'); return 'busy'; }
  try {
    const out = [];
    const pad = (x, n) => { x = String(x); return x.length >= n ? x.slice(0, n - 1) + ' ' : x + ' '.repeat(n - x.length); };
    const hm = h => otc_hm_(h);
    const iso = d => { const p = String(d).split('-'); return p[2] + p[1] + p[0]; };

    const idx = otc_entriesIndex_(ss);
    const L = getOtCompileList({ from: '01-01-2000', to: '31-12-2099' });
    if (!L || !L.success) return otc_log_(['Could not read the claims: ' + (L && L.error)]);
    const claims = L.rows || [];

    // A. claims his entries cover
    const settle = [], coveredKeys = {};
    claims.forEach(r => {
      const id = r.empId.toUpperCase();
      const k = idx[id + '|' + r.belongsTo] ? id + '|' + r.belongsTo : (idx[id + '|' + r.date] ? id + '|' + r.date : '');
      if (!k) return;
      coveredKeys[k] = true;
      settle.push({ r: r, k: k, e: idx[k] });
    });
    // B. his entries nobody claimed
    const claimKeys = {};
    claims.forEach(r => { claimKeys[r.empId.toUpperCase() + '|' + r.belongsTo] = true; claimKeys[r.empId.toUpperCase() + '|' + r.date] = true; });
    // Already-decided claims count too (an approved claim is not "unclaimed").
    const pen = ss.getSheetByName(OTC_PENDING_TAB);
    const penData = pen && pen.getLastRow() > 1 ? pen.getRange(2, 1, pen.getLastRow() - 1, pen.getLastColumn()).getValues() : [];
    penData.forEach(r => { claimKeys[String(r[OTC_COL.empId]).trim().toUpperCase() + '|' + otc_norm_(r[OTC_COL.date])] = true; });
    const noClaim = Object.keys(idx).filter(k => !claimKeys[k]);
    // C. days he compiled; claims there with no entry of his
    const compiledDays = {};
    Object.keys(idx).forEach(k => { if (idx[k].rows.some(x => x.compiler && /auto from attendance/i.test(x.remark))) compiledDays[k.split('|')[1]] = true; });
    const missed = claims.filter(r => compiledDays[r.belongsTo] && !coveredKeys[r.empId.toUpperCase() + '|' + r.belongsTo] &&
                                      !coveredKeys[r.empId.toUpperCase() + '|' + r.date]);
    // D. two rows for one man and day
    const twice = Object.keys(idx).filter(k => idx[k].rows.length > 1);

    // punches for context
    const punch = {};
    const ad = ss.getSheetByName('Attendance_Daily');
    if (ad && ad.getLastRow() > 1) {
      const H = ad.getRange(1, 1, 1, ad.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const cI = H.indexOf('Employee ID'), cD = H.indexOf('Date'), cO = H.indexOf('OT Hours');
      if (cI >= 0 && cD >= 0 && cO >= 0)
        ad.getRange(2, 1, ad.getLastRow() - 1, ad.getLastColumn()).getValues().forEach(r => {
          punch[String(r[cI]).trim().toUpperCase() + '|' + otc_norm_(r[cD])] = Number(r[cO]) || 0;
        });
    }
    const names = {};
    claims.forEach(r => { names[r.empId.toUpperCase()] = r.name; });

    out.push((commit ? 'DONE' : 'PREVIEW — nothing written yet') + ' — OT compiler\'s entries vs the claims');
    out.push('Compiler: ' + OTC_COMPILER_EMAILS.join(', ') + '.  His entry wins.');
    out.push('');

    settle.sort((a, b) => (iso(a.k.split('|')[1]) + a.k).localeCompare(iso(b.k.split('|')[1]) + b.k));
    let same = 0, lower = 0, higher = 0, tClaim = 0, tEntry = 0;
    const seenK = {};
    settle.forEach(x => {
      tClaim += x.r.claimed;
      if (!seenK[x.k]) { tEntry += x.e.hours; seenK[x.k] = true; }
      const dlt = x.e.hours - x.r.claimed;
      if (Math.abs(dlt) < 0.1) same++; else if (dlt < 0) lower++; else higher++;
    });
    out.push('A. CLAIMS HIS ENTRY ALREADY COVERS (' + settle.length + ') — ' + (commit ? 'closed' : 'will be closed') +
             ' as "' + OTC_SETTLED + '", his hours stand:');
    out.push('   claimed ' + hm(tClaim) + ' · his entries ' + hm(tEntry) +
             ' · ' + same + ' same, ' + lower + ' his lower, ' + higher + ' his higher');
    out.push('   ' + pad('date', 12) + pad('man', 30) + pad('claimed', 9) + pad('supervisor', 16) + pad('HIS ENTRY', 11) + 'punches now');
    settle.forEach(x => {
      const [id, d] = x.k.split('|');
      out.push('   ' + pad(d, 12) + pad(id + ' ' + x.r.name, 30) + pad(x.r.claimedHm, 9) +
               pad(x.r.supVerdict ? x.r.supVerdict + ' ' + x.r.supHoursHm : '—', 16) +
               pad(hm(x.e.hours), 11) + hm(punch[x.k] || 0) +
               (Math.abs((punch[x.k] || 0) - x.e.hours) > 0.25 ? '   (punches have changed since he posted)' : ''));
    });
    out.push('');

    let tNo = 0; noClaim.forEach(k => { tNo += idx[k].hours; });
    out.push('B. HIS ENTRIES WITH NO CLAIM (' + noClaim.length + ', ' + hm(tNo) + ') — stand as they are.');
    out.push('');

    out.push('C. DAYS HE COMPILED (' + Object.keys(compiledDays).sort((a, b) => iso(a).localeCompare(iso(b))).join(', ') +
             ') — MEN WHO CLAIMED BUT HAVE NO ENTRY OF HIS (' + missed.length + '):');
    out.push('   Not touched. They stay on OT Review for your decision.');
    missed.sort((a, b) => (iso(a.belongsTo) + a.empId).localeCompare(iso(b.belongsTo) + b.empId)).forEach(r => {
      const k = r.empId.toUpperCase() + '|' + r.belongsTo;
      out.push('   ' + pad(r.belongsTo, 12) + pad(r.empId + ' ' + r.name, 30) + pad('claimed ' + r.claimedHm, 17) +
               pad('punches ' + hm(punch[k] || 0), 17) + (r.supVerdict ? 'supervisor ' + r.supVerdict + ' ' + r.supHoursHm : 'supervisor —'));
    });
    out.push('');

    out.push('D. TWO ROWS IN OT_ENTRIES FOR ONE MAN AND DAY (' + twice.length + ') — payroll pays both; delete the wrong row by hand:');
    twice.forEach(k => {
      const [id, d] = k.split('|');
      out.push('   ' + pad(d, 12) + pad(id + ' ' + (names[id] || ''), 30) +
               idx[k].rows.map(x => 'row ' + x.row + ': ' + hm(x.hours) + ' "' + x.remark + '"').join('   '));
    });

    if (!commit) {
      out.push('');
      out.push(settle.length ? 'Nothing has been changed. If A is right, run commitOtEntriesReconcile().'
                             : 'Nothing to close.');
      return otc_log_(out);
    }

    const who = 'System (compiler entry wins)';
    settle.forEach(x => {
      pen.getRange(x.r.row, OTC_COL.status + 1).setValue(OTC_SETTLED);
      if (typeof logAmend_ === 'function')
        logAmend_(ss, who, 'OT SETTLED', x.r.empId, x.k.split('|')[1], 'OT',
          'claimed ' + x.r.claimedHm + (x.r.supVerdict ? ' (supervisor ' + x.r.supVerdict + ' ' + x.r.supHoursHm + ' per ' + x.r.supBy + ')' : '') +
          ' — closed: OT_Entries has ' + otc_entryText_(x.e));
    });
    SpreadsheetApp.flush();
    out.push('');
    out.push('Closed ' + settle.length + ' claim(s). OT_Entries unchanged.');
    return otc_log_(out);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
function otc_log_(lines) { const t = lines.join('\n'); Logger.log(t); return t; }

// ============================================================
// PAYROLL OT CLOSE — previewPayrollOtClose() / commitPayrollOtClose()
// ------------------------------------------------------------
// Decides EVERY pending OT claim up to yesterday by fixed rules, so payroll
// can run without waiting for answers that will not come in time.
// One decision per man per shift. Rules, in order:
//
//   0. OT_Entries already has the day      → closed, that entry stands
//                                             (the compiler's entry wins).
//   1. The shift has complete punches      → PAY THE PUNCHES:
//                                             OT = presence − shift hours.
//                                             Claim above → punches paid.
//                                             Claim below → the claim is
//                                             paid (never more than he
//                                             asked for — a big gap is
//                                             nearly always a check-out he
//                                             forgot). Listed. Set
//                                             OTC_CLOSE_PAY_ABOVE_CLAIM =
//                                             true to pay the punches.
//                                             Punches show 0 → rejected.
//   2. Punches cannot measure it (open check-in, no punches, manual day)
//      a. supervisor answered Confirmed/Partly → his hours are paid
//      b. supervisor answered "Not worked"      → rejected
//      c. nobody answered                       → HELD: stays Pending, not
//                                                 paid now, comes back next
//                                                 payroll. Nothing is lost.
//   Two or more claims for one shift → the shift is paid ONCE; the other
//   claims are rejected as duplicates of it.
//
// Every paid hour goes to OT_Entries with the rule that paid it; every
// closed claim goes to Attendance_Amendments. Preview first — it writes
// nothing and shows the totals per rule.
// ============================================================
const OTC_CLOSE_FROM = '22-08-2026';
const OTC_CLOSE_PAY_ABOVE_CLAIM = false;

function previewPayrollOtClose(params) { return otc_payrollClose_(params || {}, false); }
function commitPayrollOtClose(params)  { return otc_payrollClose_(params || {}, true); }

function otc_payrollClose_(params, commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const lock = LockService.getScriptLock();
  if (commit && !lock.tryLock(30000)) { Logger.log('Busy — try again in a minute.'); return 'busy'; }
  try {
    const iso = d => { const p = String(d).split('-'); return p[2] + p[1] + p[0]; };
    const y = new Date(); y.setHours(0, 0, 0, 0); y.setDate(y.getDate() - 1);
    const from = params.from ? otc_norm_(params.from) : OTC_CLOSE_FROM;
    const to   = params.to   ? otc_norm_(params.to)   : formatDdMmYyyy_(y);
    if (iso(to) > iso(formatDdMmYyyy_(y))) return otc_log_(['Not past yesterday — today is not over.']);
    let who = 'HR';
    try { who = Session.getActiveUser().getEmail() || 'HR'; } catch (e) {}

    const L = getOtCompileList({ from: '01-01-2000', to: to });
    if (!L || !L.success) return otc_log_(['Could not read the claims: ' + (L && L.error)]);
    const rows = (L.rows || []).filter(r => iso(r.belongsTo) >= iso(from) && iso(r.belongsTo) <= iso(to));
    const entIdx = otc_entriesIndex_(ss);

    const MEASURED = ['consistent', 'claim-above-punches', 'claim-below-punches', 'rounded-up',
                      'no-ot-in-punches', 'overnight-spillover'];
    const groups = {};
    rows.forEach(r => { const k = r.empId.toUpperCase() + '|' + r.belongsTo; (groups[k] = groups[k] || []).push(r); });

    const pay = [], reject = [], settle = [], held = [];
    const T = { punches: 0, sup: 0, claimedPaid: 0 };
    const C = { punches: 0, abovePunches: 0, belowPunches: 0, zero: 0, sup: 0, supNo: 0, held: 0, dup: 0, settled: 0 };
    let claimedAbove = 0, heldHours = 0, notAsked = 0;
    const below = [];

    Object.keys(groups).sort((a, b) => (iso(a.split('|')[1]) + a).localeCompare(iso(b.split('|')[1]) + b)).forEach(k => {
      const g = groups[k];
      const d = k.split('|')[1];
      const maxClaim = Math.max.apply(null, g.map(r => r.claimed));
      const dupNote = g.length > 1 ? ' (' + g.length + ' claims for this shift — paid once)' : '';

      // 0. already in OT_Entries
      if (entIdx[k] || g.some(r => r.flag === 'already-posted')) {
        g.forEach(r => settle.push({ r: r, why: 'OT_Entries already has ' + otc_hm_((entIdx[k] || {}).hours || r.postedHours || 0) }));
        C.settled += g.length; return;
      }

      const measured = MEASURED.indexOf(g[0].flag) !== -1;
      let hours = 0, why = '', payIt = false;
      if (measured) {
        hours = Math.round((g[0].computed || 0) * 100) / 100;
        if (hours <= 0.01) {
          g.forEach(r => reject.push({ r: r, why: 'Punches show no OT — worked ' + r.presenceHm + ' on a ' + (r.shift || 'standard') + ' shift' }));
          C.zero += g.length; return;
        }
        payIt = true;
        const punchOt = hours;
        if (maxClaim > punchOt + 0.02) { C.abovePunches++; claimedAbove += maxClaim - punchOt; }
        else if (maxClaim < punchOt - 0.02) {
          C.belowPunches++; notAsked += punchOt - maxClaim;
          below.push({ r: g[0], punchOt: punchOt, claim: maxClaim });
          if (!OTC_CLOSE_PAY_ABOVE_CLAIM) hours = maxClaim;
        }
        why = 'Paid on punches: ' + otc_hm_(hours) + ' (presence ' + g[0].presenceHm + ', punches OT ' + otc_hm_(punchOt) +
              ', claimed ' + otc_hm_(maxClaim) + ')' + dupNote;
        C.punches++; T.punches += hours;
      } else {
        const ans = g.filter(r => r.supVerdict);
        const yes = ans.filter(r => /^(confirmed|partly)$/i.test(r.supVerdict) && r.supHours > 0);
        if (yes.length) {
          hours = Math.max.apply(null, yes.map(r => r.supHours));
          payIt = true;
          why = 'Paid on supervisor: ' + yes[0].supVerdict + ' ' + otc_hm_(hours) + ' per ' + yes[0].supBy +
                ' — punches cannot measure it (' + g[0].flag + ')' + dupNote;
          C.sup++; T.sup += hours;
        } else if (ans.length) {
          g.forEach(r => reject.push({ r: r, why: 'Supervisor: ' + r.supVerdict + (r.supBy ? ' per ' + r.supBy : '') }));
          C.supNo += g.length; return;
        } else {
          g.forEach(r => held.push(r)); C.held += g.length; heldHours += maxClaim; return;
        }
      }
      if (payIt) {
        // the claim that is paid: the supervisor-answered one, else the largest
        const main = g.slice().sort((a, b) => (b.supVerdict ? 1 : 0) - (a.supVerdict ? 1 : 0) || b.claimed - a.claimed)[0];
        pay.push({ r: main, hours: hours, why: why, key: k });
        T.claimedPaid += main.claimed;
        g.filter(r => r !== main).forEach(r => { reject.push({ r: r, why: 'Duplicate — this shift is paid once on claim ' + main.otId }); C.dup++; });
      }
    });

    const out = [];
    const pad = (x, n) => { x = String(x); return x.length >= n ? x.slice(0, n - 1) + ' ' : x + ' '.repeat(n - x.length); };
    out.push((commit ? 'DONE' : 'PREVIEW — nothing written yet') + ' — payroll OT close, shifts ' + from + ' → ' + to);
    out.push('');
    out.push('PAID  ' + otc_hm_(T.punches + T.sup) + ' on ' + pay.length + ' shift(s)');
    out.push('   on punches (presence − shift)  : ' + otc_hm_(T.punches) + '  on ' + C.punches + ' shift(s)' +
             '   [' + C.abovePunches + ' claimed more than punches — ' + otc_hm_(claimedAbove) + ' not paid]');
    out.push('   ' + C.belowPunches + ' claimed LESS than punches show — ' +
             (OTC_CLOSE_PAY_ABOVE_CLAIM ? 'punches paid (' + otc_hm_(notAsked) + ' more than asked)'
                                        : 'claim paid; ' + otc_hm_(notAsked) + ' in punches not paid (listed below)'));
    out.push('   on supervisor (no punches)     : ' + otc_hm_(T.sup) + '  on ' + C.sup + ' shift(s)');
    out.push('REJECTED ' + reject.length + ' claim(s)');
    out.push('   punches show no OT             : ' + C.zero);
    out.push('   supervisor said not worked     : ' + C.supNo);
    out.push('   duplicate of a paid shift      : ' + C.dup);
    out.push('CLOSED   ' + settle.length + ' claim(s) — OT_Entries already has the day');
    out.push('HELD     ' + held.length + ' claim(s), ' + otc_hm_(heldHours) + ' claimed — nothing can measure them and no supervisor answered.');
    out.push('         They stay Pending and are not paid now. Answer them on the supervisor page and they are decided next run.');
    out.push('');
    out.push('HELD — to be answered (' + held.length + '):');
    held.forEach(r => out.push('   ' + pad(r.belongsTo, 12) + pad(r.empId + ' ' + r.name, 30) + pad(r.claimedHm, 7) + r.flag));
    out.push('');
    out.push('CLAIMED LESS THAN THE PUNCHES (' + below.length + ') — ' +
             (OTC_CLOSE_PAY_ABOVE_CLAIM ? 'punches paid' : 'claim paid') + ':');
    below.sort((a, b) => (b.punchOt - b.claim) - (a.punchOt - a.claim)).forEach(b =>
      out.push('   ' + pad(b.r.belongsTo, 12) + pad(b.r.empId + ' ' + b.r.name, 30) +
               'claimed ' + pad(otc_hm_(b.claim), 8) + 'punches ' + pad(otc_hm_(b.punchOt), 8) + '(' + b.r.checkIn + ' → ' + b.r.checkOut + ')'));
    out.push('');
    out.push('PAID (' + pay.length + '):');
    pay.forEach(p => out.push('   ' + pad(p.r.belongsTo, 12) + pad(p.r.empId + ' ' + p.r.name, 30) + pad(otc_hm_(p.hours), 8) + p.why));
    out.push('');
    out.push('REJECTED (' + reject.length + '):');
    reject.forEach(x => out.push('   ' + pad(x.r.belongsTo, 12) + pad(x.r.empId + ' ' + x.r.name, 30) + pad(x.r.claimedHm, 7) + x.why));

    if (!commit) {
      out.push('');
      out.push('Nothing has been changed. If this is right, run commitPayrollOtClose().');
      return otc_log_(out);
    }

    const pen = ss.getSheetByName(OTC_PENDING_TAB);
    let ot = ss.getSheetByName(OTC_ENTRIES_TAB);
    if (!ot) { ot = ss.insertSheet(OTC_ENTRIES_TAB); ot.appendRow(['Date','Employee ID','Name','OT Hours','Remark','Added By']); }
    const ec = otc_entryCols_(ot);
    const now = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm:ss');
    const newRows = pay.map(p => {
      const o = new Array(Math.max(ec.width, 6)).fill('');
      o[ec.date] = p.r.belongsTo; o[ec.empId] = p.r.empId; o[ec.name] = p.r.name;
      o[ec.hours] = Math.round(p.hours * 100) / 100;
      if (ec.remark >= 0) o[ec.remark] = 'Payroll close ' + now + ' — ' + p.why;
      if (ec.by >= 0) o[ec.by] = 'Payroll close (' + who + ')';
      if (ec.source >= 0) o[ec.source] = 'Kiosk claim (payroll close)';
      return o;
    });
    if (newRows.length) {
      const start = ot.getLastRow() + 1;
      ot.getRange(start, 1, newRows.length, newRows[0].length).setValues(newRows);
      ot.getRange(start, ec.date + 1, newRows.length, 1).setNumberFormat('@');
    }
    const setStatus = (r, s) => pen.getRange(r.row, OTC_COL.status + 1).setValue(s);
    pay.forEach(p => { setStatus(p.r, Math.abs(p.hours - p.r.claimed) < 0.02 ? 'Approved' : 'Approved (adjusted)'); });
    reject.forEach(x => setStatus(x.r, 'Rejected'));
    settle.forEach(x => setStatus(x.r, OTC_SETTLED));
    if (typeof logAmend_ === 'function') {
      pay.forEach(p => logAmend_(ss, 'Payroll close (' + who + ')', 'OT APPROVE', p.r.empId, p.r.belongsTo, 'OT',
                                 'claimed ' + p.r.claimedHm + ' → paid ' + otc_hm_(p.hours) + ' — ' + p.why));
      reject.forEach(x => logAmend_(ss, 'Payroll close (' + who + ')', 'OT REJECT', x.r.empId, x.r.belongsTo, 'OT',
                                    'claimed ' + x.r.claimedHm + ' — ' + x.why));
      settle.forEach(x => logAmend_(ss, 'Payroll close (' + who + ')', 'OT SETTLED', x.r.empId, x.r.belongsTo, 'OT',
                                    'claimed ' + x.r.claimedHm + ' — ' + x.why));
    }
    SpreadsheetApp.flush();
    out.push('');
    out.push('Written: ' + newRows.length + ' row(s) to OT_Entries, ' + (pay.length + reject.length + settle.length) +
             ' claim(s) decided, ' + held.length + ' held.');
    return otc_log_(out);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}