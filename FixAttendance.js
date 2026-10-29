/**
 * FixAttendance.gs — THE attendance fix for 21-09-2026 → 20-10-2026.
 *
 * THIS IS THE ONLY FILE. Delete Fix21Sep.gs, FixAnswered.gs and
 * FixCheckIns.gs if you still have them — everything they did is in here.
 * When anything changes, this file is replaced. The name never changes.
 *
 * It is IDEMPOTENT. Re-run it as often as you like: anything already correct
 * prints "already set" and is left alone. Nothing is written twice.
 *
 * WHAT IT DOES, in order
 *   1. Removes one mis-staged punch (Kiran 22-09, a logout on the wrong button).
 *   2. Writes 3 check-ins for days that had none, or had a stray tap as the start.
 *   3. Closes 30 days at the man's own rostered shift end, or at the time you
 *      confirmed. The start time is the one he really punched.
 *   4. Records Puran Kumar Thapa's day off on 23-09.
 *   5. Corrects Sudeep Limbu's sick leave to 07-10 → 08-10 via
 *      amendLeaveGrant(), which also moves his Leave_Balances figure.
 *   6. ALWAYS rebuilds Attendance_Daily for the whole period on commit —
 *      including the seven second-Check-Out rows that were appended
 *      straight into Attendance_Log, which only land through a rebuild.
 *   7. Prints what it did NOT touch, and why.
 *
 * SAFETY
 *   • Everything goes through amendAttendance(), so it lands in
 *     Attendance_Amendments under your name.
 *   • It REFUSES to overwrite a punch already there with a different time,
 *     unless that row is explicitly marked. Three rows are marked, and the
 *     log says so when it replaces one.
 *   • previewFixAttendance() writes nothing at all.
 *
 * HOW TO RUN   (no deployment — this is not the web app)
 *   1. previewFixAttendance  → Run → read the log.
 *   2. commitFixAttendance   → Run.
 */

const FX_BY   = 'Shakeel Mohammed Viam (MT-00060)';
const FX_FROM = '21-09-2026';
const FX_TO   = '20-10-2026';

// Mis-staged punches to remove. [id, name, date, stage, why]
const FX_DELETE = [
  ['MT-00063','Kiran Bhujel','22-09-2026','Break Out',
   'the 17:07 punch is his logout, recorded on the wrong button; re-written as the Check Out below'],
  ['MT-00061','Prakash Budha','04-10-2026','Day Off',
   'Bibek confirmed he worked 12:30 to 23:30 that day, so the Day Off tap at 17:29 is wrong']
];

// Check-ins to write. [id, name, date, time, mayOverwrite, why]
const FX_IN = [
  ['MT-00002','Ajay Gurung',      '23-09-2026','12:30', false,
   'the day had no punches at all'],
  ['MT-00042','Puran Kumar Thapa','25-09-2026','18:30', true,
   'replaces the 03:47 stray tap that landed 1 min after his 24th closed'],
  ['MT-00028','Keshab Alemagar',  '02-10-2026','18:54', true,
   'you confirmed 18:54 — replaces the 12:35 that made the shift ambiguous'],
  ['MT-00059','Yogendra Oli',     '22-09-2026','12:39', false,
   'you confirmed 12:39 is the start. The 20:21 Check In was a mis-tap and was '
   + 'deleted from Attendance_Log by hand on 20-10 — amendAttendance can only '
   + 'reach the FIRST matching stage on a date, which was the 12:39 we keep'],
  ['MT-00042','Puran Kumar Thapa','19-10-2026','12:30', true,
   'you confirmed 12:30 - replaces the 18:41 that was on file'],
  ['MT-00009','Bibek Nepali',     '22-09-2026','18:30', false,
   'Bibek confirmed 18:30 to 03:30. The day had no punches at all and no roster row'],
  ['MT-00024','Jit Bahadur Thakali','07-10-2026','12:30', false,
   'Bibek confirmed 12:30 to 03:30. The day had no punches at all and no roster row'],
  ['MT-00061','Prakash Budha',    '04-10-2026','12:30', false,
   'Bibek confirmed 12:30 to 23:30. He was working out of Jabriya; the Khairan '
   + 'transfer took effect 07-10'],
  ['MT-00012','Pramesh Pyakurel', '24-09-2026','12:30', false,
   'he worked; closed at the Qurain First Shift times, OT from the claim'],
  ['MT-00039','Nilam Tamang',     '14-10-2026','17:30', false,
   'he worked; closed at the Kaifan Second Shift times. His approved vacation '
   + 'does not start until 30-10, so he is working up to then']
];

// Days to close. [id, name, date, shift, checkIn, closeDate, closeTime, hours, why]
const FX_CLOSE = [
  ['MT-00001','Aaitu Newar','21-09-2026','Ardiya First Shift','12:41','21-09-2026','21:30',8.82,  // Ardiya
   'rostered shift end'],
  ['MT-00002','Ajay Gurung','21-09-2026','Jabriya Second Shift','18:49','22-09-2026','03:30',8.68,  // Jabriya
   'rostered shift end'],
  ['MT-00010','Bikash Bikram Shah','21-09-2026','Ardiya First Shift','12:41','21-09-2026','21:30',8.82,  // Ardiya
   'rostered shift end'],
  ['MT-00015','Dhirendra Khatri','21-09-2026','Kaifan Second Shift','17:35','22-09-2026','02:30',8.92,  // Kaifan
   'rostered shift end'],
  ['MT-00018','Foudman Limbu','21-09-2026','Ardiya First Shift','12:41','21-09-2026','21:30',8.82,  // Ardiya
   'rostered shift end'],
  ['MT-00024','Jit Bahadur Thakali','21-09-2026','Jabriya Second Shift','18:43','22-09-2026','03:30',8.78,  // Jabriya
   'rostered shift end'],
  ['MT-00033','Manish Gurung','21-09-2026','Jabriya First Shift','12:27','21-09-2026','21:30',9.05,  // Jabriya
   'rostered shift end'],
  ['MT-00040','Nishan Rai','21-09-2026','Kaifan First Shift','11:30','21-09-2026','20:30',9.0,  // Kaifan
   'rostered shift end'],
  ['MT-00061','Prakash Budha','21-09-2026','Jabriya Second Shift','18:44','22-09-2026','03:30',8.77,  // Jabriya
   'rostered shift end'],
  ['MT-00041','Pujan Bhattarai','21-09-2026','Khairan First Shift','13:52','21-09-2026','23:30',9.63,  // Khairan
   'you confirmed: 13:52 is kept, and that is his Khairan First start'],
  ['MT-00043','Rajan Shrestha','21-09-2026','Ardiya Second Shift','19:04','22-09-2026','03:30',8.43,  // Ardiya
   'rostered shift end'],
  ['MT-00044','Ram Bahadur Sirmal','21-09-2026','Kaifan Second Shift','17:35','22-09-2026','02:30',8.92,  // Kaifan
   'rostered shift end'],
  ['MT-00045','Ram Bahadur Thing','21-09-2026','Khairan First Shift','13:53','21-09-2026','23:30',9.62,  // Khairan
   'no roster row that day; his own pattern is the 13:5x Khairan First start'],
  ['MT-00050','Subas Pun Magar','21-09-2026','Ardiya Second Shift','19:04','22-09-2026','03:30',8.43,  // Ardiya
   'rostered shift end'],
  ['MT-00051','Sudeep Limbu','21-09-2026','Ardiya First Shift','12:41','21-09-2026','21:30',8.82,  // Ardiya
   'rostered shift end'],
  ['MT-00052','Suman Bishwakarma','21-09-2026','Jabriya First Shift','12:27','21-09-2026','21:30',9.05,  // Jabriya
   'rostered shift end'],
  ['MT-00059','Yogendra Oli','21-09-2026','Jabriya Second Shift','18:49','22-09-2026','03:30',8.68,  // Jabriya
   'rostered shift end'],
  ['MT-00063','Kiran Bhujel','22-09-2026','Store Hawally Second Shift','07:52','22-09-2026','17:07',9.25,  // Central Kitchen Hawally
   'you confirmed: he logged out at 17:07 — that punch exists, mis-staged as a Break Out'],
  ['MT-00002','Ajay Gurung','23-09-2026','Jabriya First Shift','12:30','24-09-2026','03:40',15.17,  // Jabriya
   'you confirmed: 12:30 in, 03:40 out next morning'],
  ['MT-00024','Jit Bahadur Thakali','23-09-2026','Jabriya Second Shift','18:42','24-09-2026','03:30',8.8,  // Jabriya
   'rostered shift end'],
  ['MT-00042','Puran Kumar Thapa','25-09-2026','Jabriya Second Shift','18:30','26-09-2026','03:30',9.0,  // Jabriya
   'the 03:47 tap is a stray 1 min after the 24th closed; breaks 21:31/22:34 put him on Second'],
  ['MT-00040','Nishan Rai','27-09-2026','Kaifan Second Shift','17:49','28-09-2026','02:30',8.68,  // Kaifan
   'rostered shift end'],
  ['MT-00036','Mohamed Rizwan Mohamed Hussain','05-10-2026','Central Kitchen Hawally','05:19','05-10-2026','14:00',8.68,  // Central Kitchen Hawally
   'rostered shift end'],
  ['MT-00051','Sudeep Limbu','05-10-2026','Ardiya Second Shift','19:02','06-10-2026','03:30',8.47,  // Ardiya
   'rostered shift end'],
  ['MT-00063','Kiran Bhujel','15-10-2026','Store Hawally Second Shift','08:53','15-10-2026','18:00',9.12,  // Central Kitchen Hawally
   'rostered shift end'],
  ['MT-00028','Keshab Alemagar','16-10-2026','Ardiya Second Shift','18:48','17-10-2026','03:30',8.7,  // Ardiya
   'rostered shift end'],
  ['MT-00016','Dipak Lo','17-10-2026','Qurain First Shift','12:38','17-10-2026','21:30',8.87,  // Qurain
   'rostered shift end'],
  ['MT-00009','Bibek Nepali','18-10-2026','Jabriya Second Shift','19:01','19-10-2026','03:30',8.48,  // Jabriya
   'rostered shift end'],
  ['MT-00028','Keshab Alemagar','02-10-2026','Ardiya Second Shift','18:54','03-10-2026','03:30',8.6,  // Ardiya
   'you confirmed he started 18:54 — that makes it Second, not First'],
  ['MT-00052','Suman Bishwakarma','14-10-2026','Kaifan (photo shoot)','12:52','14-10-2026','23:00',10.13,  // Mahboula
   'you confirmed he finished 23:00 after the photo-shoot duty at Kaifan'],
  ['MT-00059','Yogendra Oli','22-09-2026','Jabriya First Shift','12:39','23-09-2026','03:30',14.85,  // Jabriya
   'you confirmed 12:39 in, 03:30 out. 23-09 holds no Check Out of his, so this '
   + 'one appends cleanly; his 23-09 session (18:30) closes separately at 03:40 on the 24th'],
  ['MT-00009','Bibek Nepali','22-09-2026','Jabriya Second Shift','18:30','23-09-2026','03:30',9,  // Jabriya
   'Bibek confirmed. 23-09 holds no punch of his, so this appends cleanly'],
  ['MT-00024','Jit Bahadur Thakali','07-10-2026','Jabriya 12:30-03:30','12:30','08-10-2026','03:30',15,  // Jabriya
   'Bibek confirmed a 15-hour day. 08-10 has a Check In at 21:26 but no Check Out, '
   + 'and 03:30 falls before it, so this closes the 07-10 session and leaves the 08-10 one alone'],
];

// Days off. [id, name, date, why]
const FX_DAYOFF = [
  ['MT-00042','Puran Kumar Thapa','23-09-2026','you confirmed: he was off'],
  ['MT-00063','Kiran Bhujel','05-10-2026',
   'confirmed with Laxman Bahadur Airi, Store Hawally supervisor. His roster row '
   + 'already says Day Off but is still DRAFT, so the compile cannot see it'],
  ['MT-00063','Kiran Bhujel','13-10-2026',
   'confirmed with Laxman Bahadur Airi, Store Hawally supervisor. His roster row '
   + 'said Working and was corrected to Day Off by hand on 20-10'],
  ['MT-00063','Kiran Bhujel','19-10-2026',
   'confirmed with Laxman Bahadur Airi, Store Hawally supervisor. His roster row '
   + 'already says Day Off but is still DRAFT'],
  ['MT-00009','Bibek Nepali','23-09-2026','Bibek confirmed: off. No roster row exists'],
  ['MT-00009','Bibek Nepali','19-10-2026','Bibek confirmed: off. No roster row exists'],
  ['MT-00059','Yogendra Oli','03-10-2026','Bibek confirmed: off. No roster row exists'],
  ['MT-00059','Yogendra Oli','10-10-2026','Bibek confirmed: off. No roster row exists'],
  ['MT-00030','Kristal Ghalan','11-10-2026','Bibek confirmed: off. No roster row exists'],
  ['MT-00030','Kristal Ghalan','18-10-2026','Bibek confirmed: off. No roster row exists']
];

// Leave grants to correct. [grantId, name, newType, newStart, newEnd, why]
// Goes through amendLeaveGrant(), which moves the Leave_Balances figure,
// appends to the grant's decision note and refreshes the affected days.
// EMPTY on purpose. Sudeep Limbu is already correct, done by hand on 20-10:
// LG-20261006-164652 (06-10, 1d Sick) was CANCELLED, and LG-20261020-143528
// now carries 07-10 → 08-10, 2d Sick, Approved. His Sick balance reads
// Taken 2 / Remaining 13. amendLeaveGrant refuses a cancelled grant, so
// leaving the old row in here would fail on every run for no reason.
const FX_LEAVE = [
];

function previewFixAttendance() { return fx_run_(false); }
function commitFixAttendance()  { return fx_run_(true);  }

function fx_run_(commit) {
  const out = [];
  out.push((commit ? 'COMMIT' : 'PREVIEW') + '  —  attendance fix, ' + FX_FROM + ' to ' + FX_TO);
  out.push('');
  const n = { ok:0, failed:0, blocked:0, already:0 };

  // ── 1. remove mis-staged punches ──────────────────────────────────
  out.push('REMOVING ' + FX_DELETE.length + ' mis-staged punch');
  FX_DELETE.forEach(function (r) {
    const line = r[0] + '  ' + fx_pad_(r[1], 22) + r[2] + '  ' + r[3];
    if (!commit) { out.push('  would delete  ' + line); return; }
    const res = amendAttendance({ empId:r[0], date:r[2], stage:r[3], action:'delete', admin:FX_BY });
    const msg = (res && res.error) || '';
    if (res && res.success) { n.ok++; out.push('  DELETED       ' + line); }
    else if (/no matching punch/i.test(msg)) { n.already++; out.push('  already gone  ' + line); }
    else { n.failed++; out.push('  FAILED        ' + line + '   ' + msg); }
  });
  out.push('');

  // ── 2. check-ins ──────────────────────────────────────────────────
  out.push('WRITING ' + FX_IN.length + ' check-ins');
  FX_IN.forEach(function (r) {
    fx_set_(out, n, commit, r[0], r[1], r[2], 'Check In', r[3], r[4],
            fx_pad_(r[1], 22) + r[2] + '  Check In  ' + r[3]);
  });
  out.push('');

  // ── 3. closes ─────────────────────────────────────────────────────
  out.push('CLOSING ' + FX_CLOSE.length + ' days');
  out.push('');
  FX_CLOSE.forEach(function (r) {
    const label = fx_pad_(r[1], 22) + r[2] + '  ' + fx_pad_(r[3], 22) +
                  r[4] + ' \u2192 ' + r[6] + (r[5] !== r[2] ? ' (' + r[5] + ')' : '') +
                  '  ' + r[7] + ' h';
    fx_set_(out, n, commit, r[0], r[1], r[5], 'Check Out', r[6], false, label);
  });
  out.push('');

  // ── 4. days off ───────────────────────────────────────────────────
  out.push('DAYS OFF');
  FX_DAYOFF.forEach(function (r) {
    fx_set_(out, n, commit, r[0], r[1], r[2], 'Day Off', '00:00', false,
            fx_pad_(r[1], 22) + r[2] + '   (' + r[3] + ')');
  });
  out.push('');

  // ── 4b. leave grants ──────────────────────────────────────────────
  out.push('LEAVE GRANTS');
  FX_LEAVE.forEach(function (r) {
    const gid = r[0], label = fx_pad_(r[1], 22) + r[2] + '  ' + r[3] + ' → ' + r[4];
    let cur = null;
    try {
      const gs = ensureLeaveGrantsTab_(SpreadsheetApp.openById(MAIN_SPREADSHEET_ID));
      const f  = llFindGrant_(gs, gid);
      if (f) cur = llGrantObj_(f.row, lgDecisionCols_(gs));
    } catch (err) { /* fall through to the amend, which reports properly */ }

    if (!cur) {
      n.failed++;
      out.push('  NOT FOUND     ' + label + '   grant ' + gid + ' is not in Leave_Grants');
      return;
    }
    // llGrantObj_ hands back Date objects, not dd-mm-yyyy strings. Comparing
    // them raw never matches, so the grant would be re-amended on every run
    // and the decision note would grow a line each time. Normalise both ends
    // the same way amendLeaveGrant does before deciding anything.
    var curStart = '', curEnd = '';
    try {
      curStart = formatDdMmYyyy_(parseFlexibleDate_(cur.start));
      curEnd   = formatDdMmYyyy_(parseFlexibleDate_(cur.end));
    } catch (err) { curStart = String(cur.start); curEnd = String(cur.end); }

    if (curStart === r[3] && curEnd === r[4] &&
        String(cur.type).toLowerCase() === String(r[2]).toLowerCase()) {
      n.already++;
      out.push('  already set   ' + label);
      return;
    }
    out.push('  (currently ' + cur.type + ' ' + curStart + ' → ' + curEnd +
             ', ' + cur.days + 'd)');
    if (!commit) { out.push('  would amend   ' + label); return; }

    const res = amendLeaveGrant({ grantId:gid, type:r[2], start:r[3], end:r[4],
                                  by:FX_BY, note:r[5] });
    if (res && res.success) {
      n.ok++;
      out.push('  AMENDED       ' + label +
               (res.newRemaining !== null && res.newRemaining !== undefined
                  ? '   balance now ' + res.newRemaining : ''));
    } else {
      n.failed++;
      out.push('  FAILED        ' + label + '   ' + ((res && res.error) || 'no response'));
    }
  });
  out.push('');

  const total = FX_DELETE.length + FX_IN.length + FX_CLOSE.length +
                FX_DAYOFF.length + FX_LEAVE.length;
  out.push(commit
    ? (n.ok + ' written, ' + n.already + ' already correct, ' + n.blocked + ' blocked, ' + n.failed + ' failed.')
    : (total - n.blocked - n.already) + ' would be written, ' + n.already +
      ' already correct, ' + n.blocked + ' blocked. Nothing has changed.');

  // ── 5. rebuild ────────────────────────────────────────────────────
  // ALWAYS rebuild on commit, even when this script wrote nothing. Rows are
  // also appended straight into Attendance_Log by hand (the second Check Out
  // on a date that already has one — amendAttendance cannot write those), and
  // those only reach Attendance_Daily through this rebuild.
  if (commit) {
    out.push('');
    out.push('Rebuilding Attendance_Daily, ' + FX_FROM + ' to ' + FX_TO + '…');
    out.push('(this also picks up any rows appended directly to Attendance_Log)');
    try {
      const rb = rebuildAttendanceDaily_(FX_FROM, FX_TO);
      out.push('Rebuilt. ' + (rb && rb.message ? rb.message : ''));
    } catch (err) {
      out.push('REBUILD FAILED: ' + err);
      out.push('The punches ARE written. Set RANGE_FROM/RANGE_TO in RecompileRange.gs');
      out.push('to ' + FX_FROM + ' / ' + FX_TO + ' and run recompileFullRange().');
    }
  }

  // ── 6. what is left ───────────────────────────────────────────────
  out.push('');
  out.push('─────────────────────────────────────────────────────────────');
  out.push('NOT TOUCHED — these need a person:');
  out.push('');
  out.push('1. The driver days are CLOSED. 31 of them - Shakil Ahmad 17 and');
  out.push('   Mehdi Hasan 14 - were closed at the check-in plus NINE HOURS,');
  out.push('   on your instruction of 21-10, with OT coming from the claims.');
  out.push('   Drivers still have no row in the Shifts tab, so there was no');
  out.push('   rostered end to use and nine hours is the standard duty day.');
  out.push('   The rows went straight into Attendance_Log with Auth=ADMIN.');
  out.push('   Every close was checked first for an intervening punch, so each');
  out.push('   one lands on its own session. Sayed Sh. Anowar needed only one,');
  out.push('   on 26-09, his last day before 96 days of leave.');
  out.push('');
  out.push('   STILL TRUE: give the drivers a shift in the Shifts tab. Until');
  out.push('   they have one this has to be done by hand every period.');
  out.push('');
  out.push('2. The eight second-Check-Out rows are ALL DONE, plus Puran 19-10.');
  out.push('   They were appended straight into Attendance_Log with Auth=ADMIN,');
  out.push('   because amendAttendance can only reach the FIRST matching stage');
  out.push('   on a date and those dates already held one. The rebuild above');
  out.push('   reads them. Yogendra 22-09 also needed his 20:21 mis-tap deleted');
  out.push('   by hand, for the same reason.');
  out.push('');
  out.push('3. Two more closes went in by hand on 20-10, for the same reason -');
  out.push('   the date already held a Check Out, so amendAttendance could only');
  out.push('   have overwritten it. They are NOT in FX_CLOSE, because fx_set_');
  out.push('   reads the FIRST Check Out on a date and would report them blocked');
  out.push('   for ever:');
  out.push('      Pramesh Pyakurel  24-09  -> 21:30 on 24-09  (24-09 holds 03:10)');
  out.push('      Nilam Tamang      14-10  -> 02:30 on 15-10  (15-10 holds 23:43)');
  out.push('      Sayed Sh. Anowar  26-09  -> 23:45 on 26-09  (26-09 holds 00:54)');
  out.push('   Pramesh and Nilam also have their check-ins in FX_IN above, which');
  out.push('   are idempotent. Sayed already had his (11:30), so only the close');
  out.push('   was needed - it was his last working day before 96 days of leave.');
  out.push('');
  out.push('4. 119 roster rows in this period are still DRAFT, so the compile');
  out.push('   cannot see them. Store Hawally has not been published since 04-10.');
  out.push('   The eight Jabriya days that read "No Roster" are now answered by');
  out.push('   Bibek and recorded above. But the DRAFT rows remain, and a day');
  out.push('   with no punch and no published roster is NOT deducted - so this');
  out.push('   will happen again next period. Publish the rosters.');
  out.push('');
  out.push('5. Day Off tapped but marked Absent, because the roster row said');
  out.push('   Working and the roster wins. All EIGHT are now fixed - the roster');
  out.push('   rows were corrected to Day Off after the supervisor confirmed:');
  out.push('      Keshab Alemagar : Suman Karki 02-10, 09-10, 16-10 (all Fridays),');
  out.push('                        Sudeep Limbu 13-10, Foudman Limbu 14-10');
  out.push('      Bibek Nepali    : Prakash Budha 30-09');
  out.push('   Suman Karki is off EVERY FRIDAY and the Ardiya roster had him');
  out.push('   working all three. Fix the pattern, not just these dates.');
  out.push('');
  out.push('6. The 20th is not finished. Run this again tomorrow so the last day');
  out.push('   of the period compiles in full before you freeze it.');

  const txt = out.join('\n');
  Logger.log(txt);
  return txt;
}

// One amendAttendance call, with the overwrite guard.
function fx_set_(out, n, commit, id, nm, date, stage, time, mayOver, label) {
  const have = fx_timeOf_(id, date, stage);
  if (have === time) { n.already++; out.push('  already set   ' + id + '  ' + label); return; }
  if (have && !mayOver) {
    n.blocked++;
    out.push('  BLOCKED       ' + id + '  ' + label);
    out.push('                already has ' + stage + ' at ' + have + ' on ' + date + '. Not overwriting.');
    return;
  }
  if (have && mayOver) out.push('  (replacing the existing ' + stage + ' at ' + have + ')');
  if (!commit) { out.push('  would set     ' + id + '  ' + label); return; }
  const res = amendAttendance({ empId:id, date:date, stage:stage, action:'set', time:time, admin:FX_BY });
  if (res && res.success) {
    n.ok++;
    out.push('  SET           ' + id + '  ' + label + (res.reRunNeeded ? '   [' + res.reRunNeeded + ' flagged]' : ''));
  } else {
    n.failed++;
    out.push('  FAILED        ' + id + '  ' + label + '   ' + ((res && res.error) || 'unknown'));
  }
}

function fx_pad_(s, n) { s = String(s); while (s.length < n) s += ' '; return s; }

var fx_logCache_ = null;
function fx_timeOf_(empId, dateStr, stage) {
  if (!fx_logCache_) fx_logCache_ = fx_buildLog_();
  return fx_logCache_[empId + '|' + dateStr + '|' + stage] || '';
}
function fx_buildLog_() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const log = ss.getSheetByName('Attendance_Log');
  const map = {};
  if (!log) return map;
  const data = log.getDataRange().getValues();
  const H = data[0], C = attnCols_(H);
  const norm = function (v) {
    if (v instanceof Date) return formatDdMmYyyy_(v);
    const m = String(v).trim().match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    return m ? (m[1].length === 1 ? '0' + m[1] : m[1]) + '-' +
               (m[2].length === 1 ? '0' + m[2] : m[2]) + '-' + m[3] : String(v).trim();
  };
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][C.id]).trim();
    const d  = norm(data[i][C.date]);
    const st = String(data[i][C.stage]).trim();
    if (!id || !d || !st) continue;
    const k = id + '|' + d + '|' + st;
    if (map[k]) continue;
    const m = String(data[i][C.ts]).match(/(\d{1,2}):(\d{2})/);
    map[k] = m ? ((m[1].length === 1 ? '0' + m[1] : m[1]) + ':' + m[2]) : '';
  }
  return map;
}