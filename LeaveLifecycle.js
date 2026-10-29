// ============================================================
// LeaveLifecycle.gs — Maya HR : request → approve → attendance → payroll
// ============================================================
// Builds on LeaveEncashment.gs (Leave_Types, Leave_Balances) and
// SharedUtils.gs date helpers. Do NOT redeclare those.
//
// ------------------------------------------------------------
// THE LIFECYCLE
//
//   grantLeave            → writes the request as PENDING APPROVAL and
//                           reserves the days against the balance.
//                           Changes NOTHING about attendance or pay.
//   approveLeaveGrant     → Approved. Only now does it become real:
//                           attendance for the past days it covers is
//                           rebuilt, and payroll will see it.
//   rejectLeaveGrant      → Rejected. Reserved days handed back.
//   cancelLeaveGrant      → Cancelled (an approved leave called off).
//                           Days handed back, attendance rebuilt.
//   applyLeaveStatusChanges (daily) → on an APPROVED grant's start date,
//                           flips HR Maya status to "On Leave"; after the
//                           end date, lists reactivations (never auto).
//
// WHY THE GATE EXISTS
// grantLeave used to write "Approved" straight away, so anyone on the
// dashboard could put a man on leave and move his pay with one click, with
// nothing recorded about who agreed to it. Sixty days of annual leave is a
// two-month payroll event; it deserves a second signature.
//
// ------------------------------------------------------------
// THE ONE RULE EVERYTHING ELSE DEPENDS ON
//
//   ONLY Status === 'Approved' IS REAL.
//
// Attendance (writeDayToAttendanceDaily_), the On Leave status flip and
// the payroll deduction all test for exactly that string. A pending,
// rejected or cancelled grant is invisible to every one of them. Keep it
// that way: a new status that some readers treat as real and others do
// not is how a man gets paid for leave he never took.
//
// ------------------------------------------------------------
// BALANCES ARE RESERVED AT REQUEST, NOT AT APPROVAL
// The days come off the balance the moment leave is requested, and go back
// if it is rejected or cancelled. Two reasons: two pending requests cannot
// spend the same days, and every grant already in the sheet has already
// been deducted — so turning those into pending requests needs no balance
// arithmetic at all, which is what makes the migration below safe.
//
// ------------------------------------------------------------
// LEAVE IS AN ATTENDANCE INPUT. APPROVING IT REBUILDS THOSE DAYS.
//
// Attendance_Daily is a conclusion drawn from Leave_Grants (among other
// things), drawn once at 4 AM. Approving leave for dates already compiled
// therefore changed nothing — the sheet went on saying "Absent" or
// "No Roster" for days the man was signed off for, and every reader
// downstream believed it.
//
// Seen live on 20-09-2026: sick leave for MT-00033 covering 31-08 → 16-09
// was granted on the 20th. Four of those days sat in his supervisor's
// queue as unexplained. MT-00043, MT-00038 and MT-00006 were the same.
//
// So approve / reject / cancel each rebuild the past dates the grant
// covers, through adTouchRange_ in AttendanceDailyTrigger.gs. That call is
// deliberately forgiving: never throws, ignores future dates (they have no
// rows yet), obeys a time budget. Anything it cannot finish is picked up by
// the nightly rolling rebuild, so leave is never silently lost.
//
// HAND EDITS TO Leave_Grants ARE NOT COVERED. Nothing can hook a person
// typing into the sheet. Recent ones self-heal overnight; for anything
// older, run recompileLeaveGrants(). auditLeaveGrants() finds the damage.
//
// ENDPOINTS: grantLeave, getLeaveGrants, getLeaveRegister,
//            approveLeaveGrant, rejectLeaveGrant, cancelLeaveGrant,
//            applyLeaveStatusChanges, getLeaveStatusReminders,
//            getRemainingVacationTally, auditLeaveGrants,
//            recompileLeaveGrants,
//            previewLeaveApprovalMigration, commitLeaveApprovalMigration
// ============================================================

const LEAVE_GRANTS_TAB = 'Leave_Grants';

// The only four statuses. Anything else in the sheet is a fault and is
// reported by auditLeaveGrants().
const LG_PENDING   = 'Pending Approval';
const LG_APPROVED  = 'Approved';
const LG_REJECTED  = 'Rejected';
const LG_CANCELLED = 'Cancelled';
const LG_STATUSES  = [LG_PENDING, LG_APPROVED, LG_REJECTED, LG_CANCELLED];

// Appended once, at the end, so the positional readers elsewhere keep working.
const LG_DECISION_HEADERS = ['Decided By', 'Decided On', 'Decision Note'];

function ensureLeaveGrantsTab_(ss) {
  let sheet = ss.getSheetByName(LEAVE_GRANTS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(LEAVE_GRANTS_TAB);
    sheet.appendRow(['Grant ID', 'Employee ID', 'Name', 'Type', 'Start', 'End', 'Days', 'Status', 'Granted By', 'Granted On']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Locate (creating if needed) the decision-trail columns. Appending at the
// end never disturbs the fixed 0..9 layout the rest of the project reads.
function lgDecisionCols_(sheet) {
  const read = () => {
    const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const at = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    return { by: at('Decided By'), on: at('Decided On'), note: at('Decision Note'), width: H.length };
  };
  let c = read();
  if (c.by === -1 || c.on === -1 || c.note === -1) {
    const missing = LG_DECISION_HEADERS.filter(h =>
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .every(x => String(x).trim().toLowerCase() !== h.toLowerCase()));
    if (missing.length) {
      sheet.getRange(1, sheet.getLastColumn() + 1, 1, missing.length)
           .setValues([missing]).setFontWeight('bold');
      SpreadsheetApp.flush();
    }
    c = read();
  }
  return c;
}

// Ensure Leave_Types has an "Excludes From Payroll" column; seed sensible
// defaults (Annual & Unpaid = Yes, others = No) when creating it.
//
// NOTE: LeaveTypes.gs seeds this same column with a REGEX (/annual|unpaid/i)
// while this function used an exact match, so "Unpaid Sick Leave" got a
// different default depending on which file created the column first.
// Both now use the same test. Check the column in Leave_Types if you are
// unsure which rule your sheet was built with.
function ensureExcludeColumn_(ss) {
  const types = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (!types) return;
  const H = types.getRange(1, 1, 1, types.getLastColumn()).getValues()[0];
  let idx = H.findIndex(h => String(h).trim().toLowerCase() === 'excludes from payroll');
  if (idx !== -1) return;
  const col = types.getLastColumn() + 1;
  types.getRange(1, col).setValue('Excludes From Payroll').setFontWeight('bold');
  const data = types.getDataRange().getValues();
  const cT = H.findIndex(h => String(h).trim().toLowerCase() === 'type');
  for (let i = 1; i < data.length; i++) {
    const t = String(data[i][cT] || '').trim();
    const def = /annual|unpaid/i.test(t) ? 'Yes' : 'No';
    types.getRange(i + 1, col).setValue(def);
  }
}

function excludesFromPayroll_(ss, type) {
  const types = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (!types) return /annual|unpaid/i.test(String(type));
  const data = types.getDataRange().getValues();
  const H = data[0];
  const cT = H.findIndex(h => String(h).trim().toLowerCase() === 'type');
  const cE = H.findIndex(h => String(h).trim().toLowerCase() === 'excludes from payroll');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cT] || '').trim().toLowerCase() === String(type).trim().toLowerCase()) {
      if (cE === -1) return /annual|unpaid/i.test(String(type));
      return String(data[i][cE]).trim().toLowerCase() === 'yes';
    }
  }
  return false;
}

// ------------------------------------------------------------
// Shared plumbing
// ------------------------------------------------------------

// Rebuild the attendance days a grant covers. Past dates only — a future
// date has no Attendance_Daily row yet and gets built when it arrives.
// Never throws: a refresh must not be able to break the decision that
// caused it, and the nightly rolling rebuild is the safety net.
function llRefreshAttendance_(startDate, endDate, label) {
  try {
    if (typeof adTouchRange_ !== 'function') {
      Logger.log('llRefreshAttendance_: adTouchRange_ not found — is AttendanceDailyTrigger.gs v5 deployed?');
      return { success: false, error: 'rebuild unavailable' };
    }
    const from = formatDdMmYyyy_(startDate), to = formatDdMmYyyy_(endDate);
    const res = adTouchRange_(from, to);
    Logger.log('llRefreshAttendance_ (' + (label || 'leave') + ') ' + from + ' → ' + to +
               ': ' + (res && res.message ? res.message : JSON.stringify(res)));
    return res;
  } catch (err) {
    Logger.log('llRefreshAttendance_ failed (non-fatal): ' + err);
    return { success: false, error: err.toString() };
  }
}

// One short sentence about what the rebuild did, so a decision on past
// dates visibly fixes the attendance sheet instead of quietly leaving it wrong.
function llRefreshNote_(res) {
  if (!res) return '';
  if (!res.success) return ' Attendance for those days could not be refreshed automatically — it will correct itself on tonight\'s run.';
  const n = (res.dates || []).length;
  if (!n) return '';
  let s = ' Attendance rebuilt for ' + n + ' past day(s).';
  if (res.skipped && res.skipped.length) s += ' ' + res.skipped.length + ' more will refresh on tonight\'s run.';
  return s;
}

// Move days on or off a balance. delta > 0 gives days BACK.
// Only for balance-tracked types; Unpaid and friends have entitlement 0.
function llMoveBalance_(ss, empId, type, delta) {
  try {
    const types = getLeaveTypesMap_(ss);
    const meta = types[String(type).toLowerCase()];
    if (!meta || (meta.entitlement || 0) <= 0 || !delta) return null;
    const balSheet = ss.getSheetByName(LEAVE_BALANCES_TAB);
    const rowNum = findBalanceRow_(balSheet, empId, type);
    if (rowNum === -1) return null;
    const b = balSheet.getRange(rowNum, 1, 1, 7).getValues()[0];
    const taken = Math.max(0, (Number(b[4]) || 0) - delta);   // delta back → taken down
    const remaining = (Number(b[6]) || 0) + delta;
    balSheet.getRange(rowNum, 5).setValue(taken);
    balSheet.getRange(rowNum, 7).setValue(remaining);
    return remaining;
  } catch (e) {
    Logger.log('llMoveBalance_ failed: ' + e);
    return null;
  }
}

// Find a grant row by ID. Returns { rowNum, row } or null.
function llFindGrant_(sheet, grantId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === String(grantId).trim()) return { rowNum: i + 1, row: data[i] };
  }
  return null;
}

function llGrantObj_(row, dc) {
  return {
    grantId: String(row[0] || '').trim(),
    empId:   String(row[1] || '').trim(),
    name:    String(row[2] || '').trim(),
    type:    String(row[3] || '').trim(),
    start:   row[4], end: row[5],
    days:    Number(row[6]) || 0,
    status:  String(row[7] || '').trim(),
    grantedBy: String(row[8] || '').trim(),
    grantedOn: row[9],
    decidedBy:   dc && dc.by   >= 0 ? String(row[dc.by]   || '').trim() : '',
    decidedOn:   dc && dc.on   >= 0 ? String(row[dc.on]   || '').trim() : '',
    decisionNote:dc && dc.note >= 0 ? String(row[dc.note] || '').trim() : '',
  };
}

function llStamp_(sheet, rowNum, dc, by, note) {
  const now = formatDdMmYyyy_(new Date());
  if (dc.by   >= 0) sheet.getRange(rowNum, dc.by + 1).setValue(String(by || 'Dashboard').trim());
  if (dc.on   >= 0) sheet.getRange(rowNum, dc.on + 1).setValue(now);
  if (dc.note >= 0 && note) sheet.getRange(rowNum, dc.note + 1).setValue(String(note).trim());
}

// ------------------------------------------------------------
// REQUEST LEAVE — records the request as PENDING APPROVAL and reserves
// the days. Nothing about attendance or pay changes until it is approved.
// params: { empId, type, start (dd-mm-yyyy), end (dd-mm-yyyy), days?, by? }
// ------------------------------------------------------------
function grantLeave(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    ensureLeaveTabs(); ensureExcludeColumn_(ss);
    const grants = ensureLeaveGrantsTab_(ss);
    lgDecisionCols_(grants);

    const empId = String(params.empId || '').trim();
    const type = String(params.type || '').trim();
    if (!empId || !type) return { success: false, error: 'empId and type are required' };
    const start = parseFlexibleDate_(params.start);
    const end = parseFlexibleDate_(params.end);
    if (!start || !end) return { success: false, error: 'start and end must be valid dates (dd-mm-yyyy)' };
    if (endOfDay_(end) < startOfDay_(start)) return { success: false, error: 'End date is before start date.' };

    const name = empNameById_(ss, empId);
    if (!name) return { success: false, error: 'Employee ID "' + empId + '" not found in HR Maya' };

    const types = getLeaveTypesMap_(ss);
    if (!types[type.toLowerCase()]) return { success: false, error: 'Unknown leave type "' + type + '"' };

    // Inclusive calendar days unless overridden.
    let days = Number(params.days) || 0;
    if (days <= 0) days = Math.round((startOfDay_(end) - startOfDay_(start)) / 86400000) + 1;

    // OVERLAP GUARD. Two live grants covering the same day means attendance
    // takes whichever sits lower in the sheet and payroll may deduct both.
    const clash = llOverlapping_(grants, empId, start, end, null);
    if (clash.length && !params.force) {
      return { success: false, overlaps: clash,
        error: name + ' already has ' + (clash.length === 1 ? 'a leave' : clash.length + ' leaves') +
               ' covering those dates: ' +
               clash.map(c => c.type + ' ' + c.start + '→' + c.end + ' (' + c.status + ')').join(', ') +
               '. Cancel or reject it first.' };
    }

    // Reserve the days now, so two pending requests cannot spend the same
    // balance. Handed straight back if this is rejected or cancelled.
    const entitlement = types[type.toLowerCase()].entitlement || 0;
    let newRemaining = null;
    if (entitlement > 0) {
      const balSheet = ss.getSheetByName(LEAVE_BALANCES_TAB);
      const rowNum = findBalanceRow_(balSheet, empId, type);
      if (rowNum === -1) return { success: false, error: 'No ' + type + ' balance for ' + name + '. Set a balance first.' };
      const b = balSheet.getRange(rowNum, 1, 1, 7).getValues()[0];
      const remaining = Number(b[6]) || 0;
      // Allow going negative (no block). Balance can drop below zero.
      newRemaining = remaining - days;
      balSheet.getRange(rowNum, 5).setValue((Number(b[4]) || 0) + days); // Taken
      balSheet.getRange(rowNum, 7).setValue(newRemaining);               // Remaining (may be negative)
    }

    const grantId = 'LG-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');
    grants.appendRow([grantId, empId, name, types[type.toLowerCase()].type,
      formatDdMmYyyy_(start), formatDdMmYyyy_(end), days, LG_PENDING,
      String(params.by || 'Dashboard').trim(), formatDdMmYyyy_(new Date())]);

    const willFlip = excludesFromPayroll_(ss, type);
    const negativeWarning = (newRemaining !== null && newRemaining < 0)
      ? ' ⚠ ' + name + ' would be at ' + newRemaining + ' ' + type + ' (negative balance).' : '';
    return {
      success: true, grantId: grantId, empId: empId, name: name, type: types[type.toLowerCase()].type,
      start: formatDdMmYyyy_(start), end: formatDdMmYyyy_(end), days: days,
      status: LG_PENDING, excludesFromPayroll: willFlip,
      newRemaining: newRemaining, negative: (newRemaining !== null && newRemaining < 0),
      message: 'Requested ' + days + ' day(s) of ' + type + ' for ' + name + ' (' +
        formatDdMmYyyy_(start) + ' to ' + formatDdMmYyyy_(end) + ') — ' +
        'PENDING APPROVAL.' + negativeWarning +
        ' Days are held against the balance. Nothing changes on attendance or payroll until it is approved.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Live grants (pending or approved) for this employee that overlap a range.
function llOverlapping_(grantsSheet, empId, start, end, ignoreGrantId) {
  const out = [];
  const data = grantsSheet.getDataRange().getValues();
  const s = startOfDay_(start).getTime(), e = startOfDay_(end).getTime();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][1] || '').trim();
    if (id !== empId) continue;
    const gid = String(data[i][0] || '').trim();
    if (ignoreGrantId && gid === ignoreGrantId) continue;
    const st = String(data[i][7] || '').trim().toLowerCase();
    if (st !== LG_PENDING.toLowerCase() && st !== LG_APPROVED.toLowerCase()) continue;
    const gs = parseFlexibleDate_(data[i][4]), ge = parseFlexibleDate_(data[i][5]);
    if (!gs || !ge) continue;
    if (startOfDay_(gs).getTime() <= e && startOfDay_(ge).getTime() >= s) {
      out.push({ grantId: gid, type: String(data[i][3] || '').trim(),
                 start: formatDdMmYyyy_(gs), end: formatDdMmYyyy_(ge),
                 status: String(data[i][7] || '').trim() });
    }
  }
  return out;
}

// ------------------------------------------------------------
// DECIDE — approve / reject / cancel.
// Accepts grantId or grantIds:[...] so a screenful can be actioned at once.
// ------------------------------------------------------------
function approveLeaveGrant(params) { return llDecide_(params, 'approve'); }
function rejectLeaveGrant(params)  { return llDecide_(params, 'reject'); }
function cancelLeaveGrant(params)  { return llDecide_(params, 'cancel'); }

function llDecide_(params, action) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grants = ensureLeaveGrantsTab_(ss);
    const dc = lgDecisionCols_(grants);
    const by = String(params.by || 'Dashboard').trim();
    const note = String(params.note || params.reason || '').trim();

    let ids = params.grantIds;
    if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) { ids = [ids]; } }
    if (!ids || !ids.length) ids = params.grantId ? [params.grantId] : [];
    if (!ids.length) return { success: false, error: 'grantId (or grantIds) is required' };

    const done = [], skipped = [];
    let minStart = null, maxEnd = null, restored = 0;

    ids.forEach(gid => {
      const found = llFindGrant_(grants, gid);
      if (!found) { skipped.push(gid + ': not found'); return; }
      const g = llGrantObj_(found.row, dc);
      const st = g.status.toLowerCase();

      // What each action is allowed to act on. Being strict here is what
      // stops a balance being handed back twice.
      if (action === 'approve' && st !== LG_PENDING.toLowerCase()) {
        skipped.push(g.name + ' ' + gid + ': is "' + g.status + '", only a pending request can be approved'); return;
      }
      if (action === 'reject' && st !== LG_PENDING.toLowerCase()) {
        skipped.push(g.name + ' ' + gid + ': is "' + g.status + '", only a pending request can be rejected'); return;
      }
      if (action === 'cancel' && st !== LG_APPROVED.toLowerCase() && st !== LG_PENDING.toLowerCase()) {
        skipped.push(g.name + ' ' + gid + ': is already "' + g.status + '"'); return;
      }

      const newStatus = action === 'approve' ? LG_APPROVED
                      : action === 'reject'  ? LG_REJECTED
                      : LG_CANCELLED;
      grants.getRange(found.rowNum, 8).setValue(newStatus);
      llStamp_(grants, found.rowNum, dc, by, note);

      // Reserved days go back on anything that is not an approval.
      let newRemaining = null;
      if (action !== 'approve') {
        newRemaining = llMoveBalance_(ss, g.empId, g.type, g.days);
        restored += g.days;
      }

      const s = parseFlexibleDate_(g.start), e = parseFlexibleDate_(g.end);
      if (s && e) {
        if (!minStart || s < minStart) minStart = s;
        if (!maxEnd   || e > maxEnd)   maxEnd   = e;
      }
      done.push({ grantId: gid, empId: g.empId, name: g.name, type: g.type,
                  days: g.days, status: newStatus, newRemaining: newRemaining });
    });

    // One rebuild for everything decided, rather than one per grant.
    // Approving makes leave real; rejecting or cancelling makes it unreal.
    // Both change what those days should say, so both rebuild.
    let refresh = null;
    if (done.length && minStart && maxEnd) {
      refresh = llRefreshAttendance_(minStart, maxEnd, action + ' x' + done.length);
    }

    const verb = action === 'approve' ? 'Approved' : action === 'reject' ? 'Rejected' : 'Cancelled';
    let message = verb + ' ' + done.length + ' leave' + (done.length === 1 ? '' : 's');
    if (done.length === 1) message = verb + ' ' + done[0].days + ' day(s) of ' + done[0].type + ' for ' + done[0].name;
    message += '.';
    if (restored) message += ' ' + restored + ' day(s) returned to balance.';
    message += llRefreshNote_(refresh);
    if (skipped.length) message += ' ' + skipped.length + ' skipped.';

    return { success: true, action: action, decided: done, skipped: skipped,
             attendanceRebuilt: (refresh && refresh.dates) ? refresh.dates : [],
             message: message };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// AMEND — correct the dates (or type) of a leave that is already recorded,
// approved or not. A man's ticket moves, he comes back early, the wrong
// month was typed: all of it happens after approval, and re-cutting the
// grant loses the history and double-counts the balance.
//
// What it puts right, in order:
//   · the balance — only the DIFFERENCE is taken or returned, and a change
//     of type returns the old days to the old type before taking the new
//   · the overlap guard — the same one used when granting, ignoring itself
//   · attendance — rebuilt across BOTH the old and the new range, so days
//     the leave no longer covers stop saying "on leave"
//
// params: { grantId, start?, end?, type?, days?, by?, note?, force? }
// ------------------------------------------------------------
function amendLeaveGrant(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grants = ensureLeaveGrantsTab_(ss);
    const dc = lgDecisionCols_(grants);

    const gid = String(params.grantId || '').trim();
    if (!gid) return { success: false, error: 'grantId is required' };
    const found = llFindGrant_(grants, gid);
    if (!found) return { success: false, error: 'Leave ' + gid + ' not found.' };
    const g = llGrantObj_(found.row, dc);
    const st = String(g.status || '').toLowerCase();
    if (st === LG_REJECTED.toLowerCase() || st === LG_CANCELLED.toLowerCase()) {
      return { success: false, error: 'This leave is ' + g.status + '. Nothing to amend — grant a new one.' };
    }

    const oldStart = parseFlexibleDate_(g.start), oldEnd = parseFlexibleDate_(g.end);
    const start = params.start ? parseFlexibleDate_(params.start) : oldStart;
    const end   = params.end   ? parseFlexibleDate_(params.end)   : oldEnd;
    if (!start || !end) return { success: false, error: 'start and end must be valid dates (dd-mm-yyyy)' };
    if (endOfDay_(end) < startOfDay_(start)) return { success: false, error: 'End date is before start date.' };

    const types = getLeaveTypesMap_(ss);
    let newType = g.type;
    if (params.type && String(params.type).trim() && String(params.type).trim().toLowerCase() !== String(g.type).toLowerCase()) {
      const meta = types[String(params.type).trim().toLowerCase()];
      if (!meta) return { success: false, error: 'Unknown leave type "' + params.type + '"' };
      newType = meta.type;
    }

    let newDays = Number(params.days) || 0;
    if (newDays <= 0) newDays = Math.round((startOfDay_(end) - startOfDay_(start)) / 86400000) + 1;
    const oldDays = Number(g.days) || 0;

    // Same guard as granting — but this grant must not clash with itself.
    const clash = llOverlapping_(grants, g.empId, start, end, gid);
    if (clash.length && !params.force) {
      return { success: false, overlaps: clash,
        error: g.name + ' already has ' + (clash.length === 1 ? 'a leave' : clash.length + ' leaves') +
               ' covering those dates: ' +
               clash.map(c => c.type + ' ' + c.start + '→' + c.end + ' (' + c.status + ')').join(', ') + '.' };
    }

    // Balance. A change of type is a return to one bucket and a draw from
    // the other; otherwise only the difference in days moves.
    let newRemaining = null;
    if (newType.toLowerCase() !== String(g.type).toLowerCase()) {
      llMoveBalance_(ss, g.empId, g.type, oldDays);          // give the old days back
      newRemaining = llMoveBalance_(ss, g.empId, newType, -newDays);
    } else if (newDays !== oldDays) {
      newRemaining = llMoveBalance_(ss, g.empId, newType, oldDays - newDays);
    }

    grants.getRange(found.rowNum, 4).setValue(newType);
    grants.getRange(found.rowNum, 5).setValue(formatDdMmYyyy_(start));
    grants.getRange(found.rowNum, 6).setValue(formatDdMmYyyy_(end));
    grants.getRange(found.rowNum, 7).setValue(newDays);

    // Record the amendment without destroying who approved it.
    const who = String(params.by || 'Dashboard').trim();
    const stamp = 'Amended ' + formatDdMmYyyy_(new Date()) + ' by ' + who + ': ' +
                  g.start + '→' + g.end + ' (' + oldDays + 'd ' + g.type + ')  becomes  ' +
                  formatDdMmYyyy_(start) + '→' + formatDdMmYyyy_(end) + ' (' + newDays + 'd ' + newType + ')' +
                  (params.note ? '. ' + String(params.note).trim() : '');
    if (dc.note >= 0) {
      const prev = String(grants.getRange(found.rowNum, dc.note + 1).getValue() || '').trim();
      grants.getRange(found.rowNum, dc.note + 1).setValue(prev ? (prev + ' | ' + stamp) : stamp);
    }

    // Attendance across the whole affected span — the old days must stop
    // reading as leave, and the new ones must start.
    const spanStart = (oldStart && oldStart < start) ? oldStart : start;
    const spanEnd   = (oldEnd   && oldEnd   > end)   ? oldEnd   : end;
    let refresh = null;
    if (st === LG_APPROVED.toLowerCase()) {
      refresh = llRefreshAttendance_(spanStart, spanEnd, 'amend ' + gid);
    }

    return {
      success: true, grantId: gid, empId: g.empId, name: g.name,
      type: newType, status: g.status,
      start: formatDdMmYyyy_(start), end: formatDdMmYyyy_(end), days: newDays,
      previous: { type: g.type, start: g.start, end: g.end, days: oldDays },
      newRemaining: newRemaining,
      attendanceRebuilt: (refresh && refresh.dates) ? refresh.dates : [],
      message: g.name + ': ' + g.start + '→' + g.end + ' (' + oldDays + 'd) amended to ' +
               formatDdMmYyyy_(start) + '→' + formatDdMmYyyy_(end) + ' (' + newDays + 'd' +
               (newType.toLowerCase() !== String(g.type).toLowerCase() ? ', now ' + newType : '') + ').' +
               (newRemaining !== null ? ' Balance now ' + newRemaining + '.' : '') +
               llRefreshNote_(refresh)
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// READ — the register behind the Leave Register page.
// params: { status?, empId?, type?, from?, to?, includeDecided? }
// ------------------------------------------------------------
function getLeaveRegister(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grants = ensureLeaveGrantsTab_(ss);
    const dc = lgDecisionCols_(grants);
    const data = grants.getDataRange().getValues();

    const fStatus = String(params.status || '').trim().toLowerCase();
    const fEmp    = String(params.empId  || '').trim().toUpperCase();
    const fType   = String(params.type   || '').trim().toLowerCase();
    const from    = params.from ? startOfDay_(parseFlexibleDate_(params.from)) : null;
    const to      = params.to   ? startOfDay_(parseFlexibleDate_(params.to))   : null;

    const today = startOfDay_(new Date());
    const rows = [];
    const counts = { pending: 0, approved: 0, rejected: 0, cancelled: 0, other: 0 };

    for (let i = 1; i < data.length; i++) {
      if (!String(data[i][0] || '').trim() && !String(data[i][1] || '').trim()) continue;
      const g = llGrantObj_(data[i], dc);
      const s = parseFlexibleDate_(g.start), e = parseFlexibleDate_(g.end);
      const st = g.status.toLowerCase();

      if (st === LG_PENDING.toLowerCase()) counts.pending++;
      else if (st === LG_APPROVED.toLowerCase()) counts.approved++;
      else if (st === LG_REJECTED.toLowerCase()) counts.rejected++;
      else if (st === LG_CANCELLED.toLowerCase()) counts.cancelled++;
      else counts.other++;

      if (fStatus && st !== fStatus) continue;
      if (fEmp && g.empId.toUpperCase() !== fEmp) continue;
      if (fType && g.type.toLowerCase() !== fType) continue;
      if (from && e && startOfDay_(e) < from) continue;
      if (to   && s && startOfDay_(s) > to)   continue;

      const startsInFuture = s ? startOfDay_(s) > today : false;
      const inProgress = (s && e) ? (startOfDay_(s) <= today && startOfDay_(e) >= today) : false;

      rows.push({
        grantId: g.grantId, empId: g.empId, name: g.name, type: g.type,
        start: s ? formatDdMmYyyy_(s) : String(g.start || ''),
        end:   e ? formatDdMmYyyy_(e) : String(g.end || ''),
        days: g.days, status: g.status || '(blank)',
        grantedBy: g.grantedBy, grantedOn: g.grantedOn ? String(g.grantedOn) : '',
        decidedBy: g.decidedBy, decidedOn: g.decidedOn, decisionNote: g.decisionNote,
        isPaid: !llIsUnpaidType_(ss, g.type),
        excludesFromPayroll: excludesFromPayroll_(ss, g.type),
        startsInFuture: startsInFuture, inProgress: inProgress,
        canApprove: st === LG_PENDING.toLowerCase(),
        canReject:  st === LG_PENDING.toLowerCase(),
        canCancel:  st === LG_APPROVED.toLowerCase() || st === LG_PENDING.toLowerCase(),
      });
    }

    rows.sort((a, b) => {
      const iso = (d) => { const p = String(d).split('-'); return (p[2]||'') + (p[1]||'') + (p[0]||''); };
      return iso(b.start).localeCompare(iso(a.start));   // newest leave first
    });

    return { success: true, data: rows, counts: counts, total: rows.length,
             statuses: LG_STATUSES };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function llIsUnpaidType_(ss, type) {
  const types = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (!types) return /unpaid/i.test(String(type));
  const data = types.getDataRange().getValues();
  const H = data[0];
  const cT = H.findIndex(h => String(h).trim().toLowerCase() === 'type');
  const cP = H.findIndex(h => String(h).trim().toLowerCase() === 'is paid');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cT] || '').trim().toLowerCase() === String(type).trim().toLowerCase()) {
      if (cP === -1) return /unpaid/i.test(String(type));
      const v = String(data[i][cP]).trim().toLowerCase();
      return v === 'no' || v === 'false';
    }
  }
  return /unpaid/i.test(String(type));
}

// Kept for anything already calling it.
function getLeaveGrants(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLeaveGrantsTab_(ss);
    const data = sheet.getDataRange().getValues();
    const filterId = String(params.empId || '').trim().toUpperCase();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][1] || '').trim();
      if (!id) continue;
      if (filterId && id.toUpperCase() !== filterId) continue;
      out.push({ grantId: data[i][0], empId: id, name: data[i][2], type: data[i][3], start: data[i][4], end: data[i][5], days: data[i][6], status: data[i][7], grantedBy: data[i][8], grantedOn: data[i][9] });
    }
    return { success: true, data: out };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// AUDIT — the loopholes nothing else will catch.
// Read-only. Run it before a payroll, and after editing the sheet by hand.
// ------------------------------------------------------------
function auditLeaveGrants(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grants = ensureLeaveGrantsTab_(ss);
    const data = grants.getDataRange().getValues();
    const types = getLeaveTypesMap_(ss);
    const problems = [];
    const add = (sev, gid, name, what, why) => problems.push({ severity: sev, grantId: gid, name: name, problem: what, why: why });

    const live = [];   // pending or approved, for the overlap pass
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const gid = String(r[0] || '').trim();
      const empId = String(r[1] || '').trim();
      if (!gid && !empId) continue;
      const name = String(r[2] || '').trim() || empId;
      const type = String(r[3] || '').trim();
      const status = String(r[7] || '').trim();
      const s = parseFlexibleDate_(r[4]), e = parseFlexibleDate_(r[5]);
      const days = Number(r[6]) || 0;

      if (!empId) add('high', gid, name, 'No employee ID', 'This row can never be matched to anyone.');
      if (!status) {
        add('high', gid, name, 'Status is blank',
            'Attendance ignores it, but the payroll unpaid-leave deduction treats a blank status as approved — ' +
            'so it can be deducted without ever appearing on the attendance sheet.');
      } else if (LG_STATUSES.map(x => x.toLowerCase()).indexOf(status.toLowerCase()) === -1) {
        add('high', gid, name, 'Unknown status "' + status + '"',
            'Only ' + LG_STATUSES.join(' / ') + ' are understood. Anything else is invisible to attendance.');
      }
      if (!s || !e) add('high', gid, name, 'Unreadable dates', 'Start or end could not be parsed, so this leave affects nothing.');
      else if (startOfDay_(e) < startOfDay_(s)) add('high', gid, name, 'End is before start', 'Covers no days at all.');
      else {
        const calc = Math.round((startOfDay_(e) - startOfDay_(s)) / 86400000) + 1;
        if (days && Math.abs(calc - days) > 0.001)
          add('medium', gid, name, 'Days column says ' + days + ', dates span ' + calc,
              'The balance was moved by ' + days + ' but the leave covers ' + calc + ' days.');
      }
      if (type && !types[type.toLowerCase()])
        add('medium', gid, name, 'Unknown leave type "' + type + '"',
            'Not in Leave_Types, so nothing can tell whether it is paid.');

      const st = status.toLowerCase();
      if ((st === LG_PENDING.toLowerCase() || st === LG_APPROVED.toLowerCase()) && s && e) {
        live.push({ gid: gid, empId: empId, name: name, type: type, status: status,
                    s: startOfDay_(s).getTime(), e: startOfDay_(e).getTime() });
      }
    }

    // Overlaps: two live grants covering the same day for one person.
    for (let i = 0; i < live.length; i++) {
      for (let j = i + 1; j < live.length; j++) {
        if (live[i].empId !== live[j].empId) continue;
        if (live[i].s <= live[j].e && live[i].e >= live[j].s) {
          add('high', live[i].gid + ' / ' + live[j].gid, live[i].name,
              'Overlapping leave (' + live[i].type + ' and ' + live[j].type + ')',
              'Attendance shows whichever sits lower in the sheet; if both are unpaid, payroll may deduct the same day twice.');
        }
      }
    }

    const high = problems.filter(p => p.severity === 'high').length;
    return { success: true, problems: problems, total: problems.length, high: high,
             message: problems.length
               ? problems.length + ' problem(s) found, ' + high + ' serious.'
               : 'Leave_Grants is clean — no blank or unknown statuses, no overlaps, no bad dates.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Nothing can hook somebody typing directly into Leave_Grants. Run this
// after editing that sheet by hand to make attendance agree with it.
// params: { from?, to? }  — defaults to everything approved grants cover.
// ------------------------------------------------------------
function recompileLeaveGrants(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grants = ensureLeaveGrantsTab_(ss);
    const data = grants.getDataRange().getValues();

    const limitFrom = params.from ? startOfDay_(parseFlexibleDate_(params.from)) : null;
    const limitTo   = params.to   ? startOfDay_(parseFlexibleDate_(params.to))   : null;

    let min = null, max = null, covered = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][7] || '').trim().toLowerCase() !== LG_APPROVED.toLowerCase()) continue;
      const s = parseFlexibleDate_(data[i][4]), e = parseFlexibleDate_(data[i][5]);
      if (!s || !e) continue;
      if (limitTo   && startOfDay_(s) > limitTo)   continue;
      if (limitFrom && startOfDay_(e) < limitFrom) continue;
      covered++;
      if (!min || s < min) min = s;
      if (!max || e > max) max = e;
    }
    if (!min) return { success: true, message: 'No approved grants to rebuild.' };
    if (limitFrom && min < limitFrom) min = limitFrom;
    if (limitTo   && max > limitTo)   max = limitTo;

    const res = llRefreshAttendance_(min, max, 'all approved grants');
    return { success: true, grants: covered,
      from: formatDdMmYyyy_(min), to: formatDdMmYyyy_(max),
      rebuilt: (res && res.dates) ? res.dates : [],
      skipped: (res && res.skipped) ? res.skipped : [],
      message: covered + ' approved grant(s) span ' + formatDdMmYyyy_(min) + ' → ' +
               formatDdMmYyyy_(max) + '.' + llRefreshNote_(res) };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ============================================================
// ONE-OFF MIGRATION — existing grants that were never really approved.
//
// Before the approval gate existed, grantLeave wrote "Approved" straight
// away. Leave that has not started yet was therefore never actually
// agreed by anyone; it should be sitting in the pending queue.
//
// This is SAFE TO RUN because it only touches grants that START IN THE
// FUTURE. Those dates have no Attendance_Daily rows yet, so nothing on the
// attendance sheet changes. And because days are reserved at REQUEST
// rather than at approval, no balance moves either — the days stay held,
// exactly as they are now. It is purely a change of status.
//
// Leave already running or finished is left alone: it happened, and
// re-opening it for approval would be a lie about the past.
//
// Preview first, always. Same shape as the other migrations here.
// params: { types? (default ['Annual']), allTypes?, asOf? }
// ============================================================
function previewLeaveApprovalMigration(params) {
  return llMigrate_(params, false);
}
function commitLeaveApprovalMigration(params) {
  return llMigrate_(params, true);
}

function llMigrate_(params, commit) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grants = ensureLeaveGrantsTab_(ss);
    const dc = lgDecisionCols_(grants);
    const by = String(params.by || 'Dashboard').trim();

    let types = params.types;
    if (typeof types === 'string') { try { types = JSON.parse(types); } catch (e) { types = [types]; } }
    const allTypes = params.allTypes === true || params.allTypes === 'true';
    const typeSet = {};
    (types && types.length ? types : ['Annual']).forEach(t => typeSet[String(t).trim().toLowerCase()] = true);

    const asOf = params.asOf ? startOfDay_(parseFlexibleDate_(params.asOf)) : startOfDay_(new Date());
    const data = grants.getDataRange().getValues();

    const hits = [], leftAlone = [];
    for (let i = 1; i < data.length; i++) {
      const g = llGrantObj_(data[i], dc);
      if (!g.grantId && !g.empId) continue;
      if (g.status.toLowerCase() !== LG_APPROVED.toLowerCase()) continue;
      const s = parseFlexibleDate_(g.start), e = parseFlexibleDate_(g.end);
      if (!s || !e) continue;

      if (!allTypes && !typeSet[g.type.toLowerCase()]) continue;

      if (startOfDay_(s) <= asOf) {
        // Already started (or finished). Leave it alone — it happened.
        leftAlone.push({ grantId: g.grantId, name: g.name, type: g.type,
                         start: formatDdMmYyyy_(s), end: formatDdMmYyyy_(e),
                         why: startOfDay_(e) < asOf ? 'already finished' : 'already in progress' });
        continue;
      }

      hits.push({ rowNum: i + 1, grantId: g.grantId, empId: g.empId, name: g.name,
                  type: g.type, days: g.days,
                  start: formatDdMmYyyy_(s), end: formatDdMmYyyy_(e),
                  grantedBy: g.grantedBy });
    }

    if (commit) {
      const note = 'Moved to pending by the approval-gate migration on ' + formatDdMmYyyy_(new Date());
      hits.forEach(h => {
        grants.getRange(h.rowNum, 8).setValue(LG_PENDING);
        llStamp_(grants, h.rowNum, dc, by, note);
      });
    }

    const msg = (commit ? 'Moved ' : 'Would move ') + hits.length +
      ' future-dated grant(s) to "' + LG_PENDING + '"' +
      (allTypes ? ' (all leave types)' : ' (' + Object.keys(typeSet).join(', ') + ' only)') + '. ' +
      leftAlone.length + ' left alone because they have already started. ' +
      'No balances move and no attendance changes — future dates have no attendance rows yet.';

    return { success: true, committed: !!commit, asOf: formatDdMmYyyy_(asOf),
             willChange: hits.map(h => ({ grantId: h.grantId, empId: h.empId, name: h.name,
                                          type: h.type, start: h.start, end: h.end,
                                          days: h.days, grantedBy: h.grantedBy })),
             leftAlone: leftAlone, count: hits.length, message: msg };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// DAILY STATUS ENGINE — flip to On Leave on start date; list reactivations.
// Auto-flips Active -> On Leave when a payroll-excluding APPROVED grant
// covers today. Does NOT auto-flip back; returns a reactivation list.
// Run on a daily time-based trigger, or manually / via URL.
// ------------------------------------------------------------
function applyLeaveStatusChanges(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grantsSheet = ensureLeaveGrantsTab_(ss);
    const today = startOfDay_(new Date());

    // Which employees SHOULD be on leave today (a payroll-excluding, APPROVED grant covers today)?
    const shouldBeOnLeave = {}; const coveringGrant = {};
    const gdata = grantsSheet.getDataRange().getValues();
    for (let i = 1; i < gdata.length; i++) {
      const empId = String(gdata[i][1] || '').trim();
      const type = String(gdata[i][3] || '').trim();
      const status = String(gdata[i][7] || '').trim().toLowerCase();
      // Pending, rejected and cancelled leave must never move anybody's
      // status — only an approval does.
      if (!empId || status !== LG_APPROVED.toLowerCase()) continue;
      const s = parseFlexibleDate_(gdata[i][4]), e = parseFlexibleDate_(gdata[i][5]);
      if (!s || !e) continue;
      if (today >= startOfDay_(s) && today <= endOfDay_(e) && excludesFromPayroll_(ss, type)) {
        shouldBeOnLeave[empId] = true;
        coveringGrant[empId] = { type: type, end: formatDdMmYyyy_(e) };
      }
    }

    // Compare with current HR Maya status.
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');
    if (cStatus === -1) return { success: false, error: 'HR Maya has no Status column.' };

    const flippedToLeave = []; const toReactivate = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      const name = cName !== -1 ? String(hd[i][cName] || '').trim() : id;
      const status = String(hd[i][cStatus] || '').trim().toLowerCase();

      if (shouldBeOnLeave[id] && status !== 'on leave') {
        hr.getRange(i + 1, cStatus + 1).setValue('On Leave');
        flippedToLeave.push({ id: id, name: name, until: coveringGrant[id].end, type: coveringGrant[id].type });
      } else if (!shouldBeOnLeave[id] && status === 'on leave') {
        toReactivate.push({ id: id, name: name });
      }
    }

    return {
      success: true, ranOn: formatDdMmYyyy_(new Date()),
      flippedToLeave: flippedToLeave, toReactivate: toReactivate,
      message: 'Set ' + flippedToLeave.length + ' employee(s) to On Leave. ' +
        toReactivate.length + ' whose leave has ended need reactivation (not auto-changed): ' +
        (toReactivate.length ? toReactivate.map(r => r.name + ' (' + r.id + ')').join(', ') : 'none') + '.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Read-only: who currently needs reactivating (for the dashboard/Grant page).
function getLeaveStatusReminders(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grantsSheet = ensureLeaveGrantsTab_(ss);
    const today = startOfDay_(new Date());
    const shouldBeOnLeave = {};
    const gdata = grantsSheet.getDataRange().getValues();
    for (let i = 1; i < gdata.length; i++) {
      const empId = String(gdata[i][1] || '').trim();
      const type = String(gdata[i][3] || '').trim();
      const status = String(gdata[i][7] || '').trim().toLowerCase();
      if (!empId || status !== LG_APPROVED.toLowerCase()) continue;
      const s = parseFlexibleDate_(gdata[i][4]), e = parseFlexibleDate_(gdata[i][5]);
      if (s && e && today >= startOfDay_(s) && today <= endOfDay_(e) && excludesFromPayroll_(ss, type)) shouldBeOnLeave[empId] = true;
    }
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');
    const toReactivate = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      if (String(hd[i][cStatus] || '').trim().toLowerCase() === 'on leave' && !shouldBeOnLeave[id]) {
        toReactivate.push({ id: id, name: cName !== -1 ? String(hd[i][cName] || '').trim() : id });
      }
    }
    return { success: true, toReactivate: toReactivate };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// Reconcile HR Maya "Remaining Vacation" vs Leave_Balances (Annual).
// ------------------------------------------------------------
function getRemainingVacationTally(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const cRem = H.findIndex(h => String(h).trim().toLowerCase() === 'remaining vacation');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');

    // Leave_Balances Annual remaining by empId.
    const balAnnual = {};
    const bal = ss.getSheetByName(LEAVE_BALANCES_TAB);
    if (bal) {
      const bd = bal.getDataRange().getValues();
      for (let i = 1; i < bd.length; i++) {
        if (String(bd[i][2] || '').trim().toLowerCase() === 'annual') balAnnual[String(bd[i][0] || '').trim()] = Number(bd[i][6]) || 0;
      }
    }

    const rows = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      const hrRem = cRem !== -1 ? (Number(hd[i][cRem]) || 0) : 0;
      const balRem = (id in balAnnual) ? balAnnual[id] : null;
      rows.push({
        empId: id, name: cName !== -1 ? String(hd[i][cName] || '').trim() : id,
        status: cStatus !== -1 ? String(hd[i][cStatus] || '').trim() : '',
        hrRemaining: hrRem, balanceRemaining: balRem,
        mismatch: (balRem !== null && Math.abs(balRem - hrRem) > 0.001)
      });
    }
    const mism = rows.filter(r => r.mismatch).length;
    return { success: true, data: rows, mismatches: mism,
      message: mism ? (mism + ' employee(s) have a mismatch between HR Maya Remaining Vacation and Leave_Balances.') : 'HR Maya and Leave_Balances agree on Annual remaining.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}