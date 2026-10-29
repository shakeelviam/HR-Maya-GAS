// ============================================================
// AbsenceDeductions.gs
// ------------------------------------------------------------
// THE MISSING FUNCTION. getPayrollReview already calls
// getAbsenceDeductionsForPeriod_ behind a typeof guard, so until now
// it has silently returned {} and absence has ALWAYS deducted zero.
//
// SINGLE SOURCE OF TRUTH: Attendance_Daily. Nothing else is consulted
// for attendance. Leave_Grants is still the authority for LEAVE, and
// is deliberately NOT re-read here — see the no-double-deduct rule.
//
// THE RULE
//   Deduct one day (Basic / 26) when, and only when, the
//   Attendance_Daily row for that employee-day has Status === 'Absent'.
//
//   Everything else is explicitly NOT deducted:
//     'Present'    → worked (incl. Half Day — see below)
//     'Day Off'    → roster-planned or kiosk self-reported
//     'On Leave'   → owned by getUnpaidLeaveDaysForPeriod_ (Leave_Grants).
//                    Deducting here too would DOUBLE-DEDUCT unpaid leave.
//     'No Roster'  → an ADMIN gap, not the employee's fault. They can
//                    still check in normally. Flagged, never deducted.
//
// HALF DAY / CHECK PUNCHES
//   Attendance_Daily marks Day Type 'Half Day' when presence is under
//   the shift's required hours, and 'Check Punches' when presence
//   computes implausibly (>20h). Both keep Status 'Present', so this
//   function does not deduct for them. They are surfaced for HR review
//   instead — see getHalfDayReviewQueue() below. HR decides; the code
//   does not silently dock anyone.
//
// DIVISOR: Basic / 26, matching getUnpaidLeaveDaysForPeriod_ exactly.
//   (Encashment intentionally uses /30 — calendar days, different basis.)
// ============================================================

const ABSENCE_DIVISOR = 26;

// Statuses that cost the employee money. Deliberately a single-item
// list so adding to it is an explicit, visible decision.
const ABSENCE_DEDUCTIBLE_STATUSES = ['Absent'];

// Statuses that must NEVER be deducted here, with the reason.
// Kept as data (not just comments) so the intent survives future edits.
const ABSENCE_EXEMPT_STATUSES = {
  'Present'  : 'worked',
  'Day Off'  : 'planned or self-reported day off',
  'On Leave' : 'owned by unpaid-leave logic — deducting here double-deducts',
  'No Roster': 'admin gap, not the employee\'s fault',
};

// ── Column resolution, header-based (immune to column re-ordering) ────
function ad_cols_(sheet) {
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const c = (names) => {
    for (const n of names) {
      const i = H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
      if (i !== -1) return i;
    }
    return -1;
  };
  return {
    id:       c(['Employee ID']),
    name:     c(['Name', 'Employee Name']),
    date:     c(['Date']),
    location: c(['Location']),
    shift:    c(['Shift']),
    presence: c(['Presence (hrs)', 'Presence']),
    dayType:  c(['Day Type']),
    status:   c(['Status']),
    source:   c(['Source']),
    flags:    c(['Flags']),
  };
}

function ad_normDate_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0') + '-' + m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2, '0') + '-' + m[2].padStart(2, '0') + '-' + m[1];
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}

// ══════════════════════════════════════════════════════════════════
// THE FUNCTION getPayrollReview HAS BEEN CALLING ALL ALONG
// Returns { empId: { absentDays, amount, dates[] } }
// ══════════════════════════════════════════════════════════════════
function getAbsenceDeductionsForPeriod_(ss, periodStart, periodEnd) {
  const out = {};
  try {
    const sheet = ss.getSheetByName('Attendance_Daily');
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('getAbsenceDeductionsForPeriod_: Attendance_Daily empty or missing — no deductions.');
      return out;
    }

    const C = ad_cols_(sheet);
    if (C.id < 0 || C.date < 0 || C.status < 0) {
      Logger.log('getAbsenceDeductionsForPeriod_: Attendance_Daily missing Employee ID / Date / Status. No deductions applied.');
      return out;
    }

    // Basic salary per employee (Basic ONLY — never gross, never allowances).
    const hr = ss.getSheetByName('HR Maya');
    const hrH = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0];
    const hc = (n) => hrH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const HID = hc('Employee ID'), HBASIC = hc('Basic Salary');
    const basicById = {};
    if (HID !== -1 && HBASIC !== -1 && hr.getLastRow() > 1) {
      hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().forEach(r => {
        const id = String(r[HID] || '').trim();
        if (id) basicById[id] = Number(r[HBASIC]) || 0;
      });
    }

    const s = startOfDay_(periodStart), e = endOfDay_(periodEnd);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    rows.forEach(r => {
      const id = String(r[C.id] || '').trim();
      if (!id) return;

      const status = String(r[C.status] || '').trim();
      if (ABSENCE_DEDUCTIBLE_STATUSES.indexOf(status) === -1) return; // exempt

      const dStr = ad_normDate_(r[C.date]);
      const d = parseFlexibleDate_(dStr);
      if (!d || d < s || d > e) return;

      const basic = basicById[id] || 0;
      const daily = KD(basic / ABSENCE_DIVISOR);

      if (!out[id]) out[id] = { absentDays: 0, amount: 0, dates: [] };
      out[id].absentDays += 1;
      out[id].amount = KD(out[id].amount + daily);
      out[id].dates.push(dStr);
    });

    return out;
  } catch (err) {
    // Never let an attendance problem silently zero out or crash payroll.
    Logger.log('getAbsenceDeductionsForPeriod_ ERROR: ' + err.toString());
    return out;
  }
}

// ══════════════════════════════════════════════════════════════════
// HALF DAY / BROKEN PUNCH REVIEW QUEUE — for HR, never auto-deducted
// Returns rows needing a human decision before payroll is trusted.
// ══════════════════════════════════════════════════════════════════
function getHalfDayReviewQueue(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Attendance_Daily');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, rows: [] };

    const C = ad_cols_(sheet);
    if (C.id < 0 || C.date < 0) return { success: false, error: 'Attendance_Daily missing key columns.' };

    const from = params.from ? startOfDay_(parseFlexibleDate_(params.from)) : null;
    const to   = params.to   ? endOfDay_(parseFlexibleDate_(params.to))     : null;

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const out = [];

    rows.forEach((r, i) => {
      const id = String(r[C.id] || '').trim();
      if (!id) return;

      const dayType = C.dayType >= 0 ? String(r[C.dayType] || '').trim() : '';
      const flags   = C.flags   >= 0 ? String(r[C.flags]   || '').trim() : '';

      // Two things a human must look at:
      //   Half Day      — worked less than the shift's required hours
      //   Check Punches — presence computed implausibly (broken data)
      const needsReview = (dayType === 'Half Day' || dayType === 'Check Punches');
      if (!needsReview) return;

      const dStr = ad_normDate_(r[C.date]);
      const d = parseFlexibleDate_(dStr);
      if (from && d < from) return;
      if (to   && d > to)   return;

      out.push({
        row:      i + 2,
        empId:    id,
        name:     C.name     >= 0 ? String(r[C.name] || '')     : '',
        date:     dStr,
        location: C.location >= 0 ? String(r[C.location] || '') : '',
        shift:    C.shift    >= 0 ? String(r[C.shift] || '')    : '',
        presence: C.presence >= 0 ? Number(r[C.presence]) || 0  : 0,
        dayType:  dayType,
        status:   C.status   >= 0 ? String(r[C.status] || '')   : '',
        flags:    flags,
        reason:   dayType === 'Check Punches'
                    ? 'Presence computed implausibly — punches likely wrong'
                    : 'Worked less than required hours for the shift',
      });
    });

    out.sort((a, b) => (a.date + a.name).localeCompare(b.date + b.name));
    return { success: true, rows: out, count: out.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ══════════════════════════════════════════════════════════════════
// PRE-PAYROLL SAFETY CHECK — run before any payroll run.
// Surfaces everything that would make the numbers untrustworthy.
// ══════════════════════════════════════════════════════════════════
function getPayrollReadiness(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const from = params.from, to = params.to;
    if (!from || !to) return { success: false, error: 'from and to required (dd-mm-yyyy).' };

    const issues = [];

    // 1. Half Day / broken punches awaiting HR review
    const review = getHalfDayReviewQueue({ from: from, to: to });
    if (review.success && review.count > 0) {
      const halfDays = review.rows.filter(r => r.dayType === 'Half Day').length;
      const broken   = review.rows.filter(r => r.dayType === 'Check Punches').length;
      if (halfDays) issues.push({ severity: 'review', type: 'Half Day',
        count: halfDays, message: halfDays + ' half-day(s) need HR review before payroll.' });
      if (broken) issues.push({ severity: 'blocker', type: 'Check Punches',
        count: broken, message: broken + ' row(s) have implausible presence — punches must be fixed first.' });
    }

    // 2. "No Roster" days — not deducted, but a coverage gap worth knowing
    const sheet = ss.getSheetByName('Attendance_Daily');
    if (sheet && sheet.getLastRow() > 1) {
      const C = ad_cols_(sheet);
      const s = startOfDay_(parseFlexibleDate_(from)), e = endOfDay_(parseFlexibleDate_(to));
      let noRoster = 0, absent = 0;
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().forEach(r => {
        const d = parseFlexibleDate_(ad_normDate_(r[C.date]));
        if (!d || d < s || d > e) return;
        const st = String(r[C.status] || '').trim();
        if (st === 'No Roster') noRoster++;
        if (st === 'Absent')    absent++;
      });
      if (noRoster) issues.push({ severity: 'info', type: 'No Roster', count: noRoster,
        message: noRoster + ' day(s) have no published roster. NOT deducted, but coverage is unplanned.' });
      if (absent) issues.push({ severity: 'info', type: 'Absent', count: absent,
        message: absent + ' absent day(s) WILL be deducted at Basic/26.' });
    }

    // 3. Legacy encashment rows sitting in OT_Entries (the double-pay bug)
    const ot = ss.getSheetByName('OT_Entries');
    if (ot && ot.getLastRow() > 1) {
      const H = ot.getRange(1, 1, 1, ot.getLastColumn()).getValues()[0];
      const cSrc = H.findIndex(h => String(h).trim().toLowerCase() === 'source');
      if (cSrc !== -1) {
        let bad = 0;
        ot.getRange(2, 1, ot.getLastRow() - 1, ot.getLastColumn()).getValues().forEach(r => {
          if (String(r[cSrc] || '').trim() === 'Leave Encashment') bad++;
        });
        if (bad) issues.push({ severity: 'blocker', type: 'Encashment in OT_Entries', count: bad,
          message: bad + ' encashment row(s) are sitting in OT_Entries and would be PAID TWICE. Remove them before running payroll.' });
      }
    }

    const blockers = issues.filter(i => i.severity === 'blocker').length;
    return {
      success: true,
      ready: blockers === 0,
      blockerCount: blockers,
      issues: issues,
      message: blockers === 0
        ? 'No blockers. ' + issues.length + ' item(s) to be aware of.'
        : blockers + ' BLOCKER(S) must be resolved before running payroll.',
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}