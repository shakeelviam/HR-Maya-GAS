// ============================================================
// LeaveSalary.gs — Maya HR : leave salary paid in advance, by voucher
// ------------------------------------------------------------
// A man going on annual leave is paid before he flies. It is NOT part
// of the monthly payroll: the payroll for the month may already have
// run, and he needs the money in his hand tomorrow, not on the 25th.
// So this works the way encashment works — a voucher with its own
// number, paid separately, logged so nothing is paid twice.
//
// THE RULES THIS FILE ENCODES
//   Rate          Basic / 30 per day. Same divisor encashment uses.
//                 (OT, absence and half-day use Basic/26 — not here.)
//   A month       30 days. So 2 months = 60 days = 2 x Basic exactly.
//   Ticket        A flat amount for anyone who travels on leave, and
//                 never for encashment. Held in Script Properties
//                 (LEAVE_TICKET_KD), not written into the code, so it
//                 changes without a redeployment.
//   Already paid  Days inside the leave that a salary already covered
//                 — e.g. leave starts on the 28th and the month was
//                 paid in full. Entered by hand; the money is reduced
//                 but the period covered is not.
//   Who qualifies Leave types that are Paid AND Exclude From Payroll.
//                 Those are the people who get nothing in the monthly
//                 run while they are away, so they must be paid here.
//                 Read from Leave_Types — no list of types in code.
//
// WHAT IT REMEMBERS
//   Every payment records Paid From / Paid To against the grant, so a
//   96-day leave paid two months at a time knows, months later, which
//   days are still owed. That is the part a spreadsheet by hand loses.
//
// TAB THIS FILE OWNS
//   Leave_Salary_Log
//
// ENDPOINTS (routed in Code.gs)
//   getLeaveSalaryList, payLeaveSalary, cancelLeaveSalary,
//   getLeaveSalaryHistory, generateLeaveSalaryVoucher,
//   reprintLeaveSalaryVoucher, setLeaveTicketAmount
// ============================================================

const LEAVE_SALARY_LOG_TAB = 'Leave_Salary_Log';
const LS_TICKET_PROP       = 'LEAVE_TICKET_KD';
const LS_TICKET_DEFAULT    = 120;          // seeded once; change it in the app, not here
const LS_DAYS_PER_MONTH    = 30;           // a month of leave salary, per the Basic/30 rate
const LS_VOUCHER_FOLDER    = 'Maya Leave Salary Vouchers';
const LS_PAID              = 'Paid';
const LS_VOID              = 'Cancelled';

const LS_HEADERS = [
  'Voucher No', 'Employee ID', 'Name', 'Grant ID', 'Leave Type',
  'Leave Start', 'Leave End', 'Leave Days',
  'Paid From', 'Paid To', 'Months', 'Days Paid',
  'Basic', 'Rate Per Day', 'Leave Salary',
  'Less Days', 'Less Amount',
  'Other Deduction', 'Deduction Note',
  'Ticket', 'Net Paid',
  'Paid On', 'Paid By', 'Remark', 'Status'
];

// ------------------------------------------------------------
// Tab + settings
// ------------------------------------------------------------
function lsSheet_(ss) {
  let sh = ss.getSheetByName(LEAVE_SALARY_LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(LEAVE_SALARY_LOG_TAB);
    sh.getRange(1, 1, 1, LS_HEADERS.length).setValues([LS_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // An existing tab keeps its rows: only missing columns are appended.
  const have = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0]
                 .map(h => String(h).trim().toLowerCase());
  LS_HEADERS.forEach(h => {
    if (have.indexOf(h.toLowerCase()) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold');
      have.push(h.toLowerCase());
    }
  });
  return sh;
}

function lsCols_(sheet) {
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                 .map(h => String(h).trim().toLowerCase());
  const c = {};
  LS_HEADERS.forEach(h => { c[h] = H.indexOf(h.toLowerCase()); });
  return c;
}

// The ticket amount lives in Script Properties so it is changed in the
// app, not in this file. Run bare from the editor it just prints it.
function setLeaveTicketAmount(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const props = PropertiesService.getScriptProperties();
    const asked = (params.amount === undefined || params.amount === null || params.amount === '')
      ? null : Number(params.amount);
    if (asked === null) {
      const cur = lsTicketAmount_();
      Logger.log('Leave travel ticket is currently KD ' + cur.toFixed(3) +
                 '. To change it: setLeaveTicketAmount({amount: 135})');
      return { success: true, amount: cur, message: 'Ticket is KD ' + cur.toFixed(3) + '.' };
    }
    if (!(asked >= 0)) return { success: false, error: 'Ticket amount must be zero or more.' };
    props.setProperty(LS_TICKET_PROP, String(asked));
    Logger.log('Leave travel ticket set to KD ' + asked.toFixed(3));
    return { success: true, amount: asked, message: 'Ticket set to KD ' + asked.toFixed(3) + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function lsTicketAmount_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(LS_TICKET_PROP);
  if (raw === null || raw === '') { props.setProperty(LS_TICKET_PROP, String(LS_TICKET_DEFAULT)); return LS_TICKET_DEFAULT; }
  const n = Number(raw);
  return (n >= 0) ? n : LS_TICKET_DEFAULT;
}

function lsMoney_(v) { return Math.round((Number(v) || 0) * 1000) / 1000; }

// A cell read back from the sheet may be a real Date, because Sheets turns
// "28-09-2026" into one. Printing that raw gives
// "Mon Sep 28 2026 00:00:00 GMT+0300 (Arabian Standard Time)", so every
// date leaving this file goes through here first.
function lsDate_(v) {
  if (v === null || v === undefined || v === '') return '';
  // Anything that walks like a Date is formatted straight from its parts.
  // instanceof is not relied on: a value handed back by Sheets does not
  // always satisfy it, and that is how the timestamps got onto a voucher.
  if (v && typeof v.getFullYear === 'function' && !isNaN(v.getTime())) {
    const p = (n) => (n < 10 ? '0' + n : String(n));
    return p(v.getDate()) + '-' + p(v.getMonth() + 1) + '-' + v.getFullYear();
  }
  const d = parseFlexibleDate_(v);
  if (d) return formatDdMmYyyy_(d);
  // Last resort: a JS date string that slipped through with a time on it.
  const s = String(v).trim();
  const m = s.match(/^\w{3} (\w{3}) (\d{1,2}) (\d{4})/);
  if (m) {
    const mon = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'].indexOf(m[1]) + 1;
    if (mon > 0) {
      const p = (n) => (Number(n) < 10 ? '0' + Number(n) : String(Number(n)));
      return p(m[2]) + '-' + p(mon) + '-' + m[3];
    }
  }
  return s;
}
function lsAddDays_(d, n) { const x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function lsDaysBetween_(a, b) { return Math.round((startOfDay_(b) - startOfDay_(a)) / 86400000) + 1; }

// Which leave types need paying this way: paid leave that takes the man
// off the payroll while he is away. Read from Leave_Types every time.
function lsQualifyingTypes_(ss) {
  const out = {};
  const sh = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (!sh) return out;
  const d = sh.getDataRange().getValues();
  if (!d.length) return out;
  const H = d[0].map(h => String(h).trim().toLowerCase());
  const cType = H.indexOf('type'), cPaid = H.indexOf('is paid'),
        cExcl = H.indexOf('excludes from payroll'), cActive = H.indexOf('active');
  const yes = v => String(v).trim().toLowerCase() === 'yes';
  for (let i = 1; i < d.length; i++) {
    const t = String(d[i][cType] || '').trim();
    if (!t) continue;
    if (cActive !== -1 && String(d[i][cActive]).trim() && !yes(d[i][cActive])) continue;
    const paid = (cPaid === -1) ? true : yes(d[i][cPaid]);
    const excl = (cExcl === -1) ? false : yes(d[i][cExcl]);
    if (paid && excl) out[t.toLowerCase()] = t;
  }
  return out;
}

// Everything already paid against a grant, so we know where to resume.
// Cancelled vouchers do not count.
function lsPaidSoFar_(ss) {
  const out = {};   // grantId -> { days, amount, paidTo(Date|null), vouchers:[] }
  const sh = ss.getSheetByName(LEAVE_SALARY_LOG_TAB);
  if (!sh || sh.getLastRow() < 2) return out;
  const c = lsCols_(sh);
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const gid = String(d[i][c['Grant ID']] || '').trim();
    if (!gid) continue;
    const status = c['Status'] >= 0 ? String(d[i][c['Status']] || '').trim() : LS_PAID;
    if (status.toLowerCase() === LS_VOID.toLowerCase()) continue;
    if (!out[gid]) out[gid] = { days: 0, amount: 0, paidTo: null, vouchers: [] };
    out[gid].days   += Number(d[i][c['Days Paid']]) || 0;
    out[gid].amount += Number(d[i][c['Net Paid']]) || 0;
    out[gid].vouchers.push(String(d[i][c['Voucher No']] || '').trim());
    const pt = parseFlexibleDate_(d[i][c['Paid To']]);
    if (pt && (!out[gid].paidTo || pt > out[gid].paidTo)) out[gid].paidTo = pt;
  }
  return out;
}

// ------------------------------------------------------------
// READ — who is owed leave salary, and how much
// ------------------------------------------------------------
function getLeaveSalaryList(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const grants = ensureLeaveGrantsTab_(ss);
    lsSheet_(ss);
    const dc = lgDecisionCols_(grants);
    const qualifying = lsQualifyingTypes_(ss);
    const paid = lsPaidSoFar_(ss);
    const ticket = lsTicketAmount_();

    // Basic salary for everyone in one read, rather than one read each.
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const hH = hd[0].map(h => String(h).trim().toLowerCase());
    const hcId = hH.indexOf('employee id'), hcBasic = hH.indexOf('basic salary'),
          hcDesig = hH.indexOf('designation'), hcLoc = hH.indexOf('location'),
          hcStatus = hH.indexOf('status');
    const emp = {};
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][hcId] || '').trim();
      if (!id) continue;
      emp[id.toUpperCase()] = {
        basic: hcBasic !== -1 ? (Number(hd[i][hcBasic]) || 0) : 0,
        designation: hcDesig !== -1 ? String(hd[i][hcDesig] || '') : '',
        location: hcLoc !== -1 ? String(hd[i][hcLoc] || '') : '',
        status: hcStatus !== -1 ? String(hd[i][hcStatus] || '') : ''
      };
    }

    const gd = grants.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < gd.length; i++) {
      const g = llGrantObj_(gd[i], dc);
      if (!g.grantId || !g.empId) continue;
      if (String(g.status).toLowerCase() !== LG_APPROVED.toLowerCase()) continue;
      if (!qualifying[String(g.type).trim().toLowerCase()]) continue;

      const start = parseFlexibleDate_(g.start), end = parseFlexibleDate_(g.end);
      if (!start || !end) continue;

      const e = emp[g.empId.toUpperCase()] || { basic: 0, designation: '', location: '', status: '' };
      const rate = lsMoney_(e.basic / LS_DAYS_PER_MONTH);
      const done = paid[g.grantId] || { days: 0, amount: 0, paidTo: null, vouchers: [] };
      const totalDays = g.days || lsDaysBetween_(start, end);
      const daysLeft = Math.max(0, totalDays - done.days);

      // Where the next payment picks up.
      const nextFrom = done.paidTo ? lsAddDays_(done.paidTo, 1) : start;
      const suggestMonths = Math.min(2, Math.ceil(daysLeft / LS_DAYS_PER_MONTH)) || 0;

      // Where the leave stands today. A leave that already ended and was
      // never paid here was almost certainly settled by hand before this
      // page existed — it should stand out rather than look like a debt.
      const today = startOfDay_(new Date());
      const when = (today < startOfDay_(start)) ? 'upcoming'
                 : (today > endOfDay_(end))     ? 'ended' : 'running';
      const daysToStart = Math.round((startOfDay_(start) - today) / 86400000);

      rows.push({
        grantId: g.grantId, empId: g.empId, name: g.name, type: g.type,
        designation: e.designation, location: e.location, empStatus: e.status,
        start: formatDdMmYyyy_(start), end: formatDdMmYyyy_(end),
        leaveDays: totalDays,
        basic: lsMoney_(e.basic), ratePerDay: rate,
        daysPaid: done.days, amountPaid: lsMoney_(done.amount),
        paidTo: done.paidTo ? formatDdMmYyyy_(done.paidTo) : '',
        daysRemaining: daysLeft,
        nextFrom: formatDdMmYyyy_(nextFrom),
        suggestMonths: suggestMonths,
        when: when, daysToStart: daysToStart,
        fullyPaid: daysLeft <= 0,
        vouchers: done.vouchers.filter(String),
        noBasic: !(e.basic > 0)
      });
    }

    // Soonest leave first — the man flying tomorrow is the one who matters.
    rows.sort((a, b) => {
      const pa = parseFlexibleDate_(a.start), pb = parseFlexibleDate_(b.start);
      if (a.fullyPaid !== b.fullyPaid) return a.fullyPaid ? 1 : -1;
      return (pa ? pa.getTime() : 0) - (pb ? pb.getTime() : 0);
    });

    return {
      success: true, data: rows, ticket: ticket, daysPerMonth: LS_DAYS_PER_MONTH,
      owing: rows.filter(r => !r.fullyPaid).length, total: rows.length
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// The sum, in one place, so the preview and the payment cannot differ.
// ------------------------------------------------------------
function lsCompute_(o) {
  // o: { basic, months, daysRemaining, nextFrom(Date), lessDays,
  //      otherDeduction, ticket }
  //
  // Two different deductions, because they are two different things:
  //   lessDays        whole days of this leave a salary already covered,
  //                   valued at the man's own daily rate.
  //   otherDeduction  a flat amount — a food allowance already paid for
  //                   months he will be away, an advance, anything else.
  //                   It carries its own note onto the voucher.
  const rate = lsMoney_(o.basic / LS_DAYS_PER_MONTH);
  const asked = Math.round(Number(o.months) * LS_DAYS_PER_MONTH);
  const daysPaid = Math.max(0, Math.min(asked, o.daysRemaining));
  const capped = daysPaid < asked;
  const lessDays = Math.max(0, Number(o.lessDays) || 0);
  const leaveSalary = lsMoney_(rate * daysPaid);
  const lessAmount = lsMoney_(rate * lessDays);
  const otherDeduction = lsMoney_(Math.max(0, Number(o.otherDeduction) || 0));
  const ticket = lsMoney_(Number(o.ticket) || 0);
  const net = lsMoney_(leaveSalary - lessAmount - otherDeduction + ticket);
  const paidFrom = o.nextFrom;
  const paidTo = daysPaid > 0 ? lsAddDays_(paidFrom, daysPaid - 1) : null;
  return { rate: rate, daysPaid: daysPaid, capped: capped, lessDays: lessDays,
           leaveSalary: leaveSalary, lessAmount: lessAmount,
           otherDeduction: otherDeduction, ticket: ticket,
           net: net, paidFrom: paidFrom, paidTo: paidTo };
}

// ------------------------------------------------------------
// PAY — one man or a whole group, in one call
// ------------------------------------------------------------
// params: { items:[{grantId, months, lessDays, ticket, remark}], by, paidOn,
//           preview:true|false }
// preview:true computes and returns the same figures without writing a row,
// so the page can show the group total before anything is committed.
// ------------------------------------------------------------
function payLeaveSalary(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const items = Array.isArray(params.items) ? params.items : [];
    if (!items.length) return { success: false, error: 'Nothing selected to pay.' };

    const preview = !!params.preview;
    const by = String(params.by || 'Dashboard').trim();
    const paidOn = parseFlexibleDate_(params.paidOn) || new Date();

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const list = getLeaveSalaryList({});
    if (!list.success) return list;
    const byGrant = {};
    list.data.forEach(r => { byGrant[r.grantId] = r; });

    const sh = lsSheet_(ss);
    const c = lsCols_(sh);
    const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');

    const done = [], refused = [];
    let grandTotal = 0, n = 0;

    items.forEach(function (it, idx) {
      const gid = String(it.grantId || '').trim();
      const r = byGrant[gid];
      if (!r) { refused.push((gid || '(no grant id)') + ' — not an approved leave that needs a leave-salary payment.'); return; }
      if (r.noBasic) { refused.push(r.name + ' — no Basic Salary on his HR Maya row.'); return; }
      if (r.fullyPaid) { refused.push(r.name + ' — this leave is already paid in full.'); return; }

      const months = Number(it.months);
      if (!(months > 0)) { refused.push(r.name + ' — months must be more than zero.'); return; }

      // "Settle" records that a leave was already paid some other way —
      // before this page existed, or in cash. It consumes the days so the
      // man stops showing as owed, and pays nothing. The row stays in the
      // log, marked, so it is obvious later what happened.
      const settle = !!it.settle;

      const lessDays = Math.max(0, Number(it.lessDays) || 0);
      const calc = lsCompute_({
        basic: r.basic, months: months, daysRemaining: r.daysRemaining,
        nextFrom: parseFlexibleDate_(r.nextFrom), lessDays: settle ? 0 : lessDays,
        otherDeduction: settle ? 0 : it.otherDeduction,
        ticket: (settle || it.ticket === undefined || it.ticket === null || it.ticket === '')
                  ? 0 : Number(it.ticket)
      });
      if (settle) { calc.leaveSalary = 0; calc.lessAmount = 0; calc.otherDeduction = 0; calc.ticket = 0; calc.net = 0; }

      if (calc.daysPaid <= 0) { refused.push(r.name + ' — nothing left to pay on this leave.'); return; }
      if (calc.lessDays > calc.daysPaid) {
        refused.push(r.name + ' — you entered ' + calc.lessDays + ' days already paid, but this payment only covers ' +
                     calc.daysPaid + ' days.'); return;
      }
      if (calc.net < 0) { refused.push(r.name + ' — the deduction is larger than the payment.'); return; }

      const line = {
        settle: settle,
        voucherNo: (settle ? 'LS-SET-' : 'LS-') + stamp + '-' + String(idx + 1).padStart(2, '0'),
        grantId: gid, empId: r.empId, name: r.name, type: r.type,
        leaveStart: r.start, leaveEnd: r.end, leaveDays: r.leaveDays,
        paidFrom: formatDdMmYyyy_(calc.paidFrom),
        paidTo: formatDdMmYyyy_(calc.paidTo),
        months: months, daysPaid: calc.daysPaid, capped: calc.capped,
        basic: r.basic, ratePerDay: calc.rate,
        leaveSalary: calc.leaveSalary, lessDays: calc.lessDays, lessAmount: calc.lessAmount,
        otherDeduction: calc.otherDeduction,
        deductionNote: settle ? '' : String(it.deductionNote || '').trim(),
        ticket: calc.ticket, net: calc.net,
        daysRemainingAfter: r.daysRemaining - calc.daysPaid,
        remark: String(it.remark || '').trim()
      };

      if (!preview) {
        const row = new Array(sh.getLastColumn()).fill('');
        const put = (h, v) => { if (c[h] >= 0) row[c[h]] = v; };
        put('Voucher No', line.voucherNo);   put('Employee ID', line.empId);
        put('Name', line.name);              put('Grant ID', line.grantId);
        put('Leave Type', line.type);
        put('Leave Start', line.leaveStart); put('Leave End', line.leaveEnd);
        put('Leave Days', line.leaveDays);
        put('Paid From', line.paidFrom);     put('Paid To', line.paidTo);
        put('Months', line.months);          put('Days Paid', line.daysPaid);
        put('Basic', line.basic);            put('Rate Per Day', line.ratePerDay);
        put('Leave Salary', line.leaveSalary);
        put('Less Days', line.lessDays);     put('Less Amount', line.lessAmount);
        put('Other Deduction', line.otherDeduction);
        put('Deduction Note', line.deductionNote);
        put('Ticket', line.ticket);          put('Net Paid', line.net);
        put('Paid On', formatDdMmYyyy_(paidOn));
        put('Paid By', by);
        put('Remark', settle
          ? ('Settled outside this page — no payment made.' + (line.remark ? ' ' + line.remark : ''))
          : line.remark);
        put('Status', settle ? 'Settled' : LS_PAID);
        sh.appendRow(row);
      }

      done.push(line); grandTotal += line.net; n++;
    });

    grandTotal = lsMoney_(grandTotal);
    const settled = done.filter(x => x.settle).length;
    const payments = n - settled;
    const verb = preview ? 'would be paid' : 'paid';
    const bits = [];
    if (payments) bits.push(payments + ' leave salary payment' + (payments === 1 ? '' : 's') + ' ' + verb + ', KD ' + grandTotal.toFixed(3) + ' in total.');
    if (settled) bits.push(settled + ' leave' + (settled === 1 ? '' : 's') + ' marked settled (no payment).');
    let message = bits.length ? bits.join(' ') : 'Nothing was paid.';
    if (refused.length) message += ' ' + refused.length + ' skipped.';

    return { success: n > 0 || preview, preview: preview, paid: done,
             refused: refused, count: n, total: grandTotal, message: message };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// CANCEL — void a voucher paid by mistake. The row stays for the
// audit; it simply stops counting, so the days come back as owed.
// ------------------------------------------------------------
function cancelLeaveSalary(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const vno = String(params.voucherNo || '').trim();
    if (!vno) return { success: false, error: 'voucherNo is required' };
    const by = String(params.by || 'Dashboard').trim();
    const why = String(params.reason || params.note || '').trim();

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = lsSheet_(ss);
    const c = lsCols_(sh);
    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c['Voucher No']] || '').trim() !== vno) continue;
      const cur = c['Status'] >= 0 ? String(d[i][c['Status']] || '').trim() : LS_PAID;
      if (cur.toLowerCase() === LS_VOID.toLowerCase()) return { success: false, error: vno + ' is already cancelled.' };
      if (c['Status'] >= 0) sh.getRange(i + 1, c['Status'] + 1).setValue(LS_VOID);
      if (c['Remark'] >= 0) {
        const old = String(d[i][c['Remark']] || '').trim();
        const note = 'Cancelled ' + Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy') + ' by ' + by + (why ? ': ' + why : '');
        sh.getRange(i + 1, c['Remark'] + 1).setValue(old ? old + ' | ' + note : note);
      }
      const days = Number(d[i][c['Days Paid']]) || 0;
      return { success: true, voucherNo: vno,
               message: vno + ' cancelled. ' + days + ' day(s) are owed again on that leave.' };
    }
    return { success: false, error: 'Voucher ' + vno + ' not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// HISTORY — every payment, latest first, for reprinting.
// ------------------------------------------------------------
function getLeaveSalaryHistory(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = lsSheet_(ss);
    if (sh.getLastRow() < 2) return { success: true, rows: [] };
    const c = lsCols_(sh);
    const d = sh.getDataRange().getValues();
    const want = String(params.empId || '').trim().toUpperCase();
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][c['Employee ID']] || '').trim();
      if (!id) continue;
      if (want && id.toUpperCase() !== want) continue;
      rows.push({
        row: i + 1,
        voucherNo: String(d[i][c['Voucher No']] || '').trim(),
        empId: id, name: String(d[i][c['Name']] || '').trim(),
        grantId: String(d[i][c['Grant ID']] || '').trim(),
        type: String(d[i][c['Leave Type']] || '').trim(),
        paidFrom: lsDate_(d[i][c['Paid From']]), paidTo: lsDate_(d[i][c['Paid To']]),
        months: Number(d[i][c['Months']]) || 0,
        daysPaid: Number(d[i][c['Days Paid']]) || 0,
        leaveSalary: Number(d[i][c['Leave Salary']]) || 0,
        lessDays: Number(d[i][c['Less Days']]) || 0,
        lessAmount: Number(d[i][c['Less Amount']]) || 0,
        otherDeduction: Number(d[i][c['Other Deduction']]) || 0,
        deductionNote: String(d[i][c['Deduction Note']] || '').trim(),
        ticket: Number(d[i][c['Ticket']]) || 0,
        net: Number(d[i][c['Net Paid']]) || 0,
        paidOn: lsDate_(d[i][c['Paid On']]), paidBy: String(d[i][c['Paid By']] || '').trim(),
        remark: String(d[i][c['Remark']] || '').trim(),
        status: c['Status'] >= 0 ? String(d[i][c['Status']] || '').trim() : LS_PAID
      });
    }
    rows.reverse();
    return { success: true, rows: rows };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ============================================================
// The printable voucher — A5, same shape as the encashment one so
// the two sit together in a file without looking like strangers.
// ============================================================
function generateLeaveSalaryVoucher(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const vno = String(params.voucherNo || '').trim();
    const wantEmp = String(params.empId || '').trim().toUpperCase();
    if (!vno && !wantEmp) return { success: false, error: 'voucherNo or empId is required' };

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = lsSheet_(ss);
    const c = lsCols_(sh);
    const d = sh.getDataRange().getValues();

    let row = null;
    for (let i = d.length - 1; i >= 1; i--) {          // latest first
      const v = String(d[i][c['Voucher No']] || '').trim();
      const id = String(d[i][c['Employee ID']] || '').trim().toUpperCase();
      if (vno) { if (v === vno) { row = d[i]; break; } }
      else if (id === wantEmp) { row = d[i]; break; }
    }
    if (!row) return { success: false, error: 'Leave salary payment not found.' };

    const g = (h) => c[h] >= 0 ? row[c[h]] : '';
    const S = (h) => String(g(h) == null ? '' : g(h)).trim();
    const N = (h) => Number(g(h)) || 0;
    const D = (h) => lsDate_(g(h));          // dates only — never a timestamp

    const name = S('Name'), empId = S('Employee ID');
    const voucherNo = S('Voucher No');
    const status = S('Status') || LS_PAID;

    // Designation and Civil ID off the employee row, for the signature block.
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const hH = hd[0].map(h => String(h).trim().toLowerCase());
    const hcId = hH.indexOf('employee id'), hcDesig = hH.indexOf('designation'),
          hcCivil = hH.indexOf('civil id'), hcLoc = hH.indexOf('location');
    let desig = '', civil = '', loc = '';
    for (let i = 1; i < hd.length; i++) {
      if (String(hd[i][hcId]).trim().toUpperCase() === empId.toUpperCase()) {
        desig = hcDesig !== -1 ? String(hd[i][hcDesig] || '') : '';
        civil = hcCivil !== -1 ? String(hd[i][hcCivil] || '') : '';
        loc   = hcLoc   !== -1 ? String(hd[i][hcLoc]   || '') : '';
        break;
      }
    }

    const ticket = N('Ticket'), lessAmt = N('Less Amount'), lessDays = N('Less Days');
    const otherDed = N('Other Deduction'), dedNote = S('Deduction Note');
    const today = Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy');
    const row_ = (l, v, cls) => '<tr><td class="lbl' + (cls ? ' ' + cls : '') + '">' + ev_esc_(l) +
                                '</td><td class="r' + (cls ? ' ' + cls : '') + '">' + v + '</td></tr>';

    const html =
      // Sized to finish on ONE A5 page. If a line is ever added below,
      // check it still does — a voucher that spills onto a second sheet
      // wastes a page on every payment.
      '<html><head><meta charset="utf-8"><style>' +
      '@page{size:A5;margin:10mm;} body{font-family:Arial,Helvetica,sans-serif;color:#222;font-size:9.5px;margin:0;}' +
      'h2{text-align:center;margin:0;font-size:13px;} .sub{text-align:center;color:#666;font-size:8.5px;margin-bottom:7px;letter-spacing:1px;}' +
      '.ref{display:flex;justify-content:space-between;font-size:8.5px;color:#444;margin-bottom:7px;}' +
      'table{width:100%;border-collapse:collapse;margin-bottom:7px;} td{padding:3.5px 6px;border:0.5px solid #bbb;line-height:1.3;}' +
      '.lbl{color:#555;width:52%;} .r{text-align:right;} .amt{background:#f5f7fa;font-weight:bold;font-size:11px;}' +
      '.nw{white-space:nowrap;}' +
      '.void{color:#b00;font-weight:bold;text-align:center;letter-spacing:2px;margin:2px 0 6px;}' +
      '.words{font-style:italic;color:#333;margin:5px 0 10px;font-size:8.5px;}' +
      '.sig{display:flex;justify-content:space-between;margin-top:18px;} .sigbox{width:45%;text-align:center;}' +
      '.sigline{border-top:1px solid #444;margin-top:30px;padding-top:3px;font-size:8.5px;color:#444;line-height:1.35;}' +
      '.note{font-size:7.5px;color:#888;margin-top:12px;text-align:center;}' +
      '</style></head><body>' +
      '<h2>Maya Tex Restaurants Company</h2>' +
      '<div class="sub">LEAVE SALARY VOUCHER</div>' +
      (status.toLowerCase() === LS_VOID.toLowerCase() ? '<div class="void">CANCELLED</div>' : '') +
      '<div class="ref"><span><b>Voucher No:</b> ' + ev_esc_(voucherNo) + '</span><span><b>Date:</b> ' + today + '</span></div>' +
      '<table>' +
        '<tr><td class="lbl">Employee</td><td>' + ev_esc_(name) + '</td></tr>' +
        '<tr><td class="lbl">Employee ID</td><td>' + ev_esc_(empId) + '</td></tr>' +
        (desig ? '<tr><td class="lbl">Designation</td><td>' + ev_esc_(desig) + '</td></tr>' : '') +
        (loc   ? '<tr><td class="lbl">Location</td><td>' + ev_esc_(loc) + '</td></tr>' : '') +
        (civil ? '<tr><td class="lbl">Civil ID</td><td>' + ev_esc_(civil) + '</td></tr>' : '') +
      '</table>' +
      '<table>' +
        row_('Leave', '<span class="nw">' + ev_esc_(S('Leave Type')) + ' &nbsp; ' + ev_esc_(D('Leave Start')) +
             ' &rarr; ' + ev_esc_(D('Leave End')) + '</span> (' + N('Leave Days') + ' days)') +
        row_('This payment covers', '<span class="nw">' + ev_esc_(D('Paid From')) + ' &rarr; ' + ev_esc_(D('Paid To')) + '</span>') +
        row_('Months', N('Months') + ' month(s) &nbsp;=&nbsp; ' + N('Days Paid') + ' days') +
        row_('Basic Salary', 'KD ' + ev_money_(N('Basic'))) +
        row_('Rate (Basic ÷ ' + LS_DAYS_PER_MONTH + ')', 'KD ' + ev_money_(N('Rate Per Day')) + ' / day') +
        row_('Leave Salary', 'KD ' + ev_money_(N('Leave Salary'))) +
        (lessDays ? row_('Less: ' + lessDays + ' day(s) already paid in salary', '&minus; KD ' + ev_money_(lessAmt)) : '') +
        (otherDed ? row_('Less: ' + (dedNote || 'other deduction'), '&minus; KD ' + ev_money_(otherDed)) : '') +
        (ticket ? row_('Add: Air ticket (leave travel)', 'KD ' + ev_money_(ticket)) : '') +
        row_('NET PAID', 'KD ' + ev_money_(N('Net Paid')), 'amt') +
      '</table>' +
      '<div class="words">Amount in words: ' + ev_esc_(ev_words_(N('Net Paid'))) + '</div>' +
      '<div class="sig">' +
        '<div class="sigbox"><div class="sigline">Employee<br>' + ev_esc_(name) + '<br>Signature &amp; Date</div></div>' +
        '<div class="sigbox"><div class="sigline">' + ev_esc_(GM_NAME) + '<br>' + ev_esc_(GM_TITLE) + '<br>Signature &amp; Date</div></div>' +
      '</div>' +
      '<div class="note">Leave salary paid in advance at Basic ÷ ' + LS_DAYS_PER_MONTH +
      ' per day. Paid separately from the monthly payroll. Amounts in Kuwaiti Dinar (KD), 3 decimals.</div>' +
      '</body></html>';

    const pdf = Utilities.newBlob(html, 'text/html', 'Leave Salary ' + voucherNo + '.html')
                  .getAs('application/pdf')
                  .setName('Leave Salary - ' + name + ' - ' + voucherNo + '.pdf');
    const folder = getOrCreateFolder_(LS_VOUCHER_FOLDER);
    const file = folder.createFile(pdf);
    return { success: true, voucherNo: voucherNo, url: file.getUrl(),
             message: 'Voucher ' + voucherNo + ' ready for ' + name + ' (KD ' + ev_money_(N('Net Paid')) + ').' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Several vouchers in one go, so a group payment prints as one batch.
function reprintLeaveSalaryVoucher(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const list = Array.isArray(params.voucherNos) ? params.voucherNos
               : (params.voucherNo ? [params.voucherNo] : []);
    if (!list.length) return { success: false, error: 'No voucher numbers given.' };
    const out = [], bad = [];
    list.forEach(v => {
      const r = generateLeaveSalaryVoucher({ voucherNo: v });
      if (r.success) out.push({ voucherNo: r.voucherNo, url: r.url }); else bad.push(v + ' — ' + r.error);
    });
    return { success: out.length > 0, vouchers: out, failed: bad,
             message: out.length + ' voucher(s) ready.' + (bad.length ? ' ' + bad.length + ' failed.' : '') };
  } catch (err) { return { success: false, error: err.toString() }; }
}