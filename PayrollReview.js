// ============================================================
// PayrollReview.gs — frontend-driven payroll run
// ============================================================
// Everything is handled on the review screen, not by editing the sheet:
//   getPayrollReview  -> per-employee base + OT (ENTIRE OT_Entries tab) +
//                        encashment + auto Unpaid-leave deduction, plus
//                        empty editable Bonus / Loan / Advance / Disciplinary.
//   runPayrollFromReview -> takes the edited rows, computes Net, writes the
//                        period tab with ITEMISED deductions, logs the run,
//                        then ARCHIVES OT_Entries and starts a fresh tab.
//
// OT model: OT_Entries always holds exactly the current unpaid period's OT.
// The run consumes the WHOLE tab (no date filter), then archives it to
// "OT_Entries <period>" and creates a new empty OT_Entries. No dropped tail.
//
// Unpaid leave: any leave grant whose type has Is Paid = No is deducted at
// Basic/26 x (unpaid days that fall inside the payroll period).
//
// ENDPOINTS: getPayrollReview, runPayrollFromReview
// ============================================================

// Sum the ENTIRE OT_Entries tab per employee (no date filter).
function getAllOtHours_(ss) {
  const out = {};
  const sheet = ss.getSheetByName(OT_ENTRIES_TAB);
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][1] || '').trim();
    const hrs = Number(data[i][3]) || 0;
    if (!id || hrs <= 0) continue;
    out[id] = (out[id] || 0) + hrs;
  }
  return out;
}

// Unpaid-leave days per employee that fall within [periodStart, periodEnd].
// Reads Leave_Grants; a grant counts as unpaid when its type has Is Paid = No.
function getUnpaidLeaveDaysForPeriod_(ss, periodStart, periodEnd) {
  const out = {}; // empId -> { days, ranges:[..] }
  const grants = ss.getSheetByName('Leave_Grants');
  if (!grants) return out;
  const gd = grants.getDataRange().getValues();
  if (gd.length < 2) return out;
  const H = gd[0];
  const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const cEmp = c('Employee ID'), cType = c('Type'), cStart = c('Start'), cEnd = c('End'), cStatus = c('Status');

  // Build a set of unpaid type names from Leave_Types (Is Paid = No).
  const unpaidTypes = {};
  const types = ss.getSheetByName(LEAVE_TYPES_TAB);
  if (types) {
    const td = types.getDataRange().getValues();
    const tH = td[0];
    const tcType = tH.findIndex(h => String(h).trim().toLowerCase() === 'type');
    const tcPaid = tH.findIndex(h => String(h).trim().toLowerCase() === 'is paid');
    for (let i = 1; i < td.length; i++) {
      const t = String(td[i][tcType] || '').trim().toLowerCase();
      if (!t) continue;
      const paid = tcPaid !== -1 ? String(td[i][tcPaid]).trim().toLowerCase() : 'yes';
      if (paid === 'no' || paid === 'false') unpaidTypes[t] = true;
    }
  }

  const pS = startOfDay_(periodStart), pE = endOfDay_(periodEnd);
  for (let i = 1; i < gd.length; i++) {
    const id = String(gd[i][cEmp] || '').trim();
    const type = String(gd[i][cType] || '').trim().toLowerCase();
    const status = String(gd[i][cStatus] || '').trim().toLowerCase();
    if (!id || !type || !unpaidTypes[type]) continue;
    if (status && status !== 'approved') continue;
    const s = parseFlexibleDate_(gd[i][cStart]), e = parseFlexibleDate_(gd[i][cEnd]);
    if (!s || !e) continue;
    // Overlap of [s,e] with [periodStart, periodEnd], inclusive days.
    const from = Math.max(startOfDay_(s).getTime(), pS.getTime());
    const to = Math.min(startOfDay_(e).getTime(), startOfDay_(pE).getTime());
    if (to < from) continue;
    const days = Math.round((to - from) / 86400000) + 1;
    if (!out[id]) out[id] = { days: 0, ranges: [] };
    out[id].days += days;
    out[id].ranges.push(formatDdMmYyyy_(s) + '→' + formatDdMmYyyy_(e));
  }
  return out;
}

// Build the review payload for the screen.
// params: { periodStart?, periodEnd (dd-mm-yyyy), excludeIds? }
function getPayrollReview(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const pe = parseFlexibleDate_(params.periodEnd);
    if (!pe) return { success: false, error: 'periodEnd required (dd-mm-yyyy)' };
    const periodEnd = endOfDay_(pe);
    // Default start = day after the last run's end (run-to-run); else 1st of month.
    let periodStart;
    if (params.periodStart) { const ps = parseFlexibleDate_(params.periodStart); if (!ps) return { success: false, error: 'bad periodStart' }; periodStart = startOfDay_(ps); }
    else {
      const last = getLastPayrollEnd_(ss);
      periodStart = last ? new Date(last.getTime() + 86400000) : new Date(pe.getFullYear(), pe.getMonth(), 1);
      periodStart = startOfDay_(periodStart);
    }

    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const H = hd[0];
    const col = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cId = col('Employee ID'), cName = col('Name (English)'), cBasic = col('Basic Salary'),
          cFood = col('Food Allowance'), cAccom = col('Accommodation'), cConv = col('Conveyance'),
          cBonusHr = col('Bonus'), cRate = col('Price Per Hour'), cStatus = col('Status');

    const otAll = getAllOtHours_(ss);
    const encash = getEncashmentsForPeriod_(ss, formatDdMmYyyy_(periodEnd));
    const unpaid = getUnpaidLeaveDaysForPeriod_(ss, periodStart, periodEnd);
    const absence = (typeof getAbsenceDeductionsForPeriod_ === 'function') ? getAbsenceDeductionsForPeriod_(ss, periodStart, periodEnd) : {};
    const scheduled = (typeof getScheduledDeductionsForPeriod_ === 'function') ? getScheduledDeductionsForPeriod_(ss, periodEnd) : {};
    const excludeIds = String(params.excludeIds || '').split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

    const employees = [];
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][cId] || '').trim();
      if (!id) continue;
      if (excludeIds.indexOf(id.toUpperCase()) !== -1) continue;
      const statusVal = cStatus !== -1 ? String(hd[i][cStatus] || '').trim().toLowerCase() : '';
      if (statusVal && statusVal !== 'active') continue; // only Active paid

      const basic = Number(hd[i][cBasic]) || 0;
      const food = cFood !== -1 ? Number(hd[i][cFood]) || 0 : 0;
      const accom = cAccom !== -1 ? Number(hd[i][cAccom]) || 0 : 0;
      const conv = cConv !== -1 ? Number(hd[i][cConv]) || 0 : 0;
      const bonusHr = cBonusHr !== -1 ? Number(hd[i][cBonusHr]) || 0 : 0;
      const rate = cRate !== -1 ? Number(hd[i][cRate]) || 0 : 0;
      const gross = KD(basic + food + accom + conv + bonusHr);
      const otHours = Math.round((otAll[id] || 0) * 100) / 100;
      const otAmount = KD((otAll[id] || 0) * rate);
      const encDays = encash[id] ? (Number(encash[id].days) || 0) : 0;
      const encAmt = encash[id] ? KD(encash[id].amount) : 0;
      const unpaidDays = unpaid[id] ? unpaid[id].days : 0;
      const absentDays = absence[id] ? absence[id].absentDays : 0;
      const dailyRate = KD(basic / 26);
      const unpaidLeaveAmt = KD(dailyRate * unpaidDays);
      const absentAmt = absence[id] ? KD(absence[id].amount) : 0;
      const unpaidAmt = KD(unpaidLeaveAmt + absentAmt);
      const unpaidNoteParts = [];
      if (unpaid[id]) unpaidNoteParts.push('Unpaid leave ' + unpaidDays + 'd (' + unpaid[id].ranges.join(', ') + ')');
      if (absentDays) unpaidNoteParts.push('Absent ' + absentDays + 'd');

      employees.push({
        id: id, name: cName !== -1 ? String(hd[i][cName] || '').trim() : id,
        basic: basic, gross: gross, rate: rate,
        otHours: otHours, otAmount: otAmount,
        encDays: encDays, encAmt: encAmt,
        unpaidDays: unpaidDays + absentDays, unpaidAmt: unpaidAmt,
        unpaidNote: unpaidNoteParts.join('; '),
        // editable deductions/additions start empty (except auto unpaid)
        bonus: 0, loan: 0, advance: 0, disciplinary: 0,
        scheduled: (scheduled[id] ? KD(scheduled[id].total) : 0),
        scheduledItems: (scheduled[id] ? scheduled[id].items : []),
        net: KD(gross + otAmount + encAmt - unpaidAmt - (scheduled[id] ? scheduled[id].total : 0))
      });
    }

    return {
      success: true,
      periodStart: formatDdMmYyyy_(periodStart), periodEnd: formatDdMmYyyy_(periodEnd),
      count: employees.length, employees: employees
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Last payroll period end from Payroll_Control (for run-to-run default start).
function getLastPayrollEnd_(ss) {
  const ctrl = ss.getSheetByName(PAYROLL_CONTROL_TAB);
  if (!ctrl) return null;
  const d = ctrl.getDataRange().getValues();
  if (d.length < 2) return null;
  const H = d[0];
  const cEnd = H.findIndex(h => String(h).trim().toLowerCase().indexOf('period end') !== -1);
  if (cEnd === -1) return null;
  let latest = null;
  for (let i = 1; i < d.length; i++) {
    const e = parseFlexibleDate_(d[i][cEnd]);
    if (e && (!latest || e > latest)) latest = e;
  }
  return latest;
}

// Run payroll from the review screen's edited rows.
// params: { periodStart, periodEnd, rows: [{id, bonus, loan, advance, disciplinary, unpaidAmt, remark}], archiveOt? (default true) }
function runPayrollFromReview(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const ps = parseFlexibleDate_(params.periodStart), pe = parseFlexibleDate_(params.periodEnd);
    if (!pe) return { success: false, error: 'periodEnd required' };
    const periodStart = ps ? startOfDay_(ps) : null;
    const periodEnd = endOfDay_(pe);
    const rowsIn = params.rows || [];
    if (!rowsIn.length) return { success: false, error: 'No employee rows submitted.' };

    // Re-derive base figures server-side (never trust client for base pay).
    const review = getPayrollReview({ periodStart: params.periodStart, periodEnd: params.periodEnd });
    if (!review.success) return review;
    const base = {}; review.employees.forEach(e => base[e.id] = e);

    const editById = {};
    rowsIn.forEach(r => { editById[String(r.id).trim()] = r; });

    const headers = ['Employee ID','Name (English)','Basic Salary','Gross Salary','OT Hours','OT Amount',
      'Encashment','Bonus/Additional','Unpaid Leave','Loan','Advance','Disciplinary','Total Deductions',
      'Net Payable','Remark'];
    const rows = [headers];
    let tGross=0,tOt=0,tEnc=0,tBon=0,tUnp=0,tLoan=0,tAdv=0,tDis=0,tDed=0,tNet=0,tOtHours=0;
    const schedRowsToApply = []; // Scheduled_Deductions sheet row indexes consumed

    review.employees.forEach(e => {
      const ed = editById[e.id] || {};
      const bonus = KD(ed.bonus);
      const loan = KD(ed.loan);
      const advance = KD(ed.advance);
      const disciplinary = KD(ed.disciplinary);
      const unpaidAmt = KD(e.unpaidAmt); // server-derived, itemised
      const scheduledAmt = KD(e.scheduled || 0); // pre-logged deductions due this month
      const totalDed = KD(unpaidAmt + loan + advance + disciplinary + scheduledAmt);
      const net = KD(e.gross + e.otAmount + e.encAmt + bonus - totalDed);
      // Remark: fold in scheduled-deduction descriptions.
      const schedNote = (e.scheduledItems && e.scheduledItems.length)
        ? e.scheduledItems.map(it => it.type + ' ' + KD(it.amount) + (it.remark ? ' (' + it.remark + ')' : '')).join('; ') : '';
      const remark = [String(ed.remark || '').trim(),
        (e.unpaidDays ? 'Unpaid leave: ' + e.unpaidDays + ' day(s)' : ''),
        schedNote].filter(Boolean).join(' | ');
      if (e.scheduledItems) e.scheduledItems.forEach(it => { if (it.rowIndex) schedRowsToApply.push(it.rowIndex); });

      tGross+=e.gross; tOt+=e.otAmount; tEnc+=e.encAmt; tBon+=bonus; tUnp+=unpaidAmt;
      tLoan+=loan; tAdv+=advance; tDis+=(disciplinary+scheduledAmt); tDed+=totalDed; tNet+=net; tOtHours+=e.otHours;

      rows.push([e.id, e.name, KD(e.basic), e.gross, e.otHours, e.otAmount, e.encAmt,
        bonus, unpaidAmt, loan, advance, KD(disciplinary + scheduledAmt), totalDed, net, remark]);
    });
    rows.push(['', 'TOTALS (' + (rows.length - 1) + ' staff)', '', KD(tGross), Math.round(tOtHours*100)/100, KD(tOt),
      KD(tEnc), KD(tBon), KD(tUnp), KD(tLoan), KD(tAdv), KD(tDis), KD(tDed), KD(tNet), '']);

    const tabName = 'Payroll ' + Utilities.formatDate(periodStart || pe, tz_(), 'ddMMM') + '-' + Utilities.formatDate(pe, tz_(), 'ddMMM');
    const out = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
    out.clear();
    out.getRange(1, 1, rows.length, headers.length).setValues(rows);
    out.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    out.setFrozenRows(1);
    out.getRange(2, 3, rows.length - 1, headers.length - 2).setNumberFormat('0.000');

    // Log the run (reuse the engine's logger).
    logPayrollRun_(ss, periodStart || startOfDay_(pe), periodEnd, rows.length - 2, tOtHours, tabName, 'review');

    // Lock the scheduled deductions consumed by this run so they never repeat.
    if (typeof markScheduledDeductionsApplied_ === 'function' && schedRowsToApply.length) {
      markScheduledDeductionsApplied_(ss, schedRowsToApply, tabName);
    }

    // Archive OT_Entries -> fresh tab (the whole point: no OT ever double-counted or dropped).
    let archiveMsg = 'OT_Entries kept (archive skipped).';
    if (params.archiveOt !== false && params.archiveOt !== 'false') {
      archiveMsg = archiveOtEntries_(ss, tabName);
    }

    return {
      success: true, tab: tabName,
      periodStart: formatDdMmYyyy_(periodStart || startOfDay_(pe)), periodEnd: formatDdMmYyyy_(periodEnd),
      staff: rows.length - 2, totalNet: KD(tNet), totalOtHours: Math.round(tOtHours*100)/100,
      archive: archiveMsg,
      message: 'Ran payroll "' + tabName + '": ' + (rows.length - 2) + ' staff, net ' + KD(tNet).toFixed(3) + ' KD. ' + archiveMsg
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Rename current OT_Entries to an archive tab and create a fresh empty one.
function archiveOtEntries_(ss, periodLabel) {
  const sheet = ss.getSheetByName(OT_ENTRIES_TAB);
  if (!sheet) return 'No OT_Entries tab to archive.';
  const stamp = periodLabel.replace(/^Payroll\s*/, '');
  let archiveName = 'OT_Entries ' + stamp;
  let n = 2;
  while (ss.getSheetByName(archiveName)) { archiveName = 'OT_Entries ' + stamp + ' (' + n + ')'; n++; }
  sheet.setName(archiveName);
  // Fresh OT_Entries with the same header row.
  const fresh = ss.insertSheet(OT_ENTRIES_TAB);
  fresh.appendRow(['Date', 'Employee ID', 'Name', 'OT Hours', 'Remark', 'Added By']);
  fresh.getRange(1, 1, 1, 6).setFontWeight('bold');
  fresh.setFrozenRows(1);
  return 'Archived OT to "' + archiveName + '" and started a fresh OT_Entries.';
}