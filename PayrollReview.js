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
// Scheduled extras (22-09-2026): Pending rows in Scheduled_Extras for the
// period-end month are added to Bonus/Additional and marked Applied by the
// run — the same way Scheduled_Deductions are taken. Nothing is typed into
// the code; every bonus is a row in that tab.
//
// Encashment "(next run)" rows are NOT paid by the run — encashment is on
// hold until further notice (Shakeel, 22-09-2026). prEncashDue_ is kept, unused,
// for when it is switched back on.
//
// ENDPOINTS: getPayrollReview, runPayrollFromReview
// ============================================================

// Encashment due in this run: rows for this exact Period End, plus rows
// still marked "(next run)". Returns { empId: {days, amount, types, rows:[rowIndex]} }
function prEncashDue_(ss, periodEndKey) {
  const out = {};
  const sheet = ss.getSheetByName(ENCASHMENT_LOG_TAB);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const data = sheet.getDataRange().getValues();
  const H = data[0];
  const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const cId = c('Employee ID'), cType = c('Type'), cDays = c('Days'), cAmt = c('Amount'), cPer = c('Period End');
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][cId] || '').trim();
    if (!id) continue;
    const per = data[i][cPer];
    const due = formatKey_(per) === periodEndKey || /next\s*run/i.test(String(per || ''));
    if (!due) continue;
    if (!out[id]) out[id] = { days: 0, amount: 0, types: [], rows: [] };
    out[id].days += Number(data[i][cDays]) || 0;
    out[id].amount += Number(data[i][cAmt]) || 0;
    out[id].types.push(String(data[i][cType] || '').trim());
    if (/next\s*run/i.test(String(per || ''))) out[id].rows.push(i + 1);
  }
  Object.keys(out).forEach(id => { out[id].amount = KD(out[id].amount); });
  return out;
}
function prStampEncash_(ss, rowIndexes, periodEndKey) {
  const sheet = ss.getSheetByName(ENCASHMENT_LOG_TAB);
  if (!sheet || !rowIndexes.length) return;
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cPer = H.findIndex(h => String(h).trim().toLowerCase() === 'period end');
  if (cPer === -1) return;
  rowIndexes.forEach(r => { const cell = sheet.getRange(r, cPer + 1); cell.setNumberFormat('@'); cell.setValue(periodEndKey); });
}

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
    // Encashment: only rows already attached to this exact Period End.
    // "(next run)" rows are NOT paid — encashment is on hold (Shakeel, 22-09-2026).
    const encash = getEncashmentsForPeriod_(ss, formatDdMmYyyy_(periodEnd));
    const extras = (typeof getScheduledExtrasForPeriod_ === 'function') ? getScheduledExtrasForPeriod_(ss, periodEnd) : {};
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

      const extraAmt = extras[id] ? KD(extras[id].total) : 0;
      employees.push({
        id: id, name: cName !== -1 ? String(hd[i][cName] || '').trim() : id,
        basic: basic, food: food, accom: accom, conv: conv, bonusHr: bonusHr, gross: gross, rate: rate,
        otHours: otHours, otAmount: otAmount,
        encDays: encDays, encAmt: encAmt,
        unpaidDays: unpaidDays + absentDays, unpaidAmt: unpaidAmt,
        unpaidNote: unpaidNoteParts.join('; '),
        // editable deductions/additions start empty (except auto unpaid)
        bonus: 0, loan: 0, advance: 0, disciplinary: 0,
        scheduled: (scheduled[id] ? KD(scheduled[id].total) : 0),
        scheduledItems: (scheduled[id] ? scheduled[id].items : []),
        extra: extraAmt,
        extraItems: (extras[id] ? extras[id].items : []),
        encashRows: [],
        net: KD(gross + otAmount + encAmt + extraAmt - unpaidAmt - (scheduled[id] ? scheduled[id].total : 0))
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
    const periodEnd = endOfDay_(pe);
    const rowsIn = params.rows || [];
    if (!rowsIn.length) return { success: false, error: 'No employee rows submitted.' };

    // Re-derive base figures server-side (never trust client for base pay).
    const review = getPayrollReview({ periodStart: params.periodStart, periodEnd: params.periodEnd });
    if (!review.success) return review;
    const base = {}; review.employees.forEach(e => base[e.id] = e);
    // No start sent → the same start the review used (day after the last run).
    const periodStart = ps ? startOfDay_(ps) : startOfDay_(parseFlexibleDate_(review.periodStart));

    const editById = {};
    rowsIn.forEach(r => { editById[String(r.id).trim()] = r; });

    // Pay components itemised so the payslip shows every allowance, and the
    // scheduled deductions (loan, fine, disciplinary …) in their own column
    // instead of being folded into Disciplinary.
    const headers = ['Employee ID','Name (English)','Basic Salary','Food Allowance','Accommodation','Conveyance','Bonus',
      'Gross Salary','OT Hours','Price Per Hour','OT Amount',
      'Encashment','Bonus/Additional','Unpaid Leave','Loan','Advance','Disciplinary','Scheduled Deductions','Total Deductions',
      'Net Payable','Remark'];
    const rows = [headers];
    let tGross=0,tOt=0,tEnc=0,tBon=0,tUnp=0,tLoan=0,tAdv=0,tDis=0,tSch=0,tDed=0,tNet=0,tOtHours=0;
    let tBasic=0,tFood=0,tAccom=0,tConv=0,tBonHr=0;
    const schedRowsToApply = []; // Scheduled_Deductions sheet row indexes consumed
    const extraRowsToApply = []; // Scheduled_Extras sheet row indexes paid
    const encashRowsToStamp = []; // Encashment_Log "(next run)" rows paid

    review.employees.forEach(e => {
      const ed = editById[e.id] || {};
      const extraAmt = KD(e.extra || 0);             // scheduled extras for this month
      const bonus = KD(KD(ed.bonus) + extraAmt);     // typed on the screen + scheduled
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
      const extraNote = (e.extraItems && e.extraItems.length)
        ? e.extraItems.map(it => '+' + it.type + ' ' + KD(it.amount) + (it.remark ? ' (' + it.remark + ')' : '')).join('; ') : '';
      const remark = [String(ed.remark || '').trim(),
        (e.unpaidDays ? 'Unpaid leave: ' + e.unpaidDays + ' day(s)' : ''),
        schedNote, extraNote].filter(Boolean).join(' | ');
      if (e.scheduledItems) e.scheduledItems.forEach(it => { if (it.rowIndex) schedRowsToApply.push(it.rowIndex); });
      if (e.extraItems) e.extraItems.forEach(it => { if (it.rowIndex) extraRowsToApply.push(it.rowIndex); });
      (e.encashRows || []).forEach(r => encashRowsToStamp.push(r));

      tGross+=e.gross; tOt+=e.otAmount; tEnc+=e.encAmt; tBon+=bonus; tUnp+=unpaidAmt;
      tLoan+=loan; tAdv+=advance; tDis+=disciplinary; tSch+=scheduledAmt; tDed+=totalDed; tNet+=net; tOtHours+=e.otHours;
      tBasic+=e.basic; tFood+=e.food||0; tAccom+=e.accom||0; tConv+=e.conv||0; tBonHr+=e.bonusHr||0;

      rows.push([e.id, e.name, KD(e.basic), KD(e.food||0), KD(e.accom||0), KD(e.conv||0), KD(e.bonusHr||0),
        e.gross, e.otHours, KD(e.rate), e.otAmount, e.encAmt,
        bonus, unpaidAmt, loan, advance, disciplinary, scheduledAmt, totalDed, net, remark]);
    });
    rows.push(['', 'TOTALS (' + (rows.length - 1) + ' staff)', KD(tBasic), KD(tFood), KD(tAccom), KD(tConv), KD(tBonHr),
      KD(tGross), Math.round(tOtHours*100)/100, '', KD(tOt),
      KD(tEnc), KD(tBon), KD(tUnp), KD(tLoan), KD(tAdv), KD(tDis), KD(tSch), KD(tDed), KD(tNet), '']);

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
    if (typeof markScheduledExtrasApplied_ === 'function' && extraRowsToApply.length) {
      markScheduledExtrasApplied_(ss, extraRowsToApply, tabName);
    }
    prStampEncash_(ss, encashRowsToStamp, formatDdMmYyyy_(periodEnd));

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
// ============================================================
// LATE ADDITIONS TO A RUN — previewLateScheduledItems() / commitLateScheduledItems()
// ------------------------------------------------------------
// A scheduled deduction or extra entered AFTER the payroll ran (still
// Pending, same month as the run) is added to that run's tab without
// re-running: Scheduled Deductions / Bonus/Additional, Total Deductions,
// Net Payable, Remark and the TOTALS row are updated, and the rows are
// marked Applied. Re-running is not safe once a run has archived
// OT_Entries and marked its deductions and extras Applied.
// Only for the LATEST run, and only before the bank sheet and payslips
// are generated (or generate them again afterwards).
// params: { tab? } — defaults to the latest run in Payroll_Control.
// ============================================================
function previewLateScheduledItems(params) { return prLate_(params || {}, false); }
function commitLateScheduledItems(params)  { return prLate_(params || {}, true); }

function prLate_(params, commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const out = [];
  const last = getLastPayrollRun_(ss);
  if (!last) return prLog_(['No payroll run found in Payroll_Control.']);
  const tabName = String(params.tab || last.tabName).trim();
  if (tabName !== last.tabName) return prLog_(['Only the latest run (' + last.tabName + ') can take late additions.']);
  const sh = ss.getSheetByName(tabName);
  if (!sh) return prLog_(['Tab "' + tabName + '" not found.']);
  const data = sh.getDataRange().getValues();
  const H = data[0].map(h => String(h).trim().toLowerCase());
  const c = n => H.indexOf(n.toLowerCase());
  const C = { id: c('employee id'), sched: c('scheduled deductions'), bonus: c('bonus/additional'),
              ded: c('total deductions'), net: c('net payable'), rem: c('remark') };
  if (C.id < 0 || C.ded < 0 || C.net < 0) return prLog_(['Tab "' + tabName + '" is not a Payroll Review tab.']);

  const ded = (typeof getScheduledDeductionsForPeriod_ === 'function') ? getScheduledDeductionsForPeriod_(ss, last.periodEnd) : {};
  const ext = (typeof getScheduledExtrasForPeriod_ === 'function') ? getScheduledExtrasForPeriod_(ss, last.periodEnd) : {};
  const rowOf = {}; let totRow = -1;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][C.id] || '').trim();
    if (id) rowOf[id] = i; else if (String(data[i][1] || '').indexOf('TOTALS') === 0) totRow = i;
  }
  const changes = [], dedRows = [], extRows = [], notInRun = [];
  const touch = (id) => { if (rowOf[id] == null) { notInRun.push(id); return null; } return data[rowOf[id]]; };
  let dSum = 0, eSum = 0;
  Object.keys(ded).forEach(id => {
    const r = touch(id); if (!r) return;
    const amt = KD(ded[id].total);
    if (C.sched >= 0) r[C.sched] = KD((Number(r[C.sched]) || 0) + amt);
    r[C.ded] = KD((Number(r[C.ded]) || 0) + amt);
    r[C.net] = KD((Number(r[C.net]) || 0) - amt);
    const note = ded[id].items.map(it => it.type + ' ' + KD(it.amount) + (it.remark ? ' (' + it.remark + ')' : '')).join('; ');
    if (C.rem >= 0) r[C.rem] = [String(r[C.rem] || '').trim(), note].filter(Boolean).join(' | ');
    ded[id].items.forEach(it => dedRows.push(it.rowIndex));
    changes.push(id + '  −' + amt.toFixed(3) + '  ' + note + '  → net ' + Number(r[C.net]).toFixed(3));
    dSum += amt;
  });
  Object.keys(ext).forEach(id => {
    const r = touch(id); if (!r) return;
    const amt = KD(ext[id].total);
    if (C.bonus >= 0) r[C.bonus] = KD((Number(r[C.bonus]) || 0) + amt);
    r[C.net] = KD((Number(r[C.net]) || 0) + amt);
    const note = ext[id].items.map(it => '+' + it.type + ' ' + KD(it.amount) + (it.remark ? ' (' + it.remark + ')' : '')).join('; ');
    if (C.rem >= 0) r[C.rem] = [String(r[C.rem] || '').trim(), note].filter(Boolean).join(' | ');
    ext[id].items.forEach(it => extRows.push(it.rowIndex));
    changes.push(id + '  +' + amt.toFixed(3) + '  ' + note + '  → net ' + Number(r[C.net]).toFixed(3));
    eSum += amt;
  });
  if (totRow >= 0) {
    const t = data[totRow];
    if (C.sched >= 0) t[C.sched] = KD((Number(t[C.sched]) || 0) + dSum);
    if (C.bonus >= 0) t[C.bonus] = KD((Number(t[C.bonus]) || 0) + eSum);
    t[C.ded] = KD((Number(t[C.ded]) || 0) + dSum);
    t[C.net] = KD((Number(t[C.net]) || 0) - dSum + eSum);
  }

  out.push((commit ? 'DONE' : 'PREVIEW — nothing written yet') + ' — late scheduled items for "' + tabName + '" (' +
           formatDdMmYyyy_(last.periodEnd) + ')');
  out.push('Deductions: ' + KD(dSum).toFixed(3) + ' KD   Extras: ' + KD(eSum).toFixed(3) + ' KD');
  changes.forEach(x => out.push('   ' + x));
  if (notInRun.length) out.push('NOT IN THIS RUN (left Pending): ' + notInRun.join(', '));
  if (totRow >= 0) out.push('New total net: ' + Number(data[totRow][C.net]).toFixed(3) + ' KD');
  if (!changes.length) { out.push('Nothing Pending for this month.'); return prLog_(out); }
  if (!commit) { out.push(''); out.push('Nothing has been changed. If this is right, run commitLateScheduledItems().'); return prLog_(out); }

  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  if (dedRows.length && typeof markScheduledDeductionsApplied_ === 'function') markScheduledDeductionsApplied_(ss, dedRows, tabName);
  if (extRows.length && typeof markScheduledExtrasApplied_ === 'function') markScheduledExtrasApplied_(ss, extRows, tabName);
  SpreadsheetApp.flush();
  out.push('');
  out.push('Written to "' + tabName + '". Now generate the bank sheet and payslips.');
  return prLog_(out);
}
function prLog_(lines) { const t = lines.join('\n'); Logger.log(t); return t; }

// ============================================================
// STAFF MISSED BY A RUN — previewAddMissedStaff() / commitAddMissedStaff()
// ------------------------------------------------------------
// An Active employee in HR Maya who is not in the latest run's tab (e.g.
// added to HR Maya after the run) is appended to that tab, calculated the
// same way the run calculates everyone:
//   Gross   = Basic + Food + Accommodation + Conveyance + Bonus (HR Maya)
//   OT      = hours in the run's archived "OT_Entries <period>" tab × Price Per Hour
//   Unpaid  = unpaid leave (Leave_Grants) + absences, for the run's period
// Scheduled deductions and extras are NOT taken here — run
// previewLateScheduledItems / commitLateScheduledItems afterwards; they
// pick up the new row like any other.
// The TOTALS row is recalculated from all staff rows, and the Employees
// count in Payroll_Control is updated.
// Nothing is typed into the code: who is missing is read from the sheets.
// params: { tab?, ids? ("MT-00070,MT-00071" — limit to these) }
// ============================================================
function previewAddMissedStaff(params) { return prAddMissed_(params || {}, false); }
function commitAddMissedStaff(params)  { return prAddMissed_(params || {}, true); }

function prAddMissed_(params, commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const out = [];
  const last = getLastPayrollRun_(ss);
  if (!last) return prLog_(['No payroll run found in Payroll_Control.']);
  const tabName = String(params.tab || last.tabName).trim();
  if (tabName !== last.tabName) return prLog_(['Only the latest run (' + last.tabName + ') can take added staff.']);
  const sh = ss.getSheetByName(tabName);
  if (!sh) return prLog_(['Tab "' + tabName + '" not found.']);

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const W = headers.length;
  const tc = n => headers.findIndex(h => h.toLowerCase() === n.toLowerCase());
  const T = { id: tc('Employee ID'), name: tc('Name (English)'), basic: tc('Basic Salary'), food: tc('Food Allowance'),
              accom: tc('Accommodation'), conv: tc('Conveyance'), bonusHr: tc('Bonus'), gross: tc('Gross Salary'),
              otH: tc('OT Hours'), rate: tc('Price Per Hour'), otAmt: tc('OT Amount'), unpaid: tc('Unpaid Leave'),
              ded: tc('Total Deductions'), net: tc('Net Payable'), rem: tc('Remark') };
  if (T.id < 0 || T.net < 0 || T.ded < 0) return prLog_(['Tab "' + tabName + '" is not a Payroll Review tab.']);

  // Staff rows already in the run, and where TOTALS sits.
  const inRun = {}; let totRow = -1;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][T.id] || '').trim();
    if (id) inRun[id] = true;
    else if (String(data[i][T.name >= 0 ? T.name : 1] || '').indexOf('TOTALS') === 0) totRow = i;
  }

  // HR Maya — Active staff not in the run.
  const hd = ss.getSheetByName('HR Maya').getDataRange().getValues();
  const hH = hd[0].map(h => String(h).trim().toLowerCase());
  const hc = n => hH.indexOf(n.toLowerCase());
  const H = { id: hc('employee id'), name: hc('name (english)'), basic: hc('basic salary'), food: hc('food allowance'),
              accom: hc('accommodation'), conv: hc('conveyance'), bonusHr: hc('bonus'), rate: hc('price per hour'),
              status: hc('status') };
  const only = String(params.ids || '').split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
  const num = (r, k) => (H[k] >= 0 ? Number(r[H[k]]) || 0 : 0);

  // The run's own OT (archived when it ran) and its period.
  const stamp = tabName.replace(/^Payroll\s*/, '');
  const otTab = ss.getSheetByName('OT_Entries ' + stamp);
  const otHours = {};
  if (otTab && otTab.getLastRow() > 1) {
    const od = otTab.getDataRange().getValues();
    const oH = od[0].map(h => String(h).trim().toLowerCase());
    const oId = oH.indexOf('employee id'), oHr = oH.indexOf('ot hours');
    for (let i = 1; i < od.length; i++) {
      const id = String(od[i][oId] || '').trim(), h = Number(od[i][oHr]) || 0;
      if (id && h > 0) otHours[id] = (otHours[id] || 0) + h;
    }
  }
  const pS = startOfDay_(last.periodStart || last.periodEnd), pE = endOfDay_(last.periodEnd);
  const unpaid = getUnpaidLeaveDaysForPeriod_(ss, pS, pE);
  const absence = (typeof getAbsenceDeductionsForPeriod_ === 'function') ? getAbsenceDeductionsForPeriod_(ss, pS, pE) : {};

  const added = [];
  for (let i = 1; i < hd.length; i++) {
    const r = hd[i];
    const id = String(r[H.id] || '').trim();
    if (!id || inRun[id]) continue;
    const st = H.status >= 0 ? String(r[H.status] || '').trim().toLowerCase() : 'active';
    if (st && st !== 'active') continue;
    if (only.length && only.indexOf(id.toUpperCase()) === -1) continue;

    const basic = num(r, 'basic'), food = num(r, 'food'), accom = num(r, 'accom'), conv = num(r, 'conv'),
          bonusHr = num(r, 'bonusHr'), rate = num(r, 'rate');
    const gross = KD(basic + food + accom + conv + bonusHr);
    const hrs = Math.round((otHours[id] || 0) * 100) / 100;
    const otAmt = KD((otHours[id] || 0) * rate);
    const uDays = unpaid[id] ? unpaid[id].days : 0;
    const aDays = absence[id] ? absence[id].absentDays : 0;
    const unpaidAmt = KD(KD(basic / 26) * uDays + (absence[id] ? absence[id].amount : 0));
    const net = KD(gross + otAmt - unpaidAmt);
    const note = ['Added after the run (' + formatDdMmYyyy_(new Date()) + ')',
                  uDays ? 'Unpaid leave: ' + uDays + ' day(s)' : '', aDays ? 'Absent ' + aDays + 'd' : ''].filter(Boolean).join(' | ');

    const row = new Array(W).fill(0);
    const set = (k, v) => { if (T[k] >= 0) row[T[k]] = v; };
    set('id', id); set('name', H.name >= 0 ? String(r[H.name] || '').trim() : id);
    set('basic', KD(basic)); set('food', KD(food)); set('accom', KD(accom)); set('conv', KD(conv)); set('bonusHr', KD(bonusHr));
    set('gross', gross); set('otH', hrs); set('rate', KD(rate)); set('otAmt', otAmt);
    set('unpaid', unpaidAmt); set('ded', unpaidAmt); set('net', net); set('rem', note);
    added.push({ row: row, text: id + '  ' + row[T.name] + '  gross ' + gross.toFixed(3) + '  OT ' + hrs + 'h ' +
                 otAmt.toFixed(3) + '  unpaid ' + unpaidAmt.toFixed(3) + '  → net ' + net.toFixed(3) });
  }

  out.push((commit ? 'DONE' : 'PREVIEW — nothing written yet') + ' — staff missing from "' + tabName + '" (' +
           formatDdMmYyyy_(pS) + ' → ' + formatDdMmYyyy_(pE) + ')');
  if (!otTab) out.push('⚠ Archived OT tab "OT_Entries ' + stamp + '" not found — OT taken as 0.');
  if (!added.length) { out.push('No Active staff are missing from this run.'); return prLog_(out); }
  added.forEach(a => out.push('   ' + a.text));

  // New layout: staff rows (old + added), then TOTALS recalculated.
  const staff = data.slice(1).filter(r => String(r[T.id] || '').trim()).concat(added.map(a => a.row));
  const tot = new Array(W).fill('');
  const skip = [T.id, T.name, T.rate, T.rem];
  for (let j = 0; j < W; j++) {
    if (skip.indexOf(j) !== -1) continue;
    tot[j] = j === T.otH ? Math.round(staff.reduce((s, r) => s + (Number(r[j]) || 0), 0) * 100) / 100
                         : KD(staff.reduce((s, r) => s + (Number(r[j]) || 0), 0));
  }
  tot[T.name >= 0 ? T.name : 1] = 'TOTALS (' + staff.length + ' staff)';
  const oldNet = totRow >= 0 ? Number(data[totRow][T.net]) || 0 : 0;
  out.push('Staff: ' + (staff.length - added.length) + ' → ' + staff.length +
           '   Net total: ' + oldNet.toFixed(3) + ' → ' + Number(tot[T.net]).toFixed(3) + ' KD');
  if (!commit) { out.push(''); out.push('Nothing has been changed. If this is right, run commitAddMissedStaff().'); return prLog_(out); }

  const newData = [data[0]].concat(staff).concat([tot]);
  sh.getRange(1, 1, data.length, W).clearContent();
  sh.getRange(1, 1, newData.length, W).setValues(newData);
  sh.getRange(2, 3, newData.length - 1, W - 2).setNumberFormat('0.000');
  sh.getRange(newData.length, 1, 1, W).setFontWeight('bold');

  // Payroll_Control: Employees count for this run.
  const ctrl = ss.getSheetByName(PAYROLL_CONTROL_TAB);
  if (ctrl) {
    const cd = ctrl.getDataRange().getValues();
    const cH = cd[0].map(h => String(h).trim().toLowerCase());
    const cTab = cH.indexOf('tab name'), cEmp = cH.indexOf('employees');
    for (let i = cd.length - 1; i >= 1; i--) {
      if (String(cd[i][cTab]).trim() === tabName) { if (cEmp >= 0) ctrl.getRange(i + 1, cEmp + 1).setValue(staff.length); break; }
    }
  }
  SpreadsheetApp.flush();
  out.push('');
  out.push('Written to "' + tabName + '". Next: previewLateScheduledItems → commitLateScheduledItems, then the bank sheet and payslips.');
  return prLog_(out);
}