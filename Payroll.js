// ============================================================
// Payroll.gs — Maya HR : rolling-period payroll (attendance / manual modes)
//              + attendance amendment + manual inputs
// ============================================================
// Standalone Apps Script file. Shares global scope with Code.gs and
// SharedUtils.gs (helpers: tz_, formatDdMmYyyy_, formatDdMmYyyyHm_,
// parseFlexibleDate_, startOfDay_, endOfDay_, addDays_, sameCalendarDate_).
// Do NOT redeclare MAIN_SPREADSHEET_ID or those helpers here.
//
// WHAT'S NEW vs the previous rolling-period version:
//  A. MODE. computeMonthlyPayroll accepts mode:
//       'attendance' (default) -> OT computed from Check-In/Out cycles
//       'manual'               -> attendance skipped; OT hours read from
//                                 the 'Payroll_Inputs' tab per employee.
//  B. MANUAL INPUTS. A 'Payroll_Inputs' tab, keyed by Period End, holds
//     per-employee Manual OT Hours + Deductions + Remark. Deductions
//     apply in BOTH modes; Manual OT applies only in 'manual' mode.
//     preparePayrollInputs({periodEnd}) seeds one row per employee.
//  C. DEDUCTIONS SOURCE. Loan / Other Deductions on the HR Maya master
//     are now REFERENCE-ONLY and no longer subtracted. The single
//     'Deductions' figure from Payroll_Inputs is the only deduction.
//  D. EXCLUDE. excludeIds omits staff already paid off-cycle (advance/
//     vacation), so they aren't paid twice in the batch.
//  E. AMENDMENT SAFETY NET. amendAttendance flags any already-run period
//     in Payroll_Control as "NEEDS RE-RUN" when a punch inside it is
//     changed, so a corrected cycle can never silently drift.
// ============================================================

const DUTY_HOURS = 9; // 8 work + 1 break; OT accrues only above this
const KD = (n) => Math.round((Number(n) || 0) * 1000) / 1000; // round money to 3 dp (fils)
const PAYROLL_CONTROL_TAB = 'Payroll_Control';
const PAYROLL_INPUTS_TAB  = 'Payroll_Inputs';

// ============================================================
// COMPUTE PAYROLL
// params (rolling): { periodStart?:'dd-mm-yyyy', periodEnd:'dd-mm-yyyy',
//                     mode?:'attendance'|'manual', excludeIds?:'MT-1,MT-2' }
// params (legacy convenience): { month:1-12, year:2026, mode?, excludeIds? }
// ============================================================
function computeMonthlyPayroll(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    const mode = String(params.mode || 'attendance').trim().toLowerCase() === 'manual' ? 'manual' : 'attendance';
    const excludeIds = String(params.excludeIds || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

    // ---- 1. Resolve the period ----
    let periodStart, periodEnd, tabName, isRollingMode;

    if (params.periodEnd) {
      isRollingMode = true;
      const pe = parseFlexibleDate_(params.periodEnd);
      if (!pe) return { success: false, error: 'periodEnd must be a valid date (dd-mm-yyyy)' };
      periodEnd = endOfDay_(pe);

      // Pay cycle: salary paid by the 25th, so a run always covers the 1st of the
      // month through the chosen cutoff (21st-24th). Default the start to the 1st
      // of periodEnd's month, so picking an end date can never produce a sliver
      // window. An explicit periodStart still overrides (rare exceptions).
      if (params.periodStart) {
        const ps = parseFlexibleDate_(params.periodStart);
        if (!ps) return { success: false, error: 'periodStart must be a valid date (dd-mm-yyyy)' };
        periodStart = startOfDay_(ps);
      } else {
        periodStart = new Date(pe.getFullYear(), pe.getMonth(), 1, 0, 0, 0, 0); // 1st of the month
      }

      if (periodStart > periodEnd) {
        return { success: false, error: 'periodStart (' + formatDdMmYyyy_(periodStart) + ') is after periodEnd (' + formatDdMmYyyy_(periodEnd) + ').' };
      }

      // Backstop guard: refuse an abnormally short window unless forced, so a
      // wrong end date can't silently pay a fraction of the month's OT.
      const spanDays = Math.round((startOfDay_(periodEnd) - startOfDay_(periodStart)) / 86400000) + 1;
      const force = String(params.force || '').toLowerCase() === 'true';
      if (spanDays < 7 && !force) {
        return { success: false, error: 'This period is only ' + spanDays + ' day(s) (' + formatDdMmYyyy_(periodStart) + ' to ' + formatDdMmYyyy_(periodEnd) + '). A monthly run normally covers the 1st through the 21st-24th (~20+ days). If this really is intentional, re-run with force=true.' };
      }

      tabName = 'Payroll ' + Utilities.formatDate(periodStart, tz_(), 'ddMMM') + '-' + Utilities.formatDate(periodEnd, tz_(), 'ddMMM');
    } else {
      isRollingMode = false;
      const now = new Date();
      const month = parseInt(params.month || (now.getMonth() + 1), 10);
      const year = parseInt(params.year || now.getFullYear(), 10);
      periodStart = new Date(year, month - 1, 1, 0, 0, 0);
      periodEnd = new Date(year, month, 0, 23, 59, 59, 999);
      tabName = 'Payroll ' + String(month).padStart(2, '0') + '-' + year;
    }

    // ---- 2. Salary master from 'HR Maya' (read-only) ----
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found" };
    const hrData = hr.getDataRange().getValues();
    if (hrData.length <= 1) return { success: false, error: 'HR Maya has no employee rows' };

    const H = hrData[0];
    const col = (name) => H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
    const cId = col('Employee ID'), cName = col('Name (English)'), cBasic = col('Basic Salary'),
          cFood = col('Food Allowance'), cAccom = col('Accommodation'), cConv = col('Conveyance'),
          cBonus = col('Bonus'), cLoan = col('Loan'), cDed = col('Other Deductions'), cRate = col('Price Per Hour'), cStatus = col('Status');

    const need = { 'Employee ID': cId, 'Name (English)': cName, 'Basic Salary': cBasic, 'Price Per Hour': cRate };
    const missing = Object.keys(need).filter(k => need[k] === -1);
    if (missing.length) return { success: false, error: 'HR Maya missing column(s): ' + missing.join(', ') };

    const emp = {};
    const order = [];
    const skippedInactive = [];
    for (let i = 1; i < hrData.length; i++) {
      const id = String(hrData[i][cId] || '').trim();
      if (!id) continue;
      if (excludeIds.indexOf(id.toUpperCase()) !== -1) continue; // omit advance-paid / excluded staff
      // Only Active employees are paid. Blank status is treated as Active
      // (legacy rows); explicit 'Inactive' / 'On Leave' are skipped.
      const statusVal = cStatus !== -1 ? String(hrData[i][cStatus] || '').trim().toLowerCase() : '';
      if (statusVal && statusVal !== 'active') { skippedInactive.push({ id: id, status: String(hrData[i][cStatus]).trim() }); continue; }
      emp[id] = {
        id: id,
        name:  String(hrData[i][cName] || '').trim(),
        basic: Number(hrData[i][cBasic]) || 0,
        food:  cFood  !== -1 ? Number(hrData[i][cFood])  || 0 : 0,
        accom: cAccom !== -1 ? Number(hrData[i][cAccom]) || 0 : 0,
        conv:  cConv  !== -1 ? Number(hrData[i][cConv])  || 0 : 0,
        bonus: cBonus !== -1 ? Number(hrData[i][cBonus]) || 0 : 0,
        loanRef: cLoan !== -1 ? Number(hrData[i][cLoan]) || 0 : 0, // reference-only, not subtracted
        otherRef: cDed !== -1 ? Number(hrData[i][cDed]) || 0 : 0,  // reference-only, not subtracted
        rate:  Number(hrData[i][cRate]) || 0,
        otHours: 0, deductions: 0, remark: '', encashDays: 0, encashAmt: 0, flags: []
      };
      order.push(id);
    }

    // ---- 3a. Deductions + remark from Payroll_Inputs (both modes) ----
    const inputs = getPayrollInputs_(ss, formatDdMmYyyy_(periodEnd)); // { empId: {deductions, remark} }
    Object.keys(inputs).forEach(id => {
      if (!emp[id]) return;
      emp[id].deductions = KD(inputs[id].deductions);
      emp[id].remark = inputs[id].remark || '';
    });

    // Manual mode: OT hours come from the OT_Entries tab, summed over the period.
    if (mode === 'manual') {
      const otByEmp = getOtHoursForPeriod_(ss, periodStart, periodEnd); // { empId: totalHours }
      Object.keys(otByEmp).forEach(id => { if (emp[id]) emp[id].otHours = otByEmp[id]; });
    }

    // ---- 3a-ii. Leave encashment attached to this Period End (from Encashment_Log) ----
    const encash = getEncashmentsForPeriod_(ss, formatDdMmYyyy_(periodEnd));
    Object.keys(encash).forEach(id => {
      if (!emp[id]) return;
      emp[id].encashDays = Number(encash[id].days) || 0;
      emp[id].encashAmt = KD(encash[id].amount);
    });

    // ---- 3b. Attendance -> OT cycles (attendance mode only) ----
    if (mode === 'attendance') {
      const log = ss.getSheetByName('Attendance_Log');
      if (log) {
        const L = log.getDataRange().getValues();
        const byEmp = {};
        for (let i = 0; i < L.length; i++) {
          const empId = String(L[i][1] || '').trim();
          const stage = String(L[i][4] || '').trim();
          if (!empId || (stage !== 'Check In' && stage !== 'Check Out')) continue;
          const ts = parseLogTimestamp(L[i][0]);
          if (!ts) continue;
          (byEmp[empId] = byEmp[empId] || []).push({ ts, stage });
        }
        Object.keys(byEmp).forEach(empId => {
          if (!emp[empId]) return;
          const rows = byEmp[empId].sort((a, b) => a.ts - b.ts);
          const cycles = buildCycles_(rows);
          cycles.forEach(c => {
            if (!c.checkOut) { emp[empId].flags.push('Open shift since ' + formatDdMmYyyyHm_(c.checkIn) + ' — no Check Out yet'); return; }
            if (c.incomplete) { emp[empId].flags.push('Incomplete cycle: Check In ' + formatDdMmYyyyHm_(c.checkIn) + ' had no Check Out before the next Check In'); return; }
            if (c.checkOut < periodStart || c.checkOut > periodEnd) return; // outside this period; a later run picks it up
            const presence = (c.checkOut - c.checkIn) / 3600000;
            if (presence <= 0) { emp[empId].flags.push('Check Out not after Check In near ' + formatDdMmYyyyHm_(c.checkIn)); return; }
            const ot = presence - DUTY_HOURS;
            if (ot > 0) emp[empId].otHours += ot;
          });
        });
      }
    }

    // ---- 4. Build snapshot rows ----
    const headers = ['Employee ID','Name (English)','Basic Salary','Food Allowance','Accommodation',
      'Conveyance','Bonus','Gross Salary','OT Hours','Price Per Hour','OT Amount',
      'Encash Days','Encashment','Deductions','Deduction Remark','Total Payable','Flags'];
    const rows = [headers];
    let tGross = 0, tOtAmt = 0, tEnc = 0, tDed = 0, tPay = 0, totalOtHours = 0;

    order.forEach(id => {
      const e = emp[id];
      const gross    = KD(e.basic + e.food + e.accom + e.conv + e.bonus);
      const otHrsDisplay = Math.round(e.otHours * 100) / 100;
      const otAmount = KD(e.otHours * e.rate);
      const encAmt   = KD(e.encashAmt);
      const ded      = KD(e.deductions);
      const payable  = KD(gross + otAmount + encAmt - ded);
      totalOtHours += e.otHours;
      tGross += gross; tOtAmt += otAmount; tEnc += encAmt; tDed += ded; tPay += payable;
      rows.push([e.id, e.name, KD(e.basic), KD(e.food), KD(e.accom), KD(e.conv), KD(e.bonus),
        gross, otHrsDisplay, KD(e.rate), otAmount, (e.encashDays || ''), encAmt, ded, e.remark, payable,
        e.flags.length ? e.flags.join('; ') : '']);
    });

    // Totals row
    rows.push(['', 'TOTALS (' + (rows.length - 1) + ' staff)', '', '', '', '', '',
      KD(tGross), Math.round(totalOtHours * 100) / 100, '', KD(tOtAmt), '', KD(tEnc), KD(tDed), '', KD(tPay), '']);

    // ---- 5. Write / refresh the period tab ----
    let out = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
    out.clear();
    out.getRange(1, 1, rows.length, headers.length).setValues(rows);
    out.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    out.getRange(rows.length, 1, 1, headers.length).setFontWeight('bold').setBackground('#f2f4f7');
    if (rows.length > 2) {
      [3,4,5,6,7,8,10,11,13,14,16].forEach(c => out.getRange(2, c, rows.length - 1, 1).setNumberFormat('0.000'));
      out.getRange(2, 9, rows.length - 1, 1).setNumberFormat('0.00'); // OT Hours
    }
    out.setFrozenRows(1);
    out.autoResizeColumns(1, headers.length);

    // ---- 6. Log this run (rolling mode only) ----
    if (isRollingMode) {
      logPayrollRun_(ss, periodStart, periodEnd, order.length, totalOtHours, tabName, mode);
    }

    const flagList = order.filter(id => emp[id].flags.length)
      .map(id => ({ id: emp[id].id, name: emp[id].name, detail: emp[id].flags.join('; ') }));
    return {
      success: true, tab: tabName, mode: mode,
      periodStart: formatDdMmYyyy_(periodStart), periodEnd: formatDdMmYyyy_(periodEnd),
      employees: order.length, excluded: excludeIds,
      skippedInactive: skippedInactive,
      totalPayable: KD(tPay), totalDeductions: KD(tDed), totalOtAmount: KD(tOtAmt), totalEncashment: KD(tEnc),
      flaggedEmployees: flagList.length, flags: flagList,
      message: 'Wrote "' + tabName + '" (' + formatDdMmYyyy_(periodStart) + ' to ' + formatDdMmYyyy_(periodEnd) + ', ' + mode + ' mode): ' +
        order.length + ' employees, ' + flagList.length + ' with attendance flags' +
        (skippedInactive.length ? ', ' + skippedInactive.length + ' inactive skipped.' : '.')
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// PAYROLL_INPUTS — seed + read manual OT / deductions per period
// Columns: Employee ID | Name (English) | Period End | Manual OT Hours | Deductions | Remark
// ============================================================
function preparePayrollInputs(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const pe = parseFlexibleDate_(params.periodEnd);
    if (!pe) return { success: false, error: 'periodEnd must be a valid date (dd-mm-yyyy)' };
    const periodKey = formatDdMmYyyy_(pe);

    const excludeIds = String(params.excludeIds || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found" };
    const hrData = hr.getDataRange().getValues();
    const H = hrData[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    if (cId === -1) return { success: false, error: 'HR Maya missing Employee ID column' };

    let sheet = ss.getSheetByName(PAYROLL_INPUTS_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(PAYROLL_INPUTS_TAB);
      sheet.appendRow(['Employee ID', 'Name (English)', 'Period End', 'Deductions', 'Remark']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Which employees already have a row for this period?
    const existing = {};
    const cur = sheet.getDataRange().getValues();
    for (let i = 1; i < cur.length; i++) {
      if (formatKey_(cur[i][2]) === periodKey) existing[String(cur[i][0]).trim().toUpperCase()] = true;
    }

    const toAdd = [];
    for (let i = 1; i < hrData.length; i++) {
      const id = String(hrData[i][cId] || '').trim();
      if (!id) continue;
      if (excludeIds.indexOf(id.toUpperCase()) !== -1) continue;
      if (existing[id.toUpperCase()]) continue;
      const name = cName !== -1 ? String(hrData[i][cName] || '').trim() : '';
      toAdd.push([id, name, periodKey, '', '']);
    }

    if (toAdd.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 5).setValues(toAdd);
      sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).setNumberFormat('@'); // keep Period End as text dd-mm-yyyy
    }
    sheet.autoResizeColumns(1, 5);

    return { success: true, tab: PAYROLL_INPUTS_TAB, periodEnd: periodKey, seeded: toAdd.length,
      message: 'Seeded ' + toAdd.length + ' input rows for period ending ' + periodKey +
        '. Fill in Deductions/Remark, then run payroll. (OT is entered separately in OT_Entries.)' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getPayrollInputs_(ss, periodKey) {
  const out = {};
  const sheet = ss.getSheetByName(PAYROLL_INPUTS_TAB);
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return out;
  const H = data[0];
  const c = (name) => H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
  const cId = c('Employee ID'), cPer = c('Period End'), cDed = c('Deductions'), cRem = c('Remark');
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][cId] || '').trim();
    if (!id || formatKey_(data[i][cPer]) !== periodKey) continue;
    out[id] = {
      deductions: cDed !== -1 ? Number(data[i][cDed]) || 0 : 0,
      remark: cRem !== -1 ? String(data[i][cRem] || '').trim() : ''
    };
  }
  return out;
}

// Normalize a Period End cell (Date or string) to dd-mm-yyyy for matching.
function formatKey_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  const d = parseFlexibleDate_(v);
  return d ? formatDdMmYyyy_(d) : String(v || '').trim();
}

/**
 * Pairs a sorted list of {ts, stage} rows for ONE employee into
 * Check-In -> Check-Out cycles. Breaks are excluded before this runs.
 */
function buildCycles_(rows) {
  const cycles = [];
  let openCheckIn = null;
  rows.forEach((r) => {
    if (r.stage === 'Check In') {
      if (openCheckIn) cycles.push({ checkIn: openCheckIn, checkOut: null, incomplete: true });
      openCheckIn = r.ts;
    } else if (r.stage === 'Check Out') {
      if (openCheckIn) { cycles.push({ checkIn: openCheckIn, checkOut: r.ts, incomplete: false }); openCheckIn = null; }
    }
  });
  if (openCheckIn) cycles.push({ checkIn: openCheckIn, checkOut: null, incomplete: true });
  return cycles;
}

// Parse "DD/MM/YYYY HH:MM:SS" (kiosk/WA Timestamp), or pass through a native Date.
function parseLogTimestamp(ts) {
  if (!ts) return null;
  if (Object.prototype.toString.call(ts) === '[object Date]') return ts;
  const parts = String(ts).trim().split(' ');
  const d = parts[0].split('/');
  if (d.length !== 3) return null;
  const t = (parts[1] || '00:00:00').split(':');
  const day = parseInt(d[0], 10), mon = parseInt(d[1], 10), yr = parseInt(d[2], 10);
  if (!day || !mon || !yr) return null;
  return new Date(yr, mon - 1, day, parseInt(t[0], 10) || 0, parseInt(t[1], 10) || 0, parseInt(t[2], 10) || 0);
}

// ============================================================
// PAYROLL_CONTROL — cursor for rolling periods (+ re-run flagging)
// Columns: Run Date | Period Start | Period End | Employees | Total OT Hours | Tab Name | Status
// ============================================================
function getLastPayrollRun_(ss) {
  const sheet = ss.getSheetByName(PAYROLL_CONTROL_TAB);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  const last = data[data.length - 1];
  const periodEnd = parseFlexibleDate_(last[2]);
  if (!periodEnd) return null;
  return { runDate: last[0], periodStart: parseFlexibleDate_(last[1]), periodEnd: periodEnd, tabName: last[5] };
}

// ============================================================
// CANCEL / AMEND — rewind the most recent run so it can be re-run.
// Deletes its Payroll_Control row (moving the cursor back) and, by
// default, its period tab. Only the LATEST run can be cancelled, to
// keep the sequential rolling chain intact. Amend = cancel -> fix
// inputs (add staff, activate, load OT, set deductions) -> run again.
// params: { tab? (must match the latest run if given), deleteTab? ('false' to keep) }
// ============================================================
function cancelPayrollRun(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const ctrl = ss.getSheetByName(PAYROLL_CONTROL_TAB);
    if (!ctrl) return { success: false, error: 'No Payroll_Control tab — there are no runs to cancel.' };
    const data = ctrl.getDataRange().getValues();
    if (data.length <= 1) return { success: false, error: 'No payroll runs are logged — nothing to cancel.' };

    const lastRowNum = data.length; // 1-indexed (header is row 1)
    const last = data[data.length - 1];
    const lastTab = String(last[5] || '').trim();

    const reqTab = String(params.tab || '').trim();
    if (reqTab && reqTab !== lastTab) {
      return { success: false, error: 'Only the most recent run ("' + lastTab + '") can be cancelled, to keep the rolling periods sequential. You asked for "' + reqTab + '".' };
    }

    // Delete the period tab (default yes; pass deleteTab=false to keep it).
    const deleteTab = String(params.deleteTab || 'true').toLowerCase() !== 'false';
    let tabDeleted = false;
    if (deleteTab && lastTab) {
      const t = ss.getSheetByName(lastTab);
      if (t) { ss.deleteSheet(t); tabDeleted = true; }
    }

    // Remove the control row -> cursor recedes to the previous run.
    ctrl.deleteRow(lastRowNum);

    const newLast = getLastPayrollRun_(ss);
    return {
      success: true, cancelledTab: lastTab, tabDeleted: tabDeleted,
      newCursorEnd: newLast ? formatDdMmYyyy_(newLast.periodEnd) : null,
      message: 'Cancelled run "' + lastTab + '"' + (tabDeleted ? ' and deleted its tab' : ' (tab kept)') + '. ' +
        (newLast
          ? 'The next run will resume the day after ' + formatDdMmYyyy_(newLast.periodEnd) + '.'
          : 'No earlier runs remain — the next run is a first run, so set a Period Start.')
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function logPayrollRun_(ss, periodStart, periodEnd, employeeCount, totalOtHours, tabName, mode) {
  let sheet = ss.getSheetByName(PAYROLL_CONTROL_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(PAYROLL_CONTROL_TAB);
    sheet.appendRow(['Run Date', 'Period Start', 'Period End', 'Employees', 'Total OT Hours', 'Tab Name', 'Status']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    formatDdMmYyyy_(new Date()),
    formatDdMmYyyy_(periodStart),
    formatDdMmYyyy_(periodEnd),
    employeeCount,
    Math.round(totalOtHours * 100) / 100,
    tabName,
    'OK (' + (mode || 'attendance') + ')'
  ]);
}

// If an amended punch falls inside an already-run period, mark that run
// "NEEDS RE-RUN" so a corrected cycle can't silently drift into a later period.
function flagAffectedPayrollRun_(ss, amendedDate) {
  const sheet = ss.getSheetByName(PAYROLL_CONTROL_TAB);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  const statusCol = 7; // 1-indexed
  for (let i = 1; i < data.length; i++) {
    const ps = parseFlexibleDate_(data[i][1]);
    const pe = parseFlexibleDate_(data[i][2]);
    if (!ps || !pe) continue;
    if (amendedDate >= startOfDay_(ps) && amendedDate <= endOfDay_(pe)) {
      sheet.getRange(i + 1, statusCol).setValue('NEEDS RE-RUN (amended ' + formatDdMmYyyy_(new Date()) + ')');
      return data[i][5]; // tab name affected
    }
  }
  return null;
}

// ============================================================
// ADMIN — AMEND ATTENDANCE (audit-tracked, always allowed)
// params: { empId, date (dd-mm-yyyy / yyyy-mm-dd / dd/mm/yyyy accepted),
//           stage, time (HH:MM[:SS]), action ('set'|'delete'), admin }
// ============================================================
function amendAttendance(params) {
  try {
    const empId  = String(params.empId  || '').trim();
    const dateInput = String(params.date || '').trim();
    const stage  = String(params.stage  || '').trim();
    const action = String(params.action || 'set').trim().toLowerCase();
    const admin  = String(params.admin  || 'Unknown').trim();
    if (!empId || !dateInput) return { success: false, error: 'empId and date are required' };
    if (!['Check In','Break Out','Break In','Check Out'].includes(stage)) return { success: false, error: 'Invalid stage' };

    const targetDate = parseFlexibleDate_(dateInput);
    if (!targetDate) return { success: false, error: 'date must be a valid date (dd-mm-yyyy)' };

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    let log = ss.getSheetByName('Attendance_Log');
    if (!log) { log = ss.insertSheet('Attendance_Log'); log.appendRow(['Timestamp','Employee ID','Employee Name','Date','Stage','PIN Used']); }

    let empName = 'Unknown';
    const hr = ss.getSheetByName('HR Maya');
    if (hr) {
      const d = hr.getDataRange().getValues();
      const HH = d[0];
      const cId = HH.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
      const cName = HH.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
      for (let i = 1; i < d.length; i++) {
        if (cId !== -1 && String(d[i][cId]).trim() === empId) { empName = cName !== -1 ? String(d[i][cName]).trim() : empName; break; }
      }
    }

    const data = log.getDataRange().getValues();
    let rowNum = -1, oldTs = '';
    for (let i = 0; i < data.length; i++) {
      const rowDate = parseFlexibleDate_(data[i][3]);
      if (String(data[i][1]).trim() === empId && rowDate && sameCalendarDate_(rowDate, targetDate) && String(data[i][4]).trim() === stage) {
        rowNum = i + 1; oldTs = String(data[i][0]); break;
      }
    }

    const dateForWrite = formatDdMmYyyy_(targetDate);
    let result;

    if (action === 'delete') {
      if (rowNum === -1) return { success: false, error: 'No matching punch to delete' };
      log.deleteRow(rowNum);
      auditAmendment(ss, admin, empId, empName, dateForWrite, stage, oldTs, '(deleted)');
      result = { success: true, message: stage + ' on ' + dateForWrite + ' deleted for ' + empName };
    } else {
      const tp = String(params.time || '00:00').split(':');
      const hh = String(parseInt(tp[0], 10) || 0).padStart(2, '0');
      const mi = String(parseInt(tp[1], 10) || 0).padStart(2, '0');
      const sc = String(parseInt(tp[2], 10) || 0).padStart(2, '0');
      const newTsDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), Number(hh), Number(mi), Number(sc));
      const newTs = Utilities.formatDate(newTsDate, tz_(), 'dd/MM/yyyy HH:mm:ss');
      if (rowNum === -1) {
        log.appendRow([newTs, empId, empName, dateForWrite, stage, 'ADMIN']);
        auditAmendment(ss, admin, empId, empName, dateForWrite, stage, '(none)', newTs);
        result = { success: true, message: stage + ' on ' + dateForWrite + ' added for ' + empName + ' at ' + hh + ':' + mi };
      } else {
        log.getRange(rowNum, 1).setValue(newTs);
        log.getRange(rowNum, 4).setValue(dateForWrite);
        log.getRange(rowNum, 6).setValue('ADMIN');
        auditAmendment(ss, admin, empId, empName, dateForWrite, stage, oldTs, newTs);
        result = { success: true, message: stage + ' on ' + dateForWrite + ' updated for ' + empName + ' to ' + hh + ':' + mi };
      }
    }

    // Safety net: if this date sits inside an already-run period, flag it.
    const affectedTab = flagAffectedPayrollRun_(ss, startOfDay_(targetDate));
    if (affectedTab) {
      result.reRunNeeded = affectedTab;
      result.message += ' — NOTE: period "' + affectedTab + '" was already run and is now flagged NEEDS RE-RUN in Payroll_Control.';
    }
    return result;
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function auditAmendment(ss, admin, empId, empName, dateStr, stage, oldVal, newVal) {
  let a = ss.getSheetByName('Attendance_Amendments');
  if (!a) { a = ss.insertSheet('Attendance_Amendments'); a.appendRow(['When','Admin','Employee ID','Employee Name','Date','Stage','Old Value','New Value']); }
  const whenStr = formatDdMmYyyy_(new Date()) + ' ' + Utilities.formatDate(new Date(), tz_(), 'HH:mm:ss');
  a.appendRow([whenStr, admin, empId, empName, dateStr, stage, oldVal, newVal]);
}