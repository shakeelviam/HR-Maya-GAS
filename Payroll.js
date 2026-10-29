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
// ============================================================
// RETIRED — computeMonthlyPayroll
// ------------------------------------------------------------
// This was a SECOND payroll engine running alongside runPayrollFromReview
// (PayrollReview.gs). Both wrote the same 'Payroll ddMMM-ddMMM' tab name and
// both logged to the same control tab, but they produced DIFFERENT numbers:
// this one had no unpaid-leave, absence, encashment or scheduled deductions,
// and used a hardcoded DUTY_HOURS instead of each employee's shift hours.
// Whichever page you happened to use decided what people got paid, and the
// control log could not tell you which engine produced a given run.
//
// The two things this engine did BETTER have been ported into
// PayrollReview.gs (see resolvePayrollPeriod_):
//   1. Refusing a period shorter than 7 days unless force=true
//   2. Defaulting the period start sensibly rather than producing a sliver
//
// runPayrollFromReview is now the single payroll engine.
//
// This stub is deliberately left in place instead of deleting the function,
// so that any old bookmark, script or router case that still calls it fails
// LOUDLY with an explanation rather than silently doing nothing.
// ============================================================
function computeMonthlyPayroll(params) {
  return {
    success: false,
    retired: true,
    error: 'computeMonthlyPayroll has been retired. Use the Payroll Review page ' +
           '(getPayrollReview / runPayrollFromReview) — it is now the single payroll ' +
           'engine and includes unpaid leave, absence, encashment and scheduled ' +
           'deductions, which this function never applied.'
  };
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
    // 'Day Off' is a valid stage — the kiosk, Add Punch and the voice supervisor
    // all write it. Omitting it here meant an admin could not amend a Day Off
    // row through this path ("Invalid stage").
    if (!['Check In','Break Out','Break In','Check Out','Day Off'].includes(stage)) return { success: false, error: 'Invalid stage' };

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