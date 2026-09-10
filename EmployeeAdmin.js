// ============================================================
// EmployeeAdmin.gs — Maya HR : add / edit employees, roster search
// ============================================================
// Standalone Apps Script file. Shares global scope with Code.gs, etc.
// Writes to the HR Maya tab, resolving columns by HEADER NAME (robust to
// the real layout, which includes a PIN column). Do NOT redeclare helpers.
//
// Auto Employee ID: next available MT-##### (highest existing + 1),
//   join date IGNORED, stored IDs are permanent and never regenerated.
// Auto PIN: unique random 4-digit (none issued yet).
// Derived on save: Gross = Basic + allowances; Total Payable = Gross - Loan
//   - Other Deductions; Price Per Hour = Basic/26/8; Price Per Day = Basic/26.
//
// ENDPOINTS (add to Code.gs handleRequest):
//   saveEmployee (POST), getRoster, getEmployeeById
// ============================================================

// Header index map for the HR Maya sheet.
function hrCols_(H) {
  const c = {};
  H.forEach((h, i) => { c[String(h).trim().toLowerCase()] = i; });
  return c;
}
function col_(c, name) { const k = String(name).trim().toLowerCase(); return (k in c) ? c[k] : -1; }

// Next available Employee ID (MT-#####), highest stored + 1.
function getNextEmployeeId_(ss) {
  const hr = ss.getSheetByName('HR Maya');
  const data = hr.getDataRange().getValues();
  const c = hrCols_(data[0]);
  const cId = col_(c, 'Employee ID');
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][cId] || '').match(/MT-(\d+)/i);
    if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
  }
  return 'MT-' + String(max + 1).padStart(5, '0');
}

// Unique random 4-digit PIN not already used in the PIN column.
function generateUniquePin_(ss) {
  const hr = ss.getSheetByName('HR Maya');
  const data = hr.getDataRange().getValues();
  const c = hrCols_(data[0]);
  const cPin = col_(c, 'PIN');
  const used = {};
  if (cPin !== -1) for (let i = 1; i < data.length; i++) { const p = String(data[i][cPin] || '').trim(); if (p) used[p] = true; }
  for (let tries = 0; tries < 10000; tries++) {
    const pin = String(Math.floor(1000 + Math.random() * 9000));
    if (!used[pin]) return pin;
  }
  return String(1000 + (Date.now() % 9000)); // fallback
}

// ------------------------------------------------------------
// SAVE EMPLOYEE — append new (auto ID + PIN) or edit existing (by employeeId).
// params: employeeId? (present => edit), nameEnglish, nameArabic, civilId,
//   basic, food, accom, conv, bonus, loan, otherDed, branch, status
// ------------------------------------------------------------
function saveEmployee(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found" };

    const data = hr.getDataRange().getValues();
    const H = data[0];
    const c = hrCols_(H);
    const width = H.length;

    const nameEng = String(params.nameEnglish || '').trim();
    const isEdit = !!String(params.employeeId || '').trim();

    if (!isEdit && !nameEng) return { success: false, error: 'Name (English) is required' };

    // numeric inputs
    const num = (v) => { const n = Number(v); return isNaN(n) ? 0 : n; };
    const basic = num(params.basic), food = num(params.food), accom = num(params.accom),
          conv = num(params.conv), bonus = num(params.bonus), loan = num(params.loan), otherDed = num(params.otherDed);

    // Locate existing row (edit) or prepare a new one.
    let rowNum = -1, rowVals;
    const cId = col_(c, 'Employee ID');
    let empId;
    if (isEdit) {
      empId = String(params.employeeId).trim();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][cId] || '').trim().toUpperCase() === empId.toUpperCase()) { rowNum = i + 1; rowVals = data[i].slice(); break; }
      }
      if (rowNum === -1) return { success: false, error: 'Employee ' + empId + ' not found to edit' };
    } else {
      empId = getNextEmployeeId_(ss);
      rowVals = new Array(width).fill('');
    }

    // helper to set a cell by header (only if the column exists)
    const set = (name, val) => { const i = col_(c, name); if (i !== -1) rowVals[i] = val; };
    const hasCol_ = (name) => col_(c, name) !== -1;
    // helper: for edit, only overwrite when the param was actually provided
    const provided = (key) => params[key] !== undefined && params[key] !== '';

    set('Employee ID', empId);
    if (!isEdit) { set('PIN', generateUniquePin_(ss)); }

    if (!isEdit || provided('nameEnglish')) set('Name (English)', nameEng || rowVals[col_(c,'Name (English)')]);
    if (!isEdit || provided('nameArabic')) set('Name (Arabic)', String(params.nameArabic || '').trim());
    if (!isEdit || provided('civilId'))    set('Civil ID', String(params.civilId || '').trim());
    if (!isEdit || provided('iban'))       set('IBAN', String(params.iban || '').trim());
    if (!isEdit || provided('branch'))     set(hasCol_('Location') ? 'Location' : 'Branch', String(params.branch || '').trim());
    if (provided('shift'))          set('Shift', String(params.shift).trim());
    // Document fields from the form (Staff Submission set). Written only when provided.
    if (provided('email'))          set('Email', String(params.email).trim());
    if (provided('mobile'))         set('Mobile', String(params.mobile).trim());
    if (provided('whatsapp'))       set('WhatsApp', String(params.whatsapp).trim());
    if (provided('passportNo'))     set('Passport No', String(params.passportNo).trim());
    if (provided('cidExpiry'))      set('CID Expiry', String(params.cidExpiry).trim());
    if (provided('passportIssue'))  set('Passport Issue', String(params.passportIssue).trim());
    if (provided('passportExpiry')) set('Passport Expiry', String(params.passportExpiry).trim());
    if (provided('healthIssue'))    set('Health Issue', String(params.healthIssue).trim());
    if (provided('healthExpiry'))   set('Health Expiry', String(params.healthExpiry).trim());
    if (provided('dateOfJoin'))     set('Date of Join', String(params.dateOfJoin).trim());
    set('Status', String(params.status || (isEdit ? rowVals[col_(c,'Status')] : 'Active')).trim() || 'Active');

    // Money fields — recompute derived whenever any component is provided (or on create).
    const anyMoney = !isEdit || ['basic','food','accom','conv','bonus','loan','otherDed'].some(provided);
    if (anyMoney) {
      const B = provided('basic') || !isEdit ? basic : Number(rowVals[col_(c,'Basic Salary')]) || 0;
      const F = provided('food') || !isEdit ? food : Number(rowVals[col_(c,'Food Allowance')]) || 0;
      const A = provided('accom') || !isEdit ? accom : Number(rowVals[col_(c,'Accommodation')]) || 0;
      const C = provided('conv') || !isEdit ? conv : Number(rowVals[col_(c,'Conveyance')]) || 0;
      const Bo = provided('bonus') || !isEdit ? bonus : Number(rowVals[col_(c,'Bonus')]) || 0;
      const L = provided('loan') || !isEdit ? loan : Number(rowVals[col_(c,'Loan')]) || 0;
      const O = provided('otherDed') || !isEdit ? otherDed : Number(rowVals[col_(c,'Other Deductions')]) || 0;
      const round3 = (n) => Math.round(n * 1000) / 1000;
      // Only the INPUT components are written here. Gross / Total Payable /
      // Price Per Hour / Price Per Day are LIVE FORMULAS set after the write,
      // so a later edit (sheet or UI) recalculates automatically.
      set('Basic Salary', round3(B)); set('Food Allowance', round3(F)); set('Accommodation', round3(A));
      set('Conveyance', round3(C)); set('Bonus', round3(Bo)); set('Loan', round3(L)); set('Other Deductions', round3(O));
    }

    // Write
    if (isEdit) {
      hr.getRange(rowNum, 1, 1, width).setValues([rowVals]);
    } else {
      hr.appendRow(rowVals);
      rowNum = hr.getLastRow();
      // keep Civil ID as text (leading zeros) if that column exists
      const civI = col_(c, 'Civil ID');
      if (civI !== -1) hr.getRange(rowNum, civI + 1).setNumberFormat('@');
    }

    // Set the derived columns to LIVE FORMULAS for this row.
    setDerivedFormulas_(hr, c, rowNum);

    const pinOut = col_(c, 'PIN') !== -1 ? rowVals[col_(c, 'PIN')] : '';
    return {
      success: true, employeeId: empId, pin: pinOut, edited: isEdit,
      name: rowVals[col_(c, 'Name (English)')],
      message: (isEdit ? 'Updated ' : 'Added ') + rowVals[col_(c, 'Name (English)')] + ' (' + empId + ')' +
        (isEdit ? '' : ' — PIN ' + pinOut)
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ------------------------------------------------------------
// ROSTER — light list for search (Name / ID / Civil ID). From HR Maya.
// ------------------------------------------------------------
function getRoster(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found", data: [] };
    const data = hr.getDataRange().getValues();
    const c = hrCols_(data[0]);
    const cId = col_(c, 'Employee ID'), cName = col_(c, 'Name (English)'), cCivil = col_(c, 'Civil ID'),
          cBasic = col_(c, 'Basic Salary'), cRate = col_(c, 'Price Per Hour'), cStatus = col_(c, 'Status'), cPin = col_(c, 'PIN'), cIban = col_(c, 'IBAN');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][cId] || '').trim();
      if (!id) continue;
      out.push({
        id: id,
        name: cName !== -1 ? String(data[i][cName] || '').trim() : '',
        civilId: cCivil !== -1 ? String(data[i][cCivil] || '').trim() : '',
        basic: cBasic !== -1 ? Number(data[i][cBasic]) || 0 : 0,
        rate: cRate !== -1 ? Number(data[i][cRate]) || 0 : 0,
        status: cStatus !== -1 ? String(data[i][cStatus] || '').trim() : '',
        pin: cPin !== -1 ? String(data[i][cPin] || '').trim() : '',
        iban: cIban !== -1 ? String(data[i][cIban] || '').trim() : ''
      });
    }
    return { success: true, data: out, count: out.length };
  } catch (err) { return { success: false, error: err.toString(), data: [] }; }
}

// ------------------------------------------------------------
// One employee's full record (for prefilling the edit form).
// ------------------------------------------------------------
function getEmployeeById(params) {
  try {
    const empId = String((params && params.employeeId) || '').trim();
    if (!empId) return { success: false, error: 'employeeId required' };
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const data = hr.getDataRange().getValues();
    const H = data[0];
    const c = hrCols_(H);
    const cId = col_(c, 'Employee ID');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cId] || '').trim().toUpperCase() === empId.toUpperCase()) {
        const rec = {};
        H.forEach((h, j) => { rec[String(h).trim()] = data[i][j]; });
        return { success: true, data: rec };
      }
    }
    return { success: false, error: 'Employee ' + empId + ' not found' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ------------------------------------------------------------
// ALL employees as full HR Maya row-objects (one true source).
// Used by the dashboard employee table + profile. Fixes the old
// getIntegratedEmployees "0" by reading HR Maya directly.
// ------------------------------------------------------------
function getEmployeesFull(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found", data: [] };
    const data = hr.getDataRange().getValues();
    const H = data[0];
    const c = hrCols_(H);
    const cId = col_(c, 'Employee ID');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      if (!String(data[i][cId] || '').trim()) continue;
      const rec = {};
      H.forEach((h, j) => { rec[String(h).trim()] = data[i][j]; });
      out.push(rec);
    }
    return { success: true, data: out, count: out.length };
  } catch (err) { return { success: false, error: err.toString(), data: [] }; }
}

// ------------------------------------------------------------
// Derived columns as LIVE FORMULAS (so any edit — sheet or UI —
// recalculates): Gross, Total Payable, Price Per Hour, Price Per Day.
// ------------------------------------------------------------
function colLetter_(n) { // 1-based column index -> letter(s)
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function setDerivedFormulas_(hr, c, r) {
  const LET = (name) => { const i = col_(c, name); return i === -1 ? null : colLetter_(i + 1); };
  const basic = LET('Basic Salary'), food = LET('Food Allowance'), accom = LET('Accommodation'),
        conv = LET('Conveyance'), bonus = LET('Bonus'), loan = LET('Loan'), other = LET('Other Deductions'),
        gross = LET('Gross Salary'), pay = LET('Total Payable'), pph = LET('Price Per Hour'), ppd = LET('Price Per Day');

  if (gross && basic) {
    const parts = [basic]; [food, accom, conv, bonus].forEach(x => { if (x) parts.push(x); });
    const cell = hr.getRange(gross + r);
    cell.setFormula('=' + parts.map(x => x + r).join('+')); cell.setNumberFormat('0.000');
  }
  if (pay && gross) {
    let f = '=' + gross + r; if (loan) f += '-' + loan + r; if (other) f += '-' + other + r;
    const cell = hr.getRange(pay + r); cell.setFormula(f); cell.setNumberFormat('0.000');
  }
  if (pph && basic) { const cell = hr.getRange(pph + r); cell.setFormula('=' + basic + r + '/26/8'); cell.setNumberFormat('0.000'); }
  if (ppd && basic) { const cell = hr.getRange(ppd + r); cell.setFormula('=' + basic + r + '/26'); cell.setNumberFormat('0.000'); }
}

// One-shot: convert the derived columns of every existing employee row to
// formulas. Run once from the editor after deploying. Safe to re-run.
function convertDerivedColumnsToFormulas() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found" };
    const data = hr.getDataRange().getValues();
    const c = hrCols_(data[0]);
    const cId = col_(c, 'Employee ID');
    let n = 0;
    for (let i = 1; i < data.length; i++) {
      if (!String(data[i][cId] || '').trim()) continue;
      setDerivedFormulas_(hr, c, i + 1);
      n++;
    }
    return { success: true, rows: n,
      message: 'Converted Gross, Total Payable, Price Per Hour and Price Per Day to live formulas for ' + n + ' rows.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}