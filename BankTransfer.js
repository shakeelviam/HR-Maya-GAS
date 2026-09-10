// ============================================================
// BankTransfer.gs — Maya HR : populate the bank salary-transfer sheet
// ============================================================
// Fills the "Bank Transfer Details" tab (a replica of the Warba template:
//   Employee Civil ID | Employee Name | IBAN | Salary Amount (KD))
// with each employee's Total Payable from a payroll period tab.
//
// Matching chain (Civil ID is the bank's key, payroll's key is Employee ID):
//   Payroll tab  (Employee ID -> Total Payable)
//     -> HR Maya (Employee ID  -> Civil ID)
//       -> Bank tab (Civil ID  -> row)  => write Salary Amount
//
// Safety: flags any PAID employee with no bank row (they'd miss payment),
// and any matched bank row with a blank IBAN. Reconciles the written total
// against the payroll net total.
//
// ENDPOINT (add to Code.gs): populateBankTransfer
//   ?method=populateBankTransfer&tab=Payroll%2021Aug-21Aug
// ============================================================

const BANK_TAB_DEFAULT = 'Bank Transfer Details';

// ------------------------------------------------------------
// ONE-SHOT: ensure HR Maya has an 'IBAN' column, back-fill it from the
// existing Bank Transfer Details sheet (matched by Civil ID), and correct
// the names in the bank sheet to HR Maya's Name (English). Safe to re-run.
// Run once from the editor: syncBankDetailsToHrMaya
// ------------------------------------------------------------
function syncBankDetailsToHrMaya(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const bankTabName = String(params.bankTab || BANK_TAB_DEFAULT).trim();

    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found." };
    const hd = hr.getDataRange().getValues();
    let H = hd[0];
    const hc = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const hcId = hc('Employee ID'), hcCivil = hc('Civil ID'), hcName = hc('Name (English)');

    // Ensure IBAN column exists (append if missing).
    let hcIban = hc('IBAN');
    if (hcIban === -1) {
      hr.getRange(1, H.length + 1).setValue('IBAN').setFontWeight('bold');
      hcIban = H.length; // 0-based index of the new column
      H = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0]; // refresh
    }

    // Build Civil ID -> {row, empId, name} from HR Maya.
    const hrByCivil = {};
    for (let i = 1; i < hd.length; i++) {
      const civ = normCivil_(hd[i][hcCivil]);
      if (civ) hrByCivil[civ] = { row: i + 1, empId: String(hd[i][hcId] || '').trim(), name: String(hd[i][hcName] || '').trim() };
    }

    const bank = ss.getSheetByName(bankTabName);
    if (!bank) return { success: false, error: 'Bank tab "' + bankTabName + '" not found.' };
    const bd = bank.getDataRange().getValues();
    const bH = bd[0];
    const bc = (sub) => bH.findIndex(h => String(h).trim().toLowerCase().indexOf(sub.toLowerCase()) !== -1);
    const bcCivil = bc('civil'), bcName = bc('name'), bcIban = bc('iban');

    const ibansFilled = []; const namesFixed = []; const unmatched = [];
    for (let i = 1; i < bd.length; i++) {
      const civ = normCivil_(bd[i][bcCivil]);
      if (!civ) continue;
      const hit = hrByCivil[civ];
      if (!hit) { unmatched.push(String(bd[i][bcName] || '').trim() + ' / ' + civ); continue; }

      // Back-fill IBAN into HR Maya (only if bank has one and HR Maya's is empty/different).
      const bankIban = bcIban !== -1 ? String(bd[i][bcIban] || '').trim() : '';
      if (bankIban) {
        const cur = String(hr.getRange(hit.row, hcIban + 1).getValue() || '').trim();
        if (cur !== bankIban) { hr.getRange(hit.row, hcIban + 1).setValue(bankIban).setNumberFormat('@'); ibansFilled.push(hit.empId); }
      }
      // Fix the name in the bank sheet to HR Maya's Name (English).
      if (bcName !== -1 && hit.name) {
        const curName = String(bd[i][bcName] || '').trim();
        if (curName !== hit.name) { bank.getRange(i + 1, bcName + 1).setValue(hit.name); namesFixed.push(curName + ' -> ' + hit.name); }
      }
    }

    return {
      success: true, bankTab: bankTabName,
      ibansBackFilled: ibansFilled.length, namesCorrected: namesFixed.length,
      unmatchedRows: unmatched,
      nameChanges: namesFixed,
      message: 'Added/confirmed HR Maya "IBAN" column. Back-filled ' + ibansFilled.length + ' IBANs, corrected ' + namesFixed.length + ' names in "' + bankTabName + '"' +
        (unmatched.length ? '. ' + unmatched.length + ' bank row(s) had no Civil ID match in HR Maya (left untouched): ' + unmatched.join(', ') : '.')
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Normalize a Civil ID to digits only (handles 2.98e11 / "298...0" / floats).
function normCivil_(v) {
  if (v == null) return '';
  let s = String(v).trim();
  if (s.indexOf('E') !== -1 || s.indexOf('e') !== -1) { // scientific notation
    const n = Number(s); if (!isNaN(n)) s = n.toFixed(0);
  }
  if (s.endsWith('.0')) s = s.slice(0, -2);
  return s.replace(/[^0-9]/g, '');
}

function populateBankTransfer(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const payrollTab = String(params.tab || '').trim();
    const bankTabName = String(params.bankTab || BANK_TAB_DEFAULT).trim();
    if (!payrollTab) return { success: false, error: 'Provide the payroll tab, e.g. tab=Payroll 21Aug-21Aug' };
    // Employees paid outside the bank (e.g. cash) — skip them in the bank sheet.
    const excludeIds = String(params.excludeIds || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

    // --- 1. Payroll: Employee ID -> { pay, name } ---
    const pay = ss.getSheetByName(payrollTab);
    if (!pay) return { success: false, error: 'Payroll tab "' + payrollTab + '" not found.' };
    const pd = pay.getDataRange().getValues();
    const pH = pd[0];
    const pc = (n) => pH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const pcId = pc('Employee ID'), pcPay = pc('Total Payable'), pcName = pc('Name (English)');
    if (pcId === -1 || pcPay === -1) return { success: false, error: 'Payroll tab missing Employee ID or Total Payable column.' };
    const payByEmp = {};
    for (let i = 1; i < pd.length; i++) {
      const id = String(pd[i][pcId] || '').trim();
      if (!id) continue; // skips the TOTALS row (blank Employee ID)
      payByEmp[id] = { pay: Number(pd[i][pcPay]) || 0, name: pcName !== -1 ? String(pd[i][pcName] || '').trim() : '' };
    }

    // --- 2. HR Maya: Civil ID -> { empId, name, iban } (HR Maya is the master) ---
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found." };
    const hd = hr.getDataRange().getValues();
    const hH = hd[0];
    const hc = (n) => hH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const hcId = hc('Employee ID'), hcCivil = hc('Civil ID'), hcName = hc('Name (English)'), hcIban = hc('IBAN');
    const empByCivil = {}; const hrByEmp = {};
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][hcId] || '').trim();
      const civ = normCivil_(hd[i][hcCivil]);
      const rec = { empId: id, name: hcName !== -1 ? String(hd[i][hcName] || '').trim() : '', iban: hcIban !== -1 ? String(hd[i][hcIban] || '').trim() : '' };
      if (id && civ) empByCivil[civ] = id;
      if (id) hrByEmp[id] = rec;
    }

    // --- 3. Bank tab: match by Civil ID, write Salary Amount ---
    const bank = ss.getSheetByName(bankTabName);
    if (!bank) return { success: false, error: 'Bank tab "' + bankTabName + '" not found. Create it as a replica of the Warba template.' };
    const bd = bank.getDataRange().getValues();
    const bH = bd[0];
    const bc = (sub) => bH.findIndex(h => String(h).trim().toLowerCase().indexOf(sub.toLowerCase()) !== -1);
    const bcCivil = bc('civil'), bcIban = bc('iban'), bcName = bc('name');
    let bcAmt = bc('salary amount'); if (bcAmt === -1) bcAmt = bc('amount');
    if (bcCivil === -1 || bcAmt === -1) return { success: false, error: 'Bank tab must have a Civil ID column and a Salary Amount column.' };

    let written = 0, total = 0;
    const matched = {}; const noIban = []; const cashSkipped = [];
    for (let i = 1; i < bd.length; i++) {
      const civ = normCivil_(bd[i][bcCivil]);
      if (!civ) continue;
      const empId = empByCivil[civ];
      const hrRec = empId ? hrByEmp[empId] : null;
      // Refresh Name + IBAN from HR Maya (the master) for every matched row.
      if (hrRec) {
        if (bcName !== -1 && hrRec.name) bank.getRange(i + 1, bcName + 1).setValue(hrRec.name);
        if (bcIban !== -1 && hrRec.iban) bank.getRange(i + 1, bcIban + 1).setValue(hrRec.iban).setNumberFormat('@');
      }
      // Cash-paid (excluded) employees: zero their bank amount so nothing transfers.
      if (empId && excludeIds.indexOf(empId.toUpperCase()) !== -1) {
        bank.getRange(i + 1, bcAmt + 1).setValue(0).setNumberFormat('0.000');
        matched[empId] = true;
        if (payByEmp[empId]) cashSkipped.push(payByEmp[empId].name + ' (' + empId + ')');
        continue;
      }
      const rec = empId ? payByEmp[empId] : null;
      if (rec) {
        const amt = Math.round(rec.pay * 1000) / 1000;
        bank.getRange(i + 1, bcAmt + 1).setValue(amt).setNumberFormat('0.000');
        written++; total += amt; matched[empId] = true;
        const ibanNow = (hrRec && hrRec.iban) ? hrRec.iban : (bcIban !== -1 ? String(bd[i][bcIban] || '').trim() : '');
        if (!ibanNow) noIban.push(rec.name + ' (' + empId + ')');
      } else {
        bank.getRange(i + 1, bcAmt + 1).setValue(0).setNumberFormat('0.000'); // no pay this period
      }
    }

    // --- 4. Safety: paid employees with no bank row (excluded/cash ones don't count) ---
    let payrollNet = 0, bankNet = 0; const missingBank = [];
    Object.keys(payByEmp).forEach(id => {
      payrollNet += payByEmp[id].pay;
      if (excludeIds.indexOf(id.toUpperCase()) === -1) bankNet += payByEmp[id].pay;
      if (!matched[id] && payByEmp[id].pay > 0) missingBank.push(payByEmp[id].name + ' (' + id + ')');
    });
    payrollNet = Math.round(payrollNet * 1000) / 1000;
    bankNet = Math.round(bankNet * 1000) / 1000;
    total = Math.round(total * 1000) / 1000;

    // --- 5. Dated snapshot: copy the populated bank sheet to a period tab ---
    let snapshotTab = '';
    const periodLabel = payrollTab.replace(/^Payroll\s*/i, '').trim();
    if (periodLabel) {
      snapshotTab = 'Bank Transfer ' + periodLabel;
      const old = ss.getSheetByName(snapshotTab);
      if (old) ss.deleteSheet(old); // replace prior snapshot for the same period
      const copy = bank.copyTo(ss);
      copy.setName(snapshotTab);
    }

    return {
      success: true, tab: payrollTab, bankTab: bankTabName, snapshotTab: snapshotTab,
      written: written, totalTransfer: total, payrollNet: payrollNet, bankExpected: bankNet,
      cashSkipped: cashSkipped,
      reconciled: Math.abs(total - bankNet) < 0.001 && missingBank.length === 0,
      missingFromBankSheet: missingBank, missingIban: noIban,
      message: 'Populated ' + written + ' salary amounts into "' + bankTabName + '" — bank total ' + total.toFixed(3) + ' KD' +
        (Math.abs(total - bankNet) < 0.001 ? ' (matches expected bank total).' : ' (expected ' + bankNet.toFixed(3) + ').') +
        (snapshotTab ? ' Snapshot saved as "' + snapshotTab + '".' : '') +
        (cashSkipped.length ? ' Cash (excluded, set to 0): ' + cashSkipped.join(', ') + '.' : '') +
        (missingBank.length ? ' ⚠ ' + missingBank.length + ' PAID employee(s) have NO bank row: ' + missingBank.join(', ') + '.' : '') +
        (noIban.length ? ' ⚠ ' + noIban.length + ' matched row(s) have a blank IBAN: ' + noIban.join(', ') + '.' : '')
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}