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
// OTHER PAYEES — people paid with every payroll who are NOT staff
// (internal payments). They are rows in the "Bank_Other_Payees" tab,
// never in the code:
//   Name | Civil ID | IBAN | Amount (KD) | Location | Status | Remark
// Status Active (or blank) = paid every run; anything else = skipped.
// They go into the Warba bank sheet (after staff) and the accountant's
// Payroll by Location sheet, but never into the payroll tab or payslips.
// The tab is created with its headers the first time the bank sheet runs.
// ------------------------------------------------------------
const OTHER_PAYEES_TAB = 'Bank_Other_Payees';
const OTHER_PAYEES_HEADERS = ['Name', 'Civil ID', 'IBAN', 'Amount (KD)', 'Location', 'Status', 'Remark'];

function getOtherPayees_(ss) {
  let sh = ss.getSheetByName(OTHER_PAYEES_TAB);
  if (!sh) {
    sh = ss.insertSheet(OTHER_PAYEES_TAB);
    sh.getRange(1, 1, 1, OTHER_PAYEES_HEADERS.length).setValues([OTHER_PAYEES_HEADERS]).setFontWeight('bold');
    sh.getRange(2, 2, 200, 2).setNumberFormat('@');          // Civil ID + IBAN stay text
    sh.setFrozenRows(1);
    return { list: [], skipped: [] };
  }
  const d = sh.getDataRange().getValues();
  if (d.length < 2) return { list: [], skipped: [] };
  const H = d[0].map(h => String(h).trim().toLowerCase());
  const c = n => H.findIndex(h => h.indexOf(n) === 0);
  const cN = c('name'), cC = c('civil'), cI = c('iban'), cA = c('amount'), cL = c('location'), cS = c('status');
  const list = [], skipped = [];
  for (let i = 1; i < d.length; i++) {
    const name = cN >= 0 ? String(d[i][cN] || '').trim() : '';
    if (!name) continue;
    const status = cS >= 0 ? String(d[i][cS] || '').trim().toLowerCase() : '';
    if (status && status !== 'active') continue;
    const iban = cI >= 0 ? String(d[i][cI] || '').replace(/\s+/g, '').trim() : '';
    const amt = cA >= 0 ? Math.round((Number(d[i][cA]) || 0) * 1000) / 1000 : 0;
    if (!iban || !(amt > 0)) { skipped.push(name + (iban ? ' (no amount)' : ' (no IBAN)')); continue; }
    list.push({ name: name, civil: cC >= 0 ? normCivil_(d[i][cC]) : '', iban: iban, amount: amt,
                loc: (cL >= 0 ? String(d[i][cL] || '').trim() : '') || 'Other payees' });
  }
  return { list: list, skipped: skipped };
}

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

// ============================================================
// populateBankTransfer — TWO sheets from one payroll tab (22-09-2026)
// ------------------------------------------------------------
//   "Bank Transfer <period>"  — for Warba. Exactly the template's columns
//       (Employee Civil ID | Employee Name | IBAN | Salary Amount (KD)),
//       ONLY the people paid by bank: Net > 0 and an IBAN on file.
//   "Payroll by Location <period>" — for the accountant's cost centres.
//       The same Warba columns + Location + Payment, for EVERY paid
//       employee (bank and cash), sorted by Location, with a per-location
//       total block (staff, bank, cash, total) underneath.
//
// Payment rule: an IBAN on file (HR Maya, else the Warba template row for
// the same Civil ID) → Bank. No IBAN anywhere → Cash. excludeIds → Cash.
// Nothing about who is paid how is in the code — it comes from the IBANs.
//
// Reads Net Payable (Payroll Review tabs) or Total Payable (older tabs).
// The "Bank Transfer Details" template itself is not changed.
// Re-running refreshes both tabs in place (cleared, not deleted), so a
// link already shared to either tab keeps working. The bank sheet ends
// with a total row: the transfer total in the amount column, nothing else.
// ============================================================
function populateBankTransfer(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const payrollTab = String(params.tab || '').trim();
    const bankTabName = String(params.bankTab || BANK_TAB_DEFAULT).trim();
    if (!payrollTab) return { success: false, error: 'Provide the payroll tab, e.g. tab=Payroll 22Aug-20Sep' };
    const excludeIds = String(params.excludeIds || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

    // --- 1. Payroll tab ---
    const pay = ss.getSheetByName(payrollTab);
    if (!pay) return { success: false, error: 'Payroll tab "' + payrollTab + '" not found.' };
    const pd = pay.getDataRange().getValues();
    const pH = pd[0].map(h => String(h).trim());
    const pc = (n) => pH.findIndex(h => h.toLowerCase() === n.toLowerCase());
    const pcId = pc('Employee ID');
    let pcPay = pc('Net Payable'); if (pcPay === -1) pcPay = pc('Total Payable');
    if (pcId === -1 || pcPay === -1) return { success: false, error: 'Payroll tab needs Employee ID and Net Payable (or Total Payable).' };
    const payRows = [];
    for (let i = 1; i < pd.length; i++) {
      const id = String(pd[i][pcId] || '').trim();
      if (!id) continue;                                   // TOTALS row
      payRows.push({ id: id, row: pd[i], pay: Math.round((Number(pd[i][pcPay]) || 0) * 1000) / 1000 });
    }

    // --- 2. HR Maya (master): ID -> civil, name, IBAN, location ---
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found." };
    const hd = hr.getDataRange().getValues();
    const hH = hd[0];
    const hc = (n) => hH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const hcId = hc('Employee ID'), hcCivil = hc('Civil ID'), hcName = hc('Name (English)'), hcIban = hc('IBAN');
    let hcLoc = hc('Location'); if (hcLoc === -1) hcLoc = hc('Branch');
    const hrByEmp = {};
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][hcId] || '').trim(); if (!id) continue;
      hrByEmp[id] = {
        civil: hcCivil !== -1 ? normCivil_(hd[i][hcCivil]) : '',
        name: hcName !== -1 ? String(hd[i][hcName] || '').trim() : '',
        iban: hcIban !== -1 ? String(hd[i][hcIban] || '').replace(/\s+/g, '').trim() : '',
        loc: hcLoc !== -1 ? String(hd[i][hcLoc] || '').trim() : ''
      };
    }

    // --- 3. Warba template: headers + IBAN fallback by Civil ID ---
    const tpl = ss.getSheetByName(bankTabName);
    let tplHeaders = ['Employee Civil ID', 'Employee Name', 'IBAN', 'Salary Amount (KD)'];
    const tplIban = {};
    if (tpl && tpl.getLastRow() >= 1) {
      const td = tpl.getDataRange().getValues();
      tplHeaders = td[0].map(h => String(h).trim());
      const tc = (sub) => tplHeaders.findIndex(h => h.toLowerCase().indexOf(sub) !== -1);
      const tcCivil = tc('civil'), tcIban = tc('iban');
      for (let i = 1; i < td.length; i++) {
        const civ = tcCivil !== -1 ? normCivil_(td[i][tcCivil]) : '';
        const ib = tcIban !== -1 ? String(td[i][tcIban] || '').replace(/\s+/g, '').trim() : '';
        if (civ && ib) tplIban[civ] = ib;
      }
    }
    const tcol = (sub) => tplHeaders.findIndex(h => h.toLowerCase().indexOf(sub) !== -1);
    const bCivil = tcol('civil'), bName = tcol('name'), bIban = tcol('iban');
    let bAmt = tcol('salary amount'); if (bAmt === -1) bAmt = tcol('amount');

    // --- 4. Decide Bank / Cash per employee ---
    const bankRows = [], refRows = [], cash = [], noCivil = [];
    let bankTotal = 0, cashTotal = 0;
    payRows.forEach(p => {
      const h = hrByEmp[p.id] || { civil: '', name: '', iban: '', loc: '' };
      const iban = h.iban || (h.civil ? tplIban[h.civil] || '' : '');
      const excluded = excludeIds.indexOf(p.id.toUpperCase()) !== -1;
      let method = 'Bank';
      if (excluded || !iban) method = 'Cash';
      if (method === 'Bank' && !h.civil) { method = 'Cash'; noCivil.push((h.name || p.id) + ' (' + p.id + ')'); }
      if (p.pay <= 0) method = '—';
      if (method === 'Bank') {
        const r = new Array(tplHeaders.length).fill('');
        if (bCivil !== -1) r[bCivil] = h.civil;
        if (bName !== -1) r[bName] = h.name || String(p.row[pc('Name (English)')] || '');
        if (bIban !== -1) r[bIban] = iban;
        if (bAmt !== -1) r[bAmt] = p.pay;
        bankRows.push(r); bankTotal += p.pay;
      } else if (method === 'Cash') {
        cash.push((h.name || p.id) + ' (' + p.id + ') ' + p.pay.toFixed(3)); cashTotal += p.pay;
      }
      refRows.push({ loc: h.loc, id: p.id, pay: p.pay, vals: [p.id, h.name || String(p.row[pc('Name (English)')] || ''), h.loc, h.civil,
                     method, method === 'Bank' ? iban : ''].concat(pH.map((_, j) => j).filter(j => j !== pcId && j !== pc('Name (English)')).map(j => p.row[j])) });
    });
    // Other payees (non-staff, every payroll) — bank only, after staff.
    const others = getOtherPayees_(ss);
    let otherTotal = 0; const otherNoCivil = [];
    others.list.forEach(o => {
      const r = new Array(tplHeaders.length).fill('');
      if (bCivil !== -1) r[bCivil] = o.civil;
      if (bName !== -1) r[bName] = o.name;
      if (bIban !== -1) r[bIban] = o.iban;
      if (bAmt !== -1) r[bAmt] = o.amount;
      bankRows.push(r); bankTotal += o.amount; otherTotal += o.amount;
      if (!o.civil) otherNoCivil.push(o.name);
      refRows.push({ loc: o.loc, id: '', pay: o.amount, vals: ['', o.name, o.loc, o.civil, 'Bank', o.iban] });
    });
    otherTotal = Math.round(otherTotal * 1000) / 1000;
    bankTotal = Math.round(bankTotal * 1000) / 1000; cashTotal = Math.round(cashTotal * 1000) / 1000;

    const periodLabel = payrollTab.replace(/^Payroll\s*/i, '').trim() || payrollTab;

    // --- 5. Bank sheet for Warba ---
    const bankOut = 'Bank Transfer ' + periodLabel;
    // Re-used, not deleted: clearing keeps the tab's link (gid) working.
    let bs = ss.getSheetByName(bankOut);
    if (bs) bs.clear(); else bs = ss.insertSheet(bankOut);
    bs.getRange(1, 1, bankRows.length + 2, tplHeaders.length).setNumberFormat('@');
    bs.getRange(1, 1, 1, tplHeaders.length).setValues([tplHeaders]).setFontWeight('bold');
    if (bankRows.length) {
      bs.getRange(2, 1, bankRows.length, tplHeaders.length).setValues(bankRows);
      // Total of the transfer in the amount column, last row (other cells blank).
      if (bAmt !== -1) {
        const totRow = new Array(tplHeaders.length).fill('');
        totRow[bAmt] = bankTotal;
        bs.getRange(bankRows.length + 2, 1, 1, tplHeaders.length).setValues([totRow]).setFontWeight('bold');
        bs.getRange(2, bAmt + 1, bankRows.length + 1, 1).setNumberFormat('0.000');
      }
    }
    bs.setFrozenRows(1); bs.autoResizeColumns(1, tplHeaders.length);

    // --- 6. Accounts sheet: the Warba columns + Location (cost centre) ---
    // Every paid employee — bank AND cash — so each location carries its
    // full salary cost. Same columns as the bank sheet, then Location and
    // Payment. Sorted by Location; a per-location total block underneath.
    const refOut = 'Payroll by Location ' + periodLabel;
    let rs = ss.getSheetByName(refOut);
    if (rs) rs.clear(); else rs = ss.insertSheet(refOut);
    const oldRef = ss.getSheetByName('Payroll Reference ' + periodLabel); if (oldRef) ss.deleteSheet(oldRef);
    const refHeaders = tplHeaders.concat(['Location', 'Payment']);
    const W = refHeaders.length, cLoc = tplHeaders.length, cPay = tplHeaders.length + 1;
    const acct = refRows.filter(r => r.vals[4] === 'Bank' || r.vals[4] === 'Cash').map(r => {
      const o = new Array(W).fill('');
      if (bCivil !== -1) o[bCivil] = r.vals[3];
      if (bName !== -1) o[bName] = r.vals[1];
      if (bIban !== -1) o[bIban] = r.vals[5];
      if (bAmt !== -1) o[bAmt] = r.pay;
      o[cLoc] = r.loc || '(no location)';
      o[cPay] = r.vals[4];
      return o;
    });
    acct.sort((a, b) => String(a[cLoc]).localeCompare(String(b[cLoc])) || String(a[bName]).localeCompare(String(b[bName])));
    const byLoc = {};
    acct.forEach(o => { const k = o[cLoc]; byLoc[k] = byLoc[k] || { n: 0, bank: 0, cash: 0 }; byLoc[k].n++; byLoc[k][o[cPay] === 'Bank' ? 'bank' : 'cash'] += Number(o[bAmt]) || 0; });
    const r3 = n => Math.round(n * 1000) / 1000;
    const block = [new Array(W).fill(''), ['Location', 'Staff', 'Bank (KD)', 'Cash (KD)', 'Total (KD)'].concat(new Array(Math.max(0, W - 5)).fill(''))];
    let gN = 0, gB = 0, gC = 0;
    Object.keys(byLoc).sort().forEach(k => { const v = byLoc[k]; gN += v.n; gB += v.bank; gC += v.cash;
      block.push([k, v.n, r3(v.bank), r3(v.cash), r3(v.bank + v.cash)].concat(new Array(Math.max(0, W - 5)).fill(''))); });
    block.push(['TOTAL', gN, r3(gB), r3(gC), r3(gB + gC)].concat(new Array(Math.max(0, W - 5)).fill('')));
    const refData = [refHeaders].concat(acct).concat(block.map(r => r.slice(0, W)));
    if (bCivil !== -1) rs.getRange(1, bCivil + 1, refData.length, 1).setNumberFormat('@');
    if (bIban !== -1) rs.getRange(1, bIban + 1, refData.length, 1).setNumberFormat('@');
    rs.getRange(1, 1, refData.length, W).setValues(refData);
    if (bAmt !== -1) rs.getRange(2, bAmt + 1, acct.length, 1).setNumberFormat('0.000');
    rs.getRange(1, 1, 1, W).setFontWeight('bold');
    rs.getRange(acct.length + 3, 1, 1, W).setFontWeight('bold');
    rs.getRange(refData.length, 1, 1, W).setFontWeight('bold');
    rs.setFrozenRows(1); rs.autoResizeColumns(1, W);

    const payrollNet = Math.round(payRows.reduce((t, p) => t + p.pay, 0) * 1000) / 1000;
    const reconciled = Math.abs(bankTotal - otherTotal + cashTotal - payrollNet) < 0.001;
    return {
      success: true, tab: payrollTab, bankTab: bankOut, referenceTab: refOut,
      written: bankRows.length, totalTransfer: bankTotal, cashTotal: cashTotal, payrollNet: payrollNet,
      reconciled: reconciled, cash: cash, noCivilId: noCivil,
      otherPayees: others.list.length, otherTotal: otherTotal,
      message: 'Bank sheet "' + bankOut + '": ' + bankRows.length + ' rows, ' + bankTotal.toFixed(3) + ' KD' +
        (others.list.length ? ' (staff ' + (bankRows.length - others.list.length) + ' / ' + Math.round((bankTotal - otherTotal) * 1000) / 1000 +
          ' KD + other payees ' + others.list.length + ' / ' + otherTotal.toFixed(3) + ' KD)' : '') + '. ' +
        'Cash: ' + cash.length + ' staff, ' + cashTotal.toFixed(3) + ' KD' + (cash.length ? ' (' + cash.join(', ') + ')' : '') + '. ' +
        'Accounts sheet "' + refOut + '" (by Location). ' +
        (reconciled ? 'Bank + cash = payroll net ' + payrollNet.toFixed(3) + ' KD.' : '⚠ Bank + cash does NOT equal payroll net ' + payrollNet.toFixed(3) + ' KD.') +
        (noCivil.length ? ' ⚠ No Civil ID, paid as cash: ' + noCivil.join(', ') + '.' : '') +
        (otherNoCivil.length ? ' ⚠ Other payees with no Civil ID in ' + OTHER_PAYEES_TAB + ': ' + otherNoCivil.join(', ') + '.' : '') +
        (others.skipped.length ? ' ⚠ Other payees skipped: ' + others.skipped.join(', ') + '.' : '')
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}