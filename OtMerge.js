// ============================================================
// OtMerge.gs — one-time: fold Missing OT into a copy of a payroll run
// ============================================================
// Reads a payroll tab + the Missing OT tab, and writes a TEMP copy of the
// payroll tab where each employee's OT Hours and OT Amount have the missing
// OT added in (Total Payable recomputed = +missing amount). Original run is
// left untouched. Join by Employee ID.
//
// ENDPOINT: buildOtMergeSheet   (?payrollTab=... &missingTab=Missing OT)
// ============================================================

function buildOtMergeSheet(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const payrollTabName = String(params.payrollTab || 'Payroll 21Aug-21Aug').trim();
    const missingTabName = String(params.missingTab || 'Missing OT').trim();

    const pay = ss.getSheetByName(payrollTabName);
    if (!pay) return { success: false, error: 'Payroll tab "' + payrollTabName + '" not found.' };
    const miss = ss.getSheetByName(missingTabName);
    if (!miss) return { success: false, error: 'Missing OT tab "' + missingTabName + '" not found.' };

    // ---- Read Missing OT: Employee ID -> { hours, amount } ----
    const md = miss.getDataRange().getValues();
    const mH = md[0];
    const mc = (cands) => { for (const n of cands) { const i = mH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase()); if (i !== -1) return i; } return -1; };
    const mId = mc(['Employee ID', 'Emp ID', 'ID']);
    // missing hours: the "Missing ..." column; missing amount: "OT Amount"/"Amount"
    let mHrs = mc(['Missing (22 Aug onward)', 'Missing Hours', 'Missing OT', 'Missing']);
    if (mHrs === -1) mHrs = mH.findIndex(h => String(h).trim().toLowerCase().indexOf('missing') === 0);
    const mAmt = mc(['OT Amount', 'Missing OT Amount', 'Amount']);
    if (mId === -1 || mHrs === -1) return { success: false, error: 'Missing OT tab needs an Employee ID column and a "Missing…" hours column.' };

    const missByEmp = {};
    for (let i = 1; i < md.length; i++) {
      const id = String(md[i][mId] || '').trim();
      if (!id) continue;
      missByEmp[id] = { hours: Number(md[i][mHrs]) || 0, amount: mAmt !== -1 ? (Number(md[i][mAmt]) || 0) : 0 };
    }

    // ---- Read payroll tab ----
    const pd = pay.getDataRange().getValues();
    const pH = pd[0];
    const pc = (cands) => { for (const n of cands) { const i = pH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase()); if (i !== -1) return i; } return -1; };
    const pId = pc(['Employee ID', 'Emp ID', 'ID']);
    const pOtH = pc(['OT Hours']);
    const pOtA = pc(['OT Amount']);
    const pPay = pc(['Total Payable', 'Net Payable']);
    if (pId === -1 || pOtH === -1 || pOtA === -1) return { success: false, error: 'Payroll tab needs Employee ID, OT Hours, OT Amount columns.' };

    // ---- Build the merged copy ----
    const out = pd.map(r => r.slice()); // deep-ish copy of rows
    const KD3 = (n) => Math.round((Number(n) || 0) * 1000) / 1000;
    let updated = 0, addedHours = 0, addedAmount = 0; const noMatch = [];
    const usedEmp = {};
    for (let i = 1; i < out.length; i++) {
      const id = String(out[i][pId] || '').trim();
      if (!id) continue;
      const m = missByEmp[id];
      if (!m) continue;
      usedEmp[id] = true;
      const oldH = Number(out[i][pOtH]) || 0;
      const oldA = Number(out[i][pOtA]) || 0;
      out[i][pOtH] = KD3(oldH + m.hours);
      out[i][pOtA] = KD3(oldA + m.amount);
      if (pPay !== -1) out[i][pPay] = KD3((Number(out[i][pPay]) || 0) + m.amount);
      updated++; addedHours += m.hours; addedAmount += m.amount;
    }
    // Missing-OT people with no payroll row
    Object.keys(missByEmp).forEach(id => { if (!usedEmp[id] && (missByEmp[id].hours || missByEmp[id].amount)) noMatch.push(id); });

    // ---- Write the temp tab ----
    const outName = payrollTabName + ' (OT-fixed)';
    let dst = ss.getSheetByName(outName);
    if (dst) ss.deleteSheet(dst);
    dst = ss.insertSheet(outName);
    dst.getRange(1, 1, out.length, out[0].length).setValues(out);
    dst.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
    dst.setFrozenRows(1);

    return {
      success: true, tab: outName, updated: updated,
      addedHours: Math.round(addedHours * 100) / 100, addedAmount: KD3(addedAmount),
      unmatched: noMatch,
      message: 'Built "' + outName + '": added missing OT to ' + updated + ' employee(s) — +' +
        (Math.round(addedHours * 100) / 100) + ' hrs, +' + KD3(addedAmount).toFixed(3) + ' KD. Original run untouched.' +
        (noMatch.length ? ' Missing-OT with no payroll row: ' + noMatch.join(', ') : '')
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}