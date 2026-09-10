// ============================================================
// MissingOT.gs — one-time: price the "Missing OT" tab
// ------------------------------------------------------------
// Reads a "Missing OT" tab, looks up each employee's Price Per Hour
// from HR Maya (by Employee ID), and writes OT Amount = hours x rate
// (KD, 3 dp). Columns resolved by header NAME, so the tab's exact
// layout doesn't matter as long as it has an Employee ID column, an
// hours column, and an amount column.
//
// Recognised headers (case-insensitive, first match wins):
//   Employee ID : "Employee ID" / "Emp ID" / "ID"
//   Hours       : "Missing Hours" / "Missing OT" / "OT Hours" / "Hours"
//   Amount      : "OT Amount" / "Amount" / "Missing OT Amount"  (created if absent)
//   Rate (opt.) : "Price Per Hour" / "Rate"  (else pulled from HR Maya)
//
// Run once from the editor: computeMissingOt   (add &tab=... to override name)
// ============================================================

function computeMissingOt(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const tabName = String(params.tab || 'Missing OT').trim();
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return { success: false, error: 'Tab "' + tabName + '" not found.' };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'Tab "' + tabName + '" has no data rows.' };
    let H = data[0];
    const find = (cands) => {
      for (const cand of cands) {
        const i = H.findIndex(h => String(h).trim().toLowerCase() === cand.toLowerCase());
        if (i !== -1) return i;
      }
      return -1;
    };
    const cId = find(['Employee ID', 'Emp ID', 'ID']);
    let cHrs = find(['Missing (22 Aug onward)', 'Missing Hours', 'Missing OT Hours', 'Missing OT', 'OT Hours', 'Hours']);
    // Fallback: any header starting with "Missing" that isn't the "Paid" column.
    if (cHrs === -1) cHrs = H.findIndex(h => String(h).trim().toLowerCase().indexOf('missing') === 0);
    let cRate = find(['Price Per Hour', 'Rate']);
    let cAmt = find(['OT Amount', 'Missing OT Amount', 'Amount']);

    if (cId === -1) return { success: false, error: 'Could not find an Employee ID column in "' + tabName + '".' };
    if (cHrs === -1) return { success: false, error: 'Could not find an hours column (e.g. "Missing Hours") in "' + tabName + '".' };

    // Create an OT Amount column if the tab doesn't have one.
    if (cAmt === -1) {
      cAmt = sheet.getLastColumn();
      sheet.getRange(1, cAmt + 1).setValue('OT Amount').setFontWeight('bold');
      cAmt = cAmt; // 0-based index of the new column == old lastColumn
    }

    // HR Maya: Employee ID -> Price Per Hour (fallback for rate).
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const hH = hd[0];
    const hcId = hH.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const hcRate = hH.findIndex(h => String(h).trim().toLowerCase() === 'price per hour');
    const hcBasic = hH.findIndex(h => String(h).trim().toLowerCase() === 'basic salary');
    const rateByEmp = {};
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][hcId] || '').trim();
      if (!id) continue;
      let rate = hcRate !== -1 ? Number(hd[i][hcRate]) || 0 : 0;
      if (!rate && hcBasic !== -1) rate = (Number(hd[i][hcBasic]) || 0) / 26 / 8; // fallback compute
      rateByEmp[id] = Math.round(rate * 1000) / 1000;
    }

    let written = 0, totalHours = 0, totalAmount = 0;
    const notFound = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][cId] || '').trim();
      if (!id) continue;
      const hours = Number(data[i][cHrs]) || 0;
      if (hours === 0 && String(data[i][cHrs]).trim() === '') continue;
      let rate = (cRate !== -1 && Number(data[i][cRate])) ? Number(data[i][cRate]) : rateByEmp[id];
      if (rate === undefined) { notFound.push(id); rate = 0; }
      const amount = Math.round(hours * rate * 1000) / 1000;
      sheet.getRange(i + 1, cAmt + 1).setValue(amount).setNumberFormat('0.000');
      if (cRate !== -1 && !Number(data[i][cRate]) && rate) sheet.getRange(i + 1, cRate + 1).setValue(rate).setNumberFormat('0.000');
      written++; totalHours += hours; totalAmount += amount;
    }
    totalHours = Math.round(totalHours * 1000) / 1000;
    totalAmount = Math.round(totalAmount * 1000) / 1000;

    return {
      success: true, tab: tabName, rowsPriced: written,
      totalHours: totalHours, totalAmount: totalAmount, unmatched: notFound,
      message: 'Priced ' + written + ' row(s) in "' + tabName + '": ' + totalHours + ' hours = ' + totalAmount.toFixed(3) + ' KD.' +
        (notFound.length ? ' ⚠ No HR Maya rate for: ' + notFound.join(', ') + ' (amount set to 0).' : '')
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}