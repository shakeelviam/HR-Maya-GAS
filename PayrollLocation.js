// ============================================================
// PayrollLocation.gs — add a Location column to a payroll tab
// ------------------------------------------------------------
// Payroll tabs are written without Location (see the headers array in
// runPayrollFromReview). This inserts a Location column, populated from
// HR Maya, so a run can be read or filtered per branch.
//
// USAGE — from the Apps Script editor:
//   1. Select addLocationToPayrollTab from the function dropdown
//   2. Run
//   It defaults to the tab named in TARGET_TAB below. Change that string
//   to point at a different run.
//
// Safe to re-run: if a Location column already exists it is repopulated,
// not duplicated. Nothing else on the sheet is touched.
//
// NOTE ON WHICH LOCATION THIS IS: HR Maya holds an employee's CURRENT
// location. If someone transferred after this payroll period, this shows
// where they are NOW, not where they worked during the period. For a
// historical per-period location you would need Attendance_Daily, which
// records the branch each day was actually worked at.
// ============================================================

const TARGET_TAB = 'Payroll 21Aug-21Aug (OT-fixed)';

function addLocationToPayrollTab() {
  const tabName = TARGET_TAB;
  const res = addLocationToTab_(tabName);
  Logger.log(res.message || res.error);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(res.message || res.error, 'Location column', 8);
  } catch (e) { /* not bound to a sheet UI — log only */ }
  return res;
}

// Callable version — pass a tab name, e.g. from the router or another script.
function addLocationToTab_(tabName) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return { success: false, error: 'Tab not found: "' + tabName + '"' };
    if (sheet.getLastRow() < 2) return { success: false, error: 'Tab "' + tabName + '" has no data rows.' };

    // ---- Build empId -> location from HR Maya ----
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found." };
    const hd = hr.getDataRange().getValues();
    const HH = hd[0];
    const hcol = (n) => HH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const hId = hcol('Employee ID');
    let hLoc = hcol('Location');
    if (hLoc === -1) hLoc = hcol('Branch');
    if (hId === -1) return { success: false, error: "HR Maya has no 'Employee ID' column." };
    if (hLoc === -1) return { success: false, error: "HR Maya has no 'Location' or 'Branch' column." };

    const locById = {};
    for (let i = 1; i < hd.length; i++) {
      const id = String(hd[i][hId] || '').trim();
      if (id) locById[id] = String(hd[i][hLoc] || '').trim();
    }

    // ---- Find the payroll tab's Employee ID column ----
    const lastCol = sheet.getLastColumn();
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const pcol = (n) => header.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const pId = pcol('Employee ID');
    if (pId === -1) return { success: false, error: '"' + tabName + '" has no \'Employee ID\' column.' };

    // ---- Where does Location go? ----
    // Re-run safe: reuse the existing column if there is one, otherwise
    // insert a new one immediately after Name (English) so it reads naturally,
    // falling back to the end of the sheet.
    let locCol = pcol('Location');
    let inserted = false;
    if (locCol === -1) {
      const pName = pcol('Name (English)');
      const insertAfter = pName !== -1 ? pName + 1 : lastCol;
      sheet.insertColumnAfter(insertAfter);
      locCol = insertAfter;            // 0-based index of the new column
      sheet.getRange(1, locCol + 1).setValue('Location').setFontWeight('bold');
      inserted = true;
    }

    // ---- Populate ----
    const n = sheet.getLastRow() - 1;
    const ids = sheet.getRange(2, pId + 1, n, 1).getValues();
    const out = [];
    let matched = 0, unmatched = 0;
    const missing = [];

    ids.forEach(r => {
      const id = String(r[0] || '').trim();
      if (!id) { out.push(['']); return; }
      const loc = locById[id];
      if (loc) { out.push([loc]); matched++; }
      else { out.push(['(not found)']); unmatched++; missing.push(id); }
    });

    sheet.getRange(2, locCol + 1, n, 1).setValues(out);

    let msg = (inserted ? 'Added' : 'Refreshed') + ' Location on "' + tabName + '" — ' +
              matched + ' matched';
    if (unmatched) {
      msg += ', ' + unmatched + ' not found in HR Maya (' +
             missing.slice(0, 8).join(', ') + (missing.length > 8 ? '…' : '') + ')';
    }
    msg += '.';

    return { success: true, message: msg, matched: matched, unmatched: unmatched, missingIds: missing };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}