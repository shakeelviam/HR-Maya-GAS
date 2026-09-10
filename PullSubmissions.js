// ============================================================
// PullSubmissions.gs — refresh HR Maya's "Staff Submission" tab
// from the Staff Docs Collection workbook's "Submissions" tab.
// ------------------------------------------------------------
// Full REPLACE: the tab always mirrors the current source. Then run
// syncStaffSubmissions to push the data into HR Maya (by Civil ID).
//
// Flow:  pullSubmissionsFromDocs  ->  syncStaffSubmissions
//
// ENDPOINT: pullSubmissionsFromDocs   (?tab=Submissions to override)
// ============================================================

// Staff Docs Collection workbook (source of form responses).
const STAFF_DOCS_SPREADSHEET_ID = '1FVbLjlKxdcXSOpwMD8Z8KVkAjXReSeS5kzH3_2zt3_M';
const STAFF_DOCS_SUBMISSIONS_TAB = 'Submissions';

function pullSubmissionsFromDocs(params) {
  try {
    params = params || {};
    const srcTabName = String(params.tab || STAFF_DOCS_SUBMISSIONS_TAB).trim();

    const src = SpreadsheetApp.openById(STAFF_DOCS_SPREADSHEET_ID);
    if (!src) return { success: false, error: 'Could not open Staff Docs workbook.' };
    const srcSheet = src.getSheetByName(srcTabName);
    if (!srcSheet) {
      return { success: false, error: 'Source tab "' + srcTabName + '" not found. Available: ' + src.getSheets().map(s => s.getName()).join(', ') };
    }

    const values = srcSheet.getDataRange().getValues();
    if (!values.length) return { success: false, error: 'Source "' + srcTabName + '" is empty.' };

    const dest = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    let destSheet = dest.getSheetByName(STAFF_SUBMISSION_TAB); // 'Staff Submission' (from StaffSync.gs)
    if (!destSheet) destSheet = dest.insertSheet(STAFF_SUBMISSION_TAB);

    // Full replace: clear then write the source verbatim.
    destSheet.clear();
    destSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    destSheet.setFrozenRows(1);

    // Keep Civil ID / phone-like columns as text so they don't get mangled.
    const H = values[0];
    const textCols = ['Civil ID', 'Mobile', 'WhatsApp', 'Passport No'];
    textCols.forEach(name => {
      const idx = H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
      if (idx !== -1 && values.length > 1) destSheet.getRange(2, idx + 1, values.length - 1, 1).setNumberFormat('@');
    });

    const rows = values.length - 1;
    return {
      success: true, sourceTab: srcTabName, rowsCopied: rows,
      message: 'Replaced "' + STAFF_SUBMISSION_TAB + '" with ' + rows + ' rows from the Staff Docs "' + srcTabName + '" tab. Now run syncStaffSubmissions to push into HR Maya.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Convenience: pull then sync in one call.
function refreshAndSyncSubmissions(params) {
  const pull = pullSubmissionsFromDocs(params);
  if (!pull.success) return pull;
  const sync = syncStaffSubmissions(params || {});
  return {
    success: sync.success,
    pull: pull.message,
    sync: sync.message || sync.error,
    message: pull.message + '  •  ' + (sync.message || ('SYNC ERROR: ' + sync.error))
  };
}

// Fetch Designation from the Roster tab into the Submissions tab (both in the
// Staff Docs workbook), matching by Master Row. Run BEFORE pullSubmissionsFromDocs
// so the designation then flows Roster -> Submissions -> HR Maya.
// params.rosterTab (default 'Roster'), params.subTab (default 'Submissions')
function pullDesignationsFromRoster(params) {
  try {
    params = params || {};
    const preview = String(params.preview).toLowerCase() === 'true';
    const rosterTab = String(params.rosterTab || 'Roster').trim();
    const src = SpreadsheetApp.openById(STAFF_DOCS_SPREADSHEET_ID);
    const roster = src.getSheetByName(rosterTab);
    if (!roster) return { success: false, error: 'Roster tab "' + rosterTab + '" not found. Tabs: ' + src.getSheets().map(s => s.getName()).join(', ') };

    // Roster: Civil ID -> Designation (Roster is the source of truth; no change to the Docs script).
    const rd = roster.getDataRange().getValues();
    const rH = rd[0];
    const rc = (n) => rH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const rCivil = rc('Civil ID'), rDesig = rc('Designation'), rName = rc('Name');
    if (rDesig === -1) return { success: false, error: 'Roster has no "Designation" column.' };
    const norm = (v) => String(v == null ? '' : v).replace(/[^0-9]/g, '');
    const desigByCivil = {}, desigByName = {};
    for (let i = 1; i < rd.length; i++) {
      const dsg = String(rd[i][rDesig] || '').trim();
      if (!dsg) continue;
      if (rCivil !== -1) { const c = norm(rd[i][rCivil]); if (c) desigByCivil[c] = dsg; }
      if (rName !== -1) { const nm = String(rd[i][rName] || '').trim().toLowerCase().replace(/\\s+/g,' '); if (nm) desigByName[nm] = dsg; }
    }

    // HR Maya: ensure Designation column, fill by Civil ID (fallback unique name).
    const hr = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('HR Maya');
    let H = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0];
    const hc = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    let cDes = hc('Designation');
    if (cDes === -1 && !preview) { hr.getRange(1, hr.getLastColumn() + 1).setValue('Designation'); H = hr.getRange(1,1,1,hr.getLastColumn()).getValues()[0]; cDes = hc('Designation'); }
    const cCivil = hc('Civil ID'), cName = hc('Name (English)');

    const data = hr.getDataRange().getValues();
    let filled = 0, noMatch = 0;
    for (let i = 1; i < data.length; i++) {
      const civ = cCivil !== -1 ? norm(data[i][cCivil]) : '';
      const nm = cName !== -1 ? String(data[i][cName] || '').trim().toLowerCase().replace(/\\s+/g,' ') : '';
      let dsg = (civ && desigByCivil[civ]) ? desigByCivil[civ] : (nm && desigByName[nm] ? desigByName[nm] : '');
      if (!dsg) { if (String(data[i][0]||'').trim()) noMatch++; continue; }
      if (!preview && cDes !== -1) hr.getRange(i + 1, cDes + 1).setValue(dsg);
      filled++;
    }
    return { success: true, preview: preview, filled: filled, noMatch: noMatch,
      message: (preview ? 'PREVIEW — ' : '') + 'Designation set on ' + filled + ' HR Maya employee(s) from Roster (by Civil ID / name)' + (noMatch ? ', ' + noMatch + ' with no roster designation' : '') + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Convenience: designations from roster -> submissions -> HR Maya, all in one.
function refreshDesignationsFull(params) {
  // Designation now flows Roster -> HR Maya directly (by Civil ID). One step.
  return pullDesignationsFromRoster(params);
}

// Distinct designations (from Roster; falls back to HR Maya) for the dropdown.
function getDesignations(params) {
  try {
    const set = {};
    // Prefer Roster (source of truth).
    try {
      const src = SpreadsheetApp.openById(STAFF_DOCS_SPREADSHEET_ID);
      const roster = src.getSheetByName('Roster');
      if (roster) {
        const d = roster.getDataRange().getValues();
        const cD = d[0].findIndex(h => String(h).trim().toLowerCase() === 'designation');
        for (let i = 1; i < d.length; i++) { const v = String(d[i][cD] || '').trim(); if (v) set[v] = true; }
      }
    } catch (e) {}
    // Also include any already in HR Maya.
    try {
      const hr = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('HR Maya');
      const d = hr.getDataRange().getValues();
      const cD = d[0].findIndex(h => String(h).trim().toLowerCase() === 'designation');
      if (cD !== -1) for (let i = 1; i < d.length; i++) { const v = String(d[i][cD] || '').trim(); if (v) set[v] = true; }
    } catch (e) {}
    return { success: true, data: Object.keys(set).sort() };
  } catch (err) { return { success: false, error: err.toString() }; }
}