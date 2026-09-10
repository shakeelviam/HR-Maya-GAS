// ============================================================
// StaffSync.gs — Maya HR : sync "Staff Submission" tab into HR Maya
// ============================================================
// Matches each submission to an HR Maya employee by CIVIL ID (digits-only),
// takes the MOST RECENT submission per person (by Timestamp), and writes the
// document fields into HR Maya. Rule: the submission value wins when present
// (fills blanks AND corrects wrong values), but a blank submission field
// never wipes an existing HR Maya value. Missing HR Maya columns are created.
//
// ENDPOINT (add to Code.gs): syncStaffSubmissions
//   ?method=syncStaffSubmissions            (commit)
//   ?method=syncStaffSubmissions&preview=true   (report only, no writes)
// ============================================================

const STAFF_SUBMISSION_TAB = 'Staff Submission';

// Submission column header -> HR Maya column header.
// NOTE: Civil ID is included so a wrong HR Maya Civil ID gets CORRECTED
// when the row is matched by name (the whole point of the sync).
const SUBMISSION_MAP = {
  'Civil ID': 'Civil ID',
  'Email': 'Email',
  'Mobile': 'Mobile',
  'WhatsApp': 'WhatsApp',
  'CID Expiry': 'CID Expiry',
  'Passport No': 'Passport No',
  'Passport Issue': 'Passport Issue',
  'Passport Expiry': 'Passport Expiry',
  'Entry Date': 'Entry Date',
  'Joining Date': 'Date of Join',
  'Health Issue': 'Health Issue',
  'Health Expiry': 'Health Expiry',
  'Civil ID Copy': 'Civil ID Copy',
  'Civil ID Back Copy': 'Civil ID Back Copy',
  'Passport Entry Page': 'Passport Entry Page'
};
// Which target fields are dates (formatted dd-mm-yyyy on write).
const SUBMISSION_DATE_FIELDS = {
  'CID Expiry': 1, 'Passport Issue': 1, 'Passport Expiry': 1,
  'Entry Date': 1, 'Date of Join': 1, 'Health Issue': 1, 'Health Expiry': 1
};

function syncStaffSubmissions(params) {
  try {
    params = params || {};
    const preview = String(params.preview || '').toLowerCase() === 'true';
    const subTabName = String(params.tab || STAFF_SUBMISSION_TAB).trim();
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    const sub = ss.getSheetByName(subTabName);
    if (!sub) return { success: false, error: 'Tab "' + subTabName + '" not found.' };
    const hr = ss.getSheetByName('HR Maya');
    if (!hr) return { success: false, error: "'HR Maya' tab not found." };

    // --- Ensure all target columns exist in HR Maya (append missing) ---
    let hH = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0];
    const hcOf = (name) => hH.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
    const createdCols = [];
    if (!preview) {
      Object.keys(SUBMISSION_MAP).forEach(subCol => {
        const target = SUBMISSION_MAP[subCol];
        if (hcOf(target) === -1) {
          hr.getRange(1, hr.getLastColumn() + 1).setValue(target).setFontWeight('bold');
          hH = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0];
          createdCols.push(target);
        }
      });
    }

    // --- HR Maya: Civil ID -> row, and Name -> [rows] (for fallback) ---
    const hd = hr.getDataRange().getValues();
    const hcId = hcOf('Employee ID'), hcCivil = hcOf('Civil ID'), hcName = hcOf('Name (English)');
    const rowByCivil = {};
    const rowsByName = {};
    const normName_ = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
    for (let i = 1; i < hd.length; i++) {
      if (!String(hd[i][hcId] || '').trim()) continue;
      const civ = normCivil_(hd[i][hcCivil]);
      if (civ) rowByCivil[civ] = i + 1;
      const nm = normName_(hd[i][hcName]);
      if (nm) { (rowsByName[nm] = rowsByName[nm] || []).push(i + 1); }
    }

    // --- Submissions: keep most recent per person (key by Civil ID, else name) ---
    const sd = sub.getDataRange().getValues();
    const sH = sd[0];
    const sc = (name) => sH.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
    const scTs = sc('Timestamp'), scCivil = sc('Civil ID'), scName = sc('Name');
    const latest = {}; // key -> { row, ts }
    for (let i = 1; i < sd.length; i++) {
      const civ = normCivil_(sd[i][scCivil]);
      const nm = normName_(sd[i][scName]);
      const key = civ || ('name:' + nm);
      if (!civ && !nm) continue;
      const ts = sd[i][scTs] instanceof Date ? sd[i][scTs].getTime() : (parseFlexibleDate_(sd[i][scTs]) ? parseFlexibleDate_(sd[i][scTs]).getTime() : 0);
      if (!latest[key] || ts >= latest[key].ts) latest[key] = { row: sd[i], ts: ts };
    }

    // --- Write into HR Maya ---
    let matched = 0, cellsUpdated = 0, byName = 0, civilCorrected = 0;
    const unmatched = [], updatedPeople = [], corrections = [];
    Object.keys(latest).forEach(key => {
      const srow = latest[key].row;
      const subName = scName !== -1 ? String(srow[scName] || '').trim() : '';
      const subCivil = normCivil_(srow[scCivil]);

      // 1) match by Civil ID; 2) fall back to a UNIQUE name match.
      let rowNum = subCivil ? rowByCivil[subCivil] : undefined;
      let matchedByName = false;
      if (!rowNum) {
        const cands = rowsByName[normName_(subName)];
        if (cands && cands.length === 1) { rowNum = cands[0]; matchedByName = true; }
        else if (cands && cands.length > 1) { unmatched.push(subName + ' / ' + subCivil + ' (name matches ' + cands.length + ' rows — fix Civil ID manually)'); return; }
      }
      if (!rowNum) { unmatched.push(subName + ' / ' + (subCivil || 'no civil id') + ' (no Civil ID or name match)'); return; }

      matched++; if (matchedByName) byName++;
      const hrCivilBefore = normCivil_(hd[rowNum - 1][hcCivil]);

      let fieldsHere = 0;
      Object.keys(SUBMISSION_MAP).forEach(subCol => {
        const sci = sc(subCol);
        if (sci === -1) return;
        let val = srow[sci];
        if (val === null || val === undefined || String(val).trim() === '') return; // blank never wipes
        const target = SUBMISSION_MAP[subCol];
        const tci = hcOf(target);
        if (tci === -1) return;
        if (SUBMISSION_DATE_FIELDS[target]) {
          const d = (val instanceof Date) ? val : parseFlexibleDate_(val);
          val = d ? formatDdMmYyyy_(d) : String(val).trim();
        } else { val = String(val).trim(); }
        if (!preview) {
          const cell = hr.getRange(rowNum, tci + 1);
          cell.setValue(val);
          if (target === 'Civil ID' || target === 'Mobile' || target === 'WhatsApp' || String(target).indexOf('Passport No') !== -1) cell.setNumberFormat('@');
        }
        cellsUpdated++; fieldsHere++;
      });

      // Note a Civil ID correction (matched by name, and the IDs differed).
      if (subCivil && hrCivilBefore && subCivil !== hrCivilBefore) {
        civilCorrected++;
        corrections.push(subName + ': ' + hrCivilBefore + ' \u2192 ' + subCivil);
      }
      if (fieldsHere) updatedPeople.push(subName + (matchedByName ? ' [by name]' : '') + ' (' + fieldsHere + ' fields)');
    });

    return {
      success: true, preview: preview, tab: subTabName,
      uniquePeople: Object.keys(latest).length, matched: matched, matchedByName: byName,
      civilIdsCorrected: civilCorrected, corrections: corrections,
      unmatchedCount: unmatched.length, cellsUpdated: cellsUpdated, columnsCreated: createdCols, unmatched: unmatched,
      message: (preview ? 'PREVIEW — no writes. ' : '') +
        'Matched ' + matched + ' of ' + Object.keys(latest).length + ' submitters (' + byName + ' by name fallback)' +
        (preview ? '' : (', updated ' + cellsUpdated + ' cells')) +
        (civilCorrected ? '. Corrected ' + civilCorrected + ' Civil ID(s): ' + corrections.join('; ') : '') +
        (createdCols.length ? '. Created columns: ' + createdCols.join(', ') : '') +
        (unmatched.length ? '. Still unmatched: ' + unmatched.join(', ') : '.')
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}