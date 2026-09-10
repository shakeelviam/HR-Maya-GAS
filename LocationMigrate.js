// ============================================================
// LocationMigrate.gs — rename Branch→Location + bulk re-tag + supervisors
// ============================================================
// One-time: renames HR Maya "Branch" column to "Location" and re-tags the
// messy Arabic values to clean location names. Retired-branch staff are
// BLANKED and reported. Also seeds Central Kitchen as geofenced at Qurain's
// coordinates, and adds a Supervisors column to Branch_Locations.
//
// ENDPOINTS: migrateBranchToLocation, getLocationSupervisors,
//            saveLocationSupervisors
// ============================================================

// Arabic branch value -> clean Location. Values not listed are BLANKED.
const BRANCH_TO_LOCATION_MAP = {
  'الجابرية': 'Jabriya',
  'العارضية': 'Ardiya',
  'القرين': 'Qurain',
  'كيفان': 'Kaifan',
  'الخيران': 'Khairan',
  'مخزن اسواق القرين': 'Central Kitchen'
  // 'السعودية' and 'كل الفرع ماعدا الخيران' -> blanked (retired)
};

function normBranchVal_(s) { return String(s || '').trim().replace(/\s+/g, ' '); }

function migrateBranchToLocation(params) {
  try {
    params = params || {};
    const preview = String(params.preview).toLowerCase() === 'true';
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const H = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0];
    let cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cLocation = H.findIndex(h => String(h).trim().toLowerCase() === 'location');

    // Resolve the column to use as "Location" going forward.
    let colIndex;
    if (cLocation !== -1) colIndex = cLocation;          // already renamed
    else if (cBranch !== -1) colIndex = cBranch;         // rename this one
    else return { success: false, error: 'HR Maya has neither Branch nor Location column.' };

    if (!preview && cLocation === -1 && cBranch !== -1) {
      hr.getRange(1, cBranch + 1).setValue('Location'); // rename header
    }

    const data = hr.getDataRange().getValues();
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');

    let mapped = 0, blanked = 0, already = 0;
    const blankedList = [], unmapped = {};
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][cId] || '').trim();
      if (!id) continue;
      const cur = normBranchVal_(data[i][colIndex]);
      if (!cur) continue;
      // Already a clean location?
      const cleanNames = Object.values(BRANCH_TO_LOCATION_MAP).concat(['Head Office']);
      if (cleanNames.indexOf(cur) !== -1) { already++; continue; }
      const target = BRANCH_TO_LOCATION_MAP[cur];
      if (target) {
        if (!preview) hr.getRange(i + 1, colIndex + 1).setValue(target);
        mapped++;
      } else {
        // retired / unknown -> blank + report
        if (!preview) hr.getRange(i + 1, colIndex + 1).setValue('');
        blanked++;
        blankedList.push((cName !== -1 ? String(data[i][cName]).trim() : id) + ' (' + id + ', was "' + cur + '")');
        unmapped[cur] = (unmapped[cur] || 0) + 1;
      }
    }

    // Ensure Central Kitchen is geofenced at Qurain's coordinates.
    if (!preview) {
      try {
        saveLocation({ name: 'Central Kitchen', geofenced: 'Yes', latitude: 29.198861, longitude: 48.052833, radius: 50, active: 'Yes' });
        deleteLocation({ name: 'السعودية' }); // retire if present
      } catch (e) {}
    }

    return {
      success: true, preview: preview,
      renamed: (cLocation === -1 && cBranch !== -1),
      mapped: mapped, alreadyClean: already, blanked: blanked,
      blankedEmployees: blankedList, retiredValues: unmapped,
      message: (preview ? 'PREVIEW — ' : '') + 'Mapped ' + mapped + ' to clean locations, ' + already + ' already clean, blanked ' + blanked +
        ' on retired branches' + (blanked ? ' (' + blankedList.join('; ') + ')' : '') + '. Central Kitchen set geofenced at Qurain coords.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ---- Supervisors per location (stored in Branch_Locations "Supervisors" col) ----
function ensureSupervisorsCol_(ss) {
  const sheet = ensureLocationCols_(ss);
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (H.findIndex(h => String(h).trim().toLowerCase() === 'supervisors') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Supervisors').setFontWeight('bold');
  }
  return sheet;
}

// Save the supervisor list (comma-separated Employee IDs) for a location.
// params: { location, supervisorIds: "MT-00001,MT-00002" }
function saveLocationSupervisors(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSupervisorsCol_(ss);
    const location = String(params.location || '').trim();
    if (!location) return { success: false, error: 'location required' };
    const d = sheet.getDataRange().getValues();
    const H = d[0];
    const cB = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cS = H.findIndex(h => String(h).trim().toLowerCase() === 'supervisors');
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cB] || '').trim().toLowerCase() === location.toLowerCase()) {
        sheet.getRange(i + 1, cS + 1).setValue(String(params.supervisorIds || '').trim());
        return { success: true, message: 'Saved supervisors for ' + location + '.' };
      }
    }
    return { success: false, error: 'Location "' + location + '" not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Return locations with their supervisor employee IDs+names (for the pages).
function getLocationSupervisors(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureSupervisorsCol_(ss);
    const d = sheet.getDataRange().getValues();
    const H = d[0];
    const cB = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cS = H.findIndex(h => String(h).trim().toLowerCase() === 'supervisors');
    const cA = H.findIndex(h => String(h).trim().toLowerCase() === 'active');
    // Employee id -> name
    const hr = ss.getSheetByName('HR Maya');
    const hd = hr.getDataRange().getValues();
    const hH = hd[0];
    const hcId = hH.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const hcName = hH.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    const nameById = {};
    for (let i = 1; i < hd.length; i++) { const id = String(hd[i][hcId] || '').trim(); if (id) nameById[id] = String(hd[i][hcName] || '').trim(); }

    const out = [];
    for (let i = 1; i < d.length; i++) {
      const name = String(d[i][cB] || '').trim();
      if (!name) continue;
      if (cA !== -1 && String(d[i][cA]).trim().toLowerCase() !== 'yes') continue;
      const ids = String((cS !== -1 ? d[i][cS] : '') || '').split(',').map(s => s.trim()).filter(Boolean);
      out.push({ location: name, supervisors: ids.map(id => ({ id: id, name: nameById[id] || id })) });
    }
    return { success: true, data: out };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Employees tagged to a given location (for the temp attendance page filter).
function getEmployeesByLocation(params) {
  try {
    params = params || {};
    const location = String(params.location || '').trim();
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const d = hr.getDataRange().getValues();
    const H = d[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    let cLoc = H.findIndex(h => String(h).trim().toLowerCase() === 'location');
    if (cLoc === -1) cLoc = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cStatus = H.findIndex(h => String(h).trim().toLowerCase() === 'status');
    const out = [];
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][cId] || '').trim(); if (!id) continue;
      if (cStatus !== -1 && String(d[i][cStatus]).trim().toLowerCase() === 'inactive') continue;
      const loc = cLoc !== -1 ? String(d[i][cLoc] || '').trim() : '';
      if (location && loc.toLowerCase() !== location.toLowerCase()) continue;
      out.push({ id: id, name: cName !== -1 ? String(d[i][cName] || '').trim() : id, location: loc });
    }
    out.sort((a, b) => a.name.localeCompare(b.name));
    return { success: true, employees: out };
  } catch (err) { return { success: false, error: err.toString() }; }
}