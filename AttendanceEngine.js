// ============================================================
// AttendanceEngine.gs — location-shifts, shift resolution, pay engine
// ============================================================
// Model (locked):
//  - Shifts are assigned to LOCATIONS (Branch_Locations "Shifts" column,
//    comma-separated shift names). A location can hold multiple shifts.
//  - Kiosk check-in resolves the day's shift from the PUNCH's geofenced
//    location -> among that location's shifts, the one whose START is
//    CLOSEST (absolute) to check-in time. Fallback: employee's Location,
//    else flag "no location for shift".
//  - OT = presence - shift Normal Hours (only > shift hours, e.g. 9).
//  - Full day > 6h; Half day <= 6h (deduct Basic/26/2, admin-approved).
//  - No-punch day = Day Off (paid) by default; admin may mark Absent.
//  - Everything stamped on Attendance_Log; approvals feed payroll.
//
// ENDPOINTS: saveLocationShifts, resolveShiftForCheckin (helper)
// ============================================================

function ensureLocationShiftsCol_(ss) {
  const sheet = ensureLocationCols_(ss);
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (H.findIndex(h => String(h).trim().toLowerCase() === 'shifts') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Shifts').setFontWeight('bold');
  }
  return sheet;
}

// Save comma-separated shift names for a location.
// params: { location, shifts: "Central Kitchen,Evening" }
function saveLocationShifts(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLocationShiftsCol_(ss);
    const location = String(params.location || '').trim();
    if (!location) return { success: false, error: 'location required' };
    const d = sheet.getDataRange().getValues();
    const H = d[0];
    const cB = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cSh = H.findIndex(h => String(h).trim().toLowerCase() === 'shifts');
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cB] || '').trim().toLowerCase() === location.toLowerCase()) {
        sheet.getRange(i + 1, cSh + 1).setValue(String(params.shifts || '').trim());
        return { success: true, message: 'Saved shifts for ' + location + ': ' + (params.shifts || '(none)') };
      }
    }
    return { success: false, error: 'Location "' + location + '" not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// location name -> [shift names]
function getLocationShiftsMap_(ss) {
  const map = {};
  const sheet = ensureLocationShiftsCol_(ss);
  const d = sheet.getDataRange().getValues();
  const H = d[0];
  const cB = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
  const cSh = H.findIndex(h => String(h).trim().toLowerCase() === 'shifts');
  for (let i = 1; i < d.length; i++) {
    const name = String(d[i][cB] || '').trim();
    if (!name) continue;
    const list = String((cSh !== -1 ? d[i][cSh] : '') || '').split(',').map(s => s.trim()).filter(Boolean);
    map[name.toLowerCase()] = list;
  }
  return map;
}

// Resolve the shift for a check-in: location's shifts, closest start to checkinMin.
// Returns shift name or ''.
function resolveShiftForCheckin_(ss, locationName, checkinMin, shiftMap) {
  if (!locationName) return '';
  const locShifts = getLocationShiftsMap_(ss)[String(locationName).trim().toLowerCase()] || [];
  if (!locShifts.length) return '';
  const toMin = (s) => { const m = String(s).trim().match(/^(\d{1,2}):(\d{2})$/); return m ? (parseInt(m[1],10)*60 + parseInt(m[2],10)) : null; };
  let best = null;
  locShifts.forEach(sn => {
    const sh = shiftMap[sn.toLowerCase()];
    if (!sh || sh.startMin == null) return;
    const dist = Math.abs(checkinMin - sh.startMin);
    if (!best || dist < best.dist) best = { name: sh.name, dist: dist };
  });
  return best ? best.name : (locShifts.length === 1 ? locShifts[0] : '');
}