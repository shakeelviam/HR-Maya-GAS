// ============================================================
// RosterBackend.gs
// ------------------------------------------------------------
// Supervisor roster planning: date-range grid, draft/publish,
// emergency amend with audit trail, leave-lock enforcement,
// and punch-vs-roster discrepancy logging.
//
// Reuses supervisor auth from VoiceSupervisor.gs (same project):
//   vsAuthSupervisor_(supId, supPin) -> { ok, name, location, error }
//   vsGetTeam_(location)             -> [{ id, name }]
//   vsTeamContains_(location, empId) -> bool
//   vsEmpName_(empId)                -> string
//   vsAudit_(empId, name, date, action, supervisorName)
// VoiceSupervisor.gs MUST be deployed in this project for these
// to resolve. If it isn't, add it first.
//
// Add these cases to your doGet router (see bottom of this file
// for the exact lines to paste):
//   getRosterRange, saveRosterDraft, publishRoster,
//   amendRosterEntry, copyPreviousRange,
//   getDiscrepancies, markDiscrepancyReviewed
// ============================================================

const ROSTER_TAB        = 'Roster';
const ROSTER_DISC_TAB   = 'Roster_Discrepancies';

// Locations that never use rostering — handled by their own automatic logic:
//   Head Office     → auto-present on working days (Shifts tab, Auto Present=Yes)
//   Driver Location → present every day; drivers self-report Day Off via kiosk
const ROSTER_EXEMPT_LOCATIONS = ['Head Office', 'Driver Location'];

function rbCheckLocationEligible_(location) {
  if (ROSTER_EXEMPT_LOCATIONS.includes(location)) {
    const why = location === 'Head Office'
      ? 'staff there are automatically marked present on working days.'
      : 'drivers work every day and self-report their day off directly via the kiosk.';
    return { ok: false, error: location + ' does not use rostering — ' + why };
  }
  return { ok: true };
}

// ── Local utility helpers (rb_ prefix — no shared-name collisions) ────

function rbColMap_(sheet) {
  const h = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const m = {};
  h.forEach((v,i) => { const k = String(v).trim(); if (k) m[k] = i; });
  return m;
}

function rbPad_(n) { return String(n).padStart(2, '0'); }

function rbNormDd_(val) {
  if (!val) return '';
  if (val instanceof Date) return rbPad_(val.getDate())+'-'+rbPad_(val.getMonth()+1)+'-'+val.getFullYear();
  const s = String(val).trim();
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3]+'-'+m[2]+'-'+m[1];
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) return rbPad_(m2[1])+'-'+rbPad_(m2[2])+'-'+m2[3];
  return s;
}

function rbParseDate_(s) {
  const m = String(s).match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
}

function rbDateToStr_(d) {
  return rbPad_(d.getDate())+'-'+rbPad_(d.getMonth()+1)+'-'+d.getFullYear();
}

// Inclusive list of dd-mm-yyyy strings between from and to
function rbDateRange_(fromStr, toStr) {
  const s = rbParseDate_(fromStr);
  const e = rbParseDate_(toStr);
  const out = [];
  if (!s || !e || e < s) return out;
  for (let d = new Date(s); d <= e; d.setDate(d.getDate()+1)) {
    out.push(rbDateToStr_(new Date(d)));
  }
  return out;
}

function rbDateInRange_(dateStr, startStr, endStr) {
  function toIso(s) { const p = s.split('-'); return p[2]+'-'+p[1]+'-'+p[0]; }
  const d = toIso(dateStr), a = toIso(startStr), b = toIso(endStr);
  return d >= a && d <= b;
}

// ── Sheet bootstrap ─────────────────────────────────────────────────

function rbEnsureRosterSheet_(ss) {
  let sheet = ss.getSheetByName(ROSTER_TAB);
  const HEADERS = [
    'Employee ID', 'Name', 'Location', 'Date',
    'Planned Status', 'Shift', 'Status',
    'Published By', 'Published On',
    'Last Amended By', 'Last Amended On', 'Amend Reason'
  ];
  if (!sheet) {
    sheet = ss.insertSheet(ROSTER_TAB);
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rbEnsureDiscrepancySheet_(ss) {
  let sheet = ss.getSheetByName(ROSTER_DISC_TAB);
  const HEADERS = [
    'Employee ID', 'Name', 'Location', 'Date',
    'Roster Said', 'Actual', 'Flagged On', 'Reviewed', 'Reviewed By'
  ];
  if (!sheet) {
    sheet = ss.insertSheet(ROSTER_DISC_TAB);
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Leave lock — approved Leave_Grants covering a date ───────────────
// Returns { empId+'|'+date: leaveType } for employees on this team.
function rbGetLeaveMap_(team, fromStr, toStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lvSheet = ss.getSheetByName('Leave_Grants');
  const map = {};
  if (!lvSheet || lvSheet.getLastRow() < 2) return map;

  const teamIds = new Set(team.map(e => e.id));
  const lc      = rbColMap_(lvSheet);
  const LV_EMP  = lc['Employee ID'] ?? 1;
  const LV_TYPE = lc['Type']        ?? 3;
  const LV_ST   = lc['Start']       ?? 4;
  const LV_EN   = lc['End']         ?? 5;
  const LV_STAT = lc['Status']      ?? 7;

  lvSheet.getRange(2,1,lvSheet.getLastRow()-1,lvSheet.getLastColumn()).getValues()
    .forEach(row => {
      if (String(row[LV_STAT]||'').trim() !== 'Approved') return;
      const empId = String(row[LV_EMP]||'').trim();
      if (!teamIds.has(empId)) return;
      const type  = String(row[LV_TYPE]||'').trim();
      const start = rbNormDd_(row[LV_ST]);
      const end   = rbNormDd_(row[LV_EN]);
      if (!start || !end) return;
      rbDateRange_(start, end).forEach(dateStr => {
        if (rbDateInRange_(dateStr, fromStr, toStr) || true) { // filtered again by caller
          map[empId + '|' + dateStr] = type;
        }
      });
    });
  return map;
}

// ── Shifts available for a location — reads Branch_Locations "Shifts" col,
// falls back to any Shifts-tab row whose name contains the location. ─────
function rbGetShiftsForLocation_(location) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const locSheet = ss.getSheetByName('Branch_Locations');
  if (locSheet && locSheet.getLastRow() > 1) {
    const lc = rbColMap_(locSheet);
    const cB = lc['Branch'], cSh = lc['Shifts'];
    if (cB !== undefined && cSh !== undefined) {
      const data = locSheet.getRange(2,1,locSheet.getLastRow()-1,locSheet.getLastColumn()).getValues();
      for (const row of data) {
        if (String(row[cB]||'').trim().toLowerCase() === location.toLowerCase()) {
          const list = String(row[cSh]||'').split(',').map(s => s.trim()).filter(Boolean);
          if (list.length) return list;
        }
      }
    }
  }
  // Fallback: filter Shifts tab by name containing the location string
  const shSheet = ss.getSheetByName('Shifts');
  const out = [];
  if (shSheet && shSheet.getLastRow() > 1) {
    const sc = rbColMap_(shSheet);
    const SH_NM = sc['Shift'] ?? 0;
    const SH_AC = sc['Active'] ?? -1;
    shSheet.getRange(2,1,shSheet.getLastRow()-1,shSheet.getLastColumn()).getValues()
      .forEach(row => {
        const name = String(row[SH_NM]||'').trim();
        if (!name) return;
        if (SH_AC !== -1 && String(row[SH_AC]).trim().toLowerCase() !== 'yes') return;
        if (name.toLowerCase().includes(location.toLowerCase())) out.push(name);
      });
  }
  return out;
}

// ── Read existing roster rows for a location + date range ────────────
// Returns { empId+'|'+date: rowObj } plus the raw sheet row number (1-based)
// so callers can update in place.
function rbReadRosterRows_(location, fromStr, toStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = rbEnsureRosterSheet_(ss);
  const map = {};
  if (sheet.getLastRow() < 2) return { map, sheet };
  const rc = rbColMap_(sheet);
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach((row, i) => {
    const loc  = String(row[rc['Location']]||'').trim();
    const date = rbNormDd_(row[rc['Date']]);
    if (loc !== location) return;
    if (!rbDateInRange_(date, fromStr, toStr)) return;
    map[String(row[rc['Employee ID']]).trim() + '|' + date] = {
      rowNum: i + 2, // 1-based, +1 for header
      plannedStatus: String(row[rc['Planned Status']]||'').trim(),
      shift: String(row[rc['Shift']]||'').trim(),
      status: String(row[rc['Status']]||'').trim(),
    };
  });
  return { map, sheet, rc };
}

// ══════════════════════════════════════════════════════════════════
// ENDPOINT: getRosterRange
// ══════════════════════════════════════════════════════════════════
function getRosterRange(params) {
  try {
    const { supId, supPin, from, to } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };
    const eligible = rbCheckLocationEligible_(auth.location);
    if (!eligible.ok) return { success:false, error: eligible.error, exempt:true };
    if (!from || !to) return { success:false, error: 'from and to dates are required.' };

    const dates = rbDateRange_(from, to);
    if (!dates.length) return { success:false, error: 'Invalid date range.' };

    const team    = vsGetTeam_(auth.location);
    const shifts  = rbGetShiftsForLocation_(auth.location);
    const leaveMap = rbGetLeaveMap_(team, from, to);
    const { map: rosterMap } = rbReadRosterRows_(auth.location, from, to);

    const cells = {};
    team.forEach(emp => {
      cells[emp.id] = {};
      dates.forEach(date => {
        const leaveKey = emp.id + '|' + date;
        if (leaveMap[leaveKey]) {
          cells[emp.id][date] = { locked: true, lockReason: 'On Leave (' + leaveMap[leaveKey] + ')' };
          return;
        }
        const r = rosterMap[leaveKey];
        if (r) {
          cells[emp.id][date] = {
            locked: false,
            status: r.plannedStatus,
            shift: r.shift,
            published: r.status === 'Published',
          };
        } else {
          cells[emp.id][date] = { locked: false, status: '', shift: '', published: false };
        }
      });
    });

    return {
      success: true,
      location: auth.location,
      supervisorName: auth.name,
      employees: team,
      dates: dates,
      shifts: shifts,
      cells: cells,
    };
  } catch (e) {
    Logger.log('getRosterRange: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
// ENDPOINT: saveRosterDraft
// cells: { empId: { date: { status, shift } } }  (JSON string in POST body)
// ══════════════════════════════════════════════════════════════════
function saveRosterDraft(params) {
  try {
    const { supId, supPin, from, to } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };
    const eligible = rbCheckLocationEligible_(auth.location);
    if (!eligible.ok) return { success:false, error: eligible.error, exempt:true };

    let cells;
    try { cells = typeof params.cells === 'string' ? JSON.parse(params.cells) : params.cells; }
    catch (e) { return { success:false, error:'Invalid cells payload.' }; }
    if (!cells) return { success:false, error:'No cell data provided.' };

    const team     = vsGetTeam_(auth.location);
    const teamIds  = new Set(team.map(e => e.id));
    const leaveMap = rbGetLeaveMap_(team, from, to);
    const { map: existing, sheet, rc } = rbReadRosterRows_(auth.location, from, to);

    let saved = 0, skippedLeave = 0, skippedScope = 0;
    const appendRows = [];

    Object.keys(cells).forEach(empId => {
      if (!teamIds.has(empId)) { skippedScope++; return; }
      const empName = vsEmpName_(empId);
      Object.keys(cells[empId]).forEach(date => {
        const cell = cells[empId][date];
        if (!cell || !cell.status) return; // skip blanks
        const leaveKey = empId + '|' + date;
        if (leaveMap[leaveKey]) { skippedLeave++; return; } // never override approved leave

        const key = leaveKey;
        if (existing[key]) {
          // Don't downgrade an already-Published row back to Draft on a routine save
          const newStatus = existing[key].status === 'Published' ? 'Published' : 'Draft';
          sheet.getRange(existing[key].rowNum, rc['Planned Status']+1).setValue(cell.status);
          sheet.getRange(existing[key].rowNum, rc['Shift']+1).setValue(cell.shift || '');
          sheet.getRange(existing[key].rowNum, rc['Status']+1).setValue(newStatus);
        } else {
          const row = new Array(12).fill('');
          row[rc ? rc['Employee ID'] : 0] = empId;
          row[rc ? rc['Name']        : 1] = empName;
          row[rc ? rc['Location']    : 2] = auth.location;
          row[rc ? rc['Date']        : 3] = date;
          row[rc ? rc['Planned Status'] : 4] = cell.status;
          row[rc ? rc['Shift']       : 5] = cell.shift || '';
          row[rc ? rc['Status']      : 6] = 'Draft';
          appendRows.push(row);
        }
        saved++;
      });
    });

    if (appendRows.length) {
      sheet.getRange(sheet.getLastRow()+1, 1, appendRows.length, appendRows[0].length).setValues(appendRows);
    }

    return {
      success: true,
      message: 'Draft saved: ' + saved + ' cell(s).' +
        (skippedLeave ? ' ' + skippedLeave + ' skipped (on approved leave).' : '') +
        (skippedScope ? ' ' + skippedScope + ' skipped (not on your team).' : ''),
    };
  } catch (e) {
    Logger.log('saveRosterDraft: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
// ENDPOINT: publishRoster — flips Draft rows in range to Published
// ══════════════════════════════════════════════════════════════════
function publishRoster(params) {
  try {
    const { supId, supPin, from, to } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };
    const eligible = rbCheckLocationEligible_(auth.location);
    if (!eligible.ok) return { success:false, error: eligible.error, exempt:true };

    const { map: existing, sheet, rc } = rbReadRosterRows_(auth.location, from, to);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    let count = 0;
    Object.values(existing).forEach(row => {
      if (!row.plannedStatus) return; // nothing to publish
      sheet.getRange(row.rowNum, rc['Status']+1).setValue('Published');
      sheet.getRange(row.rowNum, rc['Published By']+1).setValue(auth.name);
      sheet.getRange(row.rowNum, rc['Published On']+1).setValue(now);
      count++;
    });

    if (!count) return { success:false, error:'No draft entries found to publish for this range. Save a draft first.' };
    return { success:true, message: 'Published ' + count + ' roster entr' + (count===1?'y':'ies') + ' for ' + from + ' to ' + to + '.' };
  } catch (e) {
    Logger.log('publishRoster: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
// ENDPOINT: amendRosterEntry — emergency edit, always publishes
// immediately, always requires a reason, works on past or future dates.
// ══════════════════════════════════════════════════════════════════
function amendRosterEntry(params) {
  try {
    const { supId, supPin, empId, date, newStatus, newShift, reason } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };
    const eligible = rbCheckLocationEligible_(auth.location);
    if (!eligible.ok) return { success:false, error: eligible.error, exempt:true };
    if (!vsTeamContains_(auth.location, empId)) return { success:false, error:'Employee is not on your team.' };
    if (!date || !newStatus) return { success:false, error:'Date and status are required.' };
    if (!reason || !reason.trim()) return { success:false, error:'A reason is required for an emergency amend.' };

    const team = vsGetTeam_(auth.location);
    const leaveMap = rbGetLeaveMap_(team, date, date);
    if (leaveMap[empId + '|' + date]) {
      return { success:false, error:'Employee is on approved leave (' + leaveMap[empId+'|'+date] + ') for this date — cannot schedule.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsureRosterSheet_(ss);
    const rc = rbColMap_(sheet);
    const empName = vsEmpName_(empId);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    // Find existing row for this employee+date
    let foundRow = -1;
    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][rc['Employee ID']]).trim() === empId &&
            rbNormDd_(data[i][rc['Date']]) === date) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow > -1) {
      sheet.getRange(foundRow, rc['Planned Status']+1).setValue(newStatus);
      sheet.getRange(foundRow, rc['Shift']+1).setValue(newShift || '');
      sheet.getRange(foundRow, rc['Status']+1).setValue('Published');
      // Preserve original Published By/On if already set; otherwise stamp now
      const existingPubBy = sheet.getRange(foundRow, rc['Published By']+1).getValue();
      if (!existingPubBy) {
        sheet.getRange(foundRow, rc['Published By']+1).setValue(auth.name);
        sheet.getRange(foundRow, rc['Published On']+1).setValue(now);
      }
      sheet.getRange(foundRow, rc['Last Amended By']+1).setValue(auth.name);
      sheet.getRange(foundRow, rc['Last Amended On']+1).setValue(now);
      sheet.getRange(foundRow, rc['Amend Reason']+1).setValue(reason.trim());
    } else {
      const row = new Array(12).fill('');
      row[rc['Employee ID']]     = empId;
      row[rc['Name']]            = empName;
      row[rc['Location']]        = auth.location;
      row[rc['Date']]            = date;
      row[rc['Planned Status']]  = newStatus;
      row[rc['Shift']]           = newShift || '';
      row[rc['Status']]          = 'Published';
      row[rc['Published By']]    = auth.name;
      row[rc['Published On']]    = now;
      row[rc['Last Amended By']] = auth.name;
      row[rc['Last Amended On']] = now;
      row[rc['Amend Reason']]    = reason.trim();
      sheet.appendRow(row);
    }

    // Cross-visible audit trail alongside other supervisor actions
    if (typeof vsAudit_ === 'function') {
      vsAudit_(empId, empName, date, 'ROSTER AMEND: ' + newStatus + (newShift ? ' (' + newShift + ')' : '') + ' — ' + reason.trim(), auth.name);
    }

    return { success:true, message: 'Roster amended for ' + empName + ' on ' + date + '.' };
  } catch (e) {
    Logger.log('amendRosterEntry: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
// ENDPOINT: copyPreviousRange — returns cells from the immediately
// preceding range of the same length, for the client to pre-fill.
// ══════════════════════════════════════════════════════════════════
function copyPreviousRange(params) {
  try {
    const { supId, supPin, from, to } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };
    const eligible = rbCheckLocationEligible_(auth.location);
    if (!eligible.ok) return { success:false, error: eligible.error, exempt:true };

    const dates = rbDateRange_(from, to);
    if (!dates.length) return { success:false, error:'Invalid date range.' };

    const rangeLen = dates.length;
    const fromDate = rbParseDate_(from);
    const prevTo   = new Date(fromDate); prevTo.setDate(prevTo.getDate() - 1);
    const prevFrom = new Date(prevTo);   prevFrom.setDate(prevFrom.getDate() - (rangeLen - 1));
    const prevFromStr = rbDateToStr_(prevFrom);
    const prevToStr   = rbDateToStr_(prevTo);
    const prevDates    = rbDateRange_(prevFromStr, prevToStr);

    const { map: prevMap } = rbReadRosterRows_(auth.location, prevFromStr, prevToStr);
    const team = vsGetTeam_(auth.location);

    const cells = {};
    team.forEach(emp => {
      cells[emp.id] = {};
      dates.forEach((newDate, i) => {
        const prevDate = prevDates[i];
        const key = emp.id + '|' + prevDate;
        const r = prevMap[key];
        if (r && r.plannedStatus) {
          cells[emp.id][newDate] = { status: r.plannedStatus, shift: r.shift };
        }
      });
    });

    return { success:true, cells: cells, sourceFrom: prevFromStr, sourceTo: prevToStr };
  } catch (e) {
    Logger.log('copyPreviousRange: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
// ENDPOINT: getDiscrepancies — this location's unreviewed mismatches
// ══════════════════════════════════════════════════════════════════
function getDiscrepancies(params) {
  try {
    const { supId, supPin } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsureDiscrepancySheet_(ss);
    if (sheet.getLastRow() < 2) return { success:true, rows: [] };

    const rc = rbColMap_(sheet);
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const rows = [];
    data.forEach((row, i) => {
      const loc = String(row[rc['Location']]||'').trim();
      const reviewed = String(row[rc['Reviewed']]||'').trim();
      if (loc !== auth.location) return;
      if (reviewed === 'Yes') return;
      rows.push({
        rowNum: i + 2,
        empId: row[rc['Employee ID']],
        name: row[rc['Name']],
        date: rbNormDd_(row[rc['Date']]),
        rosterSaid: row[rc['Roster Said']],
        actual: row[rc['Actual']],
        flaggedOn: row[rc['Flagged On']],
      });
    });

    return { success:true, rows: rows };
  } catch (e) {
    Logger.log('getDiscrepancies: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
// ENDPOINT: markDiscrepancyReviewed
// ══════════════════════════════════════════════════════════════════
function markDiscrepancyReviewed(params) {
  try {
    const { supId, supPin, rowNum } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsureDiscrepancySheet_(ss);
    const rc = rbColMap_(sheet);
    const r = parseInt(rowNum);
    if (!r || r < 2 || r > sheet.getLastRow()) return { success:false, error:'Invalid row.' };

    const loc = String(sheet.getRange(r, rc['Location']+1).getValue()).trim();
    if (loc !== auth.location) return { success:false, error:'This discrepancy belongs to a different location.' };

    sheet.getRange(r, rc['Reviewed']+1).setValue('Yes');
    sheet.getRange(r, rc['Reviewed By']+1).setValue(auth.name);
    return { success:true, message:'Marked reviewed.' };
  } catch (e) {
    Logger.log('markDiscrepancyReviewed: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
// Called by AttendanceDailyTrigger.gs during the daily compile.
// Returns published roster status for ALL employees across ALL
// locations for a single date: { empId: { status, shift } }
// ══════════════════════════════════════════════════════════════════
function rbGetPublishedRosterForDate_(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ROSTER_TAB);
  const out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  const rc = rbColMap_(sheet);
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(row => {
    if (String(row[rc['Status']]||'').trim() !== 'Published') return;
    if (rbNormDd_(row[rc['Date']]) !== dateStr) return;
    const empId = String(row[rc['Employee ID']]||'').trim();
    if (!empId) return;
    out[empId] = {
      status: String(row[rc['Planned Status']]||'').trim(),
      shift:  String(row[rc['Shift']]||'').trim(),
      location: String(row[rc['Location']]||'').trim(),
    };
  });
  return out;
}

// Called by AttendanceDailyTrigger.gs whenever a punch contradicts
// the published roster for that date.
function rbLogDiscrepancy_(empId, name, location, dateStr, rosterSaid, actual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = rbEnsureDiscrepancySheet_(ss);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  sheet.appendRow([empId, name, location, dateStr, rosterSaid, actual, now, 'No', '']);
}

// ══════════════════════════════════════════════════════════════════
// EMPLOYEE TRANSFER — request / approve / schedule / execute.
//
// Two paths:
//   STANDARD (both from & to are regular rostering locations):
//     Request → sits as "Pending Approval" → destination supervisor
//     Approves or Rejects → if Approved and effective date is today,
//     executes immediately; if future, waits for the 4 AM trigger.
//   EXEMPT (either side is Head Office or Driver Location):
//     No approval needed — Head Office/Driver have no roster to
//     protect. Executes immediately if effective date is today,
//     otherwise auto-queued as "Approved" for the trigger to apply
//     on the day.
//
// Effective date can never be in the past — a retroactive transfer
// would mean rewriting already-compiled Attendance_Daily history and
// possibly a closed payroll period. That's a manual HR correction,
// not a self-service supervisor action.
//
// No shift is chosen at request or approval time — the old location's
// shift names are meaningless at the new one. HR Maya's Shift field
// is left blank on execution; the receiving supervisor assigns a
// shift via normal roster planning afterward. getMyTeam flags any
// employee with zero future roster rows so this is never missed.
// ══════════════════════════════════════════════════════════════════

const TRANSFER_TAB         = 'Employee_Transfers';
const PENDING_TRANSFER_TAB = 'Pending_Transfers';

function rbEnsureTransferSheet_(ss) {
  let sheet = ss.getSheetByName(TRANSFER_TAB);
  const HEADERS = [
    'Employee ID', 'Name', 'From Location', 'To Location',
    'New Shift', 'Executed By', 'Executed On', 'Reason'
  ];
  if (!sheet) {
    sheet = ss.insertSheet(TRANSFER_TAB);
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rbEnsurePendingTransferSheet_(ss) {
  let sheet = ss.getSheetByName(PENDING_TRANSFER_TAB);
  const HEADERS = [
    'Request ID', 'Employee ID', 'Name', 'From Location', 'To Location',
    'Effective Date', 'Reason', 'Requested By', 'Requested On',
    'Status', 'Approved By', 'Approved On',
    'Rejected By', 'Rejected On', 'Rejection Note'
  ];
  if (!sheet) {
    sheet = ss.insertSheet(PENDING_TRANSFER_TAB);
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rbIsExemptTransfer_(fromLoc, toLoc) {
  return ROSTER_EXEMPT_LOCATIONS.includes(fromLoc) || ROSTER_EXEMPT_LOCATIONS.includes(toLoc);
}

function rbGenTransferId_() {
  return 'TR-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') +
    '-' + Math.floor(Math.random() * 1000);
}

// ── ENDPOINT: getMyTeam — lightweight team list, works for ALL
// supervisors including exempt locations. Non-exempt locations get a
// needsRoster flag per employee (no future roster rows at all — this
// is what catches a freshly transferred employee, or a new hire
// nobody's scheduled yet). ─────────────────────────────────────────
function getMyTeam(params) {
  try {
    const { supId, supPin } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const team = vsGetTeam_(auth.location);
    const isExempt = ROSTER_EXEMPT_LOCATIONS.includes(auth.location);
    let employees = team;

    if (!isExempt) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const rosterSheet = ss.getSheetByName(ROSTER_TAB);
      const hasFutureRoster = new Set();
      if (rosterSheet && rosterSheet.getLastRow() > 1) {
        const rc = rbColMap_(rosterSheet);
        const todayStr = rbDateToStr_(new Date());
        rosterSheet.getRange(2,1,rosterSheet.getLastRow()-1,rosterSheet.getLastColumn()).getValues()
          .forEach(row => {
            if (String(row[rc['Location']]||'').trim() !== auth.location) return;
            const date = rbNormDd_(row[rc['Date']]);
            if (rbDateInRange_(date, todayStr, '31-12-2099')) {
              hasFutureRoster.add(String(row[rc['Employee ID']]||'').trim());
            }
          });
      }
      employees = team.map(e => Object.assign({}, e, { needsRoster: !hasFutureRoster.has(e.id) }));
    }

    return {
      success: true,
      location: auth.location,
      supervisorName: auth.name,
      employees: employees,
      rosterExempt: isExempt,
    };
  } catch (e) {
    Logger.log('getMyTeam: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── ENDPOINT: getActiveLocations — for the destination dropdown. ────
function getActiveLocations(params) {
  try {
    const { supId, supPin } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const locSheet = ss.getSheetByName('Branch_Locations');
    const out = [];
    if (locSheet && locSheet.getLastRow() > 1) {
      const lc = rbColMap_(locSheet);
      const cB = lc['Branch'], cA = lc['Active'];
      locSheet.getRange(2,1,locSheet.getLastRow()-1,locSheet.getLastColumn()).getValues().forEach(row => {
        const name = String(row[cB]||'').trim();
        if (!name || name === auth.location) return;
        const active = cA === undefined || String(row[cA]).trim().toLowerCase() === 'yes';
        if (active) out.push(name);
      });
    }
    return { success:true, locations: out };
  } catch (e) {
    Logger.log('getActiveLocations: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── ENDPOINT: getShiftsForLocation — kept for general reuse (e.g. the
// receiving supervisor picking a shift when they build the roster
// after a transfer completes). Not called by the transfer flow itself
// since no shift is chosen at request/approval time. ────────────────
function getShiftsForLocation(params) {
  try {
    const { supId, supPin, location } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };
    if (!location) return { success:false, error:'Location is required.' };
    return { success:true, shifts: rbGetShiftsForLocation_(location) };
  } catch (e) {
    Logger.log('getShiftsForLocation: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── INTERNAL: rbApplyTransfer_ — the actual execution. Shared by the
// exempt-immediate path, the approve-today path, and the daily
// trigger's scheduled execution. Does NOT touch Pending_Transfers —
// callers handle status updates there themselves. ───────────────────
function rbApplyTransfer_(empId, empName, fromLocation, toLocation, effectiveDate, actorName, reason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hrSheet = ss.getSheetByName('HR Maya');
    const hc = rbColMap_(hrSheet);
    const locCol = hc['Location'] !== undefined ? hc['Location'] : hc['Branch'];
    if (locCol === undefined) return { success:false, error:'HR Maya has no Location/Branch column.' };

    const hrData = hrSheet.getRange(2,1,hrSheet.getLastRow()-1,hrSheet.getLastColumn()).getValues();
    let hrRow = -1;
    for (let i = 0; i < hrData.length; i++) {
      if (String(hrData[i][hc['Employee ID']]).trim() === empId) { hrRow = i + 2; break; }
    }
    if (hrRow === -1) return { success:false, error:'Employee not found in HR Maya.' };

    hrSheet.getRange(hrRow, locCol + 1).setValue(toLocation);
    // Shift intentionally left blank — assigned later via normal roster planning.
    if (hc['Shift'] !== undefined) hrSheet.getRange(hrRow, hc['Shift'] + 1).setValue('');

    // Delete stale roster rows: effectiveDate onward, at the OLD location.
    const rosterSheet = rbEnsureRosterSheet_(ss);
    if (rosterSheet.getLastRow() > 1) {
      const rc = rbColMap_(rosterSheet);
      const rData = rosterSheet.getRange(2,1,rosterSheet.getLastRow()-1,rosterSheet.getLastColumn()).getValues();
      const toDelete = [];
      rData.forEach((row, i) => {
        if (String(row[rc['Employee ID']]).trim() !== empId) return;
        if (String(row[rc['Location']]).trim() !== fromLocation) return;
        const rDate = rbNormDd_(row[rc['Date']]);
        if (rbDateInRange_(rDate, effectiveDate, '31-12-2099')) toDelete.push(i + 2);
      });
      for (let i = toDelete.length - 1; i >= 0; i--) rosterSheet.deleteRow(toDelete[i]);
    }

    const trSheet = rbEnsureTransferSheet_(ss);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    trSheet.appendRow([empId, empName, fromLocation, toLocation, '', actorName, now, reason || '']);

    return { success:true };
  } catch (e) {
    Logger.log('rbApplyTransfer_: ' + e.message);
    return { success:false, error:'Server error applying transfer: ' + e.message };
  }
}

// ── ENDPOINT: requestTransfer — always callable, branches internally. ─
function requestTransfer(params) {
  try {
    const { supId, supPin, empId, toLocation, effectiveDate, reason } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };
    if (!vsTeamContains_(auth.location, empId)) return { success:false, error:'Employee is not on your team.' };
    if (!toLocation || !toLocation.trim()) return { success:false, error:'Destination location is required.' };
    if (toLocation.trim() === auth.location) return { success:false, error:'Employee is already at this location.' };
    if (!effectiveDate) return { success:false, error:'Effective date is required.' };

    const effDateObj = rbParseDate_(effectiveDate);
    const todayObj    = rbParseDate_(rbDateToStr_(new Date()));
    if (!effDateObj) return { success:false, error:'Invalid effective date.' };
    if (effDateObj < todayObj) return { success:false, error:'Effective date cannot be in the past.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const locSheet = ss.getSheetByName('Branch_Locations');
    let validLoc = false;
    if (locSheet && locSheet.getLastRow() > 1) {
      const lc = rbColMap_(locSheet);
      const cB = lc['Branch'], cA = lc['Active'];
      locSheet.getRange(2,1,locSheet.getLastRow()-1,locSheet.getLastColumn()).getValues().forEach(row => {
        if (String(row[cB]||'').trim() === toLocation.trim()) {
          validLoc = cA === undefined || String(row[cA]).trim().toLowerCase() === 'yes';
        }
      });
    }
    if (!validLoc) return { success:false, error:'"' + toLocation + '" is not a valid active location.' };

    const empName  = vsEmpName_(empId);
    const isExempt = rbIsExemptTransfer_(auth.location, toLocation.trim());
    const sheet    = rbEnsurePendingTransferSheet_(ss);
    const now      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    const reqId    = rbGenTransferId_();
    const isToday  = effDateObj.getTime() === todayObj.getTime();

    if (isExempt) {
      if (isToday) {
        const result = rbApplyTransfer_(empId, empName, auth.location, toLocation.trim(), effectiveDate, auth.name, reason);
        if (!result.success) return result;
        sheet.appendRow([reqId, empId, empName, auth.location, toLocation.trim(), effectiveDate,
          reason || '', auth.name, now, 'Completed', auth.name, now, '', '', '']);
        return { success:true, message: empName + ' transferred to ' + toLocation.trim() + ' now.' };
      } else {
        sheet.appendRow([reqId, empId, empName, auth.location, toLocation.trim(), effectiveDate,
          reason || '', auth.name, now, 'Approved', auth.name, now, '', '', '']);
        return { success:true, message: empName + "'s transfer to " + toLocation.trim() +
          ' is scheduled for ' + effectiveDate + '. No approval needed — one side is auto-managed.' };
      }
    } else {
      sheet.appendRow([reqId, empId, empName, auth.location, toLocation.trim(), effectiveDate,
        reason || '', auth.name, now, 'Pending Approval', '', '', '', '', '']);
      return { success:true, message: 'Transfer request sent — waiting for ' +
        toLocation.trim() + "'s supervisor to approve." };
    }
  } catch (e) {
    Logger.log('requestTransfer: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── ENDPOINT: getIncomingTransferRequests — pending requests where
// MY location is the destination. ────────────────────────────────────
function getIncomingTransferRequests(params) {
  try {
    const { supId, supPin } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsurePendingTransferSheet_(ss);
    if (sheet.getLastRow() < 2) return { success:true, rows: [] };

    const rc = rbColMap_(sheet);
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const rows = [];
    data.forEach((row, i) => {
      if (String(row[rc['To Location']]||'').trim() !== auth.location) return;
      if (String(row[rc['Status']]||'').trim() !== 'Pending Approval') return;
      rows.push({
        rowNum: i + 2,
        empId: row[rc['Employee ID']],
        name: row[rc['Name']],
        fromLocation: row[rc['From Location']],
        effectiveDate: rbNormDd_(row[rc['Effective Date']]),
        reason: row[rc['Reason']],
        requestedBy: row[rc['Requested By']],
        requestedOn: row[rc['Requested On']],
      });
    });
    return { success:true, rows: rows };
  } catch (e) {
    Logger.log('getIncomingTransferRequests: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── ENDPOINT: getOutgoingTransferRequests — requests I've made,
// excluding fully Completed ones (those live in Employee_Transfers). ──
function getOutgoingTransferRequests(params) {
  try {
    const { supId, supPin } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsurePendingTransferSheet_(ss);
    if (sheet.getLastRow() < 2) return { success:true, rows: [] };

    const rc = rbColMap_(sheet);
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const rows = [];
    data.forEach((row, i) => {
      if (String(row[rc['From Location']]||'').trim() !== auth.location) return;
      const status = String(row[rc['Status']]||'').trim();
      if (status === 'Completed') return;
      rows.push({
        rowNum: i + 2,
        empId: row[rc['Employee ID']],
        name: row[rc['Name']],
        toLocation: row[rc['To Location']],
        effectiveDate: rbNormDd_(row[rc['Effective Date']]),
        status: status,
        rejectionNote: row[rc['Rejection Note']],
        requestedOn: row[rc['Requested On']],
      });
    });
    rows.reverse();
    return { success:true, rows: rows };
  } catch (e) {
    Logger.log('getOutgoingTransferRequests: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── ENDPOINT: approveTransfer — destination supervisor only. Executes
// immediately if effective date is today, otherwise just flips to
// Approved and waits for the daily trigger. ──────────────────────────
function approveTransfer(params) {
  try {
    const { supId, supPin, rowNum } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsurePendingTransferSheet_(ss);
    const rc = rbColMap_(sheet);
    const r = parseInt(rowNum);
    if (!r || r < 2 || r > sheet.getLastRow()) return { success:false, error:'Invalid request.' };

    if (String(sheet.getRange(r, rc['To Location']+1).getValue()).trim() !== auth.location)
      return { success:false, error:'This request is not for your location.' };
    if (String(sheet.getRange(r, rc['Status']+1).getValue()).trim() !== 'Pending Approval')
      return { success:false, error:'This request is no longer pending.' };

    const empId         = String(sheet.getRange(r, rc['Employee ID']+1).getValue()).trim();
    const empName        = String(sheet.getRange(r, rc['Name']+1).getValue()).trim();
    const fromLoc         = String(sheet.getRange(r, rc['From Location']+1).getValue()).trim();
    const toLoc           = String(sheet.getRange(r, rc['To Location']+1).getValue()).trim();
    const effectiveDate  = rbNormDd_(sheet.getRange(r, rc['Effective Date']+1).getValue());
    const reason           = String(sheet.getRange(r, rc['Reason']+1).getValue()).trim();
    const now              = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    const isToday          = effectiveDate === rbDateToStr_(new Date());

    sheet.getRange(r, rc['Approved By']+1).setValue(auth.name);
    sheet.getRange(r, rc['Approved On']+1).setValue(now);

    if (isToday) {
      const result = rbApplyTransfer_(empId, empName, fromLoc, toLoc, effectiveDate, auth.name, reason);
      if (!result.success) return result;
      sheet.getRange(r, rc['Status']+1).setValue('Completed');
      return { success:true, message: empName + ' transferred to ' + toLoc + ' now.' };
    } else {
      sheet.getRange(r, rc['Status']+1).setValue('Approved');
      return { success:true, message: 'Approved. ' + empName + ' will move to ' + toLoc + ' on ' + effectiveDate + '.' };
    }
  } catch (e) {
    Logger.log('approveTransfer: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── ENDPOINT: rejectTransfer — destination supervisor only. ──────────
function rejectTransfer(params) {
  try {
    const { supId, supPin, rowNum, note } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsurePendingTransferSheet_(ss);
    const rc = rbColMap_(sheet);
    const r = parseInt(rowNum);
    if (!r || r < 2 || r > sheet.getLastRow()) return { success:false, error:'Invalid request.' };

    if (String(sheet.getRange(r, rc['To Location']+1).getValue()).trim() !== auth.location)
      return { success:false, error:'This request is not for your location.' };
    if (String(sheet.getRange(r, rc['Status']+1).getValue()).trim() !== 'Pending Approval')
      return { success:false, error:'This request is no longer pending.' };

    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    sheet.getRange(r, rc['Status']+1).setValue('Rejected');
    sheet.getRange(r, rc['Rejected By']+1).setValue(auth.name);
    sheet.getRange(r, rc['Rejected On']+1).setValue(now);
    sheet.getRange(r, rc['Rejection Note']+1).setValue(note || '');

    return { success:true, message:'Request rejected.' };
  } catch (e) {
    Logger.log('rejectTransfer: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ── ENDPOINT: cancelTransferRequest — requester only, before execution. ─
function cancelTransferRequest(params) {
  try {
    const { supId, supPin, rowNum } = params;
    const auth = vsAuthSupervisor_(supId, supPin);
    if (!auth.ok) return { success:false, error: auth.error };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = rbEnsurePendingTransferSheet_(ss);
    const rc = rbColMap_(sheet);
    const r = parseInt(rowNum);
    if (!r || r < 2 || r > sheet.getLastRow()) return { success:false, error:'Invalid request.' };

    if (String(sheet.getRange(r, rc['From Location']+1).getValue()).trim() !== auth.location)
      return { success:false, error:'This is not your request.' };
    const status = String(sheet.getRange(r, rc['Status']+1).getValue()).trim();
    if (status !== 'Pending Approval' && status !== 'Approved')
      return { success:false, error:'This request can no longer be cancelled (' + status + ').' };

    sheet.getRange(r, rc['Status']+1).setValue('Cancelled');
    return { success:true, message:'Transfer request cancelled.' };
  } catch (e) {
    Logger.log('cancelTransferRequest: ' + e.message);
    return { success:false, error:'Server error: ' + e.message };
  }
}

// ============================================================
// ROUTER CASES — add these to your doGet switch statement:
// ============================================================
//   case 'getRosterRange':
//     return ContentService.createTextOutput(JSON.stringify(getRosterRange(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'saveRosterDraft':
//     return ContentService.createTextOutput(JSON.stringify(saveRosterDraft((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'publishRoster':
//     return ContentService.createTextOutput(JSON.stringify(publishRoster(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'amendRosterEntry':
//     return ContentService.createTextOutput(JSON.stringify(amendRosterEntry((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'copyPreviousRange':
//     return ContentService.createTextOutput(JSON.stringify(copyPreviousRange(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'getDiscrepancies':
//     return ContentService.createTextOutput(JSON.stringify(getDiscrepancies(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'markDiscrepancyReviewed':
//     return ContentService.createTextOutput(JSON.stringify(markDiscrepancyReviewed(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'getMyTeam':
//     return ContentService.createTextOutput(JSON.stringify(getMyTeam(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'getActiveLocations':
//     return ContentService.createTextOutput(JSON.stringify(getActiveLocations(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'getShiftsForLocation':
//     return ContentService.createTextOutput(JSON.stringify(getShiftsForLocation(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'requestTransfer':
//     return ContentService.createTextOutput(JSON.stringify(requestTransfer((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'getIncomingTransferRequests':
//     return ContentService.createTextOutput(JSON.stringify(getIncomingTransferRequests(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'getOutgoingTransferRequests':
//     return ContentService.createTextOutput(JSON.stringify(getOutgoingTransferRequests(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'approveTransfer':
//     return ContentService.createTextOutput(JSON.stringify(approveTransfer(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'rejectTransfer':
//     return ContentService.createTextOutput(JSON.stringify(rejectTransfer(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//   case 'cancelTransferRequest':
//     return ContentService.createTextOutput(JSON.stringify(cancelTransferRequest(e.parameter))).setMimeType(ContentService.MimeType.JSON);