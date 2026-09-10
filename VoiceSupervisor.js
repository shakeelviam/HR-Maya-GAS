// ============================================================
// VoiceSupervisor.gs
// ------------------------------------------------------------
// Backend methods for the Maya HR Voice Agent + shared supervisor
// scoping used by the Roster Planner (RosterBackend.gs).
//
// FIXES IN THIS VERSION:
//  1. Header lookup is now case-insensitive + whitespace-tolerant.
//     Previously vsColMap_ required an EXACT match ('Employee ID'),
//     so any difference in case, a trailing space, or a non-breaking
//     space in the HR Maya header row silently fell through to the
//     `?? 0` / `?? -1` defaults — reading the wrong column entirely.
//  2. Location matching is now case-insensitive + trimmed, and also
//     normalises internal whitespace and non-breaking spaces. This is
//     what caused supervisors to see only SOME of their team: an
//     employee whose Location cell read 'Ardiya ' or 'ardiya' did not
//     match a supervisor whose cell read 'Ardiya'.
//  3. If the Location column genuinely cannot be found, vsGetTeam_ now
//     returns an EMPTY team and logs loudly, instead of the old
//     behaviour where HR_LOC = -1 made every loc '' and matching
//     collapsed unpredictably.
//  4. pad_ renamed to vsPad_ — 'pad_' is a bare global in a shared
//     Apps Script namespace and could silently collide with a same-named
//     helper in any other .gs file in this project.
//
// All methods require supId + supPin. Scope is enforced: supervisors
// only affect employees at their own location. Every change is logged
// to Attendance_Amendments.
// ============================================================

// ── Shared helpers ────────────────────────────────────────────────────

var VS_CACHE_ = {}; // in-memory cache per execution

// Normalise any header/text cell for comparison: strip non-breaking
// spaces, collapse internal whitespace, trim, lowercase.
function vsNormKey_(v) {
  return String(v == null ? '' : v)
    .replace(/\u00A0/g, ' ')      // non-breaking space → normal space
    .replace(/\s+/g, ' ')          // collapse internal whitespace
    .trim()
    .toLowerCase();
}

// Normalise a VALUE for comparison (locations, statuses). Same rules,
// but kept separate for clarity at call sites.
function vsNormVal_(v) {
  return vsNormKey_(v);
}

// Build a header → column-index map that is case-insensitive and
// whitespace-tolerant. Returns a lookup FUNCTION, not a plain object,
// so callers can't accidentally get `undefined` and fall through to a
// hardcoded index.
function vsColMap_(sheet) {
  const h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const m = {};
  h.forEach((v, i) => {
    const k = vsNormKey_(v);
    if (k && m[k] === undefined) m[k] = i;   // first occurrence wins
  });
  return m;
}

// Look up a column index by any of several candidate header names.
// Returns -1 if none found. Callers MUST handle -1 rather than
// falling back to a guessed index.
function vsCol_(colMap, candidates) {
  for (const c of candidates) {
    const idx = colMap[vsNormKey_(c)];
    if (idx !== undefined) return idx;
  }
  return -1;
}

function vsPad_(n) { return String(n).padStart(2, '0'); }

function vsNormDd_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return vsPad_(val.getDate()) + '-' + vsPad_(val.getMonth() + 1) + '-' + val.getFullYear();
  }
  const s = String(val).trim();
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  return s;
}

function vsToday_() {
  const d = new Date();
  return vsPad_(d.getDate()) + '-' + vsPad_(d.getMonth() + 1) + '-' + d.getFullYear();
}

function vsDateStr_(offset) {
  const d = new Date();
  d.setDate(d.getDate() + offset);
  return vsPad_(d.getDate()) + '-' + vsPad_(d.getMonth() + 1) + '-' + d.getFullYear();
}

function vsResolveDate_(raw) {
  if (!raw) return vsToday_();
  const s = String(raw).trim().toLowerCase();
  if (s === 'today')     return vsToday_();
  if (s === 'yesterday') return vsDateStr_(-1);
  if (/^\d{2}-\d{2}-\d{4}$/.test(raw)) return raw;
  const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  return vsToday_();
}

// ── HR Maya column resolution (single source of truth) ────────────────
// Every function that reads HR Maya goes through this, so a header
// change only has to be handled in one place.
function vsHrCols_(hrSheet) {
  const map = vsColMap_(hrSheet);
  return {
    id:     vsCol_(map, ['Employee ID', 'EmployeeID', 'Emp ID', 'ID']),
    name:   vsCol_(map, ['Name (English)', 'Name English', 'Name', 'Employee Name']),
    loc:    vsCol_(map, ['Location', 'Branch', 'Work Location']),
    status: vsCol_(map, ['Status', 'Employment Status']),
    pin:    vsCol_(map, ['PIN', 'Pin Code', 'Kiosk PIN']),
    shift:  vsCol_(map, ['Shift', 'Shift Name']),
  };
}

// ── Supervisor authentication ─────────────────────────────────────────

function voiceLogin(p) {
  try {
    const { supId, supPin } = p;
    if (!supId || !supPin) return { success: false, error: 'ID and PIN required.' };

    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };

    const team = vsGetTeam_(profile.location);
    return {
      success:   true,
      name:      profile.name,
      location:  profile.location,
      teamCount: team.length,
      teamIds:   team.map(e => e.id),
    };
  } catch (e) {
    Logger.log('voiceLogin: ' + e.message);
    return { success: false, error: 'Server error.' };
  }
}

// ── Team status ───────────────────────────────────────────────────────

function voiceTeamStatus(p) {
  try {
    const { supId, supPin, date } = p;
    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };

    const dateStr = vsResolveDate_(date);
    const team    = vsGetTeam_(profile.location);
    const ss      = SpreadsheetApp.getActiveSpreadsheet();

    const adSheet = ss.getSheetByName('Attendance_Daily');
    if (!adSheet) return { success: false, error: 'Attendance_Daily sheet not found.' };

    const adMap   = vsColMap_(adSheet);
    const AD_ID   = vsCol_(adMap, ['Employee ID']);
    const AD_DATE = vsCol_(adMap, ['Date']);
    const AD_STS  = vsCol_(adMap, ['Status']);
    const AD_CI   = vsCol_(adMap, ['Check In']);
    const AD_CO   = vsCol_(adMap, ['Check Out']);

    if (AD_ID < 0 || AD_DATE < 0 || AD_STS < 0) {
      return { success: false, error: 'Attendance_Daily is missing Employee ID / Date / Status columns.' };
    }

    const adData = adSheet.getLastRow() > 1
      ? adSheet.getRange(2, 1, adSheet.getLastRow() - 1, adSheet.getLastColumn()).getValues()
      : [];

    const dayMap = {};
    adData.forEach(row => {
      if (vsNormDd_(row[AD_DATE]) !== dateStr) return;
      const id = String(row[AD_ID] || '').trim();
      if (id) dayMap[id] = {
        status:   String(row[AD_STS] || '').trim(),
        checkIn:  AD_CI >= 0 ? String(row[AD_CI] || '').trim() : '',
        checkOut: AD_CO >= 0 ? String(row[AD_CO] || '').trim() : '',
      };
    });

    const present = [], absent = [], onLeave = [], noData = [], openSessions = [];

    team.forEach(emp => {
      const rec = dayMap[emp.id];
      if (!rec) {
        noData.push(emp);
      } else if (rec.status === 'Present') {
        if (rec.checkIn && !rec.checkOut) openSessions.push({ id: emp.id, name: emp.name, checkIn: rec.checkIn });
        else present.push(emp);
      } else if (rec.status === 'Absent') {
        absent.push(emp);
      } else if (rec.status === 'On Leave') {
        onLeave.push(emp);
      } else {
        noData.push(emp);
      }
    });

    return { success: true, date: dateStr, present, absent, onLeave, noData, openSessions };
  } catch (e) {
    Logger.log('voiceTeamStatus: ' + e.message);
    return { success: false, error: 'Server error.' };
  }
}

// ── Employee search ───────────────────────────────────────────────────

function voiceFindEmployee(p) {
  try {
    const { supId, supPin, query } = p;
    if (!query) return { success: false, error: 'Search query required.' };
    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };

    const q    = String(query).trim().toLowerCase();
    const team = vsGetTeam_(profile.location);
    const hits = team.filter(e =>
      e.name.toLowerCase().includes(q) || e.id.toLowerCase().includes(q)
    );
    return { success: true, matches: hits };
  } catch (e) {
    return { success: false, error: 'Server error.' };
  }
}

// ── Punch management ──────────────────────────────────────────────────

function voiceAddPunch(p) {
  try {
    const { supId, supPin, empId, date, stage, time } = p;
    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };

    if (!vsTeamContains_(profile.location, empId))
      return { success: false, error: 'Employee is not in your team.' };

    const VALID_STAGES = ['Check In', 'Break Out', 'Break In', 'Check Out', 'Day Off'];
    if (!VALID_STAGES.includes(stage))
      return { success: false, error: 'Invalid stage. Say Check In, Break Out, Break In, Check Out, or Day Off.' };

    // Day Off doesn't need a time; punch stages do.
    if (stage !== 'Day Off' && (!time || !/^\d{1,2}:\d{2}$/.test(time)))
      return { success: false, error: 'Time required in HH:MM format.' };

    const dateStr  = vsResolveDate_(date);
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Attendance_Log');
    if (!logSheet) return { success: false, error: 'Attendance_Log sheet not found.' };
    const empName  = vsEmpName_(empId);

    const parts     = dateStr.split('-'); // dd-mm-yyyy
    const timeParts = (stage === 'Day Off' ? '00:00' : time).split(':');
    const ts = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]),
                        parseInt(timeParts[0]), parseInt(timeParts[1]), 0);
    const tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    logSheet.appendRow([
      tsStr, empId, empName, dateStr, stage,
      'VOICE-SUPERVISOR', '', '', '', 'Manual by ' + profile.name
    ]);

    vsAudit_(empId, empName, dateStr,
      'ADD PUNCH: ' + stage + (stage === 'Day Off' ? '' : ' at ' + time), profile.name);

    return { success: true, message: stage + ' added for ' + empName + ' on ' + dateStr +
      (stage === 'Day Off' ? '.' : ' at ' + time + '.') };
  } catch (e) {
    Logger.log('voiceAddPunch: ' + e.message);
    return { success: false, error: 'Server error: ' + e.message };
  }
}

function voiceDeletePunch(p) {
  try {
    const { supId, supPin, empId, date, stage } = p;
    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };

    if (!vsTeamContains_(profile.location, empId))
      return { success: false, error: 'Employee is not in your team.' };

    const dateStr  = vsResolveDate_(date);
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Attendance_Log');
    if (!logSheet) return { success: false, error: 'Attendance_Log sheet not found.' };

    const logMap  = vsColMap_(logSheet);
    const L_ID    = vsCol_(logMap, ['Employee ID']);
    const L_DATE  = vsCol_(logMap, ['Date']);
    const L_STAGE = vsCol_(logMap, ['Stage']);
    if (L_ID < 0 || L_DATE < 0 || L_STAGE < 0)
      return { success: false, error: 'Attendance_Log is missing Employee ID / Date / Stage columns.' };

    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'No punches found.' };

    const data = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues();
    let deleted = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][L_ID] || '').trim() === empId &&
          vsNormDd_(data[i][L_DATE])          === dateStr &&
          String(data[i][L_STAGE] || '').trim() === stage) {
        logSheet.deleteRow(i + 2);
        deleted++;
        break;
      }
    }
    if (!deleted) return { success: false, error: 'No ' + stage + ' punch found for ' + empId + ' on ' + dateStr + '.' };

    const empName = vsEmpName_(empId);
    vsAudit_(empId, empName, dateStr, 'DELETE PUNCH: ' + stage, profile.name);
    return { success: true, message: stage + ' deleted for ' + empName + ' on ' + dateStr + '.' };
  } catch (e) {
    Logger.log('voiceDeletePunch: ' + e.message);
    return { success: false, error: 'Server error: ' + e.message };
  }
}

function voiceMarkAbsent(p) {
  try {
    const { supId, supPin, empId, date } = p;
    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };
    if (!vsTeamContains_(profile.location, empId))
      return { success: false, error: 'Employee is not in your team.' };

    const dateStr = vsResolveDate_(date);
    const empName = vsEmpName_(empId);
    vsWriteManualStatus_(empId, empName, profile.location, dateStr, 'Absent', profile.name);
    vsAudit_(empId, empName, dateStr, 'MARK ABSENT', profile.name);
    return { success: true, message: empName + ' marked Absent on ' + dateStr + '.' };
  } catch (e) {
    return { success: false, error: 'Server error: ' + e.message };
  }
}

function voiceMarkDayOff(p) {
  try {
    const { supId, supPin, empId, date } = p;
    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };
    if (!vsTeamContains_(profile.location, empId))
      return { success: false, error: 'Employee is not in your team.' };

    const dateStr  = vsResolveDate_(date);
    const empName  = vsEmpName_(empId);
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Attendance_Log');
    if (!logSheet) return { success: false, error: 'Attendance_Log sheet not found.' };

    const parts = dateStr.split('-');
    const ts    = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]), 0, 0, 0);
    const tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    logSheet.appendRow([tsStr, empId, empName, dateStr, 'Day Off',
                        'VOICE-SUPERVISOR', '', '', '', 'Logged by ' + profile.name]);

    vsAudit_(empId, empName, dateStr, 'MARK DAY OFF', profile.name);
    return { success: true, message: empName + ' recorded as Day Off on ' + dateStr + '.' };
  } catch (e) {
    return { success: false, error: 'Server error: ' + e.message };
  }
}

function voiceUndoLastAction(p) {
  try {
    const { supId, supPin, empId } = p;
    const profile = vsAuthSupervisor_(supId, supPin);
    if (!profile.ok) return { success: false, error: profile.error };
    if (!vsTeamContains_(profile.location, empId))
      return { success: false, error: 'Employee is not in your team.' };

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Attendance_Log');
    if (!logSheet) return { success: false, error: 'Attendance_Log sheet not found.' };

    const logMap = vsColMap_(logSheet);
    const L_ID   = vsCol_(logMap, ['Employee ID']);
    const L_AUTH = vsCol_(logMap, ['Auth', 'PIN Used']);
    if (L_ID < 0 || L_AUTH < 0)
      return { success: false, error: 'Attendance_Log is missing Employee ID / Auth columns.' };

    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Nothing to undo.' };

    const data = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][L_ID] || '').trim() === empId &&
          String(data[i][L_AUTH] || '').trim() === 'VOICE-SUPERVISOR') {
        logSheet.deleteRow(i + 2);
        const nm = vsEmpName_(empId);
        vsAudit_(empId, nm, '', 'UNDO LAST ACTION', profile.name);
        return { success: true, message: 'Last voice action for ' + nm + ' has been undone.' };
      }
    }
    return { success: false, error: 'No recent voice action found to undo.' };
  } catch (e) {
    return { success: false, error: 'Server error: ' + e.message };
  }
}

// ── Private helpers ───────────────────────────────────────────────────

// Authenticate supervisor — returns { ok, name, location, error }
function vsAuthSupervisor_(supId, supPin) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const hrSheet = ss.getSheetByName('HR Maya');
    if (!hrSheet) return { ok: false, error: 'HR Maya sheet not found.' };

    const C = vsHrCols_(hrSheet);
    if (C.id  < 0) return { ok: false, error: 'HR Maya: Employee ID column not found.' };
    if (C.pin < 0) return { ok: false, error: 'HR Maya: PIN column not found.' };

    const data = hrSheet.getRange(2, 1, hrSheet.getLastRow() - 1, hrSheet.getLastColumn()).getValues();
    const row  = data.find(r => String(r[C.id] || '').trim() === String(supId).trim());
    if (!row) return { ok: false, error: 'Supervisor ID not found.' };

    const status = C.status >= 0 ? String(row[C.status] || '').trim() : 'Active';
    if (status === 'Inactive') return { ok: false, error: 'Account is inactive.' };

    // PIN may be stored as a number; normalise both sides (strip trailing .0)
    let rowPin = String(row[C.pin] || '').trim();
    if (rowPin.endsWith('.0')) rowPin = rowPin.slice(0, -2);
    if (rowPin !== String(supPin).trim()) return { ok: false, error: 'Incorrect PIN.' };

    const loc = C.loc >= 0 ? String(row[C.loc] || '').trim() : '';
    if (!loc) return { ok: false, error: 'No location set on your HR record. Contact HR.' };

    const isSup = vsIsSupervisorOfLocation_(supId, loc);
    if (!isSup) return { ok: false, error: 'You are not registered as a supervisor.' };

    return {
      ok: true,
      name: C.name >= 0 ? String(row[C.name] || '').trim() : String(supId),
      location: loc,
    };
  } catch (e) {
    return { ok: false, error: 'Auth error: ' + e.message };
  }
}

// Check if empId is a registered supervisor for the given location.
// Optional Admin_Settings row: key 'supervisors', value = comma-separated IDs.
// If that row is absent/empty, any active employee at the location may act
// as supervisor (existing behaviour, unchanged).
function vsIsSupervisorOfLocation_(supId, location) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminSheet = ss.getSheetByName('Admin_Settings');
    if (adminSheet && adminSheet.getLastRow() > 0) {
      const adminData = adminSheet.getRange(1, 1, adminSheet.getLastRow(), 2).getValues();
      const supRow = adminData.find(r => vsNormKey_(r[0]) === 'supervisors');
      if (supRow) {
        const supList = String(supRow[1] || '').split(',').map(s => s.trim()).filter(Boolean);
        if (supList.length) return supList.includes(String(supId).trim());
      }
    }
    return true;
  } catch (e) { return true; }
}

// Get all active employees at a location.
// FIX: comparison is now case-insensitive, trimmed, whitespace-normalised.
// FIX: if the Location column is missing entirely, return [] and log —
// never silently match everyone or no-one by accident.
function vsGetTeam_(location) {
  const key = 'team_' + vsNormVal_(location);
  if (VS_CACHE_[key]) return VS_CACHE_[key];

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hrSheet = ss.getSheetByName('HR Maya');
  if (!hrSheet) { Logger.log('vsGetTeam_: HR Maya sheet not found.'); return []; }

  const C = vsHrCols_(hrSheet);
  if (C.id < 0) {
    Logger.log('vsGetTeam_: Employee ID column not found in HR Maya. Headers must include "Employee ID".');
    return [];
  }
  if (C.loc < 0) {
    Logger.log('vsGetTeam_: Location column not found in HR Maya. Headers must include "Location" or "Branch".');
    return [];
  }

  const wanted = vsNormVal_(location);
  const data   = hrSheet.getRange(2, 1, hrSheet.getLastRow() - 1, hrSheet.getLastColumn()).getValues();

  const team = data
    .filter(r => {
      const loc = vsNormVal_(r[C.loc]);
      const st  = C.status >= 0 ? vsNormVal_(r[C.status]) : 'active';
      return loc === wanted && st !== 'inactive';
    })
    .map(r => ({
      id:   String(r[C.id] || '').trim(),
      name: C.name >= 0 ? String(r[C.name] || '').trim() : String(r[C.id] || '').trim(),
    }))
    .filter(e => e.id);

  VS_CACHE_[key] = team;
  return team;
}

function vsTeamContains_(location, empId) {
  return vsGetTeam_(location).some(e => e.id === String(empId).trim());
}

function vsEmpName_(empId) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hrSheet = ss.getSheetByName('HR Maya');
  if (!hrSheet) return empId;
  const C    = vsHrCols_(hrSheet);
  if (C.id < 0) return empId;
  const data = hrSheet.getRange(2, 1, hrSheet.getLastRow() - 1, hrSheet.getLastColumn()).getValues();
  const row  = data.find(r => String(r[C.id] || '').trim() === String(empId).trim());
  if (!row) return empId;
  return C.name >= 0 ? String(row[C.name] || '').trim() : empId;
}

function vsWriteManualStatus_(empId, empName, location, dateStr, status, supervisorName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Attendance_Log');
  if (!logSheet) throw new Error('Attendance_Log sheet not found.');
  const parts = dateStr.split('-');
  const ts    = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]), 0, 0, 0);
  const tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  logSheet.appendRow([tsStr, empId, empName, dateStr, status,
                      'VOICE-SUPERVISOR', '', '', '', 'By ' + supervisorName]);
}

function vsAudit_(empId, empName, date, action, supervisorName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let auditSheet = ss.getSheetByName('Attendance_Amendments');
  if (!auditSheet) {
    auditSheet = ss.insertSheet('Attendance_Amendments');
    auditSheet.getRange(1, 1, 1, 7).setValues([['When', 'By', 'Action', 'Employee ID', 'Date', 'Stage', 'Detail (old → new)']]);
    auditSheet.setFrozenRows(1);
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
  auditSheet.appendRow([now, supervisorName + ' (VOICE)', action, empId, date, '', empName]);
}

// ============================================================
// DIAGNOSTIC — run this from the Apps Script editor to see exactly
// what the code reads from HR Maya. Use it to confirm the team fix.
// Select vsDiagnoseTeams from the function dropdown → Run → View logs.
// ============================================================
function vsDiagnoseTeams() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hrSheet = ss.getSheetByName('HR Maya');
  if (!hrSheet) { Logger.log('HR Maya sheet NOT FOUND'); return; }

  const rawHeaders = hrSheet.getRange(1, 1, 1, hrSheet.getLastColumn()).getValues()[0];
  Logger.log('=== RAW HEADERS (exact text, quoted) ===');
  rawHeaders.forEach((h, i) => Logger.log('  col ' + i + ': "' + h + '"'));

  const C = vsHrCols_(hrSheet);
  Logger.log('=== RESOLVED COLUMNS (-1 means NOT FOUND) ===');
  Logger.log('  id=' + C.id + '  name=' + C.name + '  loc=' + C.loc +
             '  status=' + C.status + '  pin=' + C.pin + '  shift=' + C.shift);

  if (C.loc < 0) { Logger.log('!! Location column not found — that is the bug.'); return; }

  const data = hrSheet.getRange(2, 1, hrSheet.getLastRow() - 1, hrSheet.getLastColumn()).getValues();
  const counts = {};
  data.forEach(r => {
    const id = String(r[C.id] || '').trim();
    if (!id) return;
    const rawLoc  = String(r[C.loc] || '');
    const normLoc = vsNormVal_(rawLoc);
    if (!counts[normLoc]) counts[normLoc] = { raw: {}, n: 0 };
    counts[normLoc].n++;
    counts[normLoc].raw['"' + rawLoc + '"'] = (counts[normLoc].raw['"' + rawLoc + '"'] || 0) + 1;
  });

  Logger.log('=== LOCATIONS (normalised → count, and the RAW spellings behind it) ===');
  Object.keys(counts).sort().forEach(k => {
    Logger.log('  "' + k + '" → ' + counts[k].n + ' employee(s)');
    Object.keys(counts[k].raw).forEach(rawVariant => {
      Logger.log('        raw ' + rawVariant + ' × ' + counts[k].raw[rawVariant]);
    });
  });
  Logger.log('If one location shows MULTIPLE raw spellings above, those rows');
  Logger.log('used to be split across different teams. They now match correctly.');
}