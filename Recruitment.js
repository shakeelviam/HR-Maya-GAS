// ============================================================
// Recruitment.gs — candidates from abroad, up to the day they start
// ------------------------------------------------------------
// A person being recruited is not an employee yet: no Civil ID, no
// Employee ID, no attendance, no pay. They live in the Recruitment tab
// until their first working day, when ONE action turns the candidate row
// into a real HR Maya row — nothing is typed twice, and the candidate row
// stays as the record of how they got here.
//
// Five stages (Shakeel, 23-09-2026):
//   Selected → Visa Issued → Travelling → Arrived → Employee
// plus On Hold and Cancelled, which can be set at any point.
// Each stage stamps its own date column, so you can see where someone is
// stuck, not just where they are.
//
// Tabs (created on first use, never re-created):
//   Agencies         — who supplies candidates. Agency details are rows
//                      here, not values in the code.
//   Recruitment      — one row per candidate, MC-0001 upward.
//   Recruitment_Log  — every stage change: when, who, from, to, note.
//
// ENDPOINTS (routed in Code.gs):
//   getRecruitment, saveCandidate (POST), setCandidateStage (POST),
//   convertCandidateToEmployee (POST), saveAgency (POST),
//   issueAgencyPin (POST), revokeAgencyPin (POST), checkAgencyLogin (POST),
//   agentCandidates (POST), agentAddCandidate (POST)  ← the agent portal
//
// Conversion reuses saveEmployee in EmployeeAdmin.gs — the Employee ID,
// the PIN, the derived salary columns and the Status all come from there,
// so recruitment cannot drift from how employees are normally created.
// ============================================================

const RECRUIT_TAB = 'Recruitment';
const RECRUIT_LOG_TAB = 'Recruitment_Log';
const AGENCY_TAB = 'Agencies';

const RECRUIT_STAGES = ['Selected', 'Visa Issued', 'Travelling', 'Arrived', 'Employee'];
const RECRUIT_OTHER  = ['On Hold', 'Cancelled'];
// Where an agent's own submission lands. It is NOT one of the five stages:
// nobody is in the pipeline until HR has looked at them and moved them to
// Selected. Agents can create this status and nothing else.
const RECRUIT_SUBMITTED = 'Submitted';

// stage → the column its date is stamped in
const RECRUIT_STAGE_DATE = {
  'Selected':    'Selected On',
  'Visa Issued': 'Visa Issued On',
  'Travelling':  'Travel Date',
  'Arrived':     'Arrived On',
  'Employee':    'Joined On'
};

const RECRUIT_HEADERS = [
  'Candidate ID', 'Name (English)', 'Nationality', 'Passport No', 'Passport Expiry',
  'Mobile / WhatsApp', 'Agency ID', 'Agency', 'Position', 'Location', 'Shift',
  'Basic Salary', 'Food Allowance', 'Accommodation', 'Conveyance', 'Contract (years)',
  // What the walk-in form already asks every candidate. An agent submitting
  // someone from abroad answers the same questions, so the two sources of
  // candidates hold the same facts and can be read side by side.
  'Visa Status', 'Experience (years)', 'Currently Employed', 'Available From',
  'Expected Salary (KD)', 'Languages', 'Photo', 'CV',
  'Status', 'Selected On', 'Visa Issued On', 'Travel Date', 'Arrived On', 'Joined On',
  'Employee ID', 'Remark', 'Updated By', 'Updated On'
];

// Where candidate photos and CVs are filed.
const RECRUIT_FILE_FOLDER = 'Maya Recruitment Files';

const AGENCY_HEADERS = ['Agency ID', 'Agency Name', 'License No', 'Country',
                        'Contact Person', 'WhatsApp', 'Email', 'Status', 'Notes',
                        'Portal PIN', 'Portal Status', 'PIN Issued On'];

// Portal credentials. The agent signs in with their Agency ID and this PIN,
// so each agency can be cut off on its own — no shared link, no shared
// password. Six digits, unique across agencies, and re-issuing replaces the
// old one immediately. (The portal page itself is the next phase; the
// credentials live here so they are ready and revocable from day one.)
const AGENCY_PIN_LENGTH = 6;

const RECRUIT_LOG_HEADERS = ['When', 'Candidate ID', 'Name', 'From', 'To', 'By', 'Note'];

// ── sheet helpers ────────────────────────────────────────────────────
function recSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // A tab made by an earlier version is missing the newer columns. Append
  // them rather than rebuilding the tab — the rows in it are real work.
  const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
                 .map(h => String(h).trim().toLowerCase());
  const missing = headers.filter(h => have.indexOf(String(h).trim().toLowerCase()) === -1);
  if (missing.length) {
    sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh;
}
function recCols_(sh) {
  const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  return { at: (n) => H.indexOf(String(n).trim().toLowerCase()), width: H.length };
}
function recToday_() {
  return (typeof formatDdMmYyyy_ === 'function') ? formatDdMmYyyy_(new Date())
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
}
function recDateStr_(v) {
  if (v === '' || v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return (typeof formatDdMmYyyy_ === 'function') ? formatDdMmYyyy_(v)
      : Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }
  return String(v).trim();
}
function recWho_(params) {
  const given = String((params && params.by) || '').trim();
  if (given) return given;
  try { return Session.getEffectiveUser().getEmail() || 'HR'; } catch (e) { return 'HR'; }
}
function recNextId_(sh, c) {
  const last = sh.getLastRow();
  if (last < 2) return 'MC-0001';
  const ids = sh.getRange(2, c.at('Candidate ID') + 1, last - 1, 1).getValues();
  let max = 0;
  ids.forEach(r => { const m = String(r[0] || '').match(/MC-(\d+)/i); if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; } });
  return 'MC-' + String(max + 1).padStart(4, '0');
}
function recLog_(ss, candId, name, from, to, by, note) {
  const sh = recSheet_(ss, RECRUIT_LOG_TAB, RECRUIT_LOG_HEADERS);
  sh.appendRow([recToday_(), candId, name, from || '', to || '', by || '', note || '']);
}

// ── read: everything the Recruitment page needs in one call ──────────
function getRecruitment(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, RECRUIT_TAB, RECRUIT_HEADERS);
    const c = recCols_(sh);
    const rows = [];
    if (sh.getLastRow() > 1) {
      const d = sh.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        const id = String(d[i][c.at('Candidate ID')] || '').trim();
        if (!id) continue;
        const rec = {};
        RECRUIT_HEADERS.forEach(h => {
          const j = c.at(h);
          if (j === -1) return;
          rec[h] = (h.indexOf('On') === h.length - 2 || h.indexOf('Date') !== -1 || h === 'Passport Expiry')
            ? recDateStr_(d[i][j]) : d[i][j];
        });
        rows.push(rec);
      }
    }
    const wanted = String(params.status || '').trim().toLowerCase();
    const data = wanted ? rows.filter(r => String(r['Status']).trim().toLowerCase() === wanted) : rows;

    // Agencies, so the page can offer them instead of free text.
    const ag = recSheet_(ss, AGENCY_TAB, AGENCY_HEADERS);
    const ac = recCols_(ag);
    const agencies = [];
    if (ag.getLastRow() > 1) {
      const ad = ag.getDataRange().getValues();
      for (let i = 1; i < ad.length; i++) {
        const id = String(ad[i][ac.at('Agency ID')] || '').trim();
        if (!id) continue;
        agencies.push({ id: id, name: String(ad[i][ac.at('Agency Name')] || '').trim(),
          contact: String(ad[i][ac.at('Contact Person')] || '').trim(),
          whatsapp: String(ad[i][ac.at('WhatsApp')] || '').trim(),
          country: String(ad[i][ac.at('Country')] || '').trim(),
          status: String(ad[i][ac.at('Status')] || '').trim() || 'Active',
          license: String(ad[i][ac.at('License No')] || '').trim(),
          portal: (ac.at('Portal PIN') !== -1 && String(ad[i][ac.at('Portal PIN')] || '').trim())
                  ? (String(ad[i][ac.at('Portal Status')] || '').trim() || 'Active') : 'None' });
      }
    }
    const counts = {};
    [RECRUIT_SUBMITTED].concat(RECRUIT_STAGES).concat(RECRUIT_OTHER).forEach(s => { counts[s] = 0; });
    rows.forEach(r => { const s = String(r['Status'] || '').trim(); if (s in counts) counts[s]++; });

    return { success: true, data: data, agencies: agencies, stages: RECRUIT_STAGES,
             other: RECRUIT_OTHER, submitted: RECRUIT_SUBMITTED, counts: counts, count: data.length };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ── add or edit a candidate ──────────────────────────────────────────
// params: candidateId? (edit), name, nationality, passportNo, passportExpiry,
//   mobile, agencyId, position, location, shift, basic, food, accom, conv,
//   contractYears, status?, remark?
function saveCandidate(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, RECRUIT_TAB, RECRUIT_HEADERS);
    const c = recCols_(sh);
    const name = String(params.name || '').trim();
    const isEdit = !!String(params.candidateId || '').trim();
    if (!isEdit && !name) return { success: false, error: 'Name is required (as printed in the passport).' };

    const d = sh.getDataRange().getValues();
    let rowNum = -1, vals;
    let candId;
    if (isEdit) {
      candId = String(params.candidateId).trim();
      for (let i = 1; i < d.length; i++) {
        if (String(d[i][c.at('Candidate ID')] || '').trim().toUpperCase() === candId.toUpperCase()) {
          rowNum = i + 1; vals = d[i].slice(); break;
        }
      }
      if (rowNum === -1) return { success: false, error: 'Candidate ' + candId + ' not found.' };
    } else {
      candId = recNextId_(sh, c);
      vals = new Array(c.width).fill('');
    }

    const set = (h, v) => { const j = c.at(h); if (j !== -1) vals[j] = v; };
    const given = (k) => params[k] !== undefined && params[k] !== '';
    const num = (v) => { const n = Number(v); return isNaN(n) ? 0 : n; };

    set('Candidate ID', candId);
    if (!isEdit || given('name'))           set('Name (English)', name);
    if (!isEdit || given('nationality'))    set('Nationality', String(params.nationality || '').trim());
    // Passport numbers are upper-case on the document; store them that way
    // whatever case they were typed in, so a search finds them.
    if (!isEdit || given('passportNo'))     set('Passport No', String(params.passportNo || '').trim().toUpperCase());
    if (!isEdit || given('passportExpiry')) set('Passport Expiry', recNormDate_(params.passportExpiry));
    if (!isEdit || given('mobile'))         set('Mobile / WhatsApp', String(params.mobile || '').trim());
    if (!isEdit || given('position'))       set('Position', String(params.position || '').trim());
    if (!isEdit || given('location'))       set('Location', String(params.location || '').trim());
    if (!isEdit || given('shift'))          set('Shift', String(params.shift || '').trim());
    if (!isEdit || given('basic'))          set('Basic Salary', num(params.basic));
    if (!isEdit || given('food'))           set('Food Allowance', num(params.food));
    if (!isEdit || given('accom'))          set('Accommodation', num(params.accom));
    if (!isEdit || given('conv'))           set('Conveyance', num(params.conv));
    if (!isEdit || given('contractYears'))  set('Contract (years)', String(params.contractYears || '').trim());
    if (!isEdit || given('visaStatus'))     set('Visa Status', String(params.visaStatus || '').trim());
    if (!isEdit || given('experience'))     set('Experience (years)', String(params.experience || '').trim());
    if (!isEdit || given('employed'))       set('Currently Employed', String(params.employed || '').trim());
    if (!isEdit || given('availability'))   set('Available From', String(params.availability || '').trim());
    if (!isEdit || given('expectedSalary')) set('Expected Salary (KD)', String(params.expectedSalary || '').trim());
    if (!isEdit || given('languages'))      set('Languages', Array.isArray(params.languages) ? params.languages.join(', ') : String(params.languages || '').trim());
    if (given('photoUrl'))                  set('Photo', String(params.photoUrl).trim());
    if (given('cvUrl'))                     set('CV', String(params.cvUrl).trim());
    if (!isEdit || given('remark'))         set('Remark', String(params.remark || '').trim());

    // Agency: store the ID and copy the name across for reading the tab.
    if (!isEdit || given('agencyId')) {
      const agId = String(params.agencyId || '').trim();
      set('Agency ID', agId);
      set('Agency', agId ? recAgencyName_(ss, agId) : '');
    }

    // Stage: a new candidate starts at Selected unless told otherwise.
    const wantStatus = String(params.status || '').trim();
    if (!isEdit) {
      const st = wantStatus || 'Selected';
      set('Status', st);
      const dateCol = RECRUIT_STAGE_DATE[st];
      if (dateCol) set(dateCol, String(params.stageDate || '').trim() || recToday_());
    } else if (wantStatus) {
      set('Status', wantStatus);
    }
    set('Updated By', recWho_(params));
    set('Updated On', recToday_());

    if (rowNum === -1) rowNum = sh.getLastRow() + 1;
    sh.getRange(rowNum, 1, 1, c.width).setValues([vals]);
    // Passport numbers and phone numbers are text, never numbers.
    ['Passport No', 'Mobile / WhatsApp'].forEach(h => {
      const j = c.at(h); if (j !== -1) sh.getRange(rowNum, j + 1).setNumberFormat('@');
    });
    if (!isEdit) recLog_(ss, candId, name, '', String(vals[c.at('Status')] || ''), recWho_(params), 'Added');

    return { success: true, candidateId: candId,
             message: (isEdit ? 'Updated ' : 'Added ') + candId + ' — ' + (name || candId) + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
  finally { try { lock.releaseLock(); } catch (ignore) {} }
}

function recAgencyName_(ss, agencyId) {
  const ag = ss.getSheetByName(AGENCY_TAB);
  if (!ag || ag.getLastRow() < 2) return '';
  const c = recCols_(ag);
  const d = ag.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c.at('Agency ID')] || '').trim().toUpperCase() === agencyId.toUpperCase()) {
      return String(d[i][c.at('Agency Name')] || '').trim();
    }
  }
  return '';
}

// ── move a candidate to another stage ────────────────────────────────
// params: { candidateId, status, date? (dd-mm-yyyy, default today), note? }
// 'Employee' is NOT set here — that happens only through conversion, so a
// candidate can never be marked hired without an HR Maya row existing.
function setCandidateStage(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, RECRUIT_TAB, RECRUIT_HEADERS);
    const c = recCols_(sh);
    const candId = String(params.candidateId || '').trim();
    const status = String(params.status || '').trim();
    if (!candId) return { success: false, error: 'candidateId required.' };
    if (RECRUIT_STAGES.concat(RECRUIT_OTHER).concat([RECRUIT_SUBMITTED]).indexOf(status) === -1)
      return { success: false, error: 'Unknown stage "' + status + '".' };
    if (status === 'Employee')
      return { success: false, error: 'Use Convert to Employee — that is what creates the HR Maya row.' };

    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c.at('Candidate ID')] || '').trim().toUpperCase() !== candId.toUpperCase()) continue;
      const rowNum = i + 1;
      const from = String(d[i][c.at('Status')] || '').trim();
      if (from === 'Employee') return { success: false, error: candId + ' is already an employee (' + d[i][c.at('Employee ID')] + ').' };
      const when = String(params.date || '').trim() || recToday_();

      sh.getRange(rowNum, c.at('Status') + 1).setValue(status);
      const dateCol = RECRUIT_STAGE_DATE[status];
      if (dateCol && c.at(dateCol) !== -1) sh.getRange(rowNum, c.at(dateCol) + 1).setValue(when);
      if (params.note) sh.getRange(rowNum, c.at('Remark') + 1).setValue(String(params.note).trim());
      sh.getRange(rowNum, c.at('Updated By') + 1).setValue(recWho_(params));
      sh.getRange(rowNum, c.at('Updated On') + 1).setValue(recToday_());

      recLog_(ss, candId, String(d[i][c.at('Name (English)')] || ''), from, status, recWho_(params),
              (params.note ? String(params.note) : '') + (dateCol ? ' (' + dateCol + ' ' + when + ')' : ''));
      return { success: true, candidateId: candId, status: status,
               message: candId + ' → ' + status + (dateCol ? ' on ' + when : '') + '.' };
    }
    return { success: false, error: 'Candidate ' + candId + ' not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ── the one that matters: candidate → employee ───────────────────────
// Runs on the FIRST WORKING DAY (Shakeel, 23-09-2026): Date of Join is
// that day, and attendance and payroll start from it.
// params: { candidateId, joinDate? (default today), designation?, location?,
//           shift?, basic?, food?, accom?, conv?, civilId? }
// Anything not passed comes from the candidate row.
function convertCandidateToEmployee(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    if (typeof saveEmployee !== 'function')
      return { success: false, error: 'EmployeeAdmin.gs is missing — saveEmployee not found.' };

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, RECRUIT_TAB, RECRUIT_HEADERS);
    const c = recCols_(sh);
    const candId = String(params.candidateId || '').trim();
    if (!candId) return { success: false, error: 'candidateId required.' };

    const d = sh.getDataRange().getValues();
    let rowNum = -1, row = null;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c.at('Candidate ID')] || '').trim().toUpperCase() === candId.toUpperCase()) { rowNum = i + 1; row = d[i]; break; }
    }
    if (rowNum === -1) return { success: false, error: 'Candidate ' + candId + ' not found.' };

    const get = (h) => { const j = c.at(h); return j === -1 ? '' : row[j]; };
    const existing = String(get('Employee ID') || '').trim();
    if (existing) return { success: false, error: candId + ' was already converted — ' + existing + '.' };

    const name = String(get('Name (English)') || '').trim();
    if (!name) return { success: false, error: 'Candidate has no name — fill it in before converting.' };
    const status = String(get('Status') || '').trim();
    if (status === 'Cancelled') return { success: false, error: candId + ' is Cancelled. Move them back to a stage first.' };

    const joinDate = String(params.joinDate || '').trim() || recToday_();
    const pick = (p, h) => (params[p] !== undefined && params[p] !== '') ? params[p] : get(h);

    // Create the employee through the normal path: Employee ID, PIN and the
    // derived salary columns all come from EmployeeAdmin.gs.
    const res = saveEmployee({
      nameEnglish: name,
      civilId: String(params.civilId || '').trim(),          // usually still blank — comes later via the docs form
      passportNo: String(get('Passport No') || '').trim(),
      passportExpiry: recDateStr_(get('Passport Expiry')),
      mobile: String(get('Mobile / WhatsApp') || '').trim(),
      whatsapp: String(get('Mobile / WhatsApp') || '').trim(),
      branch: String(pick('location', 'Location') || '').trim(),
      shift: String(pick('shift', 'Shift') || '').trim(),
      designation: String(pick('designation', 'Position') || '').trim(),
      basic: Number(pick('basic', 'Basic Salary')) || 0,
      food: Number(pick('food', 'Food Allowance')) || 0,
      accom: Number(pick('accom', 'Accommodation')) || 0,
      conv: Number(pick('conv', 'Conveyance')) || 0,
      dateOfJoin: joinDate,
      status: 'Active'
    });
    if (!res || !res.success) return { success: false, error: 'Employee not created: ' + ((res && res.error) || 'unknown') };

    const empId = String(res.employeeId || '').trim();
    const pin = String(res.pin || '').trim();

    // saveEmployee does not carry a Designation, so write it here — the
    // column exists in HR Maya and the roster reads it.
    const designation = String(pick('designation', 'Position') || '').trim();
    if (empId && designation) {
      const hr = ss.getSheetByName('HR Maya');
      const hd = hr.getDataRange().getValues();
      const HH = hd[0].map(h => String(h).trim().toLowerCase());
      const hId = HH.indexOf('employee id'), hDes = HH.indexOf('designation');
      if (hId !== -1 && hDes !== -1) {
        for (let i = 1; i < hd.length; i++) {
          if (String(hd[i][hId] || '').trim() === empId) { hr.getRange(i + 1, hDes + 1).setValue(designation); break; }
        }
      }
    }
    // Write the result back onto the candidate row.
    sh.getRange(rowNum, c.at('Status') + 1).setValue('Employee');
    sh.getRange(rowNum, c.at('Joined On') + 1).setValue(joinDate);
    if (empId) sh.getRange(rowNum, c.at('Employee ID') + 1).setValue(empId);
    sh.getRange(rowNum, c.at('Updated By') + 1).setValue(recWho_(params));
    sh.getRange(rowNum, c.at('Updated On') + 1).setValue(recToday_());

    recLog_(ss, candId, name, status, 'Employee', recWho_(params),
            'Converted' + (empId ? ' to ' + empId : '') + ', joined ' + joinDate);

    return {
      success: true, candidateId: candId, employeeId: empId, joinDate: joinDate,
      pin: pin,
      message: name + ' is now ' + (empId || 'an employee') + ', joined ' + joinDate +
        (pin ? ', kiosk PIN ' + pin : '') + '. ' +
        'Civil ID and health card follow through the staff document form.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
  finally { try { lock.releaseLock(); } catch (ignore) {} }
}

// ── agencies ─────────────────────────────────────────────────────────
// params: { agencyId? (edit), name, license, country, contact, whatsapp, email, status?, notes? }
function saveAgency(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, AGENCY_TAB, AGENCY_HEADERS);
    const c = recCols_(sh);
    const name = String(params.name || '').trim();
    const isEdit = !!String(params.agencyId || '').trim();
    if (!isEdit && !name) return { success: false, error: 'Agency name is required.' };

    const d = sh.getDataRange().getValues();
    let rowNum = -1, vals;
    let agId;
    if (isEdit) {
      agId = String(params.agencyId).trim();
      for (let i = 1; i < d.length; i++) {
        if (String(d[i][c.at('Agency ID')] || '').trim().toUpperCase() === agId.toUpperCase()) { rowNum = i + 1; vals = d[i].slice(); break; }
      }
      if (rowNum === -1) return { success: false, error: 'Agency ' + agId + ' not found.' };
    } else {
      let max = 0;
      for (let i = 1; i < d.length; i++) {
        const m = String(d[i][c.at('Agency ID')] || '').match(/AG-(\d+)/i);
        if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
      }
      agId = 'AG-' + String(max + 1).padStart(2, '0');
      vals = new Array(c.width).fill('');
    }
    const set = (h, v) => { const j = c.at(h); if (j !== -1) vals[j] = v; };
    const given = (k) => params[k] !== undefined && params[k] !== '';
    set('Agency ID', agId);
    if (!isEdit || given('name'))     set('Agency Name', name);
    if (!isEdit || given('license'))  set('License No', String(params.license || '').trim());
    if (!isEdit || given('country'))  set('Country', String(params.country || '').trim());
    if (!isEdit || given('contact'))  set('Contact Person', String(params.contact || '').trim());
    if (!isEdit || given('whatsapp')) set('WhatsApp', String(params.whatsapp || '').trim());
    if (!isEdit || given('email'))    set('Email', String(params.email || '').trim());
    if (!isEdit || given('notes'))    set('Notes', String(params.notes || '').trim());
    set('Status', String(params.status || 'Active').trim());

    if (rowNum === -1) rowNum = sh.getLastRow() + 1;
    sh.getRange(rowNum, 1, 1, c.width).setValues([vals]);
    const j = c.at('WhatsApp'); if (j !== -1) sh.getRange(rowNum, j + 1).setNumberFormat('@');
    return { success: true, agencyId: agId, message: (isEdit ? 'Updated ' : 'Added ') + agId + ' — ' + (name || agId) + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}


// ── portal credentials for an agency ─────────────────────────────────
function recUniqueAgencyPin_(sh, c) {
  const used = {};
  if (sh.getLastRow() > 1 && c.at('Portal PIN') !== -1) {
    sh.getRange(2, c.at('Portal PIN') + 1, sh.getLastRow() - 1, 1).getValues()
      .forEach(r => { const p = String(r[0] || '').trim(); if (p) used[p] = true; });
  }
  const min = Math.pow(10, AGENCY_PIN_LENGTH - 1);
  for (let i = 0; i < 20000; i++) {
    const pin = String(Math.floor(min + Math.random() * (min * 9)));
    if (!used[pin]) return pin;
  }
  return String(min + (Date.now() % (min * 9)));
}

// params: { agencyId }  → a new PIN, shown once here so it can be sent on.
function issueAgencyPin(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, AGENCY_TAB, AGENCY_HEADERS);
    const c = recCols_(sh);
    if (c.at('Portal PIN') === -1) return { success: false, error: 'The Agencies tab has no "Portal PIN" column. Delete the header row and reload the page, or add the column by hand.' };
    const agId = String(params.agencyId || '').trim();
    if (!agId) return { success: false, error: 'agencyId required.' };

    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c.at('Agency ID')] || '').trim().toUpperCase() !== agId.toUpperCase()) continue;
      const pin = recUniqueAgencyPin_(sh, c);
      const name = String(d[i][c.at('Agency Name')] || '').trim();
      sh.getRange(i + 1, c.at('Portal PIN') + 1).setNumberFormat('@');
      sh.getRange(i + 1, c.at('Portal PIN') + 1).setValue(pin);
      sh.getRange(i + 1, c.at('Portal Status') + 1).setValue('Active');
      sh.getRange(i + 1, c.at('PIN Issued On') + 1).setValue(recToday_());
      recLog_(ss, agId, name, '', 'PIN issued', recWho_(params), 'Portal credentials issued');
      return { success: true, agencyId: agId, name: name, pin: pin,
               message: 'Portal PIN for ' + (name || agId) + ': ' + pin + '  (Agency ID ' + agId + '). Send both to the agent — this replaces any earlier PIN.' };
    }
    return { success: false, error: 'Agency ' + agId + ' not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// params: { agencyId } → the agent can no longer sign in.
function revokeAgencyPin(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, AGENCY_TAB, AGENCY_HEADERS);
    const c = recCols_(sh);
    const agId = String(params.agencyId || '').trim();
    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c.at('Agency ID')] || '').trim().toUpperCase() !== agId.toUpperCase()) continue;
      const name = String(d[i][c.at('Agency Name')] || '').trim();
      if (c.at('Portal PIN') !== -1) sh.getRange(i + 1, c.at('Portal PIN') + 1).setValue('');
      if (c.at('Portal Status') !== -1) sh.getRange(i + 1, c.at('Portal Status') + 1).setValue('Revoked');
      recLog_(ss, agId, name, 'PIN issued', 'PIN revoked', recWho_(params), 'Portal access revoked');
      return { success: true, agencyId: agId, message: 'Portal access revoked for ' + (name || agId) + '.' };
    }
    return { success: false, error: 'Agency ' + agId + ' not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Who an agent is, for the portal to check a sign-in. Returns the agency
// only when the ID and PIN match and access is Active — never the PIN.
function checkAgencyLogin(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = recSheet_(ss, AGENCY_TAB, AGENCY_HEADERS);
    const c = recCols_(sh);
    const agId = String(params.agencyId || '').trim();
    const pin = String(params.pin || '').trim();
    if (!agId || !pin) return { success: false, error: 'Agency ID and PIN are required.' };
    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c.at('Agency ID')] || '').trim().toUpperCase() !== agId.toUpperCase()) continue;
      const stored = String(d[i][c.at('Portal PIN')] || '').trim();
      const pStatus = String(d[i][c.at('Portal Status')] || '').trim();
      const aStatus = String(d[i][c.at('Status')] || '').trim();
      if (!stored || pStatus === 'Revoked' || aStatus === 'Inactive')
        return { success: false, error: 'This agency has no portal access. Ask HR.' };
      if (stored !== pin) return { success: false, error: 'Wrong Agency ID or PIN.' };
      return { success: true, agencyId: agId, name: String(d[i][c.at('Agency Name')] || '').trim(),
               contact: String(d[i][c.at('Contact Person')] || '').trim() };
    }
    return { success: false, error: 'Wrong Agency ID or PIN.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}


// ══════════════════════════════════════════════════════════════════
// AGENT PORTAL — what an agency may do, and only for its own people
// ------------------------------------------------------------------
// Every call carries the Agency ID and PIN and is checked again here:
// the portal page proves nothing on its own. An agent sees only their
// own candidates, never anyone else's, never the HR-only columns
// (salary, location), and they cannot move anyone through the stages —
// they can only submit a new candidate, which lands as 'Submitted' for
// HR to accept.
// ══════════════════════════════════════════════════════════════════

// The agent's view of one candidate: progress, no internals.
function recAgentView_(rec) {
  return {
    id: rec['Candidate ID'], name: rec['Name (English)'],
    passport: rec['Passport No'], nationality: rec['Nationality'],
    position: rec['Position'], status: rec['Status'],
    selectedOn: rec['Selected On'], visaOn: rec['Visa Issued On'],
    travelDate: rec['Travel Date'], arrivedOn: rec['Arrived On'],
    joinedOn: rec['Joined On'], remark: rec['Remark'],
    mobile: rec['Mobile / WhatsApp'], visaStatus: rec['Visa Status'],
    experience: rec['Experience (years)'], languages: rec['Languages'],
    photo: rec['Photo'], cv: rec['CV']
  };
}

// params: { agencyId, pin }
function agentCandidates(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const who = checkAgencyLogin(params);
    if (!who.success) return who;
    const all = getRecruitment({});
    if (!all.success) return all;
    const mine = all.data.filter(r => String(r['Agency ID'] || '').trim().toUpperCase() === who.agencyId.toUpperCase());
    const counts = {};
    mine.forEach(r => { const st = String(r['Status'] || '').trim(); counts[st] = (counts[st] || 0) + 1; });
    return { success: true, agencyId: who.agencyId, name: who.name, contact: who.contact,
             data: mine.map(recAgentView_), counts: counts, count: mine.length };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// params: { agencyId, pin, name, passportNo, nationality?, mobile?, position?, remark? }
// Lands as 'Submitted'. Salary, location and shift are HR's to set, so the
// agent cannot send them at all.
function agentAddCandidate(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const who = checkAgencyLogin(params);
    if (!who.success) return who;
    const name = String(params.name || '').trim();
    const passport = String(params.passportNo || '').trim();
    if (!name) return { success: false, error: 'Name is required, as printed in the passport.' };
    if (!passport) return { success: false, error: 'Passport number is required.' };

    // A passport number already on file is the same man — do not duplicate him.
    const all = getRecruitment({});
    if (all.success) {
      const clash = all.data.filter(r => String(r['Passport No'] || '').trim().toUpperCase() === passport.toUpperCase())[0];
      if (clash) return { success: false, error: 'Passport ' + passport + ' is already submitted (' + clash['Candidate ID'] + ').' };
    }

    let photoUrl = '', cvUrl = '';
    try {
      photoUrl = recSaveFile_(params.photo, name, 'Photo');
      cvUrl = recSaveFile_(params.cv, name, 'CV');
    } catch (e) {
      // A failed upload must not lose the candidate — record them, say so.
      photoUrl = photoUrl || ''; cvUrl = cvUrl || '';
    }

    const res = saveCandidate({
      name: name, passportNo: passport,
      passportExpiry: String(params.passportExpiry || '').trim(),
      nationality: String(params.nationality || '').trim(),
      mobile: String(params.mobile || '').trim(),
      position: String(params.position || '').trim(),
      visaStatus: String(params.visaStatus || '').trim(),
      experience: String(params.experience || '').trim(),
      employed: String(params.employed || '').trim(),
      availability: String(params.availability || '').trim(),
      expectedSalary: String(params.expectedSalary || '').trim(),
      languages: params.languages,
      photoUrl: photoUrl, cvUrl: cvUrl,
      remark: String(params.remark || '').trim(),
      agencyId: who.agencyId,
      status: RECRUIT_SUBMITTED,
      by: who.name || who.agencyId
    });
    if (!res.success) return res;
    return { success: true, candidateId: res.candidateId,
             message: name + ' submitted. HR will review and come back to you.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// A date from anywhere — the HR page, the agent portal's date picker
// (yyyy-mm-dd), a paste — is stored the one way the workbook reads.
function recNormDate_(v) {
  const raw = String(v == null ? '' : v).trim();
  if (!raw) return '';
  if (typeof parseFlexibleDate_ === 'function' && typeof formatDdMmYyyy_ === 'function') {
    const d = parseFlexibleDate_(raw);
    if (d) return formatDdMmYyyy_(d);
  }
  return raw;
}

// ── attachments ──────────────────────────────────────────────────────
// A photo or CV arrives as base64 from the portal. It is written to Drive
// and only the link is kept in the sheet, so the tab stays small and the
// file can be opened by anyone in HR.
function recSaveFile_(file, who, kind) {
  if (!file || !file.data) return '';
  const folders = DriveApp.getFoldersByName(RECRUIT_FILE_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(RECRUIT_FILE_FOLDER);
  const safe = String(who || 'candidate').replace(/[^A-Za-z0-9 ]/g, '').trim() || 'candidate';
  const ext = String(file.name || '').match(/\.(\w{1,5})$/);
  const name = safe + ' — ' + kind + (ext ? ext[0] : '');
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.mimeType || 'application/octet-stream', name);
  const f = folder.createFile(blob);
  return f.getUrl();
}

// ── editor view: the pipeline as text ────────────────────────────────
function previewRecruitment() {
  const r = getRecruitment({});
  if (!r.success) { Logger.log(r.error); return r.error; }
  const out = ['Recruitment — ' + recToday_() + '  (' + r.count + ' candidates)'];
  [RECRUIT_SUBMITTED].concat(RECRUIT_STAGES).concat(RECRUIT_OTHER).forEach(st => {
    const rows = r.data.filter(x => String(x['Status']).trim() === st);
    if (!rows.length) return;
    out.push('');
    out.push(st + ' (' + rows.length + ')');
    rows.forEach(x => out.push('   ' + x['Candidate ID'] + '  ' + x['Name (English)'] +
      '  passport ' + (x['Passport No'] || '—') +
      '  ' + (x['Position'] || '—') +
      (x['Travel Date'] ? '  travel ' + x['Travel Date'] : '') +
      (x['Employee ID'] ? '  → ' + x['Employee ID'] : '') +
      (x['Remark'] ? '  · ' + x['Remark'] : '')));
  });
  const text = out.join('\n');
  Logger.log(text);
  return text;
}