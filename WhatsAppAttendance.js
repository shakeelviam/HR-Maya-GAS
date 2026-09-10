/**
 * Maya Tex — WhatsApp Attendance Bot Backend
 * ------------------------------------------------------------
 * Add this as a .gs file in the SAME project as Code.gs, Payroll.gs,
 * and SharedUtils.gs. Uses MAIN_SPREADSHEET_ID from Code.gs directly —
 * do not redeclare it here.
 *
 * Validates Employee IDs against 'HR Maya' (the same tab Payroll.gs
 * already trusts for salary data) — NOT a separate master sheet.
 *
 * Writes punches directly into 'Attendance_Log', extending your
 * existing kiosk format with four trailing columns for geofencing:
 *   A Timestamp (dd/MM/yyyy HH:mm:ss) | B Employee ID | C Employee Name
 *   | D Date (dd-mm-yyyy) | E Stage | F PIN Used
 *   | G Latitude | H Longitude | I Distance (m) | J Location Flag
 * WhatsApp-sourced rows use 'WA' in column F. Kiosk rows keep writing
 * only columns A-F, leaving G-J blank — no change needed there.
 *
 * Punches: Check In, Break Out, Break In, Check Out, Day Off.
 * Day Off doesn't need a location and is only offered when there's no
 * open shift (same moment "Check In" would be offered).
 *
 * Geofencing is a SOFT check — a punch is never blocked. After picking
 * a punch type (except Day Off), the bot asks the employee to share
 * their WhatsApp location, or reply "skip". Either way the punch is
 * logged; if location was shared, distance from the employee's branch
 * (looked up in the new 'Branch_Locations' tab) is recorded and flagged
 * if it exceeds that branch's radius.
 *
 * INTEGRATION: the 'waMessage' case in Code.gs's handleRequest()
 * routes here. Dispatch is by the URL parameter ?method=waMessage,
 * with the payload (secret, phone, text, location) as the JSON POST
 * body, matching your existing convention for markAttendance/etc.
 *
 * GROUP NOTIFICATIONS: whenever a real punch is actually logged
 * (any of the 4 stages, or Day Off), the return value includes an
 * extra `groupNotify` field alongside `reply`. index.js posts that
 * text to a configured WhatsApp group, in addition to replying to
 * the employee privately. Replies that don't represent a completed
 * punch (menus, prompts, errors) omit `groupNotify` entirely.
 */

// ====== CONFIG — fill this in ======
const WA_SHARED_SECRET = 'PASTE_A_LONG_RANDOM_STRING_HERE'; // must match index.js
// ====================================

const WA_SESSIONS_TAB = 'WA_Sessions';
const WA_SESSIONS_HEADERS = ['Phone', 'Employee ID', 'Name', 'Branch', 'State', 'Pending Action', 'Last Updated'];
const BRANCH_LOCATIONS_TAB = 'Branch_Locations';

const PUNCH_LABELS = { 1: 'Check In', 2: 'Break Out', 3: 'Break In', 4: 'Check Out', 5: 'Day Off' };
const PUNCH_NUMBERS = { 'Check In': 1, 'Break Out': 2, 'Break In': 3, 'Check Out': 4, 'Day Off': 5 };
const LOCATION_REQUIRED_ACTIONS = ['Check In', 'Break Out', 'Break In', 'Check Out']; // not Day Off
const GREETING_WORDS = ['hi', 'hello', 'hey', 'start', 'menu', 'salam', 'assalamualaikum'];

/**
 * Entry point, called from Code.gs's handleRequest() for
 * ?method=waMessage. payload: { secret, phone, text, location }
 * location, when present: { lat, lng }
 * returns: { reply: string|null }
 */
function handleWaMessage(payload) {
  if (!payload || payload.secret !== WA_SHARED_SECRET) {
    return { reply: null }; // silently ignore unauthenticated calls
  }

  const phone = String(payload.phone || '').trim();
  const rawText = String(payload.text || '').trim();
  const textLower = rawText.toLowerCase();
  const hasLocation = !!(payload.location && payload.location.lat != null && payload.location.lng != null);
  if (!phone || (!rawText && !hasLocation)) return { reply: null };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    // Another message from the same person arrived almost simultaneously
    // and is still being processed. Rather than fail silently (which
    // leaves the employee with no reply at all), ask them to resend.
    return { reply: 'Still processing your last message — please resend that in a few seconds.' };
  }

  try {
    const session = getSession_(phone);

    // Greeting always redraws the menu, EXCEPT mid-location-request —
    // don't let "hi" cancel a pending punch, could confuse someone
    // who's fumbling with the WhatsApp location picker.
    if (session.state !== 'AWAITING_LOCATION' && GREETING_WORDS.includes(textLower)) {
      if (session.employeeId) {
        saveSession_(phone, session.employeeId, session.name, session.branch, 'MENU', '');
        return { reply: buildMenuMessage_(session.employeeId) };
      }
      saveSession_(phone, '', '', '', 'AWAITING_ID', '');
      return { reply: 'Welcome to Maya Tex Attendance. Please reply with your Employee ID (e.g. MT-00123).' };
    }

    switch (session.state) {
      case 'AWAITING_ID':
        return handleIdEntry_(phone, rawText);
      case 'MENU':
        return handleMenuSelection_(phone, session, textLower);
      case 'AWAITING_LOCATION':
        return handleLocationStep_(phone, session, payload);
      default:
        saveSession_(phone, '', '', '', 'AWAITING_ID', '');
        return { reply: 'Welcome to Maya Tex Attendance. Please reply with your Employee ID (e.g. MT-00123).' };
    }
  } catch (err) {
    // Any unexpected failure (sheet access issue, etc.) — still return
    // something sendable rather than silently dropping the message.
    return { reply: 'Sorry, something went wrong on our end. Please try again in a moment.' };
  } finally {
    lock.releaseLock();
  }
}

function handleIdEntry_(phone, rawText) {
  const empId = rawText.toUpperCase().replace(/\s+/g, '');
  const match = validateEmployeeId_(empId);

  if (!match.valid) {
    return { reply: `Employee ID "${rawText}" not found. Please check and resend your Employee ID (e.g. MT-00123).` };
  }

  saveSession_(phone, match.employeeId, match.name, match.branch, 'MENU', '');
  return { reply: `Hi ${match.name}! ${buildMenuMessage_(match.employeeId)}` };
}

function handleMenuSelection_(phone, session, textLower) {
  const chosenNum = Number(textLower);
  const action = PUNCH_LABELS[chosenNum];

  if (!action) {
    return { reply: `Sorry, I didn't understand that.\n\n${buildMenuMessage_(session.employeeId)}` };
  }

  const validNext = getValidNextActions_(session.employeeId);
  if (!validNext.includes(action)) {
    return { reply: `That's not available right now.\n\n${buildMenuMessage_(session.employeeId)}` };
  }

  if (action === 'Day Off') {
    logPunch_(session.employeeId, session.name, 'Day Off', null);
    saveSession_(phone, session.employeeId, session.name, session.branch, 'MENU', '');
    return {
      reply: `✅ Day Off marked for today.\n\n${buildMenuMessage_(session.employeeId)}`,
      groupNotify: `🏖️ *${session.name}* (${session.employeeId}) — Day Off marked for today.`,
    };
  }

  // Real punch — ask for location before finalizing (soft check, never blocks)
  saveSession_(phone, session.employeeId, session.name, session.branch, 'AWAITING_LOCATION', action);
  return {
    reply: `To record *${action}*, please share your location:\n📎 Attach > Location > Send Your Current Location\n\nOr reply "skip" to continue without it.`,
  };
}

function handleLocationStep_(phone, session, payload) {
  const pendingAction = session.pendingAction;
  if (!pendingAction) {
    saveSession_(phone, session.employeeId, session.name, session.branch, 'MENU', '');
    return { reply: buildMenuMessage_(session.employeeId) };
  }

  const hasLocation = !!(payload.location && payload.location.lat != null && payload.location.lng != null);
  const timeStr = Utilities.formatDate(new Date(), tz_(), 'hh:mm a');

  if (hasLocation) {
    const geofence = getBranchGeofence_(session.branch);
    let locationInfo;
    if (geofence) {
      const dist = distanceMeters_(payload.location.lat, payload.location.lng, geofence.lat, geofence.lng);
      const flag = dist > geofence.radius ? `OUT OF RANGE (${Math.round(dist)}m, limit ${geofence.radius}m)` : '';
      locationInfo = { lat: payload.location.lat, lng: payload.location.lng, distance: dist, flag };
    } else {
      locationInfo = {
        lat: payload.location.lat, lng: payload.location.lng, distance: '',
        flag: `No geofence configured for branch "${session.branch || 'unknown'}"`,
      };
    }
    logPunch_(session.employeeId, session.name, pendingAction, locationInfo);
    saveSession_(phone, session.employeeId, session.name, session.branch, 'MENU', '');
    let msg = `✅ ${pendingAction} recorded at ${timeStr}.`;
    if (locationInfo.flag) msg += `\n⚠️ ${locationInfo.flag}`;
    let groupMsg = `✅ *${session.name}* (${session.employeeId}) — ${pendingAction} at ${timeStr}`;
    if (locationInfo.flag) groupMsg += `\n⚠️ ${locationInfo.flag}`;
    return { reply: `${msg}\n\n${buildMenuMessage_(session.employeeId)}`, groupNotify: groupMsg };
  }

  const textLower = String(payload.text || '').trim().toLowerCase();
  if (textLower === 'skip') {
    logPunch_(session.employeeId, session.name, pendingAction, 'SKIPPED');
    saveSession_(phone, session.employeeId, session.name, session.branch, 'MENU', '');
    const groupMsg = `✅ *${session.name}* (${session.employeeId}) — ${pendingAction} at ${timeStr} (no location)`;
    return {
      reply: `✅ ${pendingAction} recorded at ${timeStr} (no location provided).\n\n${buildMenuMessage_(session.employeeId)}`,
      groupNotify: groupMsg,
    };
  }

  return { reply: `Please share your location (📎 Attach > Location), or reply "skip" to continue without it.` };
}

function buildMenuMessage_(employeeId) {
  const validNext = getValidNextActions_(employeeId);
  const options = validNext.map(action => `${PUNCH_NUMBERS[action]}. ${action}`).join('\n');
  return `What would you like to do?\n${options}\n\nReply with the number.`;
}

/**
 * Valid next punch action(s), based on this employee's most recent row
 * in Attendance_Log — not limited to "today", since shifts can cross
 * midnight. Day Off is offered alongside Check In whenever there's no
 * open shift.
 */
function getValidNextActions_(employeeId) {
  const lastAction = getLastAction_(employeeId);
  switch (lastAction) {
    case 'Check In': return ['Break Out', 'Check Out'];
    case 'Break Out': return ['Break In'];
    case 'Break In': return ['Check Out'];
    case 'Check Out':
    case 'Day Off':
    default: return ['Check In', 'Day Off'];
  }
}

function getLastAction_(employeeId) {
  const sheet = getAttendanceLogSheet_();
  const data = sheet.getDataRange().getValues();
  let latestTs = null;
  let latestAction = null;

  for (let r = 0; r < data.length; r++) {
    const rowEmpId = String(data[r][1]).trim().toUpperCase();
    if (rowEmpId !== employeeId.toUpperCase()) continue;
    const stage = String(data[r][4]).trim();
    if (stage !== 'Check In' && stage !== 'Break Out' && stage !== 'Break In' && stage !== 'Check Out' && stage !== 'Day Off') continue;
    const ts = parseLogTimestamp(data[r][0]); // from Payroll.gs, handles dd/MM/yyyy HH:mm:ss
    if (!ts) continue;
    if (!latestTs || ts > latestTs) {
      latestTs = ts;
      latestAction = stage;
    }
  }
  return latestAction;
}

/**
 * locationInfo: null (not applicable, e.g. Day Off) | 'SKIPPED' (asked,
 * declined) | { lat, lng, distance, flag } (location captured)
 */
function logPunch_(employeeId, name, action, locationInfo) {
  const sheet = getAttendanceLogSheet_();
  const now = new Date();
  const timestamp = Utilities.formatDate(now, tz_(), 'dd/MM/yyyy HH:mm:ss');
  const dateStr = formatDdMmYyyy_(now);

  let lat = '', lng = '', distance = '', flag = '';
  if (locationInfo === 'SKIPPED') {
    flag = 'No location provided';
  } else if (locationInfo) {
    lat = locationInfo.lat;
    lng = locationInfo.lng;
    distance = locationInfo.distance === '' ? '' : Math.round(locationInfo.distance);
    flag = locationInfo.flag || '';
  }

  sheet.appendRow([timestamp, employeeId, name, dateStr, action, 'WA', lat, lng, distance, flag]);
}

function getAttendanceLogSheet_() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Attendance_Log');
  if (!sheet) throw new Error('Attendance_Log tab not found.');
  return sheet;
}

/* ---------------- Geofencing ---------------- */

function getBranchGeofence_(branchName) {
  if (!branchName) return null;
  const sheet = getOrCreateBranchLocationsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(branchName).trim().toLowerCase()) {
      const lat = Number(data[i][1]), lng = Number(data[i][2]), radius = Number(data[i][3]);
      if (!isNaN(lat) && !isNaN(lng) && !isNaN(radius) && radius > 0) return { lat, lng, radius };
    }
  }
  return null;
}

function getOrCreateBranchLocationsSheet_() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(BRANCH_LOCATIONS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(BRANCH_LOCATIONS_TAB);
    sheet.appendRow(['Branch', 'Latitude', 'Longitude', 'Radius (meters)']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Haversine distance between two lat/lng points, in meters.
function distanceMeters_(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = (d) => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

/* ---------------- Employee ID validation against HR Maya (READ-ONLY) ---------------- */

function validateEmployeeId_(empId) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('HR Maya');
  if (!sheet) throw new Error('"HR Maya" tab not found.');

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { valid: false };

  const H = data[0];
  const col = (name) => H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
  const cId = col('Employee ID');
  const cName = col('Name (English)');
  const cBranch = col('Branch');
  if (cId === -1 || cName === -1) throw new Error('HR Maya is missing Employee ID / Name (English) columns.');

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][cId]).trim().toUpperCase();
    if (rowId === empId) {
      return {
        valid: true,
        employeeId: rowId,
        name: String(data[i][cName]).trim(),
        branch: cBranch !== -1 ? String(data[i][cBranch]).trim() : '',
      };
    }
  }
  return { valid: false };
}

/* ---------------- WA_Sessions tab (conversation state only — not attendance data) ---------------- */

function getSession_(phone) {
  const sheet = getOrCreateSessionsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === phone) {
      return {
        employeeId: data[r][1], name: data[r][2], branch: data[r][3],
        state: data[r][4], pendingAction: data[r][5], row: r + 1,
      };
    }
  }
  return { employeeId: '', name: '', branch: '', state: '', pendingAction: '', row: -1 };
}

function saveSession_(phone, employeeId, name, branch, state, pendingAction) {
  const sheet = getOrCreateSessionsSheet_();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === phone) {
      sheet.getRange(r + 1, 1, 1, 7).setValues([[phone, employeeId, name, branch, state, pendingAction, now]]);
      return;
    }
  }
  sheet.appendRow([phone, employeeId, name, branch, state, pendingAction, now]);
}

function getOrCreateSessionsSheet_() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(WA_SESSIONS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(WA_SESSIONS_TAB);
    sheet.appendRow(WA_SESSIONS_HEADERS);
    sheet.getRange(1, 1, 1, WA_SESSIONS_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}