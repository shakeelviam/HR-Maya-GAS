// ============================================================
// QrPunch.gs — the QR terminal check-in, in the HR Maya stack
// ------------------------------------------------------------
// A rehearsal of the permanent system's check-in, built with
// nothing new: Apps Script, Sheets and GitHub Pages, same as
// everything else here.
//
//   1. A tablet at the location shows a QR that changes every
//      30 seconds.
//   2. The employee, signed in on his phone with the ID and PIN
//      he already uses, scans it.
//   3. The TERMINAL greets him by name and shows the same stage
//      page the kiosk shows after login.
//   4. He taps his stage on the terminal. Punch recorded.
//
// The phone proves WHO. The QR proves WHERE and WHEN. The punch
// happens on a tablet that lives at the branch.
//
// Nothing here duplicates kiosk logic: the employee check is
// loginStaffKiosk, the stage state is getDailyKioskStatus, and
// the punch is markKioskAttendance — the same three functions
// the current kiosk uses, so the two can never disagree.
//
// SETUP — run once from the editor:
//   setTerminalSetupPin('1234')   → the PIN that pairs a tablet
//   listTerminals()               → what is paired, and when it was last seen
// ============================================================

const QR_WINDOW_SEC   = 30;               // the QR changes this often
const QR_WINDOW_GRACE = 2;                // windows either side still accepted (±60s)
const QR_CODE_CHARS   = 8;
const QR_TERMINALS_TAB = 'Terminals';
const QR_TERMINAL_HEADERS = ['Terminal ID', 'Location', 'Token', 'Status', 'Created On', 'Last Seen'];
const QR_SESSION_TTL  = 90;               // seconds a scan waits at the terminal
const QR_SETUP_PIN_PROP = 'TERMINAL_SETUP_PIN';
const QR_PHOTO_FOLDER = 'Attendance Photos';
const QR_PHOTO_KEEP_DAYS = 90;            // photographs of staff are not kept forever

// ── secrets ──────────────────────────────────────────────────
// One secret per location, kept in Script Properties — never in
// the sheet, never in a page. Created the first time it is needed.
function qrLocationSecret_(location) {
  const key = 'QR_SECRET_' + String(location).trim().toUpperCase().replace(/[^A-Z0-9]/g, '_');
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(key);
  if (!secret) {
    secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    props.setProperty(key, secret);
  }
  return secret;
}

function qrWindowFor_(ms) { return Math.floor(ms / 1000 / QR_WINDOW_SEC); }

// The code a terminal shows for one 30-second window.
function qrCodeFor_(location, terminalId, win) {
  const secret = qrLocationSecret_(location);
  const msg = String(terminalId) + ':' + String(win);
  const raw = Utilities.computeHmacSha256Signature(msg, secret);
  let hex = '';
  for (let i = 0; i < raw.length && hex.length < QR_CODE_CHARS; i++) {
    const b = (raw[i] < 0 ? raw[i] + 256 : raw[i]).toString(16);
    hex += (b.length === 1 ? '0' : '') + b;
  }
  return hex.substring(0, QR_CODE_CHARS).toUpperCase();
}

// ── the Terminals tab ────────────────────────────────────────
function qrTerminalSheet_(ss) {
  let sh = ss.getSheetByName(QR_TERMINALS_TAB);
  if (!sh) {
    sh = ss.insertSheet(QR_TERMINALS_TAB);
    sh.getRange(1, 1, 1, QR_TERMINAL_HEADERS.length).setValues([QR_TERMINAL_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function qrFindTerminal_(ss, terminalId) {
  const sh = qrTerminalSheet_(ss);
  if (sh.getLastRow() < 2) return null;
  const d = sh.getDataRange().getValues();
  const H = d[0].map(h => String(h).trim());
  const at = n => H.indexOf(n);
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][at('Terminal ID')]).trim() === String(terminalId).trim()) {
      return {
        row: i + 1,
        id: String(d[i][at('Terminal ID')]).trim(),
        location: String(d[i][at('Location')]).trim(),
        token: String(d[i][at('Token')]).trim(),
        status: String(d[i][at('Status')]).trim()
      };
    }
  }
  return null;
}

// Every terminal call proves itself with the token it was paired with.
function qrAuthTerminal_(ss, params) {
  const t = qrFindTerminal_(ss, params.terminalId);
  if (!t) return { ok: false, error: 'This tablet is not paired. Pair it again from the terminal page.' };
  if (t.status.toLowerCase() !== 'active') return { ok: false, error: 'This tablet has been revoked. Ask HR.' };
  if (t.token !== String(params.token || '').trim()) return { ok: false, error: 'This tablet is not paired correctly. Pair it again.' };
  return { ok: true, terminal: t };
}

// ── pairing ──────────────────────────────────────────────────
// Run once from the editor to set the PIN a tablet must be given.
// Run this from the editor. The Run button passes no argument, so called
// bare it simply invents a PIN and prints it — which is what is wanted 99
// times out of 100. Pass one only if you want to choose it yourself.
function setTerminalSetupPin(pin) {
  let clean = String(pin || '').trim();
  if (!clean) {
    while (clean.length < 6) clean += String(Math.floor(Math.random() * 10));
  } else if (clean.length < 4) {
    Logger.log('Use at least 4 digits.');
    return 'Use at least 4 digits.';
  }
  PropertiesService.getScriptProperties().setProperty(QR_SETUP_PIN_PROP, clean);
  const msg = 'Terminal pairing PIN is ' + clean + ' — type it into the tablet once, on the terminal page.';
  Logger.log(msg);
  return msg;
}

// The editor's Run button cannot pass an argument, so this is the one-click
// version: it invents a PIN, stores it, and prints it in the log. Nothing is
// written into any file.
function newTerminalSetupPin() {
  let pin = '';
  while (pin.length < 6) pin += String(Math.floor(Math.random() * 10));
  PropertiesService.getScriptProperties().setProperty(QR_SETUP_PIN_PROP, pin);
  const msg = 'Terminal pairing PIN is ' + pin + ' — type it into the tablet once, on the terminal page.';
  Logger.log(msg);
  return msg;
}

// What is the PIN right now?
function showTerminalSetupPin() {
  const pin = PropertiesService.getScriptProperties().getProperty(QR_SETUP_PIN_PROP);
  const msg = pin ? ('Terminal pairing PIN is ' + pin) : 'No pairing PIN set yet — run newTerminalSetupPin().';
  Logger.log(msg);
  return msg;
}

// ── a pairing PIN per location ───────────────────────────────
// One shared PIN across every branch is a PIN that leaks. Each location
// gets its own, kept in Script Properties, and a tablet can only be paired
// to the location whose PIN was typed into it.
//
// Run from the editor:
//   generateLocationPins()   → makes any that are missing, prints them all
//   showLocationPins()       → prints them again, changes nothing
//   resetLocationPin('Jabriya') → rolls one, if a PIN gets out
function qrPairPinKey_(location) {
  return 'QR_PAIR_PIN_' + String(location).trim().toUpperCase().replace(/[^A-Z0-9]/g, '_');
}

function qrLocationRows_() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = ss.getSheetByName('Branch_Locations');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const H = d[0].map(h => String(h).trim().toLowerCase());
  const cB = H.indexOf('branch'), cLa = H.indexOf('latitude'), cLo = H.indexOf('longitude'),
        cR = H.indexOf('radius'), cA = H.indexOf('active');
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const name = String(d[i][cB] || '').trim();
    if (!name) continue;
    if (cA !== -1 && String(d[i][cA]).trim().toLowerCase() === 'no') continue;   // retired locations
    out.push({
      name: name,
      lat: cLa === -1 ? '' : d[i][cLa], lng: cLo === -1 ? '' : d[i][cLo],
      radius: cR === -1 ? '' : d[i][cR]
    });
  }
  return out;
}

function qrMakePin_() {
  let pin = '';
  while (pin.length < 6) pin += String(Math.floor(Math.random() * 10));
  return pin;
}

function generateLocationPins() {
  const props = PropertiesService.getScriptProperties();
  const rows = qrLocationRows_();
  if (!rows.length) { Logger.log('No locations found in Branch_Locations.'); return 'No locations found.'; }
  const lines = [], data = [];
  rows.forEach(r => {
    const key = qrPairPinKey_(r.name);
    let pin = props.getProperty(key);
    let made = false;
    if (!pin) { pin = qrMakePin_(); props.setProperty(key, pin); made = true; }
    qrLocationSecret_(r.name);      // every location also needs its code secret
    data.push({ location: r.name, pin: pin, lat: r.lat, lng: r.lng, radius: r.radius, isNew: made });
    lines.push(
      (r.name + '                         ').substring(0, 24) + ' ' + pin +
      (made ? '  (new)' : '       ') +
      (r.lat === '' ? '   no coordinates' : ('   ' + r.lat + ', ' + r.lng + '  r=' + r.radius + 'm'))
    );
  });
  Logger.log('LOCATION                 PIN\n' + lines.join('\n'));
  return { success: true, data: data };
}

function showLocationPins() {
  const props = PropertiesService.getScriptProperties();
  const lines = qrLocationRows_().map(r => {
    const pin = props.getProperty(qrPairPinKey_(r.name));
    return (r.name + '                         ').substring(0, 24) + ' ' + (pin || '— not generated —');
  });
  Logger.log('LOCATION                 PIN\n' + lines.join('\n'));
  return lines.join('\n');
}

function resetLocationPin(location) {
  const name = String(location || '').trim();
  // Run bare from the editor: say what it would do and show the list, rather
  // than failing with a message that looks like a bug.
  if (!name) {
    const list = qrLocationRows_().map(r => '   resetLocationPin("' + r.name + '")').join('\n');
    const msg = 'This rolls one location\'s pairing PIN. Nothing has changed.\n' +
                'Edit the line at the top of the editor to one of these and Run again:\n' + list;
    Logger.log(msg);
    return msg;
  }
  const pin = qrMakePin_();
  PropertiesService.getScriptProperties().setProperty(qrPairPinKey_(name), pin);
  const msg = name + ' pairing PIN is now ' + pin + '. Tablets already paired keep working.';
  Logger.log(msg);
  return msg;
}

// When a tablet says the PIN is wrong, run this from the editor with the
// location and the PIN you typed. It shows exactly what the server holds,
// so there is nothing left to guess at.
//   diagnoseLocationPin('Jabriya', '771139')
function diagnoseLocationPin(location, typed) {
  const name = String(location || '').trim();
  const given = String(typed || '').trim();

  // Run bare from the editor — no arguments — and it prints the whole
  // picture: every location, the PIN the server holds for it, and the
  // master. That is almost always what is wanted.
  if (!name) {
    const props0 = PropertiesService.getScriptProperties();
    const rows = qrLocationRows_();
    const out = ['EVERY LOCATION AND THE PIN THIS SERVER HOLDS', ''];
    rows.forEach(r => {
      const pin = props0.getProperty(qrPairPinKey_(r.name));
      out.push((r.name + '                         ').substring(0, 24) + ' ' +
               (pin || '— none — run generateLocationPins()'));
    });
    out.push('');
    out.push('Master PIN (works for any location): ' + (props0.getProperty(QR_SETUP_PIN_PROP) || '— none —'));
    out.push('');
    out.push('If the tablet refuses a PIN that appears above, the web app is still');
    out.push('serving an older version: Manage deployments -> edit -> New version.');
    Logger.log(out.join('\n'));
    return out.join('\n');
  }
  const props = PropertiesService.getScriptProperties();
  const key = qrPairPinKey_(name);
  const stored = props.getProperty(key);
  const master = props.getProperty(QR_SETUP_PIN_PROP);
  const known = qrLocationRows_().map(r => r.name);
  const lines = [
    'Location typed      : "' + name + '"',
    'Property key        : ' + key,
    'PIN stored for it   : ' + (stored || '— none —'),
    'PIN you typed       : ' + (given || '— none —'),
    'Match               : ' + (stored && given === stored ? 'YES' : 'no'),
    'Master PIN          : ' + (master || '— none —') + (master && given === master ? '  (this one matches)' : ''),
    'Known locations     : ' + known.join(', '),
    'Exact name in sheet : ' + (known.filter(n => n.toLowerCase() === name.toLowerCase())[0] || 'NOT FOUND — the dropdown name does not match Branch_Locations')
  ];
  Logger.log(lines.join('\n'));
  return lines.join('\n');
}

// params: { location, setupPin, label? }
function pairTerminal(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const location = String(params.location || '').trim();
    if (!location) return { success: false, error: 'Choose a location.' };

    // Either this location's own PIN, or the master one. The location PIN is
    // checked first, so a branch PIN cannot pair a tablet to another branch.
    const props = PropertiesService.getScriptProperties();
    const given = String(params.setupPin || '').trim();
    const locPin = props.getProperty(qrPairPinKey_(location));
    const masterPin = props.getProperty(QR_SETUP_PIN_PROP);
    if (!locPin && !masterPin) {
      return { success: false, error: 'No pairing PIN has been set. HR must run generateLocationPins once.' };
    }
    if (given !== locPin && given !== masterPin) {
      return { success: false, error: 'Wrong pairing PIN for ' + location + '.' };
    }

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = qrTerminalSheet_(ss);
    const n = Math.max(0, sh.getLastRow() - 1);
    const terminalId = 'TRM-' + ('0' + (n + 1)).slice(-2);
    const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    const now = (typeof formatDdMmYyyy_ === 'function') ? formatDdMmYyyy_(new Date()) : new Date().toISOString();

    sh.appendRow([terminalId, location, token, 'Active', now, '']);
    qrLocationSecret_(location);   // make sure the location has a secret

    return { success: true, terminalId: terminalId, token: token, location: location };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── the codes a terminal displays ────────────────────────────
// One call gives the tablet the next half hour, so a wifi drop at
// the branch does not stop anybody punching. serverNow lets the
// tablet correct its own clock — the terminal must show the code
// for the window the SERVER thinks it is.
// params: { terminalId, token, count? }
function terminalCodes(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const auth = qrAuthTerminal_(ss, params);
    if (!auth.ok) return { success: false, error: auth.error };

    const count = Math.min(Math.max(Number(params.count) || 60, 1), 240);
    const nowMs = Date.now();
    const win0 = qrWindowFor_(nowMs);
    const out = [];
    for (let i = 0; i < count; i++) {
      const w = win0 + i;
      out.push({ win: w, startMs: w * QR_WINDOW_SEC * 1000, code: qrCodeFor_(auth.terminal.location, auth.terminal.id, w) });
    }
    qrTouchTerminal_(ss, auth.terminal);
    return {
      success: true, terminalId: auth.terminal.id, location: auth.terminal.location,
      windowSec: QR_WINDOW_SEC, serverNow: nowMs, codes: out
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// Last Seen is useful, but not 30 times a minute — once every 5 minutes is plenty.
function qrTouchTerminal_(ss, terminal) {
  try {
    const cache = CacheService.getScriptCache();
    const k = 'qrseen:' + terminal.id;
    if (cache.get(k)) return;
    cache.put(k, '1', 300);
    const sh = qrTerminalSheet_(ss);
    const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const c = H.indexOf('Last Seen');
    if (c === -1) return;
    const stamp = (typeof formatDdMmYyyy_ === 'function')
      ? formatDdMmYyyy_(new Date()) + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm')
      : new Date().toISOString();
    sh.getRange(terminal.row, c + 1).setValue(stamp);
  } catch (e) { /* never fail a poll over a timestamp */ }
}

// ── the employee badge ───────────────────────────────────────
// The phone shows a code that changes every 30 seconds. It computes it
// itself from a secret issued once at sign-in, so it keeps working with
// no signal at all — the only thing that must be online is the tablet.
//
// The tablet reads the badge with its camera. That is the whole flow:
// one gesture, and no call to this script unless somebody is standing
// in front of the terminal.
const BADGE_SECRET_PREFIX = 'BADGE_';

function badgeSecretKey_(empId) {
  return BADGE_SECRET_PREFIX + String(empId).trim().toUpperCase().replace(/[^A-Z0-9]/g, '_');
}

function badgeCodeFor_(secret, empId, win) {
  const raw = Utilities.computeHmacSha256Signature(String(empId) + ':' + String(win), secret);
  let hex = '';
  for (let i = 0; i < raw.length && hex.length < QR_CODE_CHARS; i++) {
    const b = (raw[i] < 0 ? raw[i] + 256 : raw[i]).toString(16);
    hex += (b.length === 1 ? '0' : '') + b;
  }
  return hex.substring(0, QR_CODE_CHARS).toUpperCase();
}

// Called once, when a man signs in on his phone. Hands back the secret his
// phone will keep, and the server's clock so the phone can correct its own.
// params: { empId, pin }
function issueBadge(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const who = loginStaffKiosk({ empId: params.empId, pin: params.pin });
    if (!who.success) return { success: false, error: who.error || 'Invalid Employee ID or PIN' };

    const props = PropertiesService.getScriptProperties();
    const key = badgeSecretKey_(who.id);
    let secret = props.getProperty(key);
    if (!secret) {
      secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
      props.setProperty(key, secret);
    }
    return {
      success: true, empId: who.id, name: who.name, secret: secret,
      windowSec: QR_WINDOW_SEC, serverNow: Date.now()
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── personal links ───────────────────────────────────────────
// A phone can lose its storage — a browser cleanup, private mode, a new
// handset. Rather than make a man type his ID and PIN again, every
// employee gets a personal link. Opening it registers his badge silently.
// Sent once on WhatsApp, added to his home screen, and that is the end of
// it. The link can be cancelled like anything else.
function badgeLinkSecret_() {
  const props = PropertiesService.getScriptProperties();
  let sec = props.getProperty('BADGE_LINK_SECRET');
  if (!sec) { sec = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); props.setProperty('BADGE_LINK_SECRET', sec); }
  return sec;
}

function badgeLinkKey_(empId) {
  const raw = Utilities.computeHmacSha256Signature(String(empId).trim().toUpperCase(), badgeLinkSecret_());
  let hex = '';
  for (let i = 0; i < raw.length && hex.length < 16; i++) {
    const b = (raw[i] < 0 ? raw[i] + 256 : raw[i]).toString(16);
    hex += (b.length === 1 ? '0' : '') + b;
  }
  return String(empId).trim().toUpperCase() + '.' + hex.substring(0, 16);
}

// Run bare from the editor: prints a link for every active employee.
// BADGE_PAGE_URL is read from Script Properties so no address is written
// into this file; set it once with setBadgePageUrl('https://…/badge.html').
function badgeLinks() {
  const props = PropertiesService.getScriptProperties();
  const base = props.getProperty('BADGE_PAGE_URL');
  if (!base) {
    const msg = 'Set the page address once: setBadgePageUrl("https://yoursite/badge.html")';
    Logger.log(msg); return msg;
  }
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = ss.getSheetByName('HR Maya');
  const d = sh.getDataRange().getValues();
  const H = d[0].map(h => String(h).trim().toLowerCase());
  const cId = H.indexOf('employee id'), cName = H.indexOf('name (english)'),
        cStat = H.indexOf('status'), cMob = H.indexOf('mobile');
  const lines = [];
  for (let i = 1; i < d.length; i++) {
    const id = String(d[i][cId] || '').trim();
    if (!id) continue;
    if (cStat !== -1 && String(d[i][cStat]).trim().toLowerCase() !== 'active') continue;
    const name = cName !== -1 ? String(d[i][cName]).trim() : id;
    const mob = cMob !== -1 ? String(d[i][cMob] || '').trim() : '';
    lines.push([id, name, mob, base + '?k=' + encodeURIComponent(badgeLinkKey_(id))].join('\t'));
  }
  const out = 'EMPLOYEE ID\tNAME\tMOBILE\tPERSONAL BADGE LINK\n' + lines.join('\n');
  Logger.log(out);
  return out;
}

function setBadgePageUrl(url) {
  const u = String(url || '').trim();
  if (!u) { const m = 'Pass the address, e.g. setBadgePageUrl("https://yoursite/badge.html")'; Logger.log(m); return m; }
  PropertiesService.getScriptProperties().setProperty('BADGE_PAGE_URL', u);
  const m = 'Badge page address saved: ' + u;
  Logger.log(m); return m;
}

// The phone opened a personal link. Same answer as signing in, no typing.
// params: { k }
function badgeByLink(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const k = String(params.k || '').trim();
    const empId = k.split('.')[0].toUpperCase();
    if (!empId || badgeLinkKey_(empId) !== k) return { success: false, error: 'That link is not valid. Ask HR for a new one.' };

    const props = PropertiesService.getScriptProperties();
    const key = badgeSecretKey_(empId);
    let secret = props.getProperty(key);
    if (!secret) {
      secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
      props.setProperty(key, secret);
    }
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    return {
      success: true, empId: empId, name: qrEmployeeName_(ss, empId), secret: secret,
      windowSec: QR_WINDOW_SEC, serverNow: Date.now()
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// A man's badge can be cancelled — lost phone, leaver. He simply signs in
// again on a new phone to get a fresh one.
// params: { empId }
function revokeBadge(params) {
  params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
  const empId = String(params && params.empId || '').trim();
  if (!empId) {
    const msg = 'Name the employee, e.g. revokeBadge({empId:"MT-00060"}). Nothing has changed.';
    Logger.log(msg); return { success: false, error: msg };
  }
  PropertiesService.getScriptProperties().deleteProperty(badgeSecretKey_(empId));
  const msg = 'Badge for ' + empId + ' cancelled. His phone will stop working within 30 seconds.';
  Logger.log(msg);
  return { success: true, message: msg };
}

// The tablet has read a badge. One call: who is he, and what stage is he at.
// params: { terminalId, token, payload }   payload = "MAYA2|MT-00060|<win>|<code>"
function badgeIdentify(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const auth = qrAuthTerminal_(ss, params);
    if (!auth.ok) return { success: false, error: auth.error };

    const parts = String(params.payload || '').trim().split('|');
    if (parts.length !== 4 || parts[0] !== 'MAYA2') {
      return { success: false, error: 'That is not a Maya badge.' };
    }
    const empId = String(parts[1]).trim().toUpperCase();
    const win = Number(parts[2]);
    const code = String(parts[3]).toUpperCase();

    const secret = PropertiesService.getScriptProperties().getProperty(badgeSecretKey_(empId));
    if (!secret) return { success: false, error: 'This phone is not registered. Sign in on the staff page first.' };

    const nowWin = qrWindowFor_(Date.now());
    if (!isFinite(win) || Math.abs(nowWin - win) > QR_WINDOW_GRACE) {
      return { success: false, error: 'That badge is out of date. Open the app again.' };
    }
    if (badgeCodeFor_(secret, empId, win) !== code) {
      return { success: false, error: 'That badge is not valid.' };
    }

    // One badge code, one punch.
    const cache = CacheService.getScriptCache();
    const usedKey = 'badgeused:' + empId + ':' + win;
    if (cache.get(usedKey)) return { success: false, error: 'Already punched on this code. Wait a moment and try again.' };

    const status = getDailyKioskStatus({ empId: empId });
    const name = qrEmployeeName_(ss, empId);
    const punchToken = Utilities.getUuid();
    cache.put('badgepunch:' + auth.terminal.id, JSON.stringify({
      token: punchToken, empId: empId, name: name, win: win, at: Date.now()
    }), QR_SESSION_TTL);

    return {
      success: true, empId: empId, name: name,
      stage: status.currentStage || 'none',
      greeting: qrGreeting_(), punchToken: punchToken
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function qrEmployeeName_(ss, empId) {
  const sh = ss.getSheetByName('HR Maya');
  if (!sh) return empId;
  const d = sh.getDataRange().getValues();
  const H = d[0].map(h => String(h).trim().toLowerCase());
  const cId = H.indexOf('employee id'), cName = H.indexOf('name (english)');
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][cId]).trim().toUpperCase() === String(empId).trim().toUpperCase()) {
      return cName !== -1 ? String(d[i][cName]).trim() : empId;
    }
  }
  return empId;
}

// ── the scan, from the employee's phone ──────────────────────
// params: { empId, pin, payload }  payload = "MAYA|TRM-01|<win>|<code>"
function qrScan(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const parts = String(params.payload || '').trim().split('|');
    if (parts.length !== 4 || parts[0] !== 'MAYA') {
      return { success: false, error: 'That is not a Maya terminal code. Scan the code on the tablet.' };
    }
    const terminalId = parts[1], win = Number(parts[2]), code = String(parts[3]).toUpperCase();

    // Who is he? Same check the kiosk does, so one set of credentials.
    const who = loginStaffKiosk({ empId: params.empId, pin: params.pin });
    if (!who.success) return { success: false, error: who.error || 'Invalid Employee ID or PIN' };

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const t = qrFindTerminal_(ss, terminalId);
    if (!t) return { success: false, error: 'Unknown terminal.' };
    if (t.status.toLowerCase() !== 'active') return { success: false, error: 'That terminal has been revoked. Ask HR.' };

    // Is the code current? Windows either side are accepted, because a
    // camera takes a moment and clocks are never exactly equal.
    const nowWin = qrWindowFor_(Date.now());
    if (!isFinite(win) || Math.abs(nowWin - win) > QR_WINDOW_GRACE) {
      return { success: false, error: 'That code has expired. Look at the tablet and scan again.' };
    }
    if (qrCodeFor_(t.location, t.id, win) !== code) {
      return { success: false, error: 'That code is not valid for this terminal.' };
    }

    // One code, one man, one punch.
    const cache = CacheService.getScriptCache();
    const usedKey = 'qrused:' + t.id + ':' + win;
    if (cache.get(usedKey)) {
      return { success: false, error: 'That code has already been used. Wait for the next one on the tablet.' };
    }
    cache.put(usedKey, String(who.id), (QR_WINDOW_SEC * (QR_WINDOW_GRACE + 2)));

    // Hand the man over to the terminal.
    const sessionId = Utilities.getUuid();
    const session = {
      sessionId: sessionId, empId: who.id, name: who.name,
      stage: who.currentStatus || 'none', at: Date.now()
    };
    cache.put('qrscan:' + t.id, JSON.stringify(session), QR_SESSION_TTL);

    return {
      success: true, name: who.name, empId: who.id, location: t.location,
      message: 'Go to the terminal — ' + who.name.split(' ')[0] + ', your stage is ready.'
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── the terminal asking "has anyone scanned?" ────────────────
// Reads cache only — no sheet is touched, so this stays cheap
// even at one call every two seconds.
// params: { terminalId, token }
function terminalPoll(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const cache = CacheService.getScriptCache();
    const raw = cache.get('qrscan:' + String(params.terminalId || '').trim());
    if (!raw) return { success: true, waiting: false };

    // Only check the token when there is actually something to hand over.
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const auth = qrAuthTerminal_(ss, params);
    if (!auth.ok) return { success: false, error: auth.error };

    const s = JSON.parse(raw);
    return {
      success: true, waiting: true, sessionId: s.sessionId,
      empId: s.empId, name: s.name, stage: s.stage, greeting: qrGreeting_()
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function qrGreeting_() {
  // Never let a date-format quirk decide the greeting: if the pattern comes
  // back as anything but a number, fall back to the plain clock.
  let h = NaN;
  try { h = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH')); } catch (e) { h = NaN; }
  if (!isFinite(h)) h = new Date().getHours();
  if (h < 12) return 'Good morning';
  if (h < 17) return 'Good afternoon';
  return 'Good evening';
}

// ── the punch, tapped on the terminal ────────────────────────
// params: { terminalId, token, sessionId, stage }
function terminalPunch(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const auth = qrAuthTerminal_(ss, params);
    if (!auth.ok) return { success: false, error: auth.error };

    const cache = CacheService.getScriptCache();
    const key = 'badgepunch:' + auth.terminal.id;
    const raw = cache.get(key);
    if (!raw) return { success: false, error: 'That took too long. Show your badge again.' };
    const s = JSON.parse(raw);
    const given = String(params.punchToken || params.sessionId || '').trim();
    if (s.token !== given) {
      return { success: false, error: 'Somebody else has used the terminal since. Show your badge again.' };
    }

    const stage = String(params.stage || '').trim();
    if (['Check In', 'Break Out', 'Break In', 'Check Out'].indexOf(stage) === -1) {
      return { success: false, error: 'Invalid stage' };
    }

    // The same punch path the kiosk uses. The location is the terminal's,
    // not the phone's GPS — the tablet is what is standing at the branch.
    const res = markKioskAttendance({
      empId: s.empId, stage: stage,
      auth: 'QR', branchOverride: auth.terminal.location
    });
    cache.remove(key);
    cache.put('badgeused:' + s.empId + ':' + s.win, '1', QR_WINDOW_SEC * (QR_WINDOW_GRACE + 2));
    if (!res || !res.success) return res || { success: false, error: 'Punch failed' };

    // The tablet's front camera saw whoever was standing there. Filed after
    // the punch, never before: a camera fault must not cost a man his punch.
    let photoNote = '';
    try {
      const url = qrSavePhoto_(params.photo, s, stage, auth.terminal.location);
      qrWritePhotoLink_(auth.terminal.location, s.empId, stage, url);
      if (!url) photoNote = 'NO PHOTO';
    } catch (e) { photoNote = 'NO PHOTO'; }

    return { success: true, name: s.name, empId: s.empId, stage: stage,
             photo: photoNote ? 'none' : 'saved',
             message: res.message || (stage + ' recorded') };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── the photograph ───────────────────────────────────────────
// One frame, filed in Drive by location and date. Only the link is
// kept in the sheet, so the workbook stays small.
function qrSavePhoto_(photo, session, stage, location) {
  if (!photo || !photo.data) return '';
  const root = qrFolder_(DriveApp.getRootFolder(), QR_PHOTO_FOLDER);
  const byLoc = qrFolder_(root, String(location || 'Unknown'));
  const dayStr = (typeof formatDdMmYyyy_ === 'function') ? formatDdMmYyyy_(new Date()) : new Date().toISOString().slice(0, 10);
  const byDay = qrFolder_(byLoc, dayStr);
  const hhmm = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmm');
  const name = session.empId + ' ' + stage + ' ' + hhmm + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(photo.data), photo.mimeType || 'image/jpeg', name);
  return byDay.createFile(blob).getUrl();
}

function qrFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// Put the link on the punch row that markKioskAttendance just wrote —
// the last row for this employee and stage.
function qrWritePhotoLink_(location, empId, stage, url) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = ss.getSheetByName('Attendance_Log');
  if (!sh || sh.getLastRow() < 2) return;
  let H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  let cPhoto = H.indexOf('Photo');
  if (cPhoto === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('Photo').setFontWeight('bold');
    cPhoto = sh.getLastColumn() - 1;
  }
  const last = sh.getLastRow();
  const row = sh.getRange(last, 1, 1, sh.getLastColumn()).getValues()[0];
  const cId = H.indexOf('Employee ID'), cStage = H.indexOf('Stage');
  if (String(row[cId]).trim() !== String(empId).trim() || String(row[cStage]).trim() !== stage) return;
  sh.getRange(last, cPhoto + 1).setValue(url || 'NO PHOTO');
}

// Monthly tidy-up. Install once from the editor: setPhotoPurgeTrigger()
function purgeOldAttendancePhotos() {
  const cutoff = new Date(Date.now() - QR_PHOTO_KEEP_DAYS * 86400000);
  const roots = DriveApp.getFoldersByName(QR_PHOTO_FOLDER);
  if (!roots.hasNext()) { Logger.log('No photo folder yet.'); return 'No photo folder yet.'; }
  const root = roots.next();
  let trashed = 0;
  const locs = root.getFolders();
  while (locs.hasNext()) {
    const days = locs.next().getFolders();
    while (days.hasNext()) {
      const day = days.next();
      if (day.getDateCreated() < cutoff) { day.setTrashed(true); trashed++; }
    }
  }
  const msg = 'Trashed ' + trashed + ' day folder(s) older than ' + QR_PHOTO_KEEP_DAYS + ' days.';
  Logger.log(msg);
  return msg;
}

function setPhotoPurgeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'purgeOldAttendancePhotos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('purgeOldAttendancePhotos').timeBased().onMonthDay(1).atHour(3).create();
  const msg = 'Attendance photos will be purged after ' + QR_PHOTO_KEEP_DAYS + ' days, checked monthly.';
  Logger.log(msg);
  return msg;
}

// ── HR view ──────────────────────────────────────────────────
function listTerminals(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sh = qrTerminalSheet_(ss);
    if (sh.getLastRow() < 2) return { success: true, data: [] };
    const d = sh.getDataRange().getValues();
    const H = d[0].map(h => String(h).trim());
    const out = [];
    for (let i = 1; i < d.length; i++) {
      const rec = {};
      H.forEach((h, c) => { if (h !== 'Token') rec[h] = String(d[i][c] || '').trim(); });
      out.push(rec);
    }
    return { success: true, data: out };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// params: { terminalId }
function revokeTerminal(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const t = qrFindTerminal_(ss, params.terminalId);
    if (!t) return { success: false, error: 'No such terminal.' };
    const sh = qrTerminalSheet_(ss);
    const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    sh.getRange(t.row, H.indexOf('Status') + 1).setValue('Revoked');
    sh.getRange(t.row, H.indexOf('Token') + 1).setValue('');
    return { success: true, message: t.id + ' revoked. The tablet will stop working within seconds.' };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// Editor helper: what is paired and when was it last seen.
function checkTerminals() {
  const res = listTerminals({});
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}