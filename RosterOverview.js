// ============================================================
// RosterOverview.gs — every location's roster, read-only, for HR
// ------------------------------------------------------------
// WHY THIS EXISTS
// Supervisors build and publish rosters on roster.html, one location each,
// behind their own PIN. HR had no page that showed what they had built:
// the only way to see a roster was to open the Roster tab and filter it.
//
// This reads the Roster tab directly — it does not go through
// RosterBackend.gs and changes nothing. For each location it returns:
//   • the grid: every man × every date, with the shift, Day Off, or a gap,
//     and whether that cell is Published or still a Draft
//   • who published it and when, and any amendment with its reason
//   • for dates already past, what ACTUALLY happened (Attendance_Daily), so
//     "rostered to work, was absent" is visible on the same screen
//   • coverage: how many cells are filled, published, or still missing
//
// Head Office and Driver Location are left out — they are auto-present and
// are never rostered (the 4 AM compile treats them the same way).
//
// ROUTER CASE: getRosterOverview   (UI: roster-view-ui.js on index.html)
// ============================================================

function getRosterOverview(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const pad = n => String(n).padStart(2, '0');
    const ddmm = d => pad(d.getDate()) + '-' + pad(d.getMonth() + 1) + '-' + d.getFullYear();
    const parse = s => { const p = String(s || '').split('-'); const d = new Date(+p[2], +p[1] - 1, +p[0]);
                         return isNaN(d.getTime()) ? null : d; };
    const norm = v => {
      if (v == null || v === '') return '';
      if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : ddmm(v);
      const s = String(v).trim();
      let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
      if (m) return pad(m[1]) + '-' + pad(m[2]) + '-' + m[3];
      m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) return pad(m[3]) + '-' + pad(m[2]) + '-' + m[1];
      return s;
    };
    const stamp = v => (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()))
      ? Utilities.formatDate(v, tz_(), 'dd-MM HH:mm') : String(v || '').trim();
    const iso = s => { const p = s.split('-'); return p[2] + p[1] + p[0]; };

    // ── Range: default this week, Sunday → Saturday (Kuwait week) ──
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let from = params.from ? parse(norm(params.from)) : null;
    let to   = params.to   ? parse(norm(params.to))   : null;
    if (!from) { from = new Date(today.getTime() - today.getDay() * 86400000); }
    if (!to)   { to = new Date(from.getTime() + 6 * 86400000); }
    if (to < from) { const t = from; from = to; to = t; }
    if ((to - from) / 86400000 > 41) to = new Date(from.getTime() + 41 * 86400000);   // six weeks max
    const dates = [];
    for (let d = new Date(from); d <= to; d = new Date(d.getTime() + 86400000)) dates.push(ddmm(d));
    const inRange = {}; dates.forEach(d => { inRange[d] = true; });
    const todayStr = ddmm(today);

    const hdr = (sh) => {
      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim().toLowerCase());
      return (...names) => { for (const n of names) { const i = H.indexOf(n.toLowerCase()); if (i !== -1) return i; } return -1; };
    };

    // ── Shifts: which are auto-present (never rostered) ──
    const autoShift = {};
    const shS = ss.getSheetByName('Shifts');
    if (shS && shS.getLastRow() > 1) {
      const at = hdr(shS), cN = at('Shift Name', 'Shift', 'Name'), cA = at('Auto Present');
      shS.getRange(2, 1, shS.getLastRow() - 1, shS.getLastColumn()).getValues().forEach(r => {
        if (cN >= 0 && cA >= 0 && String(r[cA] || '').trim().toLowerCase() === 'yes')
          autoShift[String(r[cN] || '').trim()] = true;
      });
    }

    // ── Staff from HR Maya ──
    const hr = ss.getSheetByName('HR Maya');
    const ha = hdr(hr);
    const hId = ha('Employee ID'), hNm = ha('Name (English)', 'Name'), hLoc = ha('Location', 'Branch'),
          hSh = ha('Shift'), hSt = ha('Status');
    const staff = {};          // id -> { id, name, location, shift }
    hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().forEach(r => {
      const id = String(r[hId] || '').trim(); if (!id) return;
      if (hSt >= 0 && String(r[hSt] || '').trim().toLowerCase() === 'inactive') return;
      const loc = hLoc >= 0 ? String(r[hLoc] || '').trim() : '';
      const shift = hSh >= 0 ? String(r[hSh] || '').trim() : '';
      if (!loc || loc === 'Driver Location' || autoShift[shift]) return;
      staff[id] = { id: id, name: hNm >= 0 ? String(r[hNm] || '').trim() : id, location: loc, shift: shift };
    });

    // ── Supervisors per location (names) ──
    const supsByLoc = {};
    const bl = ss.getSheetByName('Branch_Locations');
    const nameOf = {};
    hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues()
      .forEach(r => { nameOf[String(r[hId] || '').trim().toUpperCase()] = hNm >= 0 ? String(r[hNm] || '').trim() : ''; });
    if (bl && bl.getLastRow() > 1) {
      const ba = hdr(bl), cL = ba('Branch', 'Location'), cS = ba('Supervisors');
      bl.getRange(2, 1, bl.getLastRow() - 1, bl.getLastColumn()).getValues().forEach(r => {
        const loc = String(r[cL] || '').trim(); if (!loc || cS < 0) return;
        supsByLoc[loc] = String(r[cS] || '').split(/[,;\s]+/).filter(Boolean)
          .map(x => nameOf[x.trim().toUpperCase()] || x.trim());
      });
    }

    // ── The roster itself ──
    const locs = {};           // location -> { people:{id:true}, cells:{id:{date:cell}} }
    const L = (loc) => locs[loc] || (locs[loc] = { people: {}, cells: {} });
    const rs = ss.getSheetByName('Roster');
    if (rs && rs.getLastRow() > 1) {
      const ra = hdr(rs);
      const rId = ra('Employee ID'), rNm = ra('Name'), rLoc = ra('Location'), rDate = ra('Date'),
            rPl = ra('Planned Status'), rSh = ra('Shift'), rSt = ra('Status'),
            rPb = ra('Published By'), rPo = ra('Published On'),
            rAb = ra('Last Amended By'), rAo = ra('Last Amended On'), rAr = ra('Amend Reason');
      rs.getRange(2, 1, rs.getLastRow() - 1, rs.getLastColumn()).getValues().forEach(r => {
        const d = norm(r[rDate]); if (!inRange[d]) return;
        const id = String(r[rId] || '').trim(); if (!id) return;
        const loc = String(r[rLoc] || '').trim() || (staff[id] && staff[id].location) || '';
        if (!loc) return;
        const g = L(loc);
        g.people[id] = true;
        if (!staff[id]) staff[id] = { id: id, name: String(r[rNm] || id).trim(), location: loc, shift: '' };
        (g.cells[id] = g.cells[id] || {})[d] = {
          status: String(r[rPl] || '').trim(),                 // Working | Day Off
          shift:  String(r[rSh] || '').trim(),
          state:  String(r[rSt] || '').trim() || 'Draft',        // Published | Draft
          by:     rPb >= 0 ? String(r[rPb] || '').trim() : '',
          on:     rPo >= 0 ? stamp(r[rPo]) : '',
          amendedBy: rAb >= 0 ? String(r[rAb] || '').trim() : '',
          amendedOn: rAo >= 0 ? stamp(r[rAo]) : '',
          reason:    rAr >= 0 ? String(r[rAr] || '').trim() : '',
        };
      });
    }
    // Every location that has rosterable staff appears — including one
    // nobody has rostered at all. That gap is the thing worth seeing.
    Object.keys(staff).forEach(id => { L(staff[id].location).people[id] = true; });

    // ── What actually happened, for dates already past ──
    const actual = {};         // id -> date -> { code, text }
    const ad = ss.getSheetByName('Attendance_Daily');
    if (ad && ad.getLastRow() > 1) {
      const aa = hdr(ad);
      const aId = aa('Employee ID'), aDate = aa('Date'), aSt = aa('Status'), aIn = aa('Check In'), aOut = aa('Check Out');
      const clock = v => (Object.prototype.toString.call(v) === '[object Date]')
        ? Utilities.formatDate(v, tz_(), 'HH:mm') : String(v || '').trim();
      ad.getRange(2, 1, ad.getLastRow() - 1, ad.getLastColumn()).getValues().forEach(r => {
        const d = norm(r[aDate]); if (!inRange[d] || iso(d) >= iso(todayStr)) return;
        const id = String(r[aId] || '').trim(); if (!id) return;
        const st = String(r[aSt] || '').trim();
        const code = st === 'Present' ? 'P' : st === 'Absent' ? 'A' : st === 'Day Off' ? 'O'
                   : st === 'On Leave' ? 'L' : '·';
        const ci = aIn >= 0 ? clock(r[aIn]) : '', co = aOut >= 0 ? clock(r[aOut]) : '';
        (actual[id] = actual[id] || {})[d] = { code: code, text: st + (ci ? ' ' + ci + (co ? '–' + co : '') : '') };
      });
    }

    // ── Assemble per location ──
    const shortShift = (loc, sh) => {
      if (!sh) return '';
      let s = sh;
      if (s.toLowerCase().indexOf(loc.toLowerCase()) === 0) s = s.slice(loc.length).trim();
      s = s.replace(/\bshift\b/i, '').trim();
      return s || sh;
    };
    const out = Object.keys(locs).sort().map(loc => {
      const g = locs[loc];
      const ids = Object.keys(g.people).sort((a, b) => (staff[a].name || '').localeCompare(staff[b].name || ''));
      let filled = 0, published = 0, draft = 0, missing = 0, mismatch = 0;
      let lastBy = '', lastOn = '';
      const rows = ids.map(id => {
        const cells = dates.map(d => {
          const c = (g.cells[id] || {})[d];
          const act = (actual[id] || {})[d] || null;
          if (!c) { missing++; return { d: d, empty: true, act: act }; }
          filled++;
          if (c.state === 'Published') published++; else draft++;
          if (c.state === 'Published' && c.on && (!lastOn || c.on > lastOn)) { lastOn = c.on; lastBy = c.by; }
          // Rostered to work but absent / off — or rostered off but worked.
          const bad = act && ((c.status === 'Working' && (act.code === 'A' || act.code === 'O')) ||
                              (c.status === 'Day Off' && act.code === 'P'));
          if (bad) mismatch++;
          return { d: d, status: c.status, shift: c.shift, label: c.status === 'Day Off' ? 'Off' : shortShift(loc, c.shift),
                   state: c.state, by: c.by, on: c.on, amendedBy: c.amendedBy, amendedOn: c.amendedOn,
                   reason: c.reason, act: act, mismatch: !!bad };
        });
        return { id: id, name: staff[id].name, homeLocation: staff[id].location, cells: cells };
      });
      const total = ids.length * dates.length;
      return {
        location: loc, supervisors: supsByLoc[loc] || [],
        staffCount: ids.length, total: total, filled: filled, published: published,
        draft: draft, missing: missing, mismatch: mismatch,
        lastPublishedBy: lastBy, lastPublishedOn: lastOn, rows: rows,
      };
    });

    return { success: true, from: dates[0], to: dates[dates.length - 1], today: todayStr,
             dates: dates, locations: out };
  } catch (e) { return { success: false, error: e.toString() }; }
}


// ============================================================
// ONE-OFF: remove the repeated rows from Roster_Discrepancies.
// Every recompile used to append the same discrepancy again (1,589 rows,
// 242 distinct). Keeps ONE row per man / date / roster-said / actual —
// the reviewed one if any copy was reviewed. The compile no longer adds
// repeats (AttendanceDailyTrigger adLogDiscrepancyOnce_).
//   previewDedupeRosterDiscrepancies()  — counts only
//   commitDedupeRosterDiscrepancies()   — rewrites the tab
// ============================================================
function previewDedupeRosterDiscrepancies() { return rov_dedupe_(false); }
function commitDedupeRosterDiscrepancies()  { return rov_dedupe_(true); }

function rov_dedupe_(commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sh = ss.getSheetByName('Roster_Discrepancies');
  if (!sh || sh.getLastRow() < 2) { Logger.log('Nothing to do.'); return 'nothing'; }
  const width = sh.getLastColumn();
  const all = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
  const H = all[0].map(h => String(h || '').trim().toLowerCase());
  const cRev = H.indexOf('reviewed');
  const norm = v => (Object.prototype.toString.call(v) === '[object Date]')
    ? Utilities.formatDate(v, tz_(), 'dd-MM-yyyy') : String(v || '').trim();
  const reviewed = r => cRev >= 0 && ['true', 'yes', 'y', '1'].indexOf(String(r[cRev]).trim().toLowerCase()) !== -1;
  const keep = {}, order = [];
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const k = [String(r[0]).trim(), norm(r[3]), String(r[4]).trim(), String(r[5]).trim()].join('|');
    if (!keep[k]) { keep[k] = r; order.push(k); }
    else if (reviewed(r) && !reviewed(keep[k])) keep[k] = r;
  }
  const msg = (all.length - 1) + ' rows → ' + order.length + ' distinct (' + (all.length - 1 - order.length) + ' repeats' +
              (commit ? ' removed' : ' would be removed — run commitDedupeRosterDiscrepancies()') + ').';
  if (commit) {
    const rows = order.map(k => keep[k]);
    sh.getRange(2, 1, all.length - 1, width).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
    if (all.length - 1 > rows.length) sh.deleteRows(rows.length + 2, all.length - 1 - rows.length);
  }
  Logger.log(msg);
  return msg;
}