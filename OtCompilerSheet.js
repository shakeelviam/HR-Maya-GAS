// ============================================================
// OtCompilerSheet.gs — the OT compiler's own sheet goes into OT_Entries
// ------------------------------------------------------------
// Source: "OT maya - شهر 9.xlsx" (Google Drive), compiled by the OT
// compiler (3abwd.3@gmail.com). One row per man, "day 1" … "day 31".
//
// WHICH DATE IS "day 1"? The sheet does not say. Matched against the
// punches (presence − 9) for every man on every day, the sheet lines up
// when day 1 = 20-08-2026: 353 of 775 cells agree within an hour, against
// at most 128 for any other start date. So:
//     day 1 = 20-08-2026   …   day 31 = 19-09-2026
// Dated by the day the shift STARTED, the same as Attendance_Daily.
// He closed his last sheet at 19/08 and this one runs to 20/09: the last
// column is the shift of 19-09, which ends on the morning of 20-09. The
// shift that starts on 20-09 is not in it — it stays as the payroll close
// decided it (from the claims).
//
// DAY OFF WORKED = the whole shift is OT (Min Bahadur, Laxman, Kiran …).
// His figure for those days is imported as he wrote it.
//
// RULE (Shakeel, 21-09-2026): the compiler's entry wins. For 20-08 → 19-09:
//   • every OT_Entries row dated in that range is REPLACED by the sheet
//     (a full copy of OT_Entries is kept first as a backup tab);
//   • every claim still Pending in that range is closed "Settled
//     (compiler sheet)" — it is not paid on top of the sheet.
// OT_Entries rows outside the range (20-09, and anything before 20-08)
// are not touched.
//
// Names are the compiler's nicknames. Each was matched to an employee by
// comparing his numbers with that man's punches; the uncertain ones are
// marked below. A name with no ID is NOT imported — it is listed.
//
// previewCompilerSheetImport() writes nothing. commitCompilerSheetImport()
// does it. Run the preview first and read the back-check.
// ============================================================
const OTS_FROM = '20-08-2026';        // day 1
const OTS_TO   = '19-09-2026';        // day 31
const OTS_SOURCE = 'OT maya - شهر 9';
const OTS_BY = 'OT compiler sheet (3abwd.3@gmail.com)';
const OTS_SETTLED = 'Settled (compiler sheet)';

// [name on his sheet, Employee ID, day 1 … day 31]
// Matched by punches. Checked by hand:
//   sanju → MT-00018 Foudman Limbu (11 of 20 days match his punches)
//   Tapa p → MT-00042 Puran Kumar Thapa     dev vai → MT-00013 Deb Bahadur Thapa
//   ram → MT-00044 Ram Bahadur Sirmal       ram bahardur → MT-00045 Ram Bahadur Thing
//   suman → MT-00052 Suman Bishwakarma      Suman gang → MT-00054 Suman Tamang
//   jett → MT-00024 Jit Bahadur Thakali
//   andrew refrea → MT-00020 James Andrew (Jeffrson is on vacation — Shakeel, 21-09-2026)
//   prashmich → MT-00012 Pramesh Pyakurel   akil/toheed/shakil/sayed/Mehedi → drivers
//   milan tmg → MT-00039 Nilam Tamang (Shakeel, 22-09-2026): Nilam changed his number and
//   claimed on the group as "Milan"; his entries match Nilam's punches on 06, 08, 10, 13, 17-09.
const OTS_ROWS = [
  ['Tapa p', 'MT-00042', ',6,,,6,,,6,,6,6,,7,6,2,2,,6,6,,,6,,1,6,,2,,2,2,2'],
  ['sunil', 'MT-00055', ',,,1,,,,2,,2,2,,1,3,2,,1,,,,,,,,,,,,,,'],
  ['toheed', 'MT-00005', ',,,,,,,,,,,,,,,,,,,,,,8,,,,,8,,,'],
  ['ram', 'MT-00044', ',,,6,,,6,,,,6,,7,6,,,,6,,7,6,,,,6,,7,,6,,'],
  ['shakil', 'MT-00049', '4,4,,,,,8,2,,2,2,2,3,2,10,3,4,,,,8,4,4,,,,,8,,,'],
  ['sayed', 'MT-00064', '2,4,2,8,2,,,,,6,11,5,5,2,2,2,,10,,,,,,8,,,,,,,8'],
  ['sanju', 'MT-00018', '2,6,,2,6,1,2,6,,,6,,7,,6,,6,2,6,1,,6,,6,,6,2,,2,,6'],
  ['Dependra', 'MT-00017', ',6,,,6,1,,,6,,,6,1,6,,6,,,6,1,6,,6,,,6,1,6,,6,'],
  ['yogendra', 'MT-00059', ',,,6,2,1,,,,,,6,2,6,6,2,,6,,2,6,,2,,2,,2,,2,,'],
  ['aziz', 'MT-00007', ',1,,,,,,,,2,4,1,,3,2,,1,,,,,,,,,,,,,,'],
  ['ajay m', 'MT-00003', ',1,,,,,,,,2,4,2,1,,2,,1,,,,,,,,,,,,,,'],
  ['lakpa', 'MT-00031', ',,2,6,,7,2,,7,,6,,6,2,6,1,6,2,,6,,6,7,,2,,6,2,,7,'],
  ['nilam', 'MT-00039', '6,,,6,,,,6,,6,,,7,,,,,,,,,,,,,,,,,,'],
  ['manish', 'MT-00033', '6,,6,,,7,6,1,6,6,,,,,,,,,,,,,,,,,,,6,6,6'],
  ['suman', 'MT-00052', '2,,,,,5,,2,3,3,2,,2,,,,,,6,2,,2,2,2,5,2,7,2,2,2,2'],
  ['suraj', 'MT-00056', ',,,,,,,,,,,,,,,,,,,,,4,,,,,,,,,'],
  ['sudeep', 'MT-00051', ',,,2,6,1,2,6,,,,6,1,2,,6,,6,6,,6,2,6,2,6,2,,6,2,,6'],
  ['bikash', 'MT-00010', '2,,6,2,,7,,2,6,,6,2,2,6,2,,6,2,,7,,2,,,,6,2,6,2,,6'],
  ['kasheeb', 'MT-00028', '6,,6,,,2,6,2,,6,,,2,6,,6,6,,,7,,6,2,6,2,,7,,6,,2'],
  ['bibek', 'MT-00009', '6,2,6,,,2,6,6,,6,6,,7,6,,6,6,,,7,2,6,,7,,,1,6,,,2'],
  ['devram - bishal', 'MT-00014', '6,7,,2,6,,,6,1,1,,6,,6,2,7,,,6,,6,1,2,6,,6,6,6,2,1,6'],
  ['uttam', 'MT-00058', '2,7,,6,2,7,1,2,7,,2,6,,6,6,2,,6,,6,2,6,2,,6,,2,,6,2,6'],
  ['nisan', 'MT-00040', ',,6,,,1,,,,6,,6,,,6,,6,,6,,,,6,,6,6,,,,6,'],
  ['dev vai', 'MT-00013', '6,,6,,2,7,,2,,,,6,2,,6,,2,,6,2,6,,6,2,,6,,6,,6,2'],
  ['Laxman luky', 'MT-00032', ',,,,,,,,6,5,,3,9,5,5,13,,,,,,,,,,,,,,,'],
  ['gopal', 'MT-00019', ',1,,,,,,,,2,4,10,1,3,2,,1,,,,,,,,,,,,,,'],
  ['kapil', 'MT-00027', '4,4,,,,,,,4,4,,,,,4,4,4,,,,,4,4,,,,,,4,4,'],
  ['pujan', 'MT-00041', '4,4,,,,,,4,4,,,,,,4,4,,,,,,,4,,,,,,4,4,'],
  ['aaitu', 'MT-00001', ',6,,,,2,6,,,6,6,,7,,2,2,,6,6,1,2,,6,,6,,7,2,,6,2'],
  ['ajay g', 'MT-00002', ',6,2,,2,,,6,,6,6,2,,6,6,,6,2,2,,6,,2,1,,2,,2,,,6'],
  ['dhirendra', 'MT-00015', ',,,,,7,,,,6,,6,1,6,,,6,,,7,,6,,6,,,7,6,,,6'],
  ['milan tam', 'MT-00035', ',,,,6,1,6,,,,,,7,,6,,,6,,1,6,,,6,,,1,6,,6,6'],
  ['bimal rai', 'MT-00011', ',7,,2,,5,6,2,1,6,,6,,,6,1,6,2,6,2,,6,1,6,2,,2,,6,1,6'],
  ['jaya', 'MT-00022', ',1,,,,,,,,2,4,2,,3,2,,1,,,,,,,,,,,,,,'],
  ['Suman gang', 'MT-00054', ',,,,,,,,,,4,2,1,3,2,,,,,,,,,,,,,,,,'],
  ['Jash', 'MT-00021', '6,6,,,,1,,6,6,,,,1,,6,6,,,,1,,6,6,,,,1,,6,6,'],
  ['Kristal', 'MT-00030', '4,,4,,,,,4,4,,1,,,,6,6,2,,6,1,6,2,6,,2,6,1,6,,6,6'],
  ['akil', 'MT-00004', '2,1,9,1,2,1,1,1,1,9,1,1,3,3,1,1,9,,,,,,,8,,,,,,,8'],
  ['sanamsing', 'MT-00046', ',4,,,,,,4,4,,,,,,4,4,,,,,,4,4,,6,6,,6,6,2,2'],
  ['suman karki', 'MT-00053', ',2,6,,,2,6,,6,2,6,2,1,6,,6,6,,,7,2,6,,6,2,,7,,6,,6'],
  ['prashmich', 'MT-00012', ',7,2,,6,2,6,,,6,2,,6,1,,7,6,,6,2,,,7,6,,2,,2,,7,2'],
  ['sandip', 'MT-00047', ',2,,6,2,2,3,,2,2,,2,2,2,,2,2,,2,,2,2,1,2,,2,2,,2,1,2'],
  ['ram bahardur', 'MT-00045', '4,4,,,,,,4,,,,,,,4,4,,,,,,4,4,,,,,,4,4,'],
  ['jett', 'MT-00024', ',6,2,,6,1,2,,6,2,,6,7,2,,,6,2,,7,,,6,1,6,,2,,,6,2'],
  ['mukti', 'MT-00038', '2,,2,2,6,7,2,2,,,2,,,,,,,,,,,,,7,,6,7,6,6,6,'],
  ['kabir bhattari', 'MT-00026', ',,,,6,,6,,,,6,,,6,,,6,,6,,6,,,6,,6,,6,,,6'],
  ['arjun pariyar', 'MT-00006', ',6,2,,2,,,,,,,,,,,,,,,,,,,,,,,,,,'],
  ['Mehedi Hasan', 'MT-00034', '9,,3,4,,2,2,10,4,,,,,10,4,4,3,,,,4,8,,,,,,,8,,'],
  ['Junaid Shaikh', 'MT-00025', ',1,,,,,,,2,4,,2,1,3,,,1,,,,,,,,,,,,,,'],
  ['Rajan Shrestha', 'MT-00043', '6,,6,2,,7,2,,6,,,,,,,,,,,,,,,,,,,,,2,6'],
  ['Saroj KC', 'MT-00048', '6,1,6,,6,,6,,2,,6,,2,6,,7,,6,2,,6,2,7,,2,,6,2,,2,2'],
  ['Subas Pun', 'MT-00050', '6,2,,6,2,1,,6,2,6,,6,1,2,6,2,,6,2,2,6,,2,,6,6,1,2,,2,6'],
  ['Dipak LO', 'MT-00016', '6,,6,6,,7,,,1,6,,2,6,1,2,,6,6,,6,6,1,,6,,2,,,2,,6'],
  ['andrew refrea', 'MT-00020', ',,6,,,7,,,,,,6,7,,,,6,,,7,,6,,2,,,1,2,,,2'],
  ['prakash', 'MT-00061', ',2,,6,,2,,6,2,,6,6,1,,6,6,,6,6,,6,,,7,,6,1,,6,,'],
  ['min bahadur', 'MT-00062', ',,,,,,4,,,5,4,3,17,5,5,1,,,2,,,2,1,,,,,,,,'],
  ['kiran bhujel', 'MT-00063', ',,,,,,4,,,5,4,3,9,14,2,,,,,1,1,1,1,,,,,,,,'],
  ['mohamed rizwan', 'MT-00036', ',,,,,,,,,2,,,,,,,,,,,,,,,,,,,,,'],
  ['milan tmg', 'MT-00039', ',,,,,,,,,,,,,,,,6,6,,7,,6,,,6,7,,,6,,'],
];

// HR CORRECTIONS on top of his sheet (Shakeel, 22-09-2026): the totals are
// Nilam Tamang 76h and Milan Tamang 60h. His sheet gives 75h and 64h.
//   Nilam  +1h on 01-09 (7h → 8h)
//   Milan  −4h on 18-09 (6h → 2h) — the one Milan day his punches do not
//          support (17:28 → 02:29, no overtime)
// [Employee ID, date, change in hours, why]
const OTS_ADJUST = [
  ['MT-00039', '01-09-2026', +1, 'HR correction: Nilam Tamang total 76h'],
  ['MT-00035', '18-09-2026', -4, 'HR correction: Milan Tamang total 60h'],
];

function previewCompilerSheetImport() { return ots_import_(false); }
function commitCompilerSheetImport()  { return ots_import_(true); }

function ots_import_(commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const lock = LockService.getScriptLock();
  if (commit && !lock.tryLock(30000)) { Logger.log('Busy — try again in a minute.'); return 'busy'; }
  try {
    const out = [];
    const iso = d => { const p = String(d).split('-'); return p[2] + p[1] + p[0]; };
    const inRange = d => d && iso(d) >= iso(OTS_FROM) && iso(d) <= iso(OTS_TO);
    const pad = (x, n) => { x = String(x); return x.length >= n ? x.slice(0, n - 1) + ' ' : x + ' '.repeat(n - x.length); };
    const hm = h => otc_hm_(h);
    const p0 = OTS_FROM.split('-');
    const day1 = new Date(+p0[2], +p0[1] - 1, +p0[0]);
    const dateOf = i => formatDdMmYyyy_(new Date(day1.getFullYear(), day1.getMonth(), day1.getDate() + i));

    // names from HR Maya
    const hr = ss.getSheetByName('HR Maya');
    const H = hr.getRange(1, 1, 1, hr.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const cId = H.indexOf('Employee ID');
    let cNm = H.indexOf('Name (English)'); if (cNm === -1) cNm = H.indexOf('Name');
    const names = {};
    hr.getRange(2, 1, hr.getLastRow() - 1, hr.getLastColumn()).getValues().forEach(r => {
      const id = String(r[cId] || '').trim().toUpperCase(); if (id) names[id] = String(r[cNm] || '').trim();
    });

    // ── his sheet → rows ─────────────────────────────────────────
    const add = [], unmapped = [], perMan = {};
    let total = 0;
    OTS_ROWS.forEach(x => {
      const nick = x[0], id = String(x[1] || '').toUpperCase();
      const vals = String(x[2]).split(',');
      let sum = 0;
      vals.forEach((v, i) => {
        const h = parseFloat(v);
        if (!h || h <= 0 || i > 30) return;
        sum += h;
        if (id) add.push({ id: id, name: names[id] || nick, nick: nick, date: dateOf(i), day: i + 1, hours: h });
      });
      if (!id && sum) unmapped.push(nick + ': ' + hm(sum) + ' on ' + vals.filter(v => parseFloat(v) > 0).length + ' day(s)');
      if (id) { total += sum; perMan[id] = (perMan[id] || 0) + sum; }
    });

    // ── HR corrections ───────────────────────────────────────────
    const adjusted = [];
    OTS_ADJUST.forEach(a => {
      const id = String(a[0]).toUpperCase(), d = otc_norm_(a[1]), delta = Number(a[2]);
      const e = add.find(x => x.id === id && x.date === d);
      if (e) {
        const was = e.hours;
        e.hours = Math.max(0, Math.round((e.hours + delta) * 100) / 100);
        e.note = a[3] + ' (' + was + 'h → ' + e.hours + 'h)';
        adjusted.push(d + '  ' + id + ' ' + (names[id] || '') + ': ' + was + 'h → ' + e.hours + 'h — ' + a[3]);
      } else if (delta > 0) {
        add.push({ id: id, name: names[id] || id, nick: 'HR', date: d, day: 0, hours: delta, note: a[3] });
        adjusted.push(d + '  ' + id + ' ' + (names[id] || '') + ': +' + delta + 'h (new row) — ' + a[3]);
      } else {
        adjusted.push(d + '  ' + id + ': NOT APPLIED — no entry on that day to reduce');
        return;
      }
      total += (e ? e.hours - (e.hours - delta) : delta);
      perMan[id] = (perMan[id] || 0) + delta;
    });
    add.forEach(e => { if (e.hours <= 0) e.drop = true; });
    for (let i = add.length - 1; i >= 0; i--) if (add[i].drop) add.splice(i, 1);

    // ── back-check against the punches ───────────────────────────
    const ad = ss.getSheetByName('Attendance_Daily');
    const A = ad.getRange(1, 1, ad.getLastRow(), ad.getLastColumn()).getValues();
    const AH = A[0].map(h => String(h).trim());
    const aI = AH.indexOf('Employee ID'), aD = AH.indexOf('Date'), aO = AH.indexOf('OT Hours'),
          aCo = AH.indexOf('Check Out'), aCi = AH.indexOf('Check In'), aSrc = AH.indexOf('Source'), aSt = AH.indexOf('Status');
    const att = {};
    for (let i = 1; i < A.length; i++) {
      const d = otc_norm_(A[i][aD]); if (!inRange(d)) continue;
      att[String(A[i][aI]).trim().toUpperCase() + '|' + d] = {
        ot: Number(A[i][aO]) || 0, src: String(A[i][aSrc] || ''), st: String(A[i][aSt] || ''),
        full: String(A[i][aSrc]) === 'Kiosk' && String(A[i][aSt]) === 'Present' && A[i][aCo] !== '' && A[i][aCo] != null,
        ci: otc_clock_(A[i][aCi]), co: otc_clock_(A[i][aCo]) };
    }
    const B = { agree: 0, above: 0, below: 0, noPunch: 0, noData: 0 };
    const above = [];
    add.forEach(e => {
      const a = att[e.id + '|' + e.date];
      if (!a) { B.noData += e.hours; return; }
      if (!a.full) { B.noPunch += e.hours; return; }
      if (Math.abs(e.hours - a.ot) < 1) B.agree += e.hours;
      else if (e.hours > a.ot) { B.above += e.hours - a.ot; above.push({ e: e, a: a }); }
      else B.below += a.ot - e.hours;
    });

    // Men whose punches show OT in the range but who have no line in his sheet.
    const inSheet = {}; add.forEach(e => { inSheet[e.id] = true; });
    const punchOnly = {};
    Object.keys(att).forEach(k => {
      const id = k.split('|')[0];
      if (!inSheet[id] && att[k].full && att[k].ot > 0.05) punchOnly[id] = (punchOnly[id] || 0) + att[k].ot;
    });

    // ── what it replaces in OT_Entries ───────────────────────────
    let ot = ss.getSheetByName(OTC_ENTRIES_TAB);
    const ec = otc_entryCols_(ot);
    const all = ot.getRange(1, 1, ot.getLastRow(), Math.max(ot.getLastColumn(), 6)).getValues();
    const keep = [all[0]], gone = [];
    for (let i = 1; i < all.length; i++) {
      const r = all[i];
      if (r.every(c => c === '' || c == null)) continue;
      const d = otc_anyDate_(r[ec.date]);
      if (inRange(d)) gone.push(r); else keep.push(r);
    }
    const goneBy = {};
    let goneH = 0;
    gone.forEach(r => {
      const rem = String(r[ec.remark] || '') + ' ' + String(r[ec.by] || '');
      const k = /payroll close/i.test(rem) ? 'payroll close (claims, today)'
              : /auto from attendance/i.test(rem) ? '"Auto from attendance"'
              : /compiler sheet/i.test(rem) ? 'compiler sheet (earlier import)'
              : 'approved claims / other';
      goneBy[k] = (goneBy[k] || 0) + (Number(r[ec.hours]) || 0); goneH += Number(r[ec.hours]) || 0;
    });

    // ── pending claims in range ──────────────────────────────────
    const L = getOtCompileList({ from: '01-01-2000', to: OTS_TO });
    const settle = (L.rows || []).filter(r => inRange(r.belongsTo) || inRange(r.date));

    // ── report ───────────────────────────────────────────────────
    out.push((commit ? 'DONE' : 'PREVIEW — nothing written yet') + ' — OT compiler sheet "' + OTS_SOURCE + '"');
    out.push('day 1 = ' + OTS_FROM + '   day 31 = ' + OTS_TO + '   (shift of 19-09 ends the morning of 20-09; the 20-09 shift stays as decided from the claims)');
    out.push('');
    out.push('HIS SHEET: ' + hm(total) + ' on ' + add.length + ' man-day(s), ' + Object.keys(perMan).length + ' men');
    out.push('   ' + hm(B.agree)   + '  agrees with the punches (within 1h)');
    out.push('   ' + hm(B.above)   + '  MORE than the punches show, on ' + above.length + ' day(s) — listed below');
    out.push('   ' + hm(B.below)   + '  less than the punches show');
    out.push('   ' + hm(B.noPunch) + '  on days the punches cannot measure (22-31 Aug manual days, drivers, open check-ins, day off)');
    out.push('   ' + hm(B.noData)  + '  on days with no attendance at all in the system (20-21 Aug)');
    if (adjusted.length) { out.push(''); out.push('HR CORRECTIONS:'); adjusted.forEach(a => out.push('   ' + a)); }
    if (unmapped.length) { out.push(''); out.push('NOT IMPORTED — name not matched to an employee:'); unmapped.forEach(u => out.push('   ' + u)); }
    const po = Object.keys(punchOnly);
    if (po.length) {
      out.push('');
      out.push('NOT IN HIS SHEET — punches show OT but he has no line for them (they get NO OT for these dates):');
      po.sort().forEach(id => out.push('   ' + pad(id + ' ' + (names[id] || ''), 36) + hm(punchOnly[id]) + ' in punches'));
    }
    out.push('');
    out.push('REPLACES in OT_Entries (' + OTS_FROM + ' → ' + OTS_TO + '): ' + gone.length + ' row(s), ' + hm(goneH));
    Object.keys(goneBy).forEach(k => out.push('   ' + pad(hm(goneBy[k]), 10) + k));
    out.push('   Net change to OT paid for these dates: ' + (total - goneH >= 0 ? '+' : '−') + hm(Math.abs(total - goneH)));
    out.push('CLOSES ' + settle.length + ' pending claim(s) in the range as "' + OTS_SETTLED + '"');
    out.push('');
    out.push('PER MAN (his sheet):');
    Object.keys(perMan).sort().forEach(id => out.push('   ' + pad(id + ' ' + (names[id] || ''), 36) + hm(perMan[id])));
    out.push('');
    out.push('HIS SHEET ABOVE THE PUNCHES (' + above.length + ') — check these with him:');
    above.sort((x, y) => (y.e.hours - y.a.ot) - (x.e.hours - x.a.ot)).forEach(x =>
      out.push('   ' + pad(x.e.date, 12) + pad(x.e.id + ' ' + x.e.name, 30) + 'sheet ' + pad(hm(x.e.hours), 8) +
               'punches ' + pad(hm(x.a.ot), 8) + '(' + x.a.ci + ' → ' + x.a.co + ')'));

    if (!commit) {
      out.push('');
      out.push('Nothing has been changed. If this is right, run commitCompilerSheetImport().');
      return otc_log_(out);
    }

    // backup, then rewrite OT_Entries
    const stamp = Utilities.formatDate(new Date(), tz_(), 'dd-MM HH:mm');
    ot.copyTo(ss).setName('OT_Entries backup ' + stamp);
    const width = keep[0].length;
    const now = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm:ss');
    const newRows = add.map(e => {
      const o = new Array(width).fill('');
      o[ec.date] = e.date; o[ec.empId] = e.id; o[ec.name] = e.name; o[ec.hours] = e.hours;
      if (ec.remark >= 0) o[ec.remark] = 'Compiler sheet "' + OTS_SOURCE + '" day ' + e.day + ' ("' + e.nick + '")' +
                                         (e.note ? ' — ' + e.note : '') + ' — imported ' + now;
      if (ec.by >= 0) o[ec.by] = OTS_BY;
      return o;
    });
    // Every date is written as text dd-MM-yyyy — no more "Sat Sep 05 2026 …".
    const final = keep.map((r, i) => {
      const o = r.slice(0, width); while (o.length < width) o.push('');
      if (i > 0) o[ec.date] = otc_anyDate_(o[ec.date]) || o[ec.date];
      return o;
    }).concat(newRows);
    ot.getRange(1, 1, ot.getMaxRows(), Math.max(ot.getLastColumn(), width)).clearContent();
    ot.getRange(1, ec.date + 1, final.length, 1).setNumberFormat('@');
    ot.getRange(1, 1, final.length, width).setValues(final);

    const pen = ss.getSheetByName(OTC_PENDING_TAB);
    settle.forEach(r => pen.getRange(r.row, OTC_COL.status + 1).setValue(OTS_SETTLED));
    if (typeof logAmend_ === 'function') {
      logAmend_(ss, OTS_BY, 'OT IMPORT', '', OTS_FROM + ' → ' + OTS_TO, 'OT',
        'replaced ' + gone.length + ' OT_Entries row(s) (' + hm(goneH) + ') with ' + add.length +
        ' from "' + OTS_SOURCE + '" (' + hm(total) + '); backup tab "OT_Entries backup ' + stamp + '"');
      settle.forEach(r => logAmend_(ss, OTS_BY, 'OT SETTLED', r.empId, r.belongsTo, 'OT',
        'claimed ' + r.claimedHm + ' — closed: the compiler sheet covers this day'));
    }
    SpreadsheetApp.flush();
    out.push('');
    out.push('Written: ' + newRows.length + ' row(s) to OT_Entries; ' + gone.length + ' replaced; ' + settle.length +
             ' claim(s) closed. Backup: "OT_Entries backup ' + stamp + '".');
    return otc_log_(out);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================================
// OT_ENTRIES = HIS SHEET ONLY  (Shakeel, 22-09-2026)
// previewKeepOnlyCompilerSheetOt() / commitKeepOnlyCompilerSheetOt()
// ------------------------------------------------------------
// Payroll pays the whole OT_Entries tab. This payroll's OT comes ONLY from
// the OT compiler's sheet, so every row NOT written by the sheet import
// (e.g. the 20-09 shift paid from claims by the payroll close) is removed.
// A backup tab is made first. The claims behind removed rows are left as
// they are in OT_Pending; nothing else is touched.
// ============================================================
function previewKeepOnlyCompilerSheetOt() { return ots_keepOnly_(false); }
function commitKeepOnlyCompilerSheetOt()  { return ots_keepOnly_(true); }

function ots_keepOnly_(commit) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const lock = LockService.getScriptLock();
  if (commit && !lock.tryLock(30000)) { Logger.log('Busy — try again in a minute.'); return 'busy'; }
  try {
    const ot = ss.getSheetByName(OTC_ENTRIES_TAB);
    const ec = otc_entryCols_(ot);
    const all = ot.getRange(1, 1, ot.getLastRow(), Math.max(ot.getLastColumn(), 6)).getValues();
    const keep = [all[0]], drop = [];
    let keepH = 0, dropH = 0;
    for (let i = 1; i < all.length; i++) {
      const r = all[i];
      if (r.every(c => c === '' || c == null)) continue;
      const fromSheet = String(r[ec.by] || '') === OTS_BY;
      if (fromSheet) { keep.push(r); keepH += Number(r[ec.hours]) || 0; }
      else { drop.push(r); dropH += Number(r[ec.hours]) || 0; }
    }
    const out = [];
    out.push((commit ? 'DONE' : 'PREVIEW — nothing written yet') + ' — OT_Entries = the OT compiler\'s sheet only');
    out.push('Keep   : ' + (keep.length - 1) + ' row(s), ' + otc_hm_(keepH) + '  (' + OTS_BY + ')');
    out.push('Remove : ' + drop.length + ' row(s), ' + otc_hm_(dropH));
    drop.forEach(r => out.push('   ' + otc_anyDate_(r[ec.date]) + '  ' + r[ec.empId] + '  ' + String(r[ec.name] || '') +
                               '  ' + otc_hm_(Number(r[ec.hours]) || 0) + '  — ' + String(r[ec.by] || '')));
    if (!commit) {
      out.push('');
      out.push(drop.length ? 'Nothing has been changed. If this is right, run commitKeepOnlyCompilerSheetOt().' : 'Nothing to remove.');
      return otc_log_(out);
    }
    if (!drop.length) return otc_log_(out.concat(['', 'Nothing to remove.']));
    const stamp = Utilities.formatDate(new Date(), tz_(), 'dd-MM HH:mm');
    ot.copyTo(ss).setName('OT_Entries backup ' + stamp);
    const width = keep[0].length;
    const final = keep.map(r => { const o = r.slice(0, width); while (o.length < width) o.push(''); return o; });
    ot.getRange(1, 1, ot.getMaxRows(), Math.max(ot.getLastColumn(), width)).clearContent();
    ot.getRange(1, ec.date + 1, final.length, 1).setNumberFormat('@');
    ot.getRange(1, 1, final.length, width).setValues(final);
    if (typeof logAmend_ === 'function')
      logAmend_(ss, 'System (OT = compiler sheet only)', 'OT REMOVE', '', '', 'OT',
        'removed ' + drop.length + ' OT_Entries row(s) (' + otc_hm_(dropH) + ') not from the compiler sheet; backup "OT_Entries backup ' + stamp + '"');
    SpreadsheetApp.flush();
    out.push('');
    out.push('Removed. OT_Entries now holds ' + (final.length - 1) + ' row(s), ' + otc_hm_(keepH) + '. Backup: "OT_Entries backup ' + stamp + '".');
    return otc_log_(out);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}