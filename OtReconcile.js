// ============================================================
// OtReconcile.gs — READ ONLY. Compares OT three ways.
// ------------------------------------------------------------
// Writes nothing. No deletes, no edits, no posting.
//
// THE QUESTION THIS ANSWERS
// OT currently exists in two places that can disagree:
//   COMPUTED — derived from the actual punches in Attendance_Log
//              (presence minus the shift's normal hours)
//   CLAIMED  — what the employee submitted themselves on the kiosk,
//              sitting in OT_Pending
// And a third, once approved:
//   POSTED   — what is in OT_Entries and will actually be PAID
//
// Payroll pays POSTED. So any gap between COMPUTED and CLAIMED is a
// decision someone has to make before approving, and any gap between
// those and POSTED is money already committed on a number nobody
// reconciled.
//
// ROUNDING
// Raw computed OT is ugly — 0.83h, 1.27h. Two conventions are shown
// side by side so you can pick one as policy:
//   ROUND_15 : nearest 15 minutes (0.83h → 0.75h)
//   ROUND_30 : nearest 30 minutes (0.83h → 1.00h)
// Nothing is applied. The report just shows what each would produce.
//
// USAGE:  set the range below, run reconcileOt()
// ============================================================

const OTR_FROM = '22-08-2026';
const OTR_TO   = '';            // blank = yesterday

// A gap this size or smaller is treated as agreement — not worth a
// human looking at it.
const OTR_TOLERANCE_HOURS = 0.25;   // 15 minutes

function otr_round_(hours, minutes) {
  const step = minutes / 60;
  return Math.round(hours / step) * step;
}

function otr_hm_(h) {
  if (h == null || h === '' || isNaN(Number(h))) return '—';
  h = Number(h);
  if (h === 0) return '0h';
  const neg = h < 0; h = Math.abs(h);
  const hrs = Math.floor(h), min = Math.round((h - hrs) * 60);
  const s = (hrs > 0 ? hrs + 'h' : '') + (min > 0 ? (hrs > 0 ? ' ' : '') + min + 'm' : '');
  return (neg ? '-' : '') + (s || '0h');
}

function otr_norm_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[1];
  const d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}

function reconcileOt() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

  const from = startOfDay_(parseFlexibleDate_(OTR_FROM));
  let toRaw;
  if (OTR_TO) toRaw = parseFlexibleDate_(OTR_TO);
  else { toRaw = new Date(); toRaw.setDate(toRaw.getDate() - 1); }
  const to = endOfDay_(toRaw);
  const inRange = (dStr) => {
    const d = parseFlexibleDate_(dStr);
    return d && startOfDay_(d) >= from && startOfDay_(d) <= to;
  };

  // ── names ──
  const hr = ss.getSheetByName('HR Maya');
  const hd = hr.getDataRange().getValues();
  const HH = hd[0];
  const hc = (n) => HH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
  const hId = hc('Employee ID'), hNm = hc('Name (English)');
  let hLoc = hc('Location'); if (hLoc === -1) hLoc = hc('Branch');
  const nameOf = {}, locOf = {};
  for (let i = 1; i < hd.length; i++) {
    const id = String(hd[i][hId] || '').trim();
    if (!id) continue;
    nameOf[id] = hNm !== -1 ? String(hd[i][hNm] || '').trim() : id;
    locOf[id]  = hLoc !== -1 ? String(hd[i][hLoc] || '').trim() : '';
  }

  // ── COMPUTED: from Attendance_Daily (already sessionised from punches) ──
  const computed = {};
  const presenceOf = {};
  const ad = ss.getSheetByName('Attendance_Daily');
  if (ad && ad.getLastRow() > 1) {
    const add = ad.getDataRange().getValues();
    const AH = add[0];
    const ac = (n) => AH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const aId = ac('Employee ID'), aDate = ac('Date'),
          aOt = ac('OT Hours'), aPr = ac('Presence (hrs)'), aSt = ac('Status');
    for (let i = 1; i < add.length; i++) {
      const dStr = otr_norm_(add[i][aDate]);
      if (!inRange(dStr)) continue;
      const id = String(add[i][aId] || '').trim();
      if (!id) continue;
      const k = id + '|' + dStr;
      computed[k] = Number(add[i][aOt]) || 0;
      presenceOf[k] = { presence: Number(add[i][aPr]) || 0,
                        status: String(add[i][aSt] || '').trim() };
    }
  }

  // ── CLAIMED: OT_Pending (0 ID | 1 EmpID | 2 Name | 3 Date | 4 Hours | 6 Status) ──
  const claimed = {}, claimStatus = {};
  const pen = ss.getSheetByName('OT_Pending');
  if (pen && pen.getLastRow() > 1) {
    pen.getRange(2, 1, pen.getLastRow()-1, pen.getLastColumn()).getValues().forEach(r => {
      const id = String(r[1] || '').trim();
      const dStr = otr_norm_(r[3]);
      if (!id || !inRange(dStr)) return;
      const k = id + '|' + dStr;
      claimed[k] = (claimed[k] || 0) + (Number(r[4]) || 0);
      claimStatus[k] = String(r[6] || '').trim();
    });
  }

  // ── POSTED: OT_Entries (0 Date | 1 EmpID | 3 Hours) — what gets PAID ──
  const posted = {};
  const ot = ss.getSheetByName('OT_Entries');
  if (ot && ot.getLastRow() > 1) {
    ot.getRange(2, 1, ot.getLastRow()-1, ot.getLastColumn()).getValues().forEach(r => {
      const id = String(r[1] || '').trim();
      const dStr = otr_norm_(r[0]);
      if (!id || !inRange(dStr)) return;
      const k = id + '|' + dStr;
      posted[k] = (posted[k] || 0) + (Number(r[3]) || 0);
    });
  }

  // ── Merge every key seen anywhere ──
  const keys = {};
  [computed, claimed, posted].forEach(m => Object.keys(m).forEach(k => keys[k] = true));

  const agree = [], disagree = [], claimOnly = [], computedOnly = [], postedMismatch = [];
  let sumComputed = 0, sumClaimed = 0, sumPosted = 0;

  Object.keys(keys).forEach(k => {
    const p = k.split('|'), id = p[0], dStr = p[1];
    const cm = computed[k] || 0, cl = claimed[k] || 0, po = posted[k] || 0;
    sumComputed += cm; sumClaimed += cl; sumPosted += po;

    const pr = presenceOf[k] || { presence: 0, status: '' };
    const base = [id, nameOf[id] || '?', dStr, locOf[id] || '',
                  pr.status, otr_hm_(pr.presence),
                  otr_hm_(cm), otr_hm_(cl), otr_hm_(po),
                  otr_hm_(otr_round_(cm, 15)), otr_hm_(otr_round_(cm, 30))];

    const diff = Math.abs(cm - cl);

    if (po > 0 && Math.abs(po - cm) > OTR_TOLERANCE_HOURS && Math.abs(po - cl) > OTR_TOLERANCE_HOURS) {
      postedMismatch.push(base.concat(['PAID ' + otr_hm_(po) +
        ' but punches say ' + otr_hm_(cm) + ' and claim says ' + otr_hm_(cl)]));
    } else if (cm > 0 && cl === 0) {
      computedOnly.push(base.concat(['Punches show OT but nothing was claimed on the kiosk']));
    } else if (cl > 0 && cm === 0) {
      claimOnly.push(base.concat(['Claimed on the kiosk but punches show no OT — verify']));
    } else if (diff > OTR_TOLERANCE_HOURS) {
      disagree.push(base.concat(['Differ by ' + otr_hm_(diff)]));
    } else if (cm > 0 || cl > 0) {
      agree.push(base.concat(['Agree within ' + otr_hm_(OTR_TOLERANCE_HOURS)]));
    }
  });

  const bydate = (a, b) => (a[2].slice(6) + a[2].slice(3,5) + a[2].slice(0,2) + a[1])
    .localeCompare(b[2].slice(6) + b[2].slice(3,5) + b[2].slice(0,2) + b[1]);
  [agree, disagree, claimOnly, computedOnly, postedMismatch].forEach(x => x.sort(bydate));

  // ── Report ──
  const L = [];
  L.push('══════════════════════════════════════════════');
  L.push(' OT RECONCILIATION — read only, nothing changed');
  L.push(' ' + OTR_FROM + ' → ' + formatDdMmYyyy_(toRaw));
  L.push(' Tolerance: ' + otr_hm_(OTR_TOLERANCE_HOURS));
  L.push('══════════════════════════════════════════════');
  L.push('');
  L.push('TOTALS');
  L.push('   From punches (computed) : ' + otr_hm_(sumComputed));
  L.push('   Claimed on kiosk        : ' + otr_hm_(sumClaimed));
  L.push('   Posted to OT_Entries    : ' + otr_hm_(sumPosted) + '   ← this is what PAYROLL PAYS');
  L.push('');
  L.push('   If computed and posted differ, payroll is paying a number');
  L.push('   that the punches do not support.');
  L.push('');
  L.push('ROUNDING — what the computed total would become as policy');
  L.push('   Nearest 15 min : ' + otr_hm_(otr_round_(sumComputed, 15)));
  L.push('   Nearest 30 min : ' + otr_hm_(otr_round_(sumComputed, 30)));
  L.push('');
  L.push('COUNTS');
  L.push('   Agree                       : ' + agree.length);
  L.push('   Disagree                    : ' + disagree.length);
  L.push('   Claimed but no OT in punches: ' + claimOnly.length);
  L.push('   OT in punches, not claimed  : ' + computedOnly.length);
  L.push('   Already PAID and unsupported: ' + postedMismatch.length);
  L.push('');

  const HDR = ['ID','Name','Date','Location','Status','Presence',
               'Computed','Claimed','Posted','Round15','Round30','Note'];
  const dump = (title, arr, note) => {
    L.push('──────────────────────────────────────────────');
    L.push(title + '  (' + arr.length + ')');
    if (note) L.push(note);
    if (!arr.length) { L.push('   None.'); L.push(''); return; }
    L.push('   ' + HDR.join(' | '));
    arr.forEach(r => L.push('   ' + r.join(' | ')));
    L.push('');
  };

  dump('A. ALREADY PAID BUT UNSUPPORTED', postedMismatch,
       '   OT_Entries holds a figure neither the punches nor the claim back up.');
  dump('B. DISAGREE — punches vs kiosk claim', disagree,
       '   Someone must decide which is true before approving.');
  dump('C. CLAIMED ON KIOSK, NO OT IN PUNCHES', claimOnly,
       '   Either the punches are wrong, or the claim is.');
  dump('D. OT IN PUNCHES, NEVER CLAIMED', computedOnly,
       '   The employee worked over their shift but did not submit. Unpaid unless posted.');
  dump('E. AGREE', agree, '   No action needed.');

  L.push('══════════════════════════════════════════════');
  const out = L.join('\n');
  Logger.log(out);
  return out;
}