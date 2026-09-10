// ============================================================
// WhatsAppPins.gs — build the PIN-broadcast recipient list (with dry-run)
// ============================================================
// Exports each employee's Name, Employee ID, PIN and WhatsApp number so the
// bulk sender can message each person THEIR OWN credentials. Always run the
// dry-run first and eyeball every number — a wrong number sends a PIN to the
// wrong person. Returns a verification report; getPinContactsCsv returns CSV.
//
// ENDPOINTS: getPinBroadcastList, getPinContactsCsv
// ============================================================

// Normalise a WhatsApp number to digits with country code (Kuwait 965 default).
function normWa_(raw) {
  let s = String(raw || '').trim();
  if (!s) return '';
  s = s.replace(/[^\d+]/g, '');
  s = s.replace(/^00/, '').replace(/^\+/, '');
  if (/^965\d{8}$/.test(s)) return s;          // already 965 + 8
  if (/^\d{8}$/.test(s)) return '965' + s;      // bare Kuwait 8-digit
  return s;                                     // foreign/other — leave as-is, flagged below
}

function getPinBroadcastList(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const hr = ss.getSheetByName('HR Maya');
    const d = hr.getDataRange().getValues();
    const H = d[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cId = c('Employee ID'), cName = c('Name (English)'), cPin = c('PIN'), cWa = c('WhatsApp'), cStatus = c('Status');

    const ready = [], skipped = [];
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][cId] || '').trim();
      if (!id) continue;
      const status = cStatus !== -1 ? String(d[i][cStatus] || '').trim().toLowerCase() : '';
      const name = cName !== -1 ? String(d[i][cName] || '').trim() : id;
      const pin = cPin !== -1 ? String(d[i][cPin] || '').trim() : '';
      const waRaw = cWa !== -1 ? String(d[i][cWa] || '').trim() : '';
      const wa = normWa_(waRaw);

      const reasons = [];
      if (!pin) reasons.push('no PIN');
      if (!waRaw) reasons.push('no WhatsApp');
      else if (!/^965\d{8}$/.test(wa)) reasons.push('number not a standard Kuwait 965+8 (' + wa + ')');
      if (status && status !== 'active') reasons.push('status ' + status);

      if (reasons.length) skipped.push({ id: id, name: name, whatsapp: waRaw, reasons: reasons.join(', ') });
      else ready.push({ id: id, name: name, pin: pin, whatsapp: wa });
    }
    return {
      success: true, readyCount: ready.length, skippedCount: skipped.length,
      ready: ready, skipped: skipped,
      message: ready.length + ' ready to message, ' + skipped.length + ' skipped (missing/odd data). REVIEW skipped list and all numbers before sending.'
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// CSV for the bulk sender: number,name,id,pin  (ready recipients only).
function getPinContactsCsv(params) {
  try {
    const list = getPinBroadcastList(params);
    if (!list.success) return list;
    const esc = (s) => '"' + String(s).replace(/"/g, '""') + '"';
    const lines = ['number,name,employee_id,pin'];
    list.ready.forEach(r => lines.push([r.whatsapp, esc(r.name), r.id, r.pin].join(',')));
    return { success: true, count: list.ready.length, csv: lines.join('\n'),
      message: list.ready.length + ' contacts. Save as contacts.csv for the sender.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}