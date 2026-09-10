// ============================================================
// EncashmentVoucher.gs — printable A5 leave-encashment voucher (dual signature)
// ============================================================
// generateEncashmentVoucher({ voucherNo }) or ({ empId }) -> builds an A5 PDF
// into Drive and returns its URL. Pulls the encashment from Encashment_Log,
// the employee's Basic + Designation + remaining from HR Maya. GM is the
// pre-printed management signatory. Rate shown as Basic/30.
//
// ENDPOINT: generateEncashmentVoucher
// ============================================================

const ENCASH_VOUCHER_FOLDER = 'Maya Encashment Vouchers';
const GM_NAME = 'Abdulrehman Alnajar';
const GM_TITLE = 'General Manager';

function ev_money_(v) { return (Math.round((Number(v) || 0) * 1000) / 1000).toFixed(3); }
function ev_esc_(s) { return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function ev_words_(amount) {
  const ones=['zero','one','two','three','four','five','six','seven','eight','nine','ten','eleven','twelve','thirteen','fourteen','fifteen','sixteen','seventeen','eighteen','nineteen'];
  const tens=['','','twenty','thirty','forty','fifty','sixty','seventy','eighty','ninety'];
  function three(n){let s='';if(n>=100){s+=ones[Math.floor(n/100)]+' hundred ';n%=100;}if(n>=20){s+=tens[Math.floor(n/10)]+' ';n%=10;}if(n>0)s+=ones[n]+' ';return s.trim();}
  let d=Math.floor(amount), f=Math.round((amount-d)*1000), parts=[];
  if(d===0)parts.push('zero'); else { if(d>=1000){parts.push(three(Math.floor(d/1000))+' thousand');d%=1000;} if(d>0)parts.push(three(d)); }
  let out=parts.filter(Boolean).join(' ').replace(/\b\w/g,c=>c.toUpperCase())+' Kuwaiti Dinar';
  if(f>0)out+=' and '+three(f).replace(/\b\w/g,c=>c.toUpperCase())+' Fils';
  return out+' only';
}

function generateEncashmentVoucher(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName(ENCASHMENT_LOG_TAB);
    if (!log) return { success: false, error: 'No Encashment_Log.' };
    const d = log.getDataRange().getValues();
    const H = d[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cId=c('Employee ID'), cName=c('Name'), cType=c('Type'), cDays=c('Days'), cAmt=c('Amount'),
          cPer=c('Period End'), cDate=c('Encash Date'), cBy=c('By'), cV=c('Voucher No');

    const wantVoucher = String(params.voucherNo || '').trim();
    const wantEmp = String(params.empId || '').trim().toUpperCase();
    let row = null;
    for (let i = d.length - 1; i >= 1; i--) { // latest first
      const vno = cV !== -1 ? String(d[i][cV] || '').trim() : '';
      const id = String(d[i][cId] || '').trim();
      if (wantVoucher && cV !== -1) { if (vno === wantVoucher) { row = d[i]; break; } }
      else if (wantEmp) { if (id.toUpperCase() === wantEmp) { row = d[i]; break; } }
    }
    // Fallback: if we searched by voucher but the column is missing/empty,
    // and a single latest row exists, use the most recent row.
    if (!row && wantVoucher && cV === -1 && d.length > 1) row = d[d.length - 1];
    if (!row) return { success: false, error: 'Encashment not found. (Tip: re-run the encashment after redeploying, or check the Encashment_Log has a "Voucher No" column.)' };

    const empId = String(row[cId] || '').trim();
    const name = String(row[cName] || '').trim();
    const type = String(row[cType] || '').trim();
    const days = Number(row[cDays]) || 0;
    const amount = Number(row[cAmt]) || 0;
    const encashDate = String(row[cDate] || '').trim();
    const periodEnd = String(row[cPer] || '').trim();
    const voucherNo = cV !== -1 ? String(row[cV] || '').trim() : '';

    // Employee details from HR Maya.
    const hr = ss.getSheetByName('HR Maya'); const hd = hr.getDataRange().getValues(); const hH = hd[0];
    const hc = (n) => hH.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const hcId=hc('Employee ID'), hcBasic=hc('Basic Salary'), hcDesig=hc('Designation'), hcCivil=hc('Civil ID');
    let basic=0, desig='', civil='';
    for (let i=1;i<hd.length;i++){ if(String(hd[i][hcId]).trim()===empId){ basic=Number(hd[i][hcBasic])||0; desig=hcDesig!==-1?String(hd[i][hcDesig]||''):''; civil=hcCivil!==-1?String(hd[i][hcCivil]||''):''; break; } }
    const rate = Math.round((basic/30)*1000)/1000;

    const today = Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy');
    const html =
      '<html><head><meta charset="utf-8"><style>' +
      '@page{size:A5;margin:12mm;} body{font-family:Arial,Helvetica,sans-serif;color:#222;font-size:11px;}' +
      'h2{text-align:center;margin:0;font-size:15px;} .sub{text-align:center;color:#666;font-size:10px;margin-bottom:10px;}' +
      '.ref{display:flex;justify-content:space-between;font-size:10px;color:#444;margin-bottom:10px;}' +
      'table{width:100%;border-collapse:collapse;margin-bottom:10px;} td{padding:5px 6px;border:0.5px solid #bbb;}' +
      '.lbl{color:#555;width:45%;} .r{text-align:right;} .amt{background:#f5f7fa;font-weight:bold;}' +
      '.words{font-style:italic;color:#333;margin:6px 0 14px;font-size:10px;}' +
      '.sig{display:flex;justify-content:space-between;margin-top:30px;} .sigbox{width:45%;text-align:center;}' +
      '.sigline{border-top:1px solid #444;margin-top:34px;padding-top:4px;font-size:10px;color:#444;}' +
      '.note{font-size:8.5px;color:#888;margin-top:16px;text-align:center;}' +
      '</style></head><body>' +
      '<h2>Maya Tex Restaurants Company</h2>' +
      '<div class="sub">LEAVE ENCASHMENT VOUCHER</div>' +
      '<div class="ref"><span><b>Voucher No:</b> ' + ev_esc_(voucherNo) + '</span><span><b>Date:</b> ' + today + '</span></div>' +
      '<table>' +
        '<tr><td class="lbl">Employee</td><td>' + ev_esc_(name) + '</td></tr>' +
        '<tr><td class="lbl">Employee ID</td><td>' + ev_esc_(empId) + '</td></tr>' +
        (desig ? '<tr><td class="lbl">Designation</td><td>' + ev_esc_(desig) + '</td></tr>' : '') +
        (civil ? '<tr><td class="lbl">Civil ID</td><td>' + ev_esc_(civil) + '</td></tr>' : '') +
      '</table>' +
      '<table>' +
        '<tr><td class="lbl">Leave Type</td><td>' + ev_esc_(type) + '</td></tr>' +
        '<tr><td class="lbl">Days Encashed</td><td>' + days + ' day(s)</td></tr>' +
        '<tr><td class="lbl">Basic Salary</td><td class="r">KD ' + ev_money_(basic) + '</td></tr>' +
        '<tr><td class="lbl">Rate (Basic \u00F7 30)</td><td class="r">KD ' + ev_money_(rate) + ' / day</td></tr>' +
        '<tr><td class="lbl amt">Encashment Amount</td><td class="r amt">KD ' + ev_money_(amount) + '</td></tr>' +
        '<tr><td class="lbl">Paid in period ending</td><td>' + ev_esc_(periodEnd) + '</td></tr>' +
      '</table>' +
      '<div class="words">Amount in words: ' + ev_esc_(ev_words_(amount)) + '</div>' +
      '<div class="sig">' +
        '<div class="sigbox"><div class="sigline">Employee<br>' + ev_esc_(name) + '<br>Signature &amp; Date</div></div>' +
        '<div class="sigbox"><div class="sigline">' + ev_esc_(GM_NAME) + '<br>' + ev_esc_(GM_TITLE) + '<br>Signature &amp; Date</div></div>' +
      '</div>' +
      '<div class="note">Leave encashment calculated at Basic \u00F7 30 per day (a month = 30 days). Amounts in Kuwaiti Dinar (KD), 3 decimals.</div>' +
      '</body></html>';

    const pdf = Utilities.newBlob(html, 'text/html', 'Encashment ' + voucherNo + '.html').getAs('application/pdf').setName('Encashment Voucher - ' + name + ' - ' + voucherNo + '.pdf');
    const folder = getOrCreateFolder_(ENCASH_VOUCHER_FOLDER);
    const file = folder.createFile(pdf);
    return { success: true, voucherNo: voucherNo, url: file.getUrl(),
      message: 'Voucher ' + voucherNo + ' generated for ' + name + ' (KD ' + ev_money_(amount) + ').' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// List encashment history (latest first) for the reprint panel.
function getEncashmentHistory(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName(ENCASHMENT_LOG_TAB);
    if (!log) return { success: true, rows: [] };
    const d = log.getDataRange().getValues();
    const H = d[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cId=c('Employee ID'), cName=c('Name'), cType=c('Type'), cDays=c('Days'), cAmt=c('Amount'),
          cPer=c('Period End'), cDate=c('Encash Date'), cV=c('Voucher No');
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      if (!String(d[i][cId] || '').trim()) continue;
      rows.push({ row: i + 1, empId: d[i][cId], name: d[i][cName], type: d[i][cType],
        days: Number(d[i][cDays]) || 0, amount: Number(d[i][cAmt]) || 0,
        periodEnd: d[i][cPer], encashDate: d[i][cDate],
        voucherNo: cV !== -1 ? String(d[i][cV] || '').trim() : '' });
    }
    rows.reverse(); // latest first
    return { success: true, rows: rows };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// Reprint / generate a voucher for a specific log row. Assigns a voucher no
// on the fly if the row doesn't have one (old pre-voucher encashments).
function reprintEncashmentVoucher(params) {
  try {
    params = params || {};
    const rowNum = Number(params.row);
    if (!rowNum || rowNum < 2) return { success: false, error: 'Valid row required.' };
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName(ENCASHMENT_LOG_TAB);
    let H = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0];
    let cV = H.findIndex(h => String(h).trim().toLowerCase() === 'voucher no');
    if (cV === -1) { log.getRange(1, log.getLastColumn() + 1).setValue('Voucher No').setFontWeight('bold'); H = log.getRange(1,1,1,log.getLastColumn()).getValues()[0]; cV = H.findIndex(h => String(h).trim().toLowerCase() === 'voucher no'); }
    let vno = String(log.getRange(rowNum, cV + 1).getValue() || '').trim();
    if (!vno) {
      vno = 'ENC-' + Utilities.formatDate(new Date(), tz_(), 'yyyy') + '-' + Utilities.formatDate(new Date(), tz_(), 'MMddHHmmss') + '-R';
      log.getRange(rowNum, cV + 1).setValue(vno);
    }
    return generateEncashmentVoucher({ voucherNo: vno });
  } catch (err) { return { success: false, error: err.toString() }; }
}