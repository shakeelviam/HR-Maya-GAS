// ============================================================
// PayrollDocs.gs — Maya HR : period report + A5 payslip PDFs
// ============================================================
// Standalone Apps Script file. Shares global scope with Code.gs,
// Payroll.gs, SharedUtils.gs. Uses KD (Payroll.gs) and the date
// helpers in SharedUtils.gs. Do NOT redeclare them.
//
// Reads a period tab (e.g. "Payroll 21Jul-20Aug") produced by
// computeMonthlyPayroll and generates:
//   1. ONE combined PDF of A5 payslips — one payslip per page, every
//      employee on the tab (headcount-agnostic: new hires just appear).
//   2. A one-page A5 summary report (period totals).
// Both are saved to a Drive folder and the links returned.
//
// Endpoint (add to Code.gs handleRequest switch):
//   case 'generatePayrollDocs':
//     return ContentService.createTextOutput(JSON.stringify(generatePayrollDocs(e.parameter))).setMimeType(ContentService.MimeType.JSON);
//
// Call: ?method=generatePayrollDocs&tab=Payroll%2021Jul-20Aug
//   (omit tab to use the most recent Payroll_Control run)
// ============================================================

const PAYSLIP_FOLDER_NAME = 'Maya HR Payslips';

function generatePayrollDocs(params) {
  try {
    params = params || {};
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    // Resolve which period tab to use.
    let tabName = String(params.tab || '').trim();
    let periodStart = '', periodEnd = '';
    if (!tabName) {
      const last = getLastPayrollRun_(ss);
      if (!last) return { success: false, error: 'No tab given and no runs found in Payroll_Control.' };
      tabName = last.tabName;
      periodStart = last.periodStart ? formatDdMmYyyy_(last.periodStart) : '';
      periodEnd = last.periodEnd ? formatDdMmYyyy_(last.periodEnd) : '';
    }
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return { success: false, error: 'Period tab "' + tabName + '" not found.' };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, error: 'Tab "' + tabName + '" has no rows.' };

    const H = data[0];
    const c = (name) => H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
    const idx = {
      id: c('Employee ID'), name: c('Name (English)'), basic: c('Basic Salary'),
      food: c('Food Allowance'), accom: c('Accommodation'), conv: c('Conveyance'),
      bonus: c('Bonus'), gross: c('Gross Salary'), otHrs: c('OT Hours'),
      rate: c('Price Per Hour'), otAmt: c('OT Amount'),
      encDays: c('Encash Days'), enc: c('Encashment'), ded: c('Deductions'),
      remark: c('Deduction Remark'), pay: c('Total Payable'), flags: c('Flags'),
      // New itemised review-tab columns (present when run from the Payroll Review screen)
      bonusAdd: c('Bonus/Additional'), unpaid: c('Unpaid Leave'), loan: c('Loan'),
      advance: c('Advance'), disc: c('Disciplinary'), totDed: c('Total Deductions'), net: c('Net Payable')
    };

    // Collect employee rows (skip the TOTALS row, which has a blank Employee ID).
    const emps = [];
    let totals = null;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[idx.id] || '').trim();
      if (!id) { totals = row; continue; } // TOTALS row
      emps.push(row);
    }
    if (!emps.length) return { success: false, error: 'No employee rows found on "' + tabName + '".' };

    // Period label from the tab name if we didn't get it from control.
    const periodLabel = (periodStart && periodEnd) ? (periodStart + ' to ' + periodEnd) : tabName.replace(/^Payroll\s*/, '');

    // Build one HTML doc: A5 summary page + one A5 payslip per employee.
    const html = buildPayrollHtml_(tabName, periodLabel, emps, totals, idx);

    // Convert HTML -> single PDF (A5 honored via @page).
    const pdfBlob = Utilities.newBlob(html, 'text/html', tabName + '.html').getAs('application/pdf')
                    .setName(tabName + ' — Payslips.pdf');

    // Save to a Drive folder, overwriting any previous PDF for this tab.
    const folder = getOrCreateFolder_(PAYSLIP_FOLDER_NAME);
    const fileName = tabName + ' — Payslips.pdf';
    const existing = folder.getFilesByName(fileName);
    let replaced = false;
    while (existing.hasNext()) { existing.next().setTrashed(true); replaced = true; }
    const file = folder.createFile(pdfBlob);

    return {
      success: true, tab: tabName, period: periodLabel, employees: emps.length,
      pdfUrl: file.getUrl(), folderUrl: folder.getUrl(), replacedOld: replaced,
      message: 'Generated ' + emps.length + ' A5 payslips + summary for "' + tabName + '"' +
        (replaced ? ' (replaced the previous PDF).' : '.')
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function money_(v) {
  const n = Number(v) || 0;
  return n.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
}
function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Amount-in-words (KD + fils), 0..9999 dinars.
function words_(amount) {
  const ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
  const tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  function three(n){ let w=''; if(n>=100){w+=ones[Math.floor(n/100)]+' Hundred '; n%=100;} if(n>=20){w+=tens[Math.floor(n/10)]+' '; n%=10;} if(n>0){w+=ones[n]+' ';} return w.trim(); }
  const dinars = Math.floor(amount);
  const fils = Math.round((amount - dinars) * 1000);
  let t = '';
  if (dinars === 0) t = 'Zero';
  else { if (dinars >= 1000){ t += three(Math.floor(dinars/1000)) + ' Thousand '; } const r = dinars % 1000; if (r) t += three(r); }
  t = t.trim() + (dinars === 1 ? ' Kuwaiti Dinar' : ' Kuwaiti Dinars');
  if (fils > 0) t += ' and ' + String(fils).padStart(3, '0') + ' Fils';
  return t + ' only';
}

function buildPayrollHtml_(tabName, periodLabel, emps, totals, idx) {
  const g = (row, key) => (idx[key] !== -1 ? row[idx[key]] : '');

  // --- Summary page ---
  let summaryRows = '';
  emps.forEach((r) => {
    summaryRows +=
      '<tr><td>' + esc_(g(r,'id')) + '</td><td>' + esc_(g(r,'name')) + '</td>' +
      '<td class="r">' + money_(g(r,'gross')) + '</td>' +
      '<td class="r">' + money_(g(r,'otAmt')) + '</td>' +
      '<td class="r">' + money_(g(r,'ded')) + '</td>' +
      '<td class="r b">' + money_(g(r,'pay')) + '</td></tr>';
  });
  const tG = totals ? money_(g(totals,'gross')) : '';
  const tO = totals ? money_(g(totals,'otAmt')) : '';
  const tD = totals ? money_(g(totals,'ded')) : '';
  const tP = totals ? money_((idx.net !== -1 && g(totals,'net') !== '') ? g(totals,'net') : g(totals,'pay')) : '';

  let summaryPage =
    '<section class="page">' +
      '<div class="co">MAYA TEX RESTAURANTS COMPANY</div>' +
      '<div class="sub">Payroll Summary — ' + esc_(periodLabel) + '</div>' +
      '<table class="grid"><thead><tr>' +
        '<th>ID</th><th>Name</th><th class="r">Gross</th><th class="r">OT</th><th class="r">Deduct</th><th class="r">Net Payable</th>' +
      '</tr></thead><tbody>' + summaryRows +
      '<tr class="tot"><td></td><td>TOTALS (' + emps.length + ' staff)</td>' +
        '<td class="r">' + tG + '</td><td class="r">' + tO + '</td><td class="r">' + tD + '</td><td class="r b">' + tP + '</td></tr>' +
      '</tbody></table>' +
      '<div class="note">All amounts in Kuwaiti Dinar (KD), to 3 decimals. Generated ' + formatDdMmYyyy_(new Date()) + '.</div>' +
    '</section>';

  // --- One payslip per employee ---
  let slips = '';
  emps.forEach((r) => {
    const pay = (idx.net !== -1 && g(r,'net') !== '') ? (Number(g(r,'net')) || 0) : (Number(g(r,'pay')) || 0);
    const otHrs = g(r,'otHrs');
    const remark = g(r,'remark');
    const flags = g(r,'flags');
    const lines = [];
    lines.push(['Basic Salary', money_(g(r,'basic'))]);
    if (Number(g(r,'food')))  lines.push(['Food Allowance', money_(g(r,'food'))]);
    if (Number(g(r,'accom'))) lines.push(['Accommodation', money_(g(r,'accom'))]);
    if (Number(g(r,'conv')))  lines.push(['Conveyance', money_(g(r,'conv'))]);
    if (Number(g(r,'bonus'))) lines.push(['Bonus', money_(g(r,'bonus'))]);
    lines.push(['Gross Salary', money_(g(r,'gross')), 'sub']);
    if (Number(otHrs)) lines.push(['Overtime (' + otHrs + ' hrs @ ' + money_(g(r,'rate')) + ')', money_(g(r,'otAmt'))]);
    if (Number(g(r,'enc'))) lines.push(['Leave Encashment (' + (g(r,'encDays') || 0) + ' days)', money_(g(r,'enc'))]);
    // Bonus/Additional (itemised review tab)
    if (idx.bonusAdd !== -1 && Number(g(r,'bonusAdd'))) lines.push(['Bonus / Additional', money_(g(r,'bonusAdd'))]);
    // Itemised deductions when the review-tab columns exist; else the single 'Deductions'.
    const hasItemised = idx.totDed !== -1;
    if (hasItemised) {
      if (Number(g(r,'unpaid')))  lines.push(['Unpaid Leave' + (remark ? ' (' + esc_(remark) + ')' : ''), '(' + money_(g(r,'unpaid')) + ')']);
      if (Number(g(r,'loan')))    lines.push(['Loan', '(' + money_(g(r,'loan')) + ')']);
      if (Number(g(r,'advance'))) lines.push(['Advance', '(' + money_(g(r,'advance')) + ')']);
      if (Number(g(r,'disc')))    lines.push(['Disciplinary', '(' + money_(g(r,'disc')) + ')']);
    } else if (Number(g(r,'ded'))) {
      lines.push(['Deductions' + (remark ? ' (' + esc_(remark) + ')' : ''), '(' + money_(g(r,'ded')) + ')']);
    }

    let body = '';
    lines.forEach(l => {
      const cls = l[2] === 'sub' ? ' class="subrow"' : '';
      body += '<tr' + cls + '><td>' + esc_(l[0]) + '</td><td class="r">' + l[1] + '</td></tr>';
    });

    slips +=
      '<section class="page">' +
        '<div class="co">MAYA TEX RESTAURANTS COMPANY</div>' +
        '<div class="sub">Salary Slip — ' + esc_(periodLabel) + '</div>' +
        '<table class="meta"><tr>' +
          '<td><span class="lbl">Employee</span><br><b>' + esc_(g(r,'name')) + '</b></td>' +
          '<td><span class="lbl">Employee ID</span><br><b>' + esc_(g(r,'id')) + '</b></td>' +
        '</tr></table>' +
        '<table class="grid slip"><thead><tr><th>Description</th><th class="r">Amount (KD)</th></tr></thead>' +
          '<tbody>' + body +
          '<tr class="net"><td>NET PAYABLE</td><td class="r">' + money_(pay) + '</td></tr>' +
          '</tbody></table>' +
        '<div class="words"><i>' + esc_(words_(pay)) + '</i></div>' +
        (flags ? '<div class="flag">Note: ' + esc_(flags) + '</div>' : '') +
        '<table class="sig"><tr><td>______________________<br>Received (Employee)</td>' +
          '<td>______________________<br>Prepared (HR)</td>' +
          '<td>______________________<br>Approved</td></tr></table>' +
      '</section>';
  });

  const css =
    '@page { size: A5; margin: 10mm; }' +
    '* { font-family: Arial, Helvetica, sans-serif; box-sizing: border-box; }' +
    'body { margin: 0; color: #1a1a1a; font-size: 10px; }' +
    '.page { page-break-after: always; }' +
    '.page:last-child { page-break-after: auto; }' +
    '.co { text-align: center; font-size: 13px; font-weight: bold; color: #1e3a5f; }' +
    '.sub { text-align: center; font-size: 10px; color: #555; border-bottom: 1.5px solid #1e3a5f; padding-bottom: 5px; margin-bottom: 8px; }' +
    'table { width: 100%; border-collapse: collapse; }' +
    '.meta td { padding: 3px 0; font-size: 9px; }' +
    '.lbl { color: #777; font-size: 8px; }' +
    '.grid th { background: #1e3a5f; color: #fff; text-align: left; padding: 4px 6px; font-size: 9px; }' +
    '.grid td { padding: 4px 6px; border-bottom: 0.5px solid #ccc; }' +
    '.r { text-align: right; } .b { font-weight: bold; }' +
    '.subrow td { font-weight: bold; background: #f7f9fc; }' +
    '.net td { font-weight: bold; font-size: 11px; background: #f2f4f7; border-top: 1.5px solid #1e3a5f; }' +
    '.slip { margin-top: 6px; }' +
    '.words { margin-top: 6px; font-size: 9px; }' +
    '.flag { margin-top: 4px; font-size: 8px; color: #b00; }' +
    '.tot td { font-weight: bold; background: #f2f4f7; border-top: 1.5px solid #1e3a5f; }' +
    '.sig { margin-top: 22px; text-align: center; font-size: 8px; color: #555; }' +
    '.sig td { padding-top: 14px; }' +
    '.note { margin-top: 8px; font-size: 8px; color: #777; }';

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' +
         summaryPage + slips + '</body></html>';
}