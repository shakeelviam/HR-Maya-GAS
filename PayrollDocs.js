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
      remark: (c('Deduction Remark') !== -1 ? c('Deduction Remark') : c('Remark')), pay: c('Total Payable'), flags: c('Flags'),
      // New itemised review-tab columns (present when run from the Payroll Review screen)
      bonusAdd: c('Bonus/Additional'), unpaid: c('Unpaid Leave'), loan: c('Loan'),
      advance: c('Advance'), disc: c('Disciplinary'), sched: c('Scheduled Deductions'),
      totDed: c('Total Deductions'), net: c('Net Payable')
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

    // Build a temporary Google Doc with A5 pages (summary + one payslip per
    // page) and export it as the PDF. Apps Script's HTML→PDF converter
    // ignores @page size and always prints A4, so the page size is set on
    // the Doc itself, where it is exact.
    const pdfBlob = buildPayrollPdfA5_(tabName, periodLabel, emps, totals, idx)
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

// Run from the Apps Script editor (Run ▶ runPayslipsLatest) to make the
// payslip PDF for the most recent payroll run and see the result — or the
// real error — in the Execution log. Same as the Generate button.
function runPayslipsLatest() {
  const r = generatePayrollDocs({});
  Logger.log(r.success ? (r.message + '\n' + r.pdfUrl) : ('ERROR: ' + r.error));
  return r;
}

// ------------------------------------------------------------
// A5 PDF via a temporary Google Doc.
// A5 = 148 × 210 mm = 419.53 × 595.28 pt. Margins 10 mm (28.35 pt).
// ------------------------------------------------------------
const A5_W = 419.53, A5_H = 595.28, A5_M = 28.35;
const SLIP_NAVY = '#1e3a5f';

function buildPayrollPdfA5_(tabName, periodLabel, emps, totals, idx) {
  const g = (row, key) => (idx[key] !== -1 ? row[idx[key]] : '');
  const payOf = (row) => (idx.net !== -1 && g(row,'net') !== '') ? g(row,'net') : g(row,'pay');
  const dedOf = (row) => (idx.totDed !== -1 && g(row,'totDed') !== '') ? g(row,'totDed') : g(row,'ded');
  const A = DocumentApp.Attribute, H = DocumentApp.HorizontalAlignment;
  const txt = (el, size, bold, color, italic) => {
    const st = {}; st[A.FONT_FAMILY] = 'Arial'; st[A.FONT_SIZE] = size; st[A.BOLD] = !!bold;
    st[A.ITALIC] = !!italic; st[A.FOREGROUND_COLOR] = color || '#1a1a1a';
    const t = el.editAsText();
    if (t.getText().length) t.setAttributes(st);   // nothing to style in an empty cell
    return el;
  };
  const para = (body, text, size, bold, color, align, italic) => {
    const p = body.appendParagraph(text || '');
    p.setSpacingBefore(0); p.setSpacingAfter(2); p.setLineSpacing(1);
    if (align) p.setAlignment(align);
    if (text) txt(p, size, bold, color, italic);
    return p;
  };
  const styleCell = (cell, size, bold, color, bg, right) => {
    cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);
    if (bg) cell.setBackgroundColor(bg);
    txt(cell, size, bold, color);
    if (right) cell.getChild(0).asParagraph().setAlignment(H.RIGHT);
    const p = cell.getChild(0).asParagraph(); p.setSpacingBefore(0); p.setSpacingAfter(0);
  };
  const header = (body, title) => {
    para(body, 'MAYA TEX RESTAURANTS COMPANY', 12, true, SLIP_NAVY, H.CENTER);
    para(body, title, 9, false, '#555555', H.CENTER);
  };

  const doc = DocumentApp.create('tmp payslips ' + tabName + ' ' + new Date().getTime());
  const body = doc.getBody();
  body.setPageWidth(A5_W).setPageHeight(A5_H)
      .setMarginTop(A5_M).setMarginBottom(A5_M).setMarginLeft(A5_M).setMarginRight(A5_M);
  const inner = A5_W - 2 * A5_M;

  // --- Summary (flows over as many A5 pages as it needs) ---
  const first = body.getChild(0).asParagraph();
  first.setText('MAYA TEX RESTAURANTS COMPANY'); first.setAlignment(H.CENTER);
  first.setSpacingBefore(0); first.setSpacingAfter(2); txt(first, 12, true, SLIP_NAVY);
  para(body, 'Payroll Summary — ' + periodLabel, 9, false, '#555555', H.CENTER);
  const sRows = [['ID', 'Name', 'Gross', 'OT', 'Deduct', 'Net']];
  emps.forEach(r => sRows.push([String(g(r,'id')), String(g(r,'name')), money_(g(r,'gross')),
    money_(g(r,'otAmt')), money_(dedOf(r)), money_(payOf(r))]));
  sRows.push(['', 'TOTALS (' + emps.length + ' staff)', totals ? money_(g(totals,'gross')) : '',
    totals ? money_(g(totals,'otAmt')) : '', totals ? money_(dedOf(totals)) : '', totals ? money_(payOf(totals)) : '']);
  const st = body.appendTable(sRows);
  st.setBorderWidth(0.5).setBorderColor('#cccccc');
  [48, 136, 44, 40, 44, 50].forEach((w, j) => st.setColumnWidth(j, w * inner / 362));
  for (let i = 0; i < sRows.length; i++) {
    const hdr = i === 0, tot = i === sRows.length - 1;
    for (let j = 0; j < 6; j++) styleCell(st.getRow(i).getCell(j), 7, hdr || tot || j === 5,
      hdr ? '#ffffff' : null, hdr ? SLIP_NAVY : (tot ? '#f2f4f7' : null), j >= 2);
  }
  para(body, 'All amounts in Kuwaiti Dinar (KD), to 3 decimals. Generated ' + formatDdMmYyyy_(new Date()) + '.',
       7, false, '#777777');

  // --- One A5 page per employee ---
  emps.forEach(r => {
    body.appendPageBreak();
    const pay = Number(payOf(r)) || 0;
    const otHrs = g(r,'otHrs'), remark = g(r,'remark'), flags = g(r,'flags');
    const lines = [['Description', 'Amount (KD)', 'head']];
    lines.push(['Basic Salary', money_(g(r,'basic'))]);
    if (Number(g(r,'food')))  lines.push(['Food Allowance', money_(g(r,'food'))]);
    if (Number(g(r,'accom'))) lines.push(['Accommodation', money_(g(r,'accom'))]);
    if (Number(g(r,'conv')))  lines.push(['Conveyance', money_(g(r,'conv'))]);
    if (Number(g(r,'bonus'))) lines.push(['Bonus', money_(g(r,'bonus'))]);
    lines.push(['Gross Salary', money_(g(r,'gross')), 'sub']);
    if (Number(otHrs)) lines.push(['Overtime (' + otHrs + ' hrs @ ' + money_(g(r,'rate')) + ')', money_(g(r,'otAmt'))]);
    if (Number(g(r,'enc'))) lines.push(['Leave Encashment (' + (g(r,'encDays') || 0) + ' days)', money_(g(r,'enc'))]);
    if (idx.bonusAdd !== -1 && Number(g(r,'bonusAdd'))) lines.push(['Bonus / Additional', money_(g(r,'bonusAdd'))]);
    if (idx.totDed !== -1) {
      if (Number(g(r,'unpaid')))  lines.push(['Unpaid leave / absence', '(' + money_(g(r,'unpaid')) + ')']);
      if (Number(g(r,'loan')))    lines.push(['Loan', '(' + money_(g(r,'loan')) + ')']);
      if (Number(g(r,'advance'))) lines.push(['Advance', '(' + money_(g(r,'advance')) + ')']);
      if (Number(g(r,'disc')))    lines.push(['Disciplinary', '(' + money_(g(r,'disc')) + ')']);
      if (idx.sched !== -1 && Number(g(r,'sched'))) lines.push(['Scheduled deductions', '(' + money_(g(r,'sched')) + ')']);
    } else if (Number(g(r,'ded'))) {
      lines.push(['Deductions', '(' + money_(g(r,'ded')) + ')']);
    }
    lines.push(['NET PAYABLE', money_(pay), 'net']);

    // Fill the A5 page. A slip with few lines gets big type, a long one
    // smaller, so every slip uses the sheet instead of sitting in the top
    // third of it. Worst case (16 lines) still fits on one page.
    // Sizes measured against the A5 text area (539 pt): every band lands
    // between 82% and 93% of the page, so nothing spills to a second page
    // even when a long Details line wraps.
    const n = lines.length;
    const fs    = n <= 7 ? 16 : n <= 9 ? 14 : n <= 11 ? 13 : n <= 13 ? 12 : n <= 16 ? 11 : 10;
    const pad   = n <= 7 ?  9 : n <= 9 ?  8 : n <= 11 ?  7 : n <= 13 ?  6 : n <= 16 ?  5 :  4;
    const hdrFs = n <= 7 ? 20 : n <= 9 ? 18 : n <= 13 ? 17 : 15;

    const hp = body.appendParagraph('MAYA TEX RESTAURANTS COMPANY');
    hp.setAlignment(H.CENTER); hp.setSpacingBefore(0); hp.setSpacingAfter(2); txt(hp, hdrFs, true, SLIP_NAVY);
    para(body, 'Salary Slip — ' + periodLabel, fs, false, '#555555', H.CENTER);

    const meta = body.appendTable([['Employee', 'Employee ID'], [String(g(r,'name')), String(g(r,'id'))]]);
    meta.setBorderWidth(0);
    for (let i = 0; i < 2; i++) for (let j = 0; j < 2; j++) {
      const c = meta.getRow(i).getCell(j);
      styleCell(c, i ? fs + 1 : fs - 2, !!i, i ? null : '#777777', null, j === 1);
      c.setPaddingTop(i ? 1 : pad).setPaddingBottom(1);
    }

    const t = body.appendTable(lines.map(l => [l[0], l[1]]));
    t.setBorderWidth(0.5).setBorderColor('#cccccc');
    t.setColumnWidth(0, inner * 0.68); t.setColumnWidth(1, inner * 0.32);
    lines.forEach((l, i) => {
      const kind = l[2];
      for (let j = 0; j < 2; j++) {
        const c = t.getRow(i).getCell(j);
        styleCell(c, kind === 'net' ? fs + 2 : fs, !!kind, kind === 'head' ? '#ffffff' : null,
          kind === 'head' ? SLIP_NAVY : (kind === 'sub' ? '#f7f9fc' : (kind === 'net' ? '#f2f4f7' : null)), j === 1);
        c.setPaddingTop(pad).setPaddingBottom(pad);
      }
    });

    para(body, words_(pay), fs - 1, false, '#1a1a1a', null, true).setSpacingBefore(pad);
    if (remark) para(body, 'Details: ' + remark, fs - 2, false, '#555555');
    if (flags) para(body, 'Note: ' + flags, fs - 2, false, '#bb0000');
    // Push the signature block down the page so the slip ends at the foot
    // of the A5 sheet rather than floating in the middle of it.
    const gaps = n <= 5 ? 6 : n <= 7 ? 3 : n <= 9 ? 2 : n <= 11 ? 1 : 0;
    for (let i = 0; i < gaps; i++) para(body, '', fs);
    const sig = body.appendTable([['______________________', '______________________', '______________________'],
                                  ['Received (Employee)', 'Prepared (HR)', 'Approved']]);
    sig.setBorderWidth(0);
    for (let i = 0; i < 2; i++) for (let j = 0; j < 3; j++) {
      styleCell(sig.getRow(i).getCell(j), fs - 2, false, '#555555');
      sig.getRow(i).getCell(j).getChild(0).asParagraph().setAlignment(H.CENTER);
    }
  });

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  const pdf = file.getAs('application/pdf');
  file.setTrashed(true);   // the Doc was only a vehicle for the PDF
  return pdf;
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