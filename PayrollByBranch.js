// ============================================================
// PayrollByBranch.gs — the payroll run, grouped by branch, with OT split out
// ------------------------------------------------------------
// Builds a tab that reads like "Payroll <period>" but sorted into branches,
// with a subtotal after each one and a grand total at the foot — the view a
// manager actually wants: how much OT each branch ran, and what it cost.
//
// WHERE THE THREE OT COLUMNS COME FROM
//   OT Amount      the hours x the man's hourly rate, straight off the
//                  payroll tab.
//   OT Difference  a "+OT Difference n" top-up in that row's Remark. These
//                  are OT from an earlier month that came to under 1 KD and
//                  were carried into this run. They ARE overtime money, and
//                  leaving them out would understate what a branch was paid.
//   OT Paid        the two added: what actually went out as overtime.
//
// The branch comes from "Payroll by Location <period>", because that tab is
// what the bank transfer was built from — so the branch on this tab is the
// branch the money was actually paid under, not a guess from HR Maya.
//
// RUN IT: buildPayrollByBranch() from the editor, with no arguments. It
// picks the newest payroll period by itself and rebuilds the tab from
// scratch each time, so running it twice does no harm.
// ============================================================

const PBB_TAB_PREFIX = 'Payroll by Branch ';

// Every column of the payroll tab is carried over, in its own order, with
// Branch put in front and the OT columns widened from one to three.
const PBB_MONEY_COLS = [
  'Basic Salary', 'Food Allowance', 'Accommodation', 'Conveyance', 'Bonus', 'Gross Salary'
];
const PBB_TAIL_COLS = [
  'Encashment', 'Bonus/Additional', 'Unpaid Leave', 'Loan', 'Advance',
  'Disciplinary', 'Scheduled Deductions', 'Total Deductions', 'Net Payable'
];
// Columns that get summed on the subtotal rows.
const PBB_SUM_COLS = [
  'Basic Salary', 'Food Allowance', 'Accommodation', 'Conveyance', 'Bonus', 'Gross Salary',
  'OT Hours', 'OT Amount', 'OT Difference', 'OT Paid',
  'Encashment', 'Bonus/Additional', 'Unpaid Leave', 'Loan', 'Advance',
  'Disciplinary', 'Scheduled Deductions', 'Total Deductions', 'Net Payable'
];

function pbbNum_(v) {
  const n = Number(String(v == null ? '' : v).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
}
function pbbNorm_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }

// "Payroll 22Aug-20Sep" -> "22Aug-20Sep". The period is whatever follows the
// word Payroll, so a renamed period still works.
function pbbPeriodOf_(name) {
  const m = String(name).match(/^Payroll\s+(.+)$/i);
  if (!m) return null;
  const p = m[1].trim();
  // "Payroll by Location ..." and "Payroll_Inputs" are not payroll runs.
  if (/^by location/i.test(p) || /^_/.test(p)) return null;
  return p;
}

// The newest run is the one whose tab was created last; Apps Script gives no
// creation date, so the sheet order is used — new tabs are added at the end.
function pbbLatestPeriod_(ss) {
  const found = [];
  ss.getSheets().forEach(sh => {
    const p = pbbPeriodOf_(sh.getName());
    if (p && ss.getSheetByName('Payroll by Location ' + p)) found.push(p);
  });
  return found.length ? found[found.length - 1] : null;
}

function pbbRead_(sh) {
  const d = sh.getDataRange().getValues();
  return { head: d[0].map(h => String(h).trim()), rows: d.slice(1) };
}

// ------------------------------------------------------------
// buildPayrollByBranch({ period })  — period optional
// ------------------------------------------------------------
function buildPayrollByBranch(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    const period = String(params.period || '').trim() || pbbLatestPeriod_(ss);
    if (!period) {
      const msg = 'No payroll period found. There must be a "Payroll <period>" tab ' +
                  'and a matching "Payroll by Location <period>" tab.';
      Logger.log(msg); return { success: false, error: msg };
    }

    const paySheet = ss.getSheetByName('Payroll ' + period);
    const locSheet = ss.getSheetByName('Payroll by Location ' + period);
    if (!paySheet) return { success: false, error: 'No tab called "Payroll ' + period + '".' };
    if (!locSheet) return { success: false, error: 'No tab called "Payroll by Location ' + period + '".' };

    const pay = pbbRead_(paySheet), loc = pbbRead_(locSheet);
    const P = (n) => pay.head.indexOf(n);
    const cId = P('Employee ID'), cName = P('Name (English)'), cRemark = P('Remark');
    if (cId === -1 || cName === -1) {
      return { success: false, error: 'The payroll tab has no Employee ID / Name (English) column.' };
    }

    // Branch and payment method, keyed by name — the by-location tab carries
    // no Employee ID, so the name is the only join there is.
    const L = (n) => loc.head.indexOf(n);
    const lName = L('Employee Name'), lLoc = L('Location'), lPay = L('Payment');
    // The by-location tab ends its list of people with a blank row and then
    // carries a per-branch summary block (Location | Staff | Bank | Cash |
    // Total). Reading straight through would turn "Ardiya", "TOTAL" and the
    // rest of that block into employees, so the list stops at the blank row.
    const locPeople = [];
    for (let i = 0; i < loc.rows.length; i++) {
      const nm = String(loc.rows[i][lName] || '').trim();
      if (!nm) break;                                   // the summary block starts here
      locPeople.push(loc.rows[i]);
    }
    const byName = {};
    locPeople.forEach(r => {
      const nm = pbbNorm_(r[lName]);
      if (nm) byName[nm] = { branch: String(r[lLoc] || '').trim() || '(no branch)',
                             payment: String(r[lPay] || '').trim() };
    });

    const HEAD = ['Branch', 'Employee ID', 'Name (English)']
      .concat(PBB_MONEY_COLS)
      .concat(['OT Hours', 'Price Per Hour', 'OT Amount', 'OT Difference', 'OT Paid'])
      .concat(PBB_TAIL_COLS)
      .concat(['Payment', 'Remark']);
    const H = (n) => HEAD.indexOf(n);

    const staff = [];
    const unmatched = [];
    pay.rows.forEach(r => {
      const id = String(r[cId] || '').trim();
      const name = String(r[cName] || '').trim();
      if (!id || /^totals/i.test(name)) return;          // the payroll tab's own total row

      const info = byName[pbbNorm_(name)];
      if (!info) unmatched.push(id + ' ' + name);
      const branch = info ? info.branch : '(no branch)';

      const remark = String(r[cRemark] || '');
      const m = remark.match(/\+\s*OT Difference\s+([\d.]+)/i);
      const otDiff = m ? pbbNum_(m[1]) : 0;
      const otAmt  = pbbNum_(r[P('OT Amount')]);

      const cells = new Array(HEAD.length).fill('');
      cells[0] = branch; cells[1] = id; cells[2] = name;
      PBB_MONEY_COLS.forEach(k => { if (P(k) !== -1) cells[H(k)] = pbbNum_(r[P(k)]); });
      cells[H('OT Hours')]      = pbbNum_(r[P('OT Hours')]);
      cells[H('Price Per Hour')]= pbbNum_(r[P('Price Per Hour')]);
      cells[H('OT Amount')]     = otAmt;
      cells[H('OT Difference')] = otDiff;
      cells[H('OT Paid')]       = otAmt + otDiff;
      PBB_TAIL_COLS.forEach(k => { if (P(k) !== -1) cells[H(k)] = pbbNum_(r[P(k)]); });
      cells[H('Payment')] = info ? info.payment : '';
      cells[H('Remark')]  = remark.trim();
      staff.push({ branch: branch, name: name, cells: cells });
    });

    if (!staff.length) return { success: false, error: 'No staff rows on "Payroll ' + period + '".' };

    // Anyone on the by-location tab who was not in the run at all. They are
    // not invented into the tab — they are reported, so nobody assumes the
    // grand total is the whole wage bill.
    const inRun = {};
    staff.forEach(s => inRun[pbbNorm_(s.name)] = true);
    const notInRun = locPeople
      .filter(r => String(r[lName] || '').trim() && !inRun[pbbNorm_(r[lName])])
      .map(r => String(r[lName]).trim() + ' (' + String(r[lLoc] || '').trim() + ')');

    // ── assemble ──────────────────────────────────────────────────────
    const totalRow = (label, list) => {
      const out = new Array(HEAD.length).fill('');
      out[0] = label;
      out[2] = list.length + (list.length === 1 ? ' man' : ' staff');
      PBB_SUM_COLS.forEach(k => {
        const i = H(k); if (i === -1) return;
        out[i] = list.reduce((a, x) => a + pbbNum_(x.cells[i]), 0);
      });
      return out;
    };

    const body = [];
    body.push(['Payroll by Branch — period ' + period]);
    body.push(['OT Amount = hours x rate.  OT Difference = "+OT Difference" top-ups carried into this run.  ' +
               'OT Paid = OT Amount + OT Difference.']);
    body.push(['Built ' + Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm') +
               ' from "Payroll ' + period + '" and "Payroll by Location ' + period + '".']);
    body.push([]);
    const headerRow = body.length;                      // 0-based index of HEAD
    body.push(HEAD);

    const branches = {};
    staff.forEach(s => { (branches[s.branch] = branches[s.branch] || []).push(s); });
    const order = Object.keys(branches).sort();
    const subtotalRows = [], branchStats = [];
    order.forEach(b => {
      branches[b].sort((x, y) => x.name.localeCompare(y.name)).forEach(s => body.push(s.cells));
      subtotalRows.push(body.length);
      const t = totalRow(b + ' — total', branches[b]);
      body.push(t);
      branchStats.push({ branch: b, staff: branches[b].length,
                         otHours: t[H('OT Hours')], otAmount: t[H('OT Amount')],
                         otDifference: t[H('OT Difference')], otPaid: t[H('OT Paid')],
                         netPayable: t[H('Net Payable')] });
      body.push([]);
    });
    const grandIdx = body.length;
    const grand = totalRow('ALL BRANCHES', staff);
    body.push(grand);
    if (notInRun.length) {
      body.push([]);
      body.push(['Not in this payroll run, so not counted above: ' + notInRun.join(', ')]);
    }

    // ── write ─────────────────────────────────────────────────────────
    const tabName = PBB_TAB_PREFIX + period;
    let sh = ss.getSheetByName(tabName);
    if (sh) sh.clear(); else sh = ss.insertSheet(tabName);

    const width = HEAD.length;
    const grid = body.map(r => { const c = r.slice(); while (c.length < width) c.push(''); return c; });
    sh.getRange(1, 1, grid.length, width).setValues(grid);

    // Formatting: the title, the header and every total row stand out, and
    // money reads as money. Three decimals, like the rest of the workbook.
    sh.getRange(1, 1, 1, width).merge().setFontWeight('bold').setFontSize(13);
    sh.getRange(2, 1, 2, width).setFontColor('#666666').setFontSize(9);
    sh.getRange(headerRow + 1, 1, 1, width).setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(headerRow + 1);
    sh.setFrozenColumns(3);

    const firstMoney = H('Basic Salary') + 1, lastMoney = H('Net Payable') + 1;
    sh.getRange(headerRow + 2, firstMoney, grid.length - headerRow - 1, lastMoney - firstMoney + 1)
      .setNumberFormat('#,##0.000');
    sh.getRange(headerRow + 2, H('OT Hours') + 1, grid.length - headerRow - 1, 1).setNumberFormat('#,##0');

    subtotalRows.forEach(i => sh.getRange(i + 1, 1, 1, width)
      .setFontWeight('bold').setBackground('#f1f3f4'));
    sh.getRange(grandIdx + 1, 1, 1, width)
      .setFontWeight('bold').setBackground('#d9ead3').setBorder(true, true, true, true, false, false);

    sh.autoResizeColumns(1, Math.min(width, 26));
    SpreadsheetApp.flush();

    const message = 'Built "' + tabName + '" — ' + staff.length + ' staff across ' +
      order.length + ' branches. OT paid ' + (Math.round(grand[H('OT Paid')] * 1000) / 1000).toFixed(3) +
      ' KD, net payable ' + (Math.round(grand[H('Net Payable')] * 1000) / 1000).toFixed(3) + ' KD.';
    Logger.log(message);
    if (unmatched.length) Logger.log('No branch found for: ' + unmatched.join(', '));
    if (notInRun.length) Logger.log('On the by-location tab but not in the run: ' + notInRun.join(', '));

    return { success: true, tab: tabName, period: period, staff: staff.length,
             branches: branchStats, unmatched: unmatched, notInRun: notInRun,
             otPaid: grand[H('OT Paid')], netPayable: grand[H('Net Payable')],
             message: message };
  } catch (err) { return { success: false, error: err.toString() }; }
}