import * as XLSX from 'xlsx';
import type { LedgerAnalysis } from './analyzeLedger';
import type { BalanceSheet } from './generateBalanceSheet';
import type { CashFlowStatement } from './generateCashFlow';
import type { ProfitAndLoss } from './generatePL';
import type { MoMPL } from './generateMoMPL';
import { monthLabel } from './generateMoMPL';

const fmt = (n: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n);

// ─── Styles ──────────────────────────────────────────────────────────────────

const BLUE   = '2F5496';
const GREEN  = '1D6F42';
const RED    = 'C00000';
const GOLD   = 'BF8F00';
const LGRAY  = 'F2F2F2';
const DGRAY  = 'D9D9D9';
const WHITE  = 'FFFFFF';
const BLACK  = '000000';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type CS = any;

const hdr = (bgHex: string, fgHex = WHITE, bold = true, sz = 11): CS => ({
  font: { bold, color: { rgb: fgHex }, sz, name: 'Calibri' },
  fill: { fgColor: { rgb: bgHex }, patternType: 'solid' },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: { bottom: { style: 'thin', color: { rgb: DGRAY } } },
});

const cell = (bold = false, color = BLACK, align: 'left'|'right'|'center' = 'left', bg?: string): CS => ({
  font: { bold, color: { rgb: color }, sz: 10, name: 'Calibri' },
  fill: bg ? { fgColor: { rgb: bg }, patternType: 'solid' } : { patternType: 'none' },
  alignment: { horizontal: align, vertical: 'center' },
  border: { bottom: { style: 'hair', color: { rgb: DGRAY } } },
});

const money = (bold = false, color = BLACK, bg?: string): CS => ({
  ...cell(bold, color, 'right', bg),
  numFmt: '"$"#,##0.00;[Red]"$"(#,##0.00)',
});

const pct: CS = {
  font: { sz: 10, name: 'Calibri' },
  alignment: { horizontal: 'right', vertical: 'center' },
  numFmt: '+0.0%;-0.0%;0.0%',
};

const setStyle = (ws: XLSX.WorkSheet, addr: string, style: CS) => {
  if (!ws[addr]) ws[addr] = { t: 'z', v: '' };
  ws[addr].s = style;
};

const applyRow = (ws: XLSX.WorkSheet, row: number, cols: string[], styles: CS[]) => {
  cols.forEach((c, i) => setStyle(ws, `${c}${row}`, styles[i] ?? styles[styles.length - 1]));
};

const colWidth = (w: number) => ({ wch: w });

// ─── CSV ─────────────────────────────────────────────────────────────────────

function toCsvSection(title: string, headers: string[], rows: (string | number)[][]): string {
  const esc = (v: string | number) => { const s = String(v); return s.includes(',') || s.includes('"') ? `"${s.replace(/"/g,'""')}"` : s; };
  return [title, headers.map(esc).join(','), ...rows.map(r => r.map(esc).join(',')), ''].join('\n');
}

export function exportCsv(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const sections = [
    toCsvSection('LEDGER ANALYSIS', ['Metric','Value'], [
      ['Total Transactions', analysis.totalTransactions],
      ['Inconsistent Vendors', analysis.inconsistentVendors.length],
      ['Duplicate Transactions', analysis.duplicates.length],
    ]),
    toCsvSection('INCONSISTENT VENDORS', ['Vendor','Accounts','Reason'], analysis.inconsistentVendors.map(({ vendor, accounts, reason }) => [vendor, accounts.join(' | '), reason])),
    toCsvSection('DUPLICATES', ['Vendor','Amount','Date','Occurrences'], analysis.duplicates.map(({ name, amount, transactionDate, occurrences }) => [name, amount, transactionDate, occurrences])),
    toCsvSection('P&L', ['Item','Amount'], [['Total Revenue', fmt(pl.totalRevenue)], ['Total Expenses', fmt(pl.totalExpenses)], ['Net Profit', fmt(pl.netProfit)]]),
    toCsvSection('MONTHLY P&L', ['Month','Revenue','Expenses','Net'], Object.entries(pl.monthlyBreakdown).map(([m,{revenue,expenses}]) => [m, fmt(revenue), fmt(expenses), fmt(revenue-expenses)])),
    toCsvSection('BALANCE SHEET', ['Category','Account','Balance'], [
      ...bs.assets.map(({account,value}) => ['Asset', account, fmt(value)]),
      ...bs.liabilities.map(({account,value}) => ['Liability', account, fmt(value)]),
      ...bs.equity.map(({account,value}) => ['Equity', account, fmt(value)]),
    ]),
    toCsvSection('CASH FLOW', ['Item','Amount'], [['Net Profit', fmt(cf.netProfit)], ['Operating Cash Flow', fmt(cf.operatingCashFlow)]]),
  ];
  triggerDownload(new Blob([sections.join('\n')], { type: 'text/csv;charset=utf-8;' }), `${baseName(fileName)}_analysis.csv`);
}

// ─── Excel ───────────────────────────────────────────────────────────────────

export function exportExcel(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement, mom?: MoMPL): void {
  const wb = XLSX.utils.book_new();

  // ── 1. Summary ──────────────────────────────────────────────────────────────
  {
    const ws: XLSX.WorkSheet = { '!ref': 'A1:C20' };

    const writeCell = (addr: string, v: string | number, t: string, s: CS) => { ws[addr] = { v, t, s }; };

    // Title
    ws['A1'] = { v: `Ledger Analysis — ${baseName(fileName)}`, t: 's', s: { font: { bold: true, sz: 16, color: { rgb: WHITE }, name: 'Calibri' }, fill: { fgColor: { rgb: BLUE }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' } } };
    ws['A2'] = { v: `Generated: ${new Date().toLocaleString()}`, t: 's', s: { font: { sz: 9, color: { rgb: '666666' }, name: 'Calibri' }, alignment: { horizontal: 'left' } } };

    // KPI section
    const kpiHdr = hdr(BLUE);
    const kpiVal = { font: { bold: true, sz: 20, name: 'Calibri' }, alignment: { horizontal: 'center', vertical: 'center' } };
    const kpiLbl = { font: { sz: 9, color: { rgb: '666666' }, name: 'Calibri' }, alignment: { horizontal: 'center' } };

    ws['A4'] = { v: 'TOTAL TRANSACTIONS', t: 's', s: kpiHdr };
    ws['B4'] = { v: 'TOTAL REVENUE', t: 's', s: kpiHdr };
    ws['C4'] = { v: 'TOTAL EXPENSES', t: 's', s: kpiHdr };
    ws['D4'] = { v: 'NET PROFIT', t: 's', s: kpiHdr };
    ws['E4'] = { v: 'BS BALANCED?', t: 's', s: kpiHdr };

    ws['A5'] = { v: analysis.totalTransactions, t: 'n', s: kpiVal };
    ws['B5'] = { v: pl.totalRevenue, t: 'n', s: { ...kpiVal, numFmt: '"$"#,##0.00' } };
    ws['C5'] = { v: pl.totalExpenses, t: 'n', s: { ...kpiVal, numFmt: '"$"#,##0.00' } };
    ws['D5'] = { v: pl.netProfit, t: 'n', s: { ...kpiVal, numFmt: '"$"#,##0.00', font: { bold: true, sz: 20, name: 'Calibri', color: { rgb: pl.netProfit >= 0 ? GREEN : RED } } } };
    ws['E5'] = { v: bs.isBalanced ? '✅ YES' : '❌ NO', t: 's', s: { ...kpiVal, font: { bold: true, sz: 16, name: 'Calibri', color: { rgb: bs.isBalanced ? GREEN : RED } } } };

    ws['A6'] = { v: 'Transactions', t: 's', s: kpiLbl };
    ws['B6'] = { v: 'Revenue', t: 's', s: kpiLbl };
    ws['C6'] = { v: 'Expenses', t: 's', s: kpiLbl };
    ws['D6'] = { v: 'Net Profit', t: 's', s: kpiLbl };
    ws['E6'] = { v: 'Balance Sheet', t: 's', s: kpiLbl };

    // BS reconciliation
    ws['A8'] = { v: 'BALANCE SHEET RECONCILIATION', t: 's', s: hdr(BLUE) };
    ws['B8'] = { v: '', t: 's', s: hdr(BLUE) };
    ws['C8'] = { v: '', t: 's', s: hdr(BLUE) };

    ws['A9']  = { v: 'Total Assets', t: 's', s: cell(true) };
    ws['B9']  = { v: bs.totals.assetsTotal, t: 'n', s: money(true) };
    ws['A10'] = { v: 'Total Liabilities', t: 's', s: cell() };
    ws['B10'] = { v: bs.totals.liabilitiesTotal, t: 'n', s: money() };
    ws['A11'] = { v: 'Total Equity', t: 's', s: cell() };
    ws['B11'] = { v: bs.totals.equityTotal, t: 'n', s: money() };
    ws['A12'] = { v: 'Liabilities + Equity', t: 's', s: cell(true) };
    ws['B12'] = { v: bs.totals.liabilitiesTotal + bs.totals.equityTotal, t: 'n', s: money(true) };
    ws['C12'] = { v: '=B10+B11', t: 'n', s: { ...money(), font: { color: { rgb: '888888' }, sz: 9, name: 'Calibri', italic: true } } };
    ws['A13'] = { v: 'Variance (A − L − E)', t: 's', s: cell(true, bs.isBalanced ? GREEN : RED) };
    ws['B13'] = { v: bs.variance, t: 'n', s: money(true, bs.isBalanced ? GREEN : RED) };
    ws['C13'] = { v: '=B9-B12', t: 'n', s: { ...money(), font: { color: { rgb: '888888' }, sz: 9, name: 'Calibri', italic: true } } };
    ws['A14'] = { v: 'Status', t: 's', s: cell(true) };
    ws['B14'] = { v: bs.isBalanced ? 'BALANCED ✅' : `NOT BALANCED ❌ (Δ ${fmt(bs.variance)})`, t: 's', s: { ...cell(true, bs.isBalanced ? GREEN : RED), fill: { fgColor: { rgb: bs.isBalanced ? 'E2EFDA' : 'FFDCDC' }, patternType: 'solid' } } };

    // Flags
    ws['A16'] = { v: 'FLAGS', t: 's', s: hdr(GOLD, WHITE) };
    ws['B16'] = { v: '', t: 's', s: hdr(GOLD) };
    ws['A17'] = { v: 'Inconsistent Vendors', t: 's', s: cell() };
    ws['B17'] = { v: analysis.inconsistentVendors.length, t: 'n', s: { ...cell(true, analysis.inconsistentVendors.length > 0 ? RED : GREEN), alignment: { horizontal: 'center' } } };
    ws['A18'] = { v: 'Duplicate Transactions', t: 's', s: cell() };
    ws['B18'] = { v: analysis.duplicates.length, t: 'n', s: { ...cell(true, analysis.duplicates.length > 0 ? RED : GREEN), alignment: { horizontal: 'center' } } };

    ws['!cols'] = [colWidth(28), colWidth(18), colWidth(18), colWidth(18), colWidth(16)];
    ws['!rows'] = [{ hpt: 30 }, { hpt: 16 }, {}, { hpt: 22 }, { hpt: 36 }, { hpt: 16 }];
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
      { s: { r: 7, c: 0 }, e: { r: 7, c: 2 } },
      { s: { r: 15, c: 0 }, e: { r: 15, c: 2 } },
    ];
    ws['!ref'] = 'A1:E18';
    XLSX.utils.book_append_sheet(wb, ws, '📊 Summary');
  }

  // ── 2. P&L ──────────────────────────────────────────────────────────────────
  {
    const rows: (string | number)[][] = [];
    rows.push(['PROFIT & LOSS STATEMENT', '', '']);
    rows.push([`Source: ${baseName(fileName)}`, '', '']);
    rows.push([]);
    rows.push(['', 'Amount', '% of Revenue']);
    rows.push(['REVENUE', '', '']);
    rows.push(['  Total Revenue', pl.totalRevenue, 1]);
    rows.push([]);
    rows.push(['EXPENSES', '', '']);
    rows.push(['  Total Expenses', pl.totalExpenses, pl.totalRevenue ? pl.totalExpenses / pl.totalRevenue : 0]);
    rows.push([]);
    rows.push(['NET PROFIT', pl.netProfit, pl.totalRevenue ? pl.netProfit / pl.totalRevenue : 0]);
    rows.push([]);
    rows.push(['MONTHLY BREAKDOWN', '', '', '']);
    rows.push(['Month', 'Revenue', 'Expenses', 'Net Profit', 'Margin %']);

    const monthEntries = Object.entries(pl.monthlyBreakdown).sort();
    monthEntries.forEach(([month, { revenue, expenses }]) => {
      const net = revenue - expenses;
      rows.push([month, revenue, expenses, net, revenue ? net / revenue : 0]);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // Style header rows
    ['A1','B1','C1','D1'].forEach(a => setStyle(ws, a, hdr(BLUE)));
    ['A4','B4','C4'].forEach(a => setStyle(ws, a, hdr(DGRAY, BLACK)));
    ['A5'].forEach(a => setStyle(ws, a, hdr(BLUE)));
    ['A8'].forEach(a => setStyle(ws, a, hdr(BLUE)));
    setStyle(ws, 'A11', { ...cell(true), fill: { fgColor: { rgb: pl.netProfit >= 0 ? 'E2EFDA' : 'FFDCDC' }, patternType: 'solid' } });
    setStyle(ws, 'B11', { ...money(true, pl.netProfit >= 0 ? GREEN : RED), fill: { fgColor: { rgb: pl.netProfit >= 0 ? 'E2EFDA' : 'FFDCDC' }, patternType: 'solid' } });
    setStyle(ws, 'C11', { ...pct, font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: pl.netProfit >= 0 ? GREEN : RED } }, fill: { fgColor: { rgb: pl.netProfit >= 0 ? 'E2EFDA' : 'FFDCDC' }, patternType: 'solid' } });

    // Money format for revenue/expense rows
    ['B6','B9','B11'].forEach(a => { if (ws[a]) ws[a].s = { ...money(true), ...(ws[a].s ?? {}) }; ws[a] && (ws[a].s.numFmt = '"$"#,##0.00'); });

    // Monthly table header
    const mHdrRow = 14;
    ['A','B','C','D','E'].forEach(c => setStyle(ws, `${c}${mHdrRow}`, hdr(BLUE)));

    // Monthly data rows
    monthEntries.forEach((_, i) => {
      const r = mHdrRow + 1 + i;
      const bg = i % 2 === 0 ? WHITE : LGRAY;
      setStyle(ws, `A${r}`, cell(false, BLACK, 'left', bg));
      setStyle(ws, `B${r}`, { ...money(false, BLACK, bg) });
      setStyle(ws, `C${r}`, { ...money(false, BLACK, bg) });
      const net = monthEntries[i][1].revenue - monthEntries[i][1].expenses;
      setStyle(ws, `D${r}`, { ...money(false, net >= 0 ? GREEN : RED, bg) });
      setStyle(ws, `E${r}`, { ...pct, fill: { fgColor: { rgb: bg }, patternType: 'solid' } });
    });

    ws['!cols'] = [colWidth(28), colWidth(16), colWidth(16), colWidth(16), colWidth(12)];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } }];
    ws['!freeze'] = { xSplit: 0, ySplit: mHdrRow };
    XLSX.utils.book_append_sheet(wb, ws, '📈 P&L');
  }

  // ── 3. Balance Sheet ─────────────────────────────────────────────────────────
  {
    const aoa: (string | number)[][] = [];
    aoa.push(['BALANCE SHEET', '', '']);
    aoa.push([`Source: ${baseName(fileName)}`, '', '']);
    aoa.push([]);
    aoa.push(['ASSETS', '', '']);
    aoa.push(['Account', 'Balance', '% of Total']);
    bs.assets.forEach(({ account, value }) => aoa.push([account, value, bs.totals.assetsTotal ? value / bs.totals.assetsTotal : 0]));
    aoa.push(['Total Assets', bs.totals.assetsTotal, 1]);
    aoa.push([]);
    aoa.push(['LIABILITIES', '', '']);
    aoa.push(['Account', 'Balance', '% of Total']);
    bs.liabilities.forEach(({ account, value }) => aoa.push([account, value, bs.totals.liabilitiesTotal ? value / bs.totals.liabilitiesTotal : 0]));
    aoa.push(['Total Liabilities', bs.totals.liabilitiesTotal, 1]);
    aoa.push([]);
    aoa.push(['EQUITY', '', '']);
    aoa.push(['Account', 'Balance', '% of Total']);
    bs.equity.forEach(({ account, value }) => aoa.push([account, value, bs.totals.equityTotal ? value / bs.totals.equityTotal : 0]));
    aoa.push(['Total Equity', bs.totals.equityTotal, 1]);
    aoa.push([]);
    aoa.push(['RECONCILIATION', '', '']);
    aoa.push(['Assets', bs.totals.assetsTotal, '']);
    aoa.push(['Liabilities + Equity', bs.totals.liabilitiesTotal + bs.totals.equityTotal, '']);
    aoa.push(['Variance', bs.variance, '']);
    aoa.push(['Status', bs.isBalanced ? 'BALANCED ✅' : 'NOT BALANCED ❌', '']);

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Style section headers and total rows
    let r = 1;
    const styleSection = (row: number) => ['A','B','C'].forEach(c => setStyle(ws, `${c}${row}`, hdr(BLUE)));
    const styleColHdr = (row: number) => ['A','B','C'].forEach(c => setStyle(ws, `${c}${row}`, hdr(DGRAY, BLACK)));
    const styleTotal = (row: number, color = GREEN) => {
      setStyle(ws, `A${row}`, cell(true, color));
      setStyle(ws, `B${row}`, money(true, color));
      setStyle(ws, `C${row}`, { ...pct, font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: color } } });
    };

    styleSection(1); styleSection(2);
    styleSection(4); styleColHdr(5);
    const aEnd = 5 + bs.assets.length;
    for (let i = 6; i <= aEnd; i++) { setStyle(ws, `A${i}`, cell(false, BLACK, 'left', i%2===0?WHITE:LGRAY)); setStyle(ws, `B${i}`, money(false, BLACK, i%2===0?WHITE:LGRAY)); setStyle(ws, `C${i}`, { ...pct, fill: { fgColor: { rgb: i%2===0?WHITE:LGRAY }, patternType: 'solid' } }); }
    styleTotal(aEnd + 1);

    const lStart = aEnd + 3;
    styleSection(lStart); styleColHdr(lStart + 1);
    const lEnd = lStart + 1 + bs.liabilities.length;
    for (let i = lStart + 2; i <= lEnd; i++) { setStyle(ws, `A${i}`, cell(false, BLACK, 'left', i%2===0?WHITE:LGRAY)); setStyle(ws, `B${i}`, money(false, BLACK, i%2===0?WHITE:LGRAY)); setStyle(ws, `C${i}`, { ...pct, fill: { fgColor: { rgb: i%2===0?WHITE:LGRAY }, patternType: 'solid' } }); }
    styleTotal(lEnd + 1, RED);

    const eStart = lEnd + 3;
    styleSection(eStart); styleColHdr(eStart + 1);
    const eEnd = eStart + 1 + bs.equity.length;
    for (let i = eStart + 2; i <= eEnd; i++) { setStyle(ws, `A${i}`, cell(false, BLACK, 'left', i%2===0?WHITE:LGRAY)); setStyle(ws, `B${i}`, money(false, BLACK, i%2===0?WHITE:LGRAY)); setStyle(ws, `C${i}`, { ...pct, fill: { fgColor: { rgb: i%2===0?WHITE:LGRAY }, patternType: 'solid' } }); }
    styleTotal(eEnd + 1, BLUE);

    const recStart = eEnd + 3;
    styleSection(recStart);
    setStyle(ws, `A${recStart+1}`, cell(true)); setStyle(ws, `B${recStart+1}`, money(true));
    setStyle(ws, `A${recStart+2}`, cell(true)); setStyle(ws, `B${recStart+2}`, money(true));
    setStyle(ws, `A${recStart+3}`, cell(true, bs.isBalanced ? GREEN : RED)); setStyle(ws, `B${recStart+3}`, money(true, bs.isBalanced ? GREEN : RED, bs.isBalanced ? 'E2EFDA' : 'FFDCDC'));
    setStyle(ws, `A${recStart+4}`, cell(true, bs.isBalanced ? GREEN : RED, 'left', bs.isBalanced ? 'E2EFDA' : 'FFDCDC')); setStyle(ws, `B${recStart+4}`, cell(true, bs.isBalanced ? GREEN : RED, 'left', bs.isBalanced ? 'E2EFDA' : 'FFDCDC'));

    ws['!cols'] = [colWidth(34), colWidth(18), colWidth(12)];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } }];
    XLSX.utils.book_append_sheet(wb, ws, '⚖️ Balance Sheet');
  }

  // ── 4. Cash Flow ─────────────────────────────────────────────────────────────
  {
    const aoa: (string | number)[][] = [
      ['CASH FLOW STATEMENT', ''],
      [`Source: ${baseName(fileName)}`, ''],
      [],
      ['OPERATING ACTIVITIES', ''],
      ['Net Profit', cf.netProfit],
      ['Adjustments to reconcile Net Profit:', ''],
      ...cf.adjustments.map(({ account, impact }) => [`  ${account}`, impact]),
      [],
      ['NET OPERATING CASH FLOW', cf.operatingCashFlow],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    setStyle(ws, 'A1', hdr(BLUE)); setStyle(ws, 'B1', hdr(BLUE));
    setStyle(ws, 'A2', hdr(BLUE)); setStyle(ws, 'B2', hdr(BLUE));
    setStyle(ws, 'A4', hdr(GREEN)); setStyle(ws, 'B4', hdr(GREEN));
    setStyle(ws, 'A5', cell(false)); setStyle(ws, 'B5', money());
    setStyle(ws, 'A6', cell(false, '666666'));
    cf.adjustments.forEach((_, i) => {
      const r = 7 + i;
      setStyle(ws, `A${r}`, cell(false, BLACK, 'left', i%2===0?WHITE:LGRAY));
      setStyle(ws, `B${r}`, money(false, BLACK, i%2===0?WHITE:LGRAY));
    });
    const lastR = 7 + cf.adjustments.length + 1;
    setStyle(ws, `A${lastR}`, { ...cell(true, WHITE), fill: { fgColor: { rgb: cf.operatingCashFlow >= 0 ? GREEN : RED }, patternType: 'solid' } });
    setStyle(ws, `B${lastR}`, { ...money(true, WHITE), fill: { fgColor: { rgb: cf.operatingCashFlow >= 0 ? GREEN : RED }, patternType: 'solid' } });
    ws['!cols'] = [colWidth(38), colWidth(18)];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }];
    XLSX.utils.book_append_sheet(wb, ws, '💵 Cash Flow');
  }

  // ── 5. Month-over-Month P&L ──────────────────────────────────────────────────
  if (mom && mom.months.length > 0) {
    const months = mom.months;
    const cols = ['Account', ...months.map(m => monthLabel(m)), 'Total', ...months.slice(1).map(m => `MoM% ${monthLabel(m)}`)];

    const incRows = mom.incomeCategories.map(cat => [
      cat.name,
      ...months.map(m => cat.months[m] ?? 0),
      cat.total,
      ...months.slice(1).map((m, i) => {
        const cur = cat.months[m] ?? 0;
        const prev = cat.months[months[i]] ?? 0;
        return prev !== 0 ? (cur - prev) / Math.abs(prev) : '';
      }),
    ]);

    const expRows = mom.expenseCategories.map(cat => [
      cat.name,
      ...months.map(m => cat.months[m] ?? 0),
      cat.total,
      ...months.slice(1).map((m, i) => {
        const cur = cat.months[m] ?? 0;
        const prev = cat.months[months[i]] ?? 0;
        return prev !== 0 ? (cur - prev) / Math.abs(prev) : '';
      }),
    ]);

    const revTotalRow = ['Total Revenue', ...months.map(m => mom.monthlyRevenue[m] ?? 0), mom.totalRevenue, ...months.slice(1).map((m, i) => { const c = mom.monthlyRevenue[m]??0; const p = mom.monthlyRevenue[months[i]]??0; return p!==0?(c-p)/Math.abs(p):''; })];
    const expTotalRow = ['Total Expenses', ...months.map(m => mom.monthlyExpenses[m] ?? 0), mom.totalExpenses, ...months.slice(1).map((m, i) => { const c = mom.monthlyExpenses[m]??0; const p = mom.monthlyExpenses[months[i]]??0; return p!==0?(c-p)/Math.abs(p):''; })];
    const netRow = ['Net Profit', ...months.map(m => mom.monthlyNetProfit[m] ?? 0), mom.totalNetProfit, ...months.slice(1).map((m, i) => { const c = mom.monthlyNetProfit[m]??0; const p = mom.monthlyNetProfit[months[i]]??0; return p!==0?(c-p)/Math.abs(p):''; })];

    const aoa: (string | number | '')[][] = [
      ['MONTH-OVER-MONTH P&L', ...Array(cols.length - 1).fill('')],
      [`Source: ${baseName(fileName)}`, ...Array(cols.length - 1).fill('')],
      [],
      cols,
      ['── INCOME ──', ...Array(cols.length - 1).fill('')],
      ...incRows,
      revTotalRow,
      [],
      ['── EXPENSES ──', ...Array(cols.length - 1).fill('')],
      ...expRows,
      expTotalRow,
      [],
      netRow,
    ];

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const numCols = cols.length;
    const colLetters = Array.from({ length: numCols }, (_, i) => String.fromCharCode(65 + i));

    // Title rows
    colLetters.forEach(c => { setStyle(ws, `${c}1`, hdr(BLUE)); setStyle(ws, `${c}2`, hdr(BLUE)); });

    // Header row (row 4)
    colLetters.forEach((c, i) => {
      const isMoneyCol = i > 0 && i <= months.length + 1;
      const isPctCol = i > months.length + 1;
      setStyle(ws, `${c}4`, hdr(BLUE));
      if (isPctCol) setStyle(ws, `${c}4`, { ...hdr(GOLD), alignment: { horizontal: 'center', vertical: 'center', wrapText: true } });
    });

    // Section label rows
    const incLabelRow = 5; const expLabelRowOffset = incRows.length + 3;
    colLetters.forEach(c => { setStyle(ws, `${c}${incLabelRow}`, hdr(LGRAY, BLACK, true, 10)); });

    // Income rows
    incRows.forEach((_, i) => {
      const r = incLabelRow + 1 + i;
      const bg = i % 2 === 0 ? WHITE : LGRAY;
      colLetters.forEach((c, ci) => {
        if (ci === 0) setStyle(ws, `${c}${r}`, cell(false, BLACK, 'left', bg));
        else if (ci <= months.length + 1) setStyle(ws, `${c}${r}`, money(false, BLACK, bg));
        else setStyle(ws, `${c}${r}`, { ...pct, fill: { fgColor: { rgb: bg }, patternType: 'solid' } });
      });
    });

    // Revenue total
    const revTotalR = incLabelRow + 1 + incRows.length;
    colLetters.forEach((c, ci) => {
      if (ci === 0) setStyle(ws, `${c}${revTotalR}`, cell(true, WHITE, 'left', GREEN));
      else if (ci <= months.length + 1) setStyle(ws, `${c}${revTotalR}`, { ...money(true, WHITE), fill: { fgColor: { rgb: GREEN }, patternType: 'solid' } });
      else setStyle(ws, `${c}${revTotalR}`, { ...pct, font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: WHITE } }, fill: { fgColor: { rgb: GREEN }, patternType: 'solid' } });
    });

    // Expense section
    const expLabelR = revTotalR + 2;
    colLetters.forEach(c => setStyle(ws, `${c}${expLabelR}`, hdr(LGRAY, BLACK, true, 10)));

    expRows.forEach((_, i) => {
      const r = expLabelR + 1 + i;
      const bg = i % 2 === 0 ? WHITE : LGRAY;
      colLetters.forEach((c, ci) => {
        if (ci === 0) setStyle(ws, `${c}${r}`, cell(false, BLACK, 'left', bg));
        else if (ci <= months.length + 1) setStyle(ws, `${c}${r}`, money(false, BLACK, bg));
        else setStyle(ws, `${c}${r}`, { ...pct, fill: { fgColor: { rgb: bg }, patternType: 'solid' } });
      });
    });

    // Expense total
    const expTotalR = expLabelR + 1 + expRows.length;
    colLetters.forEach((c, ci) => {
      if (ci === 0) setStyle(ws, `${c}${expTotalR}`, cell(true, WHITE, 'left', RED));
      else if (ci <= months.length + 1) setStyle(ws, `${c}${expTotalR}`, { ...money(true, WHITE), fill: { fgColor: { rgb: RED }, patternType: 'solid' } });
      else setStyle(ws, `${c}${expTotalR}`, { ...pct, font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: WHITE } }, fill: { fgColor: { rgb: RED }, patternType: 'solid' } });
    });

    // Net Profit row
    const netR = expTotalR + 2;
    const netColor = mom.totalNetProfit >= 0 ? GREEN : RED;
    colLetters.forEach((c, ci) => {
      const netVal = ci === 0 ? null : (ci <= months.length + 1 ? (aoa[netR - 1][ci] as number) : null);
      const isPos = netVal === null || (netVal as number) >= 0;
      const nc = isPos ? GREEN : RED;
      if (ci === 0) setStyle(ws, `${c}${netR}`, cell(true, WHITE, 'left', netColor));
      else if (ci <= months.length + 1) setStyle(ws, `${c}${netR}`, { ...money(true, WHITE), fill: { fgColor: { rgb: nc }, patternType: 'solid' } });
      else setStyle(ws, `${c}${netR}`, { ...pct, font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: WHITE } }, fill: { fgColor: { rgb: netColor }, patternType: 'solid' } });
    });

    ws['!cols'] = [colWidth(32), ...months.map(() => colWidth(14)), colWidth(14), ...months.slice(1).map(() => colWidth(12))];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: numCols - 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: numCols - 1 } }];
    ws['!freeze'] = { xSplit: 1, ySplit: 4 };
    XLSX.utils.book_append_sheet(wb, ws, '📅 Month-over-Month');
  }

  // ── 6. Flags ─────────────────────────────────────────────────────────────────
  {
    const aoa: (string | number)[][] = [
      ['FLAGS & AUDIT NOTES', '', '', ''],
      [],
      ['INCONSISTENT VENDORS', '', '', ''],
      ['Vendor', 'Reason', 'Accounts Affected', ''],
      ...analysis.inconsistentVendors.map(({ vendor, reason, accounts }, i) => [vendor, reason, accounts.join(', '), '']),
      [],
      ['DUPLICATE TRANSACTIONS', '', '', ''],
      ['Vendor', 'Amount', 'Date', 'Occurrences'],
      ...analysis.duplicates.map(({ name, amount, transactionDate, occurrences }) => [name, amount, transactionDate, occurrences]),
    ];

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ['A1','B1','C1','D1'].forEach(a => setStyle(ws, a, hdr(RED)));
    ['A3','B3','C3','D3'].forEach(a => setStyle(ws, a, hdr(GOLD, WHITE)));
    ['A4','B4','C4','D4'].forEach(a => setStyle(ws, a, hdr(DGRAY, BLACK)));
    analysis.inconsistentVendors.forEach((_, i) => {
      const r = 5 + i; const bg = i%2===0?WHITE:'FFF2CC';
      ['A','B','C'].forEach(c => setStyle(ws, `${c}${r}`, cell(false, BLACK, 'left', bg)));
    });

    const dupStart = analysis.inconsistentVendors.length + 7;
    ['A','B','C','D'].map(c => `${c}${dupStart}`).forEach(a => setStyle(ws, a, hdr(GOLD, WHITE)));
    ['A','B','C','D'].map(c => `${c}${dupStart+1}`).forEach(a => setStyle(ws, a, hdr(DGRAY, BLACK)));
    analysis.duplicates.forEach((_, i) => {
      const r = dupStart + 2 + i; const bg = i%2===0?WHITE:'FFDCDC';
      ['A','B','C','D'].forEach(c => setStyle(ws, `${c}${r}`, cell(false, BLACK, 'left', bg)));
    });

    ws['!cols'] = [colWidth(28), colWidth(30), colWidth(40), colWidth(14)];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    XLSX.utils.book_append_sheet(wb, ws, '🚩 Flags');
  }

  XLSX.writeFile(wb, `${baseName(fileName)}_analysis.xlsx`);
}

// ─── PDF ──────────────────────────────────────────────────────────────────────

export function exportPdf(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const section = (title: string, content: string) => `<div class="section"><h2>${title}</h2>${content}</div>`;
  const table = (headers: string[], rows: (string | number)[][]) =>
    `<table><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${baseName(fileName)} Report</title>
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:12px;color:#111;padding:32px}
h1{font-size:20px;margin-bottom:4px}.meta{color:#666;font-size:11px;margin-bottom:24px}
.section{margin-bottom:28px;page-break-inside:avoid}h2{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#2F5496;border-bottom:1px solid #2F5496;padding-bottom:4px;margin-bottom:10px}
table{width:100%;border-collapse:collapse;font-size:11px}th{background:#2F5496;color:#fff;text-align:left;padding:5px 8px;font-weight:600}
td{padding:4px 8px;border-bottom:1px solid #e5e7eb}tr:nth-child(even) td{background:#f1f5f9}
.kpi{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:12px}.kpi-card{border:1px solid #e5e7eb;border-radius:6px;padding:10px 16px;min-width:140px}
.kpi-label{font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.06em}.kpi-value{font-size:16px;font-weight:700;margin-top:2px}
.balanced{background:#d1fae5;color:#065f46}.not-balanced{background:#fee2e2;color:#991b1b}.pill{display:inline-block;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:600;margin-left:8px}
@media print{body{padding:16px}}</style></head><body>
<h1>Ledger Analysis Report</h1>
<p class="meta">Source: ${fileName} &nbsp;|&nbsp; Generated: ${new Date().toLocaleString()}</p>
${section('Summary',`<div class="kpi"><div class="kpi-card"><div class="kpi-label">Transactions</div><div class="kpi-value">${analysis.totalTransactions}</div></div><div class="kpi-card"><div class="kpi-label">Revenue</div><div class="kpi-value">${fmt(pl.totalRevenue)}</div></div><div class="kpi-card"><div class="kpi-label">Expenses</div><div class="kpi-value">${fmt(pl.totalExpenses)}</div></div><div class="kpi-card"><div class="kpi-label">Net Profit</div><div class="kpi-value">${fmt(pl.netProfit)}</div></div></div>`)}
${section('Profit & Loss',`<div class="kpi"><div class="kpi-card"><div class="kpi-label">Revenue</div><div class="kpi-value">${fmt(pl.totalRevenue)}</div></div><div class="kpi-card"><div class="kpi-label">Expenses</div><div class="kpi-value">${fmt(pl.totalExpenses)}</div></div><div class="kpi-card"><div class="kpi-label">Net Profit</div><div class="kpi-value">${fmt(pl.netProfit)}</div></div></div>${table(['Month','Revenue','Expenses','Net'],Object.entries(pl.monthlyBreakdown).map(([m,{revenue,expenses}])=>[m,fmt(revenue),fmt(expenses),fmt(revenue-expenses)]))}`)}
${section(`Balance Sheet <span class="pill ${bs.isBalanced?'balanced':'not-balanced'}">${bs.isBalanced?'Balanced ✓':'Not Balanced ✗'} (Δ ${fmt(bs.variance)})</span>`,`${table(['Category','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,fmt(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,fmt(value)]),...bs.equity.map(({account,value})=>['Equity',account,fmt(value)])])}`)}
${section('Cash Flow',`<div class="kpi"><div class="kpi-card"><div class="kpi-label">Net Profit</div><div class="kpi-value">${fmt(cf.netProfit)}</div></div><div class="kpi-card"><div class="kpi-label">Operating CF</div><div class="kpi-value">${fmt(cf.operatingCashFlow)}</div></div></div>${cf.adjustments.length>0?table(['Account','Change','Impact'],cf.adjustments.map(({account,change,impact})=>[account,fmt(change),fmt(impact)])):''}`)}
${section('Flags',analysis.inconsistentVendors.length>0||analysis.duplicates.length>0?`${analysis.inconsistentVendors.length>0?'<p style="font-weight:bold;margin-bottom:6px">Inconsistent Vendors</p>'+table(['Vendor','Reason','Accounts'],analysis.inconsistentVendors.map(({vendor,reason,accounts})=>[vendor,reason,accounts.join(', ')])):''}${analysis.duplicates.length>0?'<p style="font-weight:bold;margin:10px 0 6px">Duplicate Transactions</p>'+table(['Vendor','Amount','Date','×'],analysis.duplicates.map(({name,amount,transactionDate,occurrences})=>[name,amount,transactionDate,occurrences])):''}` :'<p style="color:#666">No flags found.</p>')}
</body></html>`;

  const win = window.open('', '_blank');
  if (!win) return;
  win.document.write(html);
  win.document.close();
  win.onload = () => win.print();
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function baseName(f: string) { return f.replace(/\.[^.]+$/, '') || 'ledger'; }
function triggerDownload(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name; a.click();
  URL.revokeObjectURL(url);
}
