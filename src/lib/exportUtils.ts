import * as XLSX from 'xlsx';
import type { LedgerAnalysis } from './analyzeLedger';
import type { BalanceSheet } from './generateBalanceSheet';
import type { CashFlowStatement } from './generateCashFlow';
import type { ProfitAndLoss } from './generatePL';

const fmt = (n: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n);

// ─── CSV ────────────────────────────────────────────────────────────────────

function toCsvSection(title: string, headers: string[], rows: (string | number)[][]): string {
  const escape = (v: string | number) => {
    const s = String(v);
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [
    title,
    headers.map(escape).join(','),
    ...rows.map((r) => r.map(escape).join(',')),
    '',
  ];
  return lines.join('\n');
}

export function exportCsv(
  fileName: string,
  analysis: LedgerAnalysis,
  pl: ProfitAndLoss,
  bs: BalanceSheet,
  cf: CashFlowStatement,
): void {
  const sections: string[] = [];

  sections.push(
    toCsvSection('LEDGER ANALYSIS', ['Metric', 'Value'], [
      ['Total Transactions', analysis.totalTransactions],
      ['Inconsistent Vendors', analysis.inconsistentVendors.length],
      ['Duplicate Transactions', analysis.duplicates.length],
    ]),
  );

  sections.push(
    toCsvSection('INCONSISTENT VENDORS', ['Vendor', 'Accounts'], analysis.inconsistentVendors.map(({ vendor, accounts }) => [vendor, accounts.join(' | ')])),
  );

  sections.push(
    toCsvSection('DUPLICATE TRANSACTIONS', ['Vendor', 'Amount', 'Date', 'Occurrences'], analysis.duplicates.map(({ name, amount, transactionDate, occurrences }) => [name, amount, transactionDate, occurrences])),
  );

  sections.push(
    toCsvSection('PROFIT & LOSS', ['Item', 'Amount'], [
      ['Total Revenue', fmt(pl.totalRevenue)],
      ['Total Expenses', fmt(pl.totalExpenses)],
      ['Net Profit', fmt(pl.netProfit)],
    ]),
  );

  sections.push(
    toCsvSection('P&L MONTHLY BREAKDOWN', ['Month', 'Revenue', 'Expenses'], Object.entries(pl.monthlyBreakdown).map(([month, { revenue, expenses }]) => [month, fmt(revenue), fmt(expenses)])),
  );

  sections.push(
    toCsvSection('BALANCE SHEET – ASSETS', ['Account', 'Balance'], bs.assets.map(({ account, value }) => [account, fmt(value)])),
  );

  sections.push(
    toCsvSection('BALANCE SHEET – LIABILITIES', ['Account', 'Balance'], bs.liabilities.map(({ account, value }) => [account, fmt(value)])),
  );

  sections.push(
    toCsvSection('BALANCE SHEET – EQUITY', ['Account', 'Balance'], bs.equity.map(({ account, value }) => [account, fmt(value)])),
  );

  sections.push(
    toCsvSection('BALANCE SHEET TOTALS', ['Category', 'Total', 'Balanced'], [
      ['Assets', fmt(bs.totals.assetsTotal), bs.isBalanced ? 'Yes' : 'No'],
      ['Liabilities', fmt(bs.totals.liabilitiesTotal), ''],
      ['Equity', fmt(bs.totals.equityTotal), ''],
    ]),
  );

  sections.push(
    toCsvSection('CASH FLOW STATEMENT', ['Item', 'Amount'], [
      ['Net Profit', fmt(cf.netProfit)],
      ['Operating Cash Flow', fmt(cf.operatingCashFlow)],
    ]),
  );

  sections.push(
    toCsvSection('CASH FLOW ADJUSTMENTS', ['Account', 'Change', 'Impact'], cf.adjustments.map(({ account, change, impact }) => [account, fmt(change), fmt(impact)])),
  );

  const blob = new Blob([sections.join('\n')], { type: 'text/csv;charset=utf-8;' });
  triggerDownload(blob, `${baseName(fileName)}_analysis.csv`);
}

// ─── Excel ──────────────────────────────────────────────────────────────────

function makeSheet(headers: string[], rows: (string | number)[][]): XLSX.WorkSheet {
  return XLSX.utils.aoa_to_sheet([headers, ...rows]);
}

export function exportExcel(
  fileName: string,
  analysis: LedgerAnalysis,
  pl: ProfitAndLoss,
  bs: BalanceSheet,
  cf: CashFlowStatement,
): void {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    makeSheet(['Metric', 'Value'], [
      ['Total Transactions', analysis.totalTransactions],
      ['Inconsistent Vendors', analysis.inconsistentVendors.length],
      ['Duplicate Transactions', analysis.duplicates.length],
    ]),
    'Summary',
  );

  XLSX.utils.book_append_sheet(
    wb,
    makeSheet(['Vendor', 'Accounts'], analysis.inconsistentVendors.map(({ vendor, accounts }) => [vendor, accounts.join(' | ')])),
    'Inconsistent Vendors',
  );

  XLSX.utils.book_append_sheet(
    wb,
    makeSheet(['Vendor', 'Amount', 'Date', 'Occurrences'], analysis.duplicates.map(({ name, amount, transactionDate, occurrences }) => [name, amount, transactionDate, occurrences])),
    'Duplicates',
  );

  XLSX.utils.book_append_sheet(
    wb,
    makeSheet(
      ['Item', 'Amount'],
      [
        ['Total Revenue', pl.totalRevenue],
        ['Total Expenses', pl.totalExpenses],
        ['Net Profit', pl.netProfit],
        [],
        ['Month', 'Revenue', 'Expenses'],
        ...Object.entries(pl.monthlyBreakdown).map(([month, { revenue, expenses }]) => [month, revenue, expenses]),
      ],
    ),
    'Profit & Loss',
  );

  XLSX.utils.book_append_sheet(
    wb,
    makeSheet(
      ['Category', 'Account', 'Balance'],
      [
        ...bs.assets.map(({ account, value }) => ['Asset', account, value]),
        ...bs.liabilities.map(({ account, value }) => ['Liability', account, value]),
        ...bs.equity.map(({ account, value }) => ['Equity', account, value]),
        [],
        ['Assets Total', '', bs.totals.assetsTotal],
        ['Liabilities Total', '', bs.totals.liabilitiesTotal],
        ['Equity Total', '', bs.totals.equityTotal],
        ['Balanced', '', bs.isBalanced ? 'Yes' : 'No'],
      ],
    ),
    'Balance Sheet',
  );

  XLSX.utils.book_append_sheet(
    wb,
    makeSheet(
      ['Item', 'Amount'],
      [
        ['Net Profit', cf.netProfit],
        ['Operating Cash Flow', cf.operatingCashFlow],
        [],
        ['Account', 'Change', 'Impact'],
        ...cf.adjustments.map(({ account, change, impact }) => [account, change, impact]),
      ],
    ),
    'Cash Flow',
  );

  XLSX.writeFile(wb, `${baseName(fileName)}_analysis.xlsx`);
}

// ─── PDF ────────────────────────────────────────────────────────────────────

export function exportPdf(
  fileName: string,
  analysis: LedgerAnalysis,
  pl: ProfitAndLoss,
  bs: BalanceSheet,
  cf: CashFlowStatement,
): void {
  const section = (title: string, content: string) => `
    <div class="section">
      <h2>${title}</h2>
      ${content}
    </div>`;

  const table = (headers: string[], rows: (string | number)[][]) => `
    <table>
      <thead><tr>${headers.map((h) => `<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${rows.map((r) => `<tr>${r.map((c) => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;

  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>${baseName(fileName)} – Analysis Report</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 12px; color: #111; padding: 32px; }
  h1 { font-size: 20px; margin-bottom: 4px; }
  .meta { color: #666; font-size: 11px; margin-bottom: 24px; }
  .section { margin-bottom: 28px; page-break-inside: avoid; }
  h2 { font-size: 13px; font-weight: 700; text-transform: uppercase;
       letter-spacing: 0.08em; color: #2F5496; border-bottom: 1px solid #2F5496;
       padding-bottom: 4px; margin-bottom: 10px; }
  table { width: 100%; border-collapse: collapse; font-size: 11px; }
  th { background: #2F5496; color: #fff; text-align: left; padding: 5px 8px; font-weight: 600; }
  td { padding: 4px 8px; border-bottom: 1px solid #e5e7eb; }
  tr:nth-child(even) td { background: #f1f5f9; }
  .kpi { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 12px; }
  .kpi-card { border: 1px solid #e5e7eb; border-radius: 6px; padding: 10px 16px; min-width: 140px; }
  .kpi-label { font-size: 10px; color: #666; text-transform: uppercase; letter-spacing: 0.06em; }
  .kpi-value { font-size: 16px; font-weight: 700; margin-top: 2px; }
  .pill { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 11px;
          font-weight: 600; margin-left: 8px; }
  .balanced { background: #d1fae5; color: #065f46; }
  .not-balanced { background: #fee2e2; color: #991b1b; }
  @media print { body { padding: 16px; } }
</style>
</head>
<body>
<h1>Ledger Analysis Report</h1>
<p class="meta">Source: ${fileName} &nbsp;|&nbsp; Generated: ${new Date().toLocaleString()}</p>

${section('Summary', `
  <div class="kpi">
    <div class="kpi-card"><div class="kpi-label">Total Transactions</div><div class="kpi-value">${analysis.totalTransactions}</div></div>
    <div class="kpi-card"><div class="kpi-label">Inconsistent Vendors</div><div class="kpi-value">${analysis.inconsistentVendors.length}</div></div>
    <div class="kpi-card"><div class="kpi-label">Duplicate Count</div><div class="kpi-value">${analysis.duplicates.length}</div></div>
  </div>
`)}

${section('Profit & Loss', `
  <div class="kpi">
    <div class="kpi-card"><div class="kpi-label">Total Revenue</div><div class="kpi-value">${fmt(pl.totalRevenue)}</div></div>
    <div class="kpi-card"><div class="kpi-label">Total Expenses</div><div class="kpi-value">${fmt(pl.totalExpenses)}</div></div>
    <div class="kpi-card"><div class="kpi-label">Net Profit</div><div class="kpi-value">${fmt(pl.netProfit)}</div></div>
  </div>
  ${table(['Month', 'Revenue', 'Expenses'], Object.entries(pl.monthlyBreakdown).map(([m, { revenue, expenses }]) => [m, fmt(revenue), fmt(expenses)]))}
`)}

${section(`Balance Sheet <span class="pill ${bs.isBalanced ? 'balanced' : 'not-balanced'}">${bs.isBalanced ? 'Balanced ✓' : 'Not Balanced ✗'}</span>`, `
  ${table(['Category', 'Account', 'Balance'], [
    ...bs.assets.map(({ account, value }) => ['Asset', account, fmt(value)]),
    ...bs.liabilities.map(({ account, value }) => ['Liability', account, fmt(value)]),
    ...bs.equity.map(({ account, value }) => ['Equity', account, fmt(value)]),
  ])}
  <div class="kpi" style="margin-top:12px">
    <div class="kpi-card"><div class="kpi-label">Assets Total</div><div class="kpi-value">${fmt(bs.totals.assetsTotal)}</div></div>
    <div class="kpi-card"><div class="kpi-label">Liabilities Total</div><div class="kpi-value">${fmt(bs.totals.liabilitiesTotal)}</div></div>
    <div class="kpi-card"><div class="kpi-label">Equity Total</div><div class="kpi-value">${fmt(bs.totals.equityTotal)}</div></div>
  </div>
`)}

${section('Cash Flow Statement', `
  <div class="kpi">
    <div class="kpi-card"><div class="kpi-label">Net Profit</div><div class="kpi-value">${fmt(cf.netProfit)}</div></div>
    <div class="kpi-card"><div class="kpi-label">Operating Cash Flow</div><div class="kpi-value">${fmt(cf.operatingCashFlow)}</div></div>
  </div>
  ${cf.adjustments.length > 0 ? table(['Account', 'Change', 'Impact'], cf.adjustments.map(({ account, change, impact }) => [account, fmt(change), fmt(impact)])) : '<p style="color:#666;font-size:11px">No adjustments.</p>'}
`)}

${section('Inconsistent Vendors', analysis.inconsistentVendors.length > 0
  ? table(['Vendor', 'Accounts'], analysis.inconsistentVendors.map(({ vendor, accounts }) => [vendor, accounts.join(', ')]))
  : '<p style="color:#666;font-size:11px">None found.</p>'
)}

${section('Duplicate Transactions', analysis.duplicates.length > 0
  ? table(['Vendor', 'Amount', 'Date', 'Occurrences'], analysis.duplicates.map(({ name, amount, transactionDate, occurrences }) => [name, amount, transactionDate, occurrences]))
  : '<p style="color:#666;font-size:11px">None found.</p>'
)}

</body>
</html>`;

  const win = window.open('', '_blank');
  if (!win) return;
  win.document.write(html);
  win.document.close();
  win.onload = () => {
    win.print();
  };
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function baseName(fileName: string): string {
  return fileName.replace(/\.[^.]+$/, '') || 'ledger';
}

function triggerDownload(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}
