'use client';

import { useCallback, useMemo, useState } from 'react';
import { analyzeLedger, type LedgerAnalysis } from '../../src/lib/analyzeLedger';
import { generateBalanceSheet, type BalanceSheet } from '../../src/lib/generateBalanceSheet';
import { generateCashFlow, type CashFlowStatement } from '../../src/lib/generateCashFlow';
import { generatePL, type ProfitAndLoss } from '../../src/lib/generatePL';
import { exportCsv, exportExcel, exportPdf } from '../../src/lib/exportUtils';
import { FileDropzone, type UploadTheme } from './components/FileDropzone';
import { PreviewTable } from './components/PreviewTable';
import { ValidationPanel } from './components/ValidationPanel';
import {
  isCsvFile,
  isExcelFile,
  parseCsvFile,
  parseXlsxFile,
  requiredHeaders,
  type LedgerRow,
  type RowIssue,
} from './upload-utils';

type ThemeClasses = {
  page: string; shell: string; navbar: string; panel: string; card: string;
  listRow: string; heading: string; body: string; muted: string; stat: string;
  label: string; successPill: string; dangerPill: string; successCard: string;
  skyPill: string; skyCard: string; warningRow: string;
  settingsControl: string; settingsOptionActive: string; settingsOptionInactive: string;
};

const themes: Record<UploadTheme, ThemeClasses> = {
  dark: {
    page: 'bg-slate-950 text-slate-100',
    shell: 'border-slate-700 bg-slate-900/80 shadow-2xl shadow-slate-950/40',
    navbar: 'border-slate-700 bg-slate-900/95 backdrop-blur',
    panel: 'border-slate-700 bg-slate-900',
    card: 'bg-slate-950/60',
    listRow: 'bg-slate-950/50 text-slate-300',
    heading: 'text-white',
    body: 'text-slate-300',
    muted: 'text-slate-400',
    stat: 'text-white',
    label: 'text-slate-400',
    successPill: 'bg-emerald-950/70 text-emerald-200',
    dangerPill: 'bg-rose-950/70 text-rose-200',
    successCard: 'border-emerald-500/40 bg-emerald-950/50 text-emerald-100',
    skyPill: 'bg-sky-950/70 text-sky-200',
    skyCard: 'border-sky-500/40 bg-sky-950/50 text-sky-100',
    warningRow: 'bg-amber-950/50 text-amber-200',
    settingsControl: 'bg-slate-800 text-slate-200',
    settingsOptionActive: 'bg-cyan-400 text-slate-950',
    settingsOptionInactive: 'text-slate-300 hover:bg-slate-700',
  },
  light: {
    page: 'bg-slate-100 text-slate-900',
    shell: 'border-slate-200 bg-white shadow-xl shadow-slate-200/80',
    navbar: 'border-slate-200 bg-white/95 backdrop-blur',
    panel: 'border-slate-200 bg-slate-50',
    card: 'bg-white',
    listRow: 'bg-slate-50 text-slate-700',
    heading: 'text-slate-950',
    body: 'text-slate-600',
    muted: 'text-slate-500',
    stat: 'text-slate-950',
    label: 'text-slate-500',
    successPill: 'bg-emerald-100 text-emerald-800',
    dangerPill: 'bg-rose-100 text-rose-800',
    successCard: 'border-emerald-200 bg-emerald-50 text-emerald-900',
    skyPill: 'bg-sky-100 text-sky-800',
    skyCard: 'border-sky-200 bg-sky-50 text-sky-900',
    warningRow: 'bg-amber-50 text-amber-900',
    settingsControl: 'bg-slate-100 text-slate-700',
    settingsOptionActive: 'bg-slate-900 text-white',
    settingsOptionInactive: 'text-slate-600 hover:bg-slate-200',
  },
};

export default function UploadPage() {
  const [theme, setTheme] = useState<UploadTheme>('dark');
  const [view, setView] = useState<'upload' | 'dashboard'>('upload');
  const [previewRows, setPreviewRows] = useState<LedgerRow[]>([]);
  const [analysis, setAnalysis] = useState<LedgerAnalysis | null>(null);
  const [profitAndLoss, setProfitAndLoss] = useState<ProfitAndLoss | null>(null);
  const [balanceSheet, setBalanceSheet] = useState<BalanceSheet | null>(null);
  const [cashFlow, setCashFlow] = useState<CashFlowStatement | null>(null);
  const [headerErrors, setHeaderErrors] = useState<string[]>([]);
  const [rowIssues, setRowIssues] = useState<RowIssue[]>([]);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [fileName, setFileName] = useState('');
  const [isDragging, setIsDragging] = useState(false);

  const ui = themes[theme];

  const clearState = () => {
    setUploadError(null); setHeaderErrors([]); setRowIssues([]);
    setPreviewRows([]); setAnalysis(null); setProfitAndLoss(null);
    setBalanceSheet(null); setCashFlow(null);
  };

  const handleParse = useCallback(async (file: File) => {
    clearState();
    setFileName(file.name);
    if (!isCsvFile(file) && !isExcelFile(file)) {
      setUploadError('Only CSV and XLSX files are accepted.');
      return;
    }
    try {
      const result = isCsvFile(file) ? await parseCsvFile(file) : await parseXlsxFile(file);
      setHeaderErrors(result.headerErrors);
      setRowIssues(result.rowIssues);
      if (result.headerErrors.length > 0) { setPreviewRows([]); return; }
      const plResult = generatePL(result.rows);
      setAnalysis(analyzeLedger(result.rows));
      setProfitAndLoss(plResult);
      setBalanceSheet(generateBalanceSheet(result.rows));
      setCashFlow(generateCashFlow(result.rows, plResult));
      setPreviewRows(result.rows.slice(0, 50));
      setView('dashboard');
    } catch {
      setUploadError('Unable to parse file. Please verify the file format.');
    }
  }, []);

  const handleFileChange = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) await handleParse(file);
  }, [handleParse]);

  const handleDrop = useCallback(async (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault(); setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) await handleParse(file);
  }, [handleParse]);

  const handleDragOver = useCallback((e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault(); setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => setIsDragging(false), []);

  const handleUploadNew = () => { clearState(); setFileName(''); setView('upload'); };

  const currencyFormatter = useMemo(
    () => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }), [],
  );

  const ThemeToggle = () => (
    <div className={`flex rounded-full p-1 text-xs font-semibold ${ui.settingsControl}`}>
      {(['dark', 'light'] as const).map((t) => (
        <button key={t} type="button" onClick={() => setTheme(t)}
          className={`rounded-full px-3 py-1.5 capitalize transition-colors ${theme === t ? ui.settingsOptionActive : ui.settingsOptionInactive}`}>
          {t}
        </button>
      ))}
    </div>
  );

  const metricCard = (label: string, value: string | number) => (
    <div className={`rounded-2xl p-4 shadow-sm ${ui.card}`}>
      <p className={`text-sm ${ui.label}`}>{label}</p>
      <p className={`mt-2 text-2xl font-semibold ${ui.stat}`}>{value}</p>
    </div>
  );

  const renderBalanceSheetEntries = (entries: BalanceSheet['assets']) => {
    if (entries.length === 0) return <p className={`mt-3 text-sm ${ui.muted}`}>No accounts.</p>;
    return (
      <div className="mt-3 space-y-2">
        {entries.map((e) => (
          <div key={e.account} className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl px-4 py-3 text-sm ${ui.listRow}`}>
            <span className={`font-semibold ${ui.stat}`}>{e.account}</span>
            <span>{currencyFormatter.format(e.value)}</span>
          </div>
        ))}
      </div>
    );
  };

  // ─── Upload View ───────────────────────────────────────────────────────────

  if (view === 'upload') {
    return (
      <div className={`min-h-screen px-4 py-10 ${ui.page}`}>
        <div className={`mx-auto w-full max-w-3xl rounded-3xl border p-8 ${ui.shell}`}>
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h1 className={`text-2xl font-bold ${ui.heading}`}>Ledger Insight AI</h1>
                <p className={`mt-1 text-sm ${ui.muted}`}>Upload a ledger file to generate financial reports.</p>
              </div>
              <ThemeToggle />
            </div>

            <div className={`rounded-2xl border p-5 ${ui.panel}`}>
              <h2 className={`text-sm font-semibold uppercase tracking-widest ${ui.label}`}>Sample Files</h2>
              <p className={`mt-1 text-sm ${ui.body}`}>Download a sample to see the required format:</p>
              <div className="mt-3 flex flex-wrap gap-3">
                <a href="/samples/sample-ledger.csv" download className="rounded-lg bg-cyan-600 px-4 py-2 text-sm font-semibold text-white shadow transition-colors hover:bg-cyan-700">↓ CSV Sample</a>
                <a href="/samples/sample-ledger.xlsx" download className="rounded-lg bg-cyan-600 px-4 py-2 text-sm font-semibold text-white shadow transition-colors hover:bg-cyan-700">↓ XLSX Sample</a>
              </div>
            </div>

            <FileDropzone fileName={fileName} isDragging={isDragging} theme={theme}
              onFileChange={handleFileChange} onDragOver={handleDragOver}
              onDragLeave={handleDragLeave} onDrop={handleDrop} />

            {(uploadError || headerErrors.length > 0 || rowIssues.length > 0) && (
              <ValidationPanel headerErrors={headerErrors} uploadError={uploadError} rowIssues={rowIssues} theme={theme} />
            )}

            <div className={`rounded-2xl border p-5 ${ui.panel}`}>
              <h2 className={`text-sm font-semibold uppercase tracking-widest ${ui.label}`}>Required Headers</h2>
              <div className="mt-3 grid gap-2 sm:grid-cols-2">
                {requiredHeaders.map((h) => (
                  <div key={h} className={`rounded-xl px-4 py-2.5 text-sm ${ui.card} ${ui.body}`}>{h}</div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ─── Dashboard View ────────────────────────────────────────────────────────

  return (
    <div className={`min-h-screen ${ui.page}`}>
      {/* Sticky Navbar */}
      <header className={`sticky top-0 z-50 border-b ${ui.navbar}`}>
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-3 px-6 py-3">
          <div className="flex items-center gap-3">
            <h1 className={`text-base font-bold ${ui.heading}`}>Ledger Insight AI</h1>
            <span className={`hidden rounded-full px-3 py-1 text-xs font-medium sm:block ${ui.card} ${ui.muted}`}>{fileName}</span>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <ThemeToggle />
            {analysis && profitAndLoss && balanceSheet && cashFlow && (
              <>
                <button type="button" onClick={() => exportCsv(fileName, analysis, profitAndLoss, balanceSheet, cashFlow)}
                  className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-emerald-700">↓ CSV</button>
                <button type="button" onClick={() => exportExcel(fileName, analysis, profitAndLoss, balanceSheet, cashFlow)}
                  className="rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-blue-700">↓ Excel</button>
                <button type="button" onClick={() => exportPdf(fileName, analysis, profitAndLoss, balanceSheet, cashFlow)}
                  className="rounded-lg bg-rose-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-rose-700">↓ PDF</button>
              </>
            )}
            <button type="button" onClick={handleUploadNew}
              className="rounded-lg bg-cyan-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-cyan-700">↑ New File</button>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-6 px-6 py-8">
        {rowIssues.length > 0 && (
          <ValidationPanel headerErrors={[]} uploadError={null} rowIssues={rowIssues} theme={theme} />
        )}

        {/* Ledger Analysis */}
        {analysis && (
          <div className={`rounded-3xl border p-6 ${ui.panel}`}>
            <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Overview</p>
            <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Ledger Analysis</h2>
            <div className="mt-5 grid gap-4 sm:grid-cols-3">
              {metricCard('Total Transactions', analysis.totalTransactions)}
              {metricCard('Inconsistent Vendors', analysis.inconsistentVendors.length)}
              {metricCard('Duplicate Count', analysis.duplicates.length)}
            </div>
            {analysis.inconsistentVendors.length > 0 && (
              <div className={`mt-4 rounded-2xl p-4 ${ui.card}`}>
                <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Inconsistent Vendors</h3>
                <div className="mt-3 space-y-2">
                  {analysis.inconsistentVendors.map(({ vendor, accounts }) => (
                    <div key={vendor} className={`rounded-2xl px-4 py-3 text-sm ${ui.warningRow}`}>
                      <span className="font-semibold">{vendor}</span>: {accounts.join(', ')}
                    </div>
                  ))}
                </div>
              </div>
            )}
            {analysis.duplicates.length > 0 && (
              <div className={`mt-4 rounded-2xl p-4 ${ui.card}`}>
                <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Duplicate Transactions</h3>
                <div className="mt-3 space-y-2">
                  {analysis.duplicates.map(({ name, amount, transactionDate, occurrences }) => (
                    <div key={`${name}-${amount}-${transactionDate}`} className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl px-4 py-3 text-sm ${ui.warningRow}`}>
                      <span className="font-semibold">{name}</span>
                      <span>{amount} · {transactionDate} · {occurrences}×</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* P&L + Balance Sheet */}
        <div className="grid gap-6 lg:grid-cols-2">
          {profitAndLoss && (
            <div className={`rounded-3xl border p-6 ${ui.panel}`}>
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div>
                  <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Financials</p>
                  <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Profit &amp; Loss</h2>
                </div>
                <div className={`rounded-full px-3 py-1.5 text-xs font-semibold ${ui.successPill}`}>
                  Net {currencyFormatter.format(profitAndLoss.netProfit)}
                </div>
              </div>
              <div className="mt-5 grid gap-4 sm:grid-cols-3">
                {metricCard('Revenue', currencyFormatter.format(profitAndLoss.totalRevenue))}
                {metricCard('Expenses', currencyFormatter.format(profitAndLoss.totalExpenses))}
                <div className={`rounded-2xl border p-4 shadow-sm ${ui.successCard}`}>
                  <p className="text-sm">Net Profit</p>
                  <p className="mt-2 text-2xl font-semibold">{currencyFormatter.format(profitAndLoss.netProfit)}</p>
                </div>
              </div>
              {Object.keys(profitAndLoss.monthlyBreakdown).length > 0 && (
                <div className={`mt-4 rounded-2xl p-4 ${ui.card}`}>
                  <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Monthly Breakdown</h3>
                  <div className="mt-3 space-y-2">
                    {Object.entries(profitAndLoss.monthlyBreakdown).map(([month, { revenue, expenses }]) => (
                      <div key={month} className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl px-4 py-3 text-sm ${ui.listRow}`}>
                        <span className={`font-semibold ${ui.stat}`}>{month}</span>
                        <span>Rev: {currencyFormatter.format(revenue)} · Exp: {currencyFormatter.format(expenses)}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}

          {balanceSheet && (
            <div className={`rounded-3xl border p-6 ${ui.panel}`}>
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div>
                  <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Financials</p>
                  <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Balance Sheet</h2>
                </div>
                <div className={`rounded-full px-3 py-1.5 text-xs font-semibold ${balanceSheet.isBalanced ? ui.successPill : ui.dangerPill}`}>
                  {balanceSheet.isBalanced ? 'Balanced ✅' : 'Not Balanced ❌'}
                </div>
              </div>
              <div className="mt-5 grid gap-4 sm:grid-cols-3">
                {metricCard('Assets', currencyFormatter.format(balanceSheet.totals.assetsTotal))}
                {metricCard('Liabilities', currencyFormatter.format(balanceSheet.totals.liabilitiesTotal))}
                {metricCard('Equity', currencyFormatter.format(balanceSheet.totals.equityTotal))}
              </div>
              <div className="mt-4 grid gap-3 lg:grid-cols-3">
                <div className={`rounded-2xl p-4 ${ui.card}`}>
                  <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Assets</h3>
                  {renderBalanceSheetEntries(balanceSheet.assets)}
                </div>
                <div className={`rounded-2xl p-4 ${ui.card}`}>
                  <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Liabilities</h3>
                  {renderBalanceSheetEntries(balanceSheet.liabilities)}
                </div>
                <div className={`rounded-2xl p-4 ${ui.card}`}>
                  <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Equity</h3>
                  {renderBalanceSheetEntries(balanceSheet.equity)}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Cash Flow */}
        {cashFlow && (
          <div className={`rounded-3xl border p-6 ${ui.panel}`}>
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div>
                <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Financials</p>
                <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Cash Flow Statement</h2>
              </div>
              <div className={`rounded-full px-3 py-1.5 text-xs font-semibold ${ui.skyPill}`}>
                Operating CF: {currencyFormatter.format(cashFlow.operatingCashFlow)}
              </div>
            </div>
            <div className="mt-5 grid gap-4 sm:grid-cols-2">
              {metricCard('Net Profit', currencyFormatter.format(cashFlow.netProfit))}
              <div className={`rounded-2xl border p-4 shadow-sm ${ui.skyCard}`}>
                <p className="text-sm">Operating Cash Flow</p>
                <p className="mt-2 text-2xl font-semibold">{currencyFormatter.format(cashFlow.operatingCashFlow)}</p>
              </div>
            </div>
            {cashFlow.adjustments.length > 0 && (
              <div className={`mt-4 rounded-2xl p-4 ${ui.card}`}>
                <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Adjustments</h3>
                <div className="mt-3 space-y-2">
                  {cashFlow.adjustments.map((adj) => (
                    <div key={adj.account} className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl px-4 py-3 text-sm ${ui.listRow}`}>
                      <span className={`font-semibold ${ui.stat}`}>{adj.account}</span>
                      <span>Change: {currencyFormatter.format(adj.change)} · Impact: {currencyFormatter.format(adj.impact)}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* Transaction Preview */}
        {previewRows.length > 0 && (
          <div className={`rounded-3xl border p-6 ${ui.panel}`}>
            <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Raw Data</p>
            <h2 className={`mt-0.5 mb-4 text-xl font-bold ${ui.heading}`}>Transaction Preview</h2>
            <PreviewTable columns={requiredHeaders} rows={previewRows} rowIssues={rowIssues} theme={theme} />
          </div>
        )}
      </main>
    </div>
  );
}
